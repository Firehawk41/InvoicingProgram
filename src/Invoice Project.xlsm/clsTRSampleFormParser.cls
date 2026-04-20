VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTRSampleFormParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsSampleFormParser
' Author: JT
' Created: 2026-04-10
' Description: Reads one sample row from an Excel testing request
'              worksheet and returns a fully-hydrated clsSample.
'              Owns ALL column-offset and cell-reading logic for samples.
' DependsOn: clsLoggingSystem, clsChemicalService, clsAnalysisService,
'            clsElementService, clsSample, modEnums, clsChemical, clsAnalysis
' ChangeLog:
'   - 1.0.0 - 2026-04-10 - Extracted from clsTestingRequestSample.Initialize
'
' ================================================================
' LAYER: Parser — form reading only. No DB access. No persistence.
'
' COLUMN MAP (per RequestType, 1-based):
'
'   Field              Chemical  Water  Wafer
'   -----------------  --------  -----  -----
'   Chemical name      col 3     -      col 4   (Water = constant "Water")
'   Processing time    col 11    col 14 col 9
'   Additional notes   col 14    col 17 col 10
'   Requested time     col 13    col 16 -       (Wafer has none)
'   Additional elems   col 5     col 5  col 6
'
' ANALYSIS LOOKUP:
'   Each non-empty cell in the sample row is looked up by its VALUE
'   in clsAnalysisService. This is a simplified 1:1 mapping that
'   replaces the old SLIMDict 1:many mapping. Extend if needed.
' ================================================================

Private m_Logger      As clsLoggingSystem
Private m_ChemicalSvc As clsChemicalService
Private m_AnalysisSvc As clsAnalysisService
Private m_ElementSvc  As clsElementService

' ----------------------------------------------------------------
' Initialization
' ----------------------------------------------------------------
Public Sub Initialize(ChemicalSvc As clsChemicalService, _
                      AnalysisSvc As clsAnalysisService, _
                      ElementSvc As clsElementService, _
                      Logger As clsLoggingSystem)
    Debug.Assert Not ChemicalSvc Is Nothing
    Debug.Assert Not AnalysisSvc Is Nothing
    Debug.Assert Not ElementSvc Is Nothing
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing        ' enforce single-init

    Set m_Logger = Logger
    Set m_ChemicalSvc = ChemicalSvc
    Set m_AnalysisSvc = AnalysisSvc
    Set m_ElementSvc = ElementSvc
End Sub

' ----------------------------------------------------------------
' Public API
' ----------------------------------------------------------------
' Returns a fully-hydrated clsSample for one row.
' SampleCell: the cell in col 2 (sample ID column) for this row.
' SampleRange: the full sample block (used to bound analysis scan).
'
Public Function BuildFromRow(SampleCell As Range, _
                             SampleRange As Range, _
                             RequestType As TestingRequestTypeEnum) As clsTRSample
    AssertInit "BuildFromRow"
    Debug.Assert Not SampleCell Is Nothing
    Debug.Assert Not SampleRange Is Nothing
    If Not m_Logger.DebugMode Then On Error GoTo EH

    ' 1. Resolve Chemical first — ID is required in the scalar dict.
    Dim FormChemName As String
    FormChemName = GetFormChemicalName(SampleCell.Worksheet, SampleCell, RequestType)
    Dim ChemicalID As Long
    ChemicalID = ResolveChemicalID(FormChemName, CStr(SampleCell.Value))

    ' 2. Build scalar dict and hydrate entity.
    Dim D As Scripting.Dictionary
    Set D = BuildSampleDict(SampleCell, RequestType, ChemicalID)

    Dim S As New clsTRSample
    S.Initialize m_Logger
    S.LoadFromDict D

    ' 3. Scan sample row for analyses (and additional elements analysis).
    AddAnalysesFromRow S, SampleCell, SampleRange, RequestType

    ' 4. Parse and attach additional elements.
    AddAdditionalElements S, SampleCell, SampleRange, RequestType

    Set BuildFromRow = S
    Exit Function
EH:                         HandleError "BuildFromRow", "Row: " & SampleCell.Row
End Function

' ----------------------------------------------------------------
' Private: scalar dict builder
' ----------------------------------------------------------------
Private Function BuildSampleDict(SampleCell As Range, _
                                 RequestType As TestingRequestTypeEnum, _
                                 ChemicalID As Long) As Scripting.Dictionary
    Dim D   As New Scripting.Dictionary
    Dim WkS As Worksheet
    Set WkS = SampleCell.Worksheet

    D.Add "sample_id", 0
    D.Add "sample_name", CStr(SampleCell.Value)
    D.Add "form_chemical_name", GetFormChemicalName(WkS, SampleCell, RequestType)
    D.Add "chemical_id", ChemicalID
    D.Add "processing_time", CLng(GetProcessingTime(WkS, SampleCell, RequestType))
    D.Add "additional_notes", GetColumnValue(WkS, SampleCell.Row, GetNotesCol(RequestType))
    D.Add "requested_time", GetColumnValue(WkS, SampleCell.Row, GetRequestedTimeCol(RequestType))

    Set BuildSampleDict = D
End Function

' ----------------------------------------------------------------
' Private: Chemical resolution
' ----------------------------------------------------------------
Private Function ResolveChemicalID(FormChemName As String, SampleName As String) As Long
    Dim C As clsChemical
    Set C = m_ChemicalSvc.GetByName(FormChemName)

    If C Is Nothing Then
        m_Logger.LogError "clsSampleFormParser.ResolveChemicalID", vbObjectError + 1, "Chemical not found: '" & FormChemName & "' — sample: " & SampleName
        ResolveChemicalID = 0
    Else
        ResolveChemicalID = C.ID
    End If
End Function

' ----------------------------------------------------------------
' Private: Analysis scanning
' ----------------------------------------------------------------
' Scans every cell in the sample row. For each non-empty cell,
' attempts a name lookup in AnalysisService.
' Also checks the additional-elements column HEADER for its analysis.
'
Private Sub AddAnalysesFromRow(S As clsTRSample, _
                               SampleCell As Range, _
                               SampleRange As Range, _
                               RequestType As TestingRequestTypeEnum)
    Dim SampleRow   As Range
    Dim Cell        As Range
    Dim A           As clsAnalysis

    Set SampleRow = SampleCell.Resize(1, SampleRange.Columns.Count)

    For Each Cell In SampleRow.Cells
        If Not IsEmpty(Cell) And Len(Trim$(CStr(Cell.Value))) > 0 Then
            Set A = m_AnalysisSvc.GetByName(CStr(Cell.Value))
            If Not A Is Nothing Then S.AddAnalysis A.ID
        End If
    Next Cell

    ' Check if the additional-elements column header maps to an analysis
    ' (e.g. "Additional Elements" is itself a billable analysis code).
    Dim AddElCol    As Long
    AddElCol = GetAdditionalElementCol(RequestType)
    If AddElCol > 0 Then
        Dim HeaderVal As String
        HeaderVal = CStr(SampleRange.Parent.Cells(SampleRange.Row - 1, AddElCol).Value)
        If Len(Trim$(HeaderVal)) > 0 Then
            Set A = m_AnalysisSvc.GetByName(HeaderVal)
            If Not A Is Nothing Then S.AddAnalysis A.ID
        End If
    End If
End Sub

' ----------------------------------------------------------------
' Private:
' ----------------------------------------------------------------
Private Function GetAnalysisStartCol(RequestType As TestingRequestTypeEnum) As Long
    Select Case RequestType
        Case TestingRequestTypeEnum.Chemical: GetAnalysisStartCol = 4
        Case TestingRequestTypeEnum.Water:    GetAnalysisStartCol = 4
        Case TestingRequestTypeEnum.Wafer:    GetAnalysisStartCol = 5
    End Select
End Function

' ----------------------------------------------------------------
' Private: Additional element parsing
' ----------------------------------------------------------------
Private Sub AddAdditionalElements(S As clsTRSample, _
                                  SampleCell As Range, _
                                  SampleRange As Range, _
                                  RequestType As TestingRequestTypeEnum)
    Dim AddElCol As Long
    AddElCol = GetAdditionalElementCol(RequestType)
    If AddElCol = 0 Then Exit Sub

    Dim WkS     As Worksheet
    Set WkS = SampleCell.Worksheet

    Dim CellVal As String
    CellVal = CStr(WkS.Cells(SampleCell.Row, AddElCol).Value)
    If Len(Trim$(CellVal)) = 0 Then Exit Sub

    Dim Symbols As Variant
    Symbols = ParseElementSymbols(CellVal)

    Dim i       As Long
    Dim E       As clsElement
    For i = LBound(Symbols) To UBound(Symbols)
        If Len(Trim$(Symbols(i))) > 0 Then
            Set E = m_ElementSvc.GetBySymbol(CStr(Symbols(i)))
            If Not E Is Nothing Then
                S.AddElement E.ID
            Else
                m_Logger.LogError "clsSampleFormParser.AddAdditionalElements", _
                    "Element symbol not found: '" & Symbols(i) & "' — sample: " & S.SampleName
            End If
        End If
    Next i
End Sub

' ----------------------------------------------------------------
' Private: column-offset helpers
' ----------------------------------------------------------------
Private Function GetFormChemicalName(WkS As Worksheet, SampleCell As Range, _
                                     RequestType As TestingRequestTypeEnum) As String
    Select Case RequestType
        Case TestingRequestTypeEnum.Chemical
            GetFormChemicalName = CStr(WkS.Cells(SampleCell.Row, 3).Value)
        Case TestingRequestTypeEnum.Water
            GetFormChemicalName = "Water"
        Case TestingRequestTypeEnum.Wafer
            GetFormChemicalName = CStr(WkS.Cells(SampleCell.Row, 4).Value)
    End Select
End Function

Private Function GetProcessingTime(WkS As Worksheet, SampleCell As Range, _
                                   RequestType As TestingRequestTypeEnum) As ProcessingTimeEnum
    Dim Col As Long
    Select Case RequestType
        Case TestingRequestTypeEnum.Chemical: Col = 11
        Case TestingRequestTypeEnum.Water:    Col = 14
        Case TestingRequestTypeEnum.Wafer:    Col = 9
    End Select
    GetProcessingTime = ProcessingTimeToEnum(CStr(WkS.Cells(SampleCell.Row, Col).Value))
End Function

Private Function GetNotesCol(RequestType As TestingRequestTypeEnum) As Long
    Select Case RequestType
        Case TestingRequestTypeEnum.Chemical: GetNotesCol = 14
        Case TestingRequestTypeEnum.Water:    GetNotesCol = 17
        Case TestingRequestTypeEnum.Wafer:    GetNotesCol = 10
    End Select
End Function

Private Function GetRequestedTimeCol(RequestType As TestingRequestTypeEnum) As Long
    Select Case RequestType
        Case TestingRequestTypeEnum.Chemical: GetRequestedTimeCol = 13
        Case TestingRequestTypeEnum.Water:    GetRequestedTimeCol = 16
        Case TestingRequestTypeEnum.Wafer:    GetRequestedTimeCol = 0   ' not present
    End Select
End Function

Private Function GetAdditionalElementCol(RequestType As TestingRequestTypeEnum) As Long
    Select Case RequestType
        Case TestingRequestTypeEnum.Chemical, TestingRequestTypeEnum.Water
            GetAdditionalElementCol = 5
        Case TestingRequestTypeEnum.Wafer
            GetAdditionalElementCol = 6
        Case Else
            GetAdditionalElementCol = 0
    End Select
End Function

Private Function GetColumnValue(WkS As Worksheet, Row As Long, Col As Long) As String
    If Col = 0 Then Exit Function   ' column not applicable for this request type
    GetColumnValue = CStr(WkS.Cells(Row, Col).Value)
End Function

' ----------------------------------------------------------------
' Private: element string parsing
' ----------------------------------------------------------------
Private Function ParseElementSymbols(ElementString As String) As Variant
    Dim S As String
    S = ElementString
    S = Replace(S, ";", ",")
    S = Replace(S, ".", ",")
    S = Replace(S, " ", ",")
    S = Replace(S, vbCrLf, ",")
    S = Replace(S, vbCr, ",")
    S = Replace(S, vbLf, ",")

    Dim Raw()    As String
    Dim Clean()  As String
    Dim Count    As Long
    Dim i        As Long
    Dim Token    As String

    Raw = Split(S, ",")
    ReDim Clean(1 To UBound(Raw) - LBound(Raw) + 1)
    Count = 0
    For i = LBound(Raw) To UBound(Raw)
        Token = Trim$(Raw(i))
        If Len(Token) > 0 Then
            Count = Count + 1
            Clean(Count) = Token
        End If
    Next i

    If Count > 0 Then
        ReDim Preserve Clean(1 To Count)
        ParseElementSymbols = Clean
    Else
        ParseElementSymbols = Array()
    End If
End Function

' ----------------------------------------------------------------
' Private helpers
' ----------------------------------------------------------------
Private Sub AssertInit(Src As String)
    If m_Logger Is Nothing Then
        Err.Raise vbObjectError + 1, "clsSampleFormParser." & Src, "Not initialized"
    End If
End Sub

Private Sub HandleError(Src As String, Optional Context As String = "")
    m_Logger.LogError "clsSampleFormParser." & Src, Err.Description & " | " & Context
    Err.Clear
End Sub
