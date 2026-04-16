VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQuoteLineItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsQuoteLineItem
' Version: 1.2.0
' Author: JT
' Created: ?
' LastModified: 2025-12-01
' Description: Object for encapsulating quote line item properties
' DependsOn: clsLoggingSystem, modEnums, clsQuotePrice
' ChangeLog:
'   - 1.1.0 - 2025-11-21 - Added metadata and DBColumnHeaders property
'   - 1.2.0 - 2025-12-01 - Added a column map and used that as the single source of truth for columns and properties
'   - 1.3.0 - 2025-12-17 - Added a pass through to avoid exposing the price manager for prices
'   - 2.0.0 - 2025-12-31 - Big refactor - moved responsibilities around to establish SRP. Introduced ChemicalService and new style AnalysesManager. Stored chemicalID and AnalysisID instead of
'                           full objects as those objects are independent of the quote object heirachy
'   - 2.1.0 - 2026-04-07 - Refactor to pair down to ToDict/FromDict/AddPrice/Validate
'   - 2.2.0 - 2026-04-14 - Removed analysis and chemical managers so that the object only carries a collection of IDs

Private m_Logger As clsLoggingSystem
Private m_ID As Long
Private m_QuoteID As Long
Private m_RequestType As TestingRequestTypeEnum
Private m_BulkPriceMin As Long
Private m_IsValid As Boolean
Private m_ChemicalID As Variant ' Variant instead of long so that it can be nullable
Private m_AnalysisIDs As Collection
Private m_ElementIDs As Collection
Private m_Prices As Scripting.Dictionary ' Key: ProcessingTimeEnum, Value: clsQuotePrice

' ==============
' Get Properties
' ==============
Public Property Get ID() As Long
    ID = m_ID
End Property
Public Property Get QuoteID() As Long
    QuoteID = m_QuoteID
End Property
Public Property Get RequestType() As TestingRequestTypeEnum
    RequestType = m_RequestType
End Property
Public Property Get BulkPriceMin() As Long
    BulkPriceMin = m_BulkPriceMin
End Property
Public Property Get ChemicalName() As String
    If IsNull(m_ChemicalID) Then
        ChemicalName = ""
    Else
        ChemicalName = m_ChemicalService.GetName(m_ChemicalID)
    End If
End Property
Public Property Get ChemicalID() As Variant
    ChemicalID = m_ChemicalID
End Property
Public Property Get IsValid() As Boolean
    IsValid = m_IsValid
End Property
Public Property Get AllPrices() As Scripting.Dictionary
    Set AllPrices = m_Prices
End Property
' ==============
' Initialization
' ==============
Public Sub Initialize(Logger As clsLoggingSystem)

    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing

    ' Initialization not logged so as to not overwhelm the log with noise when all quotes are cached.

    If Logger Is Nothing Then Err.Raise vbObjectError + 1, , "Logger not set"
    If Not m_Logger Is Nothing Then Err.Raise vbObjectError + 2, , "Already initialized"

    If Not Logger.DebugMode Then On Error GoTo ErrorHandler

    Set m_Logger = Logger
    Set m_ElementIDs = New Collection
    Set m_AnalysisIDs = New Collection
    Set m_Prices = New Scripting.Dictionary

    Exit Sub
ErrorHandler:
    Call HandleError("Initialize", "Logger " & IIf(m_Logger Is Nothing, "not", "") & " set. ID set to: " & m_ID)

End Sub

' =====
' API
' =====
Public Function ToDict() As Scripting.Dictionary
    Dim D As New Scripting.Dictionary

    AssertInitialized "ToDict"
    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    ValidateStateDebug

    D("id") = m_ID
    D("quote_id") = m_QuoteID
    D("request_type_id") = m_RequestType
    D("bulk_min") = m_BulkPriceMin
    D("valid") = m_IsValid
    D("chemical_id") = m_ChemicalID
    D("analysis_ids") = CollectionToArray(m_AnalysisIDs)
    D("element_ids") = CollectionToArray(m_ElementIDs)

    Set ToDict = D

    Exit Function
ErrorHandler:
    HandleError "ToDict", "Error exporting line item " & m_ID
End Function
Public Sub FromDict(RowDict As Scripting.Dictionary)

    AssertInitialized "FromDict"
    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    m_ID = RowDict("id")
    m_QuoteID = RowDict("quote_id")
    m_RequestType = RowDict("request_type_id")
    m_BulkPriceMin = RowDict("bulk_min")
    m_IsValid = RowDict("valid")

    If RowDict.Exists("chemical_id") Then
        If IsNull(RowDict("chemical_id")) Then
            m_ChemicalID = Null
        Else
            m_ChemicalID = CLng(RowDict("chemical_id"))
        End If
    End If

    ValidateStateDebug

    Exit Sub
ErrorHandler:
    HandleError "FromDict", "Error loading line item " & RowDict("id")
End Sub
Public Sub Validate()
    Require m_ID >= 0, "Invalid ID"
    Require m_QuoteID >= 0, "Invalid QuoteID"
    Require m_RequestType >= RequestType_Min And m_RequestType <= RequestType_Max, "Invalid RequestType"
    Require m_BulkPriceMin >= -1, "Invalid BulkPriceMin"
    Require IsNumeric(m_ChemicalID) Or IsNull(m_ChemicalID), "Invalid ChemicalID"
    Require m_Prices.Count > 0, "Line item must have at least one price"
End Sub
Public Function GetPrice(TimeType As ProcessingTimeEnum) As clsQuotePrice
    Debug.Assert TimeType >= ProcessingTime_Min And TimeType <= ProcessingTime_Max

    AssertInitialized "GetPrice"

    ' Enable error handler if the debug mode is off
    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    If m_Prices.Exists(TimeType) Then
        Set GetPrice = m_Prices(TimeType)
    Else
        Set GetPrice = Nothing
    End If

    Exit Function
ErrorHandler:
    Call HandleError("GetPrice", "TimeType=" & TimeType & "; m_Prices.Count=" & m_Prices.Count)
End Function

Public Sub UpsertPrice(P As clsQuotePrice)
    Debug.Assert Not P Is Nothing
    Debug.Assert P.RegularPrice >= -1
    Debug.Assert P.BulkPrice >= -1
    Debug.Assert P.ProcessingTime >= ProcessingTime_Min And P.ProcessingTime <= ProcessingTime_Max

    AssertInitialized "UpsertPrice"

    ' Enable error handler if the debug mode is off
    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    Set m_Prices(P.ProcessingTime) = P

    Exit Sub
ErrorHandler:
    Dim ErrMsg As String
    If P Is Nothing Then
        ErrMsg = "Price was not set."
    Else
        ErrMsg = "ID: " & P.ID & "; Processing Time enum: " & P.ProcessingTime & "; Regular Price: " & P.RegularPrice & "; Bulk Price: " & P.BulkPrice
    End If

    Call HandleError("UpsertPrice", "Error upserting price. " & ErrMsg)
End Sub

Public Sub AddAnalysisID(AnalysisID As Long)
    AssertInitialized "AddAnalysisID"
    m_AnalysisIDs.Add AnalysisID
End Sub
Public Sub AddElementID(ElementID As Long)
    AssertInitialized "AddElementID"
    m_ElementIDs.Add ElementID
End Sub
Public Sub RemovePrice(TimeType As ProcessingTimeEnum)
    Debug.Assert TimeType >= ProcessingTime_Min And TimeType <= ProcessingTime_Max

    AssertInitialized "RemovePrice"

    ' Enable error handler if the debug mode is off
    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    If m_Prices.Exists(TimeType) Then
        m_Prices.Remove TimeType
    End If

    Exit Sub
ErrorHandler:
    Call HandleError("RemovePrice", "Error removing price with processing time enum: " & TimeType)
End Sub

' ===============
' Private Helpers
' ===============
Private Function CollectionToArray(Coll As Collection) As Long()
    Dim Arr() As Long
    Dim i As Long

    ReDim Arr(1 To Coll.Count)

    For i = 1 To Coll.Count
        Arr(i) = Coll(i)
    Next i

    CollectionToArray = Arr
End Function
Private Function ArrayToCollection(Arr() As Long) As Collection
    Dim Coll As Collection
    Dim CustomerID As Variant

    Set Coll = New Collection

    For Each CustomerID In Arr
        Coll.Add CLng(CustomerID)
    Next CustomerID

    Set ArrayToCollection = Coll
End Function
Private Sub Require(Condition As Boolean, Message As String)
    If Not Condition Then
        Err.Raise 5, TypeName(Me), Message
    End If
End Sub
Private Sub ValidateStateDebug()
    Debug.Assert m_ID >= 0
    Debug.Assert m_QuoteID >= 0
    Debug.Assert m_RequestType >= RequestType_Min And m_RequestType <= RequestType_Max
    Debug.Assert m_BulkPriceMin >= -1
    Debug.Assert IsNumeric(m_ChemicalID) Or IsNull(m_ChemicalID)
    Debug.Assert m_Prices.Count > 0
End Sub
Private Sub HandleError(MethodName As String, Optional Context As String = "")
    Call m_Logger.LogError("clsQuoteLineItem." & MethodName, Err.Number, Err.Description & IIf(Context <> "", " | " & Context, ""), True)
End Sub
Private Sub AssertInitialized(MethodName As String)
    If m_Logger Is Nothing Then
        Err.Raise 911, "clsQuoteLineItem." & MethodName, "The object must be initialized before use."
    End If
End Sub

