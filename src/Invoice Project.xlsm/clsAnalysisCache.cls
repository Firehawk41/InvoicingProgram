VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAnalysisCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsAnalysisCache
' Author: JT
' Created: 2026-04-10
' Description: Materialised in-memory cache of all analyses.
'              No dependency on clsAnalysisService.
' DependsOn: clsLoggingSystem, clsAnalysisRepository, clsAnalysis
' ChangeLog:
'   - 1.0.0 - 2026-04-10 - Initial release
'
' ================================================================
' LAYER: Cache — depends on Repo only. No upward coupling to Service.
'
' BUILD STRATEGY (one query total):
'   Build() calls Repo.SelectAll() -> all analysis rows.
'   Each Analysis is hydrated in a single pass.
'
' WRITE STRATEGY (incremental, no full rebuild after writes):
'   After each Service write, the Service calls Upsert() or Remove().
'   Build() is only ever called once at init, or explicitly to refresh.
'
' LOOKUP CONTRACT:
'   GetByID / GetByName return Nothing if not found.
'   GetIDByName / GetNameByID raise if not found (use Exists* to guard).
' ================================================================

Private m_Logger As clsLoggingSystem
Private m_Repo   As clsAnalysisRepository
Private m_ByID   As Scripting.Dictionary   ' Long      -> clsAnalysis
Private m_ByName As Scripting.Dictionary   ' LCase key -> clsAnalysis

' ----------------------------------------------------------------
' Initialization
' ----------------------------------------------------------------
Public Sub Initialize(Repo As clsAnalysisRepository, Logger As clsLoggingSystem)
    Debug.Assert Not Repo Is Nothing
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing        ' enforce single-init

    Set m_Logger = Logger
    Set m_Repo = Repo
    Set m_ByID = New Scripting.Dictionary
    Set m_ByName = New Scripting.Dictionary

    Build
End Sub

' ----------------------------------------------------------------
' Build  —  full rebuild from Repo
' Called from Initialize; public to allow explicit refresh.
' ----------------------------------------------------------------
Public Sub Build()
    AssertInit "Build"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    m_ByID.RemoveAll
    m_ByName.RemoveAll

    Dim Analyses As Variant
    Analyses = m_Repo.SelectAll

    ' No analyses in DB is valid — leave the cache empty.
    If IsEmpty(Analyses) Then Exit Sub

    Dim i As Long
    Dim D As Scripting.Dictionary
    Dim A As clsAnalysis

    ' SelectAll column order: 0=ID, 1=name, 2=description
    For i = LBound(Analyses, 2) To UBound(Analyses, 2)
        Set D = New Scripting.Dictionary
        D.Add "analysis_id", Analyses(0, i)
        D.Add "analysis_name", Analyses(1, i)
        D.Add "analysis_description", Analyses(2, i)

        Set A = New clsAnalysis
        A.Initialize m_Logger
        A.LoadFromDict D

        StoreInIndex A
    Next i

    Exit Sub
EH:                         HandleError "Build", "Failed at row " & i
End Sub

' ----------------------------------------------------------------
' Incremental updates  (called by Service after each write)
' ----------------------------------------------------------------
Public Sub Upsert(A As clsAnalysis)
    AssertInit "Upsert"
    Debug.Assert Not A Is Nothing
    Debug.Assert A.ID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    ' Remove stale name key if the analysis was renamed.
    If m_ByID.Exists(A.ID) Then
        Dim OldKey As String
        OldKey = NormName(m_ByID(A.ID).Name)
        If OldKey <> NormName(A.Name) Then
            If m_ByName.Exists(OldKey) Then m_ByName.Remove OldKey
        End If
    End If

    StoreInIndex A
    Exit Sub
EH:                         HandleError "Upsert", "Failed on ID: " & A.ID
End Sub

Public Sub Remove(AnalysisID As Long)
    AssertInit "Remove"
    Debug.Assert AnalysisID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    If m_ByID.Exists(AnalysisID) Then
        Dim NameKey As String
        NameKey = NormName(m_ByID(AnalysisID).Name)
        If m_ByName.Exists(NameKey) Then m_ByName.Remove NameKey
        m_ByID.Remove AnalysisID
    End If
    Exit Sub
EH:                         HandleError "Remove", "Failed on ID: " & AnalysisID
End Sub

' ----------------------------------------------------------------
' Read API
' ----------------------------------------------------------------
Public Function ExistsByID(AnalysisID As Long) As Boolean
    AssertInit "ExistsByID"
    ExistsByID = m_ByID.Exists(AnalysisID)
End Function

Public Function ExistsByName(Name As String) As Boolean
    AssertInit "ExistsByName"
    ExistsByName = m_ByName.Exists(NormName(Name))
End Function

Public Function GetByID(AnalysisID As Long) As clsAnalysis
    AssertInit "GetByID"
    If m_ByID.Exists(AnalysisID) Then Set GetByID = m_ByID(AnalysisID)
End Function

Public Function GetByName(Name As String) As clsAnalysis
    AssertInit "GetByName"
    Dim Key As String: Key = NormName(Name)
    If m_ByName.Exists(Key) Then Set GetByName = m_ByName(Key)
End Function

Public Function GetIDByName(Name As String) As Long
    AssertInit "GetIDByName"
    Dim Key As String: Key = NormName(Name)
    If m_ByName.Exists(Key) Then
        GetIDByName = m_ByName(Key).ID
    Else
        Err.Raise vbObjectError + 103, "clsAnalysisCache.GetIDByName", _
            "Analysis not found: " & Name
    End If
End Function

Public Function GetNameByID(AnalysisID As Long) As String
    AssertInit "GetNameByID"
    If m_ByID.Exists(AnalysisID) Then
        GetNameByID = m_ByID(AnalysisID).Name
    Else
        Err.Raise vbObjectError + 103, "clsAnalysisCache.GetNameByID", _
            "Analysis ID not found: " & AnalysisID
    End If
End Function

Public Function AllAnalyses() As Collection
    AssertInit "AllAnalyses"
    Dim Coll As New Collection
    Dim Key  As Variant
    For Each Key In m_ByID.Keys
        Coll.Add m_ByID(Key)
    Next Key
    Set AllAnalyses = Coll
End Function

' ----------------------------------------------------------------
' Private helpers
' ----------------------------------------------------------------
Private Sub StoreInIndex(A As clsAnalysis)
    m_ByID(A.ID) = A
    m_ByName(NormName(A.Name)) = A
End Sub

Private Function NormName(V As String) As String
    NormName = LCase$(Trim$(V))
End Function

Private Sub AssertInit(Src As String)
    If m_Logger Is Nothing Then
        Err.Raise vbObjectError + 1, "clsAnalysisCache." & Src, "Not initialized"
    End If
End Sub

Private Sub HandleError(Src As String, Context As String)
    m_Logger.LogError "clsAnalysisCache." & Src, Err.Description & " | " & Context
    Err.Clear
End Sub


