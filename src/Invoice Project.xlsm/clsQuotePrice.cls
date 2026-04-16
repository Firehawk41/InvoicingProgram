VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQuotePrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsQuotePrice
' Author: JT
' Created: 2025-11-24
' Description: Object for encapsulating quote price properties
' DependsOn: clsLoggingSystem, modEnums
' ChangeLog:
'   - 1.1.0 - 2025-11-24 - Added metadata and DBColumnHeaders property
'   - 1.2.0 - 2025-12-01 - Added a column map and used that as the single source of truth for columns and properties
'   - 1.3.0 - 2025-12-30 - Reorganized formatting and code order for readability; Added CopyPrice
'   - 1.4.0 - 2026-04-08 - Refactored for SRP, removed legacy code, simplified initialization and property management

Private m_Logger As clsLoggingSystem
Private m_ID As Long
Private m_QuoteLineItemID As Long
Private m_RegularPrice As Long
Private m_BulkPrice As Long
Private m_ProcessingTime As ProcessingTimeEnum
Private m_IsValid As Boolean
' =====================
' Public Get Properties
' =====================
Public Property Get ID() As Long
    ID = m_ID
End Property
Public Property Get QuoteLineItemID() As Long
    QuoteLineItemID = m_QuoteLineItemID
End Property
Public Property Get ProcessingTime() As ProcessingTimeEnum
    ProcessingTime = m_ProcessingTime
End Property
Public Property Get RegularPrice() As Long
    RegularPrice = m_RegularPrice
End Property
Public Property Get BulkPrice() As Long
    BulkPrice = m_BulkPrice
End Property

' ==============
' Initialization
' ==============
Public Sub Initialize(Logger As clsLoggingSystem)
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing

    If Not m_Logger Is Nothing Then Err.Raise vbObjectError + 1, , "Price is already initialized"

    ' Enable error handler if the debug mode is off
    If Not Logger.DebugMode Then On Error GoTo ErrorHandler

    ' Set properties
    Set m_Logger = Logger

    Exit Sub
ErrorHandler:
    Call HandleError("Initialize", "Logger " & IIf(m_Logger Is Nothing, "not", "") & " set.")
End Sub

' ==========
' Public API
' ==========

Public Function ToDict() As Scripting.Dictionary
    Dim D As New Scripting.Dictionary

    AssertInitialized "ToDict"
    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    ValidateStateDebug

    D("id") = m_ID
    D("quote_line_item_id") = m_QuoteLineItemID
    D("regular_price") = m_RegularPrice
    D("bulk_price") = m_BulkPrice
    D("valid") = m_IsValid
    D("processing_time") = m_ProcessingTime

    Set ToDict = D

    Exit Function
ErrorHandler:
    HandleError "ToDict", "Error exporting line item " & m_ID
End Function
Public Sub FromDict(RowDict As Scripting.Dictionary)

    AssertInitialized "FromDict"
    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    m_ID = RowDict("id")
    m_QuoteLineItemID = RowDict("quote_line_item_id")
    m_RegularPrice = RowDict("regular_price")
    m_BulkPrice = RowDict("bulk_price")
    m_IsValid = RowDict("valid")
    m_ProcessingTime = RowDict("processing_time")

    ValidateStateDebug

    Exit Sub
ErrorHandler:
    HandleError "FromDict", "Error loading line item " & RowDict("id")
End Sub
Public Sub Validate()
    AssertInitialized "Validate"

    Require m_ID >= 0, "Invalid ID"
    Require m_QuoteLineItemID >= 0, "Invalid QuoteLineItemID"
    Require m_ProcessingTime >= ProcessingTime_Min And m_ProcessingTime <= ProcessingTime_Max, "Invalid RequestType"
    Require m_RegularPrice >= -1, "Invalid Regular Price"
    Require m_BulkPrice >= -1, "Invalid Bulk Price"

End Sub

' ================
' Private Helpers
' ================
Private Sub ValidateStateDebug()
    Debug.Assert m_ID >= 0
    Debug.Assert m_ProcessingTime >= ProcessingTime_Min And m_ProcessingTime <= ProcessingTime_Max
    Debug.Assert m_RegularPrice >= -1
    Debug.Assert m_BulkPrice >= -1
End Sub
Private Sub AssertInitialized(MethodName As String)
    If m_Logger Is Nothing Then
        Err.Raise 911, "clsQuotePrice." & MethodName, "The object must be initialized before use."
    End If
End Sub
Private Sub HandleError(MethodName As String, Optional Context As String = "")
    Call m_Logger.LogError("clsQuotePrice." & MethodName, Err.Number, Err.Description & IIf(Context <> "", " | " & Context, ""), True)
End Sub

