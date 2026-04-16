VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQuote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsQuote
' Version: 1.2.0
' Author: JT
' Created: ?
' LastModified: 2025-12-01
' Description: Object for encapsulating quote properties
' DependsOn: clsLoggingSystem, modEnums, clsQuoteLineItem
' ChangeLog:
'   - 1.1.0 - 2025-11-21 - Added metadata and DBColumnHeaders property
'   - 1.1.1 - 2025-11-24 - Changed the order of properties in 'Initialize'
'   - 1.2.0 - 2025-12-01 - Added a column map and used that as the single source of truth for columns and properties - experimental and overly complicated - but worth saving just in case
'   - 1.3.0 - 2026-04-14 - Remvoed Customer manager so that the object only stores a collection of IDs

Private m_Logger As clsLoggingSystem
Private m_CustomerIDs As Collection
Private m_LineItems As Collection
Private m_ID As Long
Private m_EffectiveDate As Date
Private m_ExpireDate As Date
Private m_PaymentType As String
Private m_PaymentTerms As Long
Private m_CreatedBy As String
Private m_Comments As String
Private m_CreatedDate As Date
Private m_ModifiedDate As Date
Private m_IsValid As Boolean
Private m_IsDefaultQuote As Boolean

' =====================
' Public Get Properties
' =====================
Public Property Get LineItems() As Collection
    Set LineItems = m_LineItems
End Property

' ===============
' Initialization
' ==============
Public Sub Initialize(Logger As clsLoggingSystem)

    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing

    If Logger Is Nothing Then Err.Raise vbObjectError + 1, , "Logger is not set"
    If Not Logger.DebugMode Then On Error GoTo ErrorHandler

    If Not m_Logger Is Nothing Then Err.Raise vbObjectError + 2, , "Quote is already initialized"

    Set m_Logger = Logger
    Set m_CustomerIDs = New Collection
    Set m_LineItems = New Collection

    Exit Sub
ErrorHandler:
    HandleError "Initialize"
End Sub
' =====================
' Public APIs
' ===================
Public Function ToDict() As Scripting.Dictionary
    Dim D As New Scripting.Dictionary

    AssertInitialized "ToDict"

    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    ValidateStateDebug

    D("ID") = m_ID
    D("customer_ids") = CollectionToArray(m_CustomerIDs)
    D("effective_date") = m_EffectiveDate
    D("expiry_date") = m_ExpireDate
    D("payment_type") = m_PaymentType
    D("payment_value") = m_PaymentTerms
    D("created_by") = m_CreatedBy
    D("comments") = m_Comments
    D("created_date") = m_CreatedDate
    If Not IsNull(m_ModifiedDate) Then
        D("modified_date") = m_ModifiedDate
    End If
    D("valid") = m_IsValid
    D("is_default_quote") = m_IsDefaultQuote

    Set ToDict = D

    Exit Function
ErrorHandler:
    HandleError "ToDict", "Error encountered while generating quote " & m_ID
End Function
Public Sub FromDict(RowDict As Scripting.Dictionary)
    Debug.Assert Not RowDict Is Nothing

    AssertInitialized "FromDict"

    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    m_ID = RowDict("ID")
    Set m_CustomerIDs = ArrayToCollection(RowDict("customer_ids"))
    m_EffectiveDate = RowDict("effective_date")
    m_ExpireDate = RowDict("expiry_date")
    m_PaymentType = RowDict("payment_type")
    m_PaymentTerms = RowDict("payment_value")
    m_CreatedBy = RowDict("created_by")
    m_Comments = RowDict("comments")
    m_CreatedDate = RowDict("created_date")
    If Not IsNull(RowDict("modified_date")) Then
        m_ModifiedDate = RowDict("modified_date")
    End If
    m_IsValid = RowDict("valid")
    m_IsDefaultQuote = RowDict("is_default_quote")

    ValidateStateDebug

    Exit Sub
ErrorHandler:
    HandleError "FromDict", "Error encountered while generating quote " & RowDict("ID")
End Sub
Public Sub UpsertLineItem(LI As clsQuoteLineItem)
    Debug.Assert Not LI Is Nothing

    AssertInitialized "UpsertLineItem"

    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    If LI.ID = 0 Then
        ' New line item
        m_LineItems.Add LI
        ' Possibly assign a new ID if coming from DB later
    Else
        ' Existing line item, find it and replace or update
        Dim ExistingLI As clsQuoteLineItem
        Set ExistingLI = FindByID(LI.ID)
        If Not ExistingLI Is Nothing Then
            ReplaceLineItem ExistingLI, LI
        Else
            ' Edge case: ID exists but not in collection
            m_LineItems.Add LI
        End If
    End If

    Exit Sub
ErrorHandler:
    Dim LI_ID As Long
    If Not LI Is Nothing Then LI_ID = LI.ID
    HandleError "UpsertLineItem", "Error upserting " & LI_ID
End Sub
Public Sub AddCustomerID(CustomerID As Long)
    AssertInitialized "AddCustomerID"
    m_CustomerIDs.Add CustomerID
End Sub
Public Sub Validate()
    Require m_ID >= 0, "Invalid ID"

    ' Optional depending on business rules
    If Not m_IsDefaultQuote Then
        Require m_CustomerIDs.Count > 0, "Quote must have at least one customer"
    End If

    Require m_LineItems.Count > 0, "Quote must have at least one line item"
    Require m_PaymentType <> "", "PaymentType required"
    Require m_CreatedBy <> "", "CreatedBy required"
    Require Not IsNull(m_EffectiveDate), "EffectiveDate required"
    Require Not IsNull(m_ExpireDate), "ExpireDate required"
    Require Not IsNull(m_CreatedDate), "CreatedDate required"
End Sub

' ========================
' Private Helper Functions
' ========================
Private Function CollectionToArray(Coll As Collection) As Long
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
Private Sub ValidateStateDebug()
    Debug.Assert m_ID >= 0
    Debug.Assert Not m_CustomerIDs Is Nothing
    Debug.Assert Not m_LineItems Is Nothing
    Debug.Assert m_PaymentType <> ""
    Debug.Assert m_CreatedBy <> ""
    Debug.Assert Not IsNull(m_EffectiveDate)
    Debug.Assert Not IsNull(m_ExpireDate)
    Debug.Assert Not IsNull(m_CreatedDate)
End Sub
Private Sub Require(Condition As Boolean, Message As String)
    If Not Condition Then
        Err.Raise 5, TypeName(Me), Message
    End If
End Sub
Private Function FindByID(LineItemID As Long) As clsQuoteLineItem
    Dim LI As clsQuoteLineItem
    For Each LI In m_LineItems
        If LI.ID = LineItemID Then
            Set FindByID = LI
            Exit Function
        End If
    Next LI
    ' Return Nothing if not found
    Set FindByID = Nothing
End Function

' Replace an existing line item in the collection with a new one
Private Sub ReplaceLineItem(OldLI As clsQuoteLineItem, NewLI As clsQuoteLineItem)
    Dim Index As Long
    Index = 1
    Dim LI As clsQuoteLineItem

    For Each LI In m_LineItems
        If LI.ID = OldLI.ID Then
            ' Remove old, add new in same position
            m_LineItems.Remove Index
            m_LineItems.Add NewLI, , Index
            Exit Sub
        End If
        Index = Index + 1
    Next LI

    ' If old item wasn't found, just add new one at the end
    m_LineItems.Add NewLI
End Sub
Private Sub HandleError(MethodName As String, Optional Context As String = "")
    Call m_Logger.LogError("clsQuote." & MethodName, Err.Number, Err.Description & IIf(Context <> "", " | " & Context, ""), True)
End Sub
Private Sub AssertInitialized(MethodName As String)
    If m_Logger Is Nothing Then
        Err.Raise 911, "clsQuote." & MethodName, "The builder must be initialized before use."
    End If
End Sub

