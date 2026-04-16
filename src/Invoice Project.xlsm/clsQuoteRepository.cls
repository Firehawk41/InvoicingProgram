VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQuoteRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsQuoteRepository
' Version: 0.1.0
' Author: JT
' Created: 2025-12-01
' LastModified: 2025-12-01
' Description: Repository class for quotes, in charge of db operations
' DependsOn: clsLoggingSystem, modEnums
' ChangeLog:
'   - 0.1.0 - 2025-12-01    - Initial prototype

Private m_AccessDB As clsAccessDatabase
Private m_Logger As clsLoggingSystem

Public Sub Initialize(AccessDB As clsAccessDatabase, Logger As clsLoggingSystem)
    Set m_AccessDB = AccessDB
    Set m_Logger = Logger

End Sub

' Public API
Public Sub SaveQuote(Q As clsQuote)
    On Error GoTo ErrorHandler

    Dim NewID As Long
    Dim Record As Scripting.Dictionary
    Set Record = Q.ToRecord

    If Q.ID = 0 Then
        NewID = InsertRecord("quotes", Record)
        Call Q.SetIDInternal(NewID)
    Else
        Call UpdateRecord("quotes", Record, Q.ID)
    End If

    ' TODO: not yet saving to quote_analysis or quote_line_item_elements tables 12/31/25

    Call SaveQuoteCustomers(Q)
    Call SaveLineItems(Q)

    Exit Sub
ErrorHandler:
    Call m_Logger.LogError("clsQuoteRepository.SaveQuote", Err.Number, Err.Description, True)
End Sub

Public Sub DeleteQuote(Q As clsQuote)
    On Error GoTo ErrorHandler

    ' 1. Delete prices first
    Dim LI As clsQuoteLineItem
    Dim P As clsQuotePrice
    For Each LI In Q.LineItems
        For Each P In LI.Prices
            DeletePrice P.ID
        Next P
        ' 2. Delete the line item
        DeleteLineItem LI.ID
    Next LI

    ' 3. Delete quote-customer links if you have them
    DeleteQuoteCustomers Q.ID

    ' 4. Delete the quote itself
    DeleteRecord "quotes", Q.ID

    Exit Sub
ErrorHandler:
    Call m_Logger.LogError("clsQuoteRepository.DeleteQuote", Err.Number, Err.Description, True)
End Sub

Public Function FetchRows(TableName As String, ColumnHeaders As String) As Collection
    Dim SQL As String
    Dim DataArray As Variant
    Dim RowDict As Scripting.Dictionary
    Dim Rows As New Collection
    Dim iIndex As Long
    Dim Headers() As String

    SQL = "SELECT " & ColumnHeaders & " FROM " & TableName & " ORDER BY ID;"
    DataArray = m_AccessDB.ExecuteQuery(SQL)

    For iIndex = LBound(DataArray, 2) To UBound(DataArray, 2)
        Headers = Split(ColumnHeaders, ", ")
        Set RowDict = ArrayRowToDict(DataArray, iIndex, Headers)
        Rows.Add RowDict
    Next iIndex

    Set FetchRows = Rows

End Function
Public Function GetQuoteAnalyses() As Variant
    GetQuoteAnalyses = m_AccessDB.ExecuteQuery("SELECT quote_line_item_id, analysis_id FROM quote_analysis")
End Function

Public Function GetQuoteElements() As Variant
    GetQuoteElements = m_AccessDB.ExecuteQuery("SELECT quote_line_item_id, element_id FROM quote_line_item_element")
End Function

Public Function GetQuoteCustomers() As Variant
    GetQuoteCustomers = m_AccessDB.ExecuteQuery("SELECT quote_id, customer_id FROM quote_customer")
End Function
Private Sub SaveQuoteCustomers(Q As clsQuote)
    Dim Key As Variant
    Dim CustomersDict As Scripting.Dictionary
    Dim C As clsTestingRequestCustomer
    Dim SQL As String

    Set CustomersDict = Q.CustomerManager.Customers
    ' Insert current links
    For Each Key In CustomersDict.Keys
        Set C = CustomersDict(Key)
        SQL = "INSERT INTO quote_customer (quote_id, customer_id) VALUES (" & Q.ID & ", " & C.CustomerID & ");"
        Call m_AccessDB.ExecuteNonQuery(SQL)
    Next Key
End Sub
' Line Items
Private Sub SaveLineItems(Q As clsQuote)
    Dim LI As clsQuoteLineItem

    For Each LI In Q.LineItems
        Call LI.SetQuoteIDInternal(Q.ID)
        Call SaveLineItem(LI)
    Next LI
End Sub
Private Sub SaveLineItem(LI As clsQuoteLineItem)
    Dim NewID As Long
    Dim Record As Scripting.Dictionary
    Set Record = LI.ToRecord

    If LI.ID = 0 Then
        NewID = InsertRecord("quote_line_items", Record)
        Call LI.SetIDInternal(NewID)
    Else
        Call UpdateRecord("quote_line_items", Record, LI.ID)
    End If

    Call SavePrices(LI)
End Sub

' Prices

Private Sub SavePrices(LI As clsQuoteLineItem)
    Dim P As clsQuotePrice

    For Each P In LI.Prices
        Call P.SetLineItemIDInternal(LI.ID)
        Call SavePrice(P)
    Next P
End Sub

Private Sub SavePrice(P As clsQuotePrice)
    Dim NewID As Long
    Dim Record As Scripting.Dictionary
    Set Record = P.ToRecord

    If P.ID = 0 Then
        NewID = InsertRecord("quote_prices", Record)
        Call P.SetIDInternal(NewID)
    Else
        Call UpdateRecord("quote_prices", Record, P.ID)
    End If

End Sub

' --- Helper deletion routines ---
Private Sub DeletePrice(PriceID As Long)
    Dim SQL As String
    SQL = "DELETE FROM quote_prices WHERE ID = " & PriceID & ";"
    Call m_AccessDB.ExecuteQuery(SQL)
End Sub

Private Sub DeleteLineItem(LineItemID As Long)
    Dim SQL As String
    SQL = "DELETE FROM quote_line_items WHERE ID = " & LineItemID & ";"
    Call m_AccessDB.ExecuteQuery(SQL)
End Sub

Private Sub DeleteQuoteCustomers(QuoteID As Long)
    Dim SQL As String
    SQL = "DELETE FROM quote_customer WHERE quote_id = " & QuoteID & ";"
    Call m_AccessDB.ExecuteQuery(SQL)
End Sub

Private Sub DeleteRecord(TableName As String, ID As Long)
    Dim SQL As String
    SQL = "DELETE FROM " & TableName & " WHERE ID = " & ID & ";"
    Call m_AccessDB.ExecuteQuery(SQL)
End Sub
' Shared SQL Helpers
Private Function InsertRecord(TableName As String, Record As Scripting.Dictionary) As Long

    ' Remove primary id if ID=0 (ie a new record) to prevent SQL failure
    If Record.Exists("ID") Then
        If Record("ID") = 0 Then Record.Remove "ID"
    End If

    Dim Fields As String: Fields = Join(Record.Keys, ", ")
    Dim Values As String: Values = BuildValuesList(Record)

    Dim SQL As String
    SQL = "INSERT INTO " & TableName & " (" & Fields & ") VALUES (" & Values & ");"

    InsertRecord = m_AccessDB.ExecuteNonQuery(SQL)
End Function

Private Sub UpdateRecord(TableName As String, Record As Scripting.Dictionary, ID As Long)
    Dim Assignments As String: Assignments = BuildAssignmentsList(Record)

    Dim SQL As String
    SQL = "UPDATE " & TableName & _
            " SET " & Assignments & _
            " WHERE ID = " & ID & ";"

    Call m_AccessDB.ExecuteQuery(SQL)
End Sub

' SQL String builders
Private Function BuildValuesList(Record As Scripting.Dictionary) As String
    Dim Key As Variant, V As Variant, Parts As Collection
    Set Parts = New Collection

    For Each Key In Record.Keys
        V = SQLFormat(Record(Key))
        Parts.Add V
    Next Key

    BuildValuesList = Join(ToArray(Parts), ", ")
End Function

Private Function BuildAssignmentsList(Record As Scripting.Dictionary) As String
    Dim Key As Variant, Parts As Collection
    Set Parts = New Collection

    For Each Key In Record.Keys
        Parts.Add Key & " = " & SQLFormat(Record(Key))
    Next Key

    BuildAssignmentsList = Join(ToArray(Parts), ", ")
End Function

Private Function SQLFormat(V As Variant) As String
    If IsNull(V) Then
        SQLFormat = "NULL"
    ElseIf VarType(V) = vbString Then
        SQLFormat = "'" & Replace(V, "'", "''") & "'"
    ElseIf VarType(V) = vbDate Then
        SQLFormat = "'" & Format(V, "yyyy-mm-dd") & "'"
    ElseIf VarType(V) = vbBoolean Then
        SQLFormat = IIf(V, "1", "0")
    Else
        SQLFormat = CStr(V)
    End If
End Function

Private Function ToArray(Coll As Collection) As Variant()
    Dim Arr() As Variant
    Dim iIndex As Long

    ReDim Arr(0 To Coll.Count - 1)

    For iIndex = 1 To Coll.Count
        Arr(iIndex - 1) = Coll(iIndex)
    Next iIndex

    ToArray = Arr
End Function
Private Function ArrayRowToDict(DataArray As Variant, RowIndex As Long, Headers() As String) As Scripting.Dictionary
    Dim RowDict As New Scripting.Dictionary
    Dim ColIndex As Long

    For ColIndex = LBound(Headers) To UBound(Headers)
        RowDict.Add Trim(Headers(ColIndex)), DataArray(ColIndex, RowIndex)
    Next ColIndex

    Set ArrayRowToDict = RowDict
End Function
