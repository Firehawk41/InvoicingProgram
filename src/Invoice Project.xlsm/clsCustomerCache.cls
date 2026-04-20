VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCustomerCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsCustomerCache
' Author: JT
' Created: 2026-04-10
' Description: Materialised in-memory cache of all customers.
'              No dependency on clsCustomerService.
' DependsOn: clsLoggingSystem, clsCustomerRepository, clsCustomer
' ChangeLog:
'   - 1.0.0 - 2026-04-10 - Initial release (replaces v0.1.0 AccessDB-direct version)
'
' ================================================================
' LAYER: Cache — depends on Repo only. No upward coupling to Service.
'
' LOOKUP CONTRACT:
'   GetByID / GetByName return Nothing if not found.
'   GetIDByName / GetNameByID raise if not found (use Exists* to guard).
' ================================================================

Private m_Logger As clsLoggingSystem
Private m_Repo   As clsCustomerRepository
Private m_ByID   As Scripting.Dictionary   ' Long      -> clsCustomer
Private m_ByName As Scripting.Dictionary   ' LCase key -> clsCustomer

' ----------------------------------------------------------------
' Initialization
' ----------------------------------------------------------------
Public Sub Initialize(Repo As clsCustomerRepository, Logger As clsLoggingSystem)
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
' ----------------------------------------------------------------
Public Sub Build()
    AssertInit "Build"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    m_ByID.RemoveAll
    m_ByName.RemoveAll

    Dim Customers As Variant
    Customers = m_Repo.SelectAll

    ' No customers in DB is valid — leave the cache empty.
    If IsEmpty(Customers) Then Exit Sub

    Dim i As Long
    Dim D As Scripting.Dictionary
    Dim C As clsCustomer

    ' SelectAll column order: 0=ID, 1=name, 2=street_address,
    '                         3=city, 4=postal_code, 5=country
    For i = LBound(Customers, 2) To UBound(Customers, 2)
        Set D = New Scripting.Dictionary
        D.Add "customer_id", NullToDefault(Customers(0, i), 0)
        D.Add "customer_name", NullToDefault(Customers(1, i), "")
        D.Add "street_address", NullToDefault(Customers(2, i), "")
        D.Add "city", NullToDefault(Customers(3, i), "")
        D.Add "postal_code", NullToDefault(Customers(4, i), "")
        D.Add "country", NullToDefault(Customers(5, i), "")

        Set C = New clsCustomer
        C.Initialize m_Logger
        C.LoadFromDict D

        StoreInIndex C
    Next i

    Exit Sub
EH:                             HandleError "Build", "Failed at row " & i
End Sub

' ----------------------------------------------------------------
' Incremental updates  (called by Service after each write)
' ----------------------------------------------------------------
Public Sub Upsert(C As clsCustomer)
    AssertInit "Upsert"
    Debug.Assert Not C Is Nothing
    Debug.Assert C.ID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    ' Remove stale name key if the customer was renamed.
    If m_ByID.Exists(C.ID) Then
        Dim OldKey As String
        OldKey = NormName(m_ByID(C.ID).Name)
        If OldKey <> NormName(C.Name) Then
            If m_ByName.Exists(OldKey) Then m_ByName.Remove OldKey
        End If
    End If

    StoreInIndex C
    Exit Sub
EH:                             HandleError "Upsert", "Failed on ID: " & C.ID
End Sub

Public Sub Remove(CustomerID As Long)
    AssertInit "Remove"
    Debug.Assert CustomerID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    If m_ByID.Exists(CustomerID) Then
        Dim NameKey As String
        NameKey = NormName(m_ByID(CustomerID).Name)
        If m_ByName.Exists(NameKey) Then m_ByName.Remove NameKey
        m_ByID.Remove CustomerID
    End If
    Exit Sub
EH:                             HandleError "Remove", "Failed on ID: " & CustomerID
End Sub

' ----------------------------------------------------------------
' Read API
' ----------------------------------------------------------------
Public Function ExistsByID(CustomerID As Long) As Boolean
    AssertInit "ExistsByID"
    ExistsByID = m_ByID.Exists(CustomerID)
End Function

Public Function ExistsByName(Name As String) As Boolean
    AssertInit "ExistsByName"
    ExistsByName = m_ByName.Exists(NormName(Name))
End Function

Public Function GetByID(CustomerID As Long) As clsCustomer
    AssertInit "GetByID"
    If m_ByID.Exists(CustomerID) Then Set GetByID = m_ByID(CustomerID)
End Function

Public Function GetByName(Name As String) As clsCustomer
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
        Err.Raise vbObjectError + 103, "clsCustomerCache.GetIDByName", _
            "Customer not found: " & Name
    End If
End Function

Public Function GetNameByID(CustomerID As Long) As String
    AssertInit "GetNameByID"
    If m_ByID.Exists(CustomerID) Then
        GetNameByID = m_ByID(CustomerID).Name
    Else
        Err.Raise vbObjectError + 103, "clsCustomerCache.GetNameByID", _
            "Customer ID not found: " & CustomerID
    End If
End Function

Public Function AllCustomers() As Collection
    AssertInit "AllCustomers"
    Dim Coll As New Collection
    Dim Key  As Variant
    For Each Key In m_ByID.Keys
        Coll.Add m_ByID(Key)
    Next Key
    Set AllCustomers = Coll
End Function

' ----------------------------------------------------------------
' Private helpers
' ----------------------------------------------------------------
Private Sub StoreInIndex(C As clsCustomer)
    Set m_ByID(C.ID) = C
    Set m_ByName(NormName(C.Name)) = C
End Sub

Private Function NormName(V As String) As String
    NormName = LCase$(Trim$(V))
End Function
Private Function NullToDefault(Value As Variant, Default As Variant) As Variant
    If IsNull(Value) Then
        NullToDefault = Default
    Else
        NullToDefault = Value
    End If
End Function
Private Sub AssertInit(Src As String)
    If m_Logger Is Nothing Then
        Err.Raise vbObjectError + 1, "clsCustomerCache." & Src, "Not initialized"
    End If
End Sub

Private Sub HandleError(Src As String, Context As String)
    m_Logger.LogError "clsCustomerCache." & Src, Err.Description & " | " & Context
    Err.Clear
End Sub
