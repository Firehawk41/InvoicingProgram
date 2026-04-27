Attribute VB_Name = "modInvoiceSystem"
Option Explicit
Option Compare Text
' MacroName: modInvoiceSystem
' Executable: True
' Author: JT
' Created: 2022-09-08
' Description: Automatically generates an invoice
' DependsOn: clsLoggingSystem, clsAccessDatabase, frmInvoiceTypeSelector, modEnums, clsStalenessChecker, clsInvoiceSubmissionManager, clsInvoiceSalesOrderManager, clsQuoteLoader, clsInvoicePricingCache, clsInvoicePricingCache, clsInvoicePricingEngine, clsInvoiceSalesOrderBuilder, clsInvoiceLineItemBuilder, IInvoiceWriter, clsInvoiceSalesOrder
' ChangeLog:
'   - 5.0.1 - 2025-06-05 - Added metadata
'   - 5.1.0 - 2025-11-07 - Added boolean to determine whether the invoices are exported to csv or to word
'   - 5.2.0 - 2025-11-18 - Refactored to integrate output choice with the userform
'   - 6.0.0 - 2026-03-04 - Refactored to remove Word support to simplify and clean codebase

Const INVOICE_DB_PATH As String = "\\PRECILAB-SERVER\LabPlusServer\Documents_In_Works\Thomson\Sample Login.accdb"
Const OUTPUT_PATH As String = "\\PRECILAB-SERVER\LabPlusServer\Macros\Invoicing Macro\Saved Invoices\"
Public Sub CreateInvoice()
    Call DisplayInvoiceUF(True)
End Sub
Private Sub DisplayInvoiceUF(DebugMode As Boolean)

    Dim Logger As clsLoggingSystem
    Dim InvoiceUserForm As frmInvoiceTypeSelector
    Dim StartDate As Date
    Dim EndDate As Date
    Dim SelectedFilePath As String
    Dim ErrorMessage As String
    Dim InvoiceType As InvoiceTypeEnum
    Dim InvoiceOutput As InvoiceOutputEnum
    Dim InvoiceSelection As InvoiceSelectionEnum
    Dim Checker As New clsStalenessChecker

    ' Initiate the log file
    Set Logger = New clsLoggingSystem
    Call Logger.Initialize("CreateInvoice", DebugMode)

    ' Authenticate the workbook
    Call Checker.Initialize(ThisWorkbook.Name, ThisWorkbook.FullName, FileDateTime(ThisWorkbook.FullName), Logger)
    If Not Checker.IsCurrent Then Err.Raise 1984, "DisplayInvoiceUF", Checker.IsObsoleteMessage

    ' Set variables depending
    If Not Logger.DebugMode Then
        On Error GoTo ErrorHandler
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
    End If

    ' Set and display the userform
    Set InvoiceUserForm = New frmInvoiceTypeSelector
    Call InvoiceUserForm.Initialize(Logger)
    InvoiceUserForm.Show

    ' Check whether the user cancelled the userform
    If InvoiceUserForm.Cancelled Then
        GoTo CleanUp
    End If

    ' Extract information from the userform
    StartDate = InvoiceUserForm.StartDate
    EndDate = InvoiceUserForm.EndDate
    SelectedFilePath = InvoiceUserForm.SelectedFilePath
    InvoiceOutput = InvoiceUserForm.OutputType
    InvoiceSelection = InvoiceUserForm.InvoiceSelection
    InvoiceType = InvoiceUserForm.InvoiceType
    ' *** Refactor in future ***
    '    If InvoiceOutput = CSV Then
    '        InvoiceType = DailyCSV
    '    Else
    '        InvoiceType = InvoiceUserForm.InvoiceType
    '    End If

    ' Close the userform
    Set InvoiceUserForm = Nothing

    Call GenerateSalesOrders(StartDate, EndDate, SelectedFilePath, Logger, InvoiceSelection)

CleanUp:

    Call Logger.LogMessage("DisplayInvoiceUF", LogLevelEnum.LogInfo, "Completed sub and entering clean up.")

    ' Close the userform
    If Not InvoiceUserForm Is Nothing Then
        Set InvoiceUserForm = Nothing
    End If

    ' Close the log file
    If Not Logger Is Nothing Then
        Logger.CloseLogFile
        Set Logger = Nothing
    End If

    ' Reset settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    Exit Sub
ErrorHandler:
    Call Logger.LogError("Create Invoice", Err.Number, Err.Description)
    GoTo CleanUp

End Sub

Public Function GenerateSalesOrders(StartDate As Date, EndDate As Date, FilePath As String, Logger As clsLoggingSystem, Optional InvoiceSelection As InvoiceSelectionEnum = InvoiceSelectionEnum.Batch) As Boolean

    Dim AccessDB As clsAccessDatabase
    Dim SubmissionManager As clsInvoiceSubmissionManager
    Dim SalesOrderManager As clsInvoiceSalesOrderManager
    Dim Writer As clsInvoiceWriterCSV
    Dim SalesOrder As clsInvoiceSalesOrder
    Dim StartTime As Double
    Dim EndTime As Double

    ' Flag the process as unsuccessful by default
    GenerateSalesOrders = False

    If Not Logger.DebugMode Then On Error GoTo ErrorHandler

    StartTime = Timer

    ' Build infrastructure once
    Set AccessDB = New clsAccessDatabase
    AccessDB.Initialize INVOICE_DB_PATH, Logger

    Set SubmissionManager = BuildSubmissionManager(AccessDB, Logger)

    Select Case InvoiceSelection
        Case InvoiceSelectionEnum.Individual
            Call SubmissionManager.LoadSingle(FilePath)
        Case InvoiceSelectionEnum.Batch
            Call SubmissionManager.LoadByDateRange(FilePath, StartDate, EndDate)
        Case Else
            Err.Raise 911, , "Unknown invoice selection enum."
    End Select

    ' Build sales orders
    Set SalesOrderManager = BuildSalesOrderManager(AccessDB, Logger)
    Call SalesOrderManager.BuildFromSubmissions(SubmissionManager.Submissions)

    ' Write CSV
    Set Writer = New clsInvoiceWriterCSV
    Call Writer.Initialize(Logger)
    Call Writer.BeginOutput

    For Each SalesOrder In SalesOrderManager.SalesOrders
        Call Writer.WriteInvoice(SalesOrder)
    Next SalesOrder

    If Not Logger.DebugMode Then
        Call Writer.SaveInvoice(OUTPUT_PATH & "PRECILAB Invoice CSV" & Format(Now, "yyyymmdd hh") & "h" & Format(Now, "nn") & "m" & Format(Now, "ss") & "s")
        Call Writer.CloseInvoice
    End If


    ' Flag the process as success
    GenerateSalesOrders = True

CleanUp:
    EndTime = Timer
    Call Logger.LogMessage("CreateInvoice", LogLevelEnum.LogInfo, "Execution time: " & Round(EndTime - StartTime, 1) & " seconds.")

    ' Terminate objects
    Set SubmissionManager = Nothing
    Set SalesOrderManager = Nothing
    Set AccessDB = Nothing

    Exit Function
ErrorHandler:
    Call Logger.LogError("modInvoiceSystem.GenerateInvoiceFromFile", Err.Number, Err.Description, False)
    GoTo CleanUp
End Function
Private Function BuildSubmissionManager(AccessDB As clsAccessDatabase, Logger As clsLoggingSystem) As clsInvoiceSubmissionManager
    Dim CustomerSvc As New clsCustomerService
    Dim ChemicalSvc As New clsChemicalService
    Dim AnalysisSvc As New clsAnalysisService
    Dim ElementSvc As New clsElementService
    Dim Resolver As New clsTRFormInputResolver
    Dim SubmissionManager As New clsInvoiceSubmissionManager

    ' Build domain services once — shared across all consumers
    CustomerSvc.Initialize AccessDB, Logger
    ChemicalSvc.Initialize AccessDB, Logger
    AnalysisSvc.Initialize AccessDB, Logger
    ElementSvc.Initialize AccessDB, Logger

    ' Build resolver — needs services
    Resolver.Initialize CustomerSvc, ChemicalSvc, ElementSvc, Logger

    ' Inject everything into manager
    SubmissionManager.Initialize AccessDB, CustomerSvc, ChemicalSvc, AnalysisSvc, ElementSvc, Resolver, Logger

    Set BuildSubmissionManager = SubmissionManager
End Function
Private Function BuildSalesOrderManager(AccessDB As clsAccessDatabase, Logger As clsLoggingSystem) As clsInvoiceSalesOrderManager
    Dim QuoteLoader As clsQuoteLoader
    Dim PricingCache As clsInvoicePricingCache
    Dim PricingEngine As clsInvoicePricingEngine
    Dim LineItemBuilder As clsInvoiceLineItemBuilder
    Dim SalesOrderBuilder As clsInvoiceSalesOrderBuilder
    Dim SalesOrderManager As clsInvoiceSalesOrderManager

    Set QuoteLoader = New clsQuoteLoader
    QuoteLoader.Initialize AccessDB, Logger

    Set PricingCache = New clsInvoicePricingCache
    PricingCache.Initialize QuoteLoader, Logger

    Set PricingEngine = New clsInvoicePricingEngine
    PricingEngine.Initialize PricingCache, Logger

    Set LineItemBuilder = New clsInvoiceLineItemBuilder
    LineItemBuilder.Initialize PricingEngine, Logger

    Set SalesOrderBuilder = New clsInvoiceSalesOrderBuilder
    SalesOrderBuilder.Initialize LineItemBuilder, Logger

    Set SalesOrderManager = New clsInvoiceSalesOrderManager
    SalesOrderManager.Initialize SalesOrderBuilder, Logger

    Set BuildSalesOrderManager = SalesOrderManager
End Function
Private Sub ValidateSubmissions(SubmissionManager As clsInvoiceSubmissionManager)

    If SubmissionManager.Submissions.Count = 0 Or SubmissionManager.Submissions Is Nothing Then
        Err.Raise 1001, "modInvoiceSystem.ValidateSubmissions", "No submissions found"
    ElseIf SubmissionManager.Submissions(1) Is Nothing Then
        Err.Raise 1001, "modInvoiceSystem.ValidateSubmissions", "No submissions found"
    End If

End Sub
