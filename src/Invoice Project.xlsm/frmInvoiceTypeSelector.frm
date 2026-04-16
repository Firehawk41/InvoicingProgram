VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInvoiceTypeSelector 
   Caption         =   "Invoicing Macro"
   ClientHeight    =   6135
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5565
   OleObjectBlob   =   "frmInvoiceTypeSelector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInvoiceTypeSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
' MacroName: frmInvoiceTypeSelector
' Version: 2.1.0
' Author: JT
' Created: ?
' LastModified: 2026-01-28
' Description: Form for users to choose which type of invoice to generate
' DependsOn: clsLoggingSystem
' ChangeLog:
'   - 2.0.1 - 2025-06-05 - Added metadata
'   - 2.1.0 - 2025-11-18 - Added output type
'   - 3.0.0 - 2026-01-28 - Refactored whole form in order to work with CSV output - this is a placeholder while we support legacy code for word output - this will be further refactored once word is no longer supported
Private m_Logger As clsLoggingSystem
Private m_Cancelled As Boolean
Private m_SelectedFilePath As String
' =================
' Public Properties
' =================
Public Property Get InvoiceType() As InvoiceTypeEnum
    If btnRegular.Value = True Then
        InvoiceType = InvoiceTypeEnum.DefaultIndividual
    ElseIf btnWeekly.Value = True Then
        InvoiceType = InvoiceTypeEnum.FujiSummary
    ElseIf btnMonthly.Value = True Then
        InvoiceType = InvoiceTypeEnum.DowSummary
    End If
End Property
Public Property Get InvoiceSelection() As InvoiceSelectionEnum
    If btnIndividualMode.Value = True Then
        InvoiceSelection = InvoiceSelectionEnum.Individual
    ElseIf btnBatchMode.Value = True Then
        InvoiceSelection = InvoiceSelectionEnum.Batch
    End If
End Property
Public Property Get OutputType() As InvoiceOutputEnum
    If Me.btnOutputToWord.Value = True Then
        OutputType = InvoiceOutputEnum.Word ' legacy default
    ElseIf Me.btnOutputToCSV.Value = True Then
        OutputType = InvoiceOutputEnum.CSV
    Else
        Call m_Logger.LogMessage("frmInvoiceTypeSelector.OutputType", LogLevelEnum.LogWarning, "Output not recognized")
    End If
End Property
Public Property Get StartDate() As Date
    ' Only relevant for summary invoices
    If btnWeekly.Value Then
        StartDate = GetStartOfWeek(CInt(cboDate.Value))
    ElseIf btnMonthly.Value Then
        StartDate = GetFirstDayOfMonth(cboDate.Value)
    End If
End Property


' Returns the file path selected by the user (only relevant for individual invoices)
Public Property Get SelectedFilePath() As Variant
    SelectedFilePath = m_SelectedFilePath
End Property
Public Property Get Cancelled() As Boolean
    Cancelled = m_Cancelled
End Property


' ==============
' Initialization
' ==============
Public Sub Initialize(Logger As clsLoggingSystem)
    Call Logger.LogMessage("Initialize", LogLevelEnum.LogDebug, "Initializing frmInvoiceTypeSelector")

    Set m_Logger = Logger

    Me.btnRegular = True
    Me.btnBatchMode = True

    Me.btnOutputToCSV = True
    Call m_Logger.LogMessage("Initialize", LogLevelEnum.LogInfo, "Invoice output formatting is legacy; NetSuite will handle final output.")


    Call m_Logger.LogMessage("Initialize", LogLevelEnum.LogDebug, "Userform initialization complete")
End Sub

' ============
' Form Actions
' ============
Private Sub btnRegular_Click()
    btnInvoiceType_Click Me.btnRegular
End Sub
Private Sub btnMonthly_Click()
    btnInvoiceType_Click Me.btnMonthly
End Sub
Private Sub btnWeekly_Click()
    btnInvoiceType_Click Me.btnWeekly
End Sub

Private Sub btnInvoiceType_Click(Sender As Object)

    Call m_Logger.LogMessage("btnInvoiceType_Click", LogLevelEnum.LogDebug, "Invoice type changed: " & Sender.Caption)

    Select Case Sender.Name
        Case "btnRegular"

            btnIndividualMode.Value = True
            EnforceOutputRules False   ' Individual can do Word or CSV

        Case "btnWeekly"

            PopulateDateComboBox True
            EnforceOutputRules True    ' Batch only ? CSV only

        Case "btnMonthly"

            PopulateDateComboBox False
            EnforceOutputRules True    ' Batch only ? CSV only
    End Select


End Sub
Private Sub btnOpenFile_Click()
    Call m_Logger.LogMessage("btnOpenFile_Click", LogLevelEnum.LogDebug, "Opening file from user input.")
    '    cboErrorFile.Enabled = False
End Sub

Private Sub CancelButton_Click()
    m_Cancelled = True
    Me.Hide
End Sub
Private Sub cboDate_Change()
    If btnWeekly.Value And Not cboDate.Value = "" Then
        Dim selectedWeek As Integer
        selectedWeek = CInt(cboDate.Value)
        lblWeeklyDateRange.Caption = "Week " & selectedWeek & ": " & Format(GetStartOfWeek(selectedWeek), "mmm dd") & " - " & Format(GetStartOfWeek(selectedWeek) + 6, "mmm dd")
        lblWeeklyDateRange.Visible = True
    End If
End Sub

Private Sub OKButton_Click()
    Call m_Logger.LogMessage("OKButton_Click", LogLevelEnum.LogDebug, "OK button clicked. Processing invoice.")

    ' Handle based on the selected invoice type
    Select Case True
        Case btnRegular.Value
            If btnIndividualMode.Value Then
                ' Let the user select a file for individual invoices
                m_SelectedFilePath = Application.GetOpenFilename(FileFilter:="Excel files (*.xlsx), *.xlsx", MultiSelect:=False)
                If m_SelectedFilePath = "False.xlsx" Then
                    Call m_Logger.LogMessage("OKButton_Click", LogLevelEnum.LogInfo, "User cancelled file selection")
                    m_Cancelled = True ' Cancel if no file selected
                End If

            ElseIf btnBatchMode.Value Then
                Dim fd As FileDialog
                Set fd = Application.FileDialog(msoFileDialogFolderPicker)

                With fd
                    .Title = "Select folder containing reports to invoice"
                    .AllowMultiSelect = False

                    If .Show = False Then
                        Call m_Logger.LogMessage("OKButton_Click", LogLevelEnum.LogInfo, "User cancelled folder selection")
                        m_Cancelled = True
                    Else
                        m_SelectedFilePath = .SelectedItems(1)  ' <-- THIS is the folder
                        Call m_Logger.LogMessage("OKButton_Click", LogLevelEnum.LogInfo, _
                            "Folder selected: " & m_SelectedFilePath)
                    End If
                End With
            End If
            ' Normalize filepath
            If Right(m_SelectedFilePath, 1) <> "\" Then
                m_SelectedFilePath = m_SelectedFilePath & "\"
            End If
            Me.Hide
        Case btnWeekly.Value Or btnMonthly.Value
            ' Check if customer or date are empty
            If Len(cboDate.Value) = 0 Then
                Call m_Logger.LogMessage("OKButton_Click", LogLevelEnum.LogWarning, "Invoice cannot be processed: Missing date.", , True)
            Else
                Call m_Logger.LogMessage("OKButton_Click", LogLevelEnum.LogInfo, "Invoice processed with date: " & cboDate.Value)
                Me.Hide
            End If

        Case Else
            ' Default case
            Call m_Logger.LogMessage("OKButton_Click", LogLevelEnum.LogWarning, "Invoice processing cancelled or invalid.", , True)
    End Select
End Sub
' ===============
' Private Helpers
' ===============
Private Sub EnforceOutputRules(IsBatch As Boolean)
    If IsBatch Then
        btnOutputToWord.Enabled = False
        btnOutputToCSV.Enabled = True
        btnOutputToCSV.Value = True
    Else
        btnOutputToWord.Enabled = True
        btnOutputToCSV.Enabled = True
    End If
End Sub

Private Function GetStartOfWeek(WeekNum As Integer) As Date
    Dim YearStart As Date
    YearStart = DateSerial(Year(Now), 1, 1)
    GetStartOfWeek = DateAdd("ww", WeekNum - 1, YearStart)
    If Weekday(GetStartOfWeek, vbMonday) <> 1 Then
        GetStartOfWeek = DateAdd("d", -Weekday(GetStartOfWeek, vbMonday) + 1, GetStartOfWeek)
    End If
End Function
Private Function GetFirstDayOfMonth(monthAbbreviation As String) As Date
    GetFirstDayOfMonth = DateSerial(Year(Now), Month(DateValue(monthAbbreviation & " 1")), 1)
End Function
Private Sub PopulateDateComboBox(isWeekly As Boolean)
    Const LIST_OF_MONTHS As String = "Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec"

    cboDate.Clear


    If isWeekly Then

        PopulateComboBox cboDate, SetWeekList()
        If cboDate.ListCount > 0 Then
            cboDate.Value = cboDate.List(0)
        End If
        lblWeeklyDateRange.Visible = True
    Else
        PopulateComboBox cboDate, Split(LIST_OF_MONTHS, ",")
        cboDate.Value = Format(DateAdd("m", -1, Date), "mmm")
        lblWeeklyDateRange.Visible = False
    End If
End Sub

Private Sub PopulateComboBox(cbo As Object, List As Variant)
    Dim iIndex As Long

    cbo.Clear
    For iIndex = LBound(List) To UBound(List)
        cbo.AddItem List(iIndex)
    Next iIndex
End Sub



Private Function SetWeekList() As Variant
    Dim WeekSelectArr(1 To 10) As Integer
    Dim aIndex As Integer
    Dim CurrentWeek As Integer

    CurrentWeek = GetWeekNumber(Now)

    For aIndex = 1 To UBound(WeekSelectArr)
        WeekSelectArr(aIndex) = IIf(CurrentWeek - aIndex < 1, CurrentWeek - aIndex + 52, CurrentWeek - aIndex)
    Next aIndex

    ' Return the array to the calling sub
    SetWeekList = WeekSelectArr
End Function
Private Function GetWeekNumber(InputDate As Date) As Integer
    Dim Jan1 As Date
    Dim WeekOneMonday As Date
    Dim WeekNum As Integer

    ' Determine January 1 for the given year.
    Jan1 = DateSerial(Year(InputDate), 1, 1)

    ' If January 1 is Monday-Friday, then the week that contains Jan 1 is week one.
    ' That week starts on the Monday before (or on) Jan 1.
    ' Otherwise, if Jan 1 is Saturday or Sunday, the first Monday-Friday week is the one
    ' beginning on the following Monday.
    If Weekday(Jan1, vbMonday) <= 5 Then
        ' For a weekday, go backward to the Monday of that week.
        WeekOneMonday = DateAdd("d", -(Weekday(Jan1, vbMonday) - 1), Jan1)
    Else
        ' For a weekend day, go forward to the next Monday.
        WeekOneMonday = DateAdd("d", 8 - Weekday(Jan1, vbMonday), Jan1)
    End If

    ' Calculate the week number by counting the number of weeks between the starting Monday
    ' and the InputDate. The +1 ensures the starting week is counted as week 1.
    WeekNum = DateDiff("ww", WeekOneMonday, InputDate, vbMonday) + 1

    GetWeekNumber = WeekNum
End Function

