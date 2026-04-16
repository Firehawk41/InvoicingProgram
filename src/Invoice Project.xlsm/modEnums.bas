Attribute VB_Name = "modEnums"
Option Explicit
' EnumModule
' MacroName: modEnums
' Version: 1.2.0
' Author: JT
' Created: 2026-01-22
' LastModified: 2026-01-22
' Description: Store public enums in one place
' DependsOn:
' ChangeLog:
'   - 1.0.6 - 2025-05-21 added metadata JT
'   - 1.1.0 - 2025-11-18 Added InvoiceOutputEnum
'   - 1.2.0 - 2026-01-22 Added InvoiceSelectionEnum
'   - 1.3.0 - 2026-03-25 Added more processing time options
Public Enum InvoiceOutputEnum
    Word = 1
    CSV
End Enum
Public Enum InvoiceTypeEnum
    DefaultIndividual = 1
    FujiSummary
    DowSummary
    DailyCSV ' <- left for legacy reasons: really obsolete 1/22/26 JT
End Enum
Public Enum InvoiceSelectionEnum
    Individual = 1
    Batch
End Enum
Public Enum LineItemColumnTypeEnum
    LogInDateColumn = 1
    TestingRequestColumn
    QuantityColumn
    DescriptionColumn
    TurnAroundTimeColumn
    UnitPriceColumn
    ExtendedPriceColumn
    RequestedByColumn
End Enum
Public Enum LogLevelEnum
    LogDebug = 1
    LogInfo
    LogWarning
    LogError
End Enum
Public Enum TestingRequestTypeEnum
    RequestType_Min = 1
    Chemical = 1
    Water
    Wafer
    RequestType_Max = 3
End Enum
Public Enum EmailTypeEnum
    ResultsMain = 1
    ResultsCC
    InvoiceMain
    InvoiceCC
End Enum
Public Enum ProcessingTimeEnum
    ProcessingTime_Min = 1
    ExtendedTime = 1
    NextDay
    TimeLimited
    SameDayRush
    CallInRush
    TwoDays
    ThreeDays
    FiveDays
    ProcessingTime_Max = 8
End Enum
Public Enum PaymentTypeEnum
    PONumber = 1
    CreditCard
End Enum
Public Enum PythonModeEnum
    ModePrototype = 0
    ModeProduction
End Enum
