Attribute VB_Name = "modUtilities"
Option Explicit

Public Function NullToDefault(Value As Variant, Default As Variant) As Variant
    If IsNull(Value) Then
        NullToDefault = Default
    Else
        NullToDefault = Value
    End If
End Function
' ================================================================
' ProcessingTime conversion utilities
' Converts between the string values used on Excel forms and
' the ProcessingTimeEnum used throughout the domain layer.
' ================================================================

Public Function ProcessingTimeToEnum(V As String) As ProcessingTimeEnum
    Select Case LCase$(Trim$(V))
        Case "extended time": ProcessingTimeToEnum = ProcessingTimeEnum.ExtendedTime
        Case "next day", "next day rush": ProcessingTimeToEnum = ProcessingTimeEnum.NextDay
        Case "time limited", "timelimited": ProcessingTimeToEnum = ProcessingTimeEnum.TimeLimited
        Case "same day rush", "samedayrush": ProcessingTimeToEnum = ProcessingTimeEnum.SameDayRush
        Case "call in rush", "callinrush": ProcessingTimeToEnum = ProcessingTimeEnum.CallInRush
        Case "two days", "2 days", "2days": ProcessingTimeToEnum = ProcessingTimeEnum.TwoDays
        Case "three days", "3 days", "3days": ProcessingTimeToEnum = ProcessingTimeEnum.ThreeDays
        Case "five days", "5 days", "5days": ProcessingTimeToEnum = ProcessingTimeEnum.FiveDays
        Case Else
            Err.Raise vbObjectError + 500, "modUtilities.ProcessingTimeToEnum", _
                      "Unrecognised processing time string: '" & V & "'"
    End Select
End Function

Public Function ProcessingTimeToString(V As ProcessingTimeEnum) As String
    Select Case V
        Case ProcessingTimeEnum.ExtendedTime: ProcessingTimeToString = "Extended Time"
        Case ProcessingTimeEnum.NextDay: ProcessingTimeToString = "Next Day"
        Case ProcessingTimeEnum.TimeLimited: ProcessingTimeToString = "Time Limited"
        Case ProcessingTimeEnum.SameDayRush: ProcessingTimeToString = "Same Day Rush"
        Case ProcessingTimeEnum.CallInRush: ProcessingTimeToString = "Call In Rush"
        Case ProcessingTimeEnum.TwoDays: ProcessingTimeToString = "Two Days"
        Case ProcessingTimeEnum.ThreeDays: ProcessingTimeToString = "Three Days"
        Case ProcessingTimeEnum.FiveDays: ProcessingTimeToString = "Five Days"
        Case Else
            Err.Raise vbObjectError + 501, "modUtilities.ProcessingTimeToString", _
                      "Unrecognised ProcessingTimeEnum value: " & V
    End Select
End Function

