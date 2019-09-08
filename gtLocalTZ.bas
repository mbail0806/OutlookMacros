Attribute VB_Name = "gtLocalTZ"

''''''''''''''''''''''''''''''''''''''''''''''''''
' Main: get_TZLtr()
' Description: Get the local time zone and convert it to a letter and a possible modifier.
' Author: Mike Bail <bail@infionline.net>
' Version: 1.0
' Build: 1
' Date: 2015-12-07
' Contains: None
' Dependancy: None
' Notes:
' ToDo:
''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
' Modified from the following code samples
' Posted by Ruddles of Glevum Castra, Britannia at:
'      http://www.mrexcel.com/forum/excel-questions/500495-can-visual-basic-applications-code-detect-time-zone.html
' See also Pearson Software Consulting
'      http://www.cpearson.com/excel/TimeZoneAndDaylightTime.aspx

    Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
    End Type


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' NOTE: If you are using the Windows WinAPI Viewer Add-In to get
    ' function declarations, not (sic) that there is an error in the
    ' TIME_ZONE_INFORMATION structure. It defines StandardName and
    ' DaylightName As 32. This is fine if you have an Option Base
    ' directive to set the lower bound of arrays to 1. However, if
    ' your Option Base directive is set to 0 or you have no
    ' Option Base diretive, the code won't work. Instead,
    ' change the (32) to (0 To 31).
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName(0 To 31) As Integer
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName(0 To 31) As Integer
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
    End Type


    ''''''''''''''''''''''''''''''''''''''''''''''
    ' These give symbolic names to the time zone
    ' values returned by GetTimeZoneInformation .
    ''''''''''''''''''''''''''''''''''''''''''''''

    Private Enum TIME_ZONE
        TIME_ZONE_ID_INVALID = 0        ' Cannot determine DST
        TIME_ZONE_STANDARD = 1          ' Standard Time, not Daylight
        TIME_ZONE_DAYLIGHT = 2          ' Daylight Time, not Standard
    End Enum


    Private Declare Function GetTimeZoneInformation Lib "kernel32" _
        (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

    Private Declare Sub GetSystemTime Lib "kernel32" _
        (lpSystemTime As SYSTEMTIME)



Function get_TZLtr() As String
    Dim gtl_TZMod As String
    Dim gtl_TZI As TIME_ZONE_INFORMATION 'Time Zone structure
    Dim gtl_DST As TIME_ZONE 'DST state indicator (0=>Unknown 1=>Standard 2=>Daylight)
    Dim gtl_TZOffset As Integer
    Dim gtl_TZOffsetR As Double

    'Query the OS
    gtl_DST = GetTimeZoneInformation(gtl_TZI)


    Select Case gtl_DST
        Case 0, 1
            gtl_TZOffset = ((-gtl_TZI.Bias - gtl_TZI.StandardBias) / 60)
            gtl_TZOffsetR = ((-gtl_TZI.Bias - gtl_TZI.StandardBias) / 60)
        Case 2
            gtl_TZOffset = ((-gtl_TZI.Bias - gtl_TZI.DaylightBias) / 60)
            gtl_TZOffsetR = ((-gtl_TZI.Bias - gtl_TZI.DaylightBias) / 60)
    End Select

    Debug.Print gtl_DST
    Debug.Print gtl_TZI.Bias
    Debug.Print gtl_TZI.DaylightBias
    Debug.Print gtl_TZI.StandardBias
    
    
    If gtl_TZOffset < gtl_TZOffsetR Then
        gtl_TZMod = "#"
    Else
        gtl_TZMod = ""
    End If


    Select Case gtl_TZOffset
        Case 0
            get_TZLtr = "Z"
        Case 1 To 9
            get_TZLtr = Chr(gtl_TZOffset + 64) & gtl_TZMod
        Case 10 To 11
            get_TZLtr = Chr(gtl_TZOffset + 65) & gtl_TZMod
        Case 12
            get_TZLtr = "M" & gtl_TZMod
        Case 13 To 14
            get_TZLtr = "M#"
        Case -11 To -1
            get_TZLtr = Chr((gtl_TZOffset * -1) + 77) & gtl_TZMod
        Case -12
            get_TZLtr = "Y" & gtl_TZMod
        Case Else
            get_TZLtr = "-" & gtl_TZMod
    End Select

End Function
