VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        clsDates
'
' Description:   Various date calculation routines.
'
'                The last country of Eastern Orthodox Europe to adopt the
'                Gregorian calendar was Greece on Thursday, 1-Mar-1923, which
'                followed Wednesday, 15-Feb-1923 (a change that also dropped
'                thirteen days).  In eastern Asia, the Nationalist Government
'                of China decreed that effective 1-Jan-1929 the Gregorian
'                calendar would be used.
'
' Important:      This module must have access to modTrimStr.bas
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-Jun-2005  Kenneth Ives  kenaso@tx.rr.com
' 09-Sep-2008  Kenneth Ives  kenaso@tx.rr.com
'              Fixed logic that evaluated the century in routines
'              JulianToDate() and JulianToSerial()
' 05-Oct-2008  Kenneth Ives  kenaso@tx.rr.com
'              Enhanced IsLeapYear() routine
' 10-Jun-2010  Kenneth Ives  kenaso@tx.rr.com
'              Updated IsLeapYear() routine
' 05-Nov-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added Y2K38_SecondsToDate() and Y2K38_DateToSeconds()
'              routines to handle the new Y2K known as Y2K38
' 21-Nov-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated SELECT statement to determine year validity.  See
'              IsLeapYear() routine.
' 29-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Added IsValidYear() routine.
'              - Added DayOfMonth(), EasterSunday() routines.
'              - Updated IsLeapYear() routine.
' 02-Feb-2012  Kenneth Ives  kenaso@tx.rr.com
'              Fixed bug in EasterSunday() routine.
' 11-Feb-2012  Kenneth Ives  kenaso@tx.rr.com
'              Added IsDaylightSavings(), CurrentUTC(),
'              DaylightSavingsBegins(), DaylightSavingsEnds() routines.
' 26-Mar-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Deleted RemoveTrailingNulls() routine from this module.
'              - Changed call to RemoveTrailingNulls() to TrimStr module
'                due to speed and accuracy.
' 10-May-2012  Kenneth Ives  kenaso@tx.rr.com
'              Made IsValidYear() routine global.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  ' pertain to Locale Settings
  Private Const MAX_SIZE              As Long = 260     ' buffer area
  Private Const LOCALE_SDATE          As Long = &H1D    ' date separator
  Private Const LOCALE_STIME          As Long = &H1E    ' time separator
  Private Const LOCALE_SSHORTDATE     As Long = &H1F    ' short date format string
  Private Const LOCALE_SLONGDATE      As Long = &H20    ' long date format string
  Private Const LOCALE_STIMEFORMAT    As Long = &H1003  ' time format string
  Private Const LOCALE_S1159          As Long = &H28    ' AM designator
  Private Const LOCALE_S2359          As Long = &H29    ' PM designator

  ' Clock rolls over on 1/19/2038 at 3:14:07 AM,
  ' commonly referred to as Y2K38
  Private Const BASE_DATE             As Date = #1/1/1970#
  Private Const BIT_31                As Double = 2147483648#   ' 0x80000000
  Private Const BIT_32                As Double = 4294967296#   ' 0x100000000
  
  ' Return values for API GetTimeZoneInformation()
  Private Const TIME_ZONE_ID_UNKNOWN  As Long = 0   ' Cannot determine current time zone
  Private Const TIME_ZONE_ID_STANDARD As Long = 1
  Private Const TIME_ZONE_ID_DAYLIGHT As Long = 2
  Private Const TIME_ZONE_ID_INVALID  As Long = &HFFFFFFFF   ' Failed API call

' ***************************************************************************
' Enumerations
' ***************************************************************************
  Public Enum enumDayOfWeek
      eSunday = 1
      eMonday = 2
      eTuesday = 3
      eWednesday = 4
      eThursday = 5
      eFriday = 6
      eSaturday = 7
  End Enum
  
  Public Enum enumDateInterval
      eYears      ' 0
      eMonths     ' 1
      eDays       ' 2
      eHours      ' 3
      eMinutes    ' 4
      eSeconds    ' 5
  End Enum

  Public Enum enumYear
      e1700 = 1700
      e1800 = 1800
      e1900 = 1900
      e2000 = 2000   ' Default
      e2100 = 2100
  End Enum

' ***************************************************************************
' Type Structures
' ***************************************************************************
  Private Type SYSTEMTIME
      wYear         As Integer
      wMonth        As Integer
      wDayOfWeek    As Integer
      wDay          As Integer
      wHour         As Integer
      wMinute       As Integer
      wSecond       As Integer
      wMilliseconds As Integer
  End Type

  Private Type TIME_ZONE_INFORMATION
      bias                  As Long         ' current offset to UTC
      StandardName(1 To 64) As Byte         ' as unicode string
      StandardDate          As SYSTEMTIME   ' date & time Stamdard time comes into effect
      StandardBias          As Long         ' offset to UTC during Stamdard time
      DaylightName(1 To 64) As Byte         ' as unicode string
      DaylightDate          As SYSTEMTIME   ' date & time Day light savings time comes into effect
      DaylightBias          As Long         ' offset to UTC during Day light savings time
  End Type

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' ZeroMemory is used for clearing contents of a type structure.
  Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" _
          (Destination As Any, ByVal Length As Long)

  ' GetTimeZoneInformation function retrieves the current time-zone parameters.
  ' These parameters control the translations between Coordinated Universal
  ' Time (UTC) and local time.
  Private Declare Function GetTimeZoneInformation Lib "kernel32" _
          (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

  '==========================================================================
  ' Declares for Locale Setings
  '
  ' The GetUserDefaultLCID function retrieves the user default locale
  ' identifier.
  Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long

  ' This function retrieves information about a user's locale.  The
  ' GetLocaleInfo function always retrieves information in text format.
  ' If the information is a numeric value, the function converts the
  ' number to text using decimal notation.
  Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" _
          (ByVal Locale As Long, ByVal LCType As Long, _
          ByVal lpLCData As String, ByVal cchData As Long) As Long

' ***************************************************************************
' Module Variables
'
'                    +-------------- Module level designator
'                    |  +----------- Data type (String)
'                    |  |     |----- Variable subname
'                    - --- ---------
' Naming standard:   m str TimeFormat
' Variable name:     mstrTimeFormat
'
' ***************************************************************************
  Private mstrTimeFormat      As String
  Private mstrTimeSeparator   As String
  Private mstrDateSeparator   As String
  Private mstrLongDateFormat  As String
  Private mstrShortDateFormat As String
   
   
' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

Public Property Let StopProcessing(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.StopProcessing = True
    gblnStopProcessing = blnData
    DoEvents
End Property

Public Property Get StopProcessing() As Boolean
    ' used when retrieving value of a property, on the right
    ' side of an assignment.   Syntax:  Debug.Print X.StopProcessing
    StopProcessing = gblnStopProcessing
    DoEvents
End Property

Public Property Get LongDateFormat() As String
    ' used when retrieving value of a property, on the right
    ' side of an assignment.   Syntax:  Debug.Print X.LongDateFormat
    LongDateFormat = mstrLongDateFormat
End Property

Public Property Get ShortDateFormat() As String
    ' used when retrieving value of a property, on the right
    ' side of an assignment.   Syntax:  Debug.Print X.ShortDateFormat
    ShortDateFormat = mstrShortDateFormat
End Property

Public Property Get DateSeparator() As String
    ' used when retrieving value of a property, on the right
    ' side of an assignment.   Syntax:  Debug.Print X.DateSeparator
    DateSeparator = mstrDateSeparator
End Property

Public Property Get TimeFormat() As String
    ' used when retrieving value of a property, on the right
    ' side of an assignment.   Syntax:  Debug.Print X.TimeFormat
    TimeFormat = mstrTimeFormat
End Property

Public Property Get TimeSeparator() As String
    ' used when retrieving value of a property, on the right
    ' side of an assignment.   Syntax:  Debug.Print X.TimeSeparator
    TimeSeparator = mstrTimeSeparator
End Property


' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Procedure:     CalcDate
'
' Description:   Use this function to calculate a future or previous date.
'
' Parameters:    datDate   - Base or starting date
'                lngNumber - Number of iterations.
'                            Positive number calculates a future date.
'                            Negative number calculates a date in the past.
'                lngDateInterval - [Optional] Segment to use in time spanning
'                            Default - eDays
'
' Example:       CalcDate = DateAdd("yyyy", 5, datDate)   ' 5 years in the future
'                CalcDate = DateAdd("yyyy", -5, datDate)  ' 5 years in the past
'
' Returns:       Past or future date
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-Jun-2005  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function CalcDate(ByVal datDate As Date, _
                         ByVal lngNumber As Long, _
                Optional ByVal lngDateInterval As enumDateInterval = eDays) As Date

    Dim astrSegments() As String
    
    Erase astrSegments()  ' Always start with an empty array
    ReDim astrSegments(6)
    
    astrSegments(0) = "yyyy"  ' 0 - Years
    astrSegments(1) = "m"     ' 1 - Months
    astrSegments(2) = "d"     ' 2 - Days
    astrSegments(3) = "h"     ' 3 - Hours
    astrSegments(4) = "n"     ' 4 - Minutes
    astrSegments(5) = "s"     ' 5 - Seconds
    
    CalcDate = DateAdd(astrSegments(lngDateInterval), lngNumber, datDate)
    
    Erase astrSegments()  ' Always empty arrays when not needed
    
End Function

' ***************************************************************************
' Routine:       DateToJulian
'
' Description:   This procedure takes a normal date format (that is, 1/1/94)
'                and converts it to the appropriate Julian date (yyddd).
'
'                Most government agencies and contractors require the use of
'                Julian dates. A Julian date starts with a two-digit year,
'                and then counts the number of days from January 1 inclusive.
'
'                The Julian Date is returned in string format so as to
'                display all 5 digits to include any leading zeroes. If it
'                were returned in numeric format the Julian Date might be
'                truncated.  For example the Julian Date "00001"
'                (Jan 1, 2000) would appear as 1 because all leading
'                zeroes would be dropped.
'
' Parameters:    datDate - [Optional] Date to be converted
'
' Returns:       Formatted date in string format to display all 5 digits to
'                include any leading zeroes.
'                Ex:  8/17/2007 --> "07229"
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-Aug-2007  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function DateToJulian(Optional ByVal datDate As Date = Empty) As String

    ' If default date value (no date, just time),
    ' then use system date and time
    If datDate = "12:00:00 AM" Then
        datDate = Now()
    End If
    
    DateToJulian = Format$(datDate, "yy") & _
                   Format$(DatePart("y", datDate), "000")

End Function

' ***************************************************************************
' Routine:       JulianToDate
'
' Description:   This routine takes a Julian Date and converts to an
'                acceptable date format.
'
'                Most government agencies and contractors require the use of
'                Julian dates. A Julian date starts with a two-digit year,
'                and then counts the number of days from January 1 inclusive.
'
' Parameters:    strJulianDate - Julian date in string format.
'                lngYear - [Optional] Designate the century.
'                          Min - 1700
'                          Max - 2100
'
' Returns:       Formatted date.
'                Ex: "07229" --> 08/17/2007
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 11-Oct-2006  Visual Basic Procedures for Julian Date Calculations
'              http://support.microsoft.com/kb/116281/en-us
' 17-Aug-2007  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 09-Sep-2008  Kenneth Ives  kenaso@tx.rr.com
'              Fixed logic that evaluated the century
' ***************************************************************************
Public Function JulianToDate(ByVal strJulianDate As String, _
                    Optional ByVal lngCentury As enumYear = e2000) As Date

    Dim lngJulianDate As Long
    
    Select Case lngCentury
           Case e1700, e1800, e1900, e2000, e2100
                ' Valid century
           Case Else
                lngCentury = e2000
    End Select
    
    If IsNumeric(strJulianDate) Then
        
        lngJulianDate = CLng(strJulianDate)  ' convert to numeric
        
        ' convert numeric values into short
        ' date format as defined by this locale
        JulianToDate = CDate(DateSerial(lngCentury + Int(lngJulianDate / 1000), _
                                        1, lngJulianDate Mod 1000))
    End If
    
End Function

' ***************************************************************************
' Routine:       JulianToSerial
'
' Description:   This procedure takes a Julian date (yyddd) and converts it
'                to the appropriate serial date.
'
' Parameters:    strJulianDate - Julian date in string format.
'                lngYear - [Optional] Designate the century.
'                          Min - 1700
'                          Max - 2100
'
' Returns:       Date serial number
'                Ex: "07229" --> 39311
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 11-Oct-2006  Visual Basic Procedures for Julian Date Calculations
'              http://support.microsoft.com/kb/116281/en-us
' 17-Aug-2007  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 09-Sep-2008  Kenneth Ives  kenaso@tx.rr.com
'              Fixed logic that evaluated the century
' ***************************************************************************
Public Function JulianToSerial(ByVal strJulianDate As String, _
                      Optional ByVal lngCentury As enumYear = e2000) As Long

    Dim lngJulianDate As Long
    
    Select Case lngCentury
           Case e1700, e1800, e1900, e2000, e2100
                ' Valid century
           Case Else
                lngCentury = e2000
    End Select
    
    If IsNumeric(strJulianDate) Then
        
        lngJulianDate = CLng(strJulianDate)  ' convert to numeric
        
        ' convert numeric values
        ' into date serial value
        JulianToSerial = DateSerial(lngCentury + Int(lngJulianDate / 1000), _
                                    1, lngJulianDate Mod 1000)
    End If
    
End Function

' ***************************************************************************
' Routine:       SerialToJulian
'
' Description:   This procedure takes a serial date number and converts it
'                to the appropriate Julian date and returns it as a string.
'
' Parameters:    lngSerial - Long integer representing a valid date.
'
' Returns:       Julian Date in string format
'                Ex:  39311 --> "07229"
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 11-Oct-2006  Visual Basic Procedures for Julian Date Calculations
'              http://support.microsoft.com/kb/116281/en-us
' 17-Aug-2007  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function SerialToJulian(ByVal lngSerial As Long) As String

    Dim datSerialDate As Date    ' The serial date.
    Dim strSerialYear As String  ' The year of the serial date.
    Dim strJulianDay  As String  ' Number of days in the serial date

    datSerialDate = CDate(lngSerial)

    ' Assign SerialYear the year number
    strSerialYear = Format$(lngSerial, "yy")

    ' Find the day number for SerialDate
    strJulianDay = Format$(CStr(datSerialDate - DateValue("1/1/" & strSerialYear) + 1), "000")
    SerialToJulian = strSerialYear & strJulianDay
    
End Function

' ***************************************************************************
' Routine:       SerialToDate
'
' Description:   Convert a long integer representing a valid date into
'                an accepted date format.
'
' Parameters:    lngSerial - Long integer representing a valid date.
'                strDateFmt - [Optional] Designates the return format of
'                    the date.  Default is the system short date format.
'
' Returns:       formatted date
'                Ex:  39311 --> 08/17/2007
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-Aug-2007  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function SerialToDate(ByVal lngSerial As Long, _
                    Optional ByVal strDateFmt As String = vbNullString) As Date
                    
    If Len(Trim$(strDateFmt)) = 0 Then
        strDateFmt = mstrShortDateFormat   ' Use system short date format
    End If
    
    ' convert to date format
    SerialToDate = CDate(Format$(lngSerial, strDateFmt))
    
End Function

' ***************************************************************************
' Routine:       DateToSerial
'
' Description:   This procedure takes a normal date format (11/11/08 08:30:15)
'                and converts it to a date serial number.
'
' Parameters:    datDate - Date to be evaluated
'
' Returns:       Long integer representing a date.
'                Ex:  8/17/2007 --> 39311
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-Aug-2007  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function DateToSerial(ByVal datDate As Date) As Long

    Dim intYear  As Integer
    Dim intMonth As Integer
    Dim intDay   As Integer

    ' Capture year, month, day
    ' from the date passed
    intYear = Year(datDate)
    intMonth = Month(datDate)
    intDay = Day(datDate)
    
    ' Convert to a date serial value
    DateToSerial = CLng(DateSerial(intYear, intMonth, intDay))

End Function

' ***************************************************************************
' Routine:       DateToWords
'
' Description:   This procedure takes a normal date and returns the date into
'                words.  Works great for party invitations.
'
' Reference:     Marlo F. Simangan  "Date to word conversion"
'                http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=70130&lngWId=1
'
' Parameters:    datDate - Date to be evaluated
'
' Returns:       String representing a date.
'
' Examples:      12/25/1999 -> Saturday, Twenty-Fifth day of December in the
'                              year One Thousand Nine Hundred Ninety-Nine
'
'                4/17/2008  -> Tuesday, Seventeenth day of April in the year
'                              Two Thousand Eight
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-Apr-2008  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 24-Apr-2008  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a logic bug in the year determination.
' ***************************************************************************
Public Function DateToWords(ByVal datDate As Date) As String

    Dim lngYear      As Long
    Dim lngMonth     As Long
    Dim lngDay       As Long
    Dim lngWeekDay   As Long
    Dim strYear      As String
    Dim strDayName   As String
    Dim strMonName   As String
    Dim avntOnesA    As Variant
    Dim avntOnesB    As Variant
    Dim avntTensA    As Variant
    Dim avntTensB    As Variant
    Dim avntHundreds As Variant
    
    ' Start with empty variants
    avntOnesA = Empty
    avntOnesB = Empty
    avntTensA = Empty
    avntTensB = Empty
    avntHundreds = Empty

    lngYear = Year(datDate)           ' Capture year
    lngMonth = Month(datDate)         ' Capture month
    lngDay = Day(datDate)             ' Capture day
    lngWeekDay = DayOfWeek(datDate)  ' Get numeric day of week (1-7)
    
    strDayName = GetWeekdayName(lngWeekDay, False)
    strMonName = GetMonthName(lngMonth, False)
    strYear = vbNullString
        
    ' Load the temp arrays.
    ' First element is not used.
    avntOnesA = Array(" ", "One", "Two", "Three", "Four", "Five", "Six", "Seven", _
                      "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", _
                      "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", _
                      "Nineteen")
    
    avntOnesB = Array(" ", "First", "Second", "Third", "Fourth", "Fifth", "Sixth", _
                      "Seventh", "Eighth", "Ninth", "Tenth", "Eleventh", "Twelveth", _
                      "Thirteenth", "Fourteenth", "Fifteenth", "Sixteenth", _
                      "Seventeenth", "Eighteenth", "Nineteenth")
    
    avntTensA = Array(" ", "Ten", "Twenty", "Thirty", "Forty", "Fifty", _
                      "Sixty", "Seventy", "Eighty", "Ninety")
        
    avntTensB = Array(" ", "Tenth", "Twentieth", "Thirtieth", "Fortieth", "Fiftieth", _
                      "Sixtieth", "Seventieth", "Eightieth", "Ninetieth")
        
    avntHundreds = Array(" ", "One", "Two", "Three", "Four", "Five", _
                         "Six", "Seven", "Eight", "Nine")
        
    ' Determine year format
    Select Case (lngYear Mod 1000)
    
           Case 0
                strYear = avntHundreds(Fix(lngYear / 1000)) & " Thousand"
         
           Case Is < 20
                strYear = avntHundreds(Fix(lngYear / 1000)) & " Thousand " & _
                          avntOnesA(lngYear Mod 100)
                        
           Case Is < 100
                strYear = avntHundreds(Fix(lngYear / 1000)) & " Thousand " & _
                          avntTensA(Fix((lngYear Mod 1000) / 10)) & "-" & _
                          avntOnesA(lngYear Mod 10)
                        
           Case Is < 1000
                strYear = avntHundreds(Fix(lngYear / 1000)) & " Thousand " & _
                          avntHundreds(Fix((lngYear Mod 1000) / 100)) & " Hundred " & _
                          avntTensA(Fix((lngYear Mod 100) / 10)) & "-" & _
                          avntOnesA(lngYear Mod 10)
    End Select
           
    ' Begin formatting output string
    DateToWords = strDayName & ", "
    
    Select Case lngDay
    
           Case Is < 20
                DateToWords = DateToWords & avntOnesB(lngDay)
                
           Case Else
                Select Case (lngDay Mod 10)
                       Case 0:    DateToWords = DateToWords & avntTensB(lngDay / 10)
                       Case Else: DateToWords = DateToWords & avntTensA(Int(lngDay / 10)) & _
                                                "-" & avntOnesB(lngDay Mod 10)
                End Select
    End Select
        
    ' Return completed string
    DateToWords = DateToWords & " day of " & strMonName & " in the year " & strYear

    ' Always empty variants and
    ' arrays when not needed
    avntOnesA = Empty
    avntOnesB = Empty
    avntTensA = Empty
    avntTensB = Empty
    avntHundreds = Empty

End Function

' ***************************************************************************
' Routine:       TimeToWords
'
' Description:   This procedure takes the time and returns the time into
'                words.  Works great for party invitations.
'
' Parameters:    strTime - Time to be evaluated
'
' Returns:       String representing time.
'
' Examples:      2:59 am     -> Fifty-Nine minutes past the hour of Two
'                               in the Morning
'                12:25:12 PM -> Twenty-Five minutes past the hour of Twelve
'                               in the Afternoon
'                0:00 or 00:00 or
'                24:00 or 12:00 AM -> At Midnight
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function TimeToWords(ByVal strTime As String) As String

    Dim lngTemp     As Long
    Dim lngHour     As Long
    Dim lngMinutes  As Long
    Dim lngPosition As Long
    Dim blnMorning  As Boolean
    Dim strTemp     As String
    Dim strHour     As String
    Dim strMinutes  As String
    Dim avntOnesA   As Variant
    Dim avntOnesB   As Variant
    Dim avntTens    As Variant
    
    avntOnesA = Empty   ' Start with empty variants
    avntOnesB = Empty
    avntTens = Empty
    blnMorning = True   ' preset flag

    strTime = Trim$(strTime)   ' remove leading and trailing blanks
    
    ' If not enough characters then leave
    ' (Ex:  0:00 = minimum length (hours and minutes))
    If Len(strTime) < 4 Then
        TimeToWords = "<Unknown time format>"
        Exit Function
    End If
    
    ' Find first time separator
    ' used by this locale  (Ex: ":")
    lngPosition = InStr(1, strTime, mstrTimeSeparator)
    
    ' If not found, then leave
    If lngPosition = 0 Then
        TimeToWords = "<Unknown time format>"
        Exit Function
    End If
    
    lngTemp = CLng(Mid$(strTime, 1, lngPosition - 1))   ' Capture hour to temp area
    strTime = Mid$(strTime, lngPosition + 1)            ' resize time string
    lngPosition = InStr(1, strTime, mstrTimeSeparator)  ' Find end of minutes
    
    ' See if hours, minutes and
    ' seconds were passed (Ex: 00:00:00)
    If lngPosition > 0 Then
        lngMinutes = CLng(Mid$(strTime, 1, lngPosition - 1))   ' Capture minutes
    Else
        ' Just hours and minutes passed (Ex:  09:05)
        lngMinutes = CLng(Trim$(Left$(strTime, 2)))
    End If
        
    ' Capture AM or PM indicator
    strTemp = Trim$(Right$(strTime, 2))
    
    ' if no AM or PM indicators then
    ' assume a 24-hour format is used
    If Len(Trim$(strTemp)) > 0 Then
        
        ' If PM indicator found then
        ' time is not in the morning
        If StrComp(strTemp, "PM", vbTextCompare) = 0 Then
            blnMorning = False    ' afternoon or evening
        End If
    End If

    ' Is the hour morning or evening?
    Select Case lngTemp
           
           Case 24: lngHour = 0  ' Midnight
           
           Case 13 To 23
                ' Adjust to not use
                ' 24-hour clock format
                lngHour = lngTemp - 12
                
                If lngHour <> 0 Then
                    blnMorning = False   ' Afternoon or evening
                End If
                
           Case 12
                lngHour = lngTemp
                blnMorning = False       ' Afternoon or evening
           
           Case Else: lngHour = lngTemp  ' Morning
    End Select
            
    ' Load temp arrays.
    avntOnesA = Array("Midnight", "One", "Two", "Three", "Four", "Five", "Six", "Seven", _
                      "Eight", "Nine", "Ten", "Eleven", "Twelve")
    
    avntOnesB = Array("Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", _
                      "Seventeen", "Eighteen", "Nineteen")
             
    ' First element is not used
    avntTens = Array(" ", "Ten", "Twenty", "Thirty", "Forty", "Fifty")
             
    ' Determine the minutes
    Select Case lngMinutes
           
           Case 0:      strMinutes = vbNullString
           Case 1:      strMinutes = avntOnesA(lngMinutes) & " minute past "
           Case 2 To 9: strMinutes = avntOnesA(lngMinutes) & " minutes past "
           
           Case 10 To 59
                lngTemp = Fix(lngMinutes / 10)
                
                Select Case lngTemp
                       
                       Case 1     ' 1 to 19 minutes
                            ' Are there any minutes not equal to 10
                            If (lngMinutes - (lngTemp * 10)) > 0 Then
                                strMinutes = avntOnesB((lngMinutes - (lngTemp * 10)))
                            Else
                                ' exactly 10 minutes
                                strMinutes = avntOnesB(lngTemp)
                            End If
                            
                       Case Else  ' 20 to 59 minutes
                            strMinutes = avntTens(lngTemp)
                            
                            If (lngMinutes - (lngTemp * 10)) > 0 Then
                                strMinutes = strMinutes & "-" & _
                                             avntOnesA(lngMinutes - (lngTemp * 10))
                            End If
                End Select
                            
                strMinutes = strMinutes & " minutes past "
    End Select
    
    ' Begin formatting return string
    strHour = avntOnesA(lngHour)
    
    If Len(strMinutes) = 0 Then
        strHour = "At " & strHour
    Else
        If StrComp("midnight", Right$(strHour, 8), vbTextCompare) = 0 Then
            strHour = strMinutes & strHour
        Else
            strHour = strMinutes & "the hour of " & strHour
        End If
    End If
          
    ' Finish formatting return string
    If StrComp("midnight", Right$(strHour, 8), vbTextCompare) = 0 Then
            
        TimeToWords = strHour
    
    Else
        If blnMorning Then
            TimeToWords = strHour & " in the Morning"
        Else
            ' Afternoon or evening
            Select Case lngHour
                   Case 6 To 11: TimeToWords = strHour & " in the Evening"
                   Case Else:    TimeToWords = strHour & " in the Afternoon"
            End Select
        End If
    End If
    
    avntOnesA = Empty   ' Always empty variants when not needed
    avntOnesB = Empty
    avntTens = Empty

End Function

' ***************************************************************************
' Routine:       DaysLeftInYear
'
' Description:   This procedure will return the number of days left in the
'                current year.
'
' Parameters:    datDate - Date to be evaluated
'
' Returns:       Number of days left in year
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-Jun-2005  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function DaysLeftInYear(ByVal datDate As Date) As Long
    
    DaysLeftInYear = DaysInYear(datDate) - DaysPassed(datDate)

End Function
    
' ***************************************************************************
' Routine:       DaysPassed
'
' Description:   This procedure will return the number of days that have
'                passed since January 1 of a given year.
'
' Parameters:    datDate - Date to be evaluated
'
' Returns:       Number of days
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-Jun-2005  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function DaysPassed(ByVal datDate As Date) As Long

    DaysPassed = Format$(CStr(datDate - DateValue("1/1/" & CStr(Year(datDate))) + 1), "000")

End Function

' ***************************************************************************
' Procedure:     DaysInYear
'
' Description:   Use this function to determine how many days there are in
'                a year. Leap years have 366 days and all others have 365.
'
' Parameters:    datDate - Date to be evaluated
'
' Returns:       Number of days in the requested year
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-Jun-2005  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function DaysInYear(ByVal datDate As Date) As Long

    Dim lngYear As Long
    
    lngYear = Year(datDate)
    DaysInYear = IIf(IsLeapYear(lngYear), 366, 365)

End Function

' ***************************************************************************
' Procedure:     IsLeapYear
'
' Description:   Determine if a 4-digit year is a leap year. The first leap
'                year in the modern sense was 1752, when 11 days were lost
'                from the month September with the adoption of the Gregorian
'                calendar by Britain and her colonies. After 1752 we adopted
'                the system still in use today where an additional day is
'                inserted in February in years wholly divisible by four,
'                other than years ending in "00" with the exception of those
'                divisible by 400 which are still leap years (like 2000).
'
' Parameters:    lngYear - Year being queried (prefer 4 digits; ex: 2012)
'                          If one to two-digit year is passed, year is
'                          evaluated to be either current or previous
'                          century.
'
' Returns:       TRUE - This is a leap year
'                FALSE - This is not a leap year
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-Jun-2005  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 05-Oct-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added functionality to assume current century
' 25-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Updated method to determine year validity
' ***************************************************************************
Public Function IsLeapYear(Optional ByVal lngYear As Long = 0) As Boolean
    
    If IsValidYear(lngYear) Then
        ' If month of February has 29 days then this is a leap year
        IsLeapYear = IIf(Day(DateSerial(lngYear, 2, 29)) = 29, True, False)
    Else
        ' Invalid year (three digits or letters)
        IsLeapYear = False
    End If
    
End Function

' ***************************************************************************
' Procedure:     WeekNumToDate
'
' Description:   Convert a week number to a specific date.
'
' Parameters:    lngWeekNbr - Number of the week in the above year
'                lngYear - 4-digit year to be used in the calculation
'                lngFirstDayOfWeek - Number representing the first day of
'                      the week (Default:  Sunday = 1)
'
' Returns:       Appropriate date value
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Jun-2005  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function WeekNumToDate(ByVal lngWeekNbr As Long, _
                              ByVal lngYear As Long, _
                     Optional ByVal lngFirstDayOfWeek As enumDayOfWeek = eSunday) As Date
    
    WeekNumToDate = CDate(DateSerial(lngYear, 1, 7 * (lngWeekNbr - 1) + (lngFirstDayOfWeek - 1)))

End Function

' ***************************************************************************
' Procedure:     DaysBetween
'
' Description:   This function check how many days there are between
'                two certain dates
'
' Parameters:    datStartDate - Starting date
'                datEndDate - Ending date
'                blnIncludeStartDay - [Optional] TRUE = Include the starting
'                      day in the total.  FALSE (Default) = Do not include
'                      the start day in the total.
'                lngFirstDayOfWeek - Number representing the first day of
'                      the week (Ex:  Sunday = 1)
'
' Returns:       Number of days between two dates
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-Jun-2005  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function DaysBetween(ByVal datStartDate As Date, _
                            ByVal datEndDate As Date, _
                   Optional ByVal blnIncludeStartDay As Boolean = False, _
                   Optional ByVal lngFirstDayOfWeek As enumDayOfWeek = eSunday) As Long
    
    DaysBetween = DateDiff("d", datEndDate, datStartDate, lngFirstDayOfWeek)
        
    If blnIncludeStartDay Then
        DaysBetween = DaysBetween + 1
    End If
    
End Function

' ***************************************************************************
' Procedure:     LastDayOfMonth
'
' Description:   Returns the number of days in a specified month
'
' Parameters:    datDate - date to be evaluated
'
' Returns:       Returns the number of days in a requested month
'                Ex:   29 in 2/1/2000
'                      28 in 2/1/2001
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-Jun-2005  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function LastDayOfMonth(ByVal datDate As Date) As Long

    LastDayOfMonth = Day(DateSerial(Year(datDate), Month(datDate) + 1, 0))
    
End Function

' ***************************************************************************
' Procedure:     LastDayPrevMonth
'
' Description:   Find last day of previous month
'
' Parameters:    datDate - Date to be evaluated
'
' Returns:       None
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Jun-2005  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function LastDayPrevMonth(ByVal datDate As Date) As Long
    
    LastDayPrevMonth = DateSerial(Year(datDate), Month(datDate), 0)

End Function

' ***************************************************************************
' Procedure:     GetWeekNumber
'
' Description:   Find the week of the year as a number (1 - 54).
'
' Parameters:    datDate - date to be evaluated
'
' Returns:       Week number of the date passed
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-Jun-2005  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function GetWeekNumber(ByVal datDate As Date) As Long

    GetWeekNumber = CLng(Format$(datDate, "ww"))
    
End Function

' ***************************************************************************
' Procedure:     DayOfWeek
'
' Description:   Returns the day of the week of a certain date in numeric
'                format
'
' Parameters:    datDate - Date to be evaluated
'                lngFirstDayOfWeek - Number representing first day of
'                      the week (Ex:  Sunday = 1)
'
' Returns:       day of the week as a number
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Jun-2005  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function DayOfWeek(ByVal datDate As Date, _
                 Optional ByVal lngFirstDayOfWeek As enumDayOfWeek = eSunday) As Long

    DayOfWeek = Weekday(datDate, lngFirstDayOfWeek)

End Function

' ***************************************************************************
' Procedure:     GetWeekdayName
'
' Description:   Get name of the week day
'
' Parameters:    lngDayOfWeek - Designates the day of the week you want.
'                    Valid values are 1 to 7.  (Ex: Sunday = 1)
'                blnAbbreviate - [Optional] Abbreviate the names
'                    TRUE - Shorten names to 3 chars (Default)
'                    FALSE - Use long nanes
'                lngFirstDayOfWeek = Designates which day is the beginning
'                    of the week.  Sunday = 1
'
' Returns:       Name of week day
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 31-Mar-2012  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Function GetWeekdayName(ByVal lngDayOfWeek As Long, _
                      Optional ByVal blnAbbreviate As Boolean = True, _
                      Optional ByVal lngFirstDayOfWeek As enumDayOfWeek = eSunday) As String

    Select Case lngDayOfWeek
    
           Case 1 To 7   ' Make API call to get day name
                If blnAbbreviate Then
                    GetWeekdayName = WeekdayName(lngDayOfWeek, True, lngFirstDayOfWeek)
                Else
                    GetWeekdayName = WeekdayName(lngDayOfWeek, False, lngFirstDayOfWeek)
                End If
                
           Case Else  ' Invalid request
                If blnAbbreviate Then
                    GetWeekdayName = "<?>"         ' short name
                Else
                    GetWeekdayName = "<Unknown>"   ' long name
                End If
    End Select
    
End Function

' ***************************************************************************
' Procedure:     GetMonthName
'
' Description:   Get the name of the month
'                Ex:   GetMonthName(2, False) will return February
'
' Parameters:    lngMonthNumber - Designates the month you want.
'                    Valid values are 1 to 12.  (Ex: January = 1)
'                blnAbbreviate - [Optional] Abbreviate the names
'                    TRUE - Shorten names to 3 chars (Default)
'                    FALSE - Use long nanes
'
' Returns:       Name of the month
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 31-Mar-2012  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Function GetMonthName(ByVal lngMonthNumber As Long, _
                    Optional ByVal blnAbbreviate As Boolean = True) As String

    Select Case lngMonthNumber
    
           Case 1 To 12   ' Make API call to get month name
                If blnAbbreviate Then
                    GetMonthName = MonthName(lngMonthNumber, True)
                Else
                    GetMonthName = MonthName(lngMonthNumber, False)
                End If
                
           Case Else  ' Invalid request
                If blnAbbreviate Then
                    GetMonthName = "<?>"         ' short name
                Else
                    GetMonthName = "<Unknown>"   ' long name
                End If
    End Select
    
End Function

' ***************************************************************************
' Routine:       DayOfMonth
'
' Description:   Determine day of a specific month of a specific year.
'                Takes in consideration if this is a leap year or not.
'
'                Some dates to remember:
'                    Martin L. King Day is 3rd Monday in January
'                    Presidents Day is 3rd Monday in February
'                    Mothers Day is 2nd Sunday in May
'                    Fathers Day is 3rd Sunday in June
'                    Labor Day is 1st Monday in September
'                    Thanksgiving Day in USA IS 4th Thursday in November
'                    Daylight savings starts 2nd Sunday in March (2:00 AM local time)
'                    Daylight savings ends 1st Sunday in November (2:00 AM local time)
'
' Reference:     http://www.usa.gov/citizens/holidays.shtml
'                http://aa.usno.navy.mil/faq/docs/daylight_time.php
'
' Parameters:    lngNthDay - Day occurance of a week (1=1st, 2=2nd, etc)
'                lngWeekday - Numeric placement of the weekday.
'                lngMonth - Number of month to be queried. (2=February)
'                lngYear - Year being queried (prefer 4 digits; ex: 2012)
'                          If one to two-digit year is passed, year is
'                          evaluated to be either current or previous
'                          century.  See IsValidYear() routine.
'                lngFirstDayOfWeek - Day considered first day in a week
'                          (ex: 1-Sunday, 2-Monday, etc).  Remember to
'                          adjust the week day accordingly.
'                          (ex:  if Sun first day of week then Thu=5
'                                if Mon first day of week then Thu=4, etc)
'
' Syntax:        Thanksgiving Day in USA (Fourth Thursday in November 2012)
'                    lngNthDay = 4    4th
'                    lngWeekday = 5   Thursday
'                    lngMonth = 11    November
'                    lngYear = 2012   Prefer 4-digit year
'
'                    DayOfMonth = 22
'
' Returns:       Day number in month and year requested
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Jan-2012  Paul Bahlawan
'              http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=74245&lngWId=1
' 25-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function DayOfMonth(ByVal lngNthDay As Long, _
                           ByVal lngWeekDay As Long, _
                           ByVal lngMonth As Long, _
                           ByVal lngYear As Long, _
                  Optional ByVal lngFirstDayOfWeek As enumDayOfWeek = eSunday) As Long
                       
    Dim lngDay       As Long  ' Temp calculated day
    Dim lngDayOfWeek As Long  ' 1-Sun, 2-Mon, etc
    
    Const DAYS_IN_WEEK As Long = 7
    
    ' Evaluate data passed to this routine
    If (lngNthDay < 0) Or (lngNthDay > 5) Or _
       (lngWeekDay < 1) Or (lngWeekDay > 7) Or _
       (lngMonth < 1) Or (lngMonth > 12) Or _
       (lngFirstDayOfWeek < 1) Or (lngFirstDayOfWeek > 7) Then

        DayOfMonth = 0   ' Invalid data passed
        Exit Function
    End If
    
    ' Evaluate year
    If IsValidYear(lngYear) Then
    
        If lngNthDay > 0 Then
               
            ' Calculate day of month based on maximum of
            ' four weeks per month. Calculated in individual
            ' steps for easier understanding.
            lngDay = ((lngNthDay - 1) * DAYS_IN_WEEK) + 1
            lngDayOfWeek = DayOfWeek(DateSerial(lngYear, lngMonth, lngDay), lngFirstDayOfWeek)
            DayOfMonth = (lngDay + (DAYS_IN_WEEK + lngWeekDay - lngDayOfWeek) Mod DAYS_IN_WEEK)
        
            ' Verify day of month does not exceed maximum
            ' days for a requested month.
            ' ex:  5th Thu in Jan 2012 [DayOfMonth = 33]
            If DayOfMonth > LastDayOfMonth(DateSerial(lngYear, lngMonth, 1)) Then
                DayOfMonth = 0
            End If
            
        Else
            ' If Nth day equal zero, return last day in month
            DayOfMonth = LastDayOfMonth(DateSerial(lngYear, lngMonth, 1))
        End If
    
    Else
        DayOfMonth = 0   ' Invalid year
    End If

End Function

' ***************************************************************************
' Routine:       EasterSunday
'
' Description:   In Western Christianity, Easter always falls on a Sunday
'                from March 22 to April 25 inclusive. The following day,
'                Easter Monday, is a legal holiday in many countries with
'                predominantly Christian traditions.
'
' Reference:     http://www.merlyn.demon.co.uk/estralgs.txt
'                http://aa.usno.navy.mil/faq/docs/easter.php
'                http://www.wheniseastersunday.com/year/1990/
'                http://www.celebrating-easter.com/celebrating/easter-dates.shtml
'
'                Easter Sunday 1990-2029 (yyyy-mm-dd format)
'                  1990-04-15    2000-04-23    2010-04-04    2020-04-12
'                  1991-03-31    2001-04-15    2011-04-24    2021-04-04
'                  1992-04-19    2002-03-31    2012-04-08    2022-04-17
'                  1993-04-11    2003-04-20    2013-03-31    2023-04-09
'                  1994-04-03    2004-04-11    2014-04-20    2024-03-31
'                  1995-04-16    2005-03-27    2015-04-05    2025-04-20
'                  1996-04-07    2006-04-16    2016-03-27    2026-04-05
'                  1997-03-30    2007-04-08    2017-04-16    2027-03-28
'                  1998-04-12    2008-03-23    2018-04-01    2028-04-16
'                  1999-04-04    2009-04-12    2019-04-21    2029-04-01
'
' Parameters:    lngYear - Year being queried (prefer 4 digits; ex: 2012)
'                          If one to two-digit year is passed, year is
'                          evaluated to be either current or previous
'                          century.  See IsValidYear() routine.
'                strMonthDay - Optional - Returns formatted short month name
'                          and day number for Easter Sunday
'
' Returns:       Full date for Easter Sunday (ex:  4/15/1990)
'                Optional - Short month name and day number for Easter Sunday
'                          (ex:  "Apr 15")
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 26-Jan-2012  Paul Bahlawan
'              http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=74245&lngWId=1
' 29-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 02-Feb-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Fixed bug:  While cleaning up code, I added too many parens
'                and was getting the wrong calculation for lngDay.
'              - Thanks to Paul Bahlawan for his suggestions.
' 05-Feb-2012  Kenneth Ives  kenaso@tx.rr.com
'              - More enhancement suggestions from Paul Bahlawan.
'              - Updated documentation.
' ***************************************************************************
Public Function EasterSunday(ByVal lngYear As Long, _
                    Optional ByRef strMonthDay As String = vbNullString) As Date

    Dim lngPhase        As Long  ' xx
    Dim lngBCPcypher    As Long  ' cy  BCP cypher
    Dim lngDayInMarch   As Long  ' DM  Day in March
    Dim lngGoldenNumber As Long  ' gn  Golden number
    
    ' Evaluate year
    If IsValidYear(lngYear) Then
    
        lngGoldenNumber = lngYear Mod 19   ' gn = YR % 19
        lngPhase = (lngYear \ 100)         ' xx = (YR/100)|0
        
        ' cy = ((3*(xx+1)/4)|0) - (((13+xx*8)/25)|0)
        lngBCPcypher = (3 * (lngPhase + 1) \ 4) - ((13 + lngPhase * 8) \ 25)
        
        ' xx = ( 6 + YR + ((YR/4)|0) - xx + ((YR/400)|0) ) % 7
        lngPhase = (6 + lngYear + (lngYear \ 4) - lngPhase + (lngYear \ 400)) Mod 7
        
        ' DM = 21 + (gn*19 + cy + 15)%30
        lngDayInMarch = 21 + (lngGoldenNumber * 19 + lngBCPcypher + 15) Mod 30
        
        '***************************************************
        ' DM -= ((gn>10) + DM > 49)
        '
        ' Above algorithm explained:
        '
        '  If (lngGoldenNumber > 10) Then
        '      lngDayInMarch = lngDayInMarch + lngDayInMarch
        '      If lngDayInMarch > 49 Then
        '          lngDayInMarch = lngDayInMarch - lngDayInMarch
        '      End If
        '  End If
        '***************************************************
        lngDayInMarch = lngDayInMarch - ((lngGoldenNumber > 10) + lngDayInMarch > 49)
        
        ' return DM + 1 + (66-xx-DM)%7
        lngDayInMarch = lngDayInMarch + 1 + (66 - lngPhase - lngDayInMarch) Mod 7
        
        ' Calculate full date.  If day in March calculation
        ' is greater than 31 days, the month becomes April
        ' and the day is adjusted accordingly.
        EasterSunday = DateSerial(lngYear, 3, lngDayInMarch)
        
        ' Format month and day (ex:  "Apr 15")
        strMonthDay = GetMonthName(Month(EasterSunday)) & " " & Day(EasterSunday)

    End If
    
End Function

' ***************************************************************************
' Routine:       CurrentUTC
'
' Description:   Calculate current Coordinated Universal Time (UTC) for
'                this locale. Coordinated Universal Time (UTC) is the
'                primary time standard by which the world regulates clocks
'                and time. It is one of several closely related successors
'                to Greenwich Mean Time (GMT) (Greenwich, England) which
'                is also known as Zulu Time.
'
'                Short list of a few time zones:
'                   MSD    Moscow Daylight Time              GMT +4 hours
'                   MSK    Moscow Standard Time              GMT +3 hours
'                   EET    Eastern European Time             GMT +2 hours
'                   CET    Central European Time             GMT +1 hour
'                   IST    Irish Standard Time               GMT +1 hour
'                   GMT    Greenwich Mean Time               GMT *********
'                   WET    Western European Time             GMT
'                   EGT    East Greenland Time               GMT -1 hour
'                   PMDT   Pierre & Miquelon Daylight Time   GMT -2 hours
'                   PMST   Pierre & Miquelon Standard Time   GMT -3 hours
'                   NDT    Newfoundland Daylight Time        GMT -2:30 hours
'                   NST    Newfoundland Standard Time        GMT -3:30 hours
'                   ADT    Atlantic Daylight Time            GMT -3 hours
'                   AST    Atlantic Standard Time            GMT -4 hours
'                   WGT    West Greenland Time               GMT -3 hours
'                   EDT    Eastern Daylight Time             GMT -4 hours
'                   EST    Eastern Standard Time             GMT -5 hours
'                   CDT    Central Daylight Time             GMT -5 hours
'                   CST    Central Standard Time             GMT -6 hours
'                   MDT    Mountain Daylight Time            GMT -6 hours
'                   MST    Mountain Standard Time            GMT -7 hours
'                   PDT    Pacific Daylight Time             GMT -7 hours
'                   PST    Pacific Standard Time             GMT -8 hours
'                   AKDT   Alaska Daylight Time              GMT -8 hours
'                   AKST   Alaska Standard Time              GMT -9 hours
'                   HADT   Hawaii-Aleutian Daylight Time     GMT -9 hours
'                   HAST   Hawaii-Aleutian Standard Time     GMT -10 hours
'
' References:    http://wwp.greenwichmeantime.com/
'                http://wwp.greenwichmeantime.com/what-is-gmt/index.htm
'                http://www.timeanddate.com/library/abbreviations/timezones/
'
' Parameters:    strHoursMinutes - Optional - Hours and minutes from
'                          Greenwich Mean Time (GMT)
'                          ex:  Berlin, Germany (CET)  GMT +1:00 hour
'                               Dallas, TX, USA (CST)  GMT -6:00 hours
'
' Returns:       Full date and time
'                Optional - Hours and minutes from Greenwich Mean Time
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 08-Feb-2012  Paul Bahlawan
'              http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=74245&lngWId=1
' 11-Feb-2012  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function CurrentUTC(Optional ByRef strHoursMinutes As String = vbNullString) As Date
    
    Dim datDate As Date
    Dim typTZI  As TIME_ZONE_INFORMATION
    
    ZeroMemory typTZI, Len(typTZI)  ' Clear type structure
    GetTimeZoneInformation typTZI   ' Make API call to load type structure
    datDate = Now()                 ' Capture current date/time
    
    ' Calc Coordinated Universal Time (UTC)
    ' by adding number of minutes difference
    ' between this locale and Greenwich Mean
    ' Time (GMT)
    CurrentUTC = DateAdd("n", typTZI.bias, datDate)
    
    Select Case CurrentUTC
    
           Case Is < datDate  ' ex:  Berlin, Germany (CET)  GMT +1:00 hours
                strHoursMinutes = "GMT +" & Format$(Hour(datDate - CurrentUTC), "#0") & ":" & _
                                            Format$(Minute(datDate - CurrentUTC), "00") & " hours"
                                        
           Case Is > datDate  ' ex:  Dallas, TX, USA (CST)  GMT -6:00 hours
                strHoursMinutes = "GMT -" & Format$(Hour(CurrentUTC - datDate), "#0") & ":" & _
                                            Format$(Minute(CurrentUTC - datDate), "00") & " hours"
                                        
           Case Else: strHoursMinutes = "0:00 hours"  ' Welcome to Greenwich, England
    End Select
    
    ZeroMemory typTZI, Len(typTZI)  ' Clear type structure

End Function

' ***************************************************************************
' Routine:       IsDaylightSavings
'
' Description:   Determine if this locale is currently observing Daylight
'                Savings Time or Standard Time.
'
' Returns:       True  - Observing Daylight Savings Time (DST)
'                False - Observing standard time or fixed time
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 08-Feb-2012  Paul Bahlawan
'              http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=74245&lngWId=1
' 11-Feb-2012  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function IsDaylightSavings() As Boolean
    
    Dim typTZI As TIME_ZONE_INFORMATION
    
    ZeroMemory typTZI, Len(typTZI)  ' Clear type structure

    If GetTimeZoneInformation(typTZI) = TIME_ZONE_ID_DAYLIGHT Then
        IsDaylightSavings = True
    Else
        IsDaylightSavings = False
    End If
    
    ZeroMemory typTZI, Len(typTZI)  ' Clear type structure

End Function

' ***************************************************************************
' Routine:       DaylightSavingsBegins
'
' Description:   Calculate when Daylight Savings time begins for this
'                locale if it is being observed.  You can also use
'                DayOfMonth() routine to calculate this date.
'
'                Daylight savings starts 2nd Sunday in Mar 2:00:00 AM
'
' Parameters:    lngYear - Year being queried (prefer 4 digits; ex: 2012)
'                          If one to two-digit year is passed, year is
'                          evaluated to be either current or previous
'                          century.  See IsValidYear() routine.
'
' Returns:       Formatted date and time
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 08-Feb-2012  Paul Bahlawan
'              http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=74245&lngWId=1
' 11-Feb-2012  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function DaylightSavingsBegins(ByVal lngYear As Long) As String
    
    Dim typTZI As TIME_ZONE_INFORMATION
        
    ' Evaluate year
    If IsValidYear(lngYear) Then
    
        ZeroMemory typTZI, Len(typTZI)  ' Clear type structure
        GetTimeZoneInformation typTZI   ' Make API call
        
        With typTZI.DaylightDate
        
            ' .wDay = week in month
            If (.wDay < 1) Or (.wDay > 4) Then
                
                DaylightSavingsBegins = vbNullString   ' Invalid week
            Else
                ' Weeks 1 to 4 (ex:  2nd Sunday)
                ' Format return date (ex: 3/11/2012 2:00:00 AM)
                DaylightSavingsBegins = CStr(DateSerial(lngYear, .wMonth, DayOfMonth(.wDay, .wDayOfWeek + 1, .wMonth, lngYear)) & _
                                             Chr$(32) & _
                                             TimeSerial(.wHour, .wMinute, .wSecond))
            End If
        End With
        
        ZeroMemory typTZI, Len(typTZI)   ' Clear type structure
    Else
        DaylightSavingsBegins = vbNullString   ' Invalid year
    End If
        
End Function

' ***************************************************************************
' Routine:       DaylightSavingsEnds
'
' Description:   Calculate when Daylight Savings time ends for this locale
'                if it is being observed.  You can also use DayOfMonth()
'                routine to calculate this date.
'
'                Daylight savings ends 1st Sunday in Nov 2:00:00 AM
'
' Parameters:    lngYear - Year being queried (prefer 4 digits; ex: 2012)
'                          If one to two-digit year is passed, year is
'                          evaluated to be either current or previous
'                          century.  See IsValidYear() routine.
'
' Returns:       Formatted date and time
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 08-Feb-2012  Paul Bahlawan
'              http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=74245&lngWId=1
' 11-Feb-2012  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function DaylightSavingsEnds(ByVal lngYear As Long) As String

    Dim typTZI As TIME_ZONE_INFORMATION
    
    ' Evaluate year
    If IsValidYear(lngYear) Then
    
        ZeroMemory typTZI, Len(typTZI)  ' Clear type structure
        GetTimeZoneInformation typTZI   ' Make API call
        
        With typTZI.StandardDate
        
            ' .wDay = week in month
            If (.wDay < 1) Or (.wDay > 4) Then
                
                DaylightSavingsEnds = vbNullString   ' Invalid week
            Else
                ' Weeks 1 to 4 (ex:  1st Sunday)
                ' Format return date (ex: 11/4/2012 2:00:00 AM)
                DaylightSavingsEnds = CStr(DateSerial(lngYear, .wMonth, DayOfMonth(.wDay, .wDayOfWeek + 1, .wMonth, lngYear)) & _
                                           Chr$(32) & _
                                           TimeSerial(.wHour, .wMinute, .wSecond))
            End If
        End With
        
        ZeroMemory typTZI, Len(typTZI)   ' Clear type structure
    Else
        DaylightSavingsEnds = vbNullString   ' Invalid year
    End If
        
End Function

' ***************************************************************************
' Procedure:     GetTimeStamp
'
' Description:   Format system date and time to be used in a log file.  If
'                no format is passed, this routine will use the current
'                system short date and time format.
'
' Parameters:    strDateFmt - [Optional] Date format desired.
'                strTimeFmt - [Optional] Time format desired
'
' Returns:       Formatted date and time stamp
'                Ex:  6/17/2005 10:12:45 AM     <-- My local settings
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-Jun-2005  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function GetTimeStamp(Optional ByVal strDateFmt As String = vbNullString, _
                             Optional ByVal strTimeFmt As String = vbNullString) As String

    ' If date format is empty then use
    ' current locale short date setting
    If Len(Trim$(strDateFmt)) = 0 Then
        strDateFmt = mstrShortDateFormat
    End If
    
    ' If time format is empty then use
    ' current locale time setting
    If Len(Trim$(strTimeFmt)) = 0 Then
        strTimeFmt = mstrTimeFormat
    End If
    
    ' Return time stamp string
    GetTimeStamp = Format$(Now(), strDateFmt) & " " & _
                   Format$(Now(), strTimeFmt)

End Function


' ***************************************************************************
' Routine:       Y2K38_SecondsToDate
'
' Description:   The next Y2K event in programming is referred to as Y2K38.
'                Ever heard the phrase "Unix time"?  That's the most
'                often-used description of the method of marking time by
'                the number of seconds that have passed since Jan 1, 1970.
'                A question arose about this, and how negative values were
'                starting to creep into VB calculations, ostensibly because
'                of the way VB uses the high-bit to indicate sign.  The
'                Clock rolls over on 1/19/2038 at 3:14:07 AM, if the base
'                date is the commonly used midnight on 1/1/1970.
'
'                Problem solved! For now, anyway, right? Well, not so fast.
'                It turns out that, even though Windows wasn't around back
'                then, Unix was indeed around before 1970 and the negative
'                numbers are by design. D'oh! Of course there's a need to
'                express dates in the past, too.
'
' Reference:     The Next Y2K (Short article - A must read)
'                http://visualstudiomagazine.com/articles/2010/02/16/the-next-y2k.aspx
'
' Parameters:    dblNetDate - Date in seconds
'
' Returns:       Correct VB date
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 16-Feb-2010  Karl E. Peterson
'              http://vb.mvps.org/
' 01-Mar-2011  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function Y2K38_SecondsToDate(ByVal dblNetDate As Double) As Date

    ' Be aware that many "Unix" times are expressed in GMT, so
    ' you may also need to adjust for local offset as needed.
    ' See:  http://vb.mvps.org/samples/TimeZone
    '
    ' Supports VB Dates through: 2/7/2106 6:28:15 AM
    
    Dim dblSeconds As Double
    
    Const SECONDS_PER_DAY As Double = 86400#
    
    If dblNetDate >= 0# Then
        dblSeconds = dblNetDate
    Else
        dblSeconds = dblNetDate + BIT_32
    End If
    
    Y2K38_SecondsToDate = CDate(CDbl(BASE_DATE) + (dblSeconds / SECONDS_PER_DAY))
    
End Function

' ***************************************************************************
' Routine:       Y2K38_DateToSeconds
'
' Description:   The next Y2K event in programming is referred to as Y2K38.
'                Ever heard the phrase "Unix time"?  That's the most
'                often-used description of the method of marking time by
'                the number of seconds that have passed since Jan 1, 1970.
'                A question arose about this, and how negative values were
'                starting to creep into VB calculations, ostensibly because
'                of the way VB uses the high-bit to indicate sign.  The
'                Clock rolls over on 1/19/2038 at 3:14:07 AM, if the base
'                date is the commonly used midnight on 1/1/1970.
'
'                Problem solved! For now, anyway, right? Well, not so fast.
'                It turns out that, even though Windows wasn't around back
'                then, Unix was indeed around before 1970 and the negative
'                numbers are by design. D'oh! Of course there's a need to
'                express dates in the past, too.
'
' Reference:     The Next Y2K (Short article - A must read)
'                http://visualstudiomagazine.com/articles/2010/02/16/the-next-y2k.aspx
'
' Parameters:    datDate - Date to be converted
'
' Returns:       Date in seconds
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 16-Feb-2010  Karl E. Peterson
'              http://vb.mvps.org/
' 01-Mar-2011  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function Y2K38_DateToSeconds(ByVal datDate As Date) As Double

    ' Be aware that many "Unix" times are expressed in GMT, so
    ' you may also need to adjust for local offset as needed.
    ' See:  http://vb.mvps.org/samples/TimeZone
    '
    ' Supports VB Dates through: 2/7/2106 6:28:15 AM
    
    Dim dblSeconds As Double
    
    ' Calculate difference in seconds between
    ' base date and date passed here
    dblSeconds = CDbl(DateDiff("s", BASE_DATE, datDate))
    
    Select Case dblSeconds
           Case Is >= BIT_32: dblSeconds = 0   ' Should never see this
           Case Is >= BIT_31: dblSeconds = dblSeconds - BIT_32
    End Select
    
    Y2K38_DateToSeconds = dblSeconds
    
End Function

' ***************************************************************************
' Routine:       IsValidYear
'
' Description:   Determines if year is a valid value.
'
' Reference:     The 2029 Rule used in Microsoft Excel:
'                http://support.microsoft.com/kb/214391
'
' Parameters:    lngYear - One, two, or four digits representing a year
'
' Returns:       Formatted 4-digit year
'                True - Valid year
'                False - Invalid year
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 25-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function IsValidYear(ByRef lngYear As Long) As Boolean

    ' Called by IsLeapYear()
    '           DayOfMonth()
    '           DaylightSavingsBegins()
    '           DaylightSavingsEnds()
        
    Select Case lngYear
           Case Is < 0: IsValidYear = False   ' Negative value passed
                
           Case 0 To 99    ' 1 to 2-digits (Same concept Microsoft Excel uses)
                If lngYear < 30 Then
                    lngYear = 2000 + lngYear   ' Add to current century
                Else
                    lngYear = 1900 + lngYear   ' Add to previous century
                End If
                
                IsValidYear = True   ' Valid year
           
           Case Else
                ' Assume 4-digit year
                ' No leap years prior to 1752.
                ' Stay within a reasonable time frame.
                If (lngYear > 1751) And (lngYear < 2200) Then
                    IsValidYear = True    ' Valid year
                Else
                    IsValidYear = False   ' Invalid range or 3-digits passed
                End If
    End Select
    
End Function



' ***************************************************************************
' ****               Internal Procedures and Functions                   ****
' ***************************************************************************

' ***************************************************************************
' Routine:       GetLocaleDateFormat
'
' Description:   Capture the date formatting for this particular locale
'                Ex:  Short date       M/d/yyyy
'                     Long date        dddd, MMMM dd, yyyy
'                     Date separator   /
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-JUN-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 06-JUL-2002  Kenneth Ives  kenaso@tx.rr.com
'              Corrected obtaining Locale Country Code ID by capturing the
'              user's personal PC locale settings.  Thanks to Alfred
'              Hellmüller for spotting this.
' 15-Aug-2009  Kenneth Ives  kenaso@tx.rr.com
'              Modified to load property values
' ***************************************************************************
Private Sub GetLocaleDateFormat()

    ' Called by Class_Initialize()
    
    Dim LCID As Long  ' local country code identifier

    On Error GoTo GetLocaleDateFormat_Error

    LCID = GetUserDefaultLCID()  ' capture local country code

    ' get the date formats that will be returned
    mstrShortDateFormat = GetUserLocaleInfo(LCID, LOCALE_SSHORTDATE)
    mstrLongDateFormat = GetUserLocaleInfo(LCID, LOCALE_SLONGDATE)
    mstrDateSeparator = GetUserLocaleInfo(LCID, LOCALE_SDATE)

GetLocaleDateFormat_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Sub

GetLocaleDateFormat_Error:
    Err.Clear
    mstrShortDateFormat = vbNullString
    mstrLongDateFormat = vbNullString
    mstrDateSeparator = vbNullString
    Resume GetLocaleDateFormat_CleanUp

End Sub

' ***************************************************************************
' Routine:       GetLocaleTimeFormat
'
' Description:   Capture the time formatting for this particular locale.
'                Ex:  Time format      h:mm:ss AMPM
'                     Time separator   :
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-JUN-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 06-JUL-2002  Kenneth Ives  kenaso@tx.rr.com
'              Corrected obtaining Locale Country Code ID by capturing the
'              user's personal PC locale settings.  Thanks to Alfred
'              Hellmüller for spotting this.
' 15-Aug-2009  Kenneth Ives  kenaso@tx.rr.com
'              Modified to load property values
' ***************************************************************************
Private Sub GetLocaleTimeFormat()

    ' Called by Class_Initialize()
    
    Dim LCID   As Long       ' local country code identifier
    Dim intPos As Integer    ' String pointer
    Dim strAM  As String     ' AM/PM designators
    Dim strPM  As String

    On Error GoTo GetLocaleTimeFormat_Error

    LCID = GetUserDefaultLCID()    ' capture local country code

    ' get the Short date and time format
    mstrTimeFormat = GetUserLocaleInfo(LCID, LOCALE_STIMEFORMAT)
    mstrTimeSeparator = GetUserLocaleInfo(LCID, LOCALE_STIME)

    ' If there is a trailing time indicator then capture what is used for AM/PM
    ' within this locale time format
    intPos = InStr(1, Trim$(mstrTimeFormat), " ")
    
    If intPos > 0 Then
        strAM = GetUserLocaleInfo(LCID, LOCALE_S1159)  ' AM indicator
        strPM = GetUserLocaleInfo(LCID, LOCALE_S2359)  ' PM indicator
        
        ' add the indicators and remove blanks accordingly
        mstrTimeFormat = Trim$(Left$(mstrTimeFormat, intPos)) & " " & _
                         Trim$(strAM) & Trim$(strPM)
    End If

GetLocaleTimeFormat_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Sub

GetLocaleTimeFormat_Error:
    Err.Clear
    mstrTimeFormat = vbNullString
    mstrTimeSeparator = vbNullString
    Resume GetLocaleTimeFormat_CleanUp

End Sub

' ***************************************************************************
' Routine:       GetUserLocaleInfo
'
' Description:   Capture the format settings for this particular locale
'
' Parameters:    lngLocaleID - Local country code identifier
'                lngLocaleType - One of the constants listed above that deal
'                       a particular format setting
'
' Returns:       data string representing a particular format setting
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-MAY-1999  Randy Birch   http://www.mvps.org/vbnet/index.html
'              Routine created
' 28-JUN-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified/documented
' ***************************************************************************
Private Function GetUserLocaleInfo(ByVal lngLocaleID As Long, _
                                   ByVal lngLocaleType As Long) As String

    ' Called by GetLocaleTimeFormat()
    '           GetLocaleDateFormat()
    
    Dim strBuffer  As String
    Dim lngRetCode As Long

    strBuffer = Space$(MAX_SIZE)    ' pad buffer with spaces

    ' call function passing Locale type variable to retrieve required
    ' size of string buffer needed
    lngRetCode = GetLocaleInfo(lngLocaleID, lngLocaleType, strBuffer, Len(strBuffer))

    ' if successful
    If lngRetCode > 0 Then

        ' pad the buffer with spaces
        strBuffer = Space$(MAX_SIZE)
    
        ' and call again passing the buffer
        lngRetCode = GetLocaleInfo(lngLocaleID, lngLocaleType, strBuffer, Len(strBuffer))
    
        ' if successful
        If lngRetCode > 0 Then
            ' lngRetCode holds the size of the string
            ' including the terminating null
            GetUserLocaleInfo = TrimStr(strBuffer)
        Else
            GetUserLocaleInfo = vbNullString
        End If
    Else
        GetUserLocaleInfo = vbNullString
    End If

End Function

Private Sub Class_Initialize()
    
    ' When this class is instantiated determine
    ' current date and time format for this locale.
    StopProcessing = False
    GetLocaleTimeFormat
    GetLocaleDateFormat
    
End Sub

