VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConvertDateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ***************************************************************************
' Routine:       cConvertDateTime
'
' Description:   Converts dates and/or time into a common format so the
'                results can be sorted.  If you find some additional
'                formats, please email me so I can update this module.
'
'                This module is only used by clsSort.cls when sorting dates.
'
' Important:     clsSort and clsConvertDateTime must always be together.
'
' In project group window, click on cConvertDateTime(clsConvertDataeTime.cls)
' Look at Properties window.
'
'    Instancing   1 - Private   Means this module can only be seen
'                               by modules within the same DLL (project)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-May-2012  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Enumerations
' ***************************************************************************
  Public Enum enumDateFormat
      eCDT_Date_0    ' MMM dd, yyyy
      eCDT_Date_1    ' MMM d, yyyy
      eCDT_Date_2    ' MMMM dd, yyyy
      eCDT_Date_3    ' MMMM d, yyyy

      eCDT_Date_4    ' dd-MMM-yyyy
      eCDT_Date_5    ' d-MMM-yyyy
      eCDT_Date_6    ' dd MMM yyyy
      eCDT_Date_7    ' d MMM yyyy
      eCDT_Date_8    ' dd.MMM.yyyy
      eCDT_Date_9    ' d.MMM.yyyy

      eCDT_Date_10   ' yyyy-MMM-dd
      eCDT_Date_11   ' yyyy-MMM-d
      eCDT_Date_12   ' yyyy MMM dd
      eCDT_Date_13   ' yyyy MMM d
      eCDT_Date_14   ' yyyy.MMM.dd
      eCDT_Date_15   ' yyyy.MMM.d
      
      eCDT_Date_16   ' mm/dd/yyyy
      eCDT_Date_17   ' m/d/yyyy
      eCDT_Date_18   ' mm-dd-yyyy
      eCDT_Date_19   ' m-d-yyyy
      eCDT_Date_20   ' mm.dd.yyyy
      eCDT_Date_21   ' m.d.yyyy

      eCDT_Date_22   ' dd/mm/yyyy
      eCDT_Date_23   ' d/m/yyyy
      eCDT_Date_24   ' dd-mm-yyyy
      eCDT_Date_25   ' d-m-yyyy
      eCDT_Date_26   ' dd.mm.yyyy
      eCDT_Date_27   ' d.m.yyyy
      
      eCDT_Date_28   ' yyyy/mm/dd
      eCDT_Date_29   ' yyyy/m/d
      eCDT_Date_30   ' yyyy-mm-dd
      eCDT_Date_31   ' yyyy-m-d
      eCDT_Date_32   ' yyyy.mm.dd
      eCDT_Date_33   ' yyyy.m.d
            
      eCDT_Date_34   ' yyyy/dd/mm
      eCDT_Date_35   ' yyyy/d/m
      eCDT_Date_36   ' yyyy-dd-mm
      eCDT_Date_37   ' yyyy-d-m
      eCDT_Date_38   ' yyyy.dd.mm
      eCDT_Date_39   ' yyyy.d.m
  End Enum
  
  Public Enum enumTimeFormat
      eCDT_Time_0    ' h:nn
      eCDT_Time_1    ' hh:nn
      eCDT_Time_2    ' hh:nn:ss
      eCDT_Time_3    ' hh:nn:ss AM/PM
      eCDT_Time_4    ' h:nna/p
      eCDT_Time_5    ' hh:nnam/pm

      eCDT_Time_6    ' hh:nn:ss A.M./P.M.
      eCDT_Time_7    ' hh:nn:ss.ttt
      eCDT_Time_8    ' hh:nn:ss:ttt
      eCDT_Time_9    ' hh:nn:ss.ttt AM/PM
      eCDT_Time_10   ' hh:nn:ss:ttt AM/PM
      eCDT_Time_11   ' hh:nn:ss.ttt A.M./P.M.
      eCDT_Time_12   ' hh:nn:ss:ttt A.M./P.M.

      eCDT_Time_13   ' h.nn
      eCDT_Time_14   ' hh.nn
      eCDT_Time_15   ' hh.nn.ss
      eCDT_Time_16   ' hh.nn.ss AM/PM
      eCDT_Time_17   ' h.nna/p
      eCDT_Time_18   ' hh.nnam/pm

      eCDT_Time_19   ' hh.nn.ss A.M./P.M.
      eCDT_Time_20   ' hh.nn.ss.ttt
      eCDT_Time_21   ' hh.nn.ss.ttt AM/PM
      eCDT_Time_22   ' hh.nn.ss.ttt A.M./P.M.
  
      eCDT_Time_23   ' hh:nn:ss.tttt
      eCDT_Time_24   ' hh:nn:ss:tttt
      eCDT_Time_25   ' hh.nn.ss.tttt
      eCDT_Time_26   ' hh:nn:ss.tttt AM/PM
      eCDT_Time_27   ' hh:nn:ss:tttt AM/PM
      eCDT_Time_28   ' hh.nn.ss.tttt AM/PM
      eCDT_Time_29   ' hh:nn:ss.tttt A.M./P.M.
      eCDT_Time_30   ' hh:nn:ss:tttt A.M./P.M.
      eCDT_Time_31   ' hh.nn.ss.tttt A.M./P.M.
  End Enum
                  
' ***************************************************************************
' Module Variables
'
'                    +-------------- Module level designator
'                    |  +----------- Data type (Boolean)
'                    |  |     |----- Variable subname
'                    - --- ---------
' Naming standard:   m bln ReturnNumeric
' Variable name:     mblnReturnNumeric
'
' ***************************************************************************
  Private mblnProcessTime As Boolean
  Private mlngDateFormat  As enumDateFormat
  Private mlngTimeFormat  As enumTimeFormat
  
  
' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

Public Property Let DateFormat(ByVal lngDateFormat As enumDateFormat)
    
    ' used when assigning a value to the property, on the left
    ' side of an assignment.        Syntax: X.DateFormat = eCDT_Date_0
    
    mlngDateFormat = lngDateFormat

End Property

Public Property Let TimeFormat(ByVal lngTimeFormat As enumTimeFormat)
    
    ' used when assigning a value to the property, on the left
    ' side of an assignment.        Syntax: X.TimeFormat = eCDT_Time_0
    
    mlngTimeFormat = lngTimeFormat

End Property

Public Property Let ProcessTime(ByVal blnData As Boolean)
    
    ' used when assigning a value to the property, on the left
    ' side of an assignment.       Syntax: X.ProcessTime = True
    '
    ' True  - Date and time, time only
    ' False - Date only
    '
    ' Designates to process time if Time only or Date and time are
    ' to be sorted.
    
    mblnProcessTime = blnData

End Property


' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       ConvertDateTime
'
' Description:   Converts date and/or time into a numeric or string
'                equivalent to be used for comparison compared within
'                a sort routine.  I realize there are many other scenarios
'                for displaying date and time.  If you encounter others,
'                then update accordingly.
'
' Parameters:    strInputData - Date and/or time in string format
'
' Returns:       Numeric data returned in numeric or string format
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-May-2012  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Function ConvertDateTime(ByVal strInputData As String) As Variant

    Dim dblDate  As Double   ' Converted input data
    Dim strDate  As String   ' Hold area for date data
    Dim strTime  As String   ' Hold area for time data
    Dim strFinal As String   ' Final data format
    Dim strMSecs As String   ' Hold area for millisecond data
    
    Const TEST_LEN As Long = 10

    On Error GoTo ConvertDateTime_Error

    strInputData = Trim$(strInputData)   ' Remove leading/trailing blanks
    
    ' Verify some data was passed
    If Len(strInputData) = 0 Then
        Exit Function
    End If
   
    strDate = vbNullString   ' Initialize variables
    strTime = vbNullString
    strFinal = vbNullString
    strMSecs = vbNullString
        
    '*************************************
    ' Step 1 - Work on time format
    '*************************************
    If mblnProcessTime Then
        strInputData = ReformatTime(strInputData, strMSecs)
    End If
    
    '*************************************
    ' Step 2 - Work on date format
    '*************************************
    ' Date only
    If InStr(1, strInputData, ":") = 0 Then
    
        strDate = ReformatDate(strInputData)   ' Convert date to common date format
        strInputData = strDate                 ' Transfer back to input data string
        
    ' Date and time
    ' ex:  m/d/yyyy 6:18  First colon is in 11th position
    ElseIf InStr(1, strInputData, ":") > TEST_LEN Then
        
        ' Parse out date and time
        strDate = Mid$(strInputData, 1, InStr(1, strInputData, ":") - 3)
        strTime = Mid$(strInputData, InStr(1, strInputData, ":") - 2)
        
        strDate = ReformatDate(Trim$(strDate))          ' Convert date to common date format
        strInputData = strDate & " " & Trim$(strTime)   ' Concatenate input data back together
        
    End If
    
    ' Final evaluation of parsed data
    strDate = IIf(Val(strDate) = 0, vbNullString, strDate)
    strTime = IIf(Val(strTime) = 0, vbNullString, strTime)
    
    '*************************************
    ' Step 3 - Calculate date and time
    '*************************************
    dblDate = CDbl(CDate(strInputData))
    strDate = CStr(Fix(dblDate))
    strTime = CStr(CDec(Abs(dblDate) - Fix(dblDate)))
    
    ' Concatenate final data string
    strFinal = strDate & strTime & strMSecs
    
ConvertDateTime_CleanUp:
    ConvertDateTime = Val(strFinal)   ' Return data in numeric format
    On Error GoTo 0                   ' Nullify this error trap
    Exit Function

ConvertDateTime_Error:
    Err.Clear
    strFinal = vbNullString
    Resume ConvertDateTime_CleanUp
    
End Function


' ***************************************************************************
' ****               Internal Procedures and Functions                   ****
' ***************************************************************************

Private Function ReformatTime(ByVal strData As String, _
                              ByRef strMSecs As String) As String

    ' Called by ConvertDateTime()
    
    Dim lngEnd   As Long     ' Ending position in string
    Dim lngPos   As Long     ' Time position in string
    Dim lngCount As Long     ' Count number of decimals
    Dim lngStart As Long     ' Starting position in string
    Dim strTmp   As String
    Dim strTime  As String
    
    lngCount = 0
    strTime = vbNullString
    strMSecs = vbNullString
    ReformatTime = vbNullString
    
    ' Format AM/PM
    ' Test for no space between last digit and AM/PM.
    ' Insert a space if needed.
    If UCase$(Right$(strData, 2)) Like "#A" Or _
       UCase$(Right$(strData, 2)) Like "#P" Then
       
        strData = Mid$(strData, 1, Len(strData) - 1) & " " & _
                  UCase$(Right$(strData, 1)) & "M"
    
    ElseIf UCase$(Right$(strData, 3)) Like "#AM" Or _
           UCase$(Right$(strData, 3)) Like "#PM" Then
    
            strData = Mid$(strData, 1, Len(strData) - 2) & " " & _
                      UCase$(Right$(strData, 2))
    
    ElseIf UCase$(Right$(strData, 5)) Like "#A.M." Or _
           UCase$(Right$(strData, 5)) Like "#P.M." Then
    
            strData = Mid$(strData, 1, Len(strData) - 4) & " " & _
                      UCase$(Right$(strData, 4))
    End If
    
    ' Test for periods in "A.M."  Remove periods.
    If UCase$(Right$(strData, 3)) = ".M." Then
                
        strData = Mid$(strData, 1, Len(strData) - 4) & _
                  IIf(UCase$(Right$(strData, 4)) = "A.M.", "AM", "PM")
    End If
    
    ' If there are decimals instead of
    ' colons then change them to colons
    Select Case mlngTimeFormat
            
           ' Without AM or milliseconds
           Case 13, 14, 15
                lngStart = InStrRev(strData, " ")   ' Start at end of string
                
                If lngStart = 0 Then
                    lngPos = 1          ' Time only - Start in first position
                Else
                    lngPos = lngStart   ' Date and time
                End If
                                
                strTime = Trim$(Mid$(strData, lngPos))   ' Capture time portion
                strTime = Replace(strTime, ".", ":")     ' Change any decimals to colons
                
                If lngStart = 0 Then
                    strData = strTime   ' Time only
                Else
                    ' Date and time
                    strData = Trim$(Mid$(strData, 1, lngStart)) & " " & strTime
                End If
                
           ' Without AM, with milliseconds
           Case 8, 20, 24, 25
                lngStart = InStrRev(strData, " ")   ' Start at end of string
                
                If lngStart = 0 Then
                    lngPos = 1          ' Time only - Start in first position
                Else
                    lngPos = lngStart   ' Date and time
                End If
                                
                strTime = Trim$(Mid$(strData, lngPos))   ' Capture time portion
                strTime = Replace(strTime, ":", ".")     ' Replace any colons with decimals
                strTmp = Left$(strTime, 8)               ' Capture first eight positions
                strTmp = Replace(strTmp, ".", ":")       ' Replace any decimals with colons
                Mid$(strTime, 1, 8) = strTmp             ' Update time string
                
                If lngStart = 0 Then
                    strData = strTime   ' Time only
                Else
                    ' Date and time
                    strData = Trim$(Mid$(strData, 1, lngStart)) & " " & strTime
                End If
                
           ' With AM and milliseconds
           Case 10, 12, 16 To 19, 21, 22, 26 To 31
                lngStart = InStrRev(strData, " ", 4)   ' Start prior to " AM/PM"
                
                If lngStart = 0 Then
                    lngPos = 1          ' Time only - Start in first position
                Else
                    lngPos = lngStart   ' Date and time
                End If
                
                strTime = Trim$(Mid$(strData, lngPos))   ' Capture time portion
                strTime = Replace(strTime, ":", ".")     ' Replace any colons with decimals
                strTmp = Left$(strTime, 8)               ' Capture first eight positions
                strTmp = Replace(strTmp, ".", ":")       ' Replace any decimals with colons
                Mid$(strTime, 1, 8) = strTmp             ' Update time string
                
                If lngStart = 0 Then
                    strData = strTime   ' Time only
                Else
                    ' Date and time
                    strData = Trim$(Mid$(strData, 1, lngStart)) & " " & strTime
                End If
    End Select
    
    ' Test for milliseconds
    ' If decimal positions are appended then
    ' IsDate() will return FALSE and CDate()
    ' will produce an error.  (ex: 06:18:46.014 PM)
    '   1. Parse out milliseconds
    '   2. Reformat data without milliseconds
    Select Case mlngTimeFormat
                   
           Case 7 To 12, 20 To 31
                lngStart = InStrRev(strData, ".")   ' Start position of milliseconds
                lngEnd = InStrRev(strData, " ")     ' End position of milliseconds
        
                ' Capture milliseconds data
                If lngEnd > lngStart Then
                    ' AM or PM is at end of data
                    strMSecs = Mid$(strData, lngStart + 1, lngEnd - (lngStart + 1))
                Else
                    ' Nothing after final digits
                    strMSecs = Mid$(strData, lngStart + 1)
                End If
                
                strMSecs = Trim$(strMSecs)   ' Remove any leading/trailing blanks
                
                ' Reformat data string without milliseconds
                strData = Mid$(strData, 1, lngStart - 1) & _
                          Mid$(strData, lngStart + 1 + Len(strMSecs))
    End Select
        
    ReformatTime = strData   ' Return data string

End Function

Private Function ReformatDate(ByVal strDate As String) As String

    ' Called by ConvertDateTime()
    
    Dim avntDate As Variant   ' Array of numbers
    
    Const COMMON_FMT As String = "yyyy/mm/dd"
    
    ReformatDate = vbNullString   ' Preload return string
    avntDate = Empty              ' Always start with an empty variant
        
    ' Common preparation for certain formats
    Select Case mlngDateFormat
           
           Case 16 To 39  ' Replace any dashes, periods,
                          ' forward slashes with a backslash
                strDate = Replace(strDate, "-", "/")
                strDate = Replace(strDate, ".", "/")
                strDate = Replace(strDate, "\", "/")
                
                strDate = QualifyPath(strDate, "/")   ' Append one forward slash
                avntDate = Split(strDate, "/")        ' Load digits into array
    End Select
    
    ' Convert to yyyy/mm/dd format (ex:  5-Dec-2012 -> 2012/12/05)
    Select Case mlngDateFormat
    
           Case 0 To 3
                ReformatDate = Format$(CDate(strDate), COMMON_FMT)
                
           Case 4 To 15
                strDate = Replace(strDate, ".", " ")   ' Replace any periods with a space
                strDate = Replace(strDate, "\", " ")   ' Replace any backslashes with a space
                ReformatDate = Format$(CDate(strDate), COMMON_FMT)
           
           Case 16 To 21
                ReformatDate = Format$(DateSerial(avntDate(2), _
                                                  avntDate(0), _
                                                  avntDate(1)), COMMON_FMT)
           Case 22 To 27
                ReformatDate = Format$(DateSerial(avntDate(2), _
                                                  avntDate(1), _
                                                  avntDate(0)), COMMON_FMT)
           Case 28 To 33
                ReformatDate = Format$(DateSerial(avntDate(0), _
                                                  avntDate(1), _
                                                  avntDate(2)), COMMON_FMT)
           Case 34 To 39
                ReformatDate = Format$(DateSerial(avntDate(0), _
                                                  avntDate(2), _
                                                  avntDate(1)), COMMON_FMT)
    End Select
            
    avntDate = Empty   ' Always empty variants when not needed
    
End Function

' ***************************************************************************
' Routine:       QualifyPath
'
' Description:   Adds a trailing character to the path, if missing.
'
' Parameters:    strPath - Current folder being processed.
'                strChar - Optional - Specific character to append.
'                          Default = "\"
'
' Returns:       Fully qualified path with a specific trailing character
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' Unknown      Randy Birch
'              http://vbnet.mvps.org/index.html
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified/documented
' ***************************************************************************
Private Function QualifyPath(ByVal strPath As String, _
                    Optional ByVal strChar As String = "\") As String

    strPath = Trim$(strPath)
    
    If StrComp(Right$(strPath, 1), strChar, vbTextCompare) = 0 Then
        QualifyPath = strPath
    Else
        QualifyPath = strPath & strChar
    End If
    
End Function

Private Sub Class_Initialize()
    
    ' Whenever a class is instantiated (activated), this routine
    ' is called automatically if it exist with code inside.
    ' Ex:  Set objCDT = new cConvertDateTime
                
    ' Preset property values
    DateFormat = eCDT_Date_0
    TimeFormat = eCDT_Time_0
    ProcessTime = True
    
End Sub

