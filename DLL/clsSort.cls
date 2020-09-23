VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        Sort
'
' Description:   Various sort routines using the three basic types of
'                data (numeric, string, dates).  They are some of the
'                faster sorting routines that include ShellSort,
'                CombSort and QuickSort.  Read the flowerbox above each
'                routine for additional information.
'
' Important:     clsSort and clsConvertDateTime must always be together.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 21-May-2005  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 07-AUG-2005  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote or modified all routines in this module to sort
'              indices rather than the original data.
' 25-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated RemoveDupes() routine to use CompareMethod() property.
' 20-Sep-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Removed BubbleSort() routine.  Decided to keep the faster
'                routines.
'              - Added SortDirection() property.
' 30-Mar-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Updated SortData() routine.
'              - Speeded up sorting processes by removing several DoEvents.
' 15-Apr-2012  Kenneth Ives  kenaso@tx.rr.com
'              Updated NumericSorting() routine to handle decimal values.
' 15-May-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Added call to new class object clsConvertDateTime if sorting
'                dates and/or time.
'              - Updated error messages.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME As String = "clsSort"
  
' ***************************************************************************
' Enumerations
' ***************************************************************************
  Public Enum enumSortMethod
      eShellSort           ' 0
      eCombSort            ' 1
      eQuickSort           ' 2
  End Enum

  Public Enum enumSortDirection
      eSort_Ascending      ' 0
      eSort_Descending     ' 1
  End Enum

  Public Enum enumSortTypeOfData
      eSort_Numeric        ' 0 - 12345
      eSort_String         ' 1 - "abc"
      eSort_Dates          ' 2 - "3/18/2006  6:18:46 AM"
  End Enum
  
  Public Enum enumSortCompare
      eSort_CaseSensitive  ' 0 - Exact byte match
      eSort_IgnoreCase     ' 1 - Uppercase/Lowercase considered same
  End Enum
  
' ***************************************************************************
' API Declares
' ***************************************************************************
  ' The CopyMemory function copies a block of memory from one location to
  ' another. For overlapped blocks, use the MoveMemory function.
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
          (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

  ' This is a rough translation of the GetTickCount API. The
  ' tick count of a PC is only valid for the first 49.7 days
  ' since the last reboot.  When you capture the tick count,
  ' you are capturing the total number of milliseconds elapsed
  ' since the last reboot.  The elapsed time is stored as a
  ' DWORD value. Therefore, the time will wrap around to zero
  ' if the system is run continuously for 49.7 days.
  Private Declare Function GetTickCount Lib "kernel32" () As Long
  
' ***************************************************************************
' Module Variables
'
'                    +-------------- Module level designator
'                    |  +----------- Data type (Long Integer)
'                    |  |     |----- Variable subname
'                    - --- ---------
' Naming standard:   m lng DateFormat
' Variable name:     mlngDateFormat
'
' ***************************************************************************
  Private mblnProcessTime   As Boolean
  Private mlngDateFormat    As Long
  Private mlngTimeFormat    As Long
  Private mlngSortMethod    As enumSortMethod
  Private mlngTypeOfData    As enumSortTypeOfData
  Private mlngCompareMethod As enumSortCompare
  Private mlngSortDirection As enumSortDirection
  Private mobjCDT           As cConvertDateTime     ' Convert date/time class
  
  
' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

Public Property Let StopProcessing(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.        Syntax: X.StopProcessing = True
    gblnStopProcessing = blnData
    DoEvents
End Property

Public Property Get StopProcessing() As Boolean
    
    ' used when retrieving value of a property, on the right
    ' side of an assignment.        Syntax: Debug.Print X.StopProcessing
    StopProcessing = gblnStopProcessing
    DoEvents

End Property

Public Property Let CompareMethod(ByVal lngCompareMethod As enumSortCompare)
    
    ' used when assigning a value to the property, on the left
    ' side of an assignment.        Syntax: X.CompareMethod = eSort_CaseSensitive
    mlngCompareMethod = lngCompareMethod
    DoEvents

End Property

Public Property Let SortDirection(ByVal lngSortDirection As enumSortDirection)
    
    ' used when assigning a value to the property, on the left
    ' side of an assignment.        Syntax: X.SortDirection = eSort_Ascending
    mlngSortDirection = lngSortDirection

End Property

Public Property Let SortMethod(ByVal lngSortMethod As enumSortMethod)
    
    ' used when assigning a value to the property, on the left
    ' side of an assignment.        Syntax: X.SortMethod = eShellSort
    mlngSortMethod = lngSortMethod

End Property

Public Property Let TypeOfData(ByVal lngTypeOfData As enumSortTypeOfData)
    
    ' used when assigning a value to the property, on the left
    ' side of an assignment.        Syntax: X.TypeOfData = eSort_String
    mlngTypeOfData = lngTypeOfData

End Property

Public Property Let DateFormat(ByVal lngData As Long)
    
    ' used when assigning a value to the property, on the left
    ' side of an assignment.        Syntax: X.DateFormat = eSort_Date_0
    '
    ' Ignored unless TypeOfData = eSort_Dates
    mlngDateFormat = lngData

End Property

Public Property Let TimeFormat(ByVal lngData As Long)
    
    ' used when assigning a value to the property, on the left
    ' side of an assignment.        Syntax: X.TimeFormat = eSort_Time_0
    '
    ' Ignored unless TypeOfData = eSort_Dates
    mlngTimeFormat = lngData

End Property

Public Property Let ProcessTime(ByVal blnData As Boolean)
    
    ' used when assigning a value to the property, on the left
    ' side of an assignment.        Syntax: X.ProcessTime = True
    ' Designates to process time if Date and time or time only
    ' is to be sorted.
    '
    ' Ignored unless TypeOfData = eSort_Dates
    mblnProcessTime = blnData

End Property



' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       SortData
'
' Description:   Various sort routines using the three basic types of data.
'                (numeric, string, dates).  The sorts include ShellSort,
'                CombSort and QuickSort.  This is the main routine to call.
'                My main concern is sorting the indices and using them to
'                correctly identify the original data.
'
' Parameters:    astrData() - Array to be sorted
'                strElapsedTime - [Optional] - Time elapased to perform a sort.
'                           Default = vbNullString
'
' Returns:       Sorted array (astrData())
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 21-May-2005  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 09-Mar-2007  Kenneth Ives  kenaso@tx.rr.com
'              Broke out separate sort criteria to unique routines
' 15-May-2012  Kenneth Ives  kenaso@tx.rr.com
'              Updated error messages
' ***************************************************************************
Public Sub SortData(ByRef astrData() As String, _
           Optional ByRef strElapsedTime As String)

    Const ROUTINE_NAME As String = "SortData"

    On Error GoTo SortData_Error

    strElapsedTime = vbNullString   ' Empty return string
    
    ' Test incoming data
    If Not CBool(IsArrayInitialized(astrData())) Then
        InfoMsg "Incoming array has not been properly initialized." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        Exit Sub
    End If

    ' Determine type of data to be sorted
    Select Case mlngTypeOfData
           
           Case eSort_Numeric: NumericSorting astrData(), strElapsedTime
           Case eSort_String:  StringSorting astrData(), strElapsedTime
                
           Case eSort_Dates
                Set mobjCDT = New cConvertDateTime   ' Instantiate class object
                    
                ' Pass conversion parameters
                With mobjCDT
                    .DateFormat = mlngDateFormat     ' Date format in data
                    .TimeFormat = mlngTimeFormat     ' Time format in data
                    .ProcessTime = mblnProcessTime   ' True=Time within data; False=Date only
                End With
                
                NumericSorting astrData(), strElapsedTime
                
           Case Else
                InfoMsg "Cannot determine type of data to sort." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
    End Select

SortData_CleanUp:
    Set mobjCDT = Nothing   ' Verify class object is freed from memory
    On Error GoTo 0         ' Nullify this error trap
    Exit Sub

SortData_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume SortData_CleanUp

End Sub

' ***************************************************************************
' Procedure:     RemoveDupes
'
' Description:   The data must be sorted prior to calling this routine.
'                This will remove all duplicates based on user selection
'                of case sensitivity.  Number of duplicates removed will
'                be returned.
'
'                Example:  RemoveDupes abytTemp(), lngDupes
'
'                          Remove duplicates from array comparing the data
'                          as case sensitive and return number of duplicates
'                          that were in the array.

'                                 Raw        Sorted
'                                 Data     (no dupes)
'                              --------     --------
'                                 105         100
'                                 107         101
'                                 101         105
'                                 107         106
'                                 106         107
'                                 106         109
'                                 106         110
'                                 110
'                                 100      lngDupeCnt = 3
'                                 109
'
' Parameters:    avntData - an array of data to be evaluated
'                lngDupeCnt - [OPTIONAL] Returns the number of duplicates
'                     removed. User can access this variable if desired.
'
' Returns:       An array of data with no duplicate values.  If requested,
'                number of items removed from original array.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 07-JAN-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 15-MAY-2008  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a bug that dropped trailing values
' 08-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Force all incoming data to go thru sorting process
' 25-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated to use CompareMethod() property
' ***************************************************************************
Public Function RemoveDupes(ByRef avntData As Variant, _
                   Optional ByRef lngDupeCnt As Long = 0) As Boolean

    Dim lngIdx      As Long
    Dim lngLow      As Long    ' lowest number of elements in the array
    Dim lngHigh     As Long    ' highest number of elements in the array
    Dim lngItem     As Long
    Dim lngIndex    As Long
    Dim strCurrItem As String
    Dim strPrevItem As String
    Dim avntTemp    As Variant

    Const ROUTINE_NAME As String = "RemoveDupes"

    On Error GoTo RemoveDupes_Error

    avntTemp = Empty     ' Always start with empty variants
    lngIdx = 0
    lngDupeCnt = 0
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Test incoming data
    If Not CBool(IsArrayInitialized(avntData)) Then
        InfoMsg "Incoming array has not been properly initialized." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        Exit Function
    End If

    lngLow = LBound(avntData)
    lngHigh = UBound(avntData)
    ReDim avntTemp(lngLow To lngHigh)
    
    ' Copy data to temp array
    For lngIndex = lngLow To lngHigh - 1
        avntTemp(lngIndex) = avntData(lngIndex)
    Next lngIndex
        
    ' loop backwards thru data array and remove
    ' any duplicates by emptying array element
    For lngItem = (lngHigh - 1) To (lngLow + 1) Step -1
        
        strCurrItem = Trim$(CStr(avntTemp(lngItem)))
        strPrevItem = Trim$(CStr(avntTemp(lngItem - 1)))

        If StrComp(strCurrItem, strPrevItem, mlngCompareMethod) = 0 Then
            avntTemp(lngItem) = vbNullString   ' Empty array element
            lngDupeCnt = lngDupeCnt + 1        ' Increment duplicate counter
        End If
    
    Next lngItem
    
    Erase avntData                     ' Empty return array
    ReDim avntData(lngLow To lngHigh)  ' resize return data array to max
    
    ' reload return array
    For lngIndex = lngLow To lngHigh - 1
        
        ' If there is data in array element
        ' then copy data to return array
        If Len(Trim$(avntTemp(lngIndex))) > 0 Then
            avntData(lngIdx) = avntTemp(lngIndex)
            lngIdx = lngIdx + 1
        End If

    Next lngIndex
    
    ' Resize original array to actual
    ' number of elements to be returned
    ReDim Preserve avntData(lngLow To lngIdx)

RemoveDupes_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        RemoveDupes = False  ' Set flag for bad finish
    Else
        RemoveDupes = True   ' Set flag for successful finish
    End If

    avntTemp = Empty  ' Always empty variants when not needed
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

RemoveDupes_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume RemoveDupes_CleanUp
    
End Function

' ***************************************************************************
' Routine:       IsArrayInitialized
'
' Description:   This is an ArrPtr function that determines if the passed
'                array is initialized, and if so will return the pointer
'                to the safearray header. If the array is not initialized,
'                it will return zero. Normally you need to declare a VarPtr
'                alias into msvbvm50.dll or msvbvm60.dll depending on the
'                VB version, but this function will work with vb5 or vb6.
'                It is handy to test if the array is initialized as the
'                return value is non-zero.  Use CBool to convert the return
'                value into a boolean value.
'
'                This function returns a pointer to the SAFEARRAY header of
'                any Visual Basic array, including a Visual Basic string
'                array. Substitutes both ArrPtr and StrArrPtr. This function
'                will work with vb5 or vb6 without modification.
'
'                ex:  If CBool(IsArrayInitialized(array_being_tested)) Then ...
'
' Parameters:    vntData - Data to be evaluated
'
' Returns:       Zero     - Bad data (FALSE)
'                Non-zero - Good data (TRUE)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 30-Mar-2008  RD Edwards
'              http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=69970
' ***************************************************************************
Public Function IsArrayInitialized(ByVal avntData As Variant) As Long

    Dim intDataType As Integer   ' Must be a short integer

    On Error GoTo IsArrayInitialized_Exit
    
    IsArrayInitialized = 0  ' preset to FALSE
    
    ' Get the real VarType of the argument, this is similar
    ' to VarType(), but returns also the VT_BYREF bit
    CopyMemory intDataType, avntData, 2&

    ' if a valid array was passed
    If (intDataType And vbArray) = vbArray Then
        
        ' get the address of the SAFEARRAY descriptor
        ' stored in the second half of the Variant
        ' parameter that has received the array.
        ' Thanks to Francesco Balena and Monte Hansen.
        CopyMemory IsArrayInitialized, ByVal VarPtr(avntData) + 8&, 4&
    
    End If
    
IsArrayInitialized_Exit:
    On Error GoTo 0   ' Nullify this error trap

End Function


' ***************************************************************************
' ****               Internal Procedures and Functions                   ****
' ***************************************************************************

Private Sub NumericSorting(ByRef astrData() As String, _
                           ByRef strElapsedTime As String)

    ' Called by SortData()
        
    Dim lngStart       As Long      ' Start time in milliseconds
    Dim lngLow         As Long      ' lowest number of elements in array
    Dim lngHigh        As Long      ' highest number of elements in array
    Dim lngIndex       As Long      ' loop counter
    Dim alngPointers() As Long      ' array of data pointers
    Dim adblWork()     As Double    ' Sort with this data not the original
    Dim astrHold()     As String    ' Temp hold array

    Const ROUTINE_NAME As String = "NumericSorting"

    On Error GoTo NumericSorting_Error
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo NumericSorting_CleanUp
    End If
        
    lngLow = LBound(astrData)
    lngHigh = UBound(astrData)
    
    ' Not enough data to sort
    If (lngHigh - lngLow) < 2 Then
        GoTo NumericSorting_CleanUp
    End If
    
    ' too much overhead if only two items in array
    If (lngHigh - lngLow) = 2 Then

        ' Get starting time
        lngStart = CCur(GetTickCount)
        
        Select Case mlngTypeOfData
        
               Case eSort_Dates
                    ReDim adblWork(lngLow To lngHigh - 1)   ' Resize temp data array
                    
                    With mobjCDT
                        For lngIndex = lngLow To (lngHigh - 1)
                            
                            ' Format date and/or time
                            adblWork(lngIndex) = .ConvertDateTime(astrData(lngIndex))
                        Next lngIndex
                    End With
                    
                    Select Case mlngSortDirection
                    
                           Case eSort_Ascending
                                If adblWork(lngLow) > adblWork(lngHigh - 1) Then
                                    SwapData astrData(lngLow), astrData(lngHigh - 1)
                                End If
                                
                           Case eSort_Descending
                                If adblWork(lngLow) < adblWork(lngHigh - 1) Then
                                    SwapData astrData(lngLow), astrData(lngHigh - 1)
                                End If
                                
                    End Select  ' mlngSortDirection
                    
               Case Else
        
                    Select Case mlngSortDirection
                    
                           Case eSort_Ascending
                                If Val(astrData(lngLow)) > Val(astrData(lngHigh - 1)) Then
                                    SwapData astrData(lngLow), astrData(lngHigh - 1)
                                End If
                                
                           Case eSort_Descending
                                If Val(astrData(lngLow)) < Val(astrData(lngHigh - 1)) Then
                                    SwapData astrData(lngLow), astrData(lngHigh - 1)
                                End If
                                
                    End Select  ' mlngSortDirection
        End Select  ' mlngTypeOfData
                
        ' Get ending time
        strElapsedTime = ElapsedTime(GetTickCount() - lngStart)   ' Calc elapsed time
                
        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            InfoMsg "Error loading numeric arrays to be sorted." & _
                    vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
            GoTo NumericSorting_CleanUp
        End If
                        
        GoTo NumericSorting_CleanUp                      ' Exit this routine
    
    End If
    
    ' More than 2 items in array
    ReDim alngPointers(lngLow To lngHigh - 1)  ' Resize pointer array
    ReDim adblWork(lngLow To lngHigh - 1)      ' Resize temp data array
    
    Select Case mlngTypeOfData
    
           Case eSort_Dates
                With mobjCDT
                    For lngIndex = lngLow To (lngHigh - 1)
                        
                        adblWork(lngIndex) = .ConvertDateTime(astrData(lngIndex))   ' Format date and/or time
                        alngPointers(lngIndex) = lngIndex                           ' Load pointer array
                        
                    Next lngIndex
                End With
                
           Case Else
                For lngIndex = LBound(adblWork) To UBound(adblWork)
                    
                    adblWork(lngIndex) = CDbl(astrData(lngIndex))   ' Copy data to temp data array
                    alngPointers(lngIndex) = lngIndex               ' load pointer array
                    
                Next lngIndex
    End Select  ' mlngTypeOfData
        
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        InfoMsg "Error loading numeric arrays to be sorted." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        GoTo NumericSorting_CleanUp
    End If
                        
    ' Get starting time
    lngStart = CCur(GetTickCount)
        
    ' Determine sort methodology to use
    Select Case mlngSortMethod

           Case eShellSort   ' 0
                ShellSortNumeric adblWork(), alngPointers()  ' Sort data

           Case eCombSort    ' 1
                CombSortNumeric adblWork(), alngPointers()   ' Sort data

           Case eQuickSort   ' 2  TriQuickSort
                ' Read QuickSort flowerbox for additional information.
                QSortNumeric adblWork(), alngPointers(), LBound(adblWork), UBound(adblWork)
                
                ' Use a modified Insertion Sort for final phase
                QSortNumericFinal adblWork(), alngPointers(), LBound(adblWork), UBound(adblWork)
                
    End Select  ' mlngSortMethod
    
    ' Get ending time
    strElapsedTime = ElapsedTime(GetTickCount() - lngStart)   ' Calc elapsed time
    
    ' Verify no pointers were lost
    If (lngHigh - 1) <> UBound(alngPointers) Then
        gblnStopProcessing = True
    End If
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        InfoMsg "Error while sorting numeric data." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        GoTo NumericSorting_CleanUp
    End If
            
    ' If return data is to be in descending
    ' order then reverse pointers
    If mlngSortDirection = eSort_Descending Then
        ReversePointers alngPointers()
    End If
    
    ' Prepare temp holding
    ' area for string data
    ReDim astrHold(lngLow To lngHigh)
    
    ' Transfer data from original
    ' array to temp hold array
    For lngIndex = lngLow To lngHigh - 1
        astrHold(lngIndex) = astrData(lngIndex)
    Next lngIndex

    ReDim astrData(lngLow To lngHigh)  ' Resize original array
    
    ' Reload original array with sorted data
    For lngIndex = lngLow To lngHigh - 1
        astrData(lngIndex) = astrHold(alngPointers(lngIndex))
    Next lngIndex
           
NumericSorting_CleanUp:
    Erase adblWork()       ' Always empty arrays when not needed
    Erase astrHold()
    Erase alngPointers()
    
    On Error GoTo 0   ' Nullify this error trap
    Exit Sub

NumericSorting_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume NumericSorting_CleanUp
        
End Sub

Private Sub StringSorting(ByRef astrData() As String, _
                          ByRef strElapsedTime As String)

    ' Called by SortData()
    Dim lngStart       As Long      ' Start time in milliseconds
    Dim lngLow         As Long      ' lowest number of elements in array
    Dim lngHigh        As Long      ' highest number of elements in array
    Dim lngIndex       As Long      ' loop counter
    Dim alngPointers() As Long      ' array of data pointers
    Dim astrWork()     As String    ' Temp work array

    Const ROUTINE_NAME As String = "StringSorting"

    On Error GoTo StringSorting_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo StringSorting_CleanUp
    End If
   
    lngLow = LBound(astrData)
    lngHigh = UBound(astrData)
    
    ' Not enough data to sort
    If (lngHigh - lngLow) < 2 Then
        GoTo StringSorting_CleanUp
    End If
    
    ' too much overhead if only two items in array
    If (lngHigh - lngLow) = 2 Then

        ' Get starting time
        lngStart = CCur(GetTickCount)
            
        Select Case mlngSortDirection
        
               Case eSort_Ascending
                    Select Case mlngCompareMethod
                    
                           Case eSort_IgnoreCase   ' NOT case sensitive comparison
                                ' Uppercase and lowercase considered as same character
                                If StrComp(astrData(lngLow), astrData(lngHigh - 1), vbTextCompare) = 1 Then
                                    SwapData astrData(lngLow), astrData(lngHigh - 1)
                                End If
                           
                           Case Else   ' Case sensitive
                                If StrComp(astrData(lngLow), astrData(lngHigh - 1), vbBinaryCompare) = 1 Then
                                    SwapData astrData(lngLow), astrData(lngHigh - 1)
                                End If
                                
                    End Select  ' mlngCompareMethod
            
               Case eSort_Descending
                    Select Case mlngCompareMethod
                    
                           Case eSort_IgnoreCase   ' NOT case sensitive comparison
                                ' Uppercase and lowercase considered as same character
                                If StrComp(astrData(lngLow), astrData(lngHigh - 1), vbTextCompare) = -1 Then
                                    SwapData astrData(lngLow), astrData(lngHigh - 1)
                                End If

                           Case Else   ' Case sensitive
                                If StrComp(astrData(lngLow), astrData(lngHigh - 1), vbBinaryCompare) = -1 Then
                                    SwapData astrData(lngLow), astrData(lngHigh - 1)
                                End If
                    
                    End Select  ' mlngCompareMethod
        End Select  ' mlngSortDirection
        
        ' Get ending time
        strElapsedTime = ElapsedTime(GetTickCount() - lngStart)   ' Calc elapsed time
        
        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            InfoMsg "Error swapping string array data." & _
                    vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
            GoTo StringSorting_CleanUp
        End If
            
        GoTo StringSorting_CleanUp                       ' Exit this routine
    
    End If
    
    ' More than 2 items in array
    ReDim alngPointers(lngLow To lngHigh - 1)  ' Resize pointer array
    ReDim astrWork(lngLow To lngHigh - 1)      ' Resize temp data array
         
    ' Load temp work array with
    ' copy of original data
    Select Case mlngCompareMethod
    
           Case eSort_IgnoreCase   ' NOT case sensitive comparison
                For lngIndex = lngLow To lngHigh - 1
                    alngPointers(lngIndex) = lngIndex                 ' load pointer array
                    astrWork(lngIndex) = UCase$(astrData(lngIndex))   ' Insert data into temp string array
                Next lngIndex
    
           Case Else   ' Case sensitive comparison
                For lngIndex = lngLow To lngHigh - 1
                    alngPointers(lngIndex) = lngIndex         ' load pointer array
                    astrWork(lngIndex) = astrData(lngIndex)   ' Insert data into temp string array
                Next lngIndex
    End Select  ' mlngCompareMethod
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        InfoMsg "Error loading string arrays to be sorted." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        GoTo StringSorting_CleanUp
    End If
            
    ' Get starting time
    lngStart = CCur(GetTickCount)
        
    Select Case mlngSortMethod

           Case eShellSort    ' 0
                ShellSortString astrWork(), alngPointers()  ' Sort data
                
           Case eCombSort    ' 1
                CombSortString astrWork(), alngPointers()   ' Sort data
                
           Case eQuickSort    ' 2  TriQuickSort
                ' Read QuickSort flowerbox for additional information.
                QSortString astrWork(), alngPointers(), LBound(astrWork), UBound(astrWork)
                
                ' Use a modified Insertion Sort for final phase
                QSortStringFinal astrWork(), alngPointers(), LBound(astrWork), UBound(astrWork)
                
    End Select  ' mlngSortMethod
    
    ' Get ending time
    strElapsedTime = ElapsedTime(GetTickCount() - lngStart)   ' Calc elapsed time
    
    ' Verify no pointers were lost
    If (lngHigh - 1) <> UBound(alngPointers) Then
        gblnStopProcessing = True
    End If
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        InfoMsg "Error while sorting string data." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        GoTo StringSorting_CleanUp
    End If
            
    ' If return data is to be in descending
    ' order then reverse pointers
    If mlngSortDirection = eSort_Descending Then
        ReversePointers alngPointers()
    End If

    ReDim astrWork(lngLow To lngHigh)   ' Resize temp hold array
    
    ' Copy raw data to temp hold array
    For lngIndex = lngLow To lngHigh - 1
        astrWork(lngIndex) = astrData(lngIndex)
    Next lngIndex
    
    ReDim astrData(lngLow To lngHigh)   ' Resize original input array
       
    ' Reload original array with sorted data
    For lngIndex = lngLow To lngHigh - 1
        astrData(lngIndex) = astrWork(alngPointers(lngIndex))
    Next lngIndex
    
StringSorting_CleanUp:
    Erase astrWork()     ' Always empty arrays when not needed
    Erase alngPointers()
    
    On Error GoTo 0      ' Nullify this error trap
    Exit Sub

StringSorting_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume StringSorting_CleanUp
        
End Sub

' ***************************************************************************
' Routine:       ReversePointers
'
' Description:   Reverse the elements in the pointer array.
'
' Parameters:    alngPointers() - Array of values that match the number of
'                      elements that were sorted.
'
' Returns:       Array of pointers in reverse order
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 13-Aug-2001  Philippe Lord
'              http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=24546&lngWId=1
' 13-Feb-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Sub ReversePointers(ByRef alngPointers() As Long)

    ' Called by NumericSorting()
    '           StringSorting()
    
    Dim lngLow  As Long
    Dim lngHigh As Long
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Sub
    End If
            
    lngLow = LBound(alngPointers)
    lngHigh = UBound(alngPointers)
    
    ' If return data is to be in descending
    ' order then reverse pointers
    Do While lngLow < lngHigh
        
        ' Swap the pointers
        SwapLong alngPointers(lngLow), alngPointers(lngHigh)
        lngLow = lngLow + 1    ' increment low position
        lngHigh = lngHigh - 1  ' decrement high position
    Loop

End Sub

' **************************************************************************
' Routine:       ShellSortNumeric
'
' Description:   The performance of Shellsort depends on the increment
'                sequence. The essential point is that it is trivial to
'                substitute another set of increments for those used in
'                this code. We seek better increments no matter how
'                sophisticated the mathematical analysis or how extensive
'                the empirical studies we use, they can be quickly verified
'                and put to immediate practical use. Moreover, despite
'                extensive research on learning properties of different
'                families of increment sequences and on lower bounds on
'                performance, we still cannot deny the possibility that some
'                other set of increments might make this program run two or
'                three times faster, which would make it as fast or faster
'                than the best known sorting methods. This possibility makes
'                Shellsort worthy of close study.
'
'                This function performs a shellsort on an array of numeric
'                data while keeping track of the element indices. The indices
'                (pointers) will be returned to arrange the original data
'                in a sorted format.
'
' Parameters:    adblData - Array to be sorted
'                alngPointers() - Array of values that match the number of
'                      elements being sorted.
'
' Returns:       Array of pointers that depict the sorted locations within
'                the original data array.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 12-May-2005  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Sub ShellSortNumeric(ByRef adblData() As Double, _
                             ByRef alngPointers() As Long)

    ' Called by NumericSorting()
    
    Dim lngGap      As Long      ' temp ending point in array
    Dim lngHigh     As Long      ' number of elements in the array
    Dim lngLimit    As Long      ' temp bounds limitation
    Dim lngPointer  As Long      ' points to current string array element
    Dim lngLastSwap As Long      ' Location of last swap
    Dim blnSwapped  As Boolean   ' flag to denote when to stop sorting

    lngHigh = UBound(adblData)   ' Capture number of elements in array
    lngGap = lngHigh \ 2         ' First ending point is half way thru

    Do While lngGap > 0

        lngLimit = lngHigh - lngGap

        Do
            ' Preset flag to FALSE thus designating
            ' nothing has been swapped
            blnSwapped = False

            ' Compare elements and switch the ones out of order
            For lngPointer = 0 To lngLimit

                If adblData(alngPointers(lngPointer)) > adblData(alngPointers(lngPointer + lngGap)) Then
                
                    SwapLong alngPointers(lngPointer), alngPointers(lngPointer + lngGap)
                    lngLastSwap = lngPointer  ' Save last swap position
                    blnSwapped = True         ' Set flag to TRUE
                
                End If

            Next lngPointer

            ' Sort on next pass only to
            ' where last swap was made
            lngLimit = lngLastSwap - lngGap

        Loop While blnSwapped

        ' No switches at last offset,
        ' try one half as big
        lngGap = lngGap \ 2
    Loop

End Sub

' **************************************************************************
' Routine:       ShellSortString
'
' Description:   The performance of Shellsort depends on the increment
'                sequence. The essential point is that it is trivial to
'                substitute another set of increments for those used in
'                this code. We seek better increments no matter how
'                sophisticated the mathematical analysis or how extensive
'                the empirical studies we use, they can be quickly verified
'                and put to immediate practical use. Moreover, despite
'                extensive research on learning properties of different
'                families of increment sequences and on lower bounds on
'                performance, we still cannot deny the possibility that some
'                other set of increments might make this program run two or
'                three times faster, which would make it as fast or faster
'                than the best known sorting methods. This possibility makes
'                Shellsort worthy of close study.
'
'                This function performs a shellsort on an array of string
'                data while keeping track of the element indices. The indices
'                (pointers) will be returned to arrange the original data
'                in a sorted format.
'
' Parameters:    astrData() - Array to be sorted
'                alngPointers() - Array of values that match the number of
'                      elements being sorted.
'
' Returns:       Array of pointers that depict the sorted locations within
'                the original data array.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 12-May-2005  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Sub ShellSortString(ByRef astrData() As String, _
                            ByRef alngPointers() As Long)

    ' Called by StringSorting()
    
    Dim lngGap      As Long      ' temp ending point in array
    Dim lngHigh     As Long      ' number of elements in the array
    Dim lngLimit    As Long      ' temp bounds limitation
    Dim lngPointer  As Long      ' points to current string array element
    Dim lngLastSwap As Long      ' Location of last swap
    Dim blnSwapped  As Boolean   ' flag to denote when to stop sorting
    
    lngHigh = UBound(astrData)   ' Capture number of elements in array
    lngGap = lngHigh \ 2         ' First ending point is half way thru

    Do While lngGap > 0

        lngLimit = lngHigh - lngGap

        Do
            ' Preset flag to FALSE thus designating
            ' nothing has been swapped
            blnSwapped = False

            ' Compare elements and switch ones out of order:
            For lngPointer = 0 To lngLimit

                If StrComp(astrData(alngPointers(lngPointer)), _
                           astrData(alngPointers(lngPointer + lngGap)), _
                           mlngCompareMethod) = 1 Then
                    
                    SwapLong alngPointers(lngPointer), alngPointers(lngPointer + lngGap)
                    lngLastSwap = lngPointer  ' Save last swap position
                    blnSwapped = True         ' Set flag to TRUE
                End If

            Next lngPointer

            ' Sort on next pass only to
            ' where last swap was made
            lngLimit = lngLastSwap - lngGap

        Loop While blnSwapped

        ' No switches at last offset,
        ' try one half as big
        lngGap = lngGap \ 2
    Loop

End Sub

' ***************************************************************************
' Procedure:     CombSortNumeric
'
' Description:   CombSort is faster than all but QuickSort and close to
'                it.  On the other hand, the code is much simpler than
'                QuickSort and can be easily customized for any array type.
'                The CombSort was first published by Richard Box and Stephen
'                Lacey in the April 1991 issue of Byte magazine. They found
'                that using a sequence for the gaps that decreased by a shrink
'                factor of 1.3 gave the best results. Trial and error produced
'                the refinement that an eventual gap size of 11, (rather than
'                9 or 10) gave a more efficient sequence of gaps below that
'                point.
'
'                Every programmer knows that QuickSort is the best in-memory
'                sort in the world. Look inside the sort function of almost
'                any library and you'll find QuickSort. Another thing that
'                every programmer knows is that BubbleSort, the "hello world"
'                of sorts, is one of the slowest in the world and accurate.
'                But what most programmers don't know is that a simple
'                modification to bubble sort turns it into CombSort, a
'                remarkably simple sort that's nearly as fast as QuickSort.
'                It also does not slow down if the incoming data is almost
'                sorted like QuickSort.
'
'                CombSort starts out comparing items that are far apart.
'                Then it makes the gap smaller and does it again. In the
'                algorithm's last passes the gap is 1, making it act identical
'                to BubbleSort. That makes it easy to see that this algorithm
'                is correct, since we know that BubbleSort is correct and this
'                algorithm always turns into BubbleSort.
'
' Parameters:    adblData() - An array of unsorted data
'                alngPointers() - Array of values that match the number of
'                      elements being sorted.
'
' Returns:       Array of pointers that depict the sorted locations within
'                the original data array.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 21-May-2005  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 10-Feb-2008  Kenneth Ives kenaso@tx.rr.com
'              Made an assumption of the swapped flag being initialized to
'              FALSE.  Now, I manually set the flag prior to sorting.
' ***************************************************************************
Private Sub CombSortNumeric(ByRef adblData() As Double, _
                            ByRef alngPointers() As Long)
       
    ' Called by NumericSorting()
    
    Dim lngGap     As Long
    Dim lngLow     As Long
    Dim lngHigh    As Long
    Dim lngIndex   As Long
    Dim dblHold    As Double
    Dim blnSwapped As Boolean
        
    Const SHRINK_FACTOR As Double = 1.3
    
    lngLow = LBound(adblData)       ' Lowest number of elements
    lngHigh = UBound(adblData)      ' Highest number of elements
    lngGap = lngHigh - lngLow + 1   ' Total element count
    blnSwapped = False              ' Set swapped flag to FALSE
    dblHold = 0#                    ' Clear holding variable

    Do
        If lngGap > 1 Then
        
            ' Divide Gap by 1.3
            lngGap = Int(lngGap / SHRINK_FACTOR)
            
            ' Most efficient series of
            ' a final GAP starts with 11
            If (lngGap = 9) Or _
               (lngGap = 10) Then
               
                lngGap = 11
            End If
        
        End If
        
        ' Preset flag to FALSE thus designating
        ' nothing has been swapped
        blnSwapped = False
        
        ' Compare and possibly swap values
        ' and pointers separated by GAP
        For lngIndex = lngLow To lngHigh - lngGap
            
            ' Capture current data for comparison
            dblHold = adblData(lngIndex)
            
            ' Is hold data greater than current array data?
            If (dblHold > adblData(lngIndex + lngGap)) Then
            
                ' Move data to new position within array
                adblData(lngIndex) = adblData(lngIndex + lngGap)
                adblData(lngIndex + lngGap) = dblHold
                
                dblHold = 0#       ' Empty holding area
                blnSwapped = True  ' Reset swap flag
                
                ' Update pointers
                SwapLong alngPointers(lngIndex), alngPointers(lngIndex + lngGap)
                
            End If
           
        Next lngIndex
        
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit Do
        End If
    
    ' Loop until the GAP is less than one
    ' or nothing is swapped
    Loop Until lngGap = 1 And Not blnSwapped

End Sub

' ***************************************************************************
' Procedure:     CombSortString
'
' Description:   CombSortString is faster than all but QuickSort and close to
'                it.  On the other hand, the code is much simpler than
'                QuickSort and can be easily customized for any array type.
'                The CombSortString was first published by Richard Box and Stephen
'                Lacey in the April 1991 issue of Byte magazine. They found
'                that using a sequence for the gaps that decreased by a shrink
'                factor of 1.3 gave the best results. Trial and error produced
'                the refinement that an eventual gap size of 11, (rather than
'                9 or 10) gave a more efficient sequence of gaps below that
'                point.
'
'                Every programmer knows that QuickSort is the best in-memory
'                sort in the world. Look inside the sort function of almost
'                any library and you'll find QuickSort. Another thing that
'                every programmer knows is that BubbleSort, the "hello world"
'                of sorts, is one of the slowest in the world and accurate.
'                But what most programmers don't know is that a simple
'                modification to bubble sort turns it into CombSortString, a
'                remarkably simple sort that's nearly as fast as QuickSort.
'                It also does not slow down if the incoming data is almost
'                sorted like QuickSort.
'
'                CombSortString starts out comparing items that are far apart.
'                Then it makes the gap smaller and does it again. In the
'                algorithm's last passes the gap is 1, making it act identical
'                to BubbleSort. That makes it easy to see that this algorithm
'                is correct, since we know that BubbleSort is correct and this
'                algorithm always turns into BubbleSort.
'
' Parameters:    astrData - An array of unsorted data
'                alngPointers() - Array of values that match the number of
'                      elements being sorted.
'
' Returns:       Array of pointers that depict the sorted locations within
'                the original data array.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 21-May-2005  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 10-Feb-2008  Kenneth Ives kenaso@tx.rr.com
'              Made an assumption of the swapped flag being initialized to
'              FALSE.  Now, I manually set the flag prior to sorting.
' ***************************************************************************
Private Sub CombSortString(ByRef astrData() As String, _
                           ByRef alngPointers() As Long)
       
    ' Called by StringSorting()

    Dim lngGap     As Long
    Dim lngLow     As Long
    Dim lngHigh    As Long
    Dim lngIndex   As Long
    Dim strHold    As String
    Dim blnSwapped As Boolean
        
    Const SHRINK_FACTOR As Double = 1.3
    
    lngLow = LBound(astrData)       ' Lowest number of elements
    lngHigh = UBound(astrData)      ' Highest number of elements
    lngGap = lngHigh - lngLow + 1   ' Total element count
    blnSwapped = False              ' Set swapped flag to FALSE
    strHold = vbNullString          ' Clear holding variable

    Do
        If lngGap > 1 Then
        
            ' Divide Gap by 1.3
            lngGap = Int(lngGap / SHRINK_FACTOR)
            
            ' Most efficient series of
            ' a final GAP starts with 11
            If (lngGap = 9) Or _
               (lngGap = 10) Then
               
                lngGap = 11
            End If
        
        End If
        
        ' Preset flag to FALSE thus designating
        ' nothing has been swapped
        blnSwapped = False
        
        ' Compare and possibly swap values
        ' and pointers separated by GAP
        For lngIndex = lngLow To lngHigh - lngGap
            
            ' Capture current data for comparison
            strHold = astrData(lngIndex)
            
            ' Is hold data greater than current array data?
            If StrComp(strHold, astrData(lngIndex + lngGap), mlngCompareMethod) = 1 Then
            
                ' Move data to new position within array
                astrData(lngIndex) = astrData(lngIndex + lngGap)
                astrData(lngIndex + lngGap) = strHold
                
                strHold = vbNullString   ' Empty holding area
                blnSwapped = True        ' Reset swap flag
                
                ' Update pointers
                SwapLong alngPointers(lngIndex), alngPointers(lngIndex + lngGap)
                
            End If
           
        Next lngIndex
        
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit Do
        End If
    
    ' Loop until the GAP is less than one
    ' or nothing is swapped
    Loop Until lngGap = 1 And Not blnSwapped
    
End Sub

' ***************************************************************************
' Procedure:     QSortNumeric (TriQuickSort)
'
' Description:   Quicksort is a divide and conquer algorithm which relies on
'                a partition operation: to partition an array, we choose an
'                element, called a pivot, move all smaller elements before
'                the pivot, and move all greater elements after it. This can
'                be done efficiently in linear time and in-place. We then
'                recursively sort the lesser and greater sublists. Efficient
'                implementations of quicksort (with in-place partitioning)
'                are typically unstable sorts and somewhat complex, but are
'                among the fastest sorting algorithms in practice. Together
'                with its modest O(log n) space usage, this makes quicksort
'                one of the most popular sorting algorithms, available in
'                many standard libraries.
'
'                The QuickSort routine also provides pretty well the fastest
'                sort amongst the common sort methods - BubleSort, ShellSort,
'                MergeSort, InsertionSort and QuickSort. Its operation is
'                fairly straight forward: an item is chosen midway between two
'                points in array() (called the pivot point). One value that
'                is higher and one value that is lower than the pivot point
'                is found and swapped, with the function calling itself
'                passing new start and stop points. The method continues in
'                this fashion until no elements meet these conditions, at
'                which point the array is sorted.
'
'                The QuickSort algorithms are generally fast, but there
'                exists an exception.  When the array is *nearly* sorted,
'                QuickSort can be slow (up to 2 times slower).
'
'                ------------------------------------------------------------
'                Comments below are excerts from Philippe Lord's source code
'                module.  Philippe is the creator of the TriQuickSort
'                algorithm.
'
'                Since the TriQuickSort algorithm is in no way a standard
'                sort algorithm, I will try and explain it here. First, I
'                must say that the main idea started from Sun Microsystems,
'                in java source code form. I found Sun's source code after a
'                search on the internet for the 'fastest' sort algorithm
'                (considering a uniprocessor configuration and a
'                nearly-sorted OR totally unsorted array). I compared the
'                ones that performed the best, and soon enough stumbled upon
'                this one (Sun's one). Generally speaking, it was a 3-median
'                QuickSort, a little twinked, of course. The 3-median
'                QuickSort has the advantage of not suffering standard
'                2-median QuickSort's problems concerning nearly-sorted
'                arrays.  It performed very well, but there was a side-note
'                suggesting using a second algorithm when the number of
'                iterations came low, under 10.  I though about it, and
'                understood why they suggested that.
'
'                After porting their 3-median QuickSort, I made it stop when
'                the delta (difference) of the 2 bounds came under 10, like
'                suggested by Sun Microsystems. Now, if you understood my
'                explanation of the QuickSort algorithm, you should
'                understand too that stopping the process at delta 10 means
'                all you have to do after QuickSorting is to sort each
'                sections of 10, without needing to do any compares with
'                anything else other than the 10 entries you're processing.
'                Imagine just that the cutted-QuickSort sorts generally, but
'                you need to finish the work off by processing packets of 10
'                entries.
'
'                Stopping the recursion tree using a delta 10 does not mean
'                in any way that you are going to have sections exactly of
'                ten well placed entries.  In fact, if you think about the
'                problem, and if you understand the QuickSort algorithm,
'                it means that your sections can vary from your input delta
'                (10) up to 2x delta -1 (19).
'
'                I searched for the fastest algorithm for processing small
'                arrays, and had in mind to call it n times, where n equals
'                the number of sections. You must keep in mind that to have a
'                good sort algorithm working on 10-19 entries it means your
'                algorithm has to be as simple as possible, because you
'                cannot even afford to do simple mathematical operations. You
'                just need something simple and fast.
'
'                I found InsertionSort to be the perfect algorithm to
'                continue the cutted-QuickSort's job.   But, like I said
'                earlier, delta 10 (which gives us a 10-19 section's range)
'                would be like saying:  "Hey, let's give out the main job to
'                InsertionSort ".  This is normally slow, but in our case it"
'                is very fast.  I lowered the QuickSort's delta to a default
'                value of 4.  This gives a section ranging from 4 to 7 in
'                length, which  gives very good results. I do not recommend
'                you put a lower value to it, because QuickSort would eat up
'                too much memory and start getting slow. If you put higher
'                than 4, the reverse happens.  You get a much lower
'                performance because InsertionSort starts bottlenecking a
'                little too much.
'
' Parameters:    adblData() - An array of unsorted data
'                alngPointers() - Array of values that match the number of
'                      elements being sorted.
'                lngLow - Lowest element position in the array
'                lngHigh - Highest element position in the array
'
' Returns:       Array of pointers that depict the sorted locations within
'                the original data array.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 13-Aug-2001  Philippe Lord
'              http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=24546&lngWId=1
' 13-Feb-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Sub QSortNumeric(ByRef adblData() As Double, _
                         ByRef alngPointers() As Long, _
                         ByVal lngLow As Long, _
                         ByVal lngHigh As Long)

    ' Called by NumericSorting()
    
    Dim dblHold      As Double  ' Temp storage area
    Dim lngLowPivot  As Long    ' Temp low pivot pointer
    Dim lngHighPivot As Long    ' Temp high pivot pointer

    Const DELTA As Long = 4     ' DO NOT CHANGE THIS VALUE.  Read flower box above.

    ' See if there is still data above
    ' the mandatory break point
    If (lngHigh - lngLow) > DELTA Then
    
        ' Calc low end pivot position
        lngLowPivot = (lngHigh + lngLow) / 2
        
        ' Test to see if the data in the lower
        ' area is greater than data in the low
        ' pivot area
        If adblData(alngPointers(lngLow)) > adblData(alngPointers(lngLowPivot)) Then
            SwapLong alngPointers(lngLow), alngPointers(lngLowPivot)
        End If
        
        ' Test to see if the data in the lower area
        ' is greater than data in the high area
        If adblData(alngPointers(lngLow)) > adblData(alngPointers(lngHigh)) Then
            SwapLong alngPointers(lngLow), alngPointers(lngHigh)
        End If
        
        ' Test to see if the data in the lower pivot
        ' area is greater than data in the high area
        If adblData(alngPointers(lngLowPivot)) > adblData(alngPointers(lngHigh)) Then
            SwapLong alngPointers(lngLowPivot), alngPointers(lngHigh)
        End If
        
        lngHighPivot = lngHigh - 1                                      ' Decrement high pivot position
        SwapLong alngPointers(lngLowPivot), alngPointers(lngHighPivot)  ' Swap pointers
        lngLowPivot = lngLow                                            ' Update low pivot position
        
        ' Temporarily store array data
        ' from temp high pivot location
        dblHold = adblData(alngPointers(lngHighPivot))
    
        Do
            ' Increment the temp low pivot position
            ' while current array data is less than
            ' the temp stored pivot data
            Do
                lngLowPivot = lngLowPivot + 1
            Loop While adblData(alngPointers(lngLowPivot)) < dblHold
            
            ' Decrement the temp high pivot position
            ' while current array data is greater than
            ' the temp stored pivot data
            Do
                lngHighPivot = lngHighPivot - 1
            Loop While adblData(alngPointers(lngHighPivot)) > dblHold
            
            ' Exit loop if temp low pivot value
            ' exceeds temp high pivot value
            If lngHighPivot < lngLowPivot Then
                Exit Do
            End If
            
            ' Swap pointers
            SwapLong alngPointers(lngLowPivot), alngPointers(lngHighPivot)
        Loop
        
        ' Swap pointers
        SwapLong alngPointers(lngLowPivot), alngPointers(lngHigh - 1)
        
        ' Perform a recursive call to see if the minimum number of elements
        ' in the array is LESS than the temp high end.  Always sort the low
        ' end of the array first.
        QSortNumeric adblData(), alngPointers(), lngLow, lngHighPivot
        
        ' Perform a recursive call to see if the temp low end is LESS than the
        ' maximum number of elements in the array.  The high end is always
        ' sorted last.
        QSortNumeric adblData(), alngPointers(), lngLowPivot + 1, lngHigh
    
    End If
   
End Sub

' ***************************************************************************
' Procedure:     QSortString  (TriQuickSort)
'
' Description:   Quicksort is a divide and conquer algorithm which relies on
'                a partition operation: to partition an array, we choose an
'                element, called a pivot, move all smaller elements before
'                the pivot, and move all greater elements after it. This can
'                be done efficiently in linear time and in-place. We then
'                recursively sort the lesser and greater sublists. Efficient
'                implementations of quicksort (with in-place partitioning)
'                are typically unstable sorts and somewhat complex, but are
'                among the fastest sorting algorithms in practice. Together
'                with its modest O(log n) space usage, this makes quicksort
'                one of the most popular sorting algorithms, available in
'                many standard libraries.
'
'                The QuickSort routine also provides pretty well the fastest
'                sort amongst the common sort methods - BubleSort, ShellSort,
'                MergeSort, InsertionSort and QuickSort. Its operation is
'                fairly straight forward: an item is chosen midway between two
'                points in array() (called the pivot point). One value that
'                is higher and one value that is lower than the pivot point
'                is found and swapped, with the function calling itself
'                passing new start and stop points. The method continues in
'                this fashion until no elements meet these conditions, at
'                which point the array is sorted.
'
'                The QuickSort algorithms are generally fast, but there
'                exists an exception.  When the array is *nearly* sorted,
'                QuickSort can be slow (up to 2 times slower).
'
'                ------------------------------------------------------------
'                Comments below are excerts from Philippe Lord's source code
'                module.  Philippe is the creator of the TriQuickSort
'                algorithm.
'
'                Since the TriQuickSort algorithm is in no way a standard
'                sort algorithm, I will try and explain it here. First, I
'                must say that the main idea started from Sun Microsystems,
'                in java source code form. I found Sun's source code after a
'                search on the internet for the 'fastest' sort algorithm
'                (considering a uniprocessor configuration and a
'                nearly-sorted OR totally unsorted array). I compared the
'                ones that performed the best, and soon enough stumbled upon
'                this one (Sun's one). Generally speaking, it was a 3-median
'                QuickSort, a little twinked, of course. The 3-median
'                QuickSort has the advantage of not suffering standard
'                2-median QuickSort's problems concerning nearly-sorted
'                arrays.  It performed very well, but there was a side-note
'                suggesting using a second algorithm when the number of
'                iterations came low, under 10.  I though about it, and
'                understood why they suggested that.
'
'                After porting their 3-median QuickSort, I made it stop when
'                the delta (difference) of the 2 bounds came under 10, like
'                suggested by Sun Microsystems. Now, if you understood my
'                explanation of the QuickSort algorithm, you should
'                understand too that stopping the process at delta 10 means
'                all you have to do after QuickSorting is to sort each
'                sections of 10, without needing to do any compares with
'                anything else other than the 10 entries you're processing.
'                Imagine just that the cutted-QuickSort sorts generally, but
'                you need to finish the work off by processing packets of 10
'                entries.
'
'                Stopping the recursion tree using a delta 10 does not mean
'                in any way that you are going to have sections exactly of
'                ten well placed entries.  In fact, if you think about the
'                problem, and if you understand the QuickSort algorithm,
'                it means that your sections can vary from your input delta
'                (10) up to 2x delta -1 (19).
'
'                I searched for the fastest algorithm for processing small
'                arrays, and had in mind to call it n times, where n equals
'                the number of sections. You must keep in mind that to have a
'                good sort algorithm working on 10-19 entries it means your
'                algorithm has to be as simple as possible, because you
'                cannot even afford to do simple mathematical operations. You
'                just need something simple and fast.
'
'                I found InsertionSort to be the perfect algorithm to
'                continue the cutted-QuickSort's job.   But, like I said
'                earlier, delta 10 (which gives us a 10-19 section's range)
'                would be like saying:  "Hey, let's give out the main job to
'                InsertionSort ".  This is normally slow, but in our case it"
'                is very fast.  I lowered the QuickSort's delta to a default
'                value of 4.  This gives a section ranging from 4 to 7 in
'                length, which  gives very good results. I do not recommend
'                you put a lower value to it, because QuickSort would eat up
'                too much memory and start getting slow. If you put higher
'                than 4, the reverse happens.  You get a much lower
'                performance because InsertionSort starts bottlenecking a
'                little too much.
'
' Parameters:    astrData() - An array of unsorted data
'                alngPointers() - Array of values that match the number of
'                      elements being sorted.
'                lngLow - Lowest element position in the array
'                lngHigh - Highest element position in the array
'
' Returns:       Array of pointers that depict the sorted locations within
'                the original data array.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 13-Aug-2001  Philippe Lord
'              http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=24546&lngWId=1
' 13-Feb-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Sub QSortString(ByRef astrData() As String, _
                        ByRef alngPointers() As Long, _
                        ByVal lngLow As Long, _
                        ByVal lngHigh As Long)
    
    ' Called by StringSorting()
    
    Dim strHold      As String  ' Temp storage area
    Dim lngLowPivot  As Long    ' Temp low pivot pointer
    Dim lngHighPivot As Long    ' Temp high pivot pointer

    Const DELTA As Long = 4     ' DO NOT CHANGE THIS VALUE.  Read flower box above.

    ' See if there is still data above
    ' the mandatory break point
    If (lngHigh - lngLow) > DELTA Then
    
        ' Calc low end pivot position
        lngLowPivot = (lngHigh + lngLow) / 2
        
        ' Test to see if the data in the lower
        ' area is greater than data in the low
        ' pivot area
        If StrComp(astrData(alngPointers(lngLow)), astrData(alngPointers(lngLowPivot)), mlngCompareMethod) = 1 Then
            SwapLong alngPointers(lngLow), alngPointers(lngLowPivot)
        End If
        
        ' Test to see if the data in the lower area
        ' is greater than data in the high area
        If StrComp(astrData(alngPointers(lngLow)), astrData(alngPointers(lngHigh)), mlngCompareMethod) = 1 Then
            SwapLong alngPointers(lngLow), alngPointers(lngHigh)
        End If
        
        ' Test to see if the data in the lower pivot
        ' area is greater than data in the high area
        If StrComp(astrData(alngPointers(lngLowPivot)), astrData(alngPointers(lngHigh)), mlngCompareMethod) = 1 Then
            SwapLong alngPointers(lngLowPivot), alngPointers(lngHigh)
        End If
        
        lngHighPivot = lngHigh - 1                                      ' Decrement high pivot position
        SwapLong alngPointers(lngLowPivot), alngPointers(lngHighPivot)  ' Swap pointers
        lngLowPivot = lngLow                                            ' Update low pivot position
        
        ' Temporarily store array data
        ' from high pivot location
        CopyMemory ByVal VarPtr(strHold), ByVal VarPtr(astrData(alngPointers(lngHighPivot))), 4&
    
        Do
            ' Increment the temp low pivot position
            ' while current array data is less than
            ' the temp stored pivot data
            Do
                lngLowPivot = lngLowPivot + 1
            Loop While astrData(alngPointers(lngLowPivot)) < strHold
            
            ' Decrement the temp high pivot position
            ' while current array data is greater than
            ' the temp stored pivot data
            Do
                lngHighPivot = lngHighPivot - 1
            Loop While astrData(alngPointers(lngHighPivot)) > strHold
            
            ' Exit loop if temp low pivot value
            ' exceeds temp high pivot value
            If lngHighPivot < lngLowPivot Then
                Exit Do
            End If
            
            ' Swap pointers
            SwapLong alngPointers(lngLowPivot), alngPointers(lngHighPivot)
        Loop
        
        ' Swap pointers
        SwapLong alngPointers(lngLowPivot), alngPointers(lngHigh - 1)
        
        ' Perform a recursive call to see if the minimum number of elements
        ' in the array is LESS than the temp high end.  Always sort the low
        ' end of the array first.
        QSortString astrData(), alngPointers(), lngLow, lngHighPivot
        
        ' Perform a recursive call to see if the temp low end is LESS than the
        ' maximum number of elements in the array.  The high end is always
        ' sorted last.
        QSortString astrData(), alngPointers(), lngLowPivot + 1, lngHigh
    
    End If
   
    ' clear temp stroage area
    lngLowPivot = 0
    CopyMemory ByVal VarPtr(strHold), ByVal VarPtr(lngLowPivot), 4&

End Sub

' ***************************************************************************
' Routine:       QSortNumericFinal
'
' Description:   This is a modified Insertion Sort.  It is called by
'                Quicksort to supplement the final portion of data sorting.
'                This is a simple sorting algorithm, a comparison sort in
'                which the sorted array (or list) is built one entry at a
'                time. It is much less efficient on large lists than more
'                advanced algorithms such as quicksort, heapsort, or merge
'                sort, but it has various advantages.  Insertion sort
'                is simple to implement, efficient on small data sets,
'                efficient on data sets which are already substantially
'                sorted, and more efficient in practice than most other
'                simple algorithms such as selection sort or bubble sort.
'
'                Insertion sort is very similar to selection sort. Just like
'                in selection sort, after k passes through the array, the
'                first k elements are in sorted order. For selection sort,
'                these are the k smallest elements, while in insertion sort
'                they are whatever the first k elements were in the unsorted
'                array. Insertion sort's advantage is that it only scans as
'                many elements as it needs to in order to place the k + 1st
'                element, while selection sort must scan all remaining
'                elements to find the absolute smallest element.
'
' Parameters:    adblData() - Unsorted array of string data
'                alngPointers() - Array of values that match the number of
'                      elements that were sorted.
'                lngLow - Lowest element position in the array
'                lngHigh - Highest element position in the array
'
' Returns:       Array of sorted pointers
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 13-Aug-2001  Philippe Lord
'              http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=24546&lngWId=1
' 13-Feb-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Sub QSortNumericFinal(ByRef adblData() As Double, _
                              ByRef alngPointers() As Long, _
                              ByVal lngLow As Long, _
                              ByVal lngHigh As Long)
                                
    ' Called by NumericSorting()  [after QSortNumeric has finished]
    
    Dim lngIdx     As Long
    Dim lngIndex   As Long
    Dim lngPointer As Long
    
    For lngIndex = lngLow + 1 To lngHigh
        
        lngPointer = alngPointers(lngIndex)   ' Temp store pointer position
        lngIdx = lngIndex                     ' starting position for pointers
        
        Do While lngIdx > lngLow
        
            ' If the indexed data is equal to or less than the
            ' data in the previous position then exit this loop
            If adblData(alngPointers(lngIdx - 1)) <= adblData(lngPointer) Then
                Exit Do
            End If
            
            alngPointers(lngIdx) = alngPointers(lngIdx - 1)  ' Update pointer position
            lngIdx = lngIdx - 1                              ' decrement pointer index
        Loop
        
        alngPointers(lngIdx) = lngPointer   ' Update pointer
    
    Next lngIndex

End Sub

' ***************************************************************************
' Routine:       QSortStringFinal
'
' Description:   This is a modified Insertion Sort.  It is called by
'                Quicksort to supplement the final portion of data sorting.
'                This is a simple sorting algorithm, a comparison sort in
'                which the sorted array (or list) is built one entry at a
'                time. It is much less efficient on large lists than more
'                advanced algorithms such as quicksort, heapsort, or merge
'                sort, but it has various advantages.  Insertion sort
'                is simple to implement, efficient on small data sets,
'                efficient on data sets which are already substantially
'                sorted, and more efficient in practice than most other
'                simple algorithms such as selection sort or bubble sort.
'
'                Insertion sort is very similar to selection sort. Just like
'                in selection sort, after k passes through the array, the
'                first k elements are in sorted order. For selection sort,
'                these are the k smallest elements, while in insertion sort
'                they are whatever the first k elements were in the unsorted
'                array. Insertion sort's advantage is that it only scans as
'                many elements as it needs to in order to place the k + 1st
'                element, while selection sort must scan all remaining
'                elements to find the absolute smallest element.
'
' Parameters:    astrData() - Unsorted array of string data
'                alngPointers() - Array of values that match the number of
'                      elements that were sorted.
'                lngLow - Lowest element position in the array
'                lngHigh - Highest element position in the array
'
' Returns:       Array of sorted pointers
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 13-Aug-2001  Philippe Lord
'              http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=24546&lngWId=1
' 13-Feb-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Sub QSortStringFinal(ByRef astrData() As String, _
                             ByRef alngPointers() As Long, _
                             ByVal lngLow As Long, _
                             ByVal lngHigh As Long)
                                
    ' Called by StringSorting()  [after QSortString has finished]
    
    Dim lngIdx     As Long
    Dim lngIndex   As Long
    Dim lngPointer As Long
    
    For lngIndex = lngLow + 1 To lngHigh
        
        lngPointer = alngPointers(lngIndex)   ' Temp store pointer position
        lngIdx = lngIndex                     ' starting position for pointers
        
        Do While lngIdx > lngLow
        
            ' If the indexed data is equal to or less
            ' than the data in the previous position
            ' then exit this loop
            If StrComp(astrData(alngPointers(lngIdx - 1)), astrData(lngPointer), mlngCompareMethod) <= 0 Then
                Exit Do
            End If
            
            alngPointers(lngIdx) = alngPointers(lngIdx - 1)  ' Update pointer position
            lngIdx = lngIdx - 1                              ' decrement pointer index
        Loop
        
        alngPointers(lngIdx) = lngPointer   ' Update pointer
    
    Next lngIndex

End Sub

' ***************************************************************************
' Routine:       SwapData
'
' Description:   Swap data with each other.  I wrote this function since
'                BASIC stopped having its own SWAP function.  I use this
'                for swapping strings, type structures, numbers with
'                decimal values, etc.
'
' Parameters:    vntValue1 - Incoming data to be swapped with Value2
'                vntValue2 - Incoming data to be swapped with Value1
'
' Returns:       Swapped data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 09-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Private Sub SwapData(ByRef vntData1 As Variant, _
                     ByRef vntData2 As Variant)

    '   Called by NumericSorting()
    '             StringSorting()

    Dim vntHold As Variant

    vntHold = Empty   ' Start with empty variants (prevents crashing)

    vntHold = vntData1
    vntData1 = vntData2
    vntData2 = vntHold

    vntHold = Empty   ' Always empty variants (prevents crashing)

End Sub

' ***************************************************************************
' Routine:       SwapLong
'
' Description:   Swap data (Byte, Integer, Long) with each other without
'                using a temp.  Using Math to do this takes a substantial
'                amount more processing than Logic Gates. The logic gates
'                foundation is in all processors. Which you could argue
'                that so is Math, but running a math process with a temp
'                variable in fact uses a considerable amount more processor
'                cycles than logic gates.
'
'                Using 3 Xor's over 3 Basic Equations greatly speeds up the
'                application. (especially when dealing with millions of
'                equations/Logic Gates)
'
'                This routine works with byte, integer and long values.
'                Change the parameter data type accordingly.
'
' Parameters:    AA - data to be swapped with Value2
'                BB - data to be swapped with Value1
'
' Returns:       Swapped data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Mar-2003  Ulli
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=43903&lngWId=1
' ***************************************************************************
Private Sub SwapLong(ByRef AA As Long, _
                     ByRef BB As Long)

    ' Called by ShellSortNumeric()
    '           ShellSortString()
    '           CombSortNumeric()
    '           CombSortString()
    '           QSortNumeric()
    '           QSortString()
    '           ReversePointers()

    AA = AA Xor BB
    BB = BB Xor AA
    AA = AA Xor BB

End Sub

' ***************************************************************************
' Routine:       ElapsedTime
'
' Description:   Formats time display
'
' Reference:     Karl E. Peterson, http://vb.mvps.org/
'
' Parameters:    lngMilliseconds - Time in milliseconds
'
' Returns:       Formatted output
'                01:23:45.678 = 1 hour 23 minutes 45 seconds 678 thousandths
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Function ElapsedTime(ByVal lngMilliseconds As Long) As String

    ' Called by NumericSorting()
    '           StringSorting()
    
    ElapsedTime = vbNullString   ' Verify output string is empty

    ' Format output string as HH:MM:SS
    ElapsedTime = Format$(DateAdd("s", (lngMilliseconds \ 1000), #12:00:00 AM#), "HH:MM:SS")
    
    ' Calc number of milliseconds left
    lngMilliseconds = lngMilliseconds - ((lngMilliseconds \ 1000) * 1000)
    
    ' Append thousandths to output string
    ElapsedTime = ElapsedTime & "." & Format$(lngMilliseconds, "000")
   
End Function

Private Sub Class_Initialize()
    
    ' Whenever a class is instantiated (activated), this routine
    ' is called automatically if it exist with code inside.
    ' Ex:  Set objSort = new cSort
                    
    Set mobjCDT = Nothing  ' Verify class object is not active
    
    ' Preset property values
    StopProcessing = False                ' Do not stop processing
    TypeOfData = eSort_String             ' Data is in string format
    SortMethod = eCombSort                ' Use Comb sort
    SortDirection = eSort_Ascending       ' Sort in ascending order
    CompareMethod = eSort_CaseSensitive   ' Ignored unless sorting string data
    
    ' Properties for sorting
    ' dates and/or time
    ProcessTime = True     ' Ignored unless sorting dates
    DateFormat = 0         ' Ignored unless sorting dates
    TimeFormat = 0         ' Ignored unless sorting dates
    
End Sub

Private Sub Class_Terminate()

    ' Whenever a class is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
    ' Ex:  Set objSort = new cSort
            
    Set mobjCDT = Nothing  ' Verify class object is freed from memory
                
End Sub
