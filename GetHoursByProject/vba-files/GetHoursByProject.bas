Attribute VB_Name = "GetHoursByProject"

Option Explicit
' Global variable here
Public logFile As String

Public counter As Integer
Public previousCounter As Integer

Public pYear As Integer         ' Year project was worked on
Public pIndex As Integer        ' Project index number (for arrays)
Public pId As String            ' Project identification number
Public pDate As Date            ' Date project was worked on
Public pCalcTime As Long        ' Time worked on project calculated (difference between start and end time)
Public pTotalTime As Double     ' Time worked on project including breaks added together
Public pWeekNumber As Integer   ' Week the project was worked on

Sub GetHoursByProject()
	'
	' Add together how much time has been spend on projects
    '
	Call deleteLog()

    Dim lastCell As Range
    ' Get last cell (with content)
    Set lastCell = Sheets(ActiveSheet.name).Cells(Rows.Count, "A").End(xlUp)
    
    Dim size As Integer
    size = 4
    ' Create array
    Dim results() As Variant
    ' Assign dynamic size
    ReDim results(lastCell.Row, size)
    ' Counter
    counter = 0

    Dim i As Integer
    Dim totalTime As Variant
    Dim calculatedTime As String
    
    For i = 1 To lastCell.Row
        Dim cell As String
        ' Get value of cell and trim empty space
        cell = Trim(Cells(i, 1).value)
        ' Check that cell isn't empty
        If Not Len(cell) = 0 Then
            Dim projectLetter As String
            Dim projectNumbers As String
            
            ' Get first character of Cell
            projectLetter = Left(cell, 1)
            ' Get the four characters that are behind the first character
            projectNumbers = Right(Left(cell, 5), 4)

            Dim dateResult As Variant
            dateResult = CheckIfDate(cell)

            ' Check if Cell contains date
            if dateResult(0) = True Then
                pDate = dateResult(1)
            ' Check if Cell contains project
            ' Check if first character is alphabetical and the following 4 characters are numbers
            ElseIf IsAlpha(projectLetter) And IsNumeric(projectNumbers) Then
                ' Add project letter and numbers to create project ID
                pId = projectLetter + projectNumbers
                
                Dim startTime As String
                Dim endTime As String
                
                ' Total written down time including breaks
                totalTime = Cells(i, "E").value
                startTime = ReplaceColon(Cells(i, "C").Text)
                endTime = ReplaceColon(Cells(i, "D").Text)
                ' Total time calculated without breaks
                calculatedTime = CalculateTimeDifference(startTime, endTime) ' Format(, "hh:mm")
                
                results(counter, 0) = pDate
                results(counter, 1) = pId
                results(counter, 2) = CDbl(calculatedTime)
                results(counter, 3) = CDbl(totalTime)

                counter = counter + 1
            End If
        End If
    Next i

    ' Remove extra element from counter
    previousCounter = counter - 1
    counter = 0

    Dim column As String
    column = "k"

    Dim rowCount As Integer

    rowCount = DisplayTotalTime(results, previousCounter, size, column)
    rowCount = rowCount + 2

    Dim filteredResult() As Variant

    filteredResult = FilterByYear(results, previousCounter)
    Call DisplayYearlyTime(filteredResult, previousCounter, column, rowCount)

    Exit Sub ' ---------------------------------------------------------------------------
        
    ' Dim b As Integer
    ' ' Loop through array
    ' For b = LBound(results) To UBound(results)
        
    '     pId = results(b, 0)
    '     calculatedTime = results(b, 1)
    '     totalTime = results(b, 2)
        
    '     ' Skip empty array items
    '     If Len(pId) > 0 And Len(calculatedTime) > 0 And Len(totalTime) > 0 Then
    '         Dim cellCounter
    '         ' Since we have two headers above move three rows down
    '         cellCounter = b + 3
            
    '         Cells(cellCounter, "K").value = pId
    '         With Cells(cellCounter, "L")
    '             .value = calculatedTime / 3600 ' Devide seconds into hours
    '             .NumberFormat = "0.0"
    '         End With
    '         With Cells(cellCounter, "N")
    '             .value = totalTime
    '             .NumberFormat = "0.00"
    '         End With
            
    '         'If you want to convert the decmial times to full times:
    '         'For example, let's convert 12.675 hours to hours, minutes, and seconds.
    '         '
    '         'Start by finding the number of hours.
    '         '12.675 hours = 12 hours + .675 hours
    '         'Full hours = 12
    '         '
    '         'Then find the number of minutes
    '         'minutes = .675 hours and 60 minutes
    '         'minutes = 40.5 minutes
    '         'Full minutes = 40
    '         '
    '         'Find the remaining seconds
    '         'seconds = .5 minutes and 60 seconds
    '         'seconds = 30 seconds
    '         '
    '         'Finally, rewrite as HH:MM:SS
    '         'time = 12:40:30
    '     End If
    ' Next b
End Sub

' ↓ Logging functions ----------------------------------------- ↓

Function GetFilePath() As String
    Dim path As String
    Dim filename As String
    Dim oWSHShell As Object
    
    ' Get path of desktop folder
    Set oWSHShell = CreateObject("WScript.Shell")
    path = oWSHShell.SpecialFolders("Desktop")
    ' File of log file
    filename = "logfile.txt"
    GetFilePath = path & "\" & filename
End Function

Function deleteLog()
	logFile = GetFilePath()
    If FileExists(logFile) Then
		Kill logFile
	End If
End Function

' Print debug content to file
Sub outputToFile(message As String, content As Variant)
    Dim result As String
    
    ' Variable containing content that needs to be written
	if Len(message) = 0 And Len(content) = 0 Then
		result = ""
	Else
	    result = message & ": " & content
	End If

    logFile = GetFilePath()

    If FileExists(logFile) Then
        ' File exists
        Dim TextFile As Integer
        
        'Determine the next file number available for use by the FileOpen function
        TextFile = FreeFile
        
        'Open the text file
        Open logFile For Append As TextFile
        
        'Write some lines of text
        Print #TextFile, result
        'Save & Close Text File
        Close TextFile
    Else
        ' File doesn't exist
		Dim fs As Object
		Dim a As Object
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set a = fs.CreateTextFile(logFile, True)
        a.WriteLine (result)
        a.Close
    End If
End Sub

' Check if file exists already
Function FileExists(ByRef strFileName As String) As Boolean
' TRUE if the argument is an existing file
' works with Unicode file names
    On Error Resume Next
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    FileExists = objFSO.FileExists(strFileName)
    On Error GoTo 0
End Function

Function VarTypeName(value As Variant) As String
    Select Case VarType(value)
        Case 0
            VarTypeName = "Empty"
        Case 1
            VarTypeName = "Null"
        Case 2
            VarTypeName = "Integer"
        Case 3
            VarTypeName = "Long"
        Case 4
            VarTypeName = "Single"
        Case 5
            VarTypeName = "Double"
        Case 6
            VarTypeName = "Currency"
        Case 7
            VarTypeName = "Date"
        Case 8
            VarTypeName = "String"
        Case 9
            VarTypeName = "Object"
        Case 10
            VarTypeName = "Error"
        Case 11
            VarTypeName = "Boolean"
        Case 12
            VarTypeName = "Variant"
        Case 13
            VarTypeName = "Data object"
        Case 14
            VarTypeName = "Decimal"
        Case 17
            VarTypeName = "Byte"
        Case 20
            VarTypeName = "LongLong"
        Case 36
            VarTypeName = "User defined"
        Case 8192
            VarTypeName = "Array"
        Case Else
            if VarType(value) > 8192 Then
                VarTypeName = "Specific array type"
            Else
                VarTypeName = "Unknown"
            End If
    End Select 
End Function

' ↑ Logging functions ----------------------------------------- ↑

Function IsAlpha(s) As Boolean
    ' Check if value is an alphabetical character
    IsAlpha = Len(s) And Not s Like "*[!a-zA-Z]*"
End Function

' Replace dot with colon
Function ReplaceColon(value) As String
    ReplaceColon = Replace(value, ".", ":")
End Function

Function Months() As Variant
    Months = [{ "januari", 1; "februari", 2; "maart", 3; "april", 4; "mei", 5; "juni", 6; "juli", 7; "augustus", 8; "september", 9; "oktober", 10; "november", 11; "december", 12 }]
End Function


' Check if value provided is a date
Function CheckIfDate(value As String) As Variant
    value = Trim(value)
    Dim result As Boolean
    Dim dateResult As Date

    ' Check if value is date
    ' Date examples:
    ' 28/3/2010
    ' 09 October 2012
    If IsDate(value) Then
        result = True
        dateResult = value
    Else
        Dim strInArray As Variant
        Dim position As Integer
        Dim length As Integer

        ' Check if value contains a month (in Dutch)
        strInArray = CheckIfStringContainsArrayElement(value, Months(), 1)
        position = strInArray(0)
        length = strInArray(1)

        ' If no month can be found in value make result false
        If position = -1 Then
            result = False
        Else
            ' If month can be found in value extract date from value
            ' Example date: Dinsdag 24 Maart 2022
            ' Date needs to be in Dutch for this
            result = True
            Dim dayLong As String
            Dim dayShort As Integer
            Dim monthLong As String
            Dim monthShort As Integer
            Dim year As Integer

            ' Extract long day. Example: Dinsdag 12
            dayLong = Left(value, position - 2)
            ' Extract short day. Example: 12
            dayShort = CInt(Right(dayLong, 2))
            ' Extract long month. Example: April
            monthLong = Mid(value, position, length)
            ' Extract short month. Example: 4
            monthShort = strInArray(2)
            ' Extract year
            year = CInt(Right(value, 4))
            ' Pass date to variable with Date type
            dateResult = DateSerial(year, monthShort, dayShort)
        End If
    End If
    CheckIfDate = Array(result, dateResult)
End Function

Function FilterByYear(arr As Variant, count As Integer) As Variant
    ' Array for storing final values. Except it's initializer size values are too high
    ' We'll determine the correct size after storing everything
    Dim allProjects() As Variant
    ReDim allProjects(count, count + 1, 5)

    ' Array containing only the different years found in the spreadsheet
    Dim yearArr() As Variant
    ReDim yearArr(count)

    ' Counters for the for loops
    Dim i As Integer
    Dim o As Integer
    ' 
    Dim a As Integer

    Dim yearCounter As Integer
    Dim projectCounter As Integer

    ' If year is found in array assign it's position in the arry to this variable
    Dim yearInArray As Integer

    ' Year starts from 0
    a = 0

    ' Add all different years in spreadsheet to array
    For i = 0 To count
        pDate = arr(i, 0)        
        pYear = Year(pDate)

        yearInArray = FindElementInsideArray(yearArr, pYear)

        If yearInArray = -1 Then
            yearArr(a) = pYear
            allProjects(a, 0, 0) = pYear
            a = a + 1
        End If
    Next i

    yearCounter = a - 1 ' Remove one from total. Since one too many gets added
    ReDim Preserve yearArr(yearCounter)

    ' Array where final values will be stored.
    Dim yearProjects() As Variant
    ReDim yearProjects(yearCounter, count + 1, 5)
    ' Add years to final array
    yearProjects = allProjects
    ' Project index starts from 1
    a = 1

    ' Based on location of years in array add yearProjects
    For i = 0 To count
        pDate = arr(i, 0)
        pId = arr(i, 1)
        pCalcTime = arr(i, 2)
        pTotalTime = arr(i, 3)

        pYear = Year(pDate)
        yearInArray = FindElementInsideArray(yearArr, pYear)

        yearProjects(yearInArray, a, 0) = a
        yearProjects(yearInArray, a, 1) = pDate
        yearProjects(yearInArray, a, 2) = pId
        yearProjects(yearInArray, a, 3) = pCalcTime
        yearProjects(yearInArray, a, 4) = pTotalTime
        a = a + 1
    Next i

    Dim result() As Variant
    ' yearCounter = 1d size
    ' count = 2d size
    ' 5 = 3d size
    ' yearProjects = final array
    result = Array(yearCounter, count + 1, 5, yearProjects)

    FilterByYear = result
End Function

Function FilterByWeek(arr As Variant) As Variant    
    Dim a As Integer
    Dim b As Integer
    Dim c As Integer
    Dim d As Integer
    Dim weekArr() As Variant
    ReDim weekArr(arr(0), arr(1), arr(1), 6)

    For a = 0 To arr(0)
        pYear = arr(3)(a, 0, 0) ' year
        d = 1
        c = 1
        weekArr(a, 0, 0, 0) = pYear
        For b = 0 To arr(1)
            pIndex = arr(3)(a, b, 0) ' index
            pDate = arr(3)(a, b, 1) ' pDate
            pId = arr(3)(a, b, 2) ' pId
            pCalcTime = arr(3)(a, b, 3) ' pCalcTime
            pTotalTime = arr(3)(a, b, 4) ' pTotalTime
            if Len(pId) > 0 Then
                pWeekNumber = WorksheetFunction.WeekNum(pDate, vbMonday) - 1

                Dim elementInArray As Variant
                ' Get position of element in array
                elementInArray = FindWeekInside4DArray(weekArr, Array(a, arr(1)), pWeekNumber)

                Select Case VarType(elementInArray(0))
                    Case 2  ' Integer
                        ' Week exists in array
                        weekArr(elementInArray(0), elementInArray(1), d, 1) = d
                        weekArr(elementInArray(0), elementInArray(1), d, 2) = pDate
                        weekArr(elementInArray(0), elementInArray(1), d, 3) = pId
                        weekArr(elementInArray(0), elementInArray(1), d, 4) = pCalcTime
                        weekArr(elementInArray(0), elementInArray(1), d, 5) = pTotalTime
                        d = d + 1
                    Case 11 ' Boolean
                        ' Week doesn't exists in array
                        weekArr(a, c, 0, 0) = pWeekNumber
                        weekArr(a, c, d, 1) = d
                        weekArr(a, c, d, 2) = pDate
                        weekArr(a, c, d, 3) = pId
                        weekArr(a, c, d, 4) = pCalcTime
                        weekArr(a, c, d, 5) = pTotalTime
                        c = c + 1
                        d = d + 1
                    Case Else
                End Select 
            End If
        Next b    
    Next a

    FilterByWeek = Array(arr(0), arr(1), arr(1), 6, weekArr)
End Function

Public Function GetDayFromWeekNumber(inYear As Integer, weekNumber As Integer, Optional dayInWeek As Integer = 1) As Date
    Dim i As Integer: i = 1

    Do While Weekday(DateSerial(inYear, 1, i), vbMonday) <> dayInWeek
        i = i + 1
    Loop

    GetDayFromWeekNumber = DateAdd("ww", weekNumber - 1, DateSerial(inYear, 1, i))
End Function


' Display total times spend on projects
Function DisplayTotalTime(arr As Variant, firstSize As Integer, secondSize As Integer, column As String, Optional ByRef row As Integer = 1) As Integer
    Dim columnNumber As Integer
    columnNumber = CheckIfChrColumn(LCase(column), row)

    if columnNumber = -1 Then
        Exit Function
    End If

    Dim results() As Variant
    Dim i As Integer
    Dim elementCounter As Integer
    Dim elementInArray As Integer
    
    ReDim results(firstSize, secondSize)

    For i = 0 To firstSize
        pId = arr(i, 1)
        pCalcTime = arr(i, 2)
        pTotalTime = arr(i, 3)

        If Len(pId) > 0 Then
            ' Get position of element in array
            elementInArray = FindElementInsideArray(results, pId, 1)

            ' If project isn't added to array yet add it
            If elementInArray = -1 Then
                results(elementCounter, 1) = pId
                results(elementCounter, 2) = pCalcTime
                results(elementCounter, 3) = pTotalTime
                elementCounter = elementCounter + 1
            ' If project is added to array update the total time
            Else
                results(elementInArray, 2) = results(elementInArray, 2) + pCalcTime
                results(elementInArray, 3) = results(elementInArray, 3) + pTotalTime
            End If
        End If
    Next i

    With Cells(row, Chr(columnNumber))
        .value = "Totale tijden"
        .Font.Bold = True
        .Font.Size = 16
    End With
    
    With Cells(row + 1, Chr(columnNumber))
        .value = "Project"
        .Font.Bold = True
        .Font.Size = 13
    End With
    
    With Cells(row + 1, Chr(columnNumber + 1))
        .value = "Tijden"
        .Font.Bold = True
        .Font.Size = 13
        .HorizontalAlignment = xlRight
    End With
    
    With Cells(row + 1, Chr(columnNumber + 3))
        .value = "Tijden met pauze"
        .Font.Bold = True
        .Font.Size = 13
    End With
    
    Dim b As Integer
    Dim cellCounter As Integer
    ' Loop through array
    For b = 0 To elementCounter - 1
        pId = results(b, 1)
        pCalcTime = results(b, 2)
        pTotalTime = results(b, 3)
        
        ' Since we have two headers above move three rows down
        cellCounter = b + 3
        
        Cells(cellCounter, Chr(columnNumber)).value = pId
        With Cells(cellCounter, Chr(columnNumber + 1))
            .value = pCalcTime / 3600 ' Devide seconds into hours
            .NumberFormat = "0.0"
        End With
        With Cells(cellCounter, Chr(columnNumber + 3))
            .value = pTotalTime
            .NumberFormat = "0.00"
        End With
        
        'If you want to convert the decmial times to full times:
        'For example, let's convert 12.675 hours to hours, minutes, and seconds.
        '
        'Start by finding the number of hours.
        '12.675 hours = 12 hours + .675 hours
        'Full hours = 12
        '
        'Then find the number of minutes
        'minutes = .675 hours and 60 minutes
        'minutes = 40.5 minutes
        'Full minutes = 40
        '
        'Find the remaining seconds
        'seconds = .5 minutes and 60 seconds
        'seconds = 30 seconds
        '
        'Finally, rewrite as HH:MM:SS
        'time = 12:40:30
    Next b
    DisplayTotalTime = cellCounter
End Function

Sub DisplayYearlyTime(arr As Variant, projectsArrSize As Integer, column As String, row As Integer)
    Dim columnNumber As Integer
    columnNumber = CheckIfChrColumn(LCase(column), row)

    if columnNumber = -1 Then
        Exit Sub
    End If

    Dim i As Integer
    Dim o As Integer

    Dim yearRow As String
    Dim weekRow As String
    Dim projectRow As String
    Dim timesRow As String
    Dim timesBreakRow As String

    yearRow = Chr(columnNumber)
    weekRow = Chr(columnNumber + 1)
    projectRow = Chr(columnNumber + 3)
    timesRow = Chr(columnNumber + 4)
    timesBreakRow = Chr(columnNumber + 6)

    With Cells(row, Chr(columnNumber))
        .value = "Tijden per week"
        .Font.Bold = True
        .Font.Size = 16
    End With
    
    With Cells(row + 1, yearRow)
        .value = "Jaar"
        .Font.Bold = True
        .Font.Size = 13
    End With

    With Cells(row + 1, weekRow)
        .value = "Week"
        .Font.Bold = True
        .Font.Size = 13
    End With

    With Cells(row + 1, projectRow)
        .value = "Project"
        .Font.Bold = True
        .Font.Size = 13
    End With
    
    With Cells(row + 1, timesRow)
        .value = "Tijden"
        .Font.Bold = True
        .Font.Size = 13
        .HorizontalAlignment = xlRight
    End With
    
    With Cells(row + 1, timesBreakRow)
        .value = "Tijden met pauze"
        .Font.Bold = True
        .Font.Size = 13
    End With

    Dim yearArr() As Variant
    yearArr = FilterByWeek(arr)

    Dim pWeek As Integer

    Dim a As Integer
    Dim b As Integer
    Dim c As Integer

    Dim cellCounter As Integer
    cellCounter = row + 1

    For a = 0 To yearArr(0)
        cellCounter = cellCounter + 1
        pYear = yearArr(4)(a, 0, 0, 0)

        With Cells(cellCounter, yearRow)
            .value = pYear
            .Font.Size = 12
            .HorizontalAlignment = xlLeft
        End With

        ' Call outputToFile("Year", pYear)
        For b = 1 To yearArr(1)
            pWeek = yearArr(4)(a, b, 0, 0)
            if pWeek > 0 Then
                cellCounter = cellCounter + 1
                With Cells(cellCounter, weekRow)
                    .value = pWeek & " (" & GetDayFromWeekNumber(pYear, pWeek) & ")"
                    .Font.Size = 12
                    .HorizontalAlignment = xlLeft
                End With

                Dim weekResultSorted() As Variant
                ReDim weekResultSorted(yearArr(1), 3)
                Dim weekResultCount As Integer

                ' Call outputToFile("Week", pWeek)
                ' Add week results together
                For c = 1 To yearArr(2)
                    pId = yearArr(4)(a, b, c, 3) ' pId

                    if Len(pId) > 0 Then
                        pIndex = yearArr(4)(a, b, c, 1) ' index
                        pDate = yearArr(4)(a, b, c, 2) ' pDate
                        pCalcTime = yearArr(4)(a, b, c, 4) ' pCalcTime
                        pTotalTime = yearArr(4)(a, b, c, 5) ' pTotalTime

                        Dim elementInArray As Integer
                        ' Get position of element in array
                        elementInArray = FindElementInsideArray(weekResultSorted, pId, 3)

                        ' If project isn't added to array yet add it
                        If elementInArray = -1 Then
                            weekResultSorted(weekResultCount, 0) = pId
                            weekResultSorted(weekResultCount, 1) = pCalcTime
                            weekResultSorted(weekResultCount, 2) = pTotalTime
                            weekResultCount = weekResultCount + 1
                        ' If project is added to array update the total time
                        Else
                            weekResultSorted(elementInArray, 1) = weekResultSorted(elementInArray, 1) + pCalcTime
                            weekResultSorted(elementInArray, 2) = weekResultSorted(elementInArray, 2) + pTotalTime
                        End If
                    End If
                Next c

                ' Display week results
                For c = 0 To weekResultCount - 1
                    pId = weekResultSorted(c, 0) ' pId

                    if Len(pId) > 0 Then
                        pCalcTime = weekResultSorted(c, 1) ' pId
                        pTotalTime = weekResultSorted(c, 2) ' pId

                        cellCounter = cellCounter + 1
                        With Cells(cellCounter, projectRow)
                            .value = pId
                            .Font.Size = 12
                            .HorizontalAlignment = xlLeft
                        End With

                        With Cells(cellCounter, timesRow)
                            .value = pCalcTime  / 3600
                            .Font.Size = 12
                            .HorizontalAlignment = xlLeft
                        End With

                        With Cells(cellCounter, timesBreakRow)
                            .value = pTotalTime
                            .Font.Size = 12
                            .HorizontalAlignment = xlLeft
                        End With
                    End If
                Next c
            End If
        Next b    
    Next a
End Sub

Function CheckIfChrColumn(column As String, row As Integer) As Integer
    If Not Len(column) = 1 Then
        CheckIfChrColumn = -1
        Exit Function
    End If

    Dim columnName As String
    columnName = LCase(column)
    Dim columnNumber As Integer
    columnNumber = Asc(columnName)

    if columnNumber < 97 And columnNumber > 117 And row = 0 Then
        CheckIfChrColumn = -1
        Exit Function
    End If
    CheckIfChrColumn = columnNumber
End Function

' Calculate difference between two times and output the amount on seconds
Function CalculateTimeDifference(startTime As String, endTime As String)
    ' Get difference between to times and output the difference in seconds
    CalculateTimeDifference = DateDiff("s", startTime, endTime)
End Function

' Check if an array item exists in string
Function CheckIfStringContainsArrayElement(content As String, arr As Variant, Optional ByRef indexNumber As Integer = 0)
    Dim dimensions As Integer
    Dim result As Integer
    Dim length As Integer
    Dim index As Integer

    dimensions = NumberOfArrayDimensions(arr)

    If Not dimensions = -1 Then
        Dim lastValue As Integer
        ' Last value of array
        lastValue = UBound(arr)

        Dim x As Integer
        Dim inStrResult

        ' Loop through array
        If dimensions = 1 Then
            For x = LBound(arr) To UBound(arr)
                inStrResult = InStr(content, arr(x))
                ' Check if array item is contained in string
                If inStrResult >= 1 Then
                    result = inStrResult
                    length = Len(arr(x))
                    index = x
                    Exit For
                ElseIf x = lastValue Then
                    result = -1
                    Exit For
                End If
            Next x
        Else
            For x = 1 To UBound(arr, 1)
                inStrResult = InStr(content, arr(x, 1))
                ' Check if array item is contained in string
                If inStrResult >= 1 Then
                    result = inStrResult
                    length = Len(arr(x, 1))
                    index = x
                    Exit For
                ElseIf x = lastValue Then
                    result = -1
                    Exit For
                End If
            Next x
        End If
    Else
        result = -1
    End If
    CheckIfStringContainsArrayElement = Array(result, length, index)
End Function

' Find something inside an array
Function FindElementInsideArray(arr As Variant, valueToBeFound As Variant, Optional ByRef indexNumber As Integer = 0) As Integer
    Dim dimensions As Integer
    dimensions = NumberOfArrayDimensions(arr)

    If Not dimensions = -1 Then
        Dim lastValue As Integer
        ' Last value of array
        lastValue = UBound(arr)
        'declare an integer
        Dim i As Integer
        Dim o As Integer

        ' Loop through array
        For i = LBound(arr) To UBound(arr)
            If dimensions = 1 Then
                ' Check if item is the one we're searching for
                If EqualTo(arr(i), valueToBeFound) Then
                    FindElementInsideArray = i
                    Exit Function
                ' If we don't find what we're looking for return -1
                ElseIf i = lastValue Then
                    FindElementInsideArray = -1
                End If
            ElseIf dimensions = 2 Then
                ' Check if item is the one we're searching for
                If EqualTo(arr(i, indexNumber), valueToBeFound) Then
                    FindElementInsideArray = i
                    Exit Function
                ' If we don't find what we're looking for return -1
                ElseIf i = lastValue Then
                    FindElementInsideArray = -1
                End If
            ElseIf dimensions = 3 Then
                ' Check if item is the one we're searching for
                If EqualTo(arr(i, 0, indexNumber), valueToBeFound) Then
                    FindElementInsideArray = i
                    Exit Function
                ' If we don't find what we're looking for return -1
                ElseIf i = lastValue Then
                    FindElementInsideArray = -1
                End If
            End If
        Next i
    Else
        FindElementInsideArray = -1
    End If
End Function

' Find something inside an array
Function FindWeekInside4DArray(arr As Variant, sizesArr As Variant, valueToBeFound As Variant) As Variant
    Dim dimensions As Integer
    dimensions = NumberOfArrayDimensions(arr)
    If Not dimensions = -1 Then
        Dim lastValue As Integer
        ' Last value of array
        lastValue = UBound(arr)
        'declare an integer
        Dim i As Integer
        Dim o As Integer

        i = sizesArr(0)

        ' Loop through array
        For o = 1 To sizesArr(1)
            If VarType(arr(i, o, 0, 0)) > 1 Then
                ' Check if item is the one we're searching for
                If EqualTo(arr(i, o, 0, 0), valueToBeFound) Then
                    FindWeekInside4DArray = Array(i, o)
                    Exit Function
                End If
            End If
        Next o
    End If
    FindWeekInside4DArray = Array(False)
End Function

Function EqualTo(valueOne As Variant, valueTwo As Variant) As Boolean
    ' Check if the values of two fields are the same no matter the type
    ' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/vartype-function
    if VarType(valueOne) > 1 And VarType(valueTwo) > 1 Then
        Select Case VarType(valueOne)
            Case 2, 8
                if valueOne = valueTwo Then
                    EqualTo = True
                    Exit Function
                End If
            Case Else 
                ' If value isn't String or Int convert both two Strings and check if they're equal
                if CStr(valueOne) = CStr(valueTwo) Then
                    EqualTo = True
                    Exit Function
                End If
        End Select 
    End If
    EqualTo = False
End Function

Public Function NumberOfArrayDimensions(arr As Variant) As Integer
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' NumberOfArrayDimensions                                                                           '
    ' This function returns the number of dimensions of an array. An unallocated dynamic array          '
    ' has 0 dimensions. This condition can also be tested with IsArrayEmpty.                            '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If IsArray(arr) Then
        Dim Ndx As Integer
        Dim Res As Integer
        On Error Resume Next
        ' Loop, increasing the dimension index Ndx, until an error occurs.
        ' An error will occur when Ndx exceeds the number of dimension
        ' in the array. Return Ndx - 1.
        Do
            Ndx = Ndx + 1
            Res = UBound(arr, Ndx)
        Loop Until Err.Number <> 0
        
        Err.Clear
        NumberOfArrayDimensions = Ndx - 1
    Else
        NumberOfArrayDimensions = -1
    End If
End Function



