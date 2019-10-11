Attribute VB_Name = "Module1"
Sub arrange_sheets()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   Script to rearrange Google Forms results to a more readable format for error checking.
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ' Go to first sheet in case of multiple runs
    Worksheets(1).Activate
    
    ' Variables for scraping data
    Dim priority() As String
    Dim teachers() As String
    Dim addrs() As String
    Dim addr As String
    Dim pNums() As String
    Dim emails() As String
    Dim tYears() As Integer
    Dim playTs() As String
    Dim playPs() As String
    Dim playable As String
    Dim repStudents() As Integer
    Dim stuFees() As Integer
    Dim memStatus() As String
    Dim memFeeStat() As String
    Dim totFees() As Integer
    Dim actStudents() As Integer
    Dim students() As String
    Dim stuYears() As Integer
    Dim stuStats() As String
    Dim pair As String
    Dim pAssigns() As String
    Dim waitList() As String
    Dim pairCts() As Integer
    Dim notes() As String
    Dim y, x, cur, aCt, sI, px, pn, t As Integer
    Dim cTeach, cEmail, cPhone As String
    Dim dupes() As Boolean
    Dim valid As Boolean
    Dim songs() As String
    Dim sheetName, dateExt As String
    
    ' Determine row count for variable sizing
    Dim rCt As Long
    rCt = Cells(Rows.Count, 1).End(xlUp).row
    
    ' Determine number of pairs, songs/pieces, times for variable sizing
    Dim pCt, sCt, tCt As Integer
    pCt = 17                                    ' Initial implementation assuming fixed # 17
    sCt = 15                                    ' Initial implementation assuming fixed # 15
    tCt = 2                                     ' Initial implementation assuming fixed # 2
    
    ' Variables for Registration fee calculations
    Dim monER, dayER, mFee, sFee, lFee, maxS As Integer
    monER = 10                                  ' Month of end of early registration
    dayER = 25                                  ' Last DAY of early registration period
    mFee = 50                                   ' Membership Fee
    sFee = 60                                   ' Early registration fee
    lFee = 70                                   ' Late registration fee
    maxS = 30                                   ' Maximum number of students allowed in form
    
    dateExt = Format(Date, "yyyy-mm-dd")
    
    ReDim priority(rCt)
    ReDim priority(rCt)
    ReDim teachers(rCt)
    ReDim addrs(rCt)
    ReDim pNums(rCt)
    ReDim emails(rCt)
    ReDim tYears(rCt)
    ReDim playTs(rCt)
    ReDim playPs(rCt)
    ReDim repStudents(rCt)
    ReDim stuFees(rCt)
    ReDim memStatus(rCt)
    ReDim memFeeStat(rCt)
    ReDim totFees(rCt)
    ReDim actStudents(rCt)
    ReDim students(rCt * maxS)
    ReDim stuYears(rCt * maxS)
    ReDim stuStats(rCt * maxS)
    ReDim pAssigns(sCt * tCt, pCt)
    ReDim waitList(sCt * tCt, pCt)
    ReDim pairCts(sCt * tCt)
    ReDim notes(rCt)
    ReDim dupes(rCt)
    ReDim songs(sCt)
    
    ' Get the song names
    For i = 1 To sCt
        songs(i) = Split((Split(Cells(1, (11 + i)), "[")(1)), "]")(0)
    Next i
    
    For i = 1 To (sCt * tCt)
        pairCts(i) = 0
    Next i
    
    ' Scrape data off of the first sheet
    y = 2
    x = 1
    cur = 1
    
    '  While there are rows of data to parse
    Do While (Cells(y, 1) <> "")
    
    '   Get the name, email and phone number
        cTeach = UCase(Cells(y, 3))
        priority(cur) = Cells(y, 1)
        emails(cur) = Cells(y, 2)
        pNums(cur) = Cells(y, 8)
        tYears(cur) = Cells(y, 9)
        
    '   Check for duplicate entry
        If (IsInArr(CStr(cTeach), teachers) = True) Then
            ' Flag previous duplicates
            For i = 1 To curr
                If (teachers(i) = cTeach) Then
                    dupes(i) = True
                End If
            Next i
        End If
        dupes(cur) = False
        
        teachers(cur) = cTeach
        
        
    '   Combine the address
        addr = StrConv((Cells(y, 4) + ", " + Cells(y, 5)), vbProperCase)
        addr = addr + ", " + Cells(y, 6)
        
        Dim pad As String
        pad = " "
        If (Len(Cells(y, 7)) = 4) Then
            pad = " 0"
        End If
        
        addr = addr + pad + CStr(Cells(y, 7))
        
        addrs(cur) = addr
        
    '   Determine the play times
        playTs(cur) = Cells(y, 10)
    
    '   Determine the playable pieces
    '    If all pieces selected move on
        If (Cells(y, 11) = "Either" Or Len(Cells(y, 11)) > 12) Then
            playPs(cur) = "All"
        Else
    '    Otherwise for the remaining piece options, gather responses
            playable = ""
            ' Case for Primo/Secondo for All pieces
            If (Cells(y, 11) <> "") Then
                playable = "00-" + Left(Cells(y, 11), 1) + "; "
            End If
            ' Individual piece indicators
            For i = 12 To 26
                If (Cells(y, i) <> "") Then     ' A piece is playable
                    playable = playable + CStr((i - 11)) + "-"      ' Use numbers as placeholders for pieces
                    If (Cells(y, i) = "Either" Or Len(Cells(y, i)) > 12) Then
                        playable = playable + "B; "
                    Else
                        playable = playable + Left(Cells(y, i), 1) + "; "
                    End If
                End If
            Next i
            playPs(cur) = playable
        End If
    '   Get the reported number of students, student fees, member status and member fee status
        repStudents(cur) = Cells(y, 27)
        stuFees(cur) = Cells(y, 28)
        memStatus(cur) = Cells(y, 29)
        memFeeStat(cur) = Cells(y, 30)
        totFees(cur) = Cells(y, 31)
    
    '   While there are students to add to the list
        x = 32
        aCt = 0
        Do While ((x < 152) And (Cells(y, x) <> ""))
    '    Get student name, participation years and senior status
    '    Increment the actual student count
            aCt = aCt + 1
            sI = ((cur - 1) * 30) + aCt
            students(sI) = Cells(y, x)
            stuYears(sI) = Cells(y, (x + 1))
            stuStats(sI) = Cells(y, (x + 2))
            x = x + 4
        Loop
        
    '   Check if the actual student count matches what was reported
        actStudents(cur) = aCt
    
    '   Parse the special notes?
        notes(cur) = Cells(y, 319)  ' No special treatment on initial implementation
        
        y = y + 1
        cur = cur + 1
    Loop
    
    ' Loop again to get pieces for queue
    For i = 2 To (y - 1)
    '  For each teacher who's not a dupe
        If (dupes(i - 1) = False) Then
            x = 152
    '   For each piece
            pn = 0
            Do While (x < 318 And pn < sCt)
    '    Get provided primo/secondo
                For j = 0 To 2
                    valid = False
                    px = x + (j * 3) + 1
                    pair = Cells(i, px)
                    If (pair <> "") Then
                        valid = True
                    End If
                    pair = pair + " / "
                    If (Cells(i, px + 1) <> "") Then
                        valid = True
                        pair = pair + Cells(i, px + 1)
                    End If
                    If (valid) Then
    '    Identify appropriate queue based on time!!!
                        If (Left(Cells(i, px + 2), 1) = "4") Then
                            t = pn + sCt
                        Else
                            t = pn
                        End If
    '    Add to queue if there is space
                        If (pairCts(t) < pCt) Then
                            pairCts(t) = pairCts(t) + 1
                            pAssigns(t, pairCts(t)) = pair
                            
    '    Add to waitlist if queue is full
                        ElseIf (pairCts(t) < 34) Then
                            pairCts(t) = pairCts(t) + 1
                            waitList(t, (pairCts(t) - 17)) = pair
                        End If
                    End If
                Next j
                
                x = x + 10
                pn = pn + 1
            Loop
        End If
    Next i
    
    ' Create a sheet for teacher information
    sheetName = "T " + dateExt
    nameSheet (sheetName)
    
    '  Create labels for the worksheet
    Cells(1, 1) = "Name"
    Cells(1, 2) = "#"
    Cells(1, 3) = "@"
    Cells(1, 4) = "Yrs"
    Cells(1, 5) = "Member"
    Cells(1, 6) = "Fee"
    Cells(1, 7) = "Students"
    Cells(1, 8) = "Actual"
    Cells(1, 9) = "Add. Fees"
    Cells(1, 10) = "Reported"
    Cells(1, 11) = "Act. Due"
    Cells(1, 12) = "Address"
    Cells(1, 13) = "Notes"
    
    '  For each non-dupe teacher on the list
    y = 2
    For i = 1 To (cur - 1)
        If dupes(i) = False Then
    '   Dump row information
            Cells(y, 1) = teachers(i)
            Cells(y, 2) = pNums(i)
            Cells(y, 3) = emails(i)
            Cells(y, 4) = tYears(i)
            Cells(y, 5) = memStatus(i)
            Cells(y, 6) = memFeeStat(i)
            Cells(y, 7) = repStudents(i)
            Cells(y, 8) = actStudents(i)
            Cells(y, 9) = stuFees(i)
            Cells(y, 10) = totFees(i)
    '   Check fees due
            x = actStudents(i)
            Dim mo, day As Integer
    '    Check early registration status for fee multiplier
            mo = CInt(Right(Left(priority(i), 7), 2))
            day = CInt(Right(Left(priority(i), 10), 2))
            If (mo < monER) Or (mo = monER And day <= dayER) Then
                x = x * sFee
            Else
                x = x * lFee
            End If
    '    Check membership status for additional dues
            If (Left(memFeeStat(i), 1) = "T") Then
                x = x + mFee
            End If
            Cells(y, 11) = x
            Cells(y, 12) = addrs(i)
            Cells(y, 13) = notes(i)
            y = y + 1
        End If
    Next i
    
    Sheets(sheetName).UsedRange.Columns.AutoFit
    
    ' Create a sheet for student information
    sheetName = "S " + dateExt
    nameSheet (sheetName)
    
    '  Create labels for student sheet
    Cells(1, 1) = "Name"
    Cells(1, 2) = "Years"
    Cells(1, 3) = "Senior"
    
    y = 2
    For i = 1 To (rCt * 30)
        If (students(i) <> "") Then
            Cells(y, 1) = students(i)
            Cells(y, 2) = stuYears(i)
            Cells(y, 3) = stuStats(i)
            y = y + 1
        End If
    Next i
    
    Sheets(sheetName).UsedRange.Columns.AutoFit
    
    ' Create a sheet for song assignment
    For i = 0 To tCt - 1                    'For each concert time...
        Dim newName As String
        newName = "P" + CStr(i + 1) + " " + dateExt
        nameSheet (newName)
        
        For j = 1 To sCt                    'For each song...
            y = 1 + (8 * (j - 1))
            
            Cells(y, 1) = songs(j)          'Place the title
            'Cells(y, 3) = comps(j)          'Place the composer
            
            y = y + 1                       'Place all the pairs
            py = (i * sCt) + j
            pn = 1
            x = 1
            Do While (pAssigns(py, pn) <> "" And pn < pCt)
                Cells(y, x) = pAssigns(py, pn)
                If (x < 3) Then
                    x = x + 1
                Else
                    y = y + 1
                    x = 1
                End If
                pn = pn + 1
                
            Loop
        Next j
        
        Sheets(newName).UsedRange.Columns.AutoFit
        
    Next i
    
End Sub

Function IsInArr(target As String, arr As Variant) As Boolean
    IsInArr = (UBound(Filter(arr, target)) > -1)
End Function

Function nameSheet(name As String)
    Dim ws As Worksheet
    With ThisWorkbook
        Set ws = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        ws.name = name
    End With
    ws.Activate
End Function
