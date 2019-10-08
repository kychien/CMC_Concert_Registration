Attribute VB_Name = "Module1"
Sub arrange_sheets()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   Script to rearrange Google Forms results to a more readable format for error checking.
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ' Go to first sheet in case of multiple runs
    Worksheets(1).Activate
    
    ' Determine row count for variable sizing
    Dim rCt As Long
    rCt = Cells(Rows.Count, 1).End(xlUp).row
    
    ' Determine number of pairs, songs/pieces, times for variable sizing
    Dim pCt, sCt, tCt As Integer
    pCt = 17                                    ' Initial implementation assuming fixed # 17
    sCt = 15                                    ' Initial implementation assuming fixed # 15
    tCt = 2                                     ' Initial implementation assuming fixed # 2
    
    ' Variables for scraping data
    Dim teachers(rCt) As String
    Dim addrs(rCt) As String
    Dim addr As String
    Dim pNums(rCt) As String
    Dim emails(rCt) As String
    Dim tYears(rCt) As Integer
    Dim playTs(rCt) As String
    Dim playPs(rCt) As String
    Dim playable As String
    Dim repStudents(rCt) As Integer
    Dim stuFees(rCt) As Integer
    Dim memStatus(rCt) As String
    Dim memFeeStat(rCt) As String
    Dim totFees(rCt) As Integer
    Dim actStudents(rCt) As Integer
    Dim students(rCt * 30) As String
    Dim stuYears(rCt * 30) As Integer
    Dim stuStats(rCt * 30) As String
    Dim pair As String
    Dim pAssigns(sCt * tCt, pCt) As String
    Dim waitList(sCt * tCt, pCt) As String
    Dim pairCts(sCt * tCt) As Integer
    Dim notes(rCt) As String
    Dim y, x, cur, aCt, sI, px, pn, t As Integer
    Dim cTeach, cEmail, cPhone As String
    Dim dupes(rCt) As Boolean
    Dim valid As Boolean
    
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
        emails(cur) = Cells(y, 2)
        pNums(cur) = Cells(y, 8)
        tYears(cur) = Cells(y, 9)
        
    '   Check for duplicate entry
        If (IsInArr(cTeach, teachers) = True) Then
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
        
        addr = addr + pad + Cells(y, 7)
        
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
    '    Add to queue if there is space
                        If (Left(Cells(i, px + 2), 1) = "4") Then
                            t = (pn * 2) + 2
                        Else
                            t = (pn * 2) + 1
                        End If
                        If (pairCts(t) < 17) Then
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
    Next i
    
    ' Create a sheet for teacher information
    
    
    ' Create a sheet for student information
    
    
    ' Create a sheet for song assignment
    
    
    
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
