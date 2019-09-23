Attribute VB_Name = "Module1"
Sub arrange_sheets()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   Script to rearrange Google Forms results to a more readable format for error checking.
    '
    '   2019/09/17 - Setup comment skeleton of macro
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
    Dim stuStats(rCt * 30) As Boolean
    Dim pair As String
    Dim pAssigns(sCt * tCt, pCt) As String
    Dim waitList(sCt * tCt, pCt) As String
    Dim pairCts(sCt * tCt) As Integer
    Dim notes(rCt) As String
    Dim y, x As Integer
    
    ' Scrape data off of the first sheet
    '  While there are rows of data to parse
    
    '   Get the name, email and phone number
    '   Check for duplicate entry
    
    '   Combine the address
    
    '   Determine the play times
    
    '   Determine the playable pieces
    '    If all pieces selected move on
    '    Otherwise for the remaining piece options, gather responses
    
    '   Get the reported number of students, student fees, member status and member fee status
    
    '   While there are students to add to the list
    '    Get student name, participation years and senior status
    '    Increment the actual student count
    '   Check if the actual student count matches what was reported
    
    '   For each piece
    '    Get provided primo/secondo
    '    Add to queue if there is space
    '    Add to waitlist if queue is full
    
    '   Parse the special notes?
    
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
