Attribute VB_Name = "fill_roster"
'cycle through each sheet of the roster.xlsx, and fill the master roster
Sub fillRoster()
    Dim rosterPath As String
    Dim rosterBook As Workbook
    Dim sh As Worksheet
    Dim i As Long
    Dim entry As Computer
    
    Dim currEntries As Collection
    Dim currRoster As Roster
    
    On Error GoTo errorHandler
    
    rosterPath = Application.ThisWorkbook.Path & "\data\roster.xlsx"
    Debug.Print (rosterPath)
    
    Set rosterBook = Workbooks.Open(Filename:=rosterPath, ReadOnly:=True)
    'cycle through all the sheets in the workbook
    For Each sh In rosterBook.Worksheets()
        i = 2
        Set currEntries = New Collection
        Set currRoster = New Roster
        currRoster.rosterName = sh.name
        'look through all the rows with data
        Do While IsEmpty(sh.Cells(i, 1).Value) = False
            Set entry = New Computer
            entry.computerName = sh.Cells(i, 1).Value
            entry.computerHostname = sh.Cells(i, 3).Value
            currEntries.add entry
            'add to seventh grade if on "7th Grade" sheet
            'If sh.name = "7th Grade" Then
            '   seventhRoster.add entry
            'End If
            'add to eight grade if on "8th Grade" sheet
            'If sh.name = "8th Grade" Then
            '    eighthRoster.add entry
            'End If
            i = i + 1
        Loop
        currRoster.rosterEntries = currEntries
        masterRoster.add currRoster
    Next sh
    rosterBook.Close
    
Done:
    Exit Sub
errorHandler:
    MsgBox "The roster containing student data is not currently available. You can still work, however, there will be no students in the roster."
End Sub
