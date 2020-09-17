Attribute VB_Name = "fill_computer_rosters"
'fills the gui with the computers from roster.xlsx
Sub fillComputerRosters()
    Dim entry As Computer
    Dim i As Integer
    
    'loop through entire master roster
    For i = 1 To masterRoster.Count
        'add the listbox for each page in the roster.xlsx
        classEditor.classStrip.Pages.add masterRoster.Item(i).rosterName & 1, masterRoster.Item(i).rosterName
        listBoxes.add classEditor.classStrip.Pages(i - 1).Controls.add("Forms.ListBox.1")
        With listBoxes.Item(i)
            .Height = 330
            .Width = 179
            .Top = 2
            .Left = 2
        End With
        'add the computers to the correct page's listbox
        For Each e In masterRoster(i).rosterEntries
            listBoxes.Item(i).AddItem e.computerName
        Next e
    Next i
    

End Sub
