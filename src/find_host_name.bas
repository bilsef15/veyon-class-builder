Attribute VB_Name = "find_host_name"
'finds the hostname of a given computer name in the roster
Public Function findHostname(r As Collection, n As String) As String 'r should always be masterRoster
    Dim i As Integer 'index within each roster (collection within each sheet
    Dim j As Integer 'index of rosters in masterRoster (collections of sheets)
    
    Dim flag As Boolean
    flag = False
    i = 1
    j = 1
    Do While (j <> r.Count + 1) And (flag <> True) 'cycle through each sheet (data stored in roster)
        i = 1
        Do While (i <> r.Item(j).rosterEntries.Count + 1) And (flag <> True) 'cycle through collection of entries in each sheet
            If (r.Item(j).rosterEntries.Item(i).computerName = n) Then
                flag = True
            End If
            i = i + 1
        Loop
        j = j + 1
    Loop
    findHostname = r.Item(j - 1).rosterEntries.Item(i - 1).computerHostname
End Function
