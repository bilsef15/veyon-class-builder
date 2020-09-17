VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "School"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'a school holds a bunch of rooms

Private rooms As New Collection 'contains all the rooms of the school

'adds a room to the school, true if successful
'false if unsuccessful, only in case of duplicate
Public Function addRoom(rn As String) As Boolean
    Dim compare As Room
    Dim r As New Room
    r.roomName = rn
    Dim duplicate As Boolean
    duplicate = False
    For Each compare In rooms
        If (r.roomName = compare.roomName) Then
            duplicate = True
        End If
    Next compare
    If duplicate = False Then
        rooms.add r
        addRoom = True
        Else
        addRoom = False
        MsgBox "Duplicate Room Name"
    End If
End Function

'remove a room by name
'true if successful
'false if unsuccessful, usually only if the room does not exist
Public Function removeRoom(r As String) As Boolean
    Dim i As Integer
    Dim flag As Boolean
    flag = False
    If (Me.roomExist(r) = True) Then
        i = 1
        Do While (i <> rooms.Count + 1) And (flag <> True)
            If (rooms(i).roomName = r) Then
                flag = True
                rooms.remove (i)
            End If
        i = i + 1
        Loop
    End If
    removeRoom = flag
End Function

'returns a collection of room names
Public Function getRoomNames() As Collection
    Dim names As New Collection
    Dim r As Room
    For Each r In rooms
        names.add (r.roomName)
    Next r
    Set getRoomNames = names
End Function

'returns a room by name
Public Function getRoom(n As String) As Room
    Dim i As Integer
    Dim flag As Boolean
    flag = False
    i = 1
    Do While (i <> rooms.Count + 1) And (flag <> True)
        If (rooms(i).roomName = n) Then
            flag = True
        End If
        i = i + 1
    Loop
    If (flag = False) Then
        MsgBox "No Room Found"
    End If
        Set getRoom = rooms(i - 1)
End Function

'returns the number of rooms
Public Function getRoomCount() As Long
    getRoomCount = rooms.Count
End Function

'adds computer via room, name, and hostname
Public Function addComputer(r As String, n As String, h As String) As Boolean
    Dim flag As Boolean
    If (Me.roomExist(r) = True) Then
        flag = Me.getRoom(r).addComputer(n, h)
    Else
        flag = Me.addRoom(r)
        flag = Me.getRoom(r).addComputer(n, h)
    End If
    addComputer = flag
End Function

'removes a computer, by room name, and computer name
'true if successful, false if unsuccessful (room does not exist)
Public Function removeComputer(r As String, n As String) As Boolean
    Dim flag As Boolean
    If (Me.roomExist(r) = True) Then
        flag = Me.getRoom(r).removeComputer(n)
    Else
        flag = False
    End If
    removeComputer = flag
End Function

'returns true if the the room exists
Public Function roomExist(n As String) As Boolean
    Dim i As Integer
    Dim flag As Boolean
    flag = False
    i = 1
    Do While (i <> rooms.Count + 1) And (flag <> True)
        If (rooms(i).roomName = n) Then
            flag = True
        End If
        i = i + 1
    Loop
    roomExist = flag
End Function


'testing purposes
Public Sub listRoomsAndComputers()
    Dim r As Room
    Dim flag As Boolean
    For Each r In rooms
        Debug.Print (r.roomName)
        r.listComputers
    Next r
End Sub
'testing purposes
Public Sub listRooms()
    Dim r As Room
    Dim flag As Boolean
    For Each r In rooms
        Debug.Print (r.roomName)
    Next r
End Sub

'exports the school info into a csv named export.txt into the workbook directory
'for example:
'   room one, bob, smith
'   room one, james, moore
'   room two, fred, jenkins
Public Sub exportSchool()
    Dim exportFilepath As String
    Dim r As Room
    Dim c As Computer
    'need to try this
    exportFilepath = Application.ActiveWorkbook.Path & "\data\export.txt"
    Open exportFilepath For Output As #1
    For Each r In rooms
        For Each c In r.getComputers()
            Print #1, r.roomName & "," & c.computerName & "," & c.computerHostname
        Next c
    Next r
    Close #1
End Sub



