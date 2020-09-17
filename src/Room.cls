VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Room"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'a room holds a bunch of computers
Private name As String 'name of the room
Private computers As New Collection 'collection of students

Private Sub Class_Initialize()
    name = "a"
End Sub

Public Property Let roomName(n As String)
    name = n
End Property

Public Property Get roomName() As String
    roomName = name
End Property

Public Property Get computerList() As Collection
    computerList = computers
End Property

Public Property Get getComputers() As Collection
    Set getComputers = computers
End Property

'returns a collection of computer names contained in the room
Public Function getComputerNames() As Collection
    Dim names As New Collection
    Dim c As Computer
    For Each c In computers
        names.add (c.computerName)
    Next c
    Set getComputerNames = names
End Function

'adds a computer to the room collection
'if it is a duplicate it will return false and not add the computer
'if new addition, return true and add computer
Public Function addComputer(n As String, h As String) As Boolean
    Dim c As New Computer
    c.computerName = n
    c.computerHostname = h
    Dim compare As Computer
    Dim duplicate As Boolean
    
    duplicate = False
    For Each compare In computers
        If c.isEqual(compare) Then
            duplicate = True
        End If
    Next compare
    If duplicate = False Then
        computers.add c
    End If
    addComputer = Not duplicate
End Function


'looks for computer by name and removes the computer
Public Function removeComputer(n As String) As Boolean
    Dim i As Integer
    Dim flag As Boolean
    flag = False
    i = 1
    Do While (i <> computers.Count + 1) And (flag <> True)
        If (computers(i).computerName = n) Then
            computers.remove (i)
            flag = True
        End If
        i = i + 1
    Loop
    If (flag = False) Then
        MsgBox "No computer found"
    End If
    removeComputer = flag
End Function

'looks for computer by name and returns the computer
Public Function getComputer(n As String) As Computer
    Dim i As Integer
    Dim flag As Boolean
    flag = False
    i = 1
    Do While (i <> computers.Count + 1) And (flag <> True)
        If (computers(i).computerName = n) Then
            flag = True
        End If
        i = i + 1
    Loop
    If (flag = False) Then
        MsgBox "No computer found"
    End If
        Set getComputer = computers(i - 1)
End Function

Public Function getComputerCount() As Long
    getComputerCount = computers.Count
End Function


'testing purposes!!!!!!!!!!!
Public Sub listComputers()
    Dim c As Computer
    For Each c In computers
        Debug.Print (c.computerName & "   " & c.computerHostname)
    Next c
End Sub
