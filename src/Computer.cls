VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Computer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'individual computer class
'holds a name and hostname
Private name As String 'name of computer
Private hostname As String 'hostname of computer

Private Sub Class_Initialize()
    name = "a"
    hostname = "b"
End Sub

Property Let computerName(n As String)
    name = n
End Property

Property Get computerName() As String
    computerName = name
End Property

Property Let computerHostname(h As String)
    hostname = h
End Property

Property Get computerHostname() As String
    computerHostname = hostname
End Property

'true if both hostname and name are equal else false
Public Function isEqual(c As Computer) As Boolean
    If (name = c.computerName) And (hostname = c.computerHostname) Then
        isEqual = True
        Else
        isEqual = False
        End If
End Function

'true if only name is equal else false
Public Function sameName(c As Computer) As Boolean
    If (name = c.computerName) Then
        isEqual = True
        Else
        isEqual = False
        End If
End Function

