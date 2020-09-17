VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Roster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'a Roster holds a collection pertaining to one page of the roster.xlsx sheet

Private name As String 'name of sheet in workbook
Private entries As New Collection 'computers in one sheet in classroom


Private Sub Class_Initialize()
    name = "a"
End Sub

Property Let rosterName(rosterNameN As String)
    name = rosterNameN
End Property

Property Get rosterName() As String
    rosterName = name
End Property

Property Let rosterEntries(rosterEntriesN As Collection)
    Set entries = rosterEntriesN
End Property

Property Get rosterEntries() As Collection
    Set rosterEntries = entries
End Property


