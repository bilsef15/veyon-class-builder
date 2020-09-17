VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} classEditor 
   Caption         =   "Veyon Class Editor"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10500
   OleObjectBlob   =   "classEditor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "classEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Userform for adding and removing computers from a particular class

'adds a selected computer to the selected room
Private Sub add_Click()
    Dim rn As String 'room name
    Dim cn As String 'computer name (student name)
    Dim hn As String 'computer hostname
    Dim selectedPage As Integer
    
    selectedPage = classEditor.classStrip.Value
    
    If (listBoxes.Item(selectedPage + 1).ListIndex = -1) Then
        MsgBox "Please select student to add."
        Exit Sub
    End If
    
    
    rn = mySchool.getRoom(classPicker.classes.List(classPicker.classes.ListIndex)).roomName
    cn = listBoxes.Item(selectedPage + 1).List(listBoxes.Item(selectedPage + 1).ListIndex)
    hn = findHostname(masterRoster, cn)
    
    
    If mySchool.addComputer(rn, cn, hn) = True Then
        classEditor.computerList.AddItem (listBoxes.Item(selectedPage + 1).List(listBoxes.Item(selectedPage + 1).ListIndex))
    End If
End Sub

Private Sub add_not_listed_Click()
    'make sure textboxes are empty
    new_unlisted.new_name.Text = ""
    new_unlisted.hostname.Text = ""
    new_unlisted.Show
End Sub

Private Sub okay_Click()
    classEditor.Hide
End Sub

Private Sub remove_Click()
    Dim n As String
    'make sure something is selected
    If (computerList.ListIndex = -1) Then
        MsgBox "Please select student to remove."
        Exit Sub
    End If
    
    'remove computer by name from the school data structure
    n = mySchool.getRoom(classPicker.classes.List(classPicker.classes.ListIndex)).roomName
    If mySchool.removeComputer(n, computerList.List(computerList.ListIndex)) = True Then
            classEditor.computerList.RemoveItem (computerList.ListIndex)
    End If
End Sub

Private Sub UserForm_Initialize()
    fillComputerRosters
End Sub
