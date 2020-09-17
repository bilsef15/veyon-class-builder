VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} new_unlisted 
   Caption         =   "Add Unlisted Computer"
   ClientHeight    =   2250
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4980
   OleObjectBlob   =   "new_unlisted.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "new_unlisted"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'userform for adding an unlisted sta-asset to a class

Private Sub add_new_Click()
    Dim n As String
    If new_unlisted.new_name.Text <> "" And new_unlisted.hostname.Text <> "" Then
        'get the name of the selected room
        n = mySchool.getRoom(classPicker.classes.List(classPicker.classes.ListIndex)).roomName
        If mySchool.addComputer(n, new_unlisted.new_name.Text, new_unlisted.hostname.Text) = True Then
            classEditor.computerList.AddItem (new_unlisted.new_name.Text)
        End If
        new_unlisted.Hide
    Else
        MsgBox "Name and Hostname boxes must have values."
    End If
End Sub

Private Sub close_box_Click()
    new_unlisted.Hide
End Sub



Private Sub UserForm_Initialize()
    'text is empty when started
    new_unlisted.new_name.Text = ""
    hostname.Text = ""
End Sub
