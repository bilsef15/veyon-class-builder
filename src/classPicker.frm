VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} classPicker 
   Caption         =   "Veyon Class Editor"
   ClientHeight    =   6750
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5640
   OleObjectBlob   =   "classPicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "classPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Userform for picking a class to edit or adding or removing class

'adds a class to mySchool
Private Sub addClass_Click()
    Dim name As String
    Dim flag As Boolean
    name = InputBox("Input Room Name", "Class Name")
    'make sure the class has a name
    If (name <> "") Then
        flag = mySchool.addRoom(name)
    End If
    updateClassList
End Sub

'deletes the selected class
Private Sub deleteClass_Click()
    'make sure a class is selected
    If (classPicker.classes.ListIndex <> -1) Then
        mySchool.removeRoom classPicker.classes.List(classPicker.classes.ListIndex)
        updateClassList
    Else
        MsgBox "Please select a class to delete."
    End If
End Sub

'edits the computers in the selected class
Private Sub editClass_Click()
    Dim n As String
    Dim r As Room
    Dim name As Variant
    Dim names As Collection
    'make sure a class is selected
    If (classPicker.classes.ListIndex <> -1) Then
        n = classPicker.classes.List(classPicker.classes.ListIndex)
        Set r = mySchool.getRoom(n) 'get selected room
        Set names = r.getComputerNames() 'get list of computer names
        classEditor.computerList.Clear 'clear the list of computers so there is a clean slate for new one to edit
        'add each computer in the room to the list
        For Each name In names
            classEditor.computerList.AddItem (name)
        Next name
        'show the class editor
        classEditor.Show
    Else 'if class is not selected
        MsgBox "Please select a class to edit."
    End If
End Sub

Private Sub editName_Click()
    Dim name As String
    Dim flag As Boolean
    'make sure a class is selected
    If (classPicker.classes.ListIndex = -1) Then
        MsgBox "Please select a class."
        Exit Sub
    End If
    'set the default entry as the current room name
    name = InputBox("Input Room Name", "Class Name", classPicker.classes.List(classPicker.classes.ListIndex))
    'make sure the class has a name, if not keep the same name
    If (name <> "") Then
        mySchool.getRoom(classPicker.classes.List(classPicker.classes.ListIndex)).roomName = name
    End If
    'update the list of classes
    updateClassList
End Sub

Private Sub exitClass_Click()
    Unload classPicker
    ThisWorkbook.Close savechanges = True
End Sub

Private Sub saveClass_Click()
    Dim pid As Long
    mySchool.exportSchool 'convert the the school structure to a csv
    'run the script to import the new school structure into veyon
    pid = Shell(Application.ThisWorkbook.Path & "\data\importComputerData.bat", vbMinimizedNoFocus)
    Application.Wait (Now + TimeValue("0:00:2"))
End Sub

