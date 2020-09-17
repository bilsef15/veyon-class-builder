Attribute VB_Name = "update_class_list"

'updates the list of classes displayed bases on rooms in mySchool
Sub updateClassList()
    Dim names As Collection
    Set names = mySchool.getRoomNames()
    classPicker.classes.Clear
    For Each name In names
        classPicker.classes.AddItem (name)
    Next name
End Sub

