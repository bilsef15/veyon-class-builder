Attribute VB_Name = "main_one"
'This programs allows for editing of veyon network objects through excel userforms.
'It provides a registry for users to select from which are stored in another excel file.
'this can manipulate current classes and new classes
Sub main()
    'pulls data from the roster.xlsx in data
    fillRoster
    'imports the current veyon computer configuration
    veyonImport
    'shows the initial menu
    classPicker.Show
End Sub
