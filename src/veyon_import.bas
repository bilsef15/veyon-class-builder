Attribute VB_Name = "veyon_import"
'Created by Will Mohr
'7/2019
Sub veyonImport()
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Dim importList As String
    Dim batScript As String
    
    importList = Application.ThisWorkbook.Path & "\data\veyon_computers.txt"
    batScript = Application.ThisWorkbook.Path & "\data\initializeVeyonData.bat"
'######################################################
    
    Dim pid As Long
    Dim line As String
    
    On Error GoTo errorHandler
    
    Shell "cmd.exe /k cd /d" & Application.ThisWorkbook.Path & " && .\data\initializeVeyonData.bat"
    Application.Wait (Now + TimeValue("0:00:5"))
    
    Dim tempArray() As String
    
    Open importList For Input As #1
    Dim flag As Boolean
    Do Until EOF(1)
        Line Input #1, line
        tempArray = Split(line, ",")
        If (UBound(tempArray) = 2) Then
            If (tempArray(0) <> "") Then
                tempArray(0) = Replace(tempArray(0), """", "")
                tempArray(1) = Replace(tempArray(1), """", "")
                tempArray(2) = Replace(tempArray(2), """", "")
                flag = mySchool.addComputer(tempArray(0), tempArray(1), tempArray(2))
            End If
        End If
    Loop
    Close #1
    Dim names As Collection
    updateClassList
Done:
    Exit Sub
errorHandler:
    MsgBox "Necessary files are missing or damaged. Please reinstall. (User data should be retained)"
    Unload classPicker
    Unload classEditor
    ActiveWorkbook.Close savechanges = True
End Sub

