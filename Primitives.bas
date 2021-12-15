Attribute VB_Name = "Primitives"
Function DocVarExists(varName As String) As Boolean

    '*** If MS Word document variable exists return true ***
    
    Dim objVar As Variant

    For Each objVar In ActiveDocument.Variables

        If objVar.Name = varName Then
            
            DocVarExists = True
            
            Exit Function
        End If
        
    Next objVar

    DocVarExists = False

End Function

Function mydocs() As String
    
    '*** Return the path to the users 'My Documents' directory ***
    
    Dim Wshshell As Object
    Set Wshshell = CreateObject("WScript.Shell")
    
    mydocs = Wshshell.SpecialFolders("MyDocuments")

End Function

Function getFolderLst(dirPath As String, files As Boolean) As String()

    '*** Return an array of folder names ***
    
    Dim objFSO As Object
    Dim objFolders As Object
    Dim objFolder As Object
    Dim arrFolders() As String
    Dim FolderCount As Long
    
    '*** get the folder names ***
    '*** could have used 'Dir' like some damn amateur but instead used the file system object like a true pro ***
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If files = False Then
        
        Set objFolders = objFSO.getfolder(dirPath).SubFolders
    Else
        
        Set objFolders = objFSO.getfolder(dirPath).files
    End If
        
    FolderCount = objFolders.Count
    
    If FolderCount > 0 Then
    
        ReDim arrFolders(1 To FolderCount)
        Dim x As Long
        
        For Each objFolder In objFolders
            
            x = x + 1
            
            arrFolders(x) = objFolder.Name
        Next objFolder
        
    End If
    
    Set objFSO = Nothing
    Set objFolders = Nothing
    Set objFolder = Nothing

    getFolderLst = arrFolders

End Function

Function isEmptyArray(arr As Variant) As Boolean

    '*** has array been initialised ***
    
    If (Not arr) = -1 Then
    
        isEmptyArray = True
    Else
        
        isEmptyArray = False
    End If
    
End Function

Sub createFolder(fldName As String)
    
    '*** create a folder ***
    
    If fldName <> "" Then
    
        Dim objFSO As Object
      
        Set objFSO = CreateObject("Scripting.FileSystemObject")
    
        objFSO.createFolder fldName
          
        Set objFSO = Nothing
    End If

End Sub

Sub deleteFolder(fldName As String)

    '*** delete a folder ***
    
    If fldName <> "" Then
    
        Dim objFSO As Object
      
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        
        objFSO.deleteFolder fldName, False
    
        Set objFSO = Nothing
    End If

End Sub

Sub deleteFile(fileName As String)

    '*** delete a file ***

    If fileName <> "" Then
    
        Kill fileName
        
    End If

End Sub

Sub moveFile(Sourcefile As String, destfile As String)

    '*** move a file ***

    If Sourcefile <> "" And destfile <> "" Then
    
        Dim fso As Object
        
        Set fso = CreateObject("scripting.filesystemobject")
    
        fso.moveFile Source:=Sourcefile, Destination:=destfile
    End If

End Sub

Sub copyFile(Sourcefile As String, destfile As String)

    '*** copy a file ***
    
    If Sourcefile <> "" And destfile <> "" Then
    
        Dim fso As Object
        
        Set fso = CreateObject("scripting.filesystemobject")
    
        fso.copyFile Source:=Sourcefile, Destination:=destfile
    End If

End Sub

Sub changeFolderName(fldName As String, newfldname As String)

    '*** change a folder name ***

    If fldName <> "" And newfldname <> "" Then
    
        Name fldName As newfldname
        
    End If

End Sub

Function selectFolder(prompt As String) As String

    '*** select a folder ***
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        
        .AllowMultiSelect = False
        .Title = prompt
        
        If .Show <> 0 Then
            
            selectFolder = .SelectedItems(1)
        Else
            
            '*** file dialog cancelled - do nothing ***
        End If
    
    End With

End Function

Function getDaySuffix(dayInt As Integer) As String

    Select Case dayInt

        Case 1, 21, 31: getDaySuffix = "st"
        Case 2, 22: getDaySuffix = "nd"
        Case 3, 23: getDaySuffix = "rd"
        Case Else: getDaySuffix = "th"
    End Select

End Function

