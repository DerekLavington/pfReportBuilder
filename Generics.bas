Attribute VB_Name = "Generics"
Option Explicit

Enum TextFileModeType

    WRITE_FILE
    APPEND_FILE
End Enum

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

Function getFileLst(dirpath As String, filesuffix As String) As String()

    '*** Return an array of file names ***
    
    Dim oFileSystem As New FileSystemObject
    
    If oFileSystem.FolderExists(dirpath) = True Then
             
        Dim fileCount As Long
        
        Dim oFolder As folder
        Set oFolder = oFileSystem.GetFolder(dirpath)
        
        If oFolder.FILES.Count > 0 Then
        
            Dim oFile As file
            
            For Each oFile In oFolder.FILES
            
                If Right$(oFile.Name, Len(filesuffix)) = filesuffix Then fileCount = fileCount + 1

            Next oFile
            
            Dim StrArr() As String
            ReDim StrArr(fileCount)
    
            Dim x As Integer
    
            For Each oFile In oFolder.FILES
            
                If Right$(oFile.Name, Len(filesuffix)) = filesuffix Then
                    
                    StrArr(x) = oFile.Name:
                    x = x + 1
                End If
                
            Next oFile
        
            Set oFile = Nothing
            
            getFileLst = StrArr
        End If
        
        Set oFolder = Nothing
    End If
     
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

Sub deleteFile(FileName As String)

    '*** delete a file ***

    If FileName <> "" Then
    
        Kill FileName
        
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

Sub changeName(fldName As String, newfldname As String)

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

Function FileExists(FileName As String) As Boolean
    
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    FileExists = fs.FileExists(FileName)

End Function

Sub CreateTextFile(FileName As String, path As String)
    
    Dim fs As Object, tf As Object

    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Set tf = fs.CreateTextFile(path & FileName, True)

    tf.Close

End Sub

Sub ResetTextFile(FileName As String, path As String)

    Dim fs As Object, tf As Object

    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Set tf = fs.OpenTextFile(path & FileName, ForWriting, TristateFalse)
    
    tf.Close

End Sub

Sub WriteToTextFile(FileName As String, path As String, text As String)

    Dim fs As Object, tf As Object

    Set fs = CreateObject("Scripting.FileSystemObject")
            
    Set tf = fs.OpenTextFile(path & FileName, ForAppending, TristateFalse)
        
    tf.WriteLine (text)
    
    tf.Close

End Sub

Function GetTextFile(FileName As String, path As String) As String()

    Dim TxtLine() As String
    Dim index As Long
    
    Dim fs As Object, tf As Object
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set tf = fs.OpenTextFile(path & FileName, 1)
    
    Do Until tf.AtEndOfStream = True
        
        tf.readline
        index = index + 1
    Loop
    
    ReDim TxtLine(index)
    
    index = 0
    
    Set tf = fs.OpenTextFile(path & FileName, 1)
    
    Do Until tf.AtEndOfStream = True
        
        TxtLine(index) = tf.readline
        index = index + 1
    Loop
    
    tf.Close

    GetTextFile = TxtLine

End Function

Sub InstallAddin(FileName As String)
    
    If MsgBox("Install pfReportBuilder as addin ?", vbYesNo, "Hey will Robinson") = vbYes Then
    
        AddIns.Add ActiveDocument.Name, Install:=True
    End If

End Sub

Function AddinInstalled(FileName As String) As Boolean

    Dim oAddin As AddIn
    
    For Each oAddin In AddIns
 
        If oAddin = FileName Then
            
            AddinInstalled = True
            Exit Function
        End If
        
    Next oAddin

End Function


Function GetFileNamePrefix(FileName) As String

    GetFileNamePrefix = Left$(FileName, InStr(FileName, ".") - 1)

End Function

Function GetFileNameSuffix(FileName) As String

    GetFileNameSuffix = Right$(FileName, Len(FileName) - InStr(FileName, "."))

End Function

Function CreateFileName(FileName As String, path As String) As String
    
    If FileExists(path & FileName) = True Then
        
        Dim suffix As String
        Dim prefix As String
    
        prefix = GetFileNamePrefix(FileName)
        suffix = GetFileNameSuffix(FileName)

        Dim x As Long
        
        Do
        
            x = x + 1

        Loop Until FileExists(path & prefix & LTrim(x) & "." & suffix) = False

        CreateFileName = path & prefix & LTrim(x) & "." & suffix
    Else

        CreateFileName = FileName
    End If

End Function

Sub ListBoxPromoteSelectedItem(lbox As listbox)

    If lbox.ListIndex > 0 Then
        
        Dim z As String
    
        z = lbox.List(lbox.ListIndex - 1)
        lbox.List(lbox.ListIndex - 1) = lbox.List(lbox.ListIndex)
        lbox.List(lbox.ListIndex) = z
    End If

End Sub

Sub ListBoxDemoteSelectedItem(lbox As listbox)

    If lbox.ListIndex < lbox.ListCount Then
        
        Dim z As String
    
        z = lbox.List(lbox.ListIndex + 1)
        lbox.List(lbox.ListIndex + 1) = lbox.List(lbox.ListIndex)
        lbox.List(lbox.ListIndex) = z
    End If

End Sub

Sub SaveListBoxToFile(lbox As listbox, FileName As String, pathname As String)

    Dim x As Long
    
    ResetTextFile FileName, pathname
    
    For x = 0 To lbox.ListCount
    
        WriteToTextFile FileName, pathname, lbox.List(x)
    
    Next x

End Sub

