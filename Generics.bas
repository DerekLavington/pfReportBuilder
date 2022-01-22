Attribute VB_Name = "Generics"
Option Explicit

Enum TextFileModeType

    WRITE_FILE
    APPEND_FILE
End Enum

Function DocVarExists(varName As String) As Boolean

    '*** If MS Word document variable exists, return true ***
    
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
    
    '*** Return the path to the 'My Documents' directory ***
    
    Dim Wshshell As Object
    Set Wshshell = CreateObject("WScript.Shell")
    
    mydocs = Wshshell.SpecialFolders("MyDocuments")

End Function

Sub LoadFileNamesToListBox(dirpath As String, filesuffix As String, lbox As listbox)
    
    '*** Clear the listbox contents ***
    lbox.Clear
    
    '*** Check if folder exists ***
    Dim oFileSystem As New FileSystemObject
    
    If oFileSystem.FolderExists(dirpath) = True Then
             
        '*** if folder exists, load listbox with filenames matching the required suffix ***
        Dim oFolder As folder
        Set oFolder = oFileSystem.GetFolder(dirpath)
      
        Dim oFile As file
    
        For Each oFile In oFolder.FILES
            
            If Right$(oFile.Name, Len(filesuffix) + 1) = "." & filesuffix Then lbox.AddItem oFile.Name
        Next oFile
        
        Set oFile = Nothing
        Set oFolder = Nothing
    End If
     
End Sub

Function isEmptyArray(arr As Variant) As Boolean

    '*** if array has been initialised, return true ***
    
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

Sub deleteFile(filename As String)

    '*** delete a file ***

    If filename <> "" Then
    
        Kill filename
        
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

    '*** select a folder using MS Dialog ***
    
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

    '*** get the day st, nd, rd, th suffix ***
    
    Select Case dayInt

        Case 1, 21, 31: getDaySuffix = "st"
        Case 2, 22: getDaySuffix = "nd"
        Case 3, 23: getDaySuffix = "rd"
        Case Else: getDaySuffix = "th"
    End Select

End Function

Function FileExists(filename As String) As Boolean
    
    '*** if file exists, return true ***
    
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    FileExists = fs.FileExists(filename)

End Function

Sub CreateTextFile(filename As String, path As String)
    
    '*** create an empty text file ***
    
    Dim fs As Object, tf As Object

    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Set tf = fs.CreateTextFile(path & filename, True)

    tf.Close

End Sub

Sub ResetTextFile(filename As String, path As String)

    '*** reset an existing text file ***

    Dim fs As Object, tf As Object

    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Set tf = fs.OpenTextFile(path & filename, ForWriting, TristateFalse)
    
    tf.Close

End Sub

Sub WriteToTextFile(filename As String, path As String, text As String)

    '*** append to an existing text file ***

    Dim fs As Object, tf As Object

    Set fs = CreateObject("Scripting.FileSystemObject")
            
    Set tf = fs.OpenTextFile(path & filename, ForAppending, TristateFalse)
        
    tf.WriteLine (text)
    
    tf.Close

End Sub

Function GetArrayFromTextFile(filename As String, path As String) As String()

    '*** return text file as array ***

    Dim TxtLine() As String
    Dim index As Long
    
    Dim fs As Object, tf As Object
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set tf = fs.OpenTextFile(path & filename, 1)
    
    Do Until tf.AtEndOfStream = True
        
        tf.readline
        index = index + 1
    Loop
    
    ReDim TxtLine(index)
    
    index = 0
    
    Set tf = fs.OpenTextFile(path & filename, 1)
    
    Do Until tf.AtEndOfStream = True
        
        TxtLine(index) = tf.readline
        index = index + 1
    Loop
    
    tf.Close

    GetArrayFromTextFile = TxtLine

End Function

Sub SaveArrayToTextFile(filename As String, pathname As String, ar() As String)

    '*** save array to text file ***
    
    ResetTextFile filename, pathname
    
    Dim x As Long
    
    For x = 0 To UBound(ar)
    
        WriteToTextFile filename, pathname, ar(x)
    Next x

End Sub
Sub InstallAddin(filename As String)
    
    '*** install MS Word Addin ***
    
    If MsgBox("Install pfReportBuilder as addin ?", vbYesNo, "Hey will Robinson") = vbYes Then
    
        AddIns.Add ActiveDocument.Name, Install:=True
    End If

End Sub

Function AddinInstalled(filename As String) As Boolean

    '*** if addin installed, return true ***
    
    Dim oAddin As AddIn
    
    For Each oAddin In AddIns
 
        If oAddin = filename Then
            
            AddinInstalled = True
            Exit Function
        End If
        
    Next oAddin

End Function

Function GetFileNamePrefix(filename) As String

    '*** get file name prefix ***

    GetFileNamePrefix = Left$(filename, InStr(filename, ".") - 1)

End Function

Function GetFileNameSuffix(filename) As String
  
    '*** get file name suffix ***

    GetFileNameSuffix = Right$(filename, Len(filename) - InStr(filename, "."))

End Function

Function CreateFileName(filename As String, path As String) As String
    
    '*** create unique file name ***
    
    If FileExists(path & filename) = True Then
        
        Dim suffix As String
        Dim prefix As String
    
        prefix = GetFileNamePrefix(filename)
        suffix = GetFileNameSuffix(filename)

        Dim x As Long
        
        Do
        
            x = x + 1

        Loop Until FileExists(path & prefix & LTrim(x) & "." & suffix) = False

        CreateFileName = path & prefix & LTrim(x) & "." & suffix
    Else

        CreateFileName = filename
    End If

End Function

Sub ListBoxPromoteSelectedItem(lbox As listbox)

    '*** promote listbox item ***

    If lbox.ListIndex > 0 Then
        
        Dim z As String
    
        z = lbox.List(lbox.ListIndex - 1)
        lbox.List(lbox.ListIndex - 1) = lbox.List(lbox.ListIndex)
        lbox.List(lbox.ListIndex) = z
    End If

End Sub

Sub ListBoxDemoteSelectedItem(lbox As Control)

    '*** demote listbox item ***

    If lbox.ListIndex < lbox.ListCount Then
        
        Dim z As String
    
        z = lbox.List(lbox.ListIndex + 1)
        lbox.List(lbox.ListIndex + 1) = lbox.List(lbox.ListIndex)
        lbox.List(lbox.ListIndex) = z
    End If

End Sub

Sub SaveListBoxToTextFile(filename As String, pathname As String, lbox As listbox)

    '*** save listbox content to text file ***

    Dim x As Long
    
    ResetTextFile filename, pathname
    
    For x = 0 To lbox.ListCount - 1
    
        Debug.Print x, lbox.List(x)
        WriteToTextFile filename, pathname, lbox.List(x)
    
    Next x

End Sub

Function LoadTextFileToListBox(filename As String, pathname As String, lbox As listbox) As String()

    '*** load text file to listbox ***
    lbox.Clear
    lbox.List = GetArrayFromTextFile(filename, pathname)
    
End Function

