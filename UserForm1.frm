VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "pfReportBuilder 3.0"
   ClientHeight    =   9576.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12780
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*** declare our enumerators ***
Enum ButtonMode

    ADDIN_INSTALLED
    ADDIN_NOT_INSTALLED
    SELECT_TEMPLATE_PAGE
    SELECT_TEMPLATE
    EDIT_TEMPLATE_PAGE
    SELECT_TEMPLATE_CONTENT
End Enum

'*** declare our global variables ***
Dim rootFolder As String
Dim currentFolder As String
Dim currentTemplate As String

'*** declare our global constants ***
Const ADDIN_FILE_NAME = "pfReportBuilder.docm"
Const TEMPLATE_FILE_SUFFIX = "rep"
Const DOCUMENT_FILE_SUFFIX = "docx"

Private Sub ChangeFolderButton_Click()
    
    '*** Get the new root folder from the user ***
    Dim folderName As String
    folderName = selectFolder("Select folder")
    
    '*** Check we have selected a valid folder ***
    If folderName <> "" Then
        
        '*** Set the new root folder ***
        rootFolder = folderName
        currentFolder = folderName

        '*** Display new folder path ***
        Label12.Caption = "Folder = " & rootFolder
    
        '*** Reload list of templates to ListBox ***
        LoadFileNamesToListBox rootFolder, TEMPLATE_FILE_SUFFIX, ListBox1
        
        '*** Clear residual selected template content from ListBox ***
        ListBox2.Clear
        
        '*** Set the document variable to the new folder ***
        ActiveDocument.Variables("Root").Value = folderName
    End If

    '*** Reset the GUI buttons ***
    SetButtonMode SELECT_TEMPLATE_PAGE

End Sub

Private Sub ChangeTemplateNameButton_Click()

    Dim filename As String
    
    '*** get the new file name from the user ***
    filename = InputBox("Enter name of template", "Change Template Name", "")
    
    '*** check that a valid file name was selected ***
    If filename <> "" Then
            
        '*** Check that new file name has correct suffix and, if not, add one ***
        If InStr(filename, ".") > 0 Then
        
            If GetFileNameSuffix(filename) <> "rep" Then filename = GetFileNamePrefix(filename) & "." & TEMPLATE_FILE_SUFFIX
        Else
        
            filename = filename & "." & TEMPLATE_FILE_SUFFIX
        End If
        
        '*** Make sure file name is unique ***
        filename = CreateFileName(filename, rootFolder & "\")
    
        '*** change the file name ***
        changeName rootFolder & "\" & currentTemplate, rootFolder & "\" & filename
        
        '*** set the current template to the new file name ***
        currentTemplate = filename
        
        '*** reload the list box with the changed template name ***
        ListBox1.List(ListBox1.ListIndex) = filename
    End If
    
    '*** Reset the GUI buttons ***
    SetButtonMode SELECT_TEMPLATE_PAGE

End Sub

Private Sub ContentDeselectButton_Click()
            
    ListBox4.RemoveItem ListBox4.ListIndex
    
    SaveListBoxToTextFile currentTemplate, rootFolder & "\", ListBox4
    
    SetButtonMode EDIT_TEMPLATE_PAGE

End Sub

Private Sub ChangeContentNameButton_Click()

Dim filename As String
    
    '*** get the new file name from the user ***
    filename = InputBox("Enter name of content template", "Change content template name", "")
    
    '*** check that a valid file name was selected ***
    If filename <> "" Then
            
        '*** Check that new file name has correct suffix and, if not, add one ***
        If InStr(filename, ".") > 0 Then
        
            If GetFileNameSuffix(filename) <> "docx" Then filename = GetFileNamePrefix(filename) & "." & DOCUMENT_FILE_SUFFIX
        Else
        
            filename = filename & "." & DOCUMENT_FILE_SUFFIX
        End If
        
        '*** Make sure file name is unique ***
        filename = CreateFileName(filename, rootFolder & "\")
    
        '*** change the file name ***
        changeName rootFolder & "\" & currentTemplate, rootFolder & "\" & filename
        
        '*** change file name in report file ***
        Dim f() As String
        Dim flen As Long
        
        flen = UBound(GetArrayFromTextFile(currentTemplate, rootFolder & "\"))
        
        ReDim f(flen)
        f = GetArrayFromTextFile(currentTemplate, rootFolder & "\")
        
        Dim x As Long
        
        For x = 0 To flen - 1
        
            If f(x) = currentTemplate Then
                
                f(x) = filename
                
                SaveArrayToTextFile currentTemplate, rootFolder, f
                
                Exit For
            End If
        Next x
        
        '*** Load the content select listbox ***
        ListBox3.Clear
        LoadFileNamesToListBox rootFolder & "\", DOCUMENT_FILE_SUFFIX, ListBox3
    
        '*** Load the template content listbox ***
        ListBox4.Clear
        ListBox4.List = GetArrayFromTextFile(currentTemplate, rootFolder & "\")
    End If
    
    '*** Reset the GUI buttons ***
    SetButtonMode SELECT_TEMPLATE_PAGE

End Sub

Private Sub CopyTemplateButton_Click()

    Dim filename As String
    
    '*** Get the new file name from the user ***
    filename = InputBox("Enter name of template", "New Template Name", "")
    
    '*** Check that a valid file name was selected ***
    If filename <> "" Then
            
        '*** Check that new file name has correct suffix and, if not, add one ***
        If InStr(filename, ".") > 0 Then
        
            If GetFileNameSuffix(filename) <> "rep" Then filename = GetFileNamePrefix(filename) & "." & TEMPLATE_FILE_SUFFIX
        Else
        
            filename = filename & "." & TEMPLATE_FILE_SUFFIX
        End If
        
        '*** Make sure file name is unique ***
        filename = CreateFileName(filename, rootFolder & "\")
    
        '*** Copy selected file to new file name ***
        copyFile rootFolder & "\" & currentTemplate, filename
    
        '*** Set the current template to the new file name ***
        currentTemplate = filename
        
        '*** Load the listbox with the new file name ***
        ListBox1.AddItem currentTemplate
    End If
    
    '*** Reset the GUI buttons ***
    SetButtonMode SELECT_TEMPLATE_PAGE

End Sub

Private Sub CreateTemplateButton_Click()
    
    Dim filename As String
    
    filename = CreateFileName("New Template.rep", rootFolder & "\")
    
    CreateTextFile filename, rootFolder & "\"
    
    currentTemplate = filename
    
    ListBox1.AddItem currentTemplate
    
    SetButtonMode EDIT_TEMPLATE_PAGE

End Sub

Private Sub DeleteTemplateButton_Click()

    deleteFile rootFolder & "\" & currentTemplate
    
    ListBox1.RemoveItem ListBox1.ListIndex
  
    SetButtonMode SELECT_TEMPLATE_PAGE
    
End Sub

Private Sub ExitButton_Click()

    If MsgBox("Are you sure you want to exit?", vbYesNo, "Hey Will Robinson") = vbYes Then End

End Sub

Private Sub GoToEditTemplatePageButton_Click()
    
    '*** Display the selected template on the GUI ***
    Label11.Caption = rootFolder & "\" & currentTemplate
    
    '*** Load the content select listbox ***
    ListBox3.Clear
    LoadFileNamesToListBox rootFolder & "\", DOCUMENT_FILE_SUFFIX, ListBox3
    
    '*** Load the template content listbox ***
    LoadTextFileToListBox currentTemplate, rootFolder & "\", ListBox4
    
    '*** Reset the GUI buttons ***
    SetButtonMode EDIT_TEMPLATE_PAGE
    
    '*** Display the edit template GUI ***
    MultiPage1.Value = 1

End Sub

Private Sub GoToSelectTemplatePageButton_Click()
       
    '*** Reset the GUI buttons ***
    SetButtonMode SELECT_TEMPLATE_PAGE
    
    '*** Display the template select GUI ***
    MultiPage1.Value = 0

End Sub

Private Sub InstallAsAddinButton_Click()

    If MsgBox("Install pfReportBuilder as addin ?", vbYesNo, "Hey will Robinson") = vbYes Then
    
        InstallAddin (ActiveDocument.Name)
        SetButtonMode ADDIN_INSTALLED
    End If
    
End Sub

Private Sub ListBox1_Click()
    
    currentTemplate = ListBox1.Value
    
    LoadTextFileToListBox currentTemplate, rootFolder & "\", ListBox2
    
    SetButtonMode SELECT_TEMPLATE
    
End Sub

Private Sub ListBox3_Click()

    SetButtonMode SELECT_TEMPLATE_CONTENT
    
End Sub

Private Sub ListBox4_Click()

    SetButtonMode SELECT_TEMPLATE_CONTENT

End Sub

Private Sub SelectTemplateButton_Click()

    SetButtonMode SELECT_TEMPLATE

    MultiPage1.Value = 0

End Sub

Private Sub TemplateContentDemoteButton_Click()
    
    ListBoxDemoteSelectedItem ListBox4
    
    SaveListBoxToTextFile currentTemplate, rootFolder & "\", ListBox4
    
    SetButtonMode EDIT_TEMPLATE_PAGE

End Sub

Private Sub TemplateContentDeselectButton_Click()

    ListBox4.RemoveItem ListBox4.ListIndex

    SaveListBoxToTextFile currentTemplate, rootFolder & "\", ListBox4
    
    SetButtonMode EDIT_TEMPLATE_PAGE

End Sub

Private Sub TemplateContentPromoteButton_Click()
    
    ListBoxPromoteSelectedItem ListBox4
    
    SaveListBoxToTextFile currentTemplate, rootFolder & "\", ListBox4
    
    SetButtonMode EDIT_TEMPLATE_PAGE

End Sub

Private Sub TemplateContentSelectButton_Click()
  
    ListBox4.AddItem ListBox3.List(ListBox3.ListIndex)
    
    SaveListBoxToTextFile currentTemplate, rootFolder & "\", ListBox4
    
    SetButtonMode EDIT_TEMPLATE_PAGE

End Sub

Private Sub UserForm_Initialize()

     '*** Check if pfReportBuilder installed as addin ***
    If AddinInstalled(ADDIN_FILE_NAME) = True Then
    
        Label4.Caption = "pfReportBuilder Addin Installed"
        SetButtonMode ADDIN_INSTALLED
    Else
    
        SetButtonMode ADDIN_NOT_INSTALLED
    End If
     
    '*** get root folder name or set to My Documents if one doesn't exist ***
    Dim folder As String
    
    If DocVarExists("Root") = False Then
    
        folder = mydocs()
        ActiveDocument.Variables.Add Name:="Root", Value:=folder
    Else

        folder = ActiveDocument.Variables("Root").Value
    End If
    
    rootFolder = folder
    currentFolder = folder

    '*** Display root folder in GUI ***
    Label12.Caption = "Folder = " & rootFolder
    
    '*** Load list of template files to GUI ***
    LoadFileNamesToListBox rootFolder, TEMPLATE_FILE_SUFFIX, ListBox1
        
    '*** Set the button mode ***
    SetButtonMode SELECT_TEMPLATE_PAGE
    
    '*** Set GUI page to select template ***
    MultiPage1.Value = 0
    
End Sub

Private Sub SetButtonMode(bMode As ButtonMode)

    Select Case bMode
    
        Case ADDIN_NOT_INSTALLED:
            
            InstallAsAddinButton.Enabled = True
            
        Case ADDIN_INSTALLED:
            
            InstallAsAddinButton.Enabled = False
        
        Case SELECT_TEMPLATE_PAGE
        
            GoToEditTemplatePageButton = False
            DeleteTemplateButton.Enabled = False
            CopyTemplateButton.Enabled = False
            ChangeTemplateNameButton.Enabled = False
            ChangeFolderButton.Enabled = True
        
        Case SELECT_TEMPLATE
        
            GoToEditTemplatePageButton.Enabled = True
            'GoToSelectReportPageButton = False
            DeleteTemplateButton.Enabled = True
            CopyTemplateButton.Enabled = True
            ChangeTemplateNameButton.Enabled = True
            
        Case EDIT_TEMPLATE_PAGE:
        
            GoToSelectTemplatePageButton.Enabled = True
            CreateTemplateButton.Enabled = True
            TemplateContentSelectButton.Enabled = False
            TemplateContentDeselectButton.Enabled = False
            TemplateContentPromoteButton.Enabled = False
            TemplateContentDemoteButton.Enabled = False
            ChangeContentNameButton.Enabled = False

        Case SELECT_TEMPLATE_CONTENT:
            
            TemplateContentSelectButton.Enabled = True
            TemplateContentDeselectButton.Enabled = True
            TemplateContentPromoteButton.Enabled = True
            TemplateContentDemoteButton.Enabled = True
            ChangeContentNameButton.Enabled = True
        
    End Select
 
End Sub

Private Sub buildReport()

    '*** build the report from templates in the report listbox ***

    Dim i As Long
    Dim docPath As String
    Dim sectioncount As Integer
    Dim newReportName As String
    
    'Application.ScreenUpdating = False
    
    Dim docNew As Document
    Set docNew = Documents.Add
    
    'If TextBox1.Value = "" Then newReportName = "New Report.docx" Else newReportName = TextBox1.Value
    
    docNew.SaveAs filename:=newReportName
    
    sectioncount = ListBox3.ListCount - 1
    
    'ProgressBar1.Min = 0
    'ProgressBar1.Max = sectioncount
    'ProgressBar1.Scrolling = ccScrollingSmooth
        
    'Selection.InsertFile fileName:="""" & currentFolder & "\" & "Header.doc" & """", ConfirmConversions:=False, Link:=False, Attachment:=False
    'Selection.InsertBreak Type:=wdPageBreak
    
    'Selection.InsertFile fileName:="""" & currentFolder & "\" & "Contentr.doc" & """", ConfirmConversions:=False, Link:=False, Attachment:=False
    'Selection.InsertBreak Type:=wdPageBreak
    
    For i = 0 To sectioncount
    
        docPath = ListBox3.Column(0, i)
        
        Selection.InsertFile filename:="""" & currentFolder & "\" & docPath & """", ConfirmConversions:=False, Link:=False, Attachment:=False
        Selection.InsertBreak Type:=wdPageBreak
        
        'ProgressBar1.Value = i
    Next i

    Dim dateStr As String
    Dim dayInt As Integer
    Dim clientStr As String
    
    Documents(newReportName).activate
    
    dayInt = Day(Date)
    dateStr = Str$(dayInt) & getDaySuffix(dayInt) & " " & Format(Date, "MMMM YYYY")
    
    'If TextBox2.Value = "" Then clientStr = "New Client" Else clientStr = TextBox2
    
    ActiveDocument.CustomDocumentProperties.Add Name:="ClientName", LinkToContent:=False, Value:=clientStr, Type:=msoPropertyTypeString
    ActiveDocument.CustomDocumentProperties.Add Name:="ReportDate", LinkToContent:=False, Value:=dateStr, Type:=msoPropertyTypeString
    ActiveDocument.Fields.Update
    
    ResequenceSectionNumbers
    
    If ActiveDocument.TablesOfContents.Count > 0 Then ActiveDocument.TablesOfContents(1).Update

End Sub

Private Sub ResequenceSectionNumbers()

    Dim para As Paragraph
    Dim paracount As Integer
    Dim Section As Integer
    Dim x As Integer
    
    paracount = ActiveDocument.Paragraphs.Count
    
    For x = 1 To paracount
       
        If ActiveDocument.Paragraphs(x).Style = "Heading 1" Then
            
            Section = Section + 1
            ActiveDocument.Paragraphs(x).Range.Font.Size = 14
            ActiveDocument.Paragraphs(x).Range.Words(1) = LTrim(Str$(Section))
            
        End If
        
    Next x

End Sub
