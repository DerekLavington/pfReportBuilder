VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "pfReportBuilder 3.0"
   ClientHeight    =   9636.001
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

    Dim folderName As String
    
    folderName = selectFolder("Select folder")
    
    If folderName <> "" Then
        
        rootFolder = folderName
        currentFolder = folderName

        Label12.Caption = "Folder = " & rootFolder
    
        '*** load listboxes ***
        ListBox1.List = getFileLst(rootFolder, TEMPLATE_FILE_SUFFIX)
        
        '*** Set the document variable to the new folder ***
        ActiveDocument.Variables("Root").Value = folderName
    End If

    SetButtonMode SELECT_TEMPLATE_PAGE

End Sub

Private Sub ChangeTemplateNameButton_Click()

    Dim NewTemplateName As String
    
    NewTemplateName = InputBox("Enter name of template (must end in ." & TEMPLATE_FILE_SUFFIX & ")", "Change Template Name", "")
    
    If NewTemplateName <> "" And GetFileNameSuffix(NewTemplateName) = TEMPLATE_FILE_SUFFIX Then
        
        NewTemplateName = CreateFileName(NewTemplateName, rootFolder & "\")
        
        changeName rootFolder & "\" & currentTemplate, rootFolder & "\" & NewTemplateName
    
        ListBox1.List(ListBox1.ListIndex) = NewTemplateName
        
        SaveListBoxToFile ListBox1, NewTemplateName, rootFolder
    Else
    
        MsgBox "Invalid File Name", vbOKOnly, "Danger Will Robinson"
    End If
    
    SetButtonMode SELECT_TEMPLATE_PAGE
 
End Sub

Private Sub ContentDeselectButton_Click()
            
    ListBox4.RemoveItem ListBox4.ListIndex
    
    SaveListBoxToFile ListBox4, currentTemplate, rootFolder & "\"
    
    SetButtonMode EDIT_TEMPLATE_PAGE

End Sub

Private Sub CopyTemplateButton_Click()

    Dim NewTemplateName As String
    
    NewTemplateName = InputBox("Enter name of template (must end in ." & TEMPLATE_FILE_SUFFIX & ")", "Change Template Name", "")
    
    If NewTemplateName <> "" And GetFileNameSuffix(NewTemplateName) = TEMPLATE_FILE_SUFFIX Then
        
        NewTemplateName = CreateFileName(NewTemplateName, rootFolder & "\")
    
        copyFile rootFolder & "\" & currentTemplate, rootFolder & "\" & NewTemplateName
    
        currentTemplate = NewTemplateName
        
        Label11.Caption = rootFolder & "\" & currentTemplate
    Else
    
        MsgBox "Invalid File Name", vbOKOnly, "Danger Will Robinson"
    End If
    
    SetButtonMode EDIT_TEMPLATE_PAGE

End Sub

Private Sub CreateTemplateButton_Click()
    
    Dim FileName As String
    
    FileName = CreateFileName("New Template.rep", rootFolder & "\")
    
    CreateTextFile FileName, ""

    currentTemplate = FileName
    
    ListBox3.Clear
    ListBox4.Clear
    
    Label11.Caption = rootFolder & "\" & currentTemplate

    SetButtonMode EDIT_TEMPLATE_PAGE

End Sub

Private Sub DeleteTemplateButton_Click()

    deleteFile rootFolder & "\" & currentTemplate
    
    ListBox1.RemoveItem ListBox1.ListIndex
 
    SaveListBoxToFile ListBox1, currentTemplate, rootFolder & "\"
    
    SetButtonMode SELECT_TEMPLATE_PAGE
    
End Sub

Private Sub ExitButton_Click()

    If MsgBox("Are you sure you want to exit?", vbYesNo, "Hey Will Robinson") = vbYes Then End

End Sub

Private Sub GoToEditTemplatePageButton_Click()
    
    Label11.Caption = rootFolder & "\" & currentTemplate
    
    ListBox3.Clear
    ListBox3.List = getFileLst(rootFolder, DOCUMENT_FILE_SUFFIX)
    
    ListBox4.Clear
    ListBox4.List = GetTextFile(currentTemplate, rootFolder & "\")
    
    SetButtonMode EDIT_TEMPLATE_PAGE
    MultiPage1.Value = 1

End Sub

Private Sub GoToSelectTemplatePageButton_Click()

    Label12.Caption = "Folder = " & rootFolder
    
    ListBox1.Clear
    ListBox1.List = getFileLst(rootFolder, TEMPLATE_FILE_SUFFIX)
       
    SetButtonMode SELECT_TEMPLATE_PAGE
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
    
    SaveListBoxToFile ListBox4, currentTemplate, rootFolder
    
    SetButtonMode EDIT_TEMPLATE_PAGE

End Sub

Private Sub TemplateContentPromoteButton_Click()
    
    ListBoxPromoteSelectedItem ListBox4
    
    SaveListBoxToFile ListBox4, currentTemplate, rootFolder & "\"
    
    SetButtonMode EDIT_TEMPLATE_PAGE

End Sub

Private Sub TemplateContentSelectButton_Click()
  
    ListBox4.AddItem ListBox3.List(ListBox3.ListIndex)
    
    SaveListBoxToFile ListBox4, currentTemplate, rootFolder & "\"
    
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

    Label12.Caption = "Folder = " & rootFolder
    
    '*** load listboxes ***
    ListBox1.List = getFileLst(rootFolder, TEMPLATE_FILE_SUFFIX)
    ListBox3.List = getFileLst(rootFolder, DOCUMENT_FILE_SUFFIX)
        
    SetButtonMode SELECT_TEMPLATE_PAGE
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

        Case SELECT_TEMPLATE_CONTENT:
            
            TemplateContentSelectButton.Enabled = True
            TemplateContentDeselectButton.Enabled = True
            TemplateContentPromoteButton.Enabled = True
            TemplateContentDemoteButton.Enabled = True
        
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
    
    docNew.SaveAs FileName:=newReportName
    
    sectioncount = ListBox3.ListCount - 1
    
    ProgressBar1.Min = 0
    ProgressBar1.Max = sectioncount
    ProgressBar1.Scrolling = ccScrollingSmooth
        
    'Selection.InsertFile fileName:="""" & currentFolder & "\" & "Header.doc" & """", ConfirmConversions:=False, Link:=False, Attachment:=False
    'Selection.InsertBreak Type:=wdPageBreak
    
    'Selection.InsertFile fileName:="""" & currentFolder & "\" & "Contentr.doc" & """", ConfirmConversions:=False, Link:=False, Attachment:=False
    'Selection.InsertBreak Type:=wdPageBreak
    
    For i = 0 To sectioncount
    
        docPath = ListBox3.Column(0, i)
        
        Selection.InsertFile FileName:="""" & currentFolder & "\" & docPath & """", ConfirmConversions:=False, Link:=False, Attachment:=False
        Selection.InsertBreak Type:=wdPageBreak
        
        ProgressBar1.Value = i
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
