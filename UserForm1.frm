VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "pfReportBuilder 3.0"
   ClientHeight    =   8964.001
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

    INITIALISE
    ADDIN_INSTALLED
    SELECT_TEMPLATE
    EDIT_TEMPLATE
    CONTENT_SELECT
End Enum

'*** declare our global variables ***
Dim rootFolder As String
Dim currentFolder As String
Dim currentTemplate As String

'*** declare our global constants ***
Const MASTER_TEMPLATE = "Master Template.rpt"
Const ADDIN_FILE_NAME = "pfReportBuilder.docm"

Private Sub ContentSelectButton_Click()
     
    WriteToTextFile currentTemplate, rootFolder & "\", ListBox3.Value

    ListBox4.Clear
    ListBox4.List = GetTextFile(currentTemplate, rootFolder & "\")
    
End Sub

Private Sub CreateTemplateButton_Click()

    Dim x As Long
    
    x = 1
    
    While FileExists(rootFolder & "\" & "New Template" & LTrim$(Str$(x)) & ".rep") = True
    
        x = x + 1
    Wend
    
    CreateTextFile "New Template" & LTrim$(Str$(x)) & ".rep", rootFolder & "\"
    
    '*** load listbox with names of rep files ***
    ListBox1.Clear
    ListBox1.List = getFileLst(rootFolder, "rep")
    
End Sub

Private Sub EditTemplateButton_Click()

    Label11.Caption = ListBox1.Value
    ListBox4.Clear
    ListBox4.List = GetTextFile(ListBox1.Value, rootFolder & "\")
    currentTemplate = ListBox1.Value
    SetButtonMode EDIT_TEMPLATE
    
End Sub

Private Sub InstallAsAddinButton_Click()

    If MsgBox("Install pfReportBuilder as addin ?", vbYesNo, "Hey will Robinson") = vbYes Then
    
        InstallAddin (ActiveDocument.Name)
        SetButtonMode ADDIN_INSTALLED
    End If
    
End Sub

Private Sub ListBox1_Click()
    
    ListBox2.Clear
    ListBox2.List = GetTextFile(ListBox1.Value, rootFolder & "\")
 
    SetButtonMode SELECT_TEMPLATE
    
End Sub

Private Sub ListBox3_Click()

    SetButtonMode CONTENT_SELECT

End Sub

Private Sub UserForm_Initialize()

     '*** Check if pfReportBuilder installed as addin ***
    Dim oAddin As AddIn
    
    For Each oAddin In AddIns
 
        If oAddin = ADDIN_FILE_NAME Then
            
            Label4.Caption = "pfReportBuilder Addin Installed"
            SetButtonMode ADDIN_INSTALLED
        End If
        
    Next oAddin
    
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
    
    '*** load listbox with names of rep files ***
    ListBox1.List = getFileLst(rootFolder, "rep")
    
    '*** load listbox with names of docx files ***
    ListBox3.List = getFileLst(rootFolder, "docx")
        
    SetButtonMode INITIALISE
    
End Sub

Private Sub SetButtonMode(bm As ButtonMode)

    Select Case bm
    
        Case INITIALISE:
            ContentSelectButton.Enabled = False
            EditTemplateButton.Enabled = False
            CreateTemplateButton.Enabled = True
            InstallAsAddinButton.Enabled = True
            MultiPage1.Value = 0
    
        Case ADDIN_INSTALLED:
            InstallAsAddinButton.Enabled = False
        
        Case SELECT_TEMPLATE:
            EditTemplateButton.Enabled = True
        
        Case EDIT_TEMPLATE:
            MultiPage1.Value = 1
            ContentSelectButton.Enabled = False
        
        Case CONTENT_SELECT:
            ContentSelectButton.Enabled = True
        
    End Select
 
End Sub

Private Sub changeTemplateFolderName()

    Dim fldName As String
    Dim newfldname As String
    
    fldName = rootFolder & "\" & ListBox1.Value
    
    newfldname = InputBox("Enter name of report folder", "Change Report Folder Name", "")
    
    If newfldname <> "" Then
        
        newfldname = rootFolder & "\" & newfldname
    
        changeFolderName fldName, newfldname
    
        ListBox1.Clear
        'ListBox1.List = getFolderLst(rootFolder, False)
    End If
    
End Sub

Private Sub deleteTemplateFolder()
            
    Dim fldName As String
    
    fldName = ListBox1.Value
    
    If fldName <> "" Then
    
        fldName = rootFolder & "\" & fldName
        
        deleteFolder fldName
        
        ListBox1.Clear
        'ListBox1.List = getFolderLst(rootFolder, False)
        ListBox1.Selected(1) = False
            
    End If
    
End Sub

Private Sub deleteReportTemplate()

    Dim fileName As String
    
    fileName = ListBox2.Value
    
    If fileName <> "" Then
    
        fileName = currentFolder & "\" & fileName
        
        If MsgBox("Are you sure you want to delete this template ?", vbOKCancel, "Hey Will Robinson !!!") = vbOK Then
        
            deleteFile fileName
        
            ListBox2.Clear
            'ListBox2.List = getFolderLst(currentFolder, False)
            ListBox2.Selected(1) = False
        End If
        
    End If

End Sub

Private Sub moveReportTemplate()

    Dim destfile As String
    
    destfile = selectFolder("Select file to move")
    
    If destfile <> "" Then

        Dim Sourcefile As String
        
        Sourcefile = currentFolder & "\" & ListBox2.Value
        destfile = rootFolder & "\" & destfile
    
        moveFile Sourcefile, destfile
        
        ListBox2.Clear
        'ListBox2.List = getFolderLst(currentFolder, True)
    End If

End Sub

Private Sub copyReportTemplate()

    Dim destfile As String
    
    destfile = selectFolder("Select file to copy")
    
    If destfile <> "" Then

        Dim Sourcefile As String
        
        Sourcefile = currentFolder & "\" & ListBox2.Value
        destfile = rootFolder & "\" & destfile
    
        copyFile Sourcefile, destfile
        
        ListBox2.Clear
        'ListBox2.List = getFolderLst(currentFolder, True)
    End If

End Sub

Private Sub changeReportTemplateName()

    Dim fileName As String
    Dim newfilename As String
    
    fileName = currentFolder & "\" & ListBox2.Value
    
    newfilename = InputBox("Enter name of report template", "Change Report template Name", "")
    
    If newfilename <> "" Then
        
        newfilename = currentFolder & "\" & newfilename
    
        changeFolderName fileName, newfilename
    
        ListBox2.Clear
        'ListBox2.List = getFolderLst(currentFolder, True)
    End If
 
End Sub

Private Sub setRootFolder()
              
    Dim fldName As String
    
    fldName = selectFolder("Select root folder")
    
    If fldName <> "" Then
        
        ListBox1.Clear
        'ListBox1.List = getFolderLst(fldName, False)
    
        rootFolder = fldName
        currentFolder = fldName
        
        Label1.Caption = rootFolder
        
        ActiveDocument.Variables("Root").Value = fldName
    End If

End Sub

Private Sub removeReportSection()
    
    ListBox3.RemoveItem (ListBox3.ListIndex)
    CommandButton3.Enabled = False

End Sub

Private Sub addReportSection()
    
    ListBox3.AddItem ListBox2.Value
    
    CommandButton3.Enabled = True

End Sub

Private Sub promoteReportSection()

    Dim fileName As String

    fileName = ListBox3.List(ListBox3.ListIndex)
    
    ListBox3.List(ListBox3.ListIndex, 0) = ListBox3.List(ListBox3.ListIndex - 1, 0)
    ListBox3.List(ListBox3.ListIndex - 1, 0) = fileName
    
    ListBox3.ListIndex = ListBox3.ListIndex - 1

End Sub

Private Sub demoteReportSection()

    Dim fileName As String

    fileName = ListBox3.List(ListBox3.ListIndex)
    ListBox3.List(ListBox3.ListIndex, 0) = ListBox3.List(ListBox3.ListIndex + 1, 0)
    ListBox3.List(ListBox3.ListIndex + 1, 0) = fileName
    
    ListBox3.ListIndex = ListBox3.ListIndex + 1

End Sub

Private Sub selectReportSection()
    
    '*** contrain promotion/demotion buttons ****
    If ListBox3.ListIndex > 0 Then CommandButton9.Enabled = True Else CommandButton9.Enabled = False
    If ListBox3.ListIndex < ListBox3.ListCount - 1 Then CommandButton10.Enabled = True Else CommandButton10.Enabled = False
    
    CommandButton3.Enabled = True

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
    
    docNew.SaveAs fileName:=newReportName
    
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
        
        Selection.InsertFile fileName:="""" & currentFolder & "\" & docPath & """", ConfirmConversions:=False, Link:=False, Attachment:=False
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
    
    exitProgram

End Sub

Private Sub exitProgram()

    '*** provides a central stub to exit the program, which may be useful for future development purposes ***

    End

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

