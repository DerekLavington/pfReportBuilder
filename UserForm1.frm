VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "pfReportBuilder 2.1"
   ClientHeight    =   8964.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8628.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*** our 'TO DO' list ***
'Code activatesectionbuttons
'Code sign off names
'Code order templates in listbox
'debug change report template name

'*** declare our global variables ***
Dim rootFolder As String
Dim currentFolder As String

'*** create programatically as these are a bitch to edit if hardcoded into control properties ***
Const ReportSectionAddButtonTip = "Adds the selected template to the report."
Const ReportTemplateCopyButtonTip = "Creates a duplicate of the the selected template in the current folder."
Const ReportTemplateMoveButtonTip = "Moves the selected template to another folder."
Const ReportTemplateDeleteButtonTip = "Deletes the selected template from the current folder."
Const FolderCreateButtonTip = "Creates a new sub folder in the root folder."
Const FolderDeleteButtonTip = "Deletes the selected folder from the root folder"
Const FolderChangeNameButtonTip = "changes the name of the selected folder."
Const ReportSectionPromoteButtonTip = "Promotes the selected section in the report."
Const ReportSectionDemoteButtonTip = "Demotes the selected section in the report."
Const ReportSectionRemoveButtonTip = "Removes the selected section from the report."
Const ReportBuildButtonTip = "Builds the report"
Const ReportClearButtonTip = "Clears the report"

'********************************************************************************************************************
'*** The following subs relate to control events. Our convention is to refer any processing to a separate routine ***
'*** This is to avoid the faff of having to copy and paste large amounts of code between subs in the event that   ***
'*** the user interface needs to be redesigned, or worse, losing code where a control is accidently deleted.         ***
'********************************************************************************************************************

Private Sub CommandButton10_Click()

    demoteReportSection

End Sub

Private Sub CommandButton12_Click()
    
    changeReportTemplateName

End Sub

Private Sub CommandButton13_Click()
    
    ListBox3.Clear

End Sub

Private Sub CommandButton14_Click()

    buildReport
    
End Sub

Private Sub CommandButton15_Click()

    exitProgram

End Sub

Private Sub CommandButton16_Click()

End Sub

Private Sub CommandButton17_Click()

    moveReportTemplate

End Sub

Private Sub CommandButton18_Click()

    copyReportTemplate

End Sub

Private Sub CommandButton19_Click()
    
    addReportSection

End Sub

Private Sub CommandButton2_Click()

    deleteReportTemplate

End Sub

Private Sub CommandButton3_Click()

    removeReportSection

End Sub

Private Sub CommandButton4_Click()

    setRootFolder
   
End Sub

Private Sub CommandButton5_Click()

    deleteTemplateFolder

End Sub

Private Sub CommandButton6_Click()

    changeTemplateFolderName

End Sub

Private Sub CommandButton7_Click()

    createTemplateFolder

End Sub

Private Sub CommandButton9_Click()

    promoteReportSection
    
End Sub

Private Sub ListBox1_Click()

    selectTemplateFolder
    
End Sub

Private Sub ListBox2_Click()

    activateTemplateButtons True
    
End Sub

Private Sub ListBox3_Click()

    selectReportSection

End Sub

Private Sub UserForm_Initialize()

    initialiseApplication
    
End Sub

'***************************************************************************************************************
'*** The following subs relate to the processing of control event and act as stubs to the use of primitives. ***
'*** This ensures that the primitives can remain generic and reusuable in other applications.                ***
'***************************************************************************************************************

Private Sub activateFolderButtons(activate As Boolean)

    If activate = True Then
    
        CommandButton5.Enabled = True
        CommandButton6.Enabled = True
    Else
        
        CommandButton5.Enabled = False
        CommandButton6.Enabled = False
    End If
    
End Sub

Private Sub activateTemplateButtons(activate As Boolean)

    If activate = True Then
     
        CommandButton2.Enabled = True
        CommandButton12.Enabled = True
        CommandButton17.Enabled = True
        CommandButton18.Enabled = True
        CommandButton19.Enabled = True
    Else
        
        CommandButton2.Enabled = False
        CommandButton12.Enabled = False
        CommandButton17.Enabled = False
        CommandButton18.Enabled = False
        CommandButton19.Enabled = False
    End If

End Sub

Private Sub activateSectionButtons(activate As Boolean)

End Sub

Private Sub createTemplateFolder()

    Dim fldName As String
    
    fldName = InputBox("Enter name of report folder", "Create New Report Folder", "")
    
    If fldName <> "" Then
    
        fldName = rootFolder & "\" & fldName
    
        createFolder (fldName)
    
        ListBox1.Clear
        ListBox1.List = getFolderLst(rootFolder, False)
    End If
    
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
        ListBox1.List = getFolderLst(rootFolder, False)
    End If
    
    activateFolderButtons False

End Sub

Private Sub deleteTemplateFolder()
            
    Dim fldName As String
    
    fldName = ListBox1.Value
    
    If fldName <> "" Then
    
        fldName = rootFolder & "\" & fldName
        
        deleteFolder fldName
        
        ListBox1.Clear
        ListBox1.List = getFolderLst(rootFolder, False)
        ListBox1.Selected(1) = False
            
        activateFolderButtons False
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
            ListBox2.List = getFolderLst(currentFolder, False)
            ListBox2.Selected(1) = False
        End If
        
        activateTemplateButtons False
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
        ListBox2.List = getFolderLst(currentFolder, True)
    End If
    
    activateTemplateButtons False

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
        ListBox2.List = getFolderLst(currentFolder, True)
    End If

    activateTemplateButtons False

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
        ListBox2.List = getFolderLst(currentFolder, True)
    End If
    
    activateTemplateButtons False

End Sub

Private Sub setRootFolder()
              
    Dim fldName As String
    
    fldName = selectFolder("Select root folder")
    
    If fldName <> "" Then
        
        ListBox1.Clear
        ListBox1.List = getFolderLst(fldName, False)
    
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
    
    activateTemplateButtons False
    
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
    
    If TextBox1.Value = "" Then newReportName = "New Report.docx" Else newReportName = TextBox1.Value
    
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
    
    If TextBox2.Value = "" Then clientStr = "New Client" Else clientStr = TextBox2
    
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

Private Sub selectTemplateFolder()

    '*** selects a template from the template folder ***

    '*** update the global variable with the new current folder path ***
    currentFolder = rootFolder & "\" & ListBox1.Value
    
    '*** activate folder buttons ****
    activateFolderButtons True
        
    '*** update the contents of the report sections listbox ***
    ListBox2.Clear
    
    Dim folderlst() As String
    folderlst = getFolderLst(currentFolder, True)
    
    '*** catch empty array and prevent loading into listbox ***
    If (Not folderlst) = -1 Then
        
        '*** do nothing ***
    Else
         '*** load folder list into listbox ***
         ListBox2.List = folderlst
    End If

End Sub

Private Sub initialiseApplication()

    '*** Initialises the application ***
    
    Dim folder As String

    '*** get root folder name or set to My Documents if one doesn't exist ***
    If DocVarExists("Root") = False Then
    
        folder = mydocs()
        ActiveDocument.Variables.Add Name:="Root", Value:=folder
    Else

        folder = ActiveDocument.Variables("Root").Value
    End If
    
    rootFolder = folder
    currentFolder = folder

    Label1.Caption = rootFolder

    '*** load listbox with names of sub folders in the root folder ***
    ListBox1.List = getFolderLst(folder, False)
        
    '*** set command button tips ***
    CommandButton19.ControlTipText = ReportSectionAddButtonTip
    CommandButton18.ControlTipText = ReportTemplateCopyButtonTip
    CommandButton17.ControlTipText = ReportTemplateMoveButtonTip
    CommandButton14.ControlTipText = ReportBuildButtonTip
    CommandButton13.ControlTipText = ReportClearButtonTip
    CommandButton10.ControlTipText = ReportSectionDemoteButtonTip
    CommandButton9.ControlTipText = ReportSectionPromoteButtonTip
    CommandButton7.ControlTipText = FolderCreateButtonTip
    CommandButton6.ControlTipText = FolderChangeNameButtonTip
    CommandButton5.ControlTipText = FolderDeleteButtonTip
    CommandButton3.ControlTipText = ReportSectionRemoveButtonTip
    CommandButton2.ControlTipText = ReportTemplateDeleteButtonTip

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

