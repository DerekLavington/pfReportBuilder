VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   7776
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   14664
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type wdPara

    paratext As String
    listCode As Integer
    StyleCode As Page
End Type

Enum ButtonMode
 
    DisableSelect = 1
    EnableSelect = 2
    DisableLineSelect = 3
    EnablelineSelect = 4
End Enum

Dim TemplateArr() As wdPara
Dim ReportArr() As wdPara

Dim TemplatePos As Integer
Dim ReportPos As Integer

Dim ParagraphCount As Integer

'Dim DisplayMode As Integer

Private Sub UserForm_Initialize()
    
    LoadTemplate
    
    '*** set label for active document ***
    Label2.Caption = ActiveDocument.Name
    
    '*** Populate display mode combox box and set mode***
    ComboBox1.AddItem "Report"
    ComboBox1.AddItem "Template"
    ComboBox1.ListIndex = 0
    
    'DisplayMode = 1
     
    SetButtonMode DisableSelect
        
End Sub

Private Sub ListBox1_Click()
        
    '*** Update the Template position ***
    TemplatePos = ListBox1.ListIndex + 1
    
    '*** Extract the paragraph from Template and display for editing ***
    Dim para As String
    para = TemplateArr(TemplatePos).paratext

    TextBox1.Value = para
    
    '*** Determine whether paragraph contains options and display in options Listbox ***
    Dim x As Integer
    x = InStr(para, "[")
    
    Select Case x:
    
        Case Is = 0, 1: 'No selection options in para
        
            SetButtonMode EnableSelect
            
        Case Is > 1: 'para contains selection options
            
            ListBox2.List = Split(GetOptionStr(para), "/"): 'extract and display the para selection options
            
            SetButtonMode EnableSelect
            SetButtonMode DisableLineSelect
    
    End Select

End Sub

Private Sub ListBox2_Click()
  
    '*** Select para selection option ***
    
    SetButtonMode EnablelineSelect
        
End Sub

Private Sub ListBox3_Click()

    '*** Update the Report position ***
    ReportPos = ListBox3.ListIndex + 1

    '*** Set Command Buttons ***
    CommandButton6.Enabled = True: 'Deselect

End Sub

Private Sub ComboBox1_Change()

    'DisplayMode = ComboBox1.ListIndex

End Sub

Private Sub TextBox1_Change()

    '*** Alert To Edit Of Template Para ***
        
    SetButtonMode EnableSelect
    SetButtonMode DisableLineSelect

End Sub

Private Sub CommandButton1_Click()

    '*** Select Line From Template ListBox And Copy to Report Listbox ***

    '*** remove brackets from selected string and load to the Report Array ***
    Select Case InStr(TemplateArr(TemplatePos).paratext, "[")
    
        Case Is = 0:
            ReportArr(TemplatePos).paratext = TextBox1
        
        Case Is = 1:
            ReportArr(TemplatePos).paratext = RemoveOptionBracesFromStr(TemplateArr(TemplatePos).paratext)
        
        Case Is > 1:
            ReportArr(TemplatePos).paratext = InsertOptionIntoStr(TemplateArr(TemplatePos).paratext, ListBox2.Value)
    End Select
    
    '*** Load report listbox ***
    LoadReportListBox
    
    '*** Clear select options listbox ***
    ListBox2.Clear
    
    SetButtonMode DisableLineSelect

End Sub

Private Sub CommandButton2_Click()

    '*** Select All Non-Optional Paras From Template And Copy To Report Listbox ***

    Dim x As Integer
       
    For x = 1 To ParagraphCount
        
        If InStr(TemplateArr(x).paratext, "[") = 0 Then ReportArr(x) = TemplateArr(x)
    Next x

    LoadReportListBox

End Sub

Private Sub CommandButton3_Click()

    '*** Save Para Edit To Template ***

    '*** Save edited para to Template ***
    TemplateArr(TemplatePos).paratext = TextBox1.Value
    
    '#### - if bullet removed,  also need to change list TemplateArr.listcode
        
    '*** Reload contents of Template and Report Listboxes ***
    LoadTemplateListBox
    LoadReportListBox
    
    SetButtonMode DisableLineSelect
    
End Sub

Private Sub CommandButton4_Click()

    '*** Insert Line to Template And Report ***

    Dim tmpTemplateArr() As wdPara
    Dim tmpReportArr() As wdPara
    
    ReDim tmpTemplateArr(ParagraphCount)
    ReDim tmpReportArr(ParagraphCount)
    
    tmpTemplateArr = TemplateArr
    tmpReportArr = ReportArr
    
    ReDim TemplateArr(ParagraphCount + 1)
    ReDim ReportArr(ParagraphCount + 1)
 
    Dim x As Integer
    
    For x = 1 To TemplatePos - 1
        
        TemplateArr(x) = tmpTemplateArr(x)
        ReportArr(x) = tmpReportArr(x)
    Next x
    
    For x = TemplatePos To ParagraphCount - 1
    
        TemplateArr(x + 1) = tmpTemplateArr(x)
        ReportArr(x + 1) = tmpReportArr(x)
    Next x

    TemplateArr(TemplatePos).paratext = Chr$(13)
    TemplateArr(TemplatePos).listCode = 0
    TemplateArr(TemplatePos).StyleCode = 0

    ReportArr(TemplatePos).paratext = Chr$(13)
    ReportArr(TemplatePos).listCode = 0
    ReportArr(TemplatePos).StyleCode = 0

    LoadTemplateListBox
    LoadReportListBox
        
    ParagraphCount = ParagraphCount + 1
 
    SetButtonMode DisableSelect
 
End Sub

Private Sub CommandButton5_Click()

    '*** Delete Line From Template And Report ***
    
    Dim tmpTemplateArr() As wdPara
    Dim tmpReportArr() As wdPara
    
    ReDim tmpTemplateArr(ParagraphCount)
    ReDim tmpReportArr(ParagraphCount)
    
    tmpTemplateArr = TemplateArr
    tmpReportArr = ReportArr
    
    ReDim TemplateArr(ParagraphCount - 1)
    ReDim ReportArr(ParagraphCount - 1)
    Dim tmpArr() As wdPara
    
    Dim x As Integer
    
    For x = 1 To TemplatePos - 1
        
        TemplateArr(x) = tmpTemplateArr(x)
        ReportArr(x) = tmpReportArr(x)
    Next x

    For x = TemplatePos To ParagraphCount - 1
    
        TemplateArr(x) = tmpArr(x + 1)
        ReportArr(x) = tmpReportArr(x + 1)
    Next x

    LoadTemplateListBox
    LoadReportListBox
        
    ParagraphCount = ParagraphCount - 1
 
    SetButtonMode DisableSelect
    
End Sub

Private Sub CommandButton6_Click()

    '*** Deselect Line From Report Listbox ***

    '*** Deselect paragraph in Report ***
    ReportArr(ReportPos).paratext = ""
    ReportArr(ReportPos).listCode = 0
    ReportArr(ReportPos).StyleCode = 0
    
    '*** Reload contents of Report Listbox ***
    LoadReportListBox

    '*** Set Command Buttons ***
    CommandButton6.Enabled = False: 'Deselect

End Sub

Private Sub CommandButton7_Click()

    '*** Deselect All ***
        
    '*** Clear ReportArr
    Erase ReportArr
    ReportPos = 0
    
    '*** Reload contents of Report Listbox ***
    LoadReportListBox

End Sub

Private Sub CommandButton8_Click()

    '*** Promote Line ***
    
    Dim tmpTemplateArr As wdPara
    Dim tmpReportArr As wdPara
 
    tmpTemplateArr = TemplateArr(TemplatePos - 1)
    tmpReportArr = ReportArr(TemplatePos - 1)
    
    TemplateArr(TemplatePos - 1) = TemplateArr(TemplatePos)
    ReportArr(TemplatePos - 1) = ReportArr(TemplatePos)
    
    TemplateArr(TemplatePos) = tmpTemplateArr
    ReportArr(TemplatePos) = tmpReportArr

    LoadTemplateListBox
    LoadReportListBox
 
    SetButtonMode DisableSelect

End Sub

Private Sub CommandButton11_Click()
    
    '*** Revert to unmodified template ***
    
    LoadTemplate

    SetButtonMode DisableSelect

End Sub

Private Sub CommandButton9_Click()

    '*** Demote Line ***
    
    Dim tmpTemplateArr As wdPara
    Dim tmpReportArr As wdPara
 
    tmpTemplateArr = TemplateArr(TemplatePos + 1)
    tmpReportArr = ReportArr(TemplatePos + 1)
    
    TemplateArr(TemplatePos + 1) = TemplateArr(TemplatePos)
    ReportArr(TemplatePos + 1) = ReportArr(TemplatePos)
    
    TemplateArr(TemplatePos) = tmpTemplateArr
    ReportArr(TemplatePos) = tmpReportArr

    LoadTemplateListBox
    LoadReportListBox
    
    SetButtonMode DisableSelect

End Sub

Private Sub CommandButton14_Click()

    '*** Exit Program ***
    
    End

End Sub

Private Function GetOptionStr(para As String) As String

    Dim x As Integer
    Dim y As Integer

    x = InStr(para, "[")
    y = InStr(para, "]")
            
    GetOptionStr = Mid$(para, x + 1, y - x - 1)

End Function

Private Function RemoveOptionBracesFromStr(str As String) As String

    Dim x As Integer
    Dim y As Integer
    Dim StrLen As Integer
    
    x = InStr(str, "[")
    y = InStr(str, "]")
    StrLen = Len(str)
    
    RemoveOptionBracesFromStr = Mid$(str, 2, y - 1) & Right$(str, StrLen - y + 1)
    
End Function

Private Function InsertOptionIntoStr(str As String, optionStr As String) As String

    Dim x As Integer
    Dim y As Integer
    Dim StrLen As Integer
       
    x = InStr(str, "[")
    y = InStr(str, "]")
    StrLen = Len(str)
    
    InsertOptionIntoStr = Left$(str, x - 1) & optionStr & Right$(str, StrLen - y + 1)
    
End Function

Private Sub LoadTemplateListBox()

    ListBox1.Clear
    
    Dim para As String
    Dim x As Integer
    
    For x = 1 To ParagraphCount
    
        'para = Left$(TemplateArr(x), Len(TemplateArr(x)) - 1): 'strip the carriage return from the listbox entry
        para = TemplateArr(x).paratext
        
        Select Case TemplateArr(x).listCode
        
            Case Is = 4: ListBox1.AddItem "*    " & para: 'insert a psuedo-bullet to the listbox entry where the template has a bullet
            Case Else: ListBox1.AddItem para
        End Select
    Next x

End Sub

Private Sub LoadReportListBox()

    ListBox3.Clear
 
    Dim para As String
    Dim x As Integer
    
    For x = 1 To ParagraphCount
    
        'para = Left$(ReportArr(x), Len(ReportArr(x)) - 1): 'strip the carriage return from the listbox entry
        para = ReportArr(x).paratext
        
        If Len(para) > 0 Then
        
            Select Case ReportArr(x).listCode
            
                Case Is = 4: ListBox3.AddItem "*    " & para: 'insert a psuedo-bullet to the listbox entry where the template has a bullet
                Case Else: ListBox3.AddItem para
            End Select
        
        End If
    Next x

End Sub

Private Sub LoadTemplate()
  
    ParagraphCount = ActiveDocument.Paragraphs.Count
    
    ReDim TemplateArr(ParagraphCount)
    ReDim ReportArr(ParagraphCount)
  
    Dim x As Integer
    
    For x = 1 To ParagraphCount
    
        TemplateArr(x).paratext = ActiveDocument.Paragraphs(x).Range.Text
        TemplateArr(x).StyleCode = ActiveDocument.Paragraphs(x).Style
        TemplateArr(x).listCode = Val(ActiveDocument.Paragraphs(x).Range.ListFormat.ListType)
    Next x
    
    TemplatePos = 0
    ReportPos = 0
    
    LoadTemplateListBox
    LoadReportListBox

End Sub

Private Sub SetButtonMode(mode As Integer)
    
    Select Case mode
    
        Case Is = 1: 'Disable all selection dependent buttons
        
            CommandButton1.Enabled = False: ' Select Line
            CommandButton3.Enabled = False: ' Save Edit
            CommandButton4.Enabled = False: ' Insert Line
            CommandButton5.Enabled = False: ' Delete Line
            CommandButton6.Enabled = False: ' Deselect Line
            CommandButton8.Enabled = False: ' Promote Line
            CommandButton9.Enabled = False: ' Demote Line

        Case Is = 2: 'Enable all selection dependent buttons
        
            CommandButton1.Enabled = True: ' Select Line
            CommandButton3.Enabled = True: ' Save Edit
            CommandButton4.Enabled = True: ' Insert Line
            CommandButton5.Enabled = True: ' Delete Line
            CommandButton6.Enabled = True: ' Deselect Line
            CommandButton8.Enabled = True: ' Promote Line
            CommandButton9.Enabled = True: ' Demote Line

        Case Is = 3: 'Disable line selection button

            CommandButton1.Enabled = False: ' Select Line
        
        Case Is = 4: 'Enable line selection button
        
            CommandButton1.Enabled = True: ' Select Line
        
    End Select
    
End Sub

