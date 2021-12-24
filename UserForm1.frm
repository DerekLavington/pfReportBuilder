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

Dim TemplateArr() As String
Dim TemplateListArr() As Integer
Dim TemplateStyleArr() As String

Dim ModTemplateArr() As String
Dim ModTemplateListArr() As Integer
Dim ModTemplateStyleArr() As String
    
Dim ReportArr() As String
Dim ReportListArr() As Integer
Dim ReportStyleArr() As String

Dim ModTemplatePos As Integer
Dim ReportPos As Integer

Dim ParagraphCount As Integer
Dim DisplayMode As Integer

'TemplateArr              ModTemplateArr            ReportArr
'Holds template Commits   Holds Template Edits      Holds Report
'Not displayed            Displayed in full         Displayed partially

'Each new line in ModTemplateArr creates a blank line in ReportArr
'For a blank line to be shown in ReportArr, it must be selected from ModTemplateArr.
'To handle this, ReportArr will need a code char inserted where a blank line is to be displayed.
'Otherwise non-selected text will cause unnecessary blank lines

'Each deleted line in ModTemplateArr deletes a corresponding line in ReportArr whether or not blank
'Changing the order of lines in ModTemplateArr changes them in ReportArr

'TO DO :-
'Implememt ReportListArr and ReportStyleArr

Private Sub UserForm_Initialize()

    ParagraphCount = ActiveDocument.Paragraphs.Count
    
    ReDim TemplateArr(ParagraphCount)
    ReDim TemplateListArr(ParagraphCount)
    ReDim TemplateStyleArr(ParagraphCount)
    
    ReDim ModTemplateArr(ParagraphCount)
    ReDim ModTemplateListArr(ParagraphCount)
    ReDim ModTemplateStyleArr(ParagraphCount)
    
    ReDim ReportArr(ParagraphCount)
    ReDim ReportListArr(ParagraphCount)
    ReDim ReportStyleArr(ParagraphCount)
    
    '*** set label for active document ***
    Label2.Caption = ActiveDocument.Name
    
    '*** Populate display mode combox box and set mode***
    ComboBox1.AddItem "Report"
    ComboBox1.AddItem "Template"
    ComboBox1.ListIndex = 0
    
    DisplayMode = 1
    
    '*** Load word document into TemplateArr ***
    Dim x As Integer
    
    For x = 1 To ParagraphCount
    
        TemplateArr(x) = ActiveDocument.Paragraphs(x).Range.Text
        TemplateStyleArr(x) = ActiveDocument.Paragraphs(x).Style
        TemplateListArr(x) = Val(ActiveDocument.Paragraphs(x).Range.ListFormat.ListType)
    Next x
    
    '*** Copy TemplateArr to ModTemplateArr ***
    ModTemplateArr = TemplateArr
    ModTemplateStyleArr = TemplateStyleArr
    ModTemplateListArr = TemplateListArr
    
    '*** Display ModTemplateArr in Listbox1 ***
    LoadModTemplateListBox
    
    '*** Set Command Buttons ***
    CommandButton1.Enabled = False: ' Select
    CommandButton3.Enabled = False: ' Save Edit
    CommandButton6.Enabled = False: ' Deselect
    
End Sub

Private Sub ListBox1_Click()
        
    '*** Set the new ModTemplateArr position ***
    ModTemplatePos = ListBox1.ListIndex + 1
    
    '*** Extract the paragraph from ModTemplateArr and display in textbox for editing ***
    Dim para As String
    para = ModTemplateArr(ModTemplatePos)

    TextBox1.Value = para
    
    '*** Determine whether paragraph contains options and display in options Listbox ***
    Dim x As Integer
    x = InStr(para, "[")
    
    Select Case x:
    
        Case Is = 0, 1: 'No options
        
            CommandButton1.Enabled = True
            CommandButton3.Enabled = False
            
            ListBox2.Clear
            
        Case Is > 1: 'para contains options
            
            CommandButton1.Enabled = False
            CommandButton3.Enabled = False
            
            ListBox2.List = Split(GetOptionStr(para), "/")
    End Select

End Sub

Private Sub ListBox2_Click()
  
    CommandButton1.Enabled = True
    
End Sub

Private Sub ListBox3_Click()

    '*** Set the new ReportArr position ***
    ReportPos = ListBox3.ListIndex + 1

    '*** Set Command Buttons ***
    CommandButton6.Enabled = True: 'Deselect

End Sub

Private Sub ComboBox1_Change()

    DisplayMode = ComboBox1.ListIndex

    'change Listbox 3 display

End Sub

Private Sub TextBox1_Change()

    'Set Command Buttons ***
    CommandButton1.Enabled = False: 'Select
    CommandButton3.Enabled = True: ' Save Edit

End Sub

Private Sub CommandButton1_Click()

    '*** Select ***

    '*** remove brackets from selected string and load to the Report Array ***
    Select Case InStr(ModTemplateArr(ModTemplatePos), "[")
    
        Case Is = 0: ReportArr(ModTemplatePos) = TextBox1
        
        Case Is = 1: ReportArr(ModTemplatePos) = RemoveOptionBracesFromStr(ModTemplateArr(ModTemplatePos))
        
        Case Is > 1: ReportArr(ModTemplatePos) = InsertOptionIntoStr(ModTemplateArr(ModTemplatePos), ListBox2.Value)
    End Select
    
    '*** Load report listbox ***
    LoadReportListBox
    
    '*** Set command buttons ***
    CommandButton1.Enabled = False: 'Select
    CommandButton3.Enabled = False: 'Save Edit
    
End Sub

Private Sub CommandButton2_Click()

    '*** Select All ***

    Dim x As Integer
    Dim para As String
       
    For x = 1 To ParagraphCount
    
        para = ModTemplateArr(x)
        
        If InStr(para, "[") = 0 Then ReportArr(x) = para
    Next x

    LoadReportListBox

End Sub

Private Sub CommandButton3_Click()

    '*** Save Edit ***

    '*** Save edited para to ModTemplateArr ***
    ModTemplateArr(ModTemplatePos) = TextBox1.Value
    
    '*** Reload contents of ModTemplate and Report Listboxes ***
    LoadModTemplateListBox
    LoadReportListBox
        
    '*** Set command buttons ***
    CommandButton1.Enabled = True: ' Select
    CommandButton3.Enabled = False: 'Save Edit

End Sub

Private Sub CommandButton6_Click()

    '*** Deselect ***

    '*** Deselect paragraph in ReportArr ***
    ReportArr(ReportPos) = ""
    
    '*** Reload contents of Report Listbox ***
    LoadReportListBox

    '*** Set Command Buttons ***
    CommandButton6.Enabled = False: 'Deselect

End Sub

Private Sub CommandButton7_Click()

    '*** Deselect All ***
    
    '*** Clear ReportArr
    Dim x As Integer
    Dim para As String
       
    For x = 1 To ParagraphCount
    
        ModTemplateArr(x) = ""
    Next x
    
    '*** Reload contents of Report Listbox ***
    LoadReportListBox

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

Private Sub LoadModTemplateListBox()

    ListBox1.Clear
    
    Dim para As String
    Dim x As Integer
    
    For x = 1 To ParagraphCount
    
        'para = Left$(ModTemplateArr(x), Len(ModTemplateArr(x)) - 1): 'strip the carriage return from the listbox entry
        para = ModTemplateArr(x)
        
        Select Case ModTemplateListArr(x)
        
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
        para = ReportArr(x)
        
        If Len(para) > 0 Then
        
            Select Case ModTemplateListArr(x)
            
                Case Is = 4: ListBox3.AddItem "*    " & para: 'insert a psuedo-bullet to the listbox entry where the template has a bullet
                Case Else: ListBox3.AddItem para
            End Select
        
        End If
    Next x

End Sub
