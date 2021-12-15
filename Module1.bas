Attribute VB_Name = "Module1"
Option Explicit

Sub ShowMyForm()

    UserForm1.Show
    
    Unload UserForm1

End Sub

Sub GetCode()

    Dim prj As VBProject
    Dim comp As VBComponent
    Dim code As CodeModule
    Dim composedFile As String
    Dim i As Integer

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Set oFile = fso.CreateTextFile("C:\Users\Derek Lavington\Documents\output.txt")
   



    Set prj = ThisDocument.VBProject
        For Each comp In prj.VBComponents
            Set code = comp.CodeModule

            composedFile = comp.Name & vbNewLine

            For i = 1 To code.CountOfLines
                composedFile = composedFile & code.Lines(i, 1) & vbNewLine
            Next
            
            oFile.WriteLine composedFile
            
        Next
    oFile.Close
Set fso = Nothing
Set oFile = Nothing

End Sub
