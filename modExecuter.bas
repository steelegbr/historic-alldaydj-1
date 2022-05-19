Attribute VB_Name = "modExecuter"
' Shell execute for VB6 - crazy huh?

Private Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As _
String, ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function Shell(Program As String, Optional ShowCmd As Long = _
vbNormalNoFocus, Optional ByVal WorkDir As Variant) As Long

    Dim FirstSpace As Integer, Slash As Integer

    If Left(Program, 1) = """" Then
        FirstSpace = InStr(2, Program, """")


        If FirstSpace <> 0 Then
            Program = Mid(Program, 2, FirstSpace - 2) & _
              Mid(Program, FirstSpace + 1)
            FirstSpace = FirstSpace - 1
        End If

    Else
        FirstSpace = InStr(Program, " ")
    End If

    If FirstSpace = 0 Then FirstSpace = Len(Program) + 1

    If IsMissing(WorkDir) Then

        For Slash = FirstSpace - 1 To 1 Step -1
            If Mid(Program, Slash, 1) = "\" Then Exit For
        Next

        If Slash = 0 Then
            WorkDir = CurDir
        ElseIf Slash = 1 Or Mid(Program, Slash - 1, 1) = ":" Then
            WorkDir = Left(Program, Slash)
        Else
            WorkDir = Left(Program, Slash - 1)
        End If

    End If

    Shell = ShellExecute(0, vbNullString, _
    Left(Program, FirstSpace - 1), LTrim(Mid(Program, _
    FirstSpace)), WorkDir, ShowCmd)
    If Shell < 32 Then VBA.Shell Program, ShowCmd 'To raise Error

End Function


