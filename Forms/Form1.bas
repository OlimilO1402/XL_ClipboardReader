Option Explicit

Private Sub BtnRead_Click()
    If TextBox1.Text = vbNullString Then
        TextBox1.Text = ClipBoard_GetText
    End If
    Dim t As String: t = TextBox1.Text
    'zuerst alle tabs nochmal in spaces umwandeln, falls ein Text 2-mal eingelesen werden muss
    t = Replace(t, vbTab, " ")
    Dim lines() As String: lines = Split(t, vbCrLf)
    Dim i As Long
    Dim onlyNewLine As Boolean: onlyNewLine = Me.cbNewlineOnly.Value
    Dim svbCrLf As String: If onlyNewLine Then svbCrLf = vbCrLf
    'jeden Wert in eine neue Zeile
    For i = LBound(lines) To UBound(lines)
        Dim line As String
        line = lines(i)
        line = Replace(line, ".", ",")
        If cbNumOnly.Value Then
            Dim sa() As String: sa = Split(line, " ")
            Dim j As Long, u As Long: u = UBound(sa)
            line = ""
            For j = 0 To u
                If IsNumeric(sa(j)) Then
                    line = line & sa(j) & svbCrLf
                    If onlyNewLine Then
                        'line = line & vbNewLine
                    Else
                        If j < u Then
                            line = line & vbTab '" "
                        End If
                    End If
                End If
            Next
        Else
            If onlyNewLine Then
                line = Replace(line, " ", vbCrLf)
            Else
                line = Replace(line, " ", vbTab)
            End If
        End If
        lines(i) = line
    Next
    TextBox1.Text = Join(lines, vbCrLf)
    ClipBorad_SetText TextBox1.Text
End Sub

Private Sub UserForm_Initialize()
    Label1.Caption = "Sie kennen das sicher, Sie haben eine pdf-Datei mit irgendwelchen Daten in irgendwelchen Tabellen, und sie wollen diese Daten in Excel einlesen. Einfach die Daten in der Tabelle ihrer pdf-Datei markieren, in die Zwischenablage kopieren ([Strg]+[C]) und den Schalter [Read] klicken. Jetzt können Sie die Daten in Excel einfügen ([Strg]+[V])."
End Sub

Private Sub UserForm_Resize()
    Dim Brdr As Single: Brdr = 8 '* screen.twipsperpixelx
    Dim L As Single: L = Brdr
    Dim t As Single: t = TextBox1.Top
    Dim W As Single: W = Me.Width - L - Brdr
    Dim H As Single: H = Me.Height - t - brd
    If W > 0 And H > 0 Then
        TextBox1.Move L, t, W, H
    End If
    
End Sub

Function ClipBoard_GetText() As String
Try: On Error GoTo Catch
    Dim docb As New DataObject
    docb.GetFromClipboard
    ClipBoard_GetText = docb.GetText
Catch:
End Function

Sub ClipBorad_SetText(ByVal aText As String)
Try: On Error GoTo Catch
    Dim docb As New DataObject
    'docb.Clear
#If Win64 Then
    'MsgBox "x64"
    'aText = StrConv(aText, vbWide)
    'so ein verfluchter Dreck, warum geht das DataObject jetzt nicht mehr
    docb.SetText aText, 1
    docb.PutInClipboard
#Else
    'MsgBox "x86"
    docb.SetText aText, 1
    docb.PutInClipboard
#End If
Catch:
End Sub

'    Dim objClipBoard As Object
'    Set objClipBoard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
'    Call objClipBoard.SetText("Hallo Welt")
'    Call objClipBoard.PutInClipboard
'    Set objClipBoard = Nothing


'Replace Recursive Delete Multi WhiteSpace WS
Public Function DeleteMultiWS(s As String) As String
    DeleteMultiWS = Trim$(s)
    If InStr(1, s, "  ") = 0 Then Exit Function
    DeleteMultiWS = Replace(s, "  ", " ")
    DeleteMultiWS = DeleteMultiWS(DeleteMultiWS)
End Function
Public Function DeleteCRLF(s As String) As String
    DeleteCRLF = Trim$(s)
    If InStr(1, s, vbLf) = 0 Then Exit Function
    If InStr(1, s, vbCr) = 0 Then Exit Function
    DeleteCRLF = Replace(Replace(Replace(s, vbCrLf, " "), vbLf, " "), vbCr, " ")
    DeleteCRLF = DeleteCRLF(DeleteCRLF)
End Function
