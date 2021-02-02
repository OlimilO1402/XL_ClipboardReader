VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ClipboardReader"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13935
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   481
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   929
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox TxtData 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   4
      Top             =   600
      Width           =   13335
   End
   Begin VB.CheckBox cbNewlineOnly 
      Caption         =   "Only NewLine"
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      ToolTipText     =   "Every value gets a new line, so all data in one column"
      Top             =   360
      Width           =   1335
   End
   Begin VB.CheckBox cbNumOnly 
      Caption         =   "Only numbers no text"
      Height          =   195
      Left            =   1560
      TabIndex        =   1
      ToolTipText     =   "Alle texts will be deleted, only numeric-data is allowed"
      Top             =   80
      Width           =   1935
   End
   Begin VB.CommandButton BtnRead 
      Caption         =   "Read"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   $"Form1.frx":3AFA
      Height          =   390
      Left            =   3720
      TabIndex        =   3
      Top             =   75
      UseMnemonic     =   0   'False
      Width           =   10095
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    'Label1.Caption = "Maybe you know the situation, you have a pdf-document with some numeric values in tables and you need it in your excel-calculation. Just copy the data to the clipboard, click the Read button, and now you can paste it to your excel sheet. Whitespaces will be replaced by one Tab."
    Label1.Caption = "Sie kennen das sicher, Sie haben eine pdf-Datei mit irgendwelchen Daten in irgendwelchen Tabellen, und sie wollen diese Daten in Excel einlesen. Einfach die Daten in der Tabelle ihrer pdf-Datei markieren, in die Zwischenablage kopieren ([Strg]+[C]) und den Schalter [Read] klicken. Jetzt können Sie die Daten in Excel einfügen ([Strg]+[V])."
End Sub
Private Sub Form_Initialize()
    Label1.Caption = "Maybe you know the situation, you have a pdf-document with some numeric values in tables and you need it in your excel-calculation. Just copy the data to the clipboard, click the Read button, and now you can paste it to your excel sheet. Whitespaces will be replaced by one Tab."
    'Label1.Caption = "Sie kennen das sicher, Sie haben eine pdf-Datei mit irgendwelchen Daten in irgendwelchen Tabellen, und sie wollen diese Daten in Excel einlesen. Einfach die Daten in der Tabelle ihrer pdf-Datei markieren, in die Zwischenablage kopieren ([Strg]+[C]) und den Schalter [Read] klicken. Jetzt können Sie die Daten in Excel einfügen ([Strg]+[V])."
End Sub

#If VB Then
Private Sub Form_Resize()
    Dim brdr As Single: brdr = 8 '* Screen.TwipsPerPixelX
    Dim L As Single: L = Label1.Left
    Dim t As Single: t = Label1.Top
    Dim W As Single: W = Me.ScaleWidth - L - brdr
    Dim H As Single: H = Me.ScaleHeight
    Dim b As Boolean
    If W > 0 And H > 0 Then
        Label1.Move L, t, W
        Label1.AutoSize = True
    Else
        b = True
    End If
    L = 0
    t = Max(IIf(b, 0, Label1.Top + Label1.Height), BtnRead.Top + BtnRead.Height)
    W = Me.ScaleWidth
    H = H - t
    If W > 0 And H > 0 Then
        TxtData.Move 0, t, W, H
    End If
End Sub
#Else
Private Sub UserForm_Resize()
    Dim brdr As Single: brdr = 8 '* Screen.TwipsPerPixelX
    Dim L As Single: L = Label1.Left
    Dim t As Single: t = Label1.Top
    Dim W As Single: W = Me.ScaleWidth - L - brdr
    Dim H As Single: H = Me.ScaleHeight
    Dim b As Boolean
    If W > 0 And H > 0 Then
        Label1.Move L, t, W
        Label1.AutoSize = True
    Else
        b = True
    End If
    L = 0
    t = Max(IIf(b, 0, Label1.Top + Label1.Height), BtnRead.Top + BtnRead.Height)
    W = Me.ScaleWidth
    H = H - t
    If W > 0 And H > 0 Then
        TxtData.Move 0, t, W, H
    End If
End Sub
#End If

Function Max(V1, V2)
    If V1 > V2 Then Max = V1 Else Max = V2
End Function
Function Min(V1, V2)
    If V1 < V2 Then Min = V1 Else Min = V2
End Function

Private Sub BtnRead_Click()
    'also OK, folgende Idee:
    'damit man hier in der TextBox noch verbessern kann
    'alle Tabs erstmal durch "|" darstellen
    'dann kann der User die falschen Spalten rauslöschen,
    'damit die Spalten besser sichtbar werden
    'alle Spalten in den Zeilen gleichlang machen mit Padright in Spaces
    'dazu zuerst die Länge jeder einzelnen Spalte rausfinden
    'nachher werden alle Spaltentrenner "|" durch Tab ersetzt die Spaces mit Trim gelöscht
    'und die Zeilen zusammengefügt
    '
    If TxtData.Text = vbNullString Then
        TxtData.Text = Clipboard.GetText
    End If
    Dim t As String: t = TxtData.Text
    'zuerst alle tabs nochmal in spaces umwandeln, falls ein Text 2-mal eingelesen werden muss
    t = Replace(t, vbTab, " ")
    'so und wo zum Henker ist jetzt die recursive Delete Ws-Funktion abgeblieben?
    Dim lines() As String: lines = Split(t, vbCrLf)
    Dim i As Long
    Dim onlyNewLine As Boolean: onlyNewLine = Me.cbNewlineOnly.Value
    Dim svbCrLf As String: If onlyNewLine Then svbCrLf = vbCrLf
    'jeden Wert in eine neue Zeile
    For i = LBound(lines) To UBound(lines)
        Dim line As String
        'alle mehrfachen Whitespaces enfernen
        line = DeleteMultiWS(lines(i))
        'für Excel: alle Zahlen mit Komma(",") statt Punkt(".")
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
    t = Join(lines, vbCrLf)
    TxtData.Text = t
    Clipboard.Clear
    Clipboard.SetText t
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    TxtData_KeyDown KeyCode, Shift
End Sub

Private Sub TxtData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KeyCodeConstants.vbKeyA And Shift = ShiftConstants.vbCtrlMask Then
        TxtData.SelStart = 0
        TxtData.SelLength = Len(TxtData.Text)
    End If
End Sub
