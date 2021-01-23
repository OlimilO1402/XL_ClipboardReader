VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ClipboardReader"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13935
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   13935
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox TxtData 
      Height          =   6495
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   4
      Top             =   720
      Width           =   13335
   End
   Begin VB.CheckBox cbNewlineOnly 
      Caption         =   "Only NewLine"
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      ToolTipText     =   "Every data will be in one column"
      Top             =   360
      Width           =   1335
   End
   Begin VB.CheckBox cbNumOnly 
      Caption         =   "Only numbers no text"
      Height          =   195
      Left            =   1560
      TabIndex        =   1
      Top             =   80
      Width           =   1935
   End
   Begin VB.CommandButton BtnRead 
      Caption         =   "Read"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   80
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   $"Form1.frx":1782
      Height          =   750
      Left            =   3720
      TabIndex        =   3
      Top             =   75
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

Private Sub Form_Resize()
    Dim L As Single, t As Single, W As Single, H As Single
    t = TxtData.Top: W = Me.ScaleWidth: H = Me.ScaleHeight - t
    If W > 0 And H > 0 Then
        TxtData.Move L, t, W, H
        Label1.Move Label1.Left, Label1.Top, W - Label1.Left
    End If
End Sub

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

Function PadLeft(this As String, _
                 ByVal totalWidth As Long, _
                 Optional ByVal paddingChar As String) As String
    If LenB(paddingChar) Then
        If Len(this) < totalWidth Then
            PadLeft = String$(totalWidth, paddingChar)
            MidB$(PadLeft, totalWidth * 2 - LenB(this) + 1) = this
        Else
            PadLeft = this
        End If
    Else
        PadLeft = Space$(totalWidth)
        RSet PadLeft = this
    End If
End Function
Function PadRight(this As String, _
                  ByVal totalWidth As Long, _
                  Optional ByVal paddingChar As String) As String
    If LenB(paddingChar) Then
        If Len(this) < totalWidth Then
            PadRight = String$(totalWidth, paddingChar)
            MidB$(PadRight, 1) = this
        Else
            PadRight = this
        End If
    Else
        PadRight = Space$(totalWidth)
        LSet PadRight = this
    End If
End Function

