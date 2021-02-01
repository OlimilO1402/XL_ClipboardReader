Attribute VB_Name = "MString"
Option Explicit

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



