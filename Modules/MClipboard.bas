Attribute VB_Name = "MClipboard"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpStr1 As Any, ByVal lpStr2 As Any) As Long
#Else
    Public Enum LongPtr
        [_]
    End Enum
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As Long
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As LongPtr
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare Function lstrcpy Lib "kernel32" (ByVal lpStr1 As Any, ByVal lpStr2 As Any) As Long
#End If

Private Const CF_TEXT       As Long = 1&
Private Const GMEM_MOVEABLE As Long = 2

'#If VBA7 Then
Public Sub SetText(strText As String)
    Dim lngIdentifier As LongPtr, lngPointer As LongPtr
    lngIdentifier = GlobalAlloc(GMEM_MOVEABLE, Len(strText) + 1)
    lngPointer = GlobalLock(lngIdentifier)
    Call lstrcpy(ByVal lngPointer, strText)
    Call GlobalUnlock(lngIdentifier)
    Call OpenClipboard(0&)
    Call EmptyClipboard
    Call SetClipboardData(CF_TEXT, lngIdentifier)
    Call CloseClipboard
    Call GlobalFree(lngIdentifier)
End Sub
'#Else
'    Public Sub SetText(strText As String)
'        Dim lngIdentifier As LongPtr, lngPointer As LongPtr
'        lngIdentifier = GlobalAlloc(GMEM_MOVEABLE, Len(strText) + 1)
'        lngPointer = GlobalLock(lngIdentifier)
'        Call lstrcpy(ByVal lngPointer, strText)
'        Call GlobalUnlock(lngIdentifier)
'        Call OpenClipboard(0&)
'        Call EmptyClipboard
'        Call SetClipboardData(CF_TEXT, lngIdentifier)
'        Call CloseClipboard
'        Call GlobalFree(lngIdentifier)
'    End Sub
'#End If

Public Sub CBSetText(aText As String)
    Dim ClipB As Object: Set ClipB = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}") 'msforms20 DataObject
    ClipB.SetText aText
    ClipB.PutInClipboard
    Set ClipB = Nothing
End Sub

Public Function CBGetText() As String
    Dim ClipB As Object: Set ClipB = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}") 'msforms20 DataObject
    ClipB.GetFromClipboard
    CBGetText = ClipB.GetText
    Set ClipB = Nothing
End Function

#If VBA Then
Function ClipBoard_GetText() As String
Try: On Error GoTo Catch
    Dim docb As New DataObject
    docb.GetFromClipboard
    ClipBoard_GetText = docb.GetText
Catch:
End Function
#End If

#If VBA Then
Sub ClipBorad_SetText(ByVal aText As String)
Try: On Error GoTo Catch
    Dim docb As DataObject
    Set docb = New DataObject
    'docb.Clear
#If Win64 Then
    'MsgBox "x64"
    'aText = StrConv(aText, vbWide)
    docb.SetText aText, 1
    docb.PutInClipboard
#Else
    'MsgBox "x86"
    docb.SetText aText, 1
    docb.PutInClipboard
#End If
Catch:
End Sub
#End If

