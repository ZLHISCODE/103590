Attribute VB_Name = "mdlBroadcastMsg"
Option Explicit

Public Const WM_REFRESH_IMAGE As Long = 3173

Private mobjList As New Scripting.Dictionary


'ע��㲥���ܶ���
Public Sub RegBroadcastRec(ByVal lngHwnd As Long)
    If mobjList.Exists(lngHwnd) Then Exit Sub
    
    mobjList.Add lngHwnd, lngHwnd
End Sub

'ж�ع㲥���ܶ���
Public Sub UnBroadcastRec(ByVal lngHwnd As Long)
    If mobjList.Exists(lngHwnd) = False Then Exit Sub
    
    Call mobjList.Remove(lngHwnd)
End Sub

'�㲥��Ϣ
Public Sub BoradcastMsg(ByVal lngMsgData As Long)
    Dim i As Long
    
    For i = 0 To mobjList.Count - 1
        PostMessage mobjList.Keys(i), WM_REFRESH_IMAGE, lngMsgData, 0
'        DoEvents
'        SendMessage Val(mobjList.Keys(i)), WM_REFRESH_IMAGE, lngMsgData, 0
    Next i
    
    DoEvents
End Sub




