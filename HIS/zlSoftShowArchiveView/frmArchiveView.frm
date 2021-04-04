VERSION 5.00
Begin VB.Form frmArchiveView 
   AutoRedraw      =   -1  'True
   Caption         =   "�������Ӳ�������"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   13545
   Icon            =   "frmArchiveView.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   13545
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "frmArchiveView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mclsArchive As clsArchive
Private mlngSubFormHwnd As Long

Public Sub ShowMe(lngPatientID As Long, lngClinicID As Long, blnShow As Boolean)

On Error GoTo ErrorHand
    If mclsArchive Is Nothing Then Set mclsArchive = New clsArchive
    
    Call mclsArchive.zlInitCommon

    mlngSubFormHwnd = mclsArchive.zlGetFormHwnd
    
    Call ShowSubWindow(mlngSubFormHwnd, Me.hWnd)
    Call SetWindowStyle(mlngSubFormHwnd)
    Call UpdateSize(mlngSubFormHwnd, Me.hWnd)
    Call zlRefresh(lngPatientID, lngClinicID)
    If blnShow Then Call Me.Show
    
    Me.Caption = "�������Ӳ�������"
    Exit Sub
ErrorHand:
    If errHandle("zlSoftShowHisForms.frmArchiveView.ShowMe", "��ʾ�������Ĵ��ڳ��ִ���") = 1 Then Resume
End Sub

Private Sub SetWindowStyle(ByVal lngHandle As Long)
    Dim lngWindowStyle As Long
On Error Resume Next
    lngWindowStyle = GetWindowLong(lngHandle, GWL_STYLE)
    
    lngWindowStyle = lngWindowStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)

    Call SetWindowLong(lngHandle, GWL_STYLE, lngWindowStyle Or WS_CHILD)
End Sub

Public Sub UpdateSize(ByVal lngHwnd As Long, Optional ByVal lngMainHwnd As Long)
'����Ƕ�뱨��Ĵ��ڴ�С
    Dim vRect As RECT
On Error Resume Next
    If lngHwnd <= 0 Then Exit Sub
    
    If lngMainHwnd <> 0 Then
        SetParent lngHwnd, lngMainHwnd
    Else
        SetParent lngHwnd, 0
    End If
    
    GetWindowRect lngMainHwnd, vRect
    MoveWindow lngHwnd, 0, 0, Abs(vRect.Right - vRect.Left - 15), Abs(vRect.Bottom - vRect.Top - 40), 1
End Sub

'******************************************************************************************************************
'���ܣ� ˢ��Ƕ��ʽ�������Ĵ���
'������ lngPatientID - ����ID
'       lngClinicID - ����ID������Ϊ�Һ�ID ���˹Һż�¼.ID��סԺΪ��ҳID
'���أ� 0 �ɹ�����0��ʧ��
'˵����
'******************************************************************************************************************
Public Function zlRefresh(lngPatientID As Long, lngClinicID As Long) As Long
    
On Error GoTo ErrorHand
    
    mclsArchive.zlRefresh lngPatientID, lngClinicID
    
    Me.Caption = "�������Ӳ�������"
    Exit Function
ErrorHand:
    If errHandle("zlSoftShowHisForms.frmArchiveView.zlRefresh", "ˢ�²������Ĵ��ڳ��ִ��󣬲���ID=" & lngPatientID & "������ID=" & lngClinicID) = 1 Then Resume
End Function


Private Sub Form_Load()
    On Error Resume Next
        
    Call gzlComLib.RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.WindowState = vbMinimized
        Me.Hide
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call UpdateSize(mlngSubFormHwnd, Me.hWnd)
End Sub

'------------------------------------------------
'���ܣ��رյ��Ӳ������Ĵ���
'������ ��
'���أ�True -- �ɹ��� False -- ʧ��
'------------------------------------------------
Public Function zlCloseMe()
    Unload Me
End Function
