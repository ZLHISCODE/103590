VERSION 5.00
Begin VB.Form frmVideoDockWindow 
   Caption         =   "��Ƶ�ɼ�"
   ClientHeight    =   9045
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   10980
   Icon            =   "frmVideoDockWindow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9045
   ScaleWidth      =   10980
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "frmVideoDockWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    '����д�����ʱ�򣬸òɼ�������Ҫ��ʾ����ǰ��
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3 '�������ö�
    
    '�ָ�����״̬
    Call RestoreWinState(Me, App.ProductName)
    
    Call frmWork_Video.ShowVideoWindow(Me)
End Sub

Private Sub Form_Resize()
On Error Resume Next
    '������ڽ�����С��ʱ����������Ƶ���ֵ���
    If Me.WindowState = 1 Then Exit Sub
    
    Call frmWork_Video.UpdateSize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmWork_Video.RestoreContainer
    Call SaveWinState(Me, App.ProductName)
End Sub
