VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mfrmAdviceMain As Object    '������

Public Function zlCloseOrderForm()
    On Error Resume Next
    '������ڲ����ڣ���ֱ�ӷ��سɹ�
'    If mfrmMain Is Nothing Then
'        zlCloseOrderForm = True
'        Exit Function
'    End If
    
    '�رմ���
    mfrmAdviceMain.zlCloseMe
    Set mfrmAdviceMain = Nothing
    zlCloseOrderForm = True
End Function

Public Function zlOpenOrderForm()
    If mfrmAdviceMain Is Nothing Then
        Set mfrmAdviceMain = New frmAdviceMain
        mfrmAdviceMain.Height = 0: mfrmAdviceMain.Width = 0
        mfrmAdviceMain.Show
    End If
End Function
