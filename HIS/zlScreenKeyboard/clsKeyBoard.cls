VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKeyBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function StartUp()
    frmScreenKeyboard.Show
End Function

Public Function SetPos()
'���ܣ��̶�λ��
'���������ʱ����Ҫ������Ϊ�˲��޸Ĳ�Ʒ�еĴ��룬�������Ҫ����
End Function

Public Function Closed()
    Unload frmScreenKeyboard
End Function

Private Sub Class_Terminate()
    Unload frmScreenKeyboard
End Sub