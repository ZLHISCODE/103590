Attribute VB_Name = "mdlPubDefine"
Option Explicit

'��������ť
Public Const conMenu_Manage_Save = 3503         '����
Public Const conMenu_Manage_Preview = 102       'Ԥ��
Public Const conMenu_Manage_Print = 103         '��ӡ
Public Const conMenu_Manage_Exit = 191          '�˳�
Public Const conMenu_Manage_Modify = 3003       '�޸�
Public Const conMenu_Manage_Audit = 3010        '���

Public Const conMenu_EditPopup = 3    '�༭

Public Const conMenu_Manage_Up = 743
Public Const conMenu_Manage_Down = 744
Public Const conMenu_Manage_Help = 901        '*��������(&H)

Public Const conMenu_Tool_PlugIn = 890          '���
Public Const conMenu_Tool_PlugIn_Item = 89000   '�����,ʵ������Ϊ conMenu_Tool_PlugIn_Item + n, 1<=n<=99


Public Const FCONTROL = 8
            
Public Const P����������� As Integer = 1132

Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long


'*************************************************************************
'**�� �� ����HIWORD
'**��    �룺LongIn(Long) - 32λֵ
'**��    ����(Integer) - 32λֵ�ĵ�16λ
'**����������ȡ��32λֵ�ĸ�16λ
'*************************************************************************
Public Function HIWORD(LongIn As Long) As Integer
   ' ȡ��32λֵ�ĸ�16λ
     HIWORD = (LongIn And &HFFFF0000) \ &H10000
End Function

'*************************************************************************
'**�� �� ����LOWORD
'**��    �룺LongIn(Long) - 32λֵ
'**��    ����(Integer) - 32λֵ�ĵ�16λ
'**����������ȡ��32λֵ�ĵ�16λ
'*************************************************************************
Public Function LOWORD(LongIn As Long) As Integer
   ' ȡ��32λֵ�ĵ�16λ
     LOWORD = LongIn And &HFFFF&
End Function
