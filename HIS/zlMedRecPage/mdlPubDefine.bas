Attribute VB_Name = "mdlPubDefine"
Option Explicit

'工具栏按钮
Public Const conMenu_Manage_Save = 3503         '保存
Public Const conMenu_Manage_Preview = 102       '预览
Public Const conMenu_Manage_Print = 103         '打印
Public Const conMenu_Manage_Exit = 191          '退出
Public Const conMenu_Manage_Modify = 3003       '修改
Public Const conMenu_Manage_Audit = 3010        '审核

Public Const conMenu_EditPopup = 3    '编辑

Public Const conMenu_Manage_Up = 743
Public Const conMenu_Manage_Down = 744
Public Const conMenu_Manage_Help = 901        '*帮助主题(&H)

Public Const conMenu_Tool_PlugIn = 890          '插件
Public Const conMenu_Tool_PlugIn_Item = 89000   '插件项,实际依次为 conMenu_Tool_PlugIn_Item + n, 1<=n<=99


Public Const FCONTROL = 8
            
Public Const P病人入出管理 As Integer = 1132

Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long


'*************************************************************************
'**函 数 名：HIWORD
'**输    入：LongIn(Long) - 32位值
'**输    出：(Integer) - 32位值的低16位
'**功能描述：取出32位值的高16位
'*************************************************************************
Public Function HIWORD(LongIn As Long) As Integer
   ' 取出32位值的高16位
     HIWORD = (LongIn And &HFFFF0000) \ &H10000
End Function

'*************************************************************************
'**函 数 名：LOWORD
'**输    入：LongIn(Long) - 32位值
'**输    出：(Integer) - 32位值的低16位
'**功能描述：取出32位值的低16位
'*************************************************************************
Public Function LOWORD(LongIn As Long) As Integer
   ' 取出32位值的低16位
     LOWORD = LongIn And &HFFFF&
End Function
