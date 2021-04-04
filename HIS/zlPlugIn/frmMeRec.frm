VERSION 5.00
Begin VB.Form frmMeRec 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "外挂附页测试"
   ClientHeight    =   1890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   Icon            =   "frmMeRec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdTest 
      Appearance      =   0  'Flat
      Caption         =   "按钮测试"
      Height          =   855
      Left            =   3840
      TabIndex        =   2
      Top             =   840
      Width           =   3855
   End
   Begin VB.TextBox txtPic 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Text            =   "外挂附页文本框测试"
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label lblTest 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "外挂附页测试"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmMeRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'窗口功能:加载自定义病案外挂附页

'窗口说明:1.zlhis35以上窗体宽度固定为10995 保持首页格式；zlhis34住院首页宽度为10785，病案首页宽度为11985
'         2.窗体的Caption,用与首页加载页签内容
'         3.注:在修改函数SavePlugMec的时候请不要写入耗时的代码
'         4.函数CheckPlugMec:病案外挂附页输入有效性检查
'         5.函数SavePlugMec:组建病案外挂附页保存数据SQL
'         6.函数LoadPlugMec:病案外挂附页加载数据
'         7.窗体的Tag值:用于保存窗体对应病案首页图片的index
'         8.控件的Tag值:用于保存检查的提醒信息 格式:((提醒:1/禁止:0) | 提示消息| 窗体Tag值)
'         9.gblnChange:判断本窗体控件值是否发生改变


Public gblnChange As Boolean '是否改变控件数据

'首页数据
Public glngSys As Long
Public glngModule As Long
Public glngPatiID As Long
Public glngPageID As Long
Public glngPatiType As Long


Private Sub cmdTest_Click()
    MsgBox "外挂附页测试"
End Sub

Private Sub Form_Load()
    'zlhis35以上设置宽度固定为10995 保持病案首页标准格式
    Me.Width = 10995
    
'   zlhis34设置窗体窗体宽度
    '设置宽度固定为10785 保持住院首页标准格式
'    Me.Width = 10785
    '设置宽度固定为11985 保持病案首页标准格式
'    Me.Width = 11985
End Sub
'

Public Function CheckPlugMec(ByVal lngSys As Long, ByVal lngModule As Long, ByVal lngPatiID As Long, ByVal lngPageID As Long, ByRef colErr As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病案外挂附页输入有效性检查
    '返回：True是成功，False是失败
    '参数：colErr 返回集合参数 格式例: 控件对象,key值(控件名+ 窗体Tag值+(控件index)) 不是数组控件则不拼接index
    '      控件的tag值:保存提示信息  例 : ((提醒:1/禁止:0) | 提示消息| 窗体Tag值)
    '      lngSys,lngModual=当前调用接口的主程序系统号及模块号
    '      lngPageID－主页ID；
    
    '返回:检查通过返回true,否则返回False
    '编制:蒋廷中
    '日期:2017年6月20日 11:52:48
    Dim strKey As String
    On Error GoTo errHandle
    CheckPlugMec = True
    


    If txtPic.Text = "" Then
        txtPic.Tag = "0|外挂病案首页文本框不能为空" & "|" & Me.Tag
        
        '检查控件是否为数组控件
        If VarType(txtPic) <> vbObject Then
            strKey = txtPic.Name & Me.Tag
        Else
            strKey = txtPic.Name & Me.Tag & txtPic.Index
        End If
        colErr.Add txtPic, strKey
    End If

    '返回参数格式示例
    If colErr.Count > 0 Then
        CheckPlugMec = False
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function SavePlugMec(ByVal lngSys As Long, ByVal lngModule As Long, ByVal lngPatiID As Long, ByVal lngPageID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:组建病案外挂附页保存数据SQL并通过公共部件的gOracle进行执行
    '返回：True是成功，False是失败
    '参数： lngSys,lngModual=当前调用接口的主程序系统号及模块号
    '       lngPatiID:病人id
    '      lngPageID－主页ID；
    '返回:保存通过返回true,否则返回False
    '编制:蒋廷中
    '日期:2017年6月20日 11:52:48
    Dim strSql As String
    On Error GoTo errHandle
    
    strSql = "zl_病人信息从表_Update(301,'身份证号状态','遗失待办')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    SavePlugMec = True
    
    gblnChange = False
    Exit Function
errHandle:
    SavePlugMec = False
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function LoadPlugMec(ByVal lngSys As Long, ByVal lngModule As Long, ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal lngPatiType As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病案外挂附页加载数据
    '返回：True是成功，False是失败
    '      lngSys,lngModual=当前调用接口的主程序系统号及模块号
    '      lngPageID－主页ID；
    '      lngPatiType－病人类型:1-门诊,2-住院
    '      lngPatiID-病人id
    '编制:蒋廷中
    '日期:2017年6月21日 9:52:48

   On Error GoTo errHandle
    LoadPlugMec = True
    txtPic.Text = lngSys & "|" & lngModule & "|" & lngPatiID & "|" & lngPageID & "|" & lngPatiType
    gblnChange = False
    Exit Function
errHandle:
    MsgBox Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
End Function





'是否改变控件数据
Private Sub txtPic_Change()
    gblnChange = True
End Sub




'支持外挂附页滚轮,35以上版本启用
Private Sub Form_Activate()
    glngPreHWnd = GetWindowLong(Me.hwnd, GWL_WNDPROC)
    SetWindowLong Me.hwnd, GWL_WNDPROC, AddressOf MecFlexScroll
End Sub

'支持外挂附页滚轮,35以上版本启用
Private Sub Form_Deactivate()
    SetWindowLong Me.hwnd, GWL_WNDPROC, glngPreHWnd
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "病案外挂附页 卸载了！！！！！"
End Sub

