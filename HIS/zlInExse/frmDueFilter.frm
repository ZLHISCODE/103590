VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDueFilter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "过滤设置"
   ClientHeight    =   3600
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   7335
   ControlBox      =   0   'False
   Icon            =   "frmDueFilter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   7335
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fra 
      Height          =   3270
      Left            =   120
      TabIndex        =   18
      Top             =   15
      Width           =   5610
      Begin VB.CheckBox chk仅显示欠费 
         Caption         =   "仅显示存在欠款的病人"
         Height          =   225
         Left            =   1665
         TabIndex        =   15
         Top             =   2910
         Width           =   2115
      End
      Begin VB.CommandButton cmd病人 
         Height          =   300
         Left            =   5100
         Picture         =   "frmDueFilter.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "查找(F3)"
         Top             =   2535
         Width           =   330
      End
      Begin VB.TextBox txtUnit 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   13
         Top             =   2520
         Width           =   3780
      End
      Begin VB.TextBox Txt姓名 
         Height          =   300
         Left            =   1680
         MaxLength       =   64
         TabIndex        =   5
         Top             =   1065
         Width           =   3750
      End
      Begin VB.TextBox txt住院号 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1680
         MaxLength       =   18
         TabIndex        =   7
         Top             =   1425
         Width           =   3750
      End
      Begin VB.TextBox txtInvoice 
         Height          =   300
         Left            =   1680
         TabIndex        =   11
         Top             =   2145
         Width           =   3750
      End
      Begin VB.TextBox txtNO 
         Height          =   300
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   9
         Top             =   1785
         Width           =   3750
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   1680
         TabIndex        =   3
         Top             =   720
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   183959555
         CurrentDate     =   39083
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1680
         TabIndex        =   1
         Top             =   300
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   183959555
         CurrentDate     =   39078
      End
      Begin VB.Label lbl合约单位 
         Caption         =   "合约单位(&H)"
         Height          =   210
         Left            =   600
         TabIndex        =   12
         Top             =   2580
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名(&3)"
         Height          =   180
         Left            =   960
         TabIndex        =   4
         Top             =   1125
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号(&4)"
         Height          =   180
         Left            =   780
         TabIndex        =   6
         Top             =   1485
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结帐票据号(&6)"
         Height          =   180
         Left            =   420
         TabIndex        =   10
         Top             =   2205
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结帐单据号(&5)"
         Height          =   180
         Left            =   420
         TabIndex        =   8
         Top             =   1845
         Width           =   1170
      End
      Begin VB.Label lblDateE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结帐结束时间(&2)"
         Height          =   180
         Left            =   240
         TabIndex        =   2
         Top             =   780
         Width           =   1350
      End
      Begin VB.Label lblDateB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结帐开始时间(&1)"
         Height          =   180
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   1350
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5970
      TabIndex        =   16
      Top             =   120
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5970
      TabIndex        =   17
      Top             =   540
      Width           =   1100
   End
End
Attribute VB_Name = "frmDueFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Private Sub cmdCancel_Click()
    gblnOK = False
    Hide
End Sub

Private Sub cmdOK_Click()
    Dim DatTmp As Date
    If dtpBegin.Value > dtpEnd.Value Then
        DatTmp = dtpBegin.Value: dtpBegin.Value = dtpEnd.Value: dtpEnd.Value = DatTmp
    End If
    gblnOK = True
    Hide
End Sub

Private Sub cmd病人_Click()
    If SelectUnits(txtUnit, "") = False Then Exit Sub
End Sub

Private Sub Form_Activate()
    dtpBegin.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    gblnOK = False
    txtInvoice.MaxLength = gbytFactLength
End Sub


Private Sub txtInvoice_GotFocus()
    zlcontrol.TxtSelAll txtInvoice
End Sub

Private Sub txtInvoice_Validate(Cancel As Boolean)
    txtInvoice.Text = Trim(txtInvoice.Text)
End Sub

Private Sub txtNO_GotFocus()
    zlcontrol.TxtSelAll txtNO
End Sub

Private Sub txtNO_Validate(Cancel As Boolean)
    txtNO.Text = Trim(txtNO.Text)
    If txtNO.Text <> "" Then txtNO.Text = GetFullNO(txtNO.Text, 15)
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii <> 13 Then
        If Not (txtNO.Text = "" Or txtNO.SelLength = Len(txtNO.Text) Or txtNO.SelStart = 0) And _
            InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0: Beep: Exit Sub
        End If
    End If
End Sub

Private Sub txtUnit_Change()
    txtUnit.Tag = ""
End Sub

Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txtUnit.Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If txtUnit.Text = "" And txtUnit.Tag = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If SelectUnits(txtUnit, Trim(txtUnit.Text)) = False Then Exit Sub
End Sub
 

Private Sub txt姓名_GotFocus()
    zlcontrol.TxtSelAll txt姓名
End Sub

Private Sub txt姓名_Validate(Cancel As Boolean)
    txt姓名.Text = Replace(Trim(txt姓名.Text), "'", "")
End Sub

Private Sub txt住院号_GotFocus()
    zlcontrol.TxtSelAll txt住院号
End Sub

Private Sub txt住院号_Validate(Cancel As Boolean)
    txt住院号.Text = Trim(txt住院号.Text)
End Sub

Private Sub txt住院号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Function SelectUnits(ByVal objCtl As Control, Optional strKey As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择合约单位
    '入参:strKey-输入值
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2011-11-08 15:00:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    Dim vRect As RECT, strWhere As String, bytStyle As Byte
    Dim sngX As Single, sngY As Single, lngH As Long
    Dim blnCancel As Boolean
 
    On Error GoTo errH
    bytStyle = 2
    strWhere = " Start with 上级id is null Connect by prior ID=上级ID"
    If strKey <> "" Then
        strWhere = " Where 1=1 "
        If zlCommFun.IsCharChinese(strKey) Then
            strWhere = strWhere & " And 名称 like [1]  Order by 名称"
        ElseIf zlCommFun.IsCharAlpha(strKey) Then
            strWhere = strWhere & " And 简码 like upper([1]) Order by 简码"
        ElseIf zlCommFun.IsNumOrChar(strKey) Then
            strWhere = strWhere & " And 编码 like upper([1])  Order by 编码"
        Else
            strWhere = strWhere & " And  (名称 like [1] or 编码 like upper([1]) or 简码 like upper([1])) Order by 编码"
        End If
        bytStyle = 0
        strKey = gstrLike & strKey & "%"
    End If
    
    strSql = "" & _
    "   Select ID,上级ID,编码,名称,简码,地址, 末级,说明," & _
    "               To_Char(建档时间, 'YYYY-MM-DD HH24:MI') 建档时间 " & _
    "   From 合约单位" & _
        strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strKey)
    'ShowSelect:
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    vRect = zlcontrol.GetControlRect(objCtl.hWnd)
    lngH = objCtl.Height
    sngX = vRect.Left - 15: sngY = vRect.Top
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSql, bytStyle, "合约单位选择", IIf(bytStyle = 2, True, False), "", "请选择符合条件的合约单位", IIf(bytStyle = 2, True, False), True, True, sngX, sngY, lngH, blnCancel, False, True, strKey)
    If blnCancel Then
        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
        zlcontrol.TxtSelAll objCtl
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgbox "不存在符何条件的合约单位,请检查!"
        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
        zlcontrol.TxtSelAll objCtl
        Exit Function
        Exit Function
    End If
    If rsTemp.State <> 1 Then
        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
        zlcontrol.TxtSelAll objCtl
        Exit Function
        Exit Function
    End If
    With rsTemp
        objCtl.Text = Nvl(!名称): objCtl.Tag = Nvl(!ID)
    End With
    If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
    zlcontrol.TxtSelAll objCtl
    zlCommFun.PressKey vbKeyTab
    SelectUnits = True
    Exit Function
errH:
    If ErrCenter = 1 Then Resume
End Function
