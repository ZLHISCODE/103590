VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frm扣率Set 
   BorderStyle     =   0  'None
   Caption         =   "零售价计算器"
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3345
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox PicInput 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2385
      ScaleWidth      =   3315
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3345
      Begin VB.CommandButton CmdNO 
         Caption         =   "取消"
         Height          =   345
         Left            =   2415
         TabIndex        =   2
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton CmdYes 
         Caption         =   "确定"
         Height          =   345
         Left            =   1425
         TabIndex        =   1
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox Txt加价率 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1110
         MaxLength       =   8
         TabIndex        =   3
         Text            =   "15.0000"
         Top             =   1095
         Width           =   2130
      End
      Begin XtremeSuiteControls.ShortcutCaption stcTittle 
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   3990
         _Version        =   589884
         _ExtentX        =   7038
         _ExtentY        =   661
         _StockProps     =   6
         Caption         =   "售价计算器"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "    请输入加成率，零售价的计算公式：零售价=成本价*(1+加成率%)"
         ForeColor       =   &H00400000&
         Height          =   600
         Left            =   60
         TabIndex        =   5
         Top             =   555
         Width           =   3405
      End
      Begin VB.Label Lbl加价率 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "加成率(&J)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   4
         Top             =   1155
         Width           =   870
      End
   End
End
Attribute VB_Name = "frm扣率Set"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private mdbl零售价 As Double
Private mdbl加成率 As Double
Private mdbl结算价 As Double '
Private mlng材料ID As Long
Private mintUnit As Integer
Private mbln不强制控制指导价格 As Boolean

Private msngX As Single, msngY As Single, mlngTxtH As Long
Public Function ShowCalc(ByVal frmMain As Form, _
    ByVal sngX As Single, ByVal sngY As Single, lngTxtH As Long, lng材料ID As Long, intUnit As Integer, _
    ByRef dbl零售价 As Double, ByRef dbl结算价 As Double, ByRef dbl加成率 As Double, ByVal bln不强制控制指导价格 As Boolean) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:零售计算器
    '入参:
    '出参:
    '返回:选择,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-11-07 11:23:35
    '-----------------------------------------------------------------------------------------------------------
    mdbl零售价 = dbl零售价: mdbl结算价 = dbl结算价: mdbl加成率 = dbl加成率: mlng材料ID = lng材料ID: mbln不强制控制指导价格 = bln不强制控制指导价格
    msngX = sngX: msngY = sngY: mlngTxtH = lngTxtH: mintUnit = intUnit
    mblnOk = False
    Me.Show 1, frmMain
    dbl零售价 = mdbl零售价: dbl结算价 = mdbl结算价: dbl加成率 = mdbl加成率
    ShowCalc = mblnOk
End Function

Private Sub InitData()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-11-07 11:30:47
    '-----------------------------------------------------------------------------------------------------------
    
    If mdbl零售价 <> 0 And mdbl结算价 <> 0 Then
        Txt加价率 = Format(计算加成率(), "###0.0000000;-###0.0000000;0;0")
    End If
    Txt加价率.Tag = Txt加价率
    
End Sub

Private Function 计算加成率() As Single
    Dim sin指导零售价 As Single, sin差价让利比 As Single
    Dim rsTemp As New ADODB.Recordset
    '根据零售价反算成本价,由于时价卫材公式的变化,导致原来计算加成率的公式无效,需重新计算
    '原公式:(零售价/成本价-1)*100
    '现公式的理论:由于零售价是按加成率算出来后,再加上了让利外那部分金额,因此实际按加成率算出的零售价=指导零售价-(指导零售价-零售价)/差价让利比
    '再套用原公式算出实际的加成率
    计算加成率 = 0.15
    
    On Error GoTo ErrHandle
    gstrSQL = "Select a.换算系数,a.指导零售价,Nvl(a.差价让利比,100) 差价让利比,Nvl(b.是否变价,0) 时价 From 材料特性 A, 收费项目目录 b Where a.材料ID=b.id  and b.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取指导零售价", mlng材料ID)
    If rsTemp.EOF Then Exit Function
    
    sin指导零售价 = rsTemp!指导零售价
    sin差价让利比 = rsTemp!差价让利比
    If rsTemp!时价 = 0 Then Exit Function
    
    '指导零售价-(指导零售价-零售价)/差价让利比
    sin指导零售价 = sin指导零售价 * IIf(mintUnit = 0, 1, Val(NVL(rsTemp!换算系数)))
    If sin差价让利比 <> 100 And sin差价让利比 > 0 Then
        mdbl零售价 = sin指导零售价 - (sin指导零售价 - mdbl零售价) / sin差价让利比 * 100
    Else
        mdbl零售价 = sin指导零售价 - (sin指导零售价 - mdbl零售价)
    End If
    计算加成率 = (mdbl零售价 / mdbl结算价 - 1) * 100
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function 校正零售价(ByVal dbl零售价 As Double) As Double
    '得到按当前单位系数计算出来的指导零售价，如果时价卫材且强制控制指导价格计算出来的零售价大于指导零售价，以指导零售价为准
    Dim sin指导零售价 As Single
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select 指导零售价, " & IIf(mintUnit = 0, 1, " 换算系数 ") & "   as 换算系数,Nvl(差价让利比,100) 差价让利比 From 材料特性 Where 材料ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取指导零售价", mlng材料ID)
    If rsTemp.EOF Then Exit Function
    sin指导零售价 = zlStr.NVL(rsTemp!指导零售价, 0)
    sin指导零售价 = sin指导零售价 * Val(zlStr.NVL(rsTemp!换算系数))
    If sin指导零售价 = 0 Then sin指导零售价 = dbl零售价
    校正零售价 = IIf(dbl零售价 > sin指导零售价 And Not mbln不强制控制指导价格, sin指导零售价, dbl零售价)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub CmdNO_Click()
    mblnOk = False
    Unload Me
End Sub
Private Sub CmdYes_Click()
    If Val(Txt加价率) > 9900 Or Val(Txt加价率) < 0 Then
        MsgBox "请输入合法的加成率！（0-9900）", vbInformation, gstrSysName
        Txt加价率.SetFocus
        Exit Sub
    End If
    
    mdbl加成率 = Val(Txt加价率)
    '重新计算零售价、差价
    mdbl零售价 = 校正零售价(mdbl结算价 * (1 + (Val(Txt加价率) / 100)) + _
    时价材料零售价(mlng材料ID, mdbl结算价, Val(Txt加价率) / 100))
    mblnOk = True
    Unload Me
End Sub
Private Function 时价材料零售价(ByVal lng材料ID As Long, ByVal sin采购价 As Single, ByVal sin加成率 As Single, _
    Optional sng售价 As Single = -99999999) As Double
    '------------------------------------------------------------------------------------------------------
    '功能:根据指导价格或差价比计算出时价材料的差价让利情况
    '入参:lng材料ID-材料ID
    '     sin采购价-采购价格
    '     sin加成率-加成率(如果传入0,同时又传入dbl零售价,则将按传入的零售价进行计算)
    '     LngLastRow-单据的行号
    '     sng售价-传入的零售价
    '出参:
    '返回:零售价的让利情况
    '修改人:刘兴宏
    '修改时间:2007/2/25
    '------------------------------------------------------------------------------------------------------
    '时价材料零售价计算公式:采购价*(1+加成率)
    '改为:采购价*(1+加成率)+(指导零售价-采购价*(1+加成率))*(1-差价让利比)
    '由于差价让利比的存在,以前所有按指导差价率计算的地方,均需要将差价率转换成加成率进行计算,此函数用于返回本次公式增加的部分金额：(指导零售价-采购价*(1+加成率))*(1-差价让利比)
    
    Dim sin零售价 As Single, sin指导零售价 As Single, sin差价让利比 As Single
    Dim dbl系数 As Double
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select 指导零售价,Nvl(差价让利比,100) 差价让利比," & IIf(mintUnit = 0, 1, "换算系数") & " As  换算系数 From 材料特性 Where 材料ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取指导零售价", lng材料ID)
    
    If rsTemp.EOF Then Exit Function
    
    dbl系数 = Val(zlStr.NVL(rsTemp!换算系数))
    sin指导零售价 = rsTemp!指导零售价
    sin差价让利比 = rsTemp!差价让利比
    
    时价材料零售价 = 0
    If sin差价让利比 = 100 Then Exit Function
    If sin指导零售价 = 0 Then Exit Function
    
    sin零售价 = sin采购价 * (1 + sin加成率)
    If sin零售价 / dbl系数 >= sin指导零售价 Then Exit Function
    sin指导零售价 = sin指导零售价 * dbl系数
    时价材料零售价 = (sin指导零售价 - sin零售价) * (1 - sin差价让利比 / 100)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub Form_Activate()
    Call zlControl.ControlSetFocus(Txt加价率)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call CmdYes_Click
        Exit Sub
    End If
    If KeyCode = vbKeyEscape Then '
        mblnOk = False
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Call InitData
    With Me
        If msngX + .Width > Screen.Width Then
            .Left = Screen.Width - .Width
        Else
            .Left = msngX
        End If
        If msngY + .Height > Screen.Height Then
           .Top = msngY - mlngTxtH - .Height
        Else
            .Top = msngY
        End If
    End With
    
End Sub
Private Sub Txt加价率_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress Txt加价率, KeyAscii, m金额式
End Sub
