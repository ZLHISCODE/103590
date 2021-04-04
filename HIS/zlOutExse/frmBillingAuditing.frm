VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBillingAuditing 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人划价单据审核"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8595
   Icon            =   "frmBillingAuditing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   90
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   8400
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   3780
      Width           =   8400
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshBill 
      Height          =   2265
      Left            =   30
      TabIndex        =   8
      ToolTipText     =   "双击单据查看明细"
      Top             =   3825
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   3995
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmBillingAuditing.frx":058A
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   960
      Left            =   0
      ScaleHeight     =   960
      ScaleWidth      =   8595
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6120
      Width           =   8595
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   300
         TabIndex        =   16
         Top             =   525
         Width           =   1200
      End
      Begin VB.CommandButton cmdFlash 
         Caption         =   "刷新(&R)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   300
         TabIndex        =   15
         Top             =   90
         Width           =   1200
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "选择(&S)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1695
         TabIndex        =   11
         Top             =   90
         Width           =   1200
      End
      Begin VB.CommandButton cmdCls 
         Caption         =   "清除(&M)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3000
         TabIndex        =   12
         Top             =   90
         Width           =   1200
      End
      Begin VB.CommandButton cmdClsAll 
         Caption         =   "全清(&C)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3000
         TabIndex        =   14
         Top             =   525
         Width           =   1200
      End
      Begin VB.CommandButton cmdSelAll 
         Caption         =   "全选(&A)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1695
         TabIndex        =   13
         Top             =   525
         Width           =   1200
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "退出(&X)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   7140
         TabIndex        =   10
         Top             =   525
         Width           =   1200
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "审核(&O)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   7140
         TabIndex        =   9
         Top             =   90
         Width           =   1200
      End
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   -15
      ScaleHeight     =   240
      ScaleWidth      =   9585
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1110
      Width           =   9585
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "未审核划价单,当前合计:"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Tag             =   "未审核划价单,当前合计:"
         Top             =   30
         Width           =   1980
      End
   End
   Begin VB.Frame fraInfo 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   45
      TabIndex        =   17
      Top             =   -45
      Width           =   8505
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   360
         Left            =   555
         TabIndex        =   30
         Top             =   180
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   635
         Appearance      =   2
         IDKindStr       =   "姓|姓名|0;医|医保号|0;身|身份证号|0;IC|IC卡号|1;门|门诊号|0;就|就诊卡|0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   12
         FontName        =   "宋体"
         IDKind          =   -1
         ShowPropertySet =   -1  'True
         BackColor       =   -2147483633
      End
      Begin VB.TextBox txt剩余 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6800
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   675
         Width           =   1440
      End
      Begin VB.TextBox txt费用 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   675
         Width           =   1200
      End
      Begin VB.TextBox txt预交 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   675
         Width           =   1200
      End
      Begin VB.TextBox txt费别 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6795
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   1440
      End
      Begin VB.TextBox txt年龄 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5205
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   180
         Width           =   960
      End
      Begin VB.TextBox txt性别 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   180
         Width           =   615
      End
      Begin VB.TextBox txtPatient 
         BackColor       =   &H00EBFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1230
         MaxLength       =   100
         TabIndex        =   0
         ToolTipText     =   "热键:F6"
         Top             =   180
         Width           =   1680
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   30
         X2              =   8450
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000015&
         X1              =   30
         X2              =   8450
         Y1              =   585
         Y2              =   585
      End
      Begin VB.Label lbl剩余款 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "剩余款"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6000
         TabIndex        =   24
         Top             =   765
         Width           =   630
      End
      Begin VB.Label lbl未结费用 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "未结费用"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3000
         TabIndex        =   23
         Top             =   765
         Width           =   840
      End
      Begin VB.Label lbl预交余额 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预交余额"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   22
         Top             =   765
         Width           =   840
      End
      Begin VB.Label lbl费别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6255
         TabIndex        =   21
         Top             =   255
         Width           =   420
      End
      Begin VB.Label lbl年龄 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4680
         TabIndex        =   20
         Top             =   255
         Width           =   420
      End
      Begin VB.Label lbl性别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3480
         TabIndex        =   19
         Top             =   255
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   420
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   2415
      Left            =   30
      TabIndex        =   7
      ToolTipText     =   "双击单据查看明细"
      Top             =   1365
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   4260
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmBillingAuditing.frx":08A4
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   29
      Top             =   7080
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBillingAuditing.frx":0BBE
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10557
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmBillingAuditing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Private mstrPrivs As String
Private mlngModule As Long
Private mrsInfo As New ADODB.Recordset
Private mrsList As ADODB.Recordset
Attribute mrsList.VB_VarHelpID = -1
Private mlngCurRow As Long, mlngTopRow As Long
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mintSucces As Integer
Private mblnNotClick As Boolean

'-----------------------------------------------------------------------------------
'结算卡相关
Private mstrPassWord As String
'-----------------------------------------------------------------------------------
Private mobjDrugMachine As Object '自动发药机(新）
Private mblnDrugMachine As Boolean

Public Function zlShowCard(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口
    '返回:审核成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-03 15:23:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mintSucces = 0: mlngModule = lngModule: mstrPrivs = strPrivs
    Me.Show 1, frmMain
    zlShowCard = mintSucces > 0
End Function
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCls_Click()
    Dim i As Long, intS As Integer
    intS = 1
    If mshList.Row > mshList.RowSel Then intS = -1
    For i = mshList.Row To mshList.RowSel Step intS
        If mshList.TextMatrix(i, 1) <> "" Then
            mshList.TextMatrix(i, 0) = ""
        End If
    Next
    lblTotal.Caption = lblTotal.Tag & Format(CalcTotal, gstrDec)
End Sub

Private Sub cmdClsAll_Click()
    Dim i As Long
    mshList.Redraw = False
    For i = 1 To mshList.Rows - 1
        If mshList.TextMatrix(i, 1) <> "" Then
            mshList.TextMatrix(i, 0) = ""
        End If
    Next
    lblTotal.Caption = lblTotal.Tag & Format(CalcTotal, gstrDec)
    mshList.Redraw = True
End Sub

Private Sub cmdFlash_Click()
    If mrsInfo.State = 0 Then
        MsgBox "没有确定病人,请先输入病人信息！", vbInformation, gstrSysName
        txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
    End If
    Call ShowBills
End Sub

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim strDel As String, i As Long, str审核时间 As String, Curdate As Date, curTotal As Currency
    Dim arrSQL As Variant, strNos As String, strNo As String, blnTrans As Boolean
    
    arrSQL = Array()
    For i = 1 To mshList.Rows - 1
        If mshList.TextMatrix(i, 0) <> "" And mshList.TextMatrix(i, 1) <> "" Then
            If str审核时间 = "" Then
                Curdate = zlDatabase.Currentdate
                str审核时间 = "To_Date('" & Format(Curdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
            End If
            strNo = mshList.TextMatrix(i, 1)
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_门诊记帐记录_Verify('" & strNo & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "',Null," & str审核时间 & ")"
            strDel = strDel & "," & i
            
            strNos = strNos & "," & strNo
        End If
    Next
    If UBound(arrSQL) = -1 Then
        MsgBox "没有选择要审核的划价单据！", vbInformation, gstrSysName
        Exit Sub
    End If
    curTotal = CalcTotal
    If curTotal <> 0 And gdbl预存款消费验卡 <> 0 Then
        If Not zlDatabase.PatiIdentify(Me, glngSys, Val(mrsInfo!病人ID), curTotal, , , , IIf(-1 * gdbl预存款消费验卡 >= curTotal, False, True), , , , (gdbl预存款消费验卡 = 2)) Then Exit Sub
    End If
    
    strNos = Mid(strNos, 2)
    
    '费用报警
    If Not AuditingWarnByPatient(strNos) Then Exit Sub
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    If gbln审核打印 Then
        For i = 0 To UBound(Split(strNos, ","))
            strNo = Split(strNos, ",")(i)
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1122", Me, "NO=" & strNo, "药品单位=" & IIf(gbln药房单位, 1, 0), "PrintEmpty=0", 2)
        Next
    End If
    
    '110319
    If mblnDrugMachine Then
        '门诊格式：1|单据1,处方号1;单据2,处方号2
        Dim strData As String, strReturn As String
        strData = "1|" & "9," & Replace(strNos, ",", ";9,")
        Call mobjDrugMachine.Operation(gstrDBUser, Val("21-配药[门诊和住院处方明细上传]"), strData, strReturn)
    End If
    
    On Error GoTo 0
    
    '注意方向
    strDel = Mid(strDel, 2)
    For i = UBound(Split(strDel, ",")) To 0 Step -1
        If mshList.Rows > 2 Then
            mshList.RemoveItem CLng(Split(strDel, ",")(i))
        Else
            mshList.Clear
            mshList.Rows = 2
            Call SetHeader
        End If
    Next
    
    Call mshList_EnterCell
    
    lblTotal.Caption = lblTotal.Tag & Format(CalcTotal, gstrDec)
    Call RefreshMoney
    
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    mintSucces = mintSucces + 1
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSel_Click()
    Dim i As Long, intS As Integer
    intS = 1
    If mshList.Row > mshList.RowSel Then intS = -1
    For i = mshList.Row To mshList.RowSel Step intS
        If mshList.TextMatrix(i, 1) <> "" Then
            mshList.TextMatrix(i, 0) = "√"
        End If
    Next
    lblTotal.Caption = lblTotal.Tag & Format(CalcTotal, gstrDec)
End Sub

Private Sub cmdSelAll_Click()
    Dim i As Long
    mshList.Redraw = False
    For i = 1 To mshList.Rows - 1
        If mshList.TextMatrix(i, 1) <> "" Then
            mshList.TextMatrix(i, 0) = "√"
        End If
    Next
    lblTotal.Caption = lblTotal.Tag & Format(CalcTotal, gstrDec)
    mshList.Redraw = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF4
            If Shift <> vbCtrlMask Then Exit Sub
            If IDKind.Enabled Then
                Dim intIndex As Integer
                intIndex = IDKind.GetKindIndex("IC卡号")
                If intIndex <= 0 Then Exit Sub
                IDKind.IDKind = intIndex: Call IDKind_Click(IDKind.GetCurCard)
            End If
        Case vbKeyF6
            txtPatient.SetFocus
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    Call initCardSquareData
    Me.Height = 7815 '没有写Resize
    Set mobjIDCard = New clsIDCard
    Call mobjIDCard.SetParent(Me.hWnd)
    Call CreateDrugPacker
    Call SetHeader
    Call SetBill
    txtPatient.MaxLength = zlGetPatiInforMaxLen.intPatiName
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    Set mobjICCard = Nothing
    
    Set mrsList = Nothing
    Set mrsInfo = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub

 
Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If mobjICCard Is Nothing Then Exit Sub
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text <> "" Then
            Call FindPati(objCard, True, txtPatient.Text)
        End If
        Exit Sub
    End If
    
    lng卡类别ID = objCard.接口序号
    If lng卡类别ID <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOlnyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strOutPatiInforXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败\
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModule, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then
        Call FindPati(objCard, True, txtPatient.Text)
    End If
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    If mblnNotClick Then Exit Sub
    Set gobjSquare.objCurCard = objCard
    If txtPatient.Text <> "" Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub
Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    Dim lngPreIDKind As Long
      '问题:60010
    If txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.卡号
    If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    Dim lngPreIDKind As Long
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    txtPatient.Text = strCardNo
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("IC卡", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    '身份证识别
    Dim lngPreIDKind As Long
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("身份证", CardTypeName)
    txtPatient.Text = strID
    If objCard Is Nothing Then Exit Sub
    If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
End Sub

Private Sub mshList_DblClick()
    If mshList.MouseRow = 0 Then Exit Sub
    If mshList.TextMatrix(mshList.Row, 1) = "" Then Exit Sub
    
    If mshList.MouseCol = 0 Then
        If mshList.TextMatrix(mshList.Row, 0) = "" Then
            mshList.TextMatrix(mshList.Row, 0) = "√"
        Else
            mshList.TextMatrix(mshList.Row, 0) = ""
        End If
        lblTotal.Caption = lblTotal.Tag & Format(CalcTotal, gstrDec)
    Else
       Call ShowBill
    End If
End Sub

Private Sub mshList_EnterCell()
    If mshList.Row = 0 Or mshList.TextMatrix(mshList.Row, 1) = "" Then
        mshBill.Clear
        mshBill.Rows = 2
        Call SetBill
        Exit Sub
    End If
    mlngCurRow = mshList.Row: mlngTopRow = mshList.TopRow
    
    Call ShowDetail(mshList.TextMatrix(mshList.Row, 1))
End Sub

Private Sub ShowDetail(Optional strNo As String)
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim strSQL As String
    
    On Error GoTo errH
    
    '明细费用行的剩余数量和金额
    strSQL = _
    " Select C.名称 as 类别,Nvl(E.名称,B.名称) as 名称," & IIf(gTy_System_Para.byt药品名称显示 = 2, "E1.名称 as 商品名,", "") & "B.规格," & _
            IIf(gbln药房单位, "Decode(X.药品ID,NULL,A.计算单位,X." & gstr药房单位 & ")", "A.计算单位") & " as 单位," & _
    "       Avg(Nvl(A.付数,1)*A.数次)" & IIf(gbln药房单位, "/Nvl(X." & gstr药房包装 & ",1)", "") & " as 数量, " & _
    "       Ltrim(To_Char(Sum(A.标准单价)" & IIf(gbln药房单位, "*Nvl(X." & gstr药房包装 & ",1)", "") & ",'99999" & gstrFeePrecisionFmt & "')) as 单价," & _
    "       Ltrim(To_Char(Sum(A.应收金额),'99999" & gstrDec & "')) as 应收金额," & _
    "       Ltrim(To_Char(Sum(A.实收金额),'99999" & gstrDec & "')) as 实收金额," & _
    "       D.名称 as 执行科室" & _
    " From 门诊费用记录 A,收费项目目录 B,收费项目类别 C,部门表 D,收费项目别名 E,药品规格 X" & _
            IIf(gTy_System_Para.byt药品名称显示 = 2, ",收费项目别名 E1", "") & _
    " Where A.收费细目ID=B.ID and A.收费类别=C.编码 And A.执行部门ID=D.ID(+)" & _
    "       And A.NO=[1] And A.记录性质=2 And A.门诊标志 In(1,3,4) And A.记录状态=0" & _
    "       And A.病人ID+0=[2]" & _
    "       And A.收费细目ID=X.药品ID(+) And A.操作员姓名 is NULL And A.划价人 is Not NULL" & _
    "       And A.收费细目ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
            IIf(gTy_System_Para.byt药品名称显示 = 2, "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3", "") & _
    " Group by Nvl(A.价格父号,A.序号),C.名称," & _
    "       Nvl(E.名称,B.名称),B.规格,A.计算单位,D.名称," & IIf(gTy_System_Para.byt药品名称显示 = 2, "E1.名称,", "") & " X.药品ID,X." & gstr药房单位 & ",Nvl(X." & gstr药房包装 & ",1)" & _
    " Order by Nvl(A.价格父号,A.序号)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, CLng(mrsInfo!病人ID))
    
    mshBill.Redraw = False
    mshBill.ClearStructure
    mshBill.Clear
    mshBill.Rows = 2
    If Not rsTmp.EOF Then Set mshBill.DataSource = rsTmp
    Call SetBill
    mshBill.Redraw = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetBill()
    Dim strHead As String
    Dim i As Long
    
    strHead = "类别,1,650|名称,1,1500" & IIf(gTy_System_Para.byt药品名称显示 = 2, "|商品名,1,2000", "") & "|规格,1,1500|单位,1,500|数量,1,750|单价,7,750|应收金额,7,850|实收金额,7,850|执行科室,1,1000"
    With mshBill
        .COLS = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshBill, App.ProductName & "\" & Me.Name)
        For i = 0 To .COLS - 1
            If .TextMatrix(0, i) = "商品名" Then
                If gTy_System_Para.byt药品名称显示 = 2 Then
                    If .ColWidth(i) <= 0 Then .ColWidth(i) = 2000
                Else
                    .ColWidth(i) = 0
                End If
            End If
        Next
        
        .RowHeight(0) = 320
        
        .Col = 0: .ColSel = .COLS - 1
    End With
End Sub

Private Sub mshList_KeyPress(KeyAscii As Integer)
    
    If mshList.TextMatrix(mshList.Row, 1) = "" Then Exit Sub
    
    If KeyAscii = 32 Then
        If mshList.TextMatrix(mshList.Row, 0) = "" Then
            mshList.TextMatrix(mshList.Row, 0) = "√"
        Else
            mshList.TextMatrix(mshList.Row, 0) = ""
        End If
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        Call ShowBill
    End If
End Sub

Private Sub ShowBill()
    Dim strNo As String

    On Error Resume Next
        
    strNo = mshList.TextMatrix(mshList.Row, 1)
    
    frmCharge.mlngModul = 1122
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mbytInFun = 2
    frmCharge.mbytInState = 1
    frmCharge.mstrTime = ""
    frmCharge.mblnDelete = False
    frmCharge.mstrInNO = strNo
    frmCharge.mblnNOMoved = False
    frmCharge.mbytBilling = 1
    frmCharge.Show 1, Me
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If mshList.Height + Y < 1000 Or mshBill.Height - Y < 1000 Then Exit Sub
        pic.Top = pic.Top + Y
        mshList.Height = mshList.Height + Y
        mshBill.Top = mshBill.Top + Y
        mshBill.Height = mshBill.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub txtPatient_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    IDKind.SetAutoReadCard (txtPatient.Text = "")
End Sub

Private Sub txtPatient_GotFocus()
    zlControl.TxtSelAll txtPatient
     If txtPatient.Locked Then Exit Sub
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPatient.Text = "")
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtPatient.Text = "")
    IDKind.SetAutoReadCard (txtPatient.Text = "")

End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
    If IDKind.GetCurCard Is Nothing Then Exit Sub
       If IDKind.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean, blnICCard As Boolean
    
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
    '问题:51488
    If (IDKind.Cards.读卡快键 = "空格键" Or IDKind.Cards.读卡快键 = " ") And Chr(KeyAscii) = " " Then KeyAscii = 0: Exit Sub
 
    If IDKind.GetCurCard.名称 Like "姓名*" Then
        '103563,只要输入的第一个字符是“-+*”，后面是全数字，都认为不是刷卡
        If Not (InStr("-+*", Left(txtPatient.Text, 1)) > 0 And IsNumeric(Mid(txtPatient.Text, 2))) Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        End If
    ElseIf IDKind.GetCurCard.名称 = "门诊号" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
    End If
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        Call FindPati(IDKind.GetCurCard, blnCard, txtPatient.Text)
    End If
End Sub
Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查找病人
    '编制:刘兴洪
    '日期:2012-08-31 17:54:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnICCard As Boolean, blnIDCard As Boolean
    
    '读取病人信息
    Call ClearPati
    mshList.Clear: mshList.Rows = 2
    Call SetHeader

    If objCard.名称 Like "IC卡*" And objCard.系统 And mstrPassWord <> "" Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If objCard.名称 Like "*身份证*" And objCard.系统 And mstrPassWord <> "" Then blnIDCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    
    If Not GetPatient(objCard, txtPatient.Text, blnCard) Then
        txtPatient.Text = ""
        If blnCard Then
            sta.Panels(2) = "不能确定病人信息，请检查是否正确刷卡！"
            txtPatient.SetFocus: Exit Sub
        End If
        sta.Panels(2) = "输入的标识不能读取病人信息，请检查输入是否正确！"
        txtPatient.SetFocus: Exit Sub
    End If
    '就诊卡密码检查
    If (objCard.名称 Like "IC卡*" Or objCard.名称 Like "*身份证*") And objCard.系统 = True And blnCard Then blnCard = False
    If Mid(gstrCardPass, 4, 1) = "1" And (blnCard Or blnICCard Or blnIDCard) Then
        If Not zlCommFun.VerifyPassWord(Me, mstrPassWord, mrsInfo!姓名, mrsInfo!性别, "" & mrsInfo!年龄) Then
            Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
        End If
    End If
    txtPatient.PasswordChar = ""
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    txtPatient.Text = "" & mrsInfo!姓名
    txt性别.Text = "" & mrsInfo!性别
    txt年龄.Text = "" & mrsInfo!年龄
    txt费别.Text = "" & mrsInfo!费别
    Call RefreshMoney
    Call ShowBills
    mshList.SetFocus
 
End Sub
 
 Private Sub RefreshMoney()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetMoneyInfo(mrsInfo!病人ID, , , 1)
    If Not rsTmp Is Nothing Then
        txt预交.Text = Format(rsTmp!预交余额, "0.00")
        txt费用.Text = Format(rsTmp!费用余额, gstrDec)
        txt剩余.Text = Format(rsTmp!预交余额 - rsTmp!费用余额, "0.00")
    Else
        txt预交.Text = ""
        txt费用.Text = ""
        txt剩余.Text = ""
    End If
End Sub

Private Sub ClearPati()
    txt性别.Text = ""
    txt年龄.Text = ""
    txt费别.Text = ""
    
    txt费用.Text = ""
    txt预交.Text = ""
    txt剩余.Text = ""
End Sub

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, ByVal blnCard As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息
    '入参:blnCard=是否就诊卡刷卡
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-03 16:49:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    Dim strPati As String, vRect As RECT, blnCancel As Boolean
    Dim strIF As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSQL = _
            "Select A.病人ID,A.就诊卡号,A.卡验证码,A.姓名,A.性别,A.年龄,A.费别,A.险类,A.病人类型" & _
            " From 病人信息 A" & _
            " Where A.停用时间 is NULL "
            
    If blnCard = True And objCard.名称 Like "姓名*" And InStr("-+*", Left(strInput, 1)) = 0 Then    '103563
        lng卡类别ID = IDKind.GetDefaultCardTypeID
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg, lng卡类别ID) = False Then GoTo NotFoundPati:
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        strSQL = strSQL & " And A.病人ID=[1] "
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '病人ID
        strSQL = strSQL & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号(病人出院)
        strSQL = strSQL & " And A.住院号=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号
        strSQL = strSQL & " And A.门诊号=[1]"
        '75087,冉俊明,2014-7-29,门诊病人收费时,不需要输入完整的门诊号,只需要输入门诊号的最后顺序号即能找到当天就诊的病人信息、费用
        strInput = "*" & zlCommFun.GetFullNO(Mid(strInput, 2), 3)
    Else
        Select Case objCard.名称
            Case "姓名", "姓名或就诊卡"  '当作姓名
                 strPati = _
                    " Select A.病人ID as ID,A.病人ID,A.姓名,A.性别,A.年龄,A.就诊卡号,A.卡验证码,A.门诊号,A.出生日期,A.身份证号,A.家庭地址,A.工作单位" & _
                    " From 病人信息 A Where  A.姓名 Like  [1]  " & _
                                IIf(gintNameDays = 0, "", " And (A.就诊时间>Trunc(Sysdate-" & gintNameDays & ") Or A.登记时间>Trunc(Sysdate-" & gintNameDays & "))") & _
                    " And Rownum<101" & _
                    " Order by A.姓名"
                vRect = zlControl.GetControlRect(txtPatient.hWnd)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "病人Find", False, "", "", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%")
                If blnCancel Then Exit Function
                If Not rsTmp Is Nothing Then
                    strInput = rsTmp!病人ID
                    strSQL = strSQL & " And A.病人ID=[2]"
                Else
                    Exit Function
                End If
                
            Case "医保号"
                strInput = UCase(strInput)
                strSQL = strSQL & " And A.医保号=[2]"
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.门诊号=[2]"
                '75087,冉俊明,2014-7-29,门诊病人收费时,不需要输入完整的门诊号,只需要输入门诊号的最后顺序号即能找到当天就诊的病人信息、费用
                strInput = zlCommFun.GetFullNO(strInput, 3)
            Case Else
                '其他类别的,获取相关的病人ID
                If objCard.接口序号 > 0 Then
                    lng卡类别ID = objCard.接口序号
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng病人ID = 0 Then GoTo NotFoundPati:
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng病人ID <= 0 Then GoTo NotFoundPati:
                strSQL = strSQL & " And A.病人ID=[1]"
                strInput = "-" & lng病人ID
                blnHavePassWord = True
        End Select
    End If
        
    txtPatient.ForeColor = Me.ForeColor
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    If mrsInfo.RecordCount <> 0 Then '
        '75259:李南春，2014-7-10，病人姓名显示颜色处理
        Call SetPatiColor(txtPatient, Nvl(mrsInfo!病人类型), IIf(IsNull(mrsInfo!险类), Me.ForeColor, vbRed))
        mstrPassWord = strPassWord
        If Not blnHavePassWord Then mstrPassWord = Nvl(mrsInfo!卡验证码)
        GetPatient = True
    Else
        Set mrsInfo = New ADODB.Recordset
    End If
    Exit Function
NotFoundPati:
    Set mrsInfo = New ADODB.Recordset
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set mrsInfo = New ADODB.Recordset
End Function

Private Sub ShowBills(Optional blnSort As Boolean)
'功能:按条件读取单据列表(过滤功能)
'参数:strIF=以"AND"开始的条件串
'     blnSort=不重新读取数据,仅重新显示已排序的内容
    Dim i As Long, Curdate As Date
    
    On Error GoTo errH
    
    If Not blnSort Then
        sta.Panels(2).Text = "正在读取病人划价单据,请稍候 ..."
        Screen.MousePointer = 11
        DoEvents
        Me.Refresh
        
        gstrSQL = _
        " Select NULL as 审核,A.NO as 单据号," & _
        "       B.名称 as 开单科室,A.开单人 as 医生,A.费别," & _
        "       LTrim(To_Char(Sum(A.应收金额),'999999999" & gstrDec & "')) as 应收金额," & _
        "       LTrim(To_Char(Sum(A.实收金额),'999999999" & gstrDec & "')) as 实收金额," & _
        "       A.划价人,To_Char(A.登记时间,'YYYY-MM-DD HH24:MI:SS') as 划价时间" & _
        " From 门诊费用记录 A,部门表 B" & _
        " Where A.记录性质=2 And A.门诊标志 In(1,3,4) And A.记录状态=0" & _
        "       And A.划价人 is Not Null And A.操作员姓名 is NULL And A.开单部门ID=B.ID" & _
        "       And A.病人ID=[1]" & _
        " Group by A.NO,B.名称,A.开单人,A.费别,A.登记时间,A.划价人" & _
        " Order by 划价时间 Desc,单据号 Desc"
        Set mrsList = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsInfo!病人ID))
    End If
    
    mshList.Redraw = False
    mshList.ClearStructure
    mshList.Clear
    mshList.Rows = 2
    
    If mrsList.EOF Then
        sta.Panels(2).Text = "没有发现划价单据"
    Else
        Set mshList.DataSource = mrsList
        sta.Panels(2).Text = "共 " & mrsList.RecordCount & " 张划价单据"
    End If
    Call SetHeader
        
    lblTotal.Caption = lblTotal.Tag & Format(CalcTotal, gstrDec)
    
    mshList.Redraw = True
    Screen.MousePointer = 0
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetHeader()
    Dim strHead As String
    Dim i As Long
    
    strHead = "审核,4,500|单据号,1,820|开单科室,1,1000|医生,1,750|费别,1,500|应收金额,7,850|实收金额,7,850|划价人,1,700|划价时间,4,1850"
    With mshList
        .COLS = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshList, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 320
        
        '恢复上次行
        If mlngCurRow = 0 Then mlngCurRow = 1
        If mlngTopRow = 0 Then mlngTopRow = 1
        If mlngCurRow <= .Rows - 1 Then
            .Row = mlngCurRow
        Else
            .Row = .Rows - 1
        End If
        If mlngTopRow <= .Rows - 1 Then
            .TopRow = mlngTopRow
        Else
            .TopRow = .Row
        End If
        
        .Col = 0: .ColSel = .COLS - 1
                
        Call mshList_EnterCell
    End With
End Sub

Private Sub mshList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshList.MouseRow = 0 Then
        mshList.MousePointer = 99
    Else
        mshList.MousePointer = 0
    End If
End Sub

Private Sub mshList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = mshList.MouseCol
    
    If Button = 1 And mshList.MousePointer = 99 Then
        If mshList.TextMatrix(0, lngCol) = "" Then Exit Sub
        If mshList.TextMatrix(mshList.Row, 1) = "" Then Exit Sub
        If mrsList Is Nothing Then Exit Sub
        
        Set mshList.DataSource = Nothing

        mrsList.Sort = mshList.TextMatrix(0, lngCol) & IIf(mshList.ColData(lngCol) = 0, "", " DESC")
        mshList.ColData(lngCol) = (mshList.ColData(lngCol) + 1) Mod 2
        
        Call ShowBills(True)
    End If
End Sub

Private Function CalcTotal() As Currency
    Dim i As Long
    For i = 1 To mshList.Rows - 1
        If mshList.TextMatrix(i, 0) <> "" Then
            CalcTotal = CalcTotal + Val(mshList.TextMatrix(i, 6))
        End If
    Next
End Function

Private Function AuditingWarnByPatient(ByVal strNos As String) As Boolean
'功能：审核划价单时，对费用进行报警
'参数：str序号=指定单据中要审核的行号,为空表示所有行
    Dim rsWarn As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str类别s As String, cur金额 As Currency, cur余额 As Currency
    Dim strWarn As String, intWarn As Integer
    
    On Error GoTo errH
    
    '费用相关信息
    strSQL = _
        " Select A.收费类别,B.名称 as 类别名称,Sum(A.实收金额) as 金额" & _
        " From 门诊费用记录 A,收费项目类别 B" & _
        " Where A.记录性质=2 And A.门诊标志 In(1,3,4) And A.记录状态=0" & _
        "       And A.收费类别=B.编码 And A.划价人 is Not Null And A.操作员姓名 is NULL" & _
                IIf(strNos <> "", " And Instr(','||[2]||',',','||A.NO||',')>0", "") & _
        "       And A.病人ID=[1]" & _
        " Group by A.收费类别,B.名称"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mrsInfo!病人ID), strNos)
    
    For i = 1 To rsTmp.RecordCount
        If InStr(str类别s, rsTmp!收费类别 & rsTmp!类别名称) = 0 Then
            str类别s = str类别s & "," & rsTmp!收费类别 & rsTmp!类别名称
        End If
        cur金额 = cur金额 + rsTmp!金额
        rsTmp.MoveNext
    Next
    str类别s = Mid(str类别s, 2)
    
    If cur金额 > 0 Then
        '病人相关信息
        strSQL = "Select A.姓名,C.预交余额-C.费用余额 as 余额,zl_PatiDayCharge(A.病人ID) as 当日额," & _
            " Decode(A.担保额,null,A.担保额,Zl_Patientsurety(A.病人ID,Null)) 担保额,Zl_Patiwarnscheme(A.病人id, Null) As 适用病人" & _
            " From 病人信息 A,病人余额 C" & _
            " Where A.病人ID=C.病人ID(+) And C.性质(+)=1 And A.病人ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mrsInfo!病人ID))
        Set rsWarn = GetUnitWarn(rsTmp!适用病人, "0")
        
        cur余额 = Nvl(rsTmp!余额, 0)
        If gbln报警包含划价费用 Then cur余额 = Nvl(rsTmp!余额, 0) - GetPriceMoneyTotal(0, mrsInfo!病人ID) + cur金额
        '分类报警
        For i = 0 To UBound(Split(str类别s, ","))
            intWarn = BillingWarn(mstrPrivs, rsTmp!姓名, rsTmp!适用病人, rsWarn, _
                cur余额, Nvl(rsTmp!当日额, 0), cur金额, Nvl(rsTmp!担保额, 0), _
                Left(Split(str类别s, ",")(i), 1), Mid(Split(str类别s, ",")(i), 2), strWarn)
            If intWarn = 2 Or intWarn = 3 Then Exit Function
        Next
    End If
    AuditingWarnByPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub
Private Sub initCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建或关闭结算卡对象
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    Set gobjSquare.objDefaultCard = IDKind.GetfaultCard
    If IDKind.Cards.按缺省卡查找 And Not gobjSquare.objDefaultCard Is Nothing Then
        gobjSquare.bln缺省卡号密文 = gobjSquare.objDefaultCard.卡号密文规则 <> ""
        gobjSquare.int缺省卡号长度 = gobjSquare.objDefaultCard.卡号长度
    Else
        gobjSquare.bln缺省卡号密文 = IDKind.Cards.加密显示
        gobjSquare.int缺省卡号长度 = 100
    End If
    gobjSquare.bln按缺省卡查找 = IDKind.Cards.按缺省卡查找
End Sub

Private Sub CreateDrugPacker()
    '功能:创建自助发药机(自动化药房)
    Dim objComLib As New zl9ComLib.clsComLib
    Dim strPrivs As String
    Dim strMessage As String
    
    mblnDrugMachine = False

    Err = 0: On Error Resume Next
    If Val(zlDatabase.GetPara("启用药品自动化设备接口", glngSys, Val("9010-药品自动化设备接口"))) = 1 Then
        '优先新接口
        Set mobjDrugMachine = CreateObject("zlDrugMachine.clsDrugMachine")
        If Err = 0 Then mblnDrugMachine = True
    End If
    
    Err = 0: On Error GoTo 0
    If mblnDrugMachine Then
        '权限检查
        strPrivs = GetPrivFunc(glngSys, Val("9010-药品自动化设备接口"))
        If InStr(";" & strPrivs & ";", ";基本;") > 0 Then
            mblnDrugMachine = mobjDrugMachine.Init(1, objComLib, strMessage)
        Else
            mblnDrugMachine = False
        End If
    End If
End Sub

