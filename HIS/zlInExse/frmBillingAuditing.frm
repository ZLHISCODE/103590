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
   ClientWidth     =   9930
   Icon            =   "frmBillingAuditing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   9930
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
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   3780
      Width           =   8400
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshBill 
      Height          =   2265
      Left            =   30
      TabIndex        =   10
      ToolTipText     =   "双击单据查看明细"
      Top             =   3825
      Width           =   9810
      _ExtentX        =   17304
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
      ScaleWidth      =   9930
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   6120
      Width           =   9930
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   300
         TabIndex        =   18
         Top             =   525
         Width           =   1100
      End
      Begin VB.CommandButton cmdFlash 
         Caption         =   "刷新(&R)"
         Height          =   350
         Left            =   300
         TabIndex        =   17
         Top             =   90
         Width           =   1100
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "选择(&S)"
         Height          =   350
         Left            =   1695
         TabIndex        =   13
         Top             =   90
         Width           =   1100
      End
      Begin VB.CommandButton cmdCls 
         Caption         =   "清除(&M)"
         Height          =   350
         Left            =   2880
         TabIndex        =   14
         Top             =   90
         Width           =   1100
      End
      Begin VB.CommandButton cmdClsAll 
         Caption         =   "全清(&C)"
         Height          =   350
         Left            =   2880
         TabIndex        =   16
         Top             =   525
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelAll 
         Caption         =   "全选(&A)"
         Height          =   350
         Left            =   1695
         TabIndex        =   15
         Top             =   525
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "退出(&X)"
         Height          =   350
         Left            =   7140
         TabIndex        =   12
         Top             =   525
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "审核(&O)"
         Height          =   350
         Left            =   7140
         TabIndex        =   11
         Top             =   90
         Width           =   1100
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
      ScaleWidth      =   9855
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1110
      Width           =   9855
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "未审核划价单,当前合计:"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   120
         TabIndex        =   31
         Tag             =   "未审核划价单,当前合计:"
         Top             =   30
         Width           =   1980
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   1125
      Left            =   45
      TabIndex        =   19
      Top             =   -45
      Width           =   9795
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   315
         Left            =   450
         TabIndex        =   34
         Top             =   195
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   556
         Appearance      =   2
         IDKindStr       =   $"frmBillingAuditing.frx":08A4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "宋体"
         IDKind          =   -1
         ShowPropertySet =   -1  'True
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         BackColor       =   -2147483633
      End
      Begin VB.TextBox txt剩余 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   6870
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   705
         Width           =   1125
      End
      Begin VB.TextBox txt费用 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3975
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   705
         Width           =   1125
      End
      Begin VB.TextBox txt预交 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   705
         Width           =   1125
      End
      Begin VB.TextBox txt科室 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   8340
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   195
         Width           =   1320
      End
      Begin VB.TextBox txt床号 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   6870
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   195
         Width           =   720
      End
      Begin VB.TextBox txt住院号 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   5385
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   195
         Width           =   840
      End
      Begin VB.TextBox txt年龄 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3990
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   195
         Width           =   480
      End
      Begin VB.TextBox txt性别 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   195
         Width           =   495
      End
      Begin VB.TextBox txtPatient 
         BackColor       =   &H00EBFFFF&
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   1110
         MaxLength       =   100
         TabIndex        =   0
         ToolTipText     =   "热键:F6"
         Top             =   195
         Width           =   1155
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
      Begin VB.Label lbl科室 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "科室"
         Height          =   180
         Left            =   7875
         TabIndex        =   28
         Top             =   255
         Width           =   360
      End
      Begin VB.Label lbl剩余款 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "剩余款"
         Height          =   180
         Left            =   6255
         TabIndex        =   27
         Top             =   765
         Width           =   540
      End
      Begin VB.Label lbl未结费用 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "未结费用"
         Height          =   180
         Left            =   3165
         TabIndex        =   26
         Top             =   765
         Width           =   720
      End
      Begin VB.Label lbl预交余额 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预交余额"
         Height          =   180
         Left            =   330
         TabIndex        =   25
         Top             =   765
         Width           =   720
      End
      Begin VB.Label lbl床号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床号"
         Height          =   180
         Left            =   6420
         TabIndex        =   24
         Top             =   255
         Width           =   360
      End
      Begin VB.Label lbl住院号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         Height          =   180
         Left            =   4785
         TabIndex        =   23
         Top             =   255
         Width           =   540
      End
      Begin VB.Label lbl年龄 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Height          =   180
         Left            =   3540
         TabIndex        =   22
         Top             =   255
         Width           =   360
      End
      Begin VB.Label lbl性别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   180
         Left            =   2385
         TabIndex        =   21
         Top             =   255
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人"
         ForeColor       =   &H80000007&
         Height          =   180
         Left            =   90
         TabIndex        =   20
         Top             =   255
         Width           =   360
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   2415
      Left            =   30
      TabIndex        =   9
      ToolTipText     =   "双击单据查看明细"
      Top             =   1365
      Width           =   9795
      _ExtentX        =   17277
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
      MouseIcon       =   "frmBillingAuditing.frx":093A
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   33
      Top             =   7080
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBillingAuditing.frx":0C54
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12912
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
Private mlngModule As String
Private mlngUnitID As Long '当前所选择的病区ID
Private mstrUnitIDs As String   '当前操作员的所有病区ID
Private mstrPrivs As String
Private mstrPrivsOpt As String '记帐操作1150模块的授权功能
Private mrsInfo As New ADODB.Recordset
Private mrsList As ADODB.Recordset
Attribute mrsList.VB_VarHelpID = -1
Private mlngCurRow As Long, mlngTopRow As Long
Private mobjICCard As Object
Private mintSucces As Integer
'-----------------------------------------------------------------------------------
'结算卡相关
Private mstrPassWord As String
'-----------------------------------------------------------------------------------
Public Function zlCardShow(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, lngUnitID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:审核程序入口
    '入参:lngUnitID-当前所选择的病区ID
    '出参:
    '返回:审核一次成功以上,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-03 17:30:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mintSucces = 0: mlngModule = lngModule: mstrPrivs = strPrivs: mlngUnitID = lngUnitID
    Me.Show 1, frmMain
    zlCardShow = mintSucces > 0
End Function
 

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
    Dim strDel As String, i As Long, str审核时间 As String, Curdate As Date
    Dim arrSQL As Variant, strNos As String, strNO As String, blnTrans As Boolean
    
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State <> 1 Then Exit Sub
    If mrsInfo.EOF Then Exit Sub
    If zlIsAllowFeeChange(Val(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!主页ID))) = False Then
         Exit Sub
    End If
    
    arrSQL = Array()
    For i = 1 To mshList.Rows - 1
        If mshList.TextMatrix(i, 0) <> "" And mshList.TextMatrix(i, 1) <> "" Then
            If str审核时间 = "" Then
                Curdate = zlDatabase.Currentdate
                str审核时间 = "To_Date('" & Format(Curdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
            End If
            strNO = mshList.TextMatrix(i, 1)
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_住院记帐记录_Verify('" & strNO & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "',NULL," & mrsInfo!病人ID & "," & str审核时间 & ")"
            strDel = strDel & "," & i
            
            strNos = strNos & "," & strNO
        End If
    Next
    If UBound(arrSQL) = -1 Then
        MsgBox "没有选择要审核的划价单据！", vbInformation, gstrSysName
        Exit Sub
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
            strNO = Split(strNos, ",")(i)
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133", Me, "NO=" & strNO, "登记时间=" & Format(Curdate, "yyyy-MM-dd HH:mm:ss"), "药品单位=" & IIf(gbln住院单位, 1, 0), "PrintEmpty=0", "重打=0", 2)
        Next
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
    gblnOK = True
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
        Case vbKeyF6
            txtPatient.SetFocus
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    Me.Height = 7815 '没有写Resize
    gblnOK = False
    mstrPrivsOpt = ";" & GetInsidePrivs(Enum_Inside_Program.p记帐操作) & ";"
        
    Call SetHeader
    Call SetBill
    Call initCardSquareData
    mstrUnitIDs = GetUserUnits
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlngUnitID = 0
    Call SaveWinState(Me, App.ProductName)
End Sub


Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    Set gobjSquare.objCurCard = objCard
    If txtPatient.Text <> "" Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

 
Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.卡号
    Call FindPati(objCard, True, txtPatient.Text)
End Sub

Private Sub mshList_DblClick()
    Dim strNO As String
    
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
        Err.Clear
        On Error Resume Next
            
        strNO = mshList.TextMatrix(mshList.Row, 1)
        If BillisBatch(strNO) Then '批量记帐
            frmBillings.mstrPrivs = mstrPrivs
            frmBillings.mbytInState = 1
            frmBillings.mstrInNO = strNO
            frmBillings.Show 1, Me
        ElseIf BillisSimple(strNO) Then '简单记帐
            frmSimpleBilling.mstrPrivs = mstrPrivs
            frmSimpleBilling.mbytInState = 1
            frmSimpleBilling.mstrInNO = strNO
            frmSimpleBilling.Show 1, Me
        Else '记帐单
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 1
            frmCharge.mstrInNO = strNO
            frmCharge.Show 1, Me
        End If
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

Private Sub ShowDetail(Optional strNO As String)
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim strSQL As String
    
    On Error GoTo errH
    
    '明细费用行的剩余数量和金额
    strSQL = _
    " Select C.名称 as 类别,Nvl(E.名称,B.名称) as 名称" & IIf(gTy_System_Para.byt药品名称显示 = 2, ",E1.名称 as 商品名", "") & ",B.规格," & _
            IIf(gbln住院单位, "Decode(X.药品ID,NULL,A.计算单位,X.住院单位)", "A.计算单位") & " as 单位," & _
    "       Avg(Nvl(A.付数,1)*A.数次)" & IIf(gbln住院单位, "/Nvl(X.住院包装,1)", "") & " as 数量, " & _
    "       Ltrim(To_Char(Sum(A.标准单价)" & IIf(gbln住院单位, "*Nvl(X.住院包装,1)", "") & ",'99999" & gstrFeePrecisionFmt & "')) as 单价," & _
    "       Ltrim(To_Char(Sum(A.应收金额),'99999" & gstrDec & "')) as 应收金额," & _
    "       Ltrim(To_Char(Sum(A.实收金额),'99999" & gstrDec & "')) as 实收金额," & _
    "       D.名称 as 执行科室" & _
    " From 住院费用记录 A,收费项目目录 B,收费项目类别 C,部门表 D,收费项目别名 E,药品规格 X" & _
        IIf(gTy_System_Para.byt药品名称显示 = 2, ",收费项目别名 E1", "") & _
    " Where A.收费细目ID=B.ID and A.收费类别=C.编码 And A.执行部门ID=D.ID(+)" & _
    "       And A.NO=[1] And A.记录性质=2 And A.门诊标志=2 And A.记录状态=0" & _
    "       And A.病人ID+0=[2] And Nvl(A.主页ID,0)=[3]" & _
    "       And A.收费细目ID=X.药品ID(+) And A.操作员姓名 is NULL And A.划价人 is Not NULL" & _
    "       And A.收费细目ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
    IIf(gTy_System_Para.byt药品名称显示 = 2, "       And A.收费细目ID=E1.收费细目ID(+) And E.码类(+)=1 And E1.性质(+)=3", "") & _
    " Group by Nvl(A.价格父号,A.序号),C.名称," & _
    "       Nvl(E.名称,B.名称)" & IIf(gTy_System_Para.byt药品名称显示 = 2, ",E1.名称", "") & ",B.规格,A.计算单位,D.名称,X.药品ID,X.住院单位,Nvl(X.住院包装,1)" & _
    " Order by Nvl(A.价格父号,A.序号)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, CLng(mrsInfo!病人ID), Val("" & mrsInfo!主页ID))
    
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
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshBill, App.ProductName & "\" & Me.Name)
        
        For i = 0 To .Cols - 1
            If .TextMatrix(0, i) = "商品名" Then
                If gTy_System_Para.byt药品名称显示 = 2 Then
                    If .ColWidth(i) <= 0 Then .ColWidth(i) = 2000
                Else
                    .ColWidth(i) = 0
                End If
            End If
        Next
        
        .RowHeight(0) = 320
        .Col = 0: .ColSel = .Cols - 1
    End With
End Sub

Private Sub mshList_KeyPress(KeyAscii As Integer)
    Dim strNO As String
    
    If mshList.TextMatrix(mshList.Row, 1) = "" Then Exit Sub
    
    If KeyAscii = 32 Then
        If mshList.TextMatrix(mshList.Row, 0) = "" Then
            mshList.TextMatrix(mshList.Row, 0) = "√"
        Else
            mshList.TextMatrix(mshList.Row, 0) = ""
        End If
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        Err.Clear
        On Error Resume Next
            
        strNO = mshList.TextMatrix(mshList.Row, 1)
        If BillisBatch(strNO) Then '批量记帐
            frmBillings.mstrPrivs = mstrPrivs
            frmBillings.mbytInState = 1
            frmBillings.mstrInNO = strNO
            frmBillings.Show 1, Me
        ElseIf BillisSimple(strNO) Then '简单记帐
            frmSimpleBilling.mstrPrivs = mstrPrivs
            frmSimpleBilling.mbytInState = 1
            frmSimpleBilling.mstrInNO = strNO
            frmSimpleBilling.Show 1, Me
        Else '记帐单
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 1
            frmCharge.mstrInNO = strNO
            frmCharge.Show 1, Me
        End If
    End If
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
    If txtPatient.Locked Then Exit Sub
    Call IDKIND.SetAutoReadCard(txtPatient.Text = "")
End Sub

Private Sub txtPatient_GotFocus()
    zlControl.TxtSelAll txtPatient
    If txtPatient.Locked Then Exit Sub
    Call IDKIND.SetAutoReadCard(txtPatient.Text = "")
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
'    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
'    If IDKind.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13 Then
        With frmPatiSelect
            If InStr(mstrPrivs, ";所有病区;") > 0 Then
                .mlngUnitID = 0
            Else
                .mlngUnitID = mlngUnitID
            End If
            Set .mfrmParent = Me
            .mstrPrivs = mstrPrivs
            .Show 1, Me
        End With
    Else
        If IDKIND.GetCurCard.名称 Like "姓名*" Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKIND.ShowPassText)
        ElseIf IDKIND.IDKIND = IDKIND.GetKindIndex("门诊号") Or IDKIND.IDKIND = IDKIND.GetKindIndex("住院号") Then
            If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
                If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
            End If
            txtPatient.PasswordChar = IIf(IDKIND.ShowPassText, "*", "")
        End If
    End If
    Me.Refresh
    If blnCard And Len(txtPatient.Text) = IDKIND.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        Call FindPati(IDKIND.GetCurCard, blnCard, txtPatient.Text)
    End If
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查找病人
    '编制:刘兴洪
    '日期:2012-08-29 17:53:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnICCard As Boolean, blnIDCard As Boolean
    Dim strErrMsg As String
   '读取病人信息
    Call ClearPati
    mshList.Clear: mshList.Rows = 2
    Call SetHeader
    If Not GetPatient(objCard, txtPatient.Text, blnCard, strErrMsg) Then
        txtPatient.Text = ""
        If blnCard Then
            If strErrMsg <> "" Then
                sta.Panels(2) = strErrMsg
            Else
                sta.Panels(2) = "不能确定病人信息，请检查是否正确刷卡或选择的病人不是住院病人！"
            End If
            txtPatient.SetFocus: Exit Sub
        Else
            If strErrMsg <> "" Then
                sta.Panels(2) = strErrMsg
            Else
                sta.Panels(2) = "输入的标识不能读取病人信息，请检查输入是否正确或选择的病人不是住院病人！"
            End If
            txtPatient.SetFocus: Exit Sub
        End If
        Exit Sub
    End If
    
    '54899
    If objCard.名称 Like "IC卡*" And objCard.系统 = True And mstrPassWord <> "" Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If objCard.名称 Like "*身份证*" And objCard.系统 = True And mstrPassWord <> "" Then blnIDCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
     If (objCard.名称 Like "IC卡*" Or objCard.名称 Like "*身份证*") And objCard.系统 = True And blnCard Then blnCard = False
    '就诊卡密码检查
    If Mid(gstrCardPass, 6, 1) = "1" And (blnCard Or blnICCard Or blnIDCard) Then
        If Not zlCommFun.VerifyPassWord(Me, mstrPassWord, mrsInfo!姓名, mrsInfo!性别, "" & mrsInfo!年龄) Then
            Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
        End If
    End If
    
    If zlIsAllowFeeChange(Val(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!主页ID)), Val(Nvl(mrsInfo!审核标志))) = False Then
        Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
    End If
    
    txtPatient.PasswordChar = ""
    txtPatient.Text = IIf(IsNull(mrsInfo!姓名), "", mrsInfo!姓名)
    txt性别.Text = IIf(IsNull(mrsInfo!性别), "", mrsInfo!性别)
    txt年龄.Text = IIf(IsNull(mrsInfo!年龄), "", mrsInfo!年龄)
    txt床号.Text = IIf(IsNull(mrsInfo!床号), "家庭病床", mrsInfo!床号)
    txt住院号.Text = IIf(IsNull(mrsInfo!住院号), "", mrsInfo!住院号)
    txt科室.Text = GET部门名称(mrsInfo!科室ID)
    
    txtPatient.ForeColor = zlDatabase.GetPatiColor(Nvl(mrsInfo!病人类型))
    
    Call RefreshMoney
    Call ShowBills
    mshList.SetFocus
End Sub


Private Sub RefreshMoney()
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = GetMoneyInfo(mrsInfo!病人ID, , , 2)
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
    txt床号.Text = ""
    txt住院号.Text = ""
    txt科室.Text = ""
    txt费用.Text = ""
    txt预交.Text = ""
    txt剩余.Text = ""
End Sub

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, ByVal blnCard As Boolean, Optional ByRef strOut As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息
    '入参:blnCard=是否就诊卡刷卡
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2011-08-03 17:34:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String, bln所有病区 As Boolean
    Dim strIF As String, strWhere As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    Dim rsOutSel As ADODB.Recordset
    
    On Error GoTo errH
        
    'a.是否具有强制记帐权限
    If InStr(mstrPrivsOpt, ";出院未结强制记帐;") > 0 And InStr(mstrPrivsOpt, ";出院结清强制记帐;") > 0 Then
        strIF = ""
    ElseIf InStr(mstrPrivsOpt, ";出院未结强制记帐;") > 0 Then
        strIF = " And ((B.出院日期 is NULL And Nvl(B.状态,0)<>3) Or Nvl(X.费用余额,0)<>0)"
    ElseIf InStr(mstrPrivsOpt, ";出院结清强制记帐;") > 0 Then
        strIF = " And ((B.出院日期 is NULL And Nvl(B.状态,0)<>3) Or Nvl(X.费用余额,0)=0)"
    Else
        strIF = " And B.出院日期 is NULL And Nvl(B.状态,0)<>3"
    End If
    
    'b.是否可以记所有病区病人
    bln所有病区 = True
    If InStr(mstrPrivs, ";所有病区;") <= 0 Then
        bln所有病区 = False
        If InStr(1, mstrUnitIDs, ",") = 0 Then
            strIF = strIF & " And B.当前病区ID+0=[3]"
        Else
            strIF = strIF & " And B.当前病区ID+0 IN(Select Column_Value From Table(Cast(f_num2list([4]) As zlTools.t_numlist)))"
        End If
    End If
       
    'c.是否留观病人记帐权限
    If (InStr(mstrPrivsOpt, ";门诊留观记帐;") > 0 And gbln门诊留观) And (InStr(mstrPrivsOpt, ";住院留观记帐;") > 0 And gbln住院留观) Then
        strIF = strIF & " And Nvl(B.病人性质,0) IN(0,1,2)"
    ElseIf InStr(mstrPrivsOpt, ";门诊留观记帐;") > 0 And gbln门诊留观 Then
        strIF = strIF & " And Nvl(B.病人性质,0) IN(0,1)"
    ElseIf InStr(mstrPrivsOpt, ";住院留观记帐;") > 0 And gbln住院留观 Then
        strIF = strIF & " And Nvl(B.病人性质,0) IN(0,2)"
    Else
        strIF = strIF & " And Nvl(B.病人性质,0)=0"
    End If
    
    strSQL = _
            "Select A.病人ID,B.主页ID,B.当前病区ID as 病区ID,B.出院科室ID as 科室ID,B.入院日期,B.出院日期," & _
            "   A.就诊卡号,A.卡验证码,A.住院号,B.出院病床 as 床号,X.费用余额,B.状态," & _
            "   nvl(B.姓名,A.姓名) as 姓名,nvl(b.性别,A.性别) as 性别,A.年龄,B.费别,B.住院医师,B.医疗付款方式," & _
            "   A.担保人,Decode(A.担保额,null,A.担保额,Zl_Patientsurety(A.病人ID,B.主页ID)) 担保额," & _
            "   zl_PatiDayCharge(A.病人ID) as 当日额,B.险类,Nvl(B.病人性质,0) as 病人性质,B.病人类型,b.审核标志" & _
            " From 病人信息 A,病案主页 B,病人余额 X " & _
            " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID" & _
            "       And Nvl(B.主页ID,0)<>0 And A.病人ID=X.病人ID(+) And X.性质(+)=1 And X.类型(+)=2 And A.停用时间 is NULL " & strIF
    If blnCard = True And objCard.名称 Like "姓名*" Then    '刷卡
        If IDKIND.Cards.按缺省卡查找 And Not IDKIND.GetfaultCard Is Nothing Then
            lng卡类别ID = IDKIND.GetfaultCard.接口序号
        Else
            lng卡类别ID = "-1"
        End If
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        strWhere = strWhere & " And A.病人ID=[1] "
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '病人ID
        strWhere = strWhere & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "/" Then   '床位号
        '41654 And IsNumeric(Mid(strInput, 2))
        strInput = Mid(strInput, 2)
        If mlngUnitID = 0 Then '病区不确定、则不能通过床号确定病人
            Set mrsInfo = New ADODB.Recordset: Exit Function
        End If
        strSQL = _
            "Select A.病人ID,B.主页ID,B.当前病区ID as 病区ID,B.出院科室ID as 科室ID,B.入院日期,B.出院日期," & _
            "   A.就诊卡号,A.卡验证码,A.住院号,B.出院病床 as 床号,X.费用余额,B.状态," & _
            "   nvl(B.姓名,A.姓名) as 姓名,nvl(b.性别,A.性别) as 性别,A.年龄,B.费别,B.住院医师,B.医疗付款方式," & _
            "   A.担保人,Decode(A.担保额,null,A.担保额,Zl_Patientsurety(A.病人ID,B.主页ID)) 担保额," & _
            "   zl_PatiDayCharge(A.病人ID) as 当日额,B.险类,Nvl(B.病人性质,0) as 病人性质,B.病人类型,B.审核标志" & _
            " From 病人信息 A,病案主页 B,床位状况记录 C,病人余额 X" & _
            " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID" & _
            " And Nvl(B.主页ID,0)<>0 And A.病人ID=C.病人ID And A.病人ID=X.病人ID(+) And X.性质(+)=1 And X.类型(+)=2 And A.停用时间 is NULL " & _
            " And C.病区ID=[3] And C.床号=[2] " & strIF
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号(病人在院)
        strWhere = strWhere & " And A.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [1])"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号(医技记帐)
        strWhere = strWhere & " And A.门诊号=[1]"
    Else
        Select Case objCard.名称
            Case "姓名", "姓名或就诊卡"
                If mrsInfo.State = 1 Then
                    If mrsInfo.EOF = False Then
                        If mrsInfo!姓名 = Trim(txtPatient.Text) Then GetPatient = True: Exit Function
                    End If
                End If
                If zlSelectChargePatiFromInputName(Me, mstrPrivsOpt, strInput, bln所有病区, mstrUnitIDs, gintOutDay, lng病人ID, strOut, txtPatient.hWnd, txtPatient.Height) = False Then
                     Set mrsInfo = New Recordset: Exit Function
                End If
                strInput = "-" & lng病人ID
                strWhere = strWhere & " And A.病人ID=[1]"
            Case "医保号"
                strInput = UCase(strInput)
                strWhere = strWhere & " And A.医保号=[2]"
             Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.门诊号=[2]"
            Case "住院号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [2])"
            Case Else
                '其他类别的,获取相关的病人IDs
                If objCard.接口序号 > 0 Then
                    lng卡类别ID = objCard.接口序号
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng病人ID = 0 Then GoTo NotFoundPati:
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng病人ID <= 0 Then GoTo NotFoundPati:
                strWhere = strWhere & " And A.病人ID=[1]"
                strInput = "-" & lng病人ID
                blnHavePassWord = True
        End Select
    End If
       
    strSQL = strSQL & vbCrLf & strWhere
    txtPatient.ForeColor = Me.ForeColor
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput, mlngUnitID, mstrUnitIDs)
    
    If mrsInfo.RecordCount = 0 Then GoTo NotFoundPati:
    mstrPassWord = strPassWord
    If Not blnHavePassWord Then
        mstrPassWord = Nvl(mrsInfo!卡验证码)
    End If
    GetPatient = True
    Exit Function
NotFoundPati:
    Set mrsInfo = New ADODB.Recordset
    If strWhere = "" Then Exit Function
    
    '未找到病人，需要对该病人的具体错误信息进行提示
    strSQL = _
    " Select A.病人ID,B.主页ID,B.当前病区ID as 病区ID,B.出院科室ID as 科室ID,a.在院,B.入院日期,B.出院日期,X.费用余额,B.状态, " & _
    "       nvl(B.姓名,A.姓名) as 姓名,nvl(b.性别,A.性别) as 性别,nvl(b.年龄,A.年龄) as 年龄,B.费别,Nvl(B.病人性质,0) as 病人性质,B.病人类型" & _
    " From 病人信息 A,病案主页 B,病人余额 X" & _
    " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID" & _
    "   And Nvl(B.主页ID,0)<>0 And A.病人ID=X.病人ID(+) and X.性质(+)=1 and X.类型(+)=2 And A.停用时间 is NULL " & strWhere
    
    Set rsOutSel = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    If rsOutSel.EOF Then Exit Function
    '1.病区检查
    If InStr(mstrPrivs, ";所有病区;") <= 0 Then
        If InStr(1, "," & mstrUnitIDs & ",", "," & Val(rsOutSel!病区ID) & ",") = 0 Then
            strOut = "病人:『" & Nvl(rsOutSel!姓名) & "』不在你负责的病区,不能对该病人进行记账操作!"
            Exit Function
        End If
    End If
    
    '2.留观病人检查(是否留观病人记帐权限)
    If (InStr(mstrPrivsOpt, ";门诊留观记帐;") > 0 And gbln门诊留观) And (InStr(mstrPrivsOpt, ";住院留观记帐;") > 0 And gbln住院留观) Then
        '0-普通住院病人,1-门诊留观病人,2-住院留观病人
    ElseIf InStr(mstrPrivsOpt, ";门诊留观记帐;") > 0 And gbln门诊留观 Then
        If Val(Nvl(rsOutSel!病人性质)) = 2 Then
            strOut = "病人:『" & Nvl(rsOutSel!姓名) & "』为住院留观病人,你不具备『住院留观记帐』权限,不能对该病人进行记账操作!"
            Exit Function
        End If
    ElseIf InStr(mstrPrivsOpt, ";住院留观记帐;") > 0 And gbln住院留观 Then
        If Val(Nvl(rsOutSel!病人性质)) = 1 Then
            strOut = "病人:『" & Nvl(rsOutSel!姓名) & "』为门诊留观病人,你不具备『门诊留观记帐』权限,不能对该病人进行记账操作!"
            Exit Function
        End If
    Else
        If Val(Nvl(rsOutSel!病人性质)) <> 0 Then
            strOut = "病人:『" & Nvl(rsOutSel!姓名) & "』为" & IIf(Val(Nvl(rsOutSel!病人性质)) = 1, "门诊", "住院") & "留观病人,你不具备『门诊或住院 留观记帐』权限,不能对该病人进行记账操作!"
            Exit Function
        End If
    End If
    
        '124007
    If InStr(mstrPrivsOpt, ";出院未结强制记帐;") > 0 And InStr(mstrPrivsOpt, ";出院结清强制记帐;") > 0 Then
        strErrMsg = ""
    ElseIf InStr(mstrPrivsOpt, ";出院未结强制记帐;") > 0 Then
        If Not (Val(Nvl(rsOutSel!状态)) <> 3 And IsNull(rsOutSel!出院日期) Or Val(Nvl(rsOutSel!费用余额)) <> 0) Then
              
                If Val(Nvl(rsOutSel!状态)) = 3 And IsNull(rsOutSel!出院日期) Then
                    strErrMsg = "病人已经预出院，不能对病人进行记账操作!"
                Else
                    strErrMsg = "病人于" & Format(rsOutSel!出院日期, "yyyy年mm月DD日") & " 出院，不能对病人进行记账操作!"
                End If
        End If
    ElseIf InStr(mstrPrivsOpt, ";出院结清强制记帐;") > 0 Then
        If Not (Val(Nvl(rsOutSel!状态)) <> 3 And IsNull(rsOutSel!出院日期) Or Val(Nvl(rsOutSel!费用余额)) = 0) Then
                If Val(Nvl(rsOutSel!状态)) = 3 And IsNull(rsOutSel!出院日期) Then
                strErrMsg = "病人已经预出院，不能对病人进行记账操作!"
                Else
                strErrMsg = "病人于" & Format(rsOutSel!出院日期, "yyyy年mm月DD日") & " 出院，不能对病人进行记账操作!"
                End If
        End If
    Else
        If Not (Val(Nvl(rsOutSel!状态)) <> 3 And IsNull(rsOutSel!出院日期)) Then
            If Val(Nvl(rsOutSel!状态)) = 3 And IsNull(rsOutSel!出院日期) Then
                strErrMsg = "病人已经预出院，不能对病人进行记账操作!"
            Else
                strErrMsg = "病人于" & Format(rsOutSel!出院日期, "yyyy年mm月DD日") & " 出院，不能对病人进行记账操作!"
            End If
        End If
    End If
    If strErrMsg <> "" Then
        strOut = strErrMsg
        Exit Function
    End If
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
        " From 住院费用记录 A,部门表 B" & _
        " Where A.记录性质=2 And A.门诊标志=2 And A.记录状态=0" & _
        "       And A.划价人 is Not Null And A.操作员姓名 is NULL" & _
        "       And A.开单部门ID=B.ID" & _
        "       And A.病人ID=[1] And Nvl(A.主页ID,0)=[2]" & _
        " Group by A.NO,B.名称,A.开单人,A.费别,A.登记时间,A.划价人" & _
        " Order by 划价时间 Desc,单据号 Desc"
        Set mrsList = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsInfo!病人ID), Val("" & mrsInfo!主页ID))
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
        .Cols = UBound(Split(strHead, "|")) + 1
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
        
        i = MshGetColNum(mshList, "医生")
        'If InStr(mstrPrivsOpt, "医生查询") = 0 Then .ColWidth(i) = 0
        
        .Col = 0: .ColSel = .Cols - 1
                
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
        If mshList.TextMatrix(0, lngCol) = "审核" Then
           mshList.Col = lngCol
            If mshList.ColData(lngCol) = 1 Then
                mshList.Sort = flexSortStringNoCaseAscending
            Else
               mshList.Sort = flexSortStringNoCaseDescending
            End If
            mshList.ColData(lngCol) = (mshList.ColData(lngCol) + 1) Mod 2
            Exit Sub
        End If
        
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
        " From 住院费用记录 A,收费项目类别 B" & _
        " Where A.记录性质=2 And A.门诊标志=2 And A.记录状态=0" & _
        " And A.收费类别=B.编码 And A.划价人 is Not Null And A.操作员姓名 is NULL" & _
        IIf(strNos <> "", " And Instr(','||[3]||',',','||A.NO||',')>0", "") & _
        " And A.病人ID=[1] And Nvl(A.主页ID,0)=[2]" & _
        " Group by A.收费类别,B.名称"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mrsInfo!病人ID), Val("" & mrsInfo!主页ID), strNos)
    
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
        strSQL = "Select B.当前病区ID 病区ID,A.住院号,A.当前床号 As 床号,nvl(B.姓名,A.姓名) as 姓名,C.预交余额-C.费用余额 as 余额,zl_PatiDayCharge(A.病人ID) as 当日额," & _
            " Decode(A.担保额,null,A.担保额,Zl_Patientsurety(A.病人ID,B.主页ID)) 担保额,Zl_Patiwarnscheme(B.病人id, B.主页id) As 适用病人" & _
            " From 病人信息 A,病案主页 B,病人余额 C" & _
            " Where A.病人ID=B.病人ID(+) And Nvl(A.主页ID,0)=B.主页ID(+)" & _
            " And A.病人ID=C.病人ID(+) And C.性质(+)=1 And C.类型(+)=2 " & _
            " And A.病人ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mrsInfo!病人ID))
        Set rsWarn = GetUnitWarn(rsTmp!适用病人, Val(Nvl(rsTmp!病区ID)))    '问题:43862
        
        cur余额 = Nvl(rsTmp!余额, 0)
        If gbln报警包含划价费用 Then cur余额 = Nvl(rsTmp!余额, 0) - GetPriceMoneyTotal(1, mrsInfo!病人ID) + cur金额
        '分类报警
        For i = 0 To UBound(Split(str类别s, ","))
            intWarn = BillingWarn(mstrPrivsOpt, rsTmp!姓名 & IIf(Nvl(rsTmp!住院号) = "", "", "(住院号:" & rsTmp!住院号 & " 床号:" & rsTmp!床号 & ")"), Val("" & rsTmp!病区ID), rsTmp!适用病人, rsWarn, _
                cur余额, Nvl(rsTmp!当日额, 0), cur金额, Nvl(rsTmp!担保额, 0), _
                Left(Split(str类别s, ",")(i), 1), Mid(Split(str类别s, ",")(i), 2), strWarn)
            If intWarn = 2 Or intWarn = 3 Then Exit Function
        Next
    End If
    AuditingWarnByPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub initCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取结算卡对象的相关信息
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    Call IDKIND.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    Set objCard = IDKIND.GetfaultCard
    Set gobjSquare.objDefaultCard = objCard
    If IDKIND.Cards.按缺省卡查找 And Not objCard Is Nothing Then
        gobjSquare.bln缺省卡号密文 = objCard.卡号密文规则 <> ""
        gobjSquare.int缺省卡号长度 = objCard.卡号长度
    Else
        gobjSquare.bln缺省卡号密文 = IDKIND.Cards.加密显示
        gobjSquare.int缺省卡号长度 = 100
    End If
    gobjSquare.bln按缺省卡查找 = IDKIND.Cards.按缺省卡查找
End Sub
Private Sub txtPatient_LostFocus()
    Call IDKIND.SetAutoReadCard(False)
End Sub
