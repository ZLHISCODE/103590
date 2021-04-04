VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frm应付款编辑 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "应付款编辑"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9555
   Icon            =   "frm应付款编辑.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox Picture1 
      Height          =   5250
      Left            =   30
      ScaleHeight     =   5190
      ScaleWidth      =   9420
      TabIndex        =   49
      Top             =   60
      Width           =   9480
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   915
         TabIndex        =   1
         Top             =   735
         Width           =   4200
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   13
         Left            =   975
         MaxLength       =   50
         TabIndex        =   38
         Tag             =   "摘要"
         Top             =   4095
         Width           =   8340
      End
      Begin VB.CommandButton cmdSelDept 
         Caption         =   "…"
         Height          =   300
         Left            =   5100
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   720
         Width           =   300
      End
      Begin VB.Frame fraTemp 
         Height          =   3045
         Left            =   75
         TabIndex        =   50
         Top             =   990
         Width           =   9255
         Begin VB.OptionButton optClass 
            Caption         =   "设备(&4)"
            Height          =   180
            Index           =   3
            Left            =   6120
            TabIndex        =   7
            Top             =   220
            Width           =   1000
         End
         Begin VB.OptionButton optClass 
            Caption         =   "物资(&3)"
            Height          =   180
            Index           =   2
            Left            =   5040
            TabIndex        =   6
            Top             =   220
            Width           =   1000
         End
         Begin VB.OptionButton optClass 
            Caption         =   "卫材(&2)"
            Height          =   180
            Index           =   1
            Left            =   3960
            TabIndex        =   5
            Top             =   220
            Width           =   1000
         End
         Begin VB.OptionButton optClass 
            Caption         =   "药品(&1)"
            Height          =   180
            Index           =   0
            Left            =   2880
            TabIndex        =   4
            Top             =   220
            Width           =   1000
         End
         Begin VB.CheckBox chkSelector 
            Caption         =   "选择器录入(&D)"
            Height          =   180
            Left            =   1110
            TabIndex        =   3
            Top             =   220
            Width           =   1500
         End
         Begin VB.CommandButton cmdSelName 
            Caption         =   "…"
            Height          =   300
            Left            =   8835
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   480
            Width           =   300
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   14
            Left            =   1110
            MaxLength       =   200
            TabIndex        =   18
            Tag             =   "随货单号"
            Top             =   1560
            Width           =   4875
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   9
            Left            =   7260
            MaxLength       =   16
            TabIndex        =   30
            Tag             =   "发票金额"
            Top             =   2250
            Width           =   1890
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   5
            Left            =   7260
            MaxLength       =   8
            TabIndex        =   20
            Tag             =   "入库单据号"
            Top             =   1560
            Width           =   1890
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   7
            Left            =   4110
            MaxLength       =   16
            TabIndex        =   24
            Tag             =   "单据金额"
            Top             =   1920
            Width           =   1875
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   1
            Left            =   1110
            MaxLength       =   50
            TabIndex        =   9
            Tag             =   "品名"
            Top             =   480
            Width           =   7710
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   2
            Left            =   1110
            MaxLength       =   50
            TabIndex        =   12
            Tag             =   "规格"
            Top             =   840
            Width           =   8025
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   3
            Left            =   1110
            MaxLength       =   50
            TabIndex        =   14
            Tag             =   "产地"
            Top             =   1200
            Width           =   4890
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   4
            Left            =   7260
            MaxLength       =   8
            TabIndex        =   16
            Tag             =   "计量单位"
            Top             =   1200
            Width           =   1890
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   10
            Left            =   1110
            MaxLength       =   16
            TabIndex        =   32
            Tag             =   "数量"
            Top             =   2640
            Width           =   1875
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   12
            Left            =   7260
            MaxLength       =   16
            TabIndex        =   36
            Tag             =   "采购金额"
            Top             =   2640
            Width           =   1890
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   8
            Left            =   1110
            MaxLength       =   200
            TabIndex        =   28
            Tag             =   "发票号"
            Top             =   2280
            Width           =   4875
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   6
            Left            =   1110
            MaxLength       =   20
            TabIndex        =   22
            Tag             =   "批号"
            Top             =   1920
            Width           =   1875
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.0000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
            Height          =   300
            Index           =   11
            Left            =   4110
            MaxLength       =   16
            TabIndex        =   34
            Tag             =   "采购价"
            Top             =   2640
            Width           =   1875
         End
         Begin MSComCtl2.DTPicker Dtp发票日期 
            Height          =   300
            Left            =   7260
            TabIndex        =   26
            Tag             =   "发票日期"
            Top             =   1920
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   314638336
            CurrentDate     =   37904
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "随货单号(&B)"
            Height          =   180
            Index           =   2
            Left            =   105
            TabIndex        =   17
            Top             =   1620
            Width           =   990
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "入库单号(&R)"
            Height          =   180
            Index           =   3
            Left            =   6195
            TabIndex        =   19
            Top             =   1620
            Width           =   990
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "单据金额(&J)"
            Height          =   180
            Index           =   4
            Left            =   3105
            TabIndex        =   23
            Top             =   1960
            Width           =   990
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "发票金额(&E)"
            Height          =   180
            Index           =   1
            Left            =   6195
            TabIndex        =   29
            Top             =   2310
            Width           =   990
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "发票日期(&F)"
            Height          =   180
            Index           =   16
            Left            =   6195
            TabIndex        =   25
            Top             =   1960
            Width           =   990
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "品名(&N)"
            Height          =   180
            Index           =   7
            Left            =   465
            TabIndex        =   8
            Top             =   530
            Width           =   630
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "规格(&G)"
            Height          =   180
            Index           =   8
            Left            =   465
            TabIndex        =   11
            Top             =   890
            Width           =   630
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "产地(&A)"
            Height          =   180
            Index           =   9
            Left            =   465
            TabIndex        =   13
            Top             =   1240
            Width           =   630
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "计量单位(&U)"
            Height          =   180
            Index           =   11
            Left            =   6195
            TabIndex        =   15
            Top             =   1240
            Width           =   990
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "数量(&S)"
            Height          =   180
            Index           =   12
            Left            =   465
            TabIndex        =   31
            Top             =   2700
            Width           =   630
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "采购金额(&I)"
            Height          =   180
            Index           =   14
            Left            =   6195
            TabIndex        =   35
            Top             =   2700
            Width           =   990
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "发票号(&K)"
            Height          =   180
            Index           =   0
            Left            =   285
            TabIndex        =   27
            Top             =   2340
            Width           =   810
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "批号(&P)"
            Height          =   180
            Index           =   10
            Left            =   465
            TabIndex        =   21
            Top             =   1960
            Width           =   630
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "采购价(&T)"
            Height          =   180
            Index           =   13
            Left            =   3300
            TabIndex        =   33
            Tag             =   "采购价"
            Top             =   2700
            Width           =   810
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
         Height          =   3645
         Left            =   705
         TabIndex        =   55
         Top             =   5115
         Visible         =   0   'False
         Width           =   6870
         _ExtentX        =   12118
         _ExtentY        =   6429
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   32768
         AllowBigSelection=   0   'False
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Txt填制人 
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   975
         TabIndex        =   40
         Top             =   4470
         Width           =   1875
      End
      Begin VB.Label Txt审核人 
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   975
         TabIndex        =   44
         Top             =   4830
         Width           =   1875
      End
      Begin VB.Label Lbl审核日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核日期"
         Height          =   180
         Left            =   6645
         TabIndex        =   45
         Top             =   4890
         Width           =   720
      End
      Begin VB.Label Lbl审核人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核人"
         Height          =   180
         Left            =   375
         TabIndex        =   43
         Top             =   4890
         Width           =   540
      End
      Begin VB.Label Lbl填制日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "填制日期"
         Height          =   180
         Left            =   6645
         TabIndex        =   41
         Top             =   4530
         Width           =   720
      End
      Begin VB.Label Lbl填制人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "填制人"
         Height          =   180
         Left            =   375
         TabIndex        =   39
         Top             =   4530
         Width           =   540
      End
      Begin VB.Label Txt填制日期 
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7425
         TabIndex        =   42
         Top             =   4470
         Width           =   1875
      End
      Begin VB.Label Txt审核日期 
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7425
         TabIndex        =   46
         Top             =   4830
         Width           =   1875
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "摘要(&W)"
         Height          =   180
         Index           =   5
         Left            =   315
         TabIndex        =   37
         Top             =   4155
         Width           =   630
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "应付记录单"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   90
         TabIndex        =   53
         Top             =   15
         Width           =   8850
      End
      Begin VB.Label LblNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NO."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   6945
         TabIndex        =   52
         Top             =   390
         Width           =   1140
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   8040
         TabIndex        =   51
         Top             =   345
         Width           =   1290
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "供应商(&M)"
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   0
         Top             =   810
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8325
      TabIndex        =   48
      Top             =   5520
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   7050
      TabIndex        =   47
      Top             =   5520
      Width           =   1100
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   54
      Top             =   5940
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm应付款编辑.frx":030A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12250
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
Attribute VB_Name = "frm应付款编辑"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mEditType As gEditType
Private mint记录状态 As Integer        '  RecBillStatus       '1:正常记录;2-冲销记录;3-已经冲销的原记录
Private mErrBillStatusInfor As ErrBillStatusInfor       '对于新增后单据并行执行的处理： 1、代表正常情况；2、已经删除的记录；3、已经审核的记录
Private mblnEdit As Boolean             '编辑状态
Private mblnSuccess As Boolean          '是否有单据保存成功
Private mstrPrivs  As String
Private mstrNo As String                   '单据号
Private mlng单位ID As Long
Private mint发票号Len As Integer          '数据库长度
Private mblnFirst As Boolean
Private mblnChange As Boolean
Private mblnSave As Boolean
Private mlngID As Long      '单据ID
Private mfrmMain  As Object
Private mstrSelectTag As String
Private mintPreCol As Integer
Private mintsort As Integer
Private Const mlngModule = 1322
'检查数据依赖性
Private Function GetDepend() As Boolean
    Dim strSQL As String
    Dim rsDepend As New Recordset

    GetDepend = False
    Dim str权限 As String
    str权限 = " and  " & Get分类权限(mstrPrivs)
    strSQL = "" & _
        "   SELECT  Id " & _
        "   FROM 供应商 " & _
        "   Where   末级=1 " & zl_获取站点限制 & "  " & str权限
    Err = 0
    On Error GoTo ErrHand:
    zlDatabase.OpenRecordset rsDepend, strSQL, Me.Caption
    If rsDepend.EOF Then
        ShowMsgbox "没有设置供应商或权限不足，请在供应商管理中设置！"
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Close
    GetDepend = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Sub ShowCard(FrmMain As Form, ByVal lngID As Long, _
    ByVal int编辑状态 As gEditType, ByVal strPrivs As String, Optional lng单位ID As Long = 0, _
    Optional int记录状态 As Integer = 1, _
    Optional blnSuccess As Boolean = False)
    
    mblnSave = False
    mblnSuccess = False
    mEditType = int编辑状态
    mint记录状态 = int记录状态
    mlngID = lngID
    mlng单位ID = lng单位ID
    mblnSuccess = blnSuccess
    mblnChange = False
    mErrBillStatusInfor = 1
    
    mstrPrivs = strPrivs
    
    Set mfrmMain = FrmMain
    
    '检查数据依赖关系
    If Not GetDepend Then Exit Sub
    
    If mEditType = g新增 Then
        mblnEdit = True
    ElseIf mEditType = g修改 Then
        mblnEdit = True
    ElseIf mEditType = g审核 Then
        mblnEdit = False
        cmdOK.Caption = "审核(&V)"
    ElseIf mEditType = g取消 Then
        mblnEdit = False
        cmdOK.Caption = "冲销(&O)"
    ElseIf mEditType = g查看 Then
        mblnEdit = False
        cmdOK.Caption = "打印(&P)"
        If InStr(mstrPrivs, "单据打印") = 0 Then
            cmdOK.Visible = False
        Else
            cmdOK.Visible = True
        End If
    End If
     
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    Me.Show vbModal, FrmMain
    blnSuccess = mblnSuccess
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim strSQL As String
    Dim rsInitCard As New Recordset
    
    On Error GoTo errHandle
    Select Case mEditType
        Case g新增
            InitControl
            Txt填制人 = gstrUserName
            Txt填制日期 = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
            Txt审核人 = ""
            Txt审核日期 = ""
            
            If mlng单位ID = 0 Then Exit Sub
            '确定供应商
            'by lesfeng 2009-12-2 性能优化
            Dim str权限 As String
            str权限 = " and  " & Get分类权限(mstrPrivs)
            
            strSQL = "Select ID,编码,名称,类型 From 供应商  where id=[1]" & str权限
            
            Set rsInitCard = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng单位ID)
            
            If rsInitCard.EOF Then Exit Sub
            
            txtEdit(0).Text = "[" & Nvl(rsInitCard!编码) & "]" & Nvl(rsInitCard!名称)
            chkSelector.Tag = Nvl(rsInitCard!类型)
            mlng单位ID = Nvl(rsInitCard!ID, 0)
        Case g审核, g修改, g查看, g取消
            InitControl
            'by lesfeng 2009-12-2 性能优化
            strSQL = "" & _
                  " SELECT A.ID,A.记录性质,A.记录状态,A.NO,A.项目ID,A.序号,A.收发ID,A.单位ID,A.品名,A.规格,A.产地,A.批号,A.计量单位," & _
                  "   A.入库单据号,A.单据金额,A.数量,A.采购价,A.采购金额,A.发票号,A.发票日期,A.发票金额,A.制定日期,A.计划金额," & _
                  "   A.计划人,A.计划日期,A.填制人,A.填制日期,A.审核人,A.审核日期,A.摘要,A.付款序号,A.计划序号,A.系统标识,A.随货单号," & _
                  "   b.编码 as 供应商编码,b.名称 as 供应商名称, b.类型 " & _
                  " From 应付记录 a, 供应商 b " & _
                  " Where a.单位ID = b.ID  And a.记录性质<>-1 and a.记录状态=[1] and a.ID=[2] Order by 计划序号"
            Set rsInitCard = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mint记录状态, mlngID)
            
            If rsInitCard.EOF Then
                If mEditType = g取消 Then
                    mErrBillStatusInfor = 已经冲销
                Else
                    mErrBillStatusInfor = 2
                End If
                Exit Sub
            End If
            
            txtNo.Caption = Nvl(rsInitCard!NO)
            mstrNo = txtNo
            
            If mEditType = g取消 Then
                Txt填制人 = gstrUserName
                Txt填制日期 = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
                Txt审核人 = gstrUserName
                Txt审核日期 = Txt填制日期
            Else
                Txt填制人 = IIf(IsNull(rsInitCard!填制人), "", rsInitCard!填制人)
                Txt填制日期 = Format(rsInitCard!填制日期, "yyyy-mm-dd hh:mm:ss")
                Txt审核人 = IIf(IsNull(rsInitCard!审核人), "", rsInitCard!审核人)
                Txt审核日期 = IIf(IsNull(rsInitCard!审核日期), "", Format(rsInitCard!审核日期, "yyyy-mm-dd hh:mm:ss"))
            End If
            If mEditType = g审核 Then
                Txt审核人 = gstrUserName
                Txt审核日期 = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
            End If
            If mEditType = g修改 Then
                Txt填制人 = gstrUserName
            End If
            txtEdit(13).Text = IIf(IsNull(rsInitCard!摘要), "", rsInitCard!摘要)
            
            If (mEditType = g修改 Or mEditType = g审核) And Nvl(rsInitCard!审核人) <> "" Then
                mErrBillStatusInfor = 3
                Exit Sub
            End If
            txtEdit(0).Text = "[" & Nvl(rsInitCard!供应商编码) & "]" & Nvl(rsInitCard!供应商名称)
            chkSelector.Tag = Nvl(rsInitCard!类型)
            mlng单位ID = Nvl(rsInitCard!单位ID, 0)
            Dim intIndex As Integer
            Dim strTmp As String
            With rsInitCard
                For intIndex = 1 To 14
                    strTmp = txtEdit(intIndex).Tag
                    If InStr(1, strTmp, "金额") <> 0 Then
                        txtEdit(intIndex).Text = Format(Nvl(.Fields(strTmp), 0), "####0.00;-####0.00; ;")
                    ElseIf strTmp = "采购价" Or strTmp = "数量" Then
                        txtEdit(intIndex).Text = Format(Nvl(.Fields(strTmp), 0), "####0.0000;-####0.0000; ;")
                    Else
                        txtEdit(intIndex).Text = Nvl(.Fields(strTmp))
                    End If
                Next
                If IsNull(!发票日期) Then
                    Dtp发票日期.Value = ""
                Else
                    Dtp发票日期.Value = Format(!发票日期, "yyyy-mm-dd")
                End If
            End With
            rsInitCard.Close
    End Select
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub InitControl()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:清除控件中的内容
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    For intIndex = 1 To 14
         txtEdit(intIndex).Text = ""
    Next
    Dtp发票日期.Value = ""
    txtNo = ""
End Sub

Private Sub chkSelector_Click()
    Call SetClass
End Sub

Private Sub chkSelector_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim blnSuccess As Boolean
    Dim strReg As String
    
    If mEditType = g查看 Then    '查看
        '打印
        printbill
        '退出
        Unload Me
        Exit Sub
    End If
    
    If mEditType = g审核 Then        '审核
        If SaveCheck = True Then
            If IIf(Val(zlDatabase.GetPara("审核打印", glngSys, mlngModule)) = 1, 1, 0) = 1 Then
                '打印
                If InStr(mstrPrivs, "单据打印") <> 0 Then
                    printbill
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If
    
    If ValidData = False Then Exit Sub
    
   If mEditType = g取消 Then
        If SaveStrike() = True Then
                Unload Me
        End If
        Exit Sub
    End If
    
    blnSuccess = SaveCard
        
    If blnSuccess = True Then
            
        If IIf(Val(zlDatabase.GetPara("存盘打印", glngSys, mlngModule)) = 1, 1, 0) = 1 Then
            '打印
            If InStr(mstrPrivs, "单据打印") <> 0 Then
                printbill
            End If
        End If
        If mEditType = g修改 Then    '修改
            Unload Me
            Exit Sub
        End If
        stbThis.Panels(2).Text = "上一张的单据号：" & mstrNo
    Else
        Exit Sub
    End If
    
    mblnSave = False
    mblnEdit = True
    
    InitControl
    If txtEdit(1).Enabled Then txtEdit(1).SetFocus
    mblnChange = False
End Sub

Private Sub cmdSelDept_Click()
    Dim rsTemp As ADODB.Recordset
    Dim strTemp As String
    strTemp = frm供应商选择.SelDept(mstrPrivs)
    If strTemp = "" Then
        Unload frm供应商选择
        If txtEdit(0).Enabled Then txtEdit(0).SetFocus
        Exit Sub
    End If
    txtEdit(0).Text = Mid(strTemp, InStr(strTemp, ",") + 1)
    mlng单位ID = Val(Left(strTemp, InStr(strTemp, ",") - 1))
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord("select 类型 from 供应商 where id=[1] ", Caption & "-提取供应商类型", mlng单位ID)
    If Not rsTemp.EOF Then
        chkSelector.Tag = Nvl(rsTemp!类型)
    End If
    rsTemp.Close
    Call SetClass
    
    If txtEdit(1).Enabled Then txtEdit(1).SetFocus
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdSelName_Click()
    Call GetItem("")
    txtEdit(2).SetFocus
End Sub

Private Sub Dtp发票日期_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
     
    Call initCard
    Call SetEditPro
    Call chkSelector_Click
    If txtEdit(1).Enabled Then txtEdit(1).SetFocus
    mblnChange = False
    setCtlEn
    Select Case mErrBillStatusInfor
        Case 1
            '正常
        Case 2
            '单据已被删除
            ShowMsgbox "该单据已被删除，请检查！"
            Unload Me
            Exit Sub
        Case 3
            '修改的单据已被审核
            ShowMsgbox "该单据已被其他人审核，请检查！"
            Unload Me
            Exit Sub
        Case 已经冲销
            '修改的单据已被审核
            ShowMsgbox "该单据没有可冲销的记录，请检查！"
            Unload Me
            Exit Sub
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    RestoreWinState Me, App.ProductName
    
    mint发票号Len = Sys.FieldsLength("应付记录", "发票号")      'Get发票号Len
    txtEdit(8).MaxLength = mint发票号Len
    mblnFirst = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim blnYes As Boolean
    If mblnChange = False Then Exit Sub
    ShowMsgbox "你已经更改了单据信息,你这样退出的话," & vbCrLf & "所更改的数据将不能保存,真的要退出吗?", True, blnYes
    If blnYes = True Then Exit Sub
    SaveWinState Me, App.ProductName
    Cancel = 1
End Sub

Private Sub mshSelect_LostFocus()
    mshSelect.Visible = False
End Sub

Private Sub optClass_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = 0 Then
        mlng单位ID = 0
    End If
    setCtlEn
End Sub

Private Sub setCtlEn()
    Dim intIndex As Integer
    If mEditType = g审核 Or mEditType = g取消 Or mEditType = g查看 Then
        Me.cmdOK.Enabled = True
    Else
        Me.cmdOK.Enabled = mblnChange
    End If
    chkSelector.Enabled = (mEditType = g新增 Or mEditType = g修改)
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    Dim strTmp As String
    Dim blnOpen As Boolean
    
    strTmp = txtEdit(Index).Tag
    If InStr(1, strTmp, "金额") <> 0 Or strTmp = "采购价" Or strTmp = "数量" Or InStr(1, strTmp, "号") <> 0 Then
            blnOpen = False
    Else
        blnOpen = True
    End If
    SetTxtGotFocus txtEdit(Index), blnOpen
End Sub

Private Function ValidData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:验证数据的合法性
    '--入参数:
    '--出参数:
    '--返  回:验证合法,返回True,否则=false
    '-----------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    Dim strTemp As String
    
   If mlng单位ID = 0 Then
        ShowMsgbox "供应商选择有误,请重新选择!"
        If txtEdit(0).Enabled Then txtEdit(0).SetFocus
        Exit Function
   End If
    
    For intIndex = 1 To 14
        strTemp = Trim(txtEdit(intIndex).Text)
        If intIndex = 1 Or txtEdit(intIndex).Tag = "发票金额" Then
            If strTemp = "" Then
                ShowMsgbox txtEdit(intIndex).Tag & "必需输入!"
                If txtEdit(intIndex).Enabled Then txtEdit(intIndex).SetFocus
                Exit Function
            End If
        End If
        
        If strTemp <> "" Then
            If LenB(StrConv(strTemp, vbFromUnicode)) > txtEdit(intIndex).MaxLength Then
                ShowMsgbox txtEdit(intIndex).Tag & "超长,最多能输入" & txtEdit(intIndex).MaxLength / 2 & "个汉字或" & txtEdit(intIndex).MaxLength & "个字符!"
                If txtEdit(intIndex).Enabled Then txtEdit(intIndex).SetFocus
                Exit Function
            End If
            If InStr(1, strTemp, "'") <> 0 Then
                ShowMsgbox txtEdit(intIndex).Tag & "不能输入单引号!"
                If txtEdit(intIndex).Enabled Then txtEdit(intIndex).SetFocus
                Exit Function
            End If
            If InStr(1, txtEdit(intIndex).Tag, "金额") <> 0 Or txtEdit(intIndex).Tag = "采购价" Or txtEdit(intIndex).Tag = "数量" Then
                If Not IsNumeric(strTemp) Then
                    ShowMsgbox txtEdit(intIndex).Tag & "不是数据型,请重输!"
                    If txtEdit(intIndex).Enabled Then txtEdit(intIndex).SetFocus
                    Exit Function
                End If
                If Val(strTemp) > 9999999999.99 Then
                    ShowMsgbox txtEdit(intIndex).Tag & "不能大于9999999999.99,请重输!"
                    If txtEdit(intIndex).Enabled Then txtEdit(intIndex).SetFocus
                    Exit Function
                End If
                If Val(strTemp) < -999999999.99 Then
                    ShowMsgbox txtEdit(intIndex).Tag & "不能小于-999999999.99,请重输!"
                    If txtEdit(intIndex).Enabled Then txtEdit(intIndex).SetFocus
                    Exit Function
                End If
            End If
        End If
    Next
    ValidData = True
End Function

Private Function SaveCard() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:保存卡片信息
    '--入参数:
    '--出参数:
    '--返  回:成功返回True,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim lngID As Long
    Dim NO_IN As String
    
    On Error GoTo errHandle
    SaveCard = False
    
    If mEditType = g新增 Then
        lngID = zlDatabase.GetNextId("应付记录")
        strSQL = "ZL_应付记录_INSERT("
        mstrNo = NextNo(67)
        NO_IN = mstrNo
    Else
        lngID = mlngID
        strSQL = "ZL_应付记录_UPDATE("
        NO_IN = Trim(txtNo)
    End If
    
    '过程参数如下:
    '  Id_In         In 应付记录.ID%Type,
    strSQL = strSQL & "" & lngID & ","
    '  No_In         In 应付记录.NO%Type,
    strSQL = strSQL & "'" & NO_IN & "',"
    '  收发id_In     In 应付记录.收发id%Type,
    strSQL = strSQL & "" & "Null" & ","
    '  单位id_In     In 应付记录.单位id%Type,
    strSQL = strSQL & "" & mlng单位ID & ","
    '  发票号_In     In 应付记录.发票号%Type,
    strSQL = strSQL & "" & IIf(Trim(txtEdit(8).Text) = "", "NULL", "'" & Trim(txtEdit(8).Text) & "'") & ","
    '  发票日期_In   In 应付记录.发票日期%Type,
    strSQL = strSQL & "" & IIf(Dtp发票日期.Value = "" Or IsNull(Dtp发票日期.Value), "NULL", "to_date('" & Format(Dtp发票日期.Value, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
    '  发票金额_In   In 应付记录.发票金额%Type,
    strSQL = strSQL & "" & Val(txtEdit(9).Text) & ","
    '  付款序号_In   In 应付记录.付款序号%Type,
    strSQL = strSQL & "" & "Null" & ","
    '  记录性质_In   In 应付记录.记录性质%Type,
    strSQL = strSQL & "" & "1" & ","
    '  入库单据号_In In 应付记录.入库单据号%Type,
    strSQL = strSQL & "" & IIf(Trim(txtEdit(5).Text) = "", "NULL", "'" & Trim(txtEdit(5).Text) & "'") & ","
    '  单据金额_In   In 应付记录.单据金额%Type,
    strSQL = strSQL & "" & Val(txtEdit(7).Text) & ","
    '  品名_In       In 应付记录.品名%Type,
    strSQL = strSQL & "" & IIf(Trim(txtEdit(1).Text) = "", "NULL", "'" & Trim(txtEdit(1).Text) & "'") & ","
    '  规格_In       In 应付记录.规格%Type,
    strSQL = strSQL & "" & IIf(Trim(txtEdit(2).Text) = "", "NULL", "'" & Trim(txtEdit(2).Text) & "'") & ","
    '  产地_In       In 应付记录.产地%Type,
    strSQL = strSQL & "" & IIf(Trim(txtEdit(3).Text) = "", "NULL", "'" & Trim(txtEdit(3).Text) & "'") & ","
    '  批号_In       In 应付记录.批号%Type,
    strSQL = strSQL & "" & IIf(Trim(txtEdit(6).Text) = "", "NULL", "'" & Trim(txtEdit(6).Text) & "'") & ","
    '  计量单位_In   In 应付记录.计量单位%Type,
    strSQL = strSQL & "" & IIf(Trim(txtEdit(4).Text) = "", "NULL", "'" & Trim(txtEdit(4).Text) & "'") & ","
    '  数量_In       In 应付记录.数量%Type,
    strSQL = strSQL & "" & Val(txtEdit(10).Text) & ","
    '  采购价_In     In 应付记录.采购价%Type,
    strSQL = strSQL & "" & Val(txtEdit(11).Text) & ","
    '  采购金额_In   In 应付记录.采购金额%Type,
    strSQL = strSQL & "" & Val(txtEdit(12).Text) & ","
    '  摘要_In       In 应付记录.摘要%Type,
    strSQL = strSQL & "" & IIf(Trim(txtEdit(13).Text) = "", "NULL", "'" & Trim(txtEdit(13).Text) & "'") & ","
    '  随货单号_In   In 应付记录.随货单号%Type := Null
    strSQL = strSQL & "" & IIf(Trim(txtEdit(14).Text) = "", "NULL", "'" & Trim(txtEdit(14).Text) & "'") & ")"
 
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    SaveCard = True
    mlngID = lngID
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SelMltProvide() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取供应商数据
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim strTmp As String
    Dim str权限 As String
    
    If Trim(txtEdit(0).Text) = "" Then Exit Function
    
    strTmp = GetMatchingSting(UCase(txtEdit(0).Text), False)
    
    str权限 = " and " & Get分类权限(mstrPrivs)
    
    SelMltProvide = False
    
    strSQL = "" & _
        "  Select   ID,编码,名称,简码,许可证号," & _
        "           to_char(许可证效期,'yyyy-mm-dd') as 许可证效期,执照号," & _
        "           to_char(执照效期,'yyyy-mm-dd') as 执照效期,税务登记号,联系人,类型 " & _
        "  From  供应商 " & _
        "  Where (撤档时间 is null or To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01') " & _
        "       " & zl_获取站点限制 & "  and 末级=1  " & _
        "       And ( 编码 Like [1] or 名称 like [1] or 简码  like upper([1])) " & str权限
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strTmp)
    
    If rsTemp.EOF Then
        ShowMsgbox "未找到指定的供应商!"
        Exit Function
    End If
    With rsTemp
        If .RecordCount > 1 Then
            mstrSelectTag = "Provide"
            Set mshSelect.Recordset = rsTemp
            With mshSelect
                .Top = txtEdit(0).Top + txtEdit(0).Height + 10
                .Left = txtEdit(0).Left
                .Visible = True
                .ColWidth(0) = 0
                .ColWidth(1) = 1400
                .ColWidth(2) = 2000
                .ColWidth(3) = 800
                .ColWidth(5) = 1000
                .ColWidth(6) = 1400
                .ColWidth(7) = 1000
                .ColWidth(8) = 1400
                .ColWidth(9) = 1000
                .ColWidth(10) = 0
                .Row = 1
                .Col = 0
                .ColSel = .Cols - 1
                .ZOrder
                .SetFocus
                Exit Function
            End With
        Else
            txtEdit(0).Text = "[" & Nvl(rsTemp!编码) & "]" & rsTemp!名称
            mlng单位ID = Nvl(rsTemp!ID, 0)
            SelMltProvide = True
        End If
    End With
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        If Index = 0 Then
            If SelMltProvide = False And mshSelect.Visible = False Then
                If txtEdit(0).Enabled Then txtEdit(0).SetFocus
            Else
                If mshSelect.Visible = False Then
                    zlCommFun.PressKey vbKeyTab
                End If
            End If
        ElseIf Index = 1 And chkSelector.Value = 1 Then
            If Trim(txtEdit(Index)) <> "" And txtEdit(Index).Tag <> txtEdit(Index).Text Then
                GetItem UCase(Trim(txtEdit(Index)))
            End If
            zlCommFun.PressKey vbKeyTab
        Else
            zlCommFun.PressKey vbKeyTab
        End If
    End If
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strTmp As String
    
    strTmp = txtEdit(Index).Tag
    
    If InStr(1, strTmp, "金额") <> 0 Or strTmp = "采购价" Or strTmp = "数量" Then
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m负金额式
    Else
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m文本式
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    Dim strTmp As String
    
    strTmp = txtEdit(Index).Tag
    
    If InStr(1, strTmp, "金额") <> 0 Then
        txtEdit(Index) = Format(Val(txtEdit(Index).Text), "####0.00;-####0.00; ;")
        If strTmp = "采购金额" Then
            txtEdit(11) = Format(Val(txtEdit(Index).Text) / IIf(Val(txtEdit(10)) = 0, 1, Val(txtEdit(10))), "####0.0000;-####0.0000; ;")
        End If
    ElseIf strTmp = "采购价" Then
        txtEdit(Index) = Format(Val(txtEdit(Index).Text), "####0.0000;-####0.0000; ;")
        txtEdit(12) = Format(Val(txtEdit(Index).Text) * Val(txtEdit(10).Text), "####0.00;-####0.00; ;")
    ElseIf strTmp = "数量" Then
        txtEdit(Index) = Format(Val(txtEdit(Index).Text), "####0.0000;-####0.0000; ;")
        txtEdit(12) = Format(Val(txtEdit(Index).Text) * Val(txtEdit(11).Text), "####0.00;-####0.00; ;")
    ElseIf strTmp = "入库单据号" Then
        Dim intYear  As Integer, strYear As String
        If IsNumeric(txtEdit(Index).Text) And txtEdit(Index).Text <> "" Then
            If Len(txtEdit(Index).Text) < 8 And Len(txtEdit(Index).Text) > 0 Then
                txtEdit(Index).Text = UCase(LTrim(txtEdit(Index).Text))
                intYear = Format(zlDatabase.Currentdate, "YYYY") - 1990
                strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
                txtEdit(Index).Text = strYear & String(7 - Len(txtEdit(Index).Text), "0") & txtEdit(Index).Text
            End If
        End If
    End If
    
    ImeLanguage False
End Sub

Private Function SaveCheck() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:审核单据
    '--入参数:
    '--出参数:
    '--返  回:成功,返回True,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    '   ZL_应付记录_Verify过程参数:
    '    ID_IN
    
    On Error GoTo errHandle:
    
    gstrSQL = "ZL_应付记录_Verify(" & _
        mlngID & ")"
    
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    SaveCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveStrike() As Boolean
 '-----------------------------------------------------------------------------------------------------------
    '--功  能:冲销单据
    '--入参数:
    '--出参数:
    '--返  回:成功,返回True,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    
    
    '   ZL_应付记录_STRIKE过程参数:
    '    ID_IN
    
    On Error GoTo errHandle:
    
    gstrSQL = "ZL_应付记录_STRIKE(" & _
        mlngID & ")"
    
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    SaveStrike = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub mshSelect_Click()
    With mshSelect
         If .Row < 1 Then Exit Sub
         If .MouseRow = 0 Then
            SetColumnSort mshSelect, mintPreCol, mintsort
            Exit Sub
         End If
    End With
End Sub

Private Sub mshSelect_DblClick()
    With mshSelect
        If .Row > 0 And .TextMatrix(.Row, 0) <> "" Then
            mshSelect_KeyPress 13
        End If
    End With
End Sub

Private Sub mshselect_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim sinWidth As Single
    
    With mshSelect
        Select Case KeyCode
            Case vbKeyRight
                If .ColPos(.Cols - 1) - .ColPos(.LeftCol) > .Width Then
                    .LeftCol = .LeftCol + 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                ElseIf .ColPos(.Cols - 1) - .ColPos(.LeftCol) + .ColWidth(.Cols - 1) > .Width Then
                    .LeftCol = .LeftCol + 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyLeft
                If .LeftCol <> 0 Then
                    .LeftCol = .LeftCol - 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyHome
                If .LeftCol <> 0 Then
                    .LeftCol = 0
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyEnd
                For i = .Cols - 1 To 0 Step -1
                    sinWidth = sinWidth + .ColWidth(i)
                    If sinWidth > .Width Then
                        .LeftCol = i + 1
                        .Col = .LeftCol
                        .ColSel = .Cols - 1
                        Exit For
                    End If
                Next
        End Select
    End With
End Sub

Private Sub mshSelect_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    With mshSelect
        Select Case mstrSelectTag
            Case "Provide"
                If KeyAscii = vbKeyReturn Then
                    If .Row = 0 Then Exit Sub
                    txtEdit(0).Text = "[" & .TextMatrix(.Row, 1) & "]" & .TextMatrix(.Row, 2)
                    
                    chkSelector.Tag = .TextMatrix(.Row, 10)
                    Call SetClass
                    
                    mlng单位ID = Val(.TextMatrix(.Row, 0))
                    If txtEdit(1).Enabled Then txtEdit(1).SetFocus
                ElseIf KeyAscii = 27 Then
                    If txtEdit(0).Enabled Then txtEdit(0).SetFocus
                End If
            Case Else
        End Select
        .Visible = False
    End With
End Sub

Private Sub SetEditPro()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:设置编辑属性
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    For intIndex = 0 To 14
        txtEdit(intIndex).Enabled = mblnEdit
    Next
    cmdSelDept.Enabled = mblnEdit
    Dtp发票日期.Enabled = mblnEdit
End Sub
'打印单据
Private Sub printbill()
    ReportOpen gcnOracle, glngSys, "ZL1_bill_1322", Me, "ID=" & mlngID
End Sub

Private Function GetClassValue() As Integer
    Dim i As Integer
    For i = 0 To optClass.Count - 1
        If optClass(i).Value And optClass(i).Enabled Then
            GetClassValue = i
            Exit Function
        End If
    Next
    GetClassValue = -1
End Function

Private Sub GetItem(ByVal strKey As String)
    Dim intClass As Integer
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim vRect As RECT
    Dim sngX As Single, sngY As Single, sngH As Single
    Dim intSysParam As Integer
    Dim strMatch As String
    
    intClass = GetClassValue()
    vRect = zlControl.GetControlRect(txtEdit(1).hwnd)
    sngX = vRect.Left
    sngY = vRect.Bottom
    
    On Error GoTo errHandle
    Select Case intClass
    Case 0
        '药品
        If strKey = "" Then
            strSQL = "Select ID, 上级id, 编码, 名称, '' 规格, '' 产地, '' 药库单位, '' 住院单位, '' 门诊单位, 0 As 末级 " & _
                     "From 诊疗分类目录 " & vbLf & _
                     "Where 类型 in ('1','2','3') " & vbLf & _
                     "Start With 上级id Is Null Connect By Prior ID = 上级id " & vbLf & _
                     "Union all " & vbLf & _
                     "Select a.Id, c.分类id As 上级id, a.编码, a.名称, a.规格, a.产地, b.药库单位, b.住院单位, b.门诊单位, 1 As 末级 " & vbLf & _
                     "From 收费项目目录 A, 药品规格 B, 诊疗项目目录 C " & vbLf & _
                     "Where a.Id = b.药品id And b.药名id = c.Id And a.类别 in ('5','6','7') " & vbLf & _
                     "  And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) "
            
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, Caption & "-药品" _
                    , False, "", "选择", False, False, False, sngX, sngY, sngH, blnCancel, False, False)
        Else
            strSQL = "Select Distinct a.ID, null 上级ID, a.编码, a.名称, a.规格, a.产地, b.药库单位, b.住院单位, b.门诊单位 " & vbLf & _
                     "From 收费项目目录 A, 药品规格 B, 收费项目别名 C " & vbLf & _
                     "Where a.Id = b.药品id And a.id = c.收费细目id And A.类别 in ('5','6','7') " & vbLf & _
                     "  And (to_char(A.撤档时间, 'yyyy-mm-dd') = '3000-01-01' or A.撤档时间 is null) " & _
                     "  And C.性质 = 1 "
            intSysParam = Val(zlDatabase.GetPara("简码方式"))
            strMatch = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = "0", "%", "")
            
            If IsNumeric(strKey) Then
                strSQL = strSQL & " And (a.编码 Like [1] Or C.简码 Like [2] And C.码类=3) "
            ElseIf zlCommFun.IsCharAlpha(strKey) Then
                strSQL = strSQL & " And C.简码 Like [2] and c.码类=" & IIf(intSysParam = 0, 1, 2) & " "
            ElseIf zlCommFun.IsCharChinese(strKey) Then
                strSQL = strSQL & " And C.名称 Like [2] "
            Else
                strSQL = strSQL & " And (a.编码 = [1] And C.名称 Like [2] Or C.简码 LIKE [2]) and c.码类=" & IIf(intSysParam = 0, 1, 2) & " "
            End If
            strSQL = strSQL & vbNewLine & "Order by a.编码 "
            
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, Caption & "-药品" _
                    , False, "", "选择", False, False, True, sngX, sngY, sngH, blnCancel, False, False _
                    , strKey & "%" _
                    , strMatch & strKey & "%")
        End If
        
    Case 1
        '卫材
        If strKey = "" Then
            strSQL = "Select ID, 上级id, 编码, 名称, '' 规格, '' 产地, '' As 计算单位, 0 As 末级 " & _
                     "From 诊疗分类目录 " & vbLf & _
                     "Where 类型 = '7' " & vbLf & _
                     "Start With 上级id Is Null Connect By Prior ID = 上级id " & vbLf & _
                     "Union all " & vbLf & _
                     "Select i.Id, b.分类id As 上级id, i.编码, i.名称, i.规格, i.产地, i.计算单位, 1 As 末级 " & vbLf & _
                     "From 收费项目目录 I, 材料特性 T, 诊疗项目目录 B " & vbLf & _
                     "Where i.Id = t.材料id And t.诊疗id = b.Id And i.类别 = '4' " & vbLf & _
                     "  And (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) "
            
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, Caption & "-卫材" _
                    , False, "", "选择", False, False, False, sngX, sngY, sngH, blnCancel, False, False)
        Else
            strSQL = "Select Distinct i.Id, i.编码, i.名称, i.规格, i.产地, i.计算单位, 1 As 末级 " & vbLf & _
                     "From 收费项目目录 I, 材料特性 T, 收费项目别名 B " & vbLf & _
                     "Where i.Id = t.材料id And i.Id = b.收费细目id And i.类别 = '4' " & vbLf & _
                     "  And (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) "
            intSysParam = Val(zlDatabase.GetPara("简码方式"))
            strMatch = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = "0", "%", "")
            
            If IsNumeric(strKey) Then
                strSQL = strSQL & " And (i.编码 Like [1] Or b.简码 Like [2] And b.码类=3) "
            ElseIf zlCommFun.IsCharAlpha(strKey) Then
                strSQL = strSQL & " And b.简码 Like [2] And b.码类 = [3] "
            ElseIf zlCommFun.IsCharChinese(strKey) Then
                strSQL = strSQL & " And b.名称 Like [2] "
            Else
                strSQL = strSQL & " And (i.编码 = [1] And b.名称 Like [2] Or b.简码 LIKE [2]) And b.码类 = [3] "
            End If
            strSQL = strSQL & vbLf & "Order by i.编码 "
        
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, Caption & "-卫材" _
                    , False, "", "选择", False, False, True, sngX, sngY, sngH, blnCancel, False, False _
                    , strKey & "%" _
                    , strMatch & strKey & "%" _
                    , IIf(intSysParam = 0, 1, 2))
        End If
    Case 2
        '物资
        If strKey = "" Then
            strSQL = "Select ID, 0 末级, 上级id, 编码, 名称, '' 规格, '' 产地, '' 散装单位, '' 包装单位 " & _
                     "From 物资分类 " & _
                     "Where 物资类别 in ('普通物资', '医用物资') " & _
                     "Start With 上级id Is Null Connect By Prior ID = 上级id " & _
                     "Union All " & _
                     "Select ID, 1 末级, 分类id 上级id, 编码, 名称, 规格, 产地, 散装单位, 包装单位 " & _
                     "From 物资目录 " & _
                     "Where (to_char(撤档时间,'yyyy-MM-DD') = '3000-01-01' or 撤档时间 is null) And 物资类别 in ('普通物资', '医用物资') "
            
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, Caption & "-物资" _
                    , False, "", "选择", False, False, False, sngX, sngY, sngH, blnCancel, False, False)
        Else
            strSQL = "Select ID, 编码, 名称, 规格, 产地, 散装单位, 包装单位 " & _
                     "From 物资目录 " & _
                     "Where (to_char(撤档时间,'yyyy-MM-DD') = '3000-01-01' or 撤档时间 is null) And 物资类别 in ('普通物资', '医用物资') "
            
            If IsNumeric(strKey) Then
                strSQL = strSQL & " And (编码 Like [1] Or 简码 Like [2]) "
            ElseIf zlCommFun.IsCharAlpha(strKey) Then
                strSQL = strSQL & " And 简码 Like [2] "
            ElseIf zlCommFun.IsCharChinese(strKey) Then
                strSQL = strSQL & " And 名称 Like [2] "
            Else
                strSQL = strSQL & " And (编码 = [1] And 名称 Like [2] Or 简码 LIKE [2]) "
            End If
            strSQL = strSQL & vbLf & "Order by 编码 "
        
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, Caption & "-物资" _
                    , False, "", "选择", False, False, True, sngX, sngY, sngH, blnCancel, False, False _
                    , strKey & "%" _
                    , "%" & strKey & "%")
        End If
    Case 3
        '设备
        If strKey = "" Then
            strSQL = "Select ID, 0 末级, 上级id, 编码, 名称, '' 规格, '' 产地, '' 单位 " & _
                     "From 设备分类 " & _
                     "Start With 上级id Is Null Connect By Prior ID = 上级id " & _
                     "Union All " & _
                     "Select ID, 1 末级, 分类id 上级id, 编码, 名称, 规格, 产地, 单位 " & _
                     "From 设备目录 " & _
                     "Where (to_char(撤档时间,'yyyy-MM-DD') = '3000-01-01' or 撤档时间 is null) "
            
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, Caption & "-设备" _
                    , False, "", "选择", False, False, False, sngX, sngY, sngH, blnCancel, False, False)
        Else
            strSQL = "Select ID, 编码, 名称, 规格, 产地, 单位 " & _
                     "From 设备目录 " & _
                     "Where (to_char(撤档时间,'yyyy-MM-DD') = '3000-01-01' or 撤档时间 is null) "
            
            If IsNumeric(strKey) Then
                strSQL = strSQL & " And (编码 Like [1] Or 简码 Like [2]) "
            ElseIf zlCommFun.IsCharAlpha(strKey) Then
                strSQL = strSQL & " And 简码 Like [2] "
            ElseIf zlCommFun.IsCharChinese(strKey) Then
                strSQL = strSQL & " And 名称 Like [2] "
            Else
                strSQL = strSQL & " And (编码 = [1] And 名称 Like [2] Or 简码 LIKE [2]) "
            End If
            strSQL = strSQL & vbLf & "Order by 编码 "
        
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, Caption & "-设备" _
                    , False, "", "选择", False, False, True, sngX, sngY, sngH, blnCancel, False, False _
                    , strKey & "%" _
                    , "%" & strKey & "%")
        End If
    End Select
    
    If blnCancel = False And Not rsTemp Is Nothing Then
        txtEdit(1).Text = Nvl(rsTemp!名称)
        txtEdit(1).Tag = Nvl(rsTemp!名称)
        txtEdit(2).Text = Nvl(rsTemp!规格)
        txtEdit(3).Text = Nvl(rsTemp!产地)
        If intClass = 1 Then
            txtEdit(4).Text = Nvl(rsTemp!计算单位)
        ElseIf intClass = 3 Then
            txtEdit(4).Text = Nvl(rsTemp!单位)
        End If
    End If
    If Not rsTemp Is Nothing Then rsTemp.Close
    Exit Sub
    
errHandle:
    Call ErrCenter
End Sub

Private Sub SetClass()
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    For i = 0 To optClass.Count - 1
        '系统
        If i >= 2 Then
            Set rsTemp = zlDatabase.OpenSQLRecord("Select Count(1) Rec From zlSystems Where 编号 = [1]", Caption, IIf(i = 2, 400, 600))
            optClass(i).Enabled = rsTemp!rec > 0
            rsTemp.Close
        Else
            optClass(i).Enabled = True
        End If
        '权限
        Select Case i
            Case 0
                optClass(i).Enabled = optClass(i).Enabled And InStr(mstrPrivs, ";药品;") > 0
            Case 1
                optClass(i).Enabled = optClass(i).Enabled And InStr(mstrPrivs, ";卫材;") > 0
            Case 2
                optClass(i).Enabled = optClass(i).Enabled And InStr(mstrPrivs, ";物资;") > 0
            Case 3
                optClass(i).Enabled = optClass(i).Enabled And InStr(mstrPrivs, ";设备;") > 0
        End Select
        optClass(i).Visible = chkSelector.Value = 1
    Next
    '供应商
    If Len(chkSelector.Tag) >= 1 Then  '药品
        optClass(0).Enabled = optClass(0).Enabled And Mid(chkSelector.Tag, 1, 1) = "1"
    Else
        optClass(0).Enabled = False
    End If
    If Len(chkSelector.Tag) >= 5 Then  '卫材
        optClass(1).Enabled = optClass(1).Enabled And Mid(chkSelector.Tag, 5, 1) = "1"
    Else
        optClass(1).Enabled = False
    End If
    If Len(chkSelector.Tag) >= 2 Then  '物资
        optClass(2).Enabled = optClass(2).Enabled And Mid(chkSelector.Tag, 2, 1) = "1"
    Else
        optClass(2).Enabled = False
    End If
    If Len(chkSelector.Tag) >= 3 Then  '设备
        optClass(3).Enabled = optClass(3).Enabled And Mid(chkSelector.Tag, 3, 1) = "1"
    Else
        optClass(3).Enabled = False
    End If
    
    cmdSelName.Visible = chkSelector.Value = 1
    If chkSelector.Value = 1 Then
        txtEdit(1).Width = txtEdit(2).Width - cmdSelName.Width
    Else
        txtEdit(1).Width = txtEdit(2).Width
    End If
    Exit Sub
    
errHandle:
    Call ErrCenter
End Sub
