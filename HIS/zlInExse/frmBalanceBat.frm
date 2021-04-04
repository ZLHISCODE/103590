VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmBalanceBat 
   Caption         =   "批量中途结帐"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBalanceBat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   11820
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picDown 
      BorderStyle     =   0  'None
      Height          =   1320
      Left            =   -15
      ScaleHeight     =   1320
      ScaleWidth      =   11730
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6540
      Width           =   11730
      Begin VB.CommandButton cmdOK 
         Caption         =   "结帐(&O)"
         Default         =   -1  'True
         Height          =   400
         Left            =   8685
         TabIndex        =   13
         Top             =   825
         Width           =   1400
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "退出(&C)"
         Height          =   400
         Left            =   10200
         TabIndex        =   14
         Top             =   825
         Width           =   1400
      End
      Begin VB.ComboBox cbo结算方式 
         Height          =   360
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   0
         Width           =   2055
      End
      Begin VB.TextBox txtInvoice 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   9675
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   1905
      End
      Begin MSMask.MaskEdBox txtDateEnd 
         Height          =   360
         Left            =   375
         TabIndex        =   7
         Top             =   0
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-mm-dd hh:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VSFlex8Ctl.VSFlexGrid vsBalance 
         Height          =   435
         Left            =   120
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   825
         Visible         =   0   'False
         Width           =   8460
         _cx             =   14922
         _cy             =   767
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483633
         ForeColor       =   -2147483640
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483633
         GridColor       =   12632256
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   3
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   0
         Cols            =   8
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   360
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBalanceBat.frx":617A
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   3
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblDateEnd 
         Caption         =   "对                     之前的费用结帐"
         Height          =   255
         Left            =   90
         TabIndex        =   6
         Top             =   60
         Width           =   4440
      End
      Begin VB.Label lbl结算方式 
         Caption         =   "结算方式"
         Height          =   255
         Left            =   5040
         TabIndex        =   8
         Top             =   60
         Width           =   975
      End
      Begin VB.Label lblFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票据号"
         Height          =   240
         Left            =   8880
         TabIndex        =   10
         Top             =   60
         Width           =   720
      End
      Begin VB.Label lblInfo 
         Caption         =   "共完成n个病人结帐"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   90
         TabIndex        =   12
         Top             =   480
         Width           =   8295
      End
   End
   Begin VB.Frame fra 
      Height          =   645
      Left            =   90
      TabIndex        =   15
      Top             =   0
      Width           =   11685
      Begin VB.ComboBox cboInsure 
         Height          =   360
         Left            =   4095
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   195
         Width           =   3885
      End
      Begin VB.ComboBox cbo使用类别 
         Height          =   360
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   1815
      End
      Begin VB.Label lblInsure 
         AutoSize        =   -1  'True
         Caption         =   "保险类别"
         Height          =   240
         Left            =   3135
         TabIndex        =   17
         Top             =   270
         Width           =   960
      End
      Begin VB.Label lblRpt 
         AutoSize        =   -1  'True
         Caption         =   "sss"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   8085
         TabIndex        =   2
         Top             =   300
         Width           =   405
      End
      Begin VB.Label lbl使用类别 
         Caption         =   "使用类别"
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   255
         Width           =   960
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsDept 
      Height          =   4860
      Left            =   2160
      TabIndex        =   4
      Top             =   675
      Width           =   2460
      _cx             =   4339
      _cy             =   8572
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   13627390
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBalanceBat.frx":6245
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPati 
      Height          =   4860
      Left            =   4680
      TabIndex        =   5
      Top             =   690
      Width           =   7065
      _cx             =   12462
      _cy             =   8572
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   12640511
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBalanceBat.frx":628D
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vsFeeType 
      Height          =   4875
      Left            =   120
      TabIndex        =   3
      Top             =   675
      Width           =   1980
      _cx             =   3492
      _cy             =   8599
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   15790320
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBalanceBat.frx":63DF
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmBalanceBat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPatis As String '用于记录选择的科室下标记为不结帐的病人ID
Private mlng领用ID As Long
Private mrsRptFormat As ADODB.Recordset
Private mobjInvoice As clsInvoice
Private mobjFact As clsFactProperty
Private mblnNotClick As Boolean
Private mlngPreInsure As Long
Private mlngModul As Long
Private mstrPrivs As String '权限串
'当前病人险类的医保支持参数
Private Type TYPE_MedicarePAR
    分币处理 As Boolean
    未结清出院 As Boolean
    结算使用个人帐户 As Boolean
    出院结算必须出院 As Boolean
    出院病人结算作废 As Boolean
    中途结算仅处理已上传部分 As Boolean
    结帐设置后调用接口 As Boolean
    结帐作废后打印回单 As Boolean
    住院结算作废 As Boolean
    批量中途结帐 As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR

'3.3 模块参数定义
Private Type Ty_ModulePara
     blnZero  As Boolean '结帐时是否处理零费用
     int费用时间 As Integer '0-按登记时间,1-按发生时间
End Type
Private mty_ModulePara As Ty_ModulePara
Private mblnOK As Boolean

Public Function ShowMe(ByVal frmMain As Object, ByVal strPrivs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口(批量中途结帐)
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-07 09:52:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mblnOK = False
    mstrPrivs = strPrivs
    Me.Show IIf(gfrmMain Is Nothing, 0, 1), frmMain
    ShowMe = mblnOK
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub zlInitModulePara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化模块参数
    '编制:刘兴洪
    '日期:2010-02-04 16:50:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With mty_ModulePara
         .blnZero = zlDatabase.GetPara("处理零费用", glngSys, mlngModul) = "1"
         .int费用时间 = IIf(zlDatabase.GetPara("结帐费用时间", glngSys, mlngModul) = "1", 1, 0)
    End With
End Sub
Private Sub LoadInsureType()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载保险类别
    '编制:刘兴洪
    '日期:2015-03-25 14:26:17
    '问题:81661
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSelect As String, strSql As String
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    
    If zlStr.IsHavePrivs(mstrPrivs, "医保批量中途结帐") = False Then
        cboInsure.Visible = False: lblInsure.Visible = False
        lblRpt.Left = lblInsure.Left
        Exit Sub
    End If
    If Not gclsInsure Is Nothing Then
        strSelect = gclsInsure.GetAvailabilityInsures
    End If
    mblnNotClick = True
    cboInsure.Clear
    cboInsure.AddItem ""
    cboInsure.ItemData(cboInsure.NewIndex) = 0
    cboInsure.ListIndex = 0
    If InStr(strSelect, ",") = 0 And Val(strSelect) = 0 Then Exit Sub
    strSql = "" & _
    "   Select A.序号,A.名称,A.说明,Nvl(A.外挂,0) AS 外挂" & _
    "   From 保险类别 A " & _
    "   Where Nvl(是否禁止,0)=0 " & _
    "       And A.序号 in (Select Column_value From Table(f_Num2List([1])))" & _
    "   Order By A.序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strSelect)
    With rsTemp
        Do While Not .EOF
            cboInsure.AddItem "" & rsTemp!序号 & "-" & rsTemp!名称
            cboInsure.ItemData(cboInsure.NewIndex) = Val(NVL(rsTemp!序号))
            .MoveNext
        Loop
    End With
    mblnNotClick = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    mblnNotClick = False
End Sub
Private Sub cboInsure_Click()
    Dim lngDeptID As Long
    Dim lngInsure As Long
    If mblnNotClick Then Exit Sub
    lngInsure = cboInsure.ItemData(cboInsure.ListIndex)
    If mlngPreInsure = lngInsure Then Exit Sub  '相同时,不改变
    mlngPreInsure = lngInsure
    If vsDept.Row > 0 Then lngDeptID = Val(vsDept.RowData(vsDept.Row))
    vsDept.Cell(flexcpChecked, 1, 0, vsDept.Rows - 1, 0) = flexChecked
    mstrPatis = ""
    Call LoadPati(lngDeptID)
End Sub

Private Sub cbo使用类别_Click()
    Dim lngDeptID As Long
    If mblnNotClick Then Exit Sub
    
    lblRpt.Caption = ""
    If mrsRptFormat Is Nothing Then Exit Sub
    mrsRptFormat.Filter = "序号=" & cbo使用类别.ItemData(cbo使用类别.ListIndex)
    If Not mrsRptFormat.EOF Then
        lblRpt.Caption = NVL(mrsRptFormat!说明)
    End If
    mlng领用ID = 0
    Call InitFact(cbo使用类别.Text)
    Call RefreshFact
    If vsDept.Row > 0 Then lngDeptID = Val(vsDept.RowData(vsDept.Row))
    vsDept.Cell(flexcpChecked, 1, 0, vsDept.Rows - 1, 0) = flexChecked
    mstrPatis = ""
    Call LoadPati(lngDeptID)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, m As Long, blnPrint As Boolean
    Dim rsPati As ADODB.Recordset
    
    For i = 1 To vsDept.Rows - 1
        If vsDept.Cell(flexcpChecked, i, 0) = flexUnchecked Or vsDept.Cell(flexcpChecked, i, 0) = flexTSUnchecked Then
            m = m + 1
        End If
    Next
    If m = vsDept.Rows - 1 Then
        MsgBox "请至少选择一个科室.", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Set rsPati = GetPatiSet
    If rsPati.RecordCount = 0 Then
        MsgBox "请至少选择一个病人.", vbInformation, gstrSysName
        Exit Sub
    End If
    If Not IsDate(txtDateEnd.Text) Then
        MsgBox "费用截止时间格式不正确.", vbInformation, gstrSysName
        txtDateEnd.SetFocus
        Exit Sub
    End If
    
    blnPrint = mobjFact.打印方式 <> 0
    If mobjFact.打印方式 = 2 Then
        If MsgBox("是否打印票据?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then blnPrint = False
    End If
    
    If blnPrint Then
        If mobjFact.严格控制 Then    '严格票据管理
            If Trim(txtInvoice.Text) = "" Then
                MsgBox "必须输入一个有效的票据号码！", vbInformation, gstrSysName
                txtInvoice.SetFocus: Exit Sub
            End If
            If mobjInvoice.zlGetInvoiceGroupID(1137, UserInfo.用户名, mobjFact.票种, mobjFact.使用类别, mlng领用ID, mobjFact.共享批次ID, mlng领用ID, 1, Trim(txtInvoice.Text)) = False Then Exit Sub
            If mlng领用ID <= 0 Then
                Select Case mlng领用ID
                    Case 0 '操作失败
                    Case -1
                        MsgBox "你没有自用和共用的结帐票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                    Case -2
                        MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                    Case -3
                        MsgBox "当前票据号码不在可用领用批次的有效票据号范围内,请重新输入", vbInformation, gstrSysName
                        txtInvoice.SetFocus
                End Select
                Exit Sub
            End If
        Else
            If Len(txtInvoice.Text) <> mobjFact.票号长度 And txtInvoice.Text <> "" Then
                MsgBox "票据号码长度应该为 " & gbytFactLength & " 位！", vbInformation, gstrSysName
                txtInvoice.SetFocus: Exit Sub
            End If
        End If
    End If
    
    If MsgBox("共选择了" & rsPati.RecordCount & "位病人,即将依次进行中途结帐!" & _
        vbCrLf & "请准备好后按确定.", vbInformation + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
        Exit Sub
    End If
    
    cmdOK.Enabled = False
    Screen.MousePointer = 11
    Call SaveBalance(blnPrint, rsPati)
    Call LoadPati(Val(vsDept.RowData(vsDept.Row)))
    Screen.MousePointer = 0
    cmdOK.Enabled = True
    mblnOK = True
End Sub

Private Sub GetMaxMinDate(ByVal lngPatiID As Long, ByVal DatEnd As Date, ByRef DatMax As Date, ByRef DatMin As Date)
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim strTable As String, strDateMode As String
    
    '要和过程Zl_结帐费用记录_Patient中的待结费用游标一致,避免产生没有结帐费用的结帐单.
    '大表拆分:经与张永康落实,从SQL分析上来看,可能针对门诊的费用进行结算,但实质上应该只针对住院病人,因此,本次拆分只替换成住院费用记录
    strDateMode = IIf(mty_ModulePara.int费用时间 = 1, "发生时间", "登记时间")
    
    strSql = "" & _
    " Select Max(Max时间) DatMax, Min(Min时间) DatMin" & vbNewLine & _
    " From ( Select Max(" & strDateMode & ") Max时间, Min(" & strDateMode & ") Min时间" & vbNewLine & _
    "        From 住院费用记录 A" & vbNewLine & _
    "        Where A.病人id = [1] And A.结帐id Is Null And A.记录状态 <> 0 And Mod(记录性质, 10) In (2, 3) And" & vbNewLine & _
    "             " & strDateMode & " < [2] " & vbCrLf & _
    "             And Not Exists ( Select 1" & vbNewLine & _
    "                              From 住院费用记录 B" & vbNewLine & _
    "                              Where B.NO = A.NO And B.记录性质 = A.记录性质 And B.序号 = A.序号" & vbNewLine & _
    "                              Group By B.NO, B.记录性质, B.序号" & vbNewLine & _
    "                              Having Nvl(Sum(B.实收金额), 0) = Decode(" & IIf(gblnZero, 1, 0) & ", 1, 1 + Nvl(Sum(B.实收金额), 0), 0))" & vbNewLine & _
    "       Union All" & vbNewLine & _
    "       Select Max(" & strDateMode & ") Max时间, Min(" & strDateMode & ") Min时间" & vbNewLine & _
    "       From " & zlGetFullFieldsTable("住院费用记录") & vbNewLine & _
    "       Where A.病人id = [1] And A.结帐id Is Not Null And Mod(记录性质, 10) In (2, 3) And Nvl(A.实收金额, 0) <> Nvl(A.结帐金额, 0) And" & vbNewLine & _
    "             " & strDateMode & " < [2]" & vbNewLine & _
    "       Group By A.NO, A.序号, Mod(A.记录性质, 10), A.记录状态, A.执行状态" & vbNewLine & _
    "       Having Nvl(Sum(A.实收金额), 0) - Nvl(Sum(A.结帐金额), 0) <> 0)"


    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatiID, DatEnd)
    DatMax = NVL(rsTmp!DatMax, CDate(0))
    DatMin = NVL(rsTmp!DatMin, CDate(0))
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetDateStr(DatTmp As Date) As String
    GetDateStr = "To_Date('" & Format(DatTmp, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
End Function

Private Function GetBalanceSum(ByVal Dat收款时间 As Date) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取本次批量结算的结算信息
    '编制:刘兴洪
    '日期:2015-07-07 16:26:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    strSql = "" & _
    " Select Decode(Mod(B.记录性质,10),1,0,C.性质)  as 序号," & _
    "       DECODE(mod(B.记录性质,10),1,'冲预交', B.结算方式) as 结算方式," & _
    "       min(Decode(Mod(B.记录性质,10),1,0,C.性质)) as 结算性质," & _
    "       Sum(B.冲预交) 结算金额" & vbNewLine & _
    " From 病人结帐记录 A, 病人预交记录 B,结算方式 C" & vbNewLine & _
    " Where A.收费时间 = [1] And A.操作员姓名 = [2] and B.结算方式=C.名称(+) And A.ID = B.结帐id" & vbNewLine & _
    " Group By decode(Mod(B.记录性质,10),1,0,C.性质),DECODE(mod(B.记录性质,10),1,'冲预交', B.结算方式)" & _
    " order by 序号"
    On Error GoTo errH
    Set GetBalanceSum = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Dat收款时间, UserInfo.姓名)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SaveBalance(ByRef blnPrint As Boolean, ByRef rsPati As ADODB.Recordset)
    Dim strNO As String, i As Long, j As Long
    Dim lng主页ID As Long, lng病人ID As Long, lng结帐ID As Long
    Dim dtEndDate As Date, dtStartDate As Date, dtBalanceDate As Date
    Dim intCol As Integer
    Dim arrSQL As Variant, lngNum As Long, blnTrans As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strPreBalance As String '预结算信息:
    Dim intInsure As Integer, cllPro As Collection
    Dim strSql As String
    Dim str保险类别 As String
    Dim intScreenMouse As Integer
    
    vsBalance.Visible = False
    intScreenMouse = Screen.MousePointer
    Err = 0: On Error GoTo ErrHand:
    dtBalanceDate = zlDatabase.Currentdate '记录为统一的结帐时间

    Set cllPro = New Collection
    
    For i = 1 To rsPati.RecordCount
        arrSQL = Array()
        lng病人ID = rsPati!病人ID
        lng主页ID = Val(NVL(rsPati!主页ID))
        str保险类别 = NVL(rsPati!保险名称)
        Call GetMaxMinDate(lng病人ID, CDate(txtDateEnd.Text), dtEndDate, dtStartDate)
        
        If Not (dtEndDate = dtStartDate And dtEndDate = CDate(0)) Then '没有待结费用不结帐
            lblInfo.Caption = "当前进度:共" & rsPati.RecordCount & "位,正在进行第" & i & "位," & rsPati!科室 & ":" & rsPati!姓名
            Me.Refresh
            
            intInsure = Val(NVL(rsPati!险类))
            
            If zlStr.IsHavePrivs(mstrPrivs, "医保批量中途结帐") = False Then intInsure = 0
            
            MCPAR.批量中途结帐 = False
            If intInsure <> 0 Then
                '初始化参数
                Call InitInsurePara(lng病人ID, intInsure)
            End If
            
            If MCPAR.批量中途结帐 = False Then intInsure = 0
            If intInsure = 0 Then str保险类别 = ""
            
            
            '医保预结算
            strPreBalance = "" '报销方式|金额||....
            If InsureBudgeting(lng病人ID, lng主页ID, intInsure, dtStartDate, dtEndDate, strPreBalance) = False Then GoTo GoNextPati:
            lng结帐ID = zlDatabase.GetNextId("病人结帐记录")
            strNO = zlDatabase.GetNextNo(15)
            
            '1.病人结帐记录
            '58758
            'Zl_病人结帐记录_Insert
            strSql = "Zl_病人结帐记录_Insert("
            '  Id_In           病人结帐记录.Id%Type,
            strSql = strSql & "" & lng结帐ID & ","
            '  单据号_In       病人结帐记录.No%Type,
            strSql = strSql & "'" & strNO & "',"
            '  病人id_In       病人结帐记录.病人id%Type,
            strSql = strSql & "" & lng病人ID & ","
            '  收费时间_In     病人结帐记录.收费时间%Type,
            strSql = strSql & "" & GetDateStr(dtBalanceDate) & ","
            '  开始日期_In     病人结帐记录.开始日期%Type,
            strSql = strSql & "" & GetDateStr(dtStartDate) & ","
            '  结束日期_In     病人结帐记录.结束日期%Type,
            strSql = strSql & "" & GetDateStr(dtEndDate) & ","
            '  中途结帐_In     病人结帐记录.中途结帐%Type := 0,
            strSql = strSql & "1,"
            '  多病人结帐_In   Number := 0,
            strSql = strSql & "0,"
            '  最大结帐次数_In Number := 0,
            strSql = strSql & "" & lng主页ID & ","
            '  备注_In         病人结帐记录.备注%Type := Null,
            strSql = strSql & "NULL,"
            '  来源_In         Number := 1,
            strSql = strSql & "2,"
            '  原因_In         病人结帐记录.原因%Type := Null,
            strSql = strSql & "NULL,"
            '  结帐类型_In     病人结帐记录.结帐类型%Type := 2,
            strSql = strSql & "2,"
            '  结算状态_In     病人结帐记录.结算状态%Type := 0,
            strSql = strSql & "1,"
            '  住院次数_In     病人结帐记录.住院次数%Type := Null,  '住院次数及结帐金额在Zl_结帐费用记录_Patient过程中处理
            strSql = strSql & "" & IIf(lng主页ID = 0, "NULL", lng主页ID) & ","
            '  结帐金额_In     病人结帐记录.结帐金额%Type := Null
            strSql = strSql & 0 & ")"
            zlAddArray cllPro, strSql
            
            '3.住院费用记录
            'Zl_结帐费用记录_Patient
            strSql = "Zl_结帐费用记录_Patient("
            '  病人id_In     病人预交记录.病人id%Type,
            strSql = strSql & "" & lng病人ID & ","
            '  结帐id_In     住院费用记录.结帐id%Type,
            strSql = strSql & "" & lng结帐ID & ","
            '  截止时间_In   住院费用记录.登记时间%Type,
            strSql = strSql & "" & GetDateStr(CDate(txtDateEnd.Text)) & ","
            '  时间模式_In Number, --1: 发生时间 , 0: 登记时间
            strSql = strSql & "" & mty_ModulePara.int费用时间 & ","
            '  零费用结帐_In Number
            strSql = strSql & "" & IIf(mty_ModulePara.blnZero, 1, 0) & ")"
            zlAddArray cllPro, strSql
            
            '4.保存结帐的结算信息
            'Zl_批量结帐结算_Update
            strSql = "Zl_批量结帐结算_Update("
            '  病人id_In     门诊费用记录.病人id%Type,
            strSql = strSql & "" & lng病人ID & ","
            '  结帐id_In     病人预交记录.结帐id%Type,
            strSql = strSql & "" & lng结帐ID & ","
            '  保险结算_In   Varchar2,
            strSql = strSql & "" & IIf(strPreBalance = "", "null", "'" & strPreBalance & "'") & ","
            '  保险类别_In   保险类别.名称%Type,
            strSql = strSql & "" & IIf(str保险类别 = "", "null", "'" & str保险类别 & "'") & ","
            '  支付方式_In   结算方式.名称%Type,
            strSql = strSql & "'" & cbo结算方式.Text & "',"
            '  操作员编号_In 病人预交记录.操作员编号%Type,
            strSql = strSql & "'" & UserInfo.编号 & "',"
            '  操作员姓名_In 病人预交记录.操作员姓名%Type,
            strSql = strSql & "'" & UserInfo.姓名 & "',"
            '  收款时间_In   病人预交记录.收款时间%Type,
            strSql = strSql & "" & GetDateStr(dtBalanceDate) & ","
            '  完成结算_In Number:=0
            strSql = strSql & "" & IIf(intInsure <> 0, 0, 1) & ")"
            zlAddArray cllPro, strSql
            
            '4.开始票据号
            If blnPrint And Trim(txtInvoice.Text) <> "" Then
                strSql = "Zl_票据起始号_Update('" & strNO & "','" & Trim(txtInvoice.Text) & "',3)"
                zlAddArray cllPro, strSql
            End If
            
            On Error GoTo errH
            blnTrans = True
            zlExecuteProcedureArrAy cllPro, Me.Caption, True
            '医保结算
            If InsureBalance(lng病人ID, lng结帐ID, intInsure, strPreBalance, str保险类别, dtBalanceDate) Then
                gcnOracle.CommitTrans: blnTrans = False
                lngNum = lngNum + 1 '记录实际结帐人数
                '票据打印
                If blnPrint Then
                    mobjFact.LastUseID = mlng领用ID
                    Call frmPrint.ReportPrint(1, strNO, lng结帐ID, mobjFact, txtInvoice.Text, dtBalanceDate, "", "", lng病人ID, mobjFact.打印格式)
                    Call RefreshFact
                End If
            End If
            Set cllPro = New Collection
            blnTrans = False
        End If
GoNextPati:
        rsPati.MoveNext
    Next
    If lngNum = 0 Then
        lblInfo.Caption = "选择了" & rsPati.RecordCount & "位病人,但在指定的截止时间前都不存在未结费用!"
        vsBalance.Visible = False
    Else
        lblInfo.Caption = "对" & rsPati.RecordCount & "位病人中,存在未结费用的" & lngNum & "位完成了中途结帐."
        vsBalance.Visible = True
        Set rsTmp = GetBalanceSum(dtBalanceDate)
        With vsBalance
            intCol = 0: .Cols = rsTmp.RecordCount * 2
            .Rows = 1
            Do While Not rsTmp.EOF
                .TextMatrix(0, intCol) = NVL(rsTmp!结算方式)
                .Cell(flexcpFontBold, 0, intCol) = True
                .TextMatrix(0, intCol + 1) = zlStr.FormatEx(Val(NVL(rsTmp!结算金额)), 6)
                intCol = intCol + 2
                rsTmp.MoveNext
            Loop
            .AutoSizeMode = flexAutoSizeColWidth
            .AutoSize 0, .Cols - 1
        End With
    End If
    Exit Sub
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Screen.MousePointer = intScreenMouse
        Resume
    End If
    
    Call SaveErrLog
    If lngNum > 0 Then
        lblInfo.Caption = "选择了" & rsPati.RecordCount & "位病人,实际对" & lngNum & "位病人完成了中途结帐."
    End If
    Exit Sub
ErrHand:
     If ErrCenter = 1 Then
        Resume
     End If
End Sub
Private Function InsureBudgeting(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal intInsure As Integer, ByVal dtStartDate As Date, _
    ByVal dtEndDate As Date, ByRef strPreBalance As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保预结算
    '入参: intInsure-险类
    '出参:strBalance-返回预结算信息:报销方式|金额||....
    '返回:预算成功(含普通病人未行医保虚拟结算),返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-06 16:48:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln发生时间 As Boolean, str医保号 As String
    Dim strBalance As String, varData As Variant, varTemp As Variant
    Dim strNotBalance As String '不存在的结算方式
    Dim lngRow As Long, blnOk As Boolean
    Dim cur个人帐户 As Currency, cur统筹支付 As Currency
    Dim curMoney As Currency
    Dim rsDetail As ADODB.Recordset
    
    Dim i As Long, byt状态 As Byte, bytEditSta As Byte
    On Error GoTo errHandle
    strPreBalance = ""
    If intInsure = 0 Then InsureBudgeting = True: Exit Function
    
    bln发生时间 = mty_ModulePara.int费用时间 = 1 '0-按登记时间,1-按发生时间
    '医保预结算
    Set rsDetail = GetZYBalance_Insure(intInsure, lng病人ID, _
         IIf(lng主页ID = 0, "", CStr(lng主页ID)), dtStartDate, dtEndDate, _
        MCPAR.中途结算仅处理已上传部分, False, 0, "", "", "", "", bln发生时间)
    
    '结算方式;金额;是否允许修改|...
    strBalance = gclsInsure.WipeoffMoney(rsDetail, lng病人ID, str医保号, "1", intInsure, "|1")
    varData = Split(strBalance, "|")
    
    strPreBalance = ""
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ";;;;", ";")
        If varTemp(0) <> "" Then
            strPreBalance = strPreBalance & "||" & varTemp(0) & "|" & varTemp(1)
        End If
    Next
    If strPreBalance <> "" Then strPreBalance = Mid(strPreBalance, 3)
    InsureBudgeting = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function InsureBalance(ByVal lng病人ID As Long, ByVal lng结帐ID As Long, _
    ByVal intInsure As Integer, ByVal str预结算 As String, ByVal str保险类别 As String, ByVal dtBalanceDate As Date) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保结算接口
    '入参:intInsure-险类
    '返回:结算成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-06 15:30:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strAdvance As String, blnTruns As Boolean
    Dim cllPro As New Collection
    On Error GoTo errHandle
    
    '住院医保结算
    If intInsure = 0 Then InsureBalance = True: Exit Function
    
    If Not gclsInsure.SettleSwap(lng结帐ID, intInsure, strAdvance) Then
        gcnOracle.RollbackTrans: Screen.MousePointer = 0: Exit Function
    End If
    If strAdvance <> "" Then
        If zlInsure_Check(str预结算, strAdvance) Then
            blnTruns = True
            Call 医保数据更正(lng病人ID, lng结帐ID, strAdvance, str保险类别, dtBalanceDate, cllPro)
            zlExecuteProcedureArrAy cllPro, Me.Caption, True, True
        End If
    End If
    InsureBalance = True
    Exit Function
errHandle:
    Call gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Function

Private Function 医保数据更正(ByVal lng病人ID As Long, ByVal lng结帐ID As Long, _
    ByVal str医保结算 As String, ByVal str保险类别 As String, ByVal dtBalanceDate As Date, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保数据校对更正
    '返回:校对成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-12 17:45:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    '4.保存结帐的结算信息
    'Zl_批量结帐结算_Update
    strSql = "Zl_批量结帐结算_Update("
    '  病人id_In     门诊费用记录.病人id%Type,
    strSql = strSql & "" & lng病人ID & ","
    '  结帐id_In     病人预交记录.结帐id%Type,
    strSql = strSql & "" & lng结帐ID & ","
    '  保险结算_In   Varchar2,
    strSql = strSql & "" & IIf(str医保结算 = "", "null", "'" & str医保结算 & "'") & ","
    '  保险类别_In   保险类别.名称%Type,
    strSql = strSql & "" & IIf(str保险类别 = "", "null", "'" & str保险类别 & "'") & ","
    '  支付方式_In   结算方式.名称%Type,
    strSql = strSql & "'" & cbo结算方式.Text & "',"
    '  操作员编号_In 病人预交记录.操作员编号%Type,
    strSql = strSql & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In 病人预交记录.操作员姓名%Type,
    strSql = strSql & "'" & UserInfo.姓名 & "',"
    '  收款时间_In   病人预交记录.收款时间%Type,
    strSql = strSql & "" & GetDateStr(dtBalanceDate) & ","
    '  完成结算_In Number:=0
    strSql = strSql & "1)"
    zlAddArray cllPro, strSql
    医保数据更正 = True
End Function
Public Function zlInsure_Check(ByVal str保险结算 As String, ByVal strAdvance As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前的医保是否需要较对
    '入参:str保险结算-保险结算
    '       strAdvance-医保返回的结算
    '出参:
    '返回:需要较对,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-20 18:03:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnMedicareCheck As Boolean, strTmp As String, i As Long, j As Long
    Dim varData As Variant, varData1 As Variant
    Dim varTemp As Variant, varTemp1 As Variant
    
    On Error GoTo errHandle
    If Not (strAdvance <> "" And str保险结算 <> strAdvance) Then Exit Function
    '正式结算前后,结算方式和结算金额未发生变化时不校对
    blnMedicareCheck = True
    varData = Split(str保险结算, "||"): varData1 = Split(strAdvance, "||")
    
    If UBound(varData) = UBound(varData1) Then
        For i = 0 To UBound(varData)
            blnMedicareCheck = True
            strTmp = varData(i)
            varTemp = Split(strTmp, "|")
            For j = 0 To UBound(varData1)
                varTemp1 = Split(varData1(j), "|")
                If varTemp(0) = varTemp1(0) Then
                    If Val(varTemp(1)) = Val(varTemp1(1)) Then
                        blnMedicareCheck = False
                    End If
                End If
            Next
            If blnMedicareCheck Then Exit For
        Next
    End If
    zlInsure_Check = blnMedicareCheck
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub InitInsurePara(ByVal lng病人ID As Long, ByVal intInsure As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化医保参数
    '编制:刘兴洪
    '日期:2015-03-25 17:59:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    MCPAR.分币处理 = gclsInsure.GetCapability(support分币处理, lng病人ID, intInsure)
    MCPAR.未结清出院 = gclsInsure.GetCapability(support未结清出院, lng病人ID, intInsure)
    MCPAR.结算使用个人帐户 = gclsInsure.GetCapability(support结算使用个人帐户, lng病人ID, intInsure)
    MCPAR.出院结算必须出院 = gclsInsure.GetCapability(support出院结算必须出院, lng病人ID, intInsure)
    MCPAR.中途结算仅处理已上传部分 = gclsInsure.GetCapability(support中途结算仅处理已上传部分, lng病人ID, intInsure)
    MCPAR.结帐设置后调用接口 = gclsInsure.GetCapability(support结帐_结帐设置后调用接口, lng病人ID, intInsure)
    MCPAR.住院结算作废 = gclsInsure.GetCapability(support住院结算作废, lng病人ID, intInsure)
    MCPAR.批量中途结帐 = gclsInsure.GetCapability(support批量中途结帐, lng病人ID, intInsure)
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitFact(ByVal str使用类别 As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化发票信息
    '编制:刘兴洪
    '日期:2015-02-05 11:26:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bytInvoiceKind     As Byte, intFormat As Integer
    Dim intPrintMode As Integer, lngShareUseID As Long
    On Error GoTo errHandle
 
    
    bytInvoiceKind = Val(zlDatabase.GetPara("住院结帐票据类型", glngSys, 1137, "0"))

    Set mobjInvoice = New clsInvoice: Set mobjFact = New clsFactProperty
    mobjInvoice.zlInitCommon glngSys, gcnOracle, gstrDBUser
    
    Call mobjInvoice.zlGetInvoicePreperty(1137, IIf(bytInvoiceKind = 0, 3, 1), 0, 0, 0, mobjFact, , , 2)
    
    mobjFact.使用类别 = str使用类别
    mobjFact.票种 = IIf(bytInvoiceKind = 0, 3, 1)
    Call mobjInvoice.zlGetInvoicePrintFormat(1137, mobjFact.票种, mobjFact.使用类别, intFormat, 2)
    mobjFact.打印格式 = intFormat
    If mobjInvoice.zlGetInvoicePrintMode(1137, mobjFact.票种, mobjFact.使用类别, intPrintMode) = False Then Exit Sub
    mobjFact.打印方式 = intPrintMode
    If mobjInvoice.zlGetInvoiceShareID(1137, mobjFact.票种, mobjFact.使用类别, lngShareUseID) = False Then Exit Sub
    mobjFact.共享批次ID = lngShareUseID
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub RefreshFact()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷新收费票据号
    '编制:刘兴洪
    '日期:2015-02-05 11:40:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFactNO As String
       
    On Error GoTo errHandle
        
    If mobjFact.打印方式 = 0 Then Exit Sub
    If Not mobjFact.严格控制 Then
        '非严格控制下
        '松散：取下一个号码
        txtInvoice.Text = zlCommFun.IncStr(UCase(zlDatabase.GetPara("当前结帐票据号", glngSys, 1137, "")))
        txtInvoice.Tag = txtInvoice.Text
        txtInvoice.SelStart = Len(txtInvoice.Text)
        Exit Sub
    End If
    If zlGetInvoiceGroupUseID(mlng领用ID, 1, "") = False Then
          txtInvoice.Text = "": txtInvoice.Tag = ""
        Exit Sub
    End If
    
    '严格：取下一个号码
    If mobjInvoice.zlGetNextBill(1137, mlng领用ID, strFactNO) = False Then strFactNO = ""
    txtInvoice.Text = strFactNO
    'Tag：问题：24363:刘兴洪：主要是解决自动生成的号是否被用户更改，主要解决：
    '    1.更改的票据号需要检查是否重复，重复后直接返回不更改发票号
    '    2.并发操作，不更改的情况下，检查是否重复，如果重复，自动取下一个号码！
    txtInvoice.Tag = txtInvoice.Text
    lblFact.Tag = txtInvoice.Tag
    txtInvoice.SelStart = Len(txtInvoice.Text)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function GetPatiSet() As ADODB.Recordset
    Dim strSql As String, str费别 As String, strDeptIDs As String, i As Long
    Dim intInsure As Integer
    Dim strWhere As String
    
    str费别 = Get费别选择
    If str费别 <> "" Then
        If UBound(Split(str费别, ",")) + 1 < vsFeeType.Rows - 1 Then
            str费别 = "," & str费别 & ","
            strWhere = " And Instr([2],','||A.费别||',') > 0"
        End If
    End If
    
    For i = 1 To vsDept.Rows - 1
        If Not (vsDept.Cell(flexcpChecked, i, 0) = flexUnchecked Or vsDept.Cell(flexcpChecked, i, 0) = flexTSUnchecked) Then
            strDeptIDs = strDeptIDs & "," & vsDept.RowData(i)
        End If
    Next
    strDeptIDs = Mid(strDeptIDs, 2)
    If UBound(Split(strDeptIDs, ",")) + 1 = vsDept.Rows - 1 Then strDeptIDs = ""
    
    If strDeptIDs <> "" Then
        strWhere = strWhere & " And B.科室ID In(" & strDeptIDs & ")"
    End If
    
    If mstrPatis <> "" Then
        mstrPatis = "," & mstrPatis & ","
        strWhere = strWhere & " And Instr([1],','||B.病人ID||',') = 0"
    End If
    
    intInsure = 0
    If cboInsure.ListIndex >= 0 Then intInsure = cboInsure.ItemData(cboInsure.ListIndex)
    strWhere = strWhere & IIf(intInsure = 0, "", " And nvl(A.险类,0) =[4]")
    
    strSql = "" & _
    "Select Distinct C.名称 as 科室,A.姓名,A.病人ID,A.住院次数,A.主页ID," & _
    "   A.住院号,nvl(A.险类,0) as 险类,J.名称 as 保险名称 " & vbNewLine & _
    "From 病人信息 A, 床位状况记录 B, 部门表 C,病案主页 M,保险类别 J" & vbNewLine & _
    "Where A.病人id = B.病人ID   " & _
    "       And B.科室ID = C.ID And A.险类=J.序号(+)  " & _
    "       And Zl_Billclass(A.病人ID,A.主页ID,A.险类)=[3]  " & strWhere & vbNewLine & _
    "Order by 科室,住院号"

    On Error GoTo errH
    Set GetPatiSet = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mstrPatis, str费别, Trim(cbo使用类别.Text), intInsure)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub Form_Load()
    Set mrsRptFormat = Nothing
    lblInfo.Caption = ""
    mlngModul = 1137
    
    Call zlInitModulePara
    Call LoadUseType    '加载使用类别
    Call LoadInsureType '加载有效的险类
    Call InitFact(cbo使用类别.Text)
    txtInvoice.Locked = Not zlStr.IsHavePrivs(mstrPrivs, "修改票据号") And gblnStrictCtrl '89302
    
    mblnNotClick = True
    If Not InitData Then Unload Me: mblnNotClick = False: Exit Sub
    If vsDept.Rows > 1 Then
        vsDept.Row = 1
    Else
        cmdOK.Enabled = False
    End If
    If vsDept.Row > 0 Then
        Call LoadPati(Val(vsDept.RowData(vsDept.Row)))
        lblRpt.Caption = ""
        If mrsRptFormat Is Nothing Then Exit Sub
        mrsRptFormat.Filter = "序号=" & cbo使用类别.ItemData(cbo使用类别.ListIndex)
        If Not mrsRptFormat.EOF Then
            lblRpt.Caption = NVL(mrsRptFormat!说明)
        End If
    End If
    mblnNotClick = False
End Sub

Private Function InitData() As Boolean
    Dim rsTmp As ADODB.Recordset, i As Long

    Set rsTmp = Get费别
    If rsTmp.RecordCount = 0 Then
        MsgBox "费别未设置,不能使用此功能!", vbInformation, gstrSysName
        Exit Function
    Else
        vsFeeType.Rows = rsTmp.RecordCount + 1
        vsFeeType.ColDataType(0) = flexDTBoolean
        vsFeeType.Cell(flexcpChecked, 1, 0, vsFeeType.Rows - 1, 0) = flexChecked
        vsFeeType.Row = 1: vsFeeType.Col = 1: vsFeeType.Col = 0
    End If
    For i = 1 To rsTmp.RecordCount
        vsFeeType.TextMatrix(i, 1) = NVL(rsTmp!名称)
        rsTmp.MoveNext
    Next
    Call LoadDept
    
    txtDateEnd.Text = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    Set rsTmp = Get结算方式("结帐", 2)
    If rsTmp.RecordCount = 0 Then
        MsgBox "没有设置用于结帐场合的非现金结算方式,不能使用此功能!", vbInformation, gstrSysName
        Exit Function
    End If
    For i = 1 To rsTmp.RecordCount
        cbo结算方式.AddItem rsTmp!名称
        rsTmp.MoveNext
    Next
    cbo结算方式.ListIndex = 0
    
    Call RefreshFact
    
    InitData = True
End Function

Private Function Get费别() As ADODB.Recordset
    Dim strSql As String
 
    strSql = "Select 名称,编码 From 费别 Where 服务对象 In (2, 3) And 属性 = 1 Order by 编码"
    On Error GoTo errH
    Set Get费别 = zlDatabase.OpenSQLRecord(strSql, Me.Caption)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadDept()
    Dim rsTmp As ADODB.Recordset, strSql As String, i As Long
 
    strSql = "" & _
    "Select A.ID, A.名称" & vbNewLine & _
    "From 部门表 A, 部门性质说明 B" & vbNewLine & _
    "Where A.ID = B.部门id And B.服务对象 In (2, 3) And B.工作性质 = '临床'" & vbNewLine & _
    "   And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & vbNewLine & _
    "   And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
    "   And Exists(Select 1 From 床位状况记录 C Where C.病人id Is Not Null And C.科室id = A.ID) " & _
    " Order by 名称"
    On Error GoTo errH
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    vsDept.Rows = rsTmp.RecordCount + 1
    If rsTmp.RecordCount > 0 Then
        vsDept.Cell(flexcpChecked, 1, 0, vsDept.Rows - 1, 0) = flexChecked
        vsDept.Row = 1: vsDept.Col = 1: vsDept.Col = 0
    End If
    For i = 1 To rsTmp.RecordCount
        vsDept.TextMatrix(i, 1) = rsTmp!名称
        vsDept.RowData(i) = Val(rsTmp!ID)
        rsTmp.MoveNext
    Next
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Get费别选择() As String
    Dim i As Long, strTmp As String
    
    For i = 1 To vsFeeType.Rows - 1
        If vsFeeType.Cell(flexcpChecked, i, 0) = flexChecked Then strTmp = strTmp & "," & vsFeeType.TextMatrix(i, 1)
    Next
    Get费别选择 = Mid(strTmp, 2)
End Function

Private Sub LoadPati(ByVal lngDeptID As Long)
    Dim rsTmp As ADODB.Recordset, strSql As String, i As Long, str费别 As String
    Dim intInsure As Integer
    Dim strWhere As String
    
    str费别 = Get费别选择
    If str费别 <> "" Then
        If UBound(Split(str费别, ",")) + 1 < vsFeeType.Rows - 1 Then
            str费别 = "," & str费别 & ","
            strSql = " And Instr([2],','||A.费别||',')>0"
        End If
    End If
    intInsure = 0
    If cboInsure.ListIndex >= 0 Then intInsure = cboInsure.ItemData(cboInsure.ListIndex)
    strWhere = IIf(intInsure = 0, "", " And nvl(A.险类,0) =[4] ")
    
    On Error GoTo errH
    strSql = "" & _
    "   Select Distinct A.病人ID,A.住院号, Nvl(D.姓名,A.姓名) as 姓名, Nvl(D.性别,A.性别) as 性别, " & _
    "           Nvl(D.年龄,A.年龄) as 年龄, B.费用余额 未结费用, 预交余额 可用预交, A.费别," & _
    "           nvl(A.险类,0) as 险类,M.名称 as 保险名称" & vbNewLine & _
    "   From 病人信息 A, 病人余额 B,床位状况记录 C,病案主页 D,保险类别 M" & vbNewLine & _
    "   Where C.科室id = [1]  " & strWhere & _
    "         And A.病人id=D.病人ID(+) And A.主页id = D.主页id(+) " & _
    "         And A.病人id=C.病人ID  And A.病人id = B.病人id(+) " & _
    "         And B.性质(+) = 1  And B.类型(+)=2 and A.险类=M.序号(+) " & _
    "         And Zl_Billclass(A.病人ID, A.主页ID, A.险类)=[3] " & strSql & vbNewLine & _
    "   Order by A.住院号"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngDeptID, str费别, Trim(cbo使用类别.Text), intInsure)
    
    vsPati.Rows = 1 '清除数据,但不清除列标头
    vsPati.Rows = rsTmp.RecordCount + 1
    If rsTmp.RecordCount > 0 Then
        If vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexUnchecked Or vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexTSUnchecked Then
            vsPati.Cell(flexcpChecked, 1, 0, vsPati.Rows - 1, 0) = flexUnchecked
        Else
            vsPati.Cell(flexcpChecked, 1, 0, vsPati.Rows - 1, 0) = flexChecked
        End If
    Else
        vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexUnchecked
    End If
    
    With vsPati
        For i = 1 To rsTmp.RecordCount
            .TextMatrix(i, .ColIndex("住院号")) = "" & rsTmp!住院号
            .TextMatrix(i, .ColIndex("姓名")) = "" & rsTmp!姓名
            .TextMatrix(i, .ColIndex("性别")) = "" & rsTmp!性别
            .TextMatrix(i, .ColIndex("年龄")) = "" & rsTmp!年龄
            .TextMatrix(i, .ColIndex("未结费用")) = Format(Val(NVL(rsTmp!未结费用)), "###0.00;-###0.00;;")
            .TextMatrix(i, .ColIndex("可用预交")) = Format(Val(NVL(rsTmp!可用预交)), "###0.00;-###0.00;;")
            .TextMatrix(i, .ColIndex("费别")) = "" & rsTmp!费别
            .TextMatrix(i, .ColIndex("险类")) = Val(NVL(rsTmp!险类))
            .TextMatrix(i, .ColIndex("保险类别")) = NVL(rsTmp!保险名称)
            .RowData(i) = Val(rsTmp!病人ID)
            If Len(mstrPatis) > 0 Then
                If InStr("," & mstrPatis & ",", "," & rsTmp!病人ID & ",") > 0 Then
                    .Cell(flexcpChecked, i, 0) = flexUnchecked
                End If
            End If
            If Val(NVL(rsTmp!险类)) <> 0 Then
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
            Else
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = .ForeColor
            End If
            rsTmp.MoveNext
        Next
        If rsTmp.RecordCount > 0 Then .Row = 1: .Col = 1: .Col = 0
    End With
    Exit Sub
    
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    If Me.Width < 12060 Then Me.Width = 12060
    If Me.Height < 7635 Then Me.Height = 7635
    With fra
        .Width = ScaleWidth - .Left * 2
    End With
    With picDown
        .Width = ScaleWidth
        .Top = ScaleHeight - .Height - 100
    End With
     With vsFeeType
        .Height = picDown.Top - .Top - 50
        vsDept.Height = .Height
        vsPati.Height = .Height
        vsPati.Width = ScaleWidth - vsPati.Left - 50
     End With
     vsBalance.Width = cmdOK.Left - vsBalance.Left - 100
End Sub
 

Private Sub Form_Unload(Cancel As Integer)
    Set mrsRptFormat = Nothing
    mstrPatis = ""
    mlng领用ID = 0
    Set mobjInvoice = Nothing
    Set mobjFact = Nothing

End Sub

Private Sub picDown_Resize()
  Err = 0: On Error Resume Next
    With cmdCancel
        .Left = picDown.ScaleWidth - cmdCancel.Width - 100
        cmdOK.Left = .Left - cmdOK.Width - 50
    End With
End Sub

Private Sub vsDept_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    'If vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexTSGrayed Then vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexChecked    '手动点击时点为灰色，改为选择
    
    If Row <> vsDept.Row Then vsDept.Row = Row
    If vsPati.Rows < 2 Then Exit Sub
    
    If vsDept.Cell(flexcpChecked, Row, 0) = flexUnchecked Or vsDept.Cell(flexcpChecked, Row, 0) = flexTSUnchecked Then
        vsPati.Cell(flexcpChecked, 1, 0, vsPati.Rows - 1, 0) = flexUnchecked
    Else
        vsPati.Cell(flexcpChecked, 1, 0, vsPati.Rows - 1, 0) = flexChecked
    End If
    Call SetPatiLists
End Sub
Private Sub vsdept_DblClick()
    If vsDept.MouseCol = 0 And vsDept.MouseRow = 0 Then
        Call SetVSAll(vsDept)
        Call vsDept_AfterEdit(vsDept.Row, vsDept.Col)
        mstrPatis = ""
    End If
End Sub

Private Sub vsDept_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow <> OldRow And NewRow <> 0 Then Call LoadPati(Val(vsDept.RowData(NewRow)))
End Sub



Private Sub vsFeeType_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    i = vsDept.Row
    vsDept.Row = 0
    vsDept.Row = i
    
End Sub

Private Sub vsPati_DblClick()
    If vsPati.MouseCol = 0 And vsPati.MouseRow = 0 Then
        If vsPati.Rows < 2 Then Exit Sub
        
        Call SetVSAll(vsPati)
        Call SetDeptState
        Call SetPatiLists
    End If
End Sub

Private Sub vsFeeType_DblClick()
    Dim i As Long
    If vsFeeType.MouseCol = 0 And vsFeeType.MouseRow = 0 Then
        Call SetVSAll(vsFeeType)
        i = vsDept.Row
        vsDept.Row = 0
        vsDept.Row = i
    End If
End Sub

Private Sub SetVSAll(ByRef vsf As VSFlexGrid)
    If vsf.Rows < 2 Then Exit Sub
    vsf.Cell(flexcpChecked, 1, 0, vsf.Rows - 1, 0) = IIf(Val(vsf.Tag) = 1, flexChecked, flexUnchecked)
    vsf.Tag = IIf(Val(vsf.Tag) = 0, 1, 0)
End Sub


Private Sub vsPati_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexUnchecked Or vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexTSUnchecked Then
        SetPatiLists
    Else
        Call SetPatistr(Row)
    End If
    Call SetDeptState
End Sub

Private Sub SetPatistr(ByVal lngRow As Long)
'功能：记录没有选择的病人ＩＤ
    If vsPati.Cell(flexcpChecked, lngRow, 0) = flexUnchecked Then
        If InStr("," & mstrPatis & ",", "," & vsPati.RowData(lngRow) & ",") = 0 Then
            If mstrPatis = "" Then
                mstrPatis = vsPati.RowData(lngRow)
            Else
                mstrPatis = mstrPatis & "," & vsPati.RowData(lngRow)
            End If
        End If
    Else
        If InStr("," & mstrPatis & ",", "," & vsPati.RowData(lngRow) & ",") > 0 Then
            mstrPatis = Replace("," & mstrPatis & ",", "," & vsPati.RowData(lngRow) & ",", ",")
            mstrPatis = Mid(mstrPatis, 2)   '去掉前后的
            If mstrPatis <> "" Then mstrPatis = Mid(mstrPatis, 1, Len(mstrPatis) - 1)
        End If
    End If
    If mstrPatis = "," Then mstrPatis = ""
End Sub

Private Sub SetPatiLists()
'功能:检查当前病人列表，把没有选择的加入到变量中，已选择的，从变量中删除
    Dim i As Long
    
    If vsPati.Rows < 2 Then Exit Sub
    
    For i = 1 To vsPati.Rows - 1
        Call SetPatistr(i)
    Next
End Sub

Private Function SetDeptState() As Boolean
'功能：设置科室选择状态
    Dim i As Long, m As Long
    
    For i = 1 To vsPati.Rows - 1
        If vsPati.Cell(flexcpChecked, i, 0) = flexChecked Then m = m + 1
    Next
    If m = vsPati.Rows - 1 Then
        vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexChecked
    ElseIf m = 0 Then
        vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexUnchecked
    Else
        vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexTSGrayed
    End If
End Function

Private Sub vspati_EnterCell()
    If vsPati.Col = 0 Then
        vsPati.Editable = flexEDKbdMouse
    Else
        vsPati.Editable = flexEDNone
    End If
End Sub
Private Sub vsfeetype_EnterCell()
    If vsFeeType.Col = 0 Then
        vsFeeType.Editable = flexEDKbdMouse
    Else
        vsFeeType.Editable = flexEDNone
    End If
End Sub
Private Sub vsDept_EnterCell()
    If vsDept.Col = 0 Then
        vsDept.Editable = flexEDKbdMouse
    Else
        vsDept.Editable = flexEDNone
    End If
End Sub
Private Sub LoadUseType()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载使用类别
    '编制:刘兴洪
    '日期:2011-04-28 15:09:10
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim i As Long, strSql As String
    Dim varData As Variant, varTemp As Variant
    Dim strRptName As String
    Dim strShareInvoice As String
    
    On Error GoTo errHandle
    
    strShareInvoice = zlDatabase.GetPara("结帐发票格式", glngSys, 1137)
    varData = Split(strShareInvoice, "|")
    
    strRptName = IIf(gbytInvoiceKind = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2")
    
    '票据格式处理
    strSql = "" & _
    "   Select '使用本地缺省格式' as 说明,0 as 序号  From Dual Union ALL " & _
    "   Select B.说明,B.序号  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.报表ID And A.编号=[1]" & _
    "   Order by  序号"
    Set mrsRptFormat = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strRptName)
    
    mblnNotClick = True
    strSql = "" & _
    "   Select 编码 ,名称" & _
    "   From  票据使用类别" & _
    "   order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    With cbo使用类别
        .Clear
        Do While Not rsTemp.EOF
            .AddItem NVL(rsTemp!名称)
            .ItemData(.NewIndex) = 0
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                If Trim(varTemp(0)) = Trim(NVL(rsTemp!名称)) Then
                    .ItemData(.NewIndex) = Val(varTemp(1))
                    Exit For
                End If
            Next
            rsTemp.MoveNext
        Loop
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
    End With
    mblnNotClick = False
    Exit Sub
errHandle:
    mblnNotClick = False
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Function zlGetInvoiceGroupUseID(ByRef lng领用ID As Long, _
    Optional intNum As Integer = 1, Optional strInvoiceNO As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取票据的领用ID
    '入参:lng领用ID-领用id
    '       intNum-页数
    '       strInvoiceNO-输入的发票号
    '出参:lng领用ID-领用ID
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-04-29 15:36:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjInvoice.zlGetInvoiceGroupID(1137, UserInfo.姓名, mobjFact.票种, _
        mobjFact.使用类别, lng领用ID, mobjFact.共享批次ID, lng领用ID, intNum, strInvoiceNO) = False Then Exit Function
    
    If lng领用ID > 0 Then zlGetInvoiceGroupUseID = True: Exit Function
    
    Select Case lng领用ID
        Case 0 '操作失败
        Case -1
            If Trim(mobjFact.使用类别) = "" Then
                MsgBox "你没有自用和共用的结帐票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
            Else
                MsgBox "你没有自用和共用的『" & mobjFact.使用类别 & "』结帐票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
            End If
            Exit Function
        Case -2
            If Trim(mobjFact.使用类别) = "" Then
                MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
            Else
                MsgBox "本地的共用票据的『" & mobjFact.使用类别 & "』结帐票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
            End If
            Exit Function
        Case -3
            MsgBox "当前票据号码不在可用领用批次的有效票据号范围内,请重新输入！", vbInformation, gstrSysName
            If txtInvoice.Enabled And txtInvoice.Visible Then txtInvoice.SetFocus
            Exit Function
    End Select
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

