VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClinicWorkTimeEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "上班时间设置"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7380
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClinicWorkTimeEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4468.677
   ScaleMode       =   0  'User
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   378
      Left            =   4500
      TabIndex        =   31
      Top             =   4110
      Width           =   1100
   End
   Begin VB.CommandButton cmdSaveExit 
      Caption         =   "保存退出(&C)"
      Height          =   378
      Left            =   4265
      TabIndex        =   30
      Top             =   4110
      Width           =   1335
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   378
      Left            =   390
      TabIndex        =   27
      Top             =   4110
      Width           =   1245
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HasDC           =   0   'False
      Height          =   3915
      Left            =   30
      ScaleHeight     =   3915
      ScaleWidth      =   7350
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   0
      Width           =   7350
      Begin VB.TextBox txt预留时间 
         Height          =   330
         Left            =   3630
         TabIndex        =   16
         Text            =   "0"
         Top             =   1590
         Width           =   975
      End
      Begin MSComCtl2.UpDown upd预留时间 
         Height          =   300
         Left            =   4590
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1590
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txt预留时间"
         BuddyDispid     =   196613
         OrigLeft        =   6300
         OrigTop         =   1560
         OrigRight       =   6555
         OrigBottom      =   1860
         Max             =   99999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "删除(&D)"
         Enabled         =   0   'False
         Height          =   330
         Left            =   6360
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2055
         Width           =   885
      End
      Begin VB.ComboBox cbo时间段 
         Height          =   330
         ItemData        =   "frmClinicWorkTimeEdit.frx":000C
         Left            =   930
         List            =   "frmClinicWorkTimeEdit.frx":000E
         TabIndex        =   5
         Top             =   585
         Width           =   2505
      End
      Begin VB.CommandButton cmdAddRestTime 
         Caption         =   "增加(&A)"
         Height          =   330
         Left            =   5460
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2055
         Width           =   885
      End
      Begin MSComCtl2.DTPicker dtpRestEndTime 
         Height          =   330
         Left            =   2550
         TabIndex        =   21
         Top             =   2055
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm:ss"
         Format          =   159252483
         UpDown          =   -1  'True
         CurrentDate     =   42370
      End
      Begin VB.Frame fraLineBetween2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   25
         Left            =   -60
         TabIndex        =   29
         Top             =   3840
         Width           =   8145
      End
      Begin VB.Frame fraLineBetween1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   25
         Left            =   -30
         TabIndex        =   6
         Top             =   1020
         Width           =   8025
      End
      Begin VB.ComboBox cbo号类 
         Height          =   330
         Left            =   4650
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   150
         Width           =   2595
      End
      Begin MSComCtl2.DTPicker dtpDefaultTime 
         Height          =   330
         Left            =   6000
         TabIndex        =   12
         Top             =   1155
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm:ss"
         Format          =   159252483
         UpDown          =   -1  'True
         CurrentDate     =   42370
      End
      Begin MSComCtl2.DTPicker dtpPriorTime 
         Height          =   330
         Left            =   930
         TabIndex        =   14
         Top             =   1605
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm:ss"
         Format          =   159252483
         UpDown          =   -1  'True
         CurrentDate     =   42370
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf休息时段 
         Height          =   1335
         Left            =   930
         TabIndex        =   24
         Top             =   2400
         Width           =   6315
         _cx             =   11139
         _cy             =   2355
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483638
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
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
         ExplorerBar     =   0
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
      Begin VB.ComboBox cboNodeNo 
         Height          =   330
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   150
         Width           =   2505
      End
      Begin MSComCtl2.DTPicker dtpEndTime 
         Height          =   330
         Left            =   3630
         TabIndex        =   10
         Top             =   1155
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm:ss"
         Format          =   159252483
         UpDown          =   -1  'True
         CurrentDate     =   42370
      End
      Begin MSComCtl2.DTPicker dtpStartTime 
         Height          =   330
         Left            =   930
         TabIndex        =   8
         Top             =   1155
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm:ss"
         Format          =   159252483
         UpDown          =   -1  'True
         CurrentDate     =   42370
      End
      Begin MSComCtl2.DTPicker dtpRestStartTime 
         Height          =   330
         Left            =   930
         TabIndex        =   19
         Top             =   2055
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm:ss"
         Format          =   159252483
         UpDown          =   -1  'True
         CurrentDate     =   42370
      End
      Begin VB.Label lbl预留时间 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预留时间(分)"
         Height          =   210
         Left            =   2370
         TabIndex        =   15
         Top             =   1650
         Width           =   1260
      End
      Begin VB.Label lblPriorTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "提前时间"
         Height          =   210
         Left            =   60
         TabIndex        =   13
         Top             =   1650
         Width           =   840
      End
      Begin VB.Label lblDefaultTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "缺省时间"
         Height          =   210
         Left            =   5130
         TabIndex        =   11
         Top             =   1200
         Width           =   840
      End
      Begin VB.Label lblRestTimeAnd 
         AutoSize        =   -1  'True
         Caption         =   "～"
         Height          =   210
         Left            =   2250
         TabIndex        =   20
         Top             =   2130
         Width           =   210
      End
      Begin VB.Label lblRestTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "休息时间"
         Height          =   210
         Left            =   60
         TabIndex        =   18
         Top             =   2100
         Width           =   840
      End
      Begin VB.Label lblEndTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "终止时间"
         Height          =   210
         Left            =   2730
         TabIndex        =   9
         Top             =   1200
         Width           =   840
      End
      Begin VB.Label lblStartTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开始时间"
         Height          =   210
         Left            =   60
         TabIndex        =   7
         Top             =   1200
         Width           =   840
      End
      Begin VB.Label lbl时间段 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "时间段"
         Height          =   210
         Left            =   270
         TabIndex        =   4
         Top             =   645
         Width           =   630
      End
      Begin VB.Label lbl号类 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "号类"
         Height          =   210
         Left            =   4185
         TabIndex        =   2
         Top             =   210
         Width           =   420
      End
      Begin VB.Label lblNodeNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "站点"
         Height          =   210
         Left            =   480
         TabIndex        =   0
         Top             =   210
         Width           =   420
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   378
      Left            =   5790
      TabIndex        =   26
      Top             =   4110
      Width           =   1100
   End
   Begin VB.CommandButton cmdSaveAdd 
      Caption         =   "保存新增(&O)"
      Height          =   378
      Left            =   2730
      TabIndex        =   25
      Top             =   4110
      Width           =   1335
   End
End
Attribute VB_Name = "frmClinicWorkTimeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const M_CurDate As String = "2016/01/01 "
Private mbytFun As G_Enum_Fun '0-查看,1-添加,2-调整
Private mstr站点 As String
Private mstr号类 As String
Private mstr时间段 As String

Private Enum mGridHeadCol
    COL_序号 = 0
    COL_开始时间
    COL_结束时间
End Enum
Private mrs时间段 As ADODB.Recordset
Private mblnOK As Boolean

Public Function ShowMe(frmParent As Form, ByVal bytFun As G_Enum_Fun, _
    Optional ByVal str站点 As String, Optional ByVal str号类 As String, _
    Optional ByVal str时间段 As String) As Boolean
    '程序入口
    '入参：
    '   frmParent - 父窗口
    '   bytFun - 操作类型, 0-查看，1-新增，2-修改
    mbytFun = bytFun
    mstr站点 = str站点: mstr号类 = str号类: mstr时间段 = str时间段
    
    Err = 0: On Error Resume Next
    mblnOK = False
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Sub cboNodeNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo号类_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo时间段_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "-" Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo时间段_Validate(Cancel As Boolean)
    If zlCommFun.ActualLen(cbo时间段.Text) > 6 Then
        MsgBox "时间段名称只允许输入6个字符或3个汉字！", vbInformation, gstrSysName
        zlControl.TxtSelAll cbo时间段
        Cancel = True
    End If
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdSaveExit_Click()
    On Error GoTo ErrHandler
    If IsValied() = False Then Exit Sub
    If SaveData() = False Then Exit Sub
    mblnOK = True
    
    Unload Me
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdSaveAdd_Click()
    On Error GoTo ErrHandler
    If IsValied() = False Then Exit Sub
    If SaveData() = False Then Exit Sub
    mblnOK = True
    
    '连续增加
    Call ClearFaceInfor
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub dtpDefaultTime_Change()
    dtpDefaultTime.Tag = Format(dtpDefaultTime.Value, "hh:mm:ss")
End Sub

Private Sub dtpDefaultTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpEndTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpPriorTime_Change()
    dtpPriorTime.Tag = Format(dtpPriorTime.Value, "hh:mm:ss")
End Sub

Private Sub dtpPriorTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpRestEndTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpRestStartTime_Change()
    dtpRestEndTime.Value = dtpRestStartTime.Value
End Sub

Private Sub dtpRestStartTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpStartTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    Dim i As Long, strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHandler
     Me.Caption = Choose(mbytFun + 1, "查看", "新增", "修改", "删除") & "上班时间"
    If InitGridHead() = False Then Unload Me: Exit Sub
    
    If mbytFun = Fun_Add Or mbytFun = Fun_Update Then
        If InitData() = False Then Unload Me: Exit Sub
    End If
    If mbytFun = Fun_Add Then
        cmdSaveAdd.Visible = True
        cmdSaveExit.Visible = True
        cmdOK.Visible = False
        Exit Sub
    Else
        cmdSaveAdd.Visible = False
        cmdSaveExit.Visible = False
        cmdOK.Visible = True
    End If
    
    If mbytFun = Fun_View Then '不允许编辑修改
        cmdCancel.Visible = False
        cmdOK.Left = cmdCancel.Left
        Call SetEnabled(Me.Controls, False)
    Else
        MsgBox "提醒：" & vbCrLf & _
               "    请不要轻易修改上班时间段，一但修改需要及时对所有使用了当前上班时间段且启用了分时段的安排进行重新划分时段，否则，可能会导致预约挂号出错！", vbInformation, gstrSysName
    End If
    
    '加载数据
    If LoadData(mstr站点, mstr号类, mstr时间段) = False Then Unload Me: Exit Sub
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Unload Me
End Sub

Private Function InitData() As Boolean
    Dim i As Long, strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHandler
    '加载站点数据
    strSQL = "Select 编号, 名称 From Zlnodelist"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cboNodeNo.Clear
    cboNodeNo.AddItem ""
    Do While Not rsTemp.EOF
        cboNodeNo.AddItem Nvl(rsTemp!编号) & "-" & Nvl(rsTemp!名称)
        rsTemp.MoveNext
    Loop
    
    '加载号类数据
    strSQL = "Select 编码, 名称, 简码, Nvl(缺省标志, 0) As 缺省标志 From 号类"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cbo号类.Clear
    cbo号类.AddItem ""
    Do While Not rsTemp.EOF
        cbo号类.AddItem Nvl(rsTemp!名称)
        'If Nvl(rsTemp!缺省标志) = 1 Then cbo号类.ListIndex = cbo号类.NewIndex
        rsTemp.MoveNext
    Loop
    
    '加载已有上班时间段，以便选择
    strSQL = "Select Distinct 时间段 From 时间段"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cbo时间段.Clear
    Do While Not rsTemp.EOF
        cbo时间段.AddItem Nvl(rsTemp!时间段)
        rsTemp.MoveNext
    Loop
    InitData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadData(ByVal str站点 As String, ByVal str号类 As String, _
    ByVal str时间段 As String, Optional ByVal blnDefault As Boolean) As Boolean
    '加载时间段数据
    '入参：blnDefault True-选择时间段缺省加载数据
    Dim i As Long
    Dim strSQL As String, strWhere As String, rs上班时间 As ADODB.Recordset
    Dim varTims As Variant, varRow As Variant
    
    Err = 0: On Error GoTo ErrHandler
    If Not blnDefault Then
        strWhere = " And Nvl(站点, '-') = Nvl([1], '-') And Nvl(号类, '-') = Nvl([2], '-')"
    End If
    strWhere = strWhere & " And a.时间段=[3]"
    
    strSQL = "Select a.时间段, a.号类, a.开始时间, a.终止时间, a.休息时段," & vbNewLine & _
            "        a.缺省时间, a.提前时间, a.出诊预留时间, " & vbNewLine & _
            "        b.编号, b.名称 As 站点" & vbNewLine & _
            " From 时间段 A, Zlnodelist B" & vbNewLine & _
            " Where a.站点 = b.编号(+)" & strWhere & vbNewLine & _
            " Order By Nvl(b.编号, -1), Nvl(a.号类, -1)"
    Set rs上班时间 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str站点, str号类, str时间段)
    If rs上班时间.EOF Then Exit Function
    
    If Not blnDefault Then
        zlControl.CboSetText cboNodeNo, Nvl(rs上班时间!站点)
        If cboNodeNo.ListIndex = -1 Then cboNodeNo.AddItem Nvl(rs上班时间!站点): cboNodeNo.ListIndex = cboNodeNo.NewIndex
        zlControl.CboSetText cbo号类, Nvl(rs上班时间!号类)
        If cbo号类.ListIndex = -1 Then cbo号类.AddItem Nvl(rs上班时间!号类): cbo号类.ListIndex = cbo号类.NewIndex
        zlControl.CboSetText cbo时间段, Nvl(rs上班时间!时间段)
        If cbo时间段.ListIndex = -1 Then cbo时间段.AddItem Nvl(rs上班时间!时间段): cbo时间段.ListIndex = cbo时间段.NewIndex
    End If
    
    dtpStartTime.Value = Nvl(rs上班时间!开始时间)
    dtpEndTime.Value = Nvl(rs上班时间!终止时间)
    dtpDefaultTime.Value = Nvl(rs上班时间!缺省时间, Nvl(rs上班时间!开始时间))
    dtpDefaultTime.Tag = dtpDefaultTime.Value
    dtpPriorTime.Value = Nvl(rs上班时间!提前时间, Nvl(rs上班时间!开始时间))
    dtpPriorTime.Tag = dtpPriorTime.Value
    txt预留时间.Text = Val(Nvl(rs上班时间!出诊预留时间, 0))
    
    vsf休息时段.Clear 1
    vsf休息时段.Rows = 1
    If Nvl(rs上班时间!休息时段) <> "" Then
        varTims = Split(Nvl(rs上班时间!休息时段), ";")
        For i = 0 To UBound(varTims)
            If varTims(i) <> "" Then
                varRow = Split(varTims(i), "-")
                vsf休息时段.AddItem CStr(i + 1) & vbTab & Format(varRow(0), "hh:mm:ss") & vbTab & Format(varRow(1), "hh:mm:ss")
            End If
        Next
        vsf休息时段.RowHeight(-1) = vsf休息时段.RowHeight(0)
    End If
    LoadData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cbo时间段_Click()
    Dim varStr As Variant, strTims As Variant
    Dim i As Long, Row As Long
    
    Err = 0: On Error GoTo ErrHandler
    Call LoadData("", "", cbo时间段.Text, True)
    dtpRestStartTime.Value = dtpStartTime.Value
    dtpRestEndTime.Value = dtpStartTime.Value
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Err = 0: On Error GoTo ErrHandler
    If mbytFun = Fun_View Then Unload Me: Exit Sub
    
    cmdOK.Enabled = False
    If IsValied() = False Then cmdOK.Enabled = True: Exit Sub
    If SaveData() = False Then cmdOK.Enabled = True: Exit Sub
    mblnOK = True
    
    '连续增加
    If mbytFun = Fun_Add Then
        cmdOK.Enabled = True
        Exit Sub
    Else
        MsgBox "注意：" & vbCrLf & _
               "    上班时间段修改成功，请及时对所有使用了当前上班时间段且启用了分时段的安排进行重新划分时段，否则，可能会导致预约挂号出错！", vbExclamation, gstrSysName
    End If
    
    Unload Me
    Exit Sub
ErrHandler:
    cmdOK.Enabled = True
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ClearFaceInfor()
    '功能:清除界面信息，以便重新输入数据
    On Error GoTo errHandle
    cboNodeNo.ListIndex = -1
    cbo号类.ListIndex = -1
    cbo时间段.Text = "": cbo时间段.ListIndex = -1
    
    dtpStartTime.Value = "00:00:00": dtpEndTime.Value = "00:00:00"
    dtpDefaultTime.Value = "00:00:00": dtpDefaultTime.Tag = ""
    dtpPriorTime.Value = "00:00:00": dtpPriorTime.Tag = ""
    txt预留时间.Text = 0
    
    vsf休息时段.Clear 1
    vsf休息时段.Rows = 1
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function SaveData() As Boolean
    Dim str休息时段 As String, i  As Integer
    Dim strSQL As String, strColor As String
    Dim dtStartTime As Date, dtEndTime As Date
    Dim dtRestStartTime As Date, dtRestEndTime As Date
    Dim dtPriorTime As Date, dtDefaultTime As Date
    
    Err = 0: On Error GoTo ErrHandler
    If mbytFun <> Fun_Delete Then
        For i = 1 To vsf休息时段.Rows - 1
            str休息时段 = str休息时段 & ";" & vsf休息时段.TextMatrix(i, COL_开始时间) & "-" & vsf休息时段.TextMatrix(i, COL_结束时间)
        Next
        If str休息时段 <> "" Then str休息时段 = Mid(str休息时段, 2)
        
        Call FormatTime(0, dtStartTime, dtEndTime, dtRestStartTime, dtRestEndTime, dtPriorTime, dtDefaultTime)
    End If
    
    'CREATE OR REPLACE Procedure Zl_上班时段_Modify
    '(
    '  操作类型_In     Number,
    '  站点_In         时间段.站点%Type,
    '  号类_In         时间段.号类%Type,
    '  时间段_In       时间段.时间段%Type,
    '  开始时间_In     时间段.开始时间%Type := Null,
    '  终止时间_In     时间段.终止时间%Type := Null,
    '  休息时段_In     时间段.休息时段%Type := Null,
    '  缺省时间_In     时间段.缺省时间%Type := Null,
    '  提前时间_In     时间段.提前时间%Type := Null,
    '  出诊预留时间_In 时间段.出诊预留时间%Type := 0,
    '  原站点_In       时间段.站点%Type := Null,
    '  原号类_In       时间段.号类%Type := Null,
    '  原时间段_In     时间段.时间段%Type := Null
    ') As
    '  --操作类型_In 0-新增，1-修改，2-删除
    Select Case mbytFun
    Case Fun_Add
        strSQL = "Zl_上班时段_Modify("
        strSQL = strSQL & "" & 0 & ","
        strSQL = strSQL & "'" & NeedCode(cboNodeNo.Text) & "',"
        strSQL = strSQL & "'" & cbo号类.Text & "',"
        strSQL = strSQL & "'" & cbo时间段.Text & "',"
        strSQL = strSQL & "To_Date('" & Format(dtStartTime, "hh:mm:ss") & "','hh24:mi:ss'),"
        strSQL = strSQL & "To_Date('" & Format(dtEndTime, "hh:mm:ss") & "','hh24:mi:ss'),"
        strSQL = strSQL & "'" & str休息时段 & "',"
        strSQL = strSQL & "To_Date('" & Format(dtDefaultTime, "hh:mm:ss") & "','hh24:mi:ss'),"
        strSQL = strSQL & "To_Date('" & Format(dtPriorTime, "hh:mm:ss") & "','hh24:mi:ss'),"
        strSQL = strSQL & "" & Val(txt预留时间.Text) & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Case Fun_Update
        strSQL = "Zl_上班时段_Modify("
        strSQL = strSQL & "" & 1 & ","
        strSQL = strSQL & "'" & NeedCode(cboNodeNo.Text) & "',"
        strSQL = strSQL & "'" & cbo号类.Text & "',"
        strSQL = strSQL & "'" & cbo时间段.Text & "',"
        strSQL = strSQL & "To_Date('" & Format(dtStartTime, "hh:mm:ss") & "','hh24:mi:ss'),"
        strSQL = strSQL & "To_Date('" & Format(dtEndTime, "hh:mm:ss") & "','hh24:mi:ss'),"
        strSQL = strSQL & "'" & str休息时段 & "',"
        strSQL = strSQL & "To_Date('" & Format(dtDefaultTime, "hh:mm:ss") & "','hh24:mi:ss'),"
        strSQL = strSQL & "To_Date('" & Format(dtPriorTime, "hh:mm:ss") & "','hh24:mi:ss'),"
        strSQL = strSQL & "" & Val(txt预留时间.Text) & ","
        strSQL = strSQL & "'" & mstr站点 & "',"
        strSQL = strSQL & "'" & mstr号类 & "',"
        strSQL = strSQL & "'" & mstr时间段 & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End Select
    SaveData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function InitGridHead() As Boolean
    Dim strHead As String
    Dim i As Long, varData As Variant

    Err = 0: On Error GoTo ErrHandler
    strHead = "序号,4,500|开始时间,4,1300|结束时间,4,1300"
    With vsf休息时段
        .Redraw = False
        .FixedCols = 1: .FixedRows = 1
        .HighLight = flexHighlightWithFocus
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .RowHeight(-1) = 280
        
        .Rows = 1
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .Redraw = True
    End With
    InitGridHead = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdAddRestTime_Click()
    Dim i As Long
    Dim dtRestStartTime As Date, dtRestEndTime As Date
    Dim dtStartTime As Date, dtEndTime As Date
    Dim dtTempStart As Date, dtTempEnd As Date
    
    Err = 0: On Error GoTo ErrHandler
    Call FormatTime(1, dtStartTime, dtEndTime, dtRestStartTime, dtRestEndTime)
    If dtRestStartTime >= dtRestEndTime Then
        MsgBox "休息时间的结束时间必须大于开始时间！", vbInformation, gstrSysName
        If dtpRestEndTime.Visible And dtpRestEndTime.Enabled Then dtpRestEndTime.SetFocus
        Exit Sub
    End If
    If Not ((dtRestStartTime >= dtStartTime And dtRestStartTime <= dtEndTime) _
            And (dtRestEndTime >= dtStartTime And dtRestEndTime <= dtEndTime)) Then
        MsgBox "休息时间必须在上班时间(" & Format(dtStartTime, "hh:mm:ss") & "-" & Format(dtEndTime, "hh:mm:ss") & ")范围内！", vbInformation, gstrSysName
        If dtpRestStartTime.Visible And dtpRestStartTime.Enabled Then dtpRestStartTime.SetFocus
        Exit Sub
    End If
    
    For i = 1 To vsf休息时段.Rows - 1
        dtTempStart = M_CurDate & vsf休息时段.TextMatrix(i, COL_开始时间)
        dtTempEnd = M_CurDate & vsf休息时段.TextMatrix(i, COL_结束时间)
        If dtTempEnd <= dtTempStart Then dtTempEnd = DateAdd("d", 1, dtTempEnd)
        
        If Not ((dtRestStartTime < dtTempStart And dtRestEndTime < dtTempStart) _
                Or (dtRestStartTime > dtTempEnd And dtRestEndTime > dtTempEnd)) Then
            MsgBox "休息时间不能包含在已设置休息时间范围内！", vbInformation, gstrSysName
            If dtpRestStartTime.Visible And dtpRestStartTime.Enabled Then dtpRestStartTime.SetFocus
            Exit Sub
        End If
    Next
    vsf休息时段.AddItem CStr(vsf休息时段.Rows) & vbTab & Format(dtRestStartTime, "hh:mm:ss") & vbTab & Format(dtRestEndTime, "hh:mm:ss")
    vsf休息时段.RowHeight(-1) = vsf休息时段.RowHeight(0)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub FormatTime(ByVal bytType As Byte, ByRef dtStartTime As Date, ByRef dtEndTime As Date, _
    Optional ByRef dtRestStartTime As Date, Optional ByRef dtRestEndTime As Date, _
    Optional ByRef dtPriorTime As Date, Optional ByRef dtDefaultTime As Date)
    Dim blnChanged As Boolean
    
    '格式化时间
    dtStartTime = M_CurDate & Format(dtpStartTime.Value, "hh:mm:ss")
    dtEndTime = M_CurDate & Format(dtpEndTime.Value, "hh:mm:ss")
    dtRestStartTime = M_CurDate & Format(dtpRestStartTime.Value, "hh:mm:ss")
    dtRestEndTime = M_CurDate & Format(dtpRestEndTime.Value, "hh:mm:ss")
    
    dtPriorTime = M_CurDate & Format(dtpPriorTime.Value, "hh:mm:ss")
    dtDefaultTime = M_CurDate & Format(dtpDefaultTime.Value, "hh:mm:ss")
    
    If bytType = 1 Then
        blnChanged = False
        If dtEndTime <= dtStartTime Then
            blnChanged = True
            dtEndTime = DateAdd("d", 1, dtEndTime) '开始时间大于结束时间，则结束时间加一天
        End If
        If dtRestEndTime <= dtRestStartTime And blnChanged Then dtRestEndTime = DateAdd("d", 1, dtRestEndTime) '休息开始时间大于休息结束时间，则休息结束时间加一天
        If dtRestStartTime < dtStartTime Then dtRestStartTime = DateAdd("d", 1, dtRestStartTime) '开始时间大于休息开始时间，则休息开始时间加一天
        If dtRestEndTime < dtStartTime Then dtRestEndTime = DateAdd("d", 1, dtRestEndTime) '开始时间大于休息结束时间，则休息结束时间加一天
        If dtDefaultTime < dtStartTime Then dtDefaultTime = DateAdd("d", 1, dtDefaultTime) '开始时间大于缺省预约时间，则缺省预约时间加一天
        If dtPriorTime > dtStartTime Then dtPriorTime = DateAdd("d", -1, dtPriorTime) '开始时间小于提前挂号时间，则提前挂号时间减一天
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mrs时间段 Is Nothing Then Set mrs时间段 = Nothing
End Sub

Private Sub txt预留时间_GotFocus()
    zlControl.TxtSelAll txt预留时间
End Sub

Private Sub txt预留时间_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub vsf休息时段_EnterCell()
    cmdDelete.Enabled = vsf休息时段.Row > 0
End Sub

Private Sub cmdDelete_Click()
    Dim i As Integer
    
    Err = 0: On Error GoTo ErrHandler
    If vsf休息时段.Row > 0 Then
        If MsgBox("您确定要删除第 " & vsf休息时段.Row & " 行？", vbQuestion + vbOKCancel + vbDefaultButton2, gstrSysName) = vbOK Then
            vsf休息时段.RemoveItem vsf休息时段.Row
            For i = 1 To vsf休息时段.Rows - 1 '重新编号
                vsf休息时段.TextMatrix(i, 0) = i
            Next
        End If
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function IsValied() As Boolean
    Dim dtRestStartTime As Date, dtRestEndTime As Date
    Dim dtStartTime As Date, dtEndTime As Date
    Dim dtPriorTime As Date, dtDefaultTime As Date
    Dim dtTempStart As Date, dtTempEnd As Date
    Dim i As Integer, lngMinute As Long
    
    Err = 0: On Error GoTo ErrHandler
    If zlControl.FormCheckInput(Me) = False Then Exit Function
    If cbo时间段.Text = "" Then
        MsgBox "时间段不能为空！", vbInformation, gstrSysName
        If cbo时间段.Visible And cbo时间段.Enabled Then cbo时间段.SetFocus
        Exit Function
    End If
    If zlCommFun.ActualLen(cbo时间段.Text) > 6 Then
        MsgBox "时间段名称只允许输入6个字符或3个汉字！", vbInformation, gstrSysName
        If cbo时间段.Visible And cbo时间段.Enabled Then cbo时间段.SetFocus
        zlControl.TxtSelAll cbo时间段
        Exit Function
    End If
    If IsNumeric(Val(txt预留时间.Text)) = False Then
        MsgBox "预留时间只能为数字！", vbInformation, gstrSysName
        If txt预留时间.Visible And txt预留时间.Enabled Then txt预留时间.SetFocus
        zlControl.TxtSelAll txt预留时间
        Exit Function
    End If
    If mbytFun = Fun_Add Then
        If CheckExist(NeedCode(cboNodeNo.Text), cbo号类.Text, cbo时间段.Text) Then
            MsgBox NeedName(cboNodeNo.Text) & "已存在" & IIf(cbo号类.Text = "", "不分号类", "号类为“" & cbo号类.Text & "”") & "的“" & cbo时间段.Text & "”时间段！", vbInformation, gstrSysName
            If cbo时间段.Visible And cbo时间段.Enabled Then cbo时间段.SetFocus
            zlControl.TxtSelAll cbo时间段
            Exit Function
        End If
    ElseIf mbytFun = Fun_Update Then
        If mstr站点 <> NeedCode(cboNodeNo.Text) Or mstr号类 <> cbo号类.Text Or mstr时间段 <> cbo时间段.Text Then
            If CheckHaveUsed(mstr站点, mstr号类, mstr时间段) Then
                MsgBox "当前上班时间段已被使用，不能修改站点、号类及时间段名称！", vbInformation, gstrSysName
                If cbo时间段.Visible And cbo时间段.Enabled Then cbo时间段.SetFocus
                zlControl.TxtSelAll cbo时间段
                Exit Function
            End If
            If CheckExist(NeedCode(cboNodeNo.Text), cbo号类.Text, cbo时间段.Text) Then
                MsgBox NeedName(cboNodeNo.Text) & "已存在" & IIf(cbo号类.Text = "", "不分号类", "号类为“" & cbo号类.Text & "”") & "的“" & cbo时间段.Text & "”时间段！", vbInformation, gstrSysName
                If cbo时间段.Visible And cbo时间段.Enabled Then cbo时间段.SetFocus
                zlControl.TxtSelAll cbo时间段
                Exit Function
            End If
        End If
    End If
    
    Call FormatTime(1, dtStartTime, dtEndTime, dtRestStartTime, dtRestEndTime, dtPriorTime, dtDefaultTime)
    If dtPriorTime > dtStartTime Then
        MsgBox "提前挂号时间必须小于等于开始时间！", vbInformation, gstrSysName
        If dtpPriorTime.Visible And dtpPriorTime.Enabled Then dtpPriorTime.SetFocus
        Exit Function
    End If
    
    If dtDefaultTime < dtStartTime Or dtDefaultTime > dtEndTime Then
        MsgBox "缺省预约时间必须在上班时间(" & Format(dtStartTime, "hh:mm:ss") & "-" & Format(dtEndTime, "hh:mm:ss") & ")范围内！", vbInformation, gstrSysName
        If dtpDefaultTime.Visible And dtpDefaultTime.Enabled Then dtpDefaultTime.SetFocus
        Exit Function
    End If
    lngMinute = DateDiff("n", dtStartTime, dtEndTime)
    
    For i = 1 To vsf休息时段.Rows - 1
        dtTempStart = M_CurDate & vsf休息时段.TextMatrix(i, COL_开始时间)
        dtTempEnd = M_CurDate & vsf休息时段.TextMatrix(i, COL_结束时间)
        If dtTempEnd <= dtTempStart Then dtTempEnd = DateAdd("d", 1, dtTempEnd) '休息开始时间大于休息结束时间，则休息结束时间加一天
        If dtTempStart < dtStartTime Then dtTempStart = DateAdd("d", 1, dtTempStart) '开始时间大于休息开始时间，则休息开始时间加一天
        If dtTempEnd < dtStartTime Then dtTempEnd = DateAdd("d", 1, dtTempEnd) '开始时间大于休息结束时间，则休息结束时间加一天

        If Not ((dtTempStart >= dtStartTime And dtTempStart <= dtEndTime) _
            And (dtTempEnd >= dtStartTime And dtTempEnd <= dtEndTime)) Then
            MsgBox "第" & i & "行休息时间不在上班时间(" & Format(dtStartTime, "hh:mm:ss") & "-" & Format(dtEndTime, "hh:mm:ss") & ")范围内！", vbInformation, gstrSysName
            vsf休息时段.Row = i
            Exit Function
        End If
        lngMinute = lngMinute - DateDiff("n", dtTempStart, dtTempEnd)
    Next
    
    '预留时间不能大于总的分钟数
    If Val(txt预留时间.Text) > lngMinute Then
        MsgBox "预留时间不能大于上班时段的总时间！", vbInformation, gstrSysName
        If txt预留时间.Visible And txt预留时间.Enabled Then txt预留时间.SetFocus
        zlControl.TxtSelAll txt预留时间
        Exit Function
    End If
    
    IsValied = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckExist(ByVal str站点 As String, ByVal str号类 As String, ByVal str时间段 As String) As Boolean
    '检查记录是否已存在
    Dim strSQL As String, rs上班时间 As ADODB.Recordset
    Dim varTims As Variant, varRow As Variant
    
    Err = 0: On Error GoTo ErrHandler
    strSQL = "Select 1 From 时间段 A, Zlnodelist B" & vbNewLine & _
            " Where a.站点 = b.编号(+)" & vbNewLine & _
            "       And Nvl(站点, '-') = Nvl([1], '-') And Nvl(号类, '-') = Nvl([2], '-') And 时间段 = [3]"
    Set rs上班时间 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str站点, str号类, str时间段)
    CheckExist = Not rs上班时间.EOF
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckHaveUsed(ByVal str站点 As String, ByVal str号类 As String, ByVal str时间段 As String) As Boolean
    '检查当前上班时间段是否已被使用
    Dim strSQL As String, rs上班时间 As ADODB.Recordset
    Dim varTims As Variant, varRow As Variant
    
    Err = 0: On Error GoTo ErrHandler
    '检查原上班时段是否被使用，被使用的不能修改站点、号类、时间段
    '不能删除被使用的范围最广的那一个,被使用的时段只要有一个即可（不同站点，不同号类可能会有多个同名的时间段）
    '临床出诊号源限制
    strSQL = "Select 1" & vbNewLine & _
            " From (Select b.上班时段, c.站点, a.号类," & vbNewLine & _
            "              Row_Number() Over(Partition By b.上班时段 Order By b.上班时段, c.站点 Desc, a.号类 Desc) As 组号" & vbNewLine & _
            "        From 临床出诊号源 A, 临床出诊号源限制 B, 部门表 C" & vbNewLine & _
            "        Where a.Id = b.号源id And a.科室id = c.Id)" & vbNewLine & _
            " Where 组号 = 1 And Nvl(站点, '-') = Nvl([1], '-') And Nvl(号类, '-') = Nvl([2], '-') And 上班时段 = [3] And Rownum < 2"
    Set rs上班时间 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str站点, str号类, str时间段)
    If Not rs上班时间 Is Nothing Then
        If Not rs上班时间.EOF Then CheckHaveUsed = True: Exit Function
    End If
    
    '临床出诊限制(固定规则、模板)
    strSQL = "Select 1" & vbNewLine & _
            " From (Select a.上班时段, c.站点, b.号类," & vbNewLine & _
            "              Row_Number() Over(Partition By a.上班时段 Order By a.上班时段, c.站点 Desc, b.号类 Desc) As 组号" & vbNewLine & _
            "        From 临床出诊限制 A, 临床出诊安排 D, 临床出诊号源 B, 部门表 C" & vbNewLine & _
            "        Where a.安排id = d.Id And d.号源id = b.Id And b.科室id = c.Id)" & vbNewLine & _
            " Where 组号 = 1 And Nvl(站点, '-') = Nvl([1], '-') And Nvl(号类, '-') = Nvl([2], '-') And 上班时段 = [3] And Rownum < 2"
    Set rs上班时间 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str站点, str号类, str时间段)
    If Not rs上班时间 Is Nothing Then
        If Not rs上班时间.EOF Then CheckHaveUsed = True: Exit Function
    End If
    
    '临床出诊记录
    '不检查，因为该表太大，其次上班时段的信息都保存在了这个表中，没有找到上班时段时可由这个表的数据来提取
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub vsf休息时段_GotFocus()
    If vsf休息时段.Rows > 1 Then
        vsf休息时段.Row = 1
    Else
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub vsf休息时段_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub
