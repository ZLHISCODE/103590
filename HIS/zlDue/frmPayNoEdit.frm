VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPayNoEdit 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPayNO 
      AutoRedraw      =   -1  'True
      Height          =   6030
      Left            =   0
      ScaleHeight     =   5970
      ScaleWidth      =   9420
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   9480
      Begin VB.PictureBox picDown 
         BorderStyle     =   0  'None
         Height          =   1305
         Left            =   -60
         ScaleHeight     =   1305
         ScaleWidth      =   9930
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   4785
         Width           =   9930
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   0
            Left            =   810
            TabIndex        =   11
            Top             =   135
            Width           =   8820
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   1
            Left            =   810
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   510
            Width           =   3240
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   2
            Left            =   6390
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   510
            Width           =   3240
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   3
            Left            =   825
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   870
            Width           =   3240
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   4
            Left            =   6390
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   885
            Width           =   3240
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "付款说明"
            Height          =   180
            Index           =   4
            Left            =   0
            TabIndex        =   10
            Top             =   195
            Width           =   750
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "填制人"
            Height          =   180
            Index           =   5
            Left            =   180
            TabIndex        =   12
            Top             =   570
            Width           =   570
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "填制日期"
            Height          =   180
            Index           =   6
            Left            =   5580
            TabIndex        =   14
            Top             =   570
            Width           =   750
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "审核人"
            Height          =   180
            Index           =   7
            Left            =   180
            TabIndex        =   16
            Top             =   945
            Width           =   570
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "审核日期"
            Height          =   180
            Index           =   8
            Left            =   5580
            TabIndex        =   18
            Top             =   945
            Width           =   750
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsPayEdit 
         Height          =   2610
         Left            =   165
         TabIndex        =   6
         Top             =   1500
         Width           =   5055
         _cx             =   8916
         _cy             =   4604
         Appearance      =   0
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483644
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   12632256
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPayNoEdit.frx":0000
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
         ExplorerBar     =   7
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   -1  'True
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
      Begin VSFlex8Ctl.VSFlexGrid vs冲预付 
         Height          =   2625
         Left            =   5205
         TabIndex        =   8
         Top             =   1500
         Width           =   4605
         _cx             =   8123
         _cy             =   4630
         Appearance      =   0
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483644
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   12632256
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPayNoEdit.frx":00A8
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
         ExplorerBar     =   7
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   -1  'True
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
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "付款通知单"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   30
         TabIndex        =   22
         Top             =   90
         Width           =   9780
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "本次付款:"
         Height          =   180
         Index           =   4
         Left            =   7950
         TabIndex        =   5
         Top             =   1260
         Width           =   810
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   8355
         TabIndex        =   21
         Top             =   390
         Width           =   1425
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "NO."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Index           =   10
         Left            =   8055
         TabIndex        =   20
         Top             =   450
         Width           =   315
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "税务登记号:"
         Height          =   180
         Index           =   3
         Left            =   210
         TabIndex        =   4
         Top             =   1290
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "开户银行:"
         Height          =   180
         Index           =   2
         Left            =   405
         TabIndex        =   3
         Top             =   1050
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "地址电话:"
         Height          =   180
         Index           =   1
         Left            =   390
         TabIndex        =   2
         Top             =   825
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "单位名称:"
         Height          =   180
         Index           =   7
         Left            =   390
         TabIndex        =   1
         Top             =   600
         Width           =   810
      End
      Begin VB.Label lblEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "本次冲预付款:"
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   5
         Left            =   5205
         TabIndex        =   9
         Top             =   4110
         Width           =   4605
      End
      Begin VB.Label lblEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "合计:"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   6
         Left            =   165
         TabIndex        =   7
         Top             =   4095
         Width           =   5055
      End
   End
End
Attribute VB_Name = "frmPayNoEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '
Private mlngModule As Long
Private mstrPrivs As String
Private mfrmMain As Form
Private mblnChange As Boolean

Private mEditType As gEditType  '编辑类型
Private mstrNo As String        '单据号
Private mint记录状态 As Integer '记录状态
Private mlng付款序号 As Long    '付款序号
Private mlng单位ID As Long      '单位ID
Private mdbl本次应付 As Double, mdbl本次预交 As Double

Private mblnEdit As Boolean     '是否允许编辑
Private Enum mlblIdx
    idx_lbl地址电话 = 1
    idx_lbl开户银行 = 2
    idx_lbl税务登记号 = 3
    idx_lbl本次付款 = 4
    idx_lbl冲预交合计 = 5
    idx_lbl付款合计 = 6
    idx_lbl单位名称 = 7
End Enum
Private mrs结算方式 As ADODB.Recordset
Private mint标记 As Integer

Public Event initCard(ByVal lng付款序号 As Long, ByVal lng单位ID As Long, ByVal str单位名称 As String)
Public Event zlChangeData(ByVal blnChange As Boolean)

Private Sub SetEditPro()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:设置编辑属性
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Integer
    For i = 0 To 4
        txtInfo(i).Enabled = mblnEdit And i = 0
    Next
    If mEditType = g审核 Then
        vsPayEdit.Editable = flexEDKbdMouse
    Else
        vsPayEdit.Editable = IIf(mblnEdit, flexEDKbdMouse, flexEDNone)
    End If
End Sub

Private Sub InitvsPayEdit()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化网格控件的默认属性
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-08-19 11:55:12
    '-----------------------------------------------------------------------------------------------------------
    'Dim rsTemp As New ADODB.Recordset
    '问题27930 by lesfeng 2010-03-23
    If mint标记 = 0 Then
        gstrSQL = "Select 结算方式 From 结算方式应用 Where 应用场合='付货款' Order by 缺省标志 desc"
    Else
        gstrSQL = "Select '标记' As 结算方式 From dual "
'        gstrSQL = "Select '  ' As 结算方式 From dual Union All " & _
'                  "Select 结算方式 From 结算方式应用 Where 应用场合='付货款'"
    End If
    On Error GoTo errHandle
    Set mrs结算方式 = New ADODB.Recordset
    zlDatabase.OpenRecordset mrs结算方式, gstrSQL, Me.Caption
    With vsPayEdit
        .ColComboList(.ColIndex("付款方式")) = .BuildComboList(mrs结算方式, "结算方式", "结算方式")
    End With
    Call vs冲预付_LostFocus
    Call vsPayEdit_LostFocus
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub picDown_Resize()
    Err = 0: On Error Resume Next
    With picDown
        txtInfo(0).Width = .ScaleWidth - txtInfo(0).Left
        txtInfo(2).Left = .ScaleWidth - txtInfo(2).Width
        txtInfo(4).Left = txtInfo(2).Left
        lblInfo(6).Left = txtInfo(2).Left - lblInfo(6).Width
        lblInfo(8).Left = lblInfo(6).Left
    End With
End Sub

Private Sub txtInfo_Change(Index As Integer)
    mblnChange = True
    RaiseEvent zlChangeData(mblnChange)
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txtInfo(Index)
End Sub

Private Sub txtInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    mblnEdit = False
    '问题27930 by lesfeng 2010-03-23
'    lblTitle.Caption = GetUnitName & lblTitle.Caption
    RestoreWinState Me, App.ProductName
    zl_vsGrid_Para_Restore mlngModule, vsPayEdit, Me.Caption, "付款列表"
    zl_vsGrid_Para_Restore mlngModule, vs冲预付, Me.Caption, "冲预付列表"
'    Call InitvsPayEdit
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With picPayNO
        .Left = ScaleLeft
        .Width = ScaleWidth
        .Top = ScaleTop
        .Height = ScaleHeight
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsPayEdit, Me.Caption, "付款列表"
    zl_vsGrid_Para_Save mlngModule, vs冲预付, Me.Caption, "冲预付列表"
End Sub

Private Sub picPayNO_Resize()
    Err = 0: On Error Resume Next
    With picPayNO
        txtNo.Left = .ScaleWidth - txtNo.Width - 50
        lblInfo(10).Left = txtNo.Left - lblInfo(10).Width
        lblTitle.Left = .ScaleLeft
        lblTitle.Width = .ScaleWidth
        picDown.Top = .ScaleHeight - picDown.Height
        picDown.Width = .ScaleWidth
        picDown.Left = .ScaleLeft
        '问题27930 by lesfeng 2010-03-23
        If mint标记 = 0 Then
            lblEdit(mlblIdx.idx_lbl付款合计).Top = picDown.Top - lblEdit(mlblIdx.idx_lbl付款合计).Height - 30
            lblEdit(mlblIdx.idx_lbl冲预交合计).Top = lblEdit(mlblIdx.idx_lbl付款合计).Top
            lblEdit(mlblIdx.idx_lbl冲预交合计).Width = .ScaleWidth - lblEdit(mlblIdx.idx_lbl冲预交合计).Left - 100
            lblEdit(mlblIdx.idx_lbl冲预交合计).Height = lblEdit(mlblIdx.idx_lbl付款合计).Height
            vs冲预付.Width = .ScaleWidth - vs冲预付.Left - 100
            vs冲预付.Height = lblEdit(mlblIdx.idx_lbl冲预交合计).Top - vs冲预付.Top + 10
            
            vsPayEdit.Top = vs冲预付.Top
            vsPayEdit.Height = vs冲预付.Height
            lblEdit(mlblIdx.idx_lbl本次付款).Left = .ScaleWidth - lblEdit(mlblIdx.idx_lbl本次付款).Width - 50
            
            lblTitle.Caption = GetUnitName & lblTitle.Caption
        Else
            lblEdit(mlblIdx.idx_lbl付款合计).Top = picDown.Top - lblEdit(mlblIdx.idx_lbl付款合计).Height - 30
            lblEdit(mlblIdx.idx_lbl付款合计).Width = .ScaleWidth - lblEdit(mlblIdx.idx_lbl付款合计).Left - 100
            
            vsPayEdit.Width = .ScaleWidth - vsPayEdit.Left - 100
            vsPayEdit.Height = lblEdit(mlblIdx.idx_lbl付款合计).Top - vsPayEdit.Top + 10
            
            lblEdit(mlblIdx.idx_lbl本次付款).Left = .ScaleWidth - lblEdit(mlblIdx.idx_lbl本次付款).Width - 50
            
            lblTitle.Caption = GetUnitName & "标记付款单"
            vs冲预付.Visible = False
            lblEdit(mlblIdx.idx_lbl冲预交合计).Visible = False
            lblEdit(mlblIdx.idx_lbl本次付款).Caption = "本次标记付款:"
            lblInfo(4).Caption = "标记说明"
        End If
    End With
End Sub

Private Function initCard(ByRef intErrInfor As Integer) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化卡片信息
    '入参:
    '出参:intErrInfor-返回错误信息代码(1-已经删除,2-已经审核)
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-08-19 13:00:28
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim rsTemp As New Recordset
    '初始表格
    With vsPayEdit
        .Clear 1
        .Rows = 2
        '问题27930 by lesfeng 2010-03-23
        If mint标记 = 1 Then
            .ColHidden(.ColIndex("付款方式")) = True: .ColWidth(.ColIndex("付款方式")) = 0
        Else
            .ColHidden(.ColIndex("付款方式")) = False: .ColWidth(.ColIndex("付款方式")) = 1200
        End If
    End With
    With vs冲预付
        .Clear 1
        .Rows = 2
    End With
    On Error GoTo errHandle
    Select Case mEditType
        Case g新增
                txtInfo(1).Text = UserInfo.姓名
                txtInfo(2).Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss")
                txtInfo(3).Text = ""
                txtInfo(4).Text = ""
        Case g审核, g修改, g查看, g取消, g预审
            '读取付款序号
            'by lesfeng 2009-12-2 性能优化 '问题27930 by lesfeng 2010-03-23
            gstrSQL = "Select ID,记录状态,NO,序号,预付款,单位ID,金额,结算方式,结算号码,摘要,填制人,填制日期,审核人,审核日期,付款序号," & _
                      "       decode(拒付标志,0,'正常',1,'标记','标记') as 标记付款" & _
                      " From 付款记录 Where NO=[1] And 记录状态=[2] order by 序号"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrNo, mint记录状态)
            
            If rsTemp.EOF Then
                intErrInfor = 1
                Exit Function
            End If
            
            mlng付款序号 = Nvl(rsTemp!付款序号, 0)
            mlng单位ID = Nvl(rsTemp!单位ID, 0)
            
            txtInfo(0).Text = Nvl(rsTemp!摘要)
            txtInfo(1).Text = Nvl(rsTemp!填制人)
            txtInfo(2).Text = Format(rsTemp!填制日期, "yyyy-MM-dd hh:mm:ss")
            txtInfo(3).Text = Nvl(rsTemp!审核人)
            txtInfo(4).Text = Format(rsTemp!审核日期, "yyyy-MM-dd hh:mm:ss")
            txtNo = Nvl(rsTemp!NO)
            txtNo.Tag = Nvl(rsTemp!NO)
            If Nvl(rsTemp!审核人) <> "" And mEditType = g审核 Then
                intErrInfor = 2
                Exit Function
            End If
            
            If mEditType = g审核 Or mEditType = g取消 Then
                txtInfo(3).Text = UserInfo.姓名
                txtInfo(4).Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss")
            End If
            
            With vsPayEdit
                .Rows = rsTemp.RecordCount + 1
                i = 1
                Do While Not rsTemp.EOF
                    .TextMatrix(i, .ColIndex("标记付款")) = Nvl(rsTemp!标记付款)
                    .TextMatrix(i, .ColIndex("付款方式")) = Nvl(rsTemp!结算方式)
                    .Cell(flexcpData, i, .ColIndex("付款方式")) = Nvl(rsTemp!ID)
                    .TextMatrix(i, .ColIndex("付款金额")) = Format(Val(Nvl(rsTemp!金额)), gVbFmtString.FM_金额)
                    .TextMatrix(i, .ColIndex("结算号码")) = Nvl(rsTemp!结算号码)
                    .Cell(flexcpData, i, .ColIndex("结算号码")) = Nvl(rsTemp!结算号码)
                    i = i + 1
                    rsTemp.MoveNext
                Loop
            End With
            '读取预付记录
    End Select
    
    Call zlLoadPrivder(mlng单位ID)
    RaiseEvent initCard(mlng付款序号, mlng单位ID, lblEdit(mlblIdx.idx_lbl单位名称).Tag)
    Call SetEditPro
    Call 付款合计
    initCard = True
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Public Function zlLoadPrivder(ByVal lng单位ID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:加载供应商信息
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-08-19 13:06:23
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    '如果提供了供应商ID则读取该供应商信息
    On Error GoTo errHandle
    gstrSQL = "Select 名称,地址,电话,开户银行,税务登记号 From 供应商 Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng单位ID)
    mlng单位ID = lng单位ID
    If Not rsTemp.EOF Then
        lblEdit(mlblIdx.idx_lbl单位名称).Caption = "单位名称:" & rsTemp!名称
        lblEdit(mlblIdx.idx_lbl单位名称).Tag = Nvl(rsTemp!名称)
        lblEdit(mlblIdx.idx_lbl地址电话).Caption = "地址电话:" & Nvl(rsTemp!地址) & Nvl(rsTemp!地址)
        lblEdit(mlblIdx.idx_lbl开户银行).Caption = "开户银行:" & Nvl(rsTemp!开户银行)
        lblEdit(mlblIdx.idx_lbl税务登记号).Caption = "税务登记号:" & Nvl(rsTemp!税务登记号)
    Else
        lblEdit(mlblIdx.idx_lbl单位名称).Caption = "单位名称:"
        lblEdit(mlblIdx.idx_lbl单位名称).Tag = ""
        lblEdit(mlblIdx.idx_lbl地址电话).Caption = "地址电话:"
        lblEdit(mlblIdx.idx_lbl开户银行).Caption = "开户银行:"
        lblEdit(mlblIdx.idx_lbl税务登记号).Caption = "税务登记号:"
        zlLoadPrivder = False: Exit Function
    End If
    zlLoadPrivder = True
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Public Sub zlInitPara(ByVal FrmMain As Form, ByVal lngModuel As Long, ByVal strPrivs As String, ByVal int标记 As Integer)
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化参数设置(第一步必需初始化)
    '入参:frmMain-调用的主窗口
    '     lngModuel-模块号
    '     strPrivs-权限串
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-08-19 11:47:39
    '-----------------------------------------------------------------------------------------------------------
    Set mfrmMain = FrmMain: mlngModule = lngModuel: mstrPrivs = strPrivs: mint标记 = int标记
    Call InitvsPayEdit
End Sub

Public Function zlLoadData(ByVal EditType As gEditType, ByVal lng单位ID As Long, _
    ByVal strNO As String, ByVal int记录状态 As Integer, _
    intErrInfor As Integer) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:加载数据接口
    '入参:
    '出参:intErrInfor-返回错误信息代码(1-已经删除,2-已经审核)
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-08-19 12:53:05
    '-----------------------------------------------------------------------------------------------------------
    mEditType = EditType: mstrNo = strNO: mint记录状态 = int记录状态: mlng单位ID = lng单位ID
    mblnEdit = mEditType = g新增 Or mEditType = g修改
    zlLoadData = initCard(intErrInfor)
End Function

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtInfo(Index), KeyAscii, m文本式
End Sub

Private Sub txtInfo_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
End Sub

Private Sub Init计算余额()
    '计算相关余额
    Dim dbl余额 As Double, i As Integer
    If mEditType <> g新增 And mEditType <> g修改 Then Exit Sub
    dbl余额 = 0
    With vsPayEdit
        For i = 1 To .Rows - 1
            dbl余额 = dbl余额 + Val(.TextMatrix(i, .ColIndex("付款金额")))
        Next
        If (mdbl本次应付 - mdbl本次预交) - dbl余额 <> 0 Then
            If .Row = .Rows - 1 And Val(.TextMatrix(.Row, .ColIndex("付款金额"))) = 0 Then
                .TextMatrix(.Row, .ColIndex("付款金额")) = Format((mdbl本次应付 - mdbl本次预交) - dbl余额, gVbFmtString.FM_金额)
            End If
        End If
    End With
    Call 付款合计
End Sub

Private Sub vsPayEdit_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '--------------------------------------------------------------------------------
    '设置相关的格式
    '刘兴宏:2007/09/17
    '--------------------------------------------------------------------------------
    With vsPayEdit
        Select Case Col
        Case .ColIndex("付款金额")
            .TextMatrix(Row, .Col) = Format(Val(.TextMatrix(Row, .Col)), gVbFmtString.FM_金额)
        Case .ColIndex("结算号码")
            If mEditType = g审核 Or mEditType = g修改 Then
                If Trim(.TextMatrix(Row, Col)) <> Trim(.Cell(flexcpData, Row, Col)) Then
                    .Cell(flexcpForeColor, Row, Col) = vbRed
                Else
                    .Cell(flexcpForeColor, Row, Col) = .ForeColor
                End If
            End If
        End Select
    End With
End Sub

Private Sub vsPayEdit_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsPayEdit, OldRow, NewRow, OldCol, NewCol)
End Sub

Private Sub vsPayEdit_AfterSort(ByVal Col As Long, Order As Integer)
    Call zl_VsGridAfterSort(vsPayEdit, Col, Order)
End Sub

Private Sub vsPayEdit_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsPayEdit
        Select Case Col
        Case .ColIndex("付款方式"), .ColIndex("付款金额")
            If mEditType <> g新增 And mEditType <> g修改 Then
                Cancel = True: Exit Sub
            End If
        Case .ColIndex("结算号码")
            If mEditType <> g新增 And mEditType <> g修改 And mEditType <> g审核 Then
                Cancel = True: Exit Sub
            End If
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vsPayEdit_ChangeEdit()
    mblnChange = True
    RaiseEvent zlChangeData(mblnChange)
End Sub

Private Sub vsPayEdit_EnterCell()
    If mEditType <> g新增 And mEditType <> g修改 Then Exit Sub
    
    With vsPayEdit
        .EditMaxLength = 0
        Select Case .Col
        Case .ColIndex("结算方式")
            '             .ColComboList(.Col) = "..."
        Case .ColIndex("付款金额")
            .EditMaxLength = 16
        Case .ColIndex("结算号码")
            .EditMaxLength = 10
        End Select
    End With
End Sub

Private Sub vsPayEdit_GotFocus()
    zl_VsGridGotFocus vsPayEdit
End Sub

Private Sub vsPayEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long
    
    With vsPayEdit
        If (.Col = .ColIndex("结算方式")) And KeyCode <> vbKeyReturn Then
           ' .ColComboList(.Col) = ""
        End If
        
        If KeyCode = vbKeyDelete And (mEditType = g新增 Or mEditType = g修改) Then
            If .Row = .Rows - 1 And .Row = 1 Then
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(.Row, lngCol) = ""
                    .Cell(flexcpData, .Row, lngCol) = ""
                Next
            Else
                .RemoveItem .Row
            End If
            Call Init计算余额
        End If
    End With
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vsPayEdit
        Select Case .Col
        Case .ColIndex("付款方式")
            If Trim(.TextMatrix(.Row, .Col)) = "" Then
               If zlControl.IsCtrlSetFocus(txtInfo(0)) Then
                  zlControl.IsCtrlSetFocus txtInfo(0)
               Else
                  zlCommFun.PressKey vbKeyTab
               End If
                Exit Sub
            End If
        End Select
        Call zlVsMoveGridCell(vsPayEdit, vsPayEdit.ColIndex("付款方式"), vsPayEdit.Cols - 1, mblnEdit)
        If mblnEdit Then
            Call Init计算余额
            '设置默认的结算方式
            Call Local结算方式
        End If
    End With
End Sub

Private Sub Local结算方式()
    '-----------------------------------------------------------------------------------------------------------
    '功能:结算方式定位
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-08-21 13:45:52
    '-----------------------------------------------------------------------------------------------------------
    If mrs结算方式 Is Nothing Then Exit Sub
    If mrs结算方式.State <> 1 Then Exit Sub
    With vsPayEdit
        If mrs结算方式.RecordCount <> 0 Then mrs结算方式.MoveFirst
        If Val(.TextMatrix(.Row, .ColIndex("付款金额"))) <> 0 _
            And Trim(.TextMatrix(.Row, .ColIndex("付款方式"))) = "" Then
            If mrs结算方式.EOF = False Then mrs结算方式.MoveFirst
            If .Row > 1 Then
                If Trim(.TextMatrix(.Row - 1, .ColIndex("付款方式"))) <> "" Then
                    mrs结算方式.Find "结算方式='" & Trim(.TextMatrix(.Row - 1, .ColIndex("付款方式"))) & "'"
                    If mrs结算方式.EOF = False Then mrs结算方式.MoveNext
                    If mrs结算方式.EOF = False Then
                        .TextMatrix(.Row, .ColIndex("付款方式")) = Nvl(mrs结算方式!结算方式)
                    End If
                End If
            ElseIf mrs结算方式.EOF = False Then
                .TextMatrix(.Row, .ColIndex("付款方式")) = Nvl(mrs结算方式!结算方式)
            End If
        End If
    End With
End Sub

Private Sub vsPayEdit_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim intCol As Integer
    Dim strKey As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsPayEdit
        Select Case Col
        Case .ColIndex("付款方式")
        Case .ColIndex("付款金额")
            .TextMatrix(Row, Col) = Format(Val(strKey), gVbFmtString.FM_金额)
        Case Else
        End Select
        Call zlVsMoveGridCell(vsPayEdit, .ColIndex("付款方式"), .Cols - 1, mblnEdit)
        Call Init计算余额
        '设置默认的结算方式
        Call Local结算方式
    End With
 End Sub
 
Private Sub vsPayEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub vsPayEdit_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Exit Sub
    End If
    With vsPayEdit
        Select Case Col
        Case .ColIndex("结算号码")
            Call VsFlxGridCheckKeyPress(vsPayEdit, Row, Col, KeyAscii, m文本式)
        Case .ColIndex("付款金额")
            '主要可能存在退款情况
            Call VsFlxGridCheckKeyPress(vsPayEdit, Row, Col, KeyAscii, m负金额式)
        Case Else
        End Select
    End With
End Sub

Private Sub vsPayEdit_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsPayEdit)
End Sub

Private Sub vsPayEdit_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String, intCol As Integer, dbl余额 As Double, i As Long
    
    If mEditType <> g新增 And mEditType <> g修改 Then
        If Col <> vsPayEdit.ColIndex("结算号码") And mEditType <> g审核 Then
            Cancel = True: Exit Sub
        End If
    End If
    With vsPayEdit
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        Select Case Col
        Case .ColIndex("结算号码")
            If strKey <> "" Then
                If LenB(StrConv(strKey, vbFromUnicode)) > 10 Then
                    ShowMsgbox "结算号码超长,最多能输入5个汉字或10个字符!"
                    zlControl.IsCtrlSetFocus vsPayEdit
                    Cancel = True
                    Exit Sub
                End If
                If InStr(1, strKey, "'") <> 0 Then
                    ShowMsgbox "结算号码不能输入单引号!"
                    zlControl.IsCtrlSetFocus vsPayEdit
                    Cancel = True
                    Exit Sub
                End If
            End If
        Case .ColIndex("付款金额")
            If strKey <> "" Then
                If Not IsNumeric(strKey) Then
                    ShowMsgbox "结算金额不是数据型,请重输!"
                    zlControl.IsCtrlSetFocus vsPayEdit
                    Cancel = True
                    Exit Sub
                End If
    '            If Val(strKey) < 0 Then
    '                ShowMsgbox "结算金额不能小于零,请重输!"
    '                zlCtlSetFocus vsPayEdit, True
    '                Cancel = True
    '                Exit Sub
    '            End If
                If Abs(Val(strKey)) > 10 ^ 12 - 1 Then
                    ShowMsgbox "结算金额只能在-" & 10 ^ 12 - 1 & "至" & 10 ^ 12 - 1 & "之间的数据,请重输!"
                    zlControl.IsCtrlSetFocus vsPayEdit
                    Cancel = True
                    Exit Sub
                End If
                
                dbl余额 = 0
                For i = 1 To .Rows - 1
                    If i <> .Row Then
                        dbl余额 = dbl余额 + Val(.TextMatrix(i, .ColIndex("付款金额")))
                    End If
                Next
    '            dbl余额 = (mdbl本次应付 - mdbl本次预交) - dbl余额
    '            dbl余额 = dbl余额 - Val(strKey)
    '            If dbl余额 < 0 Then
    '                ShowMsgbox "付款金额超出总额!"
    '                zlCtlSetFocus vsPayEdit, True
    '                Cancel = True
    '                Exit Sub
    '            End If
            End If
        End Select
    End With
End Sub

Private Sub 付款合计()
    Dim lngRow As Long, dblCount As Double
   '获取付款合计数
    With vsPayEdit
        For lngRow = 1 To .Rows - 1
            dblCount = dblCount + Val(.TextMatrix(lngRow, .ColIndex("付款金额")))
        Next
    End With
    '问题27930 by lesfeng 2010-03-23
    If mint标记 = 0 Then
        lblEdit(mlblIdx.idx_lbl付款合计).Caption = "结算合计:" & Format(dblCount, gVbFmtString.FM_金额) & "元"
    Else
        lblEdit(mlblIdx.idx_lbl付款合计).Caption = "标记结算合计:" & Format(dblCount, gVbFmtString.FM_金额) & "元"
    End If
End Sub

Private Sub vs冲预付_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vs冲预付, OldRow, NewRow, OldCol, NewCol)
End Sub

Private Sub vs冲预付_GotFocus()
    Call zl_VsGridGotFocus(vs冲预付)
End Sub

Private Sub vs冲预付_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vs冲预付
        Select Case .Col
        Case .ColIndex("付款方式")
        Case .Cols - 1
            If .Row = .Rows - 1 Then
                zlCommFun.PressKey vbKeyTab
                Exit Sub
            End If
        End Select
        End With
    Call zlVsMoveGridCell(vs冲预付, vs冲预付.ColIndex("付款方式"), vs冲预付.Cols - 1, False)
    
End Sub
'共公属性
Public Property Get zldbl本次应付() As Double
    zldbl本次应付 = mdbl本次应付
End Property

Public Property Let zldbl本次应付(ByVal vNewValue As Double)
    mdbl本次应付 = vNewValue
     
    Call InitPayData
End Property

Public Property Get zldbl本次预交() As Double
    zldbl本次预交 = mdbl本次预交
End Property

Public Property Let zldbl本次预交(ByVal vNewValue As Double)
    mdbl本次预交 = vNewValue
    Call InitPayData
End Property

Private Sub InitPayData()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化付款单数据
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-08-19 15:48:58
    '-----------------------------------------------------------------------------------------------------------
    lblEdit(mlblIdx.idx_lbl冲预交合计).Caption = "冲预付款合计：" & Format(mdbl本次预交, gVbFmtString.FM_金额) & "元"
    '问题27930 by lesfeng 2010-03-23
    If mint标记 = 0 Then
        lblEdit(mlblIdx.idx_lbl本次付款).Caption = "本次付款：" & Format(mdbl本次应付, gVbFmtString.FM_金额) & "元"
    Else
        lblEdit(mlblIdx.idx_lbl本次付款).Caption = "本次标记付款：" & Format(mdbl本次应付, gVbFmtString.FM_金额) & "元"
    End If
    With vsPayEdit
    
        If .Rows = 2 Then
            .Row = 1
'            If .TextMatrix(.Row, .ColIndex("付款方式")) = "" Then
                .TextMatrix(.Row, .ColIndex("付款金额")) = Format(mdbl本次应付 - mdbl本次预交, gVbFmtString.FM_金额)
'            End If
            '问题27930 by lesfeng 2010-03-23
            If mint标记 = 0 Then
                .TextMatrix(.Row, .ColIndex("标记付款")) = "正常"
            Else
                .TextMatrix(.Row, .ColIndex("标记付款")) = "标记"
            End If
        End If
        '设置默认的结算方式
        Call Local结算方式
    End With
End Sub

Public Function zlValidData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:验证数据的合法性
    '--入参数:
    '--出参数:
    '--返  回:验证合法,返回True,否则=false
    '-----------------------------------------------------------------------------------------------------------
    Dim strTemp As String
    Dim intIndex As Integer, lngRow As Long, dblCount As Double
    
    With vsPayEdit
        For lngRow = 1 To .Rows - 1
            If Trim(.TextMatrix(lngRow, .ColIndex("付款方式"))) <> "" Then
                strTemp = Trim(.TextMatrix(lngRow, .ColIndex("付款金额")))
                If strTemp = "" Then
                    ShowMsgbox "付款金额必需输入!"
                    zlControl.IsCtrlSetFocus vsPayEdit
                    Exit Function
                End If
                If Not IsNumeric(strTemp) Then
                    ShowMsgbox "付款金额不是数据型,请重输!"
                    zlControl.IsCtrlSetFocus vsPayEdit
                    Exit Function
                End If
                If Abs(Val(strTemp)) > 10 ^ 12 - 1 Then
                    ShowMsgbox "付款金额只能在-" & 10 ^ 12 - 1 & "至" & 10 ^ 12 - 1 & "之间的数据,请重输!"
                    zlControl.IsCtrlSetFocus vsPayEdit
                    Exit Function
                End If
                
                dblCount = dblCount + Val(strTemp)
                strTemp = Trim(.TextMatrix(lngRow, .ColIndex("结算号码")))
                If strTemp <> "" Then
                    If LenB(StrConv(strTemp, vbFromUnicode)) > 10 Then
                        ShowMsgbox "结算号码超长,最多能输入5个汉字或10个字符!"
                        zlControl.IsCtrlSetFocus vsPayEdit
                        Exit Function
                    End If
                    If InStr(1, strTemp, "'") <> 0 Then
                        ShowMsgbox "结算号码不能输入单引号!"
                        zlControl.IsCtrlSetFocus vsPayEdit
                        Exit Function
                    End If
                End If
            End If
        Next
    End With
    
    If Round(mdbl本次应付 - (dblCount + mdbl本次预交), g_小数位数.金额小数) <> 0 Then
        ShowMsgbox "付款金额不平,请检查付款金额与入库单" & vbCrLf & "发票金额和预付款之差是否相同!"
        zlControl.IsCtrlSetFocus vsPayEdit
        Exit Function
    End If
    If mdbl本次应付 = 0 Then
        ShowMsgbox "本次不存在任何应付记录,请检查!"
        zlControl.IsCtrlSetFocus vsPayEdit
        Exit Function
    End If
    
    If LenB(StrConv(txtInfo(0).Text, vbFromUnicode)) > 50 Then
        ShowMsgbox "付款说明的长度超长!(最多为50个字符或25个汉字)"
        zlControl.IsCtrlSetFocus txtInfo(0)
        Exit Function
    End If
    zlValidData = True
End Function

Public Function zlSaveCard(ByRef cllPro As Collection, ByRef lng付款序号 As Long, ByRef strNO As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:数据保存
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-08-19 15:06:38
    '-----------------------------------------------------------------------------------------------------------
    Dim int序号_IN As Integer
    Dim dbl金额_IN As Double
    Dim str结算方式_IN As String
    Dim str结算号码_IN As String
    
    Dim str填制人_IN As String
    Dim str填制日期_IN As String
    Dim lng付款序号_IN As Long
    Dim str摘要_IN As String
    Dim lngRow As Long
    Dim intCol As Integer
    
    zlSaveCard = False
    'txtNo = NextNo(31)
    
    str填制人_IN = UserInfo.姓名
    str填制日期_IN = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    
    str摘要_IN = txtInfo(0).Text
    
    
    On Error GoTo errHandle:
    strNO = txtNo.Caption
    If mEditType = g新增 Then
        lng付款序号_IN = zlDatabase.GetNextId("付款记录")
        lng付款序号 = lng付款序号_IN
        strNO = NextNo(31)
        txtNo.Tag = strNO
    Else
        lng付款序号_IN = mlng付款序号
        lng付款序号 = lng付款序号_IN
        gstrSQL = "zl_付款记录_DELETE('" & strNO & "')"
        AddArray cllPro, gstrSQL
    End If
     Dim blnData As Boolean
     blnData = False
    '循环保存每行数据
    With vsPayEdit
        For lngRow = 1 To .Rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("付款金额"))) <> 0 _
                And Trim(.TextMatrix(lngRow, .ColIndex("付款方式"))) <> "" Then
                blnData = True
                dbl金额_IN = .TextMatrix(lngRow, .ColIndex("付款金额"))
                '问题27930 by lesfeng 2010-03-23
                If mint标记 = 0 Then
                    str结算方式_IN = .TextMatrix(lngRow, .ColIndex("付款方式"))
                Else
                    str结算方式_IN = ""
                End If
                str结算号码_IN = .TextMatrix(lngRow, .ColIndex("结算号码"))
                            
                'Zl_付款管理_Insert
                gstrSQL = " zl_付款管理_INSERT("
                '  No_In       IN 付款记录.NO%TYPE,
                gstrSQL = gstrSQL & "'" & strNO & "',"
                '  序号_In     IN 付款记录.序号%TYPE,
                gstrSQL = gstrSQL & "" & lngRow & ","
                '  预付款_In   IN 付款记录.预付款%TYPE := 0,
                gstrSQL = gstrSQL & "" & 0 & ","
                '  单位id_In   IN 付款记录.单位id%TYPE,
                gstrSQL = gstrSQL & "" & mlng单位ID & ","
                '  金额_In     IN 付款记录.金额%TYPE,
                gstrSQL = gstrSQL & "" & dbl金额_IN & ","
                '  结算方式_In IN 付款记录.结算方式%TYPE,
                gstrSQL = gstrSQL & "'" & str结算方式_IN & "',"
                '  结算号码_In IN 付款记录.结算号码%TYPE := NULL,
                gstrSQL = gstrSQL & "'" & str结算号码_IN & "',"
                '  填制人_In   IN 付款记录.填制人%TYPE,
                gstrSQL = gstrSQL & "'" & str填制人_IN & "',"
                '  填制日期_In IN 付款记录.填制日期%TYPE,
                gstrSQL = gstrSQL & "to_date('" & str填制日期_IN & "','yyyy-mm-dd HH24:MI:SS'),"
                '  付款序号_In IN 付款记录.付款序号%TYPE := NULL,
                gstrSQL = gstrSQL & "" & lng付款序号_IN & ","
                '  摘要_In     IN 付款记录.摘要%TYPE := NULL
                gstrSQL = gstrSQL & "'" & str摘要_IN & "',"
                '问题27930 by lesfeng 2010-03-23
                '  拒付标志_In IN 付款记录.拒付标志%TYPE := 0
                gstrSQL = gstrSQL & "" & mint标记 & ")"
                AddArray cllPro, gstrSQL
            End If
        Next
    End With
    
    If blnData = False Then
        'Zl_付款管理_Insert
        gstrSQL = " zl_付款管理_INSERT("
        '  No_In       IN 付款记录.NO%TYPE,
        gstrSQL = gstrSQL & "'" & strNO & "',"
        '  序号_In     IN 付款记录.序号%TYPE,
        gstrSQL = gstrSQL & "" & lngRow & ","
        '  预付款_In   IN 付款记录.预付款%TYPE := 0,
        gstrSQL = gstrSQL & "" & 0 & ","
        '  单位id_In   IN 付款记录.单位id%TYPE,
        gstrSQL = gstrSQL & "" & mlng单位ID & ","
        '  金额_In     IN 付款记录.金额%TYPE,
        gstrSQL = gstrSQL & "" & dbl金额_IN & ","
        '  结算方式_In IN 付款记录.结算方式%TYPE,
        gstrSQL = gstrSQL & "'" & "" & "',"
        '  结算号码_In IN 付款记录.结算号码%TYPE := NULL,
        gstrSQL = gstrSQL & "'" & "" & "',"
        '  填制人_In   IN 付款记录.填制人%TYPE,
        gstrSQL = gstrSQL & "'" & str填制人_IN & "',"
        '  填制日期_In IN 付款记录.填制日期%TYPE,
        gstrSQL = gstrSQL & "to_date('" & str填制日期_IN & "','yyyy-mm-dd HH24:MI:SS'),"
        '  付款序号_In IN 付款记录.付款序号%TYPE := NULL,
        gstrSQL = gstrSQL & "" & lng付款序号_IN & ","
        '  摘要_In     IN 付款记录.摘要%TYPE := NULL
        gstrSQL = gstrSQL & "'" & str摘要_IN & "',"
        '问题27930 by lesfeng 2010-03-23
        '  拒付标志_In IN 付款记录.拒付标志%TYPE := 0
        gstrSQL = gstrSQL & "" & mint标记 & ")"
        AddArray cllPro, gstrSQL
    End If
    zlSaveCard = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Exit Function
End Function

Public Function ClearData()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化网格控件的默认属性
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-08-19 15:57:36
    '-----------------------------------------------------------------------------------------------------------
    txtInfo(0).Text = ""
    vsPayEdit.Clear 1
    vsPayEdit.Rows = 2
    vs冲预付.Clear 1
    vs冲预付.Rows = 2
    mlng单位ID = 0
    Call zlLoadPrivder(0)
    mblnChange = False
End Function

Private Sub vs冲预付_LostFocus()
    Call zl_VsGridLOSTFOCUS(vs冲预付)
End Sub

Public Function zlCheck(ByRef cllPro As Collection) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:审核数据
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-08-19 15:06:38
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, str结算号码_IN As String
    '循环保存每行数据
    Err = 0: On Error GoTo ErrHand:
    With vsPayEdit
        For lngRow = 1 To .Rows - 1
            If Val(.Cell(flexcpData, lngRow, .ColIndex("付款方式"))) <> 0 Then
                str结算号码_IN = .TextMatrix(lngRow, .ColIndex("结算号码"))
                If Trim(.Cell(flexcpData, lngRow, .ColIndex("结算号码"))) <> str结算号码_IN Then
                     If str结算号码_IN <> "" Then
                        If LenB(StrConv(str结算号码_IN, vbFromUnicode)) > 10 Then
                            ShowMsgbox "结算号码超长,最多能输入5个汉字或10个字符!"
                            .Col = .ColIndex("结算号码"): .Row = lngRow: .TopRow = lngRow
                            Exit Function
                        End If
                        If InStr(1, str结算号码_IN, "'") <> 0 Then
                            ShowMsgbox "结算号码不能输入单引号!"
                            .Col = .ColIndex("结算号码"): .Row = lngRow: .TopRow = lngRow
                            Exit Function
                        End If
                    End If
                   
                    ' Zl_付款记录_结算号update
                    gstrSQL = " Zl_付款记录_结算号update("
                    '  Id_In       付款记录.ID%Type,
                    gstrSQL = gstrSQL & "" & Val(.Cell(flexcpData, lngRow, .ColIndex("付款方式"))) & ","
                    '  结算号码_In In 付款记录.结算号码%Type
                    gstrSQL = gstrSQL & "'" & str结算号码_IN & "')"
                    AddArray cllPro, gstrSQL
                End If
            End If
        Next
    End With
    zlCheck = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Exit Function
End Function
