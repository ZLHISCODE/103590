VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCompoundPack 
   Caption         =   "输液批量打包"
   ClientHeight    =   10230
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15210
   Icon            =   "frmCompoundPack.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   15210
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picExecuted 
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   4560
      ScaleHeight     =   5415
      ScaleWidth      =   13095
      TabIndex        =   0
      Top             =   4440
      Width           =   13095
      Begin VSFlex8Ctl.VSFlexGrid vsgExecUnpack 
         Height          =   4860
         Left            =   0
         TabIndex        =   1
         Top             =   480
         Width           =   8505
         _cx             =   15002
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
         BackColorSel    =   16771802
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCompoundPack.frx":6852
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
      Begin MSComCtl2.DTPicker dpkExecuted 
         Height          =   300
         Index           =   1
         Left            =   3120
         TabIndex        =   2
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   126222339
         CurrentDate     =   40945.9999884259
      End
      Begin MSComCtl2.DTPicker dpkExecuted 
         Height          =   300
         Index           =   0
         Left            =   960
         TabIndex        =   3
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   126222339
         CurrentDate     =   40945
      End
      Begin VB.Label lblInfo 
         Caption         =   "~"
         Height          =   135
         Index           =   10
         Left            =   2880
         TabIndex        =   5
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblInfo 
         Caption         =   "执行时间"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   4
         Top             =   165
         Width           =   855
      End
   End
   Begin VB.PictureBox picWaitExecute 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9135
      Left            =   360
      ScaleHeight     =   9135
      ScaleWidth      =   14655
      TabIndex        =   6
      Top             =   600
      Width           =   14655
      Begin VB.PictureBox picWaitExecAdvice 
         BorderStyle     =   0  'None
         Height          =   8895
         Left            =   3000
         ScaleHeight     =   8895
         ScaleWidth      =   11535
         TabIndex        =   7
         Top             =   0
         Width           =   11535
         Begin VSFlex8Ctl.VSFlexGrid vsgWaitUnpack 
            Height          =   4860
            Left            =   0
            TabIndex        =   8
            Top             =   720
            Width           =   8505
            _cx             =   15002
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
            BackColorSel    =   16771802
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   2
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   250
            RowHeightMax    =   2000
            ColWidthMin     =   0
            ColWidthMax     =   5000
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmCompoundPack.frx":68ED
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
            AutoSizeMode    =   1
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
            OwnerDraw       =   1
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
      End
      Begin VB.Frame fraPatiInfo 
         Height          =   9015
         Left            =   0
         TabIndex        =   9
         Top             =   -80
         Width           =   2895
         Begin XtremeReportControl.ReportControl rptPati 
            Height          =   6495
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   2655
            _Version        =   589884
            _ExtentX        =   4683
            _ExtentY        =   11456
            _StockProps     =   0
            BorderStyle     =   2
            AutoColumnSizing=   0   'False
         End
         Begin VB.PictureBox picFitter 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1455
            Left            =   120
            ScaleHeight     =   1455
            ScaleWidth      =   2655
            TabIndex        =   15
            Top             =   7440
            Width           =   2655
            Begin VB.CheckBox chk期效 
               Caption         =   "临嘱"
               Height          =   180
               Index           =   1
               Left            =   1800
               TabIndex        =   17
               Top             =   1117
               Value           =   1  'Checked
               Width           =   735
            End
            Begin VB.CheckBox chk期效 
               Caption         =   "长嘱"
               Height          =   180
               Index           =   0
               Left            =   960
               TabIndex        =   16
               Top             =   1117
               Value           =   1  'Checked
               Width           =   735
            End
            Begin MSComCtl2.DTPicker dpkReqTime 
               Height          =   300
               Index           =   0
               Left            =   720
               TabIndex        =   18
               Top             =   330
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd HH:mm"
               Format          =   126222339
               CurrentDate     =   40945
            End
            Begin MSComCtl2.DTPicker dpkReqTime 
               Height          =   300
               Index           =   1
               Left            =   720
               TabIndex        =   19
               Top             =   720
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd HH:mm"
               Format          =   126222339
               CurrentDate     =   40945.9999884259
            End
            Begin VB.Label lblInfo 
               Caption         =   "期效"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   23
               Top             =   1117
               Width           =   495
            End
            Begin VB.Label lblInfo 
               Caption         =   "到"
               Height          =   255
               Index           =   3
               Left            =   360
               TabIndex        =   22
               Top             =   750
               Width           =   255
            End
            Begin VB.Label lblInfo 
               Caption         =   "从"
               Height          =   255
               Index           =   2
               Left            =   360
               TabIndex        =   21
               Top             =   360
               Width           =   255
            End
            Begin VB.Label lblInfo 
               Caption         =   "执行时间"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   20
               Top             =   50
               Width           =   975
            End
         End
         Begin VB.Frame fraBaby 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   240
            TabIndex        =   11
            Top             =   7200
            Visible         =   0   'False
            Width           =   2600
            Begin VB.OptionButton optBaby 
               Caption         =   "病人"
               Height          =   180
               Index           =   1
               Left            =   1080
               TabIndex        =   14
               Top             =   0
               Width           =   660
            End
            Begin VB.OptionButton optBaby 
               Caption         =   "所有医嘱"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   13
               Top             =   0
               Value           =   -1  'True
               Width           =   1020
            End
            Begin VB.OptionButton optBaby 
               Caption         =   "婴儿"
               Height          =   180
               Index           =   2
               Left            =   1815
               TabIndex        =   12
               Top             =   0
               Width           =   660
            End
         End
         Begin VB.Label lblInfo 
            Caption         =   "当前病区："
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   24
            Top             =   240
            Width           =   2500
         End
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   840
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompoundPack.frx":6988
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompoundPack.frx":6F22
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompoundPack.frx":74BC
            Key             =   "签名"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompoundPack.frx":780E
            Key             =   "Woman"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompoundPack.frx":E070
            Key             =   "Man"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompoundPack.frx":148D2
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompoundPack.frx":14E6C
            Key             =   "AllCheck"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   9975
      Left            =   120
      TabIndex        =   25
      Top             =   0
      Width           =   15015
      _Version        =   589884
      _ExtentX        =   26485
      _ExtentY        =   17595
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   26
      Top             =   9870
      Width           =   15210
      _ExtentX        =   26829
      _ExtentY        =   635
      SimpleText      =   $"frmCompoundPack.frx":15406
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCompoundPack.frx":1544D
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21749
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCompoundPack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum PatiCol
    COL_病人ID = 0
    COL_主页ID = 1
    COL_选择 = 2
    COL_床号 = 3
    COL_姓名 = 4
    COL_性别 = 5
    COL_住院号 = 6
End Enum

Private Enum AdviceCol
    col选择 = 0
    COL床位 = 1
    col姓名 = 2
    col性别 = 3
    col期效 = 4
    col医嘱内容 = 5
    col总量 = 6
    col单量 = 7
    COL给药途径 = 8
    
    col执行时间 = 9
    col配药批次 = 10
    col配药工作时间 = 11
    col瓶签号 = 12
    col状态 = 13
    Col发送时间 = 14
    col销帐申请人 = 15
    
    col医嘱ID = 16
    col相关ID = 17
    col诊疗类别 = 18
    Col病人ID = 19
    COL主页ID = 20
    COL频率 = 21
    col发送号 = 22
    col在院 = 23
    col配药ID = 24
End Enum

Private mlng病区ID As Long
Private mlng病人ID As Long
Private mint医嘱处理范围 As Integer    '医嘱处理范围   0-所有医嘱,1-病人医嘱,2-婴儿医嘱
Private mlng医护科室ID As Long
Private mlng婴儿科室ID As Long
Private mlng婴儿病区ID As Long
Private mrsDefine As New Recordset
Private mbln摆药后不能改状态 As Boolean

Public Sub ShowMe(ByVal intType As Integer, ByRef frmParent As Object, ByVal lng病区ID As Long, ByVal lng病人ID As Long, Optional ByVal lng医护科室ID As Long, _
    Optional ByVal lng婴儿科室ID As Long, Optional ByVal lng婴儿病区ID As Long)
    mlng病区ID = lng病区ID
    mlng病人ID = lng病人ID
    mlng医护科室ID = lng医护科室ID
    mlng婴儿科室ID = lng婴儿科室ID
    mlng婴儿病区ID = lng婴儿病区ID
    Me.lblInfo(0).Caption = "当前病区：" & Sys.RowValue("部门表", IIF(mlng婴儿病区ID <> 0 And mlng婴儿病区ID = mlng医护科室ID, lng婴儿病区ID, lng病区ID), "名称")
    
    Me.Show intType, frmParent
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    tbcSub.SetFocus
    Select Case Control.ID
    
    Case conMenu_View_Refresh
        If tbcSub.Selected.Tag = "待打包医嘱" Then
            Call LoadAdvice
        Else
            Call LoadAdvice(True)
        End If
    Case conMenu_Edit_Save
        Call FuncUnpack
    Case conMenu_Manage_Undone
        Call FuncCancleUnpack
    Case conMenu_File_Exit
        Unload Me
    
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Dim lngLW As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
        
    'TabControl
    tbcSub.Left = lngLeft
    tbcSub.Top = lngTop
    tbcSub.Width = Me.Width
    tbcSub.Height = Me.Height - stbThis.Height - 560 - lngTop
    
    
       
    Me.Refresh
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    
    Case conMenu_Manage_Undone
        If tbcSub.Selected.Tag = "待打包医嘱" Then
            Control.Visible = False
        Else
            Control.Visible = True
        End If
    Case conMenu_Edit_Save
        If tbcSub.Selected.Tag = "待打包医嘱" Then
            Control.Visible = True
        Else
            Control.Visible = False
        End If
    End Select
End Sub

Private Sub chk期效_Click(Index As Integer)
    If chk期效(0).value = 0 And chk期效(1).value = 0 Then
        chk期效(Index).value = 1
    End If
End Sub

Private Sub FuncCancleUnpack()
'功能：取消执行
    Dim arrSQL() As Variant
    Dim i As Long, j As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim blnTrans As Boolean
    Dim strCurDate As String
    Dim strIDs As String, rsTmp As Recordset, strSQL As String
    
    strCurDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    
    arrSQL = Array()
    With vsgExecUnpack
        For i = 1 To .Rows - 1
            If .TextMatrix(i, col医嘱ID) <> "" And .RowData(i) = "Begin" And .Cell(flexcpData, i, col选择) = "1" Then
                
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_输液配药记录_Update(" & .TextMatrix(i, col配药ID) & ",Null," & _
                    ZVal(.Cell(flexcpData, i, col配药批次)) & ",'" & UserInfo.姓名 & "'," & strCurDate & ")"
                    strIDs = strIDs & "," & .TextMatrix(i, col配药ID)
            End If
        Next
    End With
    
    Screen.MousePointer = 11
    On Error GoTo errH
    If strIDs <> "" Then
        strSQL = "select ID from 输液配药记录 where 是否锁定=1 And ID in(Select Column_Value From Table(f_Num2list([1])))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strIDs, 2))
        If rsTmp.RecordCount > 0 Then
            MsgBox "当前调整的配药记录已经被输液配药中心锁定，暂时不允许进行取消打包,已将这些记录取消勾选。", vbInformation, "输液配液记录"
            Screen.MousePointer = 0
            For i = 1 To vsgExecUnpack.Rows - 1
                rsTmp.Filter = "ID=" & Val(vsgExecUnpack.TextMatrix(i, col配药ID))
                If rsTmp.RecordCount > 0 Then
                    vsgExecUnpack.Row = i
                    Call ExecCheck(vsgExecUnpack)
                End If
            Next
            Exit Sub
        End If
    End If
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    On Error GoTo 0
    Screen.MousePointer = 0
    If vsgExecUnpack.TextMatrix(1, col医嘱ID) = "" Then
        stbThis.Panels(2).Text = "没有可取消打包的医嘱。"
    Else
        If UBound(arrSQL) = -1 Then
            MsgBox "请勾选您需要取消打包的医嘱。", vbInformation, Me.Caption
            Exit Sub
        End If
        stbThis.Panels(2).Text = "取消成功，本次共取消打包了 " & UBound(arrSQL) + 1 & " 条医嘱。"
    End If
    Call LoadAdvice(True)
    Exit Sub
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncUnpack()
'功能：打包
    Dim arrSQL() As Variant
    Dim i As Long
    Dim blnTrans As Boolean
    Dim strCurDate As String
    Dim strIDs As String, rsTmp As Recordset, strSQL As String
    
    strCurDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    arrSQL = Array()
    With vsgWaitUnpack
        For i = 1 To .Rows - 1
            If .TextMatrix(i, col医嘱ID) <> "" And .RowData(i) = "Begin" And .Cell(flexcpData, i, col选择) = "1" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_输液配药记录_Update(" & .TextMatrix(i, col配药ID) & "," & 1 & "," & _
                    ZVal(.Cell(flexcpData, i, col配药批次)) & ",'" & UserInfo.姓名 & "'," & strCurDate & ")"
                strIDs = strIDs & "," & .TextMatrix(i, col配药ID)
            End If
        Next
    End With
    
    Screen.MousePointer = 11
    On Error GoTo errH
    If strIDs <> "" Then
        strSQL = "select ID from 输液配药记录 where 是否锁定=1 And ID in(Select Column_Value From Table(f_Num2list([1])))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strIDs, 2))
        If rsTmp.RecordCount > 0 Then
            MsgBox "当前调整的配药记录已经被输液配药中心锁定，暂时不允许进行打包,已将这些记录取消勾选。", vbInformation, "输液配液记录"
            For i = 1 To vsgWaitUnpack.Rows - 1
                rsTmp.Filter = "ID=" & Val(vsgWaitUnpack.TextMatrix(i, col配药ID))
                If rsTmp.RecordCount > 0 Then
                    vsgWaitUnpack.Row = i
                    Call ExecCheck(vsgWaitUnpack)
                End If
            Next
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    On Error GoTo 0
    Screen.MousePointer = 0
    If vsgWaitUnpack.TextMatrix(1, col医嘱ID) = "" Then
        stbThis.Panels(2).Text = "没有可打包的药品。"
    Else
        stbThis.Panels(2).Text = "保存成功，本次共打包了 " & UBound(arrSQL) + 1 & " 条药品医嘱。"
    End If
    Call LoadAdvice
    Exit Sub
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        If tbcSub.Selected.Tag = "待打包医嘱" Then
            LoadAdvice
        Else
            LoadAdvice True
        End If
    ElseIf KeyCode = vbKey1 And Shift = 4 Then
        tbcSub.Item(0).Selected = True
    ElseIf KeyCode = vbKey2 And Shift = 4 Then
        tbcSub.Item(1).Selected = True
    End If
End Sub

Private Sub Form_Load()
    Dim strHead As String
    Dim strTbc As String
    
    mbln摆药后不能改状态 = Val(zlDatabase.GetPara("输液单摆药后临床不允许改变打包状态", glngSys, 1345, 0)) = 1
    
    'tabControl
    '-----------------------------------------------------
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '绑定子窗体时会Form_Load，且自动选中第一个加入的卡片
        '如果设置当前卡片隐藏,则不会自动切换选择,但显示内容未变
        '任意指定索引号无效，最终变为0-N，只是可能改变加入顺序。
        strTbc = "打包"
        .InsertItem(0, "待" & strTbc & "医嘱(&1)", picWaitExecute.hwnd, 0).Tag = "待打包医嘱"
        .InsertItem(1, "已" & strTbc & "医嘱(&2)", picExecuted.hwnd, 0).Tag = "已打包医嘱"
        
        .Item(0).Selected = True
    End With
    'commandbar
    '-----------------------------------------------------
    Call InitCommandBar
    'ReportControl
    '-----------------------------------------------------
    Call InitReportColumn
    'VSFlexGrid
    '-----------------------------------------------------
    strHead = ",400,1;床位,750,1;姓名,850,1;性别,450,1;期效,450,1;医嘱内容,2500,1;总量;单量,700,1;给药途径;执行时间,1550,1;配药批次,1470,1;配药工作时间,1550,1;瓶签号,1980,1;状态,900,1;发送时间,1550,1;销帐申请人;医嘱ID;相关ID;诊疗类别;病人ID;主页ID;频率;发送号;在院;配药ID"

    Call InitTable(vsgWaitUnpack, strHead)
    
    Call InitTable(vsgExecUnpack, strHead)

    Set mrsDefine = InitAdviceDefine
    Call InitPageData
    Call LoadPatiInfo
    
    Call RestoreWinState(Me, App.ProductName)
    
    If DeptIsWoman(0, Get科室IDs(IIF(mlng婴儿病区ID <> 0 And mlng婴儿病区ID = mlng医护科室ID, mlng婴儿病区ID, mlng病区ID))) Then
        fraBaby.Visible = True
        '医嘱处理范围
        mint医嘱处理范围 = Val(zlDatabase.GetPara("医嘱处理范围", glngSys, p住院医嘱发送, "0"))
        optBaby(mint医嘱处理范围).value = True
    End If
End Sub

Private Sub InitCommandBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objCbo As CommandBarComboBox
    
    '工具栏----------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    '生成工具栏
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, " 查询(&Q)"): objControl.BeginGroup = True
        objControl.ToolTipText = "读取待打包/已打包的数据"

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, " 打包(&S)")
        objControl.BeginGroup = True
        objControl.ToolTipText = "对已经勾选的医嘱进行打包。"
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Undone, "取消打包(&C)")
        objControl.ToolTipText = "对已经勾选的医嘱进行取消打包的操作。"
        objControl.IconId = 3651
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, " 退出(&E)"): objControl.BeginGroup = True
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh
        .Add FALT, vbKeyX, conMenu_File_Exit
        .Add 0, vbKeyEscape, conMenu_File_Exit
    End With

End Sub

Private Sub LoadAdvice(Optional ByVal blnIsUnpack As Boolean)
'功能：加载医嘱
'参数：blnIsUnpack=true加载已打包医嘱,false为加载待打包医嘱
    Dim rsTmp As Recordset
    Dim strSQL As String
    Dim i As Long, j As Long
    Dim lngID As Long       '用于定位
    Dim strFormat As String
    Dim strTmp As String
    Dim strFitter As String
    Dim strPatis As Variant
    Dim blnDo As Boolean, blnSetup As Boolean
    Dim lngCount As Long   '需要执行的医嘱数(一并给药算1条)
    Dim rsState As Recordset  '输液配药状态
    Dim strIDs As String
    Dim arrIDs() As Variant '定义一个字符串数组

    
    strSQL = "Select " & _
            " a.Id, b.发送号,b.id as 配药ID,a.相关id, a.诊疗类别,a.开始执行时间, B.姓名, p.当前床号 As 床号, B.性别, Decode(Nvl(a.医嘱期效, 0), 0, '长嘱', '临嘱') As 期效,a.医嘱状态,p.在院," & vbNewLine & _
            "       Decode(a.单次用量, Null, Null, decode(sign(1-A.单次用量),1,'0'||A.单次用量,A.单次用量) || c.计算单位) As 单量,  Decode(a.相关id,Null,a.医嘱内容 || ' ' || a.执行频次  ,a.医嘱内容) as 医嘱内容, to_char(b.执行时间,'YYYY-MM-DD HH24:MI') as 执行时间,p.当前病区ID" & _
            ", a.执行频次 As 频率, a.病人id, a.主页id, a.诊疗项目id,c.操作类型,c.执行分类,Decode(a.总给予量, Null, Null," & _
            "  Round(a.总给予量 / Decode(a.病人来源, 2, d.住院包装, d.门诊包装), 5) || Decode(a.病人来源, 2, d.住院单位, d.门诊单位)) As 总量,to_char(E.发送时间,'YYYY-MM-DD HH24:MI') as 发送时间," & _
            "b.配药批次,g.配药时间 as 配药工作时间,b.瓶签号,Decode(b.操作状态,1, '待摆药',2, '待配药', 3,'待配药', 4,'已配药', '已发送') As 状态,'' AS 销帐申请人" & vbNewLine & _
            " From 输液配药记录 B,病人医嘱发送 E, 病人医嘱记录 A,病案主页 F, 病人信息 P, 诊疗项目目录 C, 药品规格 D,配药工作批次 G" & vbNewLine & _
            " Where F.病人ID=P.病人ID And F.主页ID = P.主页ID And b.医嘱ID=e.医嘱ID And b.发送号=e.发送号 And g.批次(+)=b.配药批次 and g.配置中心id(+)=b.部门id And (a.Id = b.医嘱id Or a.相关id = b.医嘱id) And p.病人id = a.病人id And a.诊疗项目id = c.Id And a.收费细目id = d.药品id(+) And a.诊疗类别 Not In('C','7') And Not (a.诊疗类别='E' And c.操作类型='3') " & _
            Decode(mint医嘱处理范围, 1, " And nvl(a.婴儿,0) = 0 ", 2, " And nvl(a.婴儿,0) <> 0 ", "") & _
            " And (F.婴儿科室ID is null or F.婴儿科室ID is not null and (F.婴儿病区ID=[5] or F.婴儿科室ID=[5]) and NVL(A.婴儿,0)<>0 or F.婴儿科室ID is not null and (F.婴儿病区ID<>[5] and f.婴儿科室ID<>[5]) and NVL(A.婴儿,0)=0) "
            
    '病人
    If Not blnIsUnpack Then
        For i = 0 To rptPati.Rows.Count - 1
            If Not rptPati.Rows(i).GroupRow Then
                If rptPati.Rows(i).Record.Tag = "1" Then
                    strSQL = strSQL & IIF(strPatis = "", " And(", " Or") & " a.病人ID =" & rptPati.Rows(i).Record(COL_病人ID).value
                    strPatis = strPatis & "," & rptPati.Rows(i).Record(COL_病人ID).value
                End If
            End If
        Next
        strPatis = Mid(strPatis, 2)
        If strPatis = "" Then
            MsgBox "请选择需要查询的病人。", vbInformation, Me.Caption
            Exit Sub
        End If
        strSQL = strSQL & " )"
        strSQL = strSQL & " And nvl(b.是否打包,0)<>1  And b.执行时间 between [1] and [2] "
        strSQL = strSQL & " And (" & IIF(chk期效(0).value, "Nvl(a.医嘱期效, 0)=0" & IIF(chk期效(1).value, " Or Nvl(a.医嘱期效, 0)=1", ""), IIF(chk期效(1).value, "Nvl(a.医嘱期效, 0)=1", "")) & ")"
        '已经销帐和已经配药的不允许打包，参数 摆药后不能打包
        If mbln摆药后不能改状态 Then
            strSQL = strSQL & " And b.操作状态=1 "
        Else
            strSQL = strSQL & " And b.操作状态 in(1,2,3) "
        End If
        vsgWaitUnpack.Cell(flexcpPicture, 0, col选择) = img16.ListImages("AllCheck").Picture
        vsgWaitUnpack.Cell(flexcpPictureAlignment, 0, col选择) = flexPicAlignCenterCenter
        vsgWaitUnpack.ColData(col选择) = "Check"
        strSQL = strSQL & " Order By p.当前床号,B.姓名, b.执行时间,Nvl(a.相关id, a.Id),a.id,a.序号"
    Else
        '已经打包的
        strSQL = strSQL & " and f.当前病区id+0=[5] And nvl(b.是否打包,0)=1 And b.执行时间 between [3] and [4]  "
        vsgExecUnpack.Cell(flexcpPicture, 0, col选择) = img16.ListImages("UnCheck").Picture
        vsgExecUnpack.Cell(flexcpPictureAlignment, 0, col选择) = flexPicAlignCenterCenter
        vsgExecUnpack.ColData(col选择) = ""
        strSQL = strSQL & " Order By B.打包时间,p.当前床号,B.姓名, b.执行时间,Nvl(a.相关id, a.Id),a.id,a.序号"
    End If
    
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(dpkReqTime(0).value), CDate(dpkReqTime(1).value), CDate(dpkExecuted(0).value), CDate(dpkExecuted(1).value), mlng医护科室ID)
    
    i = 0
    strSQL = ""
    ReDim Preserve arrIDs(i)
    Do While Not rsTmp.EOF
        If Len(arrIDs(i) & "," & rsTmp!ID) >= 4000 Then
            i = i + 1
            ReDim Preserve arrIDs(i)
        End If
        arrIDs(i) = arrIDs(i) & "," & rsTmp!ID
        rsTmp.MoveNext
    Loop

    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    For j = 0 To i
        arrIDs(j) = Mid(arrIDs(j), 2)
        If arrIDs(j) <> "" Then
            strSQL = strSQL & "Select 配药ID,操作类型,操作人员,操作时间 from 输液配药状态 Where 配药ID in(select Column_Value From Table(Cast(f_num2list([" & j + 1 & "]) As ZLTOOLS.t_numlist)))"
        End If
        If j < i Then
            strSQL = strSQL & " union all "
        End If
    Next
    If strSQL <> "" Then
        Set rsState = zlDatabase.OpenSQLRecordByArray(strSQL, Me.Caption, arrIDs)
    End If

    With IIF(blnIsUnpack, vsgExecUnpack, vsgWaitUnpack)
        .Redraw = flexRDNone
        .Rows = 1
        If rsTmp.RecordCount > 0 Then
            i = 1
            Do While Not rsTmp.EOF
                .AddItem ""
                If .ColData(col选择) = "Check" Then
                    .Cell(flexcpPicture, i, col选择) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, i, col选择) = 1
                    .Cell(flexcpPictureAlignment, i, col选择) = flexPicAlignCenterCenter
                End If
                .TextMatrix(i, col姓名) = rsTmp!姓名 & ""
                .TextMatrix(i, col期效) = rsTmp!期效 & ""
                .TextMatrix(i, col单量) = rsTmp!单量 & ""
                .TextMatrix(i, col医嘱ID) = rsTmp!ID & ""
                .TextMatrix(i, col相关ID) = rsTmp!相关ID & ""
                .TextMatrix(i, col性别) = rsTmp!性别 & ""
                .TextMatrix(i, COL床位) = rsTmp!床号 & ""
                .TextMatrix(i, Col病人ID) = rsTmp!病人ID & ""
                .TextMatrix(i, COL主页ID) = rsTmp!主页ID & ""
                .TextMatrix(i, col诊疗类别) = rsTmp!诊疗类别 & ""
                .TextMatrix(i, col总量) = rsTmp!总量 & ""
                .TextMatrix(i, col发送号) = rsTmp!发送号 & ""
                .TextMatrix(i, COL频率) = rsTmp!频率 & ""
                .TextMatrix(i, Col发送时间) = rsTmp!发送时间 & ""
                .TextMatrix(i, col执行时间) = rsTmp!执行时间 & ""
                .TextMatrix(i, col配药批次) = "第" & rsTmp!配药批次 & "批"
                .Cell(flexcpData, i, col配药批次) = Val(rsTmp!配药批次 & "")
                .TextMatrix(i, col配药工作时间) = rsTmp!配药工作时间 & ""
                .TextMatrix(i, col状态) = rsTmp!状态 & ""
                .TextMatrix(i, col瓶签号) = rsTmp!瓶签号 & ""
                rsState.Filter = "配药ID=" & rsTmp!ID & " And 操作类型=9"
                If rsState.RecordCount > 0 Then
                    rsState.MoveFirst
                    .TextMatrix(i, col销帐申请人) = rsState!操作人员 & ""
                End If

                .TextMatrix(i, col在院) = rsTmp!在院 & ""
                .TextMatrix(i, col配药ID) = rsTmp!配药ID & ""

                .RowData(i) = IIF(.TextMatrix(i, col相关ID) = "", "Begin", "")
                '显示简洁模式下的医嘱内容
                strFormat = rsTmp!医嘱内容
                If .TextMatrix(i, COL频率) <> "一次性" Then
                    blnDo = True
                    If mrsDefine.RecordCount > 0 Then blnDo = InStr(mrsDefine!医嘱内容, "[总量]") = 0
                    If blnDo Then
                        strTmp = .TextMatrix(i, col总量)
                        If strTmp <> "" Then strFormat = strFormat & ",共" & strTmp
                    End If
                End If
                .TextMatrix(i, col医嘱内容) = strFormat
                '可编辑列颜色
                .Cell(flexcpBackColor, i, col选择, i, col选择) = COLEditBackColor
                
                '需要执行的医嘱组数
                If rsTmp!相关ID & "" = "" Then lngCount = lngCount + 1
                
                rsTmp.MoveNext
                i = i + 1
            Loop
        Else
            .AddItem ""
        End If
        If blnIsUnpack Then
            stbThis.Panels(2).Text = "共有 " & lngCount & " 条医嘱已经打包。"
        Else
            stbThis.Panels(2).Text = "共有 " & lngCount & " 条医嘱可以打包。"
        End If
        '自动调整行高
        .AutoSize col医嘱内容
        .Redraw = flexRDDirect
        '恢复前景色
        .Cell(flexcpForeColor, 1, col选择, .Rows - 1, col选择) = vbBlack
        If blnIsUnpack Then
            dpkReqTime(0).value = dpkExecuted(0).value
            dpkReqTime(1).value = dpkExecuted(1).value
        Else
            dpkExecuted(0).value = dpkReqTime(0).value
            dpkExecuted(1).value = dpkReqTime(1).value
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadPatiInfo()
'功能：加载病人列表
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer, j As Integer, k As Integer
    Dim str病人IDs As String, lng病区ID As Long, lngUnitID As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim lngSelectRow As Long
        
    On Error GoTo errH
    lngUnitID = mlng病区ID
    If mlng婴儿病区ID <> 0 Then
        If mlng婴儿科室ID = mlng医护科室ID Or mlng婴儿病区ID = mlng医护科室ID Then
            lngUnitID = mlng婴儿病区ID
        End If
    End If
    
    str病人IDs = zlDatabase.GetPara("发送病人", glngSys, p住院医嘱发送)
    If str病人IDs <> "" And InStr(str病人IDs, ":") > 0 Then
        lng病区ID = Val(Split(str病人IDs, ":")(0))
        str病人IDs = Split(str病人IDs, ":")(1)
    End If
            
    Set rsTmp = GetPatiRsByUnit(lngUnitID, mlng病人ID, False, False, False)
    With rptPati
        For i = 1 To rsTmp.RecordCount
            If Val(rsTmp!审核标志 & "") < 1 Or gbyt病人审核方式 <> 1 Then
                Set objRecord = .Records.Add()
                objRecord.Tag = "0"
                Set objItem = objRecord.AddItem(rsTmp!病人ID & "")
                Set objItem = objRecord.AddItem(rsTmp!主页ID & "")
                Set objItem = objRecord.AddItem("")
                Set objItem = objRecord.AddItem(rsTmp!床号 & "")
                Set objItem = objRecord.AddItem(rsTmp!姓名 & "")
                    objItem.Icon = img16.ListImages.Item(IIF(rsTmp!性别 & "" = "男", "Man", "Woman")).Index - 1
                Set objItem = objRecord.AddItem(rsTmp!性别 & "")
                Set objItem = objRecord.AddItem(rsTmp!住院号 & "")
                
                
                '病人颜色
                objRecord.Item(0).ForeColor = zlDatabase.GetPatiColor(Nvl(rsTmp!病人类型))
                For j = 1 To objRecord.Childs.Count - 1
                    objRecord.Item(j).ForeColor = objRecord.Item(0).ForeColor
                Next
                
                '上次是否选择
                If lngUnitID = lng病区ID And str病人IDs <> "" Then
                    If InStr("," & str病人IDs & ",", "," & rsTmp!病人ID & ",") > 0 Or str病人IDs = "ALL" Then
                        objRecord.Item(COL_选择).Icon = img16.ListImages.Item("Check").Index - 1
                        objRecord.Tag = "1"
                        lngSelectRow = i
                    End If
                ElseIf rsTmp!病人ID = mlng病人ID Then
                    objRecord.Item(COL_选择).Icon = img16.ListImages.Item("Check").Index - 1
                    objRecord.Tag = "1"
                    lngSelectRow = i
                End If
            End If
            rsTmp.MoveNext
        Next
        .Populate
        If lngSelectRow > 0 Then
            Set .FocusedRow = .Rows(lngSelectRow - 1)
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitTable(vsgInfo As VSFlexGrid, ByVal strHead As String)
    Dim arrHead As Variant, i As Long
    
    arrHead = Split(strHead, ";")
    With vsgInfo
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn, lngIdx As Long, i As Long

    With rptPati
        
        Set objCol = .Columns.Add(COL_病人ID, "病人ID", 0, False)
        Set objCol = .Columns.Add(COL_主页ID, "主页ID", 0, False)
        Set objCol = .Columns.Add(COL_选择, "", 20, True)
            objCol.Sortable = False
            objCol.AllowDrag = False
            objCol.Alignment = xtpAlignmentRight
            objCol.Icon = img16.ListImages("UnCheck").Index - 1
        Set objCol = .Columns.Add(COL_床号, "床号", 45, True)
        Set objCol = .Columns.Add(COL_姓名, "姓名", 80, True)
        Set objCol = .Columns.Add(COL_性别, "性别", 30, True)
        Set objCol = .Columns.Add(COL_住院号, "住院号", 60, True)
        
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '有分组列时，树形线边上会再有一根边线
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的病人..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '会引发SelectionChanged事件
        .ShowItemsInGroups = False
        .SetImageList Me.img16
    End With
End Sub

Private Sub InitPageData()
'功能：初始化界面
    Dim curDate As Date
    Dim strSQL As String
    Dim rsTmp As Recordset
    Dim i As Long
    Dim strTmp As String
    On Error GoTo errH
    curDate = zlDatabase.Currentdate
    
    dpkExecuted(0).value = Format(curDate, "yyyy-MM-dd 00:00:00")
    dpkExecuted(1).value = Format(curDate, "yyyy-MM-dd 23:59:59")
    dpkReqTime(0).value = Format(curDate, "yyyy-MM-dd 00:00:00")
    dpkReqTime(1).value = Format(curDate, "yyyy-MM-dd 23:59:59")
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    Dim str病人IDs As String
    
    '保存报表病人设置
    str病人IDs = ""
    For i = 0 To rptPati.Rows.Count - 1
        If rptPati.Rows(i).Record.Tag = "1" Then
            str病人IDs = str病人IDs & "," & rptPati.Rows(i).Record(COL_病人ID).value
        End If
    Next
    str病人IDs = Mid(str病人IDs, 2)
    If str病人IDs <> "" Then
        If UBound(Split(str病人IDs, ",")) = 0 And Val(str病人IDs) = mlng病人ID Then
            Call zlDatabase.SetPara("发送病人", "", glngSys, p住院医嘱发送)
        Else
            Call zlDatabase.SetPara("发送病人", mlng病区ID & ":" & str病人IDs, glngSys, p住院医嘱发送)
        End If
    End If

    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub optBaby_Click(Index As Integer)
    mint医嘱处理范围 = Index
End Sub

Private Sub picExecuted_Resize()
    On Error Resume Next
    vsgExecUnpack.Width = picExecuted.Width - 200
    vsgExecUnpack.Height = picExecuted.Height - vsgExecUnpack.Top
    
End Sub

Private Sub picWaitExecAdvice_Resize()
    On Error Resume Next
    vsgWaitUnpack.Top = 0
    vsgWaitUnpack.Height = picWaitExecAdvice.Height - vsgWaitUnpack.Top
    vsgWaitUnpack.Width = picWaitExecAdvice.Width
    
End Sub

Private Sub picWaitExecute_Resize()
    On Error Resume Next
    fraPatiInfo.Height = picWaitExecute.Height + 80
    rptPati.Height = fraPatiInfo.Height - rptPati.Top - picFitter.Height - 100
    fraBaby.Top = rptPati.Top + rptPati.Height + 50
    picFitter.Top = IIF(fraBaby.Visible, fraBaby.Top, rptPati.Top) + IIF(fraBaby.Visible, fraBaby.Height, rptPati.Height) + 50
    picWaitExecAdvice.Height = picWaitExecute.Height
    picWaitExecAdvice.Width = picWaitExecute.Width - picWaitExecAdvice.Left - 300
End Sub

Private Sub rptPati_KeyDown(KeyCode As Integer, Shift As Integer)
    If rptPati.SelectedRows.Count > 0 Then
        If KeyCode = vbKeySpace Then
            Call rptPati_RowDblClick(rptPati.SelectedRows(0), rptPati.SelectedRows(0).Record.Item(COL_选择))
        End If
    End If
End Sub

Private Sub rptPati_MouseUp(Button As Integer, Shift As Integer, x As Long, Y As Long)
    Dim objColumn As ReportColumn
    Dim i As Long
    
    '如果点击表头的图片，就选中全部
    If Button = 1 Then
        If rptPati.HitTest(x, Y).ht = xtpHitTestHeader Then
            Set objColumn = rptPati.HitTest(x, Y).Column
            If Not objColumn Is Nothing Then
                If objColumn.Index = COL_选择 Then
                    If objColumn.Caption = "" Then
                        objColumn.Caption = "1"
                        rptPati.Columns(COL_选择).Icon = img16.ListImages("AllCheck").Index - 1
                        For i = 0 To rptPati.Records.Count - 1
                            rptPati.Records(i)(COL_选择).Icon = img16.ListImages("Check").Index - 1
                            rptPati.Rows(i).Record.Tag = "1"
                        Next
                    Else
                        objColumn.Caption = ""
                        rptPati.Columns(COL_选择).Icon = img16.ListImages("UnCheck").Index - 1
                        For i = 0 To rptPati.Records.Count - 1
                            rptPati.Records(i)(COL_选择).Icon = -1
                            rptPati.Rows(i).Record.Tag = "0"
                        Next
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub rptPati_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Record.Tag = "1" Then
        Row.Record.Item(COL_选择).Icon = -1
        Row.Record.Tag = "0"
    Else
        Row.Record.Item(COL_选择).Icon = img16.ListImages.Item("Check").Index - 1
        Row.Record.Tag = "1"
    End If
    rptPati.Populate
End Sub

Private Function RowIn一并给药(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long, vsTmp As VSFlexGrid) As Boolean
'功能：判断指定行是否在一并给药的范围中,如果是,同时返回行号范围
    Dim i As Long, blnTmp As Boolean
    
    With vsTmp
        If .TextMatrix(lngRow, col诊疗类别) = "" Then Exit Function
        If .TextMatrix(lngRow, col诊疗类别) = "诊疗类别" Then Exit Function
        If Val(.TextMatrix(lngRow - 1, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) And Val(.TextMatrix(lngRow, col相关ID)) <> 0 Or Val(.TextMatrix(lngRow - 1, col相关ID)) = Val(.TextMatrix(lngRow, col医嘱ID)) Or Val(.TextMatrix(lngRow - 1, col医嘱ID)) = Val(.TextMatrix(lngRow, col相关ID)) And Val(.TextMatrix(lngRow - 1, col医嘱ID)) <> 0 Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) And Val(.TextMatrix(lngRow + 1, col相关ID)) <> 0 Or Val(.TextMatrix(lngRow + 1, col医嘱ID)) = Val(.TextMatrix(lngRow, col相关ID)) Or Val(.TextMatrix(lngRow + 1, col相关ID)) = Val(.TextMatrix(lngRow, col医嘱ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) And Val(.TextMatrix(lngRow, col相关ID)) <> 0 And Val(.TextMatrix(i, col医嘱ID)) <> Val(.TextMatrix(lngRow, col医嘱ID)) Or Val(.TextMatrix(i, col相关ID)) = Val(.TextMatrix(lngRow, col医嘱ID)) Or Val(.TextMatrix(i, col医嘱ID)) = Val(.TextMatrix(lngRow, col相关ID)) And Val(.TextMatrix(i, col医嘱ID)) <> 0 Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) And Val(.TextMatrix(lngRow, col相关ID)) <> 0 And Val(.TextMatrix(i, col医嘱ID)) <> Val(.TextMatrix(lngRow, col医嘱ID)) Or Val(.TextMatrix(i, col医嘱ID)) = Val(.TextMatrix(lngRow, col相关ID)) Or Val(.TextMatrix(i, col相关ID)) = Val(.TextMatrix(lngRow, col医嘱ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        Else
            .RowData(lngRow) = "Begin"
        End If
        RowIn一并给药 = blnTmp
    End With
End Function

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Not Me.Visible Then Exit Sub
    If Item.Tag = "待打包医嘱" Then
        'Call LoadAdvice
    Else
        Call LoadAdvice(True)
    End If
End Sub

Private Sub vsgExecUnpack_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If vsgExecUnpack.Cell(flexcpBackColor, NewRow, NewCol) = COLEditBackColor And vsgExecUnpack.RowData(NewRow) = "Begin" Then
        vsgExecUnpack.Editable = flexEDNone
    Else
        vsgExecUnpack.FocusRect = flexFocusNone
        vsgExecUnpack.Editable = flexEDNone
        vsgExecUnpack.ComboList = ""
    End If
End Sub

Private Sub vsgExecUnpack_Click()
    Dim i As Long
    
    With vsgExecUnpack
        If .MouseCol = col选择 And .MouseRow = .FixedRows - 1 Then
            If .TextMatrix(1, col医嘱ID) = "" Then Exit Sub
            If .ColData(col选择) = "Check" Then
                .Cell(flexcpPicture, 0, col选择) = img16.ListImages("UnCheck").Picture
                .ColData(col选择) = ""
            Else
                .Cell(flexcpPicture, 0, col选择) = img16.ListImages("AllCheck").Picture
                .ColData(col选择) = "Check"
            End If
            For i = 1 To .Rows - 1
                If .TextMatrix(i, col医嘱ID) = "" Then Exit For
                If .ColData(col选择) = "Check" And .TextMatrix(i, col销帐申请人) = "" And (.TextMatrix(i, col状态) = "待配药" Or .TextMatrix(i, col状态) = "待摆药") Then
                    .Cell(flexcpPicture, i, col选择) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, i, col选择) = 1
                    .Cell(flexcpPictureAlignment, i, col选择) = flexPicAlignCenterCenter
                Else
                    Set .Cell(flexcpPicture, i, col选择) = Nothing
                    .Cell(flexcpData, i, col选择) = 0
                End If
                
            Next
        End If
    End With
End Sub

Private Sub vsgExecUnpack_DblClick()
    With vsgExecUnpack
        If .MouseCol = col选择 And .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            Call vsgExecUnpack_KeyPress(vbKeySpace)
        End If
    End With
End Sub

Private Sub vsgExecUnpack_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    '说明：1.OwnerDraw要设置为Over(画出单元所有内容)
'      2.Cell的GridLine从上下左右向内都是从第1根线开始
'      3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsgExecUnpack
        lngLeft = col选择: lngRight = col期效
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        If Not RowIn一并给药(Row, lngBegin, lngEnd, vsgExecUnpack) Then Exit Sub
        
        vRect.Left = Left '擦除左边表格线
        vRect.Right = Right - 1 '保留右边表格线
        If .TextMatrix(Row, col相关ID) = "" Then
            vRect.Top = Bottom - 1 '相关ID为空的行文字保留
            vRect.Bottom = Bottom - 1
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 2 '底行保留下边线(本窗体中用到下边线粗为2)
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            If Between(Col, col选择, col选择) Then
                SetBkColor hDC, OS.SysColor2RGB(COLEditBackColor)
            Else
                SetBkColor hDC, OS.SysColor2RGB(.BackColor)
            End If
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsgExecUnpack_KeyPress(KeyAscii As Integer)
    With vsgExecUnpack
        If .Col = col选择 And KeyAscii = vbKeySpace Then
            Call ExecCheck(vsgExecUnpack)
        End If
    End With
End Sub

Private Sub vsgWaitUnpack_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If vsgWaitUnpack.Cell(flexcpBackColor, NewRow, NewCol) = COLEditBackColor And vsgWaitUnpack.RowData(NewRow) = "Begin" Then
        vsgWaitUnpack.FocusRect = flexFocusHeavy
        vsgWaitUnpack.Editable = flexEDNone
    Else
        vsgWaitUnpack.FocusRect = flexFocusNone
        vsgWaitUnpack.Editable = flexEDNone
        vsgWaitUnpack.ComboList = ""
    End If
End Sub

Private Sub vsgWaitUnpack_Click()
    Dim i As Long
    
    With vsgWaitUnpack
        If .MouseCol = col选择 And .MouseRow = .FixedRows - 1 Then
            If .TextMatrix(1, col医嘱ID) = "" Then Exit Sub
            If .ColData(col选择) = "Check" Then
                .Cell(flexcpPicture, 0, col选择) = img16.ListImages("UnCheck").Picture
                .ColData(col选择) = ""
            Else
                .Cell(flexcpPicture, 0, col选择) = img16.ListImages("AllCheck").Picture
                .ColData(col选择) = "Check"
            End If
            For i = 1 To .Rows - 1
                If .TextMatrix(i, col医嘱ID) = "" Then Exit For
                If .ColData(col选择) = "Check" Then
                    .Cell(flexcpPicture, i, col选择) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, i, col选择) = 1
                    .Cell(flexcpPictureAlignment, i, col选择) = flexPicAlignCenterCenter
                Else
                    Set .Cell(flexcpPicture, i, col选择) = Nothing
                    .Cell(flexcpData, i, col选择) = 0
                End If
                
            Next
        End If
    End With
End Sub

Private Sub vsgWaitUnpack_DblClick()
    With vsgWaitUnpack
        If .MouseCol = col选择 And .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            Call vsgWaitUnpack_KeyPress(vbKeySpace)
        End If
    End With
End Sub

Private Sub vsgWaitUnpack_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    '说明：1.OwnerDraw要设置为Over(画出单元所有内容)
'      2.Cell的GridLine从上下左右向内都是从第1根线开始
'      3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsgWaitUnpack
        lngLeft = col选择: lngRight = col期效
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        If Not RowIn一并给药(Row, lngBegin, lngEnd, vsgWaitUnpack) Then Exit Sub
        
        vRect.Left = Left '擦除左边表格线
        vRect.Right = Right - 1 '保留右边表格线
        If .TextMatrix(Row, col相关ID) = "" Then
            vRect.Top = Bottom - 1 '相关ID为空的行文字保留
            vRect.Bottom = Bottom - 1
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 2 '底行保留下边线(本窗体中用到下边线粗为2)
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            If Between(Col, col选择, col选择) Then
                SetBkColor hDC, OS.SysColor2RGB(COLEditBackColor)
            Else
                SetBkColor hDC, OS.SysColor2RGB(.BackColor)
            End If
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsgWaitUnpack_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode > 127 Then
        '解决直接输入汉字的问题
        Call vsgWaitUnpack_KeyPress(KeyCode)
    End If
End Sub

Private Sub vsgWaitUnpack_KeyPress(KeyAscii As Integer)
    With vsgWaitUnpack
        If .Col = col选择 And KeyAscii = vbKeySpace Then
            Call ExecCheck(vsgWaitUnpack)
        End If
    End With
End Sub

Private Sub ExecCheck(ByRef objVsg As VSFlexGrid)
'功能：同步选择一组医嘱
'参数：表格
    Dim lngBegin As Long, lngEnd As Long
    Dim i As Long
    
    With objVsg
        If .TextMatrix(.Row, col医嘱ID) = "" Then Exit Sub
        If Not RowIn一并给药(.Row, lngBegin, lngEnd, objVsg) Then
            lngBegin = .Row: lngEnd = .Row
        End If
        
        For i = lngBegin To lngEnd
            If .Cell(flexcpData, i, col选择) = 1 Then
                Set .Cell(flexcpPicture, i, col选择) = Nothing
                .Cell(flexcpData, i, col选择) = 0
            Else
                If objVsg.Name = "vsgExecUnpack" Then
                    '检查是否出院
                    If .TextMatrix(i, col在院) <> "1" Then
                        MsgBox "该病人已经出院，不能取消打包。", vbInformation, Me.Caption
                        Exit Sub
                    ElseIf .TextMatrix(i, col销帐申请人) <> "" Then
                        MsgBox "该医嘱已经销帐，不能取消打包。", vbInformation, Me.Caption
                        Exit Sub
                    ElseIf .TextMatrix(i, col状态) <> "待配药" And .TextMatrix(i, col状态) <> "待摆药" Then
                        MsgBox "该医嘱已经配药，不能取消打包。", vbInformation, Me.Caption
                        Exit Sub
                    End If
                End If
                .Cell(flexcpPicture, i, col选择) = img16.ListImages("Check").Picture
                .Cell(flexcpData, i, col选择) = 1
                .Cell(flexcpPictureAlignment, i, col选择) = flexPicAlignCenterCenter
            End If
        Next
    End With
End Sub



