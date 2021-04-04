VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmAppUpgradePrepare 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "升级准备"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12015
   Icon            =   "frmAppUpgradePrepare.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   12015
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraLine 
      Height          =   120
      Left            =   5400
      TabIndex        =   9
      Top             =   600
      Width           =   4935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6960
      Width           =   1100
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "执行(&E)"
      Height          =   350
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6960
      Width           =   1100
   End
   Begin VB.Frame fraCheck 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   4560
      TabIndex        =   3
      Top             =   1080
      Width           =   3735
      Begin VSFlex8Ctl.VSFlexGrid vsfShow 
         Height          =   1215
         Index           =   3
         Left            =   360
         TabIndex        =   17
         Top             =   360
         Width           =   2895
         _cx             =   5106
         _cy             =   2143
         Appearance      =   1
         BorderStyle     =   0
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
      Begin VB.Label lblCheck 
         AutoSize        =   -1  'True
         Caption         =   "说明："
         Height          =   180
         Left            =   840
         TabIndex        =   11
         Top             =   2040
         Width           =   540
      End
   End
   Begin VB.Frame fraJob 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   8640
      TabIndex        =   2
      Top             =   1200
      Width           =   3495
      Begin VB.CheckBox chkShow 
         Caption         =   "显示停用的后台作业"
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   2175
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfShow 
         Height          =   855
         Index           =   2
         Left            =   720
         TabIndex        =   16
         Top             =   600
         Width           =   1575
         _cx             =   2778
         _cy             =   1508
         Appearance      =   1
         BorderStyle     =   0
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
      Begin VB.Label lblJob 
         AutoSize        =   -1  'True
         Caption         =   "说明："
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   540
      End
   End
   Begin VB.Frame fraUser 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   6600
      TabIndex        =   1
      Top             =   4080
      Width           =   4695
      Begin VB.CheckBox chkShow 
         Caption         =   "显示停用的用户"
         Height          =   495
         Index           =   1
         Left            =   960
         TabIndex        =   20
         Top             =   1800
         Width           =   2175
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfShow 
         Height          =   975
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   2775
         _cx             =   4895
         _cy             =   1720
         Appearance      =   1
         BorderStyle     =   0
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         Caption         =   "说明："
         Height          =   180
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   540
      End
   End
   Begin VB.Frame fraClient 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   3840
      Width           =   4935
      Begin VB.CheckBox chkShow 
         Caption         =   "显示停用的客户端"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   120
         Width           =   2175
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfShow 
         Height          =   1215
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   2295
         _cx             =   4048
         _cy             =   2143
         Appearance      =   1
         BorderStyle     =   0
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
      Begin VB.CommandButton cmdkillProcess 
         Caption         =   "中断客户端连接的进程定义(&P)"
         Height          =   350
         Left            =   360
         TabIndex        =   5
         Top             =   1920
         Width           =   2790
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         Caption         =   "说明："
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   540
      End
   End
   Begin XtremeSuiteControls.TabControl tbcMain 
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   4335
      _Version        =   589884
      _ExtentX        =   7646
      _ExtentY        =   4260
      _StockProps     =   64
   End
   Begin VB.Label lblResult 
      AutoSize        =   -1  'True
      Caption         =   "结果："
      Height          =   180
      Left            =   360
      TabIndex        =   21
      Top             =   7560
      Width           =   540
   End
   Begin VB.Image imgMain 
      Height          =   615
      Left            =   120
      Picture         =   "frmAppUpgradePrepare.frx":6852
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      Caption         =   "说明"
      Height          =   180
      Left            =   840
      TabIndex        =   8
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "frmAppUpgradePrepare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SQL_CAPTION = "其他前置检查"
Private Const mstrOracleUser      As String = "'ANONYMOUS','AURORA$JIS$UTILITY$','AURORA$ORB$UNAUTHENTICATED','CTXSYS','DBSNMP','DIP','DMSYS','DVF','DVSYS','EXFSYS','HR','LBACSYS','MDDATA','MDSYS','MGMT_VIEW','OAS_PUBLIC','ODM','ODM_MTR','OE','OGG','OLAPSYS','ORDPLUGINS','ORDSYS','OSE$HTTP$ADMIN','OUTLN','PERFSTAT','PM','QS','QS_ADM','QS_CB','QS_CBADM','QS_CS','QS_ES','QS_OS','QS_WS','REPADMIN','RMAN','SCOTT','SH','SI_INFORMTN_SCHEMA','SYSMAN','TRACESVR','TSMSYS','WEBSYS','WKPROXY','WKSYS','WKUSER','WK_TEST','WMSYS','XDB'"
Private Enum enuTab
    T_客户端 = 0
    T_用户账号
    T_后台作业
    T_其他
End Enum
Private mstrSysNum As String
Private mstrUsers As String
Private mbln10g As Boolean
Private mblnFirst As Boolean
Private mstrCondition As String
Private mcnExe As ADODB.Connection
'该窗体是展示系统升级前的一些准备，为了不影响正常升级，将忽略错误
Private Sub InitTabControl()
    
    With tbcMain
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .BoldSelected = True
            .Color = xtpTabColorDefault
            .ShowIcons = False
        End With
        Call .InsertItem(T_客户端, "中断客户端连接并禁止登录", fraClient.hwnd, 0)
        Call .InsertItem(T_用户账号, "锁定用户账号", fraUser.hwnd, 0)
        Call .InsertItem(T_后台作业, "禁用后台作业", fraJob.hwnd, 0)
        Call .InsertItem(T_其他, "影响升级效率的其他调整", fraCheck.hwnd, 0)
    End With
End Sub

Private Sub chkShow_Click(Index As Integer)
    Select Case Index
        Case T_客户端
            Call LoadClient
        Case T_用户账号
            mblnFirst = True
            Call LoadUser
        Case T_后台作业
            Call LoadJob
    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdExec_Click()
    Dim i As Long, j As Long, lngNum As Long, lngInstId As Long
    Dim rsChoose As ADODB.Recordset
    Dim varSQL As Variant
    Dim strErrContent As String, strErr As String, strTemp As String
    Dim strClient As String
    Dim cnTemp As ADODB.Connection

    If ExeCheck = False Then Exit Sub
    On Error GoTo ErrH
    Set rsChoose = CopyNewRec(Nothing, True, , Array("类型", adVarChar, 10, Empty, "停用SQL", adVarChar, 500, Empty, _
                    "名称", adVarChar, 200, Empty))
    For i = vsfShow.LBound To vsfShow.UBound
        With vsfShow(i)
            For j = 1 To .Rows - 1
                If .Cell(flexcpChecked, j, .ColIndex("选择")) = flexChecked Then
                    If i = T_客户端 Then
                        If InStr(strClient, "'" & .TextMatrix(j, .ColIndex("客户端")) & "'") = 0 Then
                            strClient = IIf(strClient = "", "", strClient & ",") & "'" & .TextMatrix(j, .ColIndex("客户端")) & "'"
                        End If
                        rsChoose.AddNew Array("类型", "停用SQL", "名称"), Array("客户端", .TextMatrix(j, .ColIndex("停用SQL")), _
                        .TextMatrix(j, .ColIndex("INST_ID")) & "|" & .TextMatrix(j, .ColIndex("当前标志")))
                    ElseIf i = T_用户账号 Then
                        rsChoose.AddNew Array("类型", "停用SQL"), Array("用户账号", .TextMatrix(j, .ColIndex("停用SQL")))
                    ElseIf i = T_后台作业 Then
                        If .TextMatrix(j, .ColIndex("类型")) = "系统调度" Then
                            rsChoose.AddNew Array("类型", "停用SQL", "名称"), Array("系统调度", .TextMatrix(j, .ColIndex("停用SQL")), .TextMatrix(j, .ColIndex("内容")))
                        '非产品自动作业需记录在zlUpgradeConfig中，故将作业号放在 名称列中
                        ElseIf .TextMatrix(j, .ColIndex("类型")) = "非产品自动作业" Then
                            rsChoose.AddNew Array("类型", "停用SQL", "名称"), Array("后台作业", .TextMatrix(j, .ColIndex("停用SQL")), .TextMatrix(j, .ColIndex("作业号")))
                        Else
                            rsChoose.AddNew Array("类型", "停用SQL"), Array("后台作业", .TextMatrix(j, .ColIndex("停用SQL")))
                        End If
                    ElseIf i = T_其他 Then
                        If .TextMatrix(j, .ColIndex("分类")) = "触发器" Then
                            rsChoose.AddNew Array("类型", "停用SQL", "名称"), Array("触发器", .TextMatrix(j, .ColIndex("停用SQL")), .TextMatrix(j, .ColIndex("对象")))
                        Else
                            rsChoose.AddNew Array("类型", "停用SQL"), Array("其他设置", .TextMatrix(j, .ColIndex("停用SQL")))
                        End If
                    End If
                End If
            Next
        End With
    Next
    If rsChoose.RecordCount = 0 Then Unload Me: Exit Sub
    If MsgBox("确认要执行勾选的这些调整吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    Call ShowFlash("正在执行所勾选调整的相关SQL,请稍候...")
    '锁定用户账号
    rsChoose.Filter = "类型='用户账号'"
    Do While Not rsChoose.EOF
        strErrContent = ""
        varSQL = Split(rsChoose!停用SQL, "分隔符")
        For i = LBound(varSQL) To UBound(varSQL)
            strErrContent = strErrContent & gclsBase.ExecuteCmdText(varSQL(i), Me.Caption, mcnExe, True)
            '如果账号锁定失败，则不改变升级停用标记值
            If i = 0 And strErrContent <> "" Then Exit For
        Next
        If strErrContent <> "" Then
            lngNum = lngNum + 1
        End If
        rsChoose.MoveNext
    Loop
    strErr = IIf(lngNum = 0, strErr, strErr & "锁定用户账号失败" & lngNum & "个;")
    lngNum = 0
    '禁用系统调度
    strTemp = ""
    rsChoose.Filter = "类型='系统调度'"
    If mbln10g Then
        Call ShowFlash("")
        Set cnTemp = GetConnection("SYS")
        Call ShowFlash("正在执行所勾选调整的相关SQL,请稍候...")
    Else
        Set cnTemp = mcnExe
    End If
    Do While Not rsChoose.EOF
        strErrContent = ""
        strErrContent = gclsBase.ExecuteCmdText(rsChoose!停用SQL, Me.Caption, cnTemp, True)
        If strErrContent = "" Then
            strTemp = IIf(strTemp = "", "", strTemp & ",") & rsChoose!名称
        Else
            lngNum = lngNum + 1
        End If
        rsChoose.MoveNext
    Loop
    If strTemp <> "" Then
        gstrSQL = "Update Zlupgradeconfig Set 内容='" & strTemp & "' Where 项目='禁用的系统调度'"
        Call gclsBase.ExecuteCmdText(gstrSQL, Me.Caption, mcnExe)
    End If
    strTemp = ""
    '禁用后台作业
    rsChoose.Filter = "类型='后台作业'"
    Do While Not rsChoose.EOF
        strErrContent = ""
        varSQL = Split(rsChoose!停用SQL, "分隔符")
        For i = LBound(varSQL) To UBound(varSQL)
            On Error Resume Next
            '后台作业不能用adCmdText类型执行
            gcnOracle.Execute varSQL(i)
            '后台作业停用失败，不改变升级停用标记值
            If i = 0 And err.Number <> 0 Then Exit For
        Next
        If err.Number = 0 Then
            '此名称对于非产品自动作业，保存的是作业号，方便启用
            If "" & rsChoose!名称 <> "" Then
                strTemp = IIf(strTemp = "", "", strTemp & ",") & rsChoose!名称
            End If
        Else
            err.Clear
            lngNum = lngNum + 1
        End If
        rsChoose.MoveNext
    Loop
    On Error GoTo ErrH
    If strTemp <> "" Then
        '非产品后台作业需单独保存在Zlupgradeconfig中
        gstrSQL = "Update Zlupgradeconfig Set 内容='" & strTemp & "' Where 项目='禁用的后台作业'"
        Call gclsBase.ExecuteCmdText(gstrSQL, Me.Caption, mcnExe)
    End If
    strErr = IIf(lngNum = 0, strErr, strErr & "停用后台作业失败" & lngNum & "个;")
    lngNum = 0
    '禁用触发器,只能本身所有者运行
    rsChoose.Filter = "类型='触发器'"
    Do While Not rsChoose.EOF
        strErrContent = ""
        varSQL = Split(rsChoose!名称, ".")
        If varSQL(0) = UCase(gstrUserName) Then
            Set cnTemp = gcnOracle
        ElseIf varSQL(0) = "ZLTOOLS" Then
            Call ShowFlash("")
            Set cnTemp = GetConnection("ZLTOOLS")
            Call ShowFlash("正在执行所勾选调整的相关SQL,请稍候...")
        Else
            Call ShowFlash("")
            Set cnTemp = GetConnection(varSQL(0))
            Call ShowFlash("正在执行所勾选调整的相关SQL,请稍候...")
        End If
        strErrContent = strErrContent & gclsBase.ExecuteCmdText(rsChoose!停用SQL, Me.Caption, cnTemp, True)
        If strErrContent = "" Then
            gstrSQL = "Insert Into Zltriggers(所有者, 名称) Values ('" & varSQL(0) & "','" & varSQL(1) & "')"
            Call gclsBase.ExecuteCmdText(gstrSQL, Me.Caption, mcnExe)
        Else
            lngNum = lngNum + 1
        End If
        rsChoose.MoveNext
    Loop
    '禁用其他设置
    rsChoose.Filter = "类型='其他设置'"
    Do While Not rsChoose.EOF
        Call ExeSQL(mcnExe, rsChoose!停用SQL, lngNum)
        rsChoose.MoveNext
    Loop
    strErr = IIf(lngNum = 0, strErr, strErr & "调整影响升级执行效率的参数等失败" & lngNum & "个;")
    lngNum = 0
    '禁用客户端并杀掉会话
    rsChoose.Filter = "类型='客户端'"
    If rsChoose.RecordCount > 0 Then
        rsChoose.Sort = "名称 desc"
        gstrSQL = "Update Zlclients Set 禁止使用 = 1, 系统升级禁用 = 1 Where Nvl(禁止使用, 0) = 0 and Nvl(系统升级禁用, 0) = 0 and 工作站 in(" & strClient & ")"
        Call gclsBase.ExecuteCmdText(gstrSQL, Me.Caption, mcnExe, True)
        If mbln10g Then
            Do While Not rsChoose.EOF
                varSQL = Split(rsChoose!名称, "|")
                '10g环境，并且当前标志为0
                If varSQL(1) = 0 Then
                    If lngInstId <> varSQL(0) Then
                        lngInstId = varSQL(0)
                        Set cnTemp = GetInstance(lngInstId)
                        Call ExeSQL(cnTemp, rsChoose!停用SQL, lngNum)
                    Else
                        Call ExeSQL(cnTemp, rsChoose!停用SQL, lngNum)
                    End If
                Else
                    Call ExeSQL(mcnExe, rsChoose!停用SQL, lngNum)
                End If
                rsChoose.MoveNext
            Loop
        Else
            Do While Not rsChoose.EOF
                Call ExeSQL(mcnExe, rsChoose!停用SQL, lngNum)
                rsChoose.MoveNext
            Loop
        End If
    End If
    strErr = IIf(lngNum = 0, strErr, strErr & "禁用客户端并杀掉会话失败" & lngNum & "个;")
    Call ShowFlash
    If strErr <> "" Then
        MsgBox strErr, vbExclamation, Me.Caption
    End If
    Unload Me
    Exit Sub
ErrH:
    Call ShowFlash
    MsgBox err.Description, vbExclamation, Me.Caption
End Sub

Private Function ExeCheck() As Boolean
'执行前的检查
    
    If Not CheckAndAdjustMustTable("zlUpgradeConfig", , False) Then
        MsgBox "无法自动重建zlUpgradeConfig表，请手工杀掉会话,之后点击确定即可！", vbInformation, gstrSysName
        Exit Function
    End If
    If Not CheckAndAdjustMustTable("Zlclients", "系统升级禁用", False) Then
        MsgBox "无法自动升级Zlclients表，请手工杀掉会话,之后点击确定即可！", vbInformation, gstrSysName
        Exit Function
    End If
    
    ExeCheck = True
End Function

Private Sub ExeSQL(ByVal cnExe As ADODB.Connection, ByVal strSQL As String, ByRef lngNum As Long)
    Dim strErr As String
    
    strErr = gclsBase.ExecuteCmdText(strSQL, Me.Caption, cnExe, True)
    If strErr <> "" Then
        lngNum = lngNum + 1
    End If
End Sub

Private Function GetInstance(ByVal lngInstId As Long) As ADODB.Connection
'根据INST_ID获取实例连接
    Dim rsTemp As ADODB.Recordset
    Dim cnTemp As ADODB.Connection
    Dim strTemp As String
    
    On Error GoTo ErrH
    gstrSQL = "select a.inst_ID, a.Instance_Name, a.Host_name, b.NAME, b.DBID" & vbNewLine & _
            "  from gv$instance a, gv$database b" & vbNewLine & _
            " where a.INST_ID = b.INST_ID" & vbNewLine & _
            "   and a.INST_ID <> userenv('instance')" & vbNewLine & _
            "   and a.STATUS = 'OPEN' and a.INST_ID=[1]"
    Set rsTemp = gclsBase.OpenSQLRecord(mcnExe, gstrSQL, "获取实例信息", lngInstId)
    If Not rsTemp.EOF Then
        strTemp = rsTemp!INST_ID & "," & rsTemp!DBID & "," & rsTemp!Instance_Name & "(" & rsTemp!name & ")"
        If frmUserCheckLogin.ShowLogin(UCT_RACInsUser, cnTemp, gstrUserName, "", "", strTemp) Then
            Set GetInstance = cnTemp
        Else
            Set GetInstance = Nothing
        End If
    End If
    Exit Function
ErrH:
    MsgBox err.Description, vbExclamation, "获取实例连接"
End Function

Private Sub cmdkillProcess_Click()
    frmKillProcessManage.ShowMe ("0102")
    Call LoadClient
End Sub

Private Sub Form_Load()
    Dim strHead As String
    
    mblnFirst = True
    mstrCondition = ""
    Call ShowFlash("正在加载升级准备信息，请稍候！")
    mbln10g = GetOracleVersion(True, True) < 11
    Call InitTabControl
    lblTip.Caption = "为了保障升级操作高效执行，升级前应禁用客户端，锁定用户账号，停用后台作业、触发器，并优化调整相关数据库参数。" & vbNewLine & "升级完成后将自动恢复禁用的项目，如果升级异常中断，可在升级主界面中执行操作【恢复升级准备期间调整的项目】。"
    strHead = " ,300,1;客户端,1500,1;IP,1500,1;部门,1500,1;院区,1200,1;用途,1000,1;SID,500,1;SERIAL#,800,1;PROGRAM,2000,1;状态,500,1;INST_ID,0,1;当前标志,0,1;停用SQL,0,1"
    Call iniVsf(strHead, vsfShow(T_客户端))
    strHead = " ,300,1;部门,1800,1;用户名,1500,1;姓名,1200,1;状态,500,1;停用SQL,0,1"
    Call iniVsf(strHead, vsfShow(T_用户账号))
    strHead = " ,300,1;系统,2000,1;类型,1200,1;作业号,800,1;名称,2200,1;内容,3000,1;下次执行时间,1200,1;状态,500,1;停用SQL,0,1"
    Call iniVsf(strHead, vsfShow(T_后台作业))
    strHead = " ,300,1;分类,800,1;对象,2000,1;当前值,1000,1;建议值,1000,1;影响说明,2500,1;处理说明,1500,1;停用SQL,0,1"
    Call iniVsf(strHead, vsfShow(T_其他))
    
    Call LoadClient
    Call LoadUser
    Call LoadJob
    Call LoadOther
    Call ChooseAll
    lblClient.Caption = "说明：正在运行的客户端可能会导致升级脚本执行被阻塞、执行缓慢。"
    lblUser.Caption = "说明：如果未锁定用户帐号，可能导致杀掉的客户端会话重新连接到数据库。"
    lblJob.Caption = "说明：如果不禁用后台作业及系统调度，可能导致升级脚本执行被阻塞、执行缓慢。"
    lblCheck.Caption = "说明：如果不调整相关数据库参数、触发器，可能导致数据修正、索引创建执行缓慢。"
    Call frmColChoose.ClearCol
    vsfShow(T_用户账号).Cell(flexcpPicture, 0, vsfShow(T_用户账号).ColIndex("部门")) = frmColChoose.imgChoose.ListImages("NoFilter").Picture
    cmdkillProcess.Visible = CheckAndAdjustMustTable("zlkillprocess")
    Call ShowFlash("")
    Call FocusRow
End Sub

Private Sub ChooseAll()
'窗体加载时，如果有数据，默认全部勾选
    Dim i As Long, j As Long
    
    For i = vsfShow.LBound To vsfShow.UBound
        For j = 0 To vsfShow.Item(i).Rows - 1
            '由于有的不能自动修正，这里需判断是否可勾选
            If vsfShow.Item(i).Cell(flexcpChecked, j, vsfShow.Item(i).ColIndex("选择")) = flexUnchecked Then
                vsfShow.Item(i).Cell(flexcpChecked, j, vsfShow.Item(i).ColIndex("选择")) = flexChecked
            End If
        Next
    Next
    For i = vsfShow.LBound To vsfShow.UBound
        If vsfShow(i).Rows > 1 Then
            tbcMain.Item(i).Selected = True
            Exit For
        Else
            If i = vsfShow.UBound Then
                tbcMain.Item(0).Selected = True
            End If
        End If
    Next
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    lblTip.Move imgMain.Left + imgMain.Width + 100, 150, Me.Width - 200
    tbcMain.Move 0, imgMain.Top + imgMain.Height + 100, Me.Width, Me.Height - imgMain.Height * 3 - 100
    fraLine.Move 0, tbcMain.Top - fraLine.Height, Me.Width
    cmdCancel.Move Me.Width - cmdCancel.Width - 600, tbcMain.Height + tbcMain.Top + 200
    cmdExec.Move cmdCancel.Left - cmdExec.Width - 200, cmdCancel.Top
    lblResult.Move tbcMain.Left + 60, cmdCancel.Top
    Select Case tbcMain.Selected.Index
        Case T_客户端
            vsfShow(T_客户端).Move 50, 0, fraClient.Width - 150, fraClient.Height - 500
            chkShow(T_客户端).Move cmdExec.Left + cmdExec.Width, vsfShow(T_客户端).Height + 50
            cmdkillProcess.Move chkShow(T_客户端).Left - cmdkillProcess.Width - 500, chkShow(T_客户端).Top + 50
            lblClient.Move 60, vsfShow(T_客户端).Height + 200
        Case T_用户账号
            vsfShow(T_用户账号).Move 50, 0, fraUser.Width - 150, fraUser.Height - 500
            lblUser.Move 60, vsfShow(T_用户账号).Height + 200
            chkShow(T_用户账号).Move cmdExec.Left, lblUser.Top - 150
        Case T_后台作业
            vsfShow(T_后台作业).Move 50, 0, fraJob.Width - 150, fraJob.Height - 500
            lblJob.Move 60, vsfShow(T_后台作业).Height + 200
            chkShow(T_后台作业).Move cmdExec.Left, lblJob.Top - 150
        Case T_其他
            vsfShow(T_其他).Move 50, 0, fraCheck.Width - 150, fraCheck.Height - 500
            lblCheck.Move 60, vsfShow(T_其他).Height + 200
    End Select
End Sub

Private Sub tbcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    
    Call Form_Resize
    lblResult.Caption = "结果：共" & vsfShow(tbcMain.Selected.Index).Rows - 1 & "条数据。"
    Call FocusRow
End Sub

Private Sub FocusRow()
'窗体加载时，vsf可能全选不了，先将焦点放在表格中则可以解决此问题
    Dim lngRow As Long
    
    If vsfShow(tbcMain.Selected.Index).Rows > 1 Then
        If vsfShow(tbcMain.Selected.Index).Row > 0 Then
            lngRow = vsfShow(tbcMain.Selected.Index).Row
        Else
            lngRow = 1
        End If
        vsfShow(tbcMain.Selected.Index).Row = lngRow
        Call vsfShow(tbcMain.Selected.Index).ShowCell(lngRow, 1)
    End If
End Sub

Private Sub iniVsf(strHead As String, vsfData As VSFlexGrid)
    
    Call InitTable(vsfData, strHead)
    
    With vsfData
        .ColKey(0) = "选择"
        .Editable = flexEDKbdMouse
        .ExtendLastCol = True
        .MergeCells = flexMergeRestrictColumns
        .SelectionMode = flexSelectionByRow
        .AllowSelection = False
        .RowHeightMin = 300
        .AllowUserResizing = flexResizeColumns
        .ExplorerBar = flexExSortShow
    End With
End Sub

Private Sub LoadClient()
'加载客户端数据
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    
    Set rsTemp = GetData(T_客户端)
    With vsfShow(T_客户端)
        .Redraw = flexRDNone
        .Rows = 1
        .Rows = rsTemp.RecordCount + 1
        Do While Not rsTemp.EOF
            i = i + 1
            .TextMatrix(i, .ColIndex("客户端")) = "" & rsTemp!客户端
            .TextMatrix(i, .ColIndex("IP")) = "" & rsTemp!IP
            .TextMatrix(i, .ColIndex("部门")) = "" & rsTemp!部门
            .TextMatrix(i, .ColIndex("院区")) = "" & rsTemp!院区
            .TextMatrix(i, .ColIndex("用途")) = "" & rsTemp!用途
            .TextMatrix(i, .ColIndex("状态")) = rsTemp!状态
            .TextMatrix(i, .ColIndex("PROGRAM")) = "" & rsTemp!Program
            .TextMatrix(i, .ColIndex("SID")) = "" & rsTemp!Sid
            .TextMatrix(i, .ColIndex("SERIAL#")) = "" & rsTemp!SERIAL
            .TextMatrix(i, .ColIndex("INST_ID")) = "" & rsTemp!INST_ID
            .TextMatrix(i, .ColIndex("当前标志")) = "" & rsTemp!当前标志
            .TextMatrix(i, .ColIndex("停用SQL")) = "" & rsTemp!停用SQL
            If rsTemp!状态 = "INACTIVE" Then
                .TextMatrix(i, .ColIndex("选择")) = " "
            Else
                .Cell(flexcpChecked, i, .ColIndex("选择")) = flexUnchecked
            End If
            rsTemp.MoveNext
        Loop
        .TextMatrix(0, 0) = ""
        .Cell(flexcpChecked, 0, 0) = flexUnchecked
        .Redraw = flexRDDirect
    End With
    lblResult.Caption = "结果：共" & vsfShow(tbcMain.Selected.Index).Rows - 1 & "条数据。"
End Sub

Private Sub LoadUser()
'加载用户账号
    Dim rsTemp As ADODB.Recordset
    Dim strDept As String
    Dim i As Long
    
    Set rsTemp = GetData(T_用户账号)
    With vsfShow(T_用户账号)
        .Redraw = flexRDNone
        .Rows = 1
        .Rows = rsTemp.RecordCount + 1
        Do While Not rsTemp.EOF
            i = i + 1
            .TextMatrix(i, .ColIndex("部门")) = rsTemp!部门
            .TextMatrix(i, .ColIndex("用户名")) = rsTemp!用户名
            .TextMatrix(i, .ColIndex("姓名")) = rsTemp!姓名
            .TextMatrix(i, .ColIndex("状态")) = rsTemp!状态
            .TextMatrix(i, .ColIndex("停用SQL")) = rsTemp!停用SQL
            If rsTemp!状态 = "INACTIVE" Then
                .TextMatrix(i, .ColIndex("选择")) = " "
            Else
                .Cell(flexcpChecked, i, .ColIndex("选择")) = flexUnchecked
            End If
            If mblnFirst Then
                If InStr(strDept, "," & rsTemp!部门 & ",") = 0 Then
                    strDept = IIf(strDept = "", ",", strDept) & rsTemp!部门 & ","
                End If
            End If
            rsTemp.MoveNext
        Loop
        If mblnFirst Then
            .ColData(.ColIndex("部门")) = strDept
        End If
        .TextMatrix(0, 0) = ""
        .Cell(flexcpChecked, 0, 0) = flexUnchecked
        .Redraw = flexRDDirect
    End With
    lblResult.Caption = "结果：共" & vsfShow(tbcMain.Selected.Index).Rows - 1 & "条数据。"
    mblnFirst = False
End Sub

Private Sub LoadJob()
'加载后台作业
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    
    Set rsTemp = GetData(T_后台作业)
    With vsfShow(T_后台作业)
        .Redraw = flexRDNone
        .Rows = 1
        .Rows = rsTemp.RecordCount + 1
        Do While Not rsTemp.EOF
            i = i + 1
            .TextMatrix(i, .ColIndex("系统")) = "" & rsTemp!系统
            .TextMatrix(i, .ColIndex("类型")) = "" & rsTemp!类型
            .TextMatrix(i, .ColIndex("名称")) = "" & rsTemp!名称
            .TextMatrix(i, .ColIndex("内容")) = "" & rsTemp!内容
            .TextMatrix(i, .ColIndex("作业号")) = "" & rsTemp!作业号
            .TextMatrix(i, .ColIndex("状态")) = rsTemp!状态
            .TextMatrix(i, .ColIndex("下次执行时间")) = "" & rsTemp!下次执行时间
            .TextMatrix(i, .ColIndex("停用SQL")) = rsTemp!停用SQL
            If rsTemp!状态 = "INACTIVE" Then
                .TextMatrix(i, .ColIndex("选择")) = " "
            Else
                .Cell(flexcpChecked, i, .ColIndex("选择")) = flexUnchecked
            End If
            rsTemp.MoveNext
        Loop
        .TextMatrix(0, 0) = ""
        .Cell(flexcpChecked, 0, 0) = flexUnchecked
        .Redraw = flexRDDirect
    End With
    lblResult.Caption = "结果：共" & vsfShow(tbcMain.Selected.Index).Rows - 1 & "条数据。"
End Sub

Private Sub LoadOther()
'加载其他数据
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    
    Set rsTemp = GetData(T_其他)
    With vsfShow(T_其他)
        .Redraw = flexRDNone
        .Rows = 1
        .Rows = rsTemp.RecordCount + 1
        If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            i = i + 1
            .TextMatrix(i, .ColIndex("分类")) = rsTemp!分类
            .TextMatrix(i, .ColIndex("对象")) = rsTemp!对象
            .TextMatrix(i, .ColIndex("当前值")) = rsTemp!当前值
            .TextMatrix(i, .ColIndex("建议值")) = rsTemp!建议值
            .TextMatrix(i, .ColIndex("影响说明")) = rsTemp!影响说明
            .TextMatrix(i, .ColIndex("处理说明")) = rsTemp!处理说明
            .TextMatrix(i, .ColIndex("停用SQL")) = "" & rsTemp!停用SQL
            If rsTemp!停用SQL = "" Then
                .TextMatrix(i, 0) = " "
            Else
                .Cell(flexcpChecked, i, 0) = flexUnchecked
            End If
            rsTemp.MoveNext
        Loop
        .TextMatrix(0, 0) = ""
        .Cell(flexcpChecked, 0, 0) = flexUnchecked
        .Redraw = flexRDDirect
    End With
    lblResult.Caption = "结果：共" & vsfShow(tbcMain.Selected.Index).Rows - 1 & "条数据。"
End Sub

Private Function GetData(ByVal intIndex As Integer) As ADODB.Recordset
'获取所需要的数据
    Dim rsTemp As ADODB.Recordset
    Dim strKillProcess As String
    
    On Error GoTo ErrH
    Select Case intIndex
        Case T_客户端
            strKillProcess = GetkillProcess
            gstrSQL = "Select b.工作站 客户端, b.Ip, Decode(b.站点, Null, '全院', c.名称) 院区,b.部门, a.Program, b.用途, decode(b.禁止使用,0,'ACTIVE',1,'INACTIVE') 状态, a.Sid," & vbNewLine & _
                "       a.Serial# Serial," & IIf(gblnRac, " a.INST_ID,  Decode(INST_ID, userenv('instance'), 1, 0) 当前标志", "userenv('instance') INST_ID,1 当前标志") & "," & vbNewLine & _
                "'alter system kill session ' || Chr(39) || a.Sid || ',' || a.Serial# || " & IIf(mbln10g, "", IIf(gblnRac, "',@' || a.INST_ID ||", "',@' || userenv('instance') || ")) & " Chr(39) || ' immediate' 停用SQL" & vbNewLine & _
                "From " & IIf(gblnRac, "G", "") & "v$session a, Zlclients b, Zlnodelist c" & vbNewLine & _
                "Where a.Terminal = b.工作站(+) And Upper(a.Program) In (" & strKillProcess & ") And" & vbNewLine & _
                "b.站点= c.编号(+) And a.STATUS != 'KILLED' And a.USERNAME is Not Null And" & vbNewLine & _
                "      (a.Terminal <> Userenv('terminal') Or" & vbNewLine & _
                "      a.Terminal = Userenv('terminal') And Upper(a.Program) Not In ('VB6.EXE', 'ZLSVRSTUDIO.EXE'))" & vbNewLine & _
                IIf(chkShow(T_客户端).value = 0, " And b.禁止使用=0 ", "") & vbNewLine & _
                "Order By INST_ID, a.Terminal, a.Program"
        Case T_用户账号
            gstrSQL = "Select Null As 选择, e.名称 部门, b.用户名, c.姓名, Decode(a.Account_Status, 'OPEN', 'ACTIVE', 'INACTIVE') 状态," & vbNewLine & _
                "'alter user ' || b.用户名 || ' account lock 分隔符Update 上机人员表 Set 系统升级锁定 = 1 Where 用户名='||Chr(39)||b.用户名||Chr(39) 停用sql" & vbNewLine & _
                "From Dba_Users a, 上机人员表 b, 人员表 c, 部门人员 d, 部门表 e" & vbNewLine & _
                "Where a.Username = b.用户名 And b.用户名 <> '" & gstrUserName & "' " & "And b.人员id = c.Id And c.Id = d.人员id And d.部门id = e.Id And d.缺省=1 " & vbNewLine & _
                IIf(chkShow(T_用户账号).value = 0, " And a.Account_Status = 'OPEN' ", "") & mstrCondition & vbNewLine & _
                "Order By 部门, b.用户名, c.姓名"
        Case T_后台作业
            '系统调度
            If mbln10g Then
                gstrSQL = "Select Null 编号, Null 系统, '系统调度' 类型, decode(a.Job_Name,'GATHER_STATS_JOB','自动统计信息收集','自动分段顾问') 名称, a.Owner || '.' || a.Job_Name 内容, Null 作业号," & vbNewLine & _
                    "       Decode(a.Enabled, 'TRUE', 'ACTIVE', 'FALSE', 'INACTIVE') 状态,Null 下次执行时间," & vbNewLine & _
                    "       'Call dbms_scheduler.disable(' || Chr(39) || a.Owner || '.' || a.Job_Name || Chr(39) || ')' 停用sql" & vbNewLine & _
                    "From Dba_Scheduler_Jobs a" & vbNewLine & _
                    "Where a.Job_Name In ('GATHER_STATS_JOB', 'AUTO_SPACE_ADVISOR_JOB')" & IIf(chkShow(T_后台作业).value = 0, " And a.Enabled = 'TRUE' ", "")
            Else
                gstrSQL = "Select Null 编号, Null 系统, '系统调度' 类型,decode(a.Client_Name,'auto optimizer stats collection','自动优化器统计收集','自动分段顾问') 名称, a.Client_Name 内容, Null 作业号," & vbNewLine & _
                    "       Decode(a.Status, 'ENABLED', 'ACTIVE', 'DISABLED', 'INACTIVE') 状态,Null 下次执行时间," & vbNewLine & _
                    "       'Call DBMS_AUTO_TASK_ADMIN.DISABLE(client_name => ' || Chr(39) || a.Client_Name || Chr(39) ||" & vbNewLine & _
                    "        ',operation => NULL,window_name => NULL)' 停用sql" & vbNewLine & _
                    "From Dba_Autotask_Client a" & vbNewLine & _
                    "Where a.Client_Name In ('auto optimizer stats collection', 'auto space advisor')" & IIf(chkShow(T_后台作业).value = 0, " And a.Status = 'ENABLED' ", "")
            End If
            '后台作业
            gstrSQL = gstrSQL & " Union All " & "Select c.编号, c.名称 系统, Decode(a.类型, 1, '系统设定', 2, '数据转移', 3, '用户自定义') 类型, a.名称, a.内容, a.作业号," & vbNewLine & _
                "       Decode(b.Broken, 'N', 'ACTIVE', 'INACTIVE') 状态,b.Next_date 下次执行时间," & vbNewLine & _
                " 'Dbms_Job.Broken('||a.作业号||',True)分隔符Update Zlautojobs Set 系统升级停用 = 1 Where 作业号='||a.作业号 停用sql" & vbNewLine & _
                "From Zlautojobs a, User_Jobs b, Zlsystems c" & vbNewLine & _
                "Where b.Job = a.作业号 And a.系统 = c.编号" & IIf(chkShow(T_后台作业).value = 0, " And b.Broken = 'N' ", "") & IIf(mstrSysNum = "", "", " And c.编号 In(" & mstrSysNum & ") ")
            '非产品自动作业
            gstrSQL = gstrSQL & " Union All " & "Select Null 编号, Null 系统, '非产品自动作业' 类型, Null 名称, a.What 内容, a.Job 作业号," & vbNewLine & _
                "       Decode(a.Broken, 'N', 'ACTIVE', 'INACTIVE') 状态,a.Next_date 下次执行时间,'dbms_Job.Broken('||a.Job||',True)' 停用sql" & vbNewLine & _
                "From User_Jobs a" & vbNewLine & _
                "Where a.Job Not In (Select 作业号 From Zltools.Zlautojobs) And a.Schema_User Not In (" & mstrOracleUser & ")" & vbNewLine & _
                IIf(chkShow(T_后台作业).value = 0, " And a.Broken = 'N' ", "") & vbNewLine & _
                "Order By 编号, 类型"
        Case T_其他
            mstrUsers = GetUsers
            Set rsTemp = CopyNewRec(Nothing, True, , _
                        Array("分类", adVarChar, 20, Empty, "对象", adVarChar, 100, Empty, _
                              "当前值", adVarChar, 50, Empty, "建议值", adVarChar, 50, Empty, _
                              "影响说明", adVarChar, 100, Empty, "处理说明", adVarChar, 100, Empty, _
                              "停用SQL", adVarChar, 200, Empty))
            Call CheckSysPara(rsTemp)
            Call CheckDBFile(rsTemp)
            Call CheckTriggers(rsTemp)
            Call CheckPrivs(rsTemp)
            
            Set GetData = rsTemp
    End Select
    If intIndex <> T_其他 Then Set GetData = gclsBase.OpenSQLRecord(mcnExe, gstrSQL, Decode(intIndex, T_客户端, "获取客户端", T_用户账号, "获取用户账号", T_后台作业, "获取自动作业"))
    Exit Function
ErrH:
    MsgBox err.Description, vbExclamation, Me.Caption
End Function

'******************************************************************************************************************
'功能：检查数据库参数
'******************************************************************************************************************
Private Sub CheckSysPara(ByRef rsData As ADODB.Recordset)
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrH
    gstrSQL = "Select Name , Value From V$parameter Where Name =[1] And Value =[2]"
    Set rsTemp = gclsBase.OpenSQLRecord(mcnExe, gstrSQL, SQL_CAPTION, "optimizer_index_cost_adj", "100")
    If Not rsTemp.EOF Then rsData.AddNew Array("分类", "对象", "当前值", "建议值", "影响说明", "处理说明", "停用SQL"), _
        Array("数据库参数", rsTemp!name, rsTemp!value, "20", "缺省值100会导致产品性能问题", "", "alter system set " & rsTemp!name & "=20")
    
    Set rsTemp = gclsBase.OpenSQLRecord(mcnExe, gstrSQL, SQL_CAPTION, "optimizer_index_caching", "0")
    If Not rsTemp.EOF Then rsData.AddNew Array("分类", "对象", "当前值", "建议值", "影响说明", "处理说明", "停用SQL"), _
        Array("数据库参数", rsTemp!name, rsTemp!value, "80", "缺省0会导致产品性能问题", "", "alter system set " & rsTemp!name & "=80")
        
    Set rsTemp = gclsBase.OpenSQLRecord(mcnExe, gstrSQL, SQL_CAPTION, "O7_DICTIONARY_ACCESSIBILITY", "FALSE")
    If Not rsTemp.EOF Then rsData.AddNew Array("分类", "对象", "当前值", "建议值", "影响说明", "处理说明", "停用SQL"), _
        Array("数据库参数", rsTemp!name, rsTemp!value, "TRUE", "导致系统视图无法授权，影响升级以及产品功能", "请手工调整为TRUE后重启数据库", "")
        
    gstrSQL = "Select Name , Value From V$parameter Where Name = [1] And Zl_To_Number(Value) < [2]"
    Set rsTemp = gclsBase.OpenSQLRecord(mcnExe, gstrSQL, SQL_CAPTION, "log_buffer", "209715200")
    If Not rsTemp.EOF Then rsData.AddNew Array("分类", "对象", "当前值", "建议值", "影响说明", "处理说明", "停用SQL"), _
        Array("数据库参数", rsTemp!name, Int(Val(rsTemp!value & "") / 1024 / 1024) & "M", ">=200M", "影响系统升级中数据修正效率", "请手工调整后重启数据库", "")
    
    Set rsTemp = gclsBase.OpenSQLRecord(mcnExe, gstrSQL, SQL_CAPTION, "parallel_execution_message_size", "8192")
    If Not rsTemp.EOF Then rsData.AddNew Array("分类", "对象", "当前值", "建议值", "影响说明", "处理说明", "停用SQL"), _
        Array("数据库参数", rsTemp!name, rsTemp!value, ">=8192", "影响系统升级并行执行", "请手工调整为8192后重启数据库", "")
    Exit Sub
ErrH:
    MsgBox err.Description, vbExclamation, Me.Caption
End Sub
'******************************************************************************************************************
'功能：检查日志文件
'******************************************************************************************************************
Private Sub CheckDBFile(ByRef rsData As ADODB.Recordset)
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrH
    gstrSQL = "Select 'INST_ID:' || a.Inst_Id || ',GROUP:' || a.Group# Name,b.Member," & vbNewLine & _
            "       a.Bytes Value" & vbNewLine & _
            "From Gv$log A" & vbNewLine & _
            "Join Gv$logfile B" & vbNewLine & _
            "On (a.Group# = b.Group# And a.Inst_Id = b.Inst_Id)" & vbNewLine & _
            "Where a.Bytes < 104857600" & vbNewLine & _
            "Order By a.Inst_Id, a.Group#, a.Thread#, b.Member"
    Set rsTemp = gclsBase.OpenSQLRecord(mcnExe, gstrSQL, SQL_CAPTION)
    Do While Not rsTemp.EOF
        rsData.AddNew Array("分类", "对象", "当前值", "建议值", "影响说明", "处理说明", "停用SQL"), _
            Array("数据库文件", rsTemp!name & "," & GetFileNameByPath(rsTemp!Member & ""), Int(Val(rsTemp!value & "") / 1024 / 1024) & "M", ">=100M", "影响系统升级中数据修正效率", "请手工调整为至少100M", "")
        rsTemp.MoveNext
    Loop
    Exit Sub
ErrH:
    MsgBox err.Description, vbExclamation, Me.Caption
End Sub
'******************************************************************************************************************
'功能：检查触发器
'******************************************************************************************************************
Private Sub CheckTriggers(ByRef rsData As ADODB.Recordset)
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrH
    'ZLHIS所有的触发器不用进一步判断对象
    gstrSQL = "Select a.Owner, a.Trigger_Name,a.Status From Dba_Triggers A Where a.Status = 'ENABLED' And a.Table_Owner In (" & mstrUsers & ") And a.Trigger_Type <> 'INSTEAD OF'"
    Set rsTemp = gclsBase.OpenSQLRecord(mcnExe, gstrSQL, SQL_CAPTION)
    Do While Not rsTemp.EOF
        rsData.AddNew Array("分类", "对象", "当前值", "建议值", "影响说明", "处理说明", "停用SQL"), _
            Array("触发器", rsTemp!Owner & "." & rsTemp!trigger_name, "ENABLED", "DISABLED", "影响该表的数据修正脚本执行效率", "升级期间禁用", _
            "alter trigger " & rsTemp!Owner & "." & rsTemp!trigger_name & " disable")
        rsTemp.MoveNext
    Loop
    Exit Sub
ErrH:
    MsgBox err.Description, vbExclamation, Me.Caption
End Sub
'******************************************************************************************************************
'功能：检查特殊用户的对象权限
'******************************************************************************************************************
Private Sub CheckPrivs(ByRef rsData As ADODB.Recordset)
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrH
    '升级是使用该表进行数据修正
    '检查ZLTOOLS与PUBLIC权限
    gstrSQL = "Select a.Grantee, a.Owner, a.Table_Name, a.Privilege" & vbNewLine & _
            "From (Select 'ZLTOOLS' Grantee, 'SYS' Owner, 'DBA_ROLE_PRIVS' Table_Name, 'SELECT' Privilege From Dual) A" & vbNewLine & _
            "Where Not Exists (Select 1" & vbNewLine & _
            "       From Dba_Tab_Privs C" & vbNewLine & _
            "       Where c.Owner = 'SYS' And (c.Grantee = 'PUBLIC' Or a.Grantee<>'PUBLIC' And c.Grantee = 'ZLTOOLS') And" & vbNewLine & _
            "             c.Table_Name = a.Table_Name And c.Privilege = a.Privilege)"
    Set rsTemp = gclsBase.OpenSQLRecord(mcnExe, gstrSQL, SQL_CAPTION)
    Do While Not rsTemp.EOF
        rsData.AddNew Array("分类", "对象", "当前值", "建议值", "影响说明", "处理说明", "停用SQL"), _
            Array("特殊用户对象权限", rsTemp!Grantee & " " & rsTemp!Privilege & " On " & rsTemp!Owner & "." & rsTemp!Table_Name, rsTemp!Privilege, "", "升级可能会出现异常，以及影响产品使用", "", "Grant " & rsTemp!Privilege & " On " & rsTemp!Owner & "." & rsTemp!Table_Name)
        rsTemp.MoveNext
    Loop
    Exit Sub
ErrH:
    MsgBox err.Description, vbExclamation, Me.Caption
End Sub

Private Function GetFileNameByPath(ByVal strFilePath As String) As String
    Dim lngPos As Long
    
    lngPos = InStrRev(strFilePath, "/")
    If lngPos = 0 Then
        lngPos = InStrRev(strFilePath, "\")

    End If
    If lngPos = 0 Then
        GetFileNameByPath = strFilePath
    Else
        GetFileNameByPath = Mid(strFilePath, lngPos + 1)
    End If
End Function

Private Function GetUsers() As String
'功能：获取勾选系统的所有者
    Dim strTemp As String, strUser As String
    Dim rsTmp As ADODB.Recordset
    
    On Error Resume Next
    gstrSQL = ""
    strTemp = "," & mstrSysNum & ","
    If InStr(strTemp, ",0,") > 0 Then
        gstrSQL = "Select 'ZLTOOLS' 所有者 From Dual"
        strTemp = Replace(strTemp, ",0,", "")
        strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
    Else
        strTemp = mstrSysNum
    End If
    If strTemp <> "" Then
        gstrSQL = IIf(gstrSQL = "", "", gstrSQL & " Union ") & "Select distinct 所有者 From Zlsystems Where 编号 In (" & strTemp & ")" & vbNewLine & _
                "Union" & vbNewLine & _
                "Select 所有者 From Zlbakspaces Where 系统 In (" & strTemp & ")"
    End If
    If gstrSQL = "" Then Exit Function
    Set rsTmp = gclsBase.OpenSQLRecord(mcnExe, gstrSQL, SQL_CAPTION)
    Do While Not rsTmp.EOF
        strUser = IIf(strUser = "", "", strUser & ",") & "'" & rsTmp!所有者 & "'"
        rsTmp.MoveNext
    Loop
    GetUsers = strUser
End Function

Private Function GetkillProcess() As String
'获取产品中的会话
    Dim strKillProcess As String
    Dim rsTemp As ADODB.Recordset
    
    On Error Resume Next
    gstrSQL = "Select Count(1) 计数 From Zltools.Zlkillprocess Where Rownum < 2"
    Set rsTemp = gclsBase.OpenSQLRecord(mcnExe, gstrSQL, "zlkillprocess数据判断")
    If err.Number <> 0 Then
        strKillProcess = ""
    Else
        If rsTemp!计数 <> 0 Then
            strKillProcess = "zlkillprocess"
        End If
    End If
    If strKillProcess <> "" Then
        strKillProcess = "Select Upper(名称) From Zltools.Zlkillprocess Union All" & vbNewLine & _
                        "Select 'VB6.EXE' From Zltools.Zlkillprocess"
    Else
        strKillProcess = "'ZL9LABPRINTSVR.EXE','ZL9LABRECEIV.EXE','ZL9LABTCPSVR.EXE','ZL9LISCOMM.EXE'," & _
                        "'ZL9WIZARDMAIN.EXE','ZLACTMAIN.EXE','ZLHIS+.EXE','ZLHISCRUST.EXE','ZLLISRECEIVESEND.EXE'," & _
                        "'ZLNEWQUERY.EXE','ZLORCLCONFIG.EXE','ZLPACSBROWSERSTATION.EXE','ZLPACSSRV.EXE'," & _
                        "'ZLPEISAUTOANALYSE.EXE','ZLRPTSQLADJUST.EXE','ZLRUNAS.EXE','ZLSVRNOTICE.EXE'," & _
                        "'ZLSVRSTUDIO.EXE','ZLWIZARDSTART.EXE','VB6.EXE'"
    End If
    GetkillProcess = strKillProcess
End Function

Public Sub ShowMe(ByVal strSys As String, ByVal cnExe As ADODB.Connection)
    mstrSysNum = strSys
    Set mcnExe = cnExe
    Me.Show 1
End Sub

Private Sub vsfShow_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim i As Long

    With vsfShow(Index)
        If Col = 0 Then
            If Row = 0 Then
                If .Cell(flexcpChecked, 0, 0) = flexChecked Then
                    .Cell(flexcpChecked, 0, 0) = flexChecked
                    For i = 1 To .Rows - 1
                        If .Cell(flexcpChecked, i, 0) = flexUnchecked Then
                            .Cell(flexcpChecked, i, 0) = flexChecked
                        End If
                    Next
                Else
                    .Cell(flexcpChecked, 0, 0) = flexUnchecked
                    For i = 1 To .Rows - 1
                        If .Cell(flexcpChecked, i, 0) = flexChecked Then
                            .Cell(flexcpChecked, i, 0) = flexUnchecked
                        End If
                    Next
                End If
            Else
                If .Cell(flexcpChecked, 0, 0) = flexChecked Then
                    .Cell(flexcpChecked, 0, 0) = flexUnchecked
                End If
                For i = 1 To .Rows - 1
                    If .Cell(flexcpChecked, i, 0) = flexUnchecked Then
                        Exit For
                    Else
                        If i = .Rows - 1 Then
                            .Cell(flexcpChecked, 0, 0) = flexChecked
                        End If
                    End If
                Next
            End If
        End If
    End With
End Sub

Private Sub vsfShow_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfShow(Index)
        If (Col = 0 And .TextMatrix(Row, 0) = " ") Or Col <> 0 Then Cancel = True
    End With
End Sub

Private Sub vsfShow_BeforeSort(Index As Integer, ByVal Col As Long, Order As Integer)
    Dim strFild As String
    
    If Col = 0 Then Order = 0
    If Index = T_用户账号 And Col = vsfShow(T_用户账号).ColIndex("部门") Then
        Order = 0
        strFild = "e.名称"
        If frmColChoose.ShowMe(vsfShow(T_用户账号), strFild, mstrCondition) Then
            Call LoadUser
        End If
    End If
End Sub

Private Sub vsfShow_BeforeUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub
