VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTransfusionLog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "操作日志"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10605
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTransfusionLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VSFlex8Ctl.VSFlexGrid vfgLog 
      Height          =   5175
      Left            =   105
      TabIndex        =   8
      Top             =   1260
      Width           =   8940
      _cx             =   15769
      _cy             =   9128
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
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
   Begin VB.Frame fraWhere 
      Height          =   870
      Left            =   135
      TabIndex        =   0
      Top             =   450
      Width           =   9000
      Begin VB.TextBox txtOper 
         Height          =   285
         Left            =   6930
         TabIndex        =   6
         Top             =   345
         Width           =   1785
      End
      Begin VB.TextBox txtNo 
         Height          =   285
         Left            =   4245
         TabIndex        =   4
         Top             =   345
         Width           =   1785
      End
      Begin MSComCtl2.DTPicker dtpS 
         Height          =   315
         Left            =   630
         TabIndex        =   1
         Top             =   345
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   233439235
         CurrentDate     =   41169
      End
      Begin MSComCtl2.DTPicker dtpE 
         Height          =   315
         Left            =   2175
         TabIndex        =   2
         Top             =   345
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   233439235
         CurrentDate     =   41169
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "操作员"
         Height          =   195
         Index           =   2
         Left            =   6270
         TabIndex        =   7
         Top             =   405
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "挂号单"
         Height          =   195
         Index           =   1
         Left            =   3585
         TabIndex        =   5
         Top             =   405
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "日期"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Top             =   405
         Width           =   360
      End
   End
   Begin MSComctlLib.StatusBar sbarSub 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   6465
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18653
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   90
      Top             =   105
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmTransfusionLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'－－操作日志查看
Private mlngDeptID As Long
Private mstrNo As String
Private mstrPrivs As String
Public Sub ShowMe(ByVal lngDeptID As Long, Optional strNO As String)
    mlngDeptID = lngDeptID
    mstrNo = strNO
    mstrPrivs = gstrPrivs
    Me.Show vbModal
End Sub

Private Sub RefData()
    '刷新数据
    Dim strSQL As String, strErr As String, dateS As Date, dateE As Date, strWhere As String
    Dim rsTmp As ADODB.Recordset
    
    dateS = Format(dtpS.Value, "yyyy-MM-dd")
    dateE = Format(dtpE.Value, "yyyy-MM-dd 23:59:59")
    
    strSQL = "select a.ID,a.类别,To_char(a.时间,'yy-MM-dd HH24:MI:SS') as 时间,a.挂号单,c.姓名 操作员,a.内容 " & _
             "from 门诊输液操作日志 a, 上机人员表 b, 人员表 c " & _
             "Where a.操作员=b.用户名 and b.人员id=c.id and a.科室ID=[1] and a.时间 Between [2] And [3] " & _
             "Order by a.时间 Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptID, dateS, dateE)
    Call vfgSetting(0, Me.vfgLog)
    
    If Not rsTmp.EOF Then
        strWhere = ""
        If txtNo <> "" Then
            strWhere = strWhere & " 挂号单 Like '" & txtNo & "*'"
        End If
        If txtOper <> "" Then
            If strWhere <> "" Then strWhere = strWhere & " And "
            strWhere = strWhere & " 操作员 Like '" & UCase(txtOper) & "*'"
        End If
        
        If strWhere <> "" Then
            rsTmp.Filter = strWhere
        End If
    End If
    
    If vfgLoadFromRecord(Me.vfgLog, rsTmp, strErr) Then
        With vfgLog
            .ColWidth(.ColIndex("时间")) = 1600: .ColHidden(.ColIndex("时间")) = False
            .ColWidth(.ColIndex("挂号单")) = 900: .ColHidden(.ColIndex("挂号单")) = False
            .ColWidth(.ColIndex("操作员")) = 700: .ColHidden(.ColIndex("操作员")) = False
            .ColWidth(.ColIndex("内容")) = 1200: .ColHidden(.ColIndex("内容")) = False
            
            .AutoResize = True
            .AutoSizeMode = flexAutoSizeColWidth
            .AutoSize .ColIndex("内容")
            
            .ExtendLastCol = True
            
        End With
    Else
        MsgBox strErr, vbQuestion, Me.Caption
        Exit Sub
    End If

End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim iRow As Integer, strSQL As String
    Select Case Control.ID
    
        Case conMenu_View_Refresh
            Call RefData
        Case conMenu_Edit_Delete
            If MsgBox("是否删除当前查询出的日志？", vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
                For iRow = vfgLog.FixedRows To vfgLog.Rows - 1
                    strSQL = "Zl_门诊输液操作日志_Delete(" & vfgLog.TextMatrix(iRow, vfgLog.ColIndex("ID")) & ")"
                    zlDatabase.ExecuteProcedure strSQL, Me.Caption
                Next
                Call RefData
            End If
        Case conMenu_File_Exit
            Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    fraWhere.Left = lngLeft + 15
    fraWhere.Top = lngTop
    fraWhere.Width = lngRight - lngLeft
    
    vfgLog.Left = lngLeft
    vfgLog.Top = fraWhere.Top + fraWhere.Height
    vfgLog.Width = lngRight - lngLeft
    vfgLog.Height = lngBottom - lngTop - sbarSub.Height - fraWhere.Height
    
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_Edit_Delete
        Control.Enabled = InStr(mstrPrivs, ";删除操作日志;") > 0
    End Select
End Sub

Private Sub Form_Load()

    Dim Menus As New Collection
    Menus.Add conMenu_View_Refresh & ",查询(&F),False"
    Menus.Add conMenu_Edit_Delete & ",删除(&D),False"
    
    Menus.Add conMenu_File_Exit & ",退出(&Q),True"
    
    Call CbsButtonInit(cbsMain, Menus, False, xtpBarTop)
    
    Set Menus = Nothing
    
    Call vfgSetting(0, Me.vfgLog, "")
    
    Me.dtpS = Format(Now, "yyyy-MM-dd")
    Me.dtpE = Format(Now, "yyyy-MM-dd")
    
    If mstrNo <> "" Then txtNo = mstrNo
    Call RefData
End Sub

Private Sub Form_Resize()
    Call cbsMain_Resize
End Sub
