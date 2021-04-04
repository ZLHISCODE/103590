VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmDoctorShift 
   Caption         =   "医生交接班管理"
   ClientHeight    =   11475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17685
   Icon            =   "frmDoctorShift.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11475
   ScaleWidth      =   17685
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picDataIn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8895
      Left            =   0
      ScaleHeight     =   8895
      ScaleWidth      =   6015
      TabIndex        =   11
      Top             =   480
      Width           =   6015
      Begin XtremeReportControl.ReportControl rptData 
         Height          =   2985
         Left            =   120
         TabIndex        =   12
         Top             =   4200
         Width           =   5805
         _Version        =   589884
         _ExtentX        =   10239
         _ExtentY        =   5265
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.PictureBox picNumBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1305
         ScaleWidth      =   5745
         TabIndex        =   34
         Top             =   7560
         Width           =   5775
         Begin VB.PictureBox picNumDown 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5520
            Picture         =   "frmDoctorShift.frx":6852
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   37
            Top             =   1080
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.PictureBox picNumUp 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5520
            Picture         =   "frmDoctorShift.frx":7254
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   36
            Top             =   0
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.PictureBox picNum 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1095
            Left            =   0
            ScaleHeight     =   1095
            ScaleWidth      =   5490
            TabIndex        =   35
            Top             =   0
            Width           =   5490
            Begin VB.Label lblTypeNum 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "病人数量"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   38
               Top             =   120
               Visible         =   0   'False
               Width           =   720
            End
         End
         Begin VB.Line lineNumY 
            Visible         =   0   'False
            X1              =   5520
            X2              =   5520
            Y1              =   0
            Y2              =   1320
         End
      End
      Begin VB.Frame fraFilter 
         Caption         =   "查询条件"
         Height          =   3375
         Left            =   120
         TabIndex        =   14
         Top             =   0
         Width           =   5820
         Begin VB.CommandButton cmdRef 
            Caption         =   "刷新(&R)"
            Height          =   350
            Left            =   960
            TabIndex        =   24
            Top             =   2400
            Width           =   855
         End
         Begin VB.ComboBox cboTime 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtSubject 
            Height          =   300
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton cmdSubject 
            Caption         =   "…"
            Height          =   290
            Left            =   3000
            TabIndex        =   21
            Top             =   240
            Width           =   255
         End
         Begin VB.ComboBox cboDept 
            Height          =   300
            Left            =   3840
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   240
            Width           =   1815
         End
         Begin VB.PictureBox picBack 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1095
            Left            =   960
            ScaleHeight     =   1065
            ScaleWidth      =   4665
            TabIndex        =   15
            Top             =   1200
            Width           =   4695
            Begin VB.PictureBox picShift 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   975
               Left            =   0
               ScaleHeight     =   975
               ScaleWidth      =   4425
               TabIndex        =   18
               Top             =   0
               Width           =   4425
               Begin VB.CheckBox chkType 
                  BackColor       =   &H80000005&
                  Caption         =   "值班班次日"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   19
                  Top             =   120
                  Width           =   1335
               End
            End
            Begin VB.PictureBox picUp 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000E&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4440
               Picture         =   "frmDoctorShift.frx":7C56
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   17
               Top             =   0
               Width           =   255
            End
            Begin VB.PictureBox picDown 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4440
               Picture         =   "frmDoctorShift.frx":8658
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   16
               Top             =   840
               Width           =   255
            End
            Begin VB.Line lineY 
               X1              =   4440
               X2              =   4440
               Y1              =   0
               Y2              =   1200
            End
         End
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   285
            Left            =   1920
            TabIndex        =   25
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            CalendarTitleBackColor=   -2147483638
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   209584131
            CurrentDate     =   42675
            MaxDate         =   402133
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   285
            Left            =   3840
            TabIndex        =   26
            Top             =   720
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            CalendarTitleBackColor=   -2147483638
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   209584131
            CurrentDate     =   42702
            MaxDate         =   402133
         End
         Begin VB.Label lblSubject 
            AutoSize        =   -1  'True
            Caption         =   "学    科"
            Height          =   180
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lblDept 
            AutoSize        =   -1  'True
            Caption         =   "科 室"
            Height          =   180
            Left            =   3360
            TabIndex        =   30
            Top             =   315
            Width           =   450
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            Caption         =   "时    间"
            Height          =   180
            Left            =   120
            TabIndex        =   29
            Top             =   795
            Width           =   720
         End
         Begin VB.Label lblSplit1 
            AutoSize        =   -1  'True
            Caption         =   "~"
            Height          =   180
            Left            =   3480
            TabIndex        =   28
            Top             =   840
            Width           =   90
         End
         Begin VB.Label lblType 
            AutoSize        =   -1  'True
            Caption         =   "交班班次"
            Height          =   180
            Left            =   120
            TabIndex        =   27
            Top             =   1200
            Width           =   720
         End
      End
      Begin MSComctlLib.TreeView tvwSubject 
         Height          =   1935
         Left            =   4440
         TabIndex        =   13
         Top             =   3360
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   3413
         _Version        =   393217
         Indentation     =   353
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imgList"
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label lblAllNum 
         AutoSize        =   -1  'True
         Caption         =   "病人汇总情况"
         Height          =   180
         Left            =   120
         TabIndex        =   33
         Top             =   7320
         Width           =   1080
      End
      Begin VB.Label lblRecord 
         AutoSize        =   -1  'True
         Caption         =   "交接班记录"
         Height          =   180
         Left            =   120
         TabIndex        =   32
         Top             =   3720
         Width           =   900
      End
   End
   Begin VB.PictureBox picSub 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   7800
      ScaleHeight     =   4665
      ScaleWidth      =   8385
      TabIndex        =   8
      Top             =   5280
      Width           =   8415
      Begin VB.PictureBox picShow 
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   480
         ScaleHeight     =   3135
         ScaleWidth      =   6735
         TabIndex        =   9
         Top             =   600
         Width           =   6735
         Begin XtremeSuiteControls.TabControl tbcSub 
            Height          =   2580
            Left            =   240
            TabIndex        =   10
            Top             =   120
            Width           =   3690
            _Version        =   589884
            _ExtentX        =   6509
            _ExtentY        =   4551
            _StockProps     =   64
         End
      End
   End
   Begin VB.PictureBox picPatient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   8040
      ScaleHeight     =   3225
      ScaleWidth      =   5625
      TabIndex        =   3
      Top             =   720
      Width           =   5655
      Begin VSFlex8Ctl.VSFlexGrid vsDetail 
         Height          =   1575
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   8850
         _cx             =   15610
         _cy             =   2778
         Appearance      =   2
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   600
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmDoctorShift.frx":905A
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
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
      Begin VB.Label lblList 
         AutoSize        =   -1  'True
         Caption         =   "交接病人清单"
         Height          =   180
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label lblColor 
         AutoSize        =   -1  'True
         Caption         =   "黄色"
         ForeColor       =   &H0080C0FF&
         Height          =   180
         Left            =   3360
         TabIndex        =   6
         Top             =   240
         Width           =   360
      End
      Begin VB.Label lblColor2 
         AutoSize        =   -1  'True
         Caption         =   "表示未生成交班描述"
         Height          =   180
         Left            =   3840
         TabIndex        =   5
         Top             =   240
         Width           =   1620
      End
   End
   Begin VB.PictureBox picSplitY 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   8280
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   5295
      TabIndex        =   2
      Top             =   4680
      Width           =   5295
   End
   Begin VB.PictureBox picSplitX 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   7320
      MousePointer    =   9  'Size W E
      ScaleHeight     =   6615
      ScaleWidth      =   45
      TabIndex        =   1
      Top             =   480
      Width           =   45
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   11115
      Width           =   17685
      _ExtentX        =   31194
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDoctorShift.frx":9240
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   28284
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   7440
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorShift.frx":9AD4
            Key             =   "unCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorShift.frx":A06E
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorShift.frx":A608
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorShift.frx":10E6A
            Key             =   "add"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorShift.frx":176CC
            Key             =   "Person"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorShift.frx":180DE
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorShift.frx":1E940
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorShift.frx":1F352
            Key             =   "Down"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.ImageManager imgPublic 
      Left            =   2640
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmDoctorShift.frx":1FD64
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   3600
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmDoctorShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum ColData
    cold_记录id = 0
    cold_科室
    cold_科室ID
    cold_交班医生
    cold_交班班次
    cold_交班时期 '交班开始时间|交班结束时间
    cold_接班医生
    cold_接班班次
    cold_接班时期 '接班开始时间|交班结束时间
    cold_交班状态
    cold_接班状态
    cold_完成时间
    cold_审阅人
    cold_审阅时间
    cold_审阅说明
    cold_交班期间
End Enum
Private mobjESign As Object '电子签名接口部件
Private mstrPriv As String
Private mintCA As Integer
Private mlngRow As Long
Private mrsPati As ADODB.Recordset '记录选择时，该记录id下的所有病人
Private mobjFrom As frmShiftEdit
Private mblnEdit As Boolean
Private mobjMenu As CommandBarPopup
Private mblnLoading As Boolean
Private mlngDeptID As Long '登录人员所属临床科室的ID,如没有，则根据界面的科室选择来传值，为所有科室时，默认是科室的第一个
Private mstrDeptId As String '登录人员有权限操作的科室id,
Private mblnClick As Boolean '是否触发cboDept的Click事件，true 触发；false-不触发

Private Sub cboDept_Click()
        
    If cboDept.Text <> "所有科室" Then
        mlngDeptID = cboDept.ItemData(cboDept.ListIndex)
    End If
    If mblnClick Then
        zlCommFun.PressKey vbKeyReturn
    End If
End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call LoadType
    End If
End Sub

Private Sub cboTime_Click()
    Dim dtToday As Date
    Dim intDay As Integer
        
    dtToday = zlDatabase.Currentdate
    Select Case cboTime.Text
        Case "当天"
            dtpBegin.Value = Format(Date, "yyyy-MM-dd")
            dtpEnd.Value = Format(Date, "yyyy-MM-dd")
        Case "昨天"
            dtpBegin.Value = Format(Date - 1, "yyyy-MM-dd")
            dtpEnd.Value = Format(Date - 1, "yyyy-MM-dd")
        Case "本周"
            dtToday = Format(Date, "yyyy-MM-dd")
            intDay = Weekday(CDate(Format(Date, "yyyy-MM-dd")))
            intDay = IIf(intDay = 1, 7, intDay - 1)
            dtpBegin.Value = Format(DateAdd("d", 0 - intDay + 1, dtToday), "yyyy-MM-dd") & " 00:00:00"
            dtpEnd.Value = Format(DateAdd("d", 7 - intDay, dtToday), "yyyy-MM-dd") & " 23:59:59"
        Case "本月"
            dtpBegin.Value = Format(dtToday, "yyyy-MM") & "-01 00:00:00"
            dtpEnd.Value = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(dtToday, "yyyy-MM") & "-01"))), "yyyy-MM-dd") & " 23:59:59"
        Case "自定义"
            
    End Select
    If cboTime.Text = "自定义" Then
        dtpBegin.Enabled = True
        dtpEnd.Enabled = True
    Else
        dtpBegin.Enabled = False
        dtpEnd.Enabled = False
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim i As Long, j As Long, lngId As Long
    Dim strDept As String, strOutTime As String, strInTime As String
    Dim strOutPer As String, strInPer As String, strTemp As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errH
    If rptData.SelectedRows.Count > 0 Then
        If rptData.SelectedRows(0).GroupRow = False Then
            lngId = rptData.SelectedRows(0).Record(cold_记录id).Value
            strOutPer = rptData.SelectedRows(0).Record(cold_交班医生).Value
            strInPer = rptData.SelectedRows(0).Record(cold_接班医生).Value
        End If
    End If
    Select Case Control.id
    Case conMenu_File_TypeManage '班次管理
        If frmShiftMange.ShowMe(mstrDeptId, mlngDeptID) Then
            Call LoadType
        End If
    Case conMenu_File_Preview '预览
        ReportMode 1
    Case conMenu_File_Print '打印
        ReportMode 2
    Case conMenu_File_Excel '输出到Excel
        ReportMode 3
    Case conMenu_Edit_NewItem '新增
        If cboDept.List(0) = "所有科室" Then j = 1
        For i = j To cboDept.ListCount - 1
            strDept = IIf(strDept = "", "", strDept & "|") & cboDept.List(i) & "," & cboDept.ItemData(i)
        Next
        strDept = cboDept.ListIndex - j & "|" & strDept
        If frmEdit.ShowMe(0, 0, strDept, grsUserInfo!姓名) Then Call RefreshRecord
    Case conMenu_Edit_Modify '修改
        strDept = rptData.SelectedRows(0).Record(cold_科室).Value & "|" & rptData.SelectedRows(0).Record(cold_科室ID).Value
        strOutTime = rptData.SelectedRows(0).Record(cold_交班班次).Value & "|" & rptData.SelectedRows(0).Record(cold_交班时期).Value
        strInTime = rptData.SelectedRows(0).Record(cold_接班班次).Value & "|" & rptData.SelectedRows(0).Record(cold_接班时期).Value
        If frmEdit.ShowMe(1, lngId, strDept, strOutPer, strOutTime, strInPer, strInTime) Then Call RefreshRecord
    Case conMenu_Edit_Delete '删除
        If MsgBox("您确认删除这条值班记录吗？", vbInformation + vbDefaultButton2 + vbYesNo) = vbNo Then Exit Sub
        gstrSQL = "Zl_医生交接班记录_State(0," & lngId & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "删除值班记录")
        Call RefreshRecord
    Case conMenu_Edit_FinOut '完成交班
        Set rsTemp = GetUserInfo(strOutPer)
        If strOutPer = "管理员" Then
            strTemp = "ZLHIS"
        Else
            strTemp = rsTemp!用户名
        End If
        Call ShiftMange(0, lngId, strTemp, strOutPer, rsTemp!部门ID)
    Case conMenu_Edit_FinIn '完成接班
        Set rsTemp = GetUserInfo(strInPer)
        Call ShiftMange(1, lngId, rsTemp!用户名, strInPer, rsTemp!部门ID)
    Case conMenu_Edit_FinRead '完成审阅
        strTemp = frmReview.ShowMe
        If strTemp = "取消JM" Then Exit Sub
        gstrSQL = "Zl_医生交接班记录_State(3," & lngId & ",'" & grsUserInfo!姓名 & "','" & strTemp & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "完成审阅")
        Call RefreshRecord
    Case conMenu_Edit_CancelOut, conMenu_Edit_CancelIn, conMenu_Edit_CancelRead '取消完成交班,取消完成接班,取消完成审阅
        Call CancelOper(Control.id, lngId)
    Case conMenu_Edit_CheckOutSign '验证签名
        Call VerifySign(1, lngId)
    Case conMenu_Edit_CheckInSign
        Call VerifySign(2, lngId)
    Case conMenu_Report_Record '交接情况查询报表
        Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1242_2", Me)
    Case conMenu_View_ToolBar_Button '工具栏
        For i = 2 To cbsMain.Count
            Control.Checked = Not Control.Checked
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Call Form_Resize
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '按钮文字
        Control.Checked = Not Control.Checked
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size '大图标
        Control.Checked = Not Control.Checked
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '状态栏
        Control.Checked = Not Control.Checked
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Call Form_Resize
        Me.cbsMain.RecalcLayout
    Case conMenu_Help_Web_Home 'Web上的中联
        Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail '发送反馈
        Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About '关于
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_File_Exit '退出
        Unload Me
    End Select
    Exit Sub
errH:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub
Private Sub ReportMode(bytMode As Byte)
'bytMode-1预览；2打印；3输出到excel
    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1242_1", Me, "记录id=" & Val(rptData.SelectedRows(0).Record(cold_记录id).Value), bytMode)
End Sub

Private Sub VerifySign(bytType As Byte, ByVal lngId As Long)
'验证签名
'bytType：1-交班；2-接班
    Dim rsTemp As ADODB.Recordset
    Dim strSource As String
        
    On Error GoTo errH
    If lngId = 0 Then Exit Sub
    gstrSQL = "Select 证书id From 医生交接班签名 Where 记录id = [1] And 签名类型 =[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngId, bytType)
    If rsTemp.RecordCount = 0 Then
        MsgBox "签名数据已被删除，电子签名验证失败！", vbInformation, Me.Caption
        Exit Sub
    End If
    Call ReadSignSource(lngId, strSource)
    If mobjESign.VerifySignature(strSource, rsTemp!证书ID, 0) = True Then
        MsgBox "电子签名验证成功!", vbInformation, Me.Caption
    Else
        MsgBox "电子签名验证失败!", vbInformation, Me.Caption
    End If
    Exit Sub
errH:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub CancelOper(ByVal lngType As Long, ByVal lngId As Long)
'取消完成交班、取消完成接班、取消完成审阅
    Dim bytType As Byte
    
    On Error GoTo errH
    bytType = Decode(lngType, conMenu_Edit_CancelOut, 4, conMenu_Edit_CancelIn, 5, conMenu_Edit_CancelRead, 6)
    gstrSQL = "Zl_医生交接班记录_State(" & bytType & "," & lngId & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "取消完成")
    Call RefreshRecord
    Exit Sub
errH:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Function CheckPati() As Boolean
'交班前的病人信息检查，主要检查交班描述是否为空
    Dim i As Long

    With vsDetail
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("交班描述")) = "" Then
                MsgBox "存在交班描述为空的病人，无法完成交班，请检查！", vbInformation, Me.Caption
                Call .ShowCell(i, .ColIndex("交班描述"))
                Exit Function
            End If
        Next
    End With
    CheckPati = True
End Function

Private Sub ShiftMange(ByVal bytType As Byte, ByVal lngId As Long, ByVal strPer As String, ByVal strDoc As String, ByVal lngDeptID As Long)
'完成交班或接班
'bytType:0-交班；1-接班；strPer-用户名；strDoc-姓名
    Dim lng证书ID As Long, lngCA As Long
    Dim strSource As String, strTimeStamp As String, strTimeStampCode As String
    Dim strSign As String, strCaInfo As String
    Dim blnBegin As Boolean, blnIndetifi As Boolean '是否已经验证
    Dim rsTemp As ADODB.Recordset

    If bytType = 0 Then
        If Not CheckPati Then Exit Sub
    End If
    If strPer <> grsUserInfo!用户名 Then
        '交班或接班时如果当前用户不是对应的交班或接班医生用户，则需身份验证
        If Not frmUserIdentify.ShowMe(Me, "身份验证，请输入密码", glngSys, strPer, True) Then
'            MsgBox "身份验证未通过，无法完成" & IIf(bytType = 0, "交班！", "接班！"), vbInformation, Me.Caption
            Exit Sub
        Else
            blnIndetifi = True
        End If
    End If
    If GetCA(strDoc) Then
        On Error Resume Next
        If mobjESign Is Nothing Then
            Set mobjESign = CreateObject("zl9ESign.clsESign")
        End If
        Err.Clear: On Error GoTo errH
        If Not mobjESign Is Nothing Then
            Call mobjESign.Initialize(gcnOracle, glngSys)
        Else
            MsgBox "电子签名部件未能正确安装，审核操作不能继续。", vbInformation, Me.Caption
            Exit Sub
        End If
        Call ReadSignSource(lngId, strSource)
        strSign = mobjESign.Signature(strSource, strPer, lng证书ID, strTimeStamp, Nothing, strTimeStampCode)
        If strSign <> "" Then
            If strTimeStamp <> "" Then
                strTimeStamp = zlStr.To_Date(strTimeStamp)
            Else
                strTimeStamp = "NULL"
            End If
        Else
'            MsgBox "电子签名失败，无法完成" & IIf(bytType = 0, "交班！", "接班！"), vbInformation, Me.Caption
            Exit Sub
        End If
        strCaInfo = "," & lng证书ID & ",'" & strSign & "','" & strTimeStampCode & "'," & strTimeStamp
    Else
        If Not blnIndetifi Then
            If Not frmUserIdentify.ShowMe(Me, "身份验证，请输入密码", glngSys, strPer, True) Then
'                MsgBox "身份验证未通过，无法完成" & IIf(bytType = 0, "交班！", "接班！"), vbInformation, Me.Caption
                Exit Sub
            End If
        End If
    End If
    gcnOracle.BeginTrans: blnBegin = True
    gstrSQL = "Zl_医生交接班签名_Edit(" & lngId & "," & IIf(bytType = 0, 1, 2) & ",'" & strDoc & "'" & strCaInfo & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "完成交班电子签名")
    gstrSQL = "Zl_医生交接班记录_State(" & IIf(bytType = 0, 1, 2) & "," & lngId & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "完成交班")
    blnBegin = True
    gcnOracle.CommitTrans
    
    Call RefreshRecord
    Exit Sub
errH:
    If blnBegin Then gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngIn As Long, lngHold As Long
    Dim strReadPer As String
    
    If mblnEdit Then
        If Control.id = conMenu_Edit_NewItem Or Control.id = conMenu_Edit_Modify Or Control.id = conMenu_Edit_Delete Or Control.id = conMenu_Edit_FinOut Or Control.id = conMenu_Edit_FinIn Or Control.id = conMenu_Edit_FinRead _
            Or Control.id = conMenu_Edit_CancelOut Or Control.id = conMenu_Edit_CancelIn Or Control.id = conMenu_Edit_CancelRead Or Control.id = conMenu_Edit_CheckOutSign Or Control.id = conMenu_Edit_CheckInSign Or Control.id = conMenu_File_TypeManage _
            Or Control.id = conMenu_File_Preview Or Control.id = conMenu_File_Print Then
            Control.Enabled = False
        Else
            Control.Enabled = True
        End If
        Exit Sub
    End If
    
    lngIn = -1
    lngHold = -1
    strReadPer = "-1"
    If rptData.SelectedRows.Count > 0 Then
        If Not rptData.SelectedRows(0).GroupRow Then
            '表示选择行时数据行
            lngIn = Val(rptData.SelectedRows(0).Record(cold_交班状态).Value)
            lngHold = Val(rptData.SelectedRows(0).Record(cold_接班状态).Value)
            strReadPer = rptData.SelectedRows(0).Record(cold_审阅人).Value
        End If
    End If
    Select Case Control.id
        '权限设置
        Case conMenu_Report_Record
            Control.Enabled = CheckPriv("临床科室交接班情况查询")
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
            Control.Enabled = CheckPriv("交接班记录") And lngHold = 1
        Case conMenu_File_TypeManage
            Control.Enabled = CheckPriv("班次管理")
        Case conMenu_Edit_NewItem
            Control.Enabled = CheckPriv("医生交接班")
        Case conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_FinOut '修改,删除，完成交班
            Control.Enabled = lngIn = 0 And CheckPriv("医生交接班")
        Case conMenu_Edit_FinIn '完成接班
            Control.Enabled = False
            Control.Enabled = lngHold = 0 And CheckPriv("医生交接班") And lngIn = 1
        Case conMenu_Edit_FinRead '完成审阅
            Control.Enabled = strReadPer = "" And CheckPriv("交接班审阅") _
                And lngHold = 1 And lngIn = 1
        Case conMenu_Edit_CancelOut '取消完成交班
            Control.Enabled = lngIn = 1 And lngHold = 0 And CheckPriv("医生交接班")
        Case conMenu_Edit_CancelIn '取消完成接班
            Control.Enabled = lngHold = 1 And strReadPer = "" And CheckPriv("医生交接班")
        Case conMenu_Edit_CancelRead '取消完成审阅
            Control.Enabled = strReadPer <> "" And strReadPer <> "-1" And CheckPriv("交接班审阅")
        Case conMenu_Edit_CheckOutSign '交班医生电子签名验证
            Control.Enabled = strReadPer <> "" And strReadPer <> "-1" And CheckPriv("医生交接班")
        Case conMenu_Edit_CheckInSign '接班医生电子签名验证
            Control.Enabled = strReadPer <> "" And strReadPer <> "-1" And CheckPriv("医生交接班")
    End Select
End Sub

Private Sub cmdRef_Click()
    Call RefreshRecord
End Sub

Private Sub cmdSubject_Click()
    
    tvwSubject.Visible = True
    tvwSubject.SetFocus
End Sub

Private Sub Form_Activate()
    With vsDetail
        .AutoSizeMode = flexAutoSizeRowHeight
        .WordWrap = True
        .AutoSize .ColIndex("交班描述")
    End With
End Sub

Private Sub Form_Load()
    
    mblnLoading = True
    mstrPriv = gstrPrivs
    Set grsUserInfo = zlDatabase.GetUserInfo
    mintCA = IIf(GetCA(grsUserInfo!姓名), 1, 0)
    Call InitCommandBar
    Call InitReportColumn
    
    cboTime.AddItem "当天"
    cboTime.AddItem "昨天"
    cboTime.AddItem "本周"
    cboTime.AddItem "本月"
    cboTime.AddItem "自定义"
    cboTime.ListIndex = 0
        
    Call LoadShowData
    Call LoadType
    Call RefreshRecord
    picSplitX.BackColor = Me.BackColor
    picSplitY.BackColor = Me.BackColor
    picSplitX.Left = 5900
    picSplitY.Top = 4200

    Call GetFrom
    Call RestoreWinState(Me, App.ProductName)
    Call LoadVsfColWidth
    
    cmdSubject.Enabled = CheckPriv("所有学科")
    
    mblnLoading = False
    If mlngRow = 0 And rptData.Rows.Count > 1 Then
        Set rptData.FocusedRow = rptData.Rows(1)
    End If
    
End Sub

Private Sub LoadVsfColWidth()
'vsf表格上一次的列宽
    Dim strCols As String
    Dim varTemp As Variant, varData As Variant
    Dim i As Long
    
    strCols = GetSetting("ZLSOFT", "私有模块\" & gstrDbaUser & "\" & gstrProductName & "\医生交接班记录", "病人清单列宽")
    If strCols = "" Or InStr(strCols, "主诉") > 0 Then
        strCols = "序号|0;病人ID|0;主页ID|0;内容ID|0;上移|300;下移|300;类型|1500;姓名|900;性别|480;年龄|720;床号|495;标识号|1020;入院时间|990;入院方式|900"
    End If
    varTemp = Split(strCols, ";")
    With vsDetail
        For i = LBound(varTemp) To UBound(varTemp)
            varData = Split(varTemp(i), "|")
            If varData(0) = .ColKey(i) Then
                .ColWidth(i) = varData(1)
            End If
        Next
    End With
End Sub

Private Sub LoadShowData()
'加载界面学科和科室的数据
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    Dim strDept As String, strTemp As String
    Dim objNode As Object
    Dim varTemp As Variant
        
    On Error GoTo errH
    gstrSQL = "Select b.用户名, d.Id 科室id, d.名称 科室, f.编码 学科编码, f.名称 学科" & vbNewLine & _
        "From 上机人员表 b, 部门表 d, 部门人员 e, 临床性质 f, 临床部门 g" & vbNewLine & _
        "Where b.用户名 =[1] And e.人员id = b.人员id And e.部门id = d.Id And e.缺省 = 1 And d.Id = g.部门id And g.工作性质 = f.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, gstrDbaUser)
    If rsTemp.RecordCount = 1 Then
        txtSubject.Text = rsTemp!学科
        txtSubject.Tag = rsTemp!学科编码
        cboDept.Tag = rsTemp!科室ID
        strDept = rsTemp!科室
        mlngDeptID = rsTemp!科室ID
    Else
        txtSubject.Tag = "所有学科"
        txtSubject.Text = "所有学科"
        strDept = "所有科室"
    End If
    If cmdSubject.Enabled Then
        '只加载有临床科室的学科
        gstrSQL = "Select 名称, 编码 From 临床性质" & vbNewLine & _
            "Where 编码 In (Select Distinct a.工作性质 From 临床部门 a, 部门性质说明 c Where a.部门id = c.部门id" & vbNewLine & _
            "And c.工作性质='临床' And c.服务对象 In (2, 3)) Order By 序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        With tvwSubject
            .Left = txtSubject.Left + 40
            .Top = txtSubject.Top + txtSubject.Height + 110
            .Width = txtSubject.Width
            .Height = fraFilter.Height
            .Nodes.Clear
            Set objNode = .Nodes.Add(, , "K所有学科", "所有学科", "Dept")
            Do Until rsTemp.EOF
                varTemp = Split(rsTemp!编码, ".")
                For i = LBound(varTemp) To UBound(varTemp) - 1
                    strTemp = IIf(strTemp = "", "", strTemp & ".") & varTemp(i)
                Next
                On Error Resume Next
                Set objNode = .Nodes.Add("K" & strTemp & "", tvwChild, "K" & CStr(rsTemp!编码), rsTemp!名称, "Dept")
                If Err.Number <> 0 Then
                    Err.Clear: On Error GoTo errH
                    Set objNode = .Nodes.Add(, , "K" & CStr(rsTemp!编码), rsTemp!名称, "Dept")
                End If
                rsTemp.MoveNext
            Loop
            .ZOrder 0
        End With
    End If
    On Error GoTo errH
    Call LoadDept(txtSubject.Tag)
    
    mblnClick = False
    For i = 0 To cboDept.ListCount - 1
        If cboDept.List(i) = strDept Then
            cboDept.ListIndex = i
            Exit For
        Else
            If i = cboDept.ListCount - 1 Then
                mlngDeptID = cboDept.ItemData(0)
            End If
        End If
    Next
    mblnClick = True
    Exit Sub
errH:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub LoadDept(ByVal strSubjectCode As String)
'选择学科时对应的部门选择变化
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errH
    cboDept.Clear
    mstrDeptId = ""
    If strSubjectCode = "所有学科" Then
        gstrSQL = "Select distinct 编码 ||'-' || 名称 as 名称,Id,编码 From 部门表" & vbNewLine & _
            "Where Id In (Select Distinct a.部门id From 临床部门 a, 部门性质说明 c Where a.部门id = c.部门id" & vbNewLine & _
            "And c.服务对象 In (2, 3) And c.工作性质='临床') And (撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or 撤档时间 is NULL) Order By 编码"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    Else
        gstrSQL = "Select distinct b.编码 || '-' || b.名称 as 名称, b.Id,b.编码 From 临床部门 a, 部门表 b,部门性质说明 c" & vbNewLine & _
            "Where a.部门id = b.Id And a.部门id = c.部门id And c.服务对象 In (2, 3) And  a.工作性质 =[1]" & vbNewLine & _
            "And c.工作性质='临床'And (b.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or b.撤档时间 is NULL) Order By 编码"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strSubjectCode)
    If rsTemp.RecordCount > 1 Then
        cboDept.AddItem "所有科室"
    End If
    Call zlcontrol.CboAddData(cboDept, rsTemp, False)
    mblnClick = False
    If cboDept.ListCount > 0 Then cboDept.ListIndex = 0
    mblnClick = True
    Do While Not rsTemp.EOF
        mstrDeptId = mstrDeptId & "," & rsTemp!id
        rsTemp.MoveNext
    Loop
    mstrDeptId = Mid(mstrDeptId, 2)
    Exit Sub
errH:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub RefreshRecord()
'刷新记录数据
    Dim i As Long
    Dim strTemp As String
    Dim rsTemp As ADODB.Recordset
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
        
    vsDetail.Rows = 1
    rptData.Records.DeleteAll
    lblTypeNum(0).Visible = False
    lblAllNum.Caption = "病人汇总情况"
    For i = 1 To lblTypeNum.UBound
        Unload lblTypeNum(i)
    Next
    gstrSQL = ""
    If cboDept.Text = "所有科室" Then
        gstrSQL = " And b.id in(Select 部门id From 临床部门" & IIf(txtSubject.Text = "所有学科", ")", " Where 工作性质 =[4])")
    Else
        gstrSQL = " And b.id=[4]"
    End If
    strTemp = ""
    For i = chkType.LBound To chkType.UBound
        If chkType(i).Value = 1 Then
            strTemp = IIf(strTemp = "", "", strTemp & ",") & chkType(i).Caption
        End If
    Next
    gstrSQL = gstrSQL & " order by a.记录id"
    On Error GoTo errH
    gstrSQL = "Select a.记录id, a.科室ID,b.名称 科室, a.交班医生, a.交班班次," & vbNewLine & _
        "       To_Char(a.交班开始时间, 'MM-DD HH24:Mi') || '～' || To_Char(a.交班结束时间, 'MM-DD HH24:Mi') As 交班期间, a.交班开始时间, a.交班结束时间," & vbNewLine & _
        "       a.接班医生, a.接班班次, a.接班开始时间, a.接班结束时间, a.交班状态, a.接班状态, a.完成时间, a.审阅人, a.审阅时间, a.审阅说明" & vbNewLine & _
        "From 医生交接班记录 a, 部门表 b" & vbNewLine & _
        "Where a.科室id = b.Id And a.接班开始时间 >=to_date([1],'yyyy-mm-dd hh24:mi:ss') And a.接班开始时间<to_date([2],'yyyy-mm-dd hh24:mi:ss')" & _
        " And a.交班班次 in(Select * From Table(f_str2list([3]))) " & gstrSQL

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Format(dtpBegin.Value, "yyyy-mm-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-mm-dd 23:59:59"), _
                strTemp, IIf(cboDept.Text = "所有科室", Trim(txtSubject.Tag), cboDept.ItemData(cboDept.ListIndex)))
    Do While Not rsTemp.EOF
        Set objRecord = rptData.Records.Add
        Set objItem = objRecord.AddItem(Val(rsTemp!记录id))
        Set objItem = objRecord.AddItem(CStr(rsTemp!科室))
        Set objItem = objRecord.AddItem(CStr(rsTemp!科室ID))
        Set objItem = objRecord.AddItem(CStr(rsTemp!交班医生))
        Set objItem = objRecord.AddItem(CStr(rsTemp!交班班次))
        Set objItem = objRecord.AddItem(CStr(rsTemp!交班开始时间) & "|" & CStr(rsTemp!交班结束时间))
        Set objItem = objRecord.AddItem(CStr(rsTemp!接班医生))
        Set objItem = objRecord.AddItem(CStr(rsTemp!接班班次))
        Set objItem = objRecord.AddItem(CStr(rsTemp!接班开始时间) & "|" & CStr(rsTemp!接班结束时间))
        Set objItem = objRecord.AddItem(CStr(rsTemp!交班状态 & ""))
        Set objItem = objRecord.AddItem(CStr(rsTemp!接班状态 & ""))
        If IsNull(rsTemp!完成时间) Then
            Set objItem = objRecord.AddItem(" ")
        Else
            Set objItem = objRecord.AddItem(Format(rsTemp!完成时间, "yyyy-mm-dd") & "")
        End If
        Set objItem = objRecord.AddItem(CStr(rsTemp!审阅人 & ""))
        Set objItem = objRecord.AddItem(Format(rsTemp!审阅时间, "yyyy-mm-dd") & "")
        Set objItem = objRecord.AddItem(CStr(rsTemp!审阅说明 & ""))
        Set objItem = objRecord.AddItem(CStr(rsTemp!交班期间))
        If rsTemp!审阅人 & "" = "" Then
            objRecord.PreviewText = "尚未审阅"
        Else
            objRecord.PreviewText = "审阅人:" & rsTemp!审阅人 & "" & "  审阅时间:" & Format(rsTemp!审阅时间, "yyyy-mm-dd") & "" & "  审阅说明:" & rsTemp!审阅说明 & ""
        End If
        rsTemp.MoveNext
    Loop
    rptData.Populate
    If mlngRow > 0 And mlngRow <= rptData.Rows.Count - 1 Then
        Set rptData.FocusedRow = rptData.Rows(mlngRow)
        rptData.SetFocus
        Exit Sub
    End If
    If mlngRow = 0 And rptData.Rows.Count > 1 Then
        Set rptData.FocusedRow = rptData.Rows(1)
'        rptData.SetFocus
    End If
    Exit Sub
errH:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub LoadType()
'动态加载值班班次
    Dim lngIndex As Long, lngNum As Long, lngHeight As Long
    Dim objChk As Object
    Dim lngMax As Long, lngMaxNum As Long
    Dim rsTemp As ADODB.Recordset
        
    Set rsTemp = GetShiftType(2, IIf(cboDept.Text = "所有科室", mstrDeptId, mlngDeptID))
    For lngIndex = 1 To chkType.UBound
        Unload chkType(lngIndex)
    Next
    chkType(0).Visible = False
    lngIndex = 0
    lngMax = rsTemp.RecordCount - 1
    If rsTemp.RecordCount > 1 Then rsTemp.MoveFirst
    For lngIndex = 0 To lngMax
        If lngIndex = 0 Then
            chkType(0).Visible = True
            chkType(0).Value = 1
            chkType(0).Caption = rsTemp!班次名称
            chkType(0).Width = 1300
        Else
            Load chkType(lngIndex)
            chkType(lngIndex).Caption = rsTemp!班次名称
            Set objChk = chkType(lngIndex)
            Set chkType(lngIndex).Container = picShift
            lngNum = Fix(lngIndex / 3)
            If lngNum = lngIndex / 3 Then
                chkType(lngIndex).Move chkType(0).Left, chkType(0).Top + (chkType(0).Height + 120) * lngNum, 1300, chkType(0).Height
            Else
                chkType(lngIndex).Move chkType(lngIndex - 1).Left + chkType(lngIndex - 1).Width + 50, chkType(lngIndex - 1).Top, 1300, chkType(0).Height
            End If
            chkType(lngIndex).Visible = True
            chkType(lngIndex).Value = 1
        End If
        rsTemp.MoveNext
    Next
    lngMaxNum = Fix(lngMax / 3) + 1
    picBack.Height = IIf(lngMaxNum > 3, 3, lngMaxNum) * (chkType(0).Height + 120) + 120
    lineY.X1 = picShift.Width
    lineY.X2 = picShift.Width
    lineY.Y1 = 0
    lineY.Y2 = picBack.Height
    lngHeight = lngMaxNum * (chkType(0).Height + 120) + 120
    If lngHeight <= picBack.Height Then lngHeight = picBack.Height
    picShift.Height = lngHeight
    If Fix(lngMax / 3) > 2 Then
        lineY.Visible = True
        picUp.Visible = False
        picDown.Visible = True
        picUp.Top = 0
        picDown.Top = picBack.Height - picDown.Height
    Else
        picBack.BackColor = picShift.BackColor
        lineY.Visible = False
        picUp.Visible = False
        picDown.Visible = False
    End If
    cmdRef.Top = picBack.Top + picBack.Height + 100
    fraFilter.Height = cmdRef.Top + cmdRef.Height + 100
    Call picDataIn_Resize
End Sub

Private Sub InitCommandBar()
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    
    Set cbsMain.Icons = imgPublic.Icons
    
    '菜单定义:包括公共部份
    '    请对xtpControlPopup类型的命令ID重新赋值
    '-----------------------------------------------------
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.id = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_TypeManage, "班次管理(&T)"): objControl.BeginGroup = True
'        Set objControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&C)"): objControl.BeginGroup = True
'        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到Excel...")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): objControl.BeginGroup = True
    End With

    Set mobjMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    mobjMenu.id = conMenu_EditPopup
    With mobjMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_FinOut, "完成交班(&O)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_CancelOut, "取消完成交班(&J)")
        If mintCA > 0 Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_CheckOutSign, "交班签名验证(&C)")
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Edit_FinIn, "完成接班(&I)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_CancelIn, "取消完成接班(&H)")
        If mintCA > 0 Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_CheckInSign, "接班签名验证(&S)")
        End If
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_FinRead, "完成审阅(&F)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_CancelRead, "取消完成审阅(&R)")
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "报表(&R)", -1, False)
    objMenu.id = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Report_Record, "临床科室交接班情况查询(&S)")
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.id = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False)
            objControl.Checked = True
            Set objControl = .Add(xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False)
            objControl.Checked = True
            Set objControl = .Add(xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False)
            objControl.Checked = True
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        objControl.Checked = True
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.id = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
'            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): objControl.BeginGroup = True
    End With

    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagHideWrap
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_FinOut, "完成交班"): objPopup.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_FinIn, "完成接班")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_FinRead, "完成审阅")
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next
    
    '命令的快键绑定:公共部份主界面已处理
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyP, conMenu_File_Print '打印
        
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '新增
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '修改
        .Add 0, vbKeyDelete, conMenu_Edit_Delete '删除
        
        .Add 0, vbKeyF1, conMenu_Help_Help '帮助
    End With
    '设置一些公共的不常用命令
'    With cbsMain.Options
'        .AddHiddenCommand conMenu_File_PrintSet '打印设置
'    End With
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPriv, "ZL1_INSIDE_1242_1")
End Sub

Private Sub Form_Resize()
    Dim lngTop As Long
    Dim lngHeight As Long
    
    On Error Resume Next
    
    If Not cbsMain(2).Visible Then
        lngTop = 500
    End If
    If stbThis.Visible Then
        lngHeight = stbThis.Height
    End If
    picSplitX.Top = 1000 - picSplitY - lngTop
    picSplitX.Height = Me.ScaleHeight - 1000 - lngHeight + lngTop
    
    picDataIn.Move 0, 900 - lngTop, picSplitX.Left, picSplitX.Height
    
    picSplitY.Left = picSplitX.Left + picSplitX.Width
    picSplitY.Width = Me.ScaleWidth - picSplitY.Left
    
    picPatient.Move picSplitY.Left, 1120 - lngTop, picSplitY.Width, picSplitY.Top - 1120 + lngTop
    picSub.Move picPatient.Left, picSplitY.Top + picSplitY.Height, picPatient.Width, picSplitX.Height - picSplitY.Top + 1000 - lngTop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strCols As String
    Dim i As Long

    Call SaveWinState(Me, App.ProductName)
    
    If Not mobjFrom Is Nothing Then
        Unload mobjFrom
        Set mobjFrom = Nothing
    End If
    Set mrsPati = Nothing
    Set mobjESign = Nothing
    Set mobjMenu = Nothing
    Set grsUserInfo = Nothing
    mlngRow = 0
    With rptData
        For i = cold_记录id To cold_交班期间
            strCols = strCols & ";" & .Columns(i).Caption & "|" & .Columns(i).Width
        Next
    End With
    strCols = Mid(strCols, 2)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbaUser & "\" & gstrProductName & "\医生交接班记录", "记录列宽", strCols
    
    strCols = ""
    With vsDetail
        For i = .ColIndex("序号") To .ColIndex("诊断")
            strCols = strCols & ";" & .ColKey(i) & "|" & .ColWidth(i)
        Next
    End With
    strCols = Mid(strCols, 2)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbaUser & "\" & gstrProductName & "\医生交接班记录", "病人清单列宽", strCols
End Sub

Private Sub picDataIn_Resize()
    
    On Error Resume Next
    If picDataIn.Width < 3000 Then Exit Sub
    If picDataIn.Height < 3000 Then Exit Sub
    fraFilter.Move 50, 150
    lblRecord.Move fraFilter.Left, fraFilter.Top + fraFilter.Height + 100
    
    picNumBack.Move fraFilter.Left, picDataIn.Height - picNumBack.Height - lblAllNum.Height
    lblAllNum.Move picNumBack.Left, picNumBack.Top - lblAllNum.Height - 50
    
    rptData.Move fraFilter.Left, lblRecord.Top + lblRecord.Height + 20, picDataIn.Width - 40, lblAllNum.Top - rptData.Top - 150
    
End Sub

Private Sub InitReportColumn()
'功能：初始化病人列表表格
    Dim objCol As ReportColumn
    Dim strRptCol As String '控件的列宽，格式:列名,长度;列名,长度...
    Dim varData As Variant
    Dim i As Long

    strRptCol = GetSetting("ZLSOFT", "私有模块\" & gstrDbaUser & "\" & gstrProductName & "\医生交接班记录", "记录列宽")
    If strRptCol = "" Then
        strRptCol = "记录id|0;科室|0;科室ID|0;交班医生|55;交班班次|55;交班时期|0;接班医生|55;接班班次|55;交班时期|0;交班状态|0;" & _
            "接班状态|0;完成时间|120;审阅人|55;审阅时间|120;审阅说明|200;交班期间|160"
    End If
    With rptData
        '当列顺序或数量(代码或人为隐藏)改变后,要用Find(列号)查找列,但仍可用Record(列号)访问数据行
        Set objCol = .Columns.Add(cold_记录id, "记录id", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(cold_科室, "科室", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(cold_科室ID, "科室ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(cold_交班医生, "交班医生", 0, False)
        Set objCol = .Columns.Add(cold_交班班次, "交班班次", 0, False)
        Set objCol = .Columns.Add(cold_交班时期, "交班时期", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(cold_接班医生, "接班医生", 0, True)
        Set objCol = .Columns.Add(cold_接班班次, "接班班次", 0, True)
        Set objCol = .Columns.Add(cold_交班时期, "交班时期", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(cold_交班状态, "交班状态", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(cold_接班状态, "接班状态", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(cold_完成时间, "完成时间", 0, True)
        Set objCol = .Columns.Add(cold_审阅人, "审阅人", 0, True)
        Set objCol = .Columns.Add(cold_审阅时间, "审阅时间", 0, True)
        Set objCol = .Columns.Add(cold_审阅说明, "审阅说明", 0, True)
        Set objCol = .Columns.Add(cold_交班期间, "交班期间", 0, False)
        varData = Split(strRptCol, ";")
        For i = cold_记录id To cold_交班期间
            .Columns(i).Width = Split(varData(i), "|")(1)
        Next
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = objCol.Index = cold_科室
        Next
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的记录..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '会引发SelectionChanged事件
        .ShowItemsInGroups = False
        
        .GroupsOrder.Add .Columns(cold_科室)
        .GroupsOrder(0).SortAscending = True
    End With
End Sub

Private Sub picnumDown_Click()

    picNum.Top = picNum.Top - (lblTypeNum(0).Height + 120) * 3
    If picNum.Top + picNum.Height > picBack.Height Then
        picNumDown.Visible = True
    Else
        picNumDown.Visible = False
    End If
    If picNum.Top < 0 Then
        picNumUp.Visible = True
    Else
        picNumUp.Visible = False
    End If
End Sub

Private Sub picnumup_Click()
    picNum.Top = picNum.Top + (lblTypeNum(0).Height + 120) * 3
    If picNum.Top > 0 Then picNum.Top = 0
    If picNum.Top + picNum.Height > picBack.Height Then
        picNumDown.Visible = True
    Else
        picNumDown.Visible = False
    End If
    If picNum.Top < 0 Then
        picNumUp.Visible = True
    Else
        picNumUp.Visible = False
    End If
End Sub

Private Sub picPatient_Resize()
    lblList.Move 50, 50
    vsDetail.Move 0, 250, picPatient.Width, picPatient.Height - lblList.Height - lblList.Top - 40
    lblColor2.Move vsDetail.Left + vsDetail.Width - lblColor2.Width - 300, lblList.Top
    lblColor.Move lblColor2.Left - lblColor.Width - 50, lblColor2.Top
End Sub


Private Sub picShow_Resize()
    tbcSub.Top = 0: tbcSub.Left = 0
    tbcSub.Width = picShow.Width: tbcSub.Height = picShow.Height
End Sub

Private Sub picSplitX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglNew As Single
    
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    
    If picSplitX.Tag <> "Draging" Then
        picSplitX.Tag = "Draging"
        picSplitX.BackColor = 0
    End If
    
    sglNew = picSplitX.Left + X
    
    picSplitX.Left = sglNew
End Sub

Private Sub picSplitX_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    If picSplitX.Tag = "Draging" Then
        picPatient.Width = Me.ScaleWidth - picSplitX.Left
        Call Form_Resize
        picSplitX.BackColor = Me.BackColor
        picSplitX.Tag = ""
    End If
End Sub

Private Sub picSplitY_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglNew As Single
    
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    
    If picSplitY.Tag <> "Draging" Then
        picSplitY.Tag = "Draging"
        picSplitY.BackColor = 0
    End If
    
    sglNew = picSplitY.Top + Y
    
    picSplitY.Top = sglNew
End Sub

Private Sub picSplitY_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    If picSplitY.Tag = "Draging" Then
        picPatient.Height = Me.ScaleHeight - picSplitY.Top
        Call Form_Resize
        picSplitY.BackColor = Me.BackColor
        picSplitY.Tag = ""
    End If
End Sub

Private Sub picSub_Resize()
    '隐藏tab标签
    picShow.Top = -360: picShow.Left = 0
    picShow.Width = picSub.Width: picShow.Height = picSub.Height + 360
    
    tbcSub.Top = 0: tbcSub.Left = 0
    tbcSub.Width = picShow.Width: tbcSub.Height = picShow.Height
End Sub

Private Sub picDown_Click()

    picShift.Top = picShift.Top - (chkType(0).Height + 120) * 2
    If picShift.Top + picShift.Height > picBack.Height Then
        picDown.Visible = True
    Else
        picDown.Visible = False
    End If
    If picShift.Top < 0 Then
        picUp.Visible = True
    Else
        picUp.Visible = False
    End If
End Sub

Private Sub picUp_Click()
    picShift.Top = picShift.Top + (chkType(0).Height + 120) * 2
    If picShift.Top > 0 Then picShift.Top = 0
    If picShift.Top + picShift.Height > picBack.Height Then
        picDown.Visible = True
    Else
        picDown.Visible = False
    End If
    If picShift.Top < 0 Then
        picUp.Visible = True
    Else
        picUp.Visible = False
    End If
End Sub

Private Sub rptData_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If Button = 2 Then
        Call ShowContenMenu.ShowPopup
    End If
End Sub
Private Function ShowContenMenu() As CommandBar
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrControl2 As CommandBarControl
    Dim cbrControl3 As CommandBarControl
    
    '弹出菜单处理
    On Error GoTo ErrHand
    Set cbrPopupBar = cbsMain.Add("弹出菜单", xtpBarPopup)
    For Each cbrControl In mobjMenu.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(cbrControl.Type, cbrControl.id, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
        cbrPopupItem.Parameter = cbrControl.Parameter
        cbrPopupItem.Visible = cbrControl.Visible
        cbrPopupItem.IconId = cbrControl.IconId
        cbrPopupItem.Checked = cbrControl.Checked
        cbrPopupItem.Style = cbrControl.Style
        If cbrControl.Type = xtpControlPopup Or cbrControl.Type = xtpControlSplitButtonPopup Then
            For Each cbrControl2 In cbrControl.CommandBar.Controls
                Set cbrControl3 = cbrPopupItem.CommandBar.Controls.Add(xtpControlButton, cbrControl2.id, cbrControl2.Caption)
                cbrControl3.BeginGroup = cbrControl2.BeginGroup
                cbrControl3.Parameter = cbrControl2.Parameter
                cbrControl3.Visible = cbrControl2.Visible
                cbrControl3.Style = cbrControl2.Style
            Next
        End If
    Next
    Set ShowContenMenu = cbrPopupBar
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub rptData_SelectionChanged()
    Dim rsTemp As ADODB.Recordset
    Dim lngId As Long, i As Long
    Dim dt开始时间 As Date, dt结束时间 As Date
    Dim btnNoEdit As Boolean
    
    If mblnLoading Then Exit Sub
    If rptData.Records.Count < 1 Then Exit Sub
    vsDetail.Rows = 1
    
    If rptData.SelectedRows(0).GroupRow Then
        Call mobjFrom.zlRefresh(0, 0, 0, 0, 0, 0, 0, False, "") '清空编辑列表
        Exit Sub
    End If
    mlngRow = rptData.FocusedRow.Index
    lngId = Val(rptData.SelectedRows(0).Record(cold_记录id).Value)
    On Error GoTo errH
    gstrSQL = "Select a.序号,a.内容id,a.病人ID,a.主页ID,a.病人类型 as 类型, a.姓名," & vbNewLine & _
        "a.性别, a.年龄, a.床号, a.标识号, to_char(a.入院时间,'yyyy-mm-dd') 入院时间, a.入院方式, a.交班描述 From 医生交接班内容 a Where 记录id =[1] order by 序号"
    Set mrsPati = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngId)
    
    If mrsPati.EOF And (Not mobjFrom Is Nothing) Then '清空编辑列表
        dt开始时间 = CDate(Split(rptData.SelectedRows(0).Record(cold_交班时期).Value, "|")(0))
        dt结束时间 = CDate(Split(rptData.SelectedRows(0).Record(cold_交班时期).Value, "|")(1))
        btnNoEdit = Val(rptData.SelectedRows(0).Record(cold_交班状态).Value) = 1 Or (Not CheckPriv("医生交接班"))
        Call mobjFrom.zlRefresh(0, 0, Val(rptData.SelectedRows(0).Record(cold_科室ID).Value), 0, Val(rptData.SelectedRows(0).Record(cold_记录id).Value), dt开始时间, dt结束时间, btnNoEdit, "")
    End If
    
    With vsDetail
        .Redraw = flexRDNone
        .Rows = mrsPati.RecordCount + 1
        Do While Not mrsPati.EOF
            .TextMatrix(mrsPati.AbsolutePosition, .ColIndex("序号")) = mrsPati!序号
            .TextMatrix(mrsPati.AbsolutePosition, .ColIndex("内容id")) = mrsPati!内容ID & ""
            .TextMatrix(mrsPati.AbsolutePosition, .ColIndex("病人id")) = mrsPati!病人ID & ""
            .TextMatrix(mrsPati.AbsolutePosition, .ColIndex("主页ID")) = mrsPati!主页ID & ""
            .TextMatrix(mrsPati.AbsolutePosition, .ColIndex("类型")) = mrsPati!类型 & ""
            .TextMatrix(mrsPati.AbsolutePosition, .ColIndex("姓名")) = mrsPati!姓名 & ""
            .TextMatrix(mrsPati.AbsolutePosition, .ColIndex("性别")) = mrsPati!性别 & ""
            .TextMatrix(mrsPati.AbsolutePosition, .ColIndex("年龄")) = mrsPati!年龄 & ""
            .TextMatrix(mrsPati.AbsolutePosition, .ColIndex("床号")) = mrsPati!床号 & ""
            .TextMatrix(mrsPati.AbsolutePosition, .ColIndex("标识号")) = mrsPati!标识号 & ""
            .TextMatrix(mrsPati.AbsolutePosition, .ColIndex("入院时间")) = mrsPati!入院时间 & ""
            .TextMatrix(mrsPati.AbsolutePosition, .ColIndex("入院方式")) = mrsPati!入院方式 & ""
            .TextMatrix(mrsPati.AbsolutePosition, .ColIndex("交班描述")) = mrsPati!交班描述 & ""
            If IsNull(mrsPati!交班描述) Then
                .Cell(flexcpBackColor, mrsPati.AbsolutePosition, 0, mrsPati.AbsolutePosition, .Cols - 1) = RGB(255, 239, 219)
            End If
            mrsPati.MoveNext
        Loop
        .Redraw = flexRDDirect
        If .Rows > 1 Then
            vsDetail.Row = 1
            vsDetail.ShowCell 1, 0
            If .Rows > 2 Then
                vsDetail.Cell(flexcpPicture, 1, .ColIndex("下移")) = imgList.ListImages("Down").Picture
            End If
        End If
        .AutoSizeMode = flexAutoSizeRowHeight
        .WordWrap = True
        .AutoSize .ColIndex("交班描述")
    End With
    Call LoadNum(lngId)
    Exit Sub
errH:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub LoadNum(ByVal lngId As Long)
'动态显示每条记录的实际汇总情况
    Dim rsTemp As ADODB.Recordset
    Dim lngIndex As Long, lngMax As Long, lngMaxNum As Long, lngNum As Long
    Dim lngHeight As Long
    Dim objlbl As Object
    
    For lngIndex = 1 To lblTypeNum.UBound
        Unload lblTypeNum(lngIndex)
    Next
    gstrSQL = "Select 项目, 数量 From 医生交接班汇总 Where 记录id =[1] Order By 序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngId)
    lblTypeNum(0).Visible = False
    lngIndex = 0
    lngMax = rsTemp.RecordCount - 1
    If rsTemp.RecordCount > 1 Then rsTemp.MoveFirst
    For lngIndex = 0 To lngMax
        If rsTemp!项目 = "住院总" Then
            lblAllNum.Caption = "病人汇总情况 【" & rptData.SelectedRows(0).Record(cold_科室).Value & "】  住院总人数：" & rsTemp!数量
            lngMax = lngMax - 1
        Else
            If lngIndex = 0 Then
                lblTypeNum(0).Visible = True
                lblTypeNum(0).Caption = rsTemp!项目 & "人数：" & rsTemp!数量
                lblTypeNum(0).Width = 1800
            Else
                Load lblTypeNum(lngIndex)
                lblTypeNum(lngIndex).Caption = rsTemp!项目 & "人数：" & rsTemp!数量
                Set objlbl = lblTypeNum(lngIndex)
                Set lblTypeNum(lngIndex).Container = picNum
                lngNum = Fix(lngIndex / 3)
                If lngNum = lngIndex / 3 Then
                    lblTypeNum(lngIndex).Move lblTypeNum(0).Left, lblTypeNum(0).Top + (lblTypeNum(0).Height + 120) * lngNum, 1800, lblTypeNum(0).Height
                Else
                    lblTypeNum(lngIndex).Move lblTypeNum(lngIndex - 1).Left + lblTypeNum(lngIndex - 1).Width + 50, lblTypeNum(lngIndex - 1).Top, 1800, lblTypeNum(0).Height
                End If
                lblTypeNum(lngIndex).Visible = True
            End If
        End If
        rsTemp.MoveNext
    Next
    lngMaxNum = Fix(lngMax / 3) + 1
    lineNumY.X1 = picNum.Width
    lineNumY.X2 = picNum.Width
    lineNumY.Y1 = 0
    lineNumY.Y2 = picNumBack.Height
    lngHeight = lngMaxNum * (lblTypeNum(0).Height + 120) + 120
    If lngHeight <= picNumBack.Height Then lngHeight = picNumBack.Height
    picNum.Height = lngHeight
    If Fix(lngMax / 3) > 3 Then
        lineNumY.Visible = True
        picNumUp.Visible = False
        picNumDown.Visible = True
        picNumUp.Top = 0
        picNumDown.Top = picNumBack.Height - picNumDown.Height
    Else
        picNumBack.BackColor = picNum.BackColor
        lineNumY.Visible = False
        picNumUp.Visible = False
        picNumDown.Visible = False
    End If
    Call picDataIn_Resize
End Sub

Private Function CheckPriv(ByVal strPri As String) As Boolean
'判断是否具有某个权限
    If InStr(";" & mstrPriv & ";", ";" & strPri & ";") > 0 Then
        CheckPriv = True
    End If
End Function

Private Sub tvwSubject_DblClick()
    
    txtSubject.Text = tvwSubject.SelectedItem.Text
    txtSubject.Tag = Mid(tvwSubject.SelectedItem.Key, 2)
    Call LoadDept(Mid(tvwSubject.SelectedItem.Key, 2))
    tvwSubject.Visible = False
    Call LoadType
End Sub

Private Sub tvwSubject_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Call tvwSubject_LostFocus
    End If
End Sub

Private Sub tvwSubject_LostFocus()
    tvwSubject.Visible = False
End Sub

Private Sub vsDetail_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim dt开始时间 As Date, dt结束时间 As Date
    Dim btnNoEdit As Boolean
    
    If mblnLoading Then Exit Sub
    If OldRow = NewRow Or NewRow < 1 Then Exit Sub
    With vsDetail
        If NewRow = 1 Then
            If .Rows = 2 Then
                .Cell(flexcpPicture, NewRow, .ColIndex("上移")) = ""
                .Cell(flexcpPicture, NewRow, .ColIndex("下移")) = ""
            Else
                .Cell(flexcpPicture, NewRow, .ColIndex("上移")) = ""
                .Cell(flexcpPicture, NewRow, .ColIndex("下移")) = imgList.ListImages("Down").Picture
            End If
        Else
            If NewRow = .Rows - 1 Then
                .Cell(flexcpPicture, NewRow, .ColIndex("下移")) = ""
                .Cell(flexcpPicture, NewRow, .ColIndex("上移")) = imgList.ListImages("Up").Picture
            Else
                .Cell(flexcpPicture, NewRow, .ColIndex("上移")) = imgList.ListImages("Up").Picture
                .Cell(flexcpPicture, NewRow, .ColIndex("下移")) = imgList.ListImages("Down").Picture
            End If
        End If
        If OldRow < .Rows Then
            .Cell(flexcpPicture, OldRow, .ColIndex("上移")) = ""
            .Cell(flexcpPicture, OldRow, .ColIndex("下移")) = ""
        End If

        btnNoEdit = Val(rptData.SelectedRows(0).Record(cold_交班状态).Value) = 1 Or (Not CheckPriv("医生交接班"))
        With vsDetail
            If (Not mobjFrom Is Nothing) And .TextMatrix(.Row, .ColIndex("内容ID")) <> "" Then
                dt开始时间 = CDate(Split(rptData.SelectedRows(0).Record(cold_交班时期).Value, "|")(0))
                dt结束时间 = CDate(Split(rptData.SelectedRows(0).Record(cold_交班时期).Value, "|")(1))
                Call mobjFrom.zlRefresh(Val(.TextMatrix(.Row, .ColIndex("病人ID"))), Val(.TextMatrix(.Row, .ColIndex("主页ID"))), Val(rptData.SelectedRows(0).Record(cold_科室ID).Value), Val(.TextMatrix(.Row, .ColIndex("内容ID"))), Val(rptData.SelectedRows(0).Record(cold_记录id).Value), dt开始时间, dt结束时间, btnNoEdit, .TextMatrix(.Row, .ColIndex("类型")))
            End If
        End With
    End With
End Sub

Private Sub vsDetail_Click()
    Dim blnBegin As Boolean
    Dim lngRow As Long, lngId As Long, lngColor As Long
    Dim i As Long
    
    On Error GoTo errH
    
    With vsDetail
        If .Row < 1 Then Exit Sub
        If .Col = .ColIndex("上移") Then
            If Not .Cell(flexcpPicture, .Row, .ColIndex("上移")) Is Nothing Then
                lngRow = .Row - 1
            End If
        ElseIf .Col = .ColIndex("下移") Then
            If Not .Cell(flexcpPicture, .Row, .ColIndex("下移")) Is Nothing Then
                lngRow = .Row + 1
            End If
        End If
        If lngRow = 0 Then Exit Sub
        '上下移动时，序号是不变的，表示病人的顺序
        '序号,类型,姓名,性别,年龄,床号,标识号,入院时间，入院方式，主诉，诊断，交班描述
        
        lngId = .TextMatrix(.Row, .ColIndex("内容id"))
        lngColor = .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1)
        mrsPati.Filter = "内容id=" & .TextMatrix(lngRow, .ColIndex("内容id"))
        If mrsPati.RecordCount = 1 Then
            For i = 0 To .Cols - 1
                '序号不需要交换
                If i <> .ColIndex("序号") And i <> .ColIndex("上移") And i <> .ColIndex("下移") Then
                    .TextMatrix(.Row, i) = mrsPati.Fields(.ColKey(i)).Value & ""
               End If
            Next
            .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1)
        Else
            MsgBox "数据已被删除，无法上移或下移!", vbExclamation, Me.Caption
            Exit Sub
        End If

        mrsPati.Filter = "内容id=" & lngId
        If mrsPati.RecordCount = 1 Then
            For i = 0 To .Cols - 1
                '序号不需要交换
                If i <> .ColIndex("序号") And i <> .ColIndex("上移") And i <> .ColIndex("下移") Then
                    .TextMatrix(lngRow, i) = mrsPati.Fields(.ColKey(i)).Value & ""
               End If
            Next
            .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1) = lngColor
        Else
            MsgBox "数据已被删除，无法上移或下移!", vbExclamation, Me.Caption
        End If
        '数据库中的交接班内容序号应一致调整
        gcnOracle.BeginTrans: blnBegin = True
        gstrSQL = "Zl_医生交接班内容_Edit(3," & .TextMatrix(.Row, .ColIndex("内容id")) & ",NULL," & .TextMatrix(.Row, .ColIndex("序号")) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "调整序号")
        
        gstrSQL = "Zl_医生交接班内容_Edit(3," & .TextMatrix(lngRow, .ColIndex("内容id")) & ",NULL," & .TextMatrix(lngRow, .ColIndex("序号")) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "调整序号")
        gcnOracle.CommitTrans
        .Row = lngRow
        .ShowCell lngRow, 1
    End With
    Exit Sub
errH:
    If blnBegin Then gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub GetFrom()
    Set mobjFrom = New frmShiftEdit
    mobjFrom.BorderStyle = FormBorderStyleConstants.vbBSNone '设置为无边框
    mobjFrom.Caption = "病人交接班内容编辑"
    Set mobjFrom.gfrmParent = Me
    
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
        .InsertItem(0, "交接班病人编辑", mobjFrom.hWnd, 0).Tag = "交接班病人编辑"
        .Item(0).Selected = True
    End With
End Sub

Public Sub SetEnable(Optional intType As Integer)
    '控制主窗体是否允许编辑
    'intType 0=可用,1=不可用
    picDataIn.Enabled = intType = 0
    picPatient.Enabled = intType = 0
    mblnEdit = intType <> 0
End Sub

Public Sub RefreshEdit(ByVal lngUPid As Long)
    '刷新交班病人列表
    Dim lngFind As Long
    Call rptData_SelectionChanged
    
    '定位病人
    lngFind = vsDetail.FindRow(lngUPid, vsDetail.FixedRows, vsDetail.ColIndex("内容ID"))
    If lngFind > 0 Then
        vsDetail.Row = lngFind
        Call vsDetail.ShowCell(lngFind, vsDetail.ColIndex("姓名"))
    Else
        If vsDetail.Rows - 1 > 0 Then
            vsDetail.Row = vsDetail.Rows - 1
            Call vsDetail.ShowCell(vsDetail.Rows - 1, vsDetail.ColIndex("姓名"))
        End If
    End If
End Sub

Private Sub vsDetail_DblClick()
    If (Not mobjFrom Is Nothing) And vsDetail.TextMatrix(vsDetail.Row, vsDetail.ColIndex("内容ID")) <> "" Then
        Call mobjFrom.EditState
    End If
End Sub


