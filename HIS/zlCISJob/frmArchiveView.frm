VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmArchiveView 
   AutoRedraw      =   -1  'True
   Caption         =   "电子病案查阅"
   ClientHeight    =   9555
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   12660
   Icon            =   "frmArchiveView.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9555
   ScaleWidth      =   12660
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   11880
      Top             =   4560
   End
   Begin VB.PictureBox picRpt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   11505
      ScaleHeight     =   780
      ScaleWidth      =   915
      TabIndex        =   46
      Top             =   2190
      Width           =   915
      Begin SHDocVwCtl.WebBrowser webRpt 
         Height          =   450
         Left            =   135
         TabIndex        =   47
         Top             =   150
         Width           =   450
         ExtentX         =   794
         ExtentY         =   794
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.PictureBox pic观片 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   350
      Left            =   11160
      ScaleHeight     =   345
      ScaleWidth      =   1095
      TabIndex        =   45
      Top             =   930
      Visible         =   0   'False
      Width           =   1100
      Begin VB.Image img观片 
         Height          =   350
         Left            =   0
         Picture         =   "frmArchiveView.frx":058A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1100
      End
   End
   Begin VB.PictureBox picTmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   11670
      ScaleHeight     =   285
      ScaleWidth      =   315
      TabIndex        =   44
      Top             =   1605
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.ComboBox cboDept 
      Height          =   300
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Top             =   0
      Width           =   3165
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   3765
      ScaleHeight     =   975
      ScaleWidth      =   7695
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   135
      Width           =   7695
      Begin VB.Frame fraInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " 基本就诊信息 "
         ForeColor       =   &H80000008&
         Height          =   840
         Left            =   60
         TabIndex        =   5
         Top             =   75
         Width           =   7500
         Begin VB.Frame fraIn 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   195
            TabIndex        =   24
            Top             =   255
            Visible         =   0   'False
            Width           =   7170
            Begin VB.Label lbl类型zy 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   4770
               TabIndex        =   42
               Top             =   0
               Width           =   1080
            End
            Begin VB.Label lbl类型zy 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "类型:"
               Height          =   180
               Index           =   0
               Left            =   4305
               TabIndex        =   41
               Top             =   0
               Width           =   450
            End
            Begin VB.Label lbl住院号zy 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "住院号:"
               Height          =   180
               Index           =   0
               Left            =   1560
               TabIndex        =   40
               Top             =   0
               Width           =   630
            End
            Begin VB.Label lbl姓名zy 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "姓名:"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   39
               Top             =   0
               Width           =   450
            End
            Begin VB.Label lbl付款zy 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "付款:"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   38
               Top             =   255
               Width           =   450
            End
            Begin VB.Label lbl床号zy 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "床号:"
               Height          =   180
               Index           =   0
               Left            =   3150
               TabIndex        =   37
               Top             =   0
               Width           =   450
            End
            Begin VB.Label lbl医保号zy 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "医保号:"
               Height          =   180
               Index           =   0
               Left            =   5940
               TabIndex        =   36
               Top             =   0
               Width           =   630
            End
            Begin VB.Label lbl入院zy 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入院:"
               Height          =   180
               Index           =   0
               Left            =   4305
               TabIndex        =   35
               Top             =   255
               Width           =   450
            End
            Begin VB.Label lbl病况zy 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "病况:"
               Height          =   180
               Index           =   0
               Left            =   3150
               TabIndex        =   34
               Top             =   255
               Width           =   450
            End
            Begin VB.Label lbl护理zy 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "护  理:"
               Height          =   180
               Index           =   0
               Left            =   1560
               TabIndex        =   33
               Top             =   255
               Width           =   630
            End
            Begin VB.Label lbl护理zy 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   2190
               TabIndex        =   32
               Top             =   255
               Width           =   900
            End
            Begin VB.Label lbl病况zy 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H000000FF&
               Height          =   180
               Index           =   1
               Left            =   3585
               TabIndex        =   31
               Top             =   255
               Width           =   675
            End
            Begin VB.Label lbl入院zy 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   4770
               TabIndex        =   30
               Top             =   255
               Width           =   90
            End
            Begin VB.Label lbl医保号zy 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00008000&
               Height          =   180
               Index           =   1
               Left            =   6600
               TabIndex        =   29
               Top             =   0
               Width           =   90
            End
            Begin VB.Label lbl床号zy 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   3585
               TabIndex        =   28
               Top             =   0
               Width           =   675
            End
            Begin VB.Label lbl付款zy 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   435
               TabIndex        =   27
               Top             =   255
               Width           =   1080
            End
            Begin VB.Label lbl姓名zy 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   435
               TabIndex        =   26
               Top             =   0
               Width           =   1080
            End
            Begin VB.Label lbl住院号zy 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   2190
               TabIndex        =   25
               Top             =   0
               Width           =   900
            End
         End
         Begin VB.Frame fraOut 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   195
            TabIndex        =   6
            Top             =   255
            Visible         =   0   'False
            Width           =   7170
            Begin VB.Label lbl急 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "急"
               BeginProperty Font 
                  Name            =   "黑体"
                  Size            =   21.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   435
               Left            =   6705
               TabIndex        =   23
               Top             =   0
               Visible         =   0   'False
               Width           =   435
            End
            Begin VB.Label lbl挂号单mz 
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   3870
               TabIndex        =   22
               Top             =   0
               Width           =   1065
            End
            Begin VB.Label lbl挂号单mz 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "挂号单:"
               Height          =   180
               Index           =   0
               Left            =   3255
               TabIndex        =   21
               Top             =   0
               Width           =   630
            End
            Begin VB.Label lbl医生mz 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   2385
               TabIndex        =   20
               Top             =   0
               Width           =   780
            End
            Begin VB.Label lbl医生mz 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "医生:"
               Height          =   180
               Index           =   0
               Left            =   1935
               TabIndex        =   19
               Top             =   0
               Width           =   450
            End
            Begin VB.Label lbl社区号mz 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00008000&
               Height          =   180
               Index           =   1
               Left            =   5655
               TabIndex        =   18
               Top             =   255
               Width           =   90
            End
            Begin VB.Label lbl社区号mz 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "社区号:"
               Height          =   180
               Index           =   0
               Left            =   5025
               TabIndex        =   17
               Top             =   255
               Width           =   630
            End
            Begin VB.Label lbl门诊号mz 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "门诊号:"
               Height          =   180
               Index           =   0
               Left            =   3240
               TabIndex        =   16
               Top             =   255
               Width           =   630
            End
            Begin VB.Label lbl姓名mz 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "姓名:"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   15
               Top             =   0
               Width           =   450
            End
            Begin VB.Label lbl费别mz 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "费别:"
               Height          =   180
               Index           =   0
               Left            =   1935
               TabIndex        =   14
               Top             =   255
               Width           =   450
            End
            Begin VB.Label lbl医保号mz 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "医保号:"
               Height          =   180
               Index           =   0
               Left            =   5025
               TabIndex        =   13
               Top             =   0
               Width           =   630
            End
            Begin VB.Label lbl医保号mz 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00008000&
               Height          =   180
               Index           =   1
               Left            =   5655
               TabIndex        =   12
               Top             =   0
               Width           =   90
            End
            Begin VB.Label lbl费别mz 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   2385
               TabIndex        =   11
               Top             =   255
               Width           =   765
            End
            Begin VB.Label lbl姓名mz 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   450
               TabIndex        =   10
               Top             =   0
               Width           =   1425
            End
            Begin VB.Label lbl门诊号mz 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   3870
               TabIndex        =   9
               Top             =   255
               Width           =   1095
            End
            Begin VB.Label lbl付款mz 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "付款:"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   8
               Top             =   255
               Width           =   450
            End
            Begin VB.Label lbl付款mz 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   450
               TabIndex        =   7
               Top             =   255
               Width           =   1455
            End
         End
      End
   End
   Begin MSComctlLib.TreeView tvwArchive 
      Height          =   5865
      Left            =   315
      TabIndex        =   3
      Top             =   1170
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   10345
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   0
   End
   Begin VB.Frame fraLR 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2745
      Left            =   3660
      MousePointer    =   9  'Size W E
      TabIndex        =   2
      Top             =   1515
      Width           =   45
   End
   Begin XtremeSuiteControls.TabControl tbcArchive 
      Height          =   6315
      Left            =   3900
      TabIndex        =   1
      Top             =   1605
      Width           =   7365
      _Version        =   589884
      _ExtentX        =   12991
      _ExtentY        =   11139
      _StockProps     =   64
   End
   Begin XtremeSuiteControls.TabControl tbcHistory 
      Height          =   7245
      Left            =   240
      TabIndex        =   0
      Top             =   735
      Width           =   3210
      _Version        =   589884
      _ExtentX        =   5662
      _ExtentY        =   12779
      _StockProps     =   64
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   120
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":21EE
            Key             =   "住院"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":8A50
            Key             =   "门诊"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":8FEA
            Key             =   "object_report"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":9584
            Key             =   "object_case"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":9B1E
            Key             =   "object_tend"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":A0B8
            Key             =   "object_first"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":A652
            Key             =   "object_advice"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":ABEC
            Key             =   "object_file"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":B186
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":119E8
            Key             =   "Path"
         EndProperty
      EndProperty
   End
   Begin VB.Image img观片普 
      Height          =   345
      Left            =   8040
      Picture         =   "frmArchiveView.frx":11F82
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Image img观片高 
      Height          =   345
      Left            =   9240
      Picture         =   "frmArchiveView.frx":13BE6
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "frmArchiveView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobjRichEMR As Object
Private mclsOutAdvices As clsDockOutAdvices
Private mclsInAdvices As clsDockInAdvices
Private mclsDockAduits As clsDockAduits
Private mclsPath As zlPublicPath.clsDockPath
Private WithEvents mclsTendsNew As zl9TendFile.clsTendFile    '新版护士工作站
Attribute mclsTendsNew.VB_VarHelpID = -1
Private mclsArchive As zlMedRecPage.clsArchive '电子病案查阅窗体类
Private mobjPublicPACS As Object

Private mlng病人ID  As Long
Private mlng就诊ID As Long '病人当前或者最后的就诊ID，门诊为挂号ID,住院号主页ID
Private mstr挂号单 As String
Private mlng科室ID As Long
Private mlng病区ID As Long
Private mblnMoved As Boolean
Private mblnNewTends As Boolean
Private mrsData As ADODB.Recordset

Private mcolSubForm As Collection
Private mblnTabTmp As Boolean
Private mlngPreDept As Long
Private mobjPatient As Object
Private mstrTempDel As String        '删除临时文件
Private mstrKey As String            '缺省显示信息

Public Sub ShowArchive(ByVal frmParent As Object, ByVal lng病人ID As Long, ByVal lng就诊ID As Long, Optional ByVal blnModal As Boolean)
'功能：公共接口方法，类似 ShowMe方法
    mblnMoved = False: mblnNewTends = False
    mlng病人ID = lng病人ID
    mlng就诊ID = lng就诊ID
    
    Me.Show IIf(blnModal, 1, 0), frmParent
End Sub

Public Sub zlRefresh(ByVal lng病人ID As Long, ByVal lng就诊ID As Long)
'功能：公共接口方法，类似 ShowMe方法
    mblnMoved = False: mblnNewTends = False
    mlng病人ID = lng病人ID
    mlng就诊ID = lng就诊ID
    
    Call InitBasicData
End Sub

Private Sub InitBasicData()
'功能：初始化一些基本数据，如下拉列表加载等
    Dim strSQL As String
    Dim objTab As TabControlItem
    Dim strTmp As String
    Dim str病人IDs As String
    Dim i As Long
    Dim rsTmp As ADODB.Recordset
    Dim strErr As String
    Dim blnTmp As Boolean
    Dim str身份证号 As String
    Dim strTemp As String
    Dim n As Long, p As Long
    Dim strThis As String
    Dim strSQLPati As String
    Dim varPar(0 To 10) As String
    
    Screen.MousePointer = 11
    LockWindowUpdate Me.hwnd
    mstr挂号单 = "": mlngPreDept = -1
    Call cboDept.Clear
    Call tbcHistory.RemoveAll
    If mlng病人ID = 0 Then
        Set objTab = tbcHistory.InsertItem(tbcHistory.ItemCount, "", tvwArchive.hwnd, 0)
        fraIn.Visible = False: fraOut.Visible = False
        Call ShowOutPatiInfo
        Call ShowArchiveTree
    Else
        On Error GoTo errH
 
        strSQL = "select a.身份证号 from 病人信息 a where a.病人id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
        strTmp = rsTmp!身份证号 & ""
        If strTmp <> "" Then
            '验证身份证号的合法性
            If mobjPatient Is Nothing Then
                On Error Resume Next
                Set mobjPatient = CreateObject("zlPublicPatient.clsPublicPatient")
                err.Clear: On Error GoTo 0
            End If
            If mobjPatient Is Nothing Then
                MsgBox "创建病人信息公共部件（zlPublicPatient.clsPublicPatient）失败！", vbInformation, Me.Caption
            Else
                Call mobjPatient.zlInitCommon(gcnOracle, glngSys, UserInfo.用户名)
                If mobjPatient.CheckPatiIdcard(strTmp) Then
                    str身份证号 = strTmp
                End If
            End If
        End If
        
        On Error GoTo errH
        
        If str身份证号 <> "" Then
            strSQL = "select a.病人id from 病人信息 a where a.病人id<>[1] and a.身份证号=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, str身份证号)
            Do While Not rsTmp.EOF
                str病人IDs = str病人IDs & "," & rsTmp!病人ID
                rsTmp.MoveNext
            Loop
        End If
        If str病人IDs = "" Then
            strSQL = " Select 病人id,ID as 就诊ID,NO,发生时间 as 开始时间,Null as 结束时间,执行部门ID as 科室ID,0 as 数据转出,-1 as 病人性质,null as 就诊号 From 病人挂号记录 Where 病人ID=[1] And 记录性质=1 And 记录状态=1" & _
                " Union ALL" & _
                " Select 病人id,ID as 就诊ID,NO,发生时间 as 开始时间,Null as 结束时间,执行部门ID as 科室ID,1 as 数据转出,-1 as 病人性质,null as 就诊号 From H病人挂号记录 Where 病人ID=[1] And 记录性质=1 And 记录状态=1" & _
                " Union ALL" & _
                " Select 病人id,主页ID as 就诊ID,Null,入院日期 as 开始时间,出院日期,出院科室ID,数据转出,NVL(病人性质,0) as 病人性质,null as 就诊号 From 病案主页 Where 病人ID=[1] And Nvl(主页ID,0)<>0"
            strSQL = "Select Rownum As 序号,a.病人ID,A.就诊ID,A.NO,A.开始时间,A.结束时间,B.名称 as 科室,A.数据转出 ,A.病人性质,a.就诊号 From (" & strSQL & ") A,部门表 B Where A.科室ID=B.ID Order by 开始时间 Desc"
            Set mrsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
        Else
            str病人IDs = mlng病人ID & str病人IDs
            strTemp = "Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) X"
            n = 0
            Do While True
                If Len(str病人IDs) < 4000 Then
                    p = Len(str病人IDs) + 1
                Else
                    p = InStrRev(Mid(str病人IDs, 1, 4000), ",")
                End If
                strThis = Mid(str病人IDs, 1, p - 1)
                If n > 10 Then
                    strSQLPati = strSQLPati & vbNewLine & " Union All " & Replace(strTemp, "[1]", "'" & strThis & "'")
                Else
                    varPar(n) = strThis
                    strSQLPati = IIf(strSQLPati = "", "", strSQLPati & vbNewLine & " Union All ") & Replace(strTemp, "[1]", "[" & (n + 1) & "]")
                End If
                n = n + 1
                str病人IDs = Mid(str病人IDs, p + 1)
                If str病人IDs = "" Then Exit Do
            Loop
            strTmp = " 病人ID In (" & strSQLPati & ")"
            strSQL = " Select 病人id,ID as 就诊ID,NO,发生时间 as 开始时间,Null as 结束时间,执行部门ID as 科室ID,0 as 数据转出,-1 as 病人性质,null as 就诊号 From 病人挂号记录 Where " & strTmp & " And 记录性质=1 And 记录状态=1 and NO is not null" & _
                " Union ALL" & _
                " Select 病人id,ID as 就诊ID,NO,发生时间 as 开始时间,Null as 结束时间,执行部门ID as 科室ID,1 as 数据转出,-1 as 病人性质,null as 就诊号 From H病人挂号记录 Where " & strTmp & " And 记录性质=1 And 记录状态=1 and NO is not null" & _
                " Union ALL" & _
                " Select 病人id,主页ID as 就诊ID,Null,入院日期 as 开始时间,出院日期,出院科室ID,数据转出,NVL(病人性质,0) as 病人性质,住院号 as 就诊号 From 病案主页 Where " & strTmp & " And Nvl(主页ID,0)<>0"
            strSQL = "Select Rownum As 序号,a.病人ID,A.就诊ID,A.NO,A.开始时间,A.结束时间,B.名称 as 科室,A.数据转出 ,A.病人性质,a.就诊号 From (" & strSQL & ") A,部门表 B Where A.科室ID=B.ID  Order by 开始时间 Desc"
            Set mrsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, varPar(0), varPar(1), varPar(2), varPar(3), varPar(4), varPar(5), varPar(6), varPar(7), varPar(8), varPar(9), varPar(10))
        End If
        
        Do While Not mrsData.EOF
        
            strTmp = IIf(IsNull(mrsData!NO), "第" & mrsData!就诊id & "次" & IIf(mrsData!病人性质 = 1, "门诊留观", IIf(mrsData!病人性质 = 2, "住院留观", "住院")), "门诊就诊") & ":" & mrsData!科室 & "," & Format(mrsData!开始时间, "yyyy-MM-dd HH:mm") & _
                IIf(Not IsNull(mrsData!结束时间), "～" & Format(mrsData!结束时间, "yyyy-MM-dd HH:mm"), "")
                
            If mrsData.AbsolutePosition = 1 Then
                Set objTab = tbcHistory.InsertItem(tbcHistory.ItemCount, strTmp, tvwArchive.hwnd, IIf(IsNull(mrsData!NO), 0, 1))
            End If

            cboDept.AddItem strTmp
            cboDept.ItemData(cboDept.NewIndex) = Val(mrsData!序号)
            
            mrsData.MoveNext
        Loop
        If cboDept.ListCount > 0 Then
            Call Cbo.SetIndex(cboDept.hwnd, 0)
            Call cboDept_Click
        End If
    End If
    LockWindowUpdate 0
    Screen.MousePointer = 0
    Exit Sub
errH:
    LockWindowUpdate 0
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim objTab As TabControlItem
    Dim frmTendBody As Object
    Dim intIdx As Integer
     
    picInfo.BackColor = fraLR.BackColor
    fraInfo.BackColor = picInfo.BackColor
    fraIn.BackColor = picInfo.BackColor
    fraOut.BackColor = picInfo.BackColor
    Call DeleteLISTempFile
    
    mstrKey = zlDatabase.GetPara("缺省显示信息", glngSys, 1259, "")
    '初始对象
    '------------------------------------------------------------------------------------------------------------------
    If Not gobjEmr Is Nothing Then
        Set mobjRichEMR = DynamicCreate("zlRichEMR.clsDockContent", "新版病历", False)
        If Not mobjRichEMR Is Nothing Then Call mobjRichEMR.Init(gobjEmr, gcnOracle, glngSys, 0)
    End If
    If mclsArchive Is Nothing Then
        Set mclsArchive = New zlMedRecPage.clsArchive
        Call mclsArchive.InitArchiveMedRec(gcnOracle, glngSys)
    End If
    Set mclsOutAdvices = New clsDockOutAdvices
    Set mclsInAdvices = New clsDockInAdvices
    Set mclsDockAduits = New clsDockAduits
    Set mclsPath = New clsDockPath
    Set mclsTendsNew = New zl9TendFile.clsTendFile
    Call mclsTendsNew.InitTendFile(gcnOracle, glngSys)
    Set frmTendBody = mclsDockAduits.zlGetFormTendBody
    Call zlControl.FormSetCaption(frmTendBody, False, False)
    Call CreateObjectPacs(mobjPublicPACS)
    
    '子窗体
    '-----------------------------------------------------
    Set mcolSubForm = New Collection
    mcolSubForm.Add mclsArchive.zlGetForm(0), "_门诊首页"
    mcolSubForm.Add mclsArchive.zlGetForm(1), "_住院首页"
    mcolSubForm.Add mclsDockAduits.zlGetFormEPR, "_病历信息"
    mcolSubForm.Add mclsOutAdvices.zlGetForm, "_门诊医嘱"
    mcolSubForm.Add mclsInAdvices.zlGetForm, "_住院医嘱"
    mcolSubForm.Add frmTendBody, "_体温记录单"
    mcolSubForm.Add mclsDockAduits.zlGetFormTendFile, "_护理记录单"
    mcolSubForm.Add mclsPath.zlGetForm, "_临床路径"
    mcolSubForm.Add mclsTendsNew.zlGetfrmInTendFile, "_新版护理"
    If Not mobjRichEMR Is Nothing Then mcolSubForm.Add mobjRichEMR.zlGetForm, "_电子病历"
    If Not mobjPublicPACS Is Nothing Then mcolSubForm.Add mobjPublicPACS.zlDocGetForm, "_检查报告"
    
    With tbcArchive
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .Color = xtpTabColorOffice2003
            .Layout = xtpTabLayoutAutoSize
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
        End With
        '隐式出发Form_Load采取添加一个图片方式，切换的时候再依次重新加载
        Set objTab = .InsertItem(intIdx, "门诊首页", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "住院首页", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "病历信息", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "门诊医嘱", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "住院医嘱", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "体温记录单", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "护理记录单", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "临床路径", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "新版护理", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        If Not mobjRichEMR Is Nothing Then
            Set objTab = .InsertItem(intIdx, "电子病历", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
                objTab.Visible = False: intIdx = intIdx + 1
        End If
        Set objTab = .InsertItem(intIdx, "检查报告", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "三方报告", picRpt.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
    End With
    
    '就诊历史
    '-----------------------------------------------------
    With tbcHistory
        With .PaintManager
            .Appearance = xtpTabAppearanceVisio
            .Color = xtpTabColorOffice2003
            .DisableLunaColors = False
            .BoldSelected = True
            .HotTracking = True
            .ShowIcons = True
        End With
        .SetImageList ils16
    End With
    Call InitBasicData
    Call RestoreWinState(Me, App.ProductName)
    If Me.WindowState = vbMinimized Then Me.WindowState = vbNormal
    Timer.Enabled = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    Me.cboDept.Width = tbcHistory.Width
    Me.tbcHistory.Top = cboDept.Height
    Me.tbcHistory.Left = 0
    
    Me.tbcHistory.Height = Me.ScaleHeight - cboDept.Height
    
    Me.fraLR.Top = 0
    Me.fraLR.Left = Me.tbcHistory.Width
    Me.fraLR.Height = Me.ScaleHeight
    
    Me.picInfo.Top = 0
    Me.picInfo.Left = Me.fraLR.Left + Me.fraLR.Width
    Me.picInfo.Width = Me.ScaleWidth - Me.tbcHistory.Width - Me.fraLR.Width
    
    Me.tbcArchive.Left = Me.fraLR.Left + Me.fraLR.Width
    Me.tbcArchive.Top = Me.picInfo.Top + Me.picInfo.Height
    Me.tbcArchive.Width = Me.ScaleWidth - Me.tbcHistory.Width - Me.fraLR.Width
    Me.tbcArchive.Height = Me.ScaleHeight - Me.picInfo.Height
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    Call zlDatabase.SetPara("缺省显示信息", Mid(tvwArchive.Tag, 1, 3), glngSys, 1259)
    
    '强行Unload,不然不会激活子窗体的事件
    For i = 1 To mcolSubForm.Count
        Unload mcolSubForm(i)
    Next
    If picRpt.Tag <> "" Then mstrTempDel = picRpt.Tag
    Set mclsArchive = Nothing
    Set mclsDockAduits = Nothing
    Set mclsOutAdvices = Nothing
    Set mclsInAdvices = Nothing
    Set mclsPath = Nothing
    Set mclsTendsNew = Nothing
    Set mobjRichEMR = Nothing
    Set mrsData = Nothing
    Set mobjPublicPACS = Nothing
    Set mobjPatient = Nothing
End Sub

Private Sub cboDept_Click()

    If cboDept.Text = "" Then Exit Sub
    
    If mlngPreDept = cboDept.ItemData(cboDept.ListIndex) Then Exit Sub
    
    mlngPreDept = cboDept.ItemData(cboDept.ListIndex)
    
    mrsData.Filter = "序号=" & mlngPreDept
    
    mlng就诊ID = mrsData!就诊id
    mlng病人ID = mrsData!病人ID

    If Not mrsData.EOF Then
        mstr挂号单 = Nvl(mrsData!NO, "")
        mblnMoved = Val(Nvl(mrsData!数据转出, "")) = 1
    End If
    '显示基本信息
    If mstr挂号单 <> "" Then
        Call ShowOutPatiInfo
    Else
        Call ShowInPatiInfo
    End If
    
    fraOut.Visible = mstr挂号单 <> ""
    fraIn.Visible = mstr挂号单 = ""

    '显示档案目录
    Me.tbcHistory(0).Caption = cboDept.Text
    Call ShowArchiveTree
    If tvwArchive.Visible And tvwArchive.Enabled Then tvwArchive.SetFocus
    Call Form_Resize
End Sub

Private Sub fraIn_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If img观片.Picture <> img观片普.Picture Then Set img观片.Picture = img观片普.Picture
End Sub

Private Sub fraInfo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If img观片.Picture <> img观片普.Picture Then Set img观片.Picture = img观片普.Picture
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If fraLR.Left + x < 1000 Or fraLR.Left + x > Me.ScaleWidth - 3000 Then Exit Sub
        
        Me.tbcHistory.Width = tbcHistory.Width + x
        Call Form_Resize
    End If
End Sub

Private Sub img观片_Click()
'观片功能
    Dim lng医嘱ID As Long
    
    If Not tvwArchive.SelectedItem Is Nothing Then
        lng医嘱ID = Val(Split(tvwArchive.SelectedItem.Tag, ";")(1) & "")
        If lng医嘱ID <> 0 Then
            If CreateObjectPacs(gobjPublicPacs) Then
                Call gobjPublicPacs.ShowImage(lng医嘱ID, Me, mblnMoved)
            End If
        End If
    End If
End Sub

Private Sub img观片_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If x <= 60 Or x >= 1040 Or y <= 60 Or y >= 300 Then
        If img观片.Picture <> img观片普.Picture Then Set img观片.Picture = img观片普.Picture
    Else
        If img观片.Picture <> img观片高.Picture Then Set img观片.Picture = img观片高.Picture
    End If
    
End Sub

Private Sub picinfo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If img观片.Picture <> img观片普.Picture Then Set img观片.Picture = img观片普.Picture
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    fraInfo.Width = picInfo.ScaleWidth - fraInfo.Left * 3
    
    fraIn.Width = fraInfo.Width - fraIn.Left * 2
    fraOut.Width = fraIn.Width
    lbl急.Left = fraOut.Width - lbl急.Width - 60
End Sub

Private Sub tbcArchive_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'功能：刷新子窗体界面及数据
'说明：仅在人为切换界面卡片激活
    Dim Index As Long, objItem As TabControlItem
    
    If mblnTabTmp Then Exit Sub
    If Item.Tag = "" Then Exit Sub '初始添卡时,还没赋值
    If Item.Handle = picTmp.hwnd Then
        Screen.MousePointer = 11
        Index = Item.Index
        mblnTabTmp = True
        On Error GoTo errH
        Select Case Item.Tag
            Case "门诊首页"
                Set objItem = tbcArchive.InsertItem(Index, "门诊首页", mcolSubForm("_门诊首页").hwnd, 0)
                objItem.Tag = "门诊首页"
            Case "住院首页"
                Set objItem = tbcArchive.InsertItem(Index, "住院首页", mcolSubForm("_住院首页").hwnd, 0)
                objItem.Tag = "住院首页"
            Case "病历信息"
                Set objItem = tbcArchive.InsertItem(Index, "病历信息", mcolSubForm("_病历信息").hwnd, 0)
                objItem.Tag = "病历信息"
            Case "门诊医嘱"
                Set objItem = tbcArchive.InsertItem(Index, "门诊医嘱", mcolSubForm("_门诊医嘱").hwnd, 0)
                objItem.Tag = "门诊医嘱"
            Case "住院医嘱"
                Set objItem = tbcArchive.InsertItem(Index, "住院医嘱", mcolSubForm("_住院医嘱").hwnd, 0)
                objItem.Tag = "住院医嘱"
            Case "体温记录单"
                Set objItem = tbcArchive.InsertItem(Index, "体温记录单", mcolSubForm("_体温记录单").hwnd, 0)
                objItem.Tag = "体温记录单"
            Case "护理记录单"
                Set objItem = tbcArchive.InsertItem(Index, "护理记录单", mcolSubForm("_护理记录单").hwnd, 0)
                objItem.Tag = "护理记录单"
            Case "临床路径"
                Set objItem = tbcArchive.InsertItem(Index, "临床路径", mcolSubForm("_临床路径").hwnd, 0)
                objItem.Tag = "临床路径"
            Case "新版护理"
                Set objItem = tbcArchive.InsertItem(Index, "新版护理", mcolSubForm("_新版护理").hwnd, 0)
                objItem.Tag = "新版护理"
            Case "电子病历"
                Set objItem = tbcArchive.InsertItem(Index, "电子病历", mcolSubForm("_电子病历").hwnd, 0)
                objItem.Tag = "电子病历"
            Case "检查报告"
                Set objItem = tbcArchive.InsertItem(Index, "检查报告", mcolSubForm("_检查报告").hwnd, 0)
                objItem.Tag = "检查报告"
        End Select
        Call tbcArchive.RemoveItem(Index + 1)
        objItem.Selected = True
        mblnTabTmp = False
        Screen.MousePointer = 0
    End If
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowArchiveTab(ByVal strShow As String, ByVal strCaption As String)
'功能：切换显示不同的档案页面，或者清空界面
    Dim i As Long
    
    For i = 0 To tbcArchive.ItemCount - 1
        If tbcArchive(i).Tag = strShow Then
            '默认的卡片跟当前界面要展示的一样时，可能窗体还未绑定上去，这里通过条件判断一下手动绑一次。不会出现多重复执行
            If tbcArchive.Item(i).Handle = picTmp.hwnd Then Call tbcArchive_SelectedChanged(tbcArchive.Item(i))
            tbcArchive(i).Caption = strCaption
            If Not tbcArchive(i).Visible Then
                tbcArchive(i).Visible = True
                tbcArchive(i).Selected = True
                Exit For
            End If
        End If
    Next
    
    For i = 0 To tbcArchive.ItemCount - 1
        If tbcArchive(i).Tag <> strShow Then
            If tbcArchive(i).Visible Then tbcArchive(i).Visible = False
        End If
    Next
End Sub

Private Sub tvwArchive_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If img观片.Picture <> img观片普.Picture Then Set img观片.Picture = img观片普.Picture
End Sub

Private Sub tvwArchive_NodeClick(ByVal node As MSComctlLib.node)
    Dim arrPar As Variant
    Dim intSel As Integer
    Dim strFile As String
        
    If tvwArchive.Tag = node.Key Then Exit Sub
    tvwArchive.Enabled = False
    
    LockWindowUpdate Me.hwnd
    
    arrPar = Split(node.Tag, ";")
    
    If node.Key Like "R1K*" Or node.Key Like "R2K*" Or node.Key Like "R4K*" Or node.Key Like "R5K*" Or node.Key Like "R6K*" Or node.Key Like "R7K*" Then
        Call ShowArchiveTab("病历信息", node.Text)
    End If
    pic观片.Visible = False
    If node.Key = "R11" Then
        Call ShowArchiveTab(IIf(mstr挂号单 <> "", "门诊首页", "住院首页"), tbcHistory.Selected.Caption)
        Call mclsArchive.zlRefresh(IIf(mstr挂号单 <> "", 0, 1), mlng病人ID, mlng就诊ID, mblnMoved)
    ElseIf node.Key = "R12" Then '医嘱记录
        If mstr挂号单 <> "" Then
            Call ShowArchiveTab("门诊医嘱", tbcHistory.Selected.Caption)
            Call mclsOutAdvices.zlRefresh(mlng病人ID, mstr挂号单, False, mblnMoved)
        Else
            Call ShowArchiveTab("住院医嘱", tbcHistory.Selected.Caption)
            Call mclsInAdvices.zlRefresh(mlng病人ID, mlng就诊ID, mlng病区ID, mlng科室ID, 0, mblnMoved)
        End If
    ElseIf node.Key Like "R1K*" Then '门诊病历
        Call mclsDockAduits.zlRefresh(1, Val(arrPar(0)), , , , , , , mblnMoved)
    ElseIf node.Key Like "R2K*" Then '住院病历
        Call mclsDockAduits.zlRefresh(2, Val(arrPar(0)), , , , , , , mblnMoved)
    ElseIf node.Key Like "R3K*" Then '护理记录
        If UBound(arrPar) >= 1 Then
            If mblnNewTends = False Then
                If Val(arrPar(1)) = -1 Then
                    Call ShowArchiveTab("体温记录单", node.Text)
                    Call mclsDockAduits.zlRefreshTendBody(mlng病人ID, mlng就诊ID, Val(arrPar(0)), 0)
                Else
                    Call ShowArchiveTab("护理记录单", node.Text)
                    Call mclsDockAduits.zlRefresh(3, Val(arrPar(3)), mlng病人ID, mlng就诊ID, Val(arrPar(0)), CStr(arrPar(2)))
                End If
            Else
                Select Case Val(arrPar(1))
                    Case -1
                        intSel = 0
                    Case 1
                        intSel = 2
                    Case Else
                        intSel = 1
                End Select
                Call ShowArchiveTab("新版护理", node.Text)
                Call mclsTendsNew.zlRefreshTendFile(mlng病人ID, mlng就诊ID, Val(arrPar(4)), Val(arrPar(0)), False, IIf(glngModul = p住院医生站, True, False), intSel, Val(arrPar(3)), 1)
            End If
        End If
    ElseIf node.Key Like "R4K*" Then '护理病历
        Call mclsDockAduits.zlRefresh(4, Val(arrPar(0)), , , , , , , mblnMoved)
    ElseIf node.Key Like "R5K*" Then '疾病证明
        Call mclsDockAduits.zlRefresh(5, Val(arrPar(0)), , , , , , , mblnMoved)
    ElseIf node.Key Like "R6K*" Then '知情文件
        Call mclsDockAduits.zlRefresh(6, Val(arrPar(0)), , , , , , , mblnMoved)
    ElseIf node.Key Like "R7K*" Then '诊疗报告
        Call mclsDockAduits.zlRefresh(7, Val(arrPar(0)), , , , , , , mblnMoved)
        If arrPar(2) = "D" Then
            pic观片.Visible = True
        Else
            pic观片.Visible = False
        End If
    ElseIf node.Key = "R8" Then
        If mstr挂号单 = "" Then
            Call ShowArchiveTab("临床路径", node.Text)
            Call mclsPath.zlRefreshReadOnly(mlng病人ID, mlng就诊ID)
        End If
    ElseIf node.Key Like "R7P*" Then  '检查报告
        pic观片.Visible = True
        Call ShowArchiveTab("检查报告", node.Text)
        If Not mobjPublicPACS Is Nothing Then Call mobjPublicPACS.zlDocRefresh(Split(node.Tag, ";")(0))
    ElseIf node.Key Like "R7L*" Then  '三方报告
        strFile = GetLisRptFile(node.Tag)
        If strFile <> "" Then
            If picRpt.Tag <> "" And picRpt.Tag <> mstrTempDel And picRpt.Tag <> strFile Then mstrTempDel = picRpt.Tag
            webRpt.Navigate strFile
            picRpt.Tag = strFile
        End If
        Call ShowArchiveTab("三方报告", node.Text)
    ElseIf InStr(node.Key, "R") = 0 And Len(node.Tag) >= 32 Then
        'EMR病历预览
        If Not mobjRichEMR Is Nothing Then
            Call ShowArchiveTab("电子病历", node.Text)
            If InStr(node.Tag, "|") > 0 Then
                Call mobjRichEMR.zlShowDoc(Split(node.Tag, "|")(0), Split(node.Tag, "|")(1))
            Else
                Call mobjRichEMR.zlShowDoc(node.Tag, "")
            End If
        End If
    Else
        LockWindowUpdate 0
        tvwArchive.Enabled = True
        Exit Sub
    End If
    tvwArchive.Tag = node.Key
    LockWindowUpdate 0
    tvwArchive.Enabled = True
    If tvwArchive.Visible And tvwArchive.Enabled Then tvwArchive.SetFocus
End Sub

Private Function ShowArchiveTree() As Boolean
'功能：显示病人档案树形目录
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, objNode As node, strSQL1 As String
    Dim blnPath As Boolean
    Dim strSel As String
    Dim strRptIDs As String

    Screen.MousePointer = 11
    
    If Not tvwArchive.SelectedItem Is Nothing Then
        If tvwArchive.SelectedItem.Key = "R11" Or tvwArchive.SelectedItem.Key = "R12" Then
            strSel = Split(tvwArchive.SelectedItem.Key, "K")(0)
        End If
    End If
    
    '病人科室存在可用的临床路径时，显示临床路径记录
    If mstr挂号单 = "" Then
        If GetInsidePrivs(p临床路径应用) <> "" Then
            blnPath = HavePath(mlng科室ID)
        End If
    End If
    
    On Error GoTo errH
    '1-门诊病历;2-住院病历;3-护理记录;4-护理病历;5-疾病证明;6-知情文件;7-诊疗报告,11-首页信息,12-医嘱记录,13-临床路径
    strSQL = _
        " Select * From (" & _
            " Select 'R11' As ID, '' As 上级id, '首页信息' As 名称, '' As 参数,1 As 末级,'object_first' As 图标,'01' As 排序 From Dual Union All" & _
            " Select 'R12' As ID, '' As 上级id, '医嘱记录' As 名称, '' As 参数,1 As 末级,'object_advice' As 图标,'02' As 排序 From Dual Union All" & _
            " Select 'R1' As ID, '' As 上级id, '门诊病历' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,'03' As 排序 From Dual Where [3]=0 Union All" & _
            " Select 'R2' As ID, '' As 上级id, '住院病历' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,'04' As 排序 From Dual Where [3]=1 Union All" & _
            " Select 'R3' As ID, '' As 上级id, '护理记录' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,'05' As 排序 From Dual Where [3]=1 Union All" & _
            " Select 'R4' As ID, '' As 上级id, '护理病历' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,'06' As 排序 From Dual Where [3]=1 Union All" & _
            " Select 'R7' As ID, '' As 上级id, '诊疗报告' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,'07' As 排序 From Dual Union All" & _
            " Select 'R5' As ID, '' As 上级id, '疾病证明' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,'08' As 排序 From Dual Union All" & _
            " Select 'R6' As ID, '' As 上级id, '知情文件' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,'09' As 排序 From Dual " & _
            IIf(blnPath, " Union All Select 'R8' As ID, '' As 上级id, '临床路径' As 名称, '' As 参数,0 As 末级,'Path' As 图标,'10' As 排序 From Dual", "")
    '病历部分
    'ID=上级ID+K病历ID,医嘱ID,0
    '参数=病历ID;医嘱ID
    strSQL = strSQL & " Union All" & _
        " Select A.上级id||'K'||Trim(To_Char(A.ID))||','||Trim(To_Char(Nvl(A.医嘱id,0)))||',0' As ID,A.上级id," & _
        "       Decode(A.医嘱id,Null,A.名称||'('||To_Char(A.创建时间, 'YYYY-MM-DD')||')',A.名称||'：'||B.医嘱内容||'('||To_Char(A.创建时间, 'YYYY-MM-DD')||')') As 名称," & _
        "       Trim(To_Char(A.ID))||';'||Decode(A.医嘱id,Null,'0',Trim(To_Char(A.医嘱id))) || ';'|| B.诊疗类别 As 参数," & _
        "       1 As 末级,Decode(病历种类,1,'object_case',2,'object_case',4,'object_case',7,'object_report','object_file') As 图标,排序 " & _
        " From (Select A.ID, 'R'||A.病历种类 As 上级id, A.病历名称 As 名称,C.医嘱id,A.病历种类,A.创建时间,To_Char(A.创建时间,'YYYY-MM-DD HH24:MI:SS') As 排序" & _
        "       From 电子病历记录 A,病人医嘱报告 C " & _
        "       Where A.病人id = [1] And A.主页id = [2] And (A.病人来源=2 And [3]=1 Or Nvl(A.病人来源,0)<>2 And [3]=0)" & _
        "           And C.病历id(+)=A.ID And A.病历种类 In (1, 2, 3, 4, 5, 6, 7)" & _
        "       ) A,病人医嘱记录 B Where A.医嘱id=B.Id(+)"
    '护理部分
    'ID=上级ID+K文件ID,0,科室ID
    '参数=科室ID;保留;开始～截止;文件ID
    '检查本次病人是使用的是老板还是新版
    strSQL1 = "Select 1 From 病人护理记录 A Where a.病人id = [1] And a.主页id = [2]"
	If mblnMoved Then
        strSQL1 = Replace(strSQL1, "病人护理记录", "H病人护理记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL1, "检查是否存在老板数据", mlng病人ID, mlng就诊ID)
    If rsTmp.RecordCount > 0 Then
        mblnNewTends = False
        strSQL = strSQL & " Union All" & _
            " Select 'R3K'||Trim(To_Char(A.ID))||',0,'||Trim(To_Char(A.科室Id)) As ID,'R3' As 上级id," & _
            "       A.名称||'('||B.名称||'：'||To_Char(A.开始, 'YYYY-MM-DD HH24:MI') || '～' ||To_Char(A.截止, 'YYYY-MM-DD HH24:MI') || ')' As 名称," & _
            "       Trim(To_Char(B.ID))||';'||Trim(To_Char(Nvl(保留,0)))||';'||To_Char(A.开始, 'YYYY-MM-DD HH24:MI')||'～'||To_Char(A.截止, 'YYYY-MM-DD HH24:MI')||';'||Trim(To_Char(A.ID)) As 参数," & _
            "       1 As 末级,'object_tend' As 图标,To_Char(a.开始,'YYYY-MM-DD HH24:MI:SS') As 排序" & _
            " From (" & _
            "   Select F.ID, F.编号, F.名称, R.开始, R.截止, R.科室id, 保留" & _
            "   From (" & _
            "       Select ID, 编号, 名称, 3 As 护理级别, 通用, 0 As 科室id, 保留" & _
            "          From 病历文件列表 Where 种类 = 3 And 保留 < 0 And NVL(子类,0)=0 " & _
            "       Union All" & _
            "       Select L.ID, L.编号, L.名称, F.报表 As 护理级别, L.通用, A.科室id, L.保留" & _
            "          From 病历页面格式 F, 病历文件列表 L, 病历应用科室 A" & _
            "          Where L.种类 = 3 And L.保留 = 0 And L.种类 = F.种类 And L.编号 = F.编号 And L.ID = A.文件id(+)" & _
            "       ) F,(" & _
            "       Select R.科室id, Nvl(Min(R.护理级别), 3) As 护理级别, Min(R.发生时间) As 开始, Max(R.发生时间) As 截止" & _
            "          From 病人护理记录 R" & _
            "          Where R.病人来源 = 2 And R.病人id = [1] And Nvl(R.主页id, 0) = [2] And Nvl(R.婴儿, 0) = 0" & _
            "          Group By R.科室id" & _
            "       ) R" & _
            "       Where (F.通用 = 1 Or F.通用 = 2 And F.科室id = R.科室id) And F.护理级别 >= R.护理级别" & _
            "   ) A, 部门表 B Where A.科室id = B.ID)" & _
            "Order By Decode(上级id,Null,' ',上级id),排序"
    Else
        mblnNewTends = True
        strSQL = strSQL & " Union All" & _
                " Select 'R3K'||Trim(To_Char(A.ID))||',0,'||Trim(To_Char(A.科室Id)) As ID,'R3' As 上级id," & vbNewLine & _
                "     A.名称||'('||B.名称||'：'||To_Char(A.开始, 'YYYY-MM-DD HH24:MI') || '～' ||To_Char(A.截止, 'YYYY-MM-DD HH24:MI') || ')' As 名称," & vbNewLine & _
                "      Trim(To_Char(B.ID))||';'||Trim(To_Char(Nvl(保留,0)))||';'||To_Char(A.开始, 'YYYY-MM-DD HH24:MI')||'～'||To_Char(A.截止, 'YYYY-MM-DD HH24:MI')||';'||Trim(To_Char(A.ID))||';'||Trim(To_Char(A.婴儿)) As 参数," & vbNewLine & _
                "       1 As 末级,'object_tend' As 图标,To_Char(a.开始,'YYYY-MM-DD HH24:MI:SS') As 排序" & vbNewLine & _
                " From (" & vbNewLine & _
                "   Select R.ID, F.编号, R.名称,R.婴儿, R.开始, NVL(R.截止,nvl(R.时间,R.开始)) 截止, R.科室id, 保留" & vbNewLine & _
                "   From (" & vbNewLine & _
                "       Select L.ID, L.编号, L.名称, F.报表 As 护理级别, L.通用, L.保留" & vbNewLine & _
                "          From 病历页面格式 F, 病历文件列表 L" & vbNewLine & _
                "          Where L.种类 = 3 And L.种类 = F.种类 And L.编号 = F.编号 And (L.通用=1 OR L.通用=2)" & vbNewLine & _
                "" & vbNewLine & _
                "       ) F,(" & vbNewLine & _
                "       Select R.ID,R.科室id,R.文件名称 名称,R.格式ID,nvl(R.婴儿,0) 婴儿,Min(R.开始时间) As 开始, Max(R.结束时间) As 截止,MAX(T.发生时间) 时间" & vbNewLine & _
                "          From 病人护理文件 R,病人护理数据 T" & vbNewLine & _
                "          Where R.ID=T.文件ID(+) And R.病人id = [1] And Nvl(R.主页id, 0) = [2]" & vbNewLine & _
                "          Group By R.ID,R.文件名称,R.科室id,R.格式ID,R.婴儿" & vbNewLine & _
                "       ) R" & vbNewLine & _
                "       Where F.ID=R.格式ID" & vbNewLine & _
                "   ) A, 部门表 B Where A.科室id = B.ID And DECODE(A.保留,-1,0,A.婴儿)=A.婴儿)" & vbNewLine & _
                " Order By Decode(上级id,Null,' ',上级id),排序"
    End If
    If mblnMoved Then
        strSQL = Replace(strSQL, "电子病历记录", "H电子病历记录")
        strSQL = Replace(strSQL, "病人护理记录", "H病人护理记录")
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱报告", "H病人医嘱报告")
        strSQL = Replace(strSQL, "病人护理文件", "H病人护理文件")
        strSQL = Replace(strSQL, "病人护理数据", "H病人护理数据")
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng就诊ID, IIf(mstr挂号单 = "", 1, 0))
    
    tvwArchive.Tag = ""
    tvwArchive.Nodes.Clear
            
    Do While Not rsTmp.EOF
        If Nvl(rsTmp!上级ID) = "" Then
            Set objNode = tvwArchive.Nodes.Add(, , CStr(rsTmp!ID), rsTmp!名称, Nvl(rsTmp!图标))
        Else
            Set objNode = tvwArchive.Nodes.Add(CStr(rsTmp!上级ID), tvwChild, CStr(rsTmp!ID), rsTmp!名称, Nvl(rsTmp!图标))
        End If
        
        objNode.Tag = Nvl(rsTmp!参数)
        objNode.Expanded = True
        
        If tvwArchive.Nodes.Count = 1 Then
            objNode.Selected = True
        ElseIf objNode.Key = strSel Then
            objNode.Selected = True
        ElseIf mstrKey <> "" Then
            If InStr(mstrKey, "K") > 0 Then
                If mstrKey = "R1K" Or mstrKey = "R2K" Then
                    If rsTmp!上级ID & "" = "R1" Or rsTmp!上级ID & "" = "R2" Then objNode.Selected = True: mstrKey = ""
                Else
                    If rsTmp!上级ID & "" = Mid(mstrKey, 1, 2) Then objNode.Selected = True: mstrKey = ""
                End If
            Else
                If objNode.Key = mstrKey Then objNode.Selected = True: mstrKey = ""
            End If
        End If
        
        rsTmp.MoveNext
    Loop
    
    Set rsTmp = Nothing
    Set rsTmp = GetEmrCISStruct(mlng病人ID, mlng就诊ID)
    
    If Not rsTmp Is Nothing Then
        If rsTmp.State = ADODB.adStateOpen Then
            If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                Do Until rsTmp.EOF
		If InStr("," & strRptIDs & ",", "," & rsTmp!ID.Value & ",") = 0 Then
                    Set objNode = tvwArchive.Nodes.Add(rsTmp!上级ID.Value, tvwChild, rsTmp!ID.Value, rsTmp!名称.Value, rsTmp!图标.Value, rsTmp!图标.Value)
                    objNode.Tag = Nvl(rsTmp!参数) '文档ID[|子文档ID]
                    If mstrKey <> "" Then
                        If rsTmp!上级ID & "" = Mid(mstrKey, 1, 2) Then objNode.Selected = True: mstrKey = ""
                    End If
		 strRptIDs = strRptIDs & "," & rsTmp!ID.Value
                    End If
                    rsTmp.MoveNext
                Loop
            End If
        End If
    End If
    
    If Not mobjPublicPACS Is Nothing Then
        Set rsTmp = Nothing
        Set rsTmp = mobjPublicPACS.zlDocGetList(mlng病人ID, mlng就诊ID, mstr挂号单)
        
        If Not rsTmp Is Nothing Then
            Do Until rsTmp.EOF
                Set objNode = tvwArchive.Nodes.Add("R7", tvwChild, "R7P" & rsTmp!报告ID & "", rsTmp!文档标题 & "", "object_report", "object_report")
                objNode.Tag = rsTmp!报告ID & ";" & rsTmp!医嘱ID
                If mstrKey <> "" Then
                   If Mid(mstrKey, 1, 2)="R7" Then objNode.Selected = True: mstrKey = ""
                End If
                rsTmp.MoveNext
            Loop
        End If
    End If
    
    '三方LIS报告
    If mstr挂号单 = "" Then
        strSQL = "select b.id as 报告ID,b.报告名 as 文档标题,c.医嘱ID,b.类型 from 病人医嘱记录 a, 医嘱报告内容 b,病人医嘱报告 c where b.id=c.报告id and a.id=c.医嘱id and c.报告id is not null and a.病人id=[1] and a.主页id=[2]"
    Else
        strSQL = "select b.id as 报告ID,b.报告名 as 文档标题,c.医嘱ID,b.类型 from 病人医嘱记录 a, 医嘱报告内容 b,病人医嘱报告 c where b.id=c.报告id and a.id=c.医嘱id and c.报告id is not null and a.挂号单=[3]"
    End If
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱报告", "H病人医嘱报告")
    End If
  
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng就诊ID, mstr挂号单)
 strRptIDs = ""
    If Not rsTmp Is Nothing Then
        Do Until rsTmp.EOF
	If InStr("," & strRptIDs & ",", "," & rsTmp!报告ID & ",") = 0 Then
            Set objNode = tvwArchive.Nodes.Add("R7", tvwChild, "R7L" & rsTmp!报告ID & "", rsTmp!文档标题 & "", "object_report", "object_report")
            objNode.Tag = rsTmp!报告ID & ";" & rsTmp!医嘱ID & ";" & rsTmp!类型 & "<sTab>" & rsTmp!文档标题
            If mstrKey <> "" Then
               If Mid(mstrKey, 1, 2)="R7" Then objNode.Selected = True: mstrKey = ""
            End If
	    strRptIDs = strRptIDs & "," & rsTmp!报告ID
            End If
            rsTmp.MoveNext
        Loop
    End If
    
    If Not tvwArchive.SelectedItem Is Nothing Then
        tvwArchive.SelectedItem.EnsureVisible
        Call tvwArchive_NodeClick(tvwArchive.SelectedItem)
    End If

    mstrKey = ""
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowOutPatiInfo() As Boolean
'功能：选择门诊病人某次历史就诊记录时，读取相关的病人信息
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    If mlng病人ID <> 0 Then
        strSQL = "Select B.Id,B.NO,B.门诊号,B.姓名,B.性别,B.年龄,A.医疗付款方式," & _
            " A.费别,A.险类,A.医保号,B.急诊,B.发生时间,B.执行人,B.执行状态,B.执行时间," & _
            " B.执行部门ID as 科室ID,B.诊室,B.社区,D.社区号,C.名称 as 科室" & _
            " From 病人信息 A,病人挂号记录 B,部门表 C,病人社区信息 D" & _
            " Where A.病人ID=B.病人ID And B.ID=[1] And B.执行部门ID=C.ID" & _
            " And B.病人ID=D.病人ID(+) And B.社区=D.社区(+) And B.记录性质=1 And B.记录状态=1"
        If mblnMoved Then
            strSQL = Replace(strSQL, "病人挂号记录", "H病人挂号记录")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng就诊ID)
        With rsTmp
            '保险病人姓名红色显示
            lbl姓名mz(1).Caption = Nvl(!姓名)
            If Not IsNull(!险类) Then
                lbl姓名mz(1).ForeColor = vbRed
            Else
                lbl姓名mz(1).ForeColor = lbl门诊号mz(1).ForeColor
            End If
            lbl医生mz(1).Caption = Nvl(!执行人)
            lbl挂号单mz(1).Caption = !NO
            lbl门诊号mz(1).Caption = Nvl(!门诊号)
            lbl付款mz(1).Caption = Nvl(!医疗付款方式)
            lbl费别mz(1).Caption = Nvl(!费别)
            lbl医保号mz(1).Caption = Nvl(!医保号)
            lbl社区号mz(1).Caption = Nvl(!社区号)
            lbl急.Visible = Nvl(!急诊, 0) <> 0
            
            mlng科室ID = Nvl(!科室ID, 0)
            mlng病区ID = 0
        End With
    Else
        fraOut.Visible = True
        lbl姓名mz(1).Caption = ""
        lbl医生mz(1).Caption = ""
        lbl挂号单mz(1).Caption = ""
        lbl门诊号mz(1).Caption = ""
        lbl付款mz(1).Caption = ""
        lbl费别mz(1).Caption = ""
        lbl医保号mz(1).Caption = ""
        lbl社区号mz(1).Caption = ""
    End If
    ShowOutPatiInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowInPatiInfo() As Boolean
'功能：选择某次住院记录时，读取相关的病人信息
'返回：blnMoved=本次住院病案是否转出了
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    If mlng病人ID <> 0 Then
        strSQL = "Select NVL(B.姓名,A.姓名) 姓名, NVL(B.性别,A.性别) 性别, NVL(B.年龄,A.年龄) 年龄,B.住院号,B.出院病床,B.医疗付款方式," & _
            " D.信息值 as 医保号,B.险类,B.当前病况,C.名称 as 护理等级,B.入院日期," & _
            " B.出院日期,B.病人类型,B.状态,B.出院科室ID,B.当前病区ID,A.住院次数" & _
            " From 病人信息 A,病案主页 B,收费项目目录 C,病案主页从表 D" & _
            " Where A.病人ID=B.病人ID And A.病人ID=[1] And B.主页ID=[2] And B.护理等级ID=C.ID(+)" & _
            " And B.病人ID=D.病人ID(+) And B.主页ID=D.主页ID(+) And D.信息名(+)='医保号'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng就诊ID)
        
        With rsTmp
            '保险病人颜色特殊显示
            lbl姓名zy(1).Caption = Nvl(!姓名)
            lbl姓名zy(1).ForeColor = zlDatabase.GetPatiColor(Nvl(!病人类型))
            
            lbl住院号zy(1).Caption = Nvl(!住院号)
            lbl床号zy(1).Caption = Nvl(!出院病床)
            lbl医保号zy(1).Caption = Nvl(!医保号)
            lbl护理zy(1).Caption = Nvl(!护理等级)
            lbl付款zy(1).Caption = Nvl(!医疗付款方式)
            
            '危重病人病况红色显示
            lbl病况zy(1).Caption = Nvl(!当前病况)
            If Nvl(!当前病况) = "危" Or Nvl(!当前病况) = "重" Or Nvl(!当前病况) = "急" Then
                lbl病况zy(1).ForeColor = vbRed
            Else
                lbl病况zy(1).ForeColor = lbl住院号zy(1).ForeColor
            End If
            
            lbl入院zy(1).Caption = Format(!入院日期, "yyyy-MM-dd HH:mm")
            If Not IsNull(!出院日期) Then
                lbl入院zy(1).Caption = lbl入院zy(1).Caption & "～" & Format(!出院日期, "yyyy-MM-dd HH:mm")
            End If
            
            lbl类型zy(1).Caption = Nvl(!病人类型)
            
            mlng科室ID = Nvl(!出院科室ID, 0)
            mlng病区ID = Nvl(!当前病区ID, 0)
        End With
    Else
        '保险病人颜色特殊显示
        fraIn.Visible = True
        lbl姓名zy(1).Caption = ""
        lbl住院号zy(1).Caption = ""
        lbl床号zy(1).Caption = ""
        lbl医保号zy(1).Caption = ""
        lbl护理zy(1).Caption = ""
        lbl付款zy(1).Caption = ""
        lbl病况zy(1).Caption = ""
        lbl入院zy(1).Caption = ""
        lbl类型zy(1).Caption = ""
    End If
    ShowInPatiInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetEmrCISStruct(ByVal lngPatiID As Long, ByVal lngPageID As Long) As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strExtendTag As String, strReturn As String, strSql As String, strSQLNew As String
    
    On Error GoTo errH
    If gobjEmr Is Nothing Then Set GetEmrCISStruct = Nothing: Exit Function
    strExtendTag = GetEMRIn_Tag(lngPatiID, lngPageID)
    If strExtendTag = "" Then Set GetEmrCISStruct = Nothing: Exit Function
    
    '上级ID，ID，名称，参数，图标
    strSql = "Select Decode(e.Kind, '01', 'R1', '02', 'R2', '03', 'R4', '04', 'R5', '05', 'R6', 'R2') 上级id," & vbNewLine & _
            "       Nvl(d.Subdoc_Id, Rawtohex(d.Real_Doc_Id)) As ID, d.Subdoc_Id As 子文档id," & vbNewLine & _
            "       e.Title ||" & vbNewLine & _
            "        Decode(d.Completor, Null, ''," & vbNewLine & _
            "               '【 ' || d.Completor || ' 在' || To_Char(d.Complete_Time, 'yyyy-mm-dd hh24:mi') || '签名】') As 名称," & vbNewLine & _
            "       Rawtohex(d.Real_Doc_Id) || Decode(d.Subdoc_Id, Null, Null, '|' || d.Subdoc_Id) As 参数, 'object_case' As 图标" & vbNewLine & _
            "From (Select Distinct d.Id, c.Antetype_Id, c.Subdoc_Id, c.Subdoc_Title, c.Real_Doc_Id, c.Complete_Time, c.Completor" & vbNewLine & _
            "       From Bz_Act_Log A, Bz_Act_Log D, Bz_Doc_Tasks C" & vbNewLine & _
            "       Where a.Extend_Tag = :etag And (a.Id = d.Id Or a.Id = d.Basiclog_Id) And d.Id = c.Actlog_Id And" & vbNewLine & _
            "             c.Real_Doc_Id Is Not Null) D, Antetype_List E" & vbNewLine & _
            "Where d.Antetype_Id = e.Id  And e.Title = Decode(e.Type, 3, d.Subdoc_Title, e.Title)" & vbNewLine & _
            "Order By Rawtohex(d.Real_Doc_Id), e.Code, d.Complete_Time"
            
    strSQLNew = "Select Decode(e.Kind, '01', 'R1', '02', 'R2', '03', 'R4', '04', 'R5', '05', 'R6', 'R2') 上级id," & vbNewLine & _
                "       Nvl(d.Subdoc_Id, Rawtohex(d.Real_Doc_Id)) As ID, d.Subdoc_Id As 子文档id," & vbNewLine & _
                "       e.Title ||" & vbNewLine & _
                "        Decode(d.Completor, Null, ''," & vbNewLine & _
                "               '【 ' || d.Completor || ' 在' || To_Char(d.Complete_Time, 'yyyy-mm-dd hh24:mi') || '签名】') As 名称," & vbNewLine & _
                "       Rawtohex(d.Real_Doc_Id) || Decode(d.Subdoc_Id, Null, Null, '|' || d.Subdoc_Id) As 参数, 'object_case' As 图标" & vbNewLine & _
                "From (Select Distinct d.Id, c.Antetype_Id, c.Subdoc_Id, c.Subdoc_Title, c.Real_Doc_Id, c.Complete_Time, c.Completor, c.Order_No" & vbNewLine & _
                "       From Bz_Act_Log A, Bz_Act_Log D, Bz_Doc_Tasks C" & vbNewLine & _
                "       Where a.Extend_Tag = :etag And (a.Id = d.Id Or a.Id = d.Basiclog_Id) And d.Id = c.Actlog_Id And" & vbNewLine & _
                "             c.Real_Doc_Id Is Not Null And Nvl(c.Intead, 0) = 0) D, Antetype_List E" & vbNewLine & _
                "Where d.Antetype_Id = e.Id " & vbNewLine & _
                "Order By Rawtohex(d.Real_Doc_Id), e.Code, d.Order_No"
    
    err.Clear
    On Error Resume Next
    strReturn = gobjEmr.OpenSQLRecordset(strSQLNew, strExtendTag & "^16^etag", rsTemp)
    If err.Number <> 0 Or strReturn <> "" Then
        err.Clear
        strReturn = gobjEmr.OpenSQLRecordset(strSql, strExtendTag & "^16^etag", rsTemp)
    End If
    
    If strReturn <> "" Then
        MsgBox strReturn, vbCritical, gstrSysName
        Set GetEmrCISStruct = Nothing: Exit Function
    End If
    
    Set GetEmrCISStruct = rsTemp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetEMRIn_Tag(ByVal lngPatiID As Long, ByVal lngPageID As Long) As String
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    If InStr(cboDept.Text, "门诊就诊") > 0 Then
        GetEMRIn_Tag = "MZ_" & mlng就诊ID
    Else
        strSQL = "Select Nvl(a.Id, b.Id) ID" & vbNewLine & _
                    "From (Select Max(ID) ID From 病人变动记录 Where 病人id = [1] And 主页id = [2] And 开始原因 = 2 And Nvl(附加床位, 0) = 0) A," & vbNewLine & _
                    "     (Select Max(ID) ID From 病人变动记录 Where 病人id = [1] And 主页id = [2] And 开始原因 = 1 And Nvl(附加床位, 0) = 0) B"
        
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取病人入院ID", lngPatiID, lngPageID)
        
        If rsTmp Is Nothing Then Exit Function
        If Nvl(rsTmp!ID) = "" Then Exit Function
        GetEMRIn_Tag = "BD_" & rsTmp!ID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetLisRptFile(ByVal strTag As String) As String
'功能：打开LIS报告文件查看，获取临时文件路径
    Dim strFile As String
    Dim lngRetu As Long, strInfo As String
    Dim objFile As New FileSystemObject
    Dim strTmp As String
    Dim lng报告ID As String
    Dim str报告名 As String
    Dim lng类型 As String
    Dim varTmp As Variant
    Dim strSuffix As String '文件后缀名
    
    Screen.MousePointer = 11
    
    varTmp = Split(strTag, ";")
    lng报告ID = varTmp(0)
    strTmp = Replace(strTag, varTmp(0) & ";" & varTmp(1) & ";", "")
    varTmp = Split(strTmp, "<sTab>")
    lng类型 = varTmp(0)
    If lng类型 = 0 Then
        strSuffix = "pdf"
    ElseIf lng类型 = 1 Then
        strSuffix = "html"
    Else
        strSuffix = "xps"
    End If
    str报告名 = varTmp(1)
    
    strFile = objFile.GetSpecialFolder(TemporaryFolder) & "\tmpReport_" & lng报告ID & "." & strSuffix
    If Not objFile.FileExists(strFile) Then
        strFile = Sys.ReadLob(glngSys, 22, lng报告ID, strFile)
        If Not objFile.FileExists(strFile) Then
            MsgBox "文件内容读取失败！", vbInformation, gstrSysName:
            Screen.MousePointer = 0: Exit Function
        End If
    End If
    GetLisRptFile = strFile
    Screen.MousePointer = 0
End Function

Private Sub picRpt_Resize()
    On Error Resume Next
    webRpt.Move 0, 0, picRpt.Width, picRpt.Height
End Sub

Private Sub Timer_Timer()
    tbcHistory.Width = tbcHistory.Width + 50
    Call Form_Resize
    Timer.Enabled = False
End Sub

Private Function DeleteLISTempFile() As Boolean
    Dim objFile As New FileSystemObject
    Dim i As Long
    If mstrTempDel = "" Then Exit Function
    If objFile.FileExists(mstrTempDel) Then
        Do While i < 1000
            On Error Resume Next
            objFile.DeleteFile mstrTempDel, True
            If err.Number = 0 Then
                mstrTempDel = ""
                Exit Do
            End If
            err.Clear: On Error GoTo 0
        Loop
    End If
End Function
