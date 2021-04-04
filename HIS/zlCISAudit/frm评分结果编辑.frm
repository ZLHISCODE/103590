VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm评分结果编辑 
   Caption         =   "评分结果编辑"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11520
   FillColor       =   &H000000FF&
   Icon            =   "frm评分结果编辑.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   11520
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdViewArchive 
      Caption         =   "电子病案(&V)"
      Height          =   350
      Left            =   1305
      TabIndex        =   43
      Top             =   7380
      Width           =   1155
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   120
      TabIndex        =   42
      Top             =   7380
      Width           =   1100
   End
   Begin VB.CommandButton cmdAuto 
      Caption         =   "自动(&A)"
      Height          =   350
      Left            =   7260
      TabIndex        =   41
      Top             =   7380
      Width           =   1100
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "终止(&S)"
      Height          =   350
      Left            =   9750
      TabIndex        =   40
      Top             =   7935
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.PictureBox picLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   7080
      Left            =   0
      ScaleHeight     =   7020
      ScaleWidth      =   3060
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   3120
      Begin VB.PictureBox pic项目信息 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2820
         Left            =   135
         ScaleHeight     =   2820
         ScaleWidth      =   2790
         TabIndex        =   28
         Top             =   4050
         Width           =   2790
         Begin VB.PictureBox imgXMXX 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2460
            Picture         =   "frm评分结果编辑.frx":000C
            ScaleHeight     =   210
            ScaleWidth      =   255
            TabIndex        =   39
            Top             =   80
            Width           =   255
         End
         Begin VB.TextBox txt项目信息 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   2175
            Left            =   135
            MultiLine       =   -1  'True
            TabIndex        =   30
            Top             =   450
            Width           =   2490
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "项目信息"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   225
            TabIndex        =   29
            Top             =   90
            Width           =   1095
         End
      End
      Begin VB.PictureBox pic方案信息 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   135
         ScaleHeight     =   1695
         ScaleWidth      =   2790
         TabIndex        =   21
         Top             =   2220
         Width           =   2790
         Begin VB.PictureBox imgFAXX 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   2460
            Picture         =   "frm评分结果编辑.frx":005B
            ScaleHeight     =   225
            ScaleWidth      =   255
            TabIndex        =   38
            Top             =   80
            Width           =   255
         End
         Begin VB.Label lbl分制 
            BackStyle       =   0  'Transparent
            Caption         =   "分制:"
            Height          =   195
            Left            =   225
            TabIndex        =   27
            Top             =   682
            Width           =   2580
         End
         Begin VB.Label lbl总分 
            BackStyle       =   0  'Transparent
            Caption         =   "总分:"
            Height          =   195
            Left            =   225
            TabIndex        =   26
            Top             =   914
            Width           =   2580
         End
         Begin VB.Label lbl下值 
            BackStyle       =   0  'Transparent
            Caption         =   "下值:"
            Height          =   195
            Left            =   225
            TabIndex        =   25
            Top             =   1380
            Width           =   2580
         End
         Begin VB.Label lbl上值 
            BackStyle       =   0  'Transparent
            Caption         =   "上值:"
            Height          =   195
            Left            =   225
            TabIndex        =   24
            Top             =   1146
            Width           =   2580
         End
         Begin VB.Label lbl方案名称 
            BackStyle       =   0  'Transparent
            Caption         =   "方案名称"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   225
            TabIndex        =   23
            Top             =   450
            Width           =   2580
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "方案信息"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   225
            TabIndex        =   22
            Top             =   90
            Width           =   1095
         End
      End
      Begin VB.PictureBox pic病人信息 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1950
         Left            =   135
         ScaleHeight     =   1950
         ScaleWidth      =   2790
         TabIndex        =   13
         Top             =   135
         Width           =   2790
         Begin VB.PictureBox imgBRXX 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   2460
            Picture         =   "frm评分结果编辑.frx":00AA
            ScaleHeight     =   225
            ScaleWidth      =   255
            TabIndex        =   37
            Top             =   80
            Width           =   255
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "病人信息"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   225
            TabIndex        =   20
            Top             =   90
            Width           =   1095
         End
         Begin VB.Label lbl住院号 
            BackStyle       =   0  'Transparent
            Caption         =   "住 院 号:"
            Height          =   195
            Left            =   225
            TabIndex        =   19
            Top             =   684
            Width           =   2580
         End
         Begin VB.Label lbl住院次数 
            BackStyle       =   0  'Transparent
            Caption         =   "住院次数:"
            Height          =   195
            Left            =   225
            TabIndex        =   18
            Top             =   918
            Width           =   2580
         End
         Begin VB.Label lbl出院科室 
            BackStyle       =   0  'Transparent
            Caption         =   "出院科室:"
            Height          =   195
            Left            =   225
            TabIndex        =   17
            Top             =   1152
            Width           =   2580
         End
         Begin VB.Label lbl姓名 
            BackStyle       =   0  'Transparent
            Caption         =   "姓   名:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   225
            TabIndex        =   16
            Top             =   450
            Width           =   2580
         End
         Begin VB.Label lbl住院医师 
            BackStyle       =   0  'Transparent
            Caption         =   "住院医师:"
            Height          =   195
            Left            =   225
            TabIndex        =   15
            Top             =   1386
            Width           =   2580
         End
         Begin VB.Label lbl编目日期 
            BackStyle       =   0  'Transparent
            Caption         =   "编目日期:"
            Height          =   195
            Left            =   225
            TabIndex        =   14
            Top             =   1620
            Width           =   2580
         End
      End
   End
   Begin VB.PictureBox picRight 
      Height          =   7080
      Left            =   3145
      Picture         =   "frm评分结果编辑.frx":00FF
      ScaleHeight     =   7020
      ScaleWidth      =   8280
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   8340
      Begin VB.ComboBox ComProName 
         Height          =   300
         Left            =   2805
         TabIndex        =   49
         Top             =   825
         Width           =   2070
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Left            =   750
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   825
         Width           =   1320
      End
      Begin VB.TextBox txt备注 
         Height          =   300
         Left            =   5460
         TabIndex        =   7
         Tag             =   "备注"
         Top             =   825
         Width           =   2550
      End
      Begin VB.CheckBox chk返回修改 
         Caption         =   "返回修改(&R)"
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   6705
         TabIndex        =   5
         Top             =   443
         Width           =   1500
      End
      Begin VSFlex8Ctl.VSFlexGrid fgMain 
         Height          =   5820
         Left            =   -45
         TabIndex        =   8
         Top             =   1200
         Width           =   8310
         _cx             =   14658
         _cy             =   10266
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483643
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16763080
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   14737632
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   4
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm评分结果编辑.frx":0318
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
         Ellipsis        =   1
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   1
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
         Begin zl9CISAudit.tipPopup tipPopup1 
            Height          =   540
            Left            =   1935
            Top             =   4800
            Width           =   3750
            _ExtentX        =   6615
            _ExtentY        =   953
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox txt评分 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFE0E0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3915
            TabIndex        =   36
            Text            =   "333"
            Top             =   945
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.ListBox lst评分 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   390
            Left            =   3960
            TabIndex        =   35
            Top             =   1485
            Visible         =   0   'False
            Width           =   1905
         End
      End
      Begin VB.TextBox txtNo 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   750
         TabIndex        =   1
         Top             =   435
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "项目"
         Height          =   180
         Left            =   2295
         TabIndex        =   50
         Top             =   870
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病理类型"
         Height          =   180
         Left            =   15
         TabIndex        =   47
         Top             =   870
         Width           =   720
      End
      Begin VB.Label lbl备注 
         AutoSize        =   -1  'True
         Caption         =   "备注"
         Height          =   180
         Left            =   5040
         TabIndex        =   6
         Top             =   885
         Width           =   360
      End
      Begin VB.Label labNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&NO."
         Height          =   180
         Left            =   435
         TabIndex        =   0
         Top             =   495
         Width           =   270
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "评分结果"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F0F0F0&
         Height          =   195
         Left            =   180
         TabIndex        =   34
         Top             =   90
         Width           =   1095
      End
      Begin VB.Label lbl得分 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   4230
         TabIndex        =   32
         Top             =   435
         Width           =   600
      End
      Begin VB.Label lbl评分人 
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   2970
         TabIndex        =   31
         Top             =   502
         Width           =   915
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "评分人:"
         Height          =   180
         Left            =   2280
         TabIndex        =   2
         Top             =   495
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "得分:"
         Height          =   180
         Left            =   3765
         TabIndex        =   3
         Top             =   495
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "等级:"
         Height          =   180
         Left            =   5025
         TabIndex        =   4
         Top             =   495
         Width           =   450
      End
      Begin VB.Label lbl等级 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   5490
         TabIndex        =   33
         Top             =   420
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   8505
      TabIndex        =   9
      Top             =   7380
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   9750
      TabIndex        =   10
      Top             =   7380
      Width           =   1100
   End
   Begin VB.Label LabStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "12.25%"
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   4545
      TabIndex        =   46
      Top             =   7440
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label labBar 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   2745
      TabIndex        =   44
      Top             =   7440
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   2670
      X2              =   7185
      Y1              =   7740
      Y2              =   7740
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   2670
      X2              =   7215
      Y1              =   7350
      Y2              =   7350
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      Visible         =   0   'False
      X1              =   15
      X2              =   11520
      Y1              =   7335
      Y2              =   7335
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   0
      X2              =   11520
      Y1              =   7230
      Y2              =   7230
   End
   Begin VB.Image imgOpen_White 
      Height          =   225
      Left            =   780
      Picture         =   "frm评分结果编辑.frx":0497
      Top             =   7980
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgClose_White 
      Height          =   225
      Left            =   1185
      Picture         =   "frm评分结果编辑.frx":04F9
      Top             =   7980
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgOpen 
      Height          =   225
      Left            =   105
      Picture         =   "frm评分结果编辑.frx":054E
      Top             =   7980
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgClose 
      Height          =   225
      Left            =   465
      Picture         =   "frm评分结果编辑.frx":05A3
      Top             =   7980
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgBGBlue 
      Height          =   1530
      Left            =   0
      Picture         =   "frm评分结果编辑.frx":05F2
      Top             =   0
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.Image imgBG 
      Height          =   1695
      Left            =   0
      Picture         =   "frm评分结果编辑.frx":07B2
      Top             =   0
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.Label pbrBar 
      Height          =   240
      Left            =   2670
      TabIndex        =   45
      Top             =   8625
      Visible         =   0   'False
      Width           =   4455
   End
End
Attribute VB_Name = "frm评分结果编辑"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private mfrmArchiveView As frmArchiveView
Private m_lng结果ID         As Long
Private m_lng病人ID         As Long
Private m_lng主页ID         As Long
Private m_lng方案ID         As Long
Private m_lng科室ID         As Long
Private m_str方式           As String     '添加、修改、重评
Private m_blnModed          As Boolean
Private edRow%, edCol%, edKey%
Private m_lngOldSJID        As Long         '旧的上级ID
Private m_lngCurSJID        As Long         '上级ID
Private m_bln多次编辑       As Boolean
Private zlCheck             As New clsCheck
Private mblnStop            As Boolean
Dim mbln编目后评分          As Boolean
Public Event AferSaveData()

Public Property Get Moded() As Boolean
   Moded = m_blnModed
End Property

Public Property Let Moded(ByVal blnModed As Boolean)
    m_blnModed = blnModed
End Property

'==============================================================================
'=功能： 窗口显示
'==============================================================================
Public Sub ShowForm(方式 As String, 结果ID As Long, 病人ID As Long, 主页ID As Long, 方案ID As Long, 科室ID As Long)
    Dim rsTemp      As ADODB.Recordset
    
    On Error GoTo errH
    
    m_bln多次编辑 = False
    
    m_blnModed = False
    m_str方式 = 方式    '添加/修改/重评
    m_lng结果ID = 结果ID
    m_lng病人ID = 病人ID
    m_lng主页ID = 主页ID
    m_lng科室ID = 科室ID
    
    '初始化数据
    Select Case 方式
        Case "新增"
            Me.Caption = "新增评分"
        Case "修改"
            Me.Caption = "修改评分结果"
        Case "重评"
            Me.Caption = "重新评分"
    End Select
    
    If m_str方式 = "重评" Or m_str方式 = "新增" Then
        '将m_lng方案ID设为默认方案ID
        gstrSQL = "select ID from 病案评分方案 where 类型= [1] and 选用 = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "住院", 1)
        
        If rsTemp.EOF Then
            MsgBox "请在【评分标准维护】中设置默认选用的评分方案。", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        Else
            m_lng方案ID = rsTemp.Fields(0).Value
        End If
    Else
        m_lng方案ID = 方案ID
    End If
    
    Call Fill病理类型
    Call Fill评分标准
    Call Fill评分项目
    
    If m_str方式 = "修改" Then
        Call Fill评分结果
    End If
    Me.Tag = "完成初始化"
    fgMain_CellChanged 0, 0
    If Me.Visible = False Then Me.Show
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 根据评分方案ID，填入评分标准网格。
'=       包括了界面文本框内容的初始化  m_lng方案ID   m_lng病人ID  m_lng主页ID
'=
'=       m_lng病人ID  m_lng主页ID  =>病人信息
'==============================================================================
Private Sub Fill评分标准()

    Dim rsTemp          As ADODB.Recordset
    Dim lngIndex        As Long
    Dim i               As Long

    On Error GoTo errH
        
    lst评分.Clear
    lst评分.AddItem "1 - 空"
    lst评分.AddItem "2 - 定级"
    txt评分.Text = ""
    
    gstrSQL = "select 姓名,住院号,出院科室,住院医师,编目日期 from 病案质量报表视图 where 病人ID=[1] And 主页ID = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng病人ID, m_lng主页ID)
    
    If rsTemp.EOF Then
        MsgBox "没有选择病人", vbInformation, gstrSysName
        CmdCancel_Click
        Exit Sub
    Else
        lbl姓名 = "姓   名:" & NVL(rsTemp("姓名"))
        txtNo = NVL(rsTemp("姓名"))
        lbl住院号 = "住 院 号:" & NVL(rsTemp("住院号"))
        lbl住院次数 = "住院次数:" & m_lng主页ID
        lbl出院科室 = "出院科室:" & NVL(rsTemp("出院科室"))
        lbl住院医师 = "住院医师:" & NVL(rsTemp("住院医师"))
        lbl编目日期 = "编目日期:" & NVL(rsTemp("编目日期"))
    End If
    rsTemp.Close
    
    'm_lng方案ID:方案信息
    gstrSQL = "select 名称,分制,总分,上值,下值 from 病案评分方案 where ID= [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng方案ID)
    If rsTemp.EOF Then
        MsgBox "未知的评分方案。", vbInformation, gstrSysName
        CmdCancel_Click
        Exit Sub
    Else
        lbl方案名称 = "名称:" & NVL(rsTemp("名称"))
        lbl分制 = "分制:" & NVL(rsTemp("分制"))
        lbl总分 = "总分:" & NVL(rsTemp("总分"))
        lbl上值 = "上值:" & NVL(rsTemp("上值"))
        lbl下值 = "下值:" & NVL(rsTemp("下值"))
    End If
    rsTemp.Close
    
    '评分信息
    If m_lng结果ID <> 0 And m_str方式 = "修改" Then
        gstrSQL = "" & _
            "   Select Id, 病人id, 主页id, 方案id, 总分, 等级, 返回修改,病理类型, 评分人, 评分时间, 审核人, 审核时间, 备注 " & _
            "   From 病案评分结果 " & _
            "   Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng结果ID)
        
        If Not rsTemp.EOF Then
            lbl评分人 = NVL(rsTemp("评分人"), "")
            If InStr(gstrPrivs, "修改他人评分") = 0 And UCase(lbl评分人) <> UCase(gstrUserName) Then  '无所有科室功能
                cmdOK.Enabled = False
            End If
            lbl得分 = NVL(rsTemp("总分"), 0)
            lbl等级 = NVL(rsTemp("等级"), "")
            chk返回修改.Value = IIf(NVL(rsTemp("返回修改"), 0) = 0, vbUnchecked, vbChecked)
            txt备注.Text = NVL(rsTemp!备注)
            If NVL(rsTemp!病理类型) = "" Then
                cbo.ListIndex = 0
            Else
                On Error Resume Next
                cbo.Text = NVL(rsTemp!病理类型)
            End If
        End If
        rsTemp.Close
    Else
        lbl评分人 = gstrUserName
        lbl得分 = ""
        lbl等级 = ""
    End If
    
    If m_bln多次编辑 = True Then    '第二次新增就不用填写评分标准了，默认标准。
        If fgMain.Rows > 1 Then
            fgMain.Row = 1
            fgMain.ShowCell 1, 4
            fgMain.SetFocus
        End If
        Exit Sub
    End If
    
    '确定分制
    Dim bln扣分制 As Boolean, intSign As Long
    gstrSQL = "select 分制 from 病案评分方案 where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng方案ID)
    
    bln扣分制 = True
    If Not rsTemp.EOF Then
        bln扣分制 = IIf(NVL(rsTemp("分制"), "加分制") = "加分制", False, True)
    End If
    rsTemp.Close
    
    If bln扣分制 Then
        intSign = -1
    Else
        intSign = 1
    End If

    With fgMain
        .Redraw = flexRDNone
        .Rows = 1
        .Clear
        '数据填入
        .Cols = 11
        .Cell(flexcpText, 0, 0) = "项目"
        .Cell(flexcpText, 0, 1) = "标准分值"
        .Cell(flexcpText, 0, 2) = "缺陷内容"
        .Cell(flexcpText, 0, 3) = "评分标准"
        .Cell(flexcpText, 0, 4) = "评分"
        .Cell(flexcpText, 0, 5) = "可否修改"
        .Cell(flexcpText, 0, 6) = "ID"
        .Cell(flexcpText, 0, 7) = "上级ID"
        .Cell(flexcpText, 0, 8) = "方案ID"
        .Cell(flexcpText, 0, 9) = "备注"
        .Cell(flexcpText, 0, 10) = "否决等级"
        .ExtendLastCol = True
        '确定方案名称
        If m_lng方案ID < 1 Then .Redraw = flexRDDirect: Exit Sub
        
        gstrSQL = "" & _
            "   Select 上级序号, 序号, Id, 上级id, 方案id, 项目, 标准分值, 基本要求, 缺陷内容, 扣分标准, 隐藏,否决等级 " & _
            "   From 病案评分标准视图 " & _
            "   Where 隐藏='否' and 方案ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng方案ID)
        .FocusRect = flexFocusSolid
        .Rows = rsTemp.RecordCount + 1
        i = 1
        Do Until rsTemp.EOF
            .Cell(flexcpText, i, 0) = IIf(IsNull(rsTemp.Fields("项目")), "", rsTemp.Fields("项目"))
            .Cell(flexcpText, i, 1) = IIf(IsNull(rsTemp.Fields("标准分值")), " ", Format(rsTemp.Fields("标准分值"), "####分"))
            .Cell(flexcpText, i, 2) = IIf(IsNull(rsTemp.Fields("缺陷内容")), "", rsTemp.Fields("缺陷内容"))
            .Cell(flexcpText, i, 3) = IIf(IsNull(rsTemp.Fields("扣分标准")), "", IIf(rsTemp.Fields("扣分标准") = "甲", "甲级", IIf(rsTemp.Fields("扣分标准") = "乙", "乙级", IIf(rsTemp.Fields("扣分标准") = "丙", "丙级", IIf(rsTemp.Fields("扣分标准") = "否", "单项否决", rsTemp.Fields("扣分标准"))))))
            .Cell(flexcpText, i, 4) = ""
            If intSign = 1 Then
                .Cell(flexcpForeColor, i, 4) = RGB(0, 0, 255)
            Else
                .Cell(flexcpForeColor, i, 4) = RGB(255, 0, 0)
            End If
            .Cell(flexcpAlignment, i, 4) = flexAlignCenterCenter
            .Cell(flexcpText, i, 5) = ""
            .Cell(flexcpForeColor, i, 5) = RGB(0, 0, 0)
            .Cell(flexcpText, i, 6) = IIf(IsNull(rsTemp.Fields("ID")), "", rsTemp.Fields("ID"))
            .Cell(flexcpText, i, 7) = IIf(IsNull(rsTemp.Fields("上级ID")), "", rsTemp.Fields("上级ID"))
            .Cell(flexcpText, i, 8) = IIf(IsNull(rsTemp.Fields("方案ID")), "", rsTemp.Fields("方案ID"))
            .Cell(flexcpText, i, 9) = ""
            .Cell(flexcpText, i, 10) = NVL(rsTemp.Fields!否决等级)
            rsTemp.MoveNext
            i = i + 1
        Loop
        '自动换行
        .WordWrap = True
        '合并单元格
        .MergeCells = 2
        .MergeCol(.ColIndex("项目")) = True
        .MergeCol(.ColIndex("标准分值")) = True
        '对齐设置
        .ColAlignment(.ColIndex("项目")) = flexAlignLeftCenter
        .ColAlignment(.ColIndex("标准分值")) = flexAlignCenterCenter
        .ColAlignment(.ColIndex("评分标准")) = flexAlignCenterCenter
        .ColAlignment(.ColIndex("评分")) = flexAlignCenterCenter
        .ColAlignment(.ColIndex("可否修改")) = flexAlignCenterCenter
        .ColAlignment(.ColIndex("备注")) = flexAlignLeftCenter
        
        '隐藏单元格
        .ColWidth(.ColIndex("ID")) = 0
        .ColWidth(.ColIndex("上级ID")) = 0
        .ColWidth(.ColIndex("方案ID")) = 0
        '宽度设置
        .ColWidth(.ColIndex("项目")) = 600
        .ColWidth(.ColIndex("标准分值")) = 600
        .ColWidth(.ColIndex("缺陷内容")) = 3200
        .ColWidth(.ColIndex("评分标准")) = 950
        .ColWidth(.ColIndex("评分")) = 1000
        .ColWidth(.ColIndex("可否修改")) = 850
        '行高设置
        .RowHeightMin = 300
        '最大宽度设置
        .ColWidthMax = 7000
        '自动适应行高、列宽
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize .ColIndex("缺陷内容")
        .SelectionMode = flexSelectionFree
        .AllowBigSelection = False
        
        .Editable = flexEDKbdMouse   '可编辑

        .Redraw = flexRDBuffered
        '选中先前的行
    End With
    If fgMain.Rows > 1 Then
        fgMain.Row = 1
        fgMain.ShowCell 1, 4
    End If

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

'==============================================================================
'=功能： 装入对应主页的评分结果
'==============================================================================
Private Function Fill评分结果() As Boolean
    Dim rs              As ADODB.Recordset
    Dim lngJGID         As Long
    Dim i               As Long
    
    On Error GoTo errH
    
    fgMain.Redraw = flexRDNone
    
    For i = 1 To fgMain.Rows - 1
        fgMain.Cell(flexcpText, i, 4) = ""
        fgMain.Cell(flexcpText, i, 5) = ""
    Next
    
    '确定分制
    Dim bln扣分制 As Boolean, intSign As Long
    gstrSQL = "select 分制 from 病案评分方案 where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng方案ID)
    
    bln扣分制 = True
    If Not rs.EOF Then
        bln扣分制 = IIf(NVL(rs("分制"), "加分制") = "加分制", False, True)
    End If
    rs.Close
    
    If bln扣分制 Then
        intSign = -1
    Else
        intSign = 1
    End If
    
    gstrSQL = "" & _
        "   select  A.ID,A.项目,A.标准分值,A.基本要求,A.缺陷内容,A.扣分标准," & _
        "           (select decode(缺陷等级,null,to_CHAR(单项分数),缺陷等级) from 病案评分明细 where 评分标准ID=A.ID and 主表ID=[1]) as 评分," & _
        "           (select 可否修改 from 病案评分明细 where 评分标准ID=A.ID and 主表ID=[1]) as 可否修改," & _
        "           (select 备注 from 病案评分明细 where 评分标准ID=A.ID and 主表ID=[1]) as 备注" & _
        "   From 病案评分标准视图 A " & _
        "   Where A.隐藏='否' and A.方案ID=(select B.方案ID from 病案评分结果 B where B.ID=[1]) " & _
        "   Order by A.上级ID,A.ID "
        
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng结果ID)
        
    If Not rs.EOF Then
        For i = 1 To fgMain.Rows - 1
            rs.MoveFirst
            rs.Find "ID=" & Val(fgMain.Cell(flexcpText, i, 6))
            If Not rs.EOF Then
                If Not IsNull(rs("评分")) Then
                    Select Case rs("评分")
                    Case "甲", "乙", "丙"
                        fgMain.Cell(flexcpText, i, 4) = rs("评分").Value + "级"
                    Case "否"
                        fgMain.Cell(flexcpText, i, 4) = "单项否决"
                    Case Else
                        fgMain.Cell(flexcpText, i, 4) = IIf(Abs(NVL(rs("评分").Value, 0)) < 1, Format(Abs(NVL(rs("评分").Value, 0)), "0.0"), Abs(NVL(rs("评分").Value, 0)))
                        If intSign = -1 Then
                            fgMain.Cell(flexcpForeColor, i, 4) = RGB(255, 0, 0)
                        Else
                            fgMain.Cell(flexcpForeColor, i, 4) = RGB(0, 0, 255)
                        End If
                    End Select
                End If
                If Not IsNull(rs("可否修改")) Then
                    If rs("可否修改") = 1 Then
                        fgMain.Cell(flexcpText, i, 5) = "√"
                    End If
                End If
                fgMain.Cell(flexcpText, i, 9) = NVL(rs!备注)
            End If
        Next
    End If
    
    fgMain.Redraw = flexRDBuffered
    If fgMain.Rows > 1 Then
        fgMain.Row = 1
        fgMain.ShowCell 1, 4
    End If

    Fill评分结果 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Fill评分结果 = False
End Function

Private Sub Fill病理类型()
    Dim rs As New ADODB.Recordset
    On Error GoTo errH
    gstrSQL = "" & _
        "Select 编码,名称,简码,缺省标志 From 病理类型"
        
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    With cbo
        .Clear
        .AddItem ""
        .ItemData(.NewIndex) = 1
        
        If Not rs.EOF Then
            rs.MoveFirst
            Do Until rs.EOF
                .AddItem zlCommFun.NVL(rs!名称)
                 .ItemData(.NewIndex) = .NewIndex + 1

                rs.MoveNext
            Loop
        End If
        
        If .ListCount > 0 Then .ListIndex = 0
        
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Fill评分项目()
    Dim rs As New ADODB.Recordset
    On Error GoTo errH
    gstrSQL = "" & _
        "    Select A.ID,A.描述 From 病案评分标准 A,病案评分方案 B Where A.方案ID= B.ID And B.选用=1"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    With ComProName
        .Clear
        .AddItem ""
        .ItemData(.NewIndex) = 1
        
        If Not rs.EOF Then
            rs.MoveFirst
            Do Until rs.EOF
                .AddItem zlCommFun.NVL(rs!描述)
                 .ItemData(.NewIndex) = .NewIndex + 1

                rs.MoveNext
            Loop
        End If
        
        If .ListCount > 0 Then .ListIndex = 0
        
    End With
    
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume Next
    End If
End Sub

'==============================================================================
'=功能： 返回修改值变量处理
'==============================================================================
Private Sub chk返回修改_Click()
    On Error GoTo errH
    If chk返回修改.Value = vbChecked Then
        chk返回修改.FontBold = True
    Else
        chk返回修改.FontBold = False
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 自动计算出相应的分值填写到评分标准
'==============================================================================
Private Sub cmdAuto_Click()
Dim lngLoop As Long, strID As String, strSQL As String
Dim strReturn As String, strMid As String, strAlidin As String
    Dim rsTemp      As ADODB.Recordset
    
    On Error GoTo errH
    
    If fgMain.Rows = 1 Then
        zlCheck.Msg_OK "无评分方案，无需评分！", vbCritical
        Exit Sub
    End If
    If zlCheck.Msg_OKC("确认进行自动评分计算吗？") Then Exit Sub
    
    cmdAuto.Visible = False
    cmdStop.Visible = True
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
    labBar.Width = 0
    labBar.Visible = True
    Line3.Visible = True
    Line4.Visible = True
    LabStatus.Visible = True
    
    DoEvents
    '读取大于当前行的记录数据
    For lngLoop = 1 To fgMain.Rows - 1
        LabStatus.Caption = Format(Round((lngLoop / (fgMain.Rows - 1)) * 100, 2), "0.00") & " %"
        labBar.Width = lngLoop * pbrBar.Width / (fgMain.Rows - 1)
        If mblnStop Then
            cmdAuto.Visible = True
            cmdOK.Enabled = True
            cmdCancel.Enabled = True
            
            cmdStop.Visible = False
            labBar.Visible = False
            Line3.Visible = False
            Line4.Visible = False
            LabStatus.Visible = False
            mblnStop = False
            Call zlCheck.Msg_OK("病案自动评分中途取消，已完成部分评分！", vbCritical)
            Exit Sub
        End If
        strID = fgMain.TextMatrix(lngLoop, fgMain.ColIndex("ID"))
        
        strSQL = "select 判断依据,数据源 from 病案评分标准 where ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strID)
        If Not zlCheck.Connection_ChkRsState(rsTemp) Then
            strSQL = "" & rsTemp.Fields!判断依据
            If strSQL <> "" Then
                If rsTemp!数据源 = 0 Then
                    strSQL = CheckAuditSql_OUT(strSQL, m_lng病人ID, m_lng主页ID)
                    Set rsTemp = zlDatabase.OpenSQLRecord("select ZL_FUN_ExecSql('" & Replace(strSQL, "'", "''") & "') from dual", Me.Caption)
                ElseIf gobjEmr Is Nothing Then
                    MsgBox "本机未安装病历组件，不能进行评分，请检查！", vbInformation, gstrSysName
                    mblnStop = True
                ElseIf Not gobjEmr Is Nothing Then
                    If strMid = "" Then Call GetEMR_MID_ALIDIN(m_lng病人ID, m_lng主页ID, strMid, strAlidin) '取新病历主体ID,活动ID
                    strSQL = Replace(rsTemp!判断依据, "[MID]", ":mid")
                    strSQL = Replace(rsTemp!判断依据, "[ALIDIN]", ":alidin")
                    strReturn = gobjEmr.OpenSQLRecordset(strSQL, IIf(strMid = "", "", strMid & "^" & DbType.T_String & "^mid") & IIf(strAlidin = "", "", IIf(strMid = "", "", "|") & strAlidin & "^" & DbType.T_String & "^alidin"), rsTemp)
                    If strReturn <> "" Then Set rsTemp = New ADODB.Recordset
                End If
                
                If Not zlCheck.Connection_ChkRsState(rsTemp) Then
                    If InStr(1, rsTemp.Fields(0), "[zlsoft]Error[zlsoft]") = 0 Then
                        fgMain.TextMatrix(lngLoop, fgMain.ColIndex("评分")) = "" & rsTemp.Fields(0)
                    Else
                        fgMain.TextMatrix(lngLoop, fgMain.ColIndex("评分")) = 0
                    End If
                End If
            End If
        End If
        DoEvents
    Next
    zlCheck.Msg_OK ("病案自动评分成功！")
    cmdAuto.Visible = True
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
    cmdStop.Visible = False
    labBar.Visible = False
    Line3.Visible = False
    Line4.Visible = False
    LabStatus.Visible = False
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Call zlCheck.Msg_OK("病案自动评分失败！", vbCritical)
    cmdAuto.Visible = True
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
    cmdStop.Visible = False
    labBar.Visible = False
    Line3.Visible = False
    Line4.Visible = False
    LabStatus.Visible = False
End Sub

'==============================================================================
'=功能： 停止自动评分
'==============================================================================
Private Sub cmdStop_Click()
    On Error GoTo errH
    
    mblnStop = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 查阅病案
'==============================================================================
Private Sub cmdViewArchive_Click()
    On Error GoTo errH
    If mfrmArchiveView Is Nothing Then Set mfrmArchiveView = New frmArchiveView
    Call mfrmArchiveView.ShowArchive(Me, m_lng病人ID, m_lng主页ID, False)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ComProName_Click()
    Dim lngRow      As Long
    Dim intCol      As Integer
    Dim bytMatch    As Byte
    Dim lngLoop     As Long
    
    On Error GoTo errH
    
    lngRow = 0
    If ComProName.Locked Then Exit Sub

    If fgMain.ColIndex("缺陷内容") = -1 Then Exit Sub
  
    '读取大于当前行的记录数据
    For lngLoop = fgMain.Row + 1 To fgMain.Rows - 1
        If InStr(UCase(fgMain.TextMatrix(lngLoop, fgMain.ColIndex("缺陷内容"))), UCase(ComProName.Text)) > 0 Then
            lngRow = lngLoop
            Exit For
        End If
    Next
    '读取小于当前行的记录数据
    If lngRow = 0 Then
        For lngLoop = 0 To fgMain.Row
            If InStr(UCase(fgMain.TextMatrix(lngLoop, fgMain.ColIndex("缺陷内容"))), UCase(ComProName.Text)) > 0 Then
                lngRow = lngLoop
                Exit For
            End If
        Next
    End If
    If fgMain.Rows > 1 And lngRow >= 1 Then fgMain.Row = lngRow
'        fgMain.Cell lngRow
    fgMain.ShowCell lngRow, 4
    Call LocationObj(ComProName)
 
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ComProName_KeyPress(KeyAscii As Integer)
    Dim lngRow      As Long
    Dim intCol      As Integer
    Dim bytMatch    As Byte
    Dim lngLoop     As Long
    Dim strTmpProName As String
    
    On Error GoTo errH
    
    lngRow = 0
    If ComProName.Locked Then Exit Sub

    If fgMain.ColIndex("缺陷内容") = -1 Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        '读取大于当前行的记录数据
        
        If zlCommFun.IsNumOrChar(ComProName.Text) Then
            For lngLoop = fgMain.Row + 1 To fgMain.Rows - 1
                If InStr(UCase(zlCommFun.SpellCode(fgMain.TextMatrix(lngLoop, fgMain.ColIndex("缺陷内容")))), UCase(ComProName.Text)) > 0 Then
                    lngRow = lngLoop
                    Exit For
                End If
            Next
            '读取小于当前行的记录数据
            If lngRow = 0 Then
                For lngLoop = 0 To fgMain.Row
                    If InStr(UCase(zlCommFun.SpellCode(fgMain.TextMatrix(lngLoop, fgMain.ColIndex("缺陷内容")))), UCase(ComProName.Text)) > 0 Then
                        lngRow = lngLoop
                        Exit For
                    End If
                Next
            End If
        Else
            For lngLoop = fgMain.Row + 1 To fgMain.Rows - 1
                If InStr(UCase(fgMain.TextMatrix(lngLoop, fgMain.ColIndex("缺陷内容"))), UCase(ComProName.Text)) > 0 Then
                    lngRow = lngLoop
                    Exit For
                End If
            Next
            '读取小于当前行的记录数据
            If lngRow = 0 Then
                For lngLoop = 0 To fgMain.Row
                    If InStr(UCase(fgMain.TextMatrix(lngLoop, fgMain.ColIndex("缺陷内容"))), UCase(ComProName.Text)) > 0 Then
                        lngRow = lngLoop
                        Exit For
                    End If
                Next
            End If
        End If
        If fgMain.Rows > 1 And lngRow >= 1 Then fgMain.Row = lngRow
'        fgMain.Cell lngRow
        fgMain.ShowCell lngRow, 4
        Call LocationObj(ComProName)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 网络可修改编辑
'==============================================================================
Private Sub fgMain_Click()
    On Error GoTo errH
    
    If fgMain.Col = 5 Then
        If fgMain.Cell(flexcpText, fgMain.Row, 5) = "" Then
            fgMain.Cell(flexcpText, fgMain.Row, 5) = "√"
        Else
            fgMain.Cell(flexcpText, fgMain.Row, 5) = ""
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 网络编辑中不能输入“'”
'==============================================================================
Private Sub fgMain_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    On Error GoTo errH
    
    If KeyAscii = Asc("'") Then
       KeyAscii = 0
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 网络单击时对m_lngCurSJID项目变量赋值
'==============================================================================
Private Sub fgMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_lngCurSJID = IIf(Len(fgMain.Cell(flexcpText, fgMain.Row, 6)) = 0, 0, Val(fgMain.Cell(flexcpText, fgMain.Row, 6)))      '获取ID
End Sub

'==============================================================================
'=功能： 网络行列变动时重取
'==============================================================================
Private Sub fgMain_RowColChange()
    Dim lngID               As Long
    Dim lngCurID            As Long
    Dim lngCurSJID          As Long
    
    On Error GoTo errH
    
    If fgMain.Row < 0 Then
        lngCurSJID = 0
        lngCurID = 0
        Exit Sub
    End If
    
    lngCurID = IIf(Len(fgMain.Cell(flexcpText, fgMain.Row, 6)) = 0, 0, Val(fgMain.Cell(flexcpText, fgMain.Row, 6)))         '获取ID
    lngCurSJID = IIf(Len(fgMain.Cell(flexcpText, fgMain.Row, 7)) = 0, 0, Val(fgMain.Cell(flexcpText, fgMain.Row, 7)))       '获取上级ID
    
    If lngCurSJID = 0 Then
        lngID = lngCurID
    Else
        lngID = lngCurSJID
    End If
    
    Show基本要求 lngID, fgMain.Cell(flexcpText, fgMain.Row, 0), fgMain.Cell(flexcpText, fgMain.Row, 1)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 网络更新后检测字符长度
'==============================================================================
Private Sub fgMain_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    On Error GoTo errH
    
    If Col = 9 Then
        If zlCommFun.ActualLen(Trim(fgMain.EditText)) > 50 Then
            MsgBox "你输入的备注大于了25个汉字或50个字符,不能继续!"
            Cancel = True
            Exit Sub
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 窗口迭件初始化
'==============================================================================
Private Sub Form_Initialize()
    On Error GoTo errH
    Call InitCommonControls
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 页面初始化
'==============================================================================
Private Sub Form_Load()
    On Error GoTo errH
    '获取系统参数：是否编目后才能评分
    mbln编目后评分 = Val(zlDatabase.GetPara(91, glngSys, 0)) = 1
    m_lngOldSJID = -1
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 窗口变化
'==============================================================================
Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.Height < 8175 Then
        Me.Height = 8175
    End If
    If Me.Width < 11520 Then
        Me.Width = 11640
    End If
    With txt备注 '
        .Width = ScaleWidth - .Left - 50
    End With
    With cmdCancel
        .Top = ScaleHeight - .Height - 85
        .Left = ScaleWidth - .Width - 100
    End With
    With cmdOK
        .Top = cmdCancel.Top
        .Left = cmdCancel.Left - .Width - 50
    End With
    With cmdAuto
        .Top = cmdOK.Top
        .Left = cmdOK.Left - .Width - 50
    End With
    With cmdStop
        .Top = cmdOK.Top
        .Left = cmdOK.Left - .Width - 50
    End With
    With cmdHelp
        .Top = cmdCancel.Top
    End With
    With cmdViewArchive
        .Top = cmdCancel.Top
    End With
    With pbrBar
        .Top = cmdHelp.Top + 30
        .Left = cmdViewArchive.Left + cmdViewArchive.Width + 200
        .Width = cmdStop.Left - cmdViewArchive.Left - cmdViewArchive.Width - 400
    End With
    With Line1
        .Y1 = cmdCancel.Top - 85
        .y2 = .Y1
        .X1 = 0
        .x2 = ScaleWidth
    End With
    With Line2
        .Y1 = Line1.Y1 + 30
        .y2 = .Y1
        .X1 = 0
        .x2 = ScaleWidth
    End With
    
    With picLeft
        .Height = Line1.Y1 - 85 - .Top
    End With
    
    With fgMain
        .Width = ScaleWidth - .Left - 50
        .Height = Line1.Y1 - .Top - 85
    End With

    With chk返回修改
        .Left = ScaleWidth - .Width
    End With
    With picRight
        .Width = ScaleWidth - .Left
        .Height = Line1.Y1 - .Top - 85
    End With
    With Line3
        .Y1 = pbrBar.Top - 10
        .y2 = pbrBar.Top - 10
        .X1 = pbrBar.Left
        .x2 = pbrBar.Left + pbrBar.Width
    End With
    With Line4
        .Y1 = pbrBar.Top + pbrBar.Height + 40
        .y2 = pbrBar.Top + pbrBar.Height + 40
        .X1 = pbrBar.Left
        .x2 = pbrBar.Left + pbrBar.Width
    End With
    With LabStatus
        .Move pbrBar.Left + pbrBar.Width / 2 - 50, pbrBar.Top + 50
    End With
    With labBar
        .Move pbrBar.Left, pbrBar.Top + 20
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Unload mfrmArchiveView
    Set mfrmArchiveView = Nothing
    Set zlCheck = Nothing
End Sub
'==============================================================================
'=功能： 病人信息缩放
'==============================================================================
Private Sub imgBRXX_Click()

    On Error GoTo errH
    
    If imgBRXX.Tag = "" Then
        imgBRXX.Tag = "Opened"
        imgBRXX.Picture = imgOpen_White.Picture
        pic病人信息.Height = 340
    Else
        imgBRXX.Tag = ""
        imgBRXX.Picture = imgClose_White.Picture
        pic病人信息.Height = 1950
    End If
    imgBRXX.Refresh
    picLeft_Resize
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：病人信息变化
'==============================================================================
Private Sub imgBRXX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If X >= 0 And X <= imgBRXX.ScaleWidth And Y >= 0 And Y <= imgBRXX.ScaleHeight Then
        SetCapture imgBRXX.hWnd
        '鼠标移入！！！
        imgBRXX.Line (0, 0)-(imgBRXX.ScaleWidth - Screen.TwipsPerPixelX, imgBRXX.ScaleHeight - Screen.TwipsPerPixelY), vbWhite, B
    Else
        '鼠标移出！！！
        imgBRXX.Cls
        ReleaseCapture
    End If
End Sub

'==============================================================================
'=功能：方案信息缩放
'==============================================================================
Private Sub imgFAXX_Click()
On Error Resume Next
    If imgFAXX.Tag = "" Then
        imgFAXX.Tag = "Opened"
        imgFAXX.Picture = imgOpen.Picture
        pic方案信息.Height = 340
    Else
        imgFAXX.Tag = ""
        imgFAXX.Picture = imgClose.Picture
        pic方案信息.Height = 1695
    End If
    imgFAXX.Refresh
    picLeft_Resize
End Sub

'==============================================================================
'=功能：方案信息变化
'==============================================================================
Private Sub imgFAXX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If X >= 0 And X <= imgFAXX.ScaleWidth And Y >= 0 And Y <= imgFAXX.ScaleHeight Then
        SetCapture imgFAXX.hWnd
        '鼠标移入！！！
        imgFAXX.Line (0, 0)-(imgFAXX.ScaleWidth - Screen.TwipsPerPixelX, imgFAXX.ScaleHeight - Screen.TwipsPerPixelY), vbBlue, B
    Else
        '鼠标移出！！！
        imgFAXX.Cls
        ReleaseCapture
    End If
End Sub

'==============================================================================
'=功能：项目信息缩放
'==============================================================================
Private Sub imgXMXX_Click()
On Error Resume Next
    If imgXMXX.Tag = "" Then
        imgXMXX.Tag = "Opened"
        imgXMXX.Picture = imgOpen.Picture
        pic项目信息.Height = 340
    Else
        imgXMXX.Tag = ""
        imgXMXX.Picture = imgClose.Picture
        pic项目信息.Height = Abs(picLeft.ScaleHeight - pic病人信息.Height - pic方案信息.Height - 135 * 4)
    End If
    imgXMXX.Refresh
    picLeft_Resize
End Sub

'==============================================================================
'=功能：项目信息变化
'==============================================================================
Private Sub imgXMXX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If X >= 0 And X <= imgXMXX.ScaleWidth And Y >= 0 And Y <= imgXMXX.ScaleHeight Then
        SetCapture imgXMXX.hWnd
        '鼠标移入！！！
        imgXMXX.Line (0, 0)-(imgXMXX.ScaleWidth - Screen.TwipsPerPixelX, imgXMXX.ScaleHeight - Screen.TwipsPerPixelY), vbBlue, B
    Else
        '鼠标移出！！！
        imgXMXX.Cls
        ReleaseCapture
    End If
End Sub

'==============================================================================
'=功能：评分双击编辑
'==============================================================================
Private Sub lst评分_DblClick()
    fgMain.SetFocus
    If fgMain.Row = fgMain.Rows - 1 Then
        If cmdOK.Enabled Then cmdOK.SetFocus
        Exit Sub
    End If
    fgMain.Row = fgMain.Row + 1
    fgMain.ShowCell fgMain.Row, 4
End Sub

'==============================================================================
'=功能： 评分编辑完成
'==============================================================================
Private Sub lst评分_LostFocus()
    On Error GoTo errH
    
    If lst评分.ListIndex = 0 Then
        fgMain.TextMatrix(edRow, edCol) = ""
    Else
        fgMain.TextMatrix(edRow, edCol) = fgMain.Cell(flexcpText, edRow, 3)
    End If
    lst评分.Visible = False
    edKey = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 点击取消
'==============================================================================
Private Sub CmdCancel_Click()
    On Error GoTo errH
    
    If m_bln多次编辑 Then
        Moded = True
    Else
        Moded = False
    End If
    Unload Me
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 点击帮助
'==============================================================================
Private Sub cmdHelp_Click()
    On Error GoTo errH
    ShowHelp App.ProductName, Me.hWnd, Me.Name, 3
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 点击确定保存数据
'==============================================================================
Private Sub CmdOK_Click()
    Dim strT            As String
    Dim r               As Long
    Dim lngID           As Long
    Dim lng明细ID       As Long
    Dim lng可否修改     As Long
    
    On Error GoTo errH
    If zlCommFun.ActualLen(Trim(txt备注.Text)) > 50 Then
        MsgBox "你输入的备注大于了25个汉字或50个字符,不能继续!"
        txt备注.SelStart = 1
        txt备注.SelLength = 100
        If txt备注.Enabled Then txt备注.SetFocus
        Exit Sub
    End If
    
    gcnOracle.BeginTrans
    '保存结果
    If m_str方式 = "重评" Or m_str方式 = "修改" Then  '清除以前的评分结果
        gstrSQL = "ZL_病案评分结果_Delete(" & m_lng结果ID & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    
    lngID = zlDatabase.GetNextId("病案评分结果")
    'Zl_病案评分结果_Insert
    gstrSQL = "ZL_病案评分结果_Insert("
    '  Id_In       In 病案评分结果.ID%Type,
    gstrSQL = gstrSQL & "" & lngID & ","
    '  病人id_In   In 病案评分结果.病人id%Type,
    gstrSQL = gstrSQL & "" & m_lng病人ID & ","
    '  主页id_In   In 病案评分结果.主页id%Type,
    gstrSQL = gstrSQL & "" & m_lng主页ID & ","
    '  方案id_In   In 病案评分结果.方案id%Type,
    gstrSQL = gstrSQL & "" & m_lng方案ID & ","
    '  总分_In     In 病案评分结果.总分%Type,
    gstrSQL = gstrSQL & "" & Val(lbl得分) & ","
    '  等级_In     In 病案评分结果.等级%Type,
    gstrSQL = gstrSQL & "'" & IIf(lbl等级 = "不合格", "否", lbl等级) & "',"
    '  病理类型_In In 病案评分结果.病理类型%Type,
    gstrSQL = gstrSQL & "" & IIf(Trim(cbo.Text) = "", "NULL", "'" & Trim(cbo.Text) & "'") & ","
    '  备注_In     In 病案评分结果.备注%Type,
    gstrSQL = gstrSQL & "" & IIf(Trim(txt备注.Text) = "", "NULL", "'" & Trim(txt备注.Text) & "'") & ","
    '  评分人_In   In 病案评分结果.评分人%Type,
    gstrSQL = gstrSQL & "'" & gstrUserName & "',"
    '  评分时间_In In 病案评分结果.评分时间%Type,
    gstrSQL = gstrSQL & "Sysdate,"
    '  审核人_In   In 病案评分结果.审核人%Type,
    gstrSQL = gstrSQL & "NULL,"
    '  审核时间_In In 病案评分结果.审核时间%Type,
    gstrSQL = gstrSQL & "NULL,"
    '  返回修改_In In 病案评分结果.返回修改%Type
    gstrSQL = gstrSQL & "" & IIf(chk返回修改.Value = vbChecked, "1", "Null") & ")"
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
     
    strT = "ZL_病案评分明细_Insert"
    
    For r = 1 To fgMain.Rows - 1
        If Trim(fgMain.Cell(flexcpText, r, 4)) <> "" Or fgMain.Cell(flexcpText, r, 5) <> "" Then
            If fgMain.Cell(flexcpText, r, 5) = "√" Then
                lng可否修改 = 1
            Else
                lng可否修改 = 0
            End If
            lng明细ID = zlDatabase.GetNextId("病案评分明细")    '？？？？

            gstrSQL = "ZL_病案评分明细_Insert("
            gstrSQL = gstrSQL & "" & lng明细ID & ","
            gstrSQL = gstrSQL & "" & lngID & ","
            gstrSQL = gstrSQL & "" & fgMain.Cell(flexcpText, r, 6) & ","
            
            Select Case fgMain.Cell(flexcpText, r, 4)
                Case "甲级"
                    gstrSQL = gstrSQL & "null,"
                    gstrSQL = gstrSQL & "'甲',"
                Case "乙级"
                    gstrSQL = gstrSQL & "null,"
                    gstrSQL = gstrSQL & "'乙',"
                Case "丙级"
                    gstrSQL = gstrSQL & "null,"
                    gstrSQL = gstrSQL & "'丙',"
                Case "单项否决"
                    gstrSQL = gstrSQL & "null,"
                    gstrSQL = gstrSQL & "'否',"
                Case "" '只有“可否修改”栏填写了！
                    gstrSQL = gstrSQL & "null,"
                    gstrSQL = gstrSQL & "null,"
                Case Else
                    gstrSQL = gstrSQL & "" & Abs(Val(fgMain.Cell(flexcpText, r, 4))) & ","
                    gstrSQL = gstrSQL & "null,"
            End Select
            gstrSQL = gstrSQL & "" & lng可否修改 & ","
            gstrSQL = gstrSQL & "" & IIf(Trim(fgMain.Cell(flexcpText, r, 9)) = "", "Null", "'" & Trim(fgMain.Cell(flexcpText, r, 9)) & "'") & ")"

            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
    Next
    gcnOracle.CommitTrans
    Moded = True
    MsgBox "评分结果保存成功！", vbOKOnly + vbInformation, gstrSysName
    RaiseEvent AferSaveData
    If m_str方式 = "新增" Then
        Call ClearResults
        cmdOK.Enabled = False
        zlControl.TxtSelAll txtNo
        txtNo.SetFocus
        m_bln多次编辑 = True
    Else
        Unload Me
    End If
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 网格编辑完成
'==============================================================================
Private Sub fgMain_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    On Error GoTo errH
    
    If txt评分.Visible Then
        txt评分.Move fgMain.CellLeft, fgMain.CellTop, fgMain.CellWidth - Screen.TwipsPerPixelX, fgMain.CellHeight - Screen.TwipsPerPixelY
    End If
    If lst评分.Visible Then
        lst评分.Move fgMain.CellLeft, fgMain.CellTop, fgMain.CellWidth - Screen.TwipsPerPixelX
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 网格移动变化
'==============================================================================
Private Sub fgMain_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo errH
    
    If txt评分.Visible Then
        txt评分.Move fgMain.CellLeft, fgMain.CellTop, fgMain.CellWidth - Screen.TwipsPerPixelX, fgMain.CellHeight - Screen.TwipsPerPixelY
    End If
    If lst评分.Visible Then
        lst评分.Move fgMain.CellLeft, fgMain.CellTop, fgMain.CellWidth - Screen.TwipsPerPixelX
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 鼠标按下之前变量值初
'==============================================================================
Private Sub fgMain_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    On Error GoTo errH
    edKey = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 网格编辑完成后，动态改变得分和等级
'==============================================================================
Private Sub fgMain_CellChanged(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo errH
    
    If Me.Tag = "" Then Exit Sub
    If fgMain.Rows > 1 Then
        lbl得分 = Get分数
    Else
        lbl得分 = ""
    End If
    
    If fgMain.Rows > 1 Then
        lbl等级 = Get等级
    Else
        lbl等级 = ""
    End If
    If lbl等级 = "不合格" Then
        lbl得分.Visible = False
    Else
        lbl得分.Visible = True
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 取得网格统计的分数
'==============================================================================
Private Function Get分数() As Single
    Dim r               As Long
    Dim SUM项目         As Single
    Dim SUM总分         As Single
    Dim 项目            As String
    Dim Num             As Single
    
    On Error GoTo errH
    
    If fgMain.Rows > 1 Then 项目 = fgMain.Cell(flexcpText, 1, 0)
    For r = 1 To fgMain.Rows - 1
        If 项目 <> fgMain.Cell(flexcpText, r, 0) Then
            If SUM项目 > Val(fgMain.Cell(flexcpText, r - 1, 1)) Then
                SUM项目 = Abs(Val(fgMain.Cell(flexcpText, r - 1, 1)))
            ElseIf Abs(SUM项目) < 0.001 Then
                SUM项目 = 0#
            Else
                SUM项目 = Abs(SUM项目)
            End If
            If Right(lbl分制, 3) = "扣分制" Then
                SUM项目 = Abs(Val(fgMain.Cell(flexcpText, r - 1, 1))) - SUM项目
            End If
            SUM总分 = SUM总分 + SUM项目
            SUM项目 = 0#
        End If
        
        Num = Abs(Val(fgMain.Cell(flexcpText, r, 4)))
        SUM项目 = SUM项目 + CDbl(Num)
        项目 = fgMain.Cell(flexcpText, r, 0)
    Next
    
    If SUM项目 > Val(fgMain.Cell(flexcpText, r - 1, 1)) Then
        SUM项目 = Abs(Val(fgMain.Cell(flexcpText, r - 1, 1)))
    ElseIf Abs(SUM项目) < 0.001 Then
        SUM项目 = 0#
    Else
        SUM项目 = Abs(SUM项目)
    End If
    If Right(lbl分制, 3) = "扣分制" Then
        SUM项目 = Abs(Val(fgMain.Cell(flexcpText, r - 1, 1))) - SUM项目
    End If
    SUM总分 = SUM总分 + SUM项目
    SUM项目 = 0#

    Get分数 = SUM总分
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=功能： 取得网格统计的等级
'==============================================================================
Private Function Get等级() As String
    Dim 等级1           As Long         '甲：3
    Dim 等级2           As Long         '乙：2
    Dim 等级            As Long         '丙：1
    Dim 分数            As Single
    Dim 下值            As Single
    Dim 上值            As Single
    Dim r               As Long
    
    On Error GoTo errH
    
    分数 = Val(lbl得分)
    下值 = Val(Mid(lbl下值, 4))
    上值 = Val(Mid(lbl上值, 4))
    If 分数 < 下值 Then
        等级1 = 1
    ElseIf 分数 < 上值 Then
        等级1 = 2
    Else
        等级1 = 3
    End If
    
    等级2 = 3
    For r = 1 To fgMain.Rows - 1
        If fgMain.Cell(flexcpText, r, 4) = "单项否决" Then
            If fgMain.Cell(flexcpText, r, 10) = "不" Then
                Get等级 = "不合格"
                Exit Function
            ElseIf fgMain.Cell(flexcpText, r, 10) = "乙" Then
                If 等级2 > 2 Then 等级2 = 2
            ElseIf fgMain.Cell(flexcpText, r, 10) = "丙" Then
                If 等级2 > 1 Then 等级2 = 1
            End If
        ElseIf fgMain.Cell(flexcpText, r, 4) = "乙级" Then
            If 等级2 > 2 Then 等级2 = 2
        ElseIf fgMain.Cell(flexcpText, r, 4) = "丙级" Then
            If 等级2 > 1 Then 等级2 = 1
        End If
    Next
    
    '取等级1与等级2的最小值：
    If 等级1 > 等级2 Then
        等级 = 等级2
    Else
        等级 = 等级1
    End If
    
    Select Case 等级
    Case 1
        Get等级 = "丙"
    Case 2
        Get等级 = "乙"
    Case 3
        Get等级 = "甲"
    End Select
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=功能： 网格键盘控制
'==============================================================================
Private Sub fgMain_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errH
    
    If KeyCode = vbKeyEscape Then
        CmdCancel_Click
    End If
    If KeyCode = vbKeyReturn Then
        If Shift = 2 Then
            CmdOK_Click
            Exit Sub
        End If
        Select Case fgMain.Cell(flexcpText, fgMain.Row, 3)
            Case "甲级", "乙级", "丙级", "单项否决"
                fgMain_StartEdit fgMain.Row, fgMain.Col, False
            Case Else
                KeyCode = 0
                If fgMain.Row < fgMain.Rows - 1 Then
                    fgMain.Row = fgMain.Row + 1
                    If fgMain.Row < fgMain.Rows - 3 Then
                        fgMain.ShowCell fgMain.Row + 2, 4
                    Else
                        fgMain.ShowCell fgMain.Row, 4
                    End If
                Else
                    If cmdOK.Enabled Then cmdOK.SetFocus
                End If
        End Select
    ElseIf KeyCode = vbKeyDelete Then
        fgMain.Cell(flexcpText, fgMain.Row, 4) = ""
    End If
    edKey = KeyCode
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 网格键盘控制
'==============================================================================
Private Sub fgMain_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error GoTo errH
    
    If Col = 9 Then Exit Sub
        
    Cancel = True
    edRow = fgMain.Row
    edCol = fgMain.Col
    If edCol = 5 Then Exit Sub
    
    Select Case fgMain.Cell(flexcpText, fgMain.Row, 3)
        Case "甲级", "乙级", "丙级", "单项否决"
            '列表动态改变
            lst评分.Clear
            lst评分.AddItem "1 - 空"
            lst评分.AddItem "2 - " + fgMain.Cell(flexcpText, fgMain.Row, 3)
            txt评分.Text = ""
        
            txt评分.Visible = False
            Select Case fgMain.Cell(flexcpText, fgMain.Row, 4)
                Case "甲级", "乙级", "丙级", "单项否决"
                    lst评分.ListIndex = 1
                Case Else
                    lst评分.ListIndex = 0
            End Select
            lst评分.Move fgMain.CellLeft, fgMain.CellTop + fgMain.CellHeight, fgMain.CellWidth
            lst评分.Visible = True
            lst评分.SetFocus
        Case Else
            txt评分.Move fgMain.CellLeft, fgMain.CellTop, fgMain.CellWidth - Screen.TwipsPerPixelX, fgMain.CellHeight - Screen.TwipsPerPixelY
            txt评分.Text = fgMain.Text
            If edKey >= 96 And edKey <= 105 Then '小键盘
                txt评分.Text = edKey - 96
                txt评分.SelStart = 1
            ElseIf edKey = vbKeyDecimal Then
                txt评分.Text = "."
                txt评分.SelStart = 1
            ElseIf edKey > 32 Then
                txt评分.Text = Chr(edKey)
                txt评分.SelStart = 1
            ElseIf edKey = 32 Then
                txt评分.SelStart = 0
                txt评分.SelStart = 32000
            Else
                txt评分.SelStart = 0
                txt评分.SelLength = 32000
            End If
            
            txt评分.Visible = True
            txt评分.SetFocus
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 空格键处理
'==============================================================================
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo errH
    
    If KeyAscii = vbKeyEscape Then
        CmdCancel_Click
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 左侧位置变化处理
'==============================================================================
Private Sub picLeft_Resize()
On Error Resume Next
    pic病人信息.Move 135, 135
    pic方案信息.Move 135, pic病人信息.Top + pic病人信息.Height + 135
    pic项目信息.Move 135, pic方案信息.Top + pic方案信息.Height + 135, pic项目信息.Width, IIf(imgXMXX.Tag <> "", pic项目信息.Height, Abs(picLeft.ScaleHeight - pic病人信息.Height - pic方案信息.Height - 135 * 4))
    pic病人信息.Cls
    pic病人信息.PaintPicture imgBGBlue.Picture, 0, 0, pic病人信息.Width, 360, 0, 0, imgBGBlue.Width, 360
    pic病人信息.PaintPicture imgBGBlue.Picture, 0, 360, Screen.TwipsPerPixelX, pic病人信息.Height - 360, 0, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
    pic病人信息.PaintPicture imgBGBlue.Picture, pic病人信息.ScaleWidth - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, pic病人信息.Height - 360, imgBGBlue.Width - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
    pic病人信息.PaintPicture imgBGBlue.Picture, 0, pic病人信息.ScaleHeight - Screen.TwipsPerPixelY, pic病人信息.Width, Screen.TwipsPerPixelY, 0, imgBGBlue.Height - Screen.TwipsPerPixelY, imgBGBlue.Width, Screen.TwipsPerPixelY
    
    pic方案信息.Cls
    pic方案信息.PaintPicture imgBG.Picture, 0, 0, pic方案信息.Width, 360, 0, 0, imgBG.Width, 360
    pic方案信息.PaintPicture imgBG.Picture, 0, 360, Screen.TwipsPerPixelX, pic方案信息.Height - 360, 0, 360, Screen.TwipsPerPixelX, imgBG.Height - 360
    pic方案信息.PaintPicture imgBG.Picture, pic方案信息.ScaleWidth - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, pic方案信息.Height - 360, imgBG.Width - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, imgBG.Height - 360
    pic方案信息.PaintPicture imgBG.Picture, 0, pic方案信息.ScaleHeight - Screen.TwipsPerPixelY, pic方案信息.Width, Screen.TwipsPerPixelY, 0, imgBG.Height - Screen.TwipsPerPixelY, imgBG.Width, Screen.TwipsPerPixelY
    pic项目信息.Cls
    
    pic项目信息.PaintPicture imgBG.Picture, 0, 0, pic项目信息.Width, 360, 0, 0, imgBG.Width, 360
    pic项目信息.PaintPicture imgBG.Picture, 0, 360, Screen.TwipsPerPixelX, pic项目信息.Height - 360, 0, 360, Screen.TwipsPerPixelX, imgBG.Height - 360
    pic项目信息.PaintPicture imgBG.Picture, pic项目信息.ScaleWidth - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, pic项目信息.Height - 360, imgBG.Width - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, imgBG.Height - 360
    pic项目信息.PaintPicture imgBG.Picture, 0, pic项目信息.ScaleHeight - Screen.TwipsPerPixelY, pic项目信息.Width, Screen.TwipsPerPixelY, 0, imgBG.Height - Screen.TwipsPerPixelY, imgBG.Width, Screen.TwipsPerPixelY
    imgBRXX.Move pic病人信息.ScaleWidth - imgBRXX.Width - 100
    imgFAXX.Move pic方案信息.ScaleWidth - imgFAXX.Width - 100
    imgXMXX.Move pic项目信息.ScaleWidth - imgXMXX.Width - 100
    Refresh
End Sub

'==============================================================================
'=功能： 右侧位置变化处理
'==============================================================================
Private Sub picRight_Resize()
On Error Resume Next
    With fgMain
        .Height = picRight.ScaleHeight - .Top
        .Width = picRight.ScaleWidth - .Left
    End With
End Sub

'==============================================================================
'=功能： 右侧位置变化处理
'==============================================================================
Private Sub pic项目信息_Resize()
On Error Resume Next
    txt项目信息.Move txt项目信息.Left, txt项目信息.Top, txt项目信息.Width, Abs(pic项目信息.ScaleHeight - txt项目信息.Top - 135)
End Sub

'==============================================================================
'=功能： 值修改后处理
'==============================================================================
Private Sub txtNo_Change()
    On Error GoTo errH
    
    txtNo.Tag = "Changed"
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 值修改后处理
'==============================================================================
Private Sub txtNo_GotFocus()
    On Error GoTo errH
    
    zlControl.TxtSelAll txtNo
    ShowTips picRight, txtNo, "以A或－开头的数字:       病人ID" & vbCrLf & _
        "以B或＋开头的数字:       住院号" & vbCrLf & _
        "以C或／开头的数字:       床位号" & vbCrLf & _
        "以D或＊开头的数字:       门诊号" & vbCrLf & _
        "纯数字:                             就诊卡号" & vbCrLf & _
        "其他情况均视为病人姓名处理。", "快速定位使用技巧" & vbCrLf
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 快速录入按键处理
'==============================================================================
Private Sub txtNo_KeyPress(KeyAscii As Integer)

    Dim StrText         As String
    Dim strTmp          As String
    Dim bytFilterMode   As Byte
    Dim lng病人ID       As Long
    Dim blnCard         As Boolean
    
    On Error GoTo errH
    
    If txtNo.Tag <> "" Then
        '就诊卡号

        blnCard = zlCommFun.InputIsCard(txtNo, KeyAscii, ParamInfo.系统号)
        If blnCard Then
            If Len(txtNo.Text) = ParamInfo.就诊卡号码长度 - 1 And KeyAscii <> 8 Or KeyAscii = 13 And txtNo.Text <> "" Then
                If KeyAscii <> 13 Then
                    txtNo.Text = txtNo.Text & Chr(KeyAscii)
                    txtNo.SelStart = Len(txtNo.Text)
                    KeyAscii = 0
                End If
                
                StrText = txtNo.Text
                bytFilterMode = 1
            End If
        End If
    End If
    
    Select Case KeyAscii
        '------------------------------------------------------------------------------------------------------------------
        Case vbKeyReturn
            KeyAscii = 0
            If txtNo.Tag = "Changed" Then
                If InStr(txtNo.Text, "'") Then
                    ShowSimpleMsg "录入的内容中有非法字符 ' ！"
                    Exit Sub
                End If
                StrText = txtNo.Text
                Select Case UCase(Left(StrText, 1))
                    Case "-", "A"                 '病人id
                        bytFilterMode = 2
                        StrText = IIf(IsNumeric(Mid(StrText, 2)), Val(Mid(StrText, 2)), "0")
                    Case "+", "B"                 '住院号
                        bytFilterMode = 3
                        StrText = IIf(IsNumeric(Mid(StrText, 2)), Val(Mid(StrText, 2)), "0")
                    Case "*", "D"                 '门诊号
                        bytFilterMode = 4
                        StrText = IIf(IsNumeric(Mid(StrText, 2)), Val(Mid(StrText, 2)), "0")
                    Case Else                     '姓名
                        txtNo.Tag = ""
                        zlCommFun.PressKey vbKeyTab
                        Exit Sub
                End Select
            Else
                zlCommFun.PressKey vbKeyTab
            End If
        '------------------------------------------------------------------------------------------------------------------
        Case vbKeyEscape
            Call CmdCancel_Click
        '------------------------------------------------------------------------------------------------------------------
        Case Else
            If Chr(KeyAscii) = "'" Then KeyAscii = 0
            If Chr(KeyAscii) = "|" Then KeyAscii = 0
    End Select
    
    If StrText <> "" And bytFilterMode > 0 Then
        Call 查找病案(StrText, bytFilterMode)
        txtNo.Tag = ""
    End If
    
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub

'==============================================================================
'=功能： 如果输入非法数值，就提示：
'==============================================================================
Private Sub txt评分_Change()
    Dim Num As Single
    On Error GoTo errH
    
    Num = Abs(Val(txt评分.Text))
    If Num > 9999 Then
        txt评分.Text = 9999
    ElseIf InStr(1, fgMain.Cell(flexcpText, fgMain.Row, 3), "/") > 0 Then
        
    ElseIf Num > Val(fgMain.Cell(flexcpText, fgMain.Row, 3)) Then
        txt评分.Text = Val(fgMain.Cell(flexcpText, fgMain.Row, 3))
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 如果输入非法数值，就提示：
'==============================================================================
Private Sub txt评分_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errH
    Select Case KeyCode
        Case vbKeyRight
            If txt评分.SelStart = Len(txt评分.Text) Then
                fgMain.SetFocus
                fgMain.Col = fgMain.Col + 1
            End If
        Case vbKeyLeft
            If txt评分.SelStart = Len(txt评分.Text) Then
                fgMain.SetFocus
                fgMain.Col = fgMain.Col - 1
            End If
        Case vbKeyUp
            fgMain.SetFocus
            fgMain.Row = fgMain.Row - 1
            fgMain.ShowCell fgMain.Row, 4
        Case vbKeyReturn, vbKeyDown
            '如果输入非法数值，就提示：
            If Len(txt评分) > 0 And IsNumeric(txt评分) = False Or Abs(Val(txt评分)) > 9999 Then
                ShowTips picRight, txt评分, "请输入正确的数字", "格式错误", 2000, fgMain.Top
                txt评分.SelStart = 0
                txt评分.SelLength = Len(txt评分)
                txt评分.SetFocus
                Exit Sub
            End If
            
            fgMain.SetFocus
            If fgMain.Row <> fgMain.Rows - 1 Then
                fgMain.Row = fgMain.Row + 1
                fgMain.ShowCell fgMain.Row, 4
            Else
                If cmdOK.Enabled Then cmdOK.SetFocus
            End If
        Case vbKeyEscape
            KeyCode = 0
            txt评分.Text = fgMain.TextMatrix(edRow, edCol)
            fgMain.SetFocus
    End Select
    edKey = KeyCode
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 评分按键控制
'==============================================================================
Private Sub lst评分_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errH
    
    Select Case KeyCode
        Case vbKeyReturn
            fgMain.SetFocus
            If fgMain.Row = fgMain.Rows - 1 Then
                If cmdOK.Enabled Then cmdOK.SetFocus
                Exit Sub
            End If
            fgMain.Row = fgMain.Row + 1
            fgMain.ShowCell fgMain.Row, 4
        Case vbKeyEscape
            KeyCode = 0
            Select Case fgMain.TextMatrix(edRow, edCol)
                Case "甲级", "乙级", "丙级", "单项否决"
                    lst评分.ListIndex = 1
                Case Else
                    lst评分.ListIndex = 0
            End Select
            fgMain.SetFocus
        Case vbKey1, vbKeyNumpad1
            lst评分.ListIndex = 0
            fgMain.SetFocus
            fgMain.Row = fgMain.Row + 1
            fgMain.ShowCell fgMain.Row, 4
        Case vbKey2, vbKeyNumpad2
            lst评分.ListIndex = 1
            fgMain.SetFocus
            If fgMain.Row = fgMain.Rows - 1 Then
                If cmdOK.Enabled Then cmdOK.SetFocus
                Exit Sub
            End If
            fgMain.Row = fgMain.Row + 1
            fgMain.ShowCell fgMain.Row, 4
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 评分按键控制
'==============================================================================
Private Sub txt评分_KeyPress(KeyAscii As Integer)
    On Error GoTo errH
    Select Case KeyAscii
        Case 13, 27:
            KeyAscii = 0
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 评分数据检测有效性
'==============================================================================
Private Sub txt评分_LostFocus()
    Dim Num         As Single
    
    On Error GoTo errH
    
    Num = Abs(Val(txt评分.Text))
    If Num > 9999 Then
        txt评分 = 9999
        Exit Sub
    End If
    If Num < 0.01 Then
        fgMain.TextMatrix(edRow, edCol) = ""
    Else
        If Num < 1 Then
            fgMain.TextMatrix(edRow, edCol) = Format(Num, "0.0")
        Else
            fgMain.TextMatrix(edRow, edCol) = Num
        End If
    End If
    
    txt评分.Visible = False
    edKey = 0
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 显示提示
'==============================================================================
Private Sub ShowTips(ctl0 As Control, ctl As Control, str内容 As String, Optional str标题 As String = "提示信息", Optional lng时间 As Long = 4500, Optional 修正高度 As Long = 0)
    Dim X       As Single
    Dim Y       As Single
    
    On Error GoTo errH
    
    X = (ctl.Left + ctl.Width / 2 + ctl0.Left) / Screen.TwipsPerPixelX
    Y = (ctl.Top + ctl.Height + ctl0.Top + 修正高度) / Screen.TwipsPerPixelY
    If Len(str内容) > 0 Then
        tipPopup1.Hide
        tipPopup1.StandardIcon = IDI_INFORMATION
        tipPopup1.ShowCloseButton = True
        
        tipPopup1.TimeOut = lng时间
        tipPopup1.Title = str标题
        tipPopup1.Text = str内容
        tipPopup1.Show Me.hWnd, X, Y
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 根据项目ID显示项目基本要求
'==============================================================================
Private Sub Show基本要求(lngID As Long, 项目 As String, 标准分值 As String)
    Dim rs              As ADODB.Recordset
    On Error GoTo errH
    
    gstrSQL = "select ID,描述 as 基本要求,上级ID from 病案评分标准 Where ID= [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID)
    
    If Not rs.EOF Then
        If m_lngOldSJID > 0 And m_lngOldSJID = lngID Then Exit Sub
        If IsNull(rs.Fields("基本要求")) Then
            txt项目信息 = "名称：" + 项目 + "  " + IIf(Len(Trim(标准分值)) = 0, "", "(" + 标准分值 + ")")
            txt项目信息 = txt项目信息 + vbCrLf
        Else
            If Len(rs.Fields("基本要求")) > 0 Then
                txt项目信息 = "名称：" + 项目 + "  " + IIf(Len(Trim(标准分值)) = 0, "", "(" + 标准分值 + ")")
                txt项目信息 = txt项目信息 + vbCrLf + rs.Fields("基本要求")
            End If
        End If
    Else
        txt项目信息 = "无":
    End If
    m_lngOldSJID = m_lngCurSJID
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 病案查询
'==============================================================================
Private Sub 查找病案(strID As String, ByVal bytFilterMode As Byte)
    Dim lngBRID         As Long
    Dim lngZYID         As Long
    Dim strSQL          As String
    Dim lngFAID         As Long
    Dim i               As Long
    Dim rs              As ADODB.Recordset
    Dim lngCurRowTMP    As Long
    Dim blnFinded       As Boolean
    
    On Error GoTo errH
    
    
    Select Case bytFilterMode
        Case 1              '就诊卡号
            
            strSQL = _
                "Select A.病人ID,B.主页ID " & _
                " From 病人信息 A,病案主页 B " & _
                " Where A.病人ID=B.病人ID " & _
                " And Nvl(B.主页ID,0)<>0 " & _
                " And A.就诊卡号=[1]"
                
        Case 2              '病人ID
            strSQL = _
                "Select A.病人ID,B.主页ID " & _
                " From 病人信息 A,病案主页 B " & _
                " Where A.病人ID=B.病人ID " & _
                " And Nvl(B.主页ID,0)<>0 " & _
                " And A.病人ID=[1]"
        Case 3              '住院号
            strSQL = _
                "Select A.病人ID,B.主页ID " & _
                " From 病人信息 A,病案主页 B " & _
                " Where A.病人ID=B.病人ID " & _
                " And Nvl(B.主页ID,0)<>0 " & _
                " And A.住院号=[1]"
        Case 4              '门诊号
            strSQL = _
                "Select A.病人ID,B.主页ID " & _
                " From 病人信息 A,病案主页 B " & _
                " Where A.病人ID=B.病人ID " & _
                " And Nvl(B.主页ID,0)<>0 " & _
                " And A.门诊号=[1]"
        Case Else            '姓名
            strSQL = _
                "Select A.病人ID,B.主页ID " & _
                " From 病人信息 A,病案主页 B " & _
                " Where A.病人ID=B.病人ID " & _
                " And Nvl(B.主页ID,0)<>0 " & _
                " And Upper(A.姓名)=[2]"
                
            strID = UCase(strID)
    End Select

    If mbln编目后评分 Then
        gstrSQL = strSQL & " and 编目日期 is not null"
    Else
        gstrSQL = strSQL
    End If
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID), strID)
    If Not rs.EOF Then '找到记录
        lngBRID = rs("病人ID")
'        lngZYID = Rs("主页ID")
        If lngBRID <= 0 Then Exit Sub
        '定位主界面病案记录
        With frm病案评分.fg病案_S
            lngCurRowTMP = .Row
            For i = lngCurRowTMP + 1 To .Rows - 1
                If Val(.Cell(flexcpText, i, 3)) = lngBRID Then
                    .Row = i
                    .ShowCell i, 2
                    blnFinded = True
                    Exit For
                End If
            Next
            If blnFinded = False Then '如果当前行下面没有匹配项，则从第一行开始重新查询。
                For i = 1 To lngCurRowTMP
                    If Val(.Cell(flexcpText, i, 3)) = lngBRID Then
                        .Row = i
                        .ShowCell i, 2
                        blnFinded = True
                        Exit For
                    End If
                Next
            End If
        End With
        
        '填入病案信息
        rs.Close
        If InStr(gstrPrivs, "所有科室") = 0 Then    '无所有科室功能
            gstrSQL = "select 姓名,性别,病人ID,主页ID,住院号,入院日期,出院日期,入院科室,出院科室,门诊医师,责任护士,住院医师,编目日期,结果ID,方案ID,总分,等级,评分人,评分时间,审核人,审核时间,返回修改,备注 " & _
                      "from 病案质量报表视图 where 评分时间 is null and 审核时间 is null and 出院科室 = [1] and 病人ID = [2]"
        Else
            gstrSQL = "select 姓名,性别,病人ID,主页ID,住院号,入院日期,出院日期,入院科室,出院科室,门诊医师,责任护士,住院医师,编目日期,结果ID,方案ID,总分,等级,评分人,评分时间,审核人,审核时间,返回修改,备注 " & _
                      "from 病案质量报表视图 where 评分时间 is null and 审核时间 is null and 病人ID=[2]"
        End If
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, gstrDeptName, lngBRID)

        If Not rs.EOF Then
            lngFAID = NVL(rs("方案ID"), 0)
            lngZYID = rs("主页ID")
            
            m_bln多次编辑 = True
            cmdOK.Enabled = True
            ShowForm "新增", 0, lngBRID, lngZYID, lngFAID, m_lng科室ID
            Exit Sub
        Else
            cmdOK.Enabled = False
        End If
    Else
        cmdOK.Enabled = False
    End If
    MsgBox "没有找到指定病案，或者该病案已经评分！请重新输入。", vbExclamation, gstrSysName
    Call ClearResults
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 清除结果病案查询
'==============================================================================
Private Sub ClearResults()
    Dim i As Long
    
    On Error GoTo errH
    
    lbl姓名.Caption = "姓   名:"
    lbl住院号.Caption = "住 院 号:"
    lbl住院次数.Caption = "住院次数:"
    lbl出院科室.Caption = "出院科室: " & gstrDeptName
    lbl住院医师.Caption = "住院医师:"
    lbl编目日期.Caption = "编目日期:"
    chk返回修改.Value = vbUnchecked
    txtNo.Text = ""
    txt备注.Text = ""
    
    For i = 1 To fgMain.Rows - 1                '清除原来评分结果
        fgMain.Cell(flexcpText, i, 4) = ""
        fgMain.Cell(flexcpText, i, 5) = ""
        fgMain.Cell(flexcpText, i, 9) = ""
    Next
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



