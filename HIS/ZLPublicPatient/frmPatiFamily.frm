VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.8#0"; "zlIDKind.ocx"
Begin VB.Form frmPatiFamily 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "家属关系"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6975
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPatiFamily.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6975
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   5760
      TabIndex        =   6
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   4560
      TabIndex        =   5
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Frame fraSplit1 
      BackColor       =   &H00C0C0C0&
      Height          =   45
      Left            =   0
      TabIndex        =   20
      Top             =   1320
      Width           =   8055
   End
   Begin VB.Frame fraPatiInfo 
      BorderStyle     =   0  'None
      Caption         =   "病人信息"
      Height          =   735
      Left            =   0
      TabIndex        =   8
      Top             =   480
      Width           =   6975
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人类型:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   4920
         TabIndex        =   28
         Tag             =   "性别:"
         Top             =   480
         Width           =   810
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "就诊卡:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2760
         TabIndex        =   27
         Tag             =   "性别:"
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   5280
         TabIndex        =   26
         Tag             =   "性别:"
         Top             =   120
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2940
         TabIndex        =   25
         Tag             =   "性别:"
         Top             =   120
         Width           =   450
      End
      Begin VB.Label lblTag 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Tag             =   "姓名:"
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   300
         TabIndex        =   23
         Tag             =   "姓名:"
         Top             =   120
         Width           =   450
      End
      Begin VB.Label lblJZK 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10101010101"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   3480
         TabIndex        =   22
         Tag             =   "病人类型:"
         Top             =   480
         Width           =   990
      End
      Begin VB.Label lblNum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "201502101"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   840
         TabIndex        =   21
         Tag             =   "姓名:"
         Top             =   480
         Width           =   810
      End
      Begin VB.Label lblPatiType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "普通患者"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   5760
         TabIndex        =   12
         Tag             =   "病人类型:"
         Top             =   480
         Width           =   720
      End
      Begin VB.Label lblAge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "30岁"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   5760
         TabIndex        =   11
         Tag             =   "年龄:"
         Top             =   120
         Width           =   360
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "未知"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   3480
         TabIndex        =   10
         Tag             =   "性别:"
         Top             =   120
         Width           =   360
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "琪玛多吉"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   840
         TabIndex        =   9
         Tag             =   "姓名:"
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame fraSplit2 
      BackColor       =   &H00C0C0C0&
      Height          =   45
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   8055
   End
   Begin VB.Frame fraGroup 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   19
      Top             =   1920
      Width           =   6975
      Begin VSFlex8Ctl.VSFlexGrid vsfamily 
         Height          =   1425
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   6750
         _cx             =   11906
         _cy             =   2514
         Appearance      =   3
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
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPatiFamily.frx":6852
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
         AllowUserFreezing=   2
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin VB.PictureBox picdel 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   6240
            Picture         =   "frmPatiFamily.frx":6913
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   29
            Top             =   360
            Visible         =   0   'False
            Width           =   240
         End
      End
   End
   Begin VB.Frame fraPatiCard 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   495
      Left            =   0
      TabIndex        =   13
      ToolTipText     =   "刷卡并输入密码后,按Enter键提取病人信息"
      Top             =   0
      Width           =   6975
      Begin VB.TextBox txtPatiPWD 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   4920
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   120
         Width           =   1935
      End
      Begin zlIDKind.PatiIdentify PatiIdentifyPati 
         Height          =   375
         Left            =   600
         TabIndex        =   0
         Top             =   120
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmPatiFamily.frx":D165
         BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSize        =   -1  'True
         IDKindAppearance=   2
         InputAppearance =   2
         ShowSortName    =   -1  'True
         DefaultCardType =   "0"
         IDkindBorderStyle=   1
         IDKindWidth     =   555
         BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CardNOForColor  =   -2147483635
         MustBrushCard   =   -1  'True
         AllowAutoCommCard=   -1  'True
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
         BackColor       =   16777215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "密码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   4500
         TabIndex        =   15
         Top             =   210
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "病人"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   210
         Width           =   360
      End
   End
   Begin VB.Frame fraFamilyCard 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   16
      ToolTipText     =   "刷卡并输入密码后,按Enter键录入家属信息"
      Top             =   1440
      Width           =   6975
      Begin VB.TextBox txtFamilyPWD 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   4950
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   0
         Width           =   1935
      End
      Begin zlIDKind.PatiIdentify PatiIdentifyFamily 
         Height          =   375
         Left            =   630
         TabIndex        =   2
         Top             =   0
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmPatiFamily.frx":D22C
         BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSize        =   -1  'True
         IDKindAppearance=   2
         InputAppearance =   2
         ShowSortName    =   -1  'True
         DefaultCardType =   "0"
         IDkindBorderStyle=   1
         IDKindWidth     =   555
         BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CardNOForColor  =   -2147483635
         MustBrushCard   =   -1  'True
         AllowAutoCommCard=   -1  'True
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
      End
      Begin VB.Label lblPWD 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "密码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   4500
         TabIndex        =   18
         Top             =   90
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "家属"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   90
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmPatiFamily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytFunc As Long    '1-查看;2-编辑
Private mbytCount As Byte  '记录密码错误次数
Private mlng病人ID As Long
Private mlngModule As Long
Private mobjKeyboard As Object
Private mblnReturn As Boolean
Private msinTime As Single

Private Type T_Pati
    病人ID As Long
    姓名 As String
    性别 As String
    年龄 As String
    就诊卡号 As String
    密码 As String
End Type

Private mPati As T_Pati
Private mFamily As T_Pati

Private Const C_FamilyColumHeader = "关系,1505,4;姓名,1370,4;性别,705,4;年龄,705,4;就诊卡号,1545,4;详细,495,4; 操作,300,4" '格式:"列名","对齐","列宽"(其中对齐取值为:1-左对齐 4-居中 7-右对齐)
Private Const C_COLOR_灰色 = &H80000000
Private Const C_COLOR_白色 = &H80000005

Public Sub ShowMe(ByRef frmMain As Object, ByVal lng病人ID As Long, ByVal bytFunc As Byte, ByVal lngModul As Long)
'功能:显示主窗体
'参数: objFrmMain-主窗体
'       =1-查看,2-编辑
'      lng病人ID-查看时传人 （bytFunc=1时传入,bytFunc=2时刷卡获取）
'      lngModul 模块号
'     str关系-用于主窗体文本显示
'     rsFamily-用于缓存病人家属
    mbytFunc = bytFunc
    mlngModule = lngModul
    If mbytFunc = 1 Then
        mlng病人ID = lng病人ID
        Me.Caption = "家属信息"
    Else
        mlng病人ID = 0
        Me.Caption = "家属登记"
    End If
    Me.Show 1, frmMain
End Sub

Private Sub cmdCancel_Click()
    Dim i As Long
    
    With vsfamily
        For i = 1 To .Rows - 1
            If InStr(",2,3,4,", "," & .RowData(i) & ",") > 0 Then
                If MsgBox("存在编辑后未保存的信息，您确定要取消？", vbOKCancel + vbQuestion + vbDefaultButton1, gstrSysName) = vbCancel Then
                    Exit Sub
                End If
            End If
        Next
    End With
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Not SavePatiFamily Then Exit Sub
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If mblnReturn Then
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0 '避免刷卡自带回车
            mblnReturn = False
        End If
    End If
End Sub

Private Sub Form_Load()
    '初始化
    
    Call ClearPatiInfo
    Call InitVsFamily
    Call LoadPatiFamily
    picdel.Visible = False
    If mbytFunc = 1 Then
        '查看
        cmdCancel.Caption = "关闭(&C)"
        Call LoadPati
    ElseIf mbytFunc = 2 Then
        '编辑
        Call CreateSquareCardObject(Me, mlngModule)
        If Not gobjSquare Is Nothing Then
            PatiIdentifyPati.MustBrushCard = True   '必须刷卡
            PatiIdentifyPati.OnlyThreeCard = True
            Call PatiIdentifyPati.zlInit(Me, glngSys, mlngModule, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "")
            
            PatiIdentifyFamily.MustBrushCard = True  '必须刷卡
            PatiIdentifyFamily.OnlyThreeCard = True
            Call PatiIdentifyFamily.zlInit(Me, glngSys, mlngModule, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "")
        End If
        '创建外接键盘
        Call CreateObjectKeyboard
        PatiIdentifyFamily.Enabled = False
        txtFamilyPWD.Enabled = False
        PatiIdentifyFamily.BackColor = C_COLOR_灰色
        txtFamilyPWD.BackColor = C_COLOR_灰色
        cmdCancel.Caption = "取消(&C)"
    End If
    
End Sub

Private Sub Form_Resize()
    Dim lngW As Long
    
    If mbytFunc = 1 Then  '查看
        lngW = 6975
        Me.Width = 7065: Me.Height = 3705
        fraPatiCard.Visible = False
        fraPatiInfo.Move 0, 0, lngW, 735
        fraSplit1.Move 0, fraPatiInfo.Top + fraPatiInfo.Height + 45, lngW, 45
        fraFamilyCard.Visible = False
        fraGroup.Move 0, fraSplit1.Top + fraSplit1.Height + 120, lngW, 1575
        fraSplit2.Move 0, fraGroup.Top + fraGroup.Height + 120, lngW, 45
        cmdCancel.Move 5760, 2805, 1095, 350
    ElseIf mbytFunc = 2 Then
        lngW = 6975
        Me.Width = 7065: Me.Height = 4740
        fraPatiCard.Visible = True
        fraFamilyCard.Visible = True
        fraPatiCard.Move 0, 0, lngW, 495
        fraPatiInfo.Move 0, fraPatiCard.Height, lngW, 735
        fraSplit1.Move 0, fraPatiInfo.Top + fraPatiInfo.Height + 45, lngW, 45
        fraFamilyCard.Move 0, fraSplit1.Top + fraSplit1.Height + 120, lngW, 375
        fraGroup.Move 0, fraFamilyCard.Top + fraFamilyCard.Height + 120, lngW, 1575
        fraSplit2.Move 0, fraGroup.Top + fraGroup.Height + 120, lngW, 45
        cmdOK.Move 4560, 3840, 1095, 350
        cmdCancel.Move 5760, 3840, 1095, 350
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjKeyboard = Nothing
End Sub

Private Sub PatiIdentifyFamily_Change()
    If Trim(PatiIdentifyFamily.Text) = "" Then
        txtFamilyPWD.Text = ""
        Call ReSetPati(mFamily)
    End If
End Sub

Private Sub PatiIdentifyFamily_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    If objHisPati Is Nothing Then
        mFamily.病人ID = 0
        blnCancel = True
        mblnReturn = True
        msinTime = Timer
        Do
            If (Timer - msinTime) > 0.25 Then Exit Do
            If mblnReturn Then
                DoEvents
            Else
                Exit Do
            End If
        Loop
        mblnReturn = False
        
        MsgBox "病人信息未找到！可能原因:" & vbCrLf & _
                Space(4) & "(1)当前选择的卡类型【" & PatiIdentifyFamily.GetCurCard.名称 & "】与所持卡的类型不符。" & vbCrLf & _
                Space(4) & "(2)所持卡未绑定病人信息。", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    Else
        mbytCount = 3
        mFamily.病人ID = Val(objHisPati.病人ID)
        mFamily.姓名 = objHisPati.姓名
        mFamily.年龄 = objHisPati.年龄
        mFamily.性别 = objHisPati.性别
        mFamily.就诊卡号 = objHisPati.就诊卡号
        mFamily.密码 = objHisPati.密码
        mblnReturn = True
        msinTime = Timer
        Do
            If (Timer - msinTime) > 0.25 Then Exit Do
            If mblnReturn Then
                DoEvents
            Else
                Exit Do
            End If
        Loop
        mblnReturn = False
        txtFamilyPWD.SetFocus
    End If
    
End Sub

Private Sub PatiIdentifyFamily_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If Val(PatiIdentifyFamily.Tag) <> Index Then
        PatiIdentifyFamily.Tag = Index
        PatiIdentifyFamily.Text = ""
    End If
End Sub

Private Sub PatiIdentifyFamily_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean
    '是否刷卡完成
    blnCard = KeyAscii <> 8 And Len(PatiIdentifyFamily.Text) = PatiIdentifyFamily.GetCurCard.卡号长度 - 1 And PatiIdentifyFamily.SelLength <> Len(PatiIdentifyFamily.Text)
    If KeyAscii = vbKeyReturn Or blnCard Then
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0
            txtFamilyPWD.SetFocus
        End If
    End If
End Sub

Private Sub PatiIdentifyFamily_LostFocus()
    If mFamily.病人ID = 0 And Len(PatiIdentifyFamily.Text) <> 0 Then
        MsgBox "病人信息未找到！可能原因:" & vbCrLf & _
                Space(4) & "(1)当前选择的卡类型【" & PatiIdentifyFamily.GetCurCard.名称 & "】与所持卡的类型不符。" & vbCrLf & _
                Space(4) & "(2)所持卡未绑定病人信息。", vbInformation + vbOKOnly, gstrSysName
        PatiIdentifyFamily.SetFocus
    End If
End Sub

Private Sub PatiIdentifyPati_Change()
    If Trim(PatiIdentifyPati.Text) = "" Then
        mlng病人ID = 0
        mbytCount = 3
        txtPatiPWD.Text = ""
        Call ClearPatiInfo
        Call ReSetPati(mPati)
    End If
End Sub

Private Sub PatiIdentifyPati_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    
    If objHisPati Is Nothing Then
        blnCancel = True
        mlng病人ID = 0  '标记未找到病人
        mblnReturn = True
        msinTime = Timer
        Do
            If (Timer - msinTime) > 0.25 Then Exit Do
            If mblnReturn Then
                DoEvents
            Else
                Exit Do
            End If
        Loop
        mblnReturn = False
        
        MsgBox "病人信息未找到！可能原因:" & vbCrLf & _
            Space(4) & "(1)当前选择的卡类型【" & PatiIdentifyPati.GetCurCard.名称 & "】与所持卡的类型不符。" & vbCrLf & _
            Space(4) & "(2)所持卡未绑定病人信息。", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    Else
        Debug.Print "1"
        mbytCount = 3
        mlng病人ID = objHisPati.病人ID
        mPati.病人ID = objHisPati.病人ID
        mPati.姓名 = objHisPati.姓名
        mPati.密码 = objHisPati.密码
        mPati.年龄 = objHisPati.年龄
        mPati.性别 = objHisPati.性别
        
        txtPatiPWD.SetFocus
        mblnReturn = True
        msinTime = Timer
        Do
            If (Timer - msinTime) > 0.25 Then Exit Do
            If mblnReturn Then
                DoEvents
            Else
                Exit Do
            End If
        Loop
        mblnReturn = False
    End If
    
End Sub

Private Sub PatiIdentifyPati_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If Val(PatiIdentifyPati.Tag) <> Index Then
        PatiIdentifyPati.Tag = Index
        PatiIdentifyPati.Text = ""
    End If
End Sub

Private Sub PatiIdentifyPati_KeyPress(KeyAscii As Integer)
     Dim blnCard As Boolean
    '是否刷卡完成
    mblnReturn = False
    blnCard = KeyAscii <> 8 And Len(PatiIdentifyPati.Text) = PatiIdentifyPati.GetCurCard.卡号长度 - 1 And PatiIdentifyPati.SelLength <> Len(PatiIdentifyPati.Text)
    If KeyAscii = vbKeyReturn Or blnCard Then
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0
            txtPatiPWD.SetFocus
        End If
    End If
End Sub

Private Sub PatiIdentifyPati_LostFocus()
    If mlng病人ID = 0 And Len(PatiIdentifyPati.Text) <> 0 Then
        MsgBox "病人信息未找到！可能原因:" & vbCrLf & _
                Space(4) & "(1)当前选择的卡类型【" & PatiIdentifyPati.GetCurCard.名称 & "】与所持卡的类型不符。" & vbCrLf & _
                Space(4) & "(2)所持卡未绑定病人信息。", vbInformation + vbOKOnly, gstrSysName
        PatiIdentifyPati.SetFocus
    End If
End Sub

Private Sub txtFamilyPWD_GotFocus()
    If PatiIdentifyFamily.Text = "" Then
        MsgBox "请先刷卡再录入密码。", vbInformation, gstrSysName
        If PatiIdentifyFamily.Enabled Then PatiIdentifyFamily.SetFocus
        Exit Sub
    ElseIf Val(mFamily.病人ID) = 0 And PatiIdentifyFamily.Text <> "" Then
        On Error Resume Next
        PatiIdentifyFamily.SetFocus '编译部件后,执行到此处 PatiIdentifyPati.SetFocus后会报错
        Err.Clear: On Error GoTo 0
        Exit Sub
    ElseIf mlng病人ID <> 0 Then
        If mFamily.密码 = "" Then Call txtFamilyPWD_KeyPress(vbKeyReturn): Exit Sub
    End If
End Sub

Private Sub txtFamilyPWD_KeyPress(KeyAscii As Integer)
    Dim strPassWord As String
    
    If KeyAscii = 22 Then
        KeyAscii = 0 '不允许粘贴
    ElseIf InStr(":：;；?？", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0 '去除特殊符号，并且不允许粘贴
    ElseIf KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        strPassWord = gobjCommFun.zlStringEncode(txtFamilyPWD.Text)
        If strPassWord <> mFamily.密码 Then
            If mbytCount = 1 Then
                MsgBox "三次密码输入错误,不能再输入！", vbExclamation, gstrSysName
            Else
                MsgBox "密码输入错误！", vbExclamation, gstrSysName
            End If
            txtFamilyPWD.Text = "": mbytCount = mbytCount - 1
            If mbytCount = 0 Then
                PatiIdentifyFamily.Text = ""
                PatiIdentifyFamily.SetFocus
            ElseIf txtFamilyPWD.Enabled Then
                txtFamilyPWD.SetFocus
            End If
            Exit Sub
        Else
            If ADDFamily Then
                If cmdOK.Enabled = False Then cmdOK.Enabled = True
            Else
                If PatiIdentifyFamily.Enabled Then PatiIdentifyFamily.SetFocus
            End If
            PatiIdentifyFamily.Text = ""
        End If
    End If
End Sub

Private Sub txtPatiPWD_GotFocus()
    If PatiIdentifyPati.Text = "" Then
        MsgBox "请先刷卡再录入密码。", vbInformation, gstrSysName
        PatiIdentifyPati.SetFocus
        Exit Sub
    ElseIf mlng病人ID = 0 And PatiIdentifyPati.Text <> "" Then
        On Error Resume Next
        PatiIdentifyPati.SetFocus '编译部件后,执行到此处 PatiIdentifyPati.SetFocus后会报错
        Err.Clear: On Error GoTo 0
        Exit Sub
    ElseIf mlng病人ID <> 0 Then
        If mPati.密码 = "" Then Call txtPatiPWD_KeyPress(vbKeyReturn): Exit Sub
    End If
    Call gobjControl.TxtSelAll(txtPatiPWD)
    Call OpenPassKeyboard(txtPatiPWD, False)
End Sub

Private Sub txtPatiPWD_KeyPress(KeyAscii As Integer)
    Dim strPassWord As String
    Dim intRet As Integer
    If KeyAscii = 22 Then
        KeyAscii = 0 '不允许粘贴
    ElseIf InStr(":：;；?？", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0 '去除特殊符号，并且不允许粘贴
    ElseIf KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If mblnReturn Then
            mblnReturn = False: Exit Sub
        End If
        strPassWord = gobjCommFun.zlStringEncode(txtPatiPWD.Text)
        If strPassWord <> mPati.密码 Then
            If mbytCount = 1 Then
                MsgBox "三次密码输入错误,不能再输入！", vbExclamation, gstrSysName
            Else
                MsgBox "密码输入错误！", vbExclamation, gstrSysName
            End If
            txtPatiPWD.Text = "": mbytCount = mbytCount - 1
            If mbytCount = 0 Then
                Unload Me '密码错误，可输入2次
            ElseIf txtPatiPWD.Enabled Then
                txtPatiPWD.SetFocus
            End If
            Exit Sub
        Else
            PatiIdentifyPati.Enabled = False
            txtPatiPWD.Enabled = False
            PatiIdentifyPati.BackColor = C_COLOR_灰色
            txtPatiPWD.BackColor = C_COLOR_灰色
            PatiIdentifyFamily.Enabled = True
            txtFamilyPWD.Enabled = True
            PatiIdentifyFamily.BackColor = C_COLOR_白色
            txtFamilyPWD.BackColor = C_COLOR_白色
            Call LoadPati
            Call LoadPatiFamily
            If PatiIdentifyFamily.Enabled Then PatiIdentifyFamily.SetFocus
        End If
    End If
End Sub

Private Sub txtPatiPWD_LostFocus()
    Call ClosePassKeyboard(txtPatiPWD)
End Sub

Private Sub vsfamily_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfamily
        If .Col = .ColIndex("关系") Then
            If .TextMatrix(Row, Col) <> Nvl(.Cell(flexcpData, Row, Col)) And CByte(Nvl(.RowData(Row))) = 1 Then
                .RowData(Row) = 3 '更新
                If cmdOK.Enabled = False Then cmdOK.Enabled = True
            ElseIf .TextMatrix(Row, Col) = Nvl(.Cell(flexcpData, Row, Col)) And CByte(Nvl(.RowData(Row))) = 3 Then
                .RowData(Row) = 1 '未更新
            End If
        End If
    End With
End Sub

Private Sub vsfamily_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfamily
        If Not Me.Visible Then Exit Sub
        If (OldRow <> NewRow Or OldRow = NewRow And OldRow = 1) And NewRow > .FixedRows - 1 Then
            If mbytFunc = 2 Then
                If Nvl(.Cell(flexcpData, NewRow, .ColIndex("姓名"))) = "" Then Exit Sub
                If Me.Visible Then
                    If picdel.Visible = False Then picdel.Visible = True
                End If
                picdel.Top = .Cell(flexcpTop, NewRow, .ColIndex("操作"))
                picdel.Left = .Cell(flexcpLeft, NewRow, .ColIndex("操作"))
            Else
                picdel.Visible = False
            End If
        End If
    End With
End Sub

Private Sub vsfamily_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsfamily.ColIndex("关系") = Col And mbytFunc = 2 And CLng(vsfamily.Cell(flexcpData, Row, vsfamily.ColIndex("姓名"))) <> 0 Then
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

Private Sub vsfamily_Click()
    With vsfamily
        If .Row < 0 Or .Col < 0 Then Exit Sub
        If CLng(.Cell(flexcpData, .Row, .ColIndex("姓名"))) = 0 Then Exit Sub
        If .ColIndex("详细") = .Col And .TextMatrix(.Row, .Col) = "详细" Then
            frmDegreeCard.mlng病人ID = CLng(.Cell(flexcpData, .Row, .ColIndex("姓名")))
            frmDegreeCard.mlng主页ID = 0
            frmDegreeCard.Show 1, Me
        End If
    End With
End Sub

Private Sub vsfamily_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long, lngCol As Long

    With vsfamily
        lngRow = .MouseRow: lngCol = .MouseCol
        If lngRow <= 0 Then Exit Sub
        If .ColData(.ColIndex("详细")) > .Rows - 1 Then .ColData(.ColIndex("详细")) = 0
        If lngCol = .ColIndex("详细") And lngRow = .ColData(.ColIndex("详细")) Then
            .Cell(flexcpFontUnderline, lngRow, lngCol) = True
            .ColData(lngCol) = lngRow
        ElseIf lngCol <> .ColIndex("详细") Or .ColData(.ColIndex("详细")) <> lngRow Then
            .Cell(flexcpFontUnderline, .ColData(.ColIndex("详细")), .ColIndex("详细")) = False
            .ColData(lngCol) = lngRow
        End If
    End With
End Sub

Private Sub InitVsFamily()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化VSGrid控件
    '编制:56599
    '日期:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    On Error GoTo ErrHand
    If mbytFunc = 2 Then
        strSQL = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 社会关系 Order by 编码"
        Call gobjDatabase.OpenRecordset(rsTemp, strSQL, "社会关系")
    
        With rsTemp
            Do While Not rsTemp.EOF
                strTmp = strTmp & "|" & Nvl(rsTemp!名称)
            rsTemp.MoveNext
            Loop
        End With
        If Left(strTmp, 1) = "|" Then strTmp = Mid(strTmp, 2)
    End If
    
    With vsfamily
        .Rows = 2
        gobjGrid.Init vsfamily, C_FamilyColumHeader
        .Editable = flexEDKbdMouse
        .SelectionMode = flexSelectionFree
        If mbytFunc = 1 Then
            .ColHidden(.ColIndex("操作")) = True
        ElseIf strTmp <> "" And mbytFunc = 2 Then
            .ColComboList(.ColIndex("关系")) = strTmp
        End If
    End With
    
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub picdel_Click()
    Dim lngRow As Long
    Dim strSQL As String
    Dim i As Long
    Dim lngFlag As Long
    
    With vsfamily
        If MsgBox("您确定要删除本行？", vbQuestion + vbOKCancel + vbDefaultButton2, gstrSysName) = vbOK Then
            If InStr(",1,3,", "," & .RowData(.Row) & ",") > 0 Then '原始,修改
                '不管有无费用都假删除,避免删除时费用产生数据（并发操作）
                .RowData(.Row) = 4    '标记假删除
                .RowHidden(.Row) = True
                If cmdOK.Enabled = False Then cmdOK.Enabled = True
            Else
                .RemoveItem .Row
            End If
            
            picdel.Visible = False
            
            For i = 1 To .Rows - 1
                If .RowHidden(i) = False Then
                    Exit For
                Else
                    lngFlag = lngFlag + 1
                End If
            Next
            If lngFlag = .Rows - 1 Then .Rows = .Rows + 1 '缺省显示一行
        End If
       
    End With
End Sub

Private Sub LoadPati()
'功能:加载病人信息
    '病人家属
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo errH

    strSQL = "Select a.门诊号, a.住院号, a.就诊卡号, a.姓名, a.性别, a.年龄, a.病人类型 From 病人信息 A Where a.病人id = [1]"

    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "病人家属", mlng病人ID)
    
    If rsTmp.RecordCount > 0 Then
        lblName.Caption = rsTmp!姓名
        lblAge.Caption = rsTmp!年龄 & ""
        lblSex.Caption = rsTmp!性别 & ""
        If rsTmp!住院号 & "" <> "" Then
            lblTag.Caption = "住院号:"
            lblNum.Caption = rsTmp!住院号 & ""
        Else
            lblTag.Caption = "门诊号:"
            lblNum.Caption = rsTmp!门诊号 & ""
        End If
        lblJZK.Caption = rsTmp!就诊卡号 & ""
        lblPatiType.Caption = rsTmp!病人类型 & ""
    End If
    
    Exit Sub
errH:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Sub


Private Sub LoadPatiFamily()
'功能:加载病人家属信息
    '病人家属
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo errH
    
    If mlng病人ID <> 0 Then
        strSQL = "Select a.家属ID, a.关系, b.就诊卡号, b.姓名, b.年龄, b.性别,1 as 状态 " & vbNewLine & _
                "From 病人家属 A, 病人信息 B" & vbNewLine & _
                "Where a.家属id = b.病人id And a.病人id = [1] And A.撤档时间 IS NULL"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "病人家属", mlng病人ID)
    End If
    With vsfamily
       .Rows = 2 '缺省显示一行
        If rsTmp Is Nothing Then Exit Sub
        For i = 1 To rsTmp.RecordCount
            .Rows = i + 1
            .TextMatrix(i, .ColIndex("关系")) = rsTmp!关系 & ""
            .TextMatrix(i, .ColIndex("姓名")) = rsTmp!姓名 & ""
            .TextMatrix(i, .ColIndex("年龄")) = rsTmp!年龄 & ""
            .TextMatrix(i, .ColIndex("性别")) = rsTmp!性别 & ""
            .TextMatrix(i, .ColIndex("就诊卡号")) = rsTmp!就诊卡号 & ""
            .TextMatrix(i, .ColIndex("详细")) = "详细"
            .RowData(i) = rsTmp!状态 & "" '1-原始加载
            
            .Cell(flexcpData, i, .ColIndex("关系")) = rsTmp!关系 & ""
            .Cell(flexcpData, i, .ColIndex("姓名")) = rsTmp!家属ID & ""
            .Cell(flexcpForeColor, .Rows - 1, .ColIndex("详细")) = &HC00000
            rsTmp.MoveNext
        Next
    End With
    
    Exit Sub
errH:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Sub

Private Function SavePatiFamily() As Boolean
'功能:保存病人家属信息
'
    Dim strSQL As String
    Dim i As Long
    Dim strDate As String
    Dim strDateDel As String
    Dim addDate As Date
    Dim blnSave As Boolean    '标记是否有有效操作
    
    Dim colSQL As Collection
    
    On Error GoTo errH
    addDate = gobjDatabase.Currentdate
    strDate = Format(addDate, "YYYY-MM-DD HH:MM:SS")
    Set colSQL = New Collection

    With vsfamily
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("关系")) = "" And CLng(.Cell(flexcpData, i, .ColIndex("姓名"))) <> 0 Then
                    MsgBox "该病人家属【" & .TextMatrix(i, .ColIndex("姓名")) & "】与病人【" & lblName.Caption & "】的关系未录入,请先录入后再保存。", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
                    .Row = i: .Col = .ColIndex("关系")
                    .SetFocus
                    .ShowCell .Row, .Col
                    Exit Function
                ElseIf InStr(",2,3,4,", "," & .RowData(i) & ",") > 0 Then
                    blnSave = True   '存在保存项目
                End If
            Next
            
            If Not blnSave Then
                If MsgBox("当前未更新任何家属信息，是否退出？", vbYesNo + vbInformation + vbDefaultButton1, gstrSysName) = vbYes Then
                    SavePatiFamily = True
                Else
                    SavePatiFamily = False
                End If
                Exit Function
            End If
            
            For i = 1 To .Rows - 1
                If CByte(Nvl(.RowData(i))) = 2 Then   '新增
                    strSQL = " Zl_病人家属_Update(1," & mlng病人ID & "," & .Cell(flexcpData, i, .ColIndex("姓名")) & ",'" & UserInfo.姓名 & _
                             "',To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS'),'" & .TextMatrix(i, .ColIndex("关系")) & "')"          '新增
                    colSQL.Add strSQL, "_" & colSQL.Count
                ElseIf CByte(Nvl(.RowData(i))) = 3 Then  '更新
                    strSQL = " Zl_病人家属_Update(2," & mlng病人ID & "," & .Cell(flexcpData, i, .ColIndex("姓名")) & ",'',NULL,'" & .TextMatrix(i, .ColIndex("关系")) & "')"                             '更新
                    colSQL.Add strSQL, "_" & colSQL.Count
                ElseIf CByte(Nvl(.RowData(i))) = 4 And .RowHidden(i) = True Then '假删除
                    '多次删除时增加一秒避免循环删除时违反唯一约束
                    addDate = addDate + 1 / 24 / 60 / 60
                    strDateDel = "To_Date('" & Format(addDate, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')"
                    strSQL = "Zl_病人家属_Update(3," & mlng病人ID & "," & .Cell(flexcpData, i, .ColIndex("姓名")) & ",'',NULL,NULL,'" & UserInfo.姓名 & "'," & strDateDel & ")"
                    colSQL.Add strSQL, "_" & colSQL.Count
                End If
            Next
        End If
    End With
    
    '批量数据提交
    For i = 1 To colSQL.Count
        Call gobjDatabase.ExecuteProcedure(CStr(colSQL(i)), "家属信息")
    Next
    
    SavePatiFamily = True
    Exit Function
errH:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Private Function CreateObjectKeyboard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建密码创建
    '返回:创建成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-24 23:59:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    Set mobjKeyboard = CreateObject("zl9Keyboard.clsKeyboard")
    If Err <> 0 Then Exit Function
    Err = 0
    CreateObjectKeyboard = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function OpenPassKeyboard(ctlText As Control, Optional bln确认密码 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否者False
    '编制:刘兴洪
    '日期:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText, bln确认密码) = False Then Exit Function
    OpenPassKeyboard = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ClosePassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否者False
    '编制:刘兴洪
    '日期:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.ColsePassKeyoardInput(Me, ctlText) = False Then Exit Function
    ClosePassKeyboard = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ADDFamily() As Boolean
    Dim i As Long
    
    With vsfamily
         '本身不能作为家属关联自己
        If mFamily.病人ID & "" = mlng病人ID & "" Then
            MsgBox "病人家属【" & mFamily.姓名 & "】不能是自己,不允许录入！", vbInformation, gstrSysName
            Exit Function
        End If
            
        For i = .FixedRows To .Rows - 1
            '检查 同一个病人不允许多次录入
            If mFamily.病人ID & "" = .Cell(flexcpData, i, .ColIndex("姓名")) & "" Then
                If .RowHidden(i) = False Then
                    MsgBox "该病人家属【" & mFamily.姓名 & "】已经录入,不允许重复录入！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            '新增家属已经被删除
        Next
        
        If .TextMatrix(.Rows - 1, .ColIndex("姓名")) <> "" Then .Rows = .Rows + 1
        .Cell(flexcpData, .Rows - 1, .ColIndex("姓名")) = mFamily.病人ID
        .TextMatrix(.Rows - 1, .ColIndex("姓名")) = mFamily.姓名
        .TextMatrix(.Rows - 1, .ColIndex("性别")) = mFamily.性别
        .TextMatrix(.Rows - 1, .ColIndex("年龄")) = mFamily.年龄
        .TextMatrix(.Rows - 1, .ColIndex("详细")) = "详细"
        .TextMatrix(.Rows - 1, .ColIndex("就诊卡号")) = mFamily.就诊卡号

        .RowData(.Rows - 1) = 2 '2-新增
        .Cell(flexcpForeColor, .Rows - 1, .ColIndex("详细")) = &HC00000
        .ShowCell .Rows - 1, .ColIndex("关系") '显示增加行
    End With
    ADDFamily = True
End Function

Private Sub ClearPatiInfo()

    lblName.Caption = ""
    lblAge.Caption = ""
    lblSex.Caption = ""
    lblTag.Caption = "住院号:"
    lblNum.Caption = ""
    lblJZK.Caption = ""
    lblPatiType.Caption = ""
    cmdOK.Enabled = False
End Sub

Private Sub ReSetPati(udtPati As T_Pati)
    udtPati.病人ID = 0
    udtPati.就诊卡号 = ""
    udtPati.密码 = ""
    udtPati.年龄 = ""
    udtPati.性别 = ""
    udtPati.姓名 = ""
End Sub
