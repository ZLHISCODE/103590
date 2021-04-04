VERSION 5.00
Begin VB.Form frmAppforCritical 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "危急值处理"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12885
   Icon            =   "frmAppforCritical.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   12885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox PicTop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6435
      Left            =   30
      ScaleHeight     =   6405
      ScaleWidth      =   12795
      TabIndex        =   0
      Top             =   15
      Width           =   12825
      Begin VB.CommandButton cmdSend 
         Caption         =   "发送"
         Height          =   375
         Left            =   9300
         TabIndex        =   38
         Top             =   5820
         Width           =   1305
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "退出"
         Height          =   375
         Left            =   10980
         TabIndex        =   37
         Top             =   5820
         Width           =   1305
      End
      Begin VB.TextBox txtReturn 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   1290
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   3840
         Width           =   11310
      End
      Begin VB.TextBox txtIPTime 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   10740
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txtStartSampleNO 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1140
         Width           =   1815
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1710
         Width           =   1815
      End
      Begin VB.TextBox TxtBad 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txtAppforDoctor 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   7650
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1710
         Width           =   1815
      End
      Begin VB.TextBox txtVerifyDoctor 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   645
         Width           =   1815
      End
      Begin VB.TextBox txtStartAge 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   10740
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1140
         Width           =   1815
      End
      Begin VB.TextBox txtRemark 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   3345
         Width           =   11310
      End
      Begin VB.TextBox TxtHealthNo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   4470
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1710
         Width           =   1815
      End
      Begin VB.TextBox TxtName 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   4470
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1140
         Width           =   1815
      End
      Begin VB.TextBox txtSimpleTyppe 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   4470
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txtPatientType 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   10740
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1710
         Width           =   1815
      End
      Begin VB.TextBox txtSex 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   7650
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1140
         Width           =   1815
      End
      Begin VB.TextBox txtIPName 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   7650
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txtSay 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   2835
         Width           =   11310
      End
      Begin VB.TextBox txtVerifyTime 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   10740
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   645
         Width           =   1815
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "处理措施"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   36
         Top             =   3840
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "通知时间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9690
         TabIndex        =   34
         Top             =   2295
         Width           =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   13
         X1              =   10740
         X2              =   12570
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检验日期"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9690
         TabIndex        =   32
         Top             =   660
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "标 本 号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   31
         Top             =   1155
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病 历 号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   30
         Top             =   1725
         Width           =   960
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床    号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   29
         Top             =   2295
         Width           =   960
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检验医师"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   28
         Top             =   660
         Width           =   960
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年    龄"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9690
         TabIndex        =   27
         Top             =   1155
         Width           =   960
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性    别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6540
         TabIndex        =   26
         Top             =   1155
         Width           =   960
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开单时间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9690
         TabIndex        =   25
         Top             =   1725
         Width           =   960
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "标本类型"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3390
         TabIndex        =   24
         Top             =   2295
         Width           =   960
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   23
         Top             =   165
         Width           =   120
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开单医师"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6540
         TabIndex        =   22
         Top             =   1725
         Width           =   960
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "备    注"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   21
         Top             =   3330
         Width           =   960
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "临床科室"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3390
         TabIndex        =   20
         Top             =   1725
         Width           =   960
      End
      Begin VB.Label lblinto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "危急值处理"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   570
         Index           =   0
         Left            =   5355
         TabIndex        =   19
         Top             =   90
         Width           =   2175
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人姓名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3390
         TabIndex        =   18
         Top             =   1155
         Width           =   960
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "通知技师"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6540
         TabIndex        =   17
         Top             =   2295
         Width           =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   0
         X1              =   1260
         X2              =   3090
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   3
         X1              =   4470
         X2              =   6300
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   1
         X1              =   1260
         X2              =   3090
         Y1              =   945
         Y2              =   945
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   2
         X1              =   7650
         X2              =   9480
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   4
         X1              =   10740
         X2              =   12570
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   5
         X1              =   1260
         X2              =   3090
         Y1              =   2010
         Y2              =   2010
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   6
         X1              =   10740
         X2              =   12570
         Y1              =   2010
         Y2              =   2010
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   7
         X1              =   7650
         X2              =   9480
         Y1              =   2010
         Y2              =   2010
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   8
         X1              =   4470
         X2              =   6300
         Y1              =   2010
         Y2              =   2010
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   9
         X1              =   1260
         X2              =   3090
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   10
         X1              =   4470
         X2              =   6300
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   11
         X1              =   7650
         X2              =   9480
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   12
         X1              =   1260
         X2              =   12540
         Y1              =   3645
         Y2              =   3645
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "通知内容"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   16
         Top             =   2835
         Width           =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   15
         X1              =   1260
         X2              =   12600
         Y1              =   3135
         Y2              =   3135
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   16
         X1              =   10740
         X2              =   12570
         Y1              =   945
         Y2              =   945
      End
   End
End
Attribute VB_Name = "frmAppforCritical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnSend As Boolean         '是否发送成功
Private mstrDicName As String       '操作技师
Private mstrReturn As String         '返回医生填写信息

Public Function ShowMe(objFrm As Object, ByVal strDicName As String, ByVal lngSampleID As Long, Optional ByRef strReturn As String) As Boolean
    mstrReturn = ""
    mstrDicName = strDicName
    Call InitData(lngSampleID)
    Me.Show 1, objFrm
    '如果发送失败,则strReturn返回空
    If mblnSend = False Then
        strReturn = ""
    Else
        strReturn = mstrReturn
    End If
    ShowMe = mblnSend
End Function

Private Sub InitData(ByVal lngSampleID As Long, Optional strErr As String)
          '初始化数据
          'lngSampleID        标本ID
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          
1         On Error GoTo InitData_Error

2         strSQL = "Select Distinct a.ID 标本ID, a.标本序号, a.病历号, a.申请科室, a.姓名, a.性别, a.年龄," & _
                  "a.病人来源, a.申请时间 开单时间, a.标本类型, a.床号, a.检验人, a.报告时间," & _
                  "a.申请人 开单医生,B.通知内容 , b.通知人 通知技师,b.通知时间, b.备注 技师备注" & _
                  " From 检验报告记录 A, 检验危急值记录 B" & _
                  " Where  a.Id = b.标本id and a.Id =[1]"
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "危急通知查询", lngSampleID)
4         If rsTmp.RecordCount > 0 Then
5             txtVerifyDoctor = rsTmp("检验人") & ""
6             txtVerifyTime = rsTmp("报告时间") & ""
7             txtStartSampleNO = rsTmp("标本序号") & ""
8             txtStartSampleNO.Tag = rsTmp("标本ID") & ""
9             txtName = rsTmp("姓名") & ""
10            If Val("" & rsTmp!性别) = 1 Then
11                txtSex = "男"
12            ElseIf Val("" & rsTmp!性别) = 2 Then
13               txtSex = "女"
14            Else
15                txtSex = "未知"
16            End If
17            txtStartAge = rsTmp("年龄") & ""
18            txtID = rsTmp("病历号") & ""
19            TxtHealthNo = rsTmp("申请科室") & ""
20            txtAppforDoctor = rsTmp("开单医生") & ""
21            txtPatientType = rsTmp("开单时间") & ""
22            TxtBad = rsTmp("床号") & ""
23            txtSimpleTyppe = rsTmp("标本类型") & ""
24            txtSay = rsTmp("通知内容") & ""
25            txtIPName = rsTmp("通知技师") & ""
26            txtIPTime = rsTmp("通知时间") & ""
27            txtRemark = rsTmp("技师备注") & ""
28        End If

29        rsTmp.Close
30        Set rsTmp = Nothing
          
31        Exit Sub
InitData_Error:
32        mblnSend = False
33        Call WriteErrLog("zlPublicHisCommLis", "frmAppforCritical", "执行(InitData)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
34        Err.Clear
End Sub

Private Sub cmdExit_Click()
    mblnSend = False
    Unload Me
End Sub

Private Sub cmdSend_Click()
          '发送反馈到LIS
          Dim strSQL As String
          Dim strTime As String
          Dim lngSampleID As String
          Dim strInfo As String           '医生填写处理措施
              
1         On Error GoTo cmdSend_Click_Error

2         lngSampleID = Val(Trim(txtStartSampleNO.Tag))           '获取标本号
3         strTime = Format(Currentdate, "yyyy-mm-dd hh:mm:ss")    '服务器时间
4         strInfo = Trim(Me.txtReturn.Text)
              
5         gcnLisOracle.BeginTrans
           '检验消息提醒
6         strSQL = "Zl_检验消息记录_Edit(1,2,null," & lngSampleID & ",null,null,null,'医生已查阅危急值','危急值')"
7         Call ComExecuteProc(Sel_Lis_DB, strSQL, "检验消息记录")
          
8         strSQL = "Zl_检验危急值记录_Message(" & lngSampleID & ",'" & mstrDicName & "',to_date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),'" & strInfo & "')"
9         Call ComExecuteProc(Sel_Lis_DB, strSQL, "危急值记录通知")
10        gcnLisOracle.CommitTrans
          
11        SaveDBLog 18, 6, Val(lngSampleID), "危急值处理", "确认处理，确认人：" & mstrDicName & " 确认时间:" & Format(strTime, "yyyy-MM-dd HH:mm:ss") & " 处理措施:" & strInfo, 2500, "临床实验室管理"
          
12        If strInfo <> "" Then mstrReturn = strInfo
          
13        mblnSend = True
14        MsgBox "发送成功", vbInformation, Me.Caption
15        Unload Me

16        Exit Sub
cmdSend_Click_Error:
17        gcnLisOracle.RollbackTrans
18        mblnSend = False
19        Call WriteErrLog("zlPublicHisCommLis", "frmAppforCritical", "执行(cmdSend_Click)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
20        Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrDicName = ""
End Sub
