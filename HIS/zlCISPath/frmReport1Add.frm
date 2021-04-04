VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmReport1Add 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " 增加报表"
   ClientHeight    =   1710
   ClientLeft      =   2760
   ClientTop       =   3720
   ClientWidth     =   4470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3240
      TabIndex        =   5
      Top             =   600
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   1100
   End
   Begin VB.Frame fraDate 
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   2950
      Begin MSMask.MaskEdBox txtScope 
         Height          =   375
         Left            =   1275
         TabIndex        =   0
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-mm"
         Mask            =   "####-##"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   1275
         TabIndex        =   2
         Top             =   1140
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
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
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   102694915
         CurrentDate     =   40240
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1275
         TabIndex        =   1
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
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
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   102694915
         CurrentDate     =   40240
      End
      Begin VB.Label lbl期间 
         Caption         =   "统计期间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   300
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结束时间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   840
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开始时间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   6
         Top             =   780
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmReport1Add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public mlng路径ID As Long   'in 传入
Public mstr期间 As String   'out传出
Public mblnOK As Boolean    'out

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String, str期间 As String
    Dim rstmp As ADODB.Recordset
    
    On Error GoTo errH
    If Not IsDate(txtScope.Text) Then
        Call MsgBox("输入的期间不符合规范，要求四位年份+两位月数字。", vbInformation, gstrSysName)
        txtScope.SetFocus
        Exit Sub
    End If
    
    str期间 = Format(txtScope.Text, "yyyymm")
    strSQL = "Select 1 From 路径报表文件 Where 期间 = [1] And 路径ID = [2]"
    Set rstmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, str期间, mlng路径ID)
    If rstmp.RecordCount > 0 Then
        Call MsgBox("该期间已存在，请重新输入一个期间。", vbInformation, gstrSysName)
        Exit Sub
    End If
    
    strSQL = "Zl_路径报表文件_Insert(1," & str期间 & ",To_date('" & dtpBegin.Value & "','yyyy-mm-dd')" & _
            ",To_date('" & dtpEnd.Value & "','yyyy-mm-dd')," & mlng路径ID & ",'" & UserInfo.姓名 & "')"
    Call zldatabase.ExecuteProcedure(strSQL, Me.Caption)
            
    mstr期间 = str期间
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call zlcommfun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    Dim Datsys As Date
    
    Datsys = zldatabase.Currentdate
    txtScope.Text = Format(Datsys, "yyyy-mm")
    dtpBegin.Value = Format(Datsys, "yyyy-mm-01")
    dtpEnd.Value = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(Datsys, "yyyy-mm-01")))))
    
End Sub
