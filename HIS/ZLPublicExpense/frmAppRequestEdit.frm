VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmAppRequestEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ԤԼ�Ǽ�"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7980
   Icon            =   "frmAppRequestEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   5250
      TabIndex        =   17
      Top             =   4530
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   6450
      TabIndex        =   18
      Top             =   4530
      Width           =   1100
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      Height          =   3195
      Left            =   30
      ScaleHeight     =   3135
      ScaleWidth      =   7830
      TabIndex        =   27
      Top             =   1200
      Width           =   7890
      Begin VB.CommandButton cmdAll 
         Height          =   315
         Left            =   3810
         Picture         =   "frmAppRequestEdit.frx":06EA
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "���к�Դ"
         Top             =   60
         Width           =   315
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   330
         Left            =   1050
         TabIndex        =   15
         Top             =   2280
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   93323267
         CurrentDate     =   42398
      End
      Begin VB.TextBox txtStyle 
         Height          =   330
         Index           =   5
         Left            =   4995
         TabIndex        =   16
         Top             =   2280
         Width           =   585
      End
      Begin VB.TextBox txt�Ǽ�ʱ�� 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4995
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   2700
         Width           =   2700
      End
      Begin VB.TextBox txt�Ǽ��� 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   2700
         Width           =   1815
      End
      Begin VB.ComboBox cboNote 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1050
         TabIndex        =   14
         Top             =   1860
         Width           =   6675
      End
      Begin VB.Frame Frame2 
         Caption         =   "������Ϣ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   165
         TabIndex        =   31
         Top             =   555
         Width           =   7560
         Begin MSComCtl2.UpDown udStyle 
            Height          =   330
            Index           =   4
            Left            =   3316
            TabIndex        =   36
            Top             =   690
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   582
            _Version        =   393216
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtStyle(4)"
            BuddyDispid     =   196612
            BuddyIndex      =   4
            OrigLeft        =   3315
            OrigTop         =   675
            OrigRight       =   3570
            OrigBottom      =   1005
            Max             =   99
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udStyle 
            Height          =   330
            Index           =   0
            Left            =   1111
            TabIndex        =   35
            Top             =   690
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   582
            _Version        =   393216
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtStyle(0)"
            BuddyDispid     =   196612
            BuddyIndex      =   0
            OrigLeft        =   1110
            OrigTop         =   690
            OrigRight       =   1365
            OrigBottom      =   1020
            Max             =   99
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udStyle 
            Height          =   330
            Index           =   3
            Left            =   5746
            TabIndex        =   34
            Top             =   285
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   582
            _Version        =   393216
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtStyle(3)"
            BuddyDispid     =   196612
            BuddyIndex      =   3
            OrigLeft        =   5745
            OrigTop         =   225
            OrigRight       =   6000
            OrigBottom      =   555
            Max             =   99
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udStyle 
            Height          =   330
            Index           =   2
            Left            =   3316
            TabIndex        =   33
            Top             =   285
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   582
            _Version        =   393216
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtStyle(2)"
            BuddyDispid     =   196612
            BuddyIndex      =   2
            OrigLeft        =   3315
            OrigTop         =   255
            OrigRight       =   3570
            OrigBottom      =   585
            Max             =   99
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtStyle 
            Height          =   330
            Index           =   4
            Left            =   3000
            TabIndex        =   13
            Top             =   690
            Width           =   315
         End
         Begin VB.TextBox txtStyle 
            Height          =   330
            Index           =   0
            Left            =   795
            TabIndex        =   12
            Top             =   690
            Width           =   315
         End
         Begin VB.TextBox txtStyle 
            Height          =   330
            Index           =   3
            Left            =   5430
            TabIndex        =   10
            Top             =   285
            Width           =   315
         End
         Begin VB.TextBox txtStyle 
            Height          =   330
            Index           =   2
            Left            =   3000
            TabIndex        =   8
            Top             =   285
            Width           =   315
         End
         Begin MSComCtl2.UpDown udStyle 
            Height          =   330
            Index           =   1
            Left            =   1111
            TabIndex        =   32
            Top             =   285
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   582
            _Version        =   393216
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtStyle(1)"
            BuddyDispid     =   196612
            BuddyIndex      =   1
            OrigLeft        =   1111
            OrigTop         =   285
            OrigRight       =   1366
            OrigBottom      =   615
            Max             =   99
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtStyle 
            Height          =   330
            Index           =   1
            Left            =   795
            TabIndex        =   6
            Top             =   285
            Width           =   315
         End
         Begin VB.OptionButton optStyle 
            Caption         =   "��      ���Ƴ�(ÿ���Ƴ�      ��)����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   300
            TabIndex        =   11
            Top             =   713
            Width           =   5085
         End
         Begin VB.OptionButton optStyle 
            Caption         =   "��      �����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   4950
            TabIndex        =   9
            Top             =   308
            Width           =   2160
         End
         Begin VB.OptionButton optStyle 
            Caption         =   "��      �ܺ���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   2520
            TabIndex        =   7
            Top             =   308
            Width           =   2175
         End
         Begin VB.OptionButton optStyle 
            Caption         =   "��      ���º���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   285
            TabIndex        =   5
            Top             =   308
            Width           =   2340
         End
      End
      Begin VB.TextBox txtDept 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5415
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   60
         Width           =   2310
      End
      Begin VB.ComboBox cboArrangeNo 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   60
         Width           =   3240
      End
      Begin MSComCtl2.UpDown udStyle 
         Height          =   330
         Index           =   5
         Left            =   5580
         TabIndex        =   44
         Top             =   2280
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   582
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtStyle(5)"
         BuddyDispid     =   196612
         BuddyIndex      =   5
         OrigLeft        =   3315
         OrigTop         =   675
         OrigRight       =   3570
         OrigBottom      =   1005
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "�Ǽ�ʱ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4110
         TabIndex        =   43
         Top             =   2760
         Width           =   840
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "�Ǽ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   330
         TabIndex        =   41
         Top             =   2760
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "��������         ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4110
         TabIndex        =   39
         Top             =   2340
         Width           =   1995
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   38
         Top             =   2340
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "����˵��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   37
         Top             =   1920
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4980
         TabIndex        =   30
         Top             =   120
         Width           =   420
      End
      Begin VB.Label lblArrangeNO 
         AutoSize        =   -1  'True
         Caption         =   "�ű�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   420
      End
   End
   Begin VB.TextBox txtBirth 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   750
      Width           =   2670
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   24
      Top             =   615
      Width           =   15000
   End
   Begin VB.TextBox txtPatient 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1200
      TabIndex        =   2
      Top             =   180
      Width           =   2670
   End
   Begin VB.TextBox txtClinic 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4740
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   750
      Width           =   3195
   End
   Begin VB.TextBox txtAge 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6345
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   180
      Width           =   1590
   End
   Begin VB.TextBox txtGender 
      BackColor       =   &H8000000F&
      Height          =   330
      Left            =   4740
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   930
   End
   Begin zlIDKind.IDKindNew IDKind 
      Height          =   330
      Left            =   660
      TabIndex        =   1
      Top             =   180
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      Appearance      =   2
      IDKindStr       =   "��|��������￨|0|0|0|0|0|;ҽ|ҽ����|0|0|0|0|0|;��|���֤��|1|0|0|0|0|;��|�����|0|0|0|0|0|"
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   11.25
      FontName        =   "����"
      IDKind          =   -1
      DefaultCardType =   "0"
      BackColor       =   -2147483633
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   195
      TabIndex        =   26
      Top             =   810
      Width           =   840
   End
   Begin VB.Label lblClinic 
      AutoSize        =   -1  'True
      Caption         =   "�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4080
      TabIndex        =   23
      Top             =   810
      Width           =   630
   End
   Begin VB.Label lblAge 
      AutoSize        =   -1  'True
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5895
      TabIndex        =   22
      Top             =   240
      Width           =   420
   End
   Begin VB.Label lblGender 
      AutoSize        =   -1  'True
      Caption         =   "�Ա�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4290
      TabIndex        =   21
      Top             =   240
      Width           =   420
   End
   Begin VB.Label lblPatient 
      AutoSize        =   -1  'True
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   195
      TabIndex        =   20
      Top             =   240
      Width           =   420
   End
End
Attribute VB_Name = "frmAppRequestEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mrsInfo As ADODB.Recordset, mintInsure As Integer
Private mstrYBPati As String, mstrPassWord As String
Private mrsPlan As ADODB.Recordset, mintIDKind As Integer
Private mrsExtra As ADODB.Recordset
Private mlng����ID As Long

Private Sub cboArrangeNo_Click()
    If mrsPlan Is Nothing Then Exit Sub
    If cboArrangeNo.ItemData(cboArrangeNo.ListIndex) = 0 Then
        mrsPlan.Filter = "����=" & cboArrangeNo.ListIndex + 1
        If Not mrsPlan.EOF Then
            txtDept.Text = Nvl(mrsPlan!����)
        End If
    Else
        mrsExtra.Filter = "ID=" & cboArrangeNo.ItemData(cboArrangeNo.ListIndex)
        If Not mrsExtra.EOF Then
            txtDept.Text = Nvl(mrsExtra!����)
        End If
    End If
End Sub

Private Sub cboArrangeNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call gobjCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboNote_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then gobjCommFun.PressKey (vbKeyTab)
End Sub

Private Sub cmdALL_Click()
    Dim strSql As String
    Dim vRect As RECT
    Dim i As Integer
    Dim rsExtra As ADODB.Recordset
    strSql = "Select a.Id, a.����, a.����, a.��Ŀ, a.ҽ������, a.ҽ��id, a.����id" & vbNewLine & _
            "From (Select a.Id, a.����, b.���� As ����, c.���� As ��Ŀ, a.ҽ������, a.ҽ��id, a.����id" & vbNewLine & _
            "       From �ٴ������Դ A, ���ű� B, �շ���ĿĿ¼ C" & vbNewLine & _
            "       Where a.����id = b.Id And Nvl(B.����ʱ��,To_Date('3000-01-01','YYYY-MM-DD')) > Sysdate And a.��Ŀid = c.Id And (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
            "             (b.վ�� = '1' Or b.վ�� Is Null) And Exists (Select 1 From �ٴ����ﰲ�� M,�ٴ������ N Where M.��ԴID = A.ID And M.����ID = N.ID And N.����ʱ�� Is Not Null)" & vbNewLine & _
            "       Order By Decode(a.ҽ��id, [1], 1, 0) Desc, a.ҽ��id, a.����id, a.����) A"
    vRect = GetControlRect(cboArrangeNo.hWnd)
    Set rsExtra = gobjDatabase.ShowSQLSelect(Me, strSql, 0, "������Դѡ��", False, "", "������Դѡ��", _
                                                False, False, True, vRect.Left, vRect.Top - 300, 600, False, True, False, UserInfo.ID)
    If rsExtra Is Nothing Then Exit Sub
    Set mrsExtra = rsExtra
    mrsPlan.Filter = "ID=" & Val(Nvl(mrsExtra!ID))
    If Not mrsPlan.EOF Then
        cboArrangeNo.ListIndex = Val(mrsPlan!����) - 1
        mrsPlan.Filter = ""
        Exit Sub
    End If
    mrsPlan.Filter = ""
    For i = 0 To cboArrangeNo.ListCount - 1
        If cboArrangeNo.ItemData(i) = Val(Nvl(mrsExtra!ID)) Then
            cboArrangeNo.ListIndex = i
            Exit Sub
        End If
    Next i
    cboArrangeNo.AddItem mrsExtra!���� & "-" & mrsExtra!��Ŀ & "(" & mrsExtra!ҽ������ & ")"
    cboArrangeNo.ItemData(cboArrangeNo.NewIndex) = Val(Nvl(mrsExtra!ID))
    cboArrangeNo.ListIndex = cboArrangeNo.NewIndex
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub ReadBill(frmParent As Object, lng��ϢID As Long)
    On Error GoTo errHandle
    Dim strSql As String, rsTemp As ADODB.Recordset
    cmdOK.Visible = False
    cmdCancel.Caption = "�˳�(&X)"
    picMain.Enabled = False
    txtPatient.Locked = True
'    IDKind.Locked = True
    
    strSql = "Select b.����, b.�����, b.�Ա�, b.����, b.��������, c.����, d.���� As ��Ŀ, e.���� As ����, a.֪ͨԭ�� As ԭ��, a.��ʼʱ�� As ����ʱ��, a.��ֹʱ��, a.����, a.���﷽ʽ, a.�Ǽ���," & vbNewLine & _
            "       a.�Ǽ�ʱ��, a.ҽ������" & vbNewLine & _
            "From ���˷�����Ϣ��¼ A, ������Ϣ B, �ٴ������Դ C, �շ���ĿĿ¼ D, ���ű� E" & vbNewLine & _
            "Where a.Id = [1] And a.����id = b.����id And a.��Դid = c.Id And c.����id = e.Id And c.��Ŀid = d.Id"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lng��ϢID)
    If rsTemp.EOF Then
        MsgBox "���ܶ�ȡ��Ϣ,�鿴ʧ��!", vbInformation, gstrSysName
        Exit Sub
    End If
    txtPatient.Text = Nvl(rsTemp!����)
    txtClinic.Text = Nvl(rsTemp!�����)
    txtGender.Text = Nvl(rsTemp!�Ա�)
    txtAge.Text = Nvl(rsTemp!����)
    txtBirth.Text = Format(Nvl(rsTemp!��������), "yyyy-MM-dd hh:mm:ss")
    txtDept.Text = Nvl(rsTemp!����)
    cboArrangeNo.AddItem rsTemp!���� & "-" & rsTemp!��Ŀ & "(" & rsTemp!ҽ������ & ")"
    cboArrangeNo.ListIndex = cboArrangeNo.NewIndex
    Select Case Val(Nvl(rsTemp!���﷽ʽ))
    Case 1
        optStyle(0).Value = 1
        txtStyle(0).Text = Nvl(rsTemp!����)
        txtStyle(4).Text = CInt(DateDiff("d", Format(Nvl(rsTemp!�Ǽ�ʱ��), "yyyy-MM-dd hh:mm:ss"), Format(Nvl(rsTemp!��ֹʱ��), "yyyy-MM-dd hh:mm:ss")) / Val(Nvl(rsTemp!����, "1")))
    Case 2
        optStyle(1).Value = 1
        txtStyle(1).Text = Nvl(rsTemp!����)
    Case 3
        optStyle(2).Value = 1
        txtStyle(2).Text = Nvl(rsTemp!����)
    Case 4
        optStyle(3).Value = 1
        txtStyle(3).Text = Nvl(rsTemp!����)
    End Select
    cboNote.Text = Nvl(rsTemp!ԭ��)
    dtpDate.Value = Format(Nvl(rsTemp!����ʱ��), "yyyy-MM-dd hh:mm:ss")
    txtStyle(5).Text = DateDiff("d", Format(Nvl(rsTemp!����ʱ��), "yyyy-MM-dd hh:mm:ss"), Format(Nvl(rsTemp!��ֹʱ��), "yyyy-MM-dd hh:mm:ss"))
    txt�Ǽ���.Text = Nvl(rsTemp!�Ǽ���)
    txt�Ǽ�ʱ��.Text = Nvl(rsTemp!�Ǽ�ʱ��)
    
    Me.Show vbModal, frmParent
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandle
    Dim strSql As String, byt���﷽ʽ As Byte
    Dim i As Integer, lngNum As Long
    Dim blnFind As Boolean
    
    If mrsInfo Is Nothing Then
        MsgBox "����ȷ������,����ѡ����!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    blnFind = False
    For i = 0 To 3
        If optStyle(i).Value = True Then
            byt���﷽ʽ = i + 1: lngNum = Val(txtStyle(i).Text)
            blnFind = True
        End If
    Next i
    If blnFind = False Then
        MsgBox "��ѡ��һ�ָ��﷽ʽ!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If cboArrangeNo.ListIndex = -1 Then
        MsgBox "��ѡ��һ����Դ!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If gobjCommFun.ActualLen(Me.cboNote.Text) > 100 Then
        MsgBox "����ĸ���˵������!(�����������50�����ֻ�100���ַ�)", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If Format(dtpDate.Value, "yyyy-mm-dd hh:mm:ss") < Format(gdatRegistTime, "yyyy-mm-dd hh:mm:ss") Then
        MsgBox "ԤԼ�Ǽǵ�ʱ��" & Format(dtpDate.Value, "yyyy-mm-dd hh:mm:ss") & "���ǳ�����Ű�ģʽ��Ч��ʱ��,���ܵǼ�!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If cboArrangeNo.ItemData(cboArrangeNo.ListIndex) = 0 Then
        mrsPlan.Filter = "����=" & cboArrangeNo.ListIndex + 1
        If mrsPlan.EOF Then
            MsgBox "����ȷ����ǰѡ��ĺ�Դ,�޷����еǼ�!", vbInformation, gstrSysName
            Exit Sub
        End If
        
        strSql = "zl_����ԤԼ�Ǽ�_Insert("
        strSql = strSql & mrsInfo!����ID & ","
        strSql = strSql & mrsPlan!ID & ","
        strSql = strSql & byt���﷽ʽ & ","
        strSql = strSql & lngNum & ",'"
        strSql = strSql & cboNote.Text & "',"
        strSql = strSql & "To_Date('" & dtpDate.Value & "','yyyy-mm-dd hh24:mi:ss'),"
        strSql = strSql & Val(txtStyle(5).Text) & ",'"
        strSql = strSql & UserInfo.���� & "','"
        strSql = strSql & UserInfo.��� & "')"
    Else
        strSql = "zl_����ԤԼ�Ǽ�_Insert("
        strSql = strSql & mrsInfo!����ID & ","
        strSql = strSql & cboArrangeNo.ItemData(cboArrangeNo.ListIndex) & ","
        strSql = strSql & byt���﷽ʽ & ","
        strSql = strSql & lngNum & ",'"
        strSql = strSql & cboNote.Text & "',"
        strSql = strSql & "To_Date('" & dtpDate.Value & "','yyyy-mm-dd hh24:mi:ss'),"
        strSql = strSql & Val(txtStyle(5).Text) & ",'"
        strSql = strSql & UserInfo.���� & "','"
        strSql = strSql & UserInfo.��� & "')"
    End If
    
    Call gobjDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Unload Me
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then gobjCommFun.PressKey (vbKeyTab)
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    If txtPatient.Locked Then Exit Sub
    
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
        'ϵͳIC��
        If Not mobjICCard Is Nothing Then
           txtPatient.Text = mobjICCard.Read_Card()
           If txtPatient.Text <> "" Then
                Call GetPatient(objCard, txtPatient.Text, True)
           End If
        End If
        Exit Sub
    End If
    
    lng�����ID = objCard.�ӿ����
    
    If lng�����ID <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '����:�����ӿ�
    '    '���:frmMain-���õĸ�����
    '    '       lngModule-���õ�ģ���
    '    '       strExpand-��չ����,������
    '    '       blnOlnyCardNO-������ȡ����
    '    '����:strOutCardNO-���صĿ���
    '    '       strOutPatiInforXML-(������Ϣ����.XML��)
    '    '����:��������    True:���óɹ�,False:����ʧ��\
    If gobjSquare.objSquareCard.zlReadCard(Me, glngModul, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    
    If txtPatient.Text <> "" Then
        Call GetPatient(objCard, txtPatient.Text, True)
    End If
    
End Sub

Public Sub ShowMe(frmMain As Object, Optional lng����ID As Long)
    mlng����ID = lng����ID
    Call LoadRegPlans
    dtpDate.Value = gobjDatabase.Currentdate
    Me.Show vbModal, frmMain
End Sub

Private Sub Form_Activate()
    If txtPatient.Enabled And txtPatient.Visible And cmdOK.Visible = True Then
        txtPatient.SetFocus
    Else
        If cmdOK.Visible = False And cmdCancel.Visible And cmdCancel.Enabled Then cmdCancel.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    Call mobjIDCard.SetParent(Me.hWnd)
    Call mobjICCard.SetParent(Me.hWnd)
    Call InitIDKind
    IDKind.RaisEffect picMain, -1
    If mlng����ID <> 0 Then
        Call GetPatient(IDKind.GetCurCard, "-" & mlng����ID, False)
    End If
End Sub

'��ʼ��IDKIND
Private Function InitIDKind() As Boolean
    Dim objCard As Card, strTemp As String
    Dim lngCardID As Long
    If gobjSquare Is Nothing Then CreateSquareCardObject Me, glngModul
    Call IDKind.zlInit(Me, glngSys, 1260, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "��|����|0;ҽ|ҽ����|0;��|���֤��|0;��|�����|0", txtPatient)
    Set objCard = IDKind.GetfaultCard
    If IDKind.Cards.��ȱʡ������ And Not objCard Is Nothing Then
        gobjSquare.blnȱʡ�������� = objCard.�������Ĺ��� <> ""
        gobjSquare.intȱʡ���ų��� = objCard.���ų���
        Set gobjSquare.objDefaultCard = objCard
    Else
        gobjSquare.blnȱʡ�������� = IDKind.Cards.������ʾ
        gobjSquare.intȱʡ���ų��� = 100
    End If
    Call GetRegInFor(g˽��ģ��, Me.Name, "idkind", strTemp)
    mintIDKind = Val(strTemp)
    If mintIDKind > 0 And mintIDKind <= IDKind.ListCount Then IDKind.IDKind = mintIDKind
End Function

Private Sub Form_Unload(Cancel As Integer)
    Dim strValues As String, strArray() As String
    Set mrsInfo = Nothing
    Set mobjICCard = Nothing
    Set mobjIDCard = Nothing
    If cmdOK.Visible Then
        Call GetRegInFor(g˽��ģ��, Me.Name, "strValues", strValues)
        If strValues = "" Then strValues = "1,14|1|2|14"
        If optStyle(0).Value = True Then
            strArray = Split(strValues & "|||", "|")
            strValues = txtStyle(0).Text & "," & txtStyle(4).Text & "|" & strArray(1) & "|" & strArray(2) & "|" & strArray(3)
        End If
        If optStyle(1).Value = True Then
            strArray = Split(strValues & "|||", "|")
            strValues = strArray(0) & "|" & txtStyle(1).Text & "|" & strArray(2) & "|" & strArray(3)
        End If
        If optStyle(2).Value = True Then
            strArray = Split(strValues & "|||", "|")
            strValues = strArray(0) & "|" & strArray(1) & "|" & txtStyle(2).Text & "|" & strArray(3)
        End If
        If optStyle(3).Value = True Then
            strArray = Split(strValues & "|||", "|")
            strValues = strArray(0) & "|" & strArray(1) & "|" & strArray(2) & "|" & txtStyle(3).Text
        End If
        
        Call SaveRegInFor(g˽��ģ��, Me.Name, "strValues", strValues)
    End If
    mlng����ID = 0
End Sub


Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    txtPatient.Text = objPatiInfor.����
    Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), True)
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNO As String)
    IDKind.IDKind = IDKind.GetKindIndex("IC����")
    txtPatient.Text = strCardNO
    Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), True)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    IDKind.IDKind = IDKind.GetKindIndex("���֤��")
    txtPatient.Text = strID
    Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), True)
End Sub

Private Sub optStyle_Click(Index As Integer)
    Dim strValues As String, strArray() As String
    Call GetRegInFor(g˽��ģ��, Me.Name, "strValues", strValues)
    If strValues = "" Then strValues = "1,14|1|2|14"
    strArray = Split(strValues & "|||", "|")
    If optStyle(0).Value = True Then
        txtStyle(0).Enabled = True
        udStyle(0).Enabled = True
        txtStyle(1).Enabled = False
        udStyle(1).Enabled = False
        txtStyle(2).Enabled = False
        udStyle(2).Enabled = False
        txtStyle(3).Enabled = False
        udStyle(3).Enabled = False
        txtStyle(4).Enabled = True
        udStyle(4).Enabled = True
        If txtStyle(0).Text = "" Then
            txtStyle(0).Text = Split(strArray(0) & ",", ",")(0)
            txtStyle(4).Text = Split(strArray(0) & ",", ",")(1)
        End If
    End If
    If optStyle(1).Value = True Then
        txtStyle(0).Enabled = False
        udStyle(0).Enabled = False
        txtStyle(1).Enabled = True
        udStyle(1).Enabled = True
        txtStyle(2).Enabled = False
        udStyle(2).Enabled = False
        txtStyle(3).Enabled = False
        udStyle(3).Enabled = False
        txtStyle(4).Enabled = False
        udStyle(4).Enabled = False
        If txtStyle(1).Text = "" Then
            txtStyle(1).Text = strArray(1)
        End If
    End If
    If optStyle(2).Value = True Then
        txtStyle(0).Enabled = False
        udStyle(0).Enabled = False
        txtStyle(1).Enabled = False
        udStyle(1).Enabled = False
        txtStyle(2).Enabled = True
        udStyle(2).Enabled = True
        txtStyle(3).Enabled = False
        udStyle(3).Enabled = False
        txtStyle(4).Enabled = False
        udStyle(4).Enabled = False
        If txtStyle(2).Text = "" Then
            txtStyle(2).Text = strArray(2)
        End If
    End If
    If optStyle(3).Value = True Then
        txtStyle(0).Enabled = False
        udStyle(0).Enabled = False
        txtStyle(1).Enabled = False
        udStyle(1).Enabled = False
        txtStyle(2).Enabled = False
        udStyle(2).Enabled = False
        txtStyle(3).Enabled = True
        udStyle(3).Enabled = True
        txtStyle(4).Enabled = False
        udStyle(4).Enabled = False
        If txtStyle(3).Text = "" Then
            txtStyle(3).Text = strArray(3)
        End If
    End If
    Call CaclDate
End Sub

Private Sub CaclDate()
    Dim intDays As Integer
    Dim strTemp As String
    If cmdOK.Visible = False Then Exit Sub
    strTemp = cboNote.Text
    cboNote.Clear
    If optStyle(0).Value = True Then
        intDays = Val(txtStyle(0).Text) * Val(txtStyle(4).Text)
        dtpDate.Value = DateAdd("d", intDays, gobjDatabase.Currentdate)
        cboNote.AddItem Val(txtStyle(0).Text) & "���Ƴ̺���"
    End If
    If optStyle(1).Value = True Then
        intDays = Val(txtStyle(1).Text)
        dtpDate.Value = DateAdd("m", intDays, gobjDatabase.Currentdate)
        cboNote.AddItem Val(txtStyle(1).Text) & "���º���"
    End If
    If optStyle(2).Value = True Then
        intDays = Val(txtStyle(2).Text) * 7
        dtpDate.Value = DateAdd("d", intDays, gobjDatabase.Currentdate)
        cboNote.AddItem Val(txtStyle(2).Text) & "�ܺ���"
    End If
    If optStyle(3).Value = True Then
        intDays = Val(txtStyle(3).Text)
        dtpDate.Value = DateAdd("d", intDays, gobjDatabase.Currentdate)
        cboNote.AddItem Val(txtStyle(3).Text) & "�����"
    End If
    cboNote.Text = strTemp
End Sub

Private Sub optStyle_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    gobjCommFun.PressKey (vbKeyTab)
End Sub

Private Sub txtPatient_Change()
    If Me.ActiveControl Is txtPatient Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard txtPatient.Text = ""
    End If
End Sub

Private Sub txtPatient_GotFocus()
    Call gobjControl.TxtSelAll(txtPatient)
    Call gobjCommFun.OpenIme(True)
    If txtPatient.Text = "" And ActiveControl Is txtPatient Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard txtPatient.Text = ""
    End If
End Sub

Private Sub LoadRegPlans()
    Dim strSql As String, rsTemp As ADODB.Recordset
    On Error GoTo errH
    strSql = "Select Rownum As ����, a.Id, a.����, a.����, a.��Ŀ, a.ҽ������, a.ҽ��id, a.����id" & vbNewLine & _
            "From (Select a.Id, a.����, b.���� As ����, c.���� As ��Ŀ, a.ҽ������, a.ҽ��id, a.����id" & vbNewLine & _
            "       From �ٴ������Դ A, ���ű� B, �շ���ĿĿ¼ C" & vbNewLine & _
            "       Where (a.ҽ��id = [1] Or (a.ҽ��id Is Null And a.����id = [2])) And a.����id = b.Id And a.��Ŀid = c.Id And (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
            "             (b.վ�� = '1' Or b.վ�� Is Null) And Exists (Select 1 From �ٴ����ﰲ�� M,�ٴ������ N Where M.��ԴID = A.ID And M.����ID = N.ID And N.����ʱ�� Is Not Null)" & vbNewLine & _
            "       Order By Decode(a.ҽ��id, [1], 1, 0) Desc, a.ҽ��id, a.����id, a.����) A"
    Set mrsPlan = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID, UserInfo.����ID)
    cboArrangeNo.Clear
    Do While Not mrsPlan.EOF
        cboArrangeNo.AddItem mrsPlan!���� & "-" & mrsPlan!��Ŀ & "(" & mrsPlan!ҽ������ & ")"
        cboArrangeNo.ItemData(cboArrangeNo.NewIndex) = 0
        mrsPlan.MoveNext
    Loop
    If cboArrangeNo.ListCount <> 0 Then cboArrangeNo.ListIndex = 0
    Exit Sub
errH:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub zlInusreIdentify()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�ҽ������鿨
    '���ƣ����˺�
    '���ڣ�2010-07-14 11:32:08
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long
    Dim str�������� As String
    Dim rsTmp As ADODB.Recordset
    Dim cur��� As Currency
    Dim curMoney As Currency
    Dim blnDeposit As Boolean, blnInsure As Boolean
    If mrsInfo Is Nothing Then
        lng����ID = 0
        str�������� = ""
    Else
        lng����ID = Val(Nvl(mrsInfo!����ID))
        str�������� = Nvl(mrsInfo!��������)
    End If

    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    IDKind.SetAutoReadCard False

    Dim strAdvance As String    '����ģʽ(0-�Ƚ�������ƻ�1-�����ƺ����)|�Һŷ���ȡ��ʽ(0-���ջ�1-����)
    Dim varData As Variant
    mstrYBPati = gclsInsure.Identify(3, lng����ID, mintInsure, strAdvance)
    If txtPatient.Text = "" And Not txtPatient.Locked Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(True)
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(True)
        IDKind.SetAutoReadCard True
    End If
    
    If mstrYBPati = "" Then
        If Not txtPatient.Enabled Then txtPatient.Enabled = True
         mstrYBPati = "": mintInsure = 0: txtPatient.SetFocus
         Exit Sub
    End If
    
    '�ջ�0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����ID
    If UBound(Split(mstrYBPati, ";")) >= 8 Then
        If IsNumeric(Split(mstrYBPati, ";")(8)) Then lng����ID = Val(Split(mstrYBPati, ";")(8))
    End If
        
    If lng����ID = 0 Then
        mstrYBPati = "": mintInsure = 0: txtPatient.SetFocus
        Exit Sub
    End If
    
    txtPatient.Text = "-" & lng����ID
    Call txtPatient_Validate(False)    '���е�Setfocus����ʹ���¼�(txtPatient_KeyPress)ִ�����,�����ٴ��Զ�ִ��txtPatient_Validate
    Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), False)
    Call SetPatiColor(txtPatient, str��������, vbRed)
    txtPatient.BackColor = &HE0E0E0
    txtPatient.Locked = True
    
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    '0-�����,1-����,2-�Һŵ�,3-���￨��,4-ҽ����
    Dim blnCard As Boolean
    Dim strKind As String, intLen As Integer
    Static sngBegin As Single
    Dim sngNow As Single
    
    'ҽ����֤
    If txtPatient.Text = "" And KeyAscii = 13 Then
        KeyAscii = 0
        Call zlInusreIdentify
    End If
    
    strKind = IDKind.GetCurCard.����
    txtPatient.PasswordChar = IIf(IDKind.GetCurCard.�������Ĺ��� <> "", "*", "")
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    
    
    'ȡȱʡ��ˢ����ʽ
            '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
            '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
            '��7λ��,��ֻ��������,��Ȼȡ������
    Select Case strKind
    Case "����"
        blnCard = gobjCommFun.InputIsCard(txtPatient, KeyAscii, gobjSquare.blnȱʡ��������)
        intLen = gobjSquare.intȱʡ���ų���
    Case "�����"
        If InStr("0123456789-" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    Case "�Һŵ�"
    Case "ҽ����"
    Case Else
            If IDKind.GetCurCard.�ӿ���� <> 0 Then
                blnCard = gobjCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.GetCurCard.�������Ĺ��� <> "")
                intLen = IDKind.GetCurCard.���ų���
            End If
    End Select
    
    'ˢ����ϻ���������س�
    If (blnCard And Len(txtPatient.Text) = intLen - 1 And KeyAscii <> 8) Or (KeyAscii = 13) Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), blnCard)
        gobjControl.TxtSelAll txtPatient
   End If
End Sub

Private Sub txtPatient_LostFocus()
    Call gobjCommFun.OpenIme
    IDKind.SetAutoReadCard False
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
End Sub

Private Sub txtPatient_Validate(Cancel As Boolean)
    txtPatient.Text = Trim(txtPatient.Text)
End Sub

Private Sub GetPatient(objCard As zlIDKind.Card, ByVal strInput As String, ByVal blnCard As Boolean, Optional blnInputIDCard As Boolean = False, Optional ByRef Cancel As Boolean)
    '���ܣ���ȡ������Ϣ
    '������blnCard=�Ƿ���￨ˢ��
    '
    '         blnInputIDCard-�Ƿ����֤ˢ��
    '����:Cancel-Ϊtrue��ʾ���صķ�����ȡ������Ϣ
    Dim strSql As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    Dim rsTmp As ADODB.Recordset, strTemp As String, rsFeeType As ADODB.Recordset
    Dim blnSame As Boolean, blnCancel As Boolean
    Dim cur��� As Currency, curMoney As Currency
    Dim strInputInfo As String '���洫��������ı� ������ʹ�����֤�� �Բ��˽��в��Һ� ���滻��"-" ����ID�����
    Dim i As Integer, strPati As String
    Dim vRect As RECT, str����Ժ As String
    Dim blnҽ���� As Boolean
    Dim IntMsg As VbMsgBoxResult
    Dim blnOtherType As Boolean '�Ƿ������

    strInputInfo = strInput
    
    On Error GoTo errH
    blnҽ���� = False
    
    If objCard Is Nothing Then Set objCard = IDKind.GetCurCard

    strSql = "Select  A.����ID,A.�����,A.סԺ��,A.���￨��,A.�ѱ�,A.ҽ�Ƹ��ʽ,A.����,A.�Ա�,A.����,A.��������,A.�����ص�,A.���֤��,A.����֤��,A.���,A.ְҵ,A.����,A.��������, " & _
             "A.����,A.����,A.����,A.ѧ��,A.����״��,A.��ͥ��ַ,A.��ͥ�绰,A.��ͥ��ַ�ʱ�,A.�໤��,A.��ϵ������,A.��ϵ�˹�ϵ,A.��ϵ�˵�ַ,A.��ϵ�˵绰,A.���ڵ�ַ, " & _
             "A.���ڵ�ַ�ʱ�,A.Email,A.QQ,A.��ͬ��λid,A.������λ,A.��λ�绰,A.��λ�ʱ�,A.��λ������,A.��λ�ʺ�,A.������,A.������,A.��������,A.����ʱ��,A.����״̬, " & _
             "A.��������,A.סԺ����,A.��ǰ����id,A.��ǰ����id,A.��ǰ����,A.��Ժʱ��,A.��Ժʱ��,A.��Ժ,A.IC����,A.������,A.ҽ����,A.����,A.��ѯ����,A.�Ǽ�ʱ��,A.ͣ��ʱ��,A.����,A.��ϵ�����֤��, " & _
             "B.���� ��������,A.��ѯ���� As ����֤��,A.����ģʽ From ������Ϣ A,������� B  Where A.���� = B.���(+) And A.ͣ��ʱ�� is NULL  "

   
    If blnCard And objCard.���� Like "����*" And mstrYBPati = "" And InStr("-+*.", Left(strInput, 1)) = 0 Then     'ˢ��
        If IDKind.Cards.��ȱʡ������ And Not IDKind.GetfaultCard Is Nothing Then
            lng�����ID = IDKind.GetfaultCard.�ӿ����
        ElseIf IDKind.GetCurCard.�ӿ���� > 0 Then
            lng�����ID = IDKind.GetCurCard.�ӿ����
'        Else
'            lng�����ID = gCurSendCard.lng�����ID
        End If
        
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0

        If lng����ID <= 0 Then GoTo NewPati:
        strInput = "-" & lng����ID
        blnHavePassWord = True
        strSql = strSql & " And A.����ID=[2] " & str����Ժ
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then
        '�����
        strSql = strSql & " And A.�����=[2]" & str����Ժ
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then
        '����ID
        strSql = strSql & " And A.����ID=[2]" & _
        IIf(mstrYBPati <> "", "", str����Ժ)
    ElseIf blnInputIDCard Then  '���������֤ʶ��
        strInput = UCase(strInput)
        If gobjSquare.objSquareCard.zlGetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
        strInput = "-" & lng����ID
        strSql = strSql & " And A.����ID=[2] " & str����Ժ
    Else
        Select Case objCard.����
            Case "����", "��������￨"
                strPati = _
                    " Select distinct 1 as ����ID,A.����ID as ID,A.����ID,A.����,A.�Ա�,A.����,A.�����,A.��������,A.���֤��,A.��ͥ��ַ,A.������λ" & _
                    " From ������Ϣ A " & _
                    " Where Rownum <101 And A.ͣ��ʱ�� is NULL And A.���� Like [1]" & str����Ժ
                    
                strPati = strPati & " Union ALL " & _
                        "Select 0,0 as ID,-NULL,'[�²���]',NULL,NULL,-NULL,To_Date(NULL),NULL,NULL,NULL From Dual"
                strPati = strPati & " Order by ����ID,����"
                    
                vRect = GetControlRect(txtPatient.hWnd)
                Set rsTmp = gobjDatabase.ShowSQLSelect(Me, strPati, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%")
                If Not rsTmp Is Nothing Then
                    If rsTmp!ID = 0 Then '�����²���
                        txtPatient.Text = ""
                        MsgBox "û���ҵ���Ӧ�Ĳ�����Ϣ������������Ϣ�Ƿ���ȷ���߲����Ƿ񽨵���", vbInformation, gstrSysName
                        Set mrsInfo = Nothing: Exit Sub
                    Else '�Բ���ID��ȡ
                        strInput = rsTmp!����ID
                        strSql = strSql & " And A.����ID=[1]"
                    End If
                Else 'ȡ��ѡ��
                    txtPatient.Text = ""
                    Set mrsInfo = Nothing: Exit Sub
                End If
            Case "ҽ����"
                strInput = UCase(strInput)
                blnҽ���� = True
                strSql = strSql & " And A.ҽ����=[1]" & str����Ժ
                
            Case "���֤��", "���֤", "�������֤"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strInput = "-" & lng����ID
                strSql = strSql & " And A.����ID=[2] " & str����Ժ
                 
            Case "IC����", "IC��"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("IC��", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strInput = "-" & lng����ID
                strSql = strSql & " And A.����ID=[2] " & str����Ժ
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSql = strSql & " And A.�����=[1]" & str����Ժ
             Case Else
                '��������,��ȡ��صĲ���ID
                If objCard.�ӿ���� > 0 Then
                    lng�����ID = objCard.�ӿ����
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                    blnOtherType = True
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, _
                        strPassWord, strErrMsg) = False Then lng����ID = 0
                End If
                strSql = strSql & " And A.����ID=[2]" & str����Ժ
                strInput = "-" & lng����ID
                blnHavePassWord = True
        End Select
    End If
ReadPati:
    If strPassWord <> "" Then
        If Not gobjCommFun.VerifyPassWord(Me, "" & strPassWord) Then
            MsgBox "���������֤ʧ�ܣ�", vbInformation, gstrSysName
            ClearPatient
            Exit Sub
        End If
    End If
    Set mrsInfo = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, strInput, Mid(strInput, 2), strTemp)
    strInput = strInputInfo
    If Not mrsInfo.EOF Then
        txtPatient.Text = Nvl(mrsInfo!����) '�����Change�¼�
        txtPatient.BackColor = &H80000005
        '�ڵ���txtPatient_Change�¼���������źͲ���������Ϊ�յ������ �޷�ʶ��ò�����Ϣ ���ִ���
        '���������ݿ����ݴ����ٽ��к����Ĵ���
        If mrsInfo Is Nothing Then Cancel = True: Exit Sub
        Call SetPatiColor(txtPatient, Nvl(mrsInfo!��������), IIf(Trim(mintInsure) = "", txtPatient.ForeColor, vbRed))
        
        mstrPassWord = strPassWord
        If Not blnHavePassWord Then mstrPassWord = Nvl(mrsInfo!����֤��)
        txtGender.Text = Nvl(mrsInfo!�Ա�)
        txtBirth.Text = Format(Nvl(mrsInfo!��������), "yyyy-MM-dd hh:mm")
        txtPatient.PasswordChar = ""
        
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtPatient.IMEMode = 0
        txtAge.Text = Nvl(mrsInfo!����)
        txtClinic.Text = Nvl(mrsInfo!�����)
        txt�Ǽ���.Text = UserInfo.����
        txt�Ǽ�ʱ��.Text = Format(gobjDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss")
        
        If cboArrangeNo.Enabled And cboArrangeNo.Visible Then cboArrangeNo.SetFocus
    Else
NewPati:
        MsgBox "û���ҵ���Ӧ�Ĳ�����Ϣ������������Ϣ�Ƿ���ȷ���߲����Ƿ񽨵���", vbInformation, gstrSysName
        ClearPatient
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub ClearPatient()
    txtPatient.Text = ""
    txtPatient.BackColor = &H80000005
    txtPatient.ForeColor = vbBlack
    txtPatient.Locked = False
    txtGender.Text = ""
    txtAge.Text = ""
    txtBirth.Text = ""
    optStyle(0).Value = 1
    mintInsure = 0
    Set mrsInfo = Nothing
End Sub

Private Sub txtStyle_Change(Index As Integer)
    Call CaclDate
End Sub

Private Sub txtStyle_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Index <> 5 Then
            If cboNote.Visible And cboNote.Enabled Then cboNote.SetFocus
        Else
            Call gobjCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub
