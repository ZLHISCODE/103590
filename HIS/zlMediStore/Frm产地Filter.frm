VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "Msdatgrd.ocx"
Begin VB.Form Frm����Filter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ѡ����"
   ClientHeight    =   4905
   ClientLeft      =   1170
   ClientTop       =   3135
   ClientWidth     =   7200
   ControlBox      =   0   'False
   Icon            =   "Frm����Filter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6000
      TabIndex        =   2
      Top             =   720
      Width           =   1100
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6000
      TabIndex        =   1
      Top             =   120
      Width           =   1100
   End
   Begin MSDataGridLib.DataGrid Dbg 
      Height          =   4785
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   8440
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Frm����Filter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public BlnSelect As Boolean 'ѡ���־
Public str���� As String 'ҩƷ����
Public WithEvents PhiscRecBound As ADODB.Recordset  '�󶨼�¼��
Attribute PhiscRecBound.VB_VarHelpID = -1
Private BlnNotҩƷ As Boolean
Private mblnHaveButton As Boolean
Public CurrentID As Long '���浱ǰID


Private Sub cmdCancel_Click()
    BlnSelect = False
    Me.Hide
    Exit Sub
End Sub

Private Sub cmdOK_Click()
    Dbg_DblClick
End Sub

Private Sub Dbg_DblClick()
    BlnSelect = True
    Me.Hide
End Sub

Private Sub Dbg_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Dbg_DblClick
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        BlnSelect = False
        Me.Hide
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Set Dbg.DataSource = PhiscRecBound
    SetVisible
End Sub

Public Property Get ҩƷ() As Boolean
    ҩƷ = BlnNotҩƷ
End Property

Public Property Let ҩƷ(ByVal vNewValue As Boolean)
    BlnNotҩƷ = vNewValue
End Property

Public Property Get HaveButton() As Boolean
    HaveButton = mblnHaveButton
End Property

Public Property Let HaveButton(ByVal vNewValue As Boolean)
    mblnHaveButton = vNewValue
End Property

Private Sub SetVisible()
    If mblnHaveButton = True Then
        CmdOK.Visible = True
        CmdCancel.Visible = True
        Dbg.Width = CmdOK.Left - Dbg.Left - 50
    Else
        CmdOK.Visible = False
        CmdCancel.Visible = False
        Dbg.Width = Me.ScaleWidth - Dbg.Left - 50
        
    End If
End Sub
