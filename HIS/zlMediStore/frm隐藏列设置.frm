VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm���������� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����������"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4620
   Icon            =   "frm����������.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdAllNotUncheck 
      Caption         =   "ȫ��(&U)"
      Height          =   350
      Left            =   3360
      TabIndex        =   6
      Top             =   1545
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllCheck 
      Caption         =   "ȫѡ(&C)"
      Height          =   350
      Left            =   3360
      TabIndex        =   5
      Top             =   1080
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3360
      TabIndex        =   2
      Top             =   3220
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3360
      TabIndex        =   1
      Top             =   3690
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwColumns 
      Height          =   3315
      Left            =   30
      TabIndex        =   0
      Top             =   960
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   5847
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2520
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����������.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "* ����ѡ����������ģ�鶼������ʾ"
      Height          =   180
      Index           =   1
      Left            =   720
      TabIndex        =   7
      Top             =   360
      Width           =   2880
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "�����ǿ������ص����嵥"
      Height          =   195
      Index           =   3
      Left            =   60
      TabIndex        =   4
      Top             =   720
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frm����������.frx":1D16
      Top             =   150
      Width           =   480
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "����Զ��̵�������н�����ʾ���ؿ���"
      Height          =   180
      Index           =   0
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   3240
   End
End
Attribute VB_Name = "frm����������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrColsName As String
Private mstrReturnColsName As String

Private Sub cmdAllCheck_Click()
    Dim i As Integer
    
    For i = 1 To lvwColumns.ListItems.count
        If Not lvwColumns.ListItems.Item(i).Checked Then lvwColumns.ListItems.Item(i).Checked = True
    Next
End Sub

Private Sub cmdAllNotUncheck_Click()
    Dim i As Integer
    
    For i = 1 To lvwColumns.ListItems.count
        If lvwColumns.ListItems.Item(i).Checked Then lvwColumns.ListItems.Item(i).Checked = False
    Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    
    mstrReturnColsName = ":"
    
    For i = 1 To lvwColumns.ListItems.count
        mstrReturnColsName = mstrReturnColsName & lvwColumns.ListItems.Item(i) & ","
        If lvwColumns.ListItems.Item(i).Checked Then
            mstrReturnColsName = mstrReturnColsName & "1:"
        Else
            mstrReturnColsName = mstrReturnColsName & "0:"
        End If
    Next
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strColName() As String
    Dim i As Integer
    
    strColName = Split(mstrColsName, ":")
    
    For i = LBound(strColName) + 1 To UBound(strColName) - 1
        If Split(strColName(i), ",")(1) = 0 Then
            lvwColumns.ListItems.Add , "K" & Split(strColName(i), ",")(0), Split(strColName(i), ",")(0), , 1
        Else
            lvwColumns.ListItems.Add(, "K" & Split(strColName(i), ",")(0), Split(strColName(i), ",")(0), , 1).Checked = True
        End If
    Next
End Sub


Public Function ShowME(ByVal frmParent As Object, ByVal strColsName As String) As String
    mstrColsName = strColsName
    mstrReturnColsName = ""
    
    Me.Show 1, frmParent
    
    ShowME = mstrReturnColsName
End Function

