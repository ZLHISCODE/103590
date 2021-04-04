VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelect 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "社区选择"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5415
   Icon            =   "frmSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ListView lvwList 
      Height          =   1590
      Left            =   330
      TabIndex        =   0
      Top             =   945
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   2805
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "社区"
         Object.Width           =   7761
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   930
      Top             =   1245
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelect.frx":058A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2595
      TabIndex        =   1
      Top             =   2865
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3810
      TabIndex        =   2
      Top             =   2865
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   5
      Top             =   2700
      Width           =   6900
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   5415
      TabIndex        =   3
      Top             =   0
      Width           =   5415
      Begin VB.Image Image1 
         Height          =   720
         Left            =   4350
         Picture         =   "frmSelect.frx":0E64
         Top             =   45
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择当前社区病人所属的社区"
         Height          =   180
         Left            =   375
         TabIndex        =   4
         Top             =   300
         Width           =   2520
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   -45
         X2              =   5500
         Y1              =   765
         Y2              =   765
      End
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint社区 As Integer
Private mblnOK As Boolean

Public Function ShowMe() As Integer
'功能：选择启用的病人所属的社区接口
'返回：根据选择返回对应的社区序号
'说明：如果只有一个启用的接口，则不弹出界面选择，直接返回
    grsCommunity.Filter = "启用=1"
    If grsCommunity.EOF Then Exit Function
    If grsCommunity.RecordCount = 1 Then
        ShowMe = grsCommunity!序号
        Exit Function
    Else
        Me.Show 1
        If mblnOK Then
            ShowMe = mint社区
        End If
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If lvwList.SelectedItem Is Nothing Then Exit Sub
    
    mint社区 = Val(Mid(lvwList.SelectedItem.Key, 2))
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()
    mblnOK = False
    mint社区 = 0
    
    grsCommunity.Filter = "启用=1"
    Do While Not grsCommunity.EOF
        lvwList.ListItems.Add , "_" & grsCommunity!序号, grsCommunity!名称, , 1
        grsCommunity.MoveNext
    Loop
End Sub

Private Sub lvwList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdOK_Click
    End If
End Sub
