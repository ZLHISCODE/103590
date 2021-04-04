VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmMedTech 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.TaskPanel tplFunc 
      Height          =   4770
      Left            =   30
      TabIndex        =   0
      Top             =   315
      Width           =   3210
      _Version        =   589884
      _ExtentX        =   5662
      _ExtentY        =   8414
      _StockProps     =   64
      Behaviour       =   1
      ItemLayout      =   2
      HotTrackStyle   =   3
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   660
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmedtech.frx":0000
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmedtech.frx":0458
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmedtech.frx":08AA
            Key             =   "Read"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.ShortcutCaption stcItem 
      Height          =   300
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   3210
      _Version        =   589884
      _ExtentX        =   5662
      _ExtentY        =   529
      _StockProps     =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
End
Attribute VB_Name = "frmMedTech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim objGroup As TaskPanelGroup
    
    On Error GoTo errHandle
    Set objGroup = tplFunc.Groups.Add(1, "医疗工作")
    
    objGroup.Items.Add(1, "质控报告词句", xtpTaskItemTypeLink, 2).Selected = True
    objGroup.Items.Add 2, "质控检验方法", xtpTaskItemTypeLink, 2
    objGroup.Items.Add 3, "质控试剂来源", xtpTaskItemTypeLink, 2
    objGroup.Items.Add 4, "细菌检测方法", xtpTaskItemTypeLink, 2
    objGroup.Items.Add 5, "细菌耐药机制", xtpTaskItemTypeLink, 2
    objGroup.Items.Add 6, "检验结果描述", xtpTaskItemTypeLink, 2
    objGroup.Items.Add 7, "检验培养文字", xtpTaskItemTypeLink, 2
    
    tplFunc.SetMargins 1, 2, 0, 2, 2
    tplFunc.SelectItemOnFocus = True
    Call tplFunc.SetImageList(ils32)
    tplFunc.SetIconSize 24, 24
    objGroup.CaptionVisible = False
    objGroup.Expanded = True
    stcItem.Caption = "医疗工作"
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    stcItem.Left = Me.Left
    stcItem.Width = Me.Width
    
    tplFunc.Height = Me.Height - Me.stcItem.Height
    tplFunc.Width = Me.Width
    tplFunc.Left = Me.Left
    tplFunc.Top = Me.stcItem.Top + Me.stcItem.Height
End Sub

Private Sub tplFunc_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    Call frmBaseInfoList.ShowItemInfo(Trim(Item.Caption))
End Sub





