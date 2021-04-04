VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmItem 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2355
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   2355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin XtremeSuiteControls.TaskPanel tplFunc 
      Height          =   3135
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1215
      _Version        =   589884
      _ExtentX        =   2143
      _ExtentY        =   5530
      _StockProps     =   64
      Behaviour       =   1
      ItemLayout      =   2
      HotTrackStyle   =   3
   End
   Begin XtremeCommandBars.ImageManager imgMain 
      Left            =   1920
      Top             =   480
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmItem.frx":0000
   End
End
Attribute VB_Name = "frmItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim tpGroup As TaskPanelGroup
    
    With tplFunc
        Set tpGroup = .Groups.Add(0, "")
        tpGroup.CaptionVisible = False
        tpGroup.Expanded = True
        tpGroup.Items.Add(1, "变动过程升级管理", xtpTaskItemTypeLink, 2).Selected = False
        tpGroup.Items.Add(2, "变动过程日常管理", xtpTaskItemTypeLink, 3).Selected = False

         
        .SetMargins 1, 2, 0, 2, 2
        .SetIconSize 24, 24
        .SelectItemOnFocus = True
        .Icons.AddIcons imgMain.Icons

    End With
    
    '默认加载升级管理
    Set gfrmActive = frmProcUpgrade
    Call FindWindowAndSetActive(gfrmActive)
    gfrmActive.Show
    gfrmActive.ZOrder 0

End Sub

Private Sub Form_Resize()
    tplFunc.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub tplFunc_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    Select Case Item.Id
    Case 1
        Set gfrmActive = frmProcUpgrade
    Case 2
        Set gfrmActive = frmProcManage
    End Select
    
    If Not gfrmActive Is Nothing Then
        Call FindWindowAndSetActive(gfrmActive)
        
        gfrmActive.Show
        gfrmActive.ZOrder 0
    End If
    
End Sub

Private Sub FindWindowAndSetActive(ByVal FrmObj As Form)
    Dim LngTargetHdl As Long
    '--如果该窗体已经打开,则激活它(这样,窗体的大小不会发生变化)--zyb
    LngTargetHdl = FindWindow(vbNullString, FrmObj.Caption)
    If LngTargetHdl <> 0 Then
        If IsIconic(LngTargetHdl) Then
            Call ShowWindow(LngTargetHdl, 9)            '还原指定窗体为原大小
        End If
        Call SetActiveWindow(LngTargetHdl)
    End If
End Sub
