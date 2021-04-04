VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmTestWindow 
   Caption         =   "Form1"
   ClientHeight    =   7560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   11250
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture3 
      Height          =   1515
      Left            =   2535
      ScaleHeight     =   1455
      ScaleWidth      =   1470
      TabIndex        =   9
      Top             =   6030
      Width           =   1530
   End
   Begin VB.PictureBox Picture2 
      Height          =   7350
      Left            =   4185
      ScaleHeight     =   7290
      ScaleWidth      =   6870
      TabIndex        =   7
      Top             =   75
      Visible         =   0   'False
      Width           =   6930
   End
   Begin VB.PictureBox Picture1 
      Height          =   7215
      Left            =   195
      ScaleHeight     =   7155
      ScaleWidth      =   3900
      TabIndex        =   0
      Top             =   75
      Width           =   3960
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   435
         Left            =   315
         TabIndex        =   8
         Top             =   180
         Width           =   885
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   525
         Left            =   2235
         TabIndex        =   6
         Top             =   300
         Width           =   990
      End
      Begin VB.CommandButton 配置 
         Caption         =   "打开配置"
         Height          =   810
         Left            =   1995
         TabIndex        =   5
         Top             =   2190
         Width           =   1020
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   810
         Left            =   2175
         TabIndex        =   4
         Top             =   1080
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         Height          =   3885
         Left            =   195
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "frmTestWindow.frx":0000
         Top             =   3270
         Width           =   3420
      End
      Begin VB.CommandButton Command2 
         Caption         =   "打开报告编辑"
         Height          =   735
         Left            =   450
         TabIndex        =   2
         Top             =   1815
         Width           =   1350
      End
      Begin VB.CommandButton Command1 
         Caption         =   "打开浮动窗口"
         Height          =   675
         Left            =   390
         TabIndex        =   1
         Top             =   675
         Width           =   1275
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   4350
      Top             =   180
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmTestWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjVideo0 As clsPacsCapture



Private Sub Command1_Click()
    mobjVideo0.zlShowPopupVideo
End Sub

Private Sub InitFace()
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane
    
    With Me.dkpMain
        .VisualTheme = ThemeOffice2003

        .Options.HideClient = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
        
        .PanelPaintManager.BoldSelected = True
        .TabPaintManager.position = xtpTabPositionLeft  'TAB放到左边显示
'        .TabPaintManager.OneNoteColors = True           '一个TAB一种颜色显示
        .TabPaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .TabPaintManager.BoldSelected = True
        dkpMain.Options.DefaultPaneOptions = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    End With
    
    '注册表中保存的界面布局Pnae数量不对，则加载默认的Pane设置
    If dkpMain.PanesCount <> 3 Then
        dkpMain.DestroyAll
        
        Set Pane1 = dkpMain.CreatePane(1, 3, 3, DockLeftOf, Nothing)
        Pane1.Title = "操控列表"
        Pane1.Handle = Picture1.hWnd
'        Pane1.Options = PaneNoCloseable
        
        Set Pane2 = dkpMain.CreatePane(2, 3, 3, DockTopOf, Pane1)
        Pane2.Title = "子窗体1"
        Pane2.Handle = mobjVideo0.ContainerHwnd
'        Pane2.Options = PaneNoCloseable
        
        Set Pane3 = dkpMain.CreatePane(3, 3, 3, DockTopOf, Pane1)
        Pane3.Title = "Test"
        Pane3.Handle = Picture3.hWnd
'        Pane3.Options = PaneNoCloseable
    End If
End Sub
    
Private Sub Command2_Click()
    Dim objTest As New frmTestWindow
    objTest.Tag = 1

    
    objTest.Show 0, Me

End Sub

Private Sub Command3_Click()
    frmCaptureHint.Show 0, Me
End Sub




Private Sub Command4_Click()
    frmOpenStudyList.Show 0, Me
End Sub

Private Sub Command5_Click()
    frmImages.Show
    frmImages.Caption = frmImages.hWnd
    

    Call frmImages.RefreshImage(901, 901, False, True)
End Sub

Private Sub Form_Load()
'SetWindowPos Me.hwnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3 '将窗口置顶
    

    Set gobjOwner = Me
    

    Set mobjVideo0 = New clsPacsCapture
    mobjVideo0.ParentWindowKey = "TestWindow"

    Caption = mobjVideo0.ContainerHwnd  ' Me.hwnd
'    mobjVideo0.ContainerObj.Left = 0
'    mobjVideo0.ContainerObj.Top = 0
'    mobjVideo0.ContainerObj.Width = Picture2.Width
'    mobjVideo0.ContainerObj.Height = Picture2.Height
'
'    SetParent mobjVideo0.ContainerHwnd, Picture2.hWnd
    InitFace

    Call mobjVideo0.zlInitModule(gcnVideoOracle, 100, 1291, ";视频采集;采集参数设置;", 66, Me.hWnd, Me, True)
    mobjVideo0.zlUpdateStudyInf 901, 0, 2, False
    mobjVideo0.zlRefreshVideoWindow

    Call mobjVideo0.zlRefreshData(True)

'    gblnTestState = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If glngInstanceCount <= 1 Then Call mobjVideo0.zlNotifyQuit
    Set mobjVideo0 = Nothing
End Sub

Private Sub 配置_Click()
    mobjVideo0.zlShowVideoConfig
End Sub
