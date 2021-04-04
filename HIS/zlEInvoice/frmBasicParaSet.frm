VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmBasicParaSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "基础参数设置"
   ClientHeight    =   6600
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7350
   Icon            =   "frmBasicParaSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6015
      TabIndex        =   16
      Top             =   1050
      Width           =   1100
   End
   Begin VB.PictureBox picNomal 
      BorderStyle     =   0  'None
      Height          =   4515
      Left            =   240
      ScaleHeight     =   4515
      ScaleWidth      =   5505
      TabIndex        =   1
      Top             =   540
      Width           =   5500
      Begin VB.Frame fra 
         Caption         =   "电子票据换开方式"
         Height          =   1800
         Index           =   0
         Left            =   285
         TabIndex        =   10
         Top             =   2160
         Width           =   2370
         Begin VB.OptionButton Option换开方式 
            Caption         =   "不换开"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   11
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Option换开方式 
            Caption         =   "自动换开"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   12
            Top             =   810
            Width           =   1095
         End
         Begin VB.OptionButton Option换开方式 
            Caption         =   "提示换开"
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   13
            Top             =   1260
            Width           =   1095
         End
      End
      Begin VB.Frame fra 
         Caption         =   "告知单打印方式"
         Height          =   1800
         Index           =   1
         Left            =   2880
         TabIndex        =   6
         Top             =   180
         Width           =   2370
         Begin VB.OptionButton Option告知单打印方式 
            Caption         =   "不打印"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   7
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Option告知单打印方式 
            Caption         =   "自动打印"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   8
            Top             =   810
            Width           =   1095
         End
         Begin VB.OptionButton Option告知单打印方式 
            Caption         =   "提示打印"
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   9
            Top             =   1260
            Width           =   1095
         End
      End
      Begin VB.Frame fra 
         Caption         =   "电子票据打印方式"
         Height          =   1800
         Index           =   2
         Left            =   285
         TabIndex        =   2
         Top             =   180
         Width           =   2370
         Begin VB.OptionButton Option电子票据打印方式 
            Caption         =   "不打印"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   3
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Option电子票据打印方式 
            Caption         =   "自动打印"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   4
            Top             =   810
            Width           =   1215
         End
         Begin VB.OptionButton Option电子票据打印方式 
            Caption         =   "提示打印"
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   5
            Top             =   1260
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmd打印设置 
         Caption         =   "告知单打印设置(&P)"
         Height          =   350
         Left            =   270
         TabIndex        =   14
         Top             =   4110
         Width           =   2370
      End
   End
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   6285
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   5685
      _Version        =   589884
      _ExtentX        =   10028
      _ExtentY        =   11086
      _StockProps     =   64
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6030
      TabIndex        =   15
      Top             =   540
      Width           =   1100
   End
End
Attribute VB_Name = "frmBasicParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngSys As Long
Private mlngModule As Long
Private mstrPrivs As String
Private mobjEinvoiceObj As clsEInvoiceModule
Private mobjFrom As Object
Private mblnOnlyCreateEInvoice As Boolean '是否仅开具电子票据，不换开纸质票据

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim intTmp As Integer
    Dim strSQL As String, blnSetUp As Boolean
    
    blnSetUp = InStr(1, mstrPrivs, ";参数设置;") > 0
    
    If Not mblnOnlyCreateEInvoice Then
        intTmp = IIf(Option换开方式(2).Value, 2, IIf(Option换开方式(1).Value, 1, 0))
        If fra(0).Tag <> intTmp Then zlDatabase.SetPara "票据换开方式", intTmp, mlngSys, mlngModule, blnSetUp
    End If
    
    intTmp = IIf(Option告知单打印方式(2).Value, 2, IIf(Option告知单打印方式(1).Value, 1, 0))
    If fra(1).Tag <> intTmp Then zlDatabase.SetPara "告知单打印方式", intTmp, mlngSys, mlngModule, blnSetUp
    
    intTmp = IIf(Option电子票据打印方式(2).Value, 2, IIf(Option电子票据打印方式(1).Value, 1, 0))
    If fra(2).Tag <> intTmp Then zlDatabase.SetPara "电子票据打印方式", intTmp, mlngSys, mlngModule, blnSetUp
    If Not mobjFrom Is Nothing Then Call mobjFrom.zlSavePara
    Unload Me
End Sub

Private Sub cmd打印设置_Click()
    Call ReportPrintSet(gcnOracle, mlngSys, "ZL1_INSIDE_1145", Me)
End Sub

Private Sub InitPara()
    '初始化参数
    Dim intTmp As Integer, blnSetUp As Boolean
    
    blnSetUp = InStr(1, mstrPrivs, ";参数设置;") > 0
    
    If Not mblnOnlyCreateEInvoice Then
        intTmp = zlDatabase.GetPara("票据换开方式", mlngSys, mlngModule, 0, Array(Option换开方式(0), Option换开方式(1), Option换开方式(2)), blnSetUp)
        fra(0).Tag = intTmp
        Option换开方式(intTmp).Value = True
    End If
    
    intTmp = zlDatabase.GetPara("告知单打印方式", mlngSys, mlngModule, 0, Array(Option告知单打印方式(0), Option告知单打印方式(1), Option告知单打印方式(2)), blnSetUp)
    fra(1).Tag = intTmp
    Option告知单打印方式(intTmp).Value = True
    
    intTmp = zlDatabase.GetPara("电子票据打印方式", mlngSys, mlngModule, 0, Array(Option电子票据打印方式(0), Option电子票据打印方式(1), Option电子票据打印方式(2)), blnSetUp)
    fra(2).Tag = intTmp
    Option电子票据打印方式(intTmp).Value = True

End Sub

Private Sub Form_Load()
    mstrPrivs = ";" & GetPrivFunc(mlngSys, mlngModule) & ";"
    
    fra(0).Visible = Not mblnOnlyCreateEInvoice
    
    Call InitPara
    Call InitPage
End Sub

Public Sub ShowMe(ByVal frmMain As Object, ByVal objEinvoiceObj As clsEInvoiceModule, _
    ByVal lngSys As Long, ByVal lngModule As Long, ByVal blnOnlyCreateEInvoice As Boolean)
    '程序入口
    '入参:
    '   blnOnlyCreateEInvoice - 是否仅开具电子票据，不换开纸质票据
    On Error GoTo errHandle
    mlngSys = lngSys: mlngModule = lngModule
    Set mobjEinvoiceObj = objEinvoiceObj
    mblnOnlyCreateEInvoice = blnOnlyCreateEInvoice
    
    Me.Show 1, frmMain
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mobjFrom Is Nothing Then Unload mobjFrom: Set mobjFrom = Nothing
End Sub

Private Sub Option电子票据打印方式_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Option告知单打印方式_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Option换开方式_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化页面控件
    Dim objItem As TabControlItem
    With tbPage
        Set objItem = .InsertItem(1, "常规", picNomal.hWnd, 0)
        objItem.Tag = 1
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    If Not mobjEinvoiceObj Is Nothing Then
        Call SetControlPosition
        Set mobjFrom = mobjEinvoiceObj.zlGetParaFrom
        If Not mobjFrom Is Nothing Then
             With tbPage
                 Set objItem = .InsertItem(2, mobjFrom.Caption, mobjFrom.hWnd, 0)
                 objItem.Tag = 1
                 .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
                 .PaintManager.BoldSelected = True
                 .PaintManager.Layout = xtpTabLayoutAutoSize
                 .PaintManager.StaticFrame = True
                 .PaintManager.ClientFrame = xtpTabFrameBorder
             End With
        End If
    End If
    tbPage.Item(0).Selected = True
End Sub

Private Sub SetControlPosition()
    '功能：调整控件位置
    tbPage.Width = tbPage.Width + 860
    Me.Width = Me.Width + 860
    cmdOK.Left = cmdOK.Left + 860
    cmdCancel.Left = cmdCancel.Left + 860
    picNomal.Width = picNomal.Width + 430
    fra(0).Left = fra(0).Left + 430: fra(1).Left = fra(1).Left + 430: fra(2).Left = fra(2).Left + 430
    cmd打印设置.Left = cmd打印设置.Left + 430
End Sub
