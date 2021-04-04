VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPreviewDock 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   6090
   ClientLeft      =   -60
   ClientTop       =   -420
   ClientWidth     =   9030
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9030
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   720
      Width1          =   810
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   8910
         _ExtentX        =   15716
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Description     =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "比例"
               Key             =   "Scale"
               Description     =   "比例"
               Object.ToolTipText     =   "显示比例"
               Object.Tag             =   "比例"
               ImageIndex      =   3
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   11
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "原始大小(&O)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "合适页宽(&W)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "合适页高(&H)"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "整页显示(&P)"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "250%"
                  EndProperty
                  BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "200%"
                  EndProperty
                  BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "150%"
                  EndProperty
                  BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "75%"
                  EndProperty
                  BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "50%"
                  EndProperty
                  BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "25%"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Par_"
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "最前"
               Key             =   "First"
               Description     =   "最前"
               Object.ToolTipText     =   "最前页(Home)"
               Object.Tag             =   "最前"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "上页"
               Key             =   "Previous"
               Description     =   "上页"
               Object.ToolTipText     =   "上一页(PageUp)"
               Object.Tag             =   "上页"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "下页"
               Key             =   "Next"
               Description     =   "下页"
               Object.ToolTipText     =   "下一页(PageDown)"
               Object.Tag             =   "下页"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "最后"
               Key             =   "Last"
               Description     =   "最后"
               Object.ToolTipText     =   "最后页(End)"
               Object.Tag             =   "最后"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         Begin VB.ComboBox cboPage 
            Height          =   300
            Left            =   5085
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   195
            Width           =   1185
         End
         Begin VB.TextBox txtPage 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   4515
            Locked          =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            Tag             =   "当前页              共"
            Text            =   "当前页"
            Top             =   255
            Width           =   2985
         End
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   5730
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPreviewDock.frx":0000
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10636
            Object.ToolTipText     =   "打印机"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   18
            Object.ToolTipText     =   "纸张"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   18
            Object.ToolTipText     =   "纸向"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4755
      Left            =   0
      ScaleHeight     =   4755
      ScaleWidth      =   8760
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   720
      Width           =   8760
      Begin VB.PictureBox picPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3390
         Left            =   270
         MouseIcon       =   "frmPreviewDock.frx":0894
         MousePointer    =   99  'Custom
         ScaleHeight     =   3390
         ScaleWidth      =   6990
         TabIndex        =   6
         Top             =   180
         Width           =   6990
      End
      Begin VB.PictureBox picShadow 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3390
         Left            =   330
         ScaleHeight     =   3390
         ScaleWidth      =   6990
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   255
         Width           =   6990
      End
   End
   Begin VB.HScrollBar scrHsc 
      DragIcon        =   "frmPreviewDock.frx":09E6
      Height          =   250
      LargeChange     =   20
      Left            =   0
      Max             =   100
      MouseIcon       =   "frmPreviewDock.frx":0CF0
      SmallChange     =   10
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5475
      Width           =   8760
   End
   Begin VB.VScrollBar scrVsc 
      DragIcon        =   "frmPreviewDock.frx":0E42
      Height          =   4755
      LargeChange     =   20
      Left            =   8775
      Max             =   100
      MouseIcon       =   "frmPreviewDock.frx":114C
      SmallChange     =   10
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   735
      Width           =   250
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   705
      Top             =   465
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewDock.frx":129E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewDock.frx":14B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewDock.frx":16D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewDock.frx":18EC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewDock.frx":1B06
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewDock.frx":1D20
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewDock.frx":1F3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewDock.frx":2154
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewDock.frx":236E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   75
      Top             =   465
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewDock.frx":2588
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewDock.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewDock.frx":29BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewDock.frx":2BD6
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewDock.frx":2DF0
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewDock.frx":300A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewDock.frx":3224
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewDock.frx":343E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewDock.frx":3658
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPreviewDock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public frmParent As Object                              '父窗体对象,用于参数传递；调用预览前,打印机已按设置进行了初始化

Private lngPreX As Long, lngPreY As Long
Private intCurPage As Integer, sngCurScale As Single
Private Const Shadow_W = 60                             '阴影厚度
Private mobjFmt As RPTFmt                               '报表格式对象
Private mintIndex As Integer

Private Sub cboPage_Click()
    Call ShowPage
End Sub

Private Sub cbr_HeightChanged(ByVal NewHeight As Single)
    txtPage.Top = (NewHeight - txtPage.Height) / 2
    cboPage.Top = (NewHeight - cboPage.Height) / 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If ActiveControl Is cboPage Then Exit Sub
    Select Case KeyCode
        Case vbKeyUp
            If scrVsc.Enabled And scrVsc.Value > scrVsc.Min Then
                If Shift = 2 Then
                    scrVsc.Value = IIF(scrVsc.Value - scrVsc.LargeChange < scrVsc.Min, scrVsc.Min, scrVsc.Value - scrVsc.LargeChange)
                Else
                    scrVsc.Value = IIF(scrVsc.Value - scrVsc.SmallChange < scrVsc.Min, scrVsc.Min, scrVsc.Value - scrVsc.SmallChange)
                End If
            End If
        Case vbKeyDown
            If scrVsc.Enabled And scrVsc.Value < scrVsc.Max Then
                If Shift = 2 Then
                    scrVsc.Value = IIF(scrVsc.Value + scrVsc.LargeChange > scrVsc.Max, scrVsc.Max, scrVsc.Value + scrVsc.LargeChange)
                Else
                    scrVsc.Value = IIF(scrVsc.Value + scrVsc.SmallChange > scrVsc.Max, scrVsc.Max, scrVsc.Value + scrVsc.SmallChange)
                End If
            End If
        Case vbKeyLeft
            If scrHsc.Enabled And scrHsc.Value > scrHsc.Min Then
                If Shift = 2 Then
                    scrHsc.Value = IIF(scrHsc.Value - scrHsc.LargeChange < scrHsc.Min, scrHsc.Min, scrHsc.Value - scrHsc.LargeChange)
                Else
                    scrHsc.Value = IIF(scrHsc.Value - scrHsc.SmallChange < scrHsc.Min, scrHsc.Min, scrHsc.Value - scrHsc.SmallChange)
                End If
            End If
        Case vbKeyRight
            If scrHsc.Enabled And scrHsc.Value < scrHsc.Max Then
                If Shift = 2 Then
                    scrHsc.Value = IIF(scrHsc.Value + scrHsc.LargeChange > scrHsc.Max, scrHsc.Max, scrHsc.Value + scrHsc.LargeChange)
                Else
                    scrHsc.Value = IIF(scrHsc.Value + scrHsc.SmallChange > scrHsc.Max, scrHsc.Max, scrHsc.Value + scrHsc.SmallChange)
                End If
            End If
        Case vbKeyHome
            mnuView_Move_Click 0
        Case vbKeyEnd
            mnuView_Move_Click 3
        Case vbKeyPageUp
            mnuView_Move_Click 1
        Case vbKeyPageDown
            mnuView_Move_Click 2
    End Select
End Sub

Private Sub ExchangeValue(a As Variant, B As Variant)
'功能：交换两个值
    Dim C As Variant
    C = a: a = B: B = C
End Sub

Private Function IntEx(vNumber As Variant) As Variant
'功能：取大于指定数值的最小整数
    IntEx = -1 * Int(-1 * Val(vNumber))
End Function

Private Function GetParityCount(ByVal intB As Integer, ByVal intE As Integer, ByVal intParity As Integer) As Long
'功能：计算指定范围的奇数或者偶数个数
'参数：intB,intE=计算数字范围(x=1-n)
'      intParity=0-所有个数,1-计算奇数个数,2-计算偶数个数
    If intB > intE Then
        Call ExchangeValue(intE, intB)
    End If
    
    If intParity = 0 Then
        GetParityCount = intE - intB + 1
    Else
        If intB = intE Then
            GetParityCount = IIF((intB Mod 2) = (intParity Mod 2), 1, 0)
        Else
            If intParity = 1 Then
                If (intB Mod 2) = 0 Then intB = intB + 1
                If (intE Mod 2) = 0 Then intE = intE - 1
            ElseIf intParity = 2 Then
                If (intB Mod 2) = 1 Then intB = intB + 1
                If (intE Mod 2) = 1 Then intE = intE - 1
            End If
            GetParityCount = IntEx((intE - intB + 1) / 2)
        End If
    End If
End Function

Private Function GetParityEnd(ByVal intB As Integer, ByVal intE As Integer, ByVal intParity As Integer) As Long
'功能：计算指定范围的最后一个奇偶数
'参数：intB,intE=计算数字范围(x=1-n)
'      intParity=0-所有个数,1-计算奇数个数,2-计算偶数个数
    If intParity = 1 Then
        If (intE Mod 2) = 0 Then intE = intE - IIF(intB > intE, -1, 1)
    ElseIf intParity = 2 Then
        If (intE Mod 2) = 1 Then intE = intE - IIF(intB > intE, -1, 1)
    End If
    GetParityEnd = intE
End Function

Private Sub mnuFile_Print_Click()
    Dim intB As Integer, intE As Integer
    Dim i As Integer, k As Integer, j As Integer
    Dim lngPrintH As Long, blnDo As Boolean
    Dim blnCancel As Boolean, blnReset As Boolean
    Dim objReport As Report, objFmt As RPTFmt
    Dim objCurDLL As clsReport, arrBill As Variant
    Dim intParity As Integer, intCopy As Integer
    Dim intParityCount As Integer, intParityTotal As Integer
    Dim frmNewPrint As New frmPrint
    
    Set objReport = frmParent.mobjReport
    'Set objFmt = objReport.Fmts("_" & objReport.bytFormat)
    Set objFmt = mobjFmt
    Set objCurDLL = frmParent.mobjCurDLL
    
    frmNewPrint.mstr编号 = objReport.编号
    frmNewPrint.mbln票据 = objReport.票据
    
    If IsArray(frmParent.marrPage) Then
        frmNewPrint.mintMax = UBound(frmParent.marrPage) + 1
    End If
    If IsArray(frmParent.marrPageCard) Then
        i = UBound(frmParent.marrPageCard) + 1
    Else
        i = 1
    End If
    If i > frmNewPrint.mintMax Then frmNewPrint.mintMax = i
    
    frmNewPrint.Show 1, Me
    If frmNewPrint.mblnOK Then
        If frmNewPrint.optPage(0).Value Then '所有页
            intB = 0: intE = cboPage.ListCount - 1
        ElseIf frmNewPrint.optPage(1).Value Then '当前页
            intB = cboPage.ListIndex: intE = intB
        ElseIf frmNewPrint.optPage(2).Value Then '指定页
            intB = Val(frmNewPrint.txtBegin.Text) - 1
            intE = Val(frmNewPrint.txtEnd.Text) - 1
        ElseIf frmNewPrint.optPage(3).Value Then '奇数页
            intParity = 1
            intB = 0: intE = cboPage.ListCount - 1
            If frmNewPrint.chkOrder.Value = 1 Then
                Call ExchangeValue(intB, intE)
            End If
        ElseIf frmNewPrint.optPage(4).Value Then '偶数页
            intParity = 2
            intB = 0: intE = cboPage.ListCount - 1
            If frmNewPrint.chkOrder.Value = 1 Then
                Call ExchangeValue(intB, intE)
            End If
        End If
        intCopy = frmNewPrint.txtCopy.Text
        Unload frmNewPrint
        
        On Error GoTo errH
                
        '再次初始化打印机
        If Not InitPrinter(frmParent, intCopy) Then
            MsgBox "设备初始化失败.可能是系统没有安装打印机或与当前设置不兼容！", vbInformation, App.Title
            gblnError = True: Exit Sub
        End If
        k = intCopy '缺省为强制循环打印k份
        If Printer.Copies = intCopy Then k = 1 '支持时使用打印机功能
        
        '输出到打印之前，激活打印事件
        If Not objCurDLL Is Nothing Then
            blnCancel = False: i = 1
            If IsArray(frmParent.marrPage) Then i = UBound(frmParent.marrPage) + 1
            
            If intB = 0 And intE = cboPage.ListCount - 1 And intParity = 0 Then
                '与全部打印等同
                Call objCurDLL.Act_BeforePrint(frmParent.mobjReport.编号, i * intCopy, blnCancel, arrBill)
            ElseIf intB = cboPage.ListIndex And intE = intB Then
                '当前页
                Call objCurDLL.Act_BeforePrint(frmParent.mobjReport.编号, -1, blnCancel, arrBill)
            Else
                '指定范围页,包括奇数页、偶数页
                Call objCurDLL.Act_BeforePrint(frmParent.mobjReport.编号, -2, blnCancel, arrBill)
            End If
            
            If blnCancel Then Exit Sub
            
            '实际要打印的票据数据
            If IsArray(arrBill) Then garrBill = arrBill
        End If
    
        Me.Refresh
        Screen.MousePointer = 11
        Do
            k = k - 1
            j = j + 1
            If intE = intB Then
                If Printer.Copies <> intCopy And intCopy <> 1 Then
                    ShowFlash "输出" & objReport.名称 & ",共 1 页 " & intCopy & " 份,当前第 " & j & " 份", j / intCopy
                Else
                    ShowFlash "输出" & objReport.名称 & "…", 1
                End If
                
                '动态计算及设置纸张高度
                If objFmt.动态纸张 And objFmt.纸向 = 1 Then
                    Call PrintPage(intB, Me, frmParent, 1, False, True, lngPrintH)
                    blnDo = lngPrintH > 0 And lngPrintH < objFmt.H
                    If blnDo Then '空白部份高于30mm且高于原纸张的1/8
                        blnDo = objFmt.H - lngPrintH > 30 * Twip_mm And objFmt.H - lngPrintH > objFmt.H / 8
                    End If
                    If blnDo Then
                        lngPrintH = lngPrintH + 567 '比实际打印多留10mm高度
                        If Not SetPrinterPaper(frmParent.hwnd, objReport, lngPrintH, intCopy) Then
                            '设置失败时恢复成原始纸张
                            Call ResetPrinterPaper(frmParent.hwnd, objReport, intCopy)
                        End If
                    End If
                End If
                
                Call PrintPage(intB, Printer, frmParent)
            Else
                '注意：奇偶计算函数是以1-n为准计算，页号变量是0-n
                intParityCount = 0
                intParityTotal = GetParityCount(intB + 1, intE + 1, intParity)
                For i = intB To intE Step IIF(intB > intE, -1, 1)
                    If intParity = 0 Or ((i + 1) Mod 2) = (intParity Mod 2) Then
                        intParityCount = intParityCount + 1
                        If Printer.Copies <> intCopy And intCopy <> 1 Then
                            ShowFlash "输出" & objReport.名称 & ",共 " & intParityTotal & " 页 " & intCopy & " 份,当前第 " & j & " 份", _
                                (intParityCount + (j - 1) * intParityTotal) / (intParityTotal * intCopy)
                        Else
                            ShowFlash "输出" & objReport.名称 & ",共 " & intParityTotal & " 页,当前第 " & intParityCount & " 页…", intParityCount / intParityTotal
                        End If
                        
                        '动态计算及设置纸张高度
                        If objFmt.动态纸张 And objFmt.纸向 = 1 Then
                            Call PrintPage(i, Me, frmParent, 1, False, True, lngPrintH)
                            blnDo = lngPrintH > 0 And lngPrintH < objFmt.H
                            If blnDo Then '空白部份高于30mm且高于原纸张的1/8
                                blnDo = objFmt.H - lngPrintH > 30 * Twip_mm And objFmt.H - lngPrintH > objFmt.H / 8
                            End If
                            If blnDo Then
                                lngPrintH = lngPrintH + 567 '比实际打印多留10mm高度
                                If Not SetPrinterPaper(frmParent.hwnd, objReport, lngPrintH, intCopy) Then
                                    '设置失败时恢复成原始纸张
                                    Call ResetPrinterPaper(frmParent.hwnd, objReport, intCopy)
                                    blnReset = False
                                Else
                                    blnReset = True '本页已设置过动态纸张,下页计算出不必设时要恢复成原始的
                                End If
                            ElseIf blnReset Then
                                Call ResetPrinterPaper(frmParent.hwnd, objReport, intCopy)
                                blnReset = False
                            End If
                        End If
                        
                        If Not PrintPage(i, Printer, frmParent) Then Exit For
                        If i <> GetParityEnd(intB + 1, intE + 1, intParity) - 1 Then Printer.NewPage
                    End If
                Next
            End If
            If k > 0 Then Printer.NewPage
        Loop Until k = 0
        
        Printer.EndDoc
        
        ShowFlash
        Screen.MousePointer = 0
        
        '输出到打印结束，激活打印事件
        If Not objCurDLL Is Nothing Then
            Call objCurDLL.Act_AfterPrint(frmParent.mobjReport.编号)
        End If
    End If
    Exit Sub
errH:
    Screen.MousePointer = 0
    Call ShowFlash
    Printer.KillDoc
    MsgBox Err.Number & ":" & Err.Description & vbCrLf & "打印过程被强行中断！", vbExclamation, App.Title
    Err.Clear
End Sub

Private Sub mnuFile_Setup_Click()
    If Not ReportLocalSet(frmParent.mobjReport.系统, frmParent.mobjReport.编号, False, frmParent.mobjReport.bytFormat, Me) Then Exit Sub
    If Not InitPrinter(frmParent) Then
        gblnError = True
        MsgBox "设备初始化失败.可能是系统没有安装打印机或与当前设置不兼容！", vbInformation, App.Title: Exit Sub
    End If
    Call Form_Load
End Sub

Private Sub mnuHelpTitle_Click()
    If frmParent.Tag = "" Then
        Call ShowHelpRpt(Me.hwnd, frmParent.mobjReport.编号, Int((frmParent.mobjReport.系统) / 100))
    Else
        Call ShowHelpRpt(Me.hwnd, frmParent.Tag, Int((frmParent.mobjReport.系统) / 100))
    End If

End Sub

Private Sub mnuView_Move_Click(Index As Integer)
    With cboPage
        Select Case Index
            Case 0
                .ListIndex = 0
            Case 1
                If .ListIndex - 1 >= 0 Then .ListIndex = .ListIndex - 1
            Case 2
                If .ListIndex + 1 <= .ListCount - 1 Then .ListIndex = .ListIndex + 1
            Case 3
                .ListIndex = .ListCount - 1
        End Select
    End With
    Call ShowPage
End Sub

Private Sub mnuView_reFlash_Click()
    Me.Refresh
End Sub

Private Sub picback_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If scrVsc.Enabled Then
            If (Y - lngPreY) / 15 > 0 Then
                scrVsc.Value = IIF(scrVsc.Value - (Y - lngPreY) / 15 < scrVsc.Min, scrVsc.Min, scrVsc.Value - (Y - lngPreY) / 15)
            Else
                scrVsc.Value = IIF(scrVsc.Value - (Y - lngPreY) / 15 > scrVsc.Max, scrVsc.Max, scrVsc.Value - (Y - lngPreY) / 15)
            End If
        End If
        If scrHsc.Enabled Then
            If (X - lngPreX) / 15 > 0 Then
                scrHsc.Value = IIF(scrHsc.Value - (X - lngPreX) / 15 < scrHsc.Min, scrHsc.Min, scrHsc.Value - (X - lngPreX) / 15)
            Else
                scrHsc.Value = IIF(scrHsc.Value - (X - lngPreX) / 15 > scrHsc.Max, scrHsc.Max, scrHsc.Value - (X - lngPreX) / 15)
            End If
        End If
    End If
End Sub

Private Sub picPage_DblClick()
    tbr_ButtonClick tbr.Buttons("Scale")
End Sub

Private Sub picPage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If scrVsc.Enabled Then
            If (Y - lngPreY) / 15 > 0 Then
                scrVsc.Value = IIF(scrVsc.Value - (Y - lngPreY) / 15 < scrVsc.Min, scrVsc.Min, scrVsc.Value - (Y - lngPreY) / 15)
            Else
                scrVsc.Value = IIF(scrVsc.Value - (Y - lngPreY) / 15 > scrVsc.Max, scrVsc.Max, scrVsc.Value - (Y - lngPreY) / 15)
            End If
        End If
        If scrHsc.Enabled Then
            If (X - lngPreX) / 15 > 0 Then
                scrHsc.Value = IIF(scrHsc.Value - (X - lngPreX) / 15 < scrHsc.Min, scrHsc.Min, scrHsc.Value - (X - lngPreX) / 15)
            Else
                scrHsc.Value = IIF(scrHsc.Value - (X - lngPreX) / 15 > scrHsc.Max, scrHsc.Max, scrHsc.Value - (X - lngPreX) / 15)
            End If
        End If
    End If
End Sub

Private Sub picPage_GotFocus()
    Oldwinproc = GetWindowLong(picPage.hwnd, GWL_WNDPROC)
    SetWindowLong picPage.hwnd, GWL_WNDPROC, AddressOf FlexScroll
End Sub

Private Sub picPage_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        '下
        If scrVsc.Value + scrVsc.Max / 10 > scrVsc.Max Then
            scrVsc.Value = scrVsc.Max
        Else
            scrVsc.Value = scrVsc.Value + scrVsc.Max / 10
        End If
    ElseIf KeyCode = vbKeyPageUp Then
        '上
        If scrVsc.Value - scrVsc.Max / 10 < 0 Then
            scrVsc.Value = 0
        Else
            scrVsc.Value = scrVsc.Value - scrVsc.Max / 10
        End If
    End If
End Sub

Private Sub picPage_LostFocus()
    SetWindowLong picPage.hwnd, GWL_WNDPROC, Oldwinproc
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long '工具条占用高度
    Dim staH As Long '状态栏占用高度
    
    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    '靠齐控件宽度和高度
    cbrH = IIF(cbr.Visible, cbr.Height, 0)
    staH = IIF(sta.Visible, sta.Height, 0)

    picBack.Top = ScaleTop + cbrH
    picBack.Left = ScaleLeft
    picBack.Width = ScaleWidth - scrVsc.Width
    picBack.Height = ScaleHeight - staH - cbrH - scrHsc.Height
    
    scrVsc.Top = picBack.Top
    scrVsc.Left = ScaleWidth - scrVsc.Width
    scrVsc.Height = picBack.Height
    
    scrHsc.Left = picBack.Left
    scrHsc.Top = picBack.Top + picBack.Height
    scrHsc.Width = picBack.Width
    
    '调整预览页
    
    If picBack.ScaleWidth >= picPage.Width + Shadow_W * 4 Then
        picPage.Left = (picBack.ScaleWidth - (picPage.Width + Shadow_W * 4)) / 2 + Shadow_W * 2
        picShadow.Left = picPage.Left + Shadow_W
        scrHsc.Enabled = False
    Else
        scrHsc.Max = (picPage.Width + Shadow_W * 4 - picBack.ScaleWidth) / 15
        If scrHsc.Max / 3 < scrHsc.SmallChange Then
            scrHsc.LargeChange = scrHsc.SmallChange
        Else
            scrHsc.LargeChange = scrHsc.Max / 3
        End If
        scrHsc.Value = 0
        scrHsc.Enabled = True
        scrhsc_Change
    End If
    If picBack.ScaleHeight >= picPage.Height + Shadow_W * 4 Then
        picPage.Top = (picBack.ScaleHeight - (picPage.Height + Shadow_W * 4)) / 2 + Shadow_W
        picShadow.Top = picPage.Top + Shadow_W
        scrVsc.Enabled = False
    Else
        scrVsc.Max = (picPage.Height + Shadow_W * 4 - picBack.ScaleHeight) / 15
        If scrVsc.Max / 3 < scrVsc.SmallChange Then
            scrVsc.LargeChange = scrVsc.SmallChange
        Else
            scrVsc.LargeChange = scrVsc.Max / 3
        End If
        scrVsc.Value = 0
        scrVsc.Enabled = True
        scrVsc_Change
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    intCurPage = 0
    Set mobjFmt = Nothing
    Unload frmFlash
    
    SaveWinState Me, App.ProductName, frmParent.mobjReport.编号
    Unload frmParent
    Set frmParent = Nothing
End Sub

Private Sub mnuFile_Quit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me)
End Sub

Private Sub picPage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Set picPage.MouseIcon = scrVsc.MouseIcon
End Sub

Private Sub scrVsc_Change()
    picPage.Top = -scrVsc.Value * 15# + Shadow_W * 2
    picShadow.Top = picPage.Top + Shadow_W
    Me.Refresh
End Sub

Private Sub scrVsc_Scroll()
    picPage.Top = -scrVsc.Value * 15# + Shadow_W * 2
    picShadow.Top = picPage.Top + Shadow_W
    Me.Refresh
End Sub

Private Sub scrhsc_Change()
    picPage.Left = -scrHsc.Value * 15# + Shadow_W * 2
    picShadow.Left = picPage.Left + Shadow_W
    Me.Refresh
End Sub

Private Sub scrhsc_Scroll()
    picPage.Left = -scrHsc.Value * 15# + Shadow_W * 2
    picShadow.Left = picPage.Left + Shadow_W
    Me.Refresh
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_Quit_Click
        Case "Scale"
            If mintIndex >= 9 Then
                mintIndex = 0
            Else
                mintIndex = mintIndex + 1
            End If
            Call ShowPage
        Case "First"
            mnuView_Move_Click 0
        Case "Previous"
            mnuView_Move_Click 1
        Case "Next"
            mnuView_Move_Click 2
        Case "Last"
            mnuView_Move_Click 3
        Case "Print"
            mnuFile_Print_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
End Sub

Private Sub tbr_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Index <= 4 Then
        mintIndex = ButtonMenu.Index - 1
    Else
        mintIndex = ButtonMenu.Index - 2
    End If
    
    Call ShowPage
End Sub

Private Sub txtPage_GotFocus()
    cboPage.SetFocus
End Sub

Private Sub Form_Load()
    LoadForm
End Sub

Public Sub LoadForm(Optional intMode As Integer)
    Dim objReport As Report
    Dim objFmt As RPTFmt
    Dim i As Integer
    Dim strPrivs As String
    Dim lngPage As Long
    Dim lngPage1 As Long
    
    intCurPage = -1
    If Not Me.Visible Then
        On Error GoTo errH
        If intMode = 0 Then
            RestoreWinState Me, App.ProductName, frmParent.mobjReport.编号
        Else
            '如果是嵌入式的则隐藏工具栏菜单栏
            tbr.Buttons("Quit").Visible = False
        End If
        
        '初始页号选择
        If Not IsArray(frmParent.marrPage) Then
            lngPage = 1
        Else
            lngPage = UBound(frmParent.marrPage) + 1
        End If
        
        If Not IsArray(frmParent.marrPageCard) Then
            lngPage1 = 1
        Else
            lngPage1 = UBound(frmParent.marrPageCard) + 1
        End If
        If lngPage1 > lngPage Then lngPage = lngPage1
        
        For i = 1 To lngPage
            cboPage.AddItem "第 " & i & " 页"
        Next
        
        cboPage.ListIndex = 0
        txtPage.Text = txtPage.Tag & " " & cboPage.ListCount & " 页"
    Else
        '重复调用的情况
        Call cboPage_Click
    End If
    
    Set objReport = frmParent.mobjReport
    Set mobjFmt = objReport.Fmts("_" & objReport.bytFormat)
    
    '描述
    Caption = "打印预览 - " & objReport.名称 & IIF(objReport.说明 = "", "", "：" & objReport.说明)
    sta.Panels(2) = Printer.Port & Printer.DeviceName
    sta.Panels(3) = GetPaperName(Printer.PaperSize, mobjFmt.W, mobjFmt.H)
    sta.Panels(4) = IIF(Printer.Orientation = 1, "纵向", "横向")
    
    '打印权限判断
    strPrivs = GetPrivFunc(0, 16)
    If InStr(";" & strPrivs & ";", ";打印;") = 0 Or frmParent.mblnDisabledPrint Then
        tbr.Buttons("Print").Visible = False
        tbr.Buttons(2).Visible = False
    End If
    Exit Sub
errH:
    Err.Clear
End Sub

Private Function GetScale() As Single
'功能：返回当前显示比例
    Dim i As Integer

    Select Case mintIndex
        Case 0 '原始大小
            GetScale = 1
        Case 1 '合适页宽
            GetScale = (picBack.ScaleWidth - Shadow_W * 4) / Printer.Width
        Case 2 '合适页高
            GetScale = (picBack.ScaleHeight - Shadow_W * 4) / Printer.Height
        Case 3 '整页显示
            If picBack.ScaleWidth / Printer.Width < picBack.ScaleHeight / Printer.Height Then
                GetScale = (picBack.ScaleWidth - Shadow_W * 4) / Printer.Width
            Else
                GetScale = (picBack.ScaleHeight - Shadow_W * 4) / Printer.Height
            End If
        Case 4
            GetScale = CDbl(Val("250%") / 100)
        Case 5
            GetScale = CDbl(Val("200%") / 100)
        Case 6
            GetScale = CDbl(Val("150%") / 100)
        Case 7
            GetScale = CDbl(Val("75%") / 100)
        Case 8
            GetScale = CDbl(Val("50%") / 100)
        Case 9
            GetScale = CDbl(Val("25%") / 100)
    End Select
End Function

Private Sub ShowPage()
'功能：显示当前选择页
    Dim sngScale As Single
    
    sngScale = GetScale
    If sngScale = 0 Then sngScale = 1
    
    '不重复处理
    If intCurPage = cboPage.ListIndex And sngCurScale = sngScale Then Exit Sub
    
    '打印机纸张宽/高会因纸向设置而自动改变,因此这时不用管纸向
    picPage.Cls
    If mobjFmt Is Nothing Then
        picPage.Width = Printer.Width * sngScale
        picPage.Height = Printer.Height * sngScale
    Else
        If mobjFmt.纸向 = Val("1-竖向") Then
            picPage.Width = mobjFmt.W
            picPage.Height = mobjFmt.H
        Else
            picPage.Width = mobjFmt.H
            picPage.Height = mobjFmt.W
        End If
    End If
    
    picShadow.Width = picPage.Width
    picShadow.Height = picPage.Height

    intCurPage = cboPage.ListIndex
    sngCurScale = sngScale
    
    Form_Resize
    Screen.MousePointer = 11
    Me.Refresh

    LockWindowUpdate picPage.hwnd
    PrintPage cboPage.ListIndex, picPage, frmParent, sngScale
    LockWindowUpdate 0

    Screen.MousePointer = 0
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hwnd
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub


