VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrintAsk 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "打印选项"
   ClientHeight    =   3630
   ClientLeft      =   2550
   ClientTop       =   2625
   ClientWidth     =   6525
   HelpContextID   =   10322
   Icon            =   "frmPrintAsk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNumber 
      Height          =   270
      Left            =   1140
      MaxLength       =   2
      TabIndex        =   33
      Text            =   "1"
      Top             =   2310
      Width           =   270
   End
   Begin VB.CommandButton Command1 
      Caption         =   "输出到&Excel"
      Height          =   350
      Left            =   5040
      TabIndex        =   3
      Top             =   1680
      Width           =   1320
   End
   Begin MSComCtl2.UpDown updnEmpty 
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   19
      Top             =   1530
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtEmpty(0)"
      BuddyDispid     =   196610
      BuddyIndex      =   0
      OrigLeft        =   1560
      OrigTop         =   1635
      OrigRight       =   1800
      OrigBottom      =   1935
      Max             =   60
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtEmpty 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   290
      Index           =   2
      Left            =   4050
      TabIndex        =   24
      Top             =   1530
      Width           =   360
   End
   Begin VB.OptionButton OptPageNo 
      Caption         =   "右对齐(&R)"
      Height          =   255
      Index           =   2
      Left            =   3840
      TabIndex        =   29
      Top             =   1965
      Value           =   -1  'True
      Width           =   1110
   End
   Begin VB.OptionButton OptPageNo 
      Caption         =   "居中(&M)"
      Height          =   255
      Index           =   1
      Left            =   2820
      TabIndex        =   28
      Top             =   1980
      Width           =   1080
   End
   Begin VB.OptionButton OptPageNo 
      Caption         =   "左对齐(&L)"
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   27
      Top             =   1980
      Width           =   1155
   End
   Begin VB.CheckBox chkPageNo 
      Caption         =   "打印页码(&A)："
      Height          =   240
      Left            =   240
      TabIndex        =   26
      Top             =   2010
      Value           =   1  'Checked
      Width           =   1500
   End
   Begin VB.TextBox txtFontStyleApp 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   280
      Left            =   3525
      TabIndex        =   14
      Top             =   990
      Width           =   765
   End
   Begin VB.TextBox txtFontSizeApp 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   280
      Left            =   3105
      TabIndex        =   13
      Top             =   990
      Width           =   435
   End
   Begin VB.TextBox txtFontStyleTitl 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   280
      Left            =   3525
      TabIndex        =   9
      Top             =   585
      Width           =   765
   End
   Begin VB.TextBox txtFontSizeTitl 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   280
      Left            =   3105
      TabIndex        =   8
      Top             =   585
      Width           =   435
   End
   Begin VB.CommandButton cmd预览 
      Caption         =   "预览(&V)"
      Height          =   350
      Left            =   5025
      TabIndex        =   2
      Top             =   1230
      Width           =   1320
   End
   Begin VB.CommandButton cmd打印 
      Caption         =   "打印(&P)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5025
      TabIndex        =   1
      Top             =   765
      Width           =   1320
   End
   Begin VB.TextBox txtFontNameApp 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   280
      Left            =   1080
      TabIndex        =   12
      Top             =   990
      Width           =   2040
   End
   Begin VB.CommandButton cmdFontApp 
      Caption         =   "…"
      Height          =   300
      Left            =   4335
      TabIndex        =   15
      Top             =   990
      Width           =   300
   End
   Begin VB.CommandButton cmdFontTitl 
      Caption         =   "…"
      Height          =   300
      Left            =   4335
      TabIndex        =   10
      Top             =   585
      Width           =   300
   End
   Begin VB.TextBox txtFontNameTitl 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   280
      Left            =   1080
      TabIndex        =   7
      Top             =   585
      Width           =   2040
   End
   Begin VB.CommandButton cmd设置 
      Caption         =   "设置(&S)..."
      Height          =   360
      Left            =   5145
      TabIndex        =   30
      Top             =   3090
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      Height          =   280
      Left            =   1080
      TabIndex        =   5
      Top             =   195
      Width           =   3585
   End
   Begin VB.CommandButton cmd关闭 
      Cancel          =   -1  'True
      Caption         =   "关闭(&C)"
      Height          =   350
      Left            =   5025
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   165
      Width           =   1320
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   105
      Top             =   3045
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtEmpty 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   290
      Index           =   0
      Left            =   1785
      TabIndex        =   18
      Top             =   1530
      Width           =   390
   End
   Begin MSComCtl2.UpDown updnEmpty 
      Height          =   285
      Index           =   1
      Left            =   3210
      TabIndex        =   22
      Top             =   1530
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtEmpty(1)"
      BuddyDispid     =   196610
      BuddyIndex      =   1
      OrigLeft        =   2970
      OrigTop         =   1635
      OrigRight       =   3210
      OrigBottom      =   1935
      Max             =   60
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown updnEmpty 
      Height          =   285
      Index           =   2
      Left            =   4395
      TabIndex        =   25
      Top             =   1530
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtEmpty(2)"
      BuddyDispid     =   196610
      BuddyIndex      =   2
      OrigLeft        =   4410
      OrigTop         =   1635
      OrigRight       =   4650
      OrigBottom      =   1935
      Max             =   60
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtEmpty 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   290
      Index           =   1
      Left            =   2880
      TabIndex        =   21
      Top             =   1530
      Width           =   375
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   270
      Left            =   1410
      TabIndex        =   32
      Top             =   2310
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   476
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtNumber"
      BuddyDispid     =   196636
      OrigLeft        =   1470
      OrigTop         =   2310
      OrigRight       =   1710
      OrigBottom      =   2580
      Max             =   99
      Min             =   1
      SyncBuddy       =   -1  'True
      Wrap            =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "份数(&C):"
      Height          =   180
      Left            =   210
      TabIndex        =   34
      Top             =   2340
      Width           =   720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   8000
      Y1              =   2670
      Y2              =   2670
   End
   Begin VB.Label lblEmpty 
      AutoSize        =   -1  'True
      Caption         =   "下"
      Height          =   180
      Index           =   2
      Left            =   3870
      TabIndex        =   23
      Top             =   1575
      Width           =   180
   End
   Begin VB.Label lblEmpty 
      AutoSize        =   -1  'True
      Caption         =   "上"
      Height          =   180
      Index           =   1
      Left            =   2685
      TabIndex        =   20
      Top             =   1575
      Width           =   180
   End
   Begin VB.Label lblEmpty 
      AutoSize        =   -1  'True
      Caption         =   "左"
      Height          =   180
      Index           =   0
      Left            =   1575
      TabIndex        =   17
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label lblEmpty 
      AutoSize        =   -1  'True
      Caption         =   "页边距(毫米)："
      Height          =   180
      Index           =   3
      Left            =   225
      TabIndex        =   16
      Top             =   1590
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   0
      X2              =   8000
      Y1              =   2670
      Y2              =   2670
   End
   Begin VB.Label lbl正文字体 
      AutoSize        =   -1  'True
      Caption         =   "附项字体"
      Height          =   180
      Left            =   240
      TabIndex        =   11
      Top             =   1050
      Width           =   720
   End
   Begin VB.Label lbl标题字体 
      AutoSize        =   -1  'True
      Caption         =   "标题字体"
      Height          =   180
      Left            =   240
      TabIndex        =   6
      Top             =   645
      Width           =   720
   End
   Begin VB.Label lblPrintStatus 
      Caption         =   "打印机信息"
      Height          =   660
      Left            =   225
      TabIndex        =   31
      Top             =   2745
      Width           =   4965
   End
   Begin VB.Label lbl标题 
      AutoSize        =   -1  'True
      Caption         =   "打印标题"
      Height          =   180
      Left            =   255
      TabIndex        =   4
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmPrintAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public byRunMode As Byte       '执行方式

Private Sub chkPageNo_Click()
    If chkPageNo.Value = 0 Then
        OptPageNo(0).Value = False
        OptPageNo(1).Value = False
        OptPageNo(2).Value = False
        OptPageNo(0).Enabled = False
        OptPageNo(1).Enabled = False
        OptPageNo(2).Enabled = False
            
    Else
        Dim blnPageNo As Boolean
        blnPageNo = False
        OptPageNo(0).Enabled = True
        OptPageNo(1).Enabled = True
        OptPageNo(2).Enabled = True
        If OptPageNo(0).Value = True Then blnPageNo = True
        If OptPageNo(1).Value = True Then blnPageNo = True
        If OptPageNo(2).Value = True Then blnPageNo = True
        If Not blnPageNo Then OptPageNo(2).Value = True
    End If
End Sub

Private Sub cmdFontApp_Click()
    setfont txtFontNameApp, txtFontSizeApp, txtFontStyleApp
End Sub

Private Sub cmdFontTitl_Click()
    setfont txtFontNameTitl, txtFontSizeTitl, txtFontStyleTitl
End Sub

Private Sub cmd关闭_Click()
    byRunMode = 0
    Me.Hide
End Sub

Private Sub cmd打印_Click()
        byRunMode = 1
        Me.Hide
End Sub

Private Sub cmd预览_Click()
    byRunMode = 2
    Me.Hide
End Sub

Private Sub cmd设置_Click()
    comDlg.PrinterDefault = True
    comDlg.Min = 1
    comDlg.Max = 9999
    comDlg.FromPage = 1
    comDlg.ToPage = 9999
    comDlg.Flags = cdlPDAllPages _
        + cdlPDCollate _
        + cdlPDNoSelection _
        + cdlPDReturnDC _
        + cdlPDUseDevModeCopies
    comDlg.ShowPrinter
    
    lblPrintStatus.Caption = "打 印 机： 连接到" & Printer.Port _
        & " 的 " & Printer.DeviceName & Chr(13) & Chr(10) _
        & "纸张尺寸： " & PaperName() & Chr(13) & Chr(10) _
        & "进纸方式： " & PaperSource()
    
End Sub

Private Sub Command1_Click()
        byRunMode = 3
        Me.Hide
End Sub

Private Sub Form_Activate()
    lblPrintStatus.Caption = "打 印 机： 连接到" & Printer.Port _
        & " 的 " & Printer.DeviceName & Chr(13) & Chr(10) _
        & "纸张尺寸： " & PaperName() & Chr(13) & Chr(10) _
        & "进纸方式： " & PaperSource()
End Sub


Private Sub setfont(txtFontName As Object, txtFontSize As Object, txtFontStyle As Object)
    comDlg.FontName = txtFontName.Text
    comDlg.FontSize = txtFontSize.Text
    Select Case txtFontStyle.Text
    Case "正常体"
        comDlg.FontBold = False
        comDlg.FontItalic = False
    Case "粗体"
        comDlg.FontBold = True
        comDlg.FontItalic = False
    Case "斜体"
        comDlg.FontBold = False
        comDlg.FontItalic = True
    Case "粗斜体"
        comDlg.FontBold = True
        comDlg.FontItalic = True
    End Select
    comDlg.Flags = cdlCFANSIOnly _
        + cdlCFApply _
        + cdlCFPrinterFonts
    
    comDlg.ShowFont
    txtFontName.Text = comDlg.FontName
    txtFontSize.Text = comDlg.FontSize
    If Not comDlg.FontBold And Not comDlg.FontItalic Then txtFontStyle.Text = "正常体"
    If comDlg.FontBold And Not comDlg.FontItalic Then txtFontStyle.Text = "粗体"
    If Not comDlg.FontBold And comDlg.FontItalic Then txtFontStyle.Text = "斜体"
    If comDlg.FontBold And comDlg.FontItalic Then txtFontStyle.Text = "粗斜体"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Visible = False
    End If
End Sub
