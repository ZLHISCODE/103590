VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCasePrintSet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病历打印选项"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   Icon            =   "frmCasePrintSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印(&P)"
      Height          =   350
      Left            =   5400
      TabIndex        =   9
      Top             =   1380
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5400
      TabIndex        =   10
      Top             =   1845
      Width           =   1100
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "预览(&V)"
      Height          =   350
      Left            =   5400
      TabIndex        =   8
      Top             =   930
      Width           =   1100
   End
   Begin VB.Frame fra打印 
      Caption         =   "打印"
      Height          =   2280
      Left            =   120
      TabIndex        =   12
      Top             =   795
      Width           =   5100
      Begin VB.TextBox TxtEnd 
         Height          =   300
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "1"
         Top             =   1860
         Width           =   360
      End
      Begin VB.TextBox TxtBegin 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "1"
         Top             =   1860
         Width           =   360
      End
      Begin VB.CheckBox ChkPrintPage 
         Caption         =   "打印指定范围"
         Height          =   225
         Left            =   270
         TabIndex        =   20
         Top             =   1560
         Width           =   1635
      End
      Begin VB.PictureBox picBack 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2130
         Left            =   2910
         ScaleHeight     =   491.128
         ScaleMode       =   0  'User
         ScaleWidth      =   491.128
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   120
         Width           =   2130
         Begin VB.PictureBox picPaper 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            ForeColor       =   &H80000008&
            Height          =   1485
            Left            =   405
            ScaleHeight     =   1455
            ScaleMode       =   0  'User
            ScaleWidth      =   1140
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "拖动蓝色线条改变起始位置"
            Top             =   270
            Width           =   1170
            Begin VB.PictureBox pic起始 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   15
               Left            =   0
               MousePointer    =   7  'Size N S
               ScaleHeight     =   15
               ScaleMode       =   0  'User
               ScaleWidth      =   1140
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   135
               Width           =   1140
            End
         End
         Begin VB.PictureBox picShadow 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1485
            Left            =   450
            ScaleHeight     =   1485
            ScaleWidth      =   1170
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   315
            Width           =   1170
         End
      End
      Begin VB.TextBox txt起始 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1050
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "25"
         Top             =   555
         Width           =   600
      End
      Begin VB.CheckBox chk病人 
         Caption         =   "在起始位置打印详细病人信息"
         Height          =   195
         Left            =   285
         TabIndex        =   2
         Top             =   300
         Value           =   1  'Checked
         Width           =   2640
      End
      Begin MSComCtl2.UpDown UD页号 
         Height          =   300
         Left            =   1665
         TabIndex        =   7
         Top             =   1185
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txt页号"
         BuddyDispid     =   196623
         OrigLeft        =   1590
         OrigTop         =   1365
         OrigRight       =   1830
         OrigBottom      =   1665
         Max             =   999
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txt页号 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1050
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "1"
         Top             =   1185
         Width           =   570
      End
      Begin VB.CheckBox chk页号 
         Alignment       =   1  'Right Justify
         Caption         =   "打印页号"
         Height          =   195
         Left            =   225
         TabIndex        =   5
         Top             =   915
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin MSComCtl2.UpDown UD起始 
         Height          =   300
         Left            =   1665
         TabIndex        =   4
         Top             =   555
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   25
         BuddyControl    =   "txt起始"
         BuddyDispid     =   196620
         OrigLeft        =   1590
         OrigTop         =   705
         OrigRight       =   1830
         OrigBottom      =   1005
         Max             =   460
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Left            =   1200
         TabIndex        =   23
         Top             =   1860
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "TxtBegin"
         BuddyDispid     =   196614
         OrigLeft        =   1590
         OrigTop         =   1365
         OrigRight       =   1830
         OrigBottom      =   1665
         Max             =   999
         Min             =   1
         Enabled         =   0   'False
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   300
         Left            =   2550
         TabIndex        =   26
         Top             =   1860
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "TxtEnd"
         BuddyDispid     =   196613
         OrigLeft        =   1590
         OrigTop         =   1365
         OrigRight       =   1830
         OrigBottom      =   1665
         Max             =   999
         Min             =   1
         Enabled         =   0   'False
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "结束页"
         Height          =   180
         Left            =   1590
         TabIndex        =   24
         Top             =   1920
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "开始页"
         Height          =   180
         Left            =   270
         TabIndex        =   21
         Top             =   1920
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         Height          =   180
         Left            =   1965
         TabIndex        =   19
         Top             =   585
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "起始位置"
         Height          =   180
         Left            =   255
         TabIndex        =   14
         Top             =   615
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "起始页号"
         Height          =   180
         Left            =   255
         TabIndex        =   13
         Top             =   1260
         Width           =   720
      End
   End
   Begin VB.Frame fra病历 
      Caption         =   "病历"
      Height          =   645
      Left            =   120
      TabIndex        =   11
      Top             =   75
      Width           =   5100
      Begin VB.OptionButton opt连续 
         Caption         =   "从当前病历开始连续打印"
         Height          =   180
         Left            =   2400
         TabIndex        =   1
         Top             =   285
         Width           =   2280
      End
      Begin VB.OptionButton opt当前 
         Caption         =   "只打印当前选择的病历"
         Height          =   180
         Left            =   225
         TabIndex        =   0
         Top             =   285
         Value           =   -1  'True
         Width           =   2100
      End
   End
End
Attribute VB_Name = "frmCasePrintSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytOpt As Byte

Private mblnFirst As Boolean
Private mblnCurCase As Boolean
Private mblnPatiInfo As Boolean
Private mlngBeginY As Long
Private mintBeginPage As Integer
Private mlng病历记录ID As Long
Private mlngPatientID As Long               '病人ID
Private mlngPageID As Integer               '病人主页ID

Private mlngPrintBeginPage As Long
Private mlngPrintEndPage As Long

Private mlngWidth As Long '自定义纸张宽度,Twip
Private mlngHeight As Long '自定义纸张高度'Twip
Private mlngLeft As Long '左边距'mm
Private mlngRight As Long '右边距'mm
Private mlngTop As Long '上边距'mm
Private mlngBottom As Long '下边距'mm

Private Sub ChkPrintPage_Click()
    If ChkPrintPage.Value = 0 Then
        Me.UpDown1.Enabled = False
        Me.UpDown2.Enabled = False
        mlngPrintBeginPage = 0
        mlngPrintEndPage = 0
        Me.TxtBegin.Locked = True
        Me.TxtEnd.Locked = True
    Else
        Me.UpDown1.Enabled = True
        Me.UpDown2.Enabled = True
        mlngPrintBeginPage = Me.TxtBegin
        mlngPrintEndPage = Me.TxtEnd
        Me.TxtBegin.Locked = False
        Me.TxtEnd.Locked = False
    End If
End Sub

Private Sub chk页号_Click()
    txt页号.Enabled = chk页号.Value = 1
    UD页号.Enabled = chk页号.Value = 1
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    If Not GetValue Then Exit Sub
    mbytOpt = 1
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If Not GetValue Then Exit Sub
    mbytOpt = 2
    Unload Me
End Sub

Private Sub Form_Load()
Dim rsTmp As New ADODB.Recordset
Dim strSQL As String

    mbytOpt = 0
    
    '显示纸张打印位置调整图
    mlngWidth = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "宽度", Printer.Width)
    mlngHeight = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "高度", Printer.Height)
    mlngLeft = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "左边距", OFFSET_LEFT)
    mlngRight = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "右边距", OFFSET_RIGHT)
    mlngTop = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "上边距", OFFSET_TOP)
    mlngBottom = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "下边距", OFFSET_BOTTOM)
    
    If mlngWidth > mlngHeight Then
        picBack.ScaleWidth = mlngWidth / 56.7 * 1.1
        picBack.ScaleHeight = mlngWidth / 56.7 * 1.1
    Else
        picBack.ScaleWidth = mlngHeight / 56.7 * 1.1
        picBack.ScaleHeight = mlngHeight / 56.7 * 1.1
    End If
    picPaper.Width = mlngWidth / 56.7
    picPaper.Height = mlngHeight / 56.7
    picPaper.Left = (picBack.ScaleWidth - picPaper.Width) / 2
    picPaper.Top = (picBack.ScaleHeight - picPaper.Height) / 2
    picShadow.Width = picPaper.Width
    picShadow.Height = picPaper.Height
    picShadow.Left = picPaper.Left + 5
    picShadow.Top = picPaper.Top + 5
    
    picPaper.ScaleWidth = mlngWidth / 56.7
    picPaper.ScaleHeight = mlngHeight / 56.7
    
    '初使化打印位置
    InitPrintPosition
    
    '其它控件值初始
    If mblnCurCase Then
        opt当前.Value = True
    Else
        opt连续.Value = True
    End If
    chk病人.Value = IIf(mblnPatiInfo, 1, 0)
    
    chk页号.Value = IIf(mintBeginPage <> 0, 1, 0)
    UD页号.Value = IIf(mintBeginPage = 0, 1, mintBeginPage)
    
    If Not mblnFirst Then
        opt当前.Enabled = False
        opt连续.Enabled = False
        
        cmdPrint.Visible = False
        cmdCancel.Top = cmdPrint.Top
        cmdPreview.Caption = "确定(&O)"
        cmdPreview.Default = True
    End If
    mlngPrintBeginPage = 0
    mlngPrintEndPage = 0
End Sub

Public Function PrintSet(objParent As Object, ByVal blnFirst As Boolean, ByRef blnCurCase As Boolean, _
    ByRef blnPatiInfo As Boolean, ByRef lngBeginY As Long, ByRef intBeginPage As Integer, Optional ByVal lng病历记录ID As Long = 0, _
    Optional ByRef lng开始页 As Long = 0, Optional ByRef lng结束页 As Long = 0, Optional ByRef lngPatientID As Long, Optional ByRef lngPageID As Long) As Byte
    '功能：调用打印选项
    '参数：blnFirst=是否第一次调用,否则只有"确定","取消",且不允许修改病历打印份数
    '      blnCurCase=T=只打印当前病历,F=从当前病历开始连续打印病历
    '      blnPatiInfo=病历前打印病人信息
    '      lngBeginY=本次病历开始打印位置'mm
    '      intBeginPage=起始页号,为0表示不打印页号
    '      lng病历记录ID = 通过这个ID可以从数据库中读最后一次的这个病历的打印位置,以便在打印时接着上次打印
    '      lngPatientID病人ID
    '      lngPageID病人主页ID
    '返回：0-取消,1-预览,2-打印
    
    mblnFirst = blnFirst
    mblnCurCase = blnCurCase
    mblnPatiInfo = blnPatiInfo
    mlngBeginY = lngBeginY
    mintBeginPage = intBeginPage
    mlngPrintBeginPage = lng开始页
    mlngPrintEndPage = lng结束页
    mlng病历记录ID = lng病历记录ID
    mlngPatientID = lngPatientID
    mlngPageID = lngPageID
    Me.Show 1, objParent
    
    blnCurCase = mblnCurCase
    blnPatiInfo = mblnPatiInfo
    lngBeginY = mlngBeginY
    intBeginPage = mintBeginPage
    lng开始页 = mlngPrintBeginPage
    lng结束页 = mlngPrintEndPage
    PrintSet = mbytOpt
End Function

Private Sub opt当前_Click()
    InitPrintPosition
End Sub

Private Sub opt连续_Click()
    InitPrintPosition
End Sub

Private Sub pic起始_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If pic起始.Top + y > UD起始.Max Or pic起始.Top + y < UD起始.Min Then Exit Sub
        pic起始.Top = pic起始.Top + y
        UD起始.Value = pic起始.Top
        Call DrawPage
        Me.Refresh
    End If
End Sub

Private Sub TxtBegin_GotFocus()
    Me.TxtBegin.SelStart = 0
    Me.TxtBegin.SelLength = Len(Me.TxtBegin)
End Sub

Private Sub TxtBegin_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub TxtEnd_GotFocus()
    Me.TxtEnd.SelStart = 0
    Me.TxtEnd.SelLength = Len(Me.TxtEnd.Text)
End Sub

Private Sub TxtEnd_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt起始_Change()
    If Val(txt起始.Text) >= UD起始.Min And Val(txt起始.Text) <= UD起始.Max Then
        UD起始.Value = Val(txt起始.Text)
    End If
End Sub

Private Sub txt起始_GotFocus()
    zlControl.TxtSelAll txt起始
End Sub

Private Sub txt起始_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt页号_GotFocus()
    zlControl.TxtSelAll txt页号
End Sub

Private Sub txt页号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Function GetValue() As Boolean
    If Not (Val(txt起始.Text) >= UD起始.Min And Val(txt起始.Text) <= UD起始.Max) Then
        MsgBox "起始位置应该在 " & UD起始.Min & " 至 " & UD起始.Max & " 之间！", vbInformation, gstrSysName
        txt起始.SetFocus: Exit Function
    End If
    If Len(Trim(TxtBegin.Text)) < 1 Then
        TxtBegin.Text = 1
    End If
    If Len(Trim(TxtEnd.Text)) < 1 Then
        TxtEnd.Text = 1
    End If
    mlngPrintBeginPage = Me.TxtBegin
    mlngPrintEndPage = Me.TxtEnd
    If Val(TxtBegin.Text) > Val(TxtEnd.Text) Then
        MsgBox "开始页应该在1--" & Val(TxtEnd.Text) & "之间！", vbInformation, gstrSysName
        TxtBegin.SetFocus: Exit Function
    End If
    mblnCurCase = opt当前.Value
    mblnPatiInfo = chk病人.Value = 1
    mlngBeginY = Val(txt起始.Text)
    If chk页号.Value = 1 Then
        mintBeginPage = Val(txt页号.Text)
    Else
        mintBeginPage = 0
    End If
    
    GetValue = True
End Function

Private Sub UD起始_Change()
    pic起始.Top = UD起始.Value
    Call DrawPage
End Sub

Private Sub DrawPage()
    picPaper.Cls
    picPaper.Line (0, mlngTop)-(picPaper.ScaleWidth, mlngTop), &H808080
    picPaper.Line (0, picPaper.ScaleHeight - mlngBottom)-(picPaper.ScaleWidth, picPaper.ScaleHeight - mlngBottom), &H808080
    picPaper.Line (mlngLeft, 0)-(mlngLeft, picPaper.ScaleHeight), &H808080
    picPaper.Line (picPaper.ScaleWidth - mlngRight, 0)-(picPaper.ScaleWidth - mlngRight, picPaper.ScaleHeight), &H808080
    
    picPaper.Line (mlngLeft, UD起始.Value)-(picPaper.ScaleWidth - mlngRight, picPaper.ScaleHeight - mlngBottom), &H808080, B
End Sub

Private Sub UpDown1_DownClick()
    If Val(Me.TxtBegin) > 1 Then
        Me.TxtBegin = Val(Me.TxtBegin) - 1
    End If
    mlngPrintBeginPage = Me.TxtBegin
End Sub

Private Sub UpDown1_UpClick()
    Me.TxtBegin = Val(Me.TxtBegin) + 1
    mlngPrintBeginPage = Me.TxtBegin
End Sub
Private Sub UpDown2_DownClick()
    Me.TxtEnd = Val(Me.TxtEnd) - 1
    mlngPrintEndPage = Me.TxtEnd
End Sub

Private Sub UpDown2_UpClick()
    Me.TxtEnd = Val(Me.TxtEnd) + 1
    mlngPrintEndPage = Me.TxtEnd
End Sub
Sub InitPrintPosition()
    '功能:          初使化打印位置
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    '重新从数据中读出最后一次的页码与Y坐标位置
    mlngBeginY = 0
    If mlng病历记录ID > 0 Then
        strSQL = _
            "SELECT nvl(起始页号,1) 页号, nvl(起始位置,0) Y" & vbCrLf & _
            "  FROM 病历打印记录" & vbCrLf & _
            " WHERE 病历记录ID = " & mlng病历记录ID
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        If rsTmp.RecordCount > 0 Then
            mintBeginPage = rsTmp!页号
            mlngBeginY = Me.picPaper.ScaleY(rsTmp!y, vbTwips, vbMillimeters)
        Else
            If opt连续.Value = True Then
                strSQL = "select  nvl(起始页号,1) 页号, nvl(结束位置,0) Y from 病人病历记录 a , 病历打印记录 b where " & _
                " a.id = b.病历记录ID " & " and a.病人id = " & mlngPatientID & " and a.主页ID = " & mlngPageID & " order by  打印时间 desc,结束页号 desc "
                Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
                If rsTmp.EOF <> True Then
                    mintBeginPage = rsTmp!页号
                    mlngBeginY = Me.picPaper.ScaleY(rsTmp!y, vbTwips, vbMillimeters)
                End If
            End If
        End If
    End If
    
    '显初始位置
    If Not (mlngBeginY >= mlngTop And mlngBeginY <= picPaper.ScaleHeight - mlngBottom * 2) Then
        mlngBeginY = mlngTop
    End If
    pic起始.Left = 0
    pic起始.Width = picPaper.ScaleWidth
    pic起始.Top = mlngBeginY
    
    UD起始.Min = mlngTop
    UD起始.Max = picPaper.ScaleHeight - 2 * mlngBottom
    UD起始.Value = mlngBeginY
     
    pic起始.ScaleHeight = 1 '不然不能拖动
    txt起始.Text = mlngBeginY
    Call DrawPage
End Sub
