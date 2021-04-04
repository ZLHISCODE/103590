VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPageSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "报表页面设置"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   Icon            =   "frmPageSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame4 
      Caption         =   "送纸器"
      Height          =   765
      Left            =   75
      TabIndex        =   21
      Top             =   2280
      Width           =   3855
      Begin VB.ComboBox cboBin 
         Height          =   300
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   285
         Width           =   2505
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "纸张来源"
         Height          =   180
         Left            =   210
         TabIndex        =   22
         Top             =   345
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4260
      TabIndex        =   8
      Top             =   2715
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4260
      TabIndex        =   7
      Top             =   2295
      Width           =   1100
   End
   Begin VB.Frame Frame3 
      Caption         =   "当前格式纸向"
      Height          =   1065
      Left            =   4035
      TabIndex        =   18
      Top             =   1125
      Width           =   1605
      Begin VB.OptionButton opt横向 
         Caption         =   "横向"
         Height          =   285
         Left            =   780
         TabIndex        =   5
         Top             =   660
         Width           =   660
      End
      Begin VB.OptionButton opt纵向 
         Caption         =   "纵向"
         Height          =   285
         Left            =   780
         TabIndex        =   4
         Top             =   285
         Value           =   -1  'True
         Width           =   660
      End
      Begin VB.Image img横向 
         Height          =   480
         Left            =   180
         Picture         =   "frmPageSetup.frx":014A
         Top             =   360
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image img纵向 
         Height          =   480
         Left            =   180
         Picture         =   "frmPageSetup.frx":0A14
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "当前格式纸张"
      Height          =   1065
      Left            =   75
      TabIndex        =   12
      Top             =   1125
      Width           =   3855
      Begin MSComCtl2.UpDown UDHeight 
         Height          =   285
         Left            =   3105
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   630
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtHeight"
         BuddyDispid     =   196622
         OrigLeft        =   2985
         OrigTop         =   630
         OrigRight       =   3225
         OrigBottom      =   930
         Max             =   765
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UDWidth 
         Height          =   285
         Left            =   1410
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   630
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtWidth"
         BuddyDispid     =   196623
         OrigLeft        =   1200
         OrigTop         =   645
         OrigRight       =   1440
         OrigBottom      =   945
         Max             =   765
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtHeight 
         Height          =   300
         Left            =   2415
         MaxLength       =   6
         TabIndex        =   3
         Top             =   630
         Width           =   690
      End
      Begin VB.TextBox txtWidth 
         Height          =   300
         Left            =   720
         MaxLength       =   6
         TabIndex        =   2
         Top             =   630
         Width           =   690
      End
      Begin VB.ComboBox cboPage 
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2955
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         Height          =   180
         Left            =   3420
         TabIndex        =   20
         Top             =   735
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         Height          =   180
         Left            =   1710
         TabIndex        =   19
         Top             =   735
         Width           =   180
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "高度"
         Height          =   180
         Left            =   2010
         TabIndex        =   15
         Top             =   690
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "宽度"
         Height          =   180
         Left            =   300
         TabIndex        =   14
         Top             =   690
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "大小"
         Height          =   180
         Left            =   285
         TabIndex        =   13
         Top             =   300
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "打印机"
      Height          =   1005
      Left            =   75
      TabIndex        =   9
      Top             =   60
      Width           =   5565
      Begin VB.ComboBox cboPrinter 
         Height          =   300
         Left            =   1635
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   225
         Width           =   3540
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   390
         Picture         =   "frmPageSetup.frx":12DE
         Top             =   330
         Width           =   480
      End
      Begin VB.Label lblLoc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "位置"
         Height          =   180
         Left            =   1185
         TabIndex        =   11
         Top             =   660
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "名称"
         Height          =   180
         Left            =   1185
         TabIndex        =   10
         Top             =   285
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmPageSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'入/出口参数
Public strPrinter As String '打印机
Public intPage As Integer '纸张
Public lngWidth As Long, lngHeight As Long '尺寸(按纵向的尺寸)
Public bytOrient As Byte, intBin As Integer  '纸向/进纸方式

'事件屏蔽标志
Private blnPrinter As Boolean
Private blnPage As Boolean
Private blnChange As Boolean

Private Sub cboPage_Click()
    '自定义纸张不支持横向打印
'    If cboPage.ListIndex <> -1 Then
'        If cboPage.ItemData(cboPage.ListIndex) = 256 Then
'            opt纵向.Value = True
'            opt纵向.Enabled = False
'            opt横向.Enabled = False
'
'            Call opt纵向_Click
'        Else
'            opt纵向.Enabled = True
'            opt横向.Enabled = True
'        End If
'    End If
    
    If Not blnPage Then Exit Sub
        
    blnChange = False
    If cboPage.ItemData(cboPage.ListIndex) <> 256 Then
        '缺省可能为横向,强行设置为纵向,以正确取要显示的宽高
        Printer.Orientation = 1
        '使用该打印机支持该幅面的真实尺寸
        Printer.PaperSize = cboPage.ItemData(cboPage.ListIndex)
        txtWidth.Tag = Printer.Width
        txtWidth.Text = FormatEx(Printer.Width / Twip_mm, 2)
        txtHeight.Tag = Printer.Height
        txtHeight.Text = FormatEx(Printer.Height / Twip_mm, 2)
    Else
        If cboPage.Text = PageCustom1 Then
            txtWidth.Text = 241: txtWidth.Tag = CInt(241 * Twip_mm)
            txtHeight.Text = 280: txtHeight.Tag = CInt(280 * Twip_mm)
        ElseIf cboPage.Text = PageCustom2 Then
            txtWidth.Text = 241: txtWidth.Tag = CInt(241 * Twip_mm)
            txtHeight.Text = 140: txtHeight.Tag = CInt(140 * Twip_mm)
        ElseIf cboPage.Text = PageCustom3 Then
            txtWidth.Text = 241: txtWidth.Tag = CInt(241 * Twip_mm)
            txtHeight.Text = 94: txtHeight.Tag = CInt(94 * Twip_mm)
        ElseIf txtWidth.Text = "" And txtHeight.Text = "" Then
            txtWidth.Text = FormatEx(INIT_WIDTH / Twip_mm, 2)
            txtWidth.Tag = INIT_WIDTH
            txtHeight.Text = FormatEx(INIT_HEIGHT / Twip_mm, 2)
            txtHeight.Tag = INIT_HEIGHT
        End If
    End If
    blnChange = True
End Sub

Private Sub cboPrinter_Click()
    Dim i As Integer, j As Integer, k As Integer
    Dim lngCount As Long, strTmp As String
    Dim strPaperSize As String * 300
    Dim strPaperBin As String * 100
    Dim strPaperBinName As String * 1000
    
    If Not blnPrinter Then Exit Sub
    
    Set Printer = Printers(cboPrinter.ItemData(cboPrinter.ListIndex))
    lblLoc.Caption = "位置: " & Printer.Port
    
    '设置可用纸张
    '返回格式见常量说明或MSDN
    cboPage.Clear
    
    '------------------------------------------------------------------------------------------
    '纸张大小
    lngCount = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_PAPERS, strPaperSize, 0)
    For i = 1 To lngCount
        j = Asc(Mid(strPaperSize, i * 2, 1)) * 256# + Asc(Mid(strPaperSize, i * 2 - 1, 1))
        If j >= 1 And j <= 41 Then '只列出标准支持的纸张
            cboPage.AddItem GetPaperName(j)
            cboPage.ItemData(cboPage.ListCount - 1) = j
            If j = intPage Then cboPage.ListIndex = cboPage.ListCount - 1 '定位在原设置上
            If cboPage.ListIndex = -1 And j = Printer.PaperSize Then
                cboPage.ListIndex = cboPage.ListCount - 1 '定位在打印机缺省设置上
            End If
        End If
    Next
    
    '------------------------------------------------------------------------------------------
    '自定义不管是否支持,都要用
    cboPage.AddItem PageCustom1: cboPage.ItemData(cboPage.NewIndex) = 256
    cboPage.AddItem PageCustom2: cboPage.ItemData(cboPage.NewIndex) = 256
    cboPage.AddItem PageCustom3: cboPage.ItemData(cboPage.NewIndex) = 256
    cboPage.AddItem GetPaperName(256): cboPage.ItemData(cboPage.ListCount - 1) = 256
    
    '不支持A4则用自定义
    If cboPage.ListIndex = -1 Then cboPage.ListIndex = cboPage.ListCount - 1
    
    '设置可用进纸方式
    cboBin.Clear
    '--------------------------------------------------------------------------------------------
    lngCount = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_BINNAMES, strPaperBinName, 0)
    lngCount = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_BINS, strPaperBin, 0)
    j = 1
    For i = 1 To lngCount
        k = 0
        '进纸名称
        Do
            If Mid(strPaperBinName, j, 1) = Chr(0) Then
                If Trim(strTmp) <> "" Then
                    cboBin.AddItem Trim(strTmp)
                    
                    '进纸编号
                    cboBin.ItemData(cboBin.ListCount - 1) = Asc(Mid(strPaperBin, i * 2, 1)) * 256# + Asc(Mid(strPaperBin, i * 2 - 1, 1))
                    If cboBin.ItemData(cboBin.ListCount - 1) = intBin Then
                        cboBin.ListIndex = cboBin.ListCount - 1 '定位在原设置上
                    End If
                    If cboBin.ListIndex = -1 And cboBin.ItemData(cboBin.ListCount - 1) = Printer.PaperBin Then
                        cboBin.ListIndex = cboBin.ListCount - 1 '定位在打印机缺省设置上
                    End If
                End If
                
                j = 24 + j - LenB(StrConv(strTmp, vbFromUnicode))
                strTmp = ""
                Exit Do
            Else
                strTmp = strTmp & Mid(strPaperBinName, j, 1)
                j = j + 1
                k = k + 1
                If k > 24 Then Exit Do
            End If
        Loop
    Next
    '--------------------------------------------------------------------------------------------
    If cboBin.ListIndex = -1 And cboBin.ListCount > 0 Then cboBin.ListIndex = 0
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Not IsNumeric(txtWidth.Text) Then
        MsgBox "请确定报表的纸张宽度。", vbInformation, App.Title
        txtWidth.SetFocus: Exit Sub
    End If
    If CLng(Val(txtWidth.Text)) > 765 Or CLng(Val(txtWidth.Text)) < 5 Then
        MsgBox "报表的纸张宽度必须在5-765毫米之间。", vbInformation, App.Title
        txtWidth.SetFocus: Exit Sub
    End If
    
    If Not IsNumeric(txtHeight.Text) Then
        MsgBox "请确定报表的纸张高度。", vbInformation, App.Title
        txtHeight.SetFocus: Exit Sub
    End If
    If CLng(Val(txtHeight.Text)) > 765 Or CLng(Val(txtHeight.Text)) < 5 Then
        MsgBox "报表的纸张高度必须在5-765毫米之间。", vbInformation, App.Title
        txtHeight.SetFocus: Exit Sub
    End If
    
    On Error Resume Next
    strPrinter = cboPrinter.Text
    bytOrient = IIF(opt纵向.Value, 1, 2)
    intPage = cboPage.ItemData(cboPage.ListIndex)
    lngWidth = CLng(txtWidth.Tag)
    lngHeight = CLng(txtHeight.Tag)
    
    If cboBin.ListIndex <> -1 Then
        intBin = cboBin.ItemData(cboBin.ListIndex)
    Else
        intBin = 15
    End If
    
    gblnOK = True
    Hide
End Sub

Private Sub Form_Load()
    Dim strPaper As String, i As Integer
    
    On Error Resume Next
    
    gblnOK = False
    
    blnPrinter = False
    blnPage = False
    blnChange = False
    
    '初始打印机列表
    With cboPrinter
        .Clear
        For i = 0 To Printers.Count - 1
            .AddItem Printers(i).DeviceName
            .ItemData(.ListCount - 1) = i '打印机索引
            
            '读取存储的打印机为当前打印机,并初始化可用页面
            If strPrinter = Printers(i).DeviceName Then blnPrinter = True: .ListIndex = i: blnPrinter = False
        Next
        
        '缺省初始化为当前打印机
        If .ListIndex = -1 Then
            For i = 0 To .ListCount - 1
                '读取系统当前的打印机为当前打印机,并初始化可用页面
                If .List(i) = Printer.DeviceName Then blnPrinter = True: .ListIndex = i: blnPrinter = False: Exit For
            Next
        End If
    End With
    
    cboPage.ListIndex = -1
    
    '初始化打印机时可用打印纸张已加入
    Select Case intPage
        Case 256 '自定义纸张
            strPaper = GetPaperName(256, lngWidth, lngHeight)
            If Not strPaper Like "用户自定义*" Then
                cboPage.ListIndex = GetCboIndex(cboPage, strPaper)
                txtWidth.Text = CInt(lngWidth / Twip_mm): txtWidth.Tag = lngWidth
                txtHeight.Text = CInt(lngHeight / Twip_mm): txtHeight.Tag = lngHeight
            ElseIf txtWidth.Text = "" And txtHeight.Text = "" Then
                '设置为自定义,并读取页面大小
                cboPage.ListIndex = cboPage.ListCount - 1
                txtWidth.Text = FormatEx(lngWidth / Twip_mm, 2)
                txtHeight.Text = FormatEx(lngHeight / Twip_mm, 2)
                txtWidth.Tag = lngWidth
                txtHeight.Tag = lngHeight
            End If
        Case Else '系统纸张
            Printer.PaperSize = intPage
            If Err.Number <> 0 Then
                '该打印机不支持存储的纸张(打印机已改变),则设为自定义
                cboPage.ListIndex = cboPage.ListCount - 1
                '非自定义一样要存放宽高
                txtWidth.Text = FormatEx(lngWidth / Twip_mm, 2)
                txtHeight.Height = FormatEx(lngHeight / Twip_mm, 2)
                txtWidth.Tag = lngWidth
                txtHeight.Tag = lngHeight
                Err.Clear
            Else
                For i = 0 To cboPage.ListCount - 1
                    If cboPage.ItemData(i) = intPage Then blnPage = True: cboPage.ListIndex = i: blnPage = False: Exit For
                Next
            End If
    End Select
    
    If bytOrient = 2 Then opt横向.Value = True: opt横向_Click
    
    blnPrinter = True
    blnPage = True
    blnChange = True
End Sub

Private Sub opt横向_Click()
    If opt横向.Value Then
        img纵向.Visible = False
        img横向.Visible = True
    End If
End Sub

Private Sub opt纵向_Click()
    If opt纵向.Value Then
        img纵向.Visible = True
        img横向.Visible = False
    End If
End Sub

Private Sub txtHeight_Change()
    If Not blnChange Then Exit Sub
    
    blnPage = False
    cboPage.ListIndex = cboPage.ListCount - 1
    If IsNumeric(txtHeight.Text) Then txtHeight.Tag = CLng(txtHeight.Text * Twip_mm)
    blnPage = True
End Sub

Private Sub txtHeight_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: VBA.Beep
End Sub

Private Sub txtWidth_Change()
    If Not blnChange Then Exit Sub
    
    blnPage = False
    cboPage.ListIndex = cboPage.ListCount - 1
    If IsNumeric(txtWidth.Text) Then txtWidth.Tag = CLng(txtWidth.Text * Twip_mm)
    blnPage = True
End Sub

Private Sub txtheight_GotFocus()
    txtHeight.SelStart = 0: txtHeight.SelLength = Len(txtHeight.Text)
End Sub

Private Sub txtwidth_GotFocus()
    txtWidth.SelStart = 0: txtWidth.SelLength = Len(txtWidth.Text)
End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: VBA.Beep
End Sub
