VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrintSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "打印机设置"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5985
      TabIndex        =   1
      Top             =   3240
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4725
      TabIndex        =   0
      Top             =   3240
      Width           =   1100
   End
   Begin VB.Frame fraOrientation 
      Caption         =   "方向"
      Height          =   1725
      Left            =   4890
      TabIndex        =   4
      Top             =   1335
      Width           =   2175
      Begin VB.OptionButton optLandscape 
         Caption         =   "横向(&A)"
         Height          =   285
         Left            =   1095
         TabIndex        =   17
         Top             =   1020
         Width           =   960
      End
      Begin VB.OptionButton optPortrait 
         Caption         =   "纵向(&O)"
         Height          =   285
         Left            =   1095
         TabIndex        =   16
         Top             =   555
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   375
         TabIndex        =   18
         Top             =   705
         Width           =   255
      End
      Begin VB.Shape shpPortrait 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000006&
         Height          =   645
         Left            =   270
         Top             =   600
         Width           =   480
      End
      Begin VB.Shape shpLandscape 
         BackStyle       =   1  'Opaque
         Height          =   480
         Left            =   180
         Top             =   690
         Visible         =   0   'False
         Width           =   645
      End
   End
   Begin VB.Frame fraPaper 
      Caption         =   "纸张"
      Height          =   1725
      Left            =   135
      TabIndex        =   3
      Top             =   1335
      Width           =   4680
      Begin VB.TextBox txtHeight 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         MaxLength       =   3
         TabIndex        =   20
         Text            =   "20"
         Top             =   750
         Width           =   465
      End
      Begin VB.TextBox txtWidth 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1590
         MaxLength       =   3
         TabIndex        =   19
         Text            =   "20"
         Top             =   750
         Width           =   435
      End
      Begin MSComCtl2.UpDown updnHeight 
         Height          =   285
         Left            =   3945
         TabIndex        =   13
         Top             =   750
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   20
         BuddyControl    =   "txtHeight"
         BuddyDispid     =   196618
         OrigLeft        =   3930
         OrigTop         =   750
         OrigRight       =   4170
         OrigBottom      =   1050
         Max             =   500
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updnWidth 
         Height          =   285
         Left            =   2025
         TabIndex        =   11
         Top             =   750
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   20
         BuddyControl    =   "txtWidth"
         BuddyDispid     =   196619
         OrigLeft        =   2040
         OrigTop         =   750
         OrigRight       =   2280
         OrigBottom      =   1050
         Max             =   500
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.ComboBox cboPaperBin 
         Height          =   300
         Left            =   855
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1200
         Width           =   3690
      End
      Begin VB.ComboBox cboPaperSize 
         Height          =   300
         Left            =   855
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   330
         Width           =   3690
      End
      Begin VB.Label lblHeightNote 
         AutoSize        =   -1  'True
         Caption         =   "高度(mm):"
         Height          =   180
         Left            =   2610
         TabIndex        =   12
         Top             =   810
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lblWidthNote 
         AutoSize        =   -1  'True
         Caption         =   "宽度(mm):"
         Height          =   180
         Left            =   735
         TabIndex        =   10
         Top             =   810
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lblPaperBin 
         AutoSize        =   -1  'True
         Caption         =   "来源(&U):"
         Height          =   180
         Left            =   105
         TabIndex        =   14
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label lblPaperSize 
         AutoSize        =   -1  'True
         Caption         =   "大小(&S):"
         Height          =   180
         Left            =   105
         TabIndex        =   8
         Top             =   390
         Width           =   720
      End
   End
   Begin VB.Frame fraPrinter 
      Caption         =   "打印机"
      Height          =   1155
      Left            =   135
      TabIndex        =   2
      Top             =   90
      Width           =   6930
      Begin VB.ComboBox cboPrinterName 
         Height          =   300
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   285
         Width           =   4635
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   345
         Picture         =   "frmPrintSet.frx":0000
         Top             =   360
         Width           =   240
      End
      Begin VB.Label lblPort 
         AutoSize        =   -1  'True
         Caption         =   "位置:"
         Height          =   180
         Left            =   1095
         TabIndex        =   7
         Top             =   750
         Width           =   450
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "名称(&N):"
         Height          =   180
         Left            =   1065
         TabIndex        =   5
         Top             =   345
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmPrintSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByVal lpOutput As String, ByVal lpDevMode As Long) As Long
Private Const DC_PAPERNAMES = 16
Private Const DC_PAPERSIZE = 3
Private Const DC_PAPERS = 2
Private Const DC_BINS = 6

Dim DefDeviceName As String
Dim DefPaperSize As Integer
Dim DefWidth As Long, DefHeight As Long
Dim DefPaperBin As Integer
Dim DefOrientation As Integer
Dim colPrinters As New zlPrinters
Dim objPrinter As zlPrinter

Private Sub cboPaperSize_Click()
    Dim blnVisible As Boolean
    Err = 0
    On Error Resume Next
    If cboPaperSize.ItemData(cboPaperSize.ListIndex) = 256 Then
        blnVisible = True
    Else
        blnVisible = False
    End If
    lblWidthNote.Visible = blnVisible
    txtWidth.Visible = blnVisible
    updnWidth.Visible = blnVisible
    
    lblHeightNote.Visible = blnVisible
    txtHeight.Visible = blnVisible
    updnHeight.Visible = blnVisible
    
    updnWidth.Value = DefWidth / conRatemmToTwip
    updnHeight.Value = DefHeight / conRatemmToTwip
    txtWidth.Text = updnWidth.Value
    txtHeight.Text = updnHeight.Value

End Sub

Private Sub cboPrinterName_Click()
    Dim iCount As Integer
    For iCount = 1 To colPrinters.Count
        If colPrinters(iCount).DeviceName = Trim(cboPrinterName.Text) Then
            Set objPrinter = colPrinters(iCount)
        End If
    Next
    PrinterChanged objPrinter

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim iCount As Integer
    Err = 0
    On Error Resume Next
    
    For iCount = 0 To Printers.Count - 1
        If Printers(iCount).DeviceName = objPrinter.DeviceName Then
            If Printer.DeviceName <> objPrinter.DeviceName Then
                Set Printer = Printers(iCount)
                Exit For
            End If
        End If
    Next
    Printer.PaperBin = Me.cboPaperBin.ItemData(Me.cboPaperBin.ListIndex)
    Printer.Orientation = IIf(Me.optPortrait.Value, 1, 2)
    If Me.cboPaperSize.ItemData(Me.cboPaperSize.ListIndex) = 256 Then
        Call SetCustonPager(updnWidth.Value * conRatemmToTwip, updnHeight.Value * conRatemmToTwip)
    Else
        Printer.PaperSize = Me.cboPaperSize.ItemData(Me.cboPaperSize.ListIndex)
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Dim tmpPaperSize As Integer
    Dim tmpPaperBin As Integer
    Dim tmpOrientation As Integer
    Dim iCount As Integer
    Dim iStep As Integer
    Dim lngCount As Long
    Dim strPaper As String * 1000
    Dim i As Long, j As Long
    
    DefDeviceName = Printer.DeviceName
    DefPaperSize = Printer.PaperSize
    DefWidth = Printer.Width
    DefHeight = Printer.Height
    
    DefPaperBin = Printer.PaperBin
    DefOrientation = Printer.Orientation
    '----------------------------------------------
    
    Err = 0
    Set colPrinters = New zlPrinters

    On Error Resume Next
    For iCount = 0 To Printers.Count - 1
        '--------------------------------------------
        Set Printer = Printers(iCount)
        If Printers(iCount).DeviceName = DefDeviceName Then
            Printer.PaperBin = DefPaperBin
            Printer.Orientation = DefOrientation
            If DefPaperSize = 256 Then
                Call SetCustonPager(DefWidth, DefHeight)
            Else
                Printer.PaperSize = DefPaperSize
            End If
        End If
        
        tmpPaperSize = Printer.PaperSize
        tmpPaperBin = Printer.PaperBin
        tmpOrientation = Printer.Orientation
        
        Set objPrinter = New zlPrinter
        objPrinter.DeviceName = Printer.DeviceName
        objPrinter.Port = Printer.Port
        objPrinter.Current = (Printer.DeviceName = DefDeviceName)
        '纸张大小
        lngCount = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_PAPERS, strPaper, 0)
        For i = 1 To lngCount
            j = Asc(Mid(strPaper, i * 2, 1)) * 256# + Asc(Mid(strPaper, i * 2 - 1, 1))
            If j >= 1 And j <= 41 Then '只列出标准支持的纸张
                If j = tmpPaperSize Then
                    objPrinter.PaperSizes = objPrinter.PaperSizes & "," & j & "*" '原有的
                Else
                    objPrinter.PaperSizes = objPrinter.PaperSizes & "," & j
                End If
            End If
        Next
        objPrinter.PaperSizes = objPrinter.PaperSizes & ",256" & IIf(tmpPaperSize = 256, "*", "")
        
        '--------------------------------------------
        lngCount = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_BINS, strPaper, 0)
        For i = 1 To lngCount
            j = Asc(Mid(strPaper, i * 2, 1)) * 256# + Asc(Mid(strPaper, i * 2 - 1, 1))
            If j >= 1 And j <= 11 Then '只列出标准支持的进纸大小
                If j = tmpPaperBin Then
                    objPrinter.PaperBins = objPrinter.PaperBins & "," & j & "*" '原有的
                Else
                    objPrinter.PaperBins = objPrinter.PaperBins & "," & j
                End If
            End If
        Next
        Err = 0
        Printer.PaperBin = 14
        If Printer.PaperBin = 14 Then
            objPrinter.PaperBins = objPrinter.PaperBins & ",14" _
                & IIf(tmpPaperBin = 14, "*", "")
        End If
        
        '--------------------------------------------
        Err = 0
        Printer.Orientation = 1
        If Printer.Orientation = 1 Then
            objPrinter.Orientations = objPrinter.Orientations & "," & IIf(tmpOrientation = 1, "1*", "1")
        End If
        Err = 0
        Printer.Orientation = 2
        If Printer.Orientation = 2 Then
            objPrinter.Orientations = objPrinter.Orientations & "," & IIf(tmpOrientation = 2, "2*", "2")
        End If
        
        '--------------------------------------------
        objPrinter.PaperSizes = Mid(objPrinter.PaperSizes, 2)
        objPrinter.PaperBins = Mid(objPrinter.PaperBins, 2)
        objPrinter.Orientations = Mid(objPrinter.Orientations, 2)
        colPrinters.Add objPrinter
    Next
    
    For iCount = 0 To Printers.Count - 1
        If Printers(iCount).DeviceName = DefDeviceName Then
            Set Printer = Printers(iCount)
            Printer.PaperBin = DefPaperBin
            Printer.Orientation = DefOrientation
            If DefPaperSize = 256 Then
                Call SetCustonPager(DefWidth, DefHeight)
            Else
                Printer.PaperSize = DefPaperSize
            End If
            Exit For
        End If
    Next
    
End Sub

Private Sub Form_Activate()
    Dim iCount As Integer
    
    cboPrinterName.Clear
    For iCount = 1 To colPrinters.Count
        cboPrinterName.AddItem colPrinters(iCount).DeviceName
        If colPrinters(iCount).DeviceName = DefDeviceName Then
            cboPrinterName.ListIndex = cboPrinterName.NewIndex
            Set objPrinter = colPrinters(iCount)
        End If
    Next
End Sub

Private Sub PrinterChanged(objPrinter As zlPrinter)
    
    Dim strCount As String, strTemp As String
    lblPort.Caption = "位置:    连接到" & objPrinter.Port
    
    '--------------------------------------------
    '纸张尺寸
    With cboPaperSize
        .Clear
        strTemp = objPrinter.PaperSizes
        Do While InStr(1, strTemp, ",") > 0
            strCount = Left(strTemp, InStr(1, strTemp, ",") - 1)
            If Right(strCount, 1) = "*" Then
                .AddItem getPapersize(CInt(Left(strCount, Len(strCount) - 1)))
                .ItemData(.NewIndex) = CInt(Left(strCount, Len(strCount) - 1))
                .ListIndex = .NewIndex
            Else
                .AddItem getPapersize(CInt(strCount))
                .ItemData(.NewIndex) = CInt(strCount)
            End If
            strTemp = Mid(strTemp, InStr(1, strTemp, ",") + 1)
        Loop
        strCount = strTemp
        If Right(strCount, 1) = "*" Then
            .AddItem getPapersize(CInt(Left(strCount, Len(strCount) - 1)))
            .ItemData(.NewIndex) = CInt(Left(strCount, Len(strCount) - 1))
            .ListIndex = .NewIndex
        Else
            If IsNumeric(strCount) Then
                .AddItem getPapersize(CInt(strCount))
                .ItemData(.NewIndex) = CInt(strCount)
            End If
        End If

    End With
    
    '--------------------------------------------
    '纸张来源
    With cboPaperBin
        .Clear
        strTemp = objPrinter.PaperBins
        Do While InStr(1, strTemp, ",") > 0
            strCount = Left(strTemp, InStr(1, strTemp, ",") - 1)
            If Right(strCount, 1) = "*" Then
                .AddItem getPaperBin(CInt(Left(strCount, Len(strCount) - 1)))
                .ItemData(.NewIndex) = CInt(Left(strCount, Len(strCount) - 1))
                .ListIndex = .NewIndex
            Else
                .AddItem getPaperBin(CInt(strCount))
                .ItemData(.NewIndex) = CInt(strCount)
            End If
            strTemp = Mid(strTemp, InStr(1, strTemp, ",") + 1)
        Loop
        strCount = strTemp
        If Right(strCount, 1) = "*" Then
            .AddItem getPaperBin(CInt(Left(strCount, Len(strCount) - 1)))
            .ItemData(.NewIndex) = CInt(Left(strCount, Len(strCount) - 1))
            .ListIndex = .NewIndex
        Else
            If IsNumeric(strCount) Then
                .AddItem getPaperBin(CInt(strCount))
                .ItemData(.NewIndex) = CInt(strCount)
            End If
        End If

    End With
    '--------------------------------------------
    
    If InStr(1, objPrinter.Orientations, "1") = 0 Then
        shpPortrait.Visible = False
        optPortrait.Value = False
        optPortrait.Enabled = False
        optLandscape.Enabled = False
    Else
        If InStr(1, objPrinter.Orientations, "1*") <> 0 Then
            optPortrait.Value = True
        End If
    End If
    
    If InStr(1, objPrinter.Orientations, "2") = 0 Then
        shpLandscape.Visible = False
        optLandscape.Value = False
        optLandscape.Enabled = False
        optPortrait.Enabled = False
    Else
        If InStr(1, objPrinter.Orientations, "2*") <> 0 Then
            optLandscape.Value = True
        End If
    End If

End Sub

Private Sub OptLandscape_Click()
    shpPortrait.Visible = optPortrait.Value
    shpLandscape.Visible = Not optPortrait.Value
End Sub

Private Sub optPortrait_Click()
    shpPortrait.Visible = optPortrait.Value
    shpLandscape.Visible = Not optPortrait.Value
End Sub

Public Function getPapersize(mSize As Integer) As String
    '------------------------------------------------
    '功能： 根据当前打印机的设置，获取纸张名称
    '返回： 纸张名称
    '------------------------------------------------
    Err = 0
    On Error GoTo ErrHand
    If mSize = 256 Then
        getPapersize = "用户自定义"
        Exit Function
    End If
    If mSize >= 1 And mSize <= 41 Then
        getPapersize = Switch( _
            mSize = 1, conSize1, mSize = 2, conSize2, mSize = 3, conSize3, mSize = 4, conSize4, mSize = 5, conSize5, _
            mSize = 6, conSize6, mSize = 7, conSize7, mSize = 8, conSize8, mSize = 9, conSize9, mSize = 10, conSize10, _
            mSize = 11, conSize11, mSize = 12, conSize12, mSize = 13, conSize13, mSize = 14, conSize14, mSize = 15, conSize15, _
            mSize = 16, conSize16, mSize = 17, conSize17, mSize = 18, conSize18, mSize = 19, conSize19, mSize = 20, conSize20, _
            mSize = 21, conSize21, mSize = 22, conSize22, mSize = 23, conSize23, mSize = 24, conSize24, mSize = 25, conSize25, _
            mSize = 26, conSize26, mSize = 27, conSize27, mSize = 28, conSize28, mSize = 29, conSize29, mSize = 30, conSize30, _
            mSize = 31, conSize31, mSize = 32, conSize32, mSize = 33, conSize33, mSize = 34, conSize34, mSize = 35, conSize35, _
            mSize = 36, conSize36, mSize = 37, conSize37, mSize = 38, conSize38, mSize = 39, conSize39, mSize = 40, conSize40, _
            mSize = 41, conSize41)
        Exit Function
    End If
ErrHand:
    getPapersize = "不可测的纸张"
End Function


Public Function getPaperBin(mBin As Integer) As String
    '------------------------------------------------
    '功能： 根据当前打印机的设置，获取送纸方式描述
    '返回： 送纸方式字符串
    '------------------------------------------------
    Err = 0
    On Error GoTo ErrHand
    
    If mBin = 14 Then
        getPaperBin = "附加的卡式纸盒进纸"
        Exit Function
    End If
    If mBin >= 1 And mBin <= 11 Then
        getPaperBin = Switch( _
            mBin = 1, conBin1, mBin = 2, conBin2, mBin = 3, conBin3, mBin = 4, conBin4, mBin = 5, conBin5, _
            mBin = 6, conBin6, mBin = 7, conBin7, mBin = 8, conBin8, mBin = 9, conBin9, mBin = 10, conBin10, _
            mBin = 11, conBin11)
        Exit Function
    End If
ErrHand:
    getPaperBin = "自动选择..."
    
End Function

Private Sub txtHeight_Validate(Cancel As Boolean)
    If Not IsNumeric(txtHeight.Text) Then
        txtHeight.Text = updnHeight.Value
        Exit Sub
    End If
    If Val(txtHeight.Text) < updnHeight.Min Or Val(txtHeight.Text) > updnHeight.Max Then
        txtHeight.Text = updnHeight.Value
        Exit Sub
    End If
    
    updnHeight.Value = Val(txtHeight.Text)
End Sub

Private Sub txtWidth_GotFocus()
    txtWidth.SelStart = 0
    txtWidth.SelLength = 3
End Sub

Private Sub txtWidth_Validate(Cancel As Boolean)
    If Not IsNumeric(txtWidth.Text) Then
        txtWidth.Text = updnWidth.Value
        Exit Sub
    End If
    If Val(txtWidth.Text) < updnWidth.Min Or Val(txtWidth.Text) > updnWidth.Max Then
        txtWidth.Text = updnWidth.Value
        Exit Sub
    End If
    
    updnWidth.Value = Val(txtWidth.Text)
End Sub

Private Sub txtHeight_GotFocus()
    txtHeight.SelStart = 0
    txtHeight.SelLength = 3
End Sub
