VERSION 5.00
Begin VB.Form frmPrintSet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "打印设置"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   Icon            =   "frmPrintSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Caption         =   "打印机"
      Height          =   1485
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   5850
      Begin VB.ComboBox cboBin 
         Height          =   300
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   960
         Width           =   3885
      End
      Begin VB.ComboBox cboPrinter 
         Height          =   300
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   225
         Width           =   3885
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "纸张来源"
         Height          =   180
         Left            =   825
         TabIndex        =   4
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "名称"
         Height          =   180
         Left            =   1185
         TabIndex        =   1
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lblLoc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "位置"
         Height          =   180
         Left            =   1185
         TabIndex        =   3
         Top             =   660
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   390
         Picture         =   "frmPrintSet.frx":058A
         Top             =   330
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   3600
      TabIndex        =   6
      Top             =   1710
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4800
      TabIndex        =   7
      Top             =   1710
      Width           =   1100
   End
End
Attribute VB_Name = "frmPrintSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnWinNT As Boolean
Private mdblW As Double  '左边不可打印比例
Private mdblH As Double  '上边不可打印比例

'打印参数变量
Private mstrPrinter As String '打印机
Private mstrBin As String '进纸方式

'事件控制
Private mblnChange As Boolean
Private mbytMode As Byte

Public Sub ShowMe(ByVal frmParent As Object, Optional ByVal bytMode As Byte = 1)
'----------------------------------------------------
'
'---------------------------------------------------
    mbytMode = bytMode
    Me.Show 1, frmParent
End Sub


Private Sub cboPrinter_Click()
    Dim i As Integer, j As Integer
    Dim lngCount As Long, strtmp As String
    Dim strPaperBinName As String * 1000
    Dim strPaperbins As String, strTemp As String, strCount As String
    
    Set Printer = Printers(cboPrinter.ItemData(cboPrinter.ListIndex))
    mstrPrinter = Printer.DeviceName
    lblLoc.Caption = "位置: " & Printer.Port
    

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '如果支持,则保持原有进纸方式
    On Error Resume Next
    Printer.PaperBin = mstrBin
    On Error GoTo 0
    mstrBin = Printer.PaperBin
    
   '设置可用进纸方式
    lngCount = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_BINS, strPaperBinName, 0)
    For i = 1 To lngCount
        j = Asc(Mid(strPaperBinName, i * 2, 1)) * 256# + Asc(Mid(strPaperBinName, i * 2 - 1, 1))
        If j >= 1 And j <= 11 Then '只列出标准支持的进纸大小
            If j = mstrBin Then
                strPaperbins = strPaperbins & "," & j & "*" '原有的
            Else
                strPaperbins = strPaperbins & "," & j
            End If
        End If
    Next
    Err = 0
    
    If Printer.PaperBin = 14 Then
        strPaperbins = strPaperbins & ",14" _
            & IIf(mstrBin = 14, "*", "")
    End If
    
    strPaperbins = Mid(strPaperbins, 2)
'    '纸张来源
    With cboBin
        .Clear
        strTemp = strPaperbins
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
End Sub

Public Function getPaperBin(mBin As Integer) As String
    '------------------------------------------------
    '功能： 根据当前打印机的设置，获取送纸方式描述
    '返回： 送纸方式字符串
    '------------------------------------------------
    Err = 0
    On Error GoTo errHand
    
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
errHand:
    getPaperBin = "自动选择..."
    
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    
    mstrBin = ""
    mstrBin = Me.cboBin.ItemData(Me.cboBin.ListIndex)
    '保存打印参数
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\Default", "DeviceName", mstrPrinter)
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\Default", "PaperBin", "")
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    If Not ExistsPrinter Then
        MsgBox "系统中没有安装任何打印机,请先安装打印机！", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    mblnChange = True
    
'    初始化打印参数
    mstrPrinter = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\Default", "DeviceName", Printers(0).DeviceName)
    mstrBin = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\Default", "PaperBin", "")
    
    '初始打印机列表
    With cboPrinter
        .Clear
        For i = 0 To Printers.Count - 1
            .AddItem Printers(i).DeviceName
            .ItemData(.ListCount - 1) = i '打印机索引
            
            '读取存储的打印机为当前打印机,并初始化可用页面
            If mstrPrinter = Printers(i).DeviceName Then .ListIndex = .NewIndex
        Next
        
        '缺省初始化为当前打印机
        If .ListIndex = -1 Then
            For i = 0 To .ListCount - 1
                '读取系统当前的打印机为当前打印机,并初始化可用页面
                If .List(i) = Printer.DeviceName Then .ListIndex = i: Exit For
            Next
        End If
    End With
End Sub

