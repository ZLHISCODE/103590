VERSION 5.00
Begin VB.Form frmCISAduitPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "打印档案选项"
   ClientHeight    =   5445
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7320
   Icon            =   "frmCISAduitPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraPrinter 
      Caption         =   "打印机"
      Height          =   1470
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   5775
      Begin VB.ComboBox cboPrinterName 
         Height          =   300
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   225
         Width           =   3915
      End
      Begin VB.Label lblPrinterInfo2 
         AutoSize        =   -1  'True
         Caption         =   "默认打印机:是"
         Height          =   180
         Left            =   1695
         TabIndex        =   4
         Top             =   990
         Width           =   1170
      End
      Begin VB.Label lblPrinterName 
         AutoSize        =   -1  'True
         Caption         =   "名称(&N)"
         Height          =   180
         Left            =   945
         TabIndex        =   1
         Top             =   300
         Width           =   630
      End
      Begin VB.Label lblPrinterInfo 
         AutoSize        =   -1  'True
         Caption         =   "位置:连接到LTP1:"
         Height          =   180
         Left            =   1680
         TabIndex        =   3
         Top             =   645
         Width           =   1440
      End
      Begin VB.Image imgPrinter 
         Height          =   360
         Left            =   270
         Picture         =   "frmCISAduitPrint.frx":000C
         Top             =   270
         Width           =   360
      End
   End
   Begin VB.Frame fraPageScope 
      Caption         =   "打印范围(&R)"
      Height          =   3795
      Left            =   30
      TabIndex        =   5
      Top             =   1590
      Width           =   5775
      Begin VB.ListBox lst 
         Height          =   3420
         Left            =   150
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   285
         Width           =   3600
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "全清(&D)"
         Height          =   350
         Left            =   3840
         TabIndex        =   8
         Top             =   660
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "全选(&A)"
         Height          =   350
         Left            =   3855
         TabIndex        =   7
         Top             =   240
         Width           =   1100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "全清：Shift+Delete"
         Height          =   180
         Left            =   3870
         TabIndex        =   10
         Top             =   1560
         Width           =   1620
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "全选：Ctrl+A"
         Height          =   180
         Left            =   3870
         TabIndex        =   9
         Top             =   1245
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6045
      TabIndex        =   11
      Top             =   135
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6045
      TabIndex        =   12
      Top             =   600
      Width           =   1100
   End
End
Attribute VB_Name = "frmCISAduitPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnOK As Boolean
Private mstrPrinterDeviceName As String
Private mstrPrintRange As String

Public Function ShowDialog(ByVal frmMain As Object, ByRef strPrinterDeviceName As String, ByRef strPrintRange As String) As Boolean
    Dim intCount As Integer
    Dim arySerial As Variant
    Dim strTmp As String
    Dim strSelect As String
    
    mblnOK = False
    mstrPrintRange = strPrintRange
    mstrPrinterDeviceName = GetRegister(私有模块, "打印档案", "打印机", Printer.DeviceName)
    If mstrPrinterDeviceName = "" Then mstrPrinterDeviceName = Printer.DeviceName
    With cboPrinterName
        .Clear
        For intCount = 0 To Printers.count - 1
            .AddItem Printers(intCount).DeviceName
            If Printers(intCount).DeviceName = mstrPrinterDeviceName Then .ListIndex = intCount
        Next
    End With
    
    '----------------------------------------------------------------------------------------
    '1-住院医嘱;2-住院病历;3-护理病历;4-护理记录;5-首页记录;6-医嘱报告;7-疾病证明;8-知情文件
    
    strSelect = "," & GetRegister(私有模块, "打印档案", "打印内容", "1,2,3,4,5,6,7,8,9") & ","
    
    strTmp = Trim(zlDatabase.GetPara("档案排序顺序", ParamInfo.系统号, 1560, "5;1;6;2;3;4;8;7;9"))
    If strTmp = "" Then strTmp = "5;1;6;2;3;4;8;7;9"
    arySerial = Split(strTmp, ";")
    
    With lst
        For intCount = 0 To UBound(arySerial)
            Select Case Val(arySerial(intCount))
            Case 1
                .AddItem "住院医嘱": .ItemData(.NewIndex) = 1
                If InStr(strSelect, ",1,") > 0 Then .Selected(.NewIndex) = True
            Case 2
                .AddItem "住院病历": .ItemData(.NewIndex) = 2
                If InStr(strSelect, ",2,") > 0 Then .Selected(.NewIndex) = True
            Case 3
                .AddItem "护理病历": .ItemData(.NewIndex) = 3
                If InStr(strSelect, ",3,") > 0 Then .Selected(.NewIndex) = True
            Case 4
                .AddItem "护理记录": .ItemData(.NewIndex) = 4
                If InStr(strSelect, ",4,") > 0 Then .Selected(.NewIndex) = True
            Case 5
                .AddItem "首页记录": .ItemData(.NewIndex) = 5
                If InStr(strSelect, ",5,") > 0 Then .Selected(.NewIndex) = True
            Case 6
                .AddItem "医嘱报告": .ItemData(.NewIndex) = 6
                If InStr(strSelect, ",6,") > 0 Then .Selected(.NewIndex) = True
            Case 7
                .AddItem "疾病证明": .ItemData(.NewIndex) = 7
                If InStr(strSelect, ",7,") > 0 Then .Selected(.NewIndex) = True
            Case 8
                .AddItem "知情文件": .ItemData(.NewIndex) = 8
                If InStr(strSelect, ",8,") > 0 Then .Selected(.NewIndex) = True
            Case 9
                .AddItem "临床路径": .ItemData(.NewIndex) = 9
                If InStr(strSelect, ",9,") > 0 Then .Selected(.NewIndex) = True
            End Select
        Next

        .ListIndex = 0
    End With
    
    Me.Show 1, frmMain
    
    If mblnOK Then
        strPrintRange = mstrPrintRange
        strPrinterDeviceName = mstrPrinterDeviceName
    End If
    
    ShowDialog = mblnOK
    
End Function

Private Sub cboPrinterName_Click()
    
    lblPrinterInfo.Caption = "位置:连接到" & Printers(cboPrinterName.ListIndex).Port
    lblPrinterInfo2.Caption = "默认打印机:" & IIf(Printers(cboPrinterName.ListIndex).DeviceName = Printer.DeviceName, "是", "否")

End Sub

Private Sub cboPrinterName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClearAll_Click()
    Dim intCount As Integer
    
    For intCount = 0 To lst.ListCount - 1
        lst.Selected(intCount) = False
    Next
    
End Sub

Private Sub cmdOK_Click()
    Dim strSelect As String
    Dim intCount As Integer
    
    mstrPrintRange = ""
    For intCount = 0 To lst.ListCount - 1
        If lst.Selected(intCount) = True Then
            mstrPrintRange = mstrPrintRange & "," & lst.ItemData(intCount)
        End If
    Next
    If mstrPrintRange <> "" Then mstrPrintRange = Mid(mstrPrintRange, 2)
    
    mstrPrinterDeviceName = cboPrinterName.Text
        
    Call SetRegister(私有模块, "打印档案", "打印机", cboPrinterName.Text)
    
    Call SetRegister(私有模块, "打印档案", "打印内容", mstrPrintRange)
    
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdSelectAll_Click()
    Dim intCount As Integer
    
    For intCount = 0 To lst.ListCount - 1
        lst.Selected(intCount) = True
    Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
    Case 1
    
        If KeyCode = vbKeyDelete Then
            Call cmdClearAll_Click
        End If
    Case 2
        If KeyCode = vbKeyA Then
            Call cmdSelectAll_Click
        End If
    End Select
End Sub

Private Sub lst_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
