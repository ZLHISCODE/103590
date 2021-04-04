VERSION 5.00
Begin VB.Form frmPrintSet 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   0
   ClientWidth     =   6930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5700
      TabIndex        =   4
      Top             =   6480
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4440
      TabIndex        =   3
      Top             =   6480
      Width           =   1100
   End
   Begin VB.Frame fraPrinter 
      Caption         =   "打印机"
      Height          =   6195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6690
      Begin VB.ComboBox cboPrinterName 
         Height          =   300
         Index           =   11
         Left            =   1695
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Tag             =   "R10"
         Top             =   5520
         Width           =   4635
      End
      Begin VB.ComboBox cboPrinterName 
         Height          =   300
         Index           =   10
         Left            =   1695
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Tag             =   "R9"
         Top             =   5040
         Width           =   4635
      End
      Begin VB.ComboBox cboPrinterName 
         Height          =   300
         Index           =   9
         Left            =   1695
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Tag             =   "R8"
         Top             =   4560
         Width           =   4635
      End
      Begin VB.ComboBox cboPrinterName 
         Height          =   300
         Index           =   8
         Left            =   1695
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Tag             =   "R6"
         Top             =   4080
         Width           =   4635
      End
      Begin VB.ComboBox cboPrinterName 
         Height          =   300
         Index           =   7
         Left            =   1695
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Tag             =   "R5"
         Top             =   3600
         Width           =   4635
      End
      Begin VB.ComboBox cboPrinterName 
         Height          =   300
         Index           =   6
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Tag             =   "R7"
         Top             =   3120
         Width           =   4635
      End
      Begin VB.ComboBox cboPrinterName 
         Height          =   300
         Index           =   5
         Left            =   1695
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Tag             =   "R4"
         Top             =   2640
         Width           =   4635
      End
      Begin VB.ComboBox cboPrinterName 
         Height          =   300
         Index           =   4
         Left            =   1695
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Tag             =   "R3"
         Top             =   2160
         Width           =   4635
      End
      Begin VB.ComboBox cboPrinterName 
         Height          =   300
         Index           =   3
         Left            =   1695
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Tag             =   "R1"
         Top             =   1680
         Width           =   4635
      End
      Begin VB.ComboBox cboPrinterName 
         Height          =   300
         Index           =   2
         Left            =   1695
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Tag             =   "R2"
         Top             =   1200
         Width           =   4635
      End
      Begin VB.ComboBox cboPrinterName 
         Height          =   300
         Index           =   1
         Left            =   1695
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "R12"
         Top             =   720
         Width           =   4635
      End
      Begin VB.ComboBox cboPrinterName 
         Height          =   300
         Index           =   0
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Tag             =   "R11"
         Top             =   285
         Width           =   4635
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "其他报表"
         Height          =   180
         Index           =   11
         Left            =   825
         TabIndex        =   26
         Top             =   5580
         Width           =   720
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "住院证"
         Height          =   180
         Index           =   10
         Left            =   1005
         TabIndex        =   24
         Top             =   5100
         Width           =   540
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "临床路径"
         Height          =   180
         Index           =   9
         Left            =   825
         TabIndex        =   22
         Top             =   4620
         Width           =   720
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "知情文件"
         Height          =   180
         Index           =   8
         Left            =   825
         TabIndex        =   20
         Top             =   4140
         Width           =   720
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "疾病证明"
         Height          =   180
         Index           =   7
         Left            =   825
         TabIndex        =   18
         Top             =   3660
         Width           =   720
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "诊疗报告"
         Height          =   180
         Index           =   6
         Left            =   825
         TabIndex        =   16
         Top             =   3180
         Width           =   720
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "护理病历"
         Height          =   180
         Index           =   5
         Left            =   825
         TabIndex        =   14
         Top             =   2700
         Width           =   720
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "护理记录 "
         Height          =   180
         Index           =   4
         Left            =   825
         TabIndex        =   12
         Top             =   2220
         Width           =   810
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "门诊病历"
         Height          =   180
         Index           =   3
         Left            =   825
         TabIndex        =   10
         Top             =   1740
         Width           =   720
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "住院病历"
         Height          =   180
         Index           =   2
         Left            =   825
         TabIndex        =   8
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "医嘱记录"
         Height          =   180
         Index           =   1
         Left            =   825
         TabIndex        =   6
         Top             =   780
         Width           =   720
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   240
         Picture         =   "frmPrintSet.frx":0000
         Top             =   360
         Width           =   240
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "首页信息"
         Height          =   180
         Index           =   0
         Left            =   825
         TabIndex        =   1
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

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    
    '1-门诊病历;2-住院病历;3-护理记录;4-护理病历;5-疾病证明;6-知情文件;7-诊疗报告,11-首页信息,12-医嘱记录,8-临床路径;9-住院证;10-其他报表
    For i = cboPrinterName.LBound To cboPrinterName.UBound
        Call SetRegister(私有模块, "打印档案", "打印机" & cboPrinterName(i).Tag, Trim(cboPrinterName(i).Text))
    Next
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strPrinterName  As String
    Dim intCount        As Integer
    Dim i               As Long
    
    If Printers.Count = 0 Then
        MsgBox "注意：" & Chr(13) _
            & "    未安装打印机，请通过系统设置的打印机" & Chr(13) _
            & "管理添加安装打印机。", vbCritical + vbOKOnly, gstrSysName
        Exit Sub
    End If
    '加载打印设备
    For i = cboPrinterName.LBound To cboPrinterName.UBound
        strPrinterName = GetRegister(私有模块, "打印档案", "打印机" & cboPrinterName(i).Tag, Printer.DeviceName)
        With cboPrinterName(i)
            .Clear
            For intCount = 0 To Printers.Count - 1
                .AddItem Printers(intCount).DeviceName
                If Printers(intCount).DeviceName = strPrinterName Then .ListIndex = intCount
            Next
        End With
    Next
End Sub

