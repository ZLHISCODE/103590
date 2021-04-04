VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRptMan 
   BackColor       =   &H80000005&
   Caption         =   "报表管理"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5445
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmRptMan.frx":0000
   ScaleHeight     =   4275
   ScaleWidth      =   5445
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdEnter 
      Caption         =   "现在进入报表工具(&E)… "
      Height          =   350
      Left            =   915
      TabIndex        =   1
      Top             =   3585
      Width           =   2190
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   120
      Top             =   1710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptMan.frx":04F9
            Key             =   "K0501"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptMan.frx":228B
            Key             =   "K0502"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptMan.frx":401D
            Key             =   "K0505"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "报表管理"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   200
      TabIndex        =   2
      Top             =   100
      Width           =   960
   End
   Begin VB.Label lblMain 
      BackStyle       =   0  'Transparent
      Height          =   3330
      Left            =   945
      TabIndex        =   0
      Top             =   600
      Width           =   4140
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   255
      Picture         =   "frmRptMan.frx":8E1F
      Top             =   600
      Width           =   480
   End
End
Attribute VB_Name = "frmRptMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstr编号 As String    '窗体编号
Private mfrmProcMain As frmProcMain

Private Sub Form_Load()
    Select Case mstr编号
        Case "0501"
            lblTitle.Caption = "报表管理"
            cmdEnter.Caption = "现在进入报表工具(&E)… "
            lblMain.Caption = "完成系统各种票据格式与输出内容的定义修改。" & _
                vbCrLf & vbCrLf & "采用面向对象的策略，独特的图元定制方式（图形元素点选描绘），精确定制票据与报表，可随心所欲地调整票据的纸张特性(大小、类型)、输出格式（字体、颜色、排列），并可立即预览打印。" & _
                vbCrLf & vbCrLf & "增强SQL，实现票据输出数据内容的改变，自动检测书写正确性。" & _
                vbCrLf & vbCrLf & "如果具有“服务器管理报表工具”选件功能，将可以在现有系统的基础上增加设置新的多种报表，实现特殊数据分析。"
        Case "0502"
            lblTitle.Caption = "函数管理"
            cmdEnter.Caption = "现在进入函数工具(&E)… "
            lblMain.Caption = "完成各系统数据传递函数的管理，包括函数文本及其参数向导的定义、修改与设置。" & _
                vbCrLf & vbCrLf & "数据传递函数是本软件各应用系统间相互抽选传递数据的重要方式，使整个应用成为一个完整的整体；较多地应用于财务总帐自动凭证、成本效益核算和报表分析提取各应用系统的发生数据。" & _
                vbCrLf & vbCrLf & "软件系统装载时，已经装入可提供一些典型数据提取的函数，普通用户获得授权后即可使用；" & _
                vbCrLf & vbCrLf & "必要时，系统所有者用户可根据应用需要，增加新的函数或修改现有函数，实现对应用系统的任意数据的提取。"
        Case "0505"
            lblTitle.Caption = "过程管理"
            cmdEnter.Caption = "现在进入过程工具(&E)"
            lblMain.Caption = "完成各系统正常升级过程中，对自定义过程的修改及管理。" & _
            vbCrLf & vbCrLf & "当前数据库与脚本对比，自动搜集调整过的过程。" & _
            vbCrLf & vbCrLf & "对比升级前后过程的变化，并得出对应的差异对比报告，升级人员可直接根据差异部分内容作出判断修改相应的过程。"
    End Select
    Me.Caption = lblTitle.Caption
    imgMain.Picture = ils32.ListImages("K" & mstr编号).Picture
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With imgMain
        .Top = 700
        .Left = ScaleLeft + 200
    End With
    
    With lblMain
        .Top = imgMain.Top
        .Height = ScaleHeight - .Top * 2
        .Left = imgMain.Left * 2 + imgMain.Width
        .Width = ScaleWidth - .Left - imgMain.Left
    End With

    Dim intCount As Integer, intRows As Integer, aryRow() As String
    intRows = 1
    aryRow() = Split(lblMain.Caption, vbCrLf)
    For intCount = 0 To UBound(aryRow)
        intRows = intRows + TextWidth(aryRow(intCount)) \ (lblMain.Width - 90) + 1
    Next
    If intRows * TextHeight("A") < lblMain.Height + TextHeight("A") Then
        cmdEnter.Top = lblMain.Top + intRows * TextHeight("A")
    Else
        cmdEnter.Top = lblMain.Top + lblMain.Height + TextHeight("A")
    End If
    cmdEnter.Left = lblMain.Left
    
End Sub

Private Sub cmdEnter_Click()
    Dim frmMain As frmConnectionsManager
    
    Select Case mstr编号
        Case "0501"
            If gobjReport Is Nothing Then
                Set gobjReport = CreateObject("zl9Report.clsReport")
            End If
            Set frmMain = New frmConnectionsManager
            gobjReport.ReportMan gcnOracle, frmMDIMain, gstrLoginUserName, frmMain
        Case "0502"
            If gobjFunction Is Nothing Then
                Set gobjFunction = CreateObject("zl9Function.clsFunction")
            End If
            
            '函数管理工具暂不支持新连接
            gobjFunction.funcmanage gcnOldOra, frmMDIMain
        Case "0505"
            If mfrmProcMain Is Nothing Then
                Set mfrmProcMain = New frmProcMain
            End If
            Call mfrmProcMain.ShowMe(frmMDIMain)
    End Select
End Sub

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口

End Sub


