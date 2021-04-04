VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmParaSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "报表设置"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7545
   Icon            =   "frmParaSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7545
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraPara 
      Caption         =   "检验报告打印"
      Height          =   1215
      Left            =   600
      TabIndex        =   10
      Top             =   3000
      Width           =   6135
      Begin VB.OptionButton optLIS 
         Caption         =   "老版LIS报表或病历"
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   12
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton optLIS 
         Caption         =   "新版LIS报表"
         Height          =   375
         Index           =   1
         Left            =   4440
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame fra 
      Height          =   45
      Left            =   0
      TabIndex        =   8
      Top             =   1080
      Width           =   7575
   End
   Begin VB.ComboBox cboPacs 
      Height          =   300
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2340
      Width           =   4935
   End
   Begin VB.ComboBox cboLis 
      Height          =   300
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1680
      Width           =   4935
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   7545
      TabIndex        =   0
      Top             =   4470
      Width           =   7545
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   5070
         TabIndex        =   2
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   6225
         TabIndex        =   1
         Top             =   240
         Width           =   1100
      End
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "2、报表参数固定：【病人ID】数字型  【主页ID】数字型  【医嘱ID】字符型  将检验/检查医嘱ID用逗号拼接构成。"
      ForeColor       =   &H00C00000&
      Height          =   360
      Index           =   1
      Left            =   840
      TabIndex        =   9
      Top             =   600
      Width           =   6555
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgInfo 
      Height          =   600
      Left            =   120
      Picture         =   "frmParaSet.frx":6852
      Top             =   240
      Width           =   600
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   120
      Top             =   4440
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "1、设置检验/检查报告对应的报表，以便在预览/打印时将所选项目按对应报表处理。若对应报表为空,则按缺省方式处理。"
      ForeColor       =   &H00C00000&
      Height          =   480
      Index           =   0
      Left            =   855
      TabIndex        =   7
      Top             =   120
      Width           =   6540
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "检查对应报表"
      Height          =   180
      Left            =   600
      TabIndex        =   4
      Top             =   2400
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "检验对应报表"
      Height          =   180
      Left            =   600
      TabIndex        =   3
      Top             =   1740
      Width           =   1080
   End
End
Attribute VB_Name = "frmParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrPrivs As String
Private mlngModul As Long
Private mcolReport As Collection

Private Sub cboLis_Click()
    If cboLis.ListIndex > 0 Then
        cboLis.Tag = mcolReport("_" & cboLis.ListIndex)
    Else
        cboLis.Tag = ""
    End If
End Sub

Private Sub cboPacs_Click()
    If cboPacs.ListIndex > 0 Then
        cboPacs.Tag = mcolReport("_" & cboPacs.ListIndex)
    Else
        cboPacs.Tag = ""
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    
    Call zlDatabase.SetPara("检查对应报表", cboPacs.Tag, glngSys, glngModul, True)
    Call zlDatabase.SetPara("检验对应报表", cboLis.Tag, glngSys, glngModul, True)
    Call zlDatabase.SetPara("检验报告打印", optLIS(0).Tag, glngSys, glngModul, True)
    Unload Me
End Sub

Private Sub Form_Load()
    Dim objControl As CommandBarControl
    Dim objPop As Object
    Dim strHide As String, strPara As String
    Dim i As Long
    
    '隐藏,此处只用做容器接收报表内容
    cbsMain.ActiveMenuBar.Visible = False
    strHide = "ZL1_INSIDE_1254_1,ZL1_INSIDE_1254_2,ZL1_INSIDE_1261_1,ZL1_INSIDE_1261_4,ZL1_INSIDE_1261_5,ZL1_INSIDE_1261_6,ZL1_INSIDE_1261_7,ZL1_INSIDE_1261_8,ZL1_INSIDE_1261_9,ZL1_INSIDE_1261_10"
    Call zlDatabase.ShowReportMenu(cbsMain, glngSys, mlngModul, mstrPrivs, strHide)
    
    For i = 1 To cbsMain.ActiveMenuBar.Controls.count
        If cbsMain.ActiveMenuBar.Controls(i).ID = conMenu_ReportPopup _
            Or cbsMain.ActiveMenuBar.Controls(i).Caption Like "报表*" Then
            Set objControl = cbsMain.ActiveMenuBar.Controls.Item(i)
            Exit For
        End If
    Next
    
    If Not objControl Is Nothing Then
        With objControl.CommandBar.Controls
            For i = 1 To .count
                Set objPop = .Item(i)
                mcolReport.Add Split(objPop.Caption, "(&")(0) & "," & objPop.Parameter, "_" & i     '报表名称,系统号,报表编号
            Next
        End With
    End If
    
    '加载下拉框
    cboLis.AddItem "", 0
    cboPacs.AddItem "", 0
    For i = 1 To mcolReport.count
        cboLis.AddItem Split(mcolReport(i), ",")(0), i
        cboPacs.AddItem Split(mcolReport(i), ",")(0), i
    Next
    If cboLis.ListCount > 0 Then cboLis.ListIndex = 0
    If cboPacs.ListCount > 0 Then cboPacs.ListIndex = 0
    
    strPara = zlDatabase.GetPara("检验对应报表", glngSys, glngModul, "", cboLis, True) '报表名称,系统号,报表编号
    If strPara <> "" Then Call cbo.Locate(cboLis, Split(strPara, ",")(0))
    strPara = zlDatabase.GetPara("检查对应报表", glngSys, glngModul, "", cboPacs, True)
    If strPara <> "" Then Call cbo.Locate(cboPacs, Split(strPara, ",")(0))
    If IsSysSetUp(2500) Then
        fraPara.Visible = True
        Me.Height = 5625
        strPara = zlDatabase.GetPara("检验报告打印", glngSys, glngModul, "0", optLIS, True)
        optLIS(0).Value = Val(strPara) = 0
        optLIS(1).Value = Val(strPara) = 1
    Else
        fraPara.Visible = False
        Me.Height = 4590
    End If
End Sub

Public Function ShowMe(objMain As Object, ByVal lngSys As Long, ByVal lngModul As Long, ByVal strPrivs As String) As Boolean
    mlngModul = lngModul
    glngSys = lngSys
    mstrPrivs = strPrivs
    Set mcolReport = New Collection
    Me.Show 1, objMain
End Function

Private Sub optLIS_Click(Index As Integer)
    If optLIS(Index).Value Then optLIS(0).Tag = Index
End Sub
