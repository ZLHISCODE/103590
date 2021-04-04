VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDrugInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "药品信息上传"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6030
   Icon            =   "frmDrugInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   6030
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdUpload 
      Caption         =   "上传(&U)"
      Height          =   360
      Left            =   3480
      TabIndex        =   4
      Top             =   5400
      Width           =   1110
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   360
      Left            =   4680
      TabIndex        =   5
      Top             =   5400
      Width           =   1110
   End
   Begin TabDlg.SSTab sstDrug 
      Height          =   4575
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   8070
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "剂型"
      TabPicture(0)   =   "frmDrugInfo.frx":0A02
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tvwDrugType"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "上传消息"
      TabPicture(1)   =   "frmDrugInfo.frx":0A1E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstMess"
      Tab(1).ControlCount=   1
      Begin VB.ListBox lstMess 
         Height          =   3840
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   5535
      End
      Begin MSComctlLib.TreeView tvwDrugType 
         Height          =   3975
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   7011
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
   End
   Begin VB.ComboBox cboLink 
      Height          =   300
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   4935
   End
   Begin VB.Label lblLink 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "连接名："
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmDrugInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const STR_ROOT = "ROOT"

Private Sub cboLink_Change()
    Dim i As Integer
    Dim blnAll As Boolean
    For i = 1 To tvwDrugType.Nodes.Count
        If tvwDrugType.Nodes(i).Checked Then
            blnAll = True
            Exit For
        End If
    Next
    cmdUpload.Enabled = cboLink.ListIndex >= 0 And blnAll
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdUpload_Click()
    Dim rsData As ADODB.Recordset
    Dim strTmp As String
    Dim i As Integer
    
    '药品剂型字符串
    If tvwDrugType.Nodes(STR_ROOT).Checked = True Then
        strTmp = "0"
    Else
        For i = 0 To tvwDrugType.Nodes.Count - 1
            If tvwDrugType.Nodes(i).Key <> STR_ROOT Then
                strTmp = strTmp & "'" & tvwDrugType.Nodes(i).Text & "'"
                If tvwDrugType.Nodes.Count - 1 > i Then
                    strTmp = strTmp & ","
                End If
            End If
        Next
    End If
    
    '
    Call mdlProcessData.ProcDrugInfo(strTmp, cboLink.Text)
    
    '循环上传
    'mdlDrugPacker.DrugInfo cboLink.Text, strContent
    
    sstDrug.TabIndex = 1
    
End Sub

Private Sub Form_Load()
    Call InitLink
    If cboLink.ListCount = 0 Then
        MsgBox "尚未设置药房自动化设备的连接！", vbInformation, GSTR_INTERFACE_NAME
        Unload Me
        Exit Sub
    End If
    Call InitDrugType
    cmdUpload.Enabled = False
End Sub

Private Sub InitDrugType()
'功能：加载药品剂型

    Dim rsTmp As ADODB.Recordset
    
    tvwDrugType.Nodes.Add , , STR_ROOT, "全部"
    
    gstrSQL = "Select 编码, 名称 From 药品剂型 Order By 名称 "
    On Error GoTo errHandle
    Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "获取药品剂型")
    Do While Not rsTmp.EOF
        tvwDrugType.Nodes.Add STR_ROOT, tvwChild, "N_" & rsTmp!编码, rsTmp!名称
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    tvwDrugType.Nodes(STR_ROOT).Expanded = True
    
    Exit Sub
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub

Private Sub InitLink()
'功能：加载连接
    
    Dim rsTmp As ADODB.Recordset
        
    gstrSQL = "Select ID, 名称, 连接类型 From 药房设备连接 Order By 名称 "
    On Error GoTo errHandle
    Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "获取药房设备连接")
    Do While Not rsTmp.EOF
        cboLink.AddItem rsTmp!名称
        cboLink.ItemData(cboLink.NewIndex) = rsTmp!连接类型
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    If cboLink.ListCount >= 0 Then cboLink.ListIndex = 0
    
    Exit Sub
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub

Private Sub tvwDrugType_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim i As Integer
    Dim blnAll As Boolean
    
    If Node.Key = STR_ROOT Then
        cmdUpload.Enabled = Node.Checked
        For i = 1 To tvwDrugType.Nodes.Count
            tvwDrugType.Nodes(i).Checked = Node.Checked
        Next
    Else
        blnAll = True
        For i = 1 To tvwDrugType.Nodes.Count
            If tvwDrugType.Nodes(i).Checked = False And tvwDrugType.Nodes(i).Key <> STR_ROOT Then
                blnAll = False
                Exit For
            End If
        Next
        tvwDrugType.Nodes(STR_ROOT).Checked = blnAll
        
        For i = 1 To tvwDrugType.Nodes.Count
            If tvwDrugType.Nodes(i).Checked Then
                cmdUpload.Enabled = True And Me.cboLink.ListIndex >= 0
                Exit Sub
            End If
        Next
        cmdUpload.Enabled = False
    End If
End Sub

