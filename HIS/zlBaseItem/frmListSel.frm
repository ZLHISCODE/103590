VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListSel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "列表选择"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ListView lvwMain 
      Height          =   4755
      Left            =   165
      TabIndex        =   2
      Top             =   210
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   8387
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4155
      TabIndex        =   1
      Top             =   660
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4155
      TabIndex        =   0
      Top             =   210
      Width           =   1100
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListSel.frx":0000
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListSel.frx":0454
            Key             =   "Root"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmListSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrID As String
Private mstr编码 As String
Private mstr名称 As String
Private mblnOk As Boolean

Public Function ShowLvw(frmParent As Object, ByVal strSQL As String, strID As String, str编码 As String, str名称 As String, Optional ByVal strCaption As String = "选择器") As Boolean
    On Error GoTo ErrHandle
    '显示列表选择器
    '参数：strSql  = 数据源
    '      strID = 返回ID
    '      str编码 = 返回编码
    '      str名称 = 返回名称
    '      strCaption = 窗体标题
    
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    Dim objList As ListItem
    
        mstrID = strID
        mstr编码 = str编码
        mstr名称 = str名称
        
        gstrSQL = strSQL
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, strCaption)
        lvwMain.ListItems.Clear
        
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            For i = 0 To rsTmp.RecordCount - 1
                Set objList = lvwMain.ListItems.Add(, "I" & rsTmp!ID, "", "Write", "Write")
                    objList.Text = "【" & zlCommFun.Nvl(rsTmp!编码) & "】"
                    objList.SubItems(1) = zlCommFun.Nvl(rsTmp!名称)
                    objList.Tag = rsTmp!ID
                rsTmp.MoveNext
            Next
        End If
        
        Me.Caption = strCaption
        Me.Show 1, frmParent
        If mblnOk = True Then
            strID = mstrID
            str编码 = mstr编码
            str名称 = mstr名称
        End If
        ShowLvw = mblnOk
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If lvwMain.ListItems.Count < 1 Then MsgBox "无任何选择可供选择！", vbInformation, gstrSysName: Exit Sub
    If lvwMain.SelectedItem Is Nothing Then MsgBox "请选择一个项目！", vbInformation, gstrSysName: Exit Sub
    lvwMain_ItemClick lvwMain.SelectedItem
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_Load()
lvwMain.ColumnHeaders.Clear
lvwMain.ColumnHeaders.Add , , "编码", 1400
lvwMain.ColumnHeaders.Add , , "名称", 2000
zlControl.LvwFlatColumnHeader Me.lvwMain

End Sub

Private Sub lvwMain_DblClick()
    cmdOK_Click
End Sub

Private Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mstrID = Item.Tag
    mstr编码 = Mid(Item.Text, 2, Len(Item.Text) - 2)
    mstr名称 = Item.SubItems(1)
End Sub
