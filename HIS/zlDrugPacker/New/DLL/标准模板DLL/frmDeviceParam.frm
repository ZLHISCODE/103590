VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDeviceParam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设备应用参数信息"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5970
   Icon            =   "frmDeviceParam.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   5970
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboDevice 
      Height          =   300
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   780
      Width           =   4935
   End
   Begin VB.Frame fraLine1 
      Height          =   75
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   5900
   End
   Begin TabDlg.SSTab sstDeviceParam 
      Height          =   5175
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "应用场合(&0)"
      TabPicture(0)   =   "frmDeviceParam.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDevice(10)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Lvw药品剂型"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin MSComctlLib.ListView Lvw药品剂型 
         Height          =   4065
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   7170
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgLvwSel"
         SmallIcons      =   "imgLvwSel"
         ColHdrIcons     =   "imgLvwSel"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "名称"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "剂型选择"
         Height          =   180
         Index           =   10
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   720
      End
   End
   Begin VB.CommandButton CmdExit 
      Cancel          =   -1  'True
      Caption         =   "退出(&E)"
      Height          =   350
      Left            =   4680
      TabIndex        =   1
      Top             =   6600
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   3480
      TabIndex        =   0
      Top             =   6600
      Width           =   1100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "设备"
      Height          =   180
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmDeviceParam.frx":0326
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblComment 
      Caption         =   "设置发药设备的使用环境和应用参数！"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "frmDeviceParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng设备ID As Long
Private mlng药房ID As Long

Public Sub ShowMeByDevice(ByVal frmOwner As Form, ByVal lng设备id As Long)
    mlng设备ID = lng设备id
    
    Call GetDevice(0, lng设备id)
    
    Me.Show vbModal, frmOwner
    
    Exit Sub
End Sub

Public Sub ShowMeByStock(ByVal frmOwner As Form, ByVal lng药房ID As Long)
    mlng药房ID = lng药房ID
    
    Call GetDevice(1, mlng药房ID)
    
    Me.Show vbModal, frmOwner
    
    Exit Sub
End Sub

Private Sub IniData(ByVal lng药房ID As Long)
    Dim rsData As ADODB.Recordset
    Dim byt性质 As Byte     '1-门诊,2-住院,3-门诊和住院
    
    On Error GoTo errHandle
    
    '服务对象
    gstrSQL = "Select 服务对象 From 部门性质说明 " & _
                  "Where 部门id = [1] And 服务对象 in (1,2,3) " & _
                  "Order By 服务对象 "
    Set rsData = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "IniData", lng药房ID)
    Do While rsData.EOF = False
        If byt性质 = 0 Then
            byt性质 = NVL(rsData!服务对象, 0)
        ElseIf byt性质 = 3 Then
            Exit Do
        Else
            Select Case NVL(rsData!服务对象, 0)
                Case 1  '门诊
                    If byt性质 = 2 Then
                        byt性质 = 3
                        Exit Do
                    End If
                Case 2  '住院
                    If byt性质 = 1 Then
                        byt性质 = 3
                        Exit Do
                    End If
                Case 3  '门诊和住院
                    byt性质 = 3
                    Exit Do
            End Select
        End If
        
        rsData.MoveNext
    Loop
    
'    If byt性质 = 1 Then
'        optObject(0).Value = True
'        optObject(0).Enabled = True
'        optObject(1).Value = False
'        optObject(1).Enabled = False
'     ElseIf byt性质 = 2 Then
'        optObject(0).Value = False
'        optObject(0).Enabled = False
'        optObject(1).Value = True
'        optObject(1).Enabled = True
'     ElseIf byt性质 = 3 Then
'        optObject(0).Value = True
'        optObject(0).Enabled = True
'        optObject(1).Value = True
'        optObject(1).Enabled = True
'    End If
        
'    '业务规则
'    With cboDispense
'        .Clear
'        .AddItem "门诊收费", 0
'        .AddItem "处方发药配药功能", 1
'        .AddItem "处方发药发药功能", 2
'    End With
'
'    With cboSend
'        .Clear
'        .AddItem "无响应", 0
'        .AddItem "药品处方发药功能", 1
'    End With
    
'    '上传规则
'    chkBillType(0).Value = 1
'    chkBillType(1).Value = 1
'    chkBillType(2).Value = 1
    
    gstrSQL = "Select Distinct J.编码||'-'||J.名称 剂型" & _
         " From 诊疗执行科室 A,药品特性 B,药品剂型 J " & _
         " Where A.诊疗项目ID=B.药名ID And B.药品剂型=J.名称" & _
         " And A.执行科室ID=[1]" & _
         " Order By j.编码 || '-' || j.名称 "
    Set rsData = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "IniData", lng药房ID)
    
    With Lvw药品剂型
        .ListItems.Clear
        .ListItems.Add , "_" & .ListItems.Count + 1, "所有药品剂型" ', 1, 1
        .ListItems(.ListItems.Count).Checked = True
        Do While Not rsData.EOF
            .ListItems.Add , "_" & .ListItems.Count + 1, Mid(rsData!剂型, InStr(1, rsData!剂型, "-") + 1) ', 1, 1
            .ListItems(.ListItems.Count).Checked = True
            rsData.MoveNext
        Loop
    End With
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub

Private Sub GetDeviceParam(ByVal lng设备id As Long)
    Dim rsData As ADODB.Recordset
    Dim n As Integer
    
    On Error GoTo errHandle
    
    gstrSQL = "Select a.参数id, a.设备id, Nvl(a.参数值, b.缺省值) As 参数值, b.参数号, b.参数名, b.参数说明 " & vbNewLine & _
        " From 药房设备参数 A, 自动发药参数 B" & vbNewLine & _
        " Where a.参数id(+) = b.Id and a.设备id(+)=[1] "

    Set rsData = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetDeviceParam", lng设备id)
    
    Do While Not rsData.EOF
        Select Case rsData!参数名
'            Case "服务对象"
'                If rsData!参数值 = 1 Then
'                    optObject(0).Value = True
'                Else
'                    optObject(1).Value = True
'                End If
'            Case "预配药响应"
'                cboDispense.ListIndex = Val(rsData!参数值) - 1
'            Case "发送响应"
'                cboSend.ListIndex = Val(rsData!参数值)
'            Case "单据类型"
'                chkBillType(0).Value = Val(Mid(rsData!参数值, 1, 1))
'                chkBillType(1).Value = Val(Mid(rsData!参数值, 2, 1))
'                chkBillType(2).Value = Val(Mid(rsData!参数值, 3, 1))
            Case "药品剂型"
                With Lvw药品剂型
                    If .ListItems.Count = 0 Then
                        Exit Sub
                    End If
                    
                    For n = 1 To .ListItems.Count
                        .ListItems(n).Checked = False
                        If NVL(rsData!参数值) = "所有" Then
                            .ListItems(n).Checked = True
                        Else
                            If InStr(1, "," & NVL(rsData!参数值) & ",", "," & .ListItems(n).Text & ",") > 0 Then
                                .ListItems(n).Checked = True
                            End If
                        End If
                    Next
                End With
        End Select
                
        rsData.MoveNext
    Loop
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub


Private Sub cboDevice_Click()
'    If mlng设备ID <> Val(Split(cboDevice.ItemData(cboDevice.ListIndex), "|")(0)) And mlng药房ID <> Val(Split(cboDevice.ItemData(cboDevice.ListIndex), "|")(1)) Then
'        mlng设备ID = Val(Split(cboDevice.ItemData(cboDevice.ListIndex), "|")(0))
'        mlng药房ID = Val(Split(cboDevice.ItemData(cboDevice.ListIndex), "|")(1))
'
'        Call IniData(mlng药房ID)
'        DoEvents
'        Call GetDeviceParam(mlng设备ID)
'    End If

    If mlng设备ID <> cboDevice.ItemData(cboDevice.ListIndex) Then
        mlng设备ID = cboDevice.ItemData(cboDevice.ListIndex)
        Call IniData(mlng药房ID)
        DoEvents
        Call GetDeviceParam(mlng设备ID)
    End If
    
End Sub


Private Sub cmdSave_Click()
    Dim str剂型 As String
    Dim n As Integer
    
    On Error GoTo errHandle
    
    gobjConn.BeginTrans
    
    '按参数号分别保存参数
'    '服务对象
'    gstrSQL = "Zl_药房设备参数_Update("
'    gstrSQL = gstrSQL & 1 & ","
'    gstrSQL = gstrSQL & mlng设备ID & ","
'    gstrSQL = gstrSQL & IIf(optObject(0).Value = True, 1, 2)
'    gstrSQL = gstrSQL & ")"
'    Call gobjComLib.zldatabase.ExecuteProcedure(gstrSQL, "CmdSave_Click")
'
'    '配药功能
'    gstrSQL = "Zl_药房设备参数_Update("
'    gstrSQL = gstrSQL & 2 & ","
'    gstrSQL = gstrSQL & mlng设备ID & ","
'    gstrSQL = gstrSQL & cboDispense.ListIndex + 1
'    gstrSQL = gstrSQL & ")"
'    Call gobjComLib.zldatabase.ExecuteProcedure(gstrSQL, "CmdSave_Click")
'
'    '发送功能
'    gstrSQL = "Zl_药房设备参数_Update("
'    gstrSQL = gstrSQL & 3 & ","
'    gstrSQL = gstrSQL & mlng设备ID & ","
'    gstrSQL = gstrSQL & cboSend.ListIndex
'    gstrSQL = gstrSQL & ")"
'    Call gobjComLib.zldatabase.ExecuteProcedure(gstrSQL, "CmdSave_Click")
'
'
'    '单据类型
'    gstrSQL = "Zl_药房设备参数_Update("
'    gstrSQL = gstrSQL & 4 & ","
'    gstrSQL = gstrSQL & mlng设备ID & ","
'    gstrSQL = gstrSQL & chkBillType(0).Value & chkBillType(1).Value & chkBillType(2).Value
'    gstrSQL = gstrSQL & ")"
'    Call gobjComLib.zldatabase.ExecuteProcedure(gstrSQL, "CmdSave_Click")
     
    '剂型
    If Lvw药品剂型.ListItems(1).Checked Then
        str剂型 = "所有"
    Else
        For n = 1 To Lvw药品剂型.ListItems.Count
            If Lvw药品剂型.ListItems(n).Checked Then
                str剂型 = IIf(str剂型 = "", "", str剂型 & ",") & Lvw药品剂型.ListItems(n).Text
            End If
        Next
    End If
    gstrSQL = "Zl_药房设备参数_Update("
    gstrSQL = gstrSQL & 1 & ","
    gstrSQL = gstrSQL & mlng设备ID & ","
    gstrSQL = gstrSQL & IIf(str剂型 = "所有", "null", "'" & str剂型 & "'")
    gstrSQL = gstrSQL & ")"
    Call gobjComLib.zldatabase.ExecuteProcedure(gstrSQL, "cmdSave_Click")
    
    gobjConn.CommitTrans
    
    MsgBox "参数已保存！", vbInformation, GSTR_INTERFACE_NAME
    
    Exit Sub
errHandle:
    gobjConn.RollbackTrans
    gobjComLib.ErrCenter
    gstrMessage = Err.Description
End Sub

Private Sub Lvw药品剂型_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim n As Integer
    Dim blnAllChecked As Boolean
    
    With Lvw药品剂型
        For n = 1 To .ListItems.Count
            .ListItems(n).Selected = False
        Next
        Item.Selected = True
        If Item.Text = "所有药品剂型" Then
            If Item.Checked Then
                blnAllChecked = True
            End If
                
            For n = 1 To .ListItems.Count
                .ListItems(n).Checked = blnAllChecked
            Next
        Else
            If Item.Checked = False Then
                .ListItems(1).Checked = False
            End If
        End If
    End With
End Sub

Private Sub GetDevice(ByVal bytType As Byte, ByVal lng标识id As Long)
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
     
    If bytType = 0 Then
        '按设备ID取设备信息
        gstrSQL = "Select a.Id As 设备id, a.使用部门id As 药房id, '【' || a.编码 || '】' || a.名称 || '(' || a.型号 || ')' || ' - ' || b.名称 As 名称 " & _
            " From 药房发药设备 A, 部门表 B " & _
            " Where a.使用部门id = b.Id And a.ID = [1] "
    Else
        '按药房ID取设备信息
        gstrSQL = "Select a.Id As 设备id, a.使用部门id As 药房id, '【' || a.编码 || '】' || a.名称 || '(' || a.型号 || ')' || ' - ' || b.名称 As 名称 " & _
            " From 药房发药设备 A, 部门表 B " & _
            " Where a.使用部门id = b.Id And a.使用部门id = [1] "
    End If
    Set rsData = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetDevice", lng标识id)
    
    cboDevice.Clear
    Do While rsData.EOF = False
        cboDevice.AddItem rsData!名称
        cboDevice.ItemData(cboDevice.NewIndex) = rsData!设备ID '"" & rsData!设备ID & "|" & rsData!药房id
        rsData.MoveNext
    Loop
    
    If bytType = 1 And cboDevice.ListCount > 0 Then
        cboDevice.ListIndex = 0
    End If
    
    Exit Sub
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub
