VERSION 5.00
Begin VB.Form frmTechnicGroup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "执行间分组"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5220
   Icon            =   "frmTechnicGroup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   5220
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1065
      Left            =   0
      ScaleHeight     =   1065
      ScaleWidth      =   5220
      TabIndex        =   7
      Top             =   0
      Width           =   5220
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "    执行间分组： 请在执行间列表中选择当前分组对应的执行间，当执行间配置完所属的分组后，其他分组将不能再次选择相同的执行间。"
         Height          =   600
         Left            =   225
         TabIndex        =   8
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.TextBox txtGrounName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   870
      TabIndex        =   3
      Top             =   4200
      Width           =   1680
   End
   Begin VB.TextBox txtPrefix 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3510
      TabIndex        =   2
      Top             =   4185
      Width           =   1620
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取 消(&C)"
      Height          =   375
      Left            =   4050
      Picture         =   "frmTechnicGroup.frx":000C
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4605
      Width           =   1100
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "确 定(&S)"
      Height          =   375
      Left            =   2955
      Picture         =   "frmTechnicGroup.frx":0156
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4605
      Width           =   1100
   End
   Begin zl9PACSWork.ucFlexGrid ufgRoomSelect 
      Height          =   2970
      Left            =   60
      TabIndex        =   4
      Top             =   1110
      Width           =   5070
      _ExtentX        =   8943
      _ExtentY        =   5239
      DefaultCols     =   ""
      ColNames        =   "|执行间,rowcheck,w2800,key|执行间前缀>号码前缀,w1400,read|分组ID,hide|"
      KeyName         =   "执行间"
      DisCellColor    =   16777215
      IsCopyAdoMode   =   0   'False
      IsEjectConfig   =   -1  'True
      IsShowPopupMenu =   0   'False
      HeadFontCharset =   134
      HeadFontWeight  =   400
      HeadColor       =   0
      DataFontCharset =   134
      DataFontWeight  =   400
      DataColor       =   0
      RowHeightMin    =   260
      ExtendLastCol   =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "分组名称"
      Height          =   240
      Left            =   75
      TabIndex        =   6
      Top             =   4260
      Width           =   795
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "分组前缀"
      Height          =   240
      Left            =   2700
      TabIndex        =   5
      Top             =   4245
      Width           =   840
   End
End
Attribute VB_Name = "frmTechnicGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngDeptId As Long  '当前科室Id
Private mlngGroupId As Long '当前分组ID
Private mstrGroupName As String
Private mstrPrefix As String

Private mblnIsModify As Boolean     '是否修改分组操作 true-修改，false-添加

Private mblnOK As Boolean    '是否确认分组


Public Function ShowGroupCfg(objOwner As Object, ByVal lngDeptID As Long, _
    ByRef lngGroupId As Long, ByRef strGroupName As String, ByRef strPrefix As String) As Boolean
'显示分组配置
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    mlngDeptId = lngDeptID
    '分组ID
    mlngGroupId = lngGroupId
    mblnIsModify = IIf(lngGroupId > 0, True, False)
    
    mblnOK = False
    ShowGroupCfg = False
    
    
    strSQL = "select a.执行间,a.号码前缀,a.分组Id from 医技执行房间 a where 科室Id=[1] and (分组ID=[2] or 分组ID is null)"
    Set ufgRoomSelect.AdoData = zlDatabase.OpenSQLRecord(strSQL, "查询分组执行间", mlngDeptId, lngGroupId)
    

    ufgRoomSelect.GridRows = ufgRoomSelect.AdoData.RecordCount + 1
    Call ufgRoomSelect.RefreshData
    
    txtGrounName.Text = strGroupName
    txtPrefix.Text = strPrefix
    
    
    Me.Show 1, objOwner
    
    lngGroupId = mlngGroupId
    strGroupName = mstrGroupName
    strPrefix = mstrPrefix
    
    ShowGroupCfg = mblnOK
End Function

Private Function CheckVerify() As Boolean
'数据有效性检查
    Dim lngMsgResult As Long
    
    CheckVerify = False
    
    If Trim(txtGrounName.Text) = "" Then
        Call MsgboxEx(Me, "分组名称不能为空，请录入有效的分组名称。", vbOKOnly, "提示")
        txtGrounName.SetFocus
        Exit Function
    End If
    
    If Not ufgRoomSelect.IsCheckedRow Then
        Call MsgboxEx(Me, "请选择该分组下所对应的执行间。", vbOKOnly, "提示")
        ufgRoomSelect.SetFocus
        Exit Function
    End If
    
    If Trim(txtPrefix.Text) = "" Then
        lngMsgResult = MsgboxEx(Me, "尚未录入分组前缀,排号时，不同组之间可能产生相同排队号码，是否继续？", vbYesNo, "提示")
        If lngMsgResult = vbNo Then
            txtPrefix.SetFocus
            Exit Function
        End If
    End If
    
    CheckVerify = True
End Function


Private Function GetSelectRoomName() As String
'获取已经选择的执行间名称
    Dim strRoomName As String
    Dim i As Long
    
    strRoomName = ""
    For i = 0 To ufgRoomSelect.GridRows - 1
        If ufgRoomSelect.GetRowCheck(i) Then
            If strRoomName <> "" Then strRoomName = strRoomName & ","
            strRoomName = strRoomName & ufgRoomSelect.Text(i, "执行间")
        End If
    Next i
    
    GetSelectRoomName = strRoomName
End Function


Private Function NewGroup() As Boolean
'新增分组
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim strRoomName As String
    
    NewGroup = False
    
    '获取当前分组下的执行间名称
    strRoomName = GetSelectRoomName
    
    strSQL = "select zl_影像执行分组_Add([1],[2],[3],[4]) as 分组ID from dual"
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "新增执行分组", txtGrounName.Text, txtPrefix.Text, mlngDeptId, strRoomName)
    If rsData.RecordCount <= 0 Then Exit Function
    
    mlngGroupId = Val(Nvl(rsData!分组id))
    
    NewGroup = True
End Function


Private Function UpdateGroup() As Boolean
'更新分组
    Dim strSQL As String
    Dim strRoomName As String
    
    UpdateGroup = False
    
    '获取当前分组下的执行间名称
    strRoomName = GetSelectRoomName
    
    strSQL = "zl_影像执行分组_Update(" & mlngGroupId & ",'" & txtGrounName.Text & "','" & txtPrefix.Text & "','" & strRoomName & "'," & mlngDeptId & ")"
    
    Call zlDatabase.ExecuteProcedure(strSQL, "更新执行分组")
    
    UpdateGroup = True
End Function

Private Sub cmdSure_Click()
'新增或更新分组
On Error GoTo ErrHandle
        
    '检查数据是否有效
    If Not CheckVerify() Then
        Exit Sub
    End If
    
    If mblnIsModify Then
        mblnOK = UpdateGroup
    Else
        mblnOK = NewGroup
    End If
    
    If mblnOK Then
        mstrGroupName = txtGrounName.Text
        mstrPrefix = txtPrefix.Text
    End If
    
    Unload Me
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub Form_Load()
'    Dim lngID As Long, strName As String, strPrefix As String
'    '调试语句
'    InitDebugObject 1290, Me, "zlhis", "HIS"
'    mlngDeptID = 63
'
'    ShowGroupCfg Nothing, lngID, strName, strPrefix
'    '调试结束
End Sub

Private Sub ufgRoomSelect_OnNewRow(ByVal Row As Long)
    If ufgRoomSelect.Text(Row, "分组ID") <> "" Then Call ufgRoomSelect.SetRowCheck(Row, True)
End Sub

