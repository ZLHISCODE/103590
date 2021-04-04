VERSION 5.00
Begin VB.Form frmTechnicGroup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ִ�м����"
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
   StartUpPosition =   1  '����������
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
         Caption         =   "    ִ�м���飺 ����ִ�м��б���ѡ��ǰ�����Ӧ��ִ�м䣬��ִ�м������������ķ�����������齫�����ٴ�ѡ����ͬ��ִ�м䡣"
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
      Caption         =   "ȡ ��(&C)"
      Height          =   375
      Left            =   4050
      Picture         =   "frmTechnicGroup.frx":000C
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4605
      Width           =   1100
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "ȷ ��(&S)"
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
      ColNames        =   "|ִ�м�,rowcheck,w2800,key|ִ�м�ǰ׺>����ǰ׺,w1400,read|����ID,hide|"
      KeyName         =   "ִ�м�"
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
      Caption         =   "��������"
      Height          =   240
      Left            =   75
      TabIndex        =   6
      Top             =   4260
      Width           =   795
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "����ǰ׺"
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

Private mlngDeptId As Long  '��ǰ����Id
Private mlngGroupId As Long '��ǰ����ID
Private mstrGroupName As String
Private mstrPrefix As String

Private mblnIsModify As Boolean     '�Ƿ��޸ķ������ true-�޸ģ�false-���

Private mblnOK As Boolean    '�Ƿ�ȷ�Ϸ���


Public Function ShowGroupCfg(objOwner As Object, ByVal lngDeptID As Long, _
    ByRef lngGroupId As Long, ByRef strGroupName As String, ByRef strPrefix As String) As Boolean
'��ʾ��������
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    mlngDeptId = lngDeptID
    '����ID
    mlngGroupId = lngGroupId
    mblnIsModify = IIf(lngGroupId > 0, True, False)
    
    mblnOK = False
    ShowGroupCfg = False
    
    
    strSQL = "select a.ִ�м�,a.����ǰ׺,a.����Id from ҽ��ִ�з��� a where ����Id=[1] and (����ID=[2] or ����ID is null)"
    Set ufgRoomSelect.AdoData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ����ִ�м�", mlngDeptId, lngGroupId)
    

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
'������Ч�Լ��
    Dim lngMsgResult As Long
    
    CheckVerify = False
    
    If Trim(txtGrounName.Text) = "" Then
        Call MsgboxEx(Me, "�������Ʋ���Ϊ�գ���¼����Ч�ķ������ơ�", vbOKOnly, "��ʾ")
        txtGrounName.SetFocus
        Exit Function
    End If
    
    If Not ufgRoomSelect.IsCheckedRow Then
        Call MsgboxEx(Me, "��ѡ��÷���������Ӧ��ִ�м䡣", vbOKOnly, "��ʾ")
        ufgRoomSelect.SetFocus
        Exit Function
    End If
    
    If Trim(txtPrefix.Text) = "" Then
        lngMsgResult = MsgboxEx(Me, "��δ¼�����ǰ׺,�ź�ʱ����ͬ��֮����ܲ�����ͬ�ŶӺ��룬�Ƿ������", vbYesNo, "��ʾ")
        If lngMsgResult = vbNo Then
            txtPrefix.SetFocus
            Exit Function
        End If
    End If
    
    CheckVerify = True
End Function


Private Function GetSelectRoomName() As String
'��ȡ�Ѿ�ѡ���ִ�м�����
    Dim strRoomName As String
    Dim i As Long
    
    strRoomName = ""
    For i = 0 To ufgRoomSelect.GridRows - 1
        If ufgRoomSelect.GetRowCheck(i) Then
            If strRoomName <> "" Then strRoomName = strRoomName & ","
            strRoomName = strRoomName & ufgRoomSelect.Text(i, "ִ�м�")
        End If
    Next i
    
    GetSelectRoomName = strRoomName
End Function


Private Function NewGroup() As Boolean
'��������
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim strRoomName As String
    
    NewGroup = False
    
    '��ȡ��ǰ�����µ�ִ�м�����
    strRoomName = GetSelectRoomName
    
    strSQL = "select zl_Ӱ��ִ�з���_Add([1],[2],[3],[4]) as ����ID from dual"
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "����ִ�з���", txtGrounName.Text, txtPrefix.Text, mlngDeptId, strRoomName)
    If rsData.RecordCount <= 0 Then Exit Function
    
    mlngGroupId = Val(Nvl(rsData!����id))
    
    NewGroup = True
End Function


Private Function UpdateGroup() As Boolean
'���·���
    Dim strSQL As String
    Dim strRoomName As String
    
    UpdateGroup = False
    
    '��ȡ��ǰ�����µ�ִ�м�����
    strRoomName = GetSelectRoomName
    
    strSQL = "zl_Ӱ��ִ�з���_Update(" & mlngGroupId & ",'" & txtGrounName.Text & "','" & txtPrefix.Text & "','" & strRoomName & "'," & mlngDeptId & ")"
    
    Call zlDatabase.ExecuteProcedure(strSQL, "����ִ�з���")
    
    UpdateGroup = True
End Function

Private Sub cmdSure_Click()
'��������·���
On Error GoTo ErrHandle
        
    '��������Ƿ���Ч
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
'    '�������
'    InitDebugObject 1290, Me, "zlhis", "HIS"
'    mlngDeptID = 63
'
'    ShowGroupCfg Nothing, lngID, strName, strPrefix
'    '���Խ���
End Sub

Private Sub ufgRoomSelect_OnNewRow(ByVal Row As Long)
    If ufgRoomSelect.Text(Row, "����ID") <> "" Then Call ufgRoomSelect.SetRowCheck(Row, True)
End Sub

