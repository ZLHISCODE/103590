VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEPRModelsContent 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picContent 
      Appearance      =   0  'Flat
      BackColor       =   &H00E7CFBA&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4290
      Left            =   0
      ScaleHeight     =   4290
      ScaleWidth      =   5700
      TabIndex        =   1
      Top             =   2865
      Width           =   5700
      Begin VB.CheckBox chklevel 
         BackColor       =   &H00E7CFBA&
         Caption         =   "ȫԺͨ��"
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   9
         Top             =   525
         Width           =   1035
      End
      Begin VB.CheckBox chklevel 
         BackColor       =   &H00E7CFBA&
         Caption         =   "����ͨ��"
         Height          =   225
         Index           =   1
         Left            =   1275
         TabIndex        =   8
         Top             =   525
         Width           =   1035
      End
      Begin VB.CheckBox chklevel 
         BackColor       =   &H00E7CFBA&
         Caption         =   "����ʹ��"
         Height          =   225
         Index           =   2
         Left            =   2370
         TabIndex        =   7
         Top             =   525
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CommandButton cmdContent 
         Caption         =   "�� ��ӵ����İ���(&A)"
         Height          =   350
         Index           =   0
         Left            =   165
         TabIndex        =   6
         Top             =   30
         Width           =   2055
      End
      Begin VB.TextBox txtSeek 
         Height          =   270
         Left            =   4305
         TabIndex        =   4
         ToolTipText     =   "�����س��������Ʋ��ң���������붨λ��"
         Top             =   495
         Width           =   1170
      End
      Begin VB.CommandButton cmdContent 
         Caption         =   "�� �ӷ��İ���ɾ��(&D)"
         Height          =   350
         Index           =   1
         Left            =   3420
         TabIndex        =   3
         Top             =   45
         Width           =   2055
      End
      Begin MSComctlLib.ListView lvwModel 
         Height          =   3420
         Left            =   0
         TabIndex        =   2
         Top             =   825
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   6033
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E7CFBA&
         Caption         =   "���ƹ���"
         Height          =   165
         Left            =   3525
         TabIndex        =   5
         Top             =   540
         Width           =   780
      End
   End
   Begin MSComctlLib.ListView lvwModelContent 
      Height          =   2805
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   4948
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmEPRModelsContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mlngModelsID As Long
Private Sub Initlvw()
    With lvwModelContent.ColumnHeaders
        .Clear
        .Add , "_ID", "", 300
        .Add , "_���", "���", 800
        .Add , "_����", "����", 2000
        .Add , "_����", "����", 1600
        .Add , "_ͨ�ü�", "ͨ�ü�", 1000
        .Add , "_˵��", "˵��", 1800
        .Add , "_����", "����", 600
        .Add , "_����ID", "����ID", 0
        .Add , "_��ԱID", "��ԱID", 0
        .Add , "_����", "����", 800
        .Add , "_��Ա", "��Ա", 800
        .Add , "_�ļ�ID", "�ļ�ID", 0
    End With
    
    With lvwModel.ColumnHeaders
        .Clear
        .Add , "_ID", "", 300
        .Add , "_���", "���", 800
        .Add , "_����", "����", 2000
        .Add , "_����", "����", 1600
        .Add , "_ͨ�ü�", "ͨ�ü�", 1000
        .Add , "_˵��", "˵��", 1800
        .Add , "_����", "����", 600
        .Add , "_����ID", "����ID", 0
        .Add , "_��ԱID", "��ԱID", 0
        .Add , "_����", "����", 800
        .Add , "_��Ա", "��Ա", 800
    End With
End Sub
Private Sub cmdContent_Click(Index As Integer)
Dim arrSQL() As Variant, blnTran As Boolean, l As Integer, strTypes As String
    On Error GoTo ErrHandle
    arrSQL = Array()
    If Index = 0 Then '����
        For l = 1 To lvwModelContent.ListItems.Count 'ȡ����ѡ�����ļ����
            strTypes = strTypes & lvwModelContent.ListItems(l).SubItems(3) & "|"
        Next
    
        For l = 1 To lvwModel.ListItems.Count
            If lvwModel.ListItems(l).Checked Then
                If InStr(strTypes, lvwModel.ListItems(l).SubItems(3) & "|") > 0 Then 'ͬһ�ֲ���ֻ�ܼ���һ��,������ֶ��"��Ժ��¼"
                    MsgBox "        ͬһ�ֲ���ֻ�ܼ���һ�ݣ����飺" & vbCrLf & "�Ƿ�ѡ������ͬ������Ѿ�������ѡ������Ĳ����ļ���", vbInformation, gstrSysName: Exit Sub
                End If
                strTypes = strTypes & lvwModel.ListItems(l).SubItems(3) & "|"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "zl_�������İ����_Update(1," & mlngModelsID & "," & lvwModel.ListItems(l).Tag & ")"
            End If
        Next
    Else              'ɾ��
        For l = 1 To lvwModelContent.ListItems.Count
            If lvwModelContent.ListItems(l).Checked Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "zl_�������İ����_Update(0," & mlngModelsID & "," & lvwModelContent.ListItems(l).Tag & ")"
            End If
        Next
    End If
    
    gcnOracle.BeginTrans '--------------------------д������
    blnTran = True
    For l = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(l)), "д�����Ĳ�������")
    Next
    gcnOracle.CommitTrans: blnTran = False
    
    Call RefreshContent
    Call RefreshModel
    Exit Sub
ErrHandle:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Sub

Private Sub Form_Load()
    Initlvw
End Sub
Public Sub zlRefresh(ByVal lngModelsID As Long, ByVal strPrivs As String, ByVal bytType As Byte)
'lngModelsID��������ǰ���İ�ID bytType-0 ��ѯ bytType-1���ķ��İ����

    mstrPrivs = strPrivs: mlngModelsID = lngModelsID
    If picContent.Enabled Then bytType = 1
    If InStr(mstrPrivs, "�������İ�����") > 0 And bytType = 1 Then
        picContent.Enabled = True
        Call Form_Resize
        If InStr(mstrPrivs, "���˲�������") <= 0 Then chklevel(2).Enabled = False: chklevel(2).Value = False
        If InStr(mstrPrivs, "���Ҳ�������") <= 0 Then chklevel(1).Enabled = False: chklevel(1).Value = False
        If InStr(mstrPrivs, "ȫԺ��������") <= 0 Then chklevel(0).Enabled = False: chklevel(0).Value = False
        Call RefreshModel
    Else
        picContent.Enabled = False
        Call Form_Resize
    End If
    Call RefreshContent
End Sub
Private Sub RefreshContent()
Dim rsTemp As ADODB.Recordset, objItem As ListItem
    On Error GoTo ErrHandle
    gstrSQL = "select /*+ rule*/ A.ID,A.���,A.����,A.����,A.˵��,A.ͨ�ü�,A.����ID,A.��ԱID ,C.���� ���,D.���� ����,E.����,A.�ļ�ID" & _
                " from ��������Ŀ¼ A,�������İ���� B ,�����ļ��б� C,���ű� D,��Ա�� E" & _
                " where B.���İ�ID=[1] AND A.ID=B.����ID And nvl(A.����,0)=0 AND A.�ļ�ID=C.ID AND C.����=2 AND A.����ID=D.ID AND A.��ԱID=E.ID" & _
                " Order by C.����,C.���,A.ͨ�ü�,A.���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngModelsID)
    lvwModelContent.ListItems.Clear
    With rsTemp
        Do Until .EOF
            Set objItem = lvwModelContent.ListItems.Add(, "_" & !ID, "")
                objItem.Tag = !ID
                objItem.SubItems(1) = !���
                objItem.SubItems(2) = !����
                objItem.SubItems(3) = NVL(!���)
                objItem.SubItems(4) = Decode(NVL(!ͨ�ü�, 0), 0, "ȫԺͨ��", 1, "����ͨ��", 2, "����ʹ��")
                objItem.SubItems(5) = NVL(!˵��)
                objItem.SubItems(6) = NVL(!����)
                objItem.SubItems(7) = NVL(!����ID, 0)
                objItem.SubItems(8) = NVL(!��ԱID, 0)
                objItem.SubItems(9) = NVL(!����)
                objItem.SubItems(10) = NVL(!����)
                objItem.SubItems(11) = NVL(!�ļ�ID, 0)
                objItem.Checked = True
            .MoveNext
        Loop
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub RefreshModel()
Dim rsTemp As ADODB.Recordset, objItem As ListItem, lngID As Long, i As Integer, debarID As String
    On Error GoTo ErrHandle
    If lvwModel.ListItems.Count > 0 Then '���µ�ǰ��ѡ
        lngID = lvwModel.SelectedItem.Tag
    End If
    
'    For i = 1 To lvwModelContent.ListItems.Count
'        debarID = debarID & "," & lvwModelContent.ListItems(i).Tag
'    Next
'    If debarID <> "" Then debarID = Mid(debarID, 2)
    
    gstrSQL = ""
    If chklevel(0).Value = vbChecked Then gstrSQL = "A.ͨ�ü�=0" 'ȫԺͨ��
    If chklevel(1).Value = vbChecked Then gstrSQL = gstrSQL & IIf(gstrSQL = "", "", " or ") & "(A.ͨ�ü�=1 and A.����ID=[1])" '����ͨ��
    If chklevel(2).Value = vbChecked Then gstrSQL = gstrSQL & IIf(gstrSQL = "", "", " or ") & "(A.ͨ�ü�=2 and A.��ԱID=[2])" '����ʹ��
    If chklevel(0).Value = vbChecked And chklevel(1).Value = vbChecked And chklevel(2).Value = vbChecked Then gstrSQL = "" 'ȫѡ
        
    If gstrSQL = "" Then 'ȫѡʱ����Ȩ�޼�����
        If chklevel(0).Enabled Then gstrSQL = "A.ͨ�ü�=0"
        If chklevel(1).Enabled Then gstrSQL = gstrSQL & IIf(gstrSQL = "", "", " OR ") & "(A.ͨ�ü�=1 and A.����ID=[1])"
        If chklevel(2).Enabled Then gstrSQL = gstrSQL & IIf(gstrSQL = "", "", " OR ") & "(A.ͨ�ü�=2 and A.��ԱID=[2])"
    End If
    
    gstrSQL = "select /*+ rule*/ A.ID,A.���,A.����,A.����,A.˵��,A.ͨ�ü�,A.����ID,A.��ԱID,B.���� ���,C.���� ����,D.���� " & _
                " from ��������Ŀ¼ A,�����ļ��б� B ,���ű� C,��Ա�� D" & _
                " where A.�ļ�ID=B.ID AND B.����=2 and A.����ID=C.ID and A.��ԱID=D.ID AND nvl(A.����,0)=0" & IIf(gstrSQL = "", "", " and (" & gstrSQL & ")")
    If Trim(txtSeek.Text) <> "" Then
        gstrSQL = gstrSQL & " And " & zlCommFun.GetLike("A", "����", Trim(txtSeek))
    End If
'    If debarID <> "" Then
'        gstrSQL = gstrSQL & " And A.�ļ�ID not in(Select Distinct �ļ�ID from ��������Ŀ¼ where ID IN (" & debarID & "))"
'    End If
    gstrSQL = gstrSQL & " Order by A.ͨ�ü�,A.����,B.����,A.���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngDeptId, glngUserId)
    lvwModel.ListItems.Clear
    With rsTemp
        Do Until .EOF
            Set objItem = lvwModel.ListItems.Add(, "_" & !ID, "")
                objItem.Tag = !ID
                objItem.SubItems(1) = !���
                objItem.SubItems(2) = !����
                objItem.SubItems(3) = NVL(!���)
                objItem.SubItems(4) = Decode(NVL(!ͨ�ü�, 0), 0, "ȫԺͨ��", 1, "����ͨ��", 2, "����ʹ��")
                objItem.SubItems(5) = NVL(!˵��)
                objItem.SubItems(6) = NVL(!����)
                objItem.SubItems(7) = NVL(!����ID, 0)
                objItem.SubItems(8) = NVL(!��ԱID, 0)
                objItem.SubItems(9) = NVL(!����)
                objItem.SubItems(10) = NVL(!����)
            If !ID = lngID Then
                objItem.Selected = True
            End If
            .MoveNext
        Loop
    End With
    If lvwModel.ListItems.Count > 0 Then
        If lvwModel.SelectedItem Is Nothing Then
            lvwModel.ListItems(1).Selected = True
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    lvwModelContent.Width = Me.Width
    picContent.Width = Me.Width
    lvwModel.Width = Me.Width
    If picContent.Enabled Then
        picContent.Visible = True
        lvwModelContent.Height = 2800
        picContent.Height = Me.ScaleHeight - lvwModelContent.Height
        lvwModel.Height = picContent.Height - (chklevel(0).Top + chklevel(0).Height)
        picContent.Top = lvwModelContent.Height + lvwModelContent.Top
    Else
        picContent.Visible = False
        lvwModelContent.Top = 0
        lvwModelContent.Height = Me.ScaleHeight
    End If
    Err = 0: Err.Clear
End Sub

Private Sub chklevel_Click(Index As Integer)
Dim i As Integer, blnOnly As Boolean
    For i = 0 To chklevel.UBound
        If chklevel(i).Enabled Then
            If chklevel(i).Value = vbChecked Then
                blnOnly = True: Exit For 'ֻҪ�б�ѡ�м��˳�
            End If
        End If
    Next
    
    If blnOnly = False Then chklevel(Index).Value = vbChecked '��֤ʼ����һ����ѡ��
    Call RefreshModel
End Sub

Private Sub lvwModel_Click()
Dim i As Integer
    For i = 1 To lvwModel.ListItems.Count
        If lvwModel.ListItems(i).Checked = True Then cmdContent(0).Enabled = True: Exit Sub
    Next
    cmdContent(0).Enabled = False
End Sub

Private Sub lvwModelContent_Click()
Dim i As Integer
    For i = 1 To lvwModelContent.ListItems.Count
        If lvwModelContent.ListItems(i).Checked = True Then cmdContent(1).Enabled = True: Exit Sub
    Next
    cmdContent(1).Enabled = False
End Sub


Private Sub txtSeek_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call RefreshModel
    ElseIf InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyAscii))) > 0 Then '���붨λ
        Dim i As Integer, strtmp As String
        If txtSeek.SelLength > 0 Then
            strtmp = ""
        Else
            strtmp = txtSeek.Text
        End If
        For i = 1 To lvwModel.ListItems.Count
            If UCase(lvwModel.ListItems(i).SubItems(6)) Like UCase(Trim(strtmp)) & UCase(Chr(KeyAscii)) & "*" Then
                lvwModel.SelectedItem.Selected = False: lvwModel.ListItems(i).Selected = True: Exit Sub
            End If
        Next
    End If
End Sub
