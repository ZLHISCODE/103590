VERSION 5.00
Begin VB.Form frmMedicalTeamEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ��С��༭"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   Icon            =   "frmMedicalTeamEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.ComboBox cboDept 
      Height          =   300
      ItemData        =   "frmMedicalTeamEdit.frx":000C
      Left            =   1080
      List            =   "frmMedicalTeamEdit.frx":000E
      TabIndex        =   1
      Text            =   "cboDept"
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1920
      TabIndex        =   6
      Top             =   1815
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3120
      TabIndex        =   7
      Top             =   1800
      Width           =   1100
   End
   Begin VB.TextBox txtExplain 
      Height          =   300
      Left            =   1080
      MaxLength       =   200
      TabIndex        =   5
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1080
      MaxLength       =   48
      TabIndex        =   3
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label lblExplain 
      AutoSize        =   -1  'True
      Caption         =   "˵��(&E)"
      Height          =   180
      Left            =   360
      TabIndex        =   4
      Top             =   1250
      Width           =   630
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "����(&N)"
      Height          =   180
      Left            =   360
      TabIndex        =   2
      Top             =   770
      Width           =   630
   End
   Begin VB.Label lblDept 
      AutoSize        =   -1  'True
      Caption         =   "����(&D)"
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   290
      Width           =   630
   End
End
Attribute VB_Name = "frmMedicalTeamEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngTeamID As Long          'ҽ��С��ID
Private mbytStatus As Byte          '����״̬
Private mstrMatching As String      '����ƥ��
Public mstrPrivs As String          'Ȩ��

Property Get Status() As Byte
'����״̬��1-���; 2-�༭
    Status = mbytStatus
End Property
Property Let Status(ByVal bytStatus As Byte)
    If bytStatus = 1 Then
        Me.Caption = "ҽ��С������"
    Else
        Me.Caption = "ҽ��С���޸�"
    End If
    mbytStatus = bytStatus
End Property

Property Get TeamID() As Long
    TeamID = mlngTeamID
End Property
Property Let TeamID(ByVal lngTeamID As Long)
    mlngTeamID = lngTeamID
End Property

Private Sub cboDept_Change()
    cboDept.Tag = ""
End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
    If InStr("~!@#$%^&*()_+|=-`;'"":/.,<>?[]{}", Chr(KeyAscii)) > 0 Then KeyAscii = 0

    If KeyAscii <> vbKeyReturn Then Exit Sub
    If cboDept.ListCount = 0 Then Exit Sub
    If cboDept.Tag <> "" Then Exit Sub
    
    Dim rsTmp As ADODB.Recordset
    Dim strNO As String, strReturn As String
    
    On Error GoTo ErrHandle
    With frmSelCur
        strNO = mstrMatching & UCase(Trim(cboDept.Text)) & "%"
        If InStr(mstrPrivs, "���п���") = 0 Then
            gstrSQL = "Select distinct b.����,b.id From ��������˵�� a, ���ű� b, ������Ա c " & _
                      "Where a.����id=b.Id and b.id=c.����id And a.��������='�ٴ�' and a.������� in (2,3) " & _
                      "  and (b.���� like [1] or b.���� like [1] or b.���� like [1]) and c.��ԱID=[2] " & _
                      "  and b.����ʱ�� = To_Date('3000-1-1', 'yyyy-mm-dd') And Substr(b.����, 1, 1) <> '-' "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNO, glngUserId)
        Else
'            gstrSQL = "Select distinct b.����,b.id From ��������˵�� a, ���ű� b, ������Ա c " & _
'                      "Where a.����id=b.Id and b.id=c.����id And a.��������='�ٴ�' and a.������� in (2,3) " & _
'                      "  and (b.���� like [1] or b.���� like [1] or b.���� like [1]) " & _
'                      "  and b.����ʱ�� = To_Date('3000-1-1', 'yyyy-mm-dd') And Substr(b.����, 1, 1) <> '-' "
            gstrSQL = "Select distinct b.����,b.id From ��������˵�� a, ���ű� b " & _
                      "Where a.����id=b.Id And a.��������='�ٴ�' and a.������� in (2,3) " & _
                      "  and (b.���� like [1] or b.���� like [1] or b.���� like [1]) " & _
                      "  and b.����ʱ�� = To_Date('3000-1-1', 'yyyy-mm-dd') And Substr(b.����, 1, 1) <> '-' "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNO)
        End If
        cboDept.Tag = ""
        If rsTmp.RecordCount > 0 Then
            If rsTmp.RecordCount = 1 Then
                strReturn = rsTmp!���� & "," & rsTmp!ID
            Else
                strReturn = .ShowCurrSel(Me, rsTmp, "����,2500,0,2;id,0,0,2", "���ѡ����", False, , 0)
            End If
            If Trim(strReturn) <> "" Then
                cboDept.Text = Left(strReturn, InStr(strReturn, ",") - 1)
                cboDept.Tag = Mid(strReturn, InStr(strReturn, ",") + 1)
            End If
        Else
            MsgBox "���κο��õĿ��ң�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
        End If
    End With
    Exit Sub

ErrHandle:
    Call ErrCenter
    Call SaveErrLog

End Sub

Private Sub cboDept_LostFocus()
    Dim i As Long
    If cboDept.ListCount = 0 Then Exit Sub
    If cboDept.ListIndex < 0 Then
        For i = 0 To cboDept.ListCount - 1
            If Val(cboDept.Tag) = cboDept.ItemData(i) Then
                cboDept.ListIndex = i: Exit For
            End If
        Next
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ErrHandle
    '���
    If cboDept.ListIndex < 0 Then
        MsgBox "[����]¼����Ϣ����ȷ��", vbInformation, gstrSysName
        cboDept.SetFocus
        Exit Sub
    End If
    If cboDept.ItemData(cboDept.ListIndex) = 0 Then
        MsgBox "[����]δ¼����Ϣ��", vbInformation, gstrSysName
        cboDept.SetFocus
        Exit Sub
    End If
    If Trim(txtName.Text) = "" Then
        MsgBox "[����]δ¼����Ϣ��", vbInformation, gstrSysName
        txtName.SetFocus
        Exit Sub
    End If
    If Left(txtName.Text, 1) = "-" Then
        MsgBox "[����]���ַ��������ǡ�-����", vbInformation, gstrSysName
        txtName.SetFocus
        Exit Sub
    End If
    If LenB(StrConv(txtName.Text, vbFromUnicode)) > txtName.MaxLength Then
        MsgBox "[����]¼�����ݲ��ܳ�29�����ֻ�48�ַ���", vbInformation, gstrSysName
        txtName.SetFocus
        Exit Sub
    End If
    If zlcommfun.StrIsValid(txtName.Text) = False Then Exit Sub
    If LenB(StrConv(txtExplain.Text, vbFromUnicode)) > txtName.MaxLength Then
        MsgBox "[˵��]¼�����ݲ��ܳ�100�����ֻ�200�ַ���", vbInformation, gstrSysName
        txtExplain.SetFocus
        Exit Sub
    End If
    If zlcommfun.StrIsValid(txtExplain.Text) = False Then Exit Sub
   
    If Me.Status = 1 Then
        gstrSQL = "ZL_�ٴ�ҽ��С��_INSERT(" & _
                  cboDept.ItemData(cboDept.ListIndex) & ",'" & _
                  Trim(txtName.Text) & "'," & _
                  IIF(Trim(txtExplain.Text) = "", "null", "'" & Trim(txtExplain.Text) & "'") & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        txtName.Text = ""
        txtExplain.Text = ""
        txtName.SetFocus
    Else
        gstrSQL = "ZL_�ٴ�ҽ��С��_UPDATE(" & _
                  TeamID & "," & _
                  cboDept.ItemData(cboDept.ListIndex) & ",'" & _
                  Trim(txtName.Text) & "'," & _
                  IIF(Trim(txtExplain.Text) = "", "null", "'" & Trim(txtExplain.Text) & "'") & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Unload Me
    End If
    Exit Sub
    
ErrHandle:
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlcommfun.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    mstrMatching = IIF(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
    
    If InStr(mstrPrivs, "���п���") = 0 Then
        gstrSQL = "Select Distinct a.����id,b.���� From ��������˵�� a, ���ű� b, ������Ա c " & _
                  "Where a.����id=b.Id and b.id=c.����id and c.��Աid=[1] And a.��������='�ٴ�' and a.������� in (2,3) " & _
                  "  And b.����ʱ�� = To_Date('3000-1-1', 'yyyy-mm-dd') And Substr(b.����, 1, 1) <> '-' " & _
                  "Order By b.���� "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngUserId)
    Else
        gstrSQL = "Select Distinct a.����id,b.���� From ��������˵�� a, ���ű� b Where a.����id=b.Id " & _
                  "  And a.��������='�ٴ�' and a.������� in (2,3) " & _
                  "  And b.����ʱ�� = To_Date('3000-1-1', 'yyyy-mm-dd') And Substr(b.����, 1, 1) <> '-' " & _
                  "Order By b.���� "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    End If
    
    With cboDept
        .Clear
        For i = 0 To rsTemp.RecordCount - 1
            .AddItem Nvl(rsTemp!����)
            .ItemData(i) = Nvl(rsTemp!����id)
            rsTemp.MoveNext
        Next
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtExplain_KeyPress(KeyAscii As Integer)
    If InStr("~!@#$%^&*()_+|=-`;'"":/.,<>?[]{}", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtName_Change()
    If Left(txtName.Text, 1) = "-" Then
        txtName.Text = "?" & Mid(txtName.Text, 2)
        MsgBox "[����]���ַ��������ǡ�-����", vbInformation, gstrSysName
    End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If InStr("~!@#$%^&*()_+|=-`;'"":/.,<>?[]{}", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
