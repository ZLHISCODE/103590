VERSION 5.00
Begin VB.Form frmUnitEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Լ��λ����"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   Icon            =   "frmUnitEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboInfo 
      Height          =   300
      Left            =   3480
      TabIndex        =   26
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   5100
      TabIndex        =   25
      Top             =   1380
      Width           =   1100
   End
   Begin VB.TextBox txtEdit 
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   1
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   1
      Tag             =   "����"
      Text            =   "111111"
      Top             =   195
      Width           =   885
   End
   Begin VB.CommandButton cmd�ϼ� 
      Caption         =   "��"
      Height          =   270
      Left            =   4510
      TabIndex        =   21
      Top             =   3120
      Width           =   285
   End
   Begin VB.Frame Frame1 
      Height          =   4485
      Left            =   4920
      TabIndex        =   23
      Top             =   -150
      Width           =   30
   End
   Begin VB.CheckBox chkĩ�� 
      Caption         =   "ĩ��(&M)"
      Height          =   225
      Left            =   480
      TabIndex        =   22
      Top             =   3510
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   8
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   15
      Tag             =   "��ϵ��"
      Top             =   2760
      Width           =   1500
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   7
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   13
      Tag             =   "�ʺ�"
      Top             =   2370
      Width           =   2400
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   6
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   11
      Tag             =   "��������"
      Top             =   1980
      Width           =   3620
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   5
      Left            =   1200
      MaxLength       =   16
      TabIndex        =   9
      Tag             =   "�绰"
      Top             =   1590
      Width           =   2400
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   4
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   7
      Tag             =   "��ַ"
      Top             =   1200
      Width           =   3620
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   2
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   3
      Tag             =   "����"
      Top             =   510
      Width           =   3620
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5100
      TabIndex        =   18
      Top             =   660
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5100
      TabIndex        =   17
      Top             =   150
      Width           =   1100
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   10
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   20
      Top             =   3120
      Width           =   3310
   End
   Begin VB.TextBox txtTemp 
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   24
      TabStop         =   0   'False
      Tag             =   "����"
      Text            =   "11"
      Top             =   150
      Width           =   1155
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   3
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   5
      Tag             =   "����"
      Top             =   870
      Width           =   1155
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "Ժ��(&P)"
      Height          =   180
      Index           =   9
      Left            =   2805
      TabIndex        =   16
      Top             =   2820
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "��ϵ��(&L)"
      Height          =   180
      Index           =   8
      Left            =   300
      TabIndex        =   14
      Top             =   2820
      Width           =   810
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "�ʺ�(&Z)"
      Height          =   180
      Index           =   7
      Left            =   480
      TabIndex        =   12
      Top             =   2430
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "��������(&B)"
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "�绰(&T)"
      Height          =   180
      Index           =   5
      Left            =   480
      TabIndex        =   8
      Top             =   1680
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "��ַ(&A)"
      Height          =   180
      Index           =   4
      Left            =   480
      TabIndex        =   6
      Top             =   1320
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����(&U)"
      Height          =   180
      Index           =   1
      Left            =   480
      TabIndex        =   0
      Top             =   210
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����(&N)"
      Height          =   180
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   570
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����(&S)"
      Height          =   180
      Index           =   3
      Left            =   480
      TabIndex        =   4
      Top             =   930
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "�ϼ�(&H)"
      Height          =   180
      Index           =   10
      Left            =   480
      TabIndex        =   19
      Top             =   3180
      Width           =   630
   End
End
Attribute VB_Name = "frmUnitEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Dim mstr�ϼ���λID As String     '��ǰ�༭���ϼ���λID
Dim mstrID As String         '��ǰ�༭�ĵ�λID

Dim mstr�ϼ����� As String    'ԭʼ���ϼ������ֵ
Dim mstr���� As String        'ԭʼ�ı��������ֵ
Dim mint���� As Integer       '�޸�ǰ�����¼����ڵı�����ĳ���
Dim mintSuccess As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    If IsValid() = False Then Exit Sub
    If Save��λ() = False Then Exit Sub
    mintSuccess = mintSuccess + 1
    If mstrID <> "" Then
        Unload Me
        Exit Sub
    End If
    mstrID = ""
    For i = 2 To 8
        txtEdit(i).Text = ""
    Next
    txtEdit(1).Text = GetMaxLocalCode(mstr�ϼ���λID, "��Լ��λ")
    cmdOK.Enabled = False
    frmUnit.FillList frmUnit.tvwMain_S.SelectedItem.Key
    txtEdit(1).SetFocus
    txtTemp.MaxLength = GetLocalCodeLength(mstr�ϼ���λID, "��Լ��λ")
    txtEdit(1).MaxLength = IIf(txtTemp.MaxLength = 0, 10, txtTemp.MaxLength) - Len(mstr�ϼ�����)
'    Unload Me
End Sub

Private Function IsValid() As Boolean
'����:���������йغ�Լ��λ�������Ƿ���Ч
'����:
'����ֵ:��Ч����True,����ΪFalse
    Dim i As Integer
    Dim strTemp As String
    For i = 1 To 8
        strTemp = Trim(txtEdit(i).Text)
        If LenB(StrConv(strTemp, vbFromUnicode)) > txtEdit(i).MaxLength Then
            MsgBox "���������ݲ��ܳ���" & Int(txtEdit(i).MaxLength / 2) & "������" & "��" & txtEdit(i).MaxLength & "����ĸ��", vbExclamation, gstrSysName
            
            txtEdit(i).SetFocus
            txtEdit(i).SelStart = 0
            txtEdit(i).SelLength = 100
            Exit Function
        End If
        If InStr(strTemp, "'") > 0 Then
            MsgBox "���������ݺ��зǷ��ַ���", vbExclamation, gstrSysName
            txtEdit(i).SetFocus
            txtEdit(i).SelStart = 0
            txtEdit(i).SelLength = 100
            Exit Function
        End If
    Next
    txtEdit(1).Text = Trim(txtEdit(1).Text)
    
    If txtTemp.MaxLength = 0 Then
        If Len(txtEdit(1).Text) = 0 Then
            MsgBox "���벻��Ϊ�ա�", vbExclamation, gstrSysName
            txtEdit(1).SetFocus
            Exit Function
        End If
    Else
        If Len(txtEdit(1).Text) < txtEdit(1).MaxLength Then
            MsgBox "����ĳ��Ȳ�����", vbExclamation, gstrSysName
            txtEdit(1).SetFocus
            Exit Function
        End If
    End If
    If Not IsNumeric(txtEdit(1).Text) Or InStr(txtEdit(1).Text, ",") > 0 Or InStr(txtEdit(1).Text, ".") > 0 Then
        MsgBox "����Ӧ��������ɡ�", vbExclamation, gstrSysName
        txtEdit(1).SetFocus
        Exit Function
    End If
    If Len(Trim(txtEdit(2).Text)) = 0 Then
        MsgBox "���Ʋ���Ϊ�ա�", vbExclamation, gstrSysName
        txtEdit(2).Text = ""
        txtEdit(2).SetFocus
        Exit Function
    End If
    IsValid = True
End Function

Private Function Save��λ() As Boolean
'����:����༭�����ݵ���Լ��λ����
'����:
'����ֵ:�ɹ�����True,����ΪFalse
    Dim lngID As Long
    Dim strվ�� As String
    
    On Error GoTo errHandle
    
    If cboInfo.Text = "" Then
        strվ�� = "Null"
    Else
        strվ�� = Mid(cboInfo.Text, 1, InStr(1, cboInfo.Text, "-") - 1)
    End If
    
    If mstrID = "" Then       '����һ����¼
        lngID = zlDatabase.GetNextId("��Լ��λ")
        gstrSQL = "zl_��Լ��λ_insert(" & lngID & "," & IIf(mstr�ϼ���λID = "", "null", mstr�ϼ���λID) & _
            ",'" & txtTemp.Text & txtEdit(1).Text & "','" & txtEdit(2).Text & _
            "','" & UCase(txtEdit(3).Text) & "','" & txtEdit(4).Text & "','" & txtEdit(5).Text & _
            "','" & txtEdit(6).Text & "','" & txtEdit(7).Text & _
            "','" & txtEdit(8).Text & "'," & chkĩ��.Value & ",null,null,'" & IIf(cboInfo.Text = "", "Null", strվ��) & "')"
    Else    '�޸�
        gstrSQL = "zl_��Լ��λ_update(" & mstrID & "," & IIf(mstr�ϼ���λID = "", "null", mstr�ϼ���λID) & _
            ",'" & txtTemp.Text & txtEdit(1).Text & "','" & txtEdit(2).Text & _
            "','" & UCase(txtEdit(3).Text) & "','" & txtEdit(4).Text & "','" & txtEdit(5).Text & _
            "','" & txtEdit(6).Text & "','" & txtEdit(7).Text & _
            "','" & txtEdit(8).Text & "'," & Len(mstr����) + 1 & ",null,null,'" & IIf(cboInfo.Text = "", "Null", strվ��) & "')"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Save��λ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function �༭��λ(ByVal str�ϼ���λ As String, ByVal str�ϼ���λID As String, ByVal str�ϼ����� As String, _
    Optional strID As String = "", Optional ByVal blnĩ����λ As Boolean) As Boolean
'����:��������õĺ�Լ��λ�����ڽ���ͨѶ�ĳ���,�������ӻ��޸�ĳ����Լ��λ��Ϣ
'����:str�ϼ���λ     �ϼ���Լ��λ������
'     str�ϼ���λID   �ϼ���Լ��λ��ID
'     str�ϼ�����     �ϼ���Լ��λ�ı���
'     strID           ����Լ��λ�ĵ�ID
'     blnĩ����Ŀ     ��������Ŀ�Ƿ�ĩ��
'����ֵ:�༭�ɹ�����True,����ΪFalse
    
    Dim rs��Լ��λ As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer
    Dim rsTmp As New ADODB.Recordset
    
    
    mintSuccess = 0
    
    mstrID = strID
    
    On Error GoTo errHandle
    
    strSQL = "Select ���, ���� From Zlnodelist"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "վ��λ��ѯ")
    If Not rsTmp Is Nothing Then
        cboInfo.AddItem ""
        For i = 0 To rsTmp.RecordCount - 1
            cboInfo.AddItem rsTmp!��� & "-" & rsTmp!����
            rsTmp.MoveNext
        Next
    End If
    
    If strID <> "" Then
        rs��Լ��λ.CursorLocation = adUseClient
        gstrSQL = "select A.ID,A.����,A.���� from ��Լ��λ A,��Լ��λ B " & _
                " where A.ID(+)=B.�ϼ�ID and B.ID=[1]"
'        Call SQLTest(App.ProductName, Me.Caption, gstrSQL)
'        rs��Լ��λ.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
'        Call SQLTest
        Set rs��Լ��λ = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, IIf(strID = "", Null, CLng(strID)))
        
        mstr�ϼ���λID = IIf(IsNull(rs��Լ��λ("ID")), "", rs��Լ��λ("ID"))
        mstr�ϼ����� = IIf(IsNull(rs��Լ��λ("����")), "", rs��Լ��λ("����"))
        
        txtTemp.Text = mstr�ϼ�����
        txtEdit(10).Text = IIf(IsNull(rs��Լ��λ("����")), "��", rs��Լ��λ("����"))
        'ȡ���ϼ����룬�������볤�ȵ�ֵ
        txtTemp.MaxLength = GetLocalCodeLength(mstr�ϼ���λID, "��Լ��λ")
        'txtTemp.MaxLengthΪ0��ʾ�ø��ڵ㻹û���ӽڵ㣬Ҫ��೤�����
        
        rs��Լ��λ.Close
        
        strSQL = "select ID,�ϼ�ID,����,����,����,ĩ��,��ַ,�绰,��������,�ʺ�,��ϵ��,����ʱ��,վ�� from ��Լ��λ  " & _
            "where ID =[1]"
'        Call SQLTest(App.ProductName, Me.Caption, strSQL)
'        rs��Լ��λ.Open strSQL, gcnOracle, adOpenStatic, adLockReadOnly
'        Call SQLTest
        Set rs��Լ��λ = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(strID = "", Null, CLng(strID)))

        txtEdit(1).Text = Mid(rs��Լ��λ("����"), Len(txtTemp.Text) + 1)
        mstr���� = rs��Լ��λ("����")
        '��������ӽڵ����ڵ������
        mint���� = GetDownCodeLength(mstrID, "��Լ��λ")
        ' 8 - (mint���� - Len(mstr����))�����ʽ����˼��ҪΪ���ĺ��ӵı����������
        txtEdit(1).MaxLength = IIf(txtTemp.MaxLength = 0, 10 - (mint���� - Len(mstr����)), txtTemp.MaxLength) - Len(mstr�ϼ�����)
        
        txtEdit(2).Text = rs��Լ��λ("����")
        txtEdit(3).Text = IIf(IsNull(rs��Լ��λ("����")), "", rs��Լ��λ("����"))
        txtEdit(4).Text = IIf(IsNull(rs��Լ��λ("��ַ")), "", rs��Լ��λ("��ַ"))
        txtEdit(5).Text = IIf(IsNull(rs��Լ��λ("�绰")), "", rs��Լ��λ("�绰"))
        txtEdit(6).Text = IIf(IsNull(rs��Լ��λ("��������")), "", rs��Լ��λ("��������"))
        txtEdit(7).Text = IIf(IsNull(rs��Լ��λ("�ʺ�")), "", rs��Լ��λ("�ʺ�"))
        txtEdit(8).Text = IIf(IsNull(rs��Լ��λ("��ϵ��")), "", rs��Լ��λ("��ϵ��"))
        cboInfo.ListIndex = cbo.FindIndex(cboInfo, IIf(IsNull(rs��Լ��λ("վ��")), "", rs��Լ��λ("վ��")))
        If rs��Լ��λ("ĩ��") Then chkĩ��.Value = 1
        chkĩ��.Enabled = False
    Else
        mstr�ϼ���λID = str�ϼ���λID
        txtEdit(10).Text = str�ϼ���λ
        mstr�ϼ����� = str�ϼ�����
        
        txtTemp.Text = str�ϼ�����
        'ȡ���ϼ����룬�������볤�ȵ�ֵ
        txtTemp.MaxLength = GetLocalCodeLength(str�ϼ���λID, "��Լ��λ")
        'txtTemp.MaxLengthΪ0��ʾ�ø��ڵ㻹û���ӽڵ㣬Ҫ��೤�����
        txtEdit(1).MaxLength = IIf(txtTemp.MaxLength = 0, 10, txtTemp.MaxLength) - Len(mstr�ϼ�����)
        txtEdit(1).Text = GetMaxLocalCode(str�ϼ���λID, "��Լ��λ")
        mstr���� = mstr�ϼ����� & txtEdit(1).Text
        If blnĩ����λ Then chkĩ��.Value = 1
    End If
    If chkĩ��.Value <> 1 Then
        For i = 4 To 8
            txtEdit(i).Visible = False
            lblEdit(i).Visible = False
        Next
        
        cboInfo.Top = txtEdit(3).Top
        lblEdit(9).Top = lblEdit(3).Top
        txtEdit(10).Top = txtEdit(4).Top
        lblEdit(10).Top = lblEdit(4).Top
        cmd�ϼ�.Top = txtEdit(10).Top
        frmUnitEdit.Height = 2300
    End If
    
'    If gstrNodeNo = "-" Then
'        txtEdit(9).Visible = False
'        lblEdit(9).Visible = False
'    End If
    frmUnitEdit.Show vbModal
    �༭��λ = mintSuccess > 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmd�ϼ�_Click()
    Dim strSQL As String
    Dim blnRe As Boolean
    Dim str���� As String
    Dim strID As String
    Dim str���� As String
    Dim int����  As Integer
    
    strSQL = "select ID,�ϼ�ID,����,���� from ��Լ��λ  " & _
        "where ĩ�� <> 1 start with �ϼ�ID is null connect by prior ID =�ϼ�ID"
    strID = mstr�ϼ���λID
    str���� = txtEdit(10).Text
    str���� = txtTemp.Text
    blnRe = frm����ѡ��.ShowTree(strSQL, strID, str����, str����, mstrID, "��Լ��λ", "���к�Լ��λ", , mstr����)
    '�ɹ�����
    If blnRe Then
        '�µı����Ŀ��
        int���� = GetLocalCodeLength(strID, "��Լ��λ")
        'ֻ���޸Ĳ��б�Ҫ���
        If mstrID <> "" Then
            If mint���� - Len(mstr����) + IIf(int���� = 0, Len(str����) + 1, int����) > 10 Then
                MsgBox "����ϼ������ʣ���Ϊ���ı���̫���ˡ�", vbExclamation, gstrSysName
                Exit Sub
            End If
        End If
        mstr�ϼ���λID = strID
        txtEdit(10).Text = str����
        txtTemp.MaxLength = int����
        txtTemp.Text = str����
        If mstrID <> "" Then
            txtEdit(1).MaxLength = IIf(txtTemp.MaxLength = 0, 10 - (mint���� - Len(mstr����)), txtTemp.MaxLength) - Len(str����)
        Else
            txtEdit(1).MaxLength = IIf(txtTemp.MaxLength = 0, 10, txtTemp.MaxLength) - Len(str����)
        End If
        txtEdit(1).Text = GetMaxLocalCode(mstr�ϼ���λID, "��Լ��λ")
        'txtEdit(1).Text = Mid(txtEdit(1).Text, Len(txtTemp.Text) + 1)
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    cmdOK.Enabled = True
    If Index = 2 Then
        txtEdit(3).Text = zlCommFun.SpellCode(txtEdit(2).Text)
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    If Index = 2 Or Index = 4 Or Index = 6 Or Index = 8 Then
        OpenIme gstrIme
    ElseIf Index = 1 Or Index = 3 Or Index = 7 Then
        OpenIme
    End If
End Sub


Private Sub txtEdit_LostFocus(Index As Integer)
    If Index = 2 Or Index = 4 Or Index = 6 Or Index = 8 Then
        OpenIme
    End If
End Sub

Private Sub txtTemp_Change()
    txtEdit(1).Width = txtTemp.Width - TextWidth(txtTemp.Text) - 120
    txtEdit(1).Left = txtTemp.Left + TextWidth(txtTemp.Text) + 60
End Sub

