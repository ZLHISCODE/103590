VERSION 5.00
Begin VB.Form frmBloodVerification 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7290
   Icon            =   "frmBloodVerification.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraLine 
      Height          =   90
      Left            =   75
      TabIndex        =   18
      Top             =   1935
      Width           =   7125
   End
   Begin VB.CommandButton CMDcancel 
      Caption         =   "ȡ��"
      Height          =   350
      Left            =   6060
      TabIndex        =   10
      Top             =   2100
      Width           =   1100
   End
   Begin VB.CommandButton CMDok 
      Caption         =   "ȷ��"
      Height          =   350
      Left            =   4860
      TabIndex        =   9
      Top             =   2100
      Width           =   1100
   End
   Begin VB.Frame Fra1 
      Caption         =   "�˶���"
      Height          =   1740
      Index           =   1
      Left            =   4035
      TabIndex        =   11
      Top             =   120
      Width           =   3135
      Begin VB.PictureBox picDown 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   2715
         Picture         =   "frmBloodVerification.frx":030A
         ScaleHeight     =   240
         ScaleWidth      =   225
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   765
         Width           =   255
      End
      Begin VB.TextBox TXT���� 
         Appearance      =   0  'Flat
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1050
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1170
         Width           =   1920
      End
      Begin VB.TextBox txt�û� 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         Left            =   1050
         TabIndex        =   5
         Top             =   330
         Width           =   1920
      End
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         Left            =   1050
         TabIndex        =   6
         Top             =   750
         Width           =   1920
      End
      Begin VB.Label Lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "��      ��"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   1215
         Width           =   900
      End
      Begin VB.Label Lbl�û��� 
         AutoSize        =   -1  'True
         Caption         =   "�˶����ʺ�"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   375
         Width           =   900
      End
      Begin VB.Label Lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "�˶�������"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   810
         Width           =   900
      End
   End
   Begin VB.Frame Fra1 
      Caption         =   "������"
      Height          =   1740
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.PictureBox picDown 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2715
         Picture         =   "frmBloodVerification.frx":06C3
         ScaleHeight     =   240
         ScaleWidth      =   225
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   765
         Width           =   255
      End
      Begin VB.TextBox TXT���� 
         Appearance      =   0  'Flat
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1050
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1170
         Width           =   1920
      End
      Begin VB.TextBox txt�û� 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   0
         Left            =   1050
         TabIndex        =   1
         Top             =   330
         Width           =   1920
      End
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   0
         Left            =   1050
         TabIndex        =   2
         Top             =   750
         Width           =   1920
      End
      Begin VB.Label Lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "��      ��"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   1215
         Width           =   900
      End
      Begin VB.Label Lbl�û��� 
         AutoSize        =   -1  'True
         Caption         =   "�������ʺ�"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   390
         Width           =   900
      End
      Begin VB.Label Lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����������"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   810
         Width           =   900
      End
   End
   Begin VB.Image Image1 
      Height          =   540
      Left            =   3405
      Picture         =   "frmBloodVerification.frx":0A7C
      Stretch         =   -1  'True
      Top             =   660
      Width           =   540
   End
End
Attribute VB_Name = "frmBloodVerification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnReceive As Boolean
Private mblnAutomatic As Boolean
Private mobjfrmMain As Object
Private mblnOK As Boolean
Private mblnUserIsOk As Boolean
Private mstr������ As String

Public Property Get str������() As String
    str������ = mstr������
End Property

Public Function ShowCheck(frmMain As Object, Optional blnAutomatic As Boolean = True) As Boolean
    '����:��ʾ���պ˶���֤ҳ�棬�Խ��ղ����������
    '����:frmmain-�����壬blnAutomatic-�û��Ƿ��Լ���д��������Ϣ��true��ʾ�Զ����ݵ�½�û�����Ϣ������ȡ��false��ʾ�û��Լ���д����֤��������Ϣ��
    Set mobjfrmMain = frmMain
    mblnAutomatic = blnAutomatic
    
    If mblnAutomatic = True Then '�������Զ���ȡ�û�����ʱ
        mblnUserIsOk = userIsOk
        If mblnUserIsOk = False Then MsgBox "�û����������������ֶ���ӽ�����", vbInformation, gstrSysName: GoTo Skip
        TXT����(0).Enabled = False
        TXT����(0).Text = "123"
        txt�û�(0).Text = UserInfo.���
        txt����(0).Text = UserInfo.����
        picDown(0).Visible = False
        txt�û�(0).Enabled = False
        txt����(0).Enabled = False
    End If
Skip:
    Me.Show 1, mobjfrmMain
    ShowCheck = mblnOK
    mblnOK = False
End Function
 
Private Function userIsOk() As Boolean
    '���ܣ��ж��û��Ƿ���Ͻ����˵�����������������Ƿ��ǻ�ʿ����ҽ��,���������ڲ����Ƿ����ٴ����ŵ�
    '����������ֱ��ʹ��userinfo��������ݣ��������ﲻ�������
    Dim strSql As String
    Dim rspeople As ADODB.Recordset
    On Error GoTo Errorhand
    strSql = " select rownum || '-' || b.id as id,b.���,b.����,b.����,a.���� as �������� " & _
             " from ���ű� a,��Ա�� b,��Ա����˵�� c,������Ա d,��������˵�� e,�ϻ���Ա�� f " & _
             " where a.id=d.����id and a.id=e.����id and Instr(',�ٴ�,����,', ',' || e.�������� || ',', 1) <> 0 and f.��Աid=b.id " & _
             " and d.��Աid=b.id and b.id=c.��Աid and c.��Ա���� in('ҽ��','��ʿ') and b.id=[1]"
    
    Set rspeople = gobjDatabase.OpenSQLRecord(strSql, "��Ա��Ϣ", UserInfo.id)
    If rspeople.RecordCount = 0 Then
        userIsOk = False
    Else
        userIsOk = True
    End If
Errorhand:
End Function

Private Function GetUserName(ByVal objControl As TextBox, ByVal intIndex As Integer, Optional ByVal StrInput As String = "") As Boolean
    Dim rsUser As ADODB.Recordset
    Dim strSql As String, strWhere As String
    Dim vPoint As POINTAPI, blnCancel As Boolean

    On Error GoTo errHand

    If StrInput <> "" Then
         If IsNumeric(StrInput) Then
            strWhere = " And a.��� Like '" & txt�û�(intIndex).Text & "%'"
         ElseIf gobjCommFun.IsNumOrChar(StrInput) Then
            strWhere = " And f.�û��� Like '" & UCase(txt�û�(intIndex).Text) & "%'"
         Else
            strWhere = " And a.���� Like '" & txt����(intIndex).Text & "%'"
         End If
    End If
    vPoint = GetCoordPos(Me.hWnd, objControl.Left + Fra1(intIndex).Left, objControl.Top + Fra1(intIndex).Top) ',b.���� as ����,b.id ||
    strSql = " Select distinct f.�û��� || '-' || a.id as ID,f.�û���,a.���,a.����,a.���� " & _
            " From ��Ա�� a, ���ű� b, ������Ա c, ��������˵�� d, ��Ա����˵�� e,�ϻ���Ա�� f " & _
            " Where a.Id = c.��Աid And b.Id = c.����id And a.Id = e.��Աid And b.Id = d.����id and f.��Աid=a.id  And Instr(',�ٴ�,����,', ',' || d.�������� || ',', 1) <> 0 And " & _
            "  e.��Ա���� In ('ҽ��', '��ʿ') And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)" & strWhere
    Set rsUser = gobjDatabase.ShowSQLSelect(Me, strSql, 0, "", False, txt�û�(intIndex).Text, "��ѡ��һ��ȡѪ��Ա", False, False, True, vPoint.X, vPoint.Y, objControl.Height, blnCancel, False, False, False)

    If Not rsUser Is Nothing Then
        If blnCancel = False Then
            If rsUser.EOF Then Exit Function
            Lbl�û���(intIndex).Tag = Split(rsUser!id, "-")(1) '�û�id
            txt�û�(intIndex).Text = Nvl(rsUser!�û���) '�û��ı��
            txt�û�(intIndex).Tag = Nvl(rsUser!�û���) '�û��ĵ�½��
            objControl.Text = Nvl(rsUser!����) '�û�������
            objControl.Tag = objControl.Text '�û�������
            objControl.SetFocus
            GetUserName = True
        End If
    Else
        If StrInput = "" And blnCancel = False Then
            MsgBox "û�ж�Ӧ���ٴ���ʿ��ҽ����Ϣ��������Ա���������ã�", vbInformation, gstrSysName
        End If
    End If
    
    Exit Function
errHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub CMDcancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub CMDok_Click()
    '���ܣ����ȷ�����ж��û��������ݵ���ȷ�Ժ͹淶�ԣ�ͬʱ�ж��û������������û���Ϣ�Ƿ���ȷ��
    Dim strNote As String
    Dim strUserNo������ As String
    Dim strUserNo������ As String
    Dim strUserName������ As String
    Dim strUserName������ As String
    Dim strPassword������ As String
    Dim strPassword������ As String

    
    Dim strServerName As String
    On Error GoTo InputError
    
    'ȡѪ����֤���
    '------�����û��Ƿ�oracle�Ϸ��û�----------------
    strUserNo������ = Trim(txt�û�(0).Text)
    strUserName������ = Trim(txt�û�(0).Tag)
    strPassword������ = Trim(TXT����(0).Text)
    strUserNo������ = Trim(txt�û�(1).Text)
    strUserName������ = Trim(txt�û�(1).Tag)
    strPassword������ = Trim(TXT����(1).Text)
    
    '��Ч�ַ���Ч��
    If mblnUserIsOk = False Then '����Ҫ�û��ֶ���д�����˵�����£�Ҫ�Խ����˵����ݽ����ж�
        If Len(Trim(txt�û�(0))) = 0 Then
            strNote = "������������ʺ�"
            Call gobjControl.ControlSetFocus(txt�û�(0))
            GoTo InputError
        End If
        If Len(strUserNo������) <> 1 Then
            If Mid(strUserNo������, 1, 1) = "/" Or Mid(strUserNo������, 1, 1) = "@" Or Mid(strUserNo������, Len(strUserNo������) - 1, 1) = "/" Or Mid(strUserNo������, Len(strUserNo������) - 1, 1) = "@" Then
                strNote = "�������ʺŴ���"
                Call gobjControl.ControlSetFocus(txt�û�(0))
                Exit Sub
            End If
        End If
        If Trim(strPassword������) <> "" And Len(strPassword������) <> 1 Then
            If Mid(strPassword������, Len(strPassword������) - 1, 1) = "/" Or Mid(strPassword������, Len(strPassword������) - 1, 1) = "@" Or Mid(strPassword������, 1, 1) = "/" Or Mid(strPassword������, 1, 1) = "@" Then
                strNote = "�������ʺ��������"
                Call gobjControl.ControlSetFocus(TXT����(0))
                GoTo InputError
            End If
        End If
        If Len(Trim(strPassword������)) = 0 Then
            strNote = "������������ʺ�����"
            Call gobjControl.ControlSetFocus(TXT����(0))
            GoTo InputError
        End If
        If GetObjectRegister = False Then Exit Sub
        strServerName = gobjRegister.GetServerName
        If gobjRegister.LoginValidate(strServerName, strUserName������, strPassword������, strNote) = False Then
            TXT����(0).Text = ""
            Call gobjControl.ControlSetFocus(TXT����(0))
            GoTo InputError
        End If
    End If
    
    If Len(Trim(txt�û�(1))) = 0 Then
        strNote = "������������ʺ�"
        Call gobjControl.ControlSetFocus(txt�û�(1))
        GoTo InputError
    End If
    If Len(strUserNo������) <> 1 Then
        If Mid(strUserNo������, 1, 1) = "/" Or Mid(strUserNo������, 1, 1) = "@" Or Mid(strUserNo������, Len(strUserNo������) - 1, 1) = "/" Or Mid(strUserNo������, Len(strUserNo������) - 1, 1) = "@" Then
            strNote = "�������ʺŴ���"
            Call gobjControl.ControlSetFocus(txt�û�(1))
            Exit Sub
        End If
    End If
    If Trim(strPassword������) <> "" And Len(strPassword������) <> 1 Then
        If Mid(strPassword������, Len(strPassword������) - 1, 1) = "/" Or Mid(strPassword������, Len(strPassword������) - 1, 1) = "@" Or Mid(strPassword������, 1, 1) = "/" Or Mid(strPassword������, 1, 1) = "@" Then
            strNote = "�������ʺ��������"
            Call gobjControl.ControlSetFocus(TXT����(1))
            GoTo InputError
        End If
    End If
    If Len(Trim(strPassword������)) = 0 Then
        strNote = "������������ʺ�����"
        Call gobjControl.ControlSetFocus(TXT����(1))
        GoTo InputError
    End If


    '�����˺ͺ����˲�����ͬһ��
    If txt����(0).Text = txt����(1).Text Or txt�û�(0).Text = txt�û�(1).Text Then
        strNote = "�����˺ͺ����˲�����ͬһ���ˣ������º˶�"
        Call gobjControl.ControlSetFocus(txt����(1))
        GoTo InputError
    End If
    '�û���¼��֤
    If GetObjectRegister = False Then Exit Sub
    strServerName = gobjRegister.GetServerName
    If gobjRegister.LoginValidate(strServerName, strUserName������, strPassword������, strNote) = False Then
        TXT����(1).Text = ""
        Call gobjControl.ControlSetFocus(TXT����(1))
        GoTo InputError
    End If
    
    mblnOK = True
    mstr������ = txt����(1).Text
    Unload Me
    Exit Sub
InputError:
    If strNote <> "" Then
        MsgBox strNote, vbInformation, gstrSysName
    End If
    Exit Sub
End Sub

Private Sub picDown_Click(Index As Integer)
    If GetUserName(txt����(Index), Index) = True Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub TXT����_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub txt����_KeyPress(Index As Integer, KeyAscii As Integer)
    TXT����(Index).Text = ""
    If KeyAscii = vbKeyReturn Then
        If GetUserName(txt����(Index), Index, txt����(Index).Text) = True Then gobjCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt�û�_KeyPress(Index As Integer, KeyAscii As Integer)
    TXT����(Index).Text = ""
    If KeyAscii = vbKeyReturn Then
        If GetUserName(txt����(Index), Index, txt�û�(Index).Text) = True Then gobjCommFun.PressKey vbKeyTab
    End If
End Sub
