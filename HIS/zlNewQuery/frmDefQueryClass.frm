VERSION 5.00
Begin VB.Form frmDefQueryClass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҳ�����"
   ClientHeight    =   1965
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4740
   Icon            =   "frmDefQueryClass.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3510
      TabIndex        =   10
      Top             =   630
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3510
      TabIndex        =   9
      Top             =   150
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   75
      TabIndex        =   11
      Top             =   60
      Width           =   3195
      Begin VB.TextBox txt 
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   0
         Left            =   1065
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "11111"
         Top             =   270
         Width           =   900
      End
      Begin VB.TextBox txtTemp 
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   825
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "����"
         Text            =   "1111111111"
         Top             =   225
         Width           =   1275
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "��"
         Height          =   285
         Left            =   2775
         TabIndex        =   8
         Top             =   1350
         Width           =   285
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   825
         MaxLength       =   30
         TabIndex        =   3
         Top             =   615
         Width           =   1935
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   825
         MaxLength       =   15
         TabIndex        =   5
         Top             =   975
         Width           =   1935
      End
      Begin VB.TextBox txt 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   3
         Left            =   825
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1350
         Width           =   1935
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����(&B)"
         Height          =   180
         Index           =   0
         Left            =   135
         TabIndex        =   0
         Top             =   285
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   675
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����(&S)"
         Height          =   180
         Index           =   2
         Left            =   135
         TabIndex        =   4
         Top             =   1065
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�ϼ�(&U)"
         Height          =   180
         Index           =   3
         Left            =   135
         TabIndex        =   6
         Top             =   1410
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmDefQueryClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const mlng���볤�� As Long = 10

Private mlngKey As Long
Private mlngUpKey As Long
Private mstr�ϼ�ID As String
Private mstr�ϼ����� As String
Private mstr���� As String
Private mblnOK As Boolean

Private Function GetTreeCode(ByVal lngUpKey As Long) As Boolean
    '��ȡ���ͽṹ�ı������,�����ϼ�����,��������
    
    If lngUpKey = 0 Then
        '���û��ָ���ϼ�
        mstr�ϼ����� = ""
        txtTemp.Text = ""
        
        txt(3).Text = "��"
        
        'ȡ���ϼ����룬�������볤�ȵ�ֵ
        txtTemp.MaxLength = GetLocalCodeLength("", "��ѯҳ��Ŀ¼")
        
    Else
        'ָ�����ϼ�
        gstrSQL = "select ���� as �ϼ�����,ҳ������ as �ϼ�����,ҳ����� as �ϼ�ID from ��ѯҳ��Ŀ¼ where ҳ�����=[1]"
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngUpKey)
                        
        mstr�ϼ�ID = IIf(IsNull(gRs("�ϼ�ID")), "", gRs("�ϼ�ID"))
        mstr�ϼ����� = IIf(IsNull(gRs("�ϼ�����")), "", gRs("�ϼ�����"))
        txt(3).Text = IIf(IsNull(gRs("�ϼ�����")), "��", gRs("�ϼ�����"))
        txt(3).Tag = lngUpKey
        txtTemp.MaxLength = 0
        txtTemp.Text = mstr�ϼ�����
        
        '�жϱ����Ƿ�����
        If Len(mstr�ϼ�����) >= mlng���볤�� Then
            MsgBox "�����������ӷ����ˣ����볤���Ѿ��þ���", vbExclamation, gstrSysName
            Exit Function
        End If
        
        'ȡ���ϼ����룬�������볤�ȵ�ֵ
        txtTemp.MaxLength = GetLocalCodeLength(mstr�ϼ�ID, "��ѯҳ��Ŀ¼")
    End If
        
    txt(0).MaxLength = IIf(txtTemp.MaxLength = 0, mlng���볤��, txtTemp.MaxLength) - Len(mstr�ϼ�����)
    txt(0).Text = Mid(txt(0).Text, Len(txtTemp.Text) + 1)
    
    If mlngKey = 0 Then txt(0).Text = GetMaxLocalCode(mstr�ϼ�ID, "��ѯҳ��Ŀ¼")
    
    GetTreeCode = True
End Function

Public Function ShowEdit(ByVal frmParent As Form, ByVal lngKey As Long, ByVal lngUpKey As Long) As Boolean
    
    mblnOK = False
    
    mlngUpKey = lngUpKey
    mlngKey = lngKey
    
    If lngKey > 0 Then
        '�޸ķ���
        gstrSQL = "Select ����,ҳ������,���� from ��ѯҳ��Ŀ¼ where ҳ�����=[1]"
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
        If gRs.BOF = False Then
            txt(0).Text = IIf(IsNull(gRs("����")), "", gRs("����"))
            txt(1).Text = IIf(IsNull(gRs("ҳ������")), "", gRs("ҳ������"))
            txt(2).Text = IIf(IsNull(gRs("����")), "", gRs("����"))
            mstr���� = txt(0).Text
        End If
    End If
    
    If GetTreeCode(mlngUpKey) = False Then Exit Function
                    
    cmdOK.Tag = ""
    
    Me.Show 1, frmParent
    
    ShowEdit = mblnOK
End Function

Private Function CheckValid() As Boolean
    txt(0).Text = Trim(txt(0).Text)

    If txtTemp.MaxLength = 0 Then
        If Len(txt(0).Text) = 0 Then
            MsgBox "���벻��Ϊ�ա�", vbExclamation, gstrSysName
            txt(0).SetFocus
            Exit Function
        End If
    Else
        If Len(txt(0).Text) < txt(0).MaxLength Then
            MsgBox "����ĳ��Ȳ�����", vbExclamation, gstrSysName
            txt(0).SetFocus
            Exit Function
        End If
    End If
    If Not IsNumeric(txt(0).Text) Or InStr(txt(0).Text, ",") > 0 Or InStr(txt(0).Text, ".") > 0 Or InStr(txt(0).Text, "-") > 0 Then
        MsgBox "����Ӧ��������ɡ�", vbExclamation, gstrSysName
        txt(0).SetFocus
        Exit Function
    End If
    If Len(Trim(txt(1).Text)) = 0 Then
        MsgBox "���Ʋ���Ϊ�ա�", vbExclamation, gstrSysName
        txt(1).Text = ""
        txt(1).SetFocus
        Exit Function
    End If
    
    CheckValid = True
End Function
Private Function SaveData() As Boolean
    Dim lng��� As Long
    
    If cmdOK.Tag = "1" Then
                                
        If CheckValid = False Then Exit Function
                                        
        If mlngKey = 0 Then
            lng��� = NextValue("��ѯҳ��Ŀ¼", "ҳ�����")
            gstrSQL = "zl_��ѯҳ��Ŀ¼_insert(" & lng��� & ",'" & txt(1).Text & "',NULL,NULL,NULL,NULL,NULL," & IIf(Val(txt(3).Tag) = 0, "NULL", Val(txt(3).Tag)) & ",0,'" & txtTemp.Text & txt(0).Text & "','" & txt(2).Text & "')"
        Else
            lng��� = mlngKey
            gstrSQL = "zl_��ѯҳ��Ŀ¼_update(" & lng��� & ",'" & txt(1).Text & "',NULL,NULL,NULL,NULL," & IIf(Val(txt(3).Tag) = 0, "NULL", Val(txt(3).Tag)) & ",'" & txtTemp.Text & txt(0).Text & "','" & txt(2).Text & "'," & Len(mstr����) + 1 & ")"
        End If
        
        On Error GoTo errHand
        
        
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        
        Call frmDefQuery.RefreshClass(CStr(lng���))
    End If
    SaveData = True
    
    Exit Function
    
errHand:
    If ErrCenter() = -1 Then Resume
    
End Function

Private Sub cmdOpen_Click()
    Dim strSQL As String
    Dim blnRe As Boolean
    Dim str���� As String
    Dim strID As String
    Dim strRerurnID As String
    Dim str���� As String
    Dim int����  As Integer
    
    strSQL = "Select ҳ����� AS id,�ϼ���� AS �ϼ�id,ҳ������ AS ����,����,0 as ĩ�� From ��ѯҳ��Ŀ¼ Where (ĩ�� IS NULL OR ĩ��=0)  Start with �ϼ���� is null connect by prior ҳ����� =�ϼ���� "
    
    strID = txt(3).Tag
    str���� = txt(3).Text
    str���� = txtTemp.Text & txt(0).Text
        
    blnRe = frm����ѡ��.ShowTree(strSQL, strID, str����, mstr�ϼ�����, mlngKey, Me.Caption, "����ҳ�����", , mstr����)

    If blnRe Then       '�µı����Ŀ��
        
        mlngUpKey = Val(strID)
        txt(3).Tag = strID
        txt(3).Text = str����
        Call GetTreeCode(mlngUpKey)
        txt(0).Text = GetMaxLocalCode(strID, "��ѯҳ��Ŀ¼")
        cmdOK.Tag = "1"
    End If
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdOK.Tag = "1" Then
        If MsgBox("���ĺ�Ĳ�ѯĿ¼���뱣�������Ч" & vbCrLf & "ȷ�ϲ�������˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = True
    End If
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If SaveData Then
        mblnOK = True
        
        If mlngKey = 0 Then
            txt(0).Text = ""
            txt(1).Text = ""
            txt(2).Text = ""
            
            txt(0).Text = GetMaxLocalCode(txt(3).Tag, "��ѯҳ��Ŀ¼")
            
            txt(0).SetFocus
            cmdOK.Tag = ""
        Else
            cmdOK.Tag = ""
            Unload Me
        End If
    End If
    
End Sub

Private Sub txt_Change(Index As Integer)
    cmdOK.Tag = "1"
    If Index = 1 Then
        txt(2).Text = zlCommFun.SpellCode(txt(1).Text)
    End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
    Call SelAll(txt(Index))
    If Index = 0 Then zlCommFun.OpenIme True
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
        If Index = 3 Then SendKeys "{TAB}"
    Else
        If Index = 2 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If Index = 3 And Chr(KeyAscii) = "*" Then Call cmdOpen_Click
    End If
    
End Sub

Private Sub txt_LostFocus(Index As Integer)
    If Index = 0 Then zlCommFun.OpenIme
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub

Private Sub txtTemp_Change()
    txt(0).Width = txtTemp.Width - TextWidth(txtTemp.Text) - 120
    txt(0).Left = txtTemp.Left + TextWidth(txtTemp.Text) + 60
End Sub
