VERSION 5.00
Begin VB.Form frm����ѡ��_���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ѡ��"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   Icon            =   "frm����ѡ��_����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txt����֢ 
      Height          =   1335
      Left            =   1350
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1515
      Width           =   3675
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2610
      TabIndex        =   8
      Top             =   3105
      Width           =   1100
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3870
      TabIndex        =   9
      Top             =   3105
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   30
      TabIndex        =   7
      Top             =   2985
      Width           =   6075
   End
   Begin VB.TextBox txt������Ϣ 
      Height          =   300
      Left            =   1350
      TabIndex        =   0
      Top             =   675
      Width           =   3375
   End
   Begin VB.CommandButton cmd������Ϣ 
      Caption         =   "��"
      Height          =   300
      Left            =   4740
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   675
      Width           =   285
   End
   Begin VB.Label lbl����֢ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����֢(&M)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   5
      Top             =   1575
      Width           =   810
   End
   Begin VB.Label lblDemo 
      Caption         =   "    ��û��Ϊ�ò������ó�Ժ���֣����Գ�Ժ����ȱʡΪ��Ժ���֣���ȷ���������Ժ����"
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1065
      Width           =   3915
   End
   Begin VB.Label lblPatient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����������ҽ��    ���ţ�01234567    "
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   405
      Left            =   330
      TabIndex        =   1
      Top             =   225
      Width           =   4785
   End
   Begin VB.Label lbl������Ϣ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��Ժ����(&U)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   300
      TabIndex        =   2
      Top             =   735
      Width           =   990
   End
End
Attribute VB_Name = "frm����ѡ��_����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mblnStart As Boolean
Private mint���� As Integer
Private mlng����ID As Long
Private mlng��ҳID As Long

Private mstr��Ժ���� As String
Private mstr����֢ As String

Private Sub cmdOK_Click()
    On Error GoTo errHand
    '����ѡ������Ϣ
    If txt������Ϣ.Tag = "" Then
        MsgBox "��Ϊ�òα�����ѡ�񼲲�������Ϣ��", vbInformation, gstrSysName
        txt������Ϣ.SetFocus
        Exit Sub
    End If
    
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & mlng����ID & "," & mint���� & ",'��Ժ����ID','" & Split(txt������Ϣ.Tag, "|")(1) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���³�Ժ����")
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & mlng����ID & "," & mint���� & ",'����֢','''" & txt����֢.Text & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���֢")
    
    mblnOK = True
    mstr��Ժ���� = Split(txt������Ϣ.Tag, "|")(0)
    mstr����֢ = txt����֢.Text
    
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmd������Ϣ_Click()
    Dim rs���� As New ADODB.Recordset
    gstrSQL = "" & _
        "   Select A.ID,A.����,A.����,A.���� " & _
        "   From ���ղ��� A " & _
        "   where a.���=0 and A.����=[1] Order by A.����"
        
    Set rs���� = New ADODB.Recordset
    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "�����֤", TYPE_����)
    
    If rs����.RecordCount > 0 Then
        If frmListSel.ShowSelect(TYPE_����, rs����, "ID", "ҽ������ѡ��", "��ѡ���Ժ���֣�") = True Then
            txt������Ϣ.Tag = rs����!���� & "|" & rs����!ID
            txt������Ϣ.Text = "(" & rs����!���� & ")" & rs����!����
            lbl������Ϣ.Tag = txt������Ϣ.Text '���ڻָ���ʾ
            
            If txt������Ϣ.Tag = "" Then
                txt������Ϣ.Text = "(" & rs����!���� & ")" & rs����!����
                txt������Ϣ.Tag = rs����!���� & "|" & rs����!ID
                lbl������Ϣ.Tag = txt������Ϣ.Text
            End If
            
        End If
    End If
End Sub

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If Not mblnStart Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim blnSet As Boolean               '˵���Ƿ����ó�Ժ����
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '��ȡ�ò��˵Ļ�����Ϣ
    gstrSQL = " Select B.����,A.����,A.ҽ����,C.ID ��Ժ����ID,C.���� ��Ժ���ֱ���,C.���� ��Ժ��������," & _
              " D.ID ��Ժ����ID,D.���� ��Ժ���ֱ���,D.���� ��Ժ��������,A.����֢" & _
              " From �����ʻ� A,������Ϣ B," & _
              " (Select * From ���ղ��� Where ����=" & mint���� & ") C," & _
              " (Select * From ���ղ��� Where ����=" & mint���� & ") D" & _
              " Where A.����ID=B.����ID And A.����ID=[1] And A.����=[2]" & _
              " And C.ID(+)=A.����ID And D.ID(+)=A.��Ժ����ID"
              
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ò��˵Ļ�����Ϣ", mlng����ID, mint����)
    
    lblPatient.Caption = "������" & Nvl(rsTemp!����) & Space(4) & "���ţ�" & Nvl(rsTemp!����) & Space(4) & "���˱�ţ�" & Nvl(rsTemp!ҽ����)
    txt����֢.Text = Nvl(rsTemp!����֢)
    
    If Not IsNull(rsTemp!��Ժ���ֱ���) Then
        blnSet = True
        txt������Ϣ.Text = "(" & rsTemp!��Ժ���ֱ��� & ")" & rsTemp!��Ժ��������
        txt������Ϣ.Tag = rsTemp!��Ժ���ֱ��� & "|" & rsTemp!��Ժ����ID
        lbl������Ϣ.Tag = txt������Ϣ.Text
    Else
        blnSet = False
        If Not IsNull(rsTemp!��Ժ���ֱ���) Then
            txt������Ϣ.Text = "(" & rsTemp!��Ժ���ֱ��� & ")" & rsTemp!��Ժ��������
            txt������Ϣ.Tag = rsTemp!��Ժ���ֱ��� & "|" & rsTemp!��Ժ����ID
            lbl������Ϣ.Tag = txt������Ϣ.Text
        End If
    End If
    
    mblnStart = True
    Exit Sub
errHand:
    MsgBox "��ȷ�ϱ����ʻ���Ľṹ�����µģ�", vbInformation, gstrSysName
End Sub

Public Function ShowSelect(ByVal int���� As Integer, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
     ByRef str��Ժ���� As String, ByRef str����֢ As String) As Boolean
    'ѡ���˵���Ժ���ּ���Ժ���֣�ͬʱ�����˱���סԺ�������Ϣ��ʾ����
    '���±����ʻ��Ĳ���ID����Ժ���֣�����Ժ���֣�������Ժ���ּ���Ժ���ֱ��뷵�ظ�����ģ��
    mblnOK = False
    mint���� = int����
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    
    Me.Show 1
    
    str��Ժ���� = mstr��Ժ����
    If asc(Right(mstr����֢, 1)) = 10 Or asc(Right(mstr����֢, 1)) = 13 Then
        mstr����֢ = Substr(mstr����֢, 1, zlCommFun.ActualLen(mstr����֢) - 1)
    End If
    str����֢ = mstr����֢
    ShowSelect = mblnOK
End Function

Private Sub txt����֢_GotFocus()
    Call zlControl.TxtSelAll(txt����֢)
End Sub

Private Sub txt����֢_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdOK.SetFocus
End Sub

Private Sub txt������Ϣ_GotFocus()
    Call zlControl.TxtSelAll(txt������Ϣ)
End Sub

Private Sub txt������Ϣ_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt������Ϣ.Text = "" And txt������Ϣ.Tag <> "" Then Exit Sub
    
    On Error GoTo errHandle
    strText = txt������Ϣ.Text
    If InStr(1, strText, "(") <> 0 Then
        If InStr(1, strText, ")") <> 0 Then
            strText = Mid(strText, 2, InStr(1, strText, ")") - 2)
        End If
    End If
    
    gstrSQL = "Select A.ID,A.����,A.����,A.����" & _
             "   FROM ���ղ��� A " & _
             "   WHERE nvl(���,0)=0 and A.����=[1] And " & _
             " (A.���� like [2] || '%' or A.���� like [2] || '%' or A.���� like [2] || '%' )" & _
             " Order by A.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint����, strText)
    If rsTemp.RecordCount = 0 Then
        MsgBox "�����ڸò��֣����������룡", vbInformation, gstrSysName
        txt������Ϣ.Text = lbl������Ϣ.Tag
        zlControl.TxtSelAll txt������Ϣ
        Exit Sub
    Else
        '����ѡ����
        If rsTemp.RecordCount > 1 Then
            '�����ֶδ���3�ģ���ʹֻ��һ����¼�ѸöԻ�����ʾ�������Ա����û��õ��������Ϣ
            blnReturn = frmListSel.ShowSelect(TYPE_����, rsTemp, "ID", "ҽ������ѡ��", "��ѡ��ҽ�����֣�")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '��¼����û�п�ѡ�������
        txt������Ϣ.Text = lbl������Ϣ.Tag
        zlControl.TxtSelAll txt������Ϣ
    Else
        '�϶����м�¼����
        txt������Ϣ.Tag = rsTemp!���� & "|" & rsTemp!ID
        txt������Ϣ.Text = "(" & rsTemp!���� & ")" & rsTemp!����
        lbl������Ϣ.Tag = txt������Ϣ.Text '���ڻָ���ʾ
        txt����֢.SetFocus
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
