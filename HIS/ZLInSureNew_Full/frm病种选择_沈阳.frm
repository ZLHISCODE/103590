VERSION 5.00
Begin VB.Form frm����ѡ��_���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ѡ��"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   Icon            =   "frm����ѡ��_����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txt����֢ 
      Height          =   300
      Left            =   1350
      TabIndex        =   9
      Top             =   2010
      Width           =   3675
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2610
      TabIndex        =   11
      Top             =   2550
      Width           =   1100
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3870
      TabIndex        =   12
      Top             =   2550
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   30
      TabIndex        =   10
      Top             =   2430
      Width           =   6075
   End
   Begin VB.TextBox txt������Ϣ 
      Height          =   300
      Index           =   1
      Left            =   1350
      TabIndex        =   5
      Top             =   1170
      Width           =   3375
   End
   Begin VB.CommandButton cmd������Ϣ 
      Caption         =   "��"
      Height          =   300
      Index           =   1
      Left            =   4740
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1170
      Width           =   285
   End
   Begin VB.TextBox txt������Ϣ 
      Height          =   300
      Index           =   0
      Left            =   1350
      TabIndex        =   2
      Top             =   780
      Width           =   3375
   End
   Begin VB.CommandButton cmd������Ϣ 
      Caption         =   "��"
      Height          =   300
      Index           =   0
      Left            =   4740
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   780
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
      TabIndex        =   8
      Top             =   2070
      Width           =   810
   End
   Begin VB.Label lblDemo 
      Caption         =   "    ��û��Ϊ�ò������ó�Ժ���֣����Գ�Ժ����ȱʡΪ��Ժ���֣���ȷ���������Ժ����"
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   1560
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
      TabIndex        =   0
      Top             =   180
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
      Index           =   1
      Left            =   300
      TabIndex        =   4
      Top             =   1230
      Width           =   990
   End
   Begin VB.Label lbl������Ϣ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��Ժ����(&R)"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Top             =   840
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
Private mstr��Ժ���� As String
Private mstr����֢ As String
Private Enum ����
    ��Ժ����
    ��Ժ����
End Enum

Private Sub cmdOK_Click()
    On Error GoTo errHand
    '����ѡ������Ϣ
    If txt������Ϣ(��Ժ����).Tag = "" Then
        MsgBox "��Ϊ�òα�����ѡ�񼲲�������Ϣ��", vbInformation, gstrSysName
        txt������Ϣ(��Ժ����).SetFocus
        Exit Sub
    End If
    If txt������Ϣ(��Ժ����).Tag = "" Then
        MsgBox "��Ϊ�òα�����ѡ�񼲲�������Ϣ��", vbInformation, gstrSysName
        txt������Ϣ(��Ժ����).SetFocus
        Exit Sub
    End If
    
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & mlng����ID & "," & mint���� & ",'����ID','" & Split(txt������Ϣ(��Ժ����).Tag, "|")(1) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ����")
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & mlng����ID & "," & mint���� & ",'��Ժ����ID','" & Split(txt������Ϣ(��Ժ����).Tag, "|")(1) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���³�Ժ����")
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & mlng����ID & "," & mint���� & ",'����֢','''" & txt����֢.Text & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���֢")
    
    mblnOK = True
    mstr��Ժ���� = Split(txt������Ϣ(��Ժ����).Tag, "|")(0)
    mstr��Ժ���� = Split(txt������Ϣ(��Ժ����).Tag, "|")(0)
    mstr����֢ = txt����֢.Text
    
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmd������Ϣ_Click(Index As Integer)
    Dim rs���� As New ADODB.Recordset
    gstrSQL = " Select A.ID,A.����,A.����,A.���� " & _
            " From ���ղ��� A where A.����=[1] Order by A.����"
    Set rs���� = New ADODB.Recordset
    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "�����֤", TYPE_������)
    If rs����.RecordCount > 0 Then
        If frmListSel.ShowSelect(TYPE_������, rs����, "ID", "ҽ������ѡ��", "��ѡ��" & IIf(Index = 0, "��Ժ", "��Ժ") & "���֣�") = True Then
            txt������Ϣ(Index).Tag = rs����!���� & "|" & rs����!ID
            txt������Ϣ(Index).Text = "(" & rs����!���� & ")" & rs����!����
            lbl������Ϣ(Index).Tag = txt������Ϣ(Index).Text '���ڻָ���ʾ
            
            If txt������Ϣ(��Ժ����).Tag = "" Then
                txt������Ϣ(��Ժ����).Text = "(" & rs����!���� & ")" & rs����!����
                txt������Ϣ(��Ժ����).Tag = rs����!���� & "|" & rs����!ID
                lbl������Ϣ(��Ժ����).Tag = txt������Ϣ(��Ժ����).Text
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
        txt������Ϣ(��Ժ����).Text = "(" & rsTemp!��Ժ���ֱ��� & ")" & rsTemp!��Ժ��������
        txt������Ϣ(��Ժ����).Tag = rsTemp!��Ժ���ֱ��� & "|" & rsTemp!��Ժ����ID
        lbl������Ϣ(��Ժ����).Tag = txt������Ϣ(��Ժ����).Text
    End If
    If Not IsNull(rsTemp!��Ժ���ֱ���) Then
        blnSet = True
        txt������Ϣ(��Ժ����).Text = "(" & rsTemp!��Ժ���ֱ��� & ")" & rsTemp!��Ժ��������
        txt������Ϣ(��Ժ����).Tag = rsTemp!��Ժ���ֱ��� & "|" & rsTemp!��Ժ����ID
        lbl������Ϣ(��Ժ����).Tag = txt������Ϣ(��Ժ����).Text
    Else
        blnSet = False
        If Not IsNull(rsTemp!��Ժ���ֱ���) Then
            txt������Ϣ(��Ժ����).Text = "(" & rsTemp!��Ժ���ֱ��� & ")" & rsTemp!��Ժ��������
            txt������Ϣ(��Ժ����).Tag = rsTemp!��Ժ���ֱ��� & "|" & rsTemp!��Ժ����ID
            lbl������Ϣ(��Ժ����).Tag = txt������Ϣ(��Ժ����).Text
        End If
    End If
    
    '���δ���ó�Ժ���֣����������С
    If blnSet Then
        Me.lblDemo.Visible = False
        Me.lbl����֢.Top = Me.lbl������Ϣ(��Ժ����).Top - Me.lbl������Ϣ(��Ժ����).Top + Me.lbl������Ϣ(��Ժ����).Top
        Me.txt����֢.Top = Me.txt������Ϣ(��Ժ����).Top - Me.txt������Ϣ(��Ժ����).Top + Me.txt������Ϣ(��Ժ����).Top
    End If
    mblnStart = True
    Exit Sub
errHand:
    MsgBox "��ȷ�ϱ����ʻ���Ľṹ�����µģ�", vbInformation, gstrSysName
End Sub

Public Function ShowSelect(ByVal int���� As Integer, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByRef str��Ժ���� As String, ByRef str��Ժ���� As String, ByRef str����֢ As String) As Boolean
    'ѡ���˵���Ժ���ּ���Ժ���֣�ͬʱ�����˱���סԺ�������Ϣ��ʾ����
    '���±����ʻ��Ĳ���ID����Ժ���֣�����Ժ���֣�������Ժ���ּ���Ժ���ֱ��뷵�ظ�����ģ��
    mblnOK = False
    mint���� = int����
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    Me.Show 1
    str��Ժ���� = mstr��Ժ����
    str��Ժ���� = mstr��Ժ����
    str����֢ = mstr����֢
    ShowSelect = mblnOK
End Function

Private Sub txt����֢_GotFocus()
    Call zlControl.TxtSelAll(txt����֢)
End Sub

Private Sub txt����֢_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdOK.SetFocus
End Sub

Private Sub txt������Ϣ_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txt������Ϣ(Index))
End Sub

Private Sub txt������Ϣ_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt������Ϣ(Index).Text = "" And txt������Ϣ(Index).Tag <> "" Then Exit Sub
    
    On Error GoTo errHandle
    strText = txt������Ϣ(Index).Text
    If InStr(1, strText, "(") <> 0 Then
        If InStr(1, strText, ")") <> 0 Then
            strText = Mid(strText, 2, InStr(1, strText, ")") - 2)
        End If
    End If
    
    gstrSQL = "Select A.ID,A.����,A.����,A.����" & _
             "   FROM ���ղ��� A WHERE A.����=[1] And (" & _
             " A.���� like [2] || '%'  or  A.���� like [2] || '%'   or  A.���� like [2] || '%')" & _
             " Order by A.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint����, strText)
    If rsTemp.RecordCount = 0 Then
        MsgBox "�����ڸò��֣����������룡", vbInformation, gstrSysName
        txt������Ϣ(Index).Text = lbl������Ϣ(Index).Tag
        zlControl.TxtSelAll txt������Ϣ(Index)
        Exit Sub
    Else
        '����ѡ����
        If rsTemp.RecordCount > 1 Then
            '�����ֶδ���3�ģ���ʹֻ��һ����¼�ѸöԻ�����ʾ�������Ա����û��õ��������Ϣ
            blnReturn = frmListSel.ShowSelect(TYPE_������, rsTemp, "ID", "ҽ������ѡ��", "��ѡ��ҽ�����֣�")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '��¼����û�п�ѡ�������
        txt������Ϣ(Index).Text = lbl������Ϣ(Index).Tag
        zlControl.TxtSelAll txt������Ϣ(Index)
    Else
        '�϶����м�¼����
        txt������Ϣ(Index).Tag = rsTemp!���� & "|" & rsTemp!ID
        txt������Ϣ(Index).Text = "(" & rsTemp!���� & ")" & rsTemp!����
        lbl������Ϣ(Index).Tag = txt������Ϣ(Index).Text '���ڻָ���ʾ
            
        If txt������Ϣ(��Ժ����).Tag = "" Then
            txt������Ϣ(��Ժ����).Text = "(" & rsTemp!���� & ")" & rsTemp!����
            txt������Ϣ(��Ժ����).Tag = rsTemp!���� & "|" & rsTemp!ID
            lbl������Ϣ(��Ժ����).Tag = txt������Ϣ(��Ժ����).Text
        End If
        
        If Index = 0 Then
            txt������Ϣ(1).SetFocus
        Else
            txt����֢.SetFocus
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
