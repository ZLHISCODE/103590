VERSION 5.00
Begin VB.Form frmIdentify���������� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   5625
   ClientLeft      =   5310
   ClientTop       =   3135
   ClientWidth     =   7725
   Icon            =   "frm����������.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4920
      TabIndex        =   46
      Top             =   5160
      Width           =   1100
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6180
      TabIndex        =   47
      Top             =   5160
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Caption         =   "�ۼ���Ϣ(Ԫ)"
      Height          =   1965
      Left            =   150
      TabIndex        =   29
      Top             =   3060
      Width           =   7425
      Begin VB.TextBox txt�ʻ���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1590
         TabIndex        =   31
         Tag             =   "13#"
         Top             =   300
         Width           =   1845
      End
      Begin VB.TextBox txtסԺ��ʷδ���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5190
         TabIndex        =   45
         Tag             =   "25#"
         Top             =   1470
         Width           =   1845
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1590
         TabIndex        =   43
         Tag             =   "24#"
         Top             =   1470
         Width           =   1845
      End
      Begin VB.TextBox txt����֢״�� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5190
         TabIndex        =   41
         Tag             =   "23#"
         Top             =   1080
         Width           =   1845
      End
      Begin VB.TextBox txt���Ϲ���Ա���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1590
         TabIndex        =   39
         Tag             =   "17#"
         Top             =   1080
         Width           =   1845
      End
      Begin VB.TextBox txt�ز�����ҽ���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5190
         TabIndex        =   37
         Tag             =   "16#"
         Top             =   690
         Width           =   1845
      End
      Begin VB.TextBox txt�ز����ۼ� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1590
         TabIndex        =   35
         Tag             =   "15#"
         Top             =   690
         Width           =   1845
      End
      Begin VB.TextBox txtͳ��֧���ۼ� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5190
         TabIndex        =   33
         Tag             =   "14#"
         Top             =   300
         Width           =   1845
      End
      Begin VB.Label lbl�ʻ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ʻ����"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   810
         TabIndex        =   30
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblסԺ��ʷδ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��ʷδ����"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3840
         TabIndex        =   44
         Top             =   1530
         Width           =   1260
      End
      Begin VB.Label lbl�ز���ʷδ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ز���ʷδ����"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   210
         TabIndex        =   42
         Top             =   1530
         Width           =   1260
      End
      Begin VB.Label lbl����֢״�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����֢״��"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4020
         TabIndex        =   40
         Top             =   1140
         Width           =   1080
      End
      Begin VB.Label lbl���Ϲ���Ա���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���Ϲ���Ա����"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   210
         TabIndex        =   38
         Top             =   1140
         Width           =   1260
      End
      Begin VB.Label lbl�ز�����ҽ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ز�����ҽ����"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3840
         TabIndex        =   36
         Top             =   750
         Width           =   1260
      End
      Begin VB.Label lbl�ز������� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ز�������"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   34
         Top             =   750
         Width           =   1080
      End
      Begin VB.Label lblͳ��֧���ۼ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ͳ��֧��"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4380
         TabIndex        =   32
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "������Ϣ"
      Height          =   2745
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   7425
      Begin VB.TextBox txt����֢ 
         Height          =   300
         Left            =   4770
         TabIndex        =   28
         Top             =   2250
         Width           =   2355
      End
      Begin VB.TextBox txt������Ϣ 
         Height          =   300
         Left            =   1080
         TabIndex        =   25
         Top             =   2250
         Width           =   2055
      End
      Begin VB.TextBox txtסԺ���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4770
         TabIndex        =   22
         Tag             =   "18"
         Top             =   1860
         Width           =   555
      End
      Begin VB.CommandButton cmd������Ϣ 
         Caption         =   "��"
         Height          =   300
         Left            =   3150
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2250
         Width           =   285
      End
      Begin VB.TextBox txt���֤�� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1080
         TabIndex        =   12
         Tag             =   "0"
         Top             =   1080
         Width           =   2355
      End
      Begin VB.ComboBox cboҽ����� 
         Height          =   300
         Left            =   4770
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   300
         Width           =   2355
      End
      Begin VB.TextBox txt��λ���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1080
         TabIndex        =   20
         Tag             =   "8+9"
         Top             =   1860
         Width           =   2355
      End
      Begin VB.TextBox txtͳ������ 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4770
         TabIndex        =   18
         Tag             =   "7-"
         Top             =   1470
         Width           =   2355
      End
      Begin VB.CheckBox chk���ܹ���Ա���� 
         Alignment       =   1  'Right Justify
         Caption         =   "���ܹ���Ա����"
         Enabled         =   0   'False
         Height          =   225
         Left            =   5490
         TabIndex        =   23
         Tag             =   "6"
         Top             =   1890
         Width           =   1635
      End
      Begin VB.TextBox txt�������� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1080
         TabIndex        =   16
         Tag             =   "5-"
         Top             =   1470
         Width           =   2355
      End
      Begin VB.TextBox txt��Ա��� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4770
         TabIndex        =   14
         Tag             =   "6-"
         Top             =   1080
         Width           =   2355
      End
      Begin VB.TextBox txt���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6570
         TabIndex        =   10
         Tag             =   "12"
         Top             =   690
         Width           =   555
      End
      Begin VB.TextBox txt�Ա� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4770
         TabIndex        =   8
         Tag             =   "2"
         Top             =   690
         Width           =   1095
      End
      Begin VB.TextBox txt���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1080
         TabIndex        =   6
         Tag             =   "1"
         Top             =   690
         Width           =   1365
      End
      Begin VB.TextBox txtҽ��֤�� 
         Height          =   300
         Left            =   1080
         TabIndex        =   2
         Top             =   300
         Width           =   2355
      End
      Begin VB.Label lbl����֢ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����֢"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4140
         TabIndex        =   27
         Top             =   2310
         Width           =   540
      End
      Begin VB.Label lbl������Ϣ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������Ϣ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   300
         TabIndex        =   24
         Top             =   2310
         Width           =   720
      End
      Begin VB.Label lblסԺ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3960
         TabIndex        =   21
         Top             =   1920
         Width           =   720
      End
      Begin VB.Label lbl���֤�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   300
         TabIndex        =   11
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lblҽ����� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ�����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3960
         TabIndex        =   3
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl��λ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��λ����"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   300
         TabIndex        =   19
         Top             =   1920
         Width           =   720
      End
      Begin VB.Label lblͳ������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ͳ������"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3960
         TabIndex        =   17
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label lbl�������� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   300
         TabIndex        =   15
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label lbl��Ա��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Ա���"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3960
         TabIndex        =   13
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6120
         TabIndex        =   9
         Top             =   750
         Width           =   360
      End
      Begin VB.Label lbl�Ա� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4320
         TabIndex        =   7
         Top             =   750
         Width           =   360
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   660
         TabIndex        =   5
         Top             =   750
         Width           =   360
      End
      Begin VB.Label lblҽ��֤�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ��֤��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   300
         TabIndex        =   1
         Top             =   360
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmIdentify����������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType As Byte            '0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�
Private mlng����ID As Long
Private mstrReturn As String

Private Sub cboҽ�����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmd������Ϣ_Click()
    Dim str���ֱ��� As String, str����֢ As String
    Dim rsTemp As New ADODB.Recordset
    str���ֱ��� = txt������Ϣ.Tag
    str����֢ = txt����֢.Text
    If frm����ѡ��_����������.ShowSelect(Me, Me.cboҽ�����.ItemData(Me.cboҽ�����.ListIndex), str���ֱ���, str����֢) = True Then
        gstrSQL = "Select ���� From ���ղ��� Where ����=[1] And ����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ֵ�����", str���ֱ���, TYPE_����������)
        txt������Ϣ.Tag = str���ֱ���
        txt������Ϣ.Text = "(" & str���ֱ��� & ")" & rsTemp!����
        lbl������Ϣ.Tag = txt������Ϣ.Text '���ڻָ���ʾ
        txt����֢.SetFocus
    End If
End Sub

Private Sub txt����֢_GotFocus()
    Call zlControl.TxtSelAll(txt����֢)
End Sub

Private Sub txt����֢_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt������Ϣ_GotFocus()
    Call zlControl.TxtSelAll(txt������Ϣ)
End Sub

Private Sub txt������Ϣ_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    
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
             "   FROM ���ղ��� A WHERE A.����=[1] And (" & _
             "A.���� like [2] || '%' or A.���� like [2] || '%' or A.���� like [2] || '%')" & Get�������
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_����������, strText)
    If rsTemp.RecordCount = 0 Then
        MsgBox "�����ڸò��֣����������룡", vbInformation, gstrSysName
        txt������Ϣ.Text = lbl������Ϣ.Tag
        zlControl.TxtSelAll txt������Ϣ
        Exit Sub
    Else
        '����ѡ����
        If rsTemp.RecordCount > 1 Then
            '�����ֶδ���3�ģ���ʹֻ��һ����¼�ѸöԻ�����ʾ�������Ա����û��õ��������Ϣ
            blnReturn = frmListSel.ShowSelect(TYPE_����������, rsTemp, "ID", "ҽ������ѡ��", "��ѡ��ҽ�����֣�")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '��¼����û�п�ѡ�������
        txt������Ϣ.Text = lbl������Ϣ.Tag
        zlControl.TxtSelAll txt������Ϣ
        Exit Sub
    Else
        '�϶����м�¼����
        txt������Ϣ.Tag = rsTemp!����
        txt������Ϣ.Text = "(" & rsTemp!���� & ")" & rsTemp!����
        lbl������Ϣ.Tag = txt������Ϣ.Text '���ڻָ���ʾ
    End If
    
    Call zlCommFun.PressKey(vbKeyTab)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    Dim lng����ID As Long
    Dim str�������� As String
    Dim strIdentify As String, strAddition As String
    Dim rsTemp As New ADODB.Recordset
    
    If Trim(txtҽ��֤��.Text) = "" Then
        MsgBox "��û������ҽ��֤�ţ�", vbInformation, gstrSysName
        txtҽ��֤��.SetFocus
        Exit Sub
    End If
    If Trim(txt����.Text) = "" Then
        MsgBox "��û�л�ȡ��ҽ�����˵������Ϣ������ͨ����֤��", vbInformation, gstrSysName
        txtҽ��֤��.SetFocus
        Exit Sub
    End If
    If txt������Ϣ.Tag = "" Then
        MsgBox "�����벡�˵ļ�����Ϣ��", vbInformation, gstrSysName
        txt������Ϣ.SetFocus
        Exit Sub
    End If
    
    '��ȡ����ID
    gstrSQL = "Select ID From ���ղ��� Where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ղ���", CStr(txt������Ϣ.Tag))
    If Not rsTemp.EOF Then
        lng����ID = rsTemp!ID
    End If
    
    If mbytType <> 2 Then
        '��鲡��״̬
        gstrSQL = "select nvl(��ǰ״̬,0) as ״̬ from �����ʻ� where ����=[1] and ҽ����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_����������, CStr(txtҽ��֤��.Text))
        If rsTemp.RecordCount > 0 Then
            If rsTemp("״̬") > 0 Then
                MsgBox "�ò����Ѿ���Ժ������ͨ�������֤��", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        '�ݲ�֧�ֹҺ�
        If mbytType = 3 Then
            Unload Me
            Exit Sub
        End If
    Else
        '�����������סԺ�ģ�ֻ��ˢ����ʾһ�����ݶ��ѣ�������
        Unload Me
        Exit Sub
    End If
    
    '�������֤�ŵõ���������
    If Len(txt���֤��) >= 15 Then
        If Len(txt���֤��) = 15 Then
            str�������� = "19" & Mid(txt���֤��, 7, 2) & "-" & Mid(txt���֤��, 9, 2) & "-" & Mid(txt���֤��, 11, 2)
        ElseIf Len(txt���֤��) = 18 Then
            str�������� = Mid(txt���֤��, 7, 4) & "-" & Mid(txt���֤��, 11, 2) & "-" & Mid(txt���֤��, 13, 2)
        End If
        If Not IsDate(str��������) Then str�������� = ""
    End If
    
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�
    strIdentify = txtҽ��֤��.Text                              '0����
    strIdentify = strIdentify & ";" & txtҽ��֤��.Text          '1ҽ����
    strIdentify = strIdentify & ";"                             '2����
    strIdentify = strIdentify & ";" & txt����.Text              '3����
    strIdentify = strIdentify & ";" & txt�Ա�.Text              '4�Ա�
    strIdentify = strIdentify & ";" & str��������               '5��������
    strIdentify = strIdentify & ";" & txt���֤��.Text          '6���֤
    strIdentify = strIdentify & ";" & txt��λ����.Text          '7.��λ����(����)
    strAddition = ";0"                                          '8.���Ĵ���
    strAddition = strAddition & ";"                             '9.˳���
    strAddition = strAddition & ";" & Split(txt��Ա���.Text, "-")(0)         '10��Ա���
    strAddition = strAddition & ";" & Val(txt�ʻ����.Text)     '11�ʻ����
    strAddition = strAddition & ";0"                            '12��ǰ״̬
    strAddition = strAddition & ";" & lng����ID                 '13����ID
    strAddition = strAddition & ";1"                            '14��ְ(1,2,3)
    strAddition = strAddition & ";"                             '15����֤��
    strAddition = strAddition & ";"                             '16�����
    strAddition = strAddition & ";"                             '17�Ҷȼ�
    strAddition = strAddition & ";" & Val(txt�ʻ����.Text)     '18�ʻ������ۼ�
    strAddition = strAddition & ";0"                            '19�ʻ�֧���ۼ�
    strAddition = strAddition & ";0"                            '20���깤���ܶ�
    strAddition = strAddition & ";" & Val(txtסԺ����.Text)     '21סԺ�����ۼ�
    
    mlng����ID = BuildPatiInfo(0, strIdentify & strAddition, mlng����ID, TYPE_����������)
    '���ظ�ʽ:�м���벡��ID
    If mlng����ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng����ID & strAddition
    End If
    
    With gComInfo_����������
        .���˱�� = txtҽ��֤��.Text
        .�������� = txt������Ϣ.Tag
        .����֢ = txt����֢.Text
        .ͳ������ = Split(txtͳ������.Text, "-")(0)
        .ҵ������ = Me.cboҽ�����.ItemData(Me.cboҽ�����.ListIndex)
        .�ʻ���� = Val(txt�ʻ����.Text)
    End With
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
'      10  ҽ�����    ҩ�깺ҩ
'      11  ҽ�����    ��ͨ����
'      13  ҽ�����    ���ⲡ����
'      14  ҽ�����    ��������
'      15  ҽ�����    ��������������
'      21  ҽ�����    ��ͨסԺ
'      22  ҽ�����    ת��ҽԺסԺ
'      23  ҽ�����    ��Ժ��ͥ����
'      24  ҽ�����    �����ͥ����
'      25  ҽ�����    ����תסԺ
    With Me.cboҽ�����
        .Clear
        '��������ҵ��
        If mbytType = 0 Or mbytType = 2 Then
'            If glngSys \ 100 <> 1 Then
'                .AddItem "ҩ�깺ҩ"
'                .ItemData(.NewIndex) = 10
'            Else
                .AddItem "��ͨ����"
                .ItemData(.NewIndex) = 11
'            End If
            .AddItem "���ⲡ����"
            .ItemData(.NewIndex) = 13
            .AddItem "��������"
            .ItemData(.NewIndex) = 14
            .AddItem "��������������"
            .ItemData(.NewIndex) = 15
        End If
        '����סԺҵ��
        If mbytType = 1 Or mbytType = 2 Then
            .AddItem "��ͨסԺ"
            .ItemData(.NewIndex) = 21
            .AddItem "ת��ҽԺסԺ"
            .ItemData(.NewIndex) = 22
            .AddItem "�����ͥ����"
            .ItemData(.NewIndex) = 24
            .AddItem "����תסԺ"
            .ItemData(.NewIndex) = 25
        End If
        .ListIndex = 0
    End With
End Sub

Private Sub txtҽ��֤��_GotFocus()
    Call zlControl.TxtSelAll(txtҽ��֤��)
End Sub

Private Sub txtҽ��֤��_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim objControl As Control           'ĳ���ؼ�
    Dim strTag As String            '����ؼ���Tagֵ
    Dim strReturn As String         '�ӿڷ��ص�����
    Dim arrReturn, arrTag                  '�ֽⷵ�ص����ݶ�����������
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtҽ��֤��.Text) = "" Then Exit Sub
    
    '���û�ȡ�α����˵Ļ�����Ϣ�ӿ�
    Call ���ýӿ�_׼��_����������("07", txtҽ��֤��.Text)
    If Not ���ýӿ�_���������� Then Exit Sub
    strReturn = gstrReturn_����������
    arrReturn = Split(strReturn, gstrSplit_Col_����������)
    
    '���    ��������    ����    ����    ˵��
    '1.      string  18      ��ݺ���
    '2.      string  20      ����
    '3.      string  10      �Ա𣬷������ĺ���
    '4.      string  20      ���壬�������ĺ���
    '5.      string  3       ��Ա��𣬼������
    '6.      string  3       �������𣬼������
    '7.      string  3       ���ܹ���Ա������־���������
    '8.      string  3       �α���Ա���ڵ�ͳ�����ţ��������
    '9.      string  14      ��λ���
    '10.     string  50      ��λ����
    '11.     string  3   0   ��ǰ����������𣬼������
    '12.     string  128     ��ǰ����ԭ��
    '13.     number  3   0   ʵ������
    '14.     number  8   2   �˻����
    '15.     number  8   2   ͳ��֧���ۼ�
    '16.     number  8   2   �ز������𸶱�׼֧���ۼ�
    '17.     number  8   2   �ز�����ҽ�����ۼ�
    '18.     number  8   2   ���Ϲ���Ա��Χ��������ۼ�
    '19.     number  3   0   ����סԺ����
    '20.     string  3       ��ǰסԺ״̬���������
    '21.     string  3       ��������֢״�Ƿ���סԺ��־���������
    '22.     number  3       �����״�����֢״סԺ����
    '23.     string  3       �����ѷ�������֢״סԺ���ҽԺ�ȼ����������
    '24.     number  8   2   ��������֢״�����׼�ۼ�
    '25.     number  8   2   �ز���ʷδ�����Ը����
    '26.     number  8   2   סԺ��ʷδ�����Ը����
    
    'Tagֵ�ĺ���˵�������ֱ�ʾ�����±꣬-��ʾ��������Ҫͨ��ת���õ���+��ʾ��ֵ��Ҫ��������Ԫ����϶���,��һ��Ԫ����Ҫ��()��#��ʾ��Ҫ��ʽ��Ϊ��λС����
    For Each objControl In Controls
        '��������ҽ�������� 204-04-07
        If objControl.Name <> "txt������Ϣ" And objControl.Name <> "lbl������Ϣ" Then
            strTag = objControl.Tag
            If Trim(strTag) <> "" Then
                If UCase(TypeName(objControl)) = "TEXTBOX" Then
                    If InStr(1, strTag, "+") <> 0 Then
                        arrTag = Split(strTag, "+")
                        objControl.Text = arrReturn(arrTag(1)) & "(" & arrReturn(arrTag(0)) & ")"
                    ElseIf InStr(1, strTag, "-") <> 0 Then
                        strTag = Replace(strTag, "-", "")
                        objControl.Text = arrReturn(Val(strTag)) & "-" & Exchange(arrReturn(Val(strTag)), strTag)
                    ElseIf InStr(1, strTag, "#") <> 0 Then
                        strTag = Replace(strTag, "#", "")
                        objControl.Text = Format(arrReturn(Val(strTag)), "#####0.00;-#####0.00; ;")
                    Else
                        objControl.Text = arrReturn(Val(strTag))
                    End If
                Else    'ֻ������Checkbox
                    chk���ܹ���Ա����.Value = Val(arrReturn(Val(strTag)))
                End If
            End If
        End If
    Next
    
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Function Exchange(ByVal strData As String, ByVal intOrder As Integer) As String
    Dim intCol As Integer, intCols As Integer
    Dim arrData
    Const str��Ա��� As String = "31,���ʱ�ڲμӹ��������ݸɲ�|32,��սʱ�ڲμӹ��������ݸɲ�|33,���ս��ʱ�ڲμӹ��������ݸɲ�" & _
            "|21,����|22,������ؾ�ס|24,��ְ|25,��ְ��ؾ�ס|26,��ְ�ɷ���Ա|27,�Ʋ���ҵ����ְ��|28,����ǰ�μӸ����������Ϲ���" & _
            "|11,��ְ|12,��ְפ��|14,��ʱ�ù�|15,������ҵ��ת��|16,�پ�ҵ�����¸�ְ��|19,���ݽɷ���Ա"
    Const str�������� As String = "030,ʡ(��)��|033,�൱ʡ(��)��|040,��ʡ(��)��|043,�൱��ʡ(��)��|050,��(��)��|053,�൱��(��)��" & _
            "|060,����(��)��|063,�൱����(��)��|070,����|073,�൱����|080,�൱������|083,�Ƽ�|090,�൱�Ƽ�|093,���Ƽ�|100,�൱���Ƽ�" & _
            "|110,Ա��|200,����"
    Const strͳ������ As String = "03,������|04,��ɿ���|05,������|06,ɳƺ����|07,��������|08,�ϰ���|14,������|15,������|20,���ݸɲ�"
    'Ŀǰֻ����Ա���[6]��ͳ������[7]����������[5]��Ҫת��
    
    Select Case intOrder
    Case 5  '��������
        arrData = Split(str��������, "|")
    Case 6  '��Ա���
        arrData = Split(str��Ա���, "|")
    Case Else 'ͳ������
        arrData = Split(strͳ������, "|")
    End Select
    intCols = UBound(arrData)
    
    For intCol = 0 To intCols
        If strData = Split(arrData(intCol), ",")(0) Then
            Exchange = Split(arrData(intCol), ",")(1)
            Exit For
        End If
    Next
End Function

Public Function GetPatient(Optional bytType As Byte, Optional lng����ID As Long = 0) As String
    mbytType = bytType
    mlng����ID = lng����ID
    mstrReturn = ""
    Me.Show 1
    lng����ID = mlng����ID
    GetPatient = mstrReturn
End Function

Private Function Get�������() As String
    '��������Ｑ��,������ѡ�����Ｑ��
    '�������������,������ѡ����������
    '�����������ѡ������Ｑ�������������Ĳ���
    If Me.cboҽ�����.ItemData(Me.cboҽ�����.ListIndex) = 13 Then
        Get������� = " And ��� In ('1')"
    ElseIf Me.cboҽ�����.ItemData(Me.cboҽ�����.ListIndex) = 14 Then
        Get������� = " And ��� In ('2')"
    Else
        Get������� = " And ��� In ('0','3','4')"
    End If
End Function
