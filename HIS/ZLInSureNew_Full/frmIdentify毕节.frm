VERSION 5.00
Begin VB.Form frmIdentify�Ͻ� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   Icon            =   "frmIdentify�Ͻ�.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmd���� 
      Caption         =   "��"
      Height          =   270
      Left            =   3510
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3390
      Width           =   285
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1530
      TabIndex        =   19
      Top             =   3390
      Width           =   2265
   End
   Begin VB.CommandButton cmdChangePass 
      Caption         =   "������(&G)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4320
      TabIndex        =   24
      Top             =   3270
      Width           =   1100
   End
   Begin VB.TextBox txtͳ�ﱨ���ۼ� 
      Height          =   300
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3030
      Width           =   2265
   End
   Begin VB.TextBox txtסԺ���� 
      Height          =   300
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2670
      Width           =   2265
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4320
      TabIndex        =   22
      Top             =   240
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4320
      TabIndex        =   23
      Top             =   690
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   4050
      TabIndex        =   21
      Top             =   -120
      Width           =   45
   End
   Begin VB.TextBox txt�ʻ���� 
      Height          =   300
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2310
      Width           =   2265
   End
   Begin VB.TextBox txt�μӹ������� 
      Height          =   300
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1950
      Width           =   2265
   End
   Begin VB.TextBox txt�������� 
      Height          =   300
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1590
      Width           =   2265
   End
   Begin VB.TextBox txt�Ա� 
      Height          =   300
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1230
      Width           =   1245
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   870
      Width           =   2265
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1530
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   150
      Width           =   2265
   End
   Begin VB.TextBox txt��ᱣ�Ϻ� 
      Height          =   300
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   510
      Width           =   2265
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   810
      TabIndex        =   18
      Top             =   3450
      Width           =   630
   End
   Begin VB.Label lblͳ�ﱨ���ۼ� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ͳ�ﱨ���ۼ�(&L)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   90
      TabIndex        =   16
      Top             =   3090
      Width           =   1350
   End
   Begin VB.Label lblסԺ���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "סԺ����(&Z)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   14
      Top             =   2730
      Width           =   990
   End
   Begin VB.Label lbl�ʻ���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�ʻ����(&Y)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   12
      Top             =   2370
      Width           =   990
   End
   Begin VB.Label lbl�μӹ������� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������(&J)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   10
      Top             =   2010
      Width           =   990
   End
   Begin VB.Label lbl�������� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������(&B)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   8
      Top             =   1650
      Width           =   990
   End
   Begin VB.Label lbl�Ա� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�Ա�(&S)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   810
      TabIndex        =   6
      Top             =   1290
      Width           =   630
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   810
      TabIndex        =   4
      Top             =   930
      Width           =   630
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "IC������(&P)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   0
      Top             =   210
      Width           =   990
   End
   Begin VB.Label lbl��ᱣ�Ϻ� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ᱣ�Ϻ�(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   270
      TabIndex        =   2
      Top             =   570
      Width           =   1170
   End
End
Attribute VB_Name = "frmIdentify�Ͻ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng����ID As Long
Private mbytType As Byte
Private mstrReturn As String
Private mstrICData As String
'ֻ��ʹ����IC������ĲŽ�������У��,�������޸�����
'��Ҫ�Զ���״̬���м��,����������������Ǽ�
'�����Ч����������һ��,������ʹ��,������Ϊ�Ͽ�

Public Function GetIdentify(ByVal bytType As Byte, Optional lng����ID As Long) As String
    mlng����ID = lng����ID
    mbytType = bytType
    mstrReturn = ""
    mstrICData = ""
    Me.Show 1
    GetIdentify = mstrReturn
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChangePass_Click()
    Dim strOldPass As String, strNewPass As String
    Dim strCardData As String
    '��������ѿ��ƣ�ֻ����ȷ������������޸����룬��ˣ�������ľ����벻���ж�
    strNewPass = frm�޸�����.ChangePassword(Me.txt����.Text, strOldPass)
    If strOldPass = strNewPass Then Exit Sub
    
    If strOldPass <> IC_Data_�Ͻ�.����IC������ Then
        MsgBox "����ľ������뿨�����벻����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    IC_Data_�Ͻ�.����IC������ = strNewPass
    Call ����ת��_�Ͻ�(strCardData, False)
    Call gobjCenter.IC_ChangePass(strCardData)
End Sub

Private Sub cmdOK_Click()
    Dim lng���� As Long
    Dim strIdentify As String, strAddition As String
    Dim rsTmp As New ADODB.Recordset
    '����Ƿ���Ժ�������Ժ�����ֹ����
    gstrSQL = "Select Nvl(��ǰ״̬,0) ״̬ From �����ʻ� Where ����=[1] And ҽ����=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_�Ͻ�, CStr(txt��ᱣ�Ϻ�.Text))
    If rsTmp.RecordCount = 1 Then
        If rsTmp!״̬ = 1 Then
            MsgBox "��ǰ����Ŀǰ����Ժ���ƣ���������Ժ�����ڼ��ٴξ��", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '�������Ժ�������ѡ����
    If mbytType = 1 Then
        If Val(txt����.Tag) = 0 Then
            MsgBox "��ѡ����Ժ���֣�", vbInformation, gstrSysName
            txt����.SetFocus
            Exit Sub
        End If
    End If
    
    lng���� = GetAge(Format(zlDatabase.Currentdate, "yyyy-MM-dd"), Me.txt��������.Text)
    '����������Ϣ
    '�����ַ���
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�
    strIdentify = IC_Data_�Ͻ�.��Ч����                         '0����
    strIdentify = strIdentify & ";" & txt��ᱣ�Ϻ�.Text        '1ҽ����
    strIdentify = strIdentify & ";"                             '2����
    strIdentify = strIdentify & ";" & txt����.Text              '3����
    strIdentify = strIdentify & ";" & txt�Ա�.Text              '4�Ա�
    strIdentify = strIdentify & ";" & txt��������.Text          '5��������
    strIdentify = strIdentify & ";"                             '6���֤
    strIdentify = strIdentify & ";" & IC_Data_�Ͻ�.��λ����     '7.��λ����(����)
    strAddition = ";0"                                          '8.���Ĵ���
    strAddition = strAddition & ";"                             '9.˳���
    strAddition = strAddition & ";"                             '10��Ա���
    strAddition = strAddition & ";" & Val(txt�ʻ����.Text)     '11�ʻ����
    strAddition = strAddition & ";0"                            '12��ǰ״̬
    strAddition = strAddition & ";" & Val(txt����.Tag)         '13����ID
    strAddition = strAddition & ";1"                            '14��ְ(1,2,3)
    strAddition = strAddition & ";"                             '15����֤��
    strAddition = strAddition & ";"                             '16�����
    strAddition = strAddition & ";"                             '17�Ҷȼ�
    strAddition = strAddition & ";" & Val(txt�ʻ����.Text)     '18�ʻ������ۼ�
    strAddition = strAddition & ";0"                            '19�ʻ�֧���ۼ�
    strAddition = strAddition & ";" & Val(txtͳ�ﱨ���ۼ�.Text) '20���깤���ܶ�
    strAddition = strAddition & ";" & Val(txtסԺ����.Text)     '21סԺ�����ۼ�

    mlng����ID = BuildPatiInfo(0, strIdentify & strAddition, mlng����ID, TYPE_�Ͻ�)
    '���ظ�ʽ:�м���벡��ID
    If mlng����ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng����ID & strAddition
    Else
        Exit Sub
    End If
    
    '�����������������ʱҪʹ�ÿ������Բ���������סԺ����Ե�����
    If mbytType = 1 Then Call IC_End(True)
    
    Unload Me
End Sub

Private Sub cmd����_Click()
    Dim blnReturn As Boolean
    Dim rsTmp As New ADODB.Recordset
        
    gstrSQL = " Select ID,���ִ��� As ����,��������,��ҽ����,�������,�����Ը�����,�����𸶽�� " & _
              " From ����Ŀ¼��"
    With rsTmp
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
    End With
    
    blnReturn = frmListSel.ShowSelect(TYPE_�Ͻ�, rsTmp, "ID", "ҽ������ѡ��", "��ѡ��ҽ�����֣�")
    If blnReturn = False Then
        '��¼����û�п�ѡ�������
        txt����.Text = lbl����.Tag
        zlControl.TxtSelAll txt����
        Exit Sub
    Else
        '�϶����м�¼����
        txt����.Tag = rsTmp!ID
        txt����.Text = "(" & rsTmp!���� & ")" & rsTmp!��������
        lbl����.Tag = txt����.Text '���ڻָ���ʾ
    End If
End Sub

Private Sub Form_Load()
    txt����.Locked = Not gCominfo_�Ͻ�.blnICPassVerify
    
    Me.txt����.Enabled = (mbytType = 1)
    Me.cmd����.Enabled = (mbytType = 1)
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTmp As New ADODB.Recordset
    Dim dbl��� As Double
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Me.cmdOK.Enabled = False
    cmdChangePass.Enabled = False
    
    If Not gobjCenter.IC_ReadCard(mstrICData) Then Exit Sub
    Call ����ת��_�Ͻ�(mstrICData, True)
    
    If gCominfo_�Ͻ�.blnICPassVerify Then
        If txt����.Text <> IC_Data_�Ͻ�.����IC������ Then
            MsgBox "IC������������������룡", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '����������Ҫ���м���л�ȡ
    gstrSQL = "Select סԺ����,�ʻ�����,����ԭ��,��Ч����,����סԺ����,����ʱ��,����˵��,��� " & _
        " From �����ʻ����� Where ��ᱣ�Ϻ�='" & IC_Data_�Ͻ�.��ᱣ�Ϻ� & "'"
    If gCominfo_�Ͻ�.blnOnLine Then
        Call gobjCenter.InitConnect("")
        If Not gobjCenter.GetRecordset(gstrSQL, rsTmp) Then
            Call IC_End(True)
            Call gobjCenter.CloseConnector
            Exit Sub
        End If
    Else
        If rsTmp.State = 1 Then rsTmp.Close
        rsTmp.Open gstrSQL, gcnGYBJYB
    End If
    
    With rsTmp
        If .RecordCount = 0 Then
            Call IC_End(True)
            MsgBox "û���ָò��˵���Ч��¼������������ϵ��", vbInformation, gstrSysName
            Exit Sub
        End If
        If Nvl(!�ʻ�����, "��") = "��" Then
            Call IC_End(True)
            MsgBox "�ò��˵��ʻ��Ѿ������ᣬֻ������ͨ���˵���ݰ���" & vbCrLf & "����ԭ��" & Nvl(!����ԭ��) & vbCrLf & "����˵����" & Nvl(!����˵��) & vbCrLf & "����ʱ�䣺" & Nvl(!����ʱ��), vbInformation, gstrSysName
            Exit Sub
        End If
        If Nvl(IC_Data_�Ͻ�.��Ч����, 0) <> Nvl(!��Ч����, 0) Then
            Call IC_End(True)
            MsgBox "��ǰ��IC��Ƭ��һ����Ч�Ŀ���", vbInformation, gstrSysName
            Exit Sub
        End If
        dbl��� = Nvl(!���, 0)
        Me.txtסԺ����.Text = Format(Nvl(!סԺ����, 0), "#####0;-#####0; ;")
        Me.txtͳ�ﱨ���ۼ�.Text = Format(Nvl(!����סԺ����, 0), "#####0.00;-#####0.00; ;")
    End With
    
    If gCominfo_�Ͻ�.blnOnLine Then gobjCenter.CloseConnector
    
    '��IC��������ʾ����
    Me.txt��ᱣ�Ϻ�.Text = IC_Data_�Ͻ�.��ᱣ�Ϻ�
    Me.txt����.Text = IC_Data_�Ͻ�.����
    Me.txt�Ա�.Text = IC_Data_�Ͻ�.�Ա�
    Me.txt��������.Text = Replace(IC_Data_�Ͻ�.��������, ".", "-")
    Me.txt�μӹ�������.Text = Replace(IC_Data_�Ͻ�.�μӹ�������, ".", "-")
    Me.txt�ʻ����.Text = Format(IC_Data_�Ͻ�.�����ʻ����, "#####0.00;-#####0.00; ;")
    
    '������ѻ�ϵͳ���ҵ������ѹ����Կ������Ϊ׼������������Ϊ׼
    If gCominfo_�Ͻ�.blnOnLine Or (gCominfo_�Ͻ�.blnOnLine = False And Format(zlDatabase.Currentdate, "yyyy.MM.dd") <> IC_Data_�Ͻ�.����������) Then
        Me.txt�ʻ����.Text = Format(dbl���, "#####0.00;-#####0.00; ;")
    End If
    IC_Data_�Ͻ�.�����ʻ���� = Val(Me.txt�ʻ����.Text)
    
    cmdOK.Enabled = True
    cmdChangePass.Enabled = gCominfo_�Ͻ�.blnICPassVerify
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim rsTmp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt����.Text = "" And txt����.Tag <> "" Then Exit Sub
    
    On Error GoTo errHandle
    
    strText = UCase(txt����.Text)
    If InStr(1, strText, "(") <> 0 Then
        If InStr(1, strText, ")") <> 0 Then
            strText = Mid(strText, 2, InStr(1, strText, ")") - 2)
        End If
    End If
    gstrSQL = " Select ID,���ִ��� As ����,��������,��ҽ����,�������,�����Ը�����,�����𸶽�� " & _
              " From ����Ŀ¼�� A" & _
              " Where (" & zlCommFun.GetLike("A", "���ִ���", strText) & " or " & zlCommFun.GetLike("A", "��������", strText) & " or zlspellcode(��������) Like '" & strText & "%')"
    With rsTmp
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
    End With
    
    If rsTmp.RecordCount = 0 Then
        MsgBox "�����ڸò��֣����������룡", vbInformation, gstrSysName
        txt����.Text = lbl����.Tag
        zlControl.TxtSelAll txt����
        Exit Sub
    Else
        '����ѡ����
        If rsTmp.RecordCount > 1 Then
            '�����ֶδ���3�ģ���ʹֻ��һ����¼�ѸöԻ�����ʾ�������Ա����û��õ��������Ϣ
            blnReturn = frmListSel.ShowSelect(TYPE_�Ͻ�, rsTmp, "ID", "ҽ������ѡ��", "��ѡ��ҽ�����֣�")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '��¼����û�п�ѡ�������
        txt����.Text = lbl����.Tag
        zlControl.TxtSelAll txt����
        Exit Sub
    Else
        '�϶����м�¼����
        txt����.Tag = rsTmp!ID
        txt����.Text = "(" & rsTmp!���� & ")" & rsTmp!��������
        lbl����.Tag = txt����.Text '���ڻָ���ʾ
    End If
    
    Call zlCommFun.PressKey(vbKeyTab)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub IC_End(Optional ByVal blnPull As Boolean = False)
    '�ڴ�IC�豸����������Ƿ����������������ڵ�����رն˿�
    Call gobjCenter.IC_PullCard
    If blnPull Then Exit Sub
    
    Call gobjCenter.IC_CloseDevice
End Sub
