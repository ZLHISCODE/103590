VERSION 5.00
Begin VB.Form frmIdentify�ɶ��ϳ� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   Icon            =   "frmIdentify�ɶ��ϳ�.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CmdSelect 
      Caption         =   "��"
      Height          =   285
      Index           =   0
      Left            =   2580
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   330
      Width           =   285
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3330
      TabIndex        =   8
      Top             =   870
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3330
      TabIndex        =   7
      Top             =   420
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   3120
      TabIndex        =   9
      Top             =   -180
      Width           =   30
   End
   Begin VB.ComboBox Cbo�Ա� 
      Height          =   300
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   1
      Left            =   1140
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1110
      Width           =   1725
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   0
      Left            =   1140
      MaxLength       =   20
      TabIndex        =   1
      Top             =   330
      Width           =   1725
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&A)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   420
      TabIndex        =   5
      Top             =   1170
      Width           =   630
   End
   Begin VB.Label Lbl�Ա� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�Ա�(&S)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   420
      TabIndex        =   3
      Top             =   780
      Width           =   630
   End
   Begin VB.Label Lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   420
      TabIndex        =   0
      Top             =   390
      Width           =   630
   End
End
Attribute VB_Name = "frmIdentify�ɶ��ϳ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum �ı�Enum
    Text���� = 0
    Text���� = 1
End Enum

Private Enum ѡ��Enum
    Select���� = 0
End Enum

Dim mstrIdentify As String
Dim mbytType As Byte
Dim mlng����ID As Long

Public Function ShowCard(Optional bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�����ҽ�����˵������Ϣ
'������bytType-ʶ�����ͣ�0-���1-סԺ
'���أ�
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�
    Dim rsTemp As New ADODB.Recordset
    mbytType = bytType
    mlng����ID = lng����ID
    
    cbo�Ա�.Clear
    gstrSQL = "select ����,���� from �Ա� order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    Do Until rsTemp.EOF
        cbo�Ա�.AddItem rsTemp("����") & "." & rsTemp("����")
        rsTemp.MoveNext
    Loop
    cbo�Ա�.ListIndex = 0
    rsTemp.Close
    
    frmIdentify�ɶ��ϳ�.Show vbModal
    ShowCard = mstrIdentify
End Function

Private Sub cmdCancel_Click()
    mstrIdentify = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim strIdentify As String, strAddition As String
    Dim lng����ID As Long, lng���� As Long
    
    '���������ݵ���ȷ��
    If IsValid() = False Then
        Exit Sub
    End If
    
    '�õ��������
    lng���� = 0
    
    '��鲡��״̬
    gstrSQL = "select nvl(��ǰ״̬,0) as ״̬ from �����ʻ� where ����=[1] and ����=[2] and ����ID=[3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_�ɶ��ϳ�, lng����, mlng����ID)
    
    If rsTemp.RecordCount > 0 Then
        If rsTemp("״̬") > 0 Then
            MsgBox "�ò����Ѿ���Ժ������ͨ�������֤��", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '�����ַ���
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�
    strIdentify = ""                         '0����
    strIdentify = strIdentify & ";"          '1ҽ����
    strIdentify = strIdentify & ";"                                    '2����
    strIdentify = strIdentify & ";" & Trim(txtEdit(Text����).Text)     '3����
    strIdentify = strIdentify & ";" & Replace(GetTextFromCombo(cbo�Ա�, True), "'", "") '4�Ա�
    strIdentify = strIdentify & ";" & "" '5��������
    strIdentify = strIdentify & ";" & ""    '6���֤
    strIdentify = strIdentify & ";" & ""  '7.��λ����(����)
    strAddition = ";" & lng����                                 '8.���Ĵ���
    strAddition = strAddition & ";"                             '9.˳���
    strAddition = strAddition & ";" & ""       '10��Ա���
    strAddition = strAddition & ";" & ""  '11�ʻ����
    strAddition = strAddition & ";0"                            '12��ǰ״̬
    strAddition = strAddition & ";" & ""     '13����ID
    strAddition = strAddition & ";" & "1" '14��ְ(1,2,3)
    strAddition = strAddition & ";" & "" '15����֤��
    strAddition = strAddition & ";" & Val(txtEdit(Text����)) '16�����
    strAddition = strAddition & ";"                             '17�Ҷȼ�
    strAddition = strAddition & ";" & 0       '18�ʻ������ۼ�
    strAddition = strAddition & ";" & 0       '19�ʻ�֧���ۼ�
    strAddition = strAddition & ";" & 0        '20���깤���ܶ�
    strAddition = strAddition & ";0;0"      '21סԺ�����ۼ�
    
    lng����ID = BuildPatiInfo(mbytType, strIdentify & strAddition, mlng����ID, TYPE_�ɶ��ϳ�)
    '���ظ�ʽ:�м���벡��ID
    If lng����ID > 0 Then
        mstrIdentify = strIdentify & ";" & lng����ID & strAddition
        'ǿ�ưѵǼ�˳��š����µ�ҽ��������
        gstrSQL = "ZL_�����ʻ�_�޸�ҽ����(" & lng����ID & "," & TYPE_�ɶ��ϳ� & _
                    ",NULL,'" & lng����ID & "',NULL)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    Unload Me
End Sub

Private Function IsValid() As Boolean
'���ܣ�������ݵ���ȷ��
    Dim lngIndex As Long
    
    For lngIndex = txtEdit.LBound To txtEdit.UBound
        If txtEdit(lngIndex).Enabled = True Then
            If zlCommFun.StrIsValid(txtEdit(lngIndex), txtEdit(lngIndex).MaxLength) = False Then
                zlControl.TxtSelAll txtEdit(lngIndex)
                txtEdit(lngIndex).SetFocus
                Exit Function
            End If
        End If
        
        If lngIndex = Text���� Then
            If IsNumeric(txtEdit(lngIndex).Text) = False Then
                MsgBox "������Ϸ�����ֵ��", vbInformation, gstrSysName
                zlControl.TxtSelAll txtEdit(lngIndex)
                txtEdit(lngIndex).SetFocus
                Exit Function
            End If
            If Val(txtEdit(lngIndex).Text) < 0 Or Val(txtEdit(lngIndex).Text) > 200 Then
                MsgBox "���䲻��С��0���Ҳ��ܳ���200��", vbInformation, gstrSysName
                zlControl.TxtSelAll txtEdit(lngIndex)
                txtEdit(lngIndex).SetFocus
                Exit Function
            End If
        End If
        If (lngIndex = Text���� Or lngIndex = Text����) And Trim(txtEdit(lngIndex).Text) = "" Then
            MsgBox "���������䶼����Ϊ�ա�", vbInformation, gstrSysName
            zlControl.TxtSelAll txtEdit(lngIndex)
            txtEdit(lngIndex).SetFocus
            Exit Function
        End If
    Next
    
    IsValid = True
End Function

Private Sub cmdSelect_Click(Index As Integer)
    Dim rsTemp As ADODB.Recordset
    
    Select Case Index
        Case Select����
            gstrSQL = " Select A.����ID as ID,A.����,A.ҽ����,B.����,B.�Ա�,B.����,B.��������,B.���֤��,C.��� as ����ID " & _
                    " ,A.��Ա���,A.��λ����,A.����ID,D.���� as ����,A.��ְ as ��ְID,A.����֤��" & _
                    " From �����ʻ� A,������Ϣ B,��������Ŀ¼ C,���ղ��� D" & _
                    "  where A.����ID=B.����ID and A.����=" & TYPE_�ɶ��ϳ� & _
                    "  and A.����=C.���� and A.����=C.��� and A.����ID=D.ID(+)"
            
            Call Get�ʻ����
            zlControl.TxtSelAll txtEdit(Text����)
            txtEdit(Text����).SetFocus
    End Select
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    Select Case Index
        Case Text����
            zlCommFun.OpenIme True
        Case Else
            zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strCode As String
    Dim str���� As String
    Dim rsTemp As New ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        
        KeyAscii = 0  '��������
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub cbo�Ա�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub Get�ʻ����()
'���Ѿ����ڵļ�¼�ж����ʻ���Ϣ
    Dim rs�ʻ� As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim lngIndex As Long
    
    
    Set rs�ʻ� = frmPubSel.ShowSelect(Me, gstrSQL, 0, "�����ʻ�", , txtEdit(Text����).Text, "", False, True)
    If Not rs�ʻ� Is Nothing Then
    
        '�������õ�����
        mlng����ID = rs�ʻ�!ID
        txtEdit(Text����).Text = IIf(IsNull(rs�ʻ�("����")), "", rs�ʻ�("����"))
        txtEdit(Text����).Text = IIf(IsNull(rs�ʻ�("����")), "", rs�ʻ�("����"))
        
        Call SetComboByText(cbo�Ա�, IIf(IsNull(rs�ʻ�("�Ա�")), "", rs�ʻ�("�Ա�")), True)
    End If
End Sub



