VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmתԺ���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "תԺ����"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10635
   Icon            =   "frmתԺ����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   10635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdȡ�� 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   9000
      TabIndex        =   43
      Top             =   5850
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7710
      TabIndex        =   42
      Top             =   5850
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Caption         =   "��������"
      Height          =   4305
      Left            =   120
      TabIndex        =   19
      Top             =   1410
      Width           =   10365
      Begin VB.TextBox txt��ע 
         Height          =   300
         Left            =   1140
         TabIndex        =   41
         Top             =   3840
         Width           =   8835
      End
      Begin VB.TextBox txt��λ��� 
         Height          =   600
         Left            =   1140
         MultiLine       =   -1  'True
         TabIndex        =   39
         Top             =   3150
         Width           =   8835
      End
      Begin VB.TextBox txtҽԺ��� 
         Height          =   600
         Left            =   1140
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   2460
         Width           =   8835
      End
      Begin VB.TextBox txt������� 
         Height          =   600
         Left            =   1140
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   1770
         Width           =   8835
      End
      Begin MSComCtl2.DTPicker Dtp��Чʱ�� 
         Height          =   285
         Left            =   8700
         TabIndex        =   33
         Top             =   1380
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   86966275
         CurrentDate     =   38063
      End
      Begin VB.TextBox txt������ 
         Height          =   300
         Left            =   5070
         TabIndex        =   31
         Top             =   1380
         Width           =   2115
      End
      Begin VB.TextBox txt����ҽ�� 
         Height          =   300
         Left            =   1140
         TabIndex        =   29
         Top             =   1380
         Width           =   2115
      End
      Begin VB.TextBox txt����ժҪ 
         Height          =   600
         Left            =   1140
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   690
         Width           =   8835
      End
      Begin VB.CommandButton cmdҽ�ƻ��� 
         Caption         =   "��"
         Enabled         =   0   'False
         Height          =   300
         Left            =   9660
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   300
         Width           =   285
      End
      Begin VB.TextBox txtҽԺ���� 
         Height          =   300
         Left            =   5970
         TabIndex        =   24
         Top             =   300
         Width           =   3675
      End
      Begin VB.CommandButton cmd������Ϣ 
         Caption         =   "��"
         Height          =   300
         Left            =   4080
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   300
         Width           =   285
      End
      Begin VB.TextBox txt������Ϣ 
         Height          =   300
         Left            =   1140
         TabIndex        =   21
         Top             =   300
         Width           =   2955
      End
      Begin VB.CheckBox chk����ҽԺ 
         Caption         =   "����ҽԺ"
         Height          =   255
         Left            =   4890
         TabIndex        =   23
         Top             =   330
         Width           =   1035
      End
      Begin VB.Label lbl��ע 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ע"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   690
         TabIndex        =   40
         Top             =   3900
         Width           =   360
      End
      Begin VB.Label lbl��λ��� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��λ���"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   38
         Top             =   3210
         Width           =   720
      End
      Begin VB.Label lblҽԺ��� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽԺ���"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   36
         Top             =   2520
         Width           =   720
      End
      Begin VB.Label lbl������� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   34
         Top             =   1830
         Width           =   720
      End
      Begin VB.Label lbl��Чʱ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Чʱ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   7890
         TabIndex        =   32
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label lbl������ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(��)����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3900
         TabIndex        =   30
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label lbl����ҽ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ҽ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   28
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label lbl����ժҪ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ժҪ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   26
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lbl������Ϣ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   20
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.TextBox txt��λ���� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   7410
      TabIndex        =   18
      Top             =   1020
      Width           =   2985
   End
   Begin VB.TextBox txtҽ���� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   3630
      TabIndex        =   16
      Top             =   1020
      Width           =   2835
   End
   Begin VB.TextBox txtסԺ�� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   990
      TabIndex        =   14
      Top             =   1020
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -60
      TabIndex        =   44
      Top             =   540
      Width           =   10935
   End
   Begin VB.TextBox txt��Ա��� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   9090
      TabIndex        =   12
      Top             =   630
      Width           =   1305
   End
   Begin VB.TextBox txt�������� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   6990
      TabIndex        =   10
      Top             =   630
      Width           =   1035
   End
   Begin VB.TextBox txt�Ա� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   5340
      TabIndex        =   8
      Top             =   630
      Width           =   585
   End
   Begin VB.TextBox txt���� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   3630
      TabIndex        =   6
      Top             =   630
      Width           =   1035
   End
   Begin VB.TextBox txt���� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   990
      MaxLength       =   20
      TabIndex        =   4
      Top             =   630
      Width           =   1935
   End
   Begin VB.ComboBox cbo�������� 
      Height          =   300
      Left            =   990
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   180
      Width           =   1965
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "����(&R)"
      Height          =   350
      Left            =   3000
      TabIndex        =   2
      Top             =   150
      Width           =   1100
   End
   Begin VB.Label lbl��λ���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��λ����"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6600
      TabIndex        =   17
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label lblҽ���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ����"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3060
      TabIndex        =   15
      Top             =   1080
      Width           =   540
   End
   Begin VB.Label lblסԺ�� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "סԺ��"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   420
      TabIndex        =   13
      Top             =   1080
      Width           =   540
   End
   Begin VB.Label lbl��Ա��� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��Ա���"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   8250
      TabIndex        =   11
      Top             =   690
      Width           =   720
   End
   Begin VB.Label lbl�������� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6180
      TabIndex        =   9
      Top             =   690
      Width           =   720
   End
   Begin VB.Label lbl�Ա� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�Ա�"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4920
      TabIndex        =   7
      Top             =   690
      Width           =   360
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3210
      TabIndex        =   5
      Top             =   690
      Width           =   360
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "IC����"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   420
      TabIndex        =   3
      Top             =   690
      Width           =   540
   End
   Begin VB.Label lbl�������� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
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
      Height          =   180
      Left            =   150
      TabIndex        =   0
      Top             =   240
      Width           =   810
   End
End
Attribute VB_Name = "frmתԺ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private intMode As Integer          '�༭ģʽ
Private blnOK As Boolean
Private blnEnable As Boolean        '�Ƿ�����ҽ��
Private rsHospital As New ADODB.Recordset
Private Const madLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Private Const madDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Private Const madDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����
'תԺ�����Ϊ��ͨתԺ�����ת������
'��ͨתԺ�����ҵ��������11��12
'���ת�����������ҵ�����к�Ϊ��ʱ�������ף�ҵ������Ϊ11

Public Function ShowME(ByVal int�༭ģʽ As Integer, ByVal frmParent As Object) As Boolean
    On Error Resume Next
    blnOK = False
    intMode = int�༭ģʽ
    Me.Show 1, frmParent
    
    ShowME = blnOK
End Function

Private Sub InitFace()
    With cbo��������
        .Clear
        .AddItem "����"
        .ItemData(.NewIndex) = 11
        .AddItem "סԺ"
        .ItemData(.NewIndex) = 12
        .AddItem "ת�����"
        .ItemData(.NewIndex) = 11
        .ListIndex = 1
    End With
    Dtp��Чʱ��.Value = Format(DateAdd("d", 1, zlDatabase.Currentdate), "yyyy-MM-dd")
End Sub

Private Sub InitInsure()
    Dim lngRecord As Long
    Dim str��� As String, str���� As String
    Dim strFields As String, strValues As String
    
    '��ʼ��ҽ���ӿ�
    If Not ҽ����ʼ��_������ Then Exit Sub
    
    '��ȡҽԺ�����嵥
    strFields = "ID" & "," & adDouble & "," & "10" & "|" & _
                "���" & "," & adVarChar & "," & "20" & "|" & _
                "����" & "," & adLongVarChar & "," & "50"
    Call Record_Init(rsHospital, strFields)
    If Not ���ýӿ�_׼��_������(Function_������.��ȡҽԺ��Ϣ) Then Exit Sub
    If Not ���ýӿ�_ִ��_������ Then Exit Sub
    strFields = "ID|���|����"
    If ���ýӿ�_��¼��_������ Then
        Do While True
            lngRecord = lngRecord + 1
            Call ���ýӿ�_��ȡ����_������("hospital_id", str���)
            Call ���ýӿ�_��ȡ����_������("hospital_name", str����)
            
            strValues = lngRecord & "|" & str��� & "|" & str����
            Call Record_Add(rsHospital, strFields, strValues)
            'todo �˴��ǵ��Դ���
'            strValues = "1|001|001"
'            Call Record_Add(rsHospital, strFields, strValues)
'            strValues = "2|002|002"
'            Call Record_Add(rsHospital, strFields, strValues)
'            strValues = "3|003|003"
'            Call Record_Add(rsHospital, strFields, strValues)
'            strValues = "4|004|004"
'            Call Record_Add(rsHospital, strFields, strValues)
'            strValues = "5|005|005"
'            Call Record_Add(rsHospital, strFields, strValues)
'            strValues = "6|006|006"
'            Call Record_Add(rsHospital, strFields, strValues)
'            strValues = "7|007|007"
'            Call Record_Add(rsHospital, strFields, strValues)
'            strValues = "8|008|008"
'            Call Record_Add(rsHospital, strFields, strValues)
'            strValues = "9|009|009"
'            Call Record_Add(rsHospital, strFields, strValues)
'            strValues = "10|010|010"
            Call Record_Add(rsHospital, strFields, strValues)
            
            blnEnable = True
            If Not ���ýӿ�_�ƶ���¼��_������(MoveNext) Then Exit Do
        Loop
    End If
End Sub

Private Sub cbo��������_Click()
    gCominfo_������.ҵ������ = cbo��������.ItemData(cbo��������.ListIndex)
End Sub

Private Sub cbo��������_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.cmd����.SetFocus
End Sub

Private Sub chk����ҽԺ_Click()
    cmdҽ�ƻ���.Enabled = (chk����ҽԺ.Value = 1)
End Sub

Private Sub cmd����_Click()
    Dim str���� As String
    Dim strReturn As String
    Dim lngReturn As Long
    '--��IC��
    '����IC���е���Ϣ
    Me.txt���� = ""
    If Not ���ýӿ�_׼��_������(Function_������.����_����) Then Exit Sub
    If Not ���ýӿ�_ִ��_������ Then Exit Sub
    'ȡ���صļ�¼��
    'If Not ���ýӿ�_ָ����¼��_������("ICInfo") Then Exit Sub
    'Modified By ���� ��������ɳ ԭ�򣺽���������ɿ��Ÿ�Ϊ���֤��
    If Not ���ýӿ�_��ȡ����_������("card_no", str����) Then Exit Sub
    
    '��ȡ���˵Ļ�����Ϣ
    gstrField_������ = "hospital_id||iccardno"
    gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & str����
    If Not ���ýӿ�_׼��_������(Function_������.תԺ����_������Ϣ) Then Exit Sub
    If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Sub
    If Not ���ýӿ�_ִ��_������ Then Exit Sub
    lngReturn = glngReturn_������           '����ִ�н��
    
    '������������ʾ�ڽ�����
    'indi_id,insr_code,name,sex,pers_name,corp_name,patient_id,idcard
    If Not ���ýӿ�_ָ����¼��_������("PersonInfo") Then Exit Sub
    Call ���ýӿ�_��ȡ����_������("indi_id", strReturn)
    gCominfo_������.���˱�� = strReturn
    Call ���ýӿ�_��ȡ����_������("name", strReturn)
    Me.txt���� = strReturn
    Call ���ýӿ�_��ȡ����_������("sex", strReturn)
    Me.txt�Ա� = strReturn
    Call ���ýӿ�_��ȡ����_������("idcard", strReturn)
    If strReturn <> "" Then
        If Len(strReturn) > 15 Then
            strReturn = Mid(strReturn, 7, 8)
        Else
            strReturn = "19" & Mid(strReturn, 7, 6)
        End If
        Me.txt�������� = Mid(strReturn, 1, 4) & "-" & Mid(strReturn, 5, 2) & "-" & Mid(strReturn, 7)
    End If
    Call ���ýӿ�_��ȡ����_������("insr_code", strReturn)
    Me.txtҽ���� = strReturn
    Call ���ýӿ�_��ȡ����_������("corp_name", strReturn)
    Me.txt��λ���� = strReturn
    If lngReturn = 1 Then
        Call ���ýӿ�_��ȡ����_������("patient_id", strReturn)
        Me.txtסԺ�� = strReturn
    End If
    
    'У��תԺ���룬��ڲ�������
'    1   hospital_id ҽ�ƻ�������   20  ��
'    2   indi_id ���˱��           8   ��
    gstrField_������ = "hospital_id||indi_id"
    gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & gCominfo_������.���˱��
    If Not ���ýӿ�_׼��_������(Function_������.תԺ����_У����Ϣ) Then Exit Sub
    If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Sub
    If Not ���ýӿ�_ִ��_������ Then Exit Sub
    If Not ���ýӿ�_ָ����¼��_������("BizInfo") Then Exit Sub
    'ȡ��һ�����У�ֻ������һ�����������ֶ�����
'    1   serial_no   ҵ�����к� 12  ��Ч��סԺҵ���ҵ�����к�
'    2   patient_id  סԺ��     20
    Call ���ýӿ�_��ȡ����_������("patient_id", strReturn)
    Me.txtסԺ�� = strReturn
    If cbo��������.ListIndex <> 2 Then  'ת����ﲻ��Ҫ��ȡҵ�����к�
        Call ���ýӿ�_��ȡ����_������("serial_no", strReturn)
        gCominfo_������.ҵ�����к� = strReturn
    End If
    
    'ȫ��ִ�гɹ��Ž�������д�ڽ�����
    Me.txt���� = str����
End Sub

Private Sub cmd������Ϣ_Click()
    Dim rs���� As ADODB.Recordset
    
    gstrSQL = " Select A.ID,A.����,A.����,A.���� " & _
            " From ���ղ��� A where A.����=[1]"
    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "�����֤", TYPE_������)
    If rs����.RecordCount > 0 Then
        If frmListSel.ShowSelect(TYPE_������, rs����, "ID", "ҽ������ѡ��", "��ѡ��ҽ�����֣�") = True Then
            txt������Ϣ.Tag = rs����!����
            txt������Ϣ.Text = "(" & rs����!���� & ")" & rs����!����
            lbl������Ϣ.Tag = txt������Ϣ.Text '���ڻָ���ʾ
        End If
    End If
End Sub

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    '���������ȷ��
    If Trim(txt����) = "" Then
        MsgBox "��û��ȡ�òα����˵Ļ�����Ϣ�����ȵ��������ť��", vbInformation, gstrSysName
        cmd����.SetFocus
        Exit Sub
    End If
    If Trim(txt������Ϣ.Tag) = "" Then
        MsgBox "�������ѡ��òα��˵ļ�����Ϣ��", vbInformation, gstrSysName
        txt������Ϣ.SetFocus
        Exit Sub
    End If
    If chk����ҽԺ.Value = 1 Then
        If Trim(chk����ҽԺ.Tag) = "" Then
            MsgBox "�������ѡ�񶨵�ҽ�ƻ�����", vbInformation, gstrSysName
            txtҽԺ����.SetFocus
            Exit Sub
        End If
    Else
        If Trim(txtҽԺ����.Text) = "" Then
            MsgBox "������ҽ�ƻ������ƣ�", vbInformation, gstrSysName
            txtҽԺ����.SetFocus
            Exit Sub
        End If
    End If
    If Trim(txt����ҽ��.Text) = "" Then
        MsgBox "����������ҽ��������", vbInformation, gstrSysName
        txt����ҽ��.SetFocus
        Exit Sub
    End If
    
    '���ñ������빦��
'    1   busi_type      ҵ������    2   ��  "12"��סԺ
'    2   indi_id        ���˱��    8   ��
'    3   apply_hospital ����ҽ�ƻ�������    20  ��
'    4   hospital_id    ת��Ƕ���ҽ�ƻ�������  20
'    5   hospital_name  ת��Ƕ���ҽ�ƻ�������  60      hospital_id��hospital_name��to_hospital����ͬʱΪ��
'    6   to_hospital    ת�붨��ҽ�ƻ�������            20
'    7   apply_content  ��������    1   ��  "4"��תԺסԺ����    "5"��ת���ҽ����
'    8   serial_no      ҵ�����к�  12  ��
'    9   icd            ��������    20  ��  ҽ������Ŀ¼����
'    10  disease_desc   ����ժҪ    500 ��
'    11  doctor_name    ����ҽʦ    10  ��
'    12  apply_opinion  ��������    500 ��
'    13  hosp_opinion   ҽԺ������    500 ��
'    14  corp_opinion   ���˵�λ���    500 ��
'    15  apply_date     ������Чʱ��        ��  ��ʽ��YYYY-MM-DD
'    16  input_man      ¼����      10  ��
'    17  input_date     ¼��ʱ��        ��  ��ʽ��YYYY-MM-DD
'    18  note           ��ע        500 ��
    gstrField_������ = "busi_type||indi_id||apply_hospital||hospital_id||hospital_name||to_hospital||" & _
                       "apply_content||serial_no||icd||disease_desc||doctor_name||apply_opinion||" & _
                       "hosp_opinion||corp_opinion||apply_date||input_man||input_date||note"
    gstrValue_������ = gCominfo_������.ҵ������ & "||" & gCominfo_������.���˱�� & "||" & gCominfo_������.ҽԺ���� & "||" & _
                       "||" & IIf(chk����ҽԺ.Value = 1, "", txtҽԺ����.Text) & "||" & IIf(chk����ҽԺ.Value = 1, chk����ҽԺ.Tag, "") & "||" & _
                       IIf(cbo��������.ListIndex = 2, "5", "4") & "||" & gCominfo_������.ҵ�����к� & "||" & _
                       txt������Ϣ.Tag & "||" & txt����ժҪ.Text & "||" & txt����ҽ��.Text & "||" & txt�������.Text & "||" & _
                       txtҽԺ���.Text & "||" & txt��λ���.Text & "||" & Format(Me.Dtp��Чʱ��.Value, "yyyy-MM-dd") & "||" & _
                       gstrUserName & "||" & Format(zlDatabase.Currentdate, "yyyy-MM-dd") & "||" & txt��ע.Text
    If Not ���ýӿ�_׼��_������(Function_������.תԺ����_����תԺ����) Then Exit Sub
    If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Sub
    If Not ���ýӿ�_ִ��_������ Then Exit Sub
    
    blnOK = True
    Unload Me
    Exit Sub
End Sub

Private Sub cmdҽ�ƻ���_Click()
    If chk����ҽԺ.Value = 0 Then Exit Sub
    If frmListSel.ShowSelect(TYPE_������, rsHospital, "ID", "ѡ�񶨵�ҽ�ƻ���", "��ѡ��һ�Ҷ���ҽ�ƻ�����") = True Then
        chk����ҽԺ.Tag = rsHospital!���
        txtҽԺ����.Text = "(" & rsHospital!��� & ")" & rsHospital!����
        lblҽԺ���.Tag = txtҽԺ����.Text '���ڻָ���ʾ
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If InStr(1, "txtҽԺ����,txt������Ϣ", ActiveControl.Name) <> 0 Then Exit Sub
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call InitFace
    Call InitInsure
    If Not blnEnable Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '��Ӽ�¼
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Private Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '��ʼ��ӳ���¼��
    'strFields:�ֶ���,����,����|�ֶ���,����,����    �������Ϊ��,��ȡĬ�ϳ���
    '�ַ���:adLongVarChar;������:adDouble;������:adDBDate
    
    '���ӣ�
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|��ĿID," & adDouble & ",18|ժҪ, " & adLongVarChar & ",50|" & _
    '"ɾ��," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '��ȡ�ֶ�ȱʡ����
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub txt��ע_GotFocus()
    Call zlControl.TxtSelAll(txt��ע)
End Sub

Private Sub txt����ժҪ_GotFocus()
    Call zlControl.TxtSelAll(txt����ժҪ)
End Sub

Private Sub txt��λ����_GotFocus()
    Call zlControl.TxtSelAll(txt��λ���)
End Sub

Private Sub txt��λ���_GotFocus()
    Call zlControl.TxtSelAll(txt��λ���)
End Sub

Private Sub txt������Ϣ_GotFocus()
    Call zlControl.TxtSelAll(txt������Ϣ)
End Sub

Private Sub txt������Ϣ_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt������Ϣ.Text = "" Then Exit Sub
    
    On Error GoTo errHandle
    
    strText = txt������Ϣ.Text
    If InStr(1, strText, "(") <> 0 Then
        If InStr(1, strText, ")") <> 0 Then
            strText = Mid(strText, 2, InStr(1, strText, ")") - 2)
        Else
            strText = Mid(strText, 2)
        End If
    End If
    gstrSQL = "Select A.ID,A.����,A.����,A.����" & _
             "   FROM ���ղ��� A WHERE A.����=" & TYPE_������ & " And " & _
             "(A.���� like [2] || '%' or A.���� like [2] || '%' or A.���� like [2] || '%')"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_������, strText)
    
    If rsTemp.RecordCount > 0 Then
        '����ѡ����
        If rsTemp.RecordCount > 1 Then
            '�����ֶδ���3�ģ���ʹֻ��һ����¼�ѸöԻ�����ʾ�������Ա����û��õ��������Ϣ
            blnReturn = frmListSel.ShowSelect(TYPE_������, rsTemp, "ID", "ҽ������ѡ��", "��ѡ���ض���ҽ�����֣�")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '��¼����û�п�ѡ�������
        txt������Ϣ.SetFocus
        Call zlControl.TxtSelAll(txt������Ϣ)
        Exit Sub
    Else
        '�϶����м�¼����
        txt������Ϣ.Tag = rsTemp!����
        txt������Ϣ.Text = "(" & rsTemp!���� & ")" & rsTemp!����
        lbl������Ϣ.Tag = txt������Ϣ.Text '���ڻָ���ʾ
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub txt������_GotFocus()
    Call zlControl.TxtSelAll(txt������)
End Sub

Private Sub txt�������_GotFocus()
    Call zlControl.TxtSelAll(txt�������)
End Sub

Private Sub txtҽԺ����_GotFocus()
    Call zlControl.TxtSelAll(txtҽԺ����)
End Sub

Private Sub txtҽԺ����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim StrInput As String
    Dim blnReturn As Boolean
    If KeyCode <> vbKeyReturn Then Exit Sub
    If chk����ҽԺ.Value = 0 Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    
    StrInput = Trim(UCase(txtҽԺ����.Text))
    If InStr(1, StrInput, "(") <> 0 Then
        If InStr(1, StrInput, ")") <> 0 Then
            StrInput = Mid(StrInput, 2, InStr(1, StrInput, ")") - 2)
        Else
            StrInput = Mid(StrInput, 2)
        End If
    End If
    If IsNumeric(StrInput) Then
        rsHospital.Filter = "��� Like '" & StrInput & "*'"
    Else
        rsHospital.Filter = "���� Like '" & StrInput & "*'"
    End If
    If rsHospital.RecordCount = 0 Then
        MsgBox "û�ҵ��ñ�ŵĶ���ҽ�ƻ�����Ϣ�����������룡", vbInformation, gstrSysName
        rsHospital.Filter = 0
        txtҽԺ����.SetFocus
        GoTo ExitSub
    Else
        If rsHospital.RecordCount > 1 Then
            '�����ֶδ���3�ģ���ʹֻ��һ����¼�ѸöԻ�����ʾ�������Ա����û��õ��������Ϣ
            blnReturn = frmListSel.ShowSelect(TYPE_������, rsHospital, "ID", "ѡ�񶨵�ҽ�ƻ���", "��ѡ��һ�Ҷ���ҽ�ƻ�����")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '��¼����û�п�ѡ�������
        Call zlControl.TxtSelAll(txtҽԺ����)
        GoTo ExitSub
    Else
        '�϶����м�¼����
        chk����ҽԺ.Tag = rsHospital!���
        txtҽԺ����.Text = "(" & rsHospital!��� & ")" & rsHospital!����
        lblҽԺ���.Tag = txtҽԺ����.Text '���ڻָ���ʾ
        Call zlCommFun.PressKey(vbKeyTab)
    End If

ExitSub:
    rsHospital.Filter = 0
End Sub

Private Sub txtҽԺ���_GotFocus()
    Call zlControl.TxtSelAll(txtҽԺ���)
End Sub

Private Sub txt����ҽ��_GotFocus()
    Call zlControl.TxtSelAll(txt����ҽ��)
End Sub
