VERSION 5.00
Begin VB.Form frm���������걨�� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���������걨��"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   Icon            =   "frm���������걨��.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txt��ǰ���� 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   1200
      TabIndex        =   28
      Top             =   2400
      Width           =   1155
   End
   Begin VB.ComboBox cbo������� 
      Height          =   300
      Left            =   3300
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Width           =   1665
   End
   Begin VB.Frame Frame1 
      Caption         =   "����סԺ"
      Enabled         =   0   'False
      Height          =   1575
      Index           =   2
      Left            =   5430
      TabIndex        =   20
      Top             =   720
      Width           =   2445
      Begin VB.TextBox txt�����˴� 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   2
         Left            =   1020
         TabIndex        =   22
         Top             =   300
         Width           =   585
      End
      Begin VB.TextBox txt�����ܶ� 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   2
         Left            =   1020
         TabIndex        =   24
         Top             =   690
         Width           =   1155
      End
      Begin VB.TextBox txtͳ����� 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   2
         Left            =   1020
         TabIndex        =   26
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lbl�����˴� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����˴�"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl�����ܶ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����ܶ�"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   23
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lblͳ����� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ͳ�����"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   25
         Top             =   1140
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "����סԺ�ǰ���"
      Enabled         =   0   'False
      Height          =   1575
      Index           =   1
      Left            =   2820
      TabIndex        =   13
      Top             =   720
      Width           =   2445
      Begin VB.TextBox txtͳ����� 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   1
         Left            =   1020
         TabIndex        =   19
         Top             =   1080
         Width           =   1155
      End
      Begin VB.TextBox txt�����ܶ� 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   1
         Left            =   1020
         TabIndex        =   17
         Top             =   690
         Width           =   1155
      End
      Begin VB.TextBox txt�����˴� 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   1020
         TabIndex        =   15
         Top             =   300
         Width           =   585
      End
      Begin VB.Label lblͳ����� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ͳ�����"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lbl�����ܶ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����ܶ�"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lbl�����˴� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����˴�"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "����סԺ����"
      Enabled         =   0   'False
      Height          =   1575
      Index           =   0
      Left            =   210
      TabIndex        =   6
      Top             =   720
      Width           =   2445
      Begin VB.TextBox txt�����˴� 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   1020
         TabIndex        =   8
         Top             =   300
         Width           =   585
      End
      Begin VB.TextBox txt�����ܶ� 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   0
         Left            =   1020
         TabIndex        =   10
         Top             =   690
         Width           =   1155
      End
      Begin VB.TextBox txtͳ����� 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   0
         Left            =   1020
         TabIndex        =   12
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lbl�����˴� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����˴�"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl�����ܶ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����ܶ�"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lblͳ����� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ͳ�����"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   1140
         Width           =   720
      End
   End
   Begin VB.ComboBox cbo�ں� 
      Height          =   300
      Left            =   690
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   1665
   End
   Begin VB.CommandButton cmdȡ�� 
      Caption         =   "ȡ��(&D)"
      Height          =   350
      Left            =   5310
      TabIndex        =   4
      Top             =   210
      Width           =   1100
   End
   Begin VB.CommandButton cmd�걨 
      Caption         =   "�걨(&O)"
      Height          =   350
      Left            =   6480
      TabIndex        =   5
      Top             =   210
      Width           =   1100
   End
   Begin VB.Label lbl��ǰ���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ǰ����"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   240
      TabIndex        =   27
      Top             =   2460
      Width           =   900
   End
   Begin VB.Label lbl������� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2520
      TabIndex        =   2
      Top             =   300
      Width           =   720
   End
   Begin VB.Label lbl�ں� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�ں�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   270
      TabIndex        =   0
      Top             =   300
      Width           =   360
   End
End
Attribute VB_Name = "frm���������걨��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngID As Long              '0-����;�����ʾ����
Private mblnOK As Boolean           '�༭�ɹ�

Private Enum ����
    ����סԺ����
    ����סԺ�ǰ���
    ����סԺ
End Enum
'2�������걨�����У�
'   a������סԺ���ɣ��������Ϊ��������Ժ��ʽ���Ǽƻ������ģ�����=5��
'   b����������Ժ��ʽΪ�ƻ�������
'   c���ǰ��ɣ�����������������-�������-������

Public Function ShowME(ByVal lngID As Long) As Boolean
    mblnOK = False
    mlngID = lngID
    Me.Show 1
    ShowME = mblnOK
End Function

Private Sub cmdȡ��_Click()
    Dim str�ں� As String, str��ʼ���� As String, str�������� As String, str���ڽ������� As String
    Dim int�����˴� As Integer, dbl�����ܶ� As Double, dblͳ����� As Double
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If mlngID <> 0 Then
        '����ģʽ
        Unload Me
        Exit Sub
    End If
    
    '���
    Call ClearCons
    
    str�ں� = Me.cbo�ں�.Text
    str��ʼ���� = Mid(str�ں�, 1, 4) & "-" & Mid(str�ں�, 5, 2) & "-01 00:00:00"
    gstrSQL = " SELECT last_day(to_date('" & Mid(str��ʼ����, 1, 10) & "','yyyy-MM-dd')) from dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�¶����һ��")
    str�������� = Format(rsTemp.Fields(0).Value, "yyyy-MM-dd") & " 23:59:59"
    str���ڽ������� = Format(DateAdd("d", -1, str��ʼ����), "yyyy-MM-dd")
    
    '�����趨������ȡ��
    '1������סԺ���ɣ��������Ϊ��������Ժ��ʽ���Ǽƻ������ģ�����=5��
    gstrSQL = "SELECT  " & _
             "        COUNT(DISTINCT A.������ˮ��) AS �����˴�, " & _
             "        NVL(SUM(NVL(B.ҽ���ܷ���,0)),0) AS ҽ���ܷ���, " & _
             "        nvl(sum(c.ͳ�����),0) As ͳ����� " & _
             " FROM ���ս����¼ A,ZLGYYB.���㸽����Ϣ B," & _
             "      (Select ����id,Nvl(Sum(Decode(���㷽ʽ, 'ҽ������', Nvl(��Ԥ��, 0), 0)), 0) As ͳ����� " & _
             "      From ����Ԥ����¼ " & _
             "      Where �տ�ʱ�� BETWEEN [3] AND [4]" & _
             "      Group By ����id) C " & _
             " WHERE A.��¼ID=B.����ID AND A.��¼ID=C.����ID And A.ҽ�����<>'32'" & _
             " AND B.���㷽ʽ=5 And A.����֢=[1] And A.����=[2]" & _
             " AND A.����ʱ�� BETWEEN [3] AND [4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "������סԺ", CInt(cbo�������.ItemData(cbo�������.ListIndex)), TYPE_������, CDate(str��ʼ����), CDate(str��������))
    Me.txt�����˴�(����סԺ����).Text = Format(rsTemp!�����˴�, "#0;-#0; ;")
    Me.txt�����ܶ�(����סԺ����).Text = Format(rsTemp!ҽ���ܷ���, "#0.00;-#0.00; ;")
    Me.txtͳ�����(����סԺ����).Text = Format(rsTemp!ͳ�����, "#0.00;-#0.00; ;")
    
    '2������סԺ
    gstrSQL = "SELECT  " & _
             "        COUNT(DISTINCT A.������ˮ��) AS �����˴�, " & _
             "        NVL(SUM(NVL(B.ҽ���ܷ���,0)),0) AS ҽ���ܷ���, " & _
             "        nvl(sum(c.ͳ�����),0) As ͳ����� " & _
             " FROM ���ս����¼ A,ZLGYYB.���㸽����Ϣ B," & _
             "      (Select ����id,Nvl(Sum(Decode(���㷽ʽ, 'ҽ������', Nvl(��Ԥ��, 0), 0)), 0) As ͳ����� " & _
             "      From ����Ԥ����¼ " & _
             "      Where �տ�ʱ�� BETWEEN [3] AND [4]" & _
             "      Group By ����id) C " & _
             " WHERE A.��¼ID=B.����ID AND A.��¼ID=C.����ID And A.ҽ�����='32'" & _
             " AND A.����֢=[1] And A.����=[2]" & _
             " AND A.����ʱ�� BETWEEN [3] AND [4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��֢סԺ", CInt(cbo�������.ItemData(cbo�������.ListIndex)), TYPE_������, CDate(str��ʼ����), CDate(str��������))
    Me.txt�����˴�(����סԺ).Text = Format(rsTemp!�����˴�, "#0;-#0; ;")
    Me.txt�����ܶ�(����סԺ).Text = Format(rsTemp!ҽ���ܷ���, "#0.00;-#0.00; ;")
    Me.txtͳ�����(����סԺ).Text = Format(rsTemp!ͳ�����, "#0.00;-#0.00; ;")
    
    '3��ͳ�Ʊ�����������������������
    gstrSQL = "SELECT  " & _
             "        COUNT(DISTINCT A.������ˮ��) AS �����˴�, " & _
             "        NVL(SUM(NVL(B.ҽ���ܷ���,0)),0) AS ҽ���ܷ���, " & _
             "        nvl(sum(c.ͳ�����),0) As ͳ����� " & _
             " FROM ���ս����¼ A,ZLGYYB.���㸽����Ϣ B," & _
             "      (Select ����id,Nvl(Sum(Decode(���㷽ʽ, 'ҽ������', Nvl(��Ԥ��, 0), 0)), 0) As ͳ����� " & _
             "      From ����Ԥ����¼ " & _
             "      Where �տ�ʱ�� BETWEEN [3] AND [4]" & _
             "      Group By ����id) C " & _
             " WHERE A.��¼ID=B.����ID AND A.��¼ID=C.����ID " & _
             " AND A.����֢=[1] And A.����=[2]" & _
             " AND A.����ʱ�� BETWEEN [3] AND [4]"
             
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��֢סԺ", CInt(cbo�������.ItemData(cbo�������.ListIndex)), TYPE_������, CDate(str��ʼ����), CDate(str��������))
    int�����˴� = rsTemp!�����˴� - Val(txt�����˴�(����סԺ����).Text) - Val(txt�����˴�(����סԺ).Text)
    dbl�����ܶ� = rsTemp!ҽ���ܷ��� - Val(txt�����ܶ�(����סԺ����).Text) - Val(txt�����ܶ�(����סԺ).Text)
    dblͳ����� = rsTemp!ͳ����� - Val(txtͳ�����(����סԺ����).Text) - Val(txtͳ�����(����סԺ).Text)
    
    '4����ȡ��ǰ���Ѻϼ�
    gstrSQL = "SELECT Nvl(Sum(Decode(���㷽ʽ, '��ǰ����', Nvl(��Ԥ��, 0), 0)), 0) As ��ǰ���� " & _
             "      From ����Ԥ����¼ " & _
             "      Where �տ�ʱ�� BETWEEN [1] AND [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ǰ����", CDate(str��ʼ����), CDate(str��������))
    Me.txt��ǰ����.Text = Format(rsTemp!��ǰ����, "#0.00;-#0.00; ;")
    
    Me.txt�����˴�(����סԺ�ǰ���).Text = Format(int�����˴�, "#0;-#0; ;")
    Me.txt�����ܶ�(����סԺ�ǰ���).Text = Format(dbl�����ܶ�, "#0.00;-#0.00; ;")
    Me.txtͳ�����(����סԺ�ǰ���).Text = Format(dblͳ�����, "#0.00;-#0.00; ;")
    
    Me.Tag = 1
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call ClearCons
End Sub

Private Sub cmd�걨_Click()
    Dim str��ˮ�� As String
    On Error GoTo errHand
    
    If Val(Me.Tag) = 0 Then
        MsgBox "��ָ��������㡰ȡ������ť��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gcnGYYB.BeginTrans
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then
        gcnGYYB.RollbackTrans
        Exit Sub
    End If
    'סԺ�������ֻҪ������˱��룬��ʽ����ʱ��Ҫ����ſ����ݼ�����
    Call InsertChild(mdomInput.documentElement, "PERIOD", cbo�ں�.Text)
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName)
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"))
    Call InsertChild(mdomInput.documentElement, "INSURETYPE", cbo�������.ItemData(cbo�������.ListIndex))
    Call InsertChild(mdomInput.documentElement, "FMBGPSNS", Val(txt�����˴�(����סԺ����).Text))                 ' ��������˴�
    Call InsertChild(mdomInput.documentElement, "FMBGFEEALL", Val(txt�����ܶ�(����סԺ����).Text))
    Call InsertChild(mdomInput.documentElement, "FMBGFUND", Val(txtͳ�����(����סԺ����).Text))
    Call InsertChild(mdomInput.documentElement, "FMPSNS", Val(txt�����˴�(����סԺ�ǰ���).Text))
    Call InsertChild(mdomInput.documentElement, "FMFEEALL", Val(txt�����ܶ�(����סԺ�ǰ���).Text))
    Call InsertChild(mdomInput.documentElement, "FMFUND", Val(txtͳ�����(����סԺ�ǰ���).Text))
    Call InsertChild(mdomInput.documentElement, "JSPSNS", Val(txt�����˴�(����סԺ).Text))
    Call InsertChild(mdomInput.documentElement, "JSFEEALL", Val(txt�����ܶ�(����סԺ).Text))
    Call InsertChild(mdomInput.documentElement, "JSFUND", Val(txtͳ�����(����סԺ).Text))
    Call InsertChild(mdomInput.documentElement, "JCF", Val(txt��ǰ����.Text))
    '���ýӿ�
    If CommRecServer("APPRECB") = False Then
        gcnGYYB.RollbackTrans
        Exit Sub
    End If
    str��ˮ�� = GetElemnetValue("APPNO")
    
    '��������
    mlngID = GetNextID("���㵥", gcnGYYB)
    gstrSQL = "ZL_���㵥_INSERT(" & mlngID & ",1,'" & Me.cbo�ں�.Text & "'," & cbo�������.ItemData(cbo�������.ListIndex) & "," & _
        "'" & cbo�������.Text & "','" & gstrUserName & "',sysdate,'" & str��ˮ�� & "',NULL)"
    gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    
    gstrSQL = "ZL_����������ϸ_INSERT(" & mlngID & "," & Val(txt�����˴�(����סԺ����).Text) & "," & Val(txt�����ܶ�(����סԺ����).Text) & "," & Val(txtͳ�����(����סԺ����).Text) & "," & _
            Val(txt�����˴�(����סԺ�ǰ���).Text) & "," & Val(txt�����ܶ�(����סԺ�ǰ���).Text) & "," & Val(txtͳ�����(����סԺ�ǰ���).Text) & "," & _
            Val(txt�����˴�(����סԺ).Text) & "," & Val(txt�����ܶ�(����סԺ).Text) & "," & Val(txtͳ�����(����סԺ).Text) & "," & Val(txt��ǰ����.Text) & ")"
    gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    gcnGYYB.CommitTrans
    
    mblnOK = True
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    gcnGYYB.RollbackTrans
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
        Exit Sub
    End If
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    Dim curDate As Date
    Dim str���� As String, str���� As String
    Dim rsData As New ADODB.Recordset
    
    If mlngID = 0 Then
        'ȱʡֻװ�����¡����¹��걨
        curDate = zlDatabase.Currentdate()
        str���� = Format(DateAdd("m", -1, curDate), "yyyyMM")
        str���� = Format(curDate, "yyyyMM")
        With cbo�ں�
            .Clear
            .AddItem str����
            .AddItem str����
            .ListIndex = 0
        End With
        With cbo�������
            .Clear
            .AddItem "��ҵ��������"
            .ItemData(.NewIndex) = 4
            .AddItem "������ҵ��λ��������"
            .ItemData(.NewIndex) = 5
            .ListIndex = 0
        End With
        Exit Sub
    End If
    
    '��ȡ�걨������
    gstrSQL = "SELECT  " & _
             "        A.ID, A.�ں�, A.�������, A.����Ա, A.���� ,B.��������˴�, B.������ɷ����ܶ�, B.�������ͳ��֧��, B.����ǰ����˴�, " & _
             "        B.����ǰ��ɷ����ܶ�, B.����ǰ���ͳ��֧��, B.�����˴�, B.���������ܶ�, B.����ͳ��֧��, A.������ˮ��, A.������� " & _
             " FROM ���㵥 A, ����������ϸ B " & _
             " WHERE A.ID=B.���㵥ID AND A.ID= [1]"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�걨������", mlngID)
    
    '����
    With rsData
        Me.cbo�ں�.AddItem !�ں�
        Me.cbo�ں�.ListIndex = 0
        
        Me.txt�����˴�(����סԺ����).Text = Format(Nvl(!��������˴�, 0), "#0;-#0; ;")
        Me.txt�����ܶ�(����סԺ����).Text = Format(Nvl(!������ɷ����ܶ�, 0), "#0.00;-#0.00; ;")
        Me.txtͳ�����(����סԺ����).Text = Format(Nvl(!�������ͳ��֧��, 0), "#0.00;-#0.00; ;")
        
        Me.txt�����˴�(����סԺ�ǰ���).Text = Format(Nvl(!����ǰ����˴�, 0), "#0;-#0; ;")
        Me.txt�����ܶ�(����סԺ�ǰ���).Text = Format(Nvl(!����ǰ��ɷ����ܶ�, 0), "#0.00;-#0.00; ;")
        Me.txtͳ�����(����סԺ�ǰ���).Text = Format(Nvl(!����ǰ���ͳ��֧��, 0), "#0.00;-#0.00; ;")
        
        Me.txt�����˴�(����סԺ).Text = Format(Nvl(!�����˴�, 0), "#0;-#0; ;")
        Me.txt�����ܶ�(����סԺ).Text = Format(Nvl(!���������ܶ�, 0), "#0.00;-#0.00; ;")
        Me.txtͳ�����(����סԺ).Text = Format(Nvl(!����ͳ��֧��, 0), "#0.00;-#0.00; ;")
    End With
    
    '���ÿؼ�״̬
    Me.cbo�ں�.Enabled = False
    
    cmd�걨.Visible = False
    cmdȡ��.Caption = "�˳�(&X)"
End Sub

Private Sub ClearCons()
    Me.Tag = ""
    Me.txt�����˴�(����סԺ����).Text = ""
    Me.txt�����ܶ�(����סԺ����).Text = ""
    Me.txtͳ�����(����סԺ����).Text = ""
    
    Me.txt�����˴�(����סԺ�ǰ���).Text = ""
    Me.txt�����ܶ�(����סԺ�ǰ���).Text = ""
    Me.txtͳ�����(����סԺ�ǰ���).Text = ""
    
    Me.txt�����˴�(����סԺ).Text = ""
    Me.txt�����ܶ�(����סԺ).Text = ""
    Me.txtͳ�����(����סԺ).Text = ""
End Sub
