VERSION 5.00
Begin VB.Form frmTechnicSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   Icon            =   "frmTechnicSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Caption         =   "����ͼ��ʾ����"
      Height          =   690
      Left            =   120
      TabIndex        =   37
      Top             =   6450
      Width           =   5475
      Begin VB.TextBox TxtShowPhotoNumber 
         Height          =   315
         Left            =   1740
         TabIndex        =   20
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ʾ����ͼ��"
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   300
         Width           =   1440
      End
   End
   Begin VB.Frame fraAction 
      Caption         =   " ִ������ "
      Height          =   1875
      Left            =   120
      TabIndex        =   24
      Top             =   4560
      Width           =   5460
      Begin VB.CheckBox chkEmergencyPrint 
         Caption         =   "������˴�ӡ"
         Height          =   255
         Left            =   2940
         TabIndex        =   41
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CheckBox chkIgnorePosi 
         Caption         =   "���Խ����������"
         Height          =   180
         Left            =   2940
         TabIndex        =   40
         Top             =   1330
         Width           =   2235
      End
      Begin VB.CheckBox chkBatchInput 
         Caption         =   "������������"
         Height          =   225
         Left            =   2940
         TabIndex        =   39
         Top             =   390
         Width           =   2475
      End
      Begin VB.CheckBox chkSample 
         Caption         =   "����ǼǺ�ֱ�Ӽ��"
         Height          =   225
         Left            =   2940
         TabIndex        =   38
         Top             =   135
         Width           =   2475
      End
      Begin VB.CheckBox chkView 
         Caption         =   "��д����ʱ�򿪹�Ƭվ"
         Height          =   180
         Left            =   2940
         TabIndex        =   18
         Top             =   1115
         Width           =   2235
      End
      Begin VB.CheckBox chkFinish 
         Caption         =   "����δ�շѲ������ִ��"
         Height          =   195
         Left            =   2940
         TabIndex        =   17
         Top             =   885
         Width           =   2280
      End
      Begin VB.CheckBox chkActLog 
         Caption         =   "���������˴���ִ�м�¼"
         Height          =   195
         Left            =   2940
         TabIndex        =   16
         Top             =   645
         Width           =   2280
      End
      Begin VB.ListBox lstRoom 
         Enabled         =   0   'False
         Height          =   690
         ItemData        =   "frmTechnicSetup.frx":000C
         Left            =   255
         List            =   "frmTechnicSetup.frx":000E
         Style           =   1  'Checkbox
         TabIndex        =   15
         Top             =   555
         Width           =   2535
      End
      Begin VB.CheckBox chkRoom 
         Caption         =   "ָ�����˵�ִ�м䷶Χ"
         Height          =   195
         Left            =   255
         TabIndex        =   14
         Top             =   285
         Width           =   2100
      End
   End
   Begin VB.Frame fraExpence 
      Caption         =   " �Ʒ����� "
      Height          =   4470
      Left            =   120
      TabIndex        =   25
      Top             =   45
      Width           =   5460
      Begin VB.Frame fraLine 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   780
         TabIndex        =   35
         Top             =   4320
         Width           =   465
      End
      Begin VB.TextBox txtRefresh 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   780
         MaxLength       =   4
         TabIndex        =   12
         Text            =   "0"
         Top             =   4140
         Width           =   465
      End
      Begin VB.Frame Frame2 
         Caption         =   " ҩ������ "
         Height          =   2505
         Left            =   195
         TabIndex        =   28
         Top             =   1560
         Width           =   3525
         Begin VB.ComboBox cboס��ҩ 
            Height          =   300
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1710
            Width           =   2190
         End
         Begin VB.ComboBox cboס��ҩ 
            Height          =   300
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1365
            Width           =   2190
         End
         Begin VB.ComboBox cboס��ҩ 
            Height          =   300
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   2055
            Width           =   2190
         End
         Begin VB.ComboBox cbo�ų�ҩ 
            Height          =   300
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   615
            Width           =   2190
         End
         Begin VB.ComboBox cbo����ҩ 
            Height          =   300
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   270
            Width           =   2190
         End
         Begin VB.ComboBox cbo����ҩ 
            Height          =   300
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   960
            Width           =   2190
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ��ҩ��"
            Height          =   180
            Left            =   165
            TabIndex        =   34
            Top             =   1770
            Width           =   900
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ��ҩ��"
            Height          =   180
            Left            =   165
            TabIndex        =   33
            Top             =   1425
            Width           =   900
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ��ҩ��"
            Height          =   180
            Left            =   165
            TabIndex        =   32
            Top             =   2115
            Width           =   900
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����ҩ��"
            Height          =   180
            Left            =   165
            TabIndex        =   31
            Top             =   675
            Width           =   900
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������ҩ��"
            Height          =   180
            Left            =   165
            TabIndex        =   30
            Top             =   330
            Width           =   900
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������ҩ��"
            Height          =   180
            Left            =   165
            TabIndex        =   29
            Top             =   1020
            Width           =   900
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " ҩƷ��λ "
         Height          =   615
         Left            =   195
         TabIndex        =   27
         Top             =   870
         Width           =   3525
         Begin VB.OptionButton optҩƷ��λ 
            Caption         =   "�ۼ۵�λ"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   0
            Left            =   465
            TabIndex        =   4
            Top             =   285
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton optҩƷ��λ 
            Caption         =   "����/סԺ��λ"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   1
            Left            =   1560
            TabIndex        =   5
            Top             =   285
            Width           =   1470
         End
      End
      Begin VB.CheckBox chkҩ�� 
         Caption         =   "��ʾ����ҩ�����"
         Height          =   195
         Left            =   1995
         TabIndex        =   2
         Top             =   285
         Width           =   1770
      End
      Begin VB.CheckBox chkҩ�� 
         Caption         =   "��ʾ����ҩ����"
         Height          =   195
         Left            =   1995
         TabIndex        =   3
         Top             =   570
         Width           =   1770
      End
      Begin VB.CheckBox chkTime 
         Caption         =   "���������������"
         Height          =   195
         Left            =   195
         TabIndex        =   0
         Top             =   285
         Width           =   1740
      End
      Begin VB.CheckBox chkPay 
         Caption         =   "��ҩ�������븶��"
         Height          =   195
         Left            =   195
         TabIndex        =   1
         Top             =   570
         Value           =   1  'Checked
         Width           =   1740
      End
      Begin VB.ListBox lst�շ���� 
         Height          =   3420
         Left            =   3930
         Style           =   1  'Checkbox
         TabIndex        =   13
         ToolTipText     =   "�븴ѡ����ʹ�õ��շ����"
         Top             =   645
         Width           =   1350
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÿ��      ���Զ�ˢ�²����嵥"
         Height          =   180
         Left            =   390
         TabIndex        =   36
         Top             =   4155
         Width           =   2520
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������:"
         Height          =   180
         Left            =   3945
         TabIndex        =   26
         Top             =   420
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3060
      TabIndex        =   21
      Top             =   7215
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4155
      TabIndex        =   22
      Top             =   7215
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   420
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   7215
      Width           =   1100
   End
End
Attribute VB_Name = "frmTechnicSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Public mlng����ID As Long 'IN:��ǰִ�п���ID
Public mblnOK As Boolean

Private Sub chkRoom_Click()
    lstRoom.Enabled = chkRoom.Value = 1
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim strPar As String, i As Long
    
    'ִ�м䷶Χ
    strPar = ""
    If chkRoom.Value = 1 Then
        For i = 0 To lstRoom.ListCount - 1
            If lstRoom.Selected(i) Then
                strPar = strPar & "|" & lstRoom.List(i)
            End If
        Next
        If strPar = "" Then
            MsgBox "������ѡ��һ��ִ�м䡣", vbInformation, gstrSysName
            lstRoom.SetFocus: Exit Sub
        End If
    End If
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\����" & mlng����ID, "ִ�м䷶Χ", Mid(strPar, 2)
        
    '����
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��ҩ����", chkPay.Value
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "�������", chkTime.Value
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "��ʾ����ҩ�����", chkҩ��.Value
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "��ʾ����ҩ����", chkҩ��.Value
    
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "ҽ��ˢ�¼��", Val(txtRefresh.Text)
    
    'ҩƷ��λ
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "ҩƷ��λ", IIf(optҩƷ��λ(0).Value, 0, 1)
    
    'ȱʡҩ��
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "����ȱʡ��ҩ��", cbo����ҩ.ItemData(cbo����ҩ.ListIndex)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "����ȱʡ��ҩ��", cbo�ų�ҩ.ItemData(cbo�ų�ҩ.ListIndex)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "����ȱʡ��ҩ��", cbo����ҩ.ItemData(cbo����ҩ.ListIndex)
    
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "סԺȱʡ��ҩ��", cboס��ҩ.ItemData(cboס��ҩ.ListIndex)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "סԺȱʡ��ҩ��", cboס��ҩ.ItemData(cboס��ҩ.ListIndex)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "סԺȱʡ��ҩ��", cboס��ҩ.ItemData(cboס��ҩ.ListIndex)
    
    '�շ����
    strPar = ""
    For i = lst�շ����.ListCount - 1 To 0 Step -1
        If lst�շ����.Selected(i) Then strPar = strPar & "'" & Chr(lst�շ����.ItemData(i)) & "',"
    Next
    If strPar <> "" Then strPar = Left(strPar, Len(strPar) - 1)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "�շ����", strPar
    
    '�Ƿ��������ִ�м�¼
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "����ִ�м�¼", chkActLog.Value

    '�Ƿ��������δ�շѲ��˵���Ŀ
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "δ�շ����", chkFinish.Value
    
    '��ʾͼ����
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "��ʾͼ����", CLng(Val(TxtShowPhotoNumber.Text))
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "����ʱ��Ƭ", chkView.Value
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "�����Ǽ�����", chkBatchInput.Value
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "�Ǽ�ֱ�Ӽ��", chkSample.Value
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "���Խ��������", chkIgnorePosi.Value
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "�������ʱ��ӡ", chkEmergencyPrint.Value
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Call cmdHelp_Click
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim objCbo As ComboBox, lngҩ��ID As Long
    Dim strSQL As String, strPar As String, i As Long
    
    mblnOK = False
    
    chkPay.Value = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��ҩ����", 1))
    chkTime.Value = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "�������", 0))
    chkҩ��.Value = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "��ʾ����ҩ�����", 0))
    chkҩ��.Value = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "��ʾ����ҩ����", 0))
    
    txtRefresh.Text = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "ҽ��ˢ�¼��", 0))
        
    'ҩƷ��λ
    i = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "ҩƷ��λ", 0))
    optҩƷ��λ(IIf(i = 0, 0, 1)).Value = True
    
    'ȱʡҩ��
    cbo����ҩ.AddItem "�ֹ�ѡ��": cbo����ҩ.ListIndex = 0
    cbo�ų�ҩ.AddItem "�ֹ�ѡ��": cbo�ų�ҩ.ListIndex = 0
    cbo����ҩ.AddItem "�ֹ�ѡ��": cbo����ҩ.ListIndex = 0
    cboס��ҩ.AddItem "�ֹ�ѡ��": cboס��ҩ.ListIndex = 0
    cboס��ҩ.AddItem "�ֹ�ѡ��": cboס��ҩ.ListIndex = 0
    cboס��ҩ.AddItem "�ֹ�ѡ��": cboס��ҩ.ListIndex = 0
    strSQL = _
        "Select Distinct A.ID,A.����,A.����,B.��������,B.�������" & _
        " From ���ű� A,��������˵�� B " & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And B.����ID=A.ID And B.������� IN(1,2,3)" & _
        " And B.�������� in('��ҩ��','��ҩ��','��ҩ��')" & _
        " Order by A.����"
    Call OpenRecord(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        If rsTmp!�������� = "��ҩ��" Then
            Set objCbo = IIf(rsTmp!������� = 1, cbo����ҩ, IIf(rsTmp!������� = 2, cboס��ҩ, Nothing))
        End If
        If rsTmp!�������� = "��ҩ��" Then
            Set objCbo = IIf(rsTmp!������� = 1, cbo�ų�ҩ, IIf(rsTmp!������� = 2, cboס��ҩ, Nothing))
        End If
        If rsTmp!�������� = "��ҩ��" Then
            Set objCbo = IIf(rsTmp!������� = 1, cbo����ҩ, IIf(rsTmp!������� = 2, cboס��ҩ, Nothing))
        End If
        If objCbo Is Nothing Then
            If rsTmp!�������� = "��ҩ��" Then
                cbo����ҩ.AddItem rsTmp!����
                cbo����ҩ.ItemData(cbo����ҩ.NewIndex) = rsTmp!ID
                cboס��ҩ.AddItem rsTmp!����
                cboס��ҩ.ItemData(cboס��ҩ.NewIndex) = rsTmp!ID
            ElseIf rsTmp!�������� = "��ҩ��" Then
                cbo�ų�ҩ.AddItem rsTmp!����
                cbo�ų�ҩ.ItemData(cbo�ų�ҩ.NewIndex) = rsTmp!ID
                cboס��ҩ.AddItem rsTmp!����
                cboס��ҩ.ItemData(cboס��ҩ.NewIndex) = rsTmp!ID
            ElseIf rsTmp!�������� = "��ҩ��" Then
                cbo����ҩ.AddItem rsTmp!����
                cbo����ҩ.ItemData(cbo����ҩ.NewIndex) = rsTmp!ID
                cboס��ҩ.AddItem rsTmp!����
                cboס��ҩ.ItemData(cboס��ҩ.NewIndex) = rsTmp!ID
            End If
        Else
            objCbo.AddItem rsTmp!����
            objCbo.ItemData(objCbo.NewIndex) = rsTmp!ID
        End If
        rsTmp.MoveNext
    Next
    lngҩ��ID = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "����ȱʡ��ҩ��", 0))
    Call FindCboIndex(cbo����ҩ, lngҩ��ID, True)
    lngҩ��ID = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "����ȱʡ��ҩ��", 0))
    Call FindCboIndex(cbo�ų�ҩ, lngҩ��ID, True)
    lngҩ��ID = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "����ȱʡ��ҩ��", 0))
    Call FindCboIndex(cbo����ҩ, lngҩ��ID, True)
    lngҩ��ID = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "סԺȱʡ��ҩ��", 0))
    Call FindCboIndex(cboס��ҩ, lngҩ��ID, True)
    lngҩ��ID = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "סԺȱʡ��ҩ��", 0))
    Call FindCboIndex(cboס��ҩ, lngҩ��ID, True)
    lngҩ��ID = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "סԺȱʡ��ҩ��", 0))
    Call FindCboIndex(cboס��ҩ, lngҩ��ID, True)
    
    '�շ����
    strSQL = "Select ����,���� as ��� From �շ���Ŀ��� Where ����<>'1' Order by ���"
    Call OpenRecord(rsTmp, strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        lst�շ����.AddItem rsTmp!���
        lst�շ����.ItemData(lst�շ����.NewIndex) = Asc(rsTmp!����)
        rsTmp.MoveNext
    Loop
    strPar = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "�շ����", "")
    If strPar = "" Then
        For i = 0 To lst�շ����.ListCount - 1
            lst�շ����.Selected(i) = True
        Next
    Else
        For i = 0 To lst�շ����.ListCount - 1
            If InStr(strPar, Chr(lst�շ����.ItemData(i))) Then lst�շ����.Selected(i) = True
        Next
    End If
    If lst�շ����.ListCount > 0 Then lst�շ����.TopIndex = 0: lst�շ����.ListIndex = 0
    
    '�Ƿ��������ִ�м�¼
    chkActLog.Value = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "����ִ�м�¼", 0))
    
    '�Ƿ��������δ�շѲ��˵���Ŀ
    chkFinish.Value = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "δ�շ����", 0))
        
    'ִ�з���
    strPar = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\����" & mlng����ID, "ִ�м䷶Χ", "")
    chkRoom.Value = IIf(strPar = "", 0, 1)
    strSQL = "Select ִ�м� From ҽ��ִ�з��� Where ����ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    Do While Not rsTmp.EOF
        lstRoom.AddItem rsTmp!ִ�м�
        If InStr("|" & strPar & "|", "|" & rsTmp!ִ�м� & "|") > 0 Then
            lstRoom.Selected(lstRoom.NewIndex) = True
        End If
        rsTmp.MoveNext
    Loop
    If lstRoom.ListCount > 0 Then
        lstRoom.TopIndex = 0
        lstRoom.ListIndex = 0
    Else
        chkRoom.Value = 0
        chkRoom.Enabled = False
    End If
    
    '��ʾͼ����
    TxtShowPhotoNumber = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "��ʾͼ����", 20))
    
    chkView.Value = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "����ʱ��Ƭ", 0))
    chkBatchInput.Value = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "�����Ǽ�����", 0))
    chkSample.Value = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "�Ǽ�ֱ�Ӽ��", 0))
    chkIgnorePosi.Value = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "���Խ��������", 0))
    chkEmergencyPrint.Value = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "�������ʱ��ӡ", 0))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlng����ID = 0
End Sub

Private Sub lst�շ����_ItemCheck(Item As Integer)
    If lst�շ����.SelCount = 0 And Not lst�շ����.Selected(Item) Then
        lst�շ����.Selected(Item) = True
    End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    
End Sub

Private Sub txtRefresh_GotFocus()
    Call zlControl.TxtSelAll(txtRefresh)
End Sub

Private Sub txtRefresh_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub TxtShowPhotoNumber_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
