VERSION 5.00
Begin VB.Form frmExpenseSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ������ѡ��"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   Icon            =   "frmExpenseSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   405
      Left            =   6180
      TabIndex        =   16
      Top             =   4155
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   405
      Left            =   5085
      TabIndex        =   15
      Top             =   4155
      Width           =   1100
   End
   Begin VB.Frame fraExpence 
      Height          =   3930
      Left            =   90
      TabIndex        =   17
      Top             =   45
      Width           =   7245
      Begin VB.OptionButton optȱʡ���� 
         Caption         =   "���˿���"
         Height          =   195
         Index           =   1
         Left            =   5205
         TabIndex        =   32
         Top             =   3600
         Width           =   1065
      End
      Begin VB.OptionButton optȱʡ���� 
         Caption         =   "ҽ������"
         Height          =   195
         Index           =   0
         Left            =   4065
         TabIndex        =   31
         Top             =   3585
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.ComboBox cboSendMateria 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   3525
         Width           =   1860
      End
      Begin VB.ListBox lst�շ���� 
         ForeColor       =   &H80000012&
         Height          =   3000
         Left            =   5700
         Style           =   1  'Checkbox
         TabIndex        =   14
         ToolTipText     =   "�븴ѡ����ʹ�õ��շ����"
         Top             =   450
         Width           =   1440
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
      Begin VB.CheckBox chkTime 
         Caption         =   "���������������"
         Height          =   195
         Left            =   195
         TabIndex        =   0
         Top             =   285
         Width           =   1740
      End
      Begin VB.CheckBox chkҩ�� 
         Caption         =   "��ʾ����ҩ����"
         Height          =   195
         Left            =   195
         TabIndex        =   3
         Top             =   1170
         Width           =   1770
      End
      Begin VB.CheckBox chkҩ�� 
         Caption         =   "��ʾ����ҩ�����"
         Height          =   195
         Left            =   195
         TabIndex        =   2
         Top             =   885
         Width           =   1770
      End
      Begin VB.Frame Frame3 
         Caption         =   " ҩƷ��λ "
         Height          =   1215
         Left            =   3120
         TabIndex        =   25
         Top             =   240
         Width           =   2445
         Begin VB.OptionButton optҩƷ��λ 
            Caption         =   "����/סԺ��λ"
            Height          =   180
            Index           =   1
            Left            =   360
            TabIndex        =   5
            Top             =   720
            Width           =   1470
         End
         Begin VB.OptionButton optҩƷ��λ 
            Caption         =   "�ۼ۵�λ"
            Height          =   180
            Index           =   0
            Left            =   345
            TabIndex        =   4
            Top             =   330
            Value           =   -1  'True
            Width           =   1020
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " ҩ������ "
         Height          =   1875
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   5445
         Begin VB.ComboBox cbo�ŷ��ϲ��� 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1290
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1320
            Width           =   1350
         End
         Begin VB.ComboBox cboס���ϲ��� 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   3915
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1320
            Width           =   1350
         End
         Begin VB.ComboBox cbo����ҩ 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   960
            Width           =   1350
         End
         Begin VB.ComboBox cbo����ҩ 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   270
            Width           =   1350
         End
         Begin VB.ComboBox cbo�ų�ҩ 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   615
            Width           =   1350
         End
         Begin VB.ComboBox cboס��ҩ 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   3915
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   975
            Width           =   1350
         End
         Begin VB.ComboBox cboס��ҩ 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   3915
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   285
            Width           =   1350
         End
         Begin VB.ComboBox cboס��ҩ 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   3915
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   630
            Width           =   1350
         End
         Begin VB.Label lbl�ŷ��ϲ��� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���﷢�ϲ���"
            Height          =   180
            Left            =   120
            TabIndex        =   28
            Top             =   1380
            Width           =   1080
         End
         Begin VB.Label lblס���ϲ��� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ���ϲ���"
            Height          =   180
            Left            =   2745
            TabIndex        =   27
            Top             =   1380
            Width           =   1080
         End
         Begin VB.Label lbl����ҩ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������ҩ��"
            Height          =   180
            Left            =   285
            TabIndex        =   24
            Top             =   1020
            Width           =   900
         End
         Begin VB.Label lbl����ҩ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������ҩ��"
            Height          =   180
            Left            =   285
            TabIndex        =   23
            Top             =   330
            Width           =   900
         End
         Begin VB.Label lbl�ų�ҩ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����ҩ��"
            Height          =   180
            Left            =   285
            TabIndex        =   22
            Top             =   675
            Width           =   900
         End
         Begin VB.Label lblס��ҩ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ��ҩ��"
            Height          =   180
            Left            =   2925
            TabIndex        =   21
            Top             =   1035
            Width           =   900
         End
         Begin VB.Label lblס��ҩ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ��ҩ��"
            Height          =   180
            Left            =   2925
            TabIndex        =   20
            Top             =   345
            Width           =   900
         End
         Begin VB.Label lblס��ҩ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ��ҩ��"
            Height          =   180
            Left            =   2925
            TabIndex        =   19
            Top             =   690
            Width           =   900
         End
      End
      Begin VB.Label lblFee 
         AutoSize        =   -1  'True
         Caption         =   "����ȱʡ����"
         Height          =   180
         Left            =   2910
         TabIndex        =   30
         Top             =   3585
         Width           =   1080
      End
      Begin VB.Label lbl��ҩ 
         Caption         =   "����֮��"
         Height          =   255
         Left            =   105
         TabIndex        =   33
         Top             =   3555
         Width           =   735
      End
      Begin VB.Label lbl�շ���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������:"
         Height          =   180
         Left            =   5745
         TabIndex        =   26
         Top             =   225
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmExpenseSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Public mMainPrivs As String
Public mblnOK As Boolean
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub
Private Sub cmdOK_Click()
    Dim strPar As String, i As Long, blnSetup As Boolean
    
    blnSetup = InStr(GetInsidePrivs(pҽ�����ѹ���), "����ѡ������") > 0
    

    If cbo����ҩ.ListIndex = -1 Then
        MsgBox "��ָ��ȱʡ��������ҩ����", vbInformation, gstrSysName
        cbo����ҩ.SetFocus: Exit Sub
    End If
    If cbo�ų�ҩ.ListIndex = -1 Then
        MsgBox "��ָ��ȱʡ�������ҩ����", vbInformation, gstrSysName
        cbo�ų�ҩ.SetFocus: Exit Sub
    End If
    If cbo����ҩ.ListIndex = -1 Then
        MsgBox "��ָ��ȱʡ��������ҩ����", vbInformation, gstrSysName
        cbo����ҩ.SetFocus: Exit Sub
    End If
    If cboס��ҩ.ListIndex = -1 Then
        MsgBox "��ָ��ȱʡ��סԺ��ҩ����", vbInformation, gstrSysName
        cboס��ҩ.SetFocus: Exit Sub
    End If
    If cboס��ҩ.ListIndex = -1 Then
        MsgBox "��ָ��ȱʡ��סԺ��ҩ����", vbInformation, gstrSysName
        cboס��ҩ.SetFocus: Exit Sub
    End If
    If cboס��ҩ.ListIndex = -1 Then
        MsgBox "��ָ��ȱʡ��סԺ��ҩ����", vbInformation, gstrSysName
        cboס��ҩ.SetFocus: Exit Sub
    End If
    
    '����
    Call zlDatabase.SetPara("��ҩ���븶��", chkPay.Value, glngSys, pҽ�����ѹ���, blnSetup)
    Call zlDatabase.SetPara("�����������", chkTime.Value, glngSys, pҽ�����ѹ���, blnSetup)
    '����:51762
    Call zlDatabase.SetPara("��ʾ����ҩ����", chkҩ��.Value, glngSys, pҽ�����ѹ���, blnSetup)
    Call zlDatabase.SetPara("��ʾ����ҩ�����", chkҩ��.Value, glngSys, pҽ�����ѹ���, blnSetup)
    Call zlDatabase.SetPara("����ȱʡ����", IIF(optȱʡ����(0).Value, 0, 1), glngSys, pҽ�����ѹ���, blnSetup)
        
    'ҩƷ��λ
    Call zlDatabase.SetPara("ҩƷ��λ", IIF(optҩƷ��λ(0).Value, 0, 1), glngSys, pҽ�����ѹ���, blnSetup)
    '��ҩ��ʽ:25490
    Call zlDatabase.SetPara("���ʺ�ҩ", cboSendMateria.ListIndex, glngSys, pҽ�����ѹ���, blnSetup)
    'ȱʡҩ��
    Call zlDatabase.SetPara("����ȱʡ��ҩ��", cbo����ҩ.ItemData(cbo����ҩ.ListIndex), glngSys, pҽ�����ѹ���, blnSetup)
    Call zlDatabase.SetPara("����ȱʡ��ҩ��", cbo�ų�ҩ.ItemData(cbo�ų�ҩ.ListIndex), glngSys, pҽ�����ѹ���, blnSetup)
    Call zlDatabase.SetPara("����ȱʡ��ҩ��", cbo����ҩ.ItemData(cbo����ҩ.ListIndex), glngSys, pҽ�����ѹ���, blnSetup)
    Call zlDatabase.SetPara("����ȱʡ���ϲ���", cbo�ŷ��ϲ���.ItemData(cbo�ŷ��ϲ���.ListIndex), glngSys, pҽ�����ѹ���, blnSetup)
    Call zlDatabase.SetPara("סԺȱʡ��ҩ��", cboס��ҩ.ItemData(cboס��ҩ.ListIndex), glngSys, pҽ�����ѹ���, blnSetup)
    Call zlDatabase.SetPara("סԺȱʡ��ҩ��", cboס��ҩ.ItemData(cboס��ҩ.ListIndex), glngSys, pҽ�����ѹ���, blnSetup)
    Call zlDatabase.SetPara("סԺȱʡ��ҩ��", cboס��ҩ.ItemData(cboס��ҩ.ListIndex), glngSys, pҽ�����ѹ���, blnSetup)
    Call zlDatabase.SetPara("סԺȱʡ���ϲ���", cboס���ϲ���.ItemData(cboס���ϲ���.ListIndex), glngSys, pҽ�����ѹ���, blnSetup)
    
    '�շ����
    strPar = ""
    For i = lst�շ����.ListCount - 1 To 0 Step -1
        If lst�շ����.Selected(i) Then strPar = strPar & "'" & Chr(lst�շ����.ItemData(i)) & "',"
    Next
    If strPar <> "" Then strPar = Left(strPar, Len(strPar) - 1)
    Call zlDatabase.SetPara("�շ����", Replace(strPar, "'", "''"), glngSys, pҽ�����ѹ���, blnSetup)
    
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
    Dim objCbo As ComboBox, strPar As String
    Dim strSQL As String, i As Long
    Dim blnSetup As Boolean
    
    On Error GoTo errH
    mblnOK = False
    blnSetup = InStr(GetInsidePrivs(pҽ�����ѹ���), "����ѡ������") > 0
    
    chkPay.Value = Val(zlDatabase.GetPara("��ҩ���븶��", glngSys, pҽ�����ѹ���, , Array(chkPay), blnSetup))
    chkTime.Value = Val(zlDatabase.GetPara("�����������", glngSys, pҽ�����ѹ���, , Array(chkTime), blnSetup))
    chkҩ��.Value = Val(zlDatabase.GetPara("��ʾ����ҩ�����", glngSys, pҽ�����ѹ���, , Array(chkҩ��), blnSetup))
    chkҩ��.Value = Val(zlDatabase.GetPara("��ʾ����ҩ����", glngSys, pҽ�����ѹ���, , Array(chkҩ��), blnSetup))
    '����:36060
    If Val(zlDatabase.GetPara("����ȱʡ����", glngSys, pҽ�����ѹ���, , Array(optȱʡ����(0), optȱʡ����(1), lblFee), blnSetup)) = 0 Then
        optȱʡ����(0).Value = True
    Else
        optȱʡ����(1).Value = True
    End If
      

    'ҩƷ��λ
    i = Val(zlDatabase.GetPara("ҩƷ��λ", glngSys, pҽ�����ѹ���, , Array(optҩƷ��λ(0), optҩƷ��λ(1)), blnSetup))
    optҩƷ��λ(IIF(i = 0, 0, 1)).Value = True
    
    '25490
    cboSendMateria.Clear
    cboSendMateria.AddItem "����ҩ"
    cboSendMateria.AddItem "�Զ���ҩ"
    cboSendMateria.AddItem "��ʾ��ҩ"
    i = Val(zlDatabase.GetPara("���ʺ�ҩ", glngSys, pҽ�����ѹ���, 0, Array(lbl��ҩ, cboSendMateria), blnSetup))
    If i > cboSendMateria.ListCount Then i = 0
    cboSendMateria.ListIndex = i
    
    
    'ȱʡҩ��
    cbo����ҩ.AddItem "�ֹ�ѡ��"
    cbo�ų�ҩ.AddItem "�ֹ�ѡ��"
    cbo����ҩ.AddItem "�ֹ�ѡ��"
    cboס��ҩ.AddItem "�ֹ�ѡ��"
    cboס��ҩ.AddItem "�ֹ�ѡ��"
    cboס��ҩ.AddItem "�ֹ�ѡ��"
    cbo�ŷ��ϲ���.AddItem "�ֹ�ѡ��"
    cboס���ϲ���.AddItem "�ֹ�ѡ��"
    strSQL = _
        "Select Distinct A.ID,A.����,A.����,B.��������,B.�������" & _
        " From ���ű� A,��������˵�� B " & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And B.����ID=A.ID And B.������� IN(1,2,3)" & _
        " And B.�������� in('��ҩ��','��ҩ��','��ҩ��','���ϲ���')" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " Order by A.����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        If rsTmp!�������� = "��ҩ��" Then
            Set objCbo = IIF(rsTmp!������� = 1, cbo����ҩ, IIF(rsTmp!������� = 2, cboס��ҩ, Nothing))
        End If
        If rsTmp!�������� = "��ҩ��" Then
            Set objCbo = IIF(rsTmp!������� = 1, cbo�ų�ҩ, IIF(rsTmp!������� = 2, cboס��ҩ, Nothing))
        End If
        If rsTmp!�������� = "��ҩ��" Then
            Set objCbo = IIF(rsTmp!������� = 1, cbo����ҩ, IIF(rsTmp!������� = 2, cboס��ҩ, Nothing))
        End If
        If rsTmp!�������� = "���ϲ���" Then
            Set objCbo = IIF(rsTmp!������� = 1, cbo�ŷ��ϲ���, IIF(rsTmp!������� = 2, cboס���ϲ���, Nothing))
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
            ElseIf rsTmp!�������� = "���ϲ���" Then
                cbo�ŷ��ϲ���.AddItem rsTmp!����
                cbo�ŷ��ϲ���.ItemData(cbo�ŷ��ϲ���.NewIndex) = rsTmp!ID
                cboס���ϲ���.AddItem rsTmp!����
                cboס���ϲ���.ItemData(cboס���ϲ���.NewIndex) = rsTmp!ID
            End If
        Else
            objCbo.AddItem rsTmp!����
            objCbo.ItemData(objCbo.NewIndex) = rsTmp!ID
        End If
        rsTmp.MoveNext
    Next
    strPar = zlDatabase.GetPara("����ȱʡ��ҩ��", glngSys, pҽ�����ѹ���, , Array(lbl����ҩ, cbo����ҩ), blnSetup)
    For i = 0 To cbo����ҩ.ListCount - 1
        If cbo����ҩ.ItemData(i) = Val(strPar) Then cbo����ҩ.ListIndex = i: Exit For
    Next
    strPar = zlDatabase.GetPara("����ȱʡ��ҩ��", glngSys, pҽ�����ѹ���, , Array(lbl�ų�ҩ, cbo�ų�ҩ), blnSetup)
    For i = 0 To cbo�ų�ҩ.ListCount - 1
        If cbo�ų�ҩ.ItemData(i) = Val(strPar) Then cbo�ų�ҩ.ListIndex = i: Exit For
    Next
    strPar = zlDatabase.GetPara("����ȱʡ��ҩ��", glngSys, pҽ�����ѹ���, , Array(lbl����ҩ, cbo����ҩ), blnSetup)
    For i = 0 To cbo����ҩ.ListCount - 1
        If cbo����ҩ.ItemData(i) = Val(strPar) Then cbo����ҩ.ListIndex = i: Exit For
    Next
    strPar = zlDatabase.GetPara("����ȱʡ���ϲ���", glngSys, pҽ�����ѹ���, , Array(lbl�ŷ��ϲ���, cbo�ŷ��ϲ���), blnSetup)
    For i = 0 To cbo�ŷ��ϲ���.ListCount - 1
        If cbo�ŷ��ϲ���.ItemData(i) = Val(strPar) Then cbo�ŷ��ϲ���.ListIndex = i: Exit For
    Next
    
    
    strPar = zlDatabase.GetPara("סԺȱʡ��ҩ��", glngSys, pҽ�����ѹ���, , Array(lblס��ҩ, cboס��ҩ), blnSetup)
    For i = 0 To cboס��ҩ.ListCount - 1
        If cboס��ҩ.ItemData(i) = Val(strPar) Then cboס��ҩ.ListIndex = i: Exit For
    Next
    strPar = zlDatabase.GetPara("סԺȱʡ��ҩ��", glngSys, pҽ�����ѹ���, , Array(lblס��ҩ, cboס��ҩ), blnSetup)
    For i = 0 To cboס��ҩ.ListCount - 1
        If cboס��ҩ.ItemData(i) = Val(strPar) Then cboס��ҩ.ListIndex = i: Exit For
    Next
    strPar = zlDatabase.GetPara("סԺȱʡ��ҩ��", glngSys, pҽ�����ѹ���, , Array(lblס��ҩ, cboס��ҩ), blnSetup)
    For i = 0 To cboס��ҩ.ListCount - 1
        If cboס��ҩ.ItemData(i) = Val(strPar) Then cboס��ҩ.ListIndex = i: Exit For
    Next
    strPar = zlDatabase.GetPara("סԺȱʡ���ϲ���", glngSys, pҽ�����ѹ���, , Array(lblס���ϲ���, cboס���ϲ���), blnSetup)
    For i = 0 To cboס���ϲ���.ListCount - 1
        If cboס���ϲ���.ItemData(i) = Val(strPar) Then cboס���ϲ���.ListIndex = i: Exit For
    Next
    
    
    '�շ����
    strSQL = "Select ����,���� as ��� From �շ���Ŀ��� Where ����<>'1' Order by ���"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        lst�շ����.AddItem rsTmp!���
        lst�շ����.ItemData(lst�շ����.NewIndex) = Asc(rsTmp!����)
        rsTmp.MoveNext
    Loop
    strPar = zlDatabase.GetPara("�շ����", glngSys, pҽ�����ѹ���, "", Array(lbl�շ����, lst�շ����), blnSetup)
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
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mMainPrivs = ""
End Sub

Private Sub lst�շ����_ItemCheck(Item As Integer)
    If lst�շ����.SelCount = 0 And Not lst�շ����.Selected(Item) Then
        lst�շ����.Selected(Item) = True
    End If
End Sub
