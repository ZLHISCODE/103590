VERSION 5.00
Begin VB.Form frmExpenseSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ������ѡ��"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   Icon            =   "frmExpenseSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   405
      Left            =   5970
      TabIndex        =   10
      Top             =   3630
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   405
      Left            =   4875
      TabIndex        =   9
      Top             =   3630
      Width           =   1100
   End
   Begin VB.Frame fraExpence 
      Height          =   3405
      Left            =   90
      TabIndex        =   11
      Top             =   45
      Width           =   7245
      Begin VB.OptionButton optȱʡ���� 
         Caption         =   "���˿���"
         Height          =   195
         Index           =   1
         Left            =   2430
         TabIndex        =   24
         Top             =   2430
         Width           =   1065
      End
      Begin VB.OptionButton optȱʡ���� 
         Caption         =   "ҽ������"
         Height          =   195
         Index           =   0
         Left            =   1290
         TabIndex        =   23
         Top             =   2430
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.ListBox lst�շ���� 
         ForeColor       =   &H80000012&
         Height          =   2790
         Left            =   5700
         Style           =   1  'Checkbox
         TabIndex        =   8
         ToolTipText     =   "�븴ѡ����ʹ�õ��շ����"
         Top             =   450
         Width           =   1440
      End
      Begin VB.Frame Frame2 
         Caption         =   " ҩ������ "
         Height          =   1875
         Left            =   120
         TabIndex        =   12
         Top             =   345
         Width           =   5445
         Begin VB.ComboBox cbo�ŷ��ϲ��� 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1290
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1320
            Width           =   1350
         End
         Begin VB.ComboBox cboס���ϲ��� 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   3915
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1320
            Width           =   1350
         End
         Begin VB.ComboBox cbo����ҩ 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   960
            Width           =   1350
         End
         Begin VB.ComboBox cbo����ҩ 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   270
            Width           =   1350
         End
         Begin VB.ComboBox cbo�ų�ҩ 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   615
            Width           =   1350
         End
         Begin VB.ComboBox cboס��ҩ 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   3915
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   975
            Width           =   1350
         End
         Begin VB.ComboBox cboס��ҩ 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   3915
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   285
            Width           =   1350
         End
         Begin VB.ComboBox cboס��ҩ 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   3915
            Style           =   2  'Dropdown List
            TabIndex        =   5
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
            TabIndex        =   21
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
            TabIndex        =   20
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
            TabIndex        =   18
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
            TabIndex        =   17
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
            TabIndex        =   16
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
            TabIndex        =   15
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
            TabIndex        =   14
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
            TabIndex        =   13
            Top             =   690
            Width           =   900
         End
      End
      Begin VB.Label lblFee 
         AutoSize        =   -1  'True
         Caption         =   "����ȱʡ����"
         Height          =   180
         Left            =   135
         TabIndex        =   22
         Top             =   2430
         Width           =   1080
      End
      Begin VB.Label lbl�շ���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������:"
         Height          =   180
         Left            =   5745
         TabIndex        =   19
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
Private mblnOK As Boolean
Public Function zlEditCard(ByVal frmMain As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '���:frmMain-������������
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-05-30 17:01:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnOK = False
    Me.Show 1, frmMain
    zlEditCard = mblnOK
End Function
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
    Call gobjDatabase.SetPara("����ȱʡ����", IIf(optȱʡ����(0).Value, 0, 1), glngSys, pҽ�����ѹ���, blnSetup)
        
    'ȱʡҩ��
    Call gobjDatabase.SetPara("����ȱʡ��ҩ��", cbo����ҩ.ItemData(cbo����ҩ.ListIndex), glngSys, pҽ�����ѹ���, blnSetup)
    Call gobjDatabase.SetPara("����ȱʡ��ҩ��", cbo�ų�ҩ.ItemData(cbo�ų�ҩ.ListIndex), glngSys, pҽ�����ѹ���, blnSetup)
    Call gobjDatabase.SetPara("����ȱʡ��ҩ��", cbo����ҩ.ItemData(cbo����ҩ.ListIndex), glngSys, pҽ�����ѹ���, blnSetup)
    Call gobjDatabase.SetPara("����ȱʡ���ϲ���", cbo�ŷ��ϲ���.ItemData(cbo�ŷ��ϲ���.ListIndex), glngSys, pҽ�����ѹ���, blnSetup)
    Call gobjDatabase.SetPara("סԺȱʡ��ҩ��", cboס��ҩ.ItemData(cboס��ҩ.ListIndex), glngSys, pҽ�����ѹ���, blnSetup)
    Call gobjDatabase.SetPara("סԺȱʡ��ҩ��", cboס��ҩ.ItemData(cboס��ҩ.ListIndex), glngSys, pҽ�����ѹ���, blnSetup)
    Call gobjDatabase.SetPara("סԺȱʡ��ҩ��", cboס��ҩ.ItemData(cboס��ҩ.ListIndex), glngSys, pҽ�����ѹ���, blnSetup)
    Call gobjDatabase.SetPara("סԺȱʡ���ϲ���", cboס���ϲ���.ItemData(cboס���ϲ���.ListIndex), glngSys, pҽ�����ѹ���, blnSetup)
    
    '�շ����
    strPar = ""
    For i = lst�շ����.ListCount - 1 To 0 Step -1
        If lst�շ����.Selected(i) Then strPar = strPar & "'" & Chr(lst�շ����.ItemData(i)) & "',"
    Next
    If strPar <> "" Then strPar = Left(strPar, Len(strPar) - 1)
    Call gobjDatabase.SetPara("�շ����", Replace(strPar, "'", "''"), glngSys, pҽ�����ѹ���, blnSetup)
    
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
    
    '����:36060
    If Val(gobjDatabase.GetPara("����ȱʡ����", glngSys, pҽ�����ѹ���, , Array(optȱʡ����(0), optȱʡ����(1), lblFee), blnSetup)) = 0 Then
        optȱʡ����(0).Value = True
    Else
        optȱʡ����(1).Value = True
    End If
    
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
    Call gobjDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
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
        If rsTmp!�������� = "���ϲ���" Then
            Set objCbo = IIf(rsTmp!������� = 1, cbo�ŷ��ϲ���, IIf(rsTmp!������� = 2, cboס���ϲ���, Nothing))
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
    strPar = gobjDatabase.GetPara("����ȱʡ��ҩ��", glngSys, pҽ�����ѹ���, , Array(lbl����ҩ, cbo����ҩ), blnSetup)
    For i = 0 To cbo����ҩ.ListCount - 1
        If cbo����ҩ.ItemData(i) = Val(strPar) Then cbo����ҩ.ListIndex = i: Exit For
    Next
    strPar = gobjDatabase.GetPara("����ȱʡ��ҩ��", glngSys, pҽ�����ѹ���, , Array(lbl�ų�ҩ, cbo�ų�ҩ), blnSetup)
    For i = 0 To cbo�ų�ҩ.ListCount - 1
        If cbo�ų�ҩ.ItemData(i) = Val(strPar) Then cbo�ų�ҩ.ListIndex = i: Exit For
    Next
    strPar = gobjDatabase.GetPara("����ȱʡ��ҩ��", glngSys, pҽ�����ѹ���, , Array(lbl����ҩ, cbo����ҩ), blnSetup)
    For i = 0 To cbo����ҩ.ListCount - 1
        If cbo����ҩ.ItemData(i) = Val(strPar) Then cbo����ҩ.ListIndex = i: Exit For
    Next
    strPar = gobjDatabase.GetPara("����ȱʡ���ϲ���", glngSys, pҽ�����ѹ���, , Array(lbl�ŷ��ϲ���, cbo�ŷ��ϲ���), blnSetup)
    For i = 0 To cbo�ŷ��ϲ���.ListCount - 1
        If cbo�ŷ��ϲ���.ItemData(i) = Val(strPar) Then cbo�ŷ��ϲ���.ListIndex = i: Exit For
    Next
    
    
    strPar = gobjDatabase.GetPara("סԺȱʡ��ҩ��", glngSys, pҽ�����ѹ���, , Array(lblס��ҩ, cboס��ҩ), blnSetup)
    For i = 0 To cboס��ҩ.ListCount - 1
        If cboס��ҩ.ItemData(i) = Val(strPar) Then cboס��ҩ.ListIndex = i: Exit For
    Next
    strPar = gobjDatabase.GetPara("סԺȱʡ��ҩ��", glngSys, pҽ�����ѹ���, , Array(lblס��ҩ, cboס��ҩ), blnSetup)
    For i = 0 To cboס��ҩ.ListCount - 1
        If cboס��ҩ.ItemData(i) = Val(strPar) Then cboס��ҩ.ListIndex = i: Exit For
    Next
    strPar = gobjDatabase.GetPara("סԺȱʡ��ҩ��", glngSys, pҽ�����ѹ���, , Array(lblס��ҩ, cboס��ҩ), blnSetup)
    For i = 0 To cboס��ҩ.ListCount - 1
        If cboס��ҩ.ItemData(i) = Val(strPar) Then cboס��ҩ.ListIndex = i: Exit For
    Next
    strPar = gobjDatabase.GetPara("סԺȱʡ���ϲ���", glngSys, pҽ�����ѹ���, , Array(lblס���ϲ���, cboס���ϲ���), blnSetup)
    For i = 0 To cboס���ϲ���.ListCount - 1
        If cboס���ϲ���.ItemData(i) = Val(strPar) Then cboס���ϲ���.ListIndex = i: Exit For
    Next
    
    
    '�շ����
    strSQL = "Select ����,���� as ��� From �շ���Ŀ��� Where ����<>'1' Order by ���"
    Call gobjDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        lst�շ����.AddItem rsTmp!���
        lst�շ����.ItemData(lst�շ����.NewIndex) = Asc(rsTmp!����)
        rsTmp.MoveNext
    Loop
    strPar = gobjDatabase.GetPara("�շ����", glngSys, pҽ�����ѹ���, "", Array(lbl�շ����, lst�շ����), blnSetup)
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
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub
Private Sub lst�շ����_ItemCheck(Item As Integer)
    If lst�շ����.SelCount = 0 And Not lst�շ����.Selected(Item) Then
        lst�շ����.Selected(Item) = True
    End If
End Sub
