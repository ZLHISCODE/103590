VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServicePara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ҺŲ�������"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6900
   Icon            =   "frmServicePara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraBespeak 
      Caption         =   "ԤԼ�Һŵ�"
      Height          =   705
      Left            =   135
      TabIndex        =   16
      Top             =   870
      Width           =   6675
      Begin VB.OptionButton optPrintBespeak 
         Caption         =   "ѡ���Ƿ��ӡ"
         Height          =   180
         Index           =   2
         Left            =   3885
         TabIndex        =   6
         Top             =   300
         Width           =   1380
      End
      Begin VB.OptionButton optPrintBespeak 
         Caption         =   "����ӡ"
         Height          =   180
         Index           =   0
         Left            =   675
         TabIndex        =   4
         Top             =   315
         Width           =   900
      End
      Begin VB.OptionButton optPrintBespeak 
         Caption         =   "�Զ���ӡ"
         Height          =   180
         Index           =   1
         Left            =   2130
         TabIndex        =   5
         Top             =   300
         Value           =   -1  'True
         Width           =   1020
      End
   End
   Begin VB.ComboBox cboDefaultStyle 
      Height          =   300
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   435
      Width           =   2160
   End
   Begin VB.TextBox txtAuto 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   1665
      TabIndex        =   2
      Text            =   "0"
      Top             =   150
      Width           =   435
   End
   Begin VB.CheckBox chkAuto 
      Caption         =   "�Զ�ˢ�¼��:      ��"
      Height          =   300
      Left            =   195
      TabIndex        =   1
      Top             =   105
      Width           =   2700
   End
   Begin VB.CommandButton cmdDeviceSetup 
      Caption         =   "�豸����(&S)"
      Height          =   330
      Left            =   135
      TabIndex        =   14
      Top             =   4680
      Width           =   1425
   End
   Begin VB.Frame fraPrintSet 
      Caption         =   "��ӡ����"
      Height          =   720
      Left            =   135
      TabIndex        =   13
      Top             =   1755
      Width           =   6675
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "ԤԼ�Һŵ���ӡ����"
         Height          =   345
         Index           =   2
         Left            =   4575
         TabIndex        =   9
         Top             =   240
         Width           =   1890
      End
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "�Һ�ƾ����ӡ����"
         Height          =   345
         Index           =   1
         Left            =   2385
         TabIndex        =   8
         Top             =   240
         Width           =   1890
      End
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "�Һ�Ʊ�ݴ�ӡ����"
         Height          =   345
         Index           =   0
         Left            =   195
         TabIndex        =   7
         Top             =   240
         Width           =   1890
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4335
      TabIndex        =   11
      Top             =   4665
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5520
      TabIndex        =   12
      Top             =   4665
      Width           =   1100
   End
   Begin VB.Frame fraTitle 
      Caption         =   "���ùҺ�Ʊ��"
      Height          =   1845
      Left            =   135
      TabIndex        =   0
      Top             =   2670
      Width           =   6675
      Begin MSComctlLib.ListView lvwBill 
         Height          =   1455
         Left            =   150
         TabIndex        =   10
         Top             =   240
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "������"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "��������"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "���뷶Χ"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "ʣ��"
            Object.Width           =   1499
         EndProperty
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ȱʡԤԼ��ʽ"
      Height          =   180
      Left            =   195
      TabIndex        =   15
      Top             =   495
      Width           =   1080
   End
End
Attribute VB_Name = "frmServicePara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub zlShowMe(ByVal frmMain As Object, ByVal lngModul As Long)
    Me.Show vbModal, frmMain
End Sub

Private Sub chkAuto_Click()
    If chkAuto.Value = 1 Then
        txtAuto.Enabled = True
    Else
        txtAuto.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1111)
End Sub

Private Sub cmdOk_Click()
    Dim strTMP As String, i As Integer
    '���ùҺ�Ʊ������
    strTMP = "0"
    For i = 1 To lvwBill.ListItems.Count
        If lvwBill.ListItems(i).Checked Then strTMP = Mid(lvwBill.ListItems(i).Key, 2)
    Next
    zlDatabase.SetPara "���ùҺ�Ʊ������", strTMP, glngSys, 1111, True
    zlDatabase.SetPara "ȱʡԤԼ��ʽ", NeedName(cboDefaultStyle.Text), glngSys, 9000, True
    If chkAuto.Value = 1 Then
        zlDatabase.SetPara "ˢ�·�ʽ", "1" & "|" & Val(txtAuto.Text), glngSys, 1115, True
    Else
        zlDatabase.SetPara "ˢ�·�ʽ", 0, glngSys, 1115, True
    End If
    For i = 0 To Me.optPrintBespeak.UBound
        If optPrintBespeak(i).Value Then
            zlDatabase.SetPara "ԤԼ�Һŵ���ӡ��ʽ", i, glngSys, 9000, True
            Exit For
        End If
    Next
    Unload Me
End Sub

Private Sub cmdPrintSet_Click(index As Integer)
    Select Case index
    Case 0
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111", Me)
    Case 1
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1802", Me)
    Case 2
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_1", Me)
    End Select
End Sub

Private Sub Form_Load()
    Dim strTMP As String, i As Integer
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim intIndex As Integer
    strTMP = zlDatabase.GetPara("ȱʡԤԼ��ʽ", glngSys, 9000, "", Array(cboDefaultStyle), True)
    strSQL = "Select ����,����,ȱʡ��־ From ԤԼ��ʽ Order By ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cboDefaultStyle.Clear
    Do While Not rsTmp.EOF
        cboDefaultStyle.AddItem rsTmp!���� & "-" & rsTmp!����
        If strTMP = Nvl(rsTmp!����) Then intIndex = cboDefaultStyle.NewIndex
        If Val(Nvl(rsTmp!ȱʡ��־)) = 1 Then cboDefaultStyle.ListIndex = cboDefaultStyle.NewIndex
        rsTmp.MoveNext
    Loop
    If cboDefaultStyle.ListCount <> 0 And intIndex <> 0 Then cboDefaultStyle.ListIndex = intIndex
    strTMP = zlDatabase.GetPara("ˢ�·�ʽ", glngSys, 1115, "0", Array(chkAuto, txtAuto), True) & "|"
    chkAuto.Value = IIf(Split(strTMP, "|")(0) = "1", 1, 0)
    If chkAuto.Value = 0 Then
        txtAuto.Text = "0"
        txtAuto.Enabled = False
    Else
        txtAuto.Text = Val(Split(strTMP, "|")(1))
        txtAuto.Enabled = True
    End If
    Call LoadFactList
    i = Val(zlDatabase.GetPara("ԤԼ�Һŵ���ӡ��ʽ", glngSys, 9000, 1, Array(optPrintBespeak(0), optPrintBespeak(1), optPrintBespeak(2)), True))
    If i <= optPrintBespeak.UBound Then optPrintBespeak(i).Value = True
End Sub

Private Function LoadFactList() As Boolean
'���ܣ���ȡ���ù��ùҺ�Ʊ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer, lngTmp As Long
    Dim objItem As ListItem
    Dim blnBill As Boolean
    
    On Error GoTo errH
    lngTmp = zlDatabase.GetPara("���ùҺ�Ʊ������", glngSys, 1111, 0, Array(lvwBill), True)
    Set rsTmp = GetShareInvoiceGroupID
    
    For i = 1 To rsTmp.RecordCount
        Set objItem = lvwBill.ListItems.Add(, "_" & rsTmp!ID, rsTmp!������)
        objItem.SubItems(1) = Format(rsTmp!�Ǽ�ʱ��, "yyyy-MM-dd")
        objItem.SubItems(2) = rsTmp!��ʼ���� & "," & rsTmp!��ֹ����
        objItem.SubItems(3) = rsTmp!ʣ������
        If rsTmp!ID = lngTmp Then
            objItem.Checked = True
            objItem.Selected = True
            blnBill = True
        End If
        rsTmp.MoveNext
    Next
    
    If Not blnBill Then
        zlDatabase.SetPara "���ùҺ�Ʊ������", "0", glngSys, 9000, True
    End If
    
    LoadFactList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetShareInvoiceGroupID() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ��Ʊ�ֵĹ���Ʊ������
    '����:���˺�
    '����:2011-04-29 10:24:48
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errH
    
    strSQL = "" & _
    "   Select A.ID,A.ʹ�����,A.������,A.�Ǽ�ʱ��,A.��ʼ����,A.��ֹ����,A.ʣ������ " & _
    "   From Ʊ�����ü�¼ A,��Ա�� B" & vbNewLine & _
    "   Where A.Ʊ��=4 And A.ʹ�÷�ʽ=2 And A.ʣ������>0 And A.������=B.����" & _
    "           And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & vbNewLine & _
    "   Order by ʹ�����,ʣ������ Desc"
    
    Set GetShareInvoiceGroupID = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
