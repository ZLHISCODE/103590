VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChargeTurnParSet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�������תסԺ��������"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "Ԥ����Ʊ�ݴ�ӡ����"
      Height          =   345
      Left            =   240
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1830
   End
   Begin VB.Frame Frame2 
      Caption         =   "���ع���Ԥ����Ʊ��"
      Height          =   1695
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   5040
      Begin MSComctlLib.ListView lvwDeposit 
         Height          =   1395
         Left            =   90
         TabIndex        =   6
         Top             =   210
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   2461
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "img16"
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "������"
            Object.Width           =   1587
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
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "ʣ��"
            Object.Width           =   1499
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   60
      TabIndex        =   3
      Top             =   3270
      Width           =   5670
   End
   Begin VB.Frame fraTopSplit 
      Height          =   45
      Left            =   -30
      TabIndex        =   2
      Top             =   855
      Width           =   5670
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4095
      TabIndex        =   1
      Top             =   3480
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2910
      TabIndex        =   0
      Top             =   3480
      Width           =   1100
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   1
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeTurnParSet.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "   �������תסԺ���õ���ز������ã������ò���ʱ����ע��������ĺ��塣"
      Height          =   375
      Left            =   735
      TabIndex        =   4
      Top             =   330
      Width           =   4380
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   105
      Picture         =   "frmChargeTurnParSet.frx":00E2
      Top             =   225
      Width           =   480
   End
End
Attribute VB_Name = "frmChargeTurnParSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mstrPrivs As String, mlngModule As Long
Private mblnOk As Boolean
Public Function ShowSet(ByVal frmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ������
    '����:�ɹ�������true,���򷵻�False
    '����:���˺�
    '����:2011-02-16 09:50:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mstrPrivs = strPrivs: mlngModule = lngModule: mblnOk = False
    Me.Show 1, frmMain
    ShowSet = mblnOk
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1131)
End Sub
 
Private Sub cmdOK_Click()
    Dim i As Integer
    '���ع���Ԥ����Ʊ��
    zlDatabase.SetPara "����Ԥ��Ʊ������", 0, glngSys, mlngModule, True
    For i = 1 To lvwDeposit.ListItems.Count
        If lvwDeposit.ListItems(i).Checked Then
            zlDatabase.SetPara "����Ԥ��Ʊ������", Mid(lvwDeposit.SelectedItem.Key, 2), glngSys, mlngModule, True
        End If
    Next
        
    mblnOk = True
    Unload Me
End Sub

Private Sub cmdPrintSet_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdOK_Click
End Sub

Private Sub Form_Load()
    Dim i As Integer, blnBill As Boolean
    Dim rsTmp As ADODB.Recordset, objItem As ListItem
    On Error GoTo errH
    '��ȡ���ù���Ԥ������:    '����:36984
    Set rsTmp = GetShareInvoiceGroupID(2)
    blnBill = False
    rsTmp.Filter = "ʹ�����=2"
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            Set objItem = lvwDeposit.ListItems.Add(, "_" & rsTmp!ID, rsTmp!������, , 1)
            objItem.SubItems(1) = Format(rsTmp!�Ǽ�ʱ��, "yyyy-MM-dd")
            objItem.SubItems(2) = rsTmp!��ʼ���� & "," & rsTmp!��ֹ����
            objItem.SubItems(3) = rsTmp!ʣ������
            If rsTmp!ID = zlDatabase.GetPara("����Ԥ��Ʊ������", glngSys, mlngModule) Then
                objItem.Selected = True
                objItem.Checked = True
                blnBill = True
            End If
            rsTmp.MoveNext
        Next
    End If
    If Not blnBill Then zlDatabase.SetPara "����Ԥ��Ʊ������", 0, glngSys, mlngModule, True
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub lvwDeposit_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer
    For i = 1 To lvwDeposit.ListItems.Count
        If lvwDeposit.ListItems(i).Key <> Item.Key Then lvwDeposit.ListItems(i).Checked = False
    Next
    Item.Selected = True
End Sub
