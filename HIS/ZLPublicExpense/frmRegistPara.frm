VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRegistPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ҺŲ�������"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6900
   Icon            =   "frmRegistPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboApp 
      ForeColor       =   &H80000012&
      Height          =   300
      Left            =   4785
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   135
      Width           =   1350
   End
   Begin VB.Frame fraPrintSet 
      Caption         =   "��ӡ����"
      Height          =   720
      Left            =   135
      TabIndex        =   10
      Top             =   885
      Width           =   6675
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "ԤԼ�Һŵ���ӡ����"
         Height          =   345
         Index           =   2
         Left            =   4575
         TabIndex        =   6
         Top             =   240
         Width           =   1890
      End
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "�Һ�ƾ����ӡ����"
         Height          =   345
         Index           =   1
         Left            =   2385
         TabIndex        =   5
         Top             =   240
         Width           =   1890
      End
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "�Һ�Ʊ�ݴ�ӡ����"
         Height          =   345
         Index           =   0
         Left            =   195
         TabIndex        =   4
         Top             =   240
         Width           =   1890
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   4335
      TabIndex        =   8
      Top             =   3675
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   5520
      TabIndex        =   9
      Top             =   3675
      Width           =   1100
   End
   Begin VB.CheckBox chkDefaultInputRemark 
      Caption         =   "�Һ�ʱĬ������ժҪ"
      Height          =   300
      Left            =   135
      TabIndex        =   1
      Top             =   135
      Width           =   2700
   End
   Begin VB.CheckBox chkDefaultMedBook 
      Caption         =   "�Һ�ʱĬ�Ϲ�ѡ������ѡ��"
      Height          =   300
      Left            =   135
      TabIndex        =   3
      Top             =   465
      Width           =   2700
   End
   Begin VB.Frame fraTitle 
      Caption         =   "���ùҺ�Ʊ��"
      Height          =   1845
      Left            =   135
      TabIndex        =   0
      Top             =   1680
      Width           =   6675
      Begin MSComctlLib.ListView lvwBill 
         Height          =   1455
         Left            =   150
         TabIndex        =   7
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
   Begin VB.Label lblAppStyle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ȱʡԤԼ��ʽ"
      Height          =   180
      Left            =   3660
      TabIndex        =   11
      Top             =   195
      Width           =   1080
   End
End
Attribute VB_Name = "frmRegistPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub zlShowMe(ByVal frmMain As Object, ByVal lngModul As Long)
    Me.Show vbModal, frmMain
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strTmp As String, i As Integer
    '���ùҺ�Ʊ������
    strTmp = "0"
    For i = 1 To lvwBill.ListItems.Count
        If lvwBill.ListItems(i).Checked Then strTmp = Mid(lvwBill.ListItems(i).Key, 2)
    Next
    gobjDatabase.SetPara "���ùҺ�Ʊ������", strTmp, glngSys, 9000, True
    gobjDatabase.SetPara "Ĭ�Ϲ�����", chkDefaultMedBook.Value, glngSys, 9000, True
    gobjDatabase.SetPara "Ĭ������ժҪ", chkDefaultInputRemark.Value, glngSys, 9000, True
    gobjDatabase.SetPara "ȱʡԤԼ��ʽ", gobjCommFun.GetNeedName(cboApp.Text, "-"), glngSys, 9000, True
    Unload Me
End Sub

Private Sub cmdPrintSet_Click(Index As Integer)
    Select Case Index
    Case 0
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111", Me)
    Case 1
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1802", Me)
    Case 2
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_1", Me)
    End Select
End Sub

Private Sub Form_Load()
    Dim strTmp As String, i As Integer
    Dim rsTmp As ADODB.Recordset, strValue As String
    chkDefaultMedBook.Value = IIf(gobjDatabase.GetPara("Ĭ�Ϲ�����", glngSys, 9000, , Array(chkDefaultMedBook), True) = "1", 1, 0)
    chkDefaultInputRemark.Value = IIf(gobjDatabase.GetPara("Ĭ������ժҪ", glngSys, 9000, , Array(chkDefaultInputRemark), True) = "1", 1, 0)
    With cboApp
        .Clear
        strTmp = "Select ����,���� From ԤԼ��ʽ"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strTmp, Me.Caption)
        Do While Not rsTmp.EOF
            .AddItem rsTmp!���� & "-" & rsTmp!����
            rsTmp.MoveNext
        Loop
    End With
    strValue = gobjDatabase.GetPara("ȱʡԤԼ��ʽ", glngSys, 9000, , Array(cboApp), True)
    With cboApp
        For i = 0 To .ListCount - 1
            If gobjCommFun.GetNeedName(.List(i), "-") = strValue Then
                .ListIndex = i: Exit For
            End If
        Next i
        If .ListIndex < 0 Then .ListIndex = 0
    End With
    Call LoadFactList
End Sub

Private Function LoadFactList() As Boolean
'���ܣ���ȡ���ù��ùҺ�Ʊ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer, lngTmp As Long
    Dim ObjItem As ListItem
    Dim blnBill As Boolean
    
    On Error GoTo errH
    lngTmp = gobjDatabase.GetPara("���ùҺ�Ʊ������", glngSys, 9000, 0, Array(lvwBill), True)
    Set rsTmp = GetShareInvoiceGroupID
    
    For i = 1 To rsTmp.RecordCount
        Set ObjItem = lvwBill.ListItems.Add(, "_" & rsTmp!ID, rsTmp!������)
        ObjItem.SubItems(1) = Format(rsTmp!�Ǽ�ʱ��, "yyyy-MM-dd")
        ObjItem.SubItems(2) = rsTmp!��ʼ���� & "," & rsTmp!��ֹ����
        ObjItem.SubItems(3) = rsTmp!ʣ������
        If rsTmp!ID = lngTmp Then
            ObjItem.Checked = True
            ObjItem.Selected = True
            blnBill = True
        End If
        rsTmp.MoveNext
    Next
    
    If Not blnBill Then
        gobjDatabase.SetPara "���ùҺ�Ʊ������", "0", glngSys, 9000, True
    End If
    
    LoadFactList = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
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
    
    Set GetShareInvoiceGroupID = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

