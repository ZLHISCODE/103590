VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdviceRollSendCond 
   AutoRedraw      =   -1  'True
   Caption         =   "�ջ�����"
   ClientHeight    =   6705
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   6180
   Icon            =   "frmAdviceRollSendCond.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   6705
   ScaleWidth      =   6180
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   6180
      TabIndex        =   9
      Top             =   6210
      Width           =   6180
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   120
         TabIndex        =   12
         Top             =   0
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   4950
         TabIndex        =   11
         Top             =   0
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   3720
         TabIndex        =   10
         Top             =   0
         Width           =   1100
      End
   End
   Begin VB.Frame fraDetail 
      Height          =   5535
      Left            =   135
      TabIndex        =   5
      Top             =   555
      Width           =   5940
      Begin VB.CheckBox chkOut 
         Alignment       =   1  'Right Justify
         Caption         =   "��ʾ�����Ժ�Ĳ���(&A)"
         Height          =   195
         Left            =   3600
         TabIndex        =   4
         Top             =   5235
         Width           =   2190
      End
      Begin VB.CommandButton cmdAllPati 
         Caption         =   "ȫѡ"
         Height          =   330
         Left            =   270
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + A"
         Top             =   4410
         Width           =   870
      End
      Begin VB.CommandButton cmdNoPati 
         Caption         =   "ȫ��"
         Height          =   330
         Left            =   270
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + R"
         Top             =   4785
         Width           =   870
      End
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   285
         Width           =   2655
      End
      Begin MSComctlLib.ListView lvwPati 
         Height          =   4500
         Left            =   1215
         TabIndex        =   1
         Top             =   645
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   7938
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "סԺ��"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "����"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "סԺҽʦ"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "�ѱ�"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "����ȼ�"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "����"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "��Ժ����"
            Object.Width           =   2857
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "��������"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����(&U)"
         Height          =   180
         Left            =   180
         TabIndex        =   7
         Top             =   345
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����(&P)"
         Height          =   180
         Left            =   180
         TabIndex        =   6
         Top             =   720
         Width           =   990
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ϵͳ���ջز��˳��ڷ����������Ķ�����ü�ҩƷ����Ӳ����嵥��ѡ����Ҫ����Ĳ��ˡ�"
      Height          =   380
      Left            =   1155
      TabIndex        =   8
      Top             =   135
      Width           =   4140
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   390
      Picture         =   "frmAdviceRollSendCond.frx":058A
      Top             =   75
      Width           =   480
   End
End
Attribute VB_Name = "frmAdviceRollSendCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mMainPrivs As String 'IN
Public mlng����ID As Long 'IN/OUT
Public mlng����ID As Long 'IN

Public mblnOK As Boolean 'OUT:�Ƿ�ȷ��
Public mstr����IDs As String 'OUT:����ID��
Public mstr��ҳIDs As String 'OUT:����ID��Ӧ��ҳID��

Private Sub cboUnit_Click()
'���ܣ���ȡָ����Χ�ڵĲ����б�
    Dim rsTmp As ADODB.Recordset
    Dim objItem As ListItem, strSQL As String
    Dim i As Integer, j As Integer, k As Integer
    Dim str����IDs As String, lng����ID As Long, lngUnitID As Long
    Dim lngColor As Long
        
    On Error GoTo errH
    lvwPati.ListItems.Clear
    lngUnitID = cboUnit.ItemData(cboUnit.ListIndex)
    
    str����IDs = zlDatabase.GetPara("���Ͳ���", glngSys, pסԺҽ������)
    If str����IDs <> "" And InStr(str����IDs, ":") > 0 Then
        lng����ID = Val(Split(str����IDs, ":")(0))
        str����IDs = Split(str����IDs, ":")(1)
    End If
            
    Set rsTmp = GetPatiRsByUnit(lngUnitID, mlng����ID, False, False, chkOut.value)
  
    For i = 1 To rsTmp.RecordCount
        If Val(rsTmp!��˱�־ & "") < 1 Or gbyt������˷�ʽ <> 1 Then
            Set objItem = lvwPati.ListItems.Add(, "_" & rsTmp!����ID & "_" & rsTmp!��ҳID, rsTmp!����)
            objItem.SubItems(1) = IIF(IsNull(rsTmp!סԺ��), "", rsTmp!סԺ��)
            objItem.SubItems(2) = IIF(IsNull(rsTmp!����), "", rsTmp!����)
            objItem.SubItems(3) = IIF(IsNull(rsTmp!סԺҽʦ), "", rsTmp!סԺҽʦ)
            objItem.SubItems(4) = IIF(IsNull(rsTmp!�ѱ�), "", rsTmp!�ѱ�)
            objItem.SubItems(5) = IIF(IsNull(rsTmp!����ȼ�), "", rsTmp!����ȼ�)
            objItem.SubItems(6) = IIF(IsNull(rsTmp!����), "", rsTmp!����)
            objItem.SubItems(7) = Format(rsTmp!��Ժ����, "yyyy-MM-dd HH:mm")
            objItem.SubItems(8) = Nvl(rsTmp!��������)
            
            '������ɫ
            lngColor = zlDatabase.GetPatiColor(Nvl(rsTmp!��������))
            objItem.ListSubItems(1).ForeColor = lngColor
            objItem.ListSubItems(8).ForeColor = lngColor
            
            '�ϴ��Ƿ�ѡ��
            If lngUnitID = lng����ID And str����IDs <> "" Then
                If str����IDs = "ALL" _
                    Or Left(str����IDs, 1) <> "-" And InStr("," & str����IDs & ",", "," & rsTmp!����ID & ",") > 0 _
                    Or Left(str����IDs, 1) = "-" And InStr("," & Mid(str����IDs, 2) & ",", "," & rsTmp!����ID & ",") = 0 Then
                    objItem.Checked = True
                    If k = 0 Then 'Ϊ�˿�����ѡ���
                        objItem.EnsureVisible
                        objItem.Selected = True
                        k = 1
                    End If
                End If
            ElseIf rsTmp!����ID = mlng����ID Then
                objItem.Checked = True 'ȱʡֻѡ��ǰ����
                objItem.EnsureVisible
                objItem.Selected = True
            End If
        End If
        rsTmp.MoveNext
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub chkOut_Click()
    If Visible Then Call cboUnit_Click
End Sub

Private Sub cmdAllPati_Click()
    Call SelectLVW(lvwPati, True)
    lvwPati.SetFocus
End Sub

Private Sub SelectLVW(objLVW As Object, ByVal blnCheck As Boolean)
    Dim i As Long
    For i = 1 To objLVW.ListItems.Count
        objLVW.ListItems(i).Checked = blnCheck
    Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNoPati_Click()
    Call SelectLVW(lvwPati, False)
    lvwPati.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim strSel As String, strUnSel As String
    Dim i As Long
    
    If cboUnit.ListIndex = -1 Then
        MsgBox "��ѡ��һ��������", vbInformation, gstrSysName
        cboUnit.SetFocus: Exit Sub
    End If
    mlng����ID = cboUnit.ItemData(cboUnit.ListIndex)
    
    'סԺ����
    mstr����IDs = "": mstr��ҳIDs = ""
    For i = 1 To lvwPati.ListItems.Count
        If lvwPati.ListItems(i).Checked Then
            mstr����IDs = mstr����IDs & "," & Split(Mid(lvwPati.ListItems(i).Key, 2), "_")(0)
            mstr��ҳIDs = mstr��ҳIDs & "," & Split(Mid(lvwPati.ListItems(i).Key, 2), "_")(1)
            strSel = strSel & "," & Split(Mid(lvwPati.ListItems(i).Key, 2), "_")(0)
        Else
            strUnSel = strUnSel & "," & Split(Mid(lvwPati.ListItems(i).Key, 2), "_")(0)
        End If
    Next
    mstr����IDs = Mid(mstr����IDs, 2)
    mstr��ҳIDs = Mid(mstr��ҳIDs, 2)
    If mstr����IDs = "" Then
        MsgBox "������ѡ��һ����Ҫ����ҽ�����ˡ�", vbInformation, gstrSysName
        lvwPati.SetFocus: Exit Sub
    End If
        
    '������������
    strSel = Mid(strSel, 2)
    strUnSel = Mid(strUnSel, 2)
    If strSel = "" Or (UBound(Split(strSel, ",")) = 0 And Val(strSel) = mlng����ID) Then
        strSel = ""
    Else
        If strUnSel = "" Then
            strSel = cboUnit.ItemData(cboUnit.ListIndex) & ":ALL"
        ElseIf UBound(Split(strSel, ",")) > UBound(Split(strUnSel, ",")) Then
            strSel = cboUnit.ItemData(cboUnit.ListIndex) & ":-" & strUnSel
        Else
            strSel = cboUnit.ItemData(cboUnit.ListIndex) & ":" & strSel
        End If
    End If
    Call zlDatabase.SetPara("���Ͳ���", strSel, glngSys, pסԺҽ������)
        
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        Call cmdAllPati_Click
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call cmdNoPati_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    mblnOK = False
    Call RestoreWinState(Me, App.ProductName)
    '����/����
    Call RestoreListViewState(Me.lvwPati, App.ProductName & Me.Name, "")
    'Call zlControl.LvwFlatColumnHeader(lvwPati)
    Call InitUnits
End Sub

Private Function InitUnits() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, strSQL As String
    
    On Error GoTo errH
    
    '��������۲���
    If InStr(mMainPrivs, "ȫԺ����") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where A.ID=B.����ID And B.������� in(1,2,3) And B.��������='����'" & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " Order by A.����"
    Else
        '����Ȩ������ֱ�����ڲ���+���ڿ�����������
        strSQL = _
            " Select A.ID,A.����,A.����,Nvl(C.ȱʡ,0) as ȱʡ" & _
            " From ���ű� A,��������˵�� B,������Ա C" & _
            " Where A.ID=B.����ID And A.ID=C.����ID And C.��ԱID=[1]" & _
            " And B.������� in(1,2,3) And B.��������='����'" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = strSQL & " Union " & _
            " Select C.ID,C.����,C.����,Nvl(B.ȱʡ,0) as ȱʡ" & _
            " From �������Ҷ�Ӧ A,������Ա B,���ű� C" & _
            " Where A.����ID=C.ID And B.����ID=A.����ID And B.��ԱID=[1]" & _
            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
            " And (C.����ʱ�� is NULL or Trunc(C.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = "Select ID,����,����,Max(ȱʡ) as ȱʡ From (" & strSQL & ") Group by ID,����,���� Order by ����"
    End If
    
    cboUnit.Clear
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!���� & "-" & rsTmp!����
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            If rsTmp!ID = mlng����ID Then cboUnit.ListIndex = cboUnit.NewIndex
            rsTmp.MoveNext
        Next
    End If
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    fraDetail.Width = Me.ScaleWidth - 240
    chkOut.Left = fraDetail.Width - chkOut.Width - 120
    lvwPati.Width = fraDetail.Width - lvwPati.Left - 120
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 120
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 60
        
    fraDetail.Height = Me.ScaleHeight - picBottom.Height - fraDetail.Top - 120
    
    chkOut.Top = fraDetail.Height - chkOut.Height - 60
    lvwPati.Height = chkOut.Top - lvwPati.Top - 60
    
    cmdNoPati.Top = lvwPati.Top + lvwPati.Height - 30 - cmdNoPati.Height
    cmdAllPati.Top = cmdNoPati.Top - cmdAllPati.Height - 30
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '�ͷ�˽�м�IN����
    mMainPrivs = ""
    mlng����ID = 0
    
    Call SaveListViewState(Me.lvwPati, App.ProductName & Me.Name, "")
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwPati, ColumnHeader.Index)
End Sub
