VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdviceSendOtherCond 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   Icon            =   "frmAdviceSendOtherCond.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraDetail 
      Height          =   5445
      Index           =   0
      Left            =   135
      TabIndex        =   22
      Top             =   60
      Width           =   5460
      Begin VB.CheckBox chkBaby 
         Caption         =   "Ӥ��ҽ��"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   14
         Top             =   2430
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.CheckBox chkBaby 
         Caption         =   "����ҽ��"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   13
         Top             =   2175
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.ComboBox cboTime 
         Height          =   300
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   540
         Width           =   1995
      End
      Begin VB.CommandButton cmdִ�п��� 
         Height          =   240
         Left            =   5010
         Picture         =   "frmAdviceSendOtherCond.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "ѡ��ִ�п���(F4)"
         Top             =   930
         Width           =   270
      End
      Begin VB.CheckBox chk�Ӱ�Ӽ� 
         Caption         =   "ִ�мӰ�Ӽ�(&V)"
         Height          =   195
         Left            =   3255
         TabIndex        =   2
         Top             =   255
         Width           =   1650
      End
      Begin VB.ListBox lstClass 
         Columns         =   4
         Height          =   1110
         ItemData        =   "frmAdviceSendOtherCond.frx":0680
         Left            =   1215
         List            =   "frmAdviceSendOtherCond.frx":0682
         Style           =   1  'Checkbox
         TabIndex        =   18
         Top             =   4230
         Width           =   4095
      End
      Begin VB.OptionButton opt��Ч 
         Caption         =   "����(&L)"
         Height          =   180
         Index           =   0
         Left            =   1215
         TabIndex        =   0
         Top             =   255
         Value           =   -1  'True
         Width           =   930
      End
      Begin VB.OptionButton opt��Ч 
         Caption         =   "����(&T)"
         Height          =   180
         Index           =   1
         Left            =   2190
         TabIndex        =   1
         Top             =   255
         Width           =   930
      End
      Begin VB.CommandButton cmdAllPati 
         Caption         =   "ȫѡ"
         Height          =   330
         Left            =   270
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + A"
         Top             =   2955
         Width           =   870
      End
      Begin VB.CommandButton cmdNoPati 
         Caption         =   "ȫ��"
         Height          =   330
         Left            =   270
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + R"
         Top             =   3330
         Width           =   870
      End
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1260
         Width           =   4095
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   3240
         TabIndex        =   5
         Top             =   540
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   65798147
         CurrentDate     =   37953
      End
      Begin MSComctlLib.ListView lvwPati 
         Height          =   2070
         Left            =   1215
         TabIndex        =   12
         Top             =   1620
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   3651
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   1764
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
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "ʣ���"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "סԺҽʦ"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "�ѱ�"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "����ȼ�"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "����"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "��Ժ����"
            Object.Width           =   2857
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "��������"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtִ�п��� 
         Height          =   300
         Left            =   1215
         TabIndex        =   7
         Top             =   900
         Width           =   4095
      End
      Begin MSComctlLib.Toolbar tbrAutoSel 
         Height          =   360
         Left            =   1215
         TabIndex        =   23
         Top             =   3750
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   635
         ButtonWidth     =   5318
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���������������ſ�Ƿ�Ѳ���   "
               Object.ToolTipText     =   "Ctrl + Q"
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label lblִ�п��� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ִ�п���(&D)"
         Height          =   180
         Left            =   180
         TabIndex        =   6
         Top             =   960
         Width           =   990
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   105
         X2              =   5360
         Y1              =   4170
         Y2              =   4170
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   105
         X2              =   5360
         Y1              =   4155
         Y2              =   4155
      End
      Begin VB.Label lbl������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������(&T)"
         Height          =   180
         Left            =   180
         TabIndex        =   17
         Top             =   4275
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����(&U)"
         Height          =   180
         Left            =   180
         TabIndex        =   9
         Top             =   1320
         Width           =   990
      End
      Begin VB.Label lbl����ʱ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��(&E)"
         Height          =   180
         Left            =   180
         TabIndex        =   3
         Top             =   600
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����(&P)"
         Height          =   180
         Left            =   180
         TabIndex        =   11
         Top             =   1695
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   495
      TabIndex        =   21
      Top             =   5610
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3990
      TabIndex        =   20
      Top             =   5610
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2760
      TabIndex        =   19
      Top             =   5610
      Width           =   1100
   End
End
Attribute VB_Name = "frmAdviceSendOtherCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mMainPrivs As String 'IN
Public mlng����ID As Long 'IN/OUT
Public mlng����ID As Long 'IN
Public mblnOK As Boolean 'OUT:�Ƿ�ȷ��
Public mstrEnd As String 'OUT:����ʱ��
Public mint��Ч As Integer 'OUT:0-����,1-����
Public mlngִ�п���ID As Long 'OUT-���͵�ִ�п���
Public mstr����IDs As String 'OUT:����ID��
Public mintӤ�� As Integer 'IN/OUT:Ӥ��ҽ������
Public mstr���s As String 'OUT:�������

Private mrsWarn As ADODB.Recordset
Private mstrLike As String

Private Sub cboTime_Click()
    Dim curDate As Date
    Dim strTmp As String, lngTmp As Long
    
    dtpEnd.Enabled = cboTime.ListIndex = cboTime.ListCount - 1
    
    curDate = zlDatabase.Currentdate
    Select Case cboTime.ListIndex
    Case 0 '����
        dtpEnd.Value = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case 1 '����
        dtpEnd.Value = Format(curDate + 1, "yyyy-MM-dd 23:59:59")
    Case 2 '����
        dtpEnd.Value = Format(curDate + 2, "yyyy-MM-dd 23:59:59")
    Case 3 '[ָ��..]
        strTmp = zlDatabase.GetPara("�������ͽ���ʱ��", glngSys, pסԺҽ������, "23:59:59", Array(dtpEnd), InStr(GetInsidePrivs(pסԺҽ������), "ҽ��ѡ������") > 0)
        lngTmp = Val(zlDatabase.GetPara("��������ʱ����", glngSys, pסԺҽ������, "0", Array(dtpEnd), InStr(GetInsidePrivs(pסԺҽ������), "ҽ��ѡ������") > 0))
        dtpEnd.Value = Format(curDate + lngTmp, "yyyy-MM-dd " & strTmp)
        If Me.Visible Then dtpEnd.SetFocus
    End Select
End Sub

Private Sub cboUnit_Click()
'���ܣ���ȡָ����Χ�ڵĲ����б�
    Dim rsTmp As New ADODB.Recordset
    Dim objItem As ListItem, strSQL As String
    Dim i As Integer, j As Integer, k As Integer
    Dim str����IDs As String, lng����ID As Long
        
    lvwPati.ListItems.Clear
    
    On Error GoTo errH
    
    strSQL = "Select ���ò���,��������,����ֵ,������־1,������־2,������־3 From ���ʱ����� Where ����ID=[1]"
    Set mrsWarn = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboUnit.ItemData(cboUnit.ListIndex))
    
    str����IDs = zlDatabase.GetPara("�������Ͳ���", glngSys, pסԺҽ������)
    If str����IDs <> "" And InStr(str����IDs, ":") > 0 Then
        lng����ID = Val(Split(str����IDs, ":")(0))
        str����IDs = Split(str����IDs, ":")(1)
    End If
        
    '��Ժ����:��Ժ���˽�ֹ��ҽ��,����ҽ��
    strSQL = _
        "Select A.����ID,A.����,B.סԺ��,B.��Ժ���� as ����," & _
        " Nvl(E.Ԥ�����,0)-Nvl(E.�������,0)+Decode(B.����,Null,0,Nvl(F.���,0)) as ʣ���," & _
        " A.������,zl_PatiWarnScheme(A.����ID,B.��ҳID) as ���ò���,B.����," & _
        " B.סԺҽʦ,B.�ѱ�,D.���� as ����ȼ�,C.���� as ����,B.��Ժ����,B.��������" & _
        " From ������Ϣ A,������ҳ B,���ű� C,�շ���ĿĿ¼ D,������� E," & _
        " (Select ����ID,��ҳID,Sum(���) As ��� From ����ģ����� Group By ����ID,��ҳID) F" & _
        " Where A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And B.��Ժ����ID=C.ID" & _
        " And A.����ID=E.����ID(+) And E.����(+)=1 And B.����ID=F.����ID(+) And B.��ҳID=F.��ҳID(+)" & _
        " And B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3 And A.��Ժ=1 And B.����ȼ�ID=D.ID(+)" & _
        IIF(cboUnit.ItemData(cboUnit.ListIndex) > 0, " And B.��ǰ����ID+0=[1]", "") & _
        IIF(cboUnit.ItemData(cboUnit.ListIndex) = 0, " Order by B.סԺ�� Desc", " Order by LPAD(B.��Ժ����,10,' ')")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboUnit.ItemData(cboUnit.ListIndex))
    For i = 1 To rsTmp.RecordCount
        Set objItem = lvwPati.ListItems.Add(, "_" & rsTmp!����ID, rsTmp!����)
        objItem.SubItems(1) = IIF(IsNull(rsTmp!סԺ��), "", rsTmp!סԺ��)
        objItem.SubItems(2) = IIF(IsNull(rsTmp!����), "", rsTmp!����)
        objItem.SubItems(3) = Format(Nvl(rsTmp!ʣ���, 0), "0.00")
        objItem.SubItems(4) = IIF(IsNull(rsTmp!סԺҽʦ), "", rsTmp!סԺҽʦ)
        objItem.SubItems(5) = IIF(IsNull(rsTmp!�ѱ�), "", rsTmp!�ѱ�)
        objItem.SubItems(6) = IIF(IsNull(rsTmp!����ȼ�), "", rsTmp!����ȼ�)
        objItem.SubItems(7) = IIF(IsNull(rsTmp!����), "", rsTmp!����)
        objItem.SubItems(8) = Format(rsTmp!��Ժ����, "yyyy-MM-dd HH:mm")
        objItem.SubItems(9) = Nvl(rsTmp!��������)
        
        '������Ϣ
        objItem.ListSubItems(1).Tag = Nvl(rsTmp!���ò���)
        objItem.ListSubItems(2).Tag = Nvl(rsTmp!������, 0)
        
        '������ɫ
        objItem.ForeColor = zlDatabase.GetPatiColor(Nvl(rsTmp!��������))
        For j = 1 To objItem.ListSubItems.Count
            objItem.ListSubItems(j).ForeColor = objItem.ForeColor
        Next
        
        '�ϴ��Ƿ�ѡ��
        If cboUnit.ItemData(cboUnit.ListIndex) = lng����ID And str����IDs <> "" Then
            If InStr("," & str����IDs & ",", "," & rsTmp!����ID & ",") > 0 Then
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
        rsTmp.MoveNext
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub chkBaby_Click(Index As Integer)
    If chkBaby(0).Value = 0 And chkBaby(1).Value = 0 Then
        chkBaby(Index).Value = 1
    End If
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
    Dim i As Long
    
    If cboUnit.ListIndex = -1 Then
        MsgBox "��ѡ��һ��������", vbInformation, gstrSysName
        cboUnit.SetFocus: Exit Sub
    End If
    mlng����ID = cboUnit.ItemData(cboUnit.ListIndex)
    
    'ʱ�����Ч
    mint��Ч = IIF(opt��Ч(1).Value, 1, 0)
    If opt��Ч(0).Value Then
        mstrEnd = Format(dtpEnd.Value, "yyyy-MM-dd HH:mm:ss")
    Else
        mstrEnd = ""
    End If
    
    'ִ�п���
    mlngִ�п���ID = Val(cmdִ�п���.Tag)
    
    'סԺ����
    mstr����IDs = ""
    For i = 1 To lvwPati.ListItems.Count
        If lvwPati.ListItems(i).Checked Then
            mstr����IDs = mstr����IDs & "," & Mid(lvwPati.ListItems(i).Key, 2)
        End If
    Next
    mstr����IDs = Mid(mstr����IDs, 2)
    If mstr����IDs = "" Then
        MsgBox "������ѡ��һ����Ҫ����ҽ�����ˡ�", vbInformation, gstrSysName
        lvwPati.SetFocus: Exit Sub
    End If
        
    '�������
    mstr���s = ""
    For i = 0 To lstClass.ListCount - 1
        If lstClass.Selected(i) Then
            mstr���s = mstr���s & ",'" & Chr(lstClass.ItemData(i)) & "'"
        End If
    Next
    mstr���s = Mid(mstr���s, 2)
    If mstr���s = "" Then
        MsgBox "������ѡ��һ���������", vbInformation, gstrSysName
        lstClass.SetFocus: Exit Sub
    End If
    If UBound(Split(mstr���s, ",")) + 1 = lstClass.ListCount Then
        mstr���s = ""
    End If
    
    gbln�Ӱ�Ӽ� = chk�Ӱ�Ӽ�.Value = 1
    
    'Ӥ��ҽ��
    If chkBaby(0).Value = 1 And chkBaby(1).Value = 1 Then
        mintӤ�� = -1
    Else
        mintӤ�� = IIF(chkBaby(0).Value = 1, 0, 1)
    End If
    
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdִ�п���_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vRect As RECT
    
    strSQL = _
        " Select 0 as ID,'-' as ����,'���п���' as ����,NULL as ���� From Dual" & _
        " Union ALL" & _
        " Select Distinct A.ID,A.����,A.����,A.����" & _
        " From ���ű� A,��������˵�� B" & _
        " Where A.ID=B.����ID And B.������� IN(2,3)" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " Order by ����"
    vRect = GetControlRect(txtִ�п���.Hwnd)
    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "ִ�п���", , , , , , True, vRect.Left, vRect.Top, txtִ�п���.Height, blnCancel, , True)
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "û�п��õĿ��ң����ȵ����Ź��������á�", vbInformation, gstrSysName
        End If
        txtִ�п���.Text = txtִ�п���.Tag
        Call zlControl.TxtSelAll(txtִ�п���)
    Else
        txtִ�п���.Text = rsTmp!����
        txtִ�п���.Tag = rsTmp!����
        cmdִ�п���.Tag = rsTmp!ID
    End If
    txtִ�п���.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, j As Long
    
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If ActiveControl Is lstClass Then
            j = lstClass.ListIndex
            For i = 0 To lstClass.ListCount - 1
                lstClass.Selected(i) = True
            Next
            lstClass.ListIndex = j
        Else
            Call cmdAllPati_Click
        End If
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        If ActiveControl Is lstClass Then
            j = lstClass.ListIndex
            For i = 0 To lstClass.ListCount - 1
                lstClass.Selected(i) = False
            Next
            lstClass.ListIndex = j
        Else
            Call cmdNoPati_Click
        End If
    ElseIf KeyCode = vbKeyQ And Shift = vbCtrlMask Then
        If tbrAutoSel.Visible Then
            Call tbrAutoSel_ButtonClick(tbrAutoSel.Buttons(1))
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not Me.ActiveControl Is txtִ�п��� Then
            KeyAscii = 0
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim curDate As Date
    Dim strTmp As String, lngTmp As Long
    
    Call RestoreListViewState(Me.lvwPati, App.ProductName & Me.Name, "")
    
    mblnOK = False
    
    '����ƥ��
    mstrLike = IIF(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
    
    'ȱʡҽ����Ч
    lngTmp = Val(zlDatabase.GetPara("��������ҽ����Ч", glngSys, pסԺҽ������, "0", Array(opt��Ч(0), opt��Ч(1)), InStr(GetInsidePrivs(pסԺҽ������), "ҽ��ѡ������") > 0))
    opt��Ч(lngTmp).Value = True
    '������һ���ſ��ܽ���
    If InStr(GetInsidePrivs(pסԺҽ������), "������������") = 0 Then
        opt��Ч(0).Value = True
        opt��Ч(1).Enabled = False
    ElseIf InStr(GetInsidePrivs(pסԺҽ������), "������������") = 0 Then
        opt��Ч(1).Value = True
        opt��Ч(0).Enabled = False
    End If
    
    'ȱʡ����ʱ��
    cboTime.AddItem "1-����"
    cboTime.AddItem "2-����"
    cboTime.AddItem "3-����"
    cboTime.AddItem "4-ָ��"
    strTmp = zlDatabase.GetPara("�������ͽ���ʱ��", glngSys, pסԺҽ������, "0", Array(lbl����ʱ��, cboTime), InStr(GetInsidePrivs(pסԺҽ������), "ҽ��ѡ������") > 0)
    cboTime.ListIndex = Val(strTmp)
    If cboTime.ListIndex = cboTime.ListCount - 1 Then
        curDate = zlDatabase.Currentdate
        strTmp = zlDatabase.GetPara("�������ͽ���ʱ��", glngSys, pסԺҽ������, "23:59:59", Array(dtpEnd), InStr(GetInsidePrivs(pסԺҽ������), "ҽ��ѡ������") > 0)
        lngTmp = Val(zlDatabase.GetPara("��������ʱ����", glngSys, pסԺҽ������, "0", Array(dtpEnd), InStr(GetInsidePrivs(pסԺҽ������), "ҽ��ѡ������") > 0))
        dtpEnd.Value = Format(curDate + lngTmp, "yyyy-MM-dd " & strTmp)
    End If
        
    'ȱʡִ�п���
    txtִ�п���.Text = "���п���"
    txtִ�п���.Tag = txtִ�п���.Text
    cmdִ�п���.Tag = ""
        
    'Ӥ��ҽ��
    If mintӤ�� <> -1 Then
        chkBaby(0).Value = IIF(mintӤ�� = 0, 1, 0)
        chkBaby(1).Value = IIF(mintӤ�� > 0, 1, 0)
    End If
        
    '����/����
    'Call zlControl.LvwFlatColumnHeader(lvwPati)
    Call InitUnits
                    
    '�������
    Call Load�������
End Sub

Private Function Load�������() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str���s As String
    
    On Error GoTo errH
    
    str���s = zlDatabase.GetPara("���������������", glngSys, pסԺҽ������, "", Array(lbl�������, lstClass), InStr(GetInsidePrivs(pסԺҽ������), "ҽ��ѡ������") > 0)
    
    strSQL = "Select ����,���� From ������Ŀ��� Where ���� Not IN('5','6','7','8','9') Order by ����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        lstClass.AddItem rsTmp!����
        lstClass.ItemData(lstClass.NewIndex) = Asc(rsTmp!����)
        If str���s <> "" Then
            If InStr(str���s, "'" & rsTmp!���� & "'") > 0 Then
                lstClass.Selected(lstClass.NewIndex) = True
            End If
        Else
            lstClass.Selected(lstClass.NewIndex) = True
        End If
        rsTmp.MoveNext
    Next
    If lstClass.ListCount > 0 Then lstClass.ListIndex = 0
    Load������� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

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
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
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

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long, strTmp As String
    
    '������������
    If mblnOK Then
        Call zlDatabase.SetPara("�������ͽ���ʱ��", cboTime.ListIndex, glngSys, pסԺҽ������)
        If cboTime.ListIndex = cboTime.ListCount - 1 Then
            Call zlDatabase.SetPara("�������ͽ���ʱ��", Format(dtpEnd.Value, "HH:mm:ss"), glngSys, pסԺҽ������)
            Call zlDatabase.SetPara("��������ʱ����", Int(CDate(Format(dtpEnd.Value, "yyyy-MM-dd")) - CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd"))), glngSys, pסԺҽ������)
        End If
        Call zlDatabase.SetPara("��������ҽ����Ч", IIF(opt��Ч(1).Value, 1, 0), glngSys, pסԺҽ������)
        Call zlDatabase.SetPara("���������������", Replace(mstr���s, "'", "''"), glngSys, pסԺҽ������)

        '���ˣ�ѡ���˽�Ϊ��ǰ����ʱ,������
        If UBound(Split(mstr����IDs, ",")) = 0 And Val(mstr����IDs) = mlng����ID Then
            Call zlDatabase.SetPara("�������Ͳ���", "", glngSys, pסԺҽ������)
        Else
            Call zlDatabase.SetPara("�������Ͳ���", cboUnit.ItemData(cboUnit.ListIndex) & ":" & mstr����IDs, glngSys, pסԺҽ������)
        End If
    End If
    
    '�ͷ�˽�м�IN����
    mMainPrivs = ""
    mlng����ID = 0
    Set mrsWarn = Nothing
    
    Call SaveListViewState(Me.lvwPati, App.ProductName & Me.Name, "")
End Sub

Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwPati, ColumnHeader.Index)
End Sub

Private Sub opt��Ч_Click(Index As Integer)
    cboTime.Enabled = opt��Ч(0).Value
    dtpEnd.Enabled = cboTime.Enabled And cboTime.ListIndex = cboTime.ListCount - 1
End Sub

Private Sub txtִ�п���_GotFocus()
    Call zlControl.TxtSelAll(txtִ�п���)
End Sub

Private Sub txtִ�п���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then Call cmdִ�п���_Click
End Sub

Private Sub txtִ�п���_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vRect As RECT
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtִ�п���.Text = txtִ�п���.Tag Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txtִ�п���.Text = "" Then
            txtִ�п���.Text = "���п���"
            txtִ�п���.Tag = txtִ�п���.Text
            cmdִ�п���.Tag = ""
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            strSQL = _
                " Select 0 as ID,'-' as ����,'���п���' as ����,NULL as ���� From Dual" & _
                " Union ALL" & _
                " Select Distinct A.ID,A.����,A.����,A.����" & _
                " From ���ű� A,��������˵�� B" & _
                " Where A.ID=B.����ID And B.������� IN(2,3)" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)"
            strSQL = "Select * From (" & strSQL & ")" & _
                " Where ���� Like [1] Or Upper(����) Like [2] Or Upper(����) Like [2]" & _
                " Order by ����"
            vRect = GetControlRect(txtִ�п���.Hwnd)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ִ�п���", False, "", "", False, False, True, _
                vRect.Left, vRect.Top, txtִ�п���.Height, blnCancel, False, True, _
                UCase(txtִ�п���.Text) & "%", mstrLike & UCase(txtִ�п���.Text) & "%")
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "δ�ҵ�ƥ��Ŀ��ҡ�", vbInformation, gstrSysName
                End If
                txtִ�п���.Text = txtִ�п���.Tag
                Call zlControl.TxtSelAll(txtִ�п���)
                txtִ�п���.SetFocus
            Else
                txtִ�п���.Text = rsTmp!����
                txtִ�п���.Tag = rsTmp!����
                cmdִ�п���.Tag = rsTmp!ID
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End If
End Sub

Private Sub txtִ�п���_Validate(Cancel As Boolean)
    If txtִ�п���.Text = "" Then
        txtִ�п���.Text = "���п���"
        txtִ�п���.Tag = txtִ�п���.Text
        cmdִ�п���.Tag = ""
    ElseIf txtִ�п���.Text <> txtִ�п���.Tag Then
        txtִ�п���.Text = txtִ�п���.Tag '�ָ���Ϊ�����
    End If
End Sub

Private Sub tbrAutoSel_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim i As Long, blnDo As Boolean
    
    If mrsWarn Is Nothing Then Exit Sub
    
    With lvwPati
        For i = 1 To .ListItems.Count
            If .ListItems(i).Checked Then
                'ֻ�����ۼƱ����������д���
                mrsWarn.Filter = "��������=1 And ���ò���='" & .ListItems(i).ListSubItems(1).Tag & "'"
                If Not mrsWarn.EOF Then
                    blnDo = False
                    Select Case BeSureMode(Nvl(mrsWarn!������־1), Nvl(mrsWarn!������־2), Nvl(mrsWarn!������־3))
                    Case 1 '���ڱ���ֵ(����Ԥ����ľ�)��ʾѯ�ʼ���
                        blnDo = Val(.ListItems(i).SubItems(3)) + Val(.ListItems(i).ListSubItems(2).Tag) <= 0
                    Case 2 '���ڱ���ֵ��ʾѯ�ʼ���,Ԥ����ľ�ʱ��ֹ����
                        blnDo = Val(.ListItems(i).SubItems(3)) + Val(.ListItems(i).ListSubItems(2).Tag) <= 0
                    Case 3 '���ڱ���ֵ��ֹ����
                        blnDo = Val(.ListItems(i).SubItems(3)) + Val(.ListItems(i).ListSubItems(2).Tag) < Nvl(mrsWarn!����ֵ, 0)
                    End Select
                    If blnDo Then
                        .ListItems(i).Checked = False
                    End If
                End If
            End If
        Next
    End With
End Sub
