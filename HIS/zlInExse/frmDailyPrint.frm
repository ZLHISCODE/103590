VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDailyPrint 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "һ���嵥��ӡ"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdPrintAll 
      Caption         =   "һ�δ�ӡ���в���(&M)"
      Height          =   350
      Left            =   4110
      TabIndex        =   18
      Top             =   4635
      Width           =   1905
   End
   Begin VB.CommandButton cmdPreviewAll 
      Caption         =   "һ��Ԥ�����в���(&A)"
      Height          =   350
      Left            =   4110
      TabIndex        =   16
      Top             =   4215
      Width           =   1905
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "��ӡ����(&S)"
      Height          =   350
      Left            =   120
      TabIndex        =   21
      Top             =   4635
      Width           =   1275
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "�����˷ֿ���ӡ(&P)"
      Height          =   350
      Left            =   2220
      TabIndex        =   17
      Top             =   4635
      Width           =   1680
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&X)"
      Height          =   350
      Left            =   6195
      TabIndex        =   19
      Top             =   4635
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   120
      TabIndex        =   20
      Top             =   4215
      Width           =   1275
   End
   Begin VB.Frame fraDetail 
      Height          =   4095
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Width           =   7305
      Begin VB.CheckBox chkReCharge 
         Caption         =   "�����˷�(&R)"
         Height          =   195
         Left            =   5640
         TabIndex        =   5
         Top             =   308
         Value           =   1  'Checked
         Width           =   1395
      End
      Begin VB.CheckBox chkZeroFee 
         Caption         =   "���������(&Z)"
         Height          =   195
         Left            =   5640
         TabIndex        =   6
         Top             =   653
         Value           =   1  'Checked
         Width           =   1515
      End
      Begin VB.OptionButton opttime 
         Caption         =   "����ʱ��(&H)"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   3600
         TabIndex        =   4
         Top             =   660
         Width           =   1620
      End
      Begin VB.OptionButton opttime 
         Caption         =   "�Ǽ�ʱ��(&D)"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   3600
         TabIndex        =   3
         Top             =   315
         Value           =   -1  'True
         Width           =   1500
      End
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1230
         Width           =   2070
      End
      Begin VB.CommandButton cmdNoPati 
         Caption         =   "ȫ��(&R)"
         Height          =   330
         Left            =   270
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + R"
         Top             =   3555
         Width           =   870
      End
      Begin VB.CommandButton cmdAllPati 
         Caption         =   "ȫѡ(&A)"
         Height          =   330
         Left            =   270
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + A"
         Top             =   3180
         Width           =   870
      End
      Begin VB.CheckBox chkPatiType 
         Caption         =   "ҽ������(&M)"
         Height          =   195
         Index           =   0
         Left            =   3720
         TabIndex        =   9
         Top             =   1290
         Value           =   1  'Checked
         Width           =   1395
      End
      Begin VB.CheckBox chkPatiType 
         Caption         =   "��ҽ������(&N)"
         Height          =   195
         Index           =   1
         Left            =   5640
         TabIndex        =   10
         Top             =   1290
         Value           =   1  'Checked
         Width           =   1515
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1215
         TabIndex        =   1
         Top             =   255
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   84279299
         CurrentDate     =   37953
      End
      Begin MSComctlLib.ListView lvwPati 
         Height          =   2340
         Left            =   1215
         TabIndex        =   12
         Top             =   1590
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   4128
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   1940
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
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   1215
         TabIndex        =   2
         Top             =   600
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   84279299
         CurrentDate     =   37953
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����(&I)"
         Height          =   180
         Left            =   180
         TabIndex        =   11
         Top             =   1665
         Width           =   990
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��(&T)"
         Height          =   180
         Left            =   180
         TabIndex        =   0
         Top             =   315
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����(&U)"
         Height          =   180
         Left            =   180
         TabIndex        =   7
         Top             =   1290
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   120
         X2              =   7200
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   105
         X2              =   7200
         Y1              =   1095
         Y2              =   1095
      End
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "�����˷ֿ�Ԥ��(&V)"
      Height          =   350
      Left            =   2220
      TabIndex        =   15
      Top             =   4215
      Width           =   1680
   End
End
Attribute VB_Name = "frmDailyPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mlng����ID As Long 'IN
Public mlng����ID As Long 'IN
Public mstrPrivs As String

Private mlngModul As Long

Private Sub cboUnit_Click()
'���ܣ���ȡָ����Χ�ڵĲ����б�
    Dim rsTmp As New ADODB.Recordset
    Dim objItem As ListItem, strSql As String
    Dim i As Integer, j As Integer
    Dim intBedLen As Integer, str����IDs As String, lng����ID As Long
        
    lvwPati.ListItems.Clear
    
    On Error GoTo errH
    If cboUnit.ListIndex <> -1 Then
        lng����ID = cboUnit.ItemData(cboUnit.ListIndex)
        intBedLen = GetMaxBedLen(lng����ID, False)
    End If
    strSql = _
        "Select A.����ID,B.��ҳID,Nvl(b.����, a.����) As ����,B.סԺ��,LPAD(B.��Ժ����," & intBedLen & ",' ') as ����," & _
        "       B.סԺҽʦ,B.�ѱ�,D.���� as ����ȼ�,C.���� as ����,B.��Ժ����,B.����,B.��������" & _
        " From ������Ϣ A,������ҳ B,���ű� C,�շ���ĿĿ¼ D,��Ժ���ˡ�E" & _
        " Where A.����ID=B.����ID And B.��ҳID=A.��ҳID And B.��Ժ����ID=C.ID" & _
        " And A.����ID=E.����ID   And B.����ȼ�ID=D.ID(+)" & _
        IIf(chkPatiType(0).Value = 0, " And B.���� Is Null", "") & _
        IIf(chkPatiType(1).Value = 0, " And B.���� Is Not Null", "") & _
        IIf(lng����ID > 0, " And B.��ǰ����ID=[1] And E.����ID=[1] ", "") & _
        IIf(lng����ID = 0, " Order by B.סԺ�� Desc", " Order by LPAD(����,10,' ')")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID)
  
    For i = 1 To rsTmp.RecordCount
        Set objItem = lvwPati.ListItems.Add(, "_" & rsTmp!����ID, rsTmp!����)
        objItem.SubItems(1) = IIf(IsNull(rsTmp!סԺ��), "", rsTmp!סԺ��)
        objItem.SubItems(2) = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        objItem.SubItems(3) = IIf(IsNull(rsTmp!סԺҽʦ), "", rsTmp!סԺҽʦ)
        objItem.SubItems(4) = IIf(IsNull(rsTmp!�ѱ�), "", rsTmp!�ѱ�)
        objItem.SubItems(5) = IIf(IsNull(rsTmp!����ȼ�), "", rsTmp!����ȼ�)
        objItem.SubItems(6) = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        objItem.SubItems(7) = Format(rsTmp!��Ժ����, "yyyy-MM-dd HH:mm")
        objItem.Tag = rsTmp!��ҳID
        
        objItem.ForeColor = zlDatabase.GetPatiColor(NVL(rsTmp!��������))
        For j = 1 To objItem.ListSubItems.Count
            objItem.ListSubItems(j).ForeColor = zlDatabase.GetPatiColor(NVL(rsTmp!��������))
        Next
        
        If rsTmp!����ID = mlng����ID Then
            objItem.Checked = True 'ȱʡֻѡ��ǰ����
            objItem.EnsureVisible
            objItem.Selected = True
        End If
        rsTmp.MoveNext
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function InitUnits() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, strSql As String
    
    On Error GoTo errH
    
    '��������۲���
    If InStr(mstrPrivs, ";���в���;") > 0 Then
        strSql = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where A.ID=B.����ID And B.������� in(1,2,3) And B.��������='����'" & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
            " Order by A.����"
    Else
        '����Ȩ������ֱ�����ڲ���+���ڿ�����������
        strSql = _
            " Select A.ID,A.����,A.����,Nvl(C.ȱʡ,0) as ȱʡ" & _
            " From ���ű� A,��������˵�� B,������Ա C" & _
            " Where A.ID=B.����ID And A.ID=C.����ID And C.��ԱID=[1]" & _
            " And B.������� in(1,2,3) And B.��������='����'" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSql = strSql & " Union " & _
            " Select C.ID,C.����,C.����,Nvl(B.ȱʡ,0) as ȱʡ" & _
            " From �������Ҷ�Ӧ A,������Ա B,���ű� C" & _
            " Where A.����ID=C.ID And B.����ID=A.����ID And B.��ԱID=[1]" & _
            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & vbNewLine & _
            " And (C.����ʱ�� is NULL or Trunc(C.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSql = "Select ID,����,����,Max(ȱʡ) as ȱʡ From (" & strSql & ") Group by ID,����,���� Order by ����"
    End If
    
    cboUnit.Clear
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!���� & "-" & rsTmp!����
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            If rsTmp!ID = mlng����ID Then cboUnit.ListIndex = cboUnit.NewIndex
            rsTmp.MoveNext
        Next
    End If
    If cboUnit.ListCount > 0 And cboUnit.ListIndex = -1 Then cboUnit.ListIndex = 0
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub chkPatiType_Click(Index As Integer)
    
    If chkPatiType(0).Tag = "1" Then chkPatiType(0).Tag = "": Exit Sub
    If chkPatiType(0).Value = 0 And chkPatiType(1).Value = 0 Then
        chkPatiType(0).Tag = "1"
        chkPatiType(Index).Value = 1
    Else
        Call cboUnit_Click
    End If
End Sub
Private Sub cmdPrintSet_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1141", Me)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, j As Long
    
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        Call cmdAllPati_Click
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call cmdNoPati_Click
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


Private Sub cmdNoPati_Click()
    Call SelectLVW(lvwPati, False)
    lvwPati.SetFocus
End Sub

Private Sub cmdPreview_Click()
    Call PrintOrPreview(1) 'Ԥ��
End Sub

Private Sub cmdPrint_Click()
    Call PrintOrPreview(2) '��ӡ
End Sub

Private Sub cmdPreviewAll_Click()
    Call PrintOrPreviewAll(1) 'Ԥ��
End Sub

Private Sub cmdPrintAll_Click()
    Call PrintOrPreviewAll(2) '��ӡ
End Sub

Private Sub PrintOrPreview(bytMode As Byte)
    Dim blnNOSelect As Boolean, Item As ListItem
    
    For Each Item In lvwPati.ListItems
        If Item.Checked Then
            blnNOSelect = False
            
            Item.Selected = True
            Item.EnsureVisible
            Me.Refresh
            Call PrintContent(bytMode, Split(Item.Key, "_")(1))
        End If
    Next
    If blnNOSelect Then MsgBox "û��ѡ��Ҫ��ӡ�嵥�Ĳ��ˣ�", vbInformation, gstrSysName
End Sub

Private Sub PrintOrPreviewAll(bytMode As Byte)
    Dim blnNOSelect As Boolean, Item As ListItem
    Dim str����ID As String
    blnNOSelect = True
    For Each Item In lvwPati.ListItems
        If Item.Checked Then
            blnNOSelect = False
            str����ID = str����ID & "," & Split(Item.Key, "_")(1)
        End If
    Next
    If blnNOSelect Then
        MsgBox "û��ѡ��Ҫ��ӡ�嵥�Ĳ��ˣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    If str����ID <> "" Then
        str����ID = Mid(str����ID, 2)
        ReportOpen gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1141_1", Me, "����ID=" & str����ID, _
        "��ʼʱ��=" & Format(dtpBegin.Value, "yyyy-MM-dd HH:mm:ss"), _
        "����ʱ��=" & Format(dtpEnd.Value, "yyyy-MM-dd HH:mm:ss"), _
        "��ʾ�˷�=" & chkReCharge.Value, _
        "��ʾ�����=" & chkZeroFee.Value, _
        "���˲���=" & cboUnit.ItemData(cboUnit.ListIndex), _
        "����ʱ��=" & IIf(opttime(0).Value = True, "�Ǽ�ʱ��", "����ʱ��"), bytMode
    End If
End Sub

Private Sub PrintContent(ByVal bytMode As Byte, ByVal str����ID As String)
    ReportOpen gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1141", Me, "����ID=" & str����ID, _
        "��ʼʱ��=" & Format(dtpBegin.Value, "yyyy-MM-dd HH:mm:ss"), _
        "����ʱ��=" & Format(dtpEnd.Value, "yyyy-MM-dd HH:mm:ss"), _
        "��ʾ�˷�=" & chkReCharge.Value, _
        "��ʾ�����=" & chkZeroFee.Value, _
        "���˲���=" & cboUnit.ItemData(cboUnit.ListIndex), _
        "��ҳID=0", _
        "����ʱ��=" & IIf(opttime(0).Value = True, "�Ǽ�ʱ��", "����ʱ��"), bytMode
End Sub

Private Sub cmdCancel_Click()
    Dim lngTmp As Long
    Dim blnHavePara As Boolean
    
    blnHavePara = InStr(1, mstrPrivs, ";��������;") > 0
    
    zlDatabase.SetPara "��ʼʱ��", Format(Me.dtpBegin.Value, "hh:mm:ss"), glngSys, mlngModul, blnHavePara
    zlDatabase.SetPara "����ʱ��", Format(Me.dtpEnd.Value, "hh:mm:ss"), glngSys, mlngModul, blnHavePara

    lngTmp = DateDiff("d", Me.dtpEnd.Value, zlDatabase.Currentdate)
    zlDatabase.SetPara "�������", lngTmp, glngSys, mlngModul, blnHavePara
    lngTmp = DateDiff("d", Me.dtpBegin.Value, Me.dtpEnd.Value)
    zlDatabase.SetPara "��ʼ���", lngTmp, glngSys, mlngModul, blnHavePara
            
    
    zlDatabase.SetPara "����ʱ��", IIf(opttime(1).Value, 1, 0), glngSys, mlngModul, blnHavePara
    If InStr(mstrPrivs, ";��������;") > 0 Then
        zlDatabase.SetPara "��ʾ�˷�", chkReCharge.Value, glngSys, mlngModul, blnHavePara
        zlDatabase.SetPara "��ʾ�����", chkZeroFee.Value, glngSys, mlngModul, blnHavePara
    End If
    
    zlDatabase.SetPara "��ҽ������", chkPatiType(0).Value, glngSys, mlngModul, blnHavePara
    zlDatabase.SetPara "ҽ������", chkPatiType(1).Value, glngSys, mlngModul, blnHavePara
        
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Long, lngTmp As Long, strStartTime As String, strEndTime As String
    Dim blnParSet As Boolean
    
    blnParSet = InStr(mstrPrivs, ";��������;") > 0
    mlngModul = 1141
       
    strEndTime = zlDatabase.GetPara("����ʱ��", glngSys, mlngModul, "23:59:59", Array(dtpEnd), blnParSet)
    lngTmp = Val(zlDatabase.GetPara("�������", glngSys, mlngModul, 0, Array(dtpEnd), blnParSet))
    If lngTmp > 7 Then lngTmp = 7
    Me.dtpEnd.Value = CDate(Format(zlDatabase.Currentdate() - lngTmp, "yyyy-MM-dd") & " " & strEndTime)
    
    strStartTime = zlDatabase.GetPara("��ʼʱ��", glngSys, mlngModul, "00:00:00", Array(dtpBegin), blnParSet)
    lngTmp = Val(zlDatabase.GetPara("��ʼ���", glngSys, mlngModul, 0, Array(dtpBegin), blnParSet))
    If lngTmp > 7 Then lngTmp = 7
    Me.dtpBegin.Value = CDate(Format(Me.dtpEnd.Value - lngTmp, "yyyy-MM-dd") & " " & strStartTime)
    
    
    i = IIf(IIf(zlDatabase.GetPara("����ʱ��", glngSys, mlngModul, "0", Array(opttime(0), opttime(1)), blnParSet) = "1", "����ʱ��", "�Ǽ�ʱ��") = "�Ǽ�ʱ��", 0, 1) 'ע���ֵΪ1��ʾ������ʱ��
    opttime(i).Value = True
    chkReCharge.Value = IIf(zlDatabase.GetPara("��ʾ�˷�", glngSys, mlngModul, "0", Array(chkReCharge), blnParSet) = "1", 1, 0)
    chkZeroFee.Value = IIf(zlDatabase.GetPara("��ʾ�����", glngSys, mlngModul, "0", Array(chkZeroFee), blnParSet) = "1", 1, 0)
    
    
    chkPatiType(0).Value = IIf(zlDatabase.GetPara("��ҽ������", glngSys, mlngModul, "1", Array(chkPatiType(0)), blnParSet) = "1", 1, 0)
    chkPatiType(1).Value = IIf(zlDatabase.GetPara("ҽ������", glngSys, mlngModul, "1", Array(chkPatiType(1)), blnParSet) = "1", 1, 0)

    Call InitUnits '��ȡ����/����
    Call zlControl.LvwFlatColumnHeader(lvwPati)
End Sub
