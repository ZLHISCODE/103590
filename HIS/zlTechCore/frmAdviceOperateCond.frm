VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdviceOperateCond 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "У������"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   Icon            =   "frmAdviceOperateCond.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraDetail 
      Height          =   5040
      Index           =   0
      Left            =   135
      TabIndex        =   14
      Top             =   60
      Width           =   5460
      Begin VB.CheckBox chkPauseLast 
         Caption         =   "Ĭ�ϴ�ҽ�����ϴ�ִ��ʱ��֮��ʼ��ͣ(&F)"
         Height          =   195
         Left            =   1215
         TabIndex        =   8
         Top             =   4260
         Width           =   3825
      End
      Begin MSComctlLib.Toolbar tbrAutoSel 
         Height          =   360
         Left            =   1215
         TabIndex        =   11
         Top             =   4575
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
               Caption         =   "��������������ѡ��Ƿ�Ѳ���   "
               Object.ToolTipText     =   "Ctrl + Q"
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.CheckBox chk��Ч 
         Caption         =   "����(&T)"
         Height          =   195
         Index           =   1
         Left            =   2145
         TabIndex        =   1
         Top             =   330
         Width           =   930
      End
      Begin VB.CheckBox chk��Ч 
         Caption         =   "����(&L)"
         Height          =   195
         Index           =   0
         Left            =   1215
         TabIndex        =   0
         Top             =   330
         Width           =   930
      End
      Begin VB.CommandButton cmdAllPati 
         Caption         =   "ȫѡ"
         Height          =   330
         Left            =   210
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + A"
         Top             =   3450
         Width           =   870
      End
      Begin VB.CommandButton cmdNoPati 
         Caption         =   "ȫ��"
         Height          =   330
         Left            =   210
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + R"
         Top             =   3825
         Width           =   870
      End
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   615
         Width           =   4095
      End
      Begin MSComctlLib.ListView lvwPati 
         Height          =   3210
         Left            =   1215
         TabIndex        =   7
         Top             =   975
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   5662
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
      End
      Begin VB.CheckBox chk��� 
         Caption         =   "����(&H)"
         Height          =   195
         Index           =   1
         Left            =   4425
         TabIndex        =   3
         Top             =   330
         Width           =   930
      End
      Begin VB.CheckBox chk��� 
         Caption         =   "ҩ��(&D)"
         Height          =   195
         Index           =   0
         Left            =   3495
         TabIndex        =   2
         Top             =   330
         Width           =   930
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����(&U)"
         Height          =   180
         Left            =   150
         TabIndex        =   4
         Top             =   675
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����(&P)"
         Height          =   180
         Left            =   150
         TabIndex        =   6
         Top             =   1050
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3990
      TabIndex        =   13
      Top             =   5235
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2760
      TabIndex        =   12
      Top             =   5235
      Width           =   1100
   End
End
Attribute VB_Name = "frmAdviceOperateCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mstrPrivs As String 'IN
Public mint���� As Integer 'IN:3-ҽ��У��,5-��ͣҽ��,6-����ҽ��
Public mlng����ID As Long 'IN/OUT
Public mlng����ID As Long 'IN
Public mstr����IDs As String 'OUT:����ID��(����ID,��ҳID;...)
Public mint��Ч As Integer 'OUT:0-����,1-����,2-����
Public mint��� As Integer 'OUT:0-ҩ��,1-����,2-����
Public mblnPauseLast As Boolean 'OUT:�Ƿ���ϴ�ִ��ʱ�俪ʼ��ͣ
Public mblnOK As Boolean 'OUT:�Ƿ�ȷ��

Private mrsWarn As ADODB.Recordset

Private Sub cboUnit_Click()
'���ܣ���ȡָ����Χ�ڵĲ����б�
    Dim rsTmp As New ADODB.Recordset
    Dim rsWarn As New ADODB.Recordset
    Dim objItem As ListItem, strSQL As String
    Dim i As Integer, j As Integer, k As Integer
    Dim str����IDs As String, lng����ID As Long
        
    lvwPati.ListItems.Clear
    
    On Error GoTo errH
    
    '��ȡ������������
    If mint���� = 5 Or mint���� = 6 Then
        strSQL = "Select ���ò���,��������,����ֵ From ���ʱ����� Where ����ID=[1]"
        Set mrsWarn = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboUnit.ItemData(cboUnit.ListIndex))
    End If
    
    str����IDs = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ҽ����������" & mint����, "")
    If str����IDs <> "" And InStr(str����IDs, ":") > 0 Then
        lng����ID = Val(Split(str����IDs, ":")(0))
        str����IDs = Split(str����IDs, ":")(1)
    End If
        
    '��Ժ����:��Ժ���˽�ֹ����ҽ��
    strSQL = _
        "Select A.����ID,B.��ҳID,A.����,A.סԺ��,B.��Ժ���� as ����," & _
        " Nvl(E.Ԥ�����,0)-Nvl(E.�������,0)+Decode(B.����,Null,0,Nvl(F.���,0)) as ʣ���," & _
        " A.������,Decode(X.����,'1',1,Decode(B.����,Null,0,1)) as ҽ��,B.����," & _
        " B.סԺҽʦ,B.�ѱ�,D.���� as ����ȼ�,C.���� as ����,B.��Ժ����" & _
        " From ������Ϣ A,������ҳ B,���ű� C,�շ���ĿĿ¼ D,������� E,ҽ�Ƹ��ʽ X," & _
        " (Select ����ID,��ҳID,Sum(���) As ��� From ����ģ����� Group By ����ID,��ҳID) F" & _
        " Where A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And B.��Ժ����ID=C.ID" & _
        " And A.����ID=E.����ID(+) And E.����(+)=1 And B.����ID=F.����ID(+) And B.��ҳID=F.��ҳID(+)" & _
        " And B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3 And B.����ȼ�ID=D.ID(+) And B.ҽ�Ƹ��ʽ=X.����(+)" & _
        IIF(cboUnit.ItemData(cboUnit.ListIndex) > 0, " And B.��ǰ����ID=[1]", "") & _
        IIF(cboUnit.ItemData(cboUnit.ListIndex) = 0, " Order by A.סԺ�� Desc", " Order by LPAD(B.��Ժ����,10,' ')")
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
        objItem.Tag = rsTmp!��ҳID
                
        '������Ϣ
        objItem.ListSubItems(1).Tag = Nvl(rsTmp!ҽ��, 0)
        objItem.ListSubItems(2).Tag = Nvl(rsTmp!������, 0)
                
        '���ղ����ú�ɫ��ʾ
        If Not IsNull(rsTmp!����) Then
            objItem.ForeColor = vbRed
            For j = 1 To objItem.ListSubItems.Count
                objItem.ListSubItems(j).ForeColor = vbRed
            Next
        End If
        
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

Private Sub chk���_Click(Index As Integer)
    If chk���(0).Value = 0 And chk���(1).Value = 0 Then chk���(Index).Value = 1
End Sub

Private Sub chk��Ч_Click(Index As Integer)
    If chk��Ч(0).Value = 0 And chk��Ч(1).Value = 0 Then chk��Ч(Index).Value = 1
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
    Dim strTmp As String, i As Long
    
    If cboUnit.ListIndex = -1 Then
        MsgBox "��ѡ��һ��������", vbInformation, gstrSysName
        cboUnit.SetFocus: Exit Sub
    End If
    mlng����ID = cboUnit.ItemData(cboUnit.ListIndex)
    
    'סԺ����
    mstr����IDs = ""
    For i = 1 To lvwPati.ListItems.Count
        If lvwPati.ListItems(i).Checked Then
            strTmp = strTmp & "," & Mid(lvwPati.ListItems(i).Key, 2) '���ڱ���
            mstr����IDs = mstr����IDs & ";" & Mid(lvwPati.ListItems(i).Key, 2) & "," & lvwPati.ListItems(i).Tag
        End If
    Next
    strTmp = Mid(strTmp, 2)
    mstr����IDs = Mid(mstr����IDs, 2)
    If mstr����IDs = "" Then
        MsgBox "������ѡ��һ�����ˡ�", vbInformation, gstrSysName
        lvwPati.SetFocus: Exit Sub
    End If
        
    'ҽ����Ч
    mint��Ч = IIF(chk��Ч(0).Value = 1 And chk��Ч(1).Value = 1, 0, IIF(chk��Ч(0).Value = 1, 1, 2))
        
    'ҽ�����
    mint��� = IIF(chk���(0).Value = 1 And chk���(1).Value = 1, 0, IIF(chk���(0).Value = 1, 1, 2))
    
    'Ĭ����ͣʱ��
    mblnPauseLast = chkPauseLast.Value = 1
    
    '������������
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ҽ��������Ч" & mint����, IIF(chk��Ч(0).Value = 1 And chk��Ч(1).Value = 1, 0, IIF(chk��Ч(0).Value = 1, 1, 2))
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ҽ���������" & mint����, IIF(chk���(0).Value = 1 And chk���(1).Value = 1, 0, IIF(chk���(0).Value = 1, 1, 2))
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�ϴο�ʼ��ͣ", chkPauseLast.Value
    If UBound(Split(strTmp, ",")) = 0 And Val(strTmp) = mlng����ID Then
        '���ˣ�ѡ���˽�Ϊ��ǰ����ʱ,������
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ҽ����������" & mint����, ""
    Else
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ҽ����������" & mint����, cboUnit.ItemData(cboUnit.ListIndex) & ":" & strTmp
    End If
    
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        Call cmdAllPati_Click
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call cmdNoPati_Click
    ElseIf KeyCode = vbKeyQ And Shift = vbCtrlMask Then
        If tbrAutoSel.Visible Then
            Call tbrAutoSel_ButtonClick(tbrAutoSel.Buttons(1))
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim lngTmp As Long
    
    Call RestoreListViewState(Me.lvwPati, App.ProductName & Me.Name, "")
    
    mblnOK = False
    Me.Caption = Decode(mint����, 3, "У��", 5, "��ͣ", 6, "����") & "����"
    If mint���� <> 5 Then
        chkPauseLast.Visible = False
        
        If mint���� = 6 Then
            tbrAutoSel.Buttons(1).Caption = "���������������ſ�Ƿ�Ѳ���   "
            lvwPati.Height = chkPauseLast.Top + chkPauseLast.Height - lvwPati.Top
            cmdAllPati.Top = cmdAllPati.Top + chkPauseLast.Height
            cmdNoPati.Top = cmdNoPati.Top + chkPauseLast.Height
        Else
            tbrAutoSel.Visible = False
            lvwPati.Height = tbrAutoSel.Top + tbrAutoSel.Height - lvwPati.Top
            cmdAllPati.Top = cmdAllPati.Top + tbrAutoSel.Height + chkPauseLast.Height
            cmdNoPati.Top = cmdNoPati.Top + tbrAutoSel.Height + chkPauseLast.Height
        End If
    End If
    
    'ȱʡҽ����Ч
    If mint���� = 5 Or mint���� = 6 Then
        chk��Ч(0).Enabled = False: chk��Ч(1).Enabled = False
        chk��Ч(0).Value = 1: chk��Ч(1).Value = 0
    Else
        chk��Ч(0).Enabled = True: chk��Ч(1).Enabled = True
        lngTmp = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ҽ��������Ч" & mint����, 0))
        If lngTmp = 0 Then
            chk��Ч(0).Value = 1: chk��Ч(1).Value = 1
        Else
            chk��Ч(lngTmp - 1).Value = 1
        End If
    End If
    'ȱʡҽ�����
    lngTmp = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ҽ���������" & mint����, 0))
    If lngTmp = 0 Then
        chk���(0).Value = 1: chk���(1).Value = 1
    Else
        chk���(lngTmp - 1).Value = 1
    End If
    
    'Ĭ����ͣʱ��
    chkPauseLast.Value = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�ϴο�ʼ��ͣ", 0))
    
    '����/����
    Call zlControl.LvwFlatColumnHeader(lvwPati)
    Call InitUnits
End Sub

Private Function InitUnits() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, strSQL As String
    
    On Error GoTo errH
    
    '��������۲���
    If InStr(mstrPrivs, "ȫԺ����") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where A.ID=B.����ID And B.������� in(1,2,3) And B.��������='����'" & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Order by A.����"
    Else
        '����Ȩ������ֱ�����ڲ���+���ڿ�����������
        strSQL = _
            " Select A.ID,A.����,A.����,Nvl(C.ȱʡ,0) as ȱʡ" & _
            " From ���ű� A,��������˵�� B,������Ա C" & _
            " Where A.ID=B.����ID And A.ID=C.����ID And C.��ԱID=[1]" & _
            " And B.������� in(1,2,3) And B.��������='����'" & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))"
        If Not gbln�������Ҷ��� Then
            strSQL = strSQL & IIF(strSQL <> "", " Union ", "") & _
                " Select C.ID,C.����,C.����,Nvl(B.ȱʡ,0) as ȱʡ" & _
                " From ��λ״����¼ A,������Ա B,���ű� C" & _
                " Where A.����ID=C.ID And B.����ID=A.����ID And B.��ԱID=[1]" & _
                " And (C.����ʱ�� is NULL or Trunc(C.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))"
        End If
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
    '�ͷ�˽�м�IN����
    mstrPrivs = ""
    mint���� = 0
    mlng����ID = 0
    Set mrsWarn = Nothing
    
    Call SaveListViewState(Me.lvwPati, App.ProductName & Me.Name, "")
End Sub

Private Sub tbrAutoSel_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim i As Long, k As Long
    
    If mrsWarn Is Nothing Then Exit Sub
    
    With lvwPati
        If mint���� = 5 Then
            k = 0
            For i = 1 To .ListItems.Count
                .ListItems(i).Checked = False
                'ֻ�����ۼƱ����������д���
                mrsWarn.Filter = "��������=1 And ���ò���=" & Val(.ListItems(i).ListSubItems(1).Tag) + 1
                If Not mrsWarn.EOF Then
                    If Val(.ListItems(i).SubItems(3)) + Val(.ListItems(i).ListSubItems(2).Tag) < Nvl(mrsWarn!����ֵ, 0) Then
                        .ListItems(i).Checked = True
                        If k = 0 Then
                            .ListItems(i).Selected = True
                            .ListItems(i).EnsureVisible
                        End If
                        k = k + 1
                    End If
                End If
            Next
        ElseIf mint���� = 6 Then
            For i = 1 To .ListItems.Count
                If .ListItems(i).Checked Then
                    'ֻ�����ۼƱ����������д���
                    mrsWarn.Filter = "��������=1 And ���ò���=" & Val(.ListItems(i).ListSubItems(1).Tag) + 1
                    If Not mrsWarn.EOF Then
                        If Val(.ListItems(i).SubItems(3)) + Val(.ListItems(i).ListSubItems(2).Tag) < Nvl(mrsWarn!����ֵ, 0) Then
                            .ListItems(i).Checked = False
                        End If
                    End If
                End If
            Next
        End If
    End With
End Sub
