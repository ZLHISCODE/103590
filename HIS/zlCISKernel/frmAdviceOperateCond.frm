VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdviceOperateCond 
   AutoRedraw      =   -1  'True
   Caption         =   "У������"
   ClientHeight    =   6555
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   6180
   Icon            =   "frmAdviceOperateCond.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
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
      Top             =   6060
      Width           =   6180
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   3720
         TabIndex        =   11
         Top             =   0
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   4950
         TabIndex        =   10
         Top             =   0
         Width           =   1100
      End
   End
   Begin VB.Frame fraDetail 
      Height          =   5880
      Left            =   135
      TabIndex        =   8
      Top             =   60
      Width           =   5940
      Begin VB.Frame fraAdviceKind 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1200
         TabIndex        =   12
         Top             =   240
         Width           =   4575
         Begin VB.CheckBox chk��� 
            Caption         =   "ҩ��(&D)"
            Height          =   195
            Index           =   0
            Left            =   2280
            TabIndex        =   16
            Top             =   0
            Width           =   930
         End
         Begin VB.CheckBox chk��� 
            Caption         =   "����(&H)"
            Height          =   195
            Index           =   1
            Left            =   3210
            TabIndex        =   15
            Top             =   0
            Width           =   930
         End
         Begin VB.CheckBox chk��Ч 
            Caption         =   "����(&L)"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   930
         End
         Begin VB.CheckBox chk��Ч 
            Caption         =   "����(&T)"
            Height          =   195
            Index           =   1
            Left            =   930
            TabIndex        =   13
            Top             =   0
            Width           =   930
         End
      End
      Begin VB.CheckBox chkPauseLast 
         Caption         =   "Ĭ�ϴ�ҽ�����ϴ�ִ��ʱ��֮��ʼ��ͣ(&F)"
         Height          =   195
         Left            =   1215
         TabIndex        =   4
         Top             =   5100
         Width           =   3825
      End
      Begin MSComctlLib.Toolbar tbrAutoSel 
         Height          =   375
         Left            =   1215
         TabIndex        =   7
         Top             =   5415
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   661
         ButtonWidth     =   5159
         ButtonHeight    =   609
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
      Begin VB.CommandButton cmdAllPati 
         Caption         =   "ȫѡ"
         Height          =   330
         Left            =   210
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + A"
         Top             =   4290
         Width           =   870
      End
      Begin VB.CommandButton cmdNoPati 
         Caption         =   "ȫ��"
         Height          =   330
         Left            =   210
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + R"
         Top             =   4665
         Width           =   870
      End
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   615
         Width           =   2520
      End
      Begin MSComctlLib.ListView lvwPati 
         Height          =   4050
         Left            =   1215
         TabIndex        =   3
         Top             =   975
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   7144
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
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����(&U)"
         Height          =   180
         Left            =   150
         TabIndex        =   0
         Top             =   675
         Width           =   990
      End
      Begin VB.Label lblPati 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����(&P)"
         Height          =   180
         Left            =   150
         TabIndex        =   2
         Top             =   1050
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmAdviceOperateCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mMainPrivs As String 'IN
Public mint���� As Integer 'IN:2-ȷ��ֹͣ,3-ҽ��У��,5-��ͣҽ��,6-����ҽ��
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
    Dim i As Integer, j As Integer, k As Integer, lngUnitID As Long
    Dim str����IDs As String, lng����ID As Long
    Dim lngColor As Long
    
    On Error GoTo errH
    lvwPati.ListItems.Clear
    lngUnitID = cboUnit.ItemData(cboUnit.ListIndex)
    
    '��ȡ������������
    If mint���� = 5 Or mint���� = 6 Then
        strSQL = "Select ���ò���,��������,����ֵ,������־1,������־2,������־3 From ���ʱ����� Where ����ID=[1]"
        Set mrsWarn = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngUnitID)
    End If
    
    str����IDs = zlDatabase.GetPara("���Ͳ���", glngSys, pסԺҽ������)
    If str����IDs <> "" And InStr(str����IDs, ":") > 0 Then
        lng����ID = Val(Split(str����IDs, ":")(0))
        str����IDs = Split(str����IDs, ":")(1)
    End If
            
    Set rsTmp = GetPatiRsByUnit(lngUnitID, mlng����ID, (mint���� = 5 Or mint���� = 6), True, False)
    For i = 1 To rsTmp.RecordCount
        If Val(rsTmp!��˱�־ & "") < 1 Or gbyt������˷�ʽ <> 1 Then
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
            objItem.Tag = rsTmp!��ҳID
                    
            '������Ϣ
            objItem.ListSubItems(1).Tag = Nvl(rsTmp!���ò���)
            objItem.ListSubItems(2).Tag = Nvl(rsTmp!������, 0)
                    
            '������ɫ
            lngColor = zlDatabase.GetPatiColor(Nvl(rsTmp!��������))
            objItem.ListSubItems(1).ForeColor = lngColor
            objItem.ListSubItems(9).ForeColor = lngColor
            
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

Private Sub chk���_Click(Index As Integer)
    If chk���(0).value = 0 And chk���(1).value = 0 Then chk���(Index).value = 1
End Sub

Private Sub chk��Ч_Click(Index As Integer)
    If chk��Ч(0).value = 0 And chk��Ч(1).value = 0 Then chk��Ч(Index).value = 1
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
    Dim strSel As String, strUnSel As String, i As Long
    
    If cboUnit.ListIndex = -1 Then
        MsgBox "��ѡ��һ��������", vbInformation, gstrSysName
        cboUnit.SetFocus: Exit Sub
    End If
    mlng����ID = cboUnit.ItemData(cboUnit.ListIndex)
    
    'סԺ����
    mstr����IDs = ""
    For i = 1 To lvwPati.ListItems.Count
        If lvwPati.ListItems(i).Checked Then
            strSel = strSel & "," & Mid(lvwPati.ListItems(i).Key, 2) '���ڱ���
            mstr����IDs = mstr����IDs & ";" & Mid(lvwPati.ListItems(i).Key, 2) & "," & lvwPati.ListItems(i).Tag
        Else
            strUnSel = strUnSel & "," & Mid(lvwPati.ListItems(i).Key, 2) '���ڱ���
        End If
    Next
    mstr����IDs = Mid(mstr����IDs, 2)
    If mstr����IDs = "" Then
        MsgBox "������ѡ��һ�����ˡ�", vbInformation, gstrSysName
        lvwPati.SetFocus: Exit Sub
    End If
        
    'ҽ����Ч
    mint��Ч = IIF(chk��Ч(0).value = 1 And chk��Ч(1).value = 1, 0, IIF(chk��Ч(0).value = 1, 1, 2))
        
    'ҽ�����
    mint��� = IIF(chk���(0).value = 1 And chk���(1).value = 1, 0, IIF(chk���(0).value = 1, 1, 2))
    
    'Ĭ����ͣʱ��
    mblnPauseLast = chkPauseLast.value = 1
    
    '������������
    If chk��Ч(0).Visible Then
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ҽ��������Ч" & mint����, _
            IIF(chk��Ч(0).value = 1 And chk��Ч(1).value = 1, 0, IIF(chk��Ч(0).value = 1, 1, 2))
    End If
    If chk���(0).Visible Then
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ҽ���������" & mint����, _
            IIF(chk���(0).value = 1 And chk���(1).value = 1, 0, IIF(chk���(0).value = 1, 1, 2))
    End If
    
    If chkPauseLast.Visible Then
        Call zlDatabase.SetPara("�ϴο�ʼ��ͣ", chkPauseLast.value, glngSys, pסԺҽ������)
    End If
        
    '����
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
    ElseIf KeyCode = vbKeyQ And Shift = vbCtrlMask Then
        If tbrAutoSel.Visible Then
            Call tbrAutoSel_ButtonClick(tbrAutoSel.Buttons(1))
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call ZLCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim lngTmp As Long
    
    Call RestoreWinState(Me, App.ProductName)
    Call RestoreListViewState(Me.lvwPati, App.ProductName & Me.Name, "")
    
    mblnOK = False
    Me.Caption = Decode(mint����, 2, "ȷ��ֹͣ", 3, "У��", 5, "��ͣ", 6, "����") & "����"
    If mint���� <> 5 Then
        chkPauseLast.Visible = False
        
        If mint���� = 6 Then
            tbrAutoSel.Buttons(1).Caption = "���������������ſ�Ƿ�Ѳ���   "
            lvwPati.Height = chkPauseLast.Top + chkPauseLast.Height - lvwPati.Top
            cmdAllPati.Top = cmdAllPati.Top + chkPauseLast.Height
            cmdNoPati.Top = cmdNoPati.Top + chkPauseLast.Height
        Else
            tbrAutoSel.Visible = False
            
            lngTmp = 0
            If mint���� = 2 Then
                fraAdviceKind.Visible = False
                lngTmp = fraAdviceKind.Height + 60
                cboUnit.Top = cboUnit.Top - lngTmp
                lblUnit.Top = lblUnit.Top - lngTmp
                lblPati.Top = lblPati.Top - lngTmp
                lvwPati.Top = lvwPati.Top - lngTmp
            End If
            
            lvwPati.Height = tbrAutoSel.Top + tbrAutoSel.Height - lvwPati.Top
            cmdAllPati.Top = cmdAllPati.Top + tbrAutoSel.Height + chkPauseLast.Height
            cmdNoPati.Top = cmdNoPati.Top + tbrAutoSel.Height + chkPauseLast.Height
        End If
    End If
    
    If mint���� = 2 Then
        chk��Ч(0).value = 1: chk��Ч(1).value = 1
        chk���(0).value = 1: chk���(1).value = 1
    Else
        'ȱʡҽ����Ч
        If mint���� = 5 Or mint���� = 6 Then
            chk��Ч(0).Enabled = False: chk��Ч(1).Enabled = False
            chk��Ч(0).value = 1: chk��Ч(1).value = 0
        Else
            chk��Ч(0).Enabled = True: chk��Ч(1).Enabled = True
            lngTmp = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ҽ��������Ч" & mint����, 0))
            If lngTmp = 0 Then
                chk��Ч(0).value = 1: chk��Ч(1).value = 1
            Else
                chk��Ч(lngTmp - 1).value = 1
            End If
        End If
        'ȱʡҽ�����
        lngTmp = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ҽ���������" & mint����, 0))
        If lngTmp = 0 Then
            chk���(0).value = 1: chk���(1).value = 1
        Else
            chk���(lngTmp - 1).value = 1
        End If
        
        'Ĭ����ͣʱ��
        chkPauseLast.value = Val(zlDatabase.GetPara("�ϴο�ʼ��ͣ", glngSys, pסԺҽ������, "0", Array(chkPauseLast)))
    End If
    
    '����/����
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
    lvwPati.Width = fraDetail.Width - lvwPati.Left - 120
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 120
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 60
    
    fraDetail.Height = Me.ScaleHeight - picBottom.Height - 120
    
    tbrAutoSel.Top = fraDetail.Height - tbrAutoSel.Height - 60
    chkPauseLast.Top = tbrAutoSel.Top - chkPauseLast.Height - 60
    lvwPati.Height = chkPauseLast.Top - lvwPati.Top - 60
    
    If tbrAutoSel.Visible = False Then lvwPati.Height = lvwPati.Height + tbrAutoSel.Height + 60
    If chkPauseLast.Visible = False Then lvwPati.Height = lvwPati.Height + chkPauseLast.Height
    
    cmdNoPati.Top = lvwPati.Top + lvwPati.Height - 30 - cmdNoPati.Height
    cmdAllPati.Top = cmdNoPati.Top - cmdAllPati.Height - 30
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '�ͷ�˽�м�IN����
    mMainPrivs = ""
    mint���� = 0
    mlng����ID = 0
    Set mrsWarn = Nothing
    
    Call SaveListViewState(Me.lvwPati, App.ProductName & Me.Name, "")
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwPati, ColumnHeader.Index)
End Sub

Private Sub tbrAutoSel_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim i As Long, k As Long
    Dim blnDo As Boolean
    
    If mrsWarn Is Nothing Then Exit Sub
    
    With lvwPati
        If mint���� = 5 Then
            k = 0
            For i = 1 To .ListItems.Count
                .ListItems(i).Checked = False
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
        End If
    End With
End Sub
