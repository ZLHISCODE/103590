VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClinicOfficeEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������������"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6780
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClinicOfficeEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fra������Ϣ 
      Caption         =   "������Ϣ"
      Height          =   2085
      Left            =   90
      TabIndex        =   14
      Top             =   90
      Width           =   5085
      Begin VB.TextBox txtEdit 
         Height          =   350
         Index           =   0
         Left            =   660
         TabIndex        =   12
         Top             =   2310
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CheckBox chk�Ƿ� 
         Caption         =   "Check1"
         Height          =   210
         Index           =   0
         Left            =   210
         TabIndex        =   13
         Top             =   2760
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txt���� 
         Height          =   350
         Left            =   660
         MaxLength       =   3
         TabIndex        =   0
         Top             =   330
         Width           =   1245
      End
      Begin VB.TextBox txtλ�� 
         Height          =   350
         Left            =   660
         MaxLength       =   40
         TabIndex        =   4
         Top             =   1650
         Width           =   4335
      End
      Begin VB.TextBox txt���� 
         Height          =   350
         Left            =   660
         MaxLength       =   20
         TabIndex        =   1
         Top             =   765
         Width           =   4335
      End
      Begin VB.TextBox txt���� 
         Height          =   350
         Left            =   660
         MaxLength       =   6
         TabIndex        =   2
         Top             =   1215
         Width           =   1245
      End
      Begin VB.ComboBox cboStationNo 
         Height          =   330
         Left            =   2790
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1225
         Width           =   2205
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   210
         Index           =   0
         Left            =   210
         TabIndex        =   21
         Top             =   2340
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label lblλ�� 
         AutoSize        =   -1  'True
         Caption         =   "λ��"
         Height          =   210
         Left            =   210
         TabIndex        =   19
         Top             =   1720
         Width           =   420
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   210
         Left            =   210
         TabIndex        =   15
         Top             =   400
         Width           =   420
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   210
         Left            =   210
         TabIndex        =   16
         Top             =   835
         Width           =   420
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   210
         Left            =   210
         TabIndex        =   17
         Top             =   1285
         Width           =   420
      End
      Begin VB.Label lblStationNo 
         AutoSize        =   -1  'True
         Caption         =   "վ��"
         Height          =   210
         Left            =   2340
         TabIndex        =   18
         Top             =   1285
         Width           =   420
      End
   End
   Begin VB.Frame fraDept 
      BorderStyle     =   0  'None
      Height          =   2565
      Left            =   90
      TabIndex        =   22
      Top             =   2250
      Width           =   5115
      Begin VB.TextBox txtSelect 
         Height          =   350
         Left            =   870
         TabIndex        =   5
         Top             =   30
         Width           =   1935
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "��"
         Height          =   345
         Left            =   2820
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   30
         Width           =   345
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "�Ƴ�(&D)"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4170
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   30
         Width           =   915
      End
      Begin MSComctlLib.ListView lvwDept 
         Height          =   2145
         Left            =   30
         TabIndex        =   8
         Top             =   405
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   3784
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "���ÿ���"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label lbl���ÿ��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ÿ���"
         Height          =   210
         Left            =   0
         TabIndex        =   23
         Top             =   105
         Width           =   840
      End
   End
   Begin VB.Frame frmSplit 
      Height          =   5205
      Left            =   5220
      TabIndex        =   20
      Top             =   -150
      Width           =   30
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   360
      Left            =   5460
      TabIndex        =   9
      Top             =   210
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   360
      Left            =   5460
      TabIndex        =   10
      Top             =   690
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   360
      Left            =   5460
      TabIndex        =   11
      Top             =   4290
      Width           =   1100
   End
End
Attribute VB_Name = "frmClinicOfficeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytFun As G_Enum_Fun '0-�鿴,1-���,2-����,3-ɾ��
Private mlngID As Long '��������ID
Private mrs���� As ADODB.Recordset

Private mblnOK As Boolean
Private mstrAddNewItem As String

Public Function ShowMe(frmParent As Form, ByVal bytFun As G_Enum_Fun, _
    Optional ByVal lngID As Long, Optional ByRef strAddNewItem As String) As Boolean
    '�������
    '��Σ�
    '   frmParent - ������
    '   bytFun - ��������, 0-�鿴��1-������2-�޸ģ�3-ɾ��
    '���Σ�
    '   strAddNewItem:������������
    mbytFun = bytFun: mlngID = lngID
    mstrAddNewItem = ""
    
    Err = 0: On Error Resume Next
    mblnOK = False
    Me.Show 1, frmParent
    If mblnOK Then strAddNewItem = mstrAddNewItem
    ShowMe = mblnOK
End Function

Private Sub cboStationNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk�Ƿ�_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdAdd_Click()
    Call SelectDept(True)
End Sub

Private Sub SelectDept(ByVal blnButton As Boolean, Optional strLike As String)
    '����ѡ������ѡ��ʹ�ÿ���
    Dim strSQL As String, rsResult As ADODB.Recordset
    Dim strID As String, str���� As String
    Dim i As Integer, vRect As RECT
    Dim blnCancel As Boolean, strIDs As String
    Dim objItem As ListItem
    
    Err = 0: On Error GoTo ErrHandler
    For i = 1 To lvwDept.ListItems.Count
        strIDs = strIDs & "," & Val(Mid(lvwDept.ListItems(i).Key, 2))
    Next
    If strIDs <> "" Then strIDs = Mid(strIDs, 2)
    
    strSQL = "Select a.ID, a.����, a.����, Upper(a.����) as ����" & vbNewLine & _
            " From ���ű� A,��������˵�� B" & vbNewLine & _
            " Where a.ID=b.����ID " & vbNewLine & _
            "       And (b.�������=1 Or b.�������=3) And b.�������� = '�ٴ�'" & vbNewLine & _
            "       And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)" & vbNewLine
    If blnButton = False Then
        'ģ������
        strSQL = strSQL & _
            "       And (a.���� Like [1] Or a.���� Like [1] Or Upper(a.����) Like Upper([1]))" & vbNewLine
    End If
    If strIDs <> "" Then
        '�ų���ѡ�����
        strSQL = strSQL & _
            "       And a.ID Not In(Select Column_Value From Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)))" & vbNewLine
    End If
    strSQL = strSQL & " Order By a.����"
    vRect = zlControl.GetControlRect(txtSelect.Hwnd)
    Set rsResult = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "����", False, "", "", False, False, IIf(blnButton = False, True, False), _
        vRect.Left, vRect.Top, txtSelect.Height, blnCancel, True, False, strLike & "%", strIDs)
    If blnCancel Then Exit Sub
    If rsResult Is Nothing Then Exit Sub
    If rsResult.EOF Then Exit Sub
    
    Do While Not rsResult.EOF
        strID = Nvl(rsResult!id): str���� = Nvl(rsResult!����)
        For i = 1 To lvwDept.ListItems.Count
            If Mid(lvwDept.ListItems(i).Key, 2) = strID Then Exit Sub
        Next
        Set objItem = lvwDept.ListItems.Add(, "K" & strID, str����)
        rsResult.MoveNext
    Loop
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdRemove_Click()
    Err = 0: On Error GoTo ErrHandler
    If lvwDept.SelectedItem Is Nothing Then Exit Sub
    
    lvwDept.ListItems.Remove lvwDept.SelectedItem.Key
    If lvwDept.ListItems.Count > 0 Then
        lvwDept.ListItems(1).Selected = True
    End If
    
    If lvwDept.SelectedItem Is Nothing Then cmdRemove.Enabled = False: Exit Sub
    Call lvwDept_ItemClick(lvwDept.SelectedItem)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Form_Activate()
    If Me.ActiveControl Is txt���� And txt����.Text <> "" Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    Err = 0: On Error GoTo ErrHandler
    
    Me.Caption = Choose(mbytFun + 1, "�鿴", "����", "�޸�", "ɾ��") & "��������"
    If InitFaceEx() = False Then Unload Me: Exit Sub
    If mbytFun = Fun_Add Or mbytFun = Fun_Update Then
        If InitData() = False Then Unload Me: Exit Sub
    End If
    If mbytFun = Fun_Add Then
        txt����.Text = GetMaxLocalCode("��������")
        Exit Sub
    End If
    
    Select Case mbytFun
    Case Fun_View
        cmdCancel.Visible = False
        cmdOk.Left = cmdCancel.Left
        Call SetEnabled(Me.Controls, False)
    Case Fun_Update
        txt����.Enabled = False
    End Select
    If LoadData(mlngID) = False Then Unload Me: Exit Sub
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function InitData() As Boolean
    Dim i As Long, strSQL As String, rsTemp As ADODB.Recordset
    Dim intRow As Integer, intCol As Integer
    
    Err = 0: On Error GoTo ErrHandler
    '����վ������
    strSQL = "Select ���, ���� From Zlnodelist"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cboStationNo.Clear
    cboStationNo.AddItem ""
    Do While Not rsTemp.EOF
        cboStationNo.AddItem Nvl(rsTemp!���) & "-" & Nvl(rsTemp!����)
        If gstrNodeNo = Nvl(rsTemp!���) Then cboStationNo.ListIndex = cboStationNo.NewIndex
        rsTemp.MoveNext
    Loop
    InitData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadData(ByVal lngID As Long) As Boolean
    '��������
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim objItem As Field, Index As Integer, i As Integer
    
    Err = 0: On Error GoTo ErrHandler
    strSQL = "Select b.��� As վ����, a.*" & vbNewLine & _
            " From �������� A,Zlnodelist B" & vbNewLine & _
            " Where a.վ��=b.����(+) And ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
    If rsTemp.EOF Then Exit Function
    
    txt����.Text = Nvl(rsTemp!����)
    txt����.Text = Nvl(rsTemp!����)
    txt����.Text = Nvl(rsTemp!����)
    txtλ��.Text = Nvl(rsTemp!λ��)
    zlControl.CboSetText cboStationNo, Nvl(rsTemp!վ��), False
    If cboStationNo.ListIndex = -1 Then
        cboStationNo.AddItem Nvl(rsTemp!վ����) & "-" & Nvl(rsTemp!վ��)
        cboStationNo.ListIndex = cboStationNo.NewIndex
    End If
    
    '������չ�ֶ�ֵ
    For Each objItem In rsTemp.Fields
        If InStr(",վ����,ID,����,����,����,λ��,ȱʡ��־,վ��,", "," & UCase(objItem.Name) & ",") = 0 Then
            Index = -1
            If objItem.Name Like "�Ƿ�*" Or (objItem.Type = adNumeric And objItem.Precision = 1) Then
                '�ֶ��������Ƿ񡱡�Numeric���ͣ����1B����CheckBox����
                For i = 1 To chk�Ƿ�.UBound
                    If chk�Ƿ�(i).Caption = objItem.Name Then Index = i: Exit For
                Next
                If Index > 0 Then
                    chk�Ƿ�(Index).Value = IIf(Val(Nvl(objItem.Value)) = 0, vbUnchecked, vbChecked)
                End If
            Else
                For i = 1 To lblEdit.UBound
                    If lblEdit(i).Caption = objItem.Name Then Index = i: Exit For
                Next
                If Index > 0 Then
                    If Val(lblEdit(Index).Tag) = 2 Then '������
                        txtEdit(Index).Text = Format(Nvl(objItem.Value), "yyyy-mm-dd")
                    Else
                        txtEdit(Index).Text = Nvl(objItem.Value)
                    End If
                End If
            End If
        End If
    Next
    
    '���ÿ���
    lvwDept.ListItems.Clear
    strSQL = "Select b.Id, b.����" & vbNewLine & _
            " From �����������ÿ��� A, ���ű� B" & vbNewLine & _
            " Where a.����id = b.Id And a.����id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
    If rsTemp.EOF Then LoadData = True: Exit Function
    
    Do Until rsTemp.EOF
        lvwDept.ListItems.Add , "K" & Nvl(rsTemp!id), Nvl(rsTemp!����)
        rsTemp.MoveNext
    Loop
        
    LoadData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    
    Err = 0: On Error GoTo ErrHandler
    If mbytFun = Fun_View Then Unload Me: Exit Sub
    
    cmdOk.Enabled = False
    If IsValied() = False Then cmdOk.Enabled = True: Exit Sub
    If SaveData() = False Then cmdOk.Enabled = True: Exit Sub
    
    mblnOK = True
    mstrAddNewItem = Trim(txt����.Text)
    If mbytFun = Fun_Add Then
        Call ClearFaceInfor
        cmdOk.Enabled = True
        Exit Sub
    End If
    Unload Me
    Exit Sub
ErrHandler:
    cmdOk.Enabled = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ClearFaceInfor()
    '����:���������Ϣ���Ա�������������
    Dim i As Integer
    
    On Error GoTo errHandle
    txt����.Text = GetMaxLocalCode("��������")
    txt����.Text = ""
    txt����.Text = ""
    txtλ��.Text = ""
    txtSelect.Text = ""
    
    For i = 1 To txtEdit.UBound
        txtEdit(i).Text = ""
    Next
    
    For i = 1 To chk�Ƿ�.UBound
        chk�Ƿ�(i).Value = vbUnchecked
    Next
    
    lvwDept.ListItems.Clear
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function SaveData() As Boolean
    Dim strSQL As String, i As Long
    Dim strTemp As String, str���ÿ��� As String
    
    Err = 0: On Error GoTo ErrHandler
    
    For i = 1 To lvwDept.ListItems.Count
        strTemp = Val(Mid(lvwDept.ListItems(i).Key, 2))
        str���ÿ��� = str���ÿ��� & ";" & strTemp
    Next
    If str���ÿ��� <> "" Then str���ÿ��� = Mid(str���ÿ���, 2)
    
    Select Case mbytFun
    Case Fun_Add
        'Zl_��������_Modify(
        strSQL = "Zl_��������_Modify("
        '��������_In Number,--0-������1-�޸�
        strSQL = strSQL & "" & 0 & ","
        'Id_In       ��������.Id%Type,
        strSQL = strSQL & "" & "NULL" & ","
        '����_In     ��������.����%Type := Null,
        strSQL = strSQL & "'" & Trim(txt����.Text) & "',"
        '����_In     ��������.����%Type := Null,
        strSQL = strSQL & "'" & Trim(txt����.Text) & "',"
        '����_In     ��������.����%Type := Null,
        strSQL = strSQL & "'" & Trim(txt����.Text) & "',"
        'λ��_In     ��������.λ��%Type := Null,
        strSQL = strSQL & "'" & Trim(txtλ��.Text) & "',"
        'վ��_In     ��������.վ��%Type := Null,
        strSQL = strSQL & "'" & NeedCode(cboStationNo.Text) & "',"
        '���ÿ���_In Varchar2:=Null--��ʽ������1;����2;����3;...
        strSQL = strSQL & "'" & str���ÿ��� & "',"
        '��չ_In Varchar2:=Null--�û���չ�ֶ�ֵ����ʽ���ֶ���1=ֵ1,�ֶ���2=ֵ2,...
        strSQL = strSQL & "'" & ExpandSaveStr() & "')"
    Case Fun_Update
        'Zl_��������_Modify(
        strSQL = "Zl_��������_Modify("
        '��������_In Number,--0-������1-�޸�
        strSQL = strSQL & "" & 1 & ","
        'Id_In       ��������.Id%Type,
        strSQL = strSQL & "" & mlngID & ","
        '����_In     ��������.����%Type := Null,
        strSQL = strSQL & "'" & Trim(txt����.Text) & "',"
        '����_In     ��������.����%Type := Null,
        strSQL = strSQL & "'" & Trim(txt����.Text) & "',"
        '����_In     ��������.����%Type := Null,
        strSQL = strSQL & "'" & Trim(txt����.Text) & "',"
        'λ��_In     ��������.λ��%Type := Null,
        strSQL = strSQL & "'" & Trim(txtλ��.Text) & "',"
        'վ��_In     ��������.վ��%Type := Null,
        strSQL = strSQL & "'" & NeedCode(cboStationNo.Text) & "',"
        '���ÿ���_In Varchar2:=Null--��ʽ������1;����2;����3;...
        strSQL = strSQL & "'" & str���ÿ��� & "',"
        '��չ_In Varchar2:=Null--�û���չ�ֶ�ֵ����ʽ���ֶ���1=ֵ1,�ֶ���2=ֵ2,...
        strSQL = strSQL & "'" & ExpandSaveStr() & "')"
    End Select
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function IsValied() As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHandler
    If zlControl.FormCheckInput(Me) = False Then Exit Function
    If zlControl.TxtCheckInput(txt����, "����", 3, False) = False Then Exit Function
    If zlControl.TxtCheckInput(txt����, "����", 20, False) = False Then Exit Function
    If zlControl.TxtCheckInput(txt����, "����", 6, False) = False Then Exit Function
    If zlControl.TxtCheckInput(txtλ��, "λ��", 40, True) = False Then Exit Function
    
    If IsValidEx() = False Then Exit Function
    
    If mbytFun = Fun_Add Then
        strSQL = "Select 1 From �������� Where ���� = [1] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Trim(txt����.Text))
        If Not rsTemp Is Nothing Then
            If Not rsTemp.EOF Then
                MsgBox Trim(txt����.Text) & " �Ѵ��ڣ�", vbInformation, gstrSysName
                If txt����.Visible And txt����.Enabled Then txt����.SetFocus
                zlControl.TxtSelAll txt����
                Exit Function
            End If
        End If
    ElseIf mbytFun = Fun_Update Then
        strSQL = "Select 1 From �������� Where ���� = [1] And ID <> [2] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Trim(txt����.Text), mlngID)
        If Not rsTemp Is Nothing Then
            If Not rsTemp.EOF Then
                MsgBox Trim(txt����.Text) & " �Ѵ��ڣ�", vbInformation, gstrSysName
                If txt����.Visible And txt����.Enabled Then txt����.SetFocus
                zlControl.TxtSelAll txt����
                Exit Function
            End If
        End If
    End If
    IsValied = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    If Not mrs���� Is Nothing Then Set mrs���� = Nothing
End Sub

Private Sub lvwDept_GotFocus()
    cmdRemove.Enabled = Not lvwDept.SelectedItem Is Nothing
    If lvwDept.ListItems.Count = 0 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub lvwDept_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvwDept.SelectedItem Is Nothing Then Exit Sub
    cmdRemove.Enabled = True
End Sub

Private Sub lvwDept_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtSelect_GotFocus()
    zlControl.TxtSelAll txtSelect
End Sub

Private Sub txtSelect_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(txtSelect.Text) = "" Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        Call SelectDept(False, Trim(txtSelect.Text))
        zlControl.TxtSelAll txtSelect
    End If
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt����_Change()
    txt����.Text = zlCommFun.SpellCode(txt����.Text)
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(txt����.Text) = "" Then
            MsgBox "���Ʋ���Ϊ�գ�", vbInformation, gstrSysName
            txt����.SetFocus: Exit Sub
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtλ��_GotFocus()
    zlControl.TxtSelAll txtλ��
End Sub

Private Sub txtλ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Function InitFaceEx() As Boolean
    '��ʼ�����棬��̬�����û���չ�ֶ�,113315
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim Index As Integer, i As Integer
    Dim objItem As Field
    Dim intTabIndex As Integer
    Dim sngAddHeight As Single
    Dim sngTop As Single, sngSplit As Single
    
    On Error GoTo errHandle
    strSQL = "Select * From �������� Where 1 = 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�������ұ�ṹ")
    
    For Each objItem In rsTemp.Fields
        If InStr(",ID,����,����,����,λ��,ȱʡ��־,վ��,", "," & UCase(objItem.Name) & ",") = 0 Then
            If objItem.Name Like "�Ƿ�*" Or (objItem.Type = adNumeric And objItem.Precision = 1) Then
                '�ֶ��������Ƿ񡱡�Numeric���ͣ����1B����CheckBox����
                Index = chk�Ƿ�.Count
                
                Load chk�Ƿ�(Index): Set chk�Ƿ�(Index).Container = fra������Ϣ
                chk�Ƿ�(Index).Visible = True
                chk�Ƿ�(Index).Caption = objItem.Name
                
                chk�Ƿ�(Index).Width = 300 + Me.TextWidth(chk�Ƿ�(Index).Caption)
                If chk�Ƿ�(Index).Left + chk�Ƿ�(Index).Width - 100 > fra������Ϣ.Width Then
                    chk�Ƿ�(Index).Width = fra������Ϣ.Width - chk�Ƿ�(Index).Left - 100
                End If
            Else
                Index = lblEdit.Count
                
                Load lblEdit(Index): Set lblEdit(Index).Container = fra������Ϣ
                Load txtEdit(Index): Set txtEdit(Index).Container = fra������Ϣ
                lblEdit(Index).Visible = True
                txtEdit(Index).Visible = True
                lblEdit(Index).Caption = objItem.Name
                
                '�ֶ�����,Ϊ1��ʾ������,2��ʾ����
                If objItem.Type = adNumeric Then
                    lblEdit(Index).Tag = 1
                    txtEdit(Index).MaxLength = objItem.Precision
                ElseIf objItem.Type = adDate Or objItem.Type = adDBTimeStamp _
                    Or objItem.Type = adDBDate Or objItem.Type = adDBTime Then
                    lblEdit(Index).Tag = 2
                    txtEdit(Index).MaxLength = 10
                Else
                    lblEdit(Index).Tag = 3
                    txtEdit(Index).MaxLength = objItem.DefinedSize
                End If
                
                txtEdit(Index).Left = lblEdit(Index).Left + lblEdit(Index).Width + 30
                txtEdit(Index).Width = fra������Ϣ.Width - txtEdit(Index).Left - 90
            End If
        End If
    Next
    
    sngTop = txtλ��.Top + txtλ��.Height
    intTabIndex = txtλ��.TabIndex + 1
    sngAddHeight = 0
    sngSplit = 85
    
    For i = 1 To lblEdit.UBound
        txtEdit(i).Top = sngTop + sngSplit
        lblEdit(i).Top = txtEdit(i).Top + 70
        txtEdit(i).TabIndex = intTabIndex '����Tab˳��
        
        sngTop = txtEdit(i).Top + txtEdit(i).Height
        intTabIndex = intTabIndex + 1
        sngAddHeight = sngAddHeight + txtEdit(i).Height + sngSplit
    Next
    
    For i = 1 To chk�Ƿ�.UBound
        chk�Ƿ�(i).Top = sngTop + sngSplit
        chk�Ƿ�(i).TabIndex = intTabIndex '����Tab˳��
        
        sngTop = chk�Ƿ�(i).Top + chk�Ƿ�(i).Height
        intTabIndex = intTabIndex + 1
        sngAddHeight = sngAddHeight + chk�Ƿ�(i).Height + sngSplit
    Next
    
    fra������Ϣ.Height = fra������Ϣ.Height + sngAddHeight
    fraDept.Top = fraDept.Top + sngAddHeight
    frmSplit.Height = frmSplit.Height + sngAddHeight
    
    Me.Height = Me.Height + sngAddHeight
    cmdHelp.Top = Me.ScaleHeight - cmdHelp.Height - 200
    
    InitFaceEx = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function IsValidEx() As Boolean
    '�����û���չ�ֶ�������������Ƿ���Ч,113315
    Dim i As Integer
    Dim strTemp As String
    
    On Error GoTo errHandle
    For i = 1 To lblEdit.UBound
        strTemp = Trim(txtEdit(i).Text)
        If zlCommFun.StrIsValid(strTemp, txtEdit(i).MaxLength, txtEdit(i).Hwnd) = False Then
            zlControl.TxtSelAll txtEdit(i)
            Exit Function
        End If
        
        Select Case Val(lblEdit(i).Tag)
        Case 1 '�������ֶ�
            If strTemp <> "" And Not IsNumeric(strTemp) Then
                MsgBox lblEdit(i).Caption & "Ӧ���������֡�", vbExclamation, gstrSysName
                zlControl.TxtSelAll txtEdit(i)
                If txtEdit(i).Visible And txtEdit(i).Enabled Then txtEdit(i).SetFocus
                Exit Function
            End If
        Case 2 '�������ֶ�
            strTemp = zlCommFun.AddDate(strTemp)
            If strTemp <> "" Then
                If Not IsDate(strTemp) Then
                    MsgBox lblEdit(i).Caption & "������Ч�����ڸ�ʽ(yyyy-mm-dd)��(yyyymmdd)��", vbExclamation, gstrSysName
                    zlControl.TxtSelAll txtEdit(i)
                    If txtEdit(i).Visible And txtEdit(i).Enabled Then txtEdit(i).SetFocus
                    Exit Function
                End If
                
                Err = 0: On Error Resume Next
                strTemp = Format(strTemp, "yyyy-mm-dd")
                If Err <> 0 Or Not IsDate(strTemp) Then
                    MsgBox lblEdit(i).Caption & "������Ч�����ڸ�ʽ(yyyy-mm-dd)��(yyyymmdd)��", vbExclamation, gstrSysName
                    zlControl.TxtSelAll txtEdit(i)
                    If txtEdit(i).Visible And txtEdit(i).Enabled Then txtEdit(i).SetFocus
                    Exit Function
                End If
                
                txtEdit(i).Text = strTemp
            End If
        End Select
    Next
    IsValidEx = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExpandSaveStr() As String
    '��ȡ��չ�ֶα����ַ���,113315
    Dim strSQL As String
    Dim i As Integer
    
    On Error GoTo errHandle
    For i = 1 To lblEdit.UBound
        strSQL = strSQL & "," & lblEdit(i).Caption & "="
        Select Case Val(lblEdit(i).Tag)
        Case 1 '��ֵ��
            strSQL = strSQL & Val(txtEdit(i).Text)
        Case 2   '������
            strSQL = strSQL & "To_Date('" & Format(Trim(txtEdit(i).Text), "yyyy-mm-dd") & "','yyyy-mm-dd')"
        Case Else
            strSQL = strSQL & "'" & Trim(txtEdit(i).Text) & "'"
        End Select
    Next
    
    For i = 1 To chk�Ƿ�.UBound
        strSQL = strSQL & "," & chk�Ƿ�(i).Caption & "=" & IIf(chk�Ƿ�(i).Value = 1, "1", "0")
    Next
    If strSQL <> "" Then strSQL = Mid(strSQL, 2)
    
    ExpandSaveStr = Replace(strSQL, "'", "''")
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
