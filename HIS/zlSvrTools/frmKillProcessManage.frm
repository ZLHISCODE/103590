VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmKillProcessManage 
   Caption         =   "�����嵥����"
   ClientHeight    =   7605
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11625
   Icon            =   "frmKillProcessManage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   11625
   StartUpPosition =   1  '����������
   Begin VB.Frame fraSplit 
      Height          =   30
      Left            =   -15
      TabIndex        =   10
      Top             =   930
      Width           =   11970
   End
   Begin VB.PictureBox picEdit 
      Height          =   5670
      Left            =   7860
      ScaleHeight     =   5610
      ScaleWidth      =   3600
      TabIndex        =   2
      Top             =   1005
      Width           =   3660
      Begin VB.CommandButton cmdDel 
         Caption         =   "ɾ��(&D)"
         Enabled         =   0   'False
         Height          =   345
         Left            =   1766
         TabIndex        =   17
         Top             =   5025
         Width           =   800
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "�޸�(&U)"
         Enabled         =   0   'False
         Height          =   345
         Left            =   958
         TabIndex        =   16
         Top             =   5025
         Width           =   800
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "����(&A)"
         Enabled         =   0   'False
         Height          =   345
         Left            =   150
         TabIndex        =   15
         Top             =   5025
         Width           =   800
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   345
         Left            =   2475
         TabIndex        =   14
         Top             =   3240
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "�˳�(&D)"
         Height          =   345
         Left            =   2575
         TabIndex        =   13
         Top             =   5025
         Width           =   800
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   345
         Left            =   1575
         TabIndex        =   8
         Top             =   3240
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   645
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   255
         Width           =   2745
      End
      Begin VB.TextBox txtDescription 
         Height          =   1600
         Left            =   645
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "��������������100�����ֻ�200���ַ�"
         Top             =   1410
         Width           =   2745
      End
      Begin VB.ComboBox cboType 
         Enabled         =   0   'False
         Height          =   300
         Left            =   645
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   840
         Width           =   2745
      End
      Begin VB.Label lblAdd 
         Height          =   180
         Left            =   660
         TabIndex        =   12
         Top             =   3315
         Width           =   720
      End
      Begin VB.Label lblShowCheckName 
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   645
         TabIndex        =   11
         Top             =   570
         Width           =   2760
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   150
         TabIndex        =   9
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   150
         TabIndex        =   7
         Top             =   870
         Width           =   360
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   150
         TabIndex        =   5
         Top             =   1410
         Width           =   360
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfProcessList 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   1005
      Width           =   7740
      _cx             =   13652
      _cy             =   9975
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   260
      RowHeightMax    =   260
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmKillProcessManage.frx":6852
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Image imgMain 
      Height          =   720
      Left            =   210
      Picture         =   "frmKillProcessManage.frx":68EB
      Top             =   105
      Width           =   720
   End
   Begin VB.Label lblCaption 
      Height          =   540
      Left            =   1170
      TabIndex        =   1
      Top             =   210
      Width           =   10395
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "�����˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuAdd 
         Caption         =   "����(&A)"
      End
      Begin VB.Menu mnuPopuModify 
         Caption         =   "�޸�(&U)"
      End
      Begin VB.Menu mnuPopuDelete 
         Caption         =   "ɾ��(&D)"
      End
   End
End
Attribute VB_Name = "frmKillProcessManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnRightClick As Boolean   '����Ƿ�Ϊ����Ҽ��������Ҫ���������Ҽ�����б���vsfProcessList_Click����ʽ����
Private mblnOk As Boolean  '��������Ƿ񱣴�ɹ�
Private Enum ProcessList
    PL_��� = 0
    PL_���� = 1
    PL_���� = 2
    PL_���� = 3
    PL_���� = 4
End Enum

Public Sub ShowMe(ByVal strModule As String)
'strModule�����øô����ģ���
    Select Case strModule
        Case "0102"   'ϵͳ��Ǩ����
            Me.Caption = "�жϿͻ������ӵĽ��̶���"
            lblCaption = "��ϵͳ��Ǩ�����У������ڽ��̶����ݿ���в�����ʹ�ã���Ӱ������Ч�ʡ�" & vbNewLine & _
                        "��һ���棬����ʱ��ṹ����������ɱ������ɱ������ʹ����ʱ��ĻỰ�������ʶ�Ự�Ľ��̣���ֹ��ɱ�����ܽ��е�����" & vbNewLine & _
                        "���ڲ�Ʒ���ݽṹ�Ľ��̶�Ӧ�ü��뵽���嵥�С�ʹ�õ����ݽṹ�Ͳ�Ʒ���ݽṹ���������ϵ�Ľ��̣�ҲӦ�ü��뵽���嵥��"
        Case "0307"   '�ͻ�����������
            Me.Caption = "�ͻ��˽��̹���"
            lblCaption = "�ͻ����Զ�����ʱ�������������ĳЩӦ�ó��򣬴��������ļ������ռ�þ��޷����滻��" & vbNewLine & _
                        "Ϊ�˱������������������б��е�Ӧ�ó��򽫻ᱻ���������Զ���ֹ��"
    End Select
    Me.Show vbModal, frmMDIMain
End Sub

Private Sub cmdCancel_Click()
    Call FillSelectData(vsfProcessList.Row)
    Call SetEnable(False)
    mblnRightClick = False
    Call vsfProcessList_Click
    lblAdd.Caption = ""
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    mblnOk = False
    '������������Ƿ����Ҫ��
    If CheckData = False Then Exit Sub
    
    txtName.Tag = txtName.Text
    cboType.Tag = cboType.Text
    txtDescription.Tag = txtDescription.Text
    If Val(mnuPopuAdd.Tag) = 1 Then
        '��������
        Call ExecuteProcedure("Zl_Zlkillprocess_Edit(1, Null, '" & UCase(Trim(txtName.Text)) & "', " & cboType.ListIndex & ", '" & txtDescription.Text & "')", Me.Caption)
        strSQL = "Select ��� From Zlkillprocess Where ���� = [1]"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, UCase(Trim(txtName.Text)))
        If rsTmp.RecordCount > 0 Then
            lblAdd.Caption = "��ӳɹ�"
            vsfProcessList.Rows = vsfProcessList.Rows + 1
            vsfProcessList.TextMatrix(vsfProcessList.Rows - 1, PL_���) = rsTmp!���
            vsfProcessList.TextMatrix(vsfProcessList.Rows - 1, PL_����) = "�Զ���"
            vsfProcessList.TextMatrix(vsfProcessList.Rows - 1, PL_����) = UCase(Trim(txtName.Text))
            vsfProcessList.TextMatrix(vsfProcessList.Rows - 1, PL_����) = cboType.Text
            vsfProcessList.TextMatrix(vsfProcessList.Rows - 1, PL_����) = txtDescription.Text
            vsfProcessList.Row = vsfProcessList.Rows - 1
            Call vsfProcessList.ShowCell(vsfProcessList.Row, PL_���)
        End If
        mblnOk = True
        '��ӳɹ����Զ���������״̬���������Ч��
        Call mnuPopuAdd_Click
    Else
        '�޸�����
        Call ExecuteProcedure("Zl_Zlkillprocess_Edit(2, " & vsfProcessList.TextMatrix(vsfProcessList.Row, PL_���) & ", '" & _
                                UCase(Trim(txtName.Text)) & "', " & cboType.ListIndex & ", '" & txtDescription.Text & "')", Me.Caption)
        lblAdd.Caption = "�޸ĳɹ�"
        vsfProcessList.TextMatrix(vsfProcessList.Row, PL_����) = UCase(Trim(txtName.Text))
        vsfProcessList.TextMatrix(vsfProcessList.Row, PL_����) = cboType.Text
        vsfProcessList.TextMatrix(vsfProcessList.Row, PL_����) = txtDescription.Text
        Call SetEnable(False)
        mblnOk = True
        mblnRightClick = False
        Call vsfProcessList_Click
    End If
    Exit Sub
errH:
    If Val(mnuPopuAdd.Tag) = 1 Then
        MsgBox "���ʧ�ܣ�" & vbNewLine & err.Description, vbInformation, gstrSysName
    Else
        MsgBox "�޸�ʧ�ܣ�" & vbNewLine & err.Description, vbInformation, gstrSysName
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '��ֹ���뵥����
    If InStr("'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    cboType.addItem "����"
    cboType.addItem "����"
    '����������
    Call FillProcessData
End Sub

Private Sub FillProcessData()
    Dim strSQL  As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo errH
    strSQL = "Select ���, ����, ����, ����, �Ƿ�̶� From Zltools.Zlkillprocess Order By ���"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    With rsTmp
        vsfProcessList.Rows = .RecordCount + 1
        For i = 1 To .RecordCount
            vsfProcessList.TextMatrix(i, PL_���) = !���
            vsfProcessList.TextMatrix(i, PL_����) = IIf(!�Ƿ�̶� = 1, "�̶�", "�Զ���")
            vsfProcessList.TextMatrix(i, PL_����) = !����
            vsfProcessList.TextMatrix(i, PL_����) = IIf(!���� = 1, "����", "����")
            vsfProcessList.TextMatrix(i, PL_����) = !���� & ""
            .MoveNext
        Next
        vsfProcessList.ScrollTrack = True
        If .RecordCount > 0 Then
            vsfProcessList.Row = 1
            Call FillSelectData(1)
            Call vsfProcessList_Click
            If txtName.Locked = True Then
                txtName.ForeColor = &H80000011
                txtDescription.ForeColor = &H80000011
            End If
        End If
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '����ǰ���ڽ����������޸Ĳ��������Ѿ������ݽ������޸ģ��򵯳���ʾ���Ƿ���Ҫ����
    If UnloadMode = vbFormControlMenu Or UnloadMode = vbFormCode Then
        Select Case CheckChange
            Case 2    '�������޸ģ�����ѡ���˱��棬�ұ���ʧ���˵�
                Cancel = 1
        End Select
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picEdit.Height = Me.ScaleHeight - picEdit.Top
    picEdit.Left = Me.ScaleWidth - picEdit.Width
    vsfProcessList.Width = Me.ScaleWidth - picEdit.Width
    vsfProcessList.Height = picEdit.Height
    cmdAdd.Top = picEdit.Height - cmdAdd.Height - 255
    cmdUpdate.Top = cmdAdd.Top
    cmdDel.Top = cmdAdd.Top
    cmdExit.Top = cmdAdd.Top
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnRightClick = False
    mblnOk = False
End Sub

Private Sub cmdAdd_Click()
    Call mnuPopuAdd_Click
End Sub

Private Sub cmdUpdate_Click()
    Call mnuPopuModify_Click
End Sub

Private Sub cmdDel_Click()
    Call mnuPopuDelete_Click
End Sub

'��������������Ϣ
Private Sub mnuPopuAdd_Click()
    mnuPopuAdd.Tag = 1   '��ǵ�ǰ״̬Ϊ����
    Call FillSelectData(-1)
    cmdAdd.Enabled = False
    cmdUpdate.Enabled = False
    cmdDel.Enabled = False
    Call SetEnable(True)
    txtName.SetFocus
End Sub

'ɾ������������Ϣ
Private Sub mnuPopuDelete_Click()
    On Error GoTo errH
    If MsgBox("ȷ��Ҫ������Ϊ��" & vsfProcessList.TextMatrix(vsfProcessList.Row, PL_����) & "�������" & vsfProcessList.TextMatrix(vsfProcessList.Row, PL_����) & "ɾ����", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbOK Then
        Call ExecuteProcedure("Zl_Zlkillprocess_Edit(3, Null, '" & vsfProcessList.TextMatrix(vsfProcessList.Row, PL_����) & "')", Me.Caption)
        vsfProcessList.RemoveItem (vsfProcessList.Row)
        Call FillSelectData(vsfProcessList.Row)
    End If
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

'�޸ķ���������Ϣ
Private Sub mnuPopuModify_Click()
    mnuPopuAdd.Tag = 2  '��ǵ�ǰ״̬Ϊ�޸�
    cmdAdd.Enabled = False
    cmdUpdate.Enabled = False
    cmdDel.Enabled = False
    Call SetEnable(True)
    txtName.SetFocus
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
    '�����ı����ȣ��Լ����λ��з�
    If (ActualLen(txtDescription.Text) >= 200 And KeyAscii <> 8) Or KeyAscii = 9 Or KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    '��������\/:*?"'<>|
    If InStr("\/:*?""<>|", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtName_LostFocus()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset

    lblAdd.Caption = ""
    If txtName.Locked = True Then Exit Sub
    If Val(mnuPopuAdd.Tag) = 1 Then
        '����״̬�£����������Ƿ����ظ��ģ�����У����·���ǩҳ�ϸ���������ʾ
        strSQL = "Select Count(1) ���� From Zlkillprocess Where ���� = [1]"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, UCase(Trim(txtName.Text)))
        If rsTmp!���� = 1 Then
            lblShowCheckName.Caption = "�����Ѵ��ڣ�"
        Else
            lblShowCheckName.Caption = ""
        End If
    Else
        '�޸�״̬�£��ȼ��������û�б��޸ģ������޸��ˣ��ټ��������Ƿ����ظ�������У����·���ǩҳ�ϸ���������ʾ
        If txtName.Text <> txtName.Tag Then
            strSQL = "Select Count(1) ���� From Zlkillprocess Where ���� = [1]"
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, UCase(Trim(txtName.Text)))
            If rsTmp!���� = 1 Then
                lblShowCheckName.Caption = "�����Ѵ��ڣ�"
            Else
                lblShowCheckName.Caption = ""
            End If
        Else
            lblShowCheckName.Caption = ""
        End If
    End If
End Sub

Private Sub vsfProcessList_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    '�ڻ���֮ǰ���Ƚ������ݼ�飬���Ƿ�����ݽ������޸�
    Select Case CheckChange
        Case 1, 4  '�������޸ģ�����ѡ���˱��棬�ұ���ɹ��˵� �� δ�����޸�
            Call FillSelectData(NewRow)
            Call SetEnable(False)
        Case 2   '�������޸ģ�����ѡ���˱��棬�ұ���ʧ���˵�
            Cancel = True
        Case 3   '�������޸ģ�����ѡ���˲�����
            Call FillSelectData(OldRow)
            Cancel = True
            Call SetEnable(False)
    End Select
End Sub

Private Sub vsfProcessList_Click()
    If mblnRightClick Then Exit Sub
    With vsfProcessList
        '��ԭ�����˵���Ǽ���ͣ״̬
        '�������ԭ�Ļ�������ǰ�����������޸�ģʽ�£�����û���޸����ݣ���ʱ�Ҽ��л��У���ť��ͣ״̬�ͻ�����
        If txtName.Locked = True Then
            mnuPopuAdd.Tag = ""
            mnuPopuAdd.Enabled = True
            mnuPopuDelete.Enabled = .TextMatrix(.Row, PL_����) <> "�̶�"
            mnuPopuModify.Enabled = mnuPopuDelete.Enabled
            cmdAdd.Enabled = mnuPopuAdd.Enabled
            cmdUpdate.Enabled = mnuPopuModify.Enabled
            cmdDel.Enabled = mnuPopuDelete.Enabled
        End If
    End With
End Sub

Private Sub vsfProcessList_DblClick()
    With vsfProcessList
        If .MouseRow <> .Row Or vsfProcessList.TextMatrix(vsfProcessList.Row, PL_����) = "�̶�" Then Exit Sub
        Call mnuPopuModify_Click
    End With
End Sub

Private Sub vsfProcessList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '���ý���Ϊϵͳ���õĽ��̣�����������޸ĺ�ɾ������
    mnuPopuDelete.Enabled = vsfProcessList.TextMatrix(vsfProcessList.Row, PL_����) <> "�̶�"
    mnuPopuModify.Enabled = mnuPopuDelete.Enabled
    '�ж����Ƿ������������޸�״̬�����ǣ��������ٴν����������޸Ļ�ɾ������
    If txtName.Locked = False Then
        mnuPopuModify.Enabled = False
        mnuPopuAdd.Enabled = False
        mnuPopuDelete.Enabled = False
    End If
    cmdAdd.Enabled = mnuPopuAdd.Enabled
    cmdUpdate.Enabled = mnuPopuModify.Enabled
    cmdDel.Enabled = mnuPopuDelete.Enabled
    
    With vsfProcessList
        '�Ҽ�ĳһ��
        If .MouseRow <> -1 And .MouseRow <> 0 And Button = 2 Then
            If .MouseRow <> .Row Then
                '��ѡ����ǲ�ͬ�У�����ѡ��
                .Row = .MouseRow
                mblnRightClick = False
                Call vsfProcessList_Click
            End If
        End If
    End With
End Sub

Private Sub vsfProcessList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnRightClick = False
    If Button = 1 Then Exit Sub
    With vsfProcessList
        If .MouseRow <> .Row Then Exit Sub
        mblnRightClick = True
        PopupMenu mnuPopu
    End With
End Sub

Private Sub SetEnable(ByVal blnEnable As Boolean)
    With vsfProcessList
        txtName.Locked = Not blnEnable
        cboType.Enabled = blnEnable
        txtDescription.Locked = Not blnEnable
        cmdOK.Visible = blnEnable
        cmdCancel.Visible = blnEnable
        '�����ǰ�༭����������״̬���ͽ�ǰ��ɫ�û�
        If txtName.Locked = True Then
            txtName.ForeColor = &H80000011
            txtDescription.ForeColor = &H80000011
        Else
            txtName.ForeColor = &H80000009
            txtDescription.ForeColor = &H80000009
        End If
    End With
    lblShowCheckName.Caption = ""
End Sub

Private Function CheckChange() As Long
    '�жϵ�ǰ���Ƿ��Ѿ����޸ģ�����ǣ��򵯳���ʾ
    '���������޸ģ�����ѡ���˱��棬�ұ���ɹ��˵ģ�����1
    '���������޸ģ�����ѡ���˱��棬�ұ���ʧ���˵ģ�����2
    '���������޸ģ�����ѡ���˲����棬����3
    '��δ�����޸ģ�����4
    If txtName.Text <> txtName.Tag Or cboType.Text <> cboType.Tag Or txtDescription.Text <> txtDescription.Tag Then
        If MsgBox("��ǰ���ѱ��޸ģ��Ƿ񱣴棿", vbInformation + vbOKCancel + vbDefaultButton1, gstrSysName) = vbOK Then
            Call cmdOK_Click
            If mblnOk Then
                CheckChange = 1
            Else
                CheckChange = 2
            End If
        Else
            CheckChange = 3
        End If
    Else
        CheckChange = 4
    End If
End Function

Private Function CheckData() As Boolean
    '�жϵ�ǰ���������Ƿ����Ҫ��
    If Right(UCase(Trim(txtName.Text)), 4) <> ".EXE" Then
        MsgBox "���Ʋ���һ����ִ���ļ���*.EXE��,����е�����", vbInformation, gstrSysName
        txtName.SetFocus
        Exit Function
    End If
    If InStr(txtName.Text, "'") > 0 Or InStr(txtName.Text, """") > 0 Or InStr(txtName.Text, "\") > 0 Or _
        InStr(txtName.Text, "/") > 0 Or InStr(txtName.Text, ":") > 0 Or InStr(txtName.Text, "*") > 0 Or _
        InStr(txtName.Text, "?") > 0 Or InStr(txtName.Text, "<") > 0 Or InStr(txtName.Text, ">") > 0 Or InStr(txtName.Text, "|") > 0 Then
        MsgBox "�������к��зǷ��ַ�(\/:*?""'<>|)����������д��", vbInformation, gstrSysName
        txtName.SetFocus
        Exit Function
    End If
    If StrIsValid(txtDescription.Text, 200) = False Then
        txtDescription.SetFocus
        Exit Function
    End If
    '�������к��л��з�������ȥ��
    If InStr(txtDescription.Text, vbNewLine) > 0 Then
        txtDescription.Text = Replace(txtDescription.Text, vbNewLine, "")
    End If
    CheckData = True
End Function

Private Sub FillSelectData(ByVal lngRow As Long)
'lngRow:vsfProcessList���к�
    If lngRow = -1 Then
        txtName.Text = ""
        cboType.ListIndex = 0
        txtDescription.Text = ""
        txtName.Tag = txtName.Text
        cboType.Tag = cboType.Text
        txtDescription.Tag = txtDescription.Text
    Else
        txtName.Text = vsfProcessList.TextMatrix(lngRow, PL_����)
        cboType.Text = vsfProcessList.TextMatrix(lngRow, PL_����)
        txtDescription.Text = vsfProcessList.TextMatrix(lngRow, PL_����)
        txtName.Tag = txtName.Text
        cboType.Tag = cboType.Text
        txtDescription.Tag = txtDescription.Text
    End If
End Sub
