VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form frmSchemeImport 
   AutoRedraw      =   -1  'True
   Caption         =   "����ҽ������"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9720
   Icon            =   "frmSchemeImport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSelNone 
      Caption         =   "ȫ��(&R)"
      Height          =   350
      Left            =   1815
      TabIndex        =   8
      ToolTipText     =   "Ctrl+R"
      Top             =   5655
      Width           =   1100
   End
   Begin VB.CommandButton cmdSelALL 
      Caption         =   "ȫѡ(&A)"
      Height          =   350
      Left            =   720
      TabIndex        =   7
      ToolTipText     =   "Ctrl+A"
      Top             =   5655
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8160
      TabIndex        =   10
      Top             =   5655
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7065
      TabIndex        =   9
      Top             =   5655
      Width           =   1100
   End
   Begin VB.ComboBox cboTime 
      Height          =   300
      Left            =   3120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   135
      Width           =   3630
   End
   Begin VB.TextBox txtPati 
      Height          =   300
      Left            =   735
      TabIndex        =   1
      ToolTipText     =   "���뷽����ˢ����-����ID��+סԺ�ţ�*����ţ�.�Һŵ�"
      Top             =   135
      Width           =   1275
   End
   Begin VB.ComboBox cboBaby 
      Height          =   300
      Left            =   7530
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   135
      Width           =   2130
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   4965
      Left            =   60
      TabIndex        =   6
      Top             =   570
      Width           =   9585
      _cx             =   16907
      _cy             =   8758
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
      BackColorSel    =   12632256
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   23
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSchemeImport.frx":058A
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   1
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
      FrozenCols      =   1
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ʱ��(&T)"
      Height          =   180
      Left            =   2115
      TabIndex        =   2
      Top             =   195
      Width           =   990
   End
   Begin VB.Label lblPati 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����(&P)"
      Height          =   180
      Left            =   90
      TabIndex        =   0
      Top             =   195
      Width           =   630
   End
   Begin VB.Label lblBaby 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ӥ��(&B)"
      Height          =   180
      Left            =   6870
      TabIndex        =   4
      Top             =   195
      Width           =   630
   End
End
Attribute VB_Name = "frmSchemeImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint��Χ As Integer '1-����,2-סԺ,3-�����סԺ
Private mlng����ID As Long
Private mstrIDs As String
Private mblnOK As Boolean

Private Enum COL���׷���
    colѡ�� = 0
    col��Ч = 1
    col���� = 2
    col���� = 3
    col��λ = 4
    col���� = 5
    col������λ = 6
    colƵ�� = 7
    col�÷� = 8
    col���� = 9
    colִ��ʱ�� = 10
    colִ�п��� = 11
    colִ������ = 12
    col��� = 13
    col��� = 14
    col��ĿID = 15
    col��� = 16
    col�շ�ϸĿID = 17
    col�걾��λ = 18
    col��鷽�� = 19
    colƵ�ʴ��� = 20
    colƵ�ʼ�� = 21
    col�����λ = 22
End Enum

Public Function ShowMe(ByVal frmParent As Object, ByVal int��Χ As Integer, lng����ID As Long) As String
'���أ���ѡ���ҽ����ID
'      lng����ID=����ҽ���Ĳ���
    
    mint��Χ = int��Χ
    lng����ID = 0
    
    Me.Show 1, frmParent
    If mblnOK Then
        ShowMe = mstrIDs
        lng����ID = mlng����ID
    End If
End Function

Private Sub cboBaby_Click()
    Call LoadAdvice
End Sub

Private Sub cboTime_Click()
    Call LoadAdviceBaby
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strIDs As String, i As Long
    Dim lngҽ��ID As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, colѡ��)) <> 0 And Val(.TextMatrix(i, col���)) <> 0 Then
                lngҽ��ID = IIF(Val(.TextMatrix(i, col���)) = 0, Val(.TextMatrix(i, col���)), Val(.TextMatrix(i, col���)))
                If InStr(strIDs & ",", "," & lngҽ��ID & ",") = 0 Then
                    strIDs = strIDs & "," & lngҽ��ID
                End If
            End If
        Next
        If strIDs = "" Then
            MsgBox "û��ѡ���κ�ҽ�����ݡ�", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    
    mstrIDs = Mid(strIDs, 2)
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdSelALL_Click()
    Dim i As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, col��ĿID)) <> 0 Then
                '��ǰ�ļ��ҽ����������Ϊ���׷���
                If .TextMatrix(i, col���) = "D" Then
                    If Val(.TextMatrix(i, col���)) = 0 Then
                        If Not CheckIsOldAdvice(i) Then
                            .TextMatrix(i, colѡ��) = -1
                            Call RowSelectSame(i)
                        End If
                    Else
                        '�������Ѵ���
                    End If
                Else
                    .TextMatrix(i, colѡ��) = -1
                End If
            End If
        Next
    End With
End Sub

Private Sub cmdSelNone_Click()
    Dim i As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, colѡ��) = 0
        Next
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        Call cmdSelALL_Click
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call cmdSelNone_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    mstrIDs = ""
    mblnOK = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    vsAdvice.Width = Me.ScaleWidth - vsAdvice.Left * 2
    
    If vsAdvice.Left + vsAdvice.Width - cboBaby.Width > cboTime.Left + cboTime.Width + lblBaby.Width + 150 Then
        cboBaby.Left = vsAdvice.Left + vsAdvice.Width - cboBaby.Width
    Else
        cboBaby.Left = cboTime.Left + cboTime.Width + lblBaby.Width + 150
    End If
    lblBaby.Left = cboBaby.Left - lblBaby.Width - 30
    
    If Me.ScaleWidth - cmdCancel.Width - cmdSelALL.Left > cmdSelNone.Left + cmdSelNone.Width + cmdOK.Width Then
        cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - cmdSelALL.Left
    Else
        cmdCancel.Left = cmdSelNone.Left + cmdSelNone.Width + cmdOK.Width
    End If
    cmdOK.Left = cmdCancel.Left - cmdOK.Width
    
    vsAdvice.Height = Me.ScaleHeight - vsAdvice.Top - cmdSelNone.Height * 1.6
    cmdSelNone.Top = vsAdvice.Top + vsAdvice.Height + cmdSelNone.Height * 0.3
    cmdSelALL.Top = cmdSelNone.Top
    cmdOK.Top = cmdSelNone.Top
    cmdCancel.Top = cmdOK.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub txtPati_GotFocus()
    Call zlControl.TxtSelAll(txtPati)
End Sub

Private Sub txtPati_KeyPress(KeyAscii As Integer)
    
    '��������س�
    If KeyAscii = 13 And Trim(txtPati.Text) <> "" Then
        KeyAscii = 0
        
        '��ȡ������Ϣ
        If Not GetPatient(Trim(txtPati.Text)) Then
            txtPati.PasswordChar = ""
            txtPati.Text = lblPati.Tag
            Call zlControl.TxtSelAll(txtPati)
        Else
            txtPati.PasswordChar = ""
            vsAdvice.SetFocus
        End If
    End If
End Sub

Private Function GetPatient(ByVal strCode As String) As Boolean
'���ܣ���ȡ������Ϣ������ʾ�ò��˴��ڵ�ҽ��ʱ��
'������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strNO As String, str���� As String, lng����ID As Long
    
    On Error GoTo errH
    
    If Left(strCode, 1) = "-" And IsNumeric(Mid(strCode, 2)) Then '����ID
        strSQL = "Select ����ID,���� From ������Ϣ Where ����ID=[3] "
    ElseIf Left(strCode, 1) = "+" And IsNumeric(Mid(strCode, 2)) Then 'סԺ��
        strSQL = "Select ����ID,���� From ������Ϣ Where סԺ��=[3] "
    ElseIf Left(strCode, 1) = "*" And IsNumeric(Mid(strCode, 2)) Then '�����
        strSQL = "Select ����ID,���� From ������Ϣ Where �����=[3] "
    ElseIf Left(strCode, 1) = "." Then '�Һŵ�
        strNO = GetFullNO(Mid(UCase(strCode), 2), 12)
        strSQL = "Select ����ID,���� From ���˹Һż�¼ Where NO=[4] And ��¼����=1 And ��¼״̬=1"
    Else '��������
        strSQL = "Select ����ID,���� From ������Ϣ Where ����=[1]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strCode, UCase(strCode), Mid(strCode, 2), strNO)
    
    If rsTmp.EOF Then
        mlng����ID = 0
        MsgBox "û���ҵ���صĲ�����Ϣ��", vbInformation, gstrSysName
    Else
        str���� = rsTmp!����: lng����ID = rsTmp!����ID
        strSQL = _
            " Select B.ID as �Һ�ID,A.�Һŵ�,B.�Ǽ�ʱ��,D.���� as �Һſ���," & _
            " A.��ҳID,C.��Ժ����,E.���� as סԺ����,Min(A.����ʱ��) as ˳��" & _
            " From ����ҽ����¼ A,���˹Һż�¼ B,������ҳ C,���ű� D,���ű� E" & _
            " Where A.����ID=[1] And A.�Һŵ�=B.NO(+) And B.��¼����(+)=1 And B.��¼״̬(+)=1" & _
            Decode(mint��Χ, 1, " And A.������Դ=1", 2, " And A.������Դ=2", 3, " And A.������Դ Not IN(3,4)") & _
            " And A.����ID=C.����ID(+) And A.��ҳID=C.��ҳID(+)" & _
            " And B.ִ�в���ID=D.ID(+) And C.��Ժ����ID=E.ID(+)" & _
            " Group by B.ID,A.�Һŵ�,B.�Ǽ�ʱ��,D.����,A.��ҳID,C.��Ժ����,E.����" & _
            " Order by ˳�� Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
        
        If rsTmp.EOF Then
            mlng����ID = 0
            MsgBox "����""" & str���� & """û��ҽ����¼��", vbInformation, gstrSysName
        Else
            txtPati.Text = str����
            lblPati.Tag = str����
            mlng����ID = lng����ID
            cboTime.Clear
            For i = 1 To rsTmp.RecordCount
                If IsNull(rsTmp!��ҳID) And Not IsNull(rsTmp!�Һ�ID) Then
                    cboTime.AddItem rsTmp!�Һſ��� & " " & Format(rsTmp!�Ǽ�ʱ��, "yyyy-MM-dd HH:mm") & " �������"
                    cboTime.ItemData(cboTime.NewIndex) = rsTmp!�Һ�ID
                    If rsTmp!�Һŵ� = strNO Then cboTime.ListIndex = cboTime.NewIndex
                ElseIf Not IsNull(rsTmp!��ҳID) And IsNull(rsTmp!�Һŵ�) Then
                    cboTime.AddItem rsTmp!סԺ���� & " " & Format(rsTmp!��Ժ����, "yyyy-MM-dd HH:mm") & " ��" & rsTmp!��ҳID & "��סԺ"
                    cboTime.ItemData(cboTime.NewIndex) = rsTmp!��ҳID
                End If
                rsTmp.MoveNext
            Next
            If cboTime.ListIndex = -1 Then cboTime.ListIndex = 0
            GetPatient = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadAdviceBaby() As Boolean
'���ܣ���ȡ��ǰ����ָ��ʱ��ҽ����Ӥ���嵥
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strNO As String
    
    If Val(mlng����ID) = 0 Then Exit Function
    If cboTime.ListIndex = -1 Then Exit Function
    
    On Error GoTo errH
    
    If InStr(cboTime.Text, "סԺ") = 0 Then
        strSQL = "Select Distinct Nvl(A.Ӥ��,0) as Ӥ��,C.Ӥ������" & _
            " From ����ҽ����¼ A,���˹Һż�¼ B,������������¼ C" & _
            " Where A.����ID=[1] And A.�Һŵ�=B.NO And B.ID=[2]" & _
            " And A.Ӥ��=C.���(+) And C.����ID(+)=[1] And C.��ҳID(+)=[2]" & _
            " Order by Ӥ��"
    Else
        strSQL = "Select Distinct Nvl(A.Ӥ��,0) as Ӥ��,C.Ӥ������" & _
            " From ����ҽ����¼ A,������������¼ C" & _
            " Where A.����ID=[1] And A.��ҳID=[2]" & _
            " And A.Ӥ��=C.���(+) And C.����ID(+)=[1] And C.��ҳID(+)=[2]" & _
            " Order by Ӥ��"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, cboTime.ItemData(cboTime.ListIndex))
    
    cboBaby.Clear
    Do While Not rsTmp.EOF
        If NVL(rsTmp!Ӥ��, 0) = 0 Then
            cboBaby.AddItem "����ҽ��"
        Else
            cboBaby.AddItem "Ӥ�� " & rsTmp!Ӥ�� & IIF(IsNull(rsTmp!Ӥ������), " ҽ��", "��" & NVL(rsTmp!Ӥ������))
        End If
        cboBaby.ItemData(cboBaby.NewIndex) = NVL(rsTmp!Ӥ��, 0)
        rsTmp.MoveNext
    Loop
    If cboBaby.ListCount > 0 Then cboBaby.ListIndex = 0
    LoadAdviceBaby = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadAdvice() As Boolean
'���ܣ���ȡ��ǰ����ָ����ҽ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    
    If mlng����ID = 0 Then Exit Function
    If cboTime.ListIndex = -1 Then Exit Function
    If cboBaby.ListIndex = -1 Then Exit Function
    
    On Error GoTo errH
    
    If InStr(cboTime.Text, "סԺ") = 0 Then
        strSQL = "Select Distinct A.ID,A.���,A.���ID,A.ҽ����Ч,A.������ĿID,A.ҽ������," & _
            " A.��������,A.ִ��Ƶ��,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,A.ҽ������,A.ִ������,A.ִ�б��," & _
            " Nvl(C.����,Decode(Nvl(A.ִ������,0),0,'<����>',5,'-')) as ִ�п���,A.ִ�п���id,A.ִ��ʱ�䷽��," & _
            " A.ִ�п���ID,A.�걾��λ,A.��鷽��,Nvl(B.���,'*') as ���,B.����,B.���㵥λ," & _
            " A.�ܸ����� as ����,D.���㵥λ as ������λ,D.id as �շ�ϸĿID" & _
            " From ����ҽ����¼ A,������ĿĿ¼ B,���ű� C,�շ���ĿĿ¼ D,���˹Һż�¼ R" & _
            " Where A.������ĿID=B.ID(+) And A.ִ�п���ID=C.ID(+) And A.�շ�ϸĿID=D.ID(+)" & _
            " And A.����ID=[1] And Nvl(A.Ӥ��,0)=[3] And A.�Һŵ�=R.NO And R.ID=[2]" & _
            " And A.��ʼִ��ʱ�� is Not NULL And Nvl(A.ҽ��״̬,0)<>-1" & _
            " Order by A.���"
    Else
        strSQL = "Select Distinct A.ID,A.���,A.���ID,A.ҽ����Ч,A.������ĿID,A.ҽ������," & _
            " A.��������,A.ִ��Ƶ��,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,A.ҽ������,A.ִ������,A.ִ�б��," & _
            " Nvl(C.����,Decode(Nvl(A.ִ������,0),0,'<����>',5,'-')) as ִ�п���,A.ִ�п���id,A.ִ��ʱ�䷽��," & _
            " A.ִ�п���ID,A.�걾��λ,A.��鷽��,Nvl(B.���,'*') as ���,B.����,B.���㵥λ," & _
            " A.�ܸ����� as ����,D.���㵥λ as ������λ,D.id as �շ�ϸĿID" & _
            " From ����ҽ����¼ A,������ĿĿ¼ B,���ű� C,�շ���ĿĿ¼ D" & _
            " Where A.������ĿID=B.ID(+) And A.ִ�п���ID=C.ID(+) And A.�շ�ϸĿID=D.ID(+)" & _
            " And A.����ID=[1] And Nvl(A.Ӥ��,0)=[3] And A.��ҳID=[2]" & _
            " And A.��ʼִ��ʱ�� is Not NULL And Nvl(A.ҽ��״̬,0)<>-1" & _
            " Order by A.���"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, cboTime.ItemData(cboTime.ListIndex), cboBaby.ItemData(cboBaby.ListIndex))
    
    With vsAdvice
        .Redraw = flexRDNone
        .Rows = .FixedRows '����������
        If rsTmp.EOF Then
            .Rows = .FixedRows + 1
        Else
            .Rows = .FixedRows + rsTmp.RecordCount
            For i = 1 To rsTmp.RecordCount
                '.TextMatrix(i, colѡ��) = -1
                .TextMatrix(i, col���) = rsTmp!ID
                .TextMatrix(i, col���) = NVL(rsTmp!���ID)
                .TextMatrix(i, col��Ч) = IIF(NVL(rsTmp!ҽ����Ч, 0) = 0, "����", "����")
                .TextMatrix(i, col����) = rsTmp!ҽ������
                .TextMatrix(i, col�걾��λ) = NVL(rsTmp!�걾��λ) '����걾
                .TextMatrix(i, col��鷽��) = NVL(rsTmp!��鷽��)
                .TextMatrix(i, col����) = FormatEx(NVL(rsTmp!��������), 4)
                If Not IsNull(rsTmp!��������) Then
                    If rsTmp!��� = "4" Then
                        .TextMatrix(i, col��λ) = NVL(rsTmp!������λ)
                    Else
                        .TextMatrix(i, col��λ) = NVL(rsTmp!���㵥λ)
                    End If
                End If
                If .TextMatrix(i, col��Ч) = "����" Then
                    If Not IsNull(rsTmp!����) Then
                        .TextMatrix(i, col����) = FormatEx(NVL(rsTmp!����), 4)
                        If Not IsNull(rsTmp!������λ) Then
                            .TextMatrix(i, col������λ) = NVL(rsTmp!������λ)
                        ElseIf InStr(",4,5,6,7,", rsTmp!���) = 0 Then
                            .TextMatrix(i, col������λ) = NVL(rsTmp!���㵥λ)
                        End If
                    End If
                End If
                .TextMatrix(i, colƵ��) = NVL(rsTmp!ִ��Ƶ��)
                .TextMatrix(i, colƵ�ʴ���) = NVL(rsTmp!Ƶ�ʴ���)
                .TextMatrix(i, colƵ�ʼ��) = NVL(rsTmp!Ƶ�ʼ��)
                .TextMatrix(i, col�����λ) = NVL(rsTmp!�����λ)
                .TextMatrix(i, col����) = NVL(rsTmp!ҽ������)
                
                If InStr(NVL(rsTmp!ִ��ʱ�䷽��), ",") > 0 Then
                    .TextMatrix(i, colִ��ʱ��) = Split(NVL(rsTmp!ִ��ʱ�䷽��), ",")(1)
                Else
                    .TextMatrix(i, colִ��ʱ��) = NVL(rsTmp!ִ��ʱ�䷽��)
                End If
                
                .TextMatrix(i, colִ�п���) = NVL(rsTmp!ִ�п���)
                .Cell(flexcpData, i, colִ�п���) = CLng(NVL(rsTmp!ִ�п���id, 0))
                .Cell(flexcpData, i, colִ������) = Val(NVL(rsTmp!ִ������, 0))
                
                If Val(NVL(rsTmp!ִ�б��, 0)) = 1 Then
                    .TextMatrix(i, colִ������) = "��ȡҩ"
                ElseIf Val(NVL(rsTmp!ִ�б��, 0)) = 2 Then
                    .TextMatrix(i, colִ������) = "��ȡҩ"
                ElseIf Val(NVL(rsTmp!ִ������, 0)) = 5 And Val(NVL(rsTmp!ִ�б��, 0)) = 0 And Val(NVL(rsTmp!ִ�п���id, 0)) = 0 Then
                    .TextMatrix(i, colִ������) = "�Ա�ҩ"
                Else
                    .TextMatrix(i, colִ������) = "����"
                End If
                .TextMatrix(i, col��ĿID) = NVL(rsTmp!������ĿID)
                .TextMatrix(i, col���) = rsTmp!���
                .TextMatrix(i, col�շ�ϸĿID) = zlCommFun.NVL(rsTmp!�շ�ϸĿID)
                
                '���������ؼ��÷���ʾ
                If InStr(",C,D,F,G,E,", rsTmp!���) > 0 And Not IsNull(rsTmp!���ID) Then
                    .RowHidden(i) = True
                    
                    '��Ѫ;��
                    If rsTmp!��� = "E" And .TextMatrix(i - 1, col���) = "K" And Val(.TextMatrix(i - 1, col���)) = rsTmp!���ID Then
                        .TextMatrix(i - 1, col�÷�) = NVL(rsTmp!����)
                    End If
                ElseIf rsTmp!��� = "7" Then
                    .RowHidden(i) = True
                ElseIf rsTmp!��� = "E" And IsNull(rsTmp!���ID) _
                    And Val(.TextMatrix(i - 1, col���)) = rsTmp!ID _
                    And InStr(",5,6,", .TextMatrix(i - 1, col���)) > 0 Then
                    '��ҩ;��
                    .RowHidden(i) = True
                    '��ʾ��ҩ;��
                    For j = i - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, col���)) = rsTmp!ID Then
                            .TextMatrix(j, col�÷�) = NVL(rsTmp!����)
                            
                            '��ʾ��ҩ��ִ������
                            If Val(.Cell(flexcpData, j, colִ������)) <> 5 And Val(.Cell(flexcpData, i, colִ������)) = 5 Then
                                .TextMatrix(j, colִ������) = "��Ժ��ҩ"
                            End If
                        Else
                            Exit For
                        End If
                    Next
                ElseIf rsTmp!��� = "E" And IsNull(rsTmp!���ID) _
                    And Val(.TextMatrix(i - 1, col���)) = rsTmp!ID _
                    And InStr(",7,E,C,", .TextMatrix(i - 1, col���)) > 0 Then
                    '��ҩ�÷������ɼ�����
                    .TextMatrix(i, col�÷�) = NVL(rsTmp!����)
                    
                    '��ҩ������ִ�п���
                    For j = i - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, col���)) = rsTmp!ID Then
                            If InStr(",7,C,", .TextMatrix(j, col���)) > 0 Then
                                .TextMatrix(i, colִ�п���) = .TextMatrix(j, colִ�п���)
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    
                    '��ҩ����
                    If .TextMatrix(i - 1, col���) <> "C" Then
                        .TextMatrix(i, col������λ) = "��"
                        
                        '��ʾ��ҩ�䷽ִ������:��ҩƷΪ׼�ж�
                        j = .FindRow(CStr(rsTmp!ID), , col���)
                        If Val(.Cell(flexcpData, j, colִ������)) <> 5 And Val(.Cell(flexcpData, i, colִ������)) = 5 Then
                            .TextMatrix(j, colִ������) = "��Ժ��ҩ"
                        End If
                    End If
                End If
                rsTmp.MoveNext
            Next
            
            '��ǰ��ʽ�ļ��ҽ����ѡ��
            For i = .FixedRows To .Rows - 1
                If Not .RowHidden(i) And .TextMatrix(i, col���) = "D" Then
                    If CheckIsOldAdvice(i) Then
                        .TextMatrix(i, colѡ��) = 0
                        Call RowSelectSame(i)
                    End If
                End If
            Next
        End If
        .Row = .FixedRows: .Col = .FixedCols
        .AutoSize col����
        .Redraw = flexRDDirect
    End With
    LoadAdvice = True
    Exit Function
errH:
    vsAdvice.Redraw = flexRDDirect
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub vsAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = colѡ�� Then Call RowSelectSame(Row)
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Col = col���� Then
        vsAdvice.AutoSize Col
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(vsAdvice.FixedRows - 1, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
    End If
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = colѡ�� Then
        Cancel = True
    End If
End Sub

Private Sub vsAdvice_DblClick()
    Call vsAdvice_KeyPress(32)
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    Dim i As Long
    
    With vsAdvice
        If KeyAscii = 32 Then
            If .Col <> colѡ�� Then
                KeyAscii = 0
                If Val(.TextMatrix(.Row, col��ĿID)) <> 0 Then
                    .TextMatrix(.Row, colѡ��) = IIF(Val(.TextMatrix(.Row, colѡ��)) = 0, -1, 0)
                    Call RowSelectSame(.Row)
                End If
            End If
        ElseIf KeyAscii = 13 Then
            KeyAscii = 0
            For i = .Row + 1 To .Rows - 1
                If Not .RowHidden(i) Then
                    .Row = i
                    Call .ShowCell(.Row, .Col)
                    Exit For
                End If
            Next
            If i > .Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> colѡ�� Then
        Cancel = True
    Else
        '��ǰ�ļ��ҽ����������Ϊ���׷���
        If CheckIsOldAdvice(Row) Then
            MsgBox "�ü��ҽ����ϵͳ������ǰ�´�ģ������з�ʽ�����ݣ����ܱ���Ϊ���׷�����", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
    End If
End Sub

Private Function CheckIsOldAdvice(ByVal lngRow As Long) As Boolean
'���ܣ����ָ���еļ��ҽ���Ƿ��Ϸ�ʽ
'������lngRow=���ҽ���ɼ���
    Dim lngIdx As Long

    With vsAdvice
        If .TextMatrix(lngRow, col���) = "D" Then
            lngIdx = .FindRow(CStr(.TextMatrix(lngRow, col���)), lngRow + 1, col���)
            If lngIdx = -1 Then
                'CheckIsOldAdvice = True '��ǰ�ĵ���λ���
            ElseIf Val(.TextMatrix(lngIdx, col��ĿID)) <> Val(.TextMatrix(lngRow, col��ĿID)) Then
                CheckIsOldAdvice = True '��ǰ�Ķಿλ��Ŀ���
            End If
        End If
    End With
End Function

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        '����һ����ҩ������еı��߼�����
        lngLeft = col��Ч: lngRight = col��Ч
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = colƵ��: lngRight = col�÷�
            If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        End If
        
        If Not RowInһ����ҩ(Row, lngBegin, lngEnd) Then Exit Sub
        
        vRect.Left = Left '������߱����
        vRect.Right = Right - 1 '�����ұ߱����
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '���б�����������
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '���б����±���
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, OS.SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, col���) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, col���)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, col���)) = Val(.TextMatrix(lngRow, col���)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, col���)) = Val(.TextMatrix(lngRow, col���)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col���)) = Val(.TextMatrix(lngRow, col���)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col���)) = Val(.TextMatrix(lngRow, col���)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInһ����ҩ = blnTmp
    End With
End Function

Private Sub RowSelectSame(ByVal lngRow As Long)
'���ܣ�����ָ����(����Ϊ������)��ѡ��״̬,�����ҽ��һ��ѡ��
    Dim i As Long
    
    With vsAdvice
        If Val(.TextMatrix(lngRow, col���)) <> 0 Then
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col���)) = Val(.TextMatrix(lngRow, col���)) _
                    Or Val(.TextMatrix(i, col���)) = Val(.TextMatrix(lngRow, col���)) Then
                    .TextMatrix(i, colѡ��) = .TextMatrix(lngRow, colѡ��)
                Else
                    Exit For
                End If
            Next
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col���)) = Val(.TextMatrix(lngRow, col���)) _
                    Or Val(.TextMatrix(i, col���)) = Val(.TextMatrix(lngRow, col���)) Then
                    .TextMatrix(i, colѡ��) = .TextMatrix(lngRow, colѡ��)
                Else
                    Exit For
                End If
            Next
        Else
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col���)) = Val(.TextMatrix(lngRow, col���)) Then
                    .TextMatrix(i, colѡ��) = .TextMatrix(lngRow, colѡ��)
                Else
                    Exit For
                End If
            Next
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col���)) = Val(.TextMatrix(lngRow, col���)) Then
                    .TextMatrix(i, colѡ��) = .TextMatrix(lngRow, colѡ��)
                Else
                    Exit For
                End If
            Next
        End If
    End With
End Sub
