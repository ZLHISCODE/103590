VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmClientUpgradeSeverConfigure 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "�ļ�����������"
   ClientHeight    =   6255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12750
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   12750
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraOption 
      BorderStyle     =   0  'None
      Height          =   252
      Left            =   120
      TabIndex        =   11
      Top             =   5880
      Width           =   4932
      Begin VB.CheckBox chkSampleServer 
         Caption         =   "����ǰ������ļ��Ƿ���ڣ������ڼ���FTP���ߣ�"
         Height          =   180
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   4932
      End
   End
   Begin VB.PictureBox picBtn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   120
      ScaleHeight     =   330
      ScaleWidth      =   5385
      TabIndex        =   10
      Top             =   60
      Width           =   5385
      Begin VB.CommandButton cmdAdd 
         Caption         =   "����(&A)"
         Height          =   300
         Left            =   0
         TabIndex        =   0
         ToolTipText     =   "����һ���������ռ�������"
         Top             =   0
         Width           =   900
      End
      Begin VB.CommandButton cmdCheck 
         Caption         =   "�����������Լ��(&X)"
         Height          =   300
         Left            =   3240
         TabIndex        =   3
         ToolTipText     =   "����У��������Ƿ������ӳɹ�"
         Top             =   0
         Width           =   2000
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "ɾ��(&D)"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2160
         TabIndex        =   2
         ToolTipText     =   "ɾ��һ����������Ϣ"
         Top             =   0
         Width           =   900
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "�޸�(&S)"
         Enabled         =   0   'False
         Height          =   300
         Left            =   1080
         TabIndex        =   1
         ToolTipText     =   "�޸�һ����������Ϣ"
         Top             =   0
         Width           =   900
      End
   End
   Begin VB.PictureBox picFilter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   165
      ScaleHeight     =   255
      ScaleWidth      =   4005
      TabIndex        =   9
      Top             =   5610
      Width           =   4000
      Begin VB.OptionButton optFilter 
         Caption         =   "ͣ��"
         Height          =   240
         Index           =   2
         Left            =   3195
         TabIndex        =   7
         ToolTipText     =   "��ʾͣ�õķ�����"
         Top             =   0
         Width           =   720
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "����"
         Height          =   240
         Index           =   1
         Left            =   2190
         TabIndex        =   6
         ToolTipText     =   "��ʾ���õķ�����"
         Top             =   0
         Width           =   720
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "ȫ��"
         Height          =   240
         Index           =   0
         Left            =   1080
         TabIndex        =   5
         ToolTipText     =   "��ʾ���з�����"
         Top             =   0
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Label lblFilter 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������б�"
         Height          =   180
         Left            =   0
         TabIndex        =   4
         Top             =   15
         Width           =   900
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfMain 
      Height          =   4995
      Left            =   90
      TabIndex        =   8
      Top             =   465
      Width           =   12495
      _cx             =   22040
      _cy             =   8811
      Appearance      =   0
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
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   7000
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmClientUpgradeSeverConfigure.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
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
End
Attribute VB_Name = "frmClientUpgradeSeverConfigure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'=================================================================
'ģ�����
'=================================================================
Private Enum ServerListCols
    Col_��� = 0
    Col_���� = 1
    Col_������״̬ = 2 '���� or ͣ��
    Col_������·�� = 3
    Col_�û��� = 4
    Col_���� = 5
    Col_�˿� = 6
    Col_�Ƿ����� = 7
    Col_�Ƿ�ȱʡ = 8
    Col_�Ƿ��ռ� = 9
    Col_�ռ����� = 10
    Col_����� = 11
End Enum
Private mblnHaveDefault As Boolean '�Ƿ����Ĭ�Ϸ�����
Private mblnAllowEdit As Boolean '��ǵ�ǰ�����Ƿ�����༭
'=================================================================
'�����ӿ�
'=================================================================
Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�
End Sub

Public Sub SetMenu()
'���ܣ�����״̬������
    frmMDIMain.stbThis.Panels(2).Text = "�б��й���ʾ��" & vsfMain.Rows - 1 & "�����ݡ�"
End Sub
'
Public Sub RefreshData()
'���ܣ���������õ�ˢ�����ݽӿ�
    Call LoadSeverListData
End Sub

'=================================================================
'˽�з���
'=================================================================
Private Sub chkSampleServer_Click()
    If chkSampleServer.Tag <> "" Then
        Call gclsBase.UpdateZLReginfo("FTP������ļ�����", chkSampleServer.value)
    End If
End Sub

Private Sub cmdAdd_Click()
    Dim frmEdit As New frmClientUpgradeSeverEdit
    If frmEdit.ShowMe(0, mblnHaveDefault) Then
        Call LoadSeverListData
    End If
End Sub

Private Sub cmdCheck_Click()
    Dim i As Long, objConn As clsConnect
    Dim strErr As String
    
    With vsfMain
        If .Rows < .FixedRows Then Exit Sub
        For i = .FixedRows To .Rows - 1
            ShowFlash "���ڼ��" & .TextMatrix(i, Col_���) & "��: " & .TextMatrix(i, Col_������·��), (i - 1) / (.Rows - 1), Me, True
            DoEvents
            Set objConn = New clsConnect
            strErr = ""
            If Not objConn.ToConnect(IIf(Trim(.TextMatrix(i, Col_����)) = "FTP", SCT_FTP, SCT_Share), .TextMatrix(i, Col_������·��), .TextMatrix(i, Col_�û���), .Cell(flexcpData, i, Col_����), Val(.TextMatrix(i, Col_�˿�)), "", False, strErr) Then
                .TextMatrix(i, Col_�����) = "�����ã�" & strErr
            Else
                .TextMatrix(i, Col_�����) = "����"
            End If
            ShowFlash "���ڼ��" & .TextMatrix(i, Col_���) & "��: " & .TextMatrix(i, Col_������·��), i / (.Rows - 1), Me, True
            Call objConn.CloseConnect
        Next
        Call ShowFlash("")
    End With
End Sub

Private Sub cmdDel_Click()
    Dim strSQL As String
    Dim strRemarks As String
    
    If vsfMain.TextMatrix(vsfMain.Row, Col_�Ƿ�ȱʡ) <> "" Then
        MsgBox vsfMain.TextMatrix(vsfMain.Row, Col_���) & " ��" & vsfMain.TextMatrix(vsfMain.Row, Col_����) & "������Ϊȱʡ����������ɾ�������л�ȱʡ��������ɾ����", vbInformation, gstrSysName
        Exit Sub
    End If
    If MsgBox("ȷ��Ҫɾ�� " & vsfMain.TextMatrix(vsfMain.Row, Col_���) & " ��" & vsfMain.TextMatrix(vsfMain.Row, Col_����) & "��������", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
        '��֤��ݲ��������˵��
        If Not CheckAuditStatus("0307", "�ļ�����������-ɾ��", strRemarks) Then Exit Sub
        strSQL = "Zl_Zlupgradeserver_Update(2," & vsfMain.TextMatrix(vsfMain.Row, Col_���) & ")"
        Call ExecuteProcedure(strSQL, Me.Caption, gcnOracle)
        '������Ҫ������־
        Call SaveAuditLog(3, "�ļ�����������-ɾ��", "ɾ�����Ϊ" & vsfMain.TextMatrix(vsfMain.Row, Col_���) & "���ļ�������", strRemarks)
        Call LoadSeverListData
    End If
End Sub

Private Sub cmdModify_Click()
    Dim frmEdit As New frmClientUpgradeSeverEdit
    If frmEdit.ShowMe(Val(vsfMain.TextMatrix(vsfMain.Row, Col_���)), mblnHaveDefault) Then
        Call LoadSeverListData
    End If
End Sub

Private Sub Form_Load()
    mblnAllowEdit = True
    Call TransOldData
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fraOption.Top = Me.ScaleHeight - fraOption.Height - 60
    picFilter.Top = fraOption.Top - picFilter.Height - 90
    vsfMain.Height = picFilter.Top - 90 - vsfMain.Top
    vsfMain.Width = Me.ScaleWidth - 120
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub optFilter_Click(Index As Integer)
    Dim i As Integer
    
    With vsfMain
        If .Rows < .FixedRows Then Exit Sub
        For i = 1 To .Rows - 1
            .RowHidden(i) = Not ((Index = .Cell(flexcpData, i, Col_������״̬)) Or (Index = 0))
        Next
    End With
End Sub

Private Sub vsfMain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnAllowEdit = False Then Exit Sub
    cmdModify.Enabled = NewRow >= vsfMain.FixedRows
    cmdDel.Enabled = NewRow >= vsfMain.FixedRows
    cmdCheck.Enabled = NewRow >= vsfMain.FixedRows
End Sub

Private Sub vsfMain_DblClick()
    Dim intUpdate       As Integer, intDefault As Integer, intCollect As Integer
    Dim strFilesType    As String, strSQL      As String
    
    If mblnAllowEdit = False Then Exit Sub
    With vsfMain
        If .MouseRow <> .Row Then Exit Sub
        intUpdate = IIf(.TextMatrix(.Row, Col_�Ƿ�����) = "��", 1, 0)
        intDefault = IIf(.TextMatrix(.Row, Col_�Ƿ�ȱʡ) = "��", 1, 0)
        intCollect = IIf(.TextMatrix(.Row, Col_�Ƿ��ռ�) = "��", 1, 0)
        strFilesType = .TextMatrix(.Row, Col_�ռ�����)
        If intDefault = 1 And (.ColSel = Col_�Ƿ����� Or .ColSel = Col_�Ƿ�ȱʡ Or .ColSel = Col_�Ƿ��ռ�) Then
            Call MsgBox("ѡ�б�� " & .TextMatrix(.Row, Col_���) & " ������Ϊȱʡ�������������Ƚ������������л�Ϊȱʡ���뱣֤��һ��ȱʡ��������", vbInformation, gstrSysName)
            Exit Sub
        ElseIf Not (.ColSel = Col_�Ƿ����� Or .ColSel = Col_�Ƿ�ȱʡ Or .ColSel = Col_�Ƿ��ռ�) Then
            Exit Sub
        End If
        On Error GoTo ErrH
        Select Case .ColSel
            Case Col_�Ƿ�����
                If intCollect = 1 Then
                    If MsgBox("ѡ�б�� " & .TextMatrix(.Row, Col_���) & " ������Ϊ�ռ����������Ƿ�Ҫ�л�Ϊ������������ ", vbInformation + vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                        Exit Sub
                    End If
                ElseIf intUpdate = 1 Then
                    If MsgBox("�Ƿ�Ҫȡ����������������ȡ���󽫻���������ù��÷�����Ϊ�����������Ŀͻ���", vbInformation + vbOKCancel, gstrSysName) = vbCancel Then
                        Exit Sub
                    End If
                End If
                strFilesType = ""
                intUpdate = IIf(intUpdate = 1, 0, 1)
                intCollect = 0
                '������Ϊ��������������û��ȱʡ�����������Զ�ȱʡ
                If intUpdate = 1 And Not mblnHaveDefault Then intDefault = 1
            Case Col_�Ƿ�ȱʡ
                If intCollect = 1 Then
                    If MsgBox("ѡ�б�� " & .TextMatrix(.Row, Col_���) & " ������Ϊ�ռ����������Ƿ�Ҫ�л�Ϊ����������������Ϊȱʡ�������� ", vbInformation + vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                        Exit Sub
                    End If
                ElseIf intUpdate = 0 Then
                    If MsgBox("ѡ�б�� " & .TextMatrix(.Row, Col_���) & " ������Ϊͣ��״̬���Ƿ�Ҫ���ø÷�����������Ϊȱʡ�������� ", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                        Exit Sub
                    End If
                End If
                '�û�������������ȱʡ
                strFilesType = ""
                intUpdate = 1
                intCollect = 0
                intDefault = 1
            Case Col_�Ƿ��ռ�
                If intUpdate = 0 Then
                    If MsgBox("ѡ�б�� " & .TextMatrix(.Row, Col_���) & " ������Ϊ�������������Ƿ�Ҫ�л�Ϊ�ռ��������� ", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                        Exit Sub
                    End If
                End If
                strFilesType = ""
                intUpdate = 0
                intCollect = IIf(intCollect = 1, 0, 1)
                intDefault = 0
        End Select
        strSQL = "Zl_Zlupgradeserver_Update(11," & .TextMatrix(.Row, Col_���) & "," & IIf(.TextMatrix(.Row, Col_����) = "����", 0, 1) & ",'" & Trim(.TextMatrix(.Row, Col_������·��)) & "','" & Trim(.TextMatrix(.Row, Col_�û���)) & "'," & SQLAdjust(Cipher(Trim(.Cell(flexcpData, .Row, Col_����)))) & ",'" & Trim(.TextMatrix(.Row, Col_�˿�)) & "'," & intUpdate & "," & intDefault & "," & intCollect & "," & SQLAdjust(strFilesType) & "," & SQLAdjust(Trim(.Cell(flexcpData, .Row, Col_����))) & ")"
        Call ExecuteProcedure(strSQL, Me.Caption, gcnOracle)
        Call LoadSeverListData
        optFilter.Item(0).value = True
    End With
    Exit Sub
ErrH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

Private Sub vsfMain_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Row < vsfMain.FixedRows
End Sub

Public Sub LoadSeverListData()
'���ܣ����������ļ��������嵥
    Dim lngRow  As Long
    Dim strSQL  As String, rsTmp As ADODB.Recordset

    On Error GoTo ErrH
    mblnHaveDefault = False
    '����ʹ�ü���FTP����
    strSQL = "Select ���� As ʹ�ü���ftp���� From Zlreginfo Where ��Ŀ =[1]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, "FTP������ļ�����")
    If rsTmp.EOF Then
        chkSampleServer.value = 0
        Call gclsBase.UpdateZLReginfo("FTP������ļ�����", 0, 1)
    Else
        chkSampleServer.value = Val(rsTmp!ʹ�ü���ftp���� & "")
    End If
    chkSampleServer.Tag = "�����Ѿ�����"
    strSQL = "Select ���, ����, λ��, �û���, ����, �˿�, �Ƿ�����, �Ƿ�ȱʡ From ZLTOOLS.Zlupgradeserver Order By ���"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    With vsfMain
        .Rows = .FixedRows
        Do While Not rsTmp.EOF
            .Rows = .Rows + 1
            lngRow = .Rows - 1
            .TextMatrix(lngRow, Col_���) = rsTmp!��� & ""
            .TextMatrix(lngRow, Col_����) = IIf(Val(rsTmp!���� & "") = 1, "FTP", "����")
            .TextMatrix(lngRow, Col_������·��) = rsTmp!λ�� & ""
            .TextMatrix(lngRow, Col_�û���) = rsTmp!�û��� & ""
            .TextMatrix(lngRow, Col_����) = "***"
            .Cell(flexcpData, lngRow, Col_����) = Decipher(rsTmp!���� & "")
            .TextMatrix(lngRow, Col_�˿�) = rsTmp!�˿� & ""
            .Cell(flexcpBackColor, lngRow, Col_�Ƿ�����, lngRow, Col_�Ƿ��ռ�) = RGB(210, 240, 255)
            .TextMatrix(lngRow, Col_�Ƿ�����) = IIf(Val(rsTmp!�Ƿ����� & "") = 1, "��", "")
            .TextMatrix(lngRow, Col_�Ƿ�ȱʡ) = IIf(Val(rsTmp!�Ƿ�ȱʡ & "") = 1, "��", "")
            If .TextMatrix(lngRow, Col_�Ƿ�ȱʡ) = "��" Then
                mblnHaveDefault = True
                .Cell(flexcpForeColor, lngRow, .FixedCols, lngRow, .Cols - 1) = vbBlue
            End If
            .TextMatrix(lngRow, Col_�����) = ""
            If .TextMatrix(lngRow, Col_�Ƿ�����) = "" And .TextMatrix(lngRow, Col_�Ƿ�ȱʡ) = "" And .TextMatrix(lngRow, Col_�Ƿ��ռ�) = "" Then
                .TextMatrix(lngRow, Col_������״̬) = "ͣ��"
                .Cell(flexcpForeColor, lngRow, .FixedCols, lngRow, .Cols - 1) = vbGrayText
                .Cell(flexcpData, lngRow, Col_������״̬) = 2
            Else
                .TextMatrix(lngRow, Col_������״̬) = "����"
                .Cell(flexcpData, lngRow, Col_������״̬) = 1
            End If
            rsTmp.MoveNext
        Loop
        If lngRow > .FixedRows Then
            .Row = .FixedRows
        End If
        Call SetMenu
    End With
    Exit Sub
ErrH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox "�������б���ش���,��Ϣ:" & err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

Private Function TransOldData() As Boolean
'���ܣ��������÷����µĴ洢������
    Dim strSQL          As String, rsTmp        As ADODB.Recordset, rsNum As ADODB.Recordset
    Dim intClientUpType As Integer, strFileType As String
    Dim lngServerNO     As Integer, strTmp      As String
    Dim strUser         As String, strPwd       As String, strPort  As String, strPath  As String
    Dim blnSetDefault   As Boolean
    
    On Error GoTo ErrH
    '���ж���������
    strSQL = "Select 1 From Zlupgradeserver Where Rownum < 2"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    If Not rsTmp.EOF Then TransOldData = True: Exit Function
    '��ȡĬ����������
    strSQL = "Select Max(����) As �������� From Zlreginfo Where ��Ŀ = [1]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, "��������")
    intClientUpType = Val(rsTmp!�������� & "")
    '��ת��FTP�������͵ķ�����
    strSQL = "Select ��Ŀ, ����" & vbNewLine & _
            "From Zlreginfo" & vbNewLine & _
            "Where (��Ŀ Like 'FTP������%' Or ��Ŀ Like 'FTP�û�%' Or ��Ŀ Like 'FTP����%' Or ��Ŀ Like 'FTP�˿�%')" & vbNewLine & _
            "And ���� Is Not Null"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    '1.�ȴ���FTP������
    rsTmp.Filter = "��Ŀ Like 'FTP������*'"
    Set rsNum = CopyNewRec(rsTmp)
    rsNum.Sort = "��Ŀ"
    Do While Not rsNum.EOF
        strTmp = Mid(rsNum!��Ŀ, Len("FTP������") + 1)
        strUser = "": strPwd = "": strPort = "": strPath = ""
        strPath = rsNum!���� & ""
        rsTmp.Filter = "��Ŀ='FTP�û�" & strTmp & "'"
        If Not rsTmp.EOF Then strUser = rsTmp!���� & ""
        rsTmp.Filter = "��Ŀ='FTP����" & strTmp & "'"
        If Not rsTmp.EOF Then strPwd = rsTmp!���� & ""
        rsTmp.Filter = "��Ŀ='FTP�˿�" & strTmp & "'"
        If Not rsTmp.EOF Then strPort = rsTmp!���� & ""
        lngServerNO = lngServerNO + 1
        strSQL = "Zl_Zlupgradeserver_Update(0," & lngServerNO & ",1," & SQLAdjust(strPath) & "," & SQLAdjust(strUser) & "," & SQLAdjust(Cipher(strPwd)) & "," & Val(strPort) & ",1," & IIf(intClientUpType = 1 And Not blnSetDefault, 1, 0) & ",0,NULL," & SQLAdjust(strPwd) & ")"
        Call ExecuteProcedure(strSQL, Me.Caption, gcnOracle)
        If Not blnSetDefault Then
            blnSetDefault = intClientUpType = 1
        End If
        If intClientUpType = 1 Then
            strSQL = "Update Zltools.Zlclients Set �����ļ������� = " & lngServerNO & " Where Ftp������ " & IIf(strTmp = "", "Is Null", "=" & strTmp)
            Call gclsBase.ExecuteCmdText(strSQL, Me.Caption, gcnOracle)
        End If
        rsNum.MoveNext
    Loop
    '��ת�ƹ����������͵ķ�����
    strSQL = "Select ��Ŀ, ����" & vbNewLine & _
            "From Zlreginfo" & vbNewLine & _
            "Where (��Ŀ Like '������Ŀ¼%' Or ��Ŀ Like '�����û�%' Or ��Ŀ Like '��������%')" & vbNewLine & _
            "And ���� Is Not Null"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)

    rsTmp.Filter = "��Ŀ Like '������Ŀ¼*'"
    Set rsNum = CopyNewRec(rsTmp)
    rsNum.Sort = "��Ŀ"
    Do While Not rsNum.EOF
        strTmp = Mid(rsNum!��Ŀ, Len("������Ŀ¼") + 1)
        strUser = "": strPwd = "": strPath = ""
        strPath = rsNum!���� & ""
        rsTmp.Filter = "��Ŀ='�����û�" & strTmp & "'"
        If Not rsTmp.EOF Then strUser = rsTmp!���� & ""
        rsTmp.Filter = "��Ŀ='��������" & strTmp & "'"
        If Not rsTmp.EOF Then strPwd = rsTmp!���� & ""
        lngServerNO = lngServerNO + 1
        strSQL = "Zl_Zlupgradeserver_Update(0," & lngServerNO & ",0," & SQLAdjust(strPath) & "," & SQLAdjust(strUser) & "," & SQLAdjust(Cipher(strPwd)) & ",NULL,1," & IIf(intClientUpType = 0 And Not blnSetDefault, 1, 0) & ",0,NULL," & SQLAdjust(strPwd) & ")"
        Call ExecuteProcedure(strSQL, Me.Caption, gcnOracle)
        If Not blnSetDefault Then
            blnSetDefault = intClientUpType = 0
        End If
        If intClientUpType = 0 Then
            strSQL = "Update Zltools.Zlclients Set �����ļ������� = " & lngServerNO & " Where ���������� " & IIf(strTmp = "", "Is Null", "=" & strTmp)
            Call gclsBase.ExecuteCmdText(strSQL, Me.Caption, gcnOracle)
        End If
        rsNum.MoveNext
    Loop
    If lngServerNO > 0 Then
        '��տͻ����������õ�����������
        strSQL = "Update Zlclients Set ���������� = Null, Ftp������ = Null"
        Call gclsBase.ExecuteCmdText(strSQL, Me.Caption, gcnOracle)
    End If
    '��ȡ�ռ��������ռ���ʽ
    strSQL = "Select Max(����) As �ռ���ʽ From Zlreginfo Where ��Ŀ = [1]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, "�ռ���ʽ")
    intClientUpType = Val(rsTmp!�ռ���ʽ & "")
    strSQL = "Select Max(����) As �ռ����� From Zlreginfo Where ��Ŀ = [1]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, "�ռ�����")
    strFileType = rsTmp!�ռ����� & ""
    '����FTP�ռ�������
    strSQL = "Select ��Ŀ, ����" & vbNewLine & _
        "From Zlreginfo" & vbNewLine & _
        "Where ��Ŀ In ('�ռ�Ŀ¼S', '�����û�S', '��������S', '�ռ�Ŀ¼F', '�����û�F', '��������F', '���ʶ˿�F')" & vbNewLine & _
        "And ���� Is Not Null"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    rsTmp.Filter = "��Ŀ = '�ռ�Ŀ¼F'"
    If Not rsTmp.EOF Then
        strUser = "": strPwd = "": strPort = "": strPath = ""
        strPath = rsNum!���� & ""
        rsTmp.Filter = "��Ŀ='�����û�F'"
        If Not rsTmp.EOF Then strUser = rsTmp!���� & ""
        rsTmp.Filter = "��Ŀ='��������F'"
        If Not rsTmp.EOF Then strPwd = rsTmp!���� & ""
        rsTmp.Filter = "��Ŀ='���ʶ˿�F'"
        If Not rsTmp.EOF Then strPort = rsTmp!���� & ""
        lngServerNO = lngServerNO + 1
        strSQL = "Zl_Zlupgradeserver_Update(0," & lngServerNO & ",1," & SQLAdjust(strPath) & "," & SQLAdjust(strUser) & "," & SQLAdjust(Cipher(strPwd)) & "," & Val(strPort) & ",0,0,1," & IIf(intClientUpType = 1, SQLAdjust(strFileType), "NULL") & "," & SQLAdjust(strPwd) & ")"
        Call ExecuteProcedure(strSQL, Me.Caption, gcnOracle)
    End If
    '�������ռ�
    rsTmp.Filter = "��Ŀ = '�ռ�Ŀ¼F'"
    If Not rsTmp.EOF Then
        strTmp = Mid(rsNum!��Ŀ, Len("������Ŀ¼") + 1)
        strUser = "": strPwd = "": strPath = ""
        strPath = rsNum!���� & ""
        rsTmp.Filter = "��Ŀ='�����û�" & strTmp & "'"
        If Not rsTmp.EOF Then strUser = rsTmp!���� & ""
        rsTmp.Filter = "��Ŀ='��������" & strTmp & "'"
        If Not rsTmp.EOF Then strPwd = rsTmp!���� & ""
        lngServerNO = lngServerNO + 1
        strSQL = "Zl_Zlupgradeserver_Update(0," & lngServerNO & ",0," & SQLAdjust(strPath) & "," & SQLAdjust(strUser) & "," & SQLAdjust(Cipher(strPwd)) & ",NULL,0,0,1," & IIf(intClientUpType = 0, SQLAdjust(strFileType), "NULL") & "," & SQLAdjust(strPwd) & ")"
        Call ExecuteProcedure(strSQL, Me.Caption, gcnOracle)
    End If
    If lngServerNO > 0 Then
        '���������
        '1-����FTP����
        strSQL = "Delete From Zlreginfo" & vbNewLine & _
            "Where (��Ŀ Like 'FTP������%' And ��Ŀ > 'FTP������0')" & vbNewLine & _
            "Or (��Ŀ Like 'FTP�û�%' And ��Ŀ > 'FTP�û�0')" & vbNewLine & _
            "Or (��Ŀ Like 'FTP����%' And ��Ŀ > 'FTP����0')" & vbNewLine & _
            "Or (��Ŀ Like 'FTP�˿�%' And ��Ŀ > 'FTP�˿�0')"
        Call gclsBase.ExecuteCmdText(strSQL, Me.Caption, gcnOracle)
        '2-����������
        strSQL = "Delete From Zlreginfo" & vbNewLine & _
            "Where (��Ŀ Like '������Ŀ¼%' And ��Ŀ > '������Ŀ¼0')" & vbNewLine & _
            "Or (��Ŀ Like '�����û�%' And ��Ŀ > '�����û�0')" & vbNewLine & _
            "Or (��Ŀ Like '��������%' And ��Ŀ > '��������0')"
        Call gclsBase.ExecuteCmdText(strSQL, Me.Caption, gcnOracle)
        '3-�����ռ�����
        strSQL = "Delete From Zlreginfo" & vbNewLine & _
                    "Where ��Ŀ In ('�ռ���ʽ'," & vbNewLine & _
                    "             '�ռ�����'," & vbNewLine & _
                    "             '�ռ�Ŀ¼S'," & vbNewLine & _
                    "             '�����û�S'," & vbNewLine & _
                    "             '��������S'," & vbNewLine & _
                    "             '�ռ�Ŀ¼F'," & vbNewLine & _
                    "             '�����û�F'," & vbNewLine & _
                    "             '��������F'," & vbNewLine & _
                    "             '���ʶ˿�F')"
        Call gclsBase.ExecuteCmdText(strSQL, Me.Caption, gcnOracle)
        '4-����ZLClients��������
        strSQL = "Update Zlclients Set ���������� = Null, Ftp������ = Null Where �����ļ������� Is Null"
        Call gclsBase.ExecuteCmdText(strSQL, Me.Caption, gcnOracle)
    End If
    TransOldData = True
    Exit Function
ErrH:
    TransOldData = False
    If 0 = 1 Then
        Resume
    End If
    MsgBox "�ɰ汾����������ת��ʧ��, ����ϵ������Ա!��Ϣ��" & err.Description, vbInformation, gstrSysName
    err.Clear
End Function

Public Sub SetControlEnable(ByVal strProgFunc As String)
'����Ȩ���ַ������ÿؼ�״̬
'strProgFunc:Ȩ���ַ���
    Dim arrFunc() As String
    Dim i As Long
    
    mblnAllowEdit = False
    arrFunc = Split(strProgFunc, "|")
    For i = 0 To UBound(arrFunc)
        If arrFunc(i) = "�ļ�����������" Then
            mblnAllowEdit = True
        End If
    Next
    '��û��Ȩ�ޣ���һЩ�ؼ���Ϊ������
    If mblnAllowEdit = False Then
        cmdAdd.Enabled = False
        cmdModify.Enabled = False
        cmdDel.Enabled = False
        chkSampleServer.Enabled = False
        vsfMain.Editable = flexEDNone
    End If
End Sub
