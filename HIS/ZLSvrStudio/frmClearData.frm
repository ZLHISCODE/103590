VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClearData 
   BackColor       =   &H80000005&
   Caption         =   "�������"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmClearData.frx":0000
   ScaleHeight     =   6180
   ScaleWidth      =   7500
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClear 
      Caption         =   "ȫ��(&L)"
      Height          =   350
      Left            =   3060
      TabIndex        =   10
      Top             =   4770
      Width           =   1100
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "ȫѡ(&C)"
      Height          =   350
      Left            =   1890
      TabIndex        =   9
      Top             =   4770
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwTable 
      Height          =   3135
      Left            =   1920
      TabIndex        =   8
      Top             =   1500
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "���ձ�"
         Object.Width           =   8819
      EndProperty
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "ִ��(&E)��"
      Height          =   350
      Left            =   5640
      TabIndex        =   6
      Top             =   4770
      Width           =   1100
   End
   Begin VB.TextBox txtFile 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   1920
      MaxLength       =   256
      TabIndex        =   3
      Top             =   1065
      Width           =   4485
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "��"
      Height          =   300
      Left            =   6420
      TabIndex        =   2
      Top             =   1050
      Width           =   300
   End
   Begin VB.ComboBox cmbSystem 
      Height          =   300
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   690
      Width           =   4815
   End
   Begin MSComDlg.CommonDialog cmmFile 
      Left            =   5070
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���������(&T)"
      Height          =   180
      Index           =   2
      Left            =   690
      TabIndex        =   7
      Top             =   1620
      Width           =   1170
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��װ�ļ�(&F)"
      Height          =   180
      Index           =   0
      Left            =   870
      TabIndex        =   5
      Top             =   1110
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ӧ��ϵͳ(&S)"
      Height          =   180
      Index           =   1
      Left            =   870
      TabIndex        =   4
      Top             =   750
      Width           =   990
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmClearData.frx":04F9
      Top             =   690
      Width           =   480
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   0
      Top             =   105
      Width           =   960
   End
End
Attribute VB_Name = "frmClearData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mrsSystem As New ADODB.Recordset
Dim mrsConstraint As New ADODB.Recordset '�����ŵ�ǰϵͳ�������Լ��
Dim mstr������ As String '���浱ǰϵͳ����������
Dim mcolTable As New Collection '��ǰϵͳ�̶��ı�

Private Sub cmdClear_Click()
    Dim i As Integer
    For i = 1 To lvwTable.ListItems.Count
        lvwTable.ListItems(i).Checked = False
    Next
End Sub

Private Sub cmdExecute_Click()
    Dim colTemp As New Collection    'Ҫɾ���ı��¼
    Dim strErr As String             '��¼�쳣�ı�
    Dim strDelete As String          '��¼�Ѿ�ɾ���ı�
    Dim strLoop As String            '��һ��ѭ��������ɾ���ı�
    Dim lngCount As Long
    Dim lst As ListItem
    Dim strTable As String
    Dim blnDelete As Boolean
    Dim strRemarks As String
    Dim strNote As String
    
    '�õ�Ҫɾ���ı�
    For Each lst In lvwTable.ListItems
        If lst.Checked = True Then
            colTemp.Add lst.Text, lst.Key
        End If
    Next
    If colTemp.Count = 0 Then Exit Sub
    If MsgBox("�������Ƿǳ�Σ�յģ���ȷ���Ѿ���ȷѡ����Ӧ��ɾ���ı���", _
        vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    '��֤��ݲ��������˵��
    If Not CheckAuditStatus("0206", "ִ��", strRemarks) Then Exit Sub
    frmMDIMain.Enabled = False
    Enabled = False
    frmWait.BeginWait "����ɾ����ѡ������ݡ���"
    On Error Resume Next
    lngCount = 1
    Do While lngCount <= colTemp.Count
        '�õ��ñ��Ӧ���б��¼
        strTable = colTemp(lngCount)           '�����ñ������棬Ч�ʸ���
        
        blnDelete = CanDelete(strTable, strDelete) '��ÿһ����ȱʡ��Ϊ���ǲ���ɾ����
        If blnDelete = True Then
            '�ñ�û��ʲô���յģ�����ֱ��ɾ��
            frmMDIMain.stbThis.Panels(2).Text = "����ɾ��" & strTable & "�����ݡ���"
            gcnOracle.Execute "delete from " & mstr������ & "." & strTable
            strNote = strNote & "," & strTable
            If Err <> 0 Then
                Debug.Print Err.Description
                Err.Clear
                strErr = strErr & strTable & vbCrLf
            End If
            '���ܳ��ִ�����񣬶��Ѹñ�Ӽ�����ɾ����������ѭ��
            colTemp.Remove lngCount
            strDelete = strDelete & "[" & strTable & "]"
        Else
            '�ж���һ��
            lngCount = lngCount + 1
        End If
        '����Ѿ���ĩβ�ˣ�����ͷ��ʼ
        If lngCount > colTemp.Count Then
            If strDelete = strLoop Then
                'ѭ����һȦ��һ����ûɾ����˵������������
                If Right(strDelete, 1) = "," Then strDelete = Mid(strDelete, 1, Len(strDelete) - 1) 'ȥ�����Ķ���
                MsgBox "������ݱ����������ֱ���ɾ����" & IIf(strDelete = "", "", "�������б��ѱ�ɾ����" & vbCrLf & vbCrLf & strDelete), vbExclamation, gstrSysName
                frmMDIMain.stbThis.Panels(2).Text = ""
                frmWait.EndWait
                Enabled = True
                frmMDIMain.Enabled = True
                '������Ҫ������־
                If strNote <> "" Then
                    Call SaveAuditLog(3, "ִ��", "�ɹ�����" & Split(cmbSystem.Text, " ")(0) & "���е����ݱ�" & Mid(strNote, 2) & "�����", strRemarks)
                End If
                Exit Sub
            End If
            lngCount = 1
            strLoop = strDelete
        End If
        
    Loop
    frmMDIMain.stbThis.Panels(2).Text = ""
    frmWait.EndWait
    Enabled = True
    frmMDIMain.Enabled = True
    If Right(strErr, 1) = "," Then strErr = Mid(strErr, 1, Len(strErr) - 1) 'ȥ�����Ķ���
    MsgBox "����ɾ������ִ����ϡ�" & IIf(strErr = "", "", "�������б�δ����ɾ����" & vbCrLf & vbCrLf & strErr), vbExclamation, gstrSysName
    '������Ҫ������־
    If strNote <> "" Then
        Call SaveAuditLog(3, "ִ��", "�ɹ�����" & Split(cmbSystem.Text, " ")(0) & "���е����ݱ�" & Mid(strNote, 2) & "�����", strRemarks)
    End If
End Sub

Private Function CanDelete(ByVal strTable As String, strDelete As String) As Boolean
    Dim lst As ListItem
    Dim strTemp As String
    Dim varRefTable As Variant
    Dim i As Long
    
    CanDelete = False
    Set lst = lvwTable.ListItems("C" & strTable)
    If lst.SubItems(1) = "" Then
        CanDelete = True
    Else
        varRefTable = Split(lst.SubItems(1), ",")
        For i = LBound(varRefTable) To UBound(varRefTable)
            If varRefTable(i) <> strTable And varRefTable(i) <> strTable & "(*)" Then
                '�Լ������ж�
                If InStr(varRefTable(i), "(*)") = 0 Then
                    '�õ�һ����Լ���Ĳ��ձ��ж����Ƿ��Ѿ�ɾ��
                    If InStr(strDelete, "[" & varRefTable(i) & "]") = 0 Then
                        '��һ���ձ�ûɾ�����Ͳ���ɾ������
                        Exit For
                    End If
                Else
                    strTemp = Mid(varRefTable(i), 1, InStr(varRefTable(i), "(*)") - 1)
                    If CanDelete(strTemp, strDelete) = False Then
                        '��һ���ձ�ûɾ�����Ͳ���ɾ������
                        Exit For
                    End If
                End If
            End If
        Next
        If i > UBound(varRefTable) Then CanDelete = True
    End If
    
End Function

Private Sub cmdFile_Click()
    Dim lst As ListItem
    Dim strTemp As String

    cmmFile.Filter = "Ӧ�ð�װ�����ļ�|zlSetup.ini"
    cmmFile.FileName = txtFile.Text
    cmmFile.ShowOpen
    If cmmFile.FileName = "" Then Exit Sub
    '������ǰѡ�����ļ�
    If txtFile.Text = cmmFile.FileName Then Exit Sub
    '���½��м��
    txtFile.Text = cmmFile.FileName
    If CheckIniFile(txtFile.Text, False) = False Then
        cmdExecute.Enabled = False
        lvwTable.ListItems.Clear
    Else
        cmdExecute.Enabled = True
        Call FillTable
    End If
End Sub

Private Sub cmbSystem_Click()
    Dim rsTemp As New ADODB.Recordset
On Error GoTo ErrHandle
    
    '���ϴ���ͬ������Ҫ����
    If mrsSystem.Filter = "���=" & cmbSystem.ItemData(cmbSystem.ListIndex) Then Exit Sub
    
    MousePointer = 11
    mrsSystem.Filter = "���=" & cmbSystem.ItemData(cmbSystem.ListIndex)
    If mrsSystem.RecordCount = 0 Then
        cmdExecute.Enabled = False
        txtFile.Text = ""
    Else
        cmdExecute.Enabled = True
        mstr������ = mrsSystem("������")
        '����ϵͳ��װ�ű���λ��
        Dim varOut As Variant
        Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Loadandunload.Get_Sysfile_name", Val(cmbSystem.ItemData(cmbSystem.ListIndex)), 1)
        'txtFile.Text = IIf(IsNull(varOut(0)), "", varOut(0))
        If rsTemp.RecordCount <= 0 Then
            txtFile.Text = ""
        Else
            txtFile.Text = IIf(IsNull(rsTemp("�ļ���")), "", rsTemp("�ļ���"))
        End If
        rsTemp.Close
        
        '������ǰϵͳ������Լ��
        If mrsConstraint.State = 1 Then mrsConstraint.Close
        gstrSQL = "select A.table_name ,B.table_name r_table_name,A.DELETE_RULE" & _
                   " from all_constraints A,all_constraints b" & _
                   " where A.owner='" & mstr������ & "' AND b.OWNER='" & mstr������ & "' and A.r_owner=B.owner" & _
                   "     and A.R_CONSTRAINT_NAME=b.constraint_name And Instr(A.Table_NAME,'BIN$')<=0"
        mrsConstraint.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    End If
    '���ļ��ļ��
    If CheckIniFile(txtFile.Text, True) = False Then
        cmdExecute.Enabled = False
        lvwTable.ListItems.Clear
    Else
        cmdExecute.Enabled = True
        Call FillTable
    End If
    MousePointer = 0
    Exit Sub
ErrHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub FillTable()
    Dim rsTemp As New ADODB.Recordset
    Dim strTable As String
    Dim strTemp As String
    Dim lst As ListItem
On Error GoTo ErrHandle
    
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open "select table_name from all_tables A where owner='" & mstr������ & "' And Instr(A.Table_NAME,'BIN$')<=0 order by table_name", gcnOracle, adOpenStatic, adLockReadOnly
    
    lvwTable.ListItems.Clear
    If rsTemp.RecordCount = 0 Then
        cmdExecute.Enabled = False
        Exit Sub
    End If
    'װ����
    On Error Resume Next
    Do Until rsTemp.EOF
        strTable = rsTemp("TABLE_NAME")
        '�õ����Ƿ��ǻ�����
        strTemp = mcolTable("C" & strTable)
        If Err <> 0 Then
            '���ǣ����Լ���
            Err.Clear
            Set lst = lvwTable.ListItems.Add(, "C" & strTable, strTable)
            '�õ����������¼���
            mrsConstraint.Filter = "r_table_name='" & strTable & "'"
            strTemp = ""
            Do Until mrsConstraint.EOF
                strTemp = strTemp & mrsConstraint("TABLE_NAME") & _
                    IIf(mrsConstraint("DELETE_RULE") = "NO ACTION", "", "(*)") & ","
                mrsConstraint.MoveNext
            Loop
            If strTemp <> "" Then strTemp = Mid(strTemp, 1, Len(strTemp) - 1) 'ȥ�����һ���Ķ���
            lst.SubItems(1) = strTemp
        End If
        rsTemp.MoveNext
    Loop

    Exit Sub
ErrHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub cmdSelect_Click()
    Dim i As Integer
    For i = 1 To lvwTable.ListItems.Count
        lvwTable.ListItems(i).Checked = True
    Next
End Sub

Private Sub Command1_Click()
    ShowHelp Me.hwnd, "zl9svrtools\" & Me.name
End Sub

Private Sub Form_Resize()
    Dim sngTemp As Single
    
    On Error Resume Next
    sngTemp = IIf(ScaleWidth > 6000, ScaleWidth, 6000)
    cmbSystem.Width = sngTemp - cmbSystem.Left - 200
    cmdFile.Left = sngTemp - cmdFile.Width - 200
    txtFile.Width = cmdFile.Left - 15 - txtFile.Left
    lvwTable.Width = cmbSystem.Width
    cmdExecute.Left = lvwTable.Left + lvwTable.Width - cmdExecute.Width
    
    sngTemp = IIf(ScaleHeight > 3000, ScaleHeight, 3000)
    cmdExecute.Top = sngTemp - cmdExecute.Height - 200
    cmdClear.Top = cmdExecute.Top
    cmdSelect.Top = cmdExecute.Top
    lvwTable.Height = cmdExecute.Top - lvwTable.Top - 100
'    lbl˵��.Width = ScaleWidth - 200 - lbl˵��.Left
'    lbl˵��.Height = ScaleHeight - 200 - lbl˵��.Top
    
End Sub

Private Function CheckIniFile(FileName As String, blnCmb As Boolean) As Boolean
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim strIniPath As String
    Dim strTemp As String
    Dim strTable As String
    Dim lngCount As Long
    On Error Resume Next
    
    strIniPath = Mid(FileName, 1, Len(FileName) - 11)
    '����ļ�ƥ���Լ��
    If Dir(strIniPath & "zlAppData.sql") = "" Then
        MsgBox "Ӧ�������ļ�" & strIniPath & "zlAppData.sql��ʧ�����ܼ�����", vbExclamation, gstrSysName
        txtFile.Text = ""
        Exit Function
    End If
    
    If mrsSystem.EOF Then
        txtFile.Text = ""
        Exit Function
    End If
    '�����ļ���ȷ�Լ��
    Set objText = objFile.OpenTextFile(FileName)
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[ϵͳ��]" Then
        If Val(Mid(strTemp, 6)) <> mrsSystem("���") \ 100 Then
            If blnCmb = False Then MsgBox "��ѡ�ļ����Ǹ�ϵͳ�İ�װ�����ļ�", vbExclamation, gstrSysName
            txtFile.Text = ""
            Exit Function
        End If
    Else
        Err.Raise 10
    End If
    strTemp = Trim(objText.ReadLine) 'ȡ��ϵͳ��
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[�汾��]" Then
        If InStr(1, mrsSystem("�汾��"), Trim(Mid(strTemp, 6))) = 0 Then
            MsgBox "ѡ���ļ����ϵͳ�汾����", vbExclamation, gstrSysName
            txtFile.Text = ""
            Exit Function
        End If
    Else
        Err.Raise 10
    End If
    
    objText.Close
    
    If Err <> 0 Then
        CheckIniFile = False
        MsgBox "��װ�����ļ�����ȷ" & vbNewLine & Err.Description, vbExclamation, gstrSysName
        txtFile.Text = ""
        Exit Function
    End If
    '�õ�������
    '��ռ���
    For lngCount = 1 To mcolTable.Count
        mcolTable.Remove 1
    Next
    '�������еĻ�����
    Set objText = objFile.OpenTextFile(strIniPath & "zlAppData.sql")
    Do Until objText.AtEndOfStream
        strTemp = UCase(objText.ReadLine())
        lngCount = InStr(strTemp, "INTO")
        If lngCount > 0 Then 'ȥ��ǰ���"insert into"
            strTemp = Mid(strTemp, lngCount + 4)
            lngCount = InStr(strTemp, "(") 'ȥ�������"("
            If lngCount > 0 Then
                strTemp = Trim(Mid(strTemp, 1, lngCount - 1))
                If strTemp <> "" And strTemp <> strTable Then
                    strTable = strTemp
                    mcolTable.Add strTable, "C" & strTable
                End If
            End If
        End If
    Loop
    objText.Close
    CheckIniFile = True
End Function

Private Sub Form_Load()
    
On Error GoTo ErrHandle
    frmMDIMain.MousePointer = 11

    '��ʾ���п���ʾ��ϵͳ
    
    If gblnDBA = True Then
        Set mrsSystem = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
    Else
        Set mrsSystem = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", UCase(gstrUserName))
    End If
    
    Do Until mrsSystem.EOF
        cmbSystem.AddItem mrsSystem("����") & " v" & mrsSystem("�汾��") & "��" & mrsSystem("���") & "��"
        cmbSystem.ItemData(cmbSystem.NewIndex) = mrsSystem("���")
        mrsSystem.MoveNext
    Loop
    If mrsSystem.RecordCount > 0 Then
        cmbSystem.ListIndex = 0
    Else
        cmdExecute.Enabled = False
    End If
    frmMDIMain.MousePointer = 0
    Exit Sub
ErrHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mrsSystem.State = 1 Then mrsSystem.Close
    Set mrsSystem = Nothing
    Set mcolTable = Nothing
    mstr������ = ""
End Sub

Private Sub lvwTable_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'
    SetCheck Item
End Sub


Private Sub SetCheck(ByVal Item As MSComctlLib.ListItem)
    Dim varTable As Variant
    Dim strTable As String
    Dim i As Integer
    Dim lst As ListItem
    
On Error GoTo ErrHandle
    If Item.Checked = False Then
        '����Ӽ��ı�û��ɾ���������������ı�Ҳ����ɾ��
        mrsConstraint.Filter = "table_name='" & Item.Text & "'"
        strTable = ""
        Do Until mrsConstraint.EOF
            If mrsConstraint("DELETE_RULE") = "NO ACTION" Then
                '����ɾ���ľͲ�������
                strTable = strTable & mrsConstraint("R_TABLE_NAME") & ","
            End If
            mrsConstraint.MoveNext
        Loop
        If strTable <> "" Then strTable = Mid(strTable, 1, Len(strTable) - 1) 'ȥ�����һ���Ķ���
        varTable = Split(strTable, ",")
        For i = LBound(varTable) To UBound(varTable)
            If varTable(i) <> Item.Text Then
                Set lst = lvwTable.ListItems("C" & varTable(i))
                lst.Checked = False
                '���еݹ����
                SetCheck lst
            End If
        Next
    Else
        varTable = Split(Item.SubItems(1), ",")
        For i = LBound(varTable) To UBound(varTable)
            If InStr(varTable(i), "(*)") = 0 And varTable(i) <> Item.Text Then
                Set lst = lvwTable.ListItems("C" & varTable(i))
                lst.Checked = True
                '���еݹ����
                SetCheck lst
            End If
        Next
    
    End If
    Exit Sub
ErrHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub


Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�

End Sub

