VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmLoadOut 
   BackColor       =   &H80000005&
   Caption         =   "���ݵ���"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7320
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmLoadOut.frx":0000
   ScaleHeight     =   6540
   ScaleWidth      =   7320
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExecute 
      Caption         =   "ִ��(&E)"
      Height          =   350
      Left            =   4035
      TabIndex        =   8
      Top             =   4980
      Width           =   1155
   End
   Begin VB.DirListBox DirBak 
      Appearance      =   0  'Flat
      Height          =   3240
      Left            =   3735
      TabIndex        =   7
      Top             =   1605
      Width           =   3180
   End
   Begin VB.DriveListBox DriveBak 
      Height          =   300
      Left            =   3735
      TabIndex        =   6
      Top             =   1290
      Width           =   3180
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "ȫѡ(&A)"
      Height          =   350
      Left            =   1065
      TabIndex        =   10
      Top             =   4980
      Width           =   1080
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "ȫ��(&R)"
      Height          =   350
      Left            =   2325
      TabIndex        =   11
      Top             =   4980
      Width           =   1080
   End
   Begin VB.CommandButton cmdSQL 
      Caption         =   "����S&QL�ű���"
      Height          =   375
      Left            =   5325
      TabIndex        =   9
      Top             =   4950
      Width           =   1575
   End
   Begin VB.ComboBox cmbSystem 
      Height          =   300
      Left            =   2070
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   690
      Width           =   4815
   End
   Begin MSComctlLib.ListView lvwTabs 
      Height          =   3555
      Left            =   990
      TabIndex        =   4
      Top             =   1290
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   6271
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label lblFileName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ļ����Ŀ¼(&D)"
      Height          =   180
      Left            =   3750
      TabIndex        =   5
      Top             =   1080
      Width           =   1350
   End
   Begin VB.Label lblTabs 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ݱ�ѡ��(T)"
      Height          =   180
      Left            =   1020
      TabIndex        =   3
      Top             =   1050
      Width           =   1170
   End
   Begin VB.Label lbl˵�� 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   990
      TabIndex        =   12
      Top             =   5490
      Width           =   6195
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ϵͳ(&S)"
      Height          =   180
      Index           =   1
      Left            =   1020
      TabIndex        =   1
      Top             =   750
      Width           =   990
   End
   Begin VB.Image imgICO 
      Height          =   480
      Left            =   240
      Picture         =   "frmLoadOut.frx":04F9
      Top             =   690
      Width           =   480
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ݵ���"
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
Attribute VB_Name = "frmLoadOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mrsSystem As New ADODB.Recordset
Dim mstr������ As String '���浱ǰϵͳ����������

Private Sub cmdSelect_Click()
    Dim i As Integer
    For i = 1 To lvwTabs.ListItems.Count
        lvwTabs.ListItems(i).Checked = True
    Next
End Sub

Private Sub cmdClear_Click()
    Dim i As Integer
    For i = 1 To lvwTabs.ListItems.Count
        lvwTabs.ListItems(i).Checked = False
    Next
End Sub

Private Sub DriveBak_Change()
    Dim strDrive As String
        
    On Error GoTo UnDo
    strDrive = DirBak.Path
    DirBak.Path = DriveBak.Drive
    Exit Sub
UnDo:
    MsgBox "������δ׼����", vbExclamation, gstrSysName
    DriveBak.Drive = strDrive
End Sub
'
'Private Sub chkDate_Click()
'    If chkDate.Value = 1 Then
'        dtpStart.Enabled = True
'        dtpEnd.Enabled = True
'    Else
'        dtpStart.Enabled = False
'        dtpEnd.Enabled = False
'    End If
'End Sub
'
'Private Sub dtpEnd_Change()
'    dtpStart.MaxDate = dtpEnd.Value
'    If dtpStart.Value > dtpStart.MaxDate Then dtpStart.Value = dtpStart.MaxDate
'End Sub
'
'Private Function GetWhere(ByVal strTable As String) As String
'    'ʱ�䷶Χ
'    Dim strWhere As String
'
'    strWhere = ""
'    If chkDate.Value = 1 Then
'        Select Case strTable
'            Case "������ҳ"
'                strWhere = "��Ժ����"
'            Case "���˴��շ���"
'                strWhere = "��ʼ����"
'            Case "���˴�λ����"
'                strWhere = "��ʼ����"
'            Case "���˴�λ��¼"
'                strWhere = "��סʱ��"
'            Case "���˷��ü�¼"
'                strWhere = "�Ǽ�ʱ��"
'            Case "���˽��ʼ�¼"
'                strWhere = "�շ�ʱ��"
'            Case "���������¼"
'                strWhere = "���ʱ��"
'            Case "�������η���"
'                strWhere = "��ʼ����"
'            Case "������Ϣ"
'                strWhere = "�Ǽ�ʱ��"
'            Case "����Ԥ����¼"
'                strWhere = "�տ�ʱ��"
'            Case "���ű�"
'                strWhere = "����ʱ��"
'            Case "��λ�ȼ�"
'                strWhere = "����ʱ��"
'            Case "��λ����"
'                strWhere = "��¼����"
'            Case "��ҩ;��"
'                strWhere = "����ʱ��"
'            Case "�������ݱ�"
'                strWhere = "��������"
'            Case "�Һ���Ŀ"
'                strWhere = "����ʱ��"
'            Case "��Լ��λ"
'                strWhere = "����ʱ��"
'            Case "����ȼ�"
'                strWhere = "����ʱ��"
'            Case "���ﲡ����¼"
'                strWhere = "��������"
'            Case "Ʊ���ش��¼"
'                strWhere = "��ӡʱ��"
'            Case "�շѼ�Ŀ"
'                strWhere = "ִ������"
'            Case "�շ�ϸĿ"
'                strWhere = "����ʱ��"
'            Case "������Ŀ"
'                strWhere = "����ʱ��"
'            Case "δ��ҩƷ��¼"
'                strWhere = "��������"
'            Case "ҩƷ�ɹ��ƻ�"
'                strWhere = "��������"
'            Case "ҩƷ�����¼"
'                strWhere = "�Ǽ�ʱ��"
'            Case "ҩƷ�����¼"
'                strWhere = "��������"
'            Case "ҩƷ��Ӧ��"
'                strWhere = "����ʱ��"
'            Case "ҩƷĿ¼"
'                strWhere = "����ʱ��"
'            Case "ҩƷ�շ���¼"
'                strWhere = "�������"
'            Case "ҽ����¼"
'                strWhere = "�Ǽ�ʱ��"
'            Case "������Ŀ"
'                strWhere = "����ʱ��"
'            Case "סԺ������¼"
'                strWhere = "��������"
'        End Select
'        If strWhere <> "" Then
'            strWhere = " Where " & strWhere & " Between To_Date('" & Format(dtpStart, "yyyy-MM-dd HH:mm:ss") _
'                    & "','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtpEnd, "yyyy-MM-dd HH:mm:ss") _
'                    & "','YYYY-MM-DD HH24:MI:SS')"
'        End If
'    End If
'    GetWhere = strWhere
'End Function

Private Sub SetEnable(ByVal blnEnable As Boolean)
    frmMDIMain.Enabled = blnEnable
    lvwTabs.Enabled = blnEnable
    DirBak.Enabled = blnEnable
    DriveBak.Enabled = blnEnable
    cmdClear.Enabled = blnEnable
    cmdSelect.Enabled = blnEnable
    cmdExecute.Enabled = blnEnable
    cmdSQL.Enabled = blnEnable
End Sub

Private Sub cmdExecute_Click()
    Dim blnGen As Boolean
    Dim strBakDir As String
    Dim strLDR As String
    Dim strField As String
    Dim strTable As String
    Dim objItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    
    blnGen = False
    If MsgBox("����û����ݹ��󣬸ó��������ٶȻ�Ƚ�����" & vbCrLf & "Ҫ����ִ����", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    SetEnable False
    strBakDir = DirBak.Path
    If Right(strBakDir, 1) <> "\" Then strBakDir = strBakDir & "\"
    
    MousePointer = 11
    For Each objItem In lvwTabs.ListItems
        If objItem.Checked Then
            strTable = objItem.Text
            frmMDIMain.stbThis.Panels(2).Text = "����" & strTable & "��"
            gstrSQL = "select COLUMN_NAME,DATA_TYPE " & _
                    " from all_tab_columns " & _
                    " where owner='" & mstr������ & "' and table_name='" & strTable & "'" & _
                    "       and DATA_TYPE not in('LONG','LONG RAW','CLOB','BLOB','BFILE','NCLOB','NBLOB')" & _
                    " order by column_id"
            With rsTemp
                If .State = adStateOpen Then .Close
                .Open gstrSQL, gcnOracle, adOpenKeyset
                
                strLDR = ""
                strField = ""
                Do Until .EOF
                    If .Fields(1).value = "DATE" Then
                        strLDR = strLDR & "||'^'||To_Char(""" & .Fields(0).value & """,'YYYY-MM-DD HH24:MI:SS')"
                        strField = strField & ",""" & .Fields(0).value & """ Date 'YYYY-MM-DD HH24:MI:SS' "
                    Else
                        strLDR = strLDR & "||'^'||""" & .Fields(0).value & """"
                        strField = strField & ",""" & .Fields(0).value & """"
                    End If
                    .MoveNext
                Loop
                
                '��ѯSQL���
                If strLDR = "" Then
                    objItem.Checked = False
                Else
                    blnGen = True
                    gstrSQL = "select " & Mid(strLDR, 8) & " from " & mstr������ & "." & strTable ' & GetWhere(strTable)
                    If .State = adStateOpen Then .Close
                    .Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
                    If .RecordCount <> 0 Then
                        On Error Resume Next
                        Kill strBakDir & strTable & ".ldr"
                        On Error GoTo 0
                        Open strBakDir & strTable & ".ldr" For Binary Access Write As #1
                        strLDR = "LOAD DATA INFILE *" & vbCrLf & _
                                "PRESERVE BLANKS" & vbCrLf & _
                                "INTO TABLE " & strTable & " APPEND" & vbCrLf & _
                                "FIELDS TERMINATED BY '^'" & vbCrLf & _
                                "TRAILING NULLCOLS(" & Mid(strField, 2) & ")" & vbCrLf & _
                                "BEGINDATA"
                        Put #1, , strLDR & vbCrLf
                        Do Until .EOF
                            Put #1, , Replace(CStr(.Fields(0).value), vbCrLf, vbCr) & vbCrLf
                            If Int(.AbsolutePosition / .RecordCount * 1000) Mod 10 = 0 Then
                                frmMDIMain.stbThis.Panels(2).Text = "����" & strTable & "��" & String(Int(.AbsolutePosition * 16 / .RecordCount), "��")
                                DoEvents
                            End If
                            .MoveNext
                        Loop
                        Close #1
                    Else
                        objItem.Checked = False
                    End If
                End If
            End With
        End If
    Next
    MousePointer = 0
    frmMDIMain.stbThis.Panels(2).Text = ""
    SetEnable True
        
    If Not blnGen Then
        MsgBox "����û��ѡ����ѡ��ı�û�б����ֶΣ��޷�����Loader�ļ���", vbExclamation, gstrSysName
    Else
        MsgBox "Loader�ļ�������ϡ�" & vbCr _
            & vbCr & "�������ѡ�е����ݱ�����ѡ��״̬" _
            & vbCr & "˵���ñ�û�����ݻ�û�м����������У�", vbExclamation, gstrSysName
    End If
End Sub

Private Sub cmdSQL_Click()
    Dim blnGen As Boolean
    Dim strBakDir As String
    Dim strLDR As String
    Dim strField As String
    Dim strTable As String
    Dim objItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    
    blnGen = False
    
    SetEnable False
    strBakDir = DirBak.Path
    If Right(strBakDir, 1) <> "\" Then strBakDir = strBakDir & "\"
    
    '����CreateLoaderData.Sql�ļ�
    On Error Resume Next
    Kill strBakDir & "CreateLoaderData.Sql"
    
    '��ӱ�ͷ����
    On Error GoTo 0
    Open strBakDir & "CreateLoaderData.Sql" For Binary Access Write As #1
    Put #1, , "set echo off heading off feedback off verify off;" & vbCrLf
    Put #1, , "set linesize 30000 pagesize 0 trimspool on;" & vbCrLf
    Put #1, , "set termout off;" & vbCrLf & vbCrLf
    
    For Each objItem In lvwTabs.ListItems
        If objItem.Checked Then
            strTable = objItem.Text
            gstrSQL = "select COLUMN_NAME,DATA_TYPE " & _
                    " from all_tab_columns " & _
                    " where owner='" & mstr������ & "' and table_name='" & strTable & "'" & _
                    "       and DATA_TYPE not in('LONG','LONG RAW','CLOB','BLOB','BFILE','NCLOB','NBLOB')" & _
                    " order by column_id"
            
            With rsTemp
                If .State = adStateOpen Then .Close
                .Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
                
                strLDR = ""
                strField = ""
                Do While Not .EOF
                    If .Fields(1).value = "DATE" Then
                        strLDR = strLDR & "||'^'||To_Char(""" & .Fields(0).value & """,'YYYY-MM-DD HH24:MI:SS')"
                        strField = strField & ",""" & .Fields(0).value & """ Date 'YYYY-MM-DD HH24:MI:SS' "
                    Else
                        strLDR = strLDR & "||'^'||""" & .Fields(0).value & """"
                        strField = strField & ",""" & .Fields(0).value & """"
                    End If
                    .MoveNext
                Loop
                If strLDR = "" Then
                    objItem.Checked = False
                Else
                    blnGen = True
                    gstrSQL = "select " & Mid(strLDR, 8) & " from " & mstr������ & "." & strTable
                    Put #1, , "spool " & strBakDir & strTable & ".ldr;" & vbCrLf
                    '�����ļ���ͷ��ʼ
                    Put #1, , "select 'LOAD DATA INFILE *' from dual;" & vbCrLf
                    Put #1, , "select 'PRESERVE BLANKS' from dual;" & vbCrLf
                    Put #1, , CStr("select 'INTO TABLE " & strTable & " APPEND' from dual;") & vbCrLf
                    Put #1, , "select 'FIELDS TERMINATED BY ''^''' from dual;" & vbCrLf
                    Put #1, , CStr("select 'TRAILING NULLCOLS(" _
                        & Replace(Mid(strField, 2), "'", "''") & ")' from dual;") & vbCrLf
                    Put #1, , "select 'BEGINDATA' from dual;" & vbCrLf
                    '�����ļ���ͷ����
                    
                    '��ѯSQL���
                    Put #1, , gstrSQL & ";" & vbCrLf ' & GetWhere(strTable)
                    Put #1, , "spool off;" & vbCrLf & vbCrLf
                End If
            End With
            frmMDIMain.stbThis.Panels(2) = "�������ɡ�" & strTable & "����Ľű�����"
        End If
    Next
    Put #1, , "exit" & vbCrLf
    Close #1
    frmMDIMain.stbThis.Panels(2) = ""
    SetEnable True
    
    If Not blnGen Then
        MsgBox "����û��ѡ����ѡ��ı�û�б����ֶΣ��޷����ɽű��ļ���", vbExclamation, gstrSysName
    Else
        If MsgBox("SQL�ű��Ѿ�������ϡ�����ġ�" & strBakDir & "CreateLoaderData.Sql���ļ���" & vbCrLf & vbCrLf & _
               "�������Ͼ����иýű���", vbQuestion Or vbYesNo, gstrSysName) = vbYes Then
            
            Dim lngTemp As Long
            Dim lngProcess  As Long
            
            frmWait.BeginWait "��������SQL�ű�����"
            lngTemp = Shell("sqlplus " & gstrUserName & "/" & gstrPassword & IIf(gstrServer = "", "", "@" & gstrServer) & " @" & strBakDir & "CreateLoaderData.sql", vbHide)
            If err <> 0 Then
                err.Clear
                MsgBox "������ȷ���ɽű������飺" & _
                    vbCr & "   1) �Ƿ����sqlplus.exe�ļ���" & _
                    vbCr & "   2) Set Path�Ƿ�ָ������ڵ�Ŀ¼��", vbExclamation, gstrSysName
                frmWait.EndWait
                Exit Sub
            End If
            
            lngProcess = OpenProcess(Process_Query_Information, False, lngTemp)
            Do
                Sleep 100
                GetExitCodeProcess lngProcess, lngTemp
                DoEvents
            Loop While lngTemp = Still_Active
            CloseHandle lngProcess
            frmWait.EndWait
                
            If lngTemp <> 0 Then
                MsgBox "�ű����ɳ���Ƿ��˳���", vbCritical, gstrSysName
            End If
        End If
    End If
End Sub

Private Sub cmbSystem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then lvwTabs.SetFocus
End Sub

Private Sub Form_Load()
    lbl˵��.Caption = "˵����" & vbCrLf & _
                    "     ��������Ҫ����һ���������Ĺ��̲�����ɡ������ʱ���ڷ������Կͻ�����Ӧ���óٶۣ�������ڷ���������ʱ��ɱ�������" & vbCrLf & _
                    "     ����ÿ�����ݱ������һ��ͬ���ĵ����ļ�������ʱ�ɸ�����Щ�ļ�������ɡ�" & vbCrLf & _
                    "     ��SQLPLUS�������ɵĽű����õ������ļ������ܱȱ�����ִ�е�Ч�ʸ��ߣ���˽���ʹ�ýű���ʽ��"
'    dtpStart.Value = Format(Date & " " & Format("0:0:0", "HH:mm:ss"), "yyyy-MM-dd HH:mm:ss")
'    dtpEnd.Value = Format(Date & " " & Format("23:59:59", "HH:mm:ss"), "yyyy-MM-dd HH:mm:ss")
    Call FillSystem
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mrsSystem.State = 1 Then mrsSystem.Close
    Set mrsSystem = Nothing
    mstr������ = ""
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    lbl˵��.Width = ScaleWidth - 200 - lbl˵��.Left
    lbl˵��.Height = ScaleHeight - 200 - lbl˵��.Top
    
End Sub

Private Sub cmbSystem_Click()
    Dim rsTemp As New ADODB.Recordset
    
    mrsSystem.Filter = "���=" & cmbSystem.ItemData(cmbSystem.ListIndex)
    lvwTabs.ListItems.Clear
    If mrsSystem.RecordCount = 0 Then
        cmdExecute.Enabled = False
    Else
        cmdExecute.Enabled = True
        mstr������ = mrsSystem("������")
        
        rsTemp.CursorLocation = adUseClient
        rsTemp.Open "select table_name from all_tables where owner='" & mstr������ & "'", gcnOracle, adOpenStatic, adLockReadOnly
        Do Until rsTemp.EOF
            lvwTabs.ListItems.Add , , rsTemp("TABLE_NAME")
            rsTemp.MoveNext
        Loop
    End If
End Sub

Private Sub FillSystem()
    '��ʾ���п���ʾ��ϵͳ
    On Error GoTo errHandle
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
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�

End Sub

