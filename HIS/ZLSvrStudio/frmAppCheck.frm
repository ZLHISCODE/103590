VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAppCheck 
   BackColor       =   &H80000005&
   Caption         =   "�������޸�"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5610
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmAppCheck.frx":0000
   ScaleHeight     =   6135
   ScaleWidth      =   5610
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdFunction 
      Caption         =   "����Ȩ������(&O)"
      Height          =   350
      Index           =   6
      Left            =   825
      TabIndex        =   17
      Top             =   3555
      Width           =   1650
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "��ʷ������(&H)"
      Height          =   350
      Index           =   5
      Left            =   825
      TabIndex        =   16
      Top             =   3550
      Width           =   1650
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "ͬ�������(&N)"
      Height          =   350
      Index           =   4
      Left            =   825
      TabIndex        =   15
      Top             =   3210
      Width           =   1650
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "��������(&S)"
      Height          =   350
      Index           =   3
      Left            =   825
      TabIndex        =   10
      Top             =   2880
      Width           =   1650
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   5550
      TabIndex        =   12
      Top             =   5595
      Visible         =   0   'False
      Width           =   5610
      Begin MSComctlLib.ProgressBar pgbState 
         Height          =   180
         Left            =   135
         TabIndex        =   13
         Top             =   255
         Width           =   6180
         _ExtentX        =   10901
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ڼ��"
         Height          =   180
         Left            =   135
         TabIndex        =   14
         Top             =   60
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "�ؽ�����(&I)"
      Height          =   350
      Index           =   2
      Left            =   825
      TabIndex        =   9
      Top             =   2550
      Width           =   1650
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "��������(&R)"
      Height          =   350
      Index           =   1
      Left            =   825
      TabIndex        =   3
      Top             =   2220
      Width           =   1650
   End
   Begin VB.CommandButton cmdGetIni 
      Caption         =   "ѡ��(&S)��"
      Height          =   350
      Left            =   4215
      TabIndex        =   5
      Top             =   1020
      Width           =   1095
   End
   Begin VB.ComboBox cmbSystem 
      Height          =   300
      Left            =   1740
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   645
      Width           =   3570
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "������(&C)"
      Height          =   350
      Index           =   0
      Left            =   825
      TabIndex        =   2
      Top             =   1860
      Width           =   1650
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "������ϵͳ�����ݿ�����밲װ�ļ��Աȣ��������ϵͳ�����С�����ͼ���������洢���̵ȶ������ȷ�ԡ�"
      ForeColor       =   &H80000008&
      Height          =   1980
      Left            =   2565
      TabIndex        =   11
      Top             =   1905
      Width           =   2730
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMain 
      BackStyle       =   0  'Transparent
      Height          =   840
      Left            =   900
      TabIndex        =   8
      Top             =   4170
      Width           =   4275
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblFileName 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   960
      TabIndex        =   7
      Top             =   1380
      Width           =   4350
   End
   Begin VB.Label lblFileCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��װ�����ļ�"
      Height          =   180
      Left            =   960
      TabIndex        =   6
      Top             =   1140
      Width           =   1080
   End
   Begin VB.Label lblSys 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ӧ��ϵͳ"
      Height          =   180
      Left            =   975
      TabIndex        =   1
      Top             =   705
      Width           =   720
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������޸�"
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
      Width           =   1440
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   240
      Picture         =   "frmAppCheck.frx":04F9
      Top             =   645
      Width           =   480
   End
End
Attribute VB_Name = "frmAppCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrIniPath    As String                 '��װ�����ļ�Ŀ¼
Private mrsErrTable As ADODB.Recordset
Private mrsExeErrTable As ADODB.Recordset

Private WithEvents mclsObjectCheck As clsObjectCheck
Attribute mclsObjectCheck.VB_VarHelpID = -1
Private mclsRunScript As New clsRunScript

Private Enum CMDFUN
    E������ = 0
    E�������� = 1
    E�ؽ����� = 2
    E�������� = 3
    Eͬ��� = 4
    E��ʷ�ṹ = 5
    E����Ȩ�� = 6
End Enum

Private Sub cmbSystem_Click()
    Dim strFilePath As String
    Dim blnTools As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    If Val(cmbSystem.ItemData(cmbSystem.ListIndex)) = -1 Then
        strFilePath = App.Path & "\Tools\zlServer.Sql"
        cmbSystem.Tag = "ZLTOOLS"
    Else
        cmbSystem.Tag = GetOwnerName(Val(cmbSystem.ItemData(cmbSystem.ListIndex)), gcnOracle)
        Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Loadandunload.Get_Sysfile_name", Val(cmbSystem.ItemData(cmbSystem.ListIndex)), 1)
        If Not rsTemp.EOF Then strFilePath = rsTemp.Fields(0).value & ""
    End If
    '���ù���״̬
    Call SetFunsState(strFilePath, False)
    
End Sub

Private Function zl��ȡ�������(ByVal str������ As String, ByRef rsData As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�������Ĺ������
    '���:str������-��������
    '����:rsData
    '����:
    '����:���˺�
    '����:2009-08-20 11:30:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As New ADODB.Recordset, strTemp As String


    If gblnDBA Then
        strSQL = "Select table_Name,Constraint_Name,OWNER,R_OWNER,R_Constraint_Name,DELETE_RULE from DBA_CONSTRAINTS Where R_Constraint_Name='" & str������ & "' And Constraint_Type='R'"
    Else
        strSQL = "Select table_Name,Constraint_Name,OWNER,R_OWNER,R_Constraint_Name,DELETE_RULE From USER_CONSTRAINTS Where  R_Constraint_Name='" & str������ & "'  And Constraint_Type='R'"
    End If

    With rsTemp
        If .State = adStateOpen Then .Close
        .Open strSQL, gcnOracle, adOpenKeyset
        If .RecordCount = 0 Then zl��ȡ������� = True: Exit Function
        .Filter = "OWNER<>'" & cmbSystem.Tag & "' AND R_OWNER='" & cmbSystem.Tag & "'"
        If .RecordCount <> 0 Then
            '��������ϵͳ����,��������ֻ���ֹ�����
            If Not rsData.EOF Then
                If UCase(Nvl(rsData!��������)) <> UCase(str������) Then
                    rsData.Filter = "��������='" & UCase(str������) & "'"
                End If
            Else
                rsData.Filter = "��������='" & UCase(str������) & "'"
            End If
            If rsData.EOF Then rsData.Filter = 0: zl��ȡ������� = True: Exit Function
            .MoveFirst: strTemp = ""
            Do While Not .EOF
                strTemp = strTemp & "," & Nvl(!Table_Name) & "(" & Nvl(!Owner) & ")"
                .MoveNext
            Loop
            If strTemp <> "" Then strTemp = Mid(strTemp, 2)
            If strTemp <> "" Then strTemp = Substr("������ϵͳ������������Ч���������:" & strTemp, 1, 500)
            '���±�־
            rsData!������־ = 4
            rsData!����˵�� = strTemp
            rsData.Update
            .Close
            zl��ȡ������� = True: Exit Function
        End If
        .Filter = 0: .MoveFirst

        '��Ҫɾ����صļ���
        Do While Not .EOF
            '��־: '1-���ڶ���,2-�����ڶ���,3-ʧЧ,4-ȱ����,5-����,6-���ڽ�ֹ״̬,7-Լ����һ��,8-��������ʱ����Ҫ�ȴ������
            Call zlInsertRecData(rsData, Nvl(!Table_Name), Nvl(!Constraint_Name), "���", 8, False, "", "��������", "�ڴ���������Ψһ��ʱ����Ҫ�ȴ��������")
            .MoveNext
        Loop
        .Close
        zl��ȡ������� = True: Exit Function
    End With
End Function

Private Sub ModifyToolsObject(ByVal cnTools As ADODB.Connection, ByVal blnDele As Boolean, ByVal strSeverScrip As String)
    '-----------------------------------------------------------------------------------------------------------
    '����:����ZLTOOLS��ض���
    '����:cnTools-���ӵ������ߵ�����
    '     blnDele-�Ƿ�ɾ����ص�Լ����
    '     strSeverScrip-�����������߽ű��ļ�
    '����:
    '����:���˺�
    '����:2007/09/12
    '-----------------------------------------------------------------------------------------------------------
        Dim rsTemp As New ADODB.Recordset
    Dim rsSys As New ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    '���ݰ�װ�ļ���������
    'һ�����У�ֱ�����½������������
    '����Լ����ɾ������Լ�����½���
    '�ġ�������ɾ�������������½���
    '�塢��ͼ��ֱ�����½���
    '��������ֱ�����½���
    '�ߡ�ͬ��ʣ����ݹ��߶�������ͬ���
    
    If blnDele Then
        lblStatus.Caption = "ɾ���������Լ����"
        strSQL = "select 'alter table '||table_name||' drop constraint '||constraint_name" & _
                " From user_constraints" & _
                " where constraint_type IN ('R','C') And Instr(Table_Name,'BIN$')<=0 And Instr(constraint_name,'SYS_')<=0 "
        OpenRecordset rsTemp, strSQL, Me.Caption, , , cnTools
        With rsTemp
            Do While Not .EOF
                cnTools.Execute .Fields(0).value
                pgbState.value = .AbsolutePosition / .RecordCount * 100
                DoEvents
                .MoveNext
            Loop
        End With
        lblStatus.Caption = "ɾ��������ΨһԼ����"
        strSQL = "select 'alter table '||table_name||' drop constraint '||constraint_name" & _
                " From user_constraints Where Instr(Table_Name,'BIN$')<=0 And Instr(constraint_name,'SYS_')<=0 "
        OpenRecordset rsTemp, strSQL, Me.Caption, , , cnTools
        With rsTemp
            Do While Not .EOF
                cnTools.Execute .Fields(0).value
                pgbState.value = .AbsolutePosition / .RecordCount * 100
                DoEvents
                .MoveNext
            Loop
        End With
        lblStatus.Caption = "ɾ��������"
        strSQL = "select 'drop Index ""'||Index_name||'""'  from user_indexes Where INDEX_TYPE='NORMAL' And Instr(Table_Name,'BIN$')<=0"
        OpenRecordset rsTemp, strSQL, Me.Caption, , , cnTools
        With rsTemp
            Do While Not .EOF
                cnTools.Execute .Fields(0).value
                pgbState.value = .AbsolutePosition / .RecordCount * 100
                DoEvents
                .MoveNext
            Loop
        End With
    End If
    '���ݰ�װ�ű�����ִ��
    
    err = 0: On Error Resume Next
    
    If Not mclsRunScript.OpenFile(strSeverScrip) Then Exit Sub
    lblStatus.Caption = "������������"
    Do While Not mclsRunScript.EOF
        strSQL = mclsRunScript.SQLInfo.SQL
'        If mclsRunScript.Line >= 25239 Then Stop
        If Not mclsRunScript.SQLInfo.Block Then
            strTmp = mclsRunScript.SQLInfo.PartSQL
            If strTmp Like "CREATE TABLE *" Or strTmp Like "CREATE GLOBAL TEMPORARY TABLE *" Then
              '������
            ElseIf strTmp Like "ALTER TABLE * CONSTRAINT *" Then
                '����Լ��
                    cnTools.Execute strSQL
            ElseIf strTmp Like "CREATE INDEX *" Then
                '��������
                cnTools.Execute strSQL
            ElseIf strTmp Like "CREATE SEQUENCE *" Then
                '�������
                cnTools.Execute strSQL
            End If
        Else
            strTmp = mclsRunScript.SQLInfo.BlockType
            If strTmp = "TYPE" Then
            
                '������
            ElseIf strTmp Like "*PROCEDURE*" Or _
                    strTmp Like "*FUNCTION*" Or _
                    strTmp Like "*PACKAGE*" Then
                '�������뺯���ĺϷ���
                cnTools.Execute strSQL
            End If
        End If
        err.Clear
        pgbState.value = mclsRunScript.ProcessValue
        Call mclsRunScript.ReadNextSQL
        DoEvents
    Loop
    lblStatus.Caption = "������������ͬ��ʡ�"
    Call ReGrantForTools(cnTools, , True)
    lblStatus.Caption = "�����������С�"
    Call AdjustSequence("ZLTOOLS", cnTools)
End Sub
Private Function zlCheckTableDataIsNull(ByVal strTableName As String, ByVal str�ֶ��� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ�����ݱ���ֶ��ֶ��Ƿ�ΪNULL
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-08-20 14:07:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    err = 0: On Error Resume Next
    
    strSQL = "Select 1 From " & cmbSystem.Tag & "." & strTableName & " where " & str�ֶ��� & " is not null and rownum<=1"
    rsTemp.Open strSQL, gcnOracle
    zlCheckTableDataIsNull = Not (rsTemp.RecordCount <> 0)
    
End Function

Private Function zlModifyObject(ByVal str���� As String, Optional blnDelete As Boolean = True) As Boolean
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '����:���ݶ�������
        '���:str����-ָ��������:���;Լ��;����
        '     blnDelete-�Ƿ�ɾ��
        '����:
        '����:
        '����:���˺�
        '����:2009-08-19 17:35:10
        '---------------------------------------------------------------------------------------------------------------------------------------------
         Dim i As Long, strSQL As String, str�޸�˵�� As String, byt������־ As Integer, varData As Variant, lng���� As Long, lng���� As Long
         Dim iCount As Integer
         
        If blnDelete Then
            lblStatus.Caption = "����ɾ������" & str���� & "��"
        Else
            lblStatus.Caption = "������������" & str���� & "��"
        End If
        mrsErrTable.Filter = "����='" & str���� & "'" & IIf(blnDelete, "", " and ������־<>2")
        With mrsErrTable
            i = 0: iCount = .RecordCount
            Do While Not .EOF
            
                strSQL = Nvl(!�������)
                Select Case str����
                Case "Լ��", "���"
                   If blnDelete Then strSQL = " Alter table " & Nvl(!������) & "  Drop constraint " & Nvl(!��������)
                Case "����"
                    If blnDelete Then strSQL = " Drop index " & Nvl(!��������)
                Case "���ݱ�"
                    '��Ҫ���,�Ƿ�������־Ϊ�ֹ�ִ�е�
                     If !������־ = 4 And blnDelete = False And Nvl(!ԭ�ֶ�����) <> "" Then  '0-δ����,1-�Ѿ�����,2-����ʧ��,4-����ִ����������Ҫ�ֹ�����
                        '����С�Ļ�,��Ҫ��������Ƿ�Ϊ��,��Ϊ��,���������
                        varData = Split(Nvl(!ԭ�ֶγ���) & ",", ",")
                        lng���� = Val(varData(0)): lng���� = Val(varData(1))
                        varData = Split(Nvl(!���ֶγ���) & ",", ",")
                        If lng���� > Val(varData(0)) Or lng���� > Val(varData(1)) Then
                           If zlCheckTableDataIsNull(Nvl(!������), Nvl(!�ֶ���)) Then
                                str�޸�˵�� = "��Ȼ���ȸ�С,�������ֶ�����Ϊ��,���,Ҳ�����˸ýṹ!"
                           Else '��Ϊ��,���ܸ���
                                strSQL = ""
                           End If
                        End If
                     End If
                End Select
                str�޸�˵�� = "�����ɹ�": byt������־ = 1
                
                If strSQL <> "" Then
                    err = 0: On Error Resume Next
                    gcnOracle.Execute strSQL
                    If err <> 0 Then
                          byt������־ = 2
                         Select Case str����
                         Case "Լ��"
                            str�޸�˵�� = Substr("������Ч����������Ψһ��,������Ϣ:" & err.Description, 1, 4000)
                         Case "���"
                            str�޸�˵�� = Substr("������Ч�������,������Ϣ:" & err.Description, 1, 4000)
                         Case "����"
                            str�޸�˵�� = Substr("������Ч��������,������Ϣ:" & err.Description, 1, 40000)
                         Case "���ݱ�"
                            str�޸�˵�� = Substr("������Ч�������ݱ�,������Ϣ:" & err.Description, 1, 40000)
                         End Select
                    End If
                    
                    If blnDelete = False Then
                        !����˵�� = str�޸�˵��
                        !������־ = byt������־
                        .Update
                    End If
                    err = 0: On Error GoTo 0
                End If
                i = i + 1
                pgbState.value = i / iCount * 100
                DoEvents
                .MoveNext
            Loop
            .Filter = 0
        End With
End Function
Private Sub cmdFunction_Click(Index As Integer)
    Dim i As Long, intVer As Integer
    Dim bytOperation As VbMsgBoxResult
    Dim blnTool As Boolean
    Dim cnTools As ADODB.Connection
    Dim lngSys As Long, lngAbort As Long
    Dim cllSQL As New Collection, cllErr As New Collection
    Dim strTmp As String
    Dim rsTemp As New ADODB.Recordset, strSQL As String
    Dim rsObjects As New ADODB.Recordset
    
    Call cmdFunction_MouseMove(Index, 0, 0, 0, 0)
    
    If MsgBox("""" & Split(cmdFunction(Index).Caption, "(")(0) & """�������������Ľ϶����Դ�ͻ��ѽϳ���ʱ�䣬Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    intVer = GetOracleVersion
    Select Case Index
    '------------------------------------------------------------------------------------------------------------------
    Case 0

        Set mclsObjectCheck = New clsObjectCheck
   
        If cmbSystem.Tag = "ZLTOOLS" Then
            
            If mclsObjectCheck.InitCheckManageTool(Me, lblFileName.Caption) = True Then
                Call mclsObjectCheck.CheckToolsObject
                Call mclsObjectCheck.ShowReport
            End If

        Else
            
            If mclsObjectCheck.InitCheck(Me, cmbSystem.Tag, lblFileName.Caption, gstrUserName, gstrSysName, gblnDBA, cmbSystem.ItemData(cmbSystem.ListIndex), cmbSystem.Text) = True Then
                Call mclsObjectCheck.CheckObject
                Call mclsObjectCheck.ShowReport
            End If
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case 1
        blnTool = (cmbSystem.Tag = "ZLTOOLS")
        If blnTool Then
            If gobjFile.FileExists(lblFileName.Caption) = False Then
                For i = 0 To cmdFunction.Ubound
                    cmdFunction(i).Enabled = False
                Next
                lblFileName.Caption = ""
                cmdGetIni.SetFocus
                Exit Sub
            End If
            
            Set cnTools = GetConnection("ZLTOOLS")
            If cnTools Is Nothing Then
                MsgBox "�򿪹�����ʧ��,����!", vbInformation, gstrSysName
                Exit Sub
            End If
        Else
            If CheckIniFile(lblFileName.Caption, True) = False Then
                For i = 0 To cmdFunction.Ubound
                    cmdFunction(i).Enabled = False
                Next
                lblFileName.Caption = ""
                cmdGetIni.SetFocus
                Exit Sub
            End If
        End If
        
        '���ݰ�װ�ļ���������
        'һ�����У�ֱ�����½������������
        '�������ݱ��ȼ������ݱ�Ĵ���,�Դ���ı��������(�ֶβ�����,����;�ֶ����Ͳ���,�����ֹ�����;�ֶξ��ȱȽű��Ĵ�ʱ,�����ֹ�����(��������ʱ,�Զ�����);�ֶξ��ȱȽű���С,���Զ�����)
        '����Լ����ɾ����������Լ��,Ȼ������ִ�ж�Ӧ��Լ���ű�
        '�ġ�������ɾ��������������,Ȼ������ִ�ж�Ӧ�������ű�
        '�塢��ͼ��ֱ�����½���
        '��������ֱ�����½���
        strTmp = "����:" & vbCrLf & _
                "    ִ�ж�����������,Ӧ���أ�Ϊ�˱������ݶ�ʧ��������ִ�иù���ǰ���������µļ�飺" & vbCrLf & _
                "1. ȷ�����е������û����ڶϿ�״̬��" & vbCrLf & _
                "2. ����Ӧ�ñ��ݣ��Ա����ݻָ���" & vbCrLf & _
                "���Ƿ����Ҫִ�С��������������ܣ�"
        
        bytOperation = MsgBox(strTmp, vbQuestion + vbYesNo + vbDefaultButton3, gstrSysName)
        If bytOperation = vbNo Then Exit Sub
        
        
        Set cllSQL = New Collection
        picStatus.Visible = True
        Enabled = False
        If blnTool Then
            '��ʼ���ڲ����ݼ�
            Call zlInitRec(mrsErrTable)
            Call ModifyToolsObject(cnTools, bytOperation = vbYes, lblFileName.Caption)
            DoEvents
            Call ReCompileProcedure(cnTools)
        Else
            '��ʼ���ڲ����ݼ�
            Call zlInitRec(mrsErrTable)
            
            '�ȼ�����
            lblStatus.Caption = "���ڼ�����ݱ�"
            Call CheckTable(mstrIniPath & "zlTable.sql", True)
            lblStatus.Caption = "���ڼ��Լ����"
            Call CheckConstraint(mstrIniPath & "zlConstraint.sql", True)
            
            lblStatus.Caption = "���ڼ��������"
            Call CheckIndex(mstrIniPath & "zlIndex.sql", True)
            
            '���ɾ��
            lblStatus.Caption = "����ɾ�������"
            Call zlModifyObject("���", True)
            'Լ��ɾ��
            lblStatus.Caption = "����ɾ��Լ����"
            Call zlModifyObject("Լ��", True)
            '����ɾ��
            lblStatus.Caption = "����ɾ��������"
            Call zlModifyObject("����", True)
            '�������ݱ�:
            lblStatus.Caption = "�����������ݱ�"
            Call zlModifyObject("���ݱ�", False)
            'Լ������:
            lblStatus.Caption = "��������Լ����"
            Call zlModifyObject("Լ��", False)
            '�������:
            lblStatus.Caption = "�������������"
            Call zlModifyObject("���", False)
            '��������:
            lblStatus.Caption = "��������������"
            Call zlModifyObject("����", False)
            
            lngSys = cmbSystem.ItemData(cmbSystem.ListIndex)
            '����ʵ���������ʹ�úۼ�
            Set mclsRunScript = New clsRunScript
            '���ò��������
            Call mclsRunScript.InitGlobalPara(Me, lngSys)
            '��ʼ���û�������Ϣ�����ܿ�����õ�
            Call mclsRunScript.InitUserList(gstrUserName, gstrPassword)
        
            '���ݰ�װ�ű����½���
            lblStatus.Caption = "�����������С�"
            Call RunSQLScript(gcnOldOra, mstrIniPath & "zlSequence.sql", True)
            lblStatus.Caption = "����������ͼ��"
            Call RunSQLScript(gcnOldOra, mstrIniPath & "zlView.sql", True)
            lblStatus.Caption = "����������������̡�"
            Call RunSQLScript(gcnOldOra, mstrIniPath & "zlProgram.sql", True)
            If Dir(mstrIniPath & "zlPackage.sql") <> "" Then
                lblStatus.Caption = "������������"
                Call RunSQLScript(gcnOldOra, mstrIniPath & "zlPackage.sql", True)
            End If
            lblStatus.Caption = "�����������С�"
            Call AdjustSequence(gstrUserName, gcnOldOra, lngSys)
            DoEvents
            Call ReCompileProcedure(gcnOldOra)
        End If
        
        picStatus.Visible = False
        '���ش�����Ϣ:���˺�
        '����,��������Ϣ��ֵ,�Ա���ʾ
        
        '"������־", adLongVarChar, 2, adFldIsNullable  '0-δ����,1-�Ѿ�����,2-�������,4-����ִ����������Ҫ�ֹ�����
        
        If Not mrsErrTable Is Nothing Then
           frmAppChkRpt.blnModiyfyCheck = True
           mrsErrTable.Filter = "������־=2"
           With mrsErrTable
               '������������
               If mrsErrTable.RecordCount <> 0 Then .MoveFirst
               Do While Not .EOF
                   Call InputErrModifyRpt("������������", Nvl(!����), Nvl(!��������), Nvl(!��������), Nvl(!������Ϣ), "��������", Nvl(!����˵��))
                   .MoveNext
               Loop
           End With
           mrsErrTable.Filter = "������־=0"
           With mrsErrTable
               '������������
               If mrsErrTable.RecordCount <> 0 Then .MoveFirst
               Do While Not .EOF
                   Call InputErrModifyRpt("δ�������Ķ���", Nvl(!����), Nvl(!��������), Nvl(!��������), Nvl(!������Ϣ), "δ����", "")
                   .MoveNext
               Loop
           End With
           
           mrsErrTable.Filter = "������־=4"
           With mrsErrTable
               '������������
               If mrsErrTable.RecordCount <> 0 Then .MoveFirst
               Do While Not .EOF
                   Call InputErrModifyRpt("��Ҫ�ֹ������Ķ���", Nvl(!����), Nvl(!��������), Nvl(!��������), Nvl(!������Ϣ), "�ֹ�����", Nvl(!����˵��))
                   .MoveNext
               Loop
           End With
           
           mrsErrTable.Filter = "������־=1"
           With mrsErrTable
               '������������
               If mrsErrTable.RecordCount <> 0 Then .MoveFirst
               Do While Not .EOF
                   Call InputErrModifyRpt("�����ɹ��Ķ���", Nvl(!����), Nvl(!��������), Nvl(!��������), Nvl(!������Ϣ), "�����ɹ�", Nvl(!����˵��))
                   .MoveNext
               Loop
           End With
    
           If frmAppChkRpt.hgdReport.Rows > 1 Then
               If MsgBox("�Ѿ�ִ����ȫ��ض��������," & vbCr & "�鿴��鱨����", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                   frmAppChkRpt.hgdReport.FixedRows = 1
                   frmAppChkRpt.Show 1
               End If
           Else
               MsgBox "����������ϣ����ٴ����С������顱��ȷ���Ƿ����δ��������", vbInformation, gstrSysName
           End If
           frmAppChkRpt.blnModiyfyCheck = False
        End If
        Enabled = True
    Case 2
        picStatus.Visible = True
        Enabled = False
        lblStatus.Caption = "�ؽ�������"
        blnTool = (cmbSystem.Tag = "ZLTOOLS")
        With rsTemp
            If gblnDBA Then
                strSQL = "select INDEX_NAME,STATUS from DBA_INDEXES where TABLESPACE_NAME is Not Null And Index_Type = 'NORMAL' And OWNER='" & cmbSystem.Tag & "' and TABLE_OWNER='" & cmbSystem.Tag & "'"
            Else
                strSQL = "select INDEX_NAME,STATUS from USER_INDEXES where TABLESPACE_NAME is Not Null And Index_Type = 'NORMAL' And TABLE_OWNER='" & cmbSystem.Tag & "'"
            End If
            If .State = adStateOpen Then .Close
            .Open strSQL, gcnOracle, adOpenKeyset
            Do While Not .EOF
                lblStatus.Caption = "�ؽ�������" & .Fields(0).value & "��"
                strSQL = "ALTER INDEX " & cmbSystem.Tag & "." & .Fields(0).value & " Rebuild nologging"
                gcnOldOra.Execute strSQL
                pgbState.value = .AbsolutePosition / .RecordCount * 100
                DoEvents
                .MoveNext
            Loop
        End With
        pgbState.value = 0
        picStatus.Visible = False
        Enabled = True
        MsgBox "�����ؽ���ϣ�", vbInformation, gstrSysName
    Case 3
        picStatus.Visible = True
        Enabled = False
        lblStatus.Caption = "���ڶ�ȡ���С�"
        
        Dim lngRealId As Long
        Dim lngNextId As Long
        
        Set rsObjects = GetSequence("", gcnOracle)
        With rsObjects
            Do Until .EOF
                lblStatus.Caption = "�������У�" & !Sequence_Name & "��"
                pgbState.value = .AbsolutePosition / .RecordCount * 100: DoEvents
                Call AdjustNameSequece(rsObjects!Owner & "." & rsObjects!Table_Name, gcnOracle, rsObjects!Column_Name)
                .MoveNext
            Loop
            
            Call Adjust����ID(gcnOracle)
        End With
        pgbState.value = 0
        picStatus.Visible = False
        Enabled = True
        
        MsgBox "����������ϣ�", vbInformation, gstrSysName
    '��������ͬ���
    Case 4
        '������ǰ�����ߵ�ȫ������Ĺ���ͬ���('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION')
        gcnOracle.Execute "Zl_Createpubsynonyms", , adCmdStoredProc
        
        MsgBox "��������ͬ�����ɣ�", vbInformation, gstrSysName
    Case 5
        Call frmHistorySpaceRepair.ShowRepair(Me, Val(cmbSystem.ItemData(cmbSystem.ListIndex)), False, , , True)
    Case 6
        '����Ȩ������
        Set cnTools = GetConnection("ZLTOOLS")
        If cnTools Is Nothing Then Exit Sub
        Call ReGrantForTools(cnTools, , True)
        MsgBox "����Ȩ��������ɣ�", vbInformation, gstrSysName
    End Select
End Sub



Private Sub cmdFunction_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngTemp As Long
    Dim i As Long
    
    For i = 0 To cmdFunction.Ubound
        If i = Index Then
            If cmdFunction(Index) Is ActiveControl And cmdFunction(Index).FontBold = True Then Exit Sub
            
            For lngTemp = 0 To cmdFunction.Ubound
                cmdFunction(lngTemp).FontBold = False
            Next
            cmdFunction(i).FontBold = True
            cmdFunction(i).SetFocus
            Select Case i
            Case 0
                lblNote.Caption = "    ������ϵͳ�����ݿ�����밲װ�ļ��Աȣ��������ϵͳ�����С�����ͼ���������洢���̡����ȶ������ȷ�ԡ�"
            Case 1
                lblNote.Caption = "    ���ݰ�װ�ļ��Զ����������½������С����ݱ���ͼ���������洢���̡��������ݶ���" & vbCrLf & _
                                  "    Ϊ���������������ݲ�������������ֶ����Ͳ�һ�»�Ӹ߾�����;��ȸı�����ݽ���������ͬʱҲ�Ͳ��ܱ�֤��Ч�������ж���" & vbCrLf & _
                                  "    �ù��ܽ�ϵͳ�����߿����С�"
            Case 2
                lblNote.Caption = "    ��ϵͳ���е���������(����������Լ����ΨһԼ��������������)��������ؽ�����(Rebuild)���Ա�֤��������Ч��"
            Case 3
                lblNote.Caption = "    �����������еĵ�ǰֵ����֤������ʵ��Ӧ�õ�ƥ�䣻" & vbCrLf & "    ��ϵͳ���֡���(ID)�����ظ�����һ�����ʱ��һ���ʹ�ñ�����������⡣"
            Case 4
                lblNote.Caption = "    ���ݵ�ǰ�����ߵı���ͼ�����̣����������еȶ��󴴽�ͬ������ͬ��ʣ��������ͬ������ͬ����򲻴�����"
            Case 5
                lblNote.Caption = "    ���ݵ�ǰϵͳ��ת�����壬�����ʷ�������߿��ڱ��С�Լ���������ȷ����һ���ԣ����ṩ�޸����ܡ�"
            Case 6
                lblNote.Caption = "    �Թ����ߵĶ������������Ȩ������ͬ��ʡ�"
            End Select
        End If
    Next
End Sub

Private Sub cmdGetIni_Click()
    Dim varData As Variant
    Dim strToolsByServer As String
    Dim strToolsVer As String
    
    With frmMDIMain.dlgMain
        .FileName = lblFileName.Caption
        .DialogTitle = "ѡ��Ӧ�ð�װ�����ļ�"
        If cmbSystem.Tag = "ZLTOOLS" Then
            .Filter = "(���������߽ű�)|zlServer.Sql"
        Else
            .Filter = "(Ӧ�ð�װ�����ļ�)|zlSetup.ini"
        End If
        .ShowOpen
        If .FileName = "" Then
            Exit Sub
        Else
            lblFileName.Caption = .FileName
        End If
    End With
    '���ù���״̬
    Call SetFunsState(lblFileName.Caption, True)
    
End Sub

Private Sub Form_Load()
    
    lblMain.Caption = "ͨ�������������Ĳ�������Ҫ�ϳ�ʱ��Ĳ���������ϵͳ�Ѿ���ȷ�ر�¶�����⣬�벻Ҫ����ʹ�ô˹��ܡ�" & _
        vbCrLf & vbCrLf & "���������������й��ܶ��漰����Ķ�ռ�������������ȷ���Ѿ�û�������û���ʹ��ϵͳ��ǧ��Ҫ���У������ϵͳ����ƻ�������Ͽ�����������������������ӣ���"
        
    Call LoadSystem
End Sub
Private Function LoadSystem() As Boolean
    '-------------------------------------------------------------------------------------------------
    '����:����ϵͳ����
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/09/10
    '-------------------------------------------------------------------------------------------------
    Dim strToolsVer As String
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long
    
    LoadSystem = False
    err = 0: On Error GoTo errHand:
    LoadSystem = True
    
    Set rsTemp = OpenCursor(gcnOracle, "zlTools.B_Public.Get_Ver")
    strToolsVer = rsTemp!����
    '��д�Ѱ�װϵͳ�嵥
    If gblnDBA Then
        Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
    Else
        Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", gstrUserName)
    End If

    With rsTemp
        Do While Not .EOF
            cmbSystem.AddItem !���� & " v" & !�汾�� & "��" & !��� & "��"
            cmbSystem.ItemData(cmbSystem.NewIndex) = !���
            .MoveNext
        Loop
        cmbSystem.AddItem "������������" & " v" & strToolsVer & "��ZLTOOLS��"
        cmbSystem.ItemData(cmbSystem.NewIndex) = -1
        If cmbSystem.ListCount = 0 Then
            cmdGetIni.Enabled = False
            For i = 0 To cmdFunction.Ubound
                cmdFunction(i).Enabled = False
            Next
        End If
        If cmbSystem.ListCount > 0 Then cmbSystem.ListIndex = 0
        If cmbSystem.ListCount = 1 Then cmbSystem.Locked = True
    End With
    '���ع�����
    Exit Function
errHand:
    cmbSystem.AddItem "������������" & " v" & strToolsVer & "��ZLTOOLS��"
    cmbSystem.ItemData(cmbSystem.NewIndex) = -1
    MsgBox err.Description, vbCritical, Me.Caption
End Function
Private Sub Form_Resize()
    On Error Resume Next
    Dim sngWidth As Long '��С���
    
    With imgMain
        .Top = 700
        .Left = ScaleLeft + 200
    End With
        
    With lblMain
        .Top = lblNote.Top + lblNote.Height + 300
        .Height = ScaleHeight - picStatus.Height - .Top - 100
        .Left = lblFileName.Left
        .Width = ScaleWidth - .Left - imgMain.Left
    End With
    
    sngWidth = IIf(ScaleWidth < 5600, 5600, ScaleWidth)
    cmbSystem.Width = sngWidth - cmbSystem.Left - 300
    cmdGetIni.Left = cmbSystem.Left + cmbSystem.Width - cmdGetIni.Width
    lblFileName.Width = sngWidth - lblFileName.Left - 300
    lblNote.Width = sngWidth - lblNote.Left - 300
    
End Sub


Private Function CheckIniFile(FileName As String, Optional blnMsg As Boolean) As Boolean
    Dim strTemp As String
    Dim objText As TextStream
    Dim intDefSysCode As Integer                'ϵͳ���
    Dim strDefSysName As String                  'ϵͳ����
    Dim strDefVersion As String                 '�汾��
    Dim strDefSpace   As String                 '��ռ�
    
    err = 0
    On Error Resume Next
    
    mstrIniPath = Mid(FileName, 1, Len(FileName) - 11)
    '����ļ�ƥ���Լ��
    strTemp = ""
    If Dir(mstrIniPath & "zlSequence.sql") = "" Then strTemp = strTemp & vbCr & "�����ļ�" & mstrIniPath & "zlSequence.sql"
    If Dir(mstrIniPath & "zlTable.sql") = "" Then strTemp = strTemp & vbCr & "���ݱ��ļ�" & mstrIniPath & "zlTable.sql"
    If Dir(mstrIniPath & "zlConstraint.sql") = "" Then strTemp = strTemp & vbCr & "Լ���ļ�" & mstrIniPath & "zlConstraint.sql"
    If Dir(mstrIniPath & "zlIndex.sql") = "" Then strTemp = strTemp & vbCr & "�����ļ�" & mstrIniPath & "zlIndex.sql"
    If Dir(mstrIniPath & "zlView.sql") = "" Then strTemp = strTemp & vbCr & "��ͼ�ļ�" & mstrIniPath & "zlView.sql"
    If Dir(mstrIniPath & "zlProgram.sql") = "" Then strTemp = strTemp & vbCr & "���������ļ�" & mstrIniPath & "zlProgram.sql"
    
    '�����,��Ϊ9ϵͳû�д��ļ�
    'If Dir(mstrIniPath & "zlPackage.sql") = "" Then strTemp = strTemp & vbCr & "���ļ�" & mstrIniPath & "zlPackage.sql"
    
    If Dir(mstrIniPath & "zlManData.sql") = "" Then strTemp = strTemp & vbCr & "���������ļ�" & mstrIniPath & "zlManData.sql"
    If Dir(mstrIniPath & "zlAppData.sql") = "" Then strTemp = strTemp & vbCr & "Ӧ�������ļ�" & mstrIniPath & "zlAppData.sql"
    If strTemp <> "" Then
        If blnMsg Then MsgBox "���·�������װ������ļ���ʧ�����ܼ�����������" & strTemp, vbExclamation, gstrSysName
        Exit Function
    End If
    
    '�����ļ���ȷ�Լ��
    Set objText = gobjFile.OpenTextFile(FileName)
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[ϵͳ��]" Then
        intDefSysCode = Mid(strTemp, 6)
    Else
        err.Raise 10
    End If
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[ϵͳ��]" Then
        strDefSysName = Mid(strTemp, 6)
    Else
        err.Raise 10
    End If
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[�汾��]" Then
        strDefVersion = Mid(strTemp, 6)
    Else
        err.Raise 10
    End If
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[��ռ�]" Then
        strDefSpace = Mid(strTemp, 6)
    Else
        err.Raise 10
    End If
    objText.Close
    
    If err <> 0 Then
        CheckIniFile = False
        If blnMsg Then MsgBox "��װ�����ļ�����ȷ", vbExclamation, gstrSysName
        Exit Function
    End If
    
    '�����ļ������Լ��
    If intDefSysCode <> cmbSystem.ItemData(cmbSystem.ListIndex) \ 100 Then
        err.Raise 10
        If blnMsg Then MsgBox "ѡ���ļ����Ǹ�ϵͳ�İ�װ�����ļ�", vbExclamation, gstrSysName
    ElseIf InStr(1, cmbSystem.Text, Trim(strDefVersion)) = 0 Then
        err.Raise 10
        If blnMsg Then MsgBox "ѡ���ļ����ϵͳ�汾����", vbExclamation, gstrSysName
    End If
    If err = 0 Then
        CheckIniFile = True
    Else
        CheckIniFile = False
    End If
End Function

Private Sub InputErrRpt(ObjType As String, ObjName As String, ErrInfo As String, Optional Advice As String)
    '----------------------------------------------------
    '��дһ�д��󱨸�
    '----------------------------------------------------
    With frmAppChkRpt.hgdReport
        .Rows = .Rows + 1
        If .Tag <> ObjType Then
            .TextMatrix(.Rows - 1, 0) = "------< " & ObjType & "������ >------"
            .TextMatrix(.Rows - 1, 1) = "------< " & ObjType & "������ >------"
            .MergeRow(.Rows - 1) = True
            .Rows = .Rows + 1
        End If
        .Tag = ObjType
        .TextMatrix(.Rows - 1, 0) = ObjName
        If .ColData(0) < Me.TextWidth(ObjName) Then
            .ColData(0) = Me.TextWidth(ObjName)
        End If
        .TextMatrix(.Rows - 1, 1) = ErrInfo
        If .ColData(1) < Me.TextWidth(ErrInfo) Then
            .ColData(1) = Me.TextWidth(ErrInfo)
        End If
        .TextMatrix(.Rows - 1, 2) = Advice
        If .ColData(2) < Me.TextWidth(Advice) Then
            .ColData(2) = Me.TextWidth(Advice)
        End If
            
    End With
        
    '��ʾ���ѱ�ǩ
    If InStr(Advice, "����") > 0 Or InStr(Advice, "����") > 0 Then
         frmAppChkRpt.lblWarn.Visible = True
    End If
End Sub


Private Sub InputErrModifyRpt(ObjType As String, ObjName As String, strTableName As String, strErrType As String, ErrInfo As String, strModifyInfor As String, Optional strModifyErrInfor As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��дָ���еĶ����������
    '���:
    '����:
    '����:
    '����:22507
    '����:���˺�
    '����:2009-08-26 14:34:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With frmAppChkRpt.hgdReport
        .Rows = .Rows + 1
        If .Tag <> ObjType Then
            .TextMatrix(.Rows - 1, 0) = "------< " & ObjType & " >------"
            .TextMatrix(.Rows - 1, 1) = "------< " & ObjType & " >------"
            .TextMatrix(.Rows - 1, 2) = "------< " & ObjType & " >------"
            .TextMatrix(.Rows - 1, 3) = "------< " & ObjType & " >------"
            .TextMatrix(.Rows - 1, 4) = "------< " & ObjType & " >------"
            .TextMatrix(.Rows - 1, 5) = "------< " & ObjType & " >------"
            .MergeRow(.Rows - 1) = True
            .Rows = .Rows + 1
        End If
        .Tag = ObjType: i = 0
        .TextMatrix(.Rows - 1, i) = ObjName
        If .ColData(i) < Me.TextWidth(ObjName) Then .ColData(i) = Me.TextWidth(ObjName)
        
        i = i + 1: .TextMatrix(.Rows - 1, i) = strTableName
        If .ColData(i) < Me.TextWidth(strTableName) Then .ColData(i) = Me.TextWidth(strTableName)
        
        
        i = i + 1: .TextMatrix(.Rows - 1, i) = strErrType
        If .ColData(i) < Me.TextWidth(strErrType) Then .ColData(i) = Me.TextWidth(strErrType)
        
        
        i = i + 1: .TextMatrix(.Rows - 1, i) = ErrInfo
        If .ColData(i) < Me.TextWidth(ErrInfo) Then .ColData(i) = Me.TextWidth(ErrInfo)
        
        i = i + 1: .TextMatrix(.Rows - 1, i) = strModifyInfor
        If .ColData(i) < Me.TextWidth(strModifyInfor) Then .ColData(i) = Me.TextWidth(strModifyInfor)
            
        i = i + 1: .TextMatrix(.Rows - 1, i) = strModifyErrInfor
        If .ColData(i) < Me.TextWidth(strModifyErrInfor) Then .ColData(i) = Me.TextWidth(strModifyErrInfor)
        
    End With
End Sub

Private Sub CheckTable(FileName As String, Optional ByVal blnAddToRsTable As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ݱ�ͬʱ�ж����ݱ�����Ƿ���ȷ
    '���:FileName-�ű��ļ���(��������·��)
    '     blnAddToRsTable-�Ƿ���صĴ�����Ϣд�ڼ�¼����
    '����:
    '����:
    '����:���˺�
    '����:2009-08-19 15:03:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strԭ�ֶξ��� As String, str���ֶξ��� As String, strSQL1 As String
    Dim arySql() As String, strObjName As String
    Dim intVer As Integer
    Dim rsObjects As New ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    Dim rsColumns As New ADODB.Recordset
    
    intVer = GetOracleVersion

    With rsObjects
        .Filter = 0
        If gblnDBA Then
            strSQL = "select TABLE_NAME from DBA_TABLES where OWNER='" & cmbSystem.Tag & "' And Instr(Table_Name, 'BIN$') <= 0"
        Else
            strSQL = "select TABLE_NAME from USER_TABLES where Instr(Table_Name, 'BIN$') <= 0"
        End If
        If .State = adStateOpen Then .Close
        .Open strSQL, gcnOracle, adOpenKeyset
        
        err = 0
        On Error Resume Next
        If Not mclsRunScript.OpenFile(FileName) Then Exit Sub
        On Error GoTo 0
        Do While Not mclsRunScript.EOF
            strSQL = UCase(TrimEx(mclsRunScript.SQLInfo.SQL))
            strSQL1 = strSQL
            arySql = Split(strSQL, " TABLE ")
            strSQL = Trim(arySql(1)) '�Ѿ�ȥ��Oracle�ؼ���
            If InStr(1, strSQL, " ") > 0 And InStr(strSQL, " ") < InStr(strSQL, "(") Then
                strObjName = Trim(Left(strSQL, InStr(strSQL, " ")))
            Else
                strObjName = Trim(Left(strSQL, InStr(strSQL, "(") - 1))
            End If
            strObjName = Replace(strObjName, vbCrLf, "")
            .Filter = "TABLE_NAME='" & strObjName & "'"
            
            If .EOF Then
                If blnAddToRsTable Then
                    '1-���ڶ���,2-�����ڶ���,3-ʧЧ
                     Call zlInsertRecData(mrsErrTable, strObjName, strObjName, "���ݱ�", 2, False, strSQL1, "������", "���أ����ֹ��ܲ�����������")
                Else
                    Call InputErrRpt("���ݱ�", strObjName, "������", "���أ����ֹ��ܲ�����������")
                End If
            Else
                'ͨ������һ����׼�ṹ����������ֶνṹ�����ж�
                strSQL = arySql(0) & " table CK" & Trim(arySql(1))
                strSQL = Split(strSQL, "TABLESPACE")(0)
                On Error Resume Next
                'gcnOracle.Execute "drop table CK" & strObjName & IIf(intVer >= 100, " Purge", "")
                gcnOldOra.Execute "drop table CK" & strObjName & IIf(intVer >= 100, " Purge", "")
                err = 0
                err.Clear
                '������֯��,�����ڴ������ͬʱ���������������������⴦��
                If InStr(UCase(strSQL), "PRIMARY") > 0 Then
                    strSQL = Replace(strSQL, strObjName & "_PK", "CK" & strObjName & "_PK")
                End If
                
                gcnOldOra.Execute strSQL
                If err = 0 Then
                    With rsColumns
                        If gblnDBA Then
                            strTemp = "SELECT COLUMN_NAME,DATA_TYPE,DATA_LENGTH,DATA_PRECISION,DATA_SCALE,DATA_DEFAULT" & _
                                    " From DBA_TAB_COLUMNS" & _
                                    " WHERE OWNER='" & cmbSystem.Tag & "' and TABLE_NAME='" & strObjName & "'"
                        Else
                            strTemp = "SELECT COLUMN_NAME,DATA_TYPE,DATA_LENGTH,DATA_PRECISION,DATA_SCALE,DATA_DEFAULT" & _
                                "       From USER_TAB_COLUMNS" & _
                                "       WHERE TABLE_NAME='" & strObjName & "'"
                        End If
                        strTemp = "select N.COLUMN_NAME as N_NAME,N.DATA_TYPE as N_TYPE,N.DATA_LENGTH as N_NLENGTH," & _
                                "        N.DATA_PRECISION as N_PRECISION,N.DATA_SCALE as N_SCALE,N.DATA_DEFAULT as N_DEFAULT," & _
                                "        O.COLUMN_NAME as O_NAME,O.DATA_TYPE as O_TYPE,O.DATA_LENGTH as O_NLENGTH," & _
                                "        O.DATA_PRECISION as O_PRECISION,O.DATA_SCALE as O_SCALE,O.DATA_DEFAULT as O_DEFAULT" & _
                                " from (SELECT COLUMN_NAME,DATA_TYPE,DATA_LENGTH,DATA_PRECISION,DATA_SCALE,DATA_DEFAULT" & _
                                "       From USER_TAB_COLUMNS" & _
                                "       WHERE TABLE_NAME='CK" & strObjName & "') N," & _
                                "      (" & strTemp & ") O" & _
                                " where N.COLUMN_NAME=O.COLUMN_NAME(+) "    'and N.DATA_TYPE=O.DATA_TYPE(+):���˺�:2007/06/30���ܴ����ֶ����ͷ����仯�����,�����Ҫȡ�����
                        If .State = adStateOpen Then .Close
                        .Open strTemp, gcnOracle, adOpenKeyset
                        
                        Do While Not .EOF
                            strԭ�ֶξ��� = "": str���ֶξ��� = ""
                            If IsNull(!O_TYPE) Then
                                'ȱ�ٵ��ֶ�
                                Select Case !N_TYPE
                                Case "NUMBER"
                                    strTemp = !N_NAME & " NUMBER(" & !N_PRECISION & "," & !N_SCALE & ")"
                                    If Not IsNull(!N_DEFAULT) Then strTemp = strTemp & " DEFAULT " & !N_DEFAULT
                                    str���ֶξ��� = !N_PRECISION & "," & !N_SCALE
                                Case "VARCHAR2"
                                    strTemp = !N_NAME & " VARCHAR2(" & !N_NLENGTH & ")"
                                    If Not IsNull(!N_DEFAULT) Then strTemp = strTemp & " DEFAULT " & !N_DEFAULT
                                     str���ֶξ��� = Nvl(!N_NLENGTH)
                                Case Else
                                    strTemp = !N_NAME & !N_TYPE
                                End Select
                                If blnAddToRsTable Then
                                    '1-���ڶ���,2-�����ڶ���,3-ʧЧ,4-ȱ����
                                     strSQL1 = " Alter Table " & strObjName & " Add(" & strTemp & ")"
                                     Call zlInsertRecData(mrsErrTable, strObjName, strObjName, "���ݱ�", 4, False, strSQL1, "ȱ���� " & strTemp, "���أ����ֹ��ܲ�����������", Nvl(!N_NAME), "", Nvl(!N_TYPE), strԭ�ֶξ���, str���ֶξ���)
                                Else
                                    Call InputErrRpt("���ݱ�", strObjName, "ȱ���� " & strTemp, "���أ����ֹ��ܲ�����������")
                                End If
                            Else
                                '���Ȳ���
                                Select Case !N_TYPE
                                Case "NUMBER"
                                    If !N_PRECISION > !O_PRECISION Or !N_SCALE > !O_SCALE Then
                                        strTemp = !N_NAME & "�г���С�ڹ涨ֵ��ӦΪ��" & "NUMBER(" & !N_PRECISION & "," & !N_SCALE & ")��" & _
                                                 " ��Ϊ��" & "NUMBER(" & !O_PRECISION & "," & !O_SCALE & ")��"
                                                 
                                        strԭ�ֶξ��� = !O_PRECISION & "," & !O_SCALE: str���ֶξ��� = !N_PRECISION & "," & !N_SCALE
                                        If blnAddToRsTable Then
                                            If Nvl(!O_TYPE) = Nvl(!N_TYPE) Then
                                                '������ͬ������£��Ŵ���
                                                '1-���ڶ���,2-�����ڶ���,3-ʧЧ,4-ȱ����,5-��������
                                                 strSQL1 = " Alter Table " & strObjName & " Modify(" & !N_NAME & " NUMBER(" & !N_PRECISION & IIf(Val(Nvl(!N_SCALE)) = 0, "", "," & Nvl(!N_SCALE) & ")") & "))"
                                                If !N_PRECISION > !O_PRECISION Then
                                                    Call zlInsertRecData(mrsErrTable, strObjName, strObjName, "���ݱ�", 5, False, strSQL1, "���ȹ�С ", "���أ��ϴ�����ݽ��޷���ȷ�洢��" & strTemp, Nvl(!N_NAME), Nvl(!N_TYPE), Nvl(!N_TYPE), strԭ�ֶξ���, str���ֶξ���)
                                                ElseIf !N_SCALE > !O_SCALE Then
                                                    Call zlInsertRecData(mrsErrTable, strObjName, strObjName, "���ݱ�", 5, False, strSQL1, "���ȹ�С ", "���أ����ܵ������ݾ��Ȳ��㣺" & strTemp, Nvl(!N_NAME), Nvl(!N_TYPE), Nvl(!N_TYPE), strԭ�ֶξ���, str���ֶξ���)
                                                Else
                                                    Call InputErrRpt("���ݱ�", strObjName, strTemp, "���᣺������Ӱ������")
                                                    Call zlInsertRecData(mrsErrTable, strObjName, strObjName, "���ݱ�", 5, False, strSQL1, "���ȹ��� ", "���᣺������Ӱ�����У�" & strTemp, Nvl(!N_NAME), Nvl(!N_TYPE), Nvl(!N_TYPE), strԭ�ֶξ���, str���ֶξ���)
                                                End If
                                            End If
                                        Else
                                            If !N_PRECISION > !O_PRECISION Then
                                                Call InputErrRpt("���ݱ�", strObjName, strTemp, "���أ��ϴ�����ݽ��޷���ȷ�洢")
                                            ElseIf !N_SCALE > !O_SCALE Then
                                                Call InputErrRpt("���ݱ�", strObjName, strTemp, "���أ����ܵ������ݾ��Ȳ���")
                                            Else
                                                Call InputErrRpt("���ݱ�", strObjName, strTemp, "���᣺������Ӱ������")
                                            End If
                                        End If
                                                 
                                    End If
                                Case "VARCHAR2"
                                    If !N_NLENGTH <> !O_NLENGTH Then
                                        strTemp = !N_NAME & "�г���С�ڹ涨ֵ��ӦΪ��" & "VARCHAR2(" & !N_NLENGTH & ")��" & _
                                                 " ��Ϊ��" & "VARCHAR2(" & !O_NLENGTH & ")��"
                                        strԭ�ֶξ��� = !O_NLENGTH: str���ֶξ��� = !N_NLENGTH
                                        If blnAddToRsTable Then
                                            If Nvl(!O_TYPE) = Nvl(!N_TYPE) Then
                                                '������ͬ������£��Ŵ���
                                                '1-���ڶ���,2-�����ڶ���,3-ʧЧ,4-ȱ����,5-��������
                                                 strSQL1 = " Alter Table " & strObjName & " Modify(" & !N_NAME & " VARCHAR2(" & !N_NLENGTH & ")" & ")"
                                                
                                                If !N_NLENGTH > !O_NLENGTH Then
                                                    Call zlInsertRecData(mrsErrTable, strObjName, strObjName, "���ݱ�", 5, False, strSQL1, "���ȹ�С", "���أ����ܵ��½ϳ��ı��޷��洢��" & strTemp, Nvl(!N_NAME), Nvl(!N_TYPE), Nvl(!N_TYPE), strԭ�ֶξ���, str���ֶξ���)
                                                Else
                                                    Call zlInsertRecData(mrsErrTable, strObjName, strObjName, "���ݱ�", 5, False, strSQL1, "���ȹ���", "���᣺������Ӱ�����У�" & strTemp, Nvl(!N_NAME), Nvl(!N_TYPE), Nvl(!N_TYPE), strԭ�ֶξ���, str���ֶξ���)
                                                End If
                                            End If
                                        Else
                                            If !N_NLENGTH > !O_NLENGTH Then
                                                Call InputErrRpt("���ݱ�", strObjName, strTemp, "���أ����ܵ��½ϳ��ı��޷��洢")
                                            Else
                                                Call InputErrRpt("���ݱ�", strObjName, strTemp, "���᣺������Ӱ������")
                                            End If
                                        End If
                                    End If
                                Case Else
                                End Select
                                '���˺�:2007/06/30������������ж�
                                If Nvl(!O_TYPE) <> Nvl(!N_TYPE) Then
                                     strTemp = !N_NAME & "�����Ͳ�һ����ӦΪ��" & Nvl(!N_TYPE) & "�� ��Ϊ����" & Nvl(!O_TYPE) & "��"
                                     If blnAddToRsTable Then
                                        Call zlInsertRecData(mrsErrTable, strObjName, strObjName, "���ݱ�", 5, False, "", "�ֶ����Ͳ���", "���أ����ܵ������ݲ��ܴ洢!" & strTemp, Nvl(!N_NAME), Nvl(!O_TYPE), Nvl(!N_TYPE), "", "")
                                     Else
                                        Call InputErrRpt("���ݱ�", strObjName, strTemp, "���أ����ܵ������ݲ��ܴ洢!")
                                     End If
                                End If
                            End If
                            
                            .MoveNext
                        Loop
                    End With
                    gcnOracle.Execute "drop table CK" & strObjName & IIf(intVer >= 100, " Purge", "")
                Else
                    '����CK��ʧ��
                    '������־,�����û�,�ñ�δ���
                    If blnAddToRsTable = False Then
                        Call InputErrRpt("���ݱ�", strObjName, "�����Ƚ϶���ʧ��", "���أ����ܶԴ˶������ȷ�������жϣ�")
                    End If
                End If
            End If
            
            pgbState.value = mclsRunScript.Line / mclsRunScript.LinesCount * 100
            Call mclsRunScript.ReadNextSQL
            DoEvents
        Loop
    End With
    pgbState.value = 0
End Sub

Private Sub CheckConstraint(FileName As String, Optional ByVal blnAddToRsTable As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Լ�����ж��Ƿ���Ч����
    '���:FileName-�ű��ļ���(��������·��)
    '     blnAddToRsTable-�Ƿ���صĴ�����Ϣд�ڼ�¼����
    '����:
    '����:
    '����:���˺�
    '����:2009-08-19 15:03:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsObjects As New ADODB.Recordset, strSQL As String
    Dim rsColumns As New ADODB.Recordset, strTemp As String
    
    
    Dim strԭ�ֶξ��� As String, str���ֶξ��� As String, strSQL1 As String, strTemp1 As String
    Dim arySql() As String, strObjName As String, strColumns As String, strTableName As String
    With rsObjects
        .Filter = 0
        If gblnDBA Then
            strSQL = "select CONSTRAINT_TYPE,CONSTRAINT_NAME,STATUS,VALIDATED,BAD,Search_Condition from DBA_CONSTRAINTS where OWNER='" & cmbSystem.Tag & "'"
        Else
            strSQL = "select CONSTRAINT_TYPE,CONSTRAINT_NAME,STATUS,VALIDATED,BAD,Search_Condition from USER_CONSTRAINTS"
        End If
        If .State = adStateOpen Then .Close
        .Open strSQL, gcnOracle, adOpenKeyset
    
        err = 0
        On Error Resume Next
        If Not mclsRunScript.OpenFile(FileName) Then Exit Sub
        On Error GoTo 0
        Do While Not mclsRunScript.EOF
            strSQL = mclsRunScript.SQLInfo.PartSQL
            strSQL1 = strSQL
            arySql = Split(strSQL, " CONSTRAINT ")
            If UBound(arySql) > 0 Then
                '���˺����:
                strTableName = Trim(Split(Trim(arySql(0)), "TABLE")(1))
                strTableName = Split(strTableName, " ")(0)
                strSQL = Trim(arySql(1)) '�Ѿ�ȥ��Oracle�ؼ���
                strObjName = Trim(Left(strSQL, InStr(strSQL, " ")))
                .Filter = "CONSTRAINT_NAME='" & strObjName & "'"
                If .EOF Then
                    If blnAddToRsTable Then
                        '1-���ڶ���,2-�����ڶ���,3-ʧЧ
                        '����:����,Ψһ,���,Լ��,����,��ͼ,...
                        If InStr(1, UCase(Replace(strSQL, " ", "")), UCase("ForeignKey")) > 0 Then
                            Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "���", 2, False, strSQL1, "������", "���أ����ܵ������ݲ�һ�£�Ӱ�������ٶ�")
                        ElseIf InStr(1, strSQL, " CHECK") > 0 Then
                            Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "Լ��", 2, False, strSQL1, "������", "���᣺������Ӱ��ϵͳ����")
                        Else
                            Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "Լ��", 2, False, strSQL1, "������", "���أ����ܵ������ݲ�һ�£�Ӱ�������ٶ�")
                        End If
                    Else
                        If InStr(1, strSQL, " CHECK") > 0 Then
                            Call InputErrRpt("Լ��", strObjName, "������", "���᣺������Ӱ��ϵͳ����")
                        Else
                            Call InputErrRpt("Լ��", strObjName, "������", "���أ����ܵ������ݲ�һ�£�Ӱ�������ٶ�")
                        End If
                    End If
                ElseIf .Fields("STATUS").value <> "ENABLED" Then
                    If blnAddToRsTable Then
                        '״̬:1-���ڶ���,2-�����ڶ���,3-ʧЧ,4-ȱ����,5-����,6-���ڽ�ֹ״̬
                        '����:����,Ψһ,���,Լ��,����,��ͼ,...
                        If InStr(1, UCase(Replace(strSQL, " ", "")), UCase("ForeignKey")) > 0 Then
                            Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "���", 6, False, strSQL1, "��ǰ���ڽ�ֹ״̬", "���أ�����ϵͳ�Ѿ���������")
                        Else
                            Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "Լ��", 6, False, strSQL1, "��ǰ���ڽ�ֹ״̬", "���أ�����ϵͳ�Ѿ���������")
                        End If
                    Else
                        Call InputErrRpt("Լ��", strObjName, "��ǰ���ڽ�ֹ״̬", "���أ�����ϵͳ�Ѿ���������")
                    End If
                ElseIf !VALIDATED <> "VALIDATED" Then
                    If blnAddToRsTable Then
                        '״̬:1-���ڶ���,2-�����ڶ���,3-ʧЧ,4-ȱ����,5-����,6-���ڽ�ֹ״̬
                        '����:����,Ψһ,���,Լ��,����,��ͼ,...
                        If InStr(1, UCase(Replace(strSQL, " ", "")), UCase("ForeignKey")) > 0 Then
                            Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "���", 3, False, strSQL1, "��ǰ������Ч״̬", "���أ���������һ�����ѱ��ƻ�")
                        Else
                            Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "Լ��", 3, False, strSQL1, "��ǰ������Ч״̬", "���أ���������һ�����ѱ��ƻ�")
                        End If
                    Else
                        Call InputErrRpt("Լ��", strObjName, "��ǰ������Ч״̬", "���أ���������һ�����ѱ��ƻ�")
                    End If
                ElseIf Not IsNull(!BAD) Then
                    If blnAddToRsTable Then
                        '״̬:1-���ڶ���,2-�����ڶ���,3-ʧЧ,4-ȱ����,5-����,6-���ڽ�ֹ״̬(��)
                        '����:����,Ψһ,���,Լ��,����,��ͼ,...
                        If InStr(1, UCase(Replace(strSQL, " ", "")), UCase("ForeignKey")) > 0 Then
                            Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "���", 6, False, strSQL1, "Լ����������", "���أ����ܴ���Ӳ������")
                        Else
                            Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "Լ��", 6, False, strSQL1, "Լ����������", "���أ����ܴ���Ӳ������")
                        End If
                    Else
                        Call InputErrRpt("Լ��", strObjName, "Լ����������", "���أ����ܴ���Ӳ������")
                    End If
                Else
                    strColumns = ""
                    With rsColumns
                        If gblnDBA Then
                            strTemp = "select COLUMN_NAME" & _
                                " from DBA_CONS_COLUMNS" & _
                                " where OWNER='" & cmbSystem.Tag & "' and CONSTRAINT_NAME='" & strObjName & "'" & _
                                " order by POSITION"
                        Else
                            strTemp = "select COLUMN_NAME" & _
                                " from USER_CONS_COLUMNS" & _
                                " where CONSTRAINT_NAME='" & strObjName & "'" & _
                                " order by POSITION"
                        End If
                        If .State = adStateOpen Then .Close
                        .Open strTemp, gcnOracle, adOpenKeyset
                        Do While Not .EOF
                            strColumns = strColumns & "," & !Column_Name
                            .MoveNext
                        Loop
                    End With
                    
                    If InStr(1, strSQL, " PRIMARY ") > 0 Then
                        If !constraint_type <> "P" Then
                            If blnAddToRsTable Then
                                '״̬:'1-���ڶ���,2-�����ڶ���,3-ʧЧ,4-ȱ����,5-����,6-���ڽ�ֹ״̬,7-Լ����һ��
                                '����:����,Ψһ,���,Լ��,����,��ͼ,...
                                Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "Լ��", 7, False, strSQL1, "Լ�����ʹ���", "���أ�����Ӱ��ϵͳ���� ,ӦΪ����Լ��")
                                If !constraint_type = "U" Then
                                    'ͬʱ����Ҫ�����صļ������
                                    Call zl��ȡ�������(strObjName, mrsErrTable)
                                End If
                            Else
                                Call InputErrRpt("Լ��", strObjName, "Լ�����ʹ���ӦΪ����Լ��", "���أ�����Ӱ��ϵͳ����")
                            End If
                        Else
                            arySql = Split(strSQL, " PRIMARY ")
                            strTemp = Replace(Replace(Replace(Left(arySql(1), InStr(1, arySql(1), ")") - 1), "KEY", ""), "(", ""), " ", "")
                            If strColumns <> "," & strTemp Then
                                If blnAddToRsTable Then
                                    '״̬:'1-���ڶ���,2-�����ڶ���,3-ʧЧ,4-ȱ����,5-����,6-���ڽ�ֹ״̬,7-Լ����һ��
                                    '����:����,Ψһ,���,Լ��,����,��ͼ,...
                                    Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "Լ��", 4, False, strSQL1, "Լ���д���", "���أ�����Ӱ��ϵͳ���У�ӦΪ(" & strTemp & ")����Ϊ(" & Mid(strColumns, 2) & ")")
                                    'ͬʱ����Ҫ�����صļ������
                                    Call zl��ȡ�������(strObjName, mrsErrTable)
                                    
                                Else
                                    Call InputErrRpt("Լ��", strObjName, "Լ���д���ӦΪ(" & strTemp & ")����Ϊ(" & Mid(strColumns, 2) & ")", "���أ�����Ӱ��ϵͳ����")
                                End If
                            End If
                        End If
                    ElseIf InStr(1, strSQL, " UNIQUE") > 0 Then
                        If !constraint_type <> "U" Then
                            If blnAddToRsTable Then
                                '״̬:'1-���ڶ���,2-�����ڶ���,3-ʧЧ,4-ȱ����,5-����,6-���ڽ�ֹ״̬,7-Լ����һ��
                                '����:����,Ψһ,���,Լ��,����,��ͼ,...
                                Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "Լ��", 7, False, strSQL1, "Լ�����ʹ���", "���أ�����Ӱ��ϵͳ���� ,ӦΪΨһԼ��")
                                If !constraint_type = "P" Then
                                    'ͬʱ����Ҫ�����صļ������
                                    Call zl��ȡ�������(strObjName, mrsErrTable)
                                End If
                            Else
                                Call InputErrRpt("Լ��", strObjName, "Լ�����ʹ���ӦΪΨһԼ��", "���أ�����Ӱ��ϵͳ����")
                            End If
                        Else
                            arySql = Split(strSQL, " UNIQUE ")
                            If UBound(arySql) = 0 Then arySql = Split(strSQL, " UNIQUE(")
                            strTemp = Replace(Replace(Left(arySql(1), InStr(1, arySql(1), ")") - 1), "(", ""), " ", "")
                            If strColumns <> "," & strTemp Then
                                If blnAddToRsTable Then
                                    '״̬:'1-���ڶ���,2-�����ڶ���,3-ʧЧ,4-ȱ����,5-����,6-���ڽ�ֹ״̬,7-Լ����һ��
                                    '����:����,Ψһ,���,Լ��,����,��ͼ,...
                                    Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "Լ��", 4, False, strSQL1, "Լ���д���", "���أ�����Ӱ��ϵͳ���У�ӦΪ(" & strTemp & ")����Ϊ(" & Mid(strColumns, 2) & ")")
                                    'ͬʱ����Ҫ�����صļ������
                                    Call zl��ȡ�������(strObjName, mrsErrTable)
                                Else
                                    Call InputErrRpt("Լ��", strObjName, "Լ���д���ӦΪ(" & strTemp & ")����Ϊ(" & Mid(strColumns, 2) & ")", "���أ�����Ӱ��ϵͳ����")
                                End If
                            End If
                        End If
                    ElseIf InStr(1, strSQL, " FOREIGN ") > 0 Then
                        If !constraint_type <> "R" Then
                            If blnAddToRsTable Then
                                '״̬:'1-���ڶ���,2-�����ڶ���,3-ʧЧ,4-ȱ����,5-����,6-���ڽ�ֹ״̬,7-Լ����һ��
                                '����:����,Ψһ,���,Լ��,����,��ͼ,...
                                Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "Լ��", 7, False, strSQL1, "Լ�����ʹ���", "���أ�����Ӱ��ϵͳ���У�ӦΪ���Լ��")
                            Else
                                Call InputErrRpt("Լ��", strObjName, "Լ�����ʹ���ӦΪ���Լ��", "���أ�����Ӱ��ϵͳ����")
                            End If
                        Else
                            arySql = Split(strSQL, " FOREIGN ")
                            strTemp = Replace(Replace(Replace(Left(arySql(1), InStr(1, arySql(1), ")") - 1), "KEY", ""), "(", ""), " ", "")
                            If strColumns <> "," & strTemp Then
                                If blnAddToRsTable Then
                                    '״̬:'1-���ڶ���,2-�����ڶ���,3-ʧЧ,4-ȱ����,5-����,6-���ڽ�ֹ״̬,7-Լ����һ��
                                    '����:����,Ψһ,���,Լ��,����,��ͼ,...
                                    Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "Լ��", 4, False, strSQL1, "Լ���д���", "���أ�����Ӱ��ϵͳ���У�ӦΪ(" & strTemp & ")����Ϊ(" & Mid(strColumns, 2) & ")")
                                Else
                                    Call InputErrRpt("Լ��", strObjName, "Լ���д���ӦΪ(" & strTemp & ")����Ϊ(" & Mid(strColumns, 2) & ")", "���أ�����Ӱ������һ����")
                                End If
                            End If
                        End If
                    ElseIf InStr(1, strSQL, " CHECK") > 0 Then
                        If !constraint_type <> "C" Then
                            If blnAddToRsTable Then
                                '״̬:'1-���ڶ���,2-�����ڶ���,3-ʧЧ,4-ȱ����,5-����,6-���ڽ�ֹ״̬,7-Լ����һ��
                                '����:����,Ψһ,���,Լ��,����,��ͼ,...
                                Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "Լ��", 7, False, strSQL1, "Լ�����ʹ���", "���أ�����Ӱ��ϵͳ���У�ӦΪ���Լ��")
                            Else
                                Call InputErrRpt("Լ��", strObjName, "Լ�����ʹ���ӦΪ���Լ��", "���أ�����Ӱ��ϵͳ����")
                            End If
                        Else
                            '25047:���˺�����
                            arySql = Split(strSQL, " CHECK")
                            strTemp = Replace(UCase(Replace(Replace(Replace(arySql(1), " ", ""), vbTab, ""), vbCrLf, "")), ";", "")
                            strTemp1 = "(" & Replace(UCase(Replace(Replace(Replace(Nvl(!Search_Condition), " ", ""), vbTab, ""), vbCrLf, "")), ";", "") & ")"
                            If strTemp <> strTemp1 Then
                                '���Լ������Ƿ�һ��
                                strTemp = Trim(Replace(Replace(Replace(arySql(1), vbCrLf, " "), vbTab, " "), ";", ""))
                                strTemp = Mid(strTemp, 2, Len(strTemp) - 2)
                                
                                If blnAddToRsTable Then
                                    '״̬:'1-���ڶ���,2-�����ڶ���,3-ʧЧ,4-ȱ����,5-����,6-���ڽ�ֹ״̬,7-Լ����һ��
                                    '����:����,Ψһ,���,Լ��,����,��ͼ,...
                                    Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "Լ��", 7, False, strSQL1, "Լ�����ݲ�һ��", "���أ�����Ӱ��ϵͳ����,�������ݲ�һ��!,ӦΪ(" & strTemp & "),����Ϊ(" & Nvl(!Search_Condition) & ")")
                                Else
                                    Call InputErrRpt("Լ��", strObjName, "Լ�����ݲ�һ��,ӦΪ(" & strTemp & "),����Ϊ(" & Nvl(!Search_Condition) & ")", "���أ�����Ӱ��ϵͳ����,�������ݲ�һ��!")
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            pgbState.value = mclsRunScript.Line / mclsRunScript.LinesCount * 100
            Call mclsRunScript.ReadNextSQL
            DoEvents
        Loop
    End With
    pgbState.value = 0
End Sub
Private Sub CheckIndex(FileName As String, Optional ByVal blnAddToRsTable As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Լ�����ж��Ƿ���Ч����
    '���:FileName-�ű��ļ���(��������·��)
    '     blnAddToRsTable-�Ƿ���صĴ�����Ϣд�ڼ�¼����
    '����:
    '����:
    '����:���˺�
    '����:2009-08-19 15:03:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strԭ�ֶξ��� As String, str���ֶξ��� As String, strSQL1 As String
    Dim arySql() As String, strObjName As String, strColumns As String
    Dim strTablenName As String
    Dim rsObjects As New ADODB.Recordset, strSQL As String
    Dim rsColumns As New ADODB.Recordset, strTemp As String
    With rsObjects
        .Filter = 0
        If gblnDBA Then
            strSQL = "select INDEX_NAME,STATUS from DBA_INDEXES where OWNER='" & cmbSystem.Tag & "' and TABLE_OWNER='" & cmbSystem.Tag & "'"
        Else
            strSQL = "select INDEX_NAME,STATUS from USER_INDEXES where TABLE_OWNER='" & cmbSystem.Tag & "'"
        End If
        If .State = adStateOpen Then .Close
        .Open strSQL, gcnOracle, adOpenKeyset
    
        err = 0
        On Error Resume Next
        If Not mclsRunScript.OpenFile(FileName) Then Exit Sub
        On Error GoTo 0
        Do While Not mclsRunScript.EOF
            strSQL = mclsRunScript.SQLInfo.PartSQL
            strSQL1 = strSQL
            arySql = Split(strSQL, " INDEX ")
            If UBound(arySql) > 0 Then
                strSQL = Trim(arySql(1)) '�Ѿ�ȥ��Oracle�ؼ���
                strTablenName = Trim(Split(Split(strSQL, "ON")(1), "(")(0))
                strObjName = Trim(Left(strSQL, InStr(strSQL, " ")))
                .Filter = "INDEX_NAME='" & strObjName & "'"
                If .EOF Then
                    If blnAddToRsTable Then
                        '״̬:1-���ڶ���,2-�����ڶ���,3-ʧЧ,4-ȱ����,5-����,6-���ڽ�ֹ״̬
                        '����:����,Ψһ,���,Լ��,����,��ͼ,...
                        Call zlInsertRecData(mrsErrTable, strTablenName, strObjName, "����", 2, False, strSQL1, "������", "���أ�����Ӱ��ϵͳ�����ٶ�")
                    Else
                        Call InputErrRpt("����", strObjName, "������", "���أ�����Ӱ��ϵͳ�����ٶ�")
                    End If
                ElseIf .Fields("STATUS").value <> "VALID" Then
                    Call InputErrRpt("����", strObjName, "��ǰ������Ч״̬")
                Else
                    With rsColumns
                        If gblnDBA Then
                            strTemp = "select TABLE_NAME,COLUMN_NAME" & _
                                    " from DBA_IND_COLUMNS" & _
                                    " where INDEX_OWNER='" & cmbSystem.Tag & "' and INDEX_NAME='" & strObjName & "'" & _
                                    " order by COLUMN_POSITION"
                        Else
                            strTemp = "select TABLE_NAME,COLUMN_NAME" & _
                                    " from USER_IND_COLUMNS" & _
                                    " where INDEX_NAME='" & strObjName & "'" & _
                                    " order by COLUMN_POSITION"
                        End If
                        If .State = adStateOpen Then .Close
                        .Open strTemp, gcnOracle, adOpenKeyset
                        Do While Not .EOF
                            If .AbsolutePosition = 1 Then
                                strColumns = !Table_Name & "(" & !Column_Name
                            Else
                                strColumns = strColumns & "," & !Column_Name
                            End If
                            .MoveNext
                        Loop
                            strColumns = strColumns & ")"
                    End With
                    arySql = Split(strSQL, " ON ")
                    strTemp = Replace(Left(arySql(1), InStr(1, arySql(1), ")")), " ", "")
                    If strColumns <> strTemp Then
                        If blnAddToRsTable Then
                            '״̬:1-���ڶ���,2-�����ڶ���,3-ʧЧ,4-ȱ����,5-����,6-���ڽ�ֹ״̬
                            '����:����,Ψһ,���,Լ��,����,��ͼ,...
                            Call zlInsertRecData(mrsErrTable, strTablenName, strObjName, "����", 4, False, strSQL1, "�����д���", "���أ�����Ӱ��ϵͳ�����ٶ�,ӦΪ��" & strTemp & "������Ϊ��" & strColumns & "��")
                        Else
                            Call InputErrRpt("����", strObjName, "�����д���ӦΪ��" & strTemp & "������Ϊ��" & strColumns & "��", "���أ�����Ӱ��ϵͳ�����ٶ�")
                        End If
                    End If
                End If
            End If
            Call mclsRunScript.ReadNextSQL
        Loop
    End With
End Sub

Private Function RunSQLScript(ByVal cnThisDB As ADODB.Connection, ByVal strFile As String, Optional blnResumeNext As Boolean = True) As Boolean
'----------------------------------------------
'����:ִ��SQL�ļ�
'����:
'    cnThisDB=��ǰϵͳ����
'    strFile=�ű��ļ�
'    blnResumeNext=�Ƿ�������
'    ���أ�true-ִ�гɹ���false-ִ��ʧ��
' ----------------------------------------------
    Dim lngLines As Long
    err = 0
    On Error Resume Next
    If Not mclsRunScript.OpenFile(strFile) Then Exit Function
    pgbState.value = 0
    Do While Not mclsRunScript.EOF
        cnThisDB.Execute mclsRunScript.SQLInfo.SQL
        If err <> 0 Then
            If blnResumeNext Then
                err.Clear
            Else
                MsgBox "�����ļ�" & strFile & "�д������������ִ���жϣ�" & vbCr & mclsRunScript.SQLInfo.SQL, vbExclamation, gstrSysName
                Exit Function
            End If
        End If
        pgbState.value = mclsRunScript.Line / mclsRunScript.LinesCount * 100
        Call mclsRunScript.ReadNextSQL
        DoEvents
    Loop
    pgbState.value = 0
    RunSQLScript = True
End Function



Private Sub Form_Unload(Cancel As Integer)
    If picStatus.Visible Then Cancel = 1
    
    Set mclsObjectCheck = Nothing
    Set mclsRunScript = Nothing
    Set mrsErrTable = Nothing
    
End Sub

Private Sub mclsObjectCheck_AfterObjectCheck()
    picStatus.Visible = False
    cmdFunction(0).Enabled = True
End Sub

Private Sub mclsObjectCheck_AfterProgress()
    lblStatus.Caption = ""
    pgbState.value = 0
End Sub

Private Sub mclsObjectCheck_BeforeObjectCheck()
    picStatus.Visible = True
    cmdFunction(0).Enabled = False
End Sub

Private Sub mclsObjectCheck_BeforeProgress(ByVal Title As String, ByVal Max As Long)
    lblStatus.Caption = Title
    pgbState.Max = Max
End Sub

Private Sub mclsObjectCheck_Exception()
    Dim lngCount As Long
    
    For lngCount = 0 To cmdFunction.Ubound
        cmdFunction(lngCount).Enabled = False
    Next
    lblFileName.Caption = ""
    
    On Error Resume Next
    cmdGetIni.SetFocus
End Sub

Private Sub mclsObjectCheck_Progressing(ByVal Progress As Long)
    pgbState.value = Progress
    DoEvents
End Sub

Private Sub picStatus_Resize()
    pgbState.Width = picStatus.ScaleWidth - pgbState.Left * 2
End Sub

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�

End Sub

Private Sub zlGetConstraintInfor(FileName As String, ByRef cllOutPara As Collection)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������صĽű�,��ȡ��ص�Լ����Ϣ
    '���:FileName-�ű��ļ�
    '����:cllOutPara-����Լ����Ϣ(����,Լ����,����(UQ,PK,CK,FK),���ڷ�(Y/N),��ʷ���ݿռ�����(Y/N)))
    '����:
    '����:���˺�
    '����:2009-07-21 15:55:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arySql() As String, strObjName As String, strColumns As String, strTableName As String
    Dim blnFK As Boolean
    Dim rsObjects As New ADODB.Recordset, strSQL As String
    
    With rsObjects
        .Filter = 0
        If gblnDBA Then
            strSQL = "select CONSTRAINT_TYPE,CONSTRAINT_NAME,STATUS,VALIDATED,BAD from DBA_CONSTRAINTS where OWNER='" & cmbSystem.Tag & "'"
        Else
            strSQL = "select CONSTRAINT_TYPE,CONSTRAINT_NAME,STATUS,VALIDATED,BAD from USER_CONSTRAINTS"
        End If
        If .State = adStateOpen Then .Close
        .Open strSQL, gcnOracle, adOpenKeyset
    
        err = 0
        On Error Resume Next
        If Not mclsRunScript.OpenFile(FileName) Then Exit Sub
        On Error GoTo 0
        Do While Not mclsRunScript.EOF
            strSQL = mclsRunScript.FormatSQL
            arySql = Split(strSQL, " CONSTRAINT ")
            
            If UBound(arySql) > 0 Then
                '���˺����:
                strTableName = Trim(Split(Trim(arySql(0)), "TABLE")(1))
                strTableName = Split(strTableName, " ")(0)
                strSQL = Trim(arySql(1)) '�Ѿ�ȥ��Oracle�ؼ���
                strObjName = Trim(Left(strSQL, InStr(strSQL, " ")))
                blnFK = InStr(1, UCase(Replace(strSQL, " ", "")), UCase("ForeignKey")) > 0
                
                
                
                
                .Filter = "CONSTRAINT_NAME='" & strObjName & "'" '
                cllOutPara.Add Array(strTableName, strObjName, IIf(blnFK, "Y", "N"), IIf(.EOF, "N", "Y"))
                
            End If
            pgbState.value = mclsRunScript.Line / mclsRunScript.LinesCount * 100
            Call mclsRunScript.ReadNextSQL
            DoEvents
        Loop
    End With
    pgbState.value = 0
End Sub
Private Sub zlGetIndexInfor(FileName As String, ByRef cllOutPara As Collection)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:
    '����:cllOutPara-����������Ϣ(����,������,���ڷ�(Y/N),��ʷ���ݿռ�����(Y/N)))
    '����:
    '����:���˺�
    '����:2009-07-21 17:11:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arySql() As String, strObjName As String, strColumns As String
    Dim strTablenName As String
    Dim rsObjects As New ADODB.Recordset, strSQL As String
    
    With rsObjects
        .Filter = 0
        If gblnDBA Then
            strSQL = "select INDEX_NAME,STATUS from DBA_INDEXES where OWNER='" & cmbSystem.Tag & "' and TABLE_OWNER='" & cmbSystem.Tag & "'"
        Else
            strSQL = "select INDEX_NAME,STATUS from USER_INDEXES where TABLE_OWNER='" & cmbSystem.Tag & "'"
        End If
        If .State = adStateOpen Then .Close
        .Open strSQL, gcnOracle, adOpenKeyset
    
        err = 0
        On Error Resume Next
        If Not mclsRunScript.OpenFile(FileName) Then Exit Sub
        
        Do While Not mclsRunScript.EOF
            strSQL = mclsRunScript.FormatSQL
            arySql = Split(strSQL, " INDEX ")
            If UBound(arySql) > 0 Then
                strSQL = Trim(arySql(1)) '�Ѿ�ȥ��Oracle�ؼ���
                strTablenName = Trim(Split(Split(strSQL, "ON")(1), "(")(0))
                strObjName = Trim(Left(strSQL, InStr(strSQL, " ")))
                .Filter = "INDEX_NAME='" & strObjName & "'"
                cllOutPara.Add Array(strTablenName, strObjName, IIf(.EOF, "N", "Y"))
            End If
            Call mclsRunScript.ReadNextSQL
        Loop
        
    End With
End Sub

Private Function CheckHistorySpaceEx(ByVal lngSys As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵱ǰϵͳ���Ƿ�����ʷ��ռ�ı���Ϣ
    '���:
    '����:CheckHistorySpace-���ڶ�Ӧ����ʷ��
    '����:
    '����:��˶
    '����:2013-04-07 10:25:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    
    On Error Resume Next '���ܵ�ǰ�û�û�б�Ȩ��
    gstrSQL = "Select ����,������ From Zltools.Zlbakspaces Where ϵͳ = " & lngSys & "  And ��ǰ = 1 And ֻ�� = 0"
    Call OpenRecordset(rsTmp, gstrSQL, "��ȡ��ʷ��ռ�������")
    CheckHistorySpaceEx = Not rsTmp.EOF
    On Error GoTo 0
End Function

Private Function GetToolsIniVersion(ByVal strFilePath As String) As String
'����:���ݷ����������߽ű��İ�װ�����ļ�·����ȡ�����߰�װ�ű���Ӧ�İ汾��
    Dim strMaxVer As String, strVer As String
    Dim objFile As Scripting.File
    Dim intType As Integer
    
    On Error Resume Next
    For Each objFile In gobjFile.GetFile(strFilePath).ParentFolder.Files
        If AnalysisFileName(objFile.name, 0, strVer) Then
            '�����ű��İ汾Ϊ��汾�Ž��бȽ�
            If strVer Like "*.*.0" Then strMaxVer = IIf(VerFull(strVer) > VerFull(strMaxVer, False), strVer, strMaxVer)
        End If
    Next
    GetToolsIniVersion = strMaxVer
End Function


Private Sub SetFunsState(ByVal strFilePath As String, Optional blnMsg As Boolean)
'���ܣ�����������������ܿ���������
    Dim blnTools As Boolean
    Dim blnTmp As Boolean
    Dim strToolsByServer As String
    Dim varData As Variant
    Dim strVer As String
    Dim i As Long
    
    blnTools = cmbSystem.Tag = "ZLTOOLS"
    '�ļ������Լ��
    blnTmp = gobjFile.FileExists(strFilePath)
    
    If Not blnTools And blnTmp Then
        blnTmp = CheckIniFile(strFilePath, blnMsg)
    End If
    
    '�ļ����ɹ�
    lblFileName.Caption = IIf(blnTmp, strFilePath, "")
    For i = 0 To cmdFunction.UBound - 1
        If i <= 1 Then
            cmdFunction(i).Enabled = (UCase(gstrUserName) = UCase(cmbSystem.Tag) And Not blnTools Or blnTools) And blnTmp
        ElseIf blnTools Then
            cmdFunction(i).Enabled = False
        ElseIf i = CMDFUN.E��ʷ�ṹ Then
            cmdFunction(i).Enabled = UCase(gstrUserName) = UCase(cmbSystem.Tag) And CheckHistorySpaceEx(Val(cmbSystem.ItemData(cmbSystem.ListIndex)))
        ElseIf i = CMDFUN.Eͬ��� Then
            cmdFunction(i).Enabled = UCase(gstrUserName) = UCase(cmbSystem.Tag)
        Else
            cmdFunction(i).Enabled = True
        End If
    Next
    '���ö���Ȩ�������Ƿ����
    cmdFunction(E����Ȩ��).Visible = blnTools
    cmdFunction(E��ʷ�ṹ).Visible = Not blnTools

    If Me.Visible And Not blnTmp Then cmdGetIni.SetFocus: Exit Sub
    '�汾���
    '��ȡ�����ߵ�ǰ���ű��İ汾
    If blnTools Then strToolsByServer = GetToolsIniVersion(lblFileName.Caption)
    varData = Split(cmbSystem.Text & "v", "v")
    varData = Split(varData(1), "��")
    varData = Split(varData(0) & "....", ".")
    strVer = Val(varData(0)) & "." & Val(varData(1)) & "." & Val(varData(2))
    If Val(varData(2)) > 0 Then
        cmdFunction(0).Enabled = False
        cmdFunction(1).Enabled = False
    Else
        If blnTools And strVer <> strToolsByServer Then
            cmdFunction(0).Enabled = False
            cmdFunction(1).Enabled = False
        End If
    End If
End Sub

