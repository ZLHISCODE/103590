VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUserLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�����ߵ�¼"
   ClientHeight    =   2595
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "frmUserLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4470
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdSet 
      Caption         =   "���÷�����"
      Height          =   350
      Left            =   150
      TabIndex        =   10
      ToolTipText     =   "����Oracle�����ַ������ó���"
      Top             =   2115
      Width           =   1100
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "��"
      Height          =   300
      Left            =   3720
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "ѡ����ڵķ������б�"
      Top             =   1455
      Width           =   300
   End
   Begin VB.TextBox txt���ݿ� 
      Height          =   300
      IMEMode         =   2  'OFF
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1455
      Width           =   1785
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1050
      Width           =   2115
   End
   Begin VB.TextBox txt�û� 
      Height          =   300
      IMEMode         =   2  'OFF
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   2
      Top             =   645
      Width           =   2115
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3135
      TabIndex        =   9
      Top             =   2115
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1875
      TabIndex        =   8
      Top             =   2115
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -150
      TabIndex        =   11
      Top             =   1860
      Width           =   4965
   End
   Begin MSComDlg.CommonDialog cdgFile 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblNote 
      Caption         =   "    ֻ�о������ݿ�DBA��ɫ�����ϵͳ�������߲���ʹ�ñ����ߡ�"
      Height          =   375
      Left            =   990
      TabIndex        =   0
      Top             =   105
      Width           =   3195
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   1485
      TabIndex        =   3
      Top             =   1110
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "�û���"
      Height          =   180
      Left            =   1305
      TabIndex        =   1
      Top             =   705
      Width           =   540
   End
   Begin VB.Label lblDataBase 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   180
      Left            =   1305
      TabIndex        =   5
      Top             =   1515
      Width           =   540
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   180
      Picture         =   "frmUserLogin.frx":1CFA
      Top             =   105
      Width           =   720
   End
End
Attribute VB_Name = "frmUserLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intTimes As Integer
Dim strNote As String
Dim strUserName As String
Dim strServerName As String
Dim strPassword As String
Private mstrCommand As String
Private mblnת�� As Boolean
Private mblnAccess As Boolean

Dim mcolServer As New Collection

Public Sub ShowMe(ByVal strCommand As String)
    mstrCommand = strCommand
    mblnת�� = True
    Me.Show vbModal
End Sub

Private Sub cmdOK_Click()
    
    intTimes = intTimes + 1
    
    '------�����û��Ƿ�oracle�Ϸ��û�----------------
    strUserName = Trim(txt�û�.Text)
    strPassword = Trim(txt����.Text)
    strServerName = Trim(txt���ݿ�.Text)
    
    '��Ч�ַ���Ч��
    If Len(Trim(txt�û�)) = 0 Then
        strNote = "�������û�����"
        txt�û�.SetFocus
        GoTo InputError
    End If
    
    If Len(strUserName) <> 1 Then
        If Mid(strUserName, 1, 1) = "/" Or Mid(strUserName, 1, 1) = "@" Or Mid(strUserName, Len(strUserName) - 1, 1) = "/" Or Mid(strUserName, Len(strUserName) - 1, 1) = "@" Then
            txt�û�.SetFocus
            strNote = "�û�������"
            Exit Sub
        End If
    End If
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            txt����.SetFocus
            strNote = "�������"
            GoTo InputError
        End If
    End If
    If Trim(strServerName) <> "" Then
        If Mid(strServerName, Len(strServerName) - 1, 1) = "/" Or Mid(strServerName, Len(strServerName) - 1, 1) = "@" Or Mid(strServerName, 1, 1) = "/" Or Mid(strServerName, 1, 1) = "@" Then
            strNote = "�������Ӵ�����"
            txt���ݿ�.SetFocus
            GoTo InputError
        End If
    End If
    
    '�����ַ���
    Dim intPos As Integer
    intPos = InStr(1, strUserName, "@")
    If intPos > 0 Then
        strServerName = Mid(strUserName, intPos + 1)
        strUserName = Mid(strUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(1, strUserName, "/")
    If intPos > 0 Then
        strPassword = Mid(strUserName, intPos + 1)
        strUserName = Mid(strUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(1, strPassword, "@")
    If intPos > 0 Then
        strServerName = Mid(strPassword, intPos + 1)
        strPassword = Mid(strPassword, 1, intPos - 1)
    End If
    
    If Len(Trim(strPassword)) = 0 Then
        strNote = "δ�������룬����ע�ᡣ"
        txt����.SetFocus
        GoTo InputError
    End If
        
    strUserName = UCase(strUserName)
    If Not OraDataOpen(strServerName, strUserName, strPassword) Then
        If Me.Visible = False Then Me.Visible = True
        If glngSysNo <> -1 Then Me.Visible = False
        mblnAccess = False
        txt����.Text = ""
        Exit Sub
    End If
    
    '�޸�ע���
    SaveSetting "ZLSOFT", "ע����Ϣ\��½��Ϣ", "MANAGER", strUserName
    SaveSetting "ZLSOFT", "ע����Ϣ\��½��Ϣ", "SERVER", strServerName
    mblnAccess = True
    Unload Me
    Exit Sub

InputError:
    If intTimes > 3 Then
        MsgBox "��������ע��ʧ�ܣ�ϵͳ���Զ��˳���", vbExclamation, gstrSysName
        cmdCancel_Click
    Else
        If strNote <> "" Then
            MsgBox strNote, vbExclamation, gstrSysName
        End If
        Exit Sub
    End If

End Sub

Private Function GetRegFunFile() As String
    With cdgFile
        .DialogTitle = "ѡ��ע�ắ���ļ������´�����غ���"
        .Filter = IIf(GetOracleVersion(True, True) > 11, "ע�ắ���ļ�(ZLREGIST12C.PLB)|ZLREGIST12C.PLB", "ע�ắ���ļ�(ZLREGIST.PLB)|ZLREGIST.PLB")
        .flags = &H200000 Or &H4 Or &H2 Or &H800 Or &H4000
        .CancelError = False
        .ShowOpen
        
        GetRegFunFile = .Filename
    End With
End Function

Private Function RegCheckAndGetUnit() As String
'���ܣ���֤ϵͳע����Ȩ����ȷ�ԣ������ص�λ����
    Dim strUnit As String, strRegFunFile As String
    Dim strRegErr As String, strPassword As String, strError As String, strSQL As String
    Dim cnTools As ADODB.Connection
    Dim rstmp As ADODB.Recordset, blnLoginAgain As Boolean
    
    strRegErr = gobjRegister.zlRegCheck(False)
    If strRegErr <> "" Then
        Me.Visible = False
        If strRegErr Like "*�ָ���ȷ��ע�ắ����*" Then
            If GetOracleVersion(True, True) > 11 Then
                strRegErr = Replace(strRegErr, "ZLREGIST.PLB", "ZLREGIST12C.PLB")
            End If
            If MsgBox(strRegErr & vbCrLf & "����Ҫ���´���ע�ắ����", vbYesNo + vbDefaultButton2, "��ʾ") = vbNo Then
                End
            Else
                '����ȱʡ����ZLTOOLSִ������
                strPassword = "ZLTOOLS"
                blnLoginAgain = True
openconn:       Set cnTools = gobjRegister.GetConnection(strServerName, "ZLTOOLS", strPassword, False, OraOLEDB, strError, False)
                If strError <> "" Then
                    If blnLoginAgain Then
                        strError = ""
                        strPassword = "ZLSOFT"
                        Set cnTools = gobjRegister.GetConnection(strServerName, "ZLTOOLS", strPassword, False, OraOLEDB, strError, False)
                        blnLoginAgain = False
                    End If
                End If
                
                If strError <> "" Then
                    strPassword = InputBox("ע�ắ����֤ʧ�ܣ��������´���ע�ắ��(" & IIf(GetOracleVersion(True, True) > 11, "ZLREGIST12C.PLB", "ZLREGIST.PLB") & ")��" & vbCrLf & "��Ҫ��ZLTOOLS�û���¼ִ�У���������û�������", "��ʾ")
                    If strPassword = "" Then
                        End
                    Else
                        strError = ""
                        GoTo openconn
                    End If
                End If
                
                On Error GoTo errH
                '1.���ע�ắ������ı�ṹ�Ƿ���Ҫ����
                strSQL = "Select Table_Name" & vbNewLine & _
                        "From User_Tab_Columns" & vbNewLine & _
                        "Where Table_Name In ('ZLREGFILE', 'ZLREGAUDIT') And Column_Name = '��Ŀ' And Data_Length <> 20"
                Set rstmp = gclsBase.OpenSQLRecord(cnTools, strSQL, "������ݽṹ")
                If rstmp.RecordCount > 0 Then
                    
                    '2.�����Ҫ�������򵯳�ѡ����ϵ�����ZLHIS�ͻ���
                    If MsgBox("��⵽ע��������ݽṹ��Ҫ������Ҫ��Ͽ�����ZLHIS�ͻ��˲�����Ӧ��ϵͳ�˻���" & vbCrLf & "��ȷ������������������ZLHIS�ͻ��˲�����Ӧ��ϵͳ�˻���" & vbCrLf & _
                            "ע���뵽ϵͳ��Ǩ�����������û��Ľ������ͻ��˵����á�", vbQuestion + vbOKCancel + vbDefaultButton2, "��ʾ") = vbCancel Then
                        End
                    Else
                        '�Ͽ�����ZLHIS�ͻ������Ӳ����޸�������ʱ��Ľṹ
                        Call LockAppUser
                        Call KillSessions
                    End If
                    
                    rstmp.Filter = "Table_Name='ZLREGFILE'"
                    If rstmp.RecordCount > 0 Then
                        strSQL = "Alter Table zlRegFile Modify ��Ŀ Varchar2(20)"
                        cnTools.Execute strSQL
                    End If
                    
                    rstmp.Filter = "Table_Name='ZLREGAUDIT'"
                    If rstmp.RecordCount > 0 Then
                        strSQL = "Alter Table ZLREGAUDIT Modify ��Ŀ Varchar2(20)"
                        cnTools.Execute strSQL
                    End If
                    
                    strSQL = "Drop Type t_Reg_Rowset Force"
                    cnTools.Execute strSQL
                    strSQL = "Drop Type t_Reg_Record Force"
                    cnTools.Execute strSQL
                    strSQL = "Create Or Replace Type t_Reg_Record  As Object(Item Varchar2(20), Prog number(18), Text Varchar2(1000))"
                    cnTools.Execute strSQL
                    strSQL = "Create Or Replace Type t_Reg_Rowset As Table Of t_Reg_Record"
                    cnTools.Execute strSQL
                    
                    
                    On Error Resume Next
                    strSQL = "Grant Execute on t_Reg_Record to Public"
                    cnTools.Execute strSQL
                    If err.Number <> 0 Then
                        MsgBox "ִ�а���Ȩʱʧ�ܣ�����������" & vbCrLf & err.Description & vbCrLf _
                            & "��Ҫ���ȷ�����Ƚ�����ش�����ֹ��������½ű���" & vbCrLf & strSQL, vbExclamation, "��ʾ"
                        err.Clear
                    End If
                    
                    '����ҽԺ����ZLHIS�İ�T_DB_ROLEUSER��BH�ܺ���صģ������˸ö��󣬵�����Ȩʧ��
                    'ORA-04045: �����±���/������֤ ZLHIS.T_DB_ROLEUSER ʱ����
                    'ORA -1031: Ȩ�޲���
                    strSQL = "Grant Execute on t_Reg_Rowset to Public"
                    cnTools.Execute strSQL
                    If err.Number <> 0 Then
                        MsgBox "ִ�а���Ȩʱʧ�ܣ�����������" & vbCrLf & err.Description & vbCrLf _
                            & "��Ҫ���ȷ�����Ƚ�����ش�����ֹ��������½ű���" & vbCrLf & strSQL, vbExclamation, "��ʾ"
                        err.Clear
                    End If
                    On Error GoTo errH
                End If
                
                
                '3.ִ��ע���ļ�
                If Not gblnInIDE Then '���Ӷ໷��֧��
                    strRegFunFile = App.Path & "\TOOLS\" & IIf(GetOracleVersion(True, True) > 11, "ZLREGIST12C.PLB", "zlRegist.plb")
                Else
                    strRegFunFile = "C:\APPSOFT\TOOLS\" & IIf(GetOracleVersion(True, True) > 11, "ZLREGIST12C.PLB", "zlRegist.plb")
                End If
                If gobjFile.FileExists(strRegFunFile) = False Then
                    strRegFunFile = GetRegFunFile
                    If strRegFunFile = "" Then
                        End
                    End If
                End If
                
                If RunRegistFile(Me, cnTools, strPassword, strServerName, strRegFunFile) Then
                    MsgBox "ע�ắ��������ɣ������½���ע�ᣡ" & vbCrLf & "ע�ắ����Դ��" & vbCrLf & strRegFunFile, vbInformation
                Else
                    End
                End If
            End If
        Else
            MsgBox "ע����֤ʧ�ܣ�������ע�ᣡ" & vbCrLf & strRegErr, vbInformation, "����"
        End If
        
        If Not frmReg.ReReg Then
            End
        End If
    End If
    strUnit = gobjRegister.zlRegInfo("��λ����", False, 0)
    If strUnit = "" Then
        MsgBox "δ�ܶ�ȡ����λ���ƣ�����ע���뼰ע�ắ������������ע��!", vbExclamation, "����"
        If Not frmReg.ReReg Then
            End
        End If
    End If
    RegCheckAndGetUnit = strUnit
    Exit Function
    
errH:
    MsgBox err.Description & vbCrLf & "���һ��ִ�е�SQL��" & strSQL, vbExclamation, "��ʾ"
    End
End Function

Private Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strPassword As String) As Boolean
'���ܣ� ��ָ�������ݿ����ӣ��������ͨ�û�����ʹ�ù���Ա�ʺ����´�����
'������
'   strServerName�������ַ���
'   strUserName���û���
'   strUserPwd������
'���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false

    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim strDest() As Byte
    Dim StrJiemi() As Byte
    Dim StrJiami() As Byte
    Dim blnGrantMgr As Boolean '��Ȩ�Ĺ���������
    Dim strPwdTxt As String, strRegErr As String, strUnit As String
    Dim blnLogin As Boolean, blnTransPassword As Boolean
    Dim strHaveProg As String
    Dim strError As String
    
    '֧��strServerName = "192.168.2.13:1521/dyyy"���ָ�ʽ
    
    gstrLoginPwd = strPassword
    
    If UCase(strUserName) = "SYS" Or UCase(strUserName) = "SYSTEM" Then
        blnTransPassword = False
    Else
        blnTransPassword = mblnת��
    End If
    Set gcnOracle = gobjRegister.GetConnection(strServerName, strUserName, strPassword, blnTransPassword, OraOLEDB, strError)
    If gcnOracle.State = adStateClosed Then
        If InStr(strError, "ORA-00604") > 0 Then
            If InStr(strError, "ORA-20002") > 0 Then
                strError = "��ǰ�û�����ʹ�ø�Ӧ�õ�¼���ݿ⣬����ϵ����Ա��"
            Else
                strError = "��ǰ�û�����ֹ��¼���ݿ⣬����ϵ����Ա��"
            End If
        End If
        MsgBox strError, vbInformation, gstrSysName
        OraDataOpen = False
        Exit Function
    End If
    Call SetSQLTrace(strServerName, strUserName, gcnOracle)
    
    
    Call gobjRegister.zlRegInit(gcnOracle)
    
    On Error Resume Next
    strSQL = "SELECT 1 FROM ZLTOOLS.ZLSYSTEMS WHERE ������=USER"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "ϵͳ�������ж�")
    If err.Number <> 0 Then
        gblnCreate = False
        gblnOwner = False
        err.Clear
    Else
        gblnCreate = True
        gblnOwner = Not rsTemp.EOF
    End If
    
    strSQL = "SELECT 1 FROM SESSION_ROLES WHERE ROLE='DBA'"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "DBA�ж�")
    gblnDBA = Not rsTemp.EOF
        gblnRac = CheckRAC(gintInstID)
    If err.Number <> 0 Then err.Clear
    
    If Not (gblnDBA) And Not (gblnCreate) Then
        OraDataOpen = False
        MsgBox "�״����У�������DBAע�ᣬ�Ա㴴�������ߣ�", vbExclamation, gstrSysName
        Exit Function
    End If
    
    
    '��ͨ�û���¼������ʱ����ϵͳ�����߽���ʵ��������
    If gblnCreate Then
        strUnit = RegCheckAndGetUnit
        If strUnit = "" Then End

        gstrHaveProg = "": blnGrantMgr = False: blnLogin = False
        gstrLoginUserName = strUserName
        gstrLoginUserPwd = gobjRegister.GetPassword
        
        If Not gblnDBA And Not gblnOwner Then
            '����Ƿ��й����ߵ�Ȩ��
            strSQL = "select ���� from zltools.Zlmgrgrant Where �û���='" & gstrLoginUserName & "'"
            Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��������Ȩ�û�")
            If rsTemp.RecordCount > 0 Then
                gstrHaveProg = rsTemp!���� & ""
                If gstrHaveProg <> "" Then
                    ReDim Preserve strDest(0): ReDim Preserve StrJiemi(0)
                    Call Func16CodeToByte(gstrHaveProg, strDest)
                    Call DES_Decode(strDest, StrJiemi, strUnit)
                    gstrHaveProg = Replace(StrConv(StrJiemi, vbUnicode), Chr(0), "")
                    
                    '��Ȩ���ַ������г�ʼ������
                    gstrHaveProg = GetProgFuncs(gstrHaveProg, True)
                    
                    blnGrantMgr = True
                    
                    '�ж��Ƿ�Ϊ��ϵͳ��¼
                    If glngSysNo <> -1 Then
                        If InStr(gstrHaveProg, "0401") Then
                            strHaveProg = "0401"
                        End If
                        If InStr(gstrHaveProg, "0402") Then
                            strHaveProg = IIf(strHaveProg = "", "", strHaveProg & ",") & "0402"
                        End If
                        gstrHaveProg = strHaveProg
                        If gstrHaveProg = "" Then
                            blnGrantMgr = False
                        End If
                    End If
                    
                End If
            End If
            If Not blnGrantMgr Then
                OraDataOpen = False
                MsgBox "��û�й����ߵ�ʹ��Ȩ�ޣ�����ϵ����Ա��", vbExclamation, gstrSysName
                Exit Function
            ElseIf gstrHaveProg = "" Then
                OraDataOpen = False
                MsgBox "���Ĺ����ߵ�ʹ��Ȩ�޶�ʧ������ϵ����Ա������Ȩ��", vbExclamation, gstrSysName
                Exit Function
            End If
            
            'ʹ��ϵͳ����Ա��¼
            If err.Number <> 0 Then err.Clear
            strUserName = "": strPassword = ""
            strSQL = "Select Max(Decode(��Ŀ,'����Ա',����,'')) AS ����Ա ,Max(Decode(��Ŀ,'��֤��',����,'')) AS ��֤�� From zltools.zlRegInfo where ��Ŀ='����Ա' Or ��Ŀ='��֤��'"
            Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��Ȩ��¼��Ϣ")
            If rsTemp!����Ա & "" <> "" And rsTemp!��֤�� & "" <> "" Then
                strUserName = rsTemp!����Ա & ""
                ReDim Preserve strDest(0): ReDim Preserve StrJiemi(0)
                Call Func16CodeToByte(rsTemp!��֤�� & "", strDest)
                Call DES_Decode(strDest, StrJiemi, strUnit)
                strPassword = Replace(StrConv(StrJiemi, vbUnicode), Chr(0), "")
                
                '���´����ݿ�����(�洢�������ݿ����룬���Բ���Ҫת��)
                Set gcnOracle = gobjRegister.GetConnection(strServerName, strUserName, strPassword, False, OraOLEDB)
                blnLogin = gcnOracle.State = adStateOpen
                
                If blnLogin Then
                    Call SetSQLTrace(strServerName, strUserName, gcnOracle)
                    '������֤�Ự
                    Call gobjRegister.zlRegInit(gcnOracle)
                    strRegErr = gobjRegister.zlRegCheck(False)
                    If strRegErr <> "" Then
                        MsgBox strRegErr, vbQuestion, "����"
                        If Not frmReg.ReReg Then
                            End
                        End If
                    End If
                End If
            End If
            
            '����ʹ�ù���Ա��¼��Ҫ�������������Ա�ʺ�����
            If Not blnLogin Then
                MsgBox "����Ա��Ȩ��Ϣ��ʧ������֤����Ա�˻���", vbExclamation, gstrSysName
                If Not frmUserCheckLogin.ShowLogin(UCT_SysOwner, gcnOracle, strUserName, strServerName) Then Exit Function
                strPassword = gobjRegister.GetPassword
                Call SetSQLTrace(strServerName, strUserName, gcnOracle)
                '������֤�Ự
                Call gobjRegister.zlRegInit(gcnOracle)
                strRegErr = gobjRegister.zlRegCheck(False)
                If strRegErr <> "" Then
                    MsgBox strRegErr, vbQuestion, "����"
                    If Not frmReg.ReReg Then
                        End
                    End If
                End If
                'δ��Ȩ���򲻸��¹���Ա��Ϣ
                If Not strPassword Like "δ��Ȩ�ĳ���:*" Then
                    '���¹���Ա�˻���Ϣ
                    strSQL = "Delete zltools.zlRegInfo where ��Ŀ='����Ա' Or ��Ŀ='��֤��'"
                    gcnOracle.Execute strSQL
                    strSQL = "Insert into zltools.zlRegInfo(��Ŀ,����) values('����Ա','" & strUserName & "')"
                    gcnOracle.Execute strSQL
                    
                    strPwdTxt = ""
                    ReDim Preserve StrJiami(0)
                    Call DES_Encode(StrConv(strPassword, vbFromUnicode), StrJiami, strUnit)
                    strPwdTxt = FuncByteTo16Code(StrJiami)
                    strSQL = "Insert into zltools.zlRegInfo(��Ŀ,����) values('��֤��','" & strPwdTxt & "')"
                    gcnOracle.Execute strSQL
                End If
            End If
            
            strSQL = "SELECT 1 FROM SESSION_ROLES WHERE ROLE='DBA'"
            Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "DBA�ж�")
            gblnDBA = Not rsTemp.EOF
            
            gblnOwner = True
        Else
            strPassword = gobjRegister.GetPassword
             'δ��Ȩ���򲻸��¹���Ա��Ϣ
            If Not strPassword Like "δ��Ȩ�ĳ���:*" Then
                strSQL = "Select 1 From zltools.zlRegInfo where ��Ŀ='����Ա'"
                Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��������Ȩģʽ")
                If rsTemp.RecordCount > 0 Then
                    strSQL = "Update zltools.zlRegInfo Set ����='" & strUserName & "' Where ��Ŀ='����Ա' And ����<>'" & strUserName & "'"
                    gcnOracle.Execute strSQL
                    '��֤��
                    strPwdTxt = ""
                    ReDim Preserve StrJiami(0)
                    Call DES_Encode(StrConv(strPassword, vbFromUnicode), StrJiami, strUnit)
                    strPwdTxt = FuncByteTo16Code(StrJiami)
                    strSQL = "Select 1 From zlRegInfo where ��Ŀ='��֤��'"
                    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��֤���ж�")
                    If rsTemp.RecordCount > 0 Then
                        strSQL = "Update zlRegInfo Set ����='" & strPwdTxt & "' Where ��Ŀ='��֤��'"
                    Else
                        strSQL = "Insert into zlRegInfo(��Ŀ,����) values('��֤��','" & strPwdTxt & "')"
                    End If
                    gcnOracle.Execute strSQL
                End If
            End If
            '��Ϊ��ϵͳ��¼����ֻ�����ɫ��Ȩ���û���Ȩ����ģ���Ȩ��
            If glngSysNo <> -1 Then
                gstrHaveProg = "0401,0402"
            End If
        End If
    End If

    OraDataOpen = True
    gstrUserName = strUserName
    gstrPassword = gobjRegister.GetPassword
    gstrServer = Trim(strServerName)
End Function

Private Sub cmdCancel_Click()
    Set gcnOracle = Nothing
    Unload Me
End Sub


Private Sub cmdSelect_Click()
    Dim strServer As String
    Dim p As POINTAPI
    
    p.x = txt���ݿ�.Left / Screen.TwipsPerPixelX
    p.y = (cmdSelect.Top + cmdSelect.Height) / Screen.TwipsPerPixelY
    ClientToScreen Me.hwnd, p
    
    strServer = frmServerSelect.GetServer(mcolServer, p.x * Screen.TwipsPerPixelX, p.y * Screen.TwipsPerPixelY, txt���ݿ�.Text)
    If strServer <> "" Then
        txt���ݿ�.Text = strServer
        txt���ݿ�.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    Dim LngStyle As Long
    
    '���õ�ǰ��������������ʾ
    LngStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    LngStyle = LngStyle Or WinStyle
    Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, LngStyle)
    
    ShowWindow Me.hwnd, 0 '������
    ShowWindow Me.hwnd, 1 '����ʾ
        
    If Len(txt�û�) <> 0 Then
        txt����.SetFocus
    End If
    If Trim(txt�û�.Text) <> "" And Trim(txt����.Text) <> "" Then Call cmdOK_Click
End Sub

Private Sub Form_Load()
    Dim strFileInfo As String
    Dim ArrCommand() As String
    
    On Error GoTo errH
    txt�û�.Text = GetSetting("ZLSOFT", "ע����Ϣ\��½��Ϣ", "MANAGER", "")
    txt���ݿ�.Text = GetSetting("ZLSOFT", "ע����Ϣ\��½��Ϣ", "SERVER", "")
    intTimes = 0
    
    Set mcolServer = LoadServer(strFileInfo)
    txt���ݿ�.ToolTipText = strFileInfo
    Call ApplyOEM_Picture(Me, "Icon")

    If Val(Me.Tag) = 1 Then
        Me.Hide
    Else
        '������һ��Ļ�����������ʾfrmSplash���壬�ڿ������뷨������£�����Դ���򣬲�����ʾ��¼���ڣ�VBֻ���쳣��ֹ�˳�
        SetActiveWindow Me.hwnd
    End If
    
    '�������/����ʾͬʱ�������û��������룬�������벻��Ҫ����ת��
    If mstrCommand <> "" Then
        ArrCommand = Split(mstrCommand, " ")
        If InStr(1, ArrCommand(0), "/") <> 0 And InStr(1, ArrCommand(0), ",") = 0 Then
            Me.txt�û�.Text = Split(ArrCommand(0), "/")(0)
            Me.txt����.Text = Split(ArrCommand(0), "/")(1)
            mblnת�� = False
        End If
    End If
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Set gcnOracle = Nothing
    End If
End Sub

Private Sub txt���ݿ�_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        '�س������д���
        If KeyAscii <> vbKeyBack Then
            Call AppendText(KeyAscii)
        End If
    End If
End Sub

Private Sub txt�û�_GotFocus()
    If Me.ActiveControl Is txt�û� Then
        SelAll txt�û�
        OpenIme False
    End If
End Sub

Private Sub TXT����_GotFocus()
    SelAll txt����
End Sub

Private Sub txt���ݿ�_GotFocus()
    If Me.ActiveControl Is txt���ݿ� Then
        SelAll txt���ݿ�
        OpenIme False
    End If
End Sub

Private Sub cmdSet_Click()
    Dim strPath As String   'Oracle��װĿ¼
    Dim strCommond As String, strError As String
    
    strPath = GetOracleHomePath(strError)
    If strPath = "" Then
        MsgBox "������Oracle�Ƿ�������װ�����顣" & vbCrLf & strError, vbInformation, "��ʾ"
        Exit Sub
    End If
    
    'ִ��Oracle 8 ��Net Easy���õĳ���
    strCommond = strPath & "\BIN\N8SW.EXE"
    If ExecuteCommand(strCommond) = True Then
        '�Ѿ��ɹ�
        Exit Sub
    End If
    
    'ִ��Oracle 8i,9i,10g,11g��Net Easy���õĳ���
    strCommond = strPath & "\BIN\launch.exe """ & strPath & "\network\tools"" " & strPath & "\network\tools\netca.cl"
    If ExecuteCommand(strCommond) = True Then
        '�Ѿ��ɹ�
        Exit Sub
    End If
    
End Sub

Private Function GetOracleHomePath(ByVal strError As String) As String
'���ܣ���ȡOracleHome·��
    Dim strPath As String
    Dim strServer As String, strComputer As String, strSID As String
    Dim arrTmp As Variant
    Dim rsOraHome As ADODB.Recordset
    Dim intVersion As Integer, intTimes As Integer, intServer As Integer
    Dim i As Long, blnRead As Boolean

    Set rsOraHome = New ADODB.Recordset
    With rsOraHome
        .Fields.Append "Name", adVarChar, 256 'Name
        .Fields.Append "VerSion", adInteger  '�汾
        .Fields.Append "Times", adInteger '�ڼ��ΰ�װ
        .Fields.Append "Server", adInteger '1-������,2-�ͻ���
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        '1:��ȡ64λ��32Ŀ¼���Զ���λ��SOFTWARE\Wow6432Node\Oracle 2����ȡ32λ��32λĿ¼
        arrTmp = GetAllSubKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Oracle")
        
        If TypeName(arrTmp) = "Empty" Then
            If Is64bit Then
                strError = "û���ҵ�ע�����HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Oracle��"
            Else
                strError = "û���ҵ�ע�����HKEY_LOCAL_MACHINE\SOFTWARE\Oracle��"
            End If
        Else
            For i = LBound(arrTmp) To UBound(arrTmp)
                If UCase(arrTmp(i)) Like "KEY_ORA*HOME*" Then
                    intVersion = 0: intTimes = 0:  intServer = 1
                    If GetOraInfoByRegKey(arrTmp(i), intVersion, intTimes, intServer) Then
                        .AddNew Array("Name", "VerSion", "Times", "Server"), Array("\" & arrTmp(i), intVersion, intTimes, intServer)
                        .Update
                    End If
                End If
            Next
            If UBound(arrTmp) <> -1 Then ''����Ŀ¼������Oracle_Home��Ϣ��Ĭ�϶�ȡ���
                .AddNew Array("Name", "VerSion", "Times", "Server"), Array("", 0, 0, 1): .Update
            End If
            .Sort = "VerSion Desc,Times Desc,Server"    '�߰汾����
            Do While Not .EOF
                strPath = ""
                blnRead = Not GetRegValue("HKEY_LOCAL_MACHINE\SOFTWARE\Oracle" & !name, "ORACLE_HOME", strPath)
                blnRead = blnRead Or strPath = "" And !name & "" = ""
                If blnRead Then
                    Call GetRegValue("HKEY_LOCAL_MACHINE\SOFTWARE\Oracle", "ORA_CRS_HOME", strPath)
                End If
                If strPath <> "" Then
                    GetOracleHomePath = strPath
                    Exit Function
                End If
                
                .MoveNext
            Loop
        End If
    End With
End Function

Private Function GetOraInfoByRegKey(ByVal strOraHome As String, ByRef intVer As Integer, ByRef intTimes As Integer, ByRef intServer As Integer) As Boolean
'����:ͨ��OracleHome����ȡOracle��Ϣ
    Dim arrTmp As Variant
    Dim i As Long, blnRetrun As Boolean
    'KEY_OraDb11g_home1_32bit
    'Key_Ora*�汾Home_32Bit
    'Key_Ora*�汾_Home*
    arrTmp = Split(UCase(strOraHome), "_")
    For i = 1 To UBound(arrTmp)
        If arrTmp(i) Like "HOME*" Then
            intTimes = ValEx(arrTmp(2))
            blnRetrun = True
        ElseIf arrTmp(i) Like "*HOME*" Then
            intTimes = Val(Mid(arrTmp(1), InStr(UCase(arrTmp(1)), "HOME") + 4))
            blnRetrun = True
        End If
        If arrTmp(i) Like "ORADB*" Then
            intVer = ValEx(Mid(arrTmp(1), 6))
            intServer = 1
            blnRetrun = True
        ElseIf arrTmp(i) Like "ORACLIENT*" Then
            intVer = ValEx(Mid(arrTmp(1), 10))
            intServer = 2
            blnRetrun = True
        ElseIf arrTmp(i) Like "*CLIENT*" Then
            intServer = 2
            intVer = ValEx(arrTmp(i))
            blnRetrun = True
        End If
    Next
    GetOraInfoByRegKey = blnRetrun
End Function

Private Function ExecuteCommand(ByVal strCommand As String) As Boolean
'���ܣ�ִ��ָ������
    Dim lngShell As Long, lngProcess As Long
    
    On Error Resume Next
    lngShell = Shell(strCommand, vbNormalFocus)
    
    If err <> 0 Then
        Exit Function
    End If
    
    ExecuteCommand = True
End Function

Private Sub AppendText(KeyAscii As Integer)
'���ܣ���TextBox�ؼ���Text׷�����ݣ������ݵ�ǰText��ֵ���б��м������õ�������Ŀ
'������KeyAscii    ��ǰ�İ���
    Dim strTemp As String
    Dim strInput As String
    Dim lngIndex As Long, lngStart As Long
    Dim varItem As Variant
    
    '���ȵ�ǰ�û�������ַ�
    If KeyAscii < 0 Or InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.", UCase(Chr(KeyAscii))) > 0 Then
        '�����ַ�ֻ�������֡�Ӣ�ĺͺ���
        strInput = Chr(KeyAscii)
        KeyAscii = 0
    End If
    
    With txt���ݿ�
        '��¼�ϴεĲ����λ��
        lngStart = .SelStart + IIf(strInput <> "", 1, 0)
        '���ŵõ��û�������ɺ��ı����г��ֵ�����
        strInput = Mid(.Text, 1, .SelStart) & strInput & Mid(.Text, .SelStart + .SelLength + 1)
    End With
    '���ݼ�������ݵõ����ܵ��б���
    strTemp = ""
    For Each varItem In mcolServer
        If UCase(varItem(0)) Like UCase(strInput & "*") Then
            strTemp = varItem(0)
        End If
    Next
    If strTemp <> "" Then
        txt���ݿ�.Text = strTemp
        txt���ݿ�.SelStart = Len(strInput)
        txt���ݿ�.SelLength = 100
    Else
        txt���ݿ�.Text = strInput
        txt���ݿ�.SelStart = lngStart
    End If

End Sub

Public Function Docmd(ByVal strCmd As String, ByRef blnAnalysis As Boolean) As Boolean
    '���ܣ�Shell���ʽ��¼������
    '����
    'strCmd��Shell�������
    '     blnAnalysis������Ե�һ�ַ�ʽ�����Ƿ�ɹ�
    '     blnAnalysis=True����ʾstrCmd�����ɹ�
    '     blnAnalysis=False����ʾstrCmd����ʧ��
    '��������в��������û��������룬����䲢ִ��
    Dim ArrCommand() As String
    Dim strUser As String, strPasswd As String, strServer As String
    Dim i As Long
    
    mblnAccess = False
    mblnת�� = True
    mstrCommand = strCmd
    ArrCommand = Split(strCmd, " ")
    If InStr(ArrCommand(0), "=") > 0 Then
        For i = LBound(ArrCommand) To UBound(ArrCommand)
            If UCase(ArrCommand(i)) Like "USER=*" Then
                strUser = Split(ArrCommand(i), "=")(1)
            ElseIf UCase(ArrCommand(i)) Like "PASS=*" Then
                strPasswd = Split(ArrCommand(i), "=")(1)
            ElseIf UCase(ArrCommand(i)) Like "SERVER=*" Then
                strServer = Split(ArrCommand(i), "=")(1)
            End If
        Next
    End If
    
    If strUser <> "" And strPasswd <> "" And strServer <> "" Then
        '��ʾ���Ե�һ��Shell�����ʽ��¼
        Me.Tag = 1
        blnAnalysis = True
        Me.txt�û�.Text = strUser
        Me.txt����.Text = strPasswd
        Me.txt���ݿ�.Text = strServer
        Call cmdOK_Click
    End If
    Docmd = mblnAccess
End Function
