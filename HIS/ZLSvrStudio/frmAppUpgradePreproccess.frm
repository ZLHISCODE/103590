VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAppUpgradePreproccess 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ǩǰ�ü��"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11655
   Icon            =   "frmAppUpgradePreproccess.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   11655
      TabIndex        =   1
      Top             =   0
      Width           =   11655
      Begin VB.Frame fraTop 
         Height          =   120
         Left            =   0
         TabIndex        =   5
         Top             =   840
         Width           =   11880
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   $"frmAppUpgradePreproccess.frx":6852
         Height          =   360
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   11400
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   11655
      TabIndex        =   0
      Top             =   6555
      Width           =   11655
      Begin VB.CommandButton cmdExit 
         Caption         =   "�˳�(&E)"
         Height          =   350
         Left            =   10440
         TabIndex        =   8
         Top             =   360
         Width           =   1100
      End
      Begin VB.CommandButton cmdRecheck 
         Caption         =   "���¼��(&R)"
         Height          =   350
         Left            =   7600
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdjust 
         Caption         =   "����(&A)"
         Height          =   350
         Left            =   9000
         TabIndex        =   6
         Top             =   360
         Width           =   1100
      End
      Begin VB.Frame fraBottom 
         Height          =   120
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   11880
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsCheckResult 
      Height          =   5820
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   11460
      _cx             =   20214
      _cy             =   10266
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
      ForeColorSel    =   0
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483628
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   100
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmAppUpgradePreproccess.frx":6908
      ScrollTrack     =   0   'False
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
      OutlineBar      =   4
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
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
      Begin MSComctlLib.ImageList imgEdit 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAppUpgradePreproccess.frx":6A20
               Key             =   "Check"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAppUpgradePreproccess.frx":6FBA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAppUpgradePreproccess.frx":7554
               Key             =   "ǩ��"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAppUpgradePreproccess.frx":78A6
               Key             =   "Woman"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAppUpgradePreproccess.frx":E108
               Key             =   "Man"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAppUpgradePreproccess.frx":1496A
               Key             =   "UnCheck"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAppUpgradePreproccess.frx":14E32
               Key             =   "AllCheck"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmAppUpgradePreproccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum SysCheck
    SC_������ = 0
    SC_��ǰֵ = 1
    SC_����ֵ = 2
    SC_���� = 3
    SC_���˵�� = 4
    SC_������ = 5
End Enum

Private mrsCheckInfo        As ADODB.Recordset
Private mstrUsers           As String
Private mbln10G             As Boolean
Private Const SQL_CAPTION = "��Ǩǰ�ü��"
Private mlngBeach           As Long
'ȥ��SYS��SYSTEM
Private Const mstrOracleUser      As String = "'ANONYMOUS','AURORA$JIS$UTILITY$','AURORA$ORB$UNAUTHENTICATED','CTXSYS','DBSNMP','DIP','DMSYS','DVF','DVSYS','EXFSYS','HR','LBACSYS','MDDATA','MDSYS','MGMT_VIEW','OAS_PUBLIC','ODM','ODM_MTR','OE','OGG','OLAPSYS','ORDPLUGINS','ORDSYS','OSE$HTTP$ADMIN','OUTLN','PERFSTAT','PM','QS','QS_ADM','QS_CB','QS_CBADM','QS_CS','QS_ES','QS_OS','QS_WS','REPADMIN','RMAN','SCOTT','SH','SI_INFORMTN_SCHEMA','SYSMAN','TRACESVR','TSMSYS','WEBSYS','WKPROXY','WKSYS','WKUSER','WK_TEST','WMSYS','XDB'"
Private mblnOK              As Boolean
Private mblnExecBefore      As Boolean
Private Enum CheckClass
    CC_SYSPARA = 0
    CC_DBFile = 1
    CC_AutoJob = 2
    CC_Trigger = 3
    CC_Scheduler = 4
    CC_Privs = 5
End Enum

Private Enum ChceckType
    CT_CheckAndLoad = 0
    CT_OnlyCheck = 1
    CT_OnlyLoad = 2
End Enum
'******************************************************************************************************************
'���ܣ����ϵͳ״̬�ܷ������Ǩ
'blnExecBefore-�Ƿ���ǰ����
'���أ�TRUE-���Խ�����Ǩ��false-���ܽ�����Ǩ
'******************************************************************************************************************
Public Function ShowMe(ByVal blnExecBefore As Boolean) As Boolean
    mblnExecBefore = blnExecBefore
    If mblnExecBefore Then ShowMe = True: Exit Function
    mstrUsers = GetUsers
    mbln10G = GetOracleVersion(True, True) < 11
    mblnOK = False
    If Not gblnDBA Then
        MsgBox """" & gstrUserName & """����DBA���޷���������ǰ��飬��Ҫ��������ǰ��飬������""" & gstrUserName & """DBAȨ�ޡ�", vbInformation, gstrSysName
        ShowMe = True
        Exit Function
    End If
    If Not LoadAllCheck(CT_OnlyCheck) Then
        ShowMe = True
        Exit Function
    End If
    '���������⣬����������ʾ�û�����
    If mrsCheckInfo.RecordCount <> 0 Then
        Me.Show vbModal, frmMDIMain
        ShowMe = mblnOK
    Else
        ShowMe = True
    End If
End Function
'******************************************************************************************************************
'���ܣ���鲢���ؼ����
'������intType=0����鲢���أ�1-ֻ��飬�����ڣ�2-ֻ���ز����
'******************************************************************************************************************
Private Function LoadAllCheck(Optional ByVal intType As Integer) As Boolean
    On Error GoTo errH
    If intType < 2 Then
        mlngBeach = 0
        Set mrsCheckInfo = CopyNewRec(Nothing, True, , Array("�������", adInteger, Empty, Empty, "������", adVarChar, 100, Empty, _
                            "�������", adInteger, Empty, Empty, "������", adVarChar, 100, Empty, "������", adVarChar, 100, Empty, "����", adVarChar, 100, Empty, _
                            "��ǰֵ", adVarChar, 100, Empty, "����ֵ", adVarChar, 100, Empty, _
                            "����SQL", adVarChar, 100, Empty, "��������", adInteger, Empty, Empty, "�Ƿ����", adInteger, Empty, Empty, "�Ƿ�DBA", adInteger, Empty, Empty, _
                            "���˵��", adVarChar, 200, Empty, "������", adVarChar, 200, Empty, "�Ƿ�������", adInteger, Empty, Empty))
        '�������� 0-���ɵ����Ҳ�����������1-�����õ���������������2-�ɵ���
        Call ShowFlash("���ڼ��ϵͳ���������Ժ�")
        If Not CheckSysPara Then Call ShowFlash(""): Exit Function
        Call ShowFlash("���ڼ�����ݿ��ļ������Ժ�")
        If Not CheckDBFile Then Call ShowFlash(""): Exit Function
        Call ShowFlash("���ڼ���Զ���ҵ��ϵͳ���ȣ����Ժ�")
        If Not CheckAutoJobs Then Call ShowFlash(""): Exit Function
        Call ShowFlash("���ڼ�鴥���������Ժ�")
        If Not CheckTriggers Then Call ShowFlash(""): Exit Function
        Call ShowFlash("���ڶ���Ȩ�ޣ����Ժ�")
        If Not CheckPrivs Then Call ShowFlash(""): Exit Function
        Call ShowFlash("")
    End If
    If intType <> 1 Then
        Call LoadCheckInfo
    End If
    LoadAllCheck = True
    Exit Function
errH:
    Call ShowFlash("")
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function


'******************************************************************************************************************
'���ܣ���ȡZLHISϵͳ��������
'******************************************************************************************************************
Private Function GetUsers() As String
    Dim strSQL  As String
    Dim rsTmp   As ADODB.Recordset
    
    On Error GoTo errH
    strSQL = "Select f_List2str(Cast(Collect(Chr(39) || ������ || Chr(39)) As t_Strlist)) ������" & vbNewLine & _
            "From (Select Upper(������) ������" & vbNewLine & _
            "       From zlBakSpaces" & vbNewLine & _
            "       Union" & vbNewLine & _
            "       Select Upper(������)" & vbNewLine & _
            "       From zlSystems" & vbNewLine & _
            "       Union" & vbNewLine & _
            "       Select 'ZLTOOLS'" & vbNewLine & _
            "       From Dual)"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, SQL_CAPTION)
    GetUsers = rsTmp!������ & ""
    Exit Function
errH:
    Call ShowFlash("")
    MsgBox "��ȡ��׼ϵͳ�û����ִ���" & err.Description, vbInformation, gstrSysName
    err.Clear
End Function
'******************************************************************************************************************
'���ܣ������Ŀ��ӵ���¼��
'������strOwner=��������������
'      strName=�����������
'      strItem=�������չʾ����
'      strCurValue=�������ǰ״̬
'      lngCheckClass=�������ķ���ID
'      strSuggestiveValue=�������Ľ���״̬
'      strAdjustSQL=������������SQL,������ʱ�������Զ��޸�
'      strCheckInfo=�������ļ��˵��
'      strAdjustWarn=���������������棬���ֶ�����Ҫ���⴦����ֹ��������ڴ˴�
'      blnIgnor=���������Ƿ���Ժ���
'      blnDBA=�������������Ƿ���ҪDBA��ݣ�û������SQL�Ķ�����ע�⴫NULL
'******************************************************************************************************************
Private Sub AddCheckItem(ByVal strOwner As String, strName As String, ByVal strItem As String, ByVal strCurValue As String, ByVal lngCheckClass As Long, ByVal strCheckClass As String, _
                        ByVal strSuggestiveValue As String, ByVal strAdjustSQL As String, ByVal strCheckInfo As String, _
                        ByVal strAdjustWarn As String, Optional ByVal blnIgnor As Boolean, Optional ByVal blnDBA As Boolean)
     mrsCheckInfo.AddNew Array("�������", "������", "�������", "������", "������", "����", "��ǰֵ", "����ֵ", "����SQL", "��������", "�Ƿ�DBA", "���˵��", "������", "�Ƿ����", "�Ƿ�������"), _
                        Array(lngCheckClass, strCheckClass, mrsCheckInfo.RecordCount + 1, strItem, strOwner, strName, strCurValue, strSuggestiveValue, IIf(strAdjustSQL = "", Null, strAdjustSQL), IIf(blnIgnor, IIf(strAdjustSQL = "", 1, 2), 0), IIf(blnDBA, 1, 0), strCheckInfo, IIf(strAdjustSQL <> "" And strAdjustWarn = "", "�Զ�����", strAdjustWarn), IIf(strAdjustSQL = "", 0, 1), 0)

End Sub
'******************************************************************************************************************
'���ܣ�������ݿ����
'******************************************************************************************************************
Private Function CheckSysPara() As Boolean
    Dim strSQL  As String
    Dim rsTmp   As ADODB.Recordset
    
    On Error GoTo errH
    strSQL = "Select Name , Value From V$parameter Where Name =[1] And Value =[2]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, SQL_CAPTION, "optimizer_index_cost_adj", "100")
    If Not rsTmp.EOF Then Call AddCheckItem("", rsTmp!name & "", rsTmp!name & "", rsTmp!value & "", CC_SYSPARA, "���ݿ����", "20", "alter system set " & rsTmp!name & "=20", "ȱʡֵ100�ᵼ�²�Ʒ��������", "", True, True)
    
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, SQL_CAPTION, "optimizer_index_caching", "0")
    If Not rsTmp.EOF Then Call AddCheckItem("", rsTmp!name & "", rsTmp!name & "", rsTmp!value & "", CC_SYSPARA, "���ݿ����", "80", "alter system set " & rsTmp!name & "=80", "ȱʡ0�ᵼ�²�Ʒ��������", "", True, True)
    
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, SQL_CAPTION, "O7_DICTIONARY_ACCESSIBILITY", "FALSE")
    If Not rsTmp.EOF Then Call AddCheckItem("", rsTmp!name & "", rsTmp!name & "", rsTmp!value & "", CC_SYSPARA, "���ݿ����", "TRUE", "", "����ϵͳ��ͼ�޷���Ȩ��Ӱ�������Լ���Ʒ����", "���ֹ�����ΪTRUE���������ݿ�", False, False)

    strSQL = "Select Name , Value From V$parameter Where Name = [1] And Zl_To_Number(Value) < [2]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, SQL_CAPTION, "log_buffer", "104857600")
    If Not rsTmp.EOF Then Call AddCheckItem("", rsTmp!name & "", rsTmp!name & "", Int(Val(rsTmp!value & "") / 1024 / 1024) & "M", CC_SYSPARA, "���ݿ����", ">=100M", "", "Ӱ��ϵͳ��������������Ч��", "���ֹ�����Ϊ����100M���������ݿ�", True, False)
    
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, SQL_CAPTION, "parallel_execution_message_size", "8192")
    If Not rsTmp.EOF Then Call AddCheckItem("", rsTmp!name & "", rsTmp!name & "", rsTmp!value & "", CC_SYSPARA, "���ݿ����", ">=8192", "", "Ӱ��ϵͳ��������ִ��", "���ֹ�����Ϊ8192���������ݿ�", True, False)
    
    CheckSysPara = True
    Exit Function
errH:
    Call ShowFlash("")
    MsgBox "������ݿ�������ִ���" & err.Description, vbInformation, gstrSysName
    err.Clear
End Function
'******************************************************************************************************************
'���ܣ������־�ļ�
'******************************************************************************************************************
Private Function CheckDBFile() As Boolean
    Dim strSQL  As String
    Dim rsTmp   As ADODB.Recordset
    Dim strFile As String
    
    On Error GoTo errH
    strSQL = "Select 'INST_ID:' || a.Inst_Id || ',GROUP:' || a.Group# Name,b.Member," & vbNewLine & _
            "       a.Bytes Value" & vbNewLine & _
            "From Gv$log A" & vbNewLine & _
            "Join Gv$logfile B" & vbNewLine & _
            "On (a.Group# = b.Group# And a.Inst_Id = b.Inst_Id)" & vbNewLine & _
            "Where a.Bytes < 104857600" & vbNewLine & _
            "Order By a.Inst_Id, a.Group#, a.Thread#, b.Member"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, SQL_CAPTION)
    Do While Not rsTmp.EOF
        strFile = GetFileNameByPath(rsTmp!Member & "")
        Call AddCheckItem("", rsTmp!name & "," & strFile, rsTmp!name & "," & strFile, Int(Val(rsTmp!value & "") / 1024 / 1024) & "M", CC_DBFile, "���ݿ��ļ�", ">=100M", "", "Ӱ��ϵͳ��������������Ч��", "���ֹ�����Ϊ����100M", True, False)
        rsTmp.MoveNext
    Loop
    CheckDBFile = True
    Exit Function
errH:
    Call ShowFlash("")
    MsgBox "�����־�ļ����ִ���" & err.Description, vbInformation, gstrSysName
    CheckDBFile = True
    err.Clear
End Function
'******************************************************************************************************************
'���ܣ�����Զ���ҵ
'******************************************************************************************************************
Private Function CheckAutoJobs() As Boolean
    Dim strSQL          As String
    Dim rsTmp           As ADODB.Recordset
    Dim strProcName     As String
    
    On Error GoTo errH
    '���ִ��ʱ���ڵ�ǰʱ������2Сʱ������5Сʱ���Ծ�δ��ֹ���Զ���ҵ
    strSQL = "Select a.Job, a.Broken, a.Schema_User, Upper(What) What" & vbNewLine & _
            "From Dba_Jobs A" & vbNewLine & _
            "Where a.Job In (Select ��ҵ�� From Zltools.Zlautojobs) And a.Broken = 'N' And" & vbNewLine & _
            "      Nvl(a.Next_Date, Sysdate + 10) Between Sysdate - 2 / 24 And Sysdate + 5 / 24"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, SQL_CAPTION)
    
    Do While Not rsTmp.EOF
        Call AddCheckItem(rsTmp!Schema_User & "", rsTmp!Job & "", rsTmp!Job & "(" & rsTmp!What & ")", rsTmp!Broken, CC_AutoJob, "�Զ���ҵ", "BROKEN=Y", "Dbms_Job.Broken(" & rsTmp!Job & ", True)", "Ӱ����ر�����������ű�ִ��Ч��", "�����ڼ����", True, rsTmp!Schema_User & "" = "SYS" Or rsTmp!Schema_User & "" = "SYSTEM")
        rsTmp.MoveNext
    Loop
    '��ZLHIS������Զ���ҵ
    strSQL = "Select a.Job, a.Broken, a.Schema_User, Upper(What) What" & vbNewLine & _
            "From Dba_Jobs A" & vbNewLine & _
            "Where a.Schema_User Not In (" & mstrOracleUser & ") And a.Job Not In (Select ��ҵ�� From Zltools.Zlautojobs) And" & vbNewLine & _
            "      a.Broken = 'N' And Nvl(a.Next_Date, Sysdate + 10) Between Sysdate - 2 / 24 And Sysdate + 5 / 24"

    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, SQL_CAPTION)
    
    Do While Not rsTmp.EOF
        strProcName = GetJobProcedure(rsTmp!What & "")
        If strProcName <> "" Then
            If ObjectReferencedZLHIS(rsTmp!Schema_User, strProcName) Then
                Call AddCheckItem(rsTmp!Schema_User & "", rsTmp!Job & "", rsTmp!Job & "(" & rsTmp!What & ")", rsTmp!Broken, CC_AutoJob, "�Զ���ҵ", "BROKEN=Y", "Dbms_Job.Broken(" & rsTmp!Job & ", True)", "Ӱ����ر�����������ű�ִ��Ч��", "�����ڼ����", True, rsTmp!Schema_User & "" = "SYS" Or rsTmp!Schema_User & "" = "SYSTEM")
            End If
        End If
        rsTmp.MoveNext
    Loop
    
    CheckAutoJobs = True
    Exit Function
errH:
    Call ShowFlash("")
    MsgBox "����Զ���ҵ���ִ���" & err.Description, vbInformation, gstrSysName
    err.Clear
End Function
'******************************************************************************************************************
'���ܣ���鷢��
'******************************************************************************************************************
Private Function CheckTriggers() As Boolean
    Dim strSQL          As String
    Dim rsTmp           As ADODB.Recordset
    
    On Error GoTo errH
    'ZLHIS���еĴ��������ý�һ���ж϶���
    If CheckAndAdjustMustTable("ZLTABLES") Then
        strSQL = "Select a.Owner, a.Trigger_Name, a.Status" & vbNewLine & _
                "From Dba_Triggers A" & vbNewLine & _
                "Where a.Table_Name In (Select b.���� From Zltables B Where b.���� Not Like 'A%') And a.Status = 'ENABLED' And a.Trigger_Type <> 'INSTEAD OF' And" & vbNewLine & _
                "      a.Table_Owner In (" & mstrUsers & ")"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, SQL_CAPTION)
    Else
        strSQL = "Select a.Owner, a.Trigger_Name,a.Status From Dba_Triggers A Where a.Status = 'ENABLED' And a.Table_Owner In (" & mstrUsers & ") And a.Trigger_Type <> 'INSTEAD OF'"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, SQL_CAPTION)
    End If
    Do While Not rsTmp.EOF
        Call AddCheckItem(rsTmp!Owner & "", rsTmp!trigger_name & "", rsTmp!Owner & "." & rsTmp!trigger_name, "ENABLED", CC_Trigger, "������", "DISABLED", "alter trigger " & rsTmp!Owner & "." & rsTmp!trigger_name & " disable", "Ӱ��ñ�����������ű�ִ��Ч��", "�����ڼ����", True, rsTmp!Owner & "" = "SYS" Or rsTmp!Owner & "" = "SYSTEM")
        rsTmp.MoveNext
    Loop
    
    CheckTriggers = True
    Exit Function
errH:
    Call ShowFlash("")
    MsgBox "��鴥�������ִ���" & err.Description, vbInformation, gstrSysName
    err.Clear
End Function
'******************************************************************************************************************
'���ܣ���������û��Ķ���Ȩ��
'******************************************************************************************************************
Private Function CheckPrivs() As Boolean
    Dim strSQL          As String
    Dim rsTmp           As ADODB.Recordset
    
    On Error GoTo errH
    '������ʹ�øñ������������
    '���ZLTOOLS��PUBLICȨ��
    strSQL = "Select a.Grantee, a.Owner, a.Table_Name, a.Privilege" & vbNewLine & _
            "From (Select 'ZLTOOLS' Grantee, 'SYS' Owner, 'DBA_ROLE_PRIVS' Table_Name, 'SELECT' Privilege From Dual) A" & vbNewLine & _
            "Where Not Exists (Select 1" & vbNewLine & _
            "       From Dba_Tab_Privs C" & vbNewLine & _
            "       Where c.Owner = 'SYS' And (c.Grantee = 'PUBLIC' Or a.Grantee<>'PUBLIC' And c.Grantee = 'ZLTOOLS') And" & vbNewLine & _
            "             c.Table_Name = a.Table_Name And c.Privilege = a.Privilege)"

    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, SQL_CAPTION)
    Do While Not rsTmp.EOF
        Call AddCheckItem(rsTmp!Owner & "", rsTmp!Table_Name & "", rsTmp!Grantee & " " & rsTmp!Privilege & " On " & rsTmp!Owner & "." & rsTmp!Table_Name, "", CC_Privs, "����Ȩ��", rsTmp!Privilege & "", "Grant " & rsTmp!Privilege & " On " & rsTmp!Owner & "." & rsTmp!Table_Name, "�������ܻ�����쳣���Լ�Ӱ���Ʒʹ��", "", False, rsTmp!Owner & "" = "SYS")
        rsTmp.MoveNext
    Loop
    CheckPrivs = True
    Exit Function
errH:
    Call ShowFlash("")
    MsgBox "������Ȩ�޳��ִ���" & err.Description, vbInformation, gstrSysName
    err.Clear
End Function
'******************************************************************************************************************
'���ܣ����һ�������Ƿ�������ZLHISϵͳ�Ļ�������(��),��Ϊ�洢���ݵĻ����������øö�����ϼ�������ܻᵼ�����������ֻ����
'������strOwner-����������
'      strObjectName=��������
'      strObjectType=�������ͣ�������ʱ��Ĭ�ϲ�Ϊͬ���
'���أ�TRUE-������ZLHIS�Ķ���false-δ����ZLHIS����
'******************************************************************************************************************
Private Function ObjectReferencedZLHIS(ByVal strOwner As String, ByVal strObjectName As String, Optional ByVal strObjectType As String) As Boolean
    Dim strSQL  As String
    Dim rsTmp   As ADODB.Recordset
    
    On Error GoTo errH
    strSQL = "Select Count(1) ����" & vbNewLine & _
            "From (Select a.Owner, a.Name, a.Type, a.Referenced_Owner, a.Referenced_Name, a.Referenced_Type" & vbNewLine & _
            "       From All_Dependencies A" & vbNewLine & _
            "       Start With a.Owner = [1] And a.Name = [2] And a.Type " & IIf(strObjectType = "", "<>'SYNONYM'", "= [3]") & vbNewLine & _
            "       Connect By Prior a.Referenced_Owner = a.Owner And Prior a.Referenced_Name = a.Name And" & vbNewLine & _
            "                  Prior a.Referenced_Type = a.Type) B" & vbNewLine & _
            "Where b.Referenced_Owner In (" & mstrUsers & ") And b.Referenced_Type = 'TABLE' Or b.Owner In (" & mstrUsers & ") And b.Type = 'TABLE'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, SQL_CAPTION, strOwner, strObjectName, strObjectType)
    ObjectReferencedZLHIS = rsTmp!���� <> 0
    Exit Function
errH:
    Call ShowFlash("")
    MsgBox "�������������ִ���" & err.Description, vbInformation, gstrSysName
    err.Clear
End Function
'******************************************************************************************************************
'���ܣ�����JOBִ�����ݣ���ȡִ�еĴ洢����
'������strWhat-JOBִ������
'���أ�JOBִ�еĴ洢����
'******************************************************************************************************************
Private Function GetJobProcedure(ByVal strWhat As String) As String
    Dim arrTmp          As Variant
    Dim strProcedure    As String
    '��������EMD_MAINTENANCE.EXECUTE_EM_DBMS_JOB_PROCS();
    arrTmp = Split(strWhat & ";", ";")
    strProcedure = arrTmp(0)
    arrTmp = Split(strProcedure & "(", "(")
    strProcedure = arrTmp(0)
    arrTmp = Split(strProcedure & ".", ".") '.�ָ����(֮����Ϊ���ܴ��Σ������д���.
    strProcedure = arrTmp(0)
    GetJobProcedure = strProcedure
End Function
'******************************************************************************************************************
'���ܣ������ļ�·����ȡ�ļ���
'������strFilePath-�ļ�·����������Linuxϵͳ�µ�·��
'���أ��ļ�����
'******************************************************************************************************************
Private Function GetFileNameByPath(ByVal strFilePath As String) As String
    Dim lngPos  As Long
    
    lngPos = InStrRev(strFilePath, "/")
    If lngPos = 0 Then
        lngPos = InStrRev(strFilePath, "\")

    End If
    If lngPos = 0 Then
        GetFileNameByPath = strFilePath
    Else
        GetFileNameByPath = Mid(strFilePath, lngPos + 1)
    End If
End Function


'******************************************************************************************************************
'���ܣ����Ѿ������ϵͳ���ȡ��Զ���ҵ����������¼�����ݿ�
'******************************************************************************************************************
Private Sub AdjustRecordToDB()
    Dim blnAutoJobs     As Boolean
    Dim strScheduler    As String
    Dim strSQL          As String
    Dim strJobs         As String
    
    On Error GoTo errH
    With mrsCheckInfo
        '������Ĵ�������¼�����ݿ�
        .Filter = "�������=" & CC_Trigger & " And �Ƿ�������=" & mlngBeach
        .Sort = "�������,�������"
        If Not .EOF Then
            Call SetUpgradeConfig("������״̬", "0")
            Do While Not .EOF
                strSQL = strSQL & " Union All Select '" & !������ & "' ������,'" & !���� & "' ���� From Dual"
                .MoveNext
            Loop
            strSQL = Mid(strSQL, Len(" Union All ") + 1)
            strSQL = "Insert Into Zltriggers (������, ����) Select ������, ����" & vbNewLine & _
                    "From (" & strSQL & ") A" & vbNewLine & _
                    "Where Not Exists (Select 1 From Zltriggers B Where b.���� = a.���� And b.������ = a.������)"
            gcnOracle.Execute strSQL, , adCmdText
        End If
        .Filter = "�������=" & CC_AutoJob & " And �Ƿ�������=" & mlngBeach
        .Sort = "�������,�������"
        blnAutoJobs = False
        If Not .EOF Then
            blnAutoJobs = True
            Do While Not .EOF
                If Not SetAutoJobs(Val(!���� & "")) Then
                    strJobs = strJobs & "," & Val(!���� & "")
                End If
                .MoveNext
            Loop
            If strJobs <> "" Then
                Call SetUpgradeConfig("���õ��Զ���ҵ", Mid(strJobs, 2))
            End If
        End If
        '�������ϵͳ�����Լ��Զ���ҵ��¼�����ݿ�
        .Filter = "�������=" & CC_Scheduler & " And �Ƿ�������=" & mlngBeach
        .Sort = "�������,�������"
        If Not .EOF Then
            blnAutoJobs = True
            Do While Not .EOF
                strScheduler = strScheduler & ",'" & !���� & "'"
                
                .MoveNext
            Loop
            Call SetUpgradeConfig("���õ�ϵͳ����", Mid(strScheduler, 2))
        End If
        If blnAutoJobs Then Call SetUpgradeConfig("��̨��ҵ״̬", "0")
    End With
    Exit Sub
errH:
    MsgBox "�Զ��������ִ���" & err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

'******************************************************************************************************************
'���ܣ���ϵͳ���ȡ��Զ���ҵ�����������ü�¼�����ݿ�
'����:strConfigName-��������
'     strConfigValue=����ֵ
'******************************************************************************************************************
Private Sub SetUpgradeConfig(ByVal strConfigName As String, ByVal strConfigValue As String)
    Dim strSQL      As String
    Dim lngAffect   As Long
    Dim strTmp      As String, rsTmp    As ADODB.Recordset
    
    On Error GoTo errH
    If strConfigName = "���õ�ϵͳ����" Or strConfigName = "���õ��Զ���ҵ" Then
        strSQL = "Select ���� From Zlupgradeconfig Where ��Ŀ = '" & strConfigName & "'"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
        If rsTmp.EOF Then
            strSQL = "Insert Into ZLTOOLS.zlUpgradeConfig(��Ŀ,����) values('" & strConfigName & "'," & strConfigValue & ")"
            gcnOracle.Execute strSQL, lngAffect, adCmdText
        Else
            If rsTmp!���� & "" = "" Then
                strTmp = strConfigValue
            Else
                strTmp = rsTmp!���� & "," & strConfigValue
            End If
            strSQL = "Update Zlupgradeconfig Set ���� = " & strConfigValue & " Where ��Ŀ = '" & strConfigName & "'"
            gcnOracle.Execute strSQL, lngAffect, adCmdText
        End If
    Else
        strSQL = "Update Zlupgradeconfig Set ���� = " & strConfigValue & " Where ��Ŀ = '" & strConfigName & "'"
        gcnOracle.Execute strSQL, lngAffect, adCmdText
        If lngAffect = 0 Then
            strSQL = "Insert Into ZLTOOLS.zlUpgradeConfig(��Ŀ,����) values('" & strConfigName & "'," & strConfigValue & ")"
            gcnOracle.Execute strSQL, lngAffect, adCmdText
        End If
    End If
    Exit Sub
errH:
    MsgBox "�Զ��������ִ���" & err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

'******************************************************************************************************************
'���ܣ����ͣ�õĺ�̨�Զ���ҵ
'����:strJobNum=�Զ���ҵ��
'���أ�True-�����������Զ���ҵ��FALSE-�������������Զ���ҵ
'******************************************************************************************************************
Private Function SetAutoJobs(ByVal strJobNum As String) As Boolean
    Dim strSQL      As String
    Dim lngAffect   As Long
    
    On Error GoTo errH
    strSQL = "Update zlAutoJobs Set ϵͳ����ͣ�� = 1 Where ��ҵ�� = " & strJobNum
    gcnOracle.Execute strSQL, lngAffect, adCmdText
    SetAutoJobs = lngAffect <> 0
    Exit Function
errH:
    MsgBox "�Զ��������ִ���" & err.Description, vbInformation, gstrSysName
    err.Clear
End Function

'******************************************************************************************************************
'���ܣ�����������ص�����
'******************************************************************************************************************
Private Sub LoadCheckInfo()
    Dim lngRow          As Long, lngLastClassRow    As Long
    Dim lngLastClass    As Long, blnHideClass       As Boolean
    Dim blnSelALl       As Boolean, lngCanCheck     As Long, lngAllCanCheck As Long
    With vsCheckResult
        .Redraw = False
        .OutlineCol = 0
        mrsCheckInfo.Filter = ""
        mrsCheckInfo.Sort = "�������,�������"
        .Rows = vsCheckResult.FixedRows
        lngLastClass = -1
        lngCanCheck = 0
        Do While Not mrsCheckInfo.EOF
            .Rows = .Rows + 1
            lngRow = .Rows - 1
            If lngLastClass <> mrsCheckInfo!������� Then
                If lngLastClass <> -1 Then
                    .RowHidden(lngLastClassRow) = blnHideClass
                    If lngCanCheck > 0 Then
                        Set .Cell(flexcpPicture, lngLastClassRow, SC_����) = imgEdit.ListImages(IIf(blnSelALl, "AllCheck", "UnCheck")).Picture
                        .Cell(flexcpData, lngLastClassRow, SC_����) = IIf(blnSelALl, 1, 0)
                    Else
                        .Cell(flexcpData, lngLastClassRow, SC_����) = 0
                    End If
                End If
                blnHideClass = True
                blnSelALl = True
                
                .Cell(flexcpData, lngRow, SC_������) = Val(mrsCheckInfo!������� & "")
                .TextMatrix(lngRow, SC_������) = mrsCheckInfo!������
                .IsSubtotal(lngRow) = True
                .RowOutlineLevel(lngRow) = 1
                .RowData(lngRow) = 0
                .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1) = &H8000000F
                
                lngLastClassRow = lngRow
                lngLastClass = mrsCheckInfo!�������
                lngCanCheck = 0
                .Rows = .Rows + 1
                lngRow = .Rows - 1
            End If
            .TextMatrix(lngRow, SC_������) = mrsCheckInfo!������
            .TextMatrix(lngRow, SC_��ǰֵ) = mrsCheckInfo!��ǰֵ
            .Cell(flexcpData, lngRow, SC_����) = Val(mrsCheckInfo!�������� & "")
            If mrsCheckInfo!�������� = 0 Then
                .Cell(flexcpForeColor, lngRow, SC_������, lngRow, SC_������) = &HFF0000
            End If
            If Val(mrsCheckInfo!�������� & "") = 2 Then
                .Cell(flexcpChecked, lngRow, SC_����, lngRow, SC_����) = IIf(mrsCheckInfo!�Ƿ���� = 0, flexUnchecked, flexChecked)
                If mrsCheckInfo!�Ƿ���� = 0 Then blnSelALl = False
                If mrsCheckInfo!�Ƿ������� = 0 Then
                    lngCanCheck = lngCanCheck + 1
                    lngAllCanCheck = lngAllCanCheck + 1
                End If
            Else
                .Cell(flexcpChecked, lngRow, SC_����, lngRow, SC_����) = 0
            End If
            .TextMatrix(lngRow, SC_����ֵ) = mrsCheckInfo!����ֵ
            .TextMatrix(lngRow, SC_���˵��) = mrsCheckInfo!���˵��
            .TextMatrix(lngRow, SC_������) = mrsCheckInfo!������
            
            .RowOutlineLevel(lngRow) = 2
            .IsSubtotal(lngRow) = True
            .RowData(lngRow) = Val(mrsCheckInfo!������� & "")
            .RowHidden(lngRow) = mrsCheckInfo!�Ƿ������� > 0
            If Not .RowHidden(lngRow) Then blnHideClass = False

            mrsCheckInfo.MoveNext
        Loop
        
        If lngLastClass <> -1 Then
            If lngCanCheck > 0 Then
                Set .Cell(flexcpPicture, lngLastClassRow, SC_����) = imgEdit.ListImages(IIf(blnSelALl, "AllCheck", "UnCheck")).Picture
                .Cell(flexcpData, lngLastClassRow, SC_����) = IIf(blnSelALl, 1, 0)
            Else
                .Cell(flexcpData, lngLastClassRow, SC_����) = 0
            End If
            .Cell(flexcpPictureAlignment, .FixedRows, SC_����, .Rows - 1, SC_����) = flexAlignCenterCenter
            .RowHidden(lngLastClassRow) = blnHideClass
        End If
        cmdAdjust.Enabled = lngAllCanCheck > 0
        If lngAllCanCheck <> 0 Then
            Set .Cell(flexcpPicture, 0, SC_����) = imgEdit.ListImages("AllCheck").Picture
            .ColData(SC_����) = 1
        End If
        .Redraw = True
    End With
End Sub

Private Sub cmdAdjust_Click()
    Dim cnDBA       As ADODB.Connection
    Dim cnTmp       As ADODB.Connection
    Dim blnNewBeach As Boolean
    mrsCheckInfo.Filter = "�Ƿ�DBA=1 And �Ƿ����=1 And ����SQL<>NULL And �Ƿ�������=0"
    If mrsCheckInfo.RecordCount <> 0 Then
        Set cnDBA = GetConnection("SYSTEM")
        If cnDBA Is Nothing Then Exit Sub
    End If
    With mrsCheckInfo
        .Filter = "�Ƿ����=1 And ����SQL<>NULL And �Ƿ�������=0"
        .Sort = "�������,�������"
        Do While Not .EOF
            If !�Ƿ�DBA = 0 Then
                Set cnTmp = gcnOracle
            Else
                Set cnTmp = cnDBA
            End If
            If ExecuteCmdText(!����SQL, Me.Caption, cnTmp, !������� = CC_AutoJob, True) = "" Then
                If Not blnNewBeach Then
                    mlngBeach = mlngBeach + 1
                    blnNewBeach = True
                End If
                .Update "�Ƿ�������", mlngBeach
            End If
            .MoveNext
        Loop
    End With
    If blnNewBeach Then Call AdjustRecordToDB
    mrsCheckInfo.Filter = "�Ƿ����=1 And �Ƿ�������=0"
    If mrsCheckInfo.RecordCount = 0 Then
        mblnOK = True
        Unload Me
        Exit Sub
    End If
    Call LoadAllCheck(CT_OnlyLoad)
End Sub

Public Function ExecuteCmdText(ByVal strSQL As String, ByVal strFormCaption As String, Optional cnOracle As ADODB.Connection, Optional ByVal blnProcedure As Boolean, Optional ByVal blnErrResume As Boolean) As String
'���ܣ�ִ���޷���ֵ���
    If blnErrResume Then
        On Error Resume Next
    Else
        On Error GoTo errH
    End If
    If blnProcedure Then
        cnOracle.Execute strSQL, , adCmdStoredProc
    Else
        cnOracle.Execute strSQL, , adCmdText
    End If
    If err.Number <> 0 Then
        ExecuteCmdText = err.Description
        err.Clear
    End If
    Exit Function
errH:
    ExecuteCmdText = err.Description
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Function

Private Sub cmdExit_Click()
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdRecheck_Click()
    Call LoadAllCheck(CT_CheckAndLoad)
End Sub

Private Sub Form_Load()
    Call LoadAllCheck(CT_OnlyLoad)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        If MsgBox("����ɵ�����ȷ����ǨԤ���������޷������������Ƿ��˳���", vbInformation + vbYesNo, gstrSysName) = vbNo Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsCheckInfo = Nothing
End Sub

Private Sub vsCheckResult_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim blnSelALl   As Boolean
    Dim lngClassRow As Long, i      As Long
    
    With vsCheckResult
        blnSelALl = True
        If Col <> SC_���� Then Exit Sub
        Call RecUpdate(mrsCheckInfo, "�������=" & Val(.RowData(Row)), "�Ƿ����", IIf(.Cell(flexcpChecked, Row, SC_����, Row, SC_����) = flexUnchecked, 0, 1))
        For i = Row + 1 To .Rows - 1
            If .RowData(i) = 0 Then
                Exit For
            Else
                If .Cell(flexcpData, i, SC_����) = 2 Then
                    If .Cell(flexcpChecked, i, SC_����, i, SC_����) = flexUnchecked Then
                        blnSelALl = False
                    End If
                End If
            End If
        Next
        
        For i = Row To .FixedRows Step -1
            If .RowData(i) = 0 Then
                lngClassRow = i
                Exit For
            Else
                If .Cell(flexcpData, i, SC_����) = 2 Then
                    If .Cell(flexcpChecked, i, SC_����, i, SC_����) = flexUnchecked Then
                        blnSelALl = False
                    End If
                End If
            End If
        Next
        If Not .Cell(flexcpPicture, lngClassRow, SC_����) Is Nothing Then
            Set .Cell(flexcpPicture, lngClassRow, SC_����) = imgEdit.ListImages(IIf(blnSelALl, "AllCheck", "UnCheck")).Picture
            .Cell(flexcpData, lngClassRow, SC_����) = IIf(blnSelALl, 1, 0)
        End If
    End With
End Sub

Private Sub vsCheckResult_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow > -1 And NewCol > -1 Then
        If vsCheckResult.RowData(NewRow) = 0 Then
            vsCheckResult.BackColorSel = &H8000000F
        Else
            vsCheckResult.BackColorSel = &H8000000D
        End If
    End If
End Sub

Private Sub vsCheckResult_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col = SC_����
End Sub

Private Sub vsCheckResult_Click()
    Dim blnSelALl       As Boolean
    Dim i               As Long
    With vsCheckResult
        If .MouseRow > 0 And .MouseCol = SC_���� Then
            If .RowData(.Row) = 0 Then
                blnSelALl = .Cell(flexcpData, .MouseRow, SC_����) = 0
                Call RecUpdate(mrsCheckInfo, "�������=" & .Cell(flexcpData, .MouseRow, SC_������) & " And �Ƿ�������=0", "�Ƿ����", IIf(blnSelALl, 1, 0))
                For i = .Row + 1 To .Rows - 1
                    If .RowData(i) = 0 Then
                        Exit For
                    Else
                        If Not .RowHidden(i) Then
                            If .Cell(flexcpData, i, SC_����) = 2 Then
                                .Cell(flexcpChecked, i, SC_����, i, SC_����) = IIf(blnSelALl, flexChecked, flexUnchecked)
                            End If
                        End If
                    End If
                Next
                If Not .Cell(flexcpPicture, .MouseRow, SC_����) Is Nothing Then
                    Set .Cell(flexcpPicture, .MouseRow, SC_����) = imgEdit.ListImages(IIf(blnSelALl, "AllCheck", "UnCheck")).Picture
                    .Cell(flexcpData, .MouseRow, SC_����) = IIf(blnSelALl, 1, 0)
                End If
            End If
        ElseIf .MouseRow = 0 And .Col = SC_���� Then
            If Not .Cell(flexcpPicture, 0, SC_����) Is Nothing Then
                blnSelALl = Val(.ColData(SC_����)) = 0
                Call RecUpdate(mrsCheckInfo, "", "�Ƿ����", IIf(blnSelALl, 1, 0))
                For i = .FixedRows To .Rows - 1
                    If .RowData(i) = 0 Then
                        If Not .Cell(flexcpPicture, i, SC_����) Is Nothing Then
                            Set .Cell(flexcpPicture, i, SC_����) = imgEdit.ListImages(IIf(blnSelALl, "AllCheck", "UnCheck")).Picture
                            .Cell(flexcpData, i, SC_����) = IIf(blnSelALl, 1, 0)
                        End If
                    Else
                        If .Cell(flexcpData, i, SC_����) = 2 Then
                            .Cell(flexcpChecked, i, SC_����, i, SC_����) = IIf(blnSelALl, flexChecked, flexUnchecked)
                        End If
                    End If
                Next
                .Cell(flexcpPicture, 0, SC_����) = imgEdit.ListImages(IIf(blnSelALl, "AllCheck", "UnCheck")).Picture
                .ColData(SC_����) = IIf(blnSelALl, 1, 0)
            End If
        End If
    End With
End Sub

Private Sub vsCheckResult_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> SC_���� Then
        Cancel = True
    Else
        If vsCheckResult.RowData(Row) = 0 Then
            Cancel = True
        End If
    End If
End Sub
