VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "���ݱ䶯֪ͨ����"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9630
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   9630
   StartUpPosition =   1  '����������
   Begin VB.Timer TimerState 
      Interval        =   1000
      Left            =   4320
      Top             =   600
   End
   Begin VB.Timer tmrDcn 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3720
      Top             =   600
   End
   Begin XtremeSuiteControls.TabControl tabMain 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   2655
      _Version        =   589884
      _ExtentX        =   4683
      _ExtentY        =   1720
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   6090
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2223
            MinWidth        =   882
            Picture         =   "frmMain.frx":4D4A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10716
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "2019/10/17"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1323
            MinWidth        =   1323
            TextSave        =   "13:14"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock winSock 
      Left            =   5040
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   1320
      Top             =   840
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgMain 
      Bindings        =   "frmMain.frx":5165
      Left            =   1800
      Top             =   840
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMain.frx":5179
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum CommandBarIDCond
    conMenu_Start = 1
    conMenu_Stop
    conMenu_Log_Clear
    conMenu_Clear
    conMenu_Setting
    conMenu_Exit
End Enum

Private Enum ChangeType
    DcnInsert = 1
    DcnUpdate = 2
    DcnDelete = 3
End Enum

Private mblnStartUp As Boolean

Private mblnDcn As Boolean  '�Ƿ�����DCN
Private mblnOciConnected As Boolean
Private mblnCancel As Boolean

Private mrsNoticeSet As ADODB.Recordset '����ע��DCN��Notice��Ϣ
Private mfrmNoticeList As New frmNoticeList
Private mfrmNoticeLog As New frmNoticeLog

Private mlngCheckInterval As Long   'DCN���ʱ����¼��
Private mlngCheck As Long

Private Sub cbsMain_Resize()
    Dim lngTop As Long, lngBottom As Long
    Dim lngLeft As Long, lngRight As Long
    On Error Resume Next
    
    cbsMain.GetClientRect lngLeft, lngTop, lngRight, lngBottom
    With tabMain
        .Left = 0
        .Top = lngTop + 10
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top - stbThis.Height
    End With
End Sub

Private Sub InitCommandBar()
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    
    With cbsMain.Options
        .ShowFullAfterDelay = True
        .ShowTextBelowIcons = True
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
    End With
    Set cbsMain.Icons = imgMain.Icons
    cbsMain.ActiveMenuBar.Visible = False
    
    '����������
    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Start, "����")
        objControl.Style = xtpButtonIconAndCaption: objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Stop, "ֹͣ")
        objControl.Style = xtpButtonIconAndCaption: objControl.BeginGroup = False
    
        Set objControl = .Add(xtpControlButton, conMenu_Log_Clear, "���������־")
        objControl.Style = xtpButtonIconAndCaption: objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Clear, "������־����")
        objControl.Style = xtpButtonIconAndCaption: objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Setting, "����")
        objControl.Style = xtpButtonIconAndCaption: objControl.BeginGroup = False
        
        Set objControl = .Add(xtpControlButton, conMenu_Exit, "�˳�")
        objControl.Style = xtpButtonIconAndCaption: objControl.BeginGroup = True
    End With
    
End Sub

Private Sub InitTab()
    '����:��ʼ��tab�ؼ�
    Dim objItem As TabControlItem
    
    With tabMain.PaintManager
        .Appearance = xtpTabAppearancePropertyPage2003
        .Color = xtpTabColorOffice2003
    End With
    
    Set objItem = tabMain.InsertItem(1, "������־", mfrmNoticeLog.hwnd, 0)
    tabMain.InsertItem 2, "���ݱ䶯֪ͨ�б�", mfrmNoticeList.hwnd, 0
    
    objItem.Selected = True
End Sub

Private Sub InitNoticeSet()
    Set mrsNoticeSet = GetNoticeList
    If mrsNoticeSet Is Nothing Then Exit Sub
    
    mfrmNoticeList.SetDataSource mrsNoticeSet
End Sub

Private Sub Form_Activate()
    If mblnStartUp Then Exit Sub
    mblnStartUp = True
    
    
    DoEvents    '��ֹע��ʱ��ϳ�  ���濨��
    
    Call ChangeDcnState(1)
    Call UpdateCmdState
    
    Me.Caption = Me.Caption & " - [" & gstrServer & "]"
End Sub

Private Sub Form_Load()
    Dim rs As New ADODB.Recordset, strSql As String
    
    glngPort = 9999 'Ĭ�϶˿�9999
    gintLog = 1 'Ĭ�ϱ��汾����־
    gintInterval = 200  'Ĭ��ˢ��Ƶ�� 200ms

    '��ʽ:IP;�˿�;״̬;�Ự��
    strSql = "SELECT ����ֵ FROM zltools.zloptions WHERE ������=[1]"
    Set rs = zlDatabase.OpenSQLRecord(strSql, Me.Caption, 27)
    If rs.RecordCount <> 0 Then
        If rs!����ֵ & "" <> "" Then
            glngPort = Split(rs!����ֵ, ";")(1)
        End If
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(strSql, Me.Caption, 28)
    If rs.RecordCount <> 0 Then
        If rs!����ֵ & "" <> "" Then
            gintLog = Val(rs!����ֵ)
        End If
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(strSql, Me.Caption, 29)
    If rs.RecordCount <> 0 Then
        If rs!����ֵ & "" <> "" Then
            gintInterval = Val(rs!����ֵ)
        End If
    End If
    
    gstrIp = winSock.LocalIP
    glngSid = GetSid
    
    tmrDcn.Interval = gintInterval
    mlngCheckInterval = GetCheckInterval
    
    zlCommFun.SetWindowsInTaskBar Me.hwnd, True
    Call InitCommandBar
    Call InitTab
    
    Call InitNoticeSet
End Sub

Private Sub Form_Unload(Cancel As Integer)

    mblnCancel = False
    
    If MsgBox("���Ƿ����Ҫ�˳��Զ����ѷ���", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
        Cancel = True
        mblnCancel = True
        Exit Sub
    End If

    On Error Resume Next
    Unload mfrmNoticeList: Set mfrmNoticeList = Nothing
    Unload mfrmNoticeLog: Set mfrmNoticeLog = Nothing
    
    If mblnDcn Then
        Call OCI_UnRigister
        Call UpdateDcnState2DB(0)
    End If
    If gcnOracle.State <> adStateClosed Then gcnOracle.Close
End Sub

Private Sub ChangeDcnState(ByVal intType As Integer)
    '����:������ر�DCN
    'intType=1: ����  intType=0:�ر�
    Dim strTmp As String
    
    If intType = 0 And mblnDcn = False Then Exit Sub
    If intType = 1 And mblnDcn = True Then Exit Sub
    
    If UpdateDcnState2DB(intType) = False Then   '������Ϣ�շ�����״̬
        Exit Sub
    End If
    
    If intType = 0 Then
        strTmp = "���ڹر����ݱ䶯֪ͨ..."
        stbThis.Panels(2).Text = strTmp: mfrmNoticeLog.WriteLog strTmp, 1
        
        Call DcnStop
        tmrDcn.Enabled = False
        If winSock.State <> sckClosed Then
            winSock.Close
        End If
        
        strTmp = "���ݱ䶯֪ͨ�ѹرա�"
        stbThis.Panels(2).Text = strTmp: mfrmNoticeLog.WriteLog strTmp, 1
    Else
        strTmp = "���ڿ������ݱ䶯֪ͨע��..."
        stbThis.Panels(2).Text = strTmp: mfrmNoticeLog.WriteLog strTmp, 1
        mfrmNoticeLog.Refresh
        
        If winSock.State <> sckOpen Then
            winSock.LocalPort = glngPort
            winSock.Bind
        End If
        
        If DcnStart = False Then
            Call UpdateDcnState2DB(0)
            strTmp = "���ݿ�����ʧ�ܣ��޷��������ݱ䶯֪ͨ ��"
            stbThis.Panels(2).Text = strTmp: mfrmNoticeLog.WriteLog strTmp, 1
            Exit Sub
        End If
        Call DcnRigister
        Call UpdateDcnTime
        
        tmrDcn.Enabled = True
        strTmp = "���ݱ䶯֪ͨ�ѿ�����"
        stbThis.Panels(2).Text = strTmp: mfrmNoticeLog.WriteLog strTmp, 1
    End If
    
    If mblnDcn = True Then
        mblnDcn = False
    Else
        mblnDcn = True
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case conMenu_Start
            Call ChangeDcnState(1)
        Case conMenu_Stop
            Call ChangeDcnState(0)
        Case conMenu_Setting
            frmNoticeSet.ShowEdit
            Exit Sub
        Case conMenu_Clear
            frmLogClear.ShowClearLog
            Exit Sub
        Case conMenu_Exit
            Unload Me
            Exit Sub
        Case conMenu_Log_Clear
            Call mfrmNoticeLog.ClearLog
            Exit Sub
    End Select
    Call UpdateCmdState
End Sub

Private Sub DcnStop()

    If lpPrevWndProc <> 0 Then
        UnHook Me.hwnd
        lpPrevWndProc = 0
    End If
End Sub

Private Function DcnStart() As Boolean
    '��֤zlNoticeLib�Ƿ����,����¼���ݿ�
    Dim strServiceName As String, strIp As String, strPort As String
    Dim i  As Integer
    
    On Error GoTo errH
    If mblnOciConnected Then    '����Ѿ�������OCI����,�Ͳ���Ҫ������
        If lpPrevWndProc = 0 Then Hook Me.hwnd  '�󶨴��ں���
        DcnStart = True
        Exit Function
    End If
    
    GetServerInfo gstrServer, strServiceName, strIp, strPort
    
    mblnDcn = OCI_ConnCreate(strIp & ":" & strPort & "/" & strServiceName, gstrUserName, gstrUserPwd)    '��¼OCI

    '�����¼ʧ��,����ʹ��ע����е�IP��Ϣ
    If mblnDcn = False Then
        strIp = GetSetting("ZLSOFT\����ģ��", "zlSvrNotice", "IP", strIp)
        strPort = GetSetting("ZLSOFT\����ģ��", "zlSvrNotice", "PORT", strPort)
        strServiceName = GetSetting("ZLSOFT\����ģ��", "zlSvrNotice", "Server", strServiceName)
        
        If strPort <> "" Then
            mblnDcn = OCI_ConnCreate(strIp & ":" & strPort & "/" & strServiceName, gstrUserName, gstrUserPwd)    '��¼OCI
        End If
    End If
    
    '���ע�����û��Ip��Ϣ���ߵ�¼ʧ��,�͵���ȷ�Ͽ�
    Do While mblnDcn = False And i < 3
        mfrmNoticeLog.WriteLog "���ݱ䶯֪ͨ����������ݿ�ʧ�ܣ�����IP���˿ڡ�����������Ϣ�Ƿ���ȷ��ͬʱ���Sqlnet.ora�ļ��в����� NAMES.DIRECTORY_PATH ����", 1
         If frmUserCheckLogin.GetSerInfo(strIp, strPort, strServiceName) = True Then
            i = i + 1
            mblnDcn = OCI_ConnCreate(strIp & ":" & strPort & "/" & strServiceName, gstrUserName, gstrUserPwd)    '��¼OCI
         Else
            Exit Function
         End If
    Loop
    mblnOciConnected = mblnDcn
    DcnStart = mblnDcn
    Exit Function
errH:
    ErrCenter
End Function

Private Sub DcnRigister()
    'ע��DCN����
    Dim strSql As String, strTmp As String
    
    If mrsNoticeSet Is Nothing Then Exit Sub
    mrsNoticeSet.Filter = "Status =  1" '���˵���ͣ�õ�����
    If mrsNoticeSet.RecordCount = 0 Then Exit Sub
    
    mrsNoticeSet.Sort = "Tableowner,Tablename"
    
    Do While Not mrsNoticeSet.EOF
        If strTmp <> mrsNoticeSet!Tableowner & "." & mrsNoticeSet!Tablename Then
            strTmp = mrsNoticeSet!Tableowner & "." & mrsNoticeSet!Tablename
            strSql = "Select  * from " & strTmp
            OCI_Register Me.hwnd, strSql 'ע����Ϣ,������
        End If
        mrsNoticeSet.MoveNext
    Loop
    If lpPrevWndProc = 0 Then Hook Me.hwnd  '�󶨴��ں���
End Sub

Private Sub SendNotice()
    '����:ѭ��֪ͨ����, ������Ϣ
    Dim strOwner As String, strTable As String, strRowid As String
    Dim arrTmp()  As String, lngNoticeCode As Long, intChangeType As Integer
    Dim strSql As String, rsData As New ADODB.Recordset, strCols As String
     
    On Error GoTo errH
    With gcolNotice
        If .Count = 0 Then Exit Sub

        '�䶯��Ϣ��ʽΪ:   �䶯����-������.����-Rowid
        arrTmp = Split(.Item(1), "-")
        intChangeType = arrTmp(0)
        strRowid = arrTmp(2)
        
        arrTmp = Split(arrTmp(1), ".")
        strOwner = arrTmp(0)
        strTable = arrTmp(1)
        
        '��һ������䶯��������>80ʱ,���ص�RowIDΪ1
        If strRowid = "1" Then
            mfrmNoticeLog.WriteLog Now & "   ��" & strTable & "���ݱ䶯���������޷���ȡRowid��������֪ͨ", 1
            .Remove 1
            Exit Sub
        End If
        
        If intChangeType = ChangeType.DcnDelete Then '����䶯������ɾ��,���Ƴ������䶯��Ϣ
            .Remove 1
            Exit Sub
        End If
        
        '����Table��Owner���й���
        mrsNoticeSet.Filter = "Tableowner = '" & strOwner & "' And Tablename = '" & strTable & "' And Status = 1"
        
        If mrsNoticeSet.RecordCount = 0 Then
            .Remove 1
            Exit Sub
        End If
        
        'һ�����ݱ䶯,�����漰����֪ͨ�趨,ѭ��֪ͨ�趨,���η�������
        Do While Not mrsNoticeSet.EOF
        
            '��ȡ�䶯����
            If mrsNoticeSet!ReceiverTab & "" = "" Then
                strSql = "Select  " & IIf(mrsNoticeSet!ReceiverCols & "" = "", "1", mrsNoticeSet!ReceiverCols) & " From " & strOwner & "." & strTable & " A Where Rowid =  [1]  " & IIf(mrsNoticeSet!Filter & "" = "", "", " And " & mrsNoticeSet!Filter)
            Else
                If mrsNoticeSet!ReceiverCols & "" = "" Then
                    strCols = "1"
                Else
                    strCols = "B." & mrsNoticeSet!ReceiverCols
                    strCols = Replace(strCols, ",", ",B.")  '�滻����
                End If
            
                strSql = "Select " & strCols & " From " & strOwner & "." & strTable & " A , " & mrsNoticeSet!ReceiverTab & " B Where A.Rowid =[1] And " & mrsNoticeSet!ReceiverRelas & IIf(mrsNoticeSet!Filter & "" = "", "", " And " & mrsNoticeSet!Filter)
            End If
            
            Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ȡ�䶯����", strRowid)
            If rsData.RecordCount > 0 Then
                Post2Client rsData, mrsNoticeSet!Noticekind, intChangeType, strRowid, mrsNoticeSet!NoticeCode, strOwner, strTable, _
                                mrsNoticeSet!ReceiverCols & "", mrsNoticeSet!SplitChar & "", _
                                mrsNoticeSet!ReceiverIP & "", mrsNoticeSet!ReceiverStaffKind & "", mrsNoticeSet!ReceiverDeptKind & ""
            End If
            mrsNoticeSet.MoveNext
        Loop
        
        .Remove 1   'ִ����ɺ�,�Ƴ������䶯��Ϣ
    End With
    
    Exit Sub
errH:
    '��¼����
    gcolNotice.Remove 1  'ɾ����ǰ
    If 0 = 1 Then
        Resume
    End If
    mfrmNoticeLog.WriteLog "ת����" & strTable & "���ݱ䶯֪ͨʱ�������� " & Err.Description, 1
End Sub

Private Sub Post2Client(ByVal rsData As ADODB.Recordset, intNoticeKind As Integer, intChangeType As Integer, _
                                strRowid As String, lngNoticeCode As Long, strOwner As String, strTable As String, _
                                strReceiverCol As String, strSplitChar As String, _
                                strReceiverIp As String, strReceiverStaffKind As String, strReceiverDeptKind As String)
    '����:���䶯��Ϣ���͵��ͻ���
    Dim strSql As String, rsTmp As New ADODB.Recordset
    Dim strField1 As String, strField2 As String
    Dim strIp  As String, lngPort As Long
    Dim strTmp As String, arrTmp() As String, i As Integer
    
    '��ȡReceiverCol��ֵ
    If strReceiverCol <> "" Then
        If InStr(1, strReceiverCol, ",") > 0 Then
            strField1 = rsData.Fields(0)
            strField2 = rsData.Fields(1)
        Else
            strField1 = rsData.Fields(0)
        End If
    End If
    
    If strSplitChar <> "" Then  '����зָ���,�ͽ��ָ����滻Ϊ����,�Ա�ʹ��f_str2list����
        strField1 = Replace(strField1, strSplitChar, ",")
        strField2 = Replace(strField2, strSplitChar, ",")
    End If
    
    Select Case intNoticeKind
        Case 0 '֪ͨ���пͻ���
            strSql = "Select IP,��Ϣ�˿�,����վ From Zltools.Zlclientsession Where ״̬=1"
            
        Case 1 'ָ�����ţ�ֻ����������վ��ǰ����
            If strSplitChar <> "" Then
                strSql = "Select IP,��Ϣ�˿�,����վ From Zltools.Zlclientsession Where ״̬=1  And ��ǰ����ID " & _
                                " In (Select /*+ cardinality(a,10)*/ Column_Value From Table(f_Str2list([1])) A)"
            Else
                strSql = "Select IP,��Ϣ�˿�,����վ From Zltools.Zlclientsession Where ״̬=1  And ��ǰ����ID = [1]"
            End If
            
        Case 2 'ָ���û�����
            If strSplitChar <> "" Then
                strSql = "Select IP,��Ϣ�˿�,����վ From Zltools.Zlclientsession Where ״̬=1  And ��Ա���� " & _
                                " In (Select /*+ cardinality(a,10)*/ Column_Value From Table(f_Str2list([1])) A)"
            Else
                strSql = "Select IP,��Ϣ�˿�,����վ  From Zltools.Zlclientsession Where ״̬=1  And ��Ա����= [1]"
            End If
            
        Case 3 'ָ������+λ��
            If strField2 = "" Then   '������λ��Ϊ��,�������в��ŷ���
                strSql = "Select IP,��Ϣ�˿�,����վ From Zltools.Zlclientsession  Where ״̬=1  And ��ǰ����ID= [1]"
            Else
                If InStr(1, strField2, ",") > 0 Then
                    strTmp = ""
                    arrTmp = Split(strField2, ",")
                    For i = 0 To UBound(arrTmp)
                        strTmp = strTmp & IIf(strTmp = "", "", "Or") & " Instr(',' || ��ǰλ�� || ',' , '," & arrTmp(i) & ",' )>0 "
                    Next
                    strSql = "Select IP,��Ϣ�˿�,����վ From Zltools.Zlclientsession  Where ״̬=1  And ��ǰ����ID= [1] " & _
                                "And ( " & strTmp & " )"
                Else
                    strSql = "Select IP,��Ϣ�˿�,����վ From Zltools.Zlclientsession  Where ״̬=1  And ��ǰ����ID= [1] And instr(',' || ��ǰλ�� || ',' , '," & strField2 & ",' )>0 "
                End If
            End If
            
        Case 4  'ָ���û����û���
            If strSplitChar <> "" Then
                strSql = "Select IP,��Ϣ�˿�,����վ From Zltools.Zlclientsession Where ״̬=1  And �û��� " & _
                                " In (Select /*+ cardinality(a,10)*/ Column_Value From Table(f_Str2list([1])) A)"
            Else
                strSql = "Select IP,��Ϣ�˿�,����վ  From Zltools.Zlclientsession Where ״̬=1  And �û���= [1]"
            End If
            
        Case 5  'ָ��IP���˿�
            strSql = "Select IP,��Ϣ�˿�,����վ From Zltools.Zlclientsession  Where ״̬=1  And IP = [1] And ��Ϣ�˿� = [2] "
            strField1 = Split(strReceiverIp, ":")(0)    '����IP\�˿ڷ�����Ϣ,ֱ��ȡReceiverIP�е�ֵ
            strField2 = Split(strReceiverIp, ":")(1)
            
        Case 6  'ָ�����ʵĹ���վ
            strSql = "Select IP,��Ϣ�˿�,����վ From Zltools.Zlclientsession  Where ״̬=1 And ��Ա���� = [1]"
            strField1 = strReceiverStaffKind    '������Ա���ʷ�����Ϣ,ֱ��ȡReceiverStaffKind�е�ֵ
            
        Case 7  'ָ�����ţ�������鹤��վ����ȫ������
            If strSplitChar = "" Then
                strSql = "Select Ip, ��Ϣ�˿�, ����վ From Zltools.Zlclientsession Where ״̬ = 1 And ��ǰ����id = [1]" & vbNewLine & _
                            "Union" & vbNewLine & _
                            "Select Ip, ��Ϣ�˿�, ����վ From Zltools.Zlclientsession A, Zltools.Zlclientdepts B Where a.״̬ = 1 And a.�Ự�� = b.�Ự�� And b.����id = [1]"
            Else
                strSql = "Select Ip, ��Ϣ�˿�, ����վ" & vbNewLine & _
                            "From Zltools.Zlclientsession" & vbNewLine & _
                            "Where ״̬ = 1 And ��ǰ����id In (Select /*+ cardinality(a,10)*/ Column_Value From Table(f_Str2list([1])) A)" & vbNewLine & _
                            "Union" & vbNewLine & _
                            "Select Ip, ��Ϣ�˿�, ����վ" & vbNewLine & _
                            "From Zltools.Zlclientsession A, Zltools.Zlclientdepts B" & vbNewLine & _
                            "Where a.״̬ = 1 And a.�Ự�� = b.�Ự�� And b.����id In (Select /*+ cardinality(a,10)*/ Column_Value From Table(f_Str2list([1])) A)"
            End If
            
        Case 8  'ָ����������
            If strField1 = "" Then   '������λ��Ϊ��,��������ͬ���ʲ��ŷ���
                strSql = "Select IP,��Ϣ�˿�,����վ From Zltools.Zlclientsession  Where ״̬=1  And �������� = [2]"
            Else
                If InStr(1, strField1, ",") > 0 Then    'һ�����ݷ��͸����λ��
                    strTmp = ""
                    arrTmp = Split(strField1, ",")
                    For i = 0 To UBound(arrTmp)
                        strTmp = strTmp & IIf(strTmp = "", "", "Or") & " Instr(',' || ��ǰλ�� || ',' , '," & arrTmp(i) & ",' )>0 "
                    Next
                    strSql = "Select IP,��Ϣ�˿�,����վ From Zltools.Zlclientsession  Where ״̬=1  And �������� = [2] " & _
                                "And ( " & strTmp & " )"
                Else
                    strSql = "Select IP,��Ϣ�˿�,����վ From Zltools.Zlclientsession  Where ״̬=1  And �������� = [2]  And instr(',' || ��ǰλ�� || ',' , '," & strField1 & ",' )>0 "  'һ̨����վͬʱ���ö��λ��
                End If
            End If
            strField2 = strReceiverDeptKind '���ղ������ʷ�����Ϣ,ֱ��ȥReceiverDeptKind�е�ֵ
    End Select
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�����͹���վ", strField1, strField2)
    If rsTmp.RecordCount <> 0 Then
        
        '��¼��־����
        gstrBuild.Clear
        gstrBuild.Append Now
        gstrBuild.Append "   ��": gstrBuild.Append strOwner
        gstrBuild.Append ".": gstrBuild.Append strTable
        gstrBuild.Append "�����ݱ䶯����: ": gstrBuild.Append Decode(intChangeType, ChangeType.DcnInsert, "����", ChangeType.DcnUpdate, "�޸�", "����")
        mfrmNoticeLog.WriteLog gstrBuild.ToString, 2
        
        gstrBuild.Clear
        Do While Not rsTmp.EOF
            winSock.RemoteHost = rsTmp!IP
            winSock.RemotePort = rsTmp!��Ϣ�˿�
            winSock.SendData lngNoticeCode & "-" & intChangeType & "-" & strOwner & "-" & strTable & "-" & strRowid
            
            'ƴ���ַ���,��¼������Ϣ
            If gstrBuild.Length <> 0 Then gstrBuild.Append vbNewLine
            gstrBuild.Append "   ����վ��"
            gstrBuild.Append rsTmp!����վ: gstrBuild.Append "("
            gstrBuild.Append rsTmp!IP: gstrBuild.Append ":"
            gstrBuild.Append rsTmp!��Ϣ�˿�: gstrBuild.Append "���� �ѷ���"
            
            rsTmp.MoveNext
        Loop
        If gstrBuild.Length > 0 Then mfrmNoticeLog.WriteLog gstrBuild.ToString, 2
    End If
    
End Sub

Private Sub UpdateCmdState()
    '���ð�ť�Ŀ�����
    Dim objControl As CommandBarControl
    
    '����
    Set objControl = cbsMain.FindControl(, conMenu_Start)
    objControl.Enabled = Not mblnDcn
    'ֹͣ
    Set objControl = cbsMain.FindControl(, conMenu_Stop)
    objControl.Enabled = mblnDcn
    
End Sub

Private Sub TimerState_Timer()
        
    TimerState.Enabled = False
    
    mlngCheck = mlngCheck + 1
        
    If mlngCheck > mlngCheckInterval Then
        Call UpdateDcnTime
        mlngCheck = 0
    End If
    
    TimerState.Enabled = True
End Sub

Private Sub tmrDcn_Timer()
    tmrDcn.Enabled = False
    Call SendNotice
    tmrDcn.Enabled = True
End Sub
