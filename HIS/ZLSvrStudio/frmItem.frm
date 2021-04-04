VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmItem 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   5700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.TaskPanel tplFunc 
      Height          =   4770
      Left            =   30
      TabIndex        =   0
      Top             =   315
      Width           =   3210
      _Version        =   589884
      _ExtentX        =   5662
      _ExtentY        =   8414
      _StockProps     =   64
      Behaviour       =   1
      ItemLayout      =   2
      HotTrackStyle   =   3
   End
   Begin XtremeSuiteControls.ShortcutCaption stcItem 
      Height          =   300
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   3210
      _Version        =   589884
      _ExtentX        =   5662
      _ExtentY        =   529
      _StockProps     =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
End
Attribute VB_Name = "frmItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==============================================================
'==��������
'==============================================================
Public gstrParentNo     As String
Public gstrParentCap    As String
Private mstrDefNo       As String 'ȱʡ����

'==============================================================
'==�����ӿ�
'==============================================================
Public Function RunByModule(Optional ByVal strModule As String) As Boolean
'���ܣ�ִ��ĳ��ģ�鹦�ܣ�������ʱ��ִ��Ĭ��ģ�鹦��
    Dim frmChild As Form
    Dim GroupTmp As TaskPanelGroup
    Dim itemTmp As TaskPanelGroupItem
    Dim strTmp As String
    
    If strModule = "" Then
        strModule = mstrDefNo
    ElseIf strModule = frmMDIMain.gstrLastModule Then
        Exit Function
    End If
    
    For Each frmChild In Forms
        If Not gfrmActive Is Nothing Then
            If gfrmActive.name <> frmChild.name And InStr(",frmMDIMain,frmPubIcons,frmItem,", "," & frmChild.name & ",") <= 0 Then
                Unload frmChild
            End If
        End If
    Next
    tplFunc.Tag = "ֱ������"
    For Each GroupTmp In tplFunc.Groups
        For Each itemTmp In GroupTmp.Items
            If Val(strModule) = itemTmp.Id Then
                itemTmp.Selected = True
            Else
                itemTmp.Selected = False
            End If
        Next
    Next
    tplFunc.Tag = ""
    frmMDIMain.gstrLastModule = strModule
    If Not gfrmActive Is Nothing Then
        Unload gfrmActive
        Set gfrmActive = Nothing
    End If
    
    'DBA����ֻ��DBA����ʹ��,�������ж�
    If strModule Like "06*" Then
            If Not gblnDBA Then
                frmMDIMain.stbThis.Panels(2).Text = "��ǰ�û�����DBA�û���Ȩ�޲��㣬�޷�ʹ�øù��ܡ�"
                Exit Function
            End If
    End If
    
    Select Case strModule
        Case "01", "02", "03", "04", "05" 'װж����
            frmDescribe.mstr��� = strModule
            Set gfrmActive = frmDescribe
        Case "0101" 'ϵͳװж����
            Set gfrmActive = frmAppStart
        Case "0102" 'ϵͳ��Ǩ
            If Not CheckAndAdjustMustTable("ZLRegInfo", , True) Then
                Exit Function
            End If
            If Not CheckAndAdjustMustTable("zlUpgradeConfig", , True) Then
                Exit Function
            End If
            Set gfrmActive = frmAppUpgrade
        Case "0103" '�������޸�
            Set gfrmActive = frmAppCheck
        Case "0105" '������Ч����
            Set gfrmActive = frmCompileInvalid
        Case "0104" '�û���װ�ű�
            Set gfrmActive = frmAppScript
        Case "0201" '���˺�:��ʷ���ݿռ����
            Set gfrmActive = frmHistoryDataMgr 'frmDataMove
        Case "0202" '���ݵ���
            Set gfrmActive = frmExp
        Case "0203" '���ݵ���
            Set gfrmActive = frmImp
        Case "0204" '���ݵ���
            Set gfrmActive = frmLoadOut
        Case "0205" '���ݵ���
            Set gfrmActive = frmLoadIn
        Case "0206" '�������
            If MsgBox("�ó���������Ҫ�ȴ�һ��ʱ�䣬�Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
            Set gfrmActive = frmClearData
        Case "0207" '��������
            Set gfrmActive = frmConnManagerParent
        Case "0208" '����ͼƬ����ת��
            If Not gblnSystemUser Then
                frmMDIMain.stbThis.Panels(2).Text = "��ǰ�û����Ǳ�׼��ϵͳ�����ߣ��޷�ʹ�øù���"
                Exit Function
            End If
            Set gfrmActive = frmLisPic2Ftp
        Case "0301" '�û�ע�����
            Set gfrmActive = frmRegist
        Case "0302" '����״̬���
            Set gfrmActive = frmStatus
        Case "0303" '��̨��ҵ����
            Set gfrmActive = frmAutoJobs
        Case "0304" '������־����
            Set gfrmActive = FrmRunLog
        Case "0305" '������־����
            Set gfrmActive = FrmErrLog
        Case "0306" 'ϵͳ����ѡ��
            Set gfrmActive = FrmRunOption
        Case "0307" '�ͻ�����������
            Set gfrmActive = frmClientUpgradeManage
        Case "0308" 'վ�����п���
            Set gfrmActive = frmClientsParas
        Case "0310" 'ϵͳ��������
            Set gfrmActive = frmParameters
        Case "0312" 'ҽԺ��Ϣά��
            Set gfrmActive = frmUnitInfoEdit
        Case "0314" '��Ҫ�����䶯��־����
            Set gfrmActive = frmAuditLogManage
        Case "0315" '������ʱ����
            Set gfrmActive = frmRunLimitManage
        Case "0401" '��ɫ��Ȩ����
            Set gfrmActive = frmRole
        Case "0402" '�û���Ȩ����
            Set gfrmActive = frmUser
        Case "0403" '�˵�����滮
            Set gfrmActive = frmMenu
        Case "0404" '��������Ȩ
            Set gfrmActive = frmMgrGrant
        Case "0501", "0502", "0505" '�������
            frmRptMan.mstr��� = strModule
            Set gfrmActive = frmRptMan
        Case "0503"
            Set gfrmActive = frmInputTools
        Case "0504"
            Set gfrmActive = frmNoticeTools
        Case "0601", "0602", "0606", "0604", "0605" 'DBA������
            Set gfrmActive = frmDbatoolsParent
            strTmp = strModule
        Case "0603"  'SQL���ٹ���
            Set gfrmActive = frmSQLTrace
        Case "0607"     '�û���IP����
            Set gfrmActive = frmUserLimit
        Case "0608"     'Ӧ�ó�����Ȩ
            Set gfrmActive = frmAppLimit
        Case "0609"     '�û���¼��־
            Set gfrmActive = frmLoginLog
        Case "0610"     '������ƹ���
            Set gfrmActive = frmFga
        Case ""
    End Select
    If Not gfrmActive Is Nothing Then
        frmMDIMain.stbThis.Panels(2).Text = ""
        Call FindWindowAndSetActive(gfrmActive)
        
        If gfrmActive.name = "frmDbatoolsParent" Then
            gfrmActive.ShowToolsForm strTmp
        Else
            gfrmActive.Show
        End If
        gfrmActive.ZOrder 0
    End If
    RunByModule = True
End Function

'==============================================================
'=�ؼ��¼�
'==============================================================
Private Sub Form_Load()
    Dim tpGroup As TaskPanelGroup
    On Error GoTo errH
    
    With frmMDIMain.grsToolsMenu
        If Not frmMDIMain.grsToolsMenu Is Nothing Then
            .Filter = "�ϼ�=" & IIf(gstrParentNo = "", "NULL", "'" & gstrParentNo & "'")
            If .Sort = "" Then
                If CheckAndAdjustMustTable("Zlsvrtools", "����", False) Then
                    .Sort = "����,���"
                Else
                    .Sort = "���"
                End If
            End If
            Set tpGroup = tplFunc.Groups.Add(Val(gstrParentNo), gstrParentCap)
            Do While Not .EOF
                If mstrDefNo = "" Then mstrDefNo = !��� & ""
                If !��� & "" <> "0404" Or gblnSystemUser Then
                    tpGroup.Items.Add(Val(!���), !����, xtpTaskItemTypeLink, Val(!���) + 1).Selected = False
                End If
                .MoveNext
            Loop
            
            tplFunc.SetMargins 1, 2, 0, 2, 2
            tplFunc.SelectItemOnFocus = True
            Call tplFunc.Icons.AddIcons(frmMDIMain.GetIcons.Icons)
            tplFunc.SetIconSize 24, 24
            tpGroup.CaptionVisible = False
            tpGroup.Expanded = True
            stcItem.Caption = gstrParentCap
        End If
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, Me.Caption
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    stcItem.Left = Me.Left
    stcItem.Width = Me.Width
    
    tplFunc.Height = Me.Height - Me.stcItem.Height
    tplFunc.Width = Me.Width
    tplFunc.Left = Me.Left
    tplFunc.Top = Me.stcItem.Top + Me.stcItem.Height
End Sub

Private Sub tplFunc_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    If tplFunc.Tag = "" Then
        If frmMDIMain.gstrLastModule <> "0" & Item.Id Then
            Call RunByModule("0" & Item.Id)
        End If
    End If
End Sub

'==============================================================
'=˽�з���
'==============================================================
Private Sub FindWindowAndSetActive(ByVal FrmObj As Form)
    Dim LngTargetHdl As Long
    '--����ô����Ѿ���,�򼤻���(����,����Ĵ�С���ᷢ���仯)--zyb
    LngTargetHdl = FindWindow(vbNullString, FrmObj.Caption)
    If LngTargetHdl <> 0 Then
        If IsIconic(LngTargetHdl) Then
            Call ShowWindow(LngTargetHdl, 9)            '��ԭָ������Ϊԭ��С
        End If
        Call SetActiveWindow(LngTargetHdl)
    End If
End Sub



