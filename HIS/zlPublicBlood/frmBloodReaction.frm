VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmBloodReaction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��Ѫ��Ӧ"
   ClientHeight    =   10545
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11400
   Icon            =   "frmBloodReaction.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10545
   ScaleWidth      =   11400
   StartUpPosition =   2  '��Ļ����
   Begin zlPublicBlood.usrCardEdit UCE 
      Height          =   10725
      Left            =   -30
      TabIndex        =   0
      Top             =   345
      Width           =   10980
      _extentx        =   19368
      _extenty        =   18918
      tabsposition    =   1
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   120
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmBloodReaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng�׶� As Long      '1��ҽ������׶�  2����Ѫ�ƴ���׶�
Private mlng����ID As Long
Private mlng��ҳid As Long
Private mlng������Դ As Long  '1-����  2-סԺ
Private mstrPrivs As String   'Ȩ�޴�
Private mfrmMain As Object    '������
Private mlngSys As Long
Private mlngģ��� As Long
Public mblnBloodReactionIsOpen As Boolean '��ģ̬״̬�£��жϴ����Ƿ���

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    '------------------------------------------------------------------------------------------------------------------
    '��ʼ������
    
    Call CommandBarInit(cbsMain)
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    '------------------------------------------------------------------------------------------------------------------
    '����������:������������
    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����"): objControl.BeginGroup = True 'objControl.BeginGroup = True���ǻ�����
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ��")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "�ύ"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "����")
        Set objControl = .Add(xtpControlButton, conMenu_View_Detail, "��Ѫִ��"): objControl.ToolTipText = "��Ѫִ���������": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each objControl In objBar.Controls
        If objControl.Type = xtpControlButton Then objControl.Style = xtpButtonIconAndCaption
    Next
    
    '------------------------------------------------------------------------------------------------------------------
    '����Ŀ����:���������������Ѵ���
    With cbsMain.KeyBindings '
        .Add 0, vbKeyDelete, conMenu_Edit_Delete            'ɾ��
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem         '����
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify          '�޸�
        .Add FCONTROL, vbKeyP, conMenu_File_Print           '��ӡ
        .Add FCONTROL, vbKeyS, conMenu_Edit_Save            '����
        .Add FCONTROL, vbKeyC, conMenu_Edit_Transf_Cancle   'ȡ��
        .Add FCONTROL, vbKeyX, conMenu_File_Exit            '�˳�
        .Add 0, vbKeyF1, conMenu_Help_Help                  '����
    End With
    
    InitCommandBar = True
    Exit Function
ErrHand:
    
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
        Case conMenu_File_Preview '��Ԥ��
            Call UCE.ShowPrint(2)
        Case conMenu_File_Print '��ӡ
            Call UCE.ShowPrint(1)
        Case conMenu_Edit_NewItem: '����
            UCE.AddPage
        Case conMenu_Edit_Modify: '�޸�
            UCE.ShowModify
        Case conMenu_Edit_Delete: 'ɾ��
            If IsPrivs(mstrPrivs, "ɾ������") = False Then
                If UCE.Doctor <> "" And UCE.Doctor <> UserInfo.���� Then
                    MsgBox "��û��Ȩ��ɾ�����˼�¼�����ݣ�", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            UCE.ShowDelete
        Case conMenu_Edit_Save: '����
            UCE.ShowSave
        Case conMenu_Edit_Transf_Cancle: 'ȡ��
            UCE.ShowCancel
        Case conMenu_Edit_Audit: '�ύ
            UCE.SubmitData
        Case conMenu_Edit_Untread: '����
            UCE.ShowUntread
        Case conMenu_View_Detail 'ִ������鿴
            Call frmBloodExecEdit.ViewExecution(Me, UCE.BloodID)
        Case conMenu_Help_Help
            Call gobjComlib.ShowHelp(App.ProductName, Me.hWnd, Me.name, Int((2200) / 100))
        Case conMenu_File_Exit
            Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long
    Dim lngTop  As Long
    Dim lngRight  As Long
    Dim lngBottom  As Long
    On Error GoTo Errorhand
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    '���������ؼ�Resize����
    UCE.Move lngLeft, lngTop + 50, lngRight - lngLeft, lngBottom - lngTop
Errorhand:
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
        Case conMenu_File_Preview, conMenu_File_Print
            Control.Visible = IsPrivs(mstrPrivs, "���ݴ�ӡ")
        Case conMenu_Edit_Modify: '�޸�
            '�޼�¼��Ӧ��Ȩ�����޸İ�ť���ɼ���
            Control.Visible = IsPrivs(mstrPrivs, "��¼��Ӧ")
            'ҽ���׶����ύ״̬ ���� ��Ѫ�����ύ״̬ ���� ҽ���׶���Ѫ������������ ���� ����״̬ ���� �޸�״̬ ���޸İ�ť��ʹ�ܣ��������ʹ�ܡ�
            Control.Enabled = Not ((mlng�׶� <> 2 And UCE.lng״̬ <> 0) Or (mlng�׶� = 2 And UCE.lng״̬ = 2) Or (mlng�׶� <> 2 And UCE.��Ѫ������ = True) Or UCE.strST = ���� Or UCE.strST = �޸�)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Edit_Delete: 'ɾ��
            '��ɾ����¼��Ȩ�޻�������Ѫ�ƽ׶Σ���ɾ����ť���ɼ���
            Control.Visible = IsPrivs(mstrPrivs, "ɾ����¼")   'and not (mlng�׶�=2 and not IsPrivs(mstrPrivs, "��Ѫ������"))������������Ѫ���������Ȩ�޵����������ɾ��
            'ҽ���׶����ύ״̬ ���� ��Ѫ�ƽ׶����ύ״̬ ���� ҽ���׶���Ѫ������������ ���� ����״̬ ���� �޸�״̬ �����ɾ����ť��ʹ�ܣ��������ʹ�ܡ�
            Control.Enabled = Not ((mlng�׶� <> 2 And UCE.lng״̬ <> 0) Or (mlng�׶� = 2 And UCE.lng״̬ <> 0) Or (mlng�׶� <> 2 And UCE.��Ѫ������ = True) Or UCE.strST = ���� Or UCE.strST = �޸�)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Edit_NewItem: '����
            '��Ѫ��û������Ȩ��ʱ��������ť���ɼ����������������ť�ɼ���
            Control.Visible = IsPrivs(mstrPrivs, "��¼��Ӧ")
            
            '���� 2017��6��22��  blnAddPage=trueҲ���ǵ�ǰ���ݿ����޸ĵ�����£�������ť��ʹ�ܣ��������ʹ��
            Control.Enabled = Not (UCE.blnAddPage = True)
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Edit_Save: '����
            Control.Visible = IsPrivs(mstrPrivs, "��¼��Ӧ")
            'δ�ύ�����ݱ仯ʱ������ʹ��
            Control.Enabled = UCE.DataChanged And UCE.lng״̬ = 0
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Edit_Transf_Cancle: 'ȡ��
            Control.Visible = IsPrivs(mstrPrivs, "��¼��Ӧ")
            'δ�ύ�����ݱ仯ʱ��ȡ��ʹ��
            Control.Enabled = UCE.DataChanged And UCE.lng״̬ = 0

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Edit_Audit: '�ύ
            Control.Visible = IsPrivs(mstrPrivs, "�ύ����")
            'ҽ���׶����ύ���� ���� ��Ѫ�ƽ׶����ύ���� ���� ҽ���׶���Ѫ������״̬ ���� ���������޸�״̬ ���� �޲��˻���δѡ�в���ʱ�ύ��ʹ�ܣ�����״̬�ύʹ�ܡ�
            Control.Enabled = Not ((mlng�׶� <> 2 And UCE.lng״̬ <> 0) Or (mlng�׶� = 2 And UCE.lng״̬ = 2) Or (mlng�׶� <> 2 And UCE.��Ѫ������ = True) Or UCE.strST = ���� Or UCE.strST = �޸�)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Edit_Untread: '����
            Control.Visible = IsPrivs(mstrPrivs, "�ύ����")
            'ҽ���׶η�ҽ�����ύ״̬ ���� ��Ѫ�ƽ׶�δ�ύ״̬ ���� ҽ���׶���Ѫ���������� ���� ��Ѫ�ƽ׶���Ѫ����������δ�ύ״̬ ���� ���������޸�״̬ ���� �޲��˻���δѡ�в��� ʱ���˲�ʹ�ܣ�����״̬����ʹ�ܡ�
            Control.Enabled = Not ((mlng�׶� <> 2 And UCE.lng״̬ <> 1) Or (mlng�׶� = 2 And UCE.lng״̬ <> 2) Or (mlng�׶� <> 2 And UCE.��Ѫ������ = True) Or (mlng�׶� = 2 And UCE.lng״̬ = 0 And UCE.��Ѫ������) Or UCE.strST = ���� Or UCE.strST = �޸�)
            
    Case conMenu_View_Detail
        Control.Enabled = UCE.BloodID > 0
    End Select
End Sub

Public Sub BloodReaction(frmMain As Object, lng�׶� As Long, lng����ID As Long, lng��ҳid As Long, lng������Դ As Long, ByVal lngSys As Long, _
                    lngģ��� As Long, Optional strPrivs As String, Optional lngisModul As Long = 0, Optional ByVal lng�շ�ID As Long = 0)
    '���ܣ���Ѫ��Ӧ��Ҫ�Ĵ������
    '������lng�׶�-ҽ������׶λ�����Ѫ�ƴ���׶Σ�lng����id-���˵�id��lng��ҳid-���˵���ҳid��lng������Դ-1�����2��סԺ��strPrivs-Ȩ�޴���lng�շ�id-���������շ�id��������ҳ��
    Set mfrmMain = frmMain
    If mblnBloodReactionIsOpen = True Then GoTo TOSHOW
    If zlGetComLib = False Then MsgBox "��ȡ����ʧ�ܣ�", vbInformation, gstrSysName: Exit Sub
    InitCommandBar
    mlng�׶� = lng�׶� '1��ҽ������׶�  2����Ѫ�ƴ���׶�
    mlng����ID = lng����ID
    mlng��ҳid = lng��ҳid
    mlng������Դ = lng������Դ '1������  2��סԺ
    mstrPrivs = strPrivs
    mlngSys = lngSys
    mlngģ��� = lngģ���
    If zlGetComLib = False Then MsgBox "��ȡ����ʧ�ܣ�", vbInformation, gstrSysName: Exit Sub
    UCE.InitEdit
    UCE.showInfor mlng����ID, mlng������Դ, mlng��ҳid, mlng�׶�, gcnOracle, Me, mlngģ���, , , lng�շ�ID
TOSHOW:
    mblnBloodReactionIsOpen = True
    Me.Show lngisModul, mfrmMain '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If (UCE.strST = ���� And UCE.DataChanged = True) Or UCE.strST = �޸� Then
        Cancel = (MsgBox("����δ���棬�Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
    End If
    mblnBloodReactionIsOpen = False
End Sub
