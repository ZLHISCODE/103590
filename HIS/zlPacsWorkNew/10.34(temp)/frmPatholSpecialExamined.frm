VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPatholSpecialExamined 
   Caption         =   "������"
   ClientHeight    =   9525
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   10875
   Icon            =   "frmPatholSpecialExamined.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9525
   ScaleWidth      =   10875
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.ImageList imgTbrS 
      Left            =   9195
      Top             =   675
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecialExamined.frx":000C
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecialExamined.frx":0C7E
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecialExamined.frx":18F0
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecialExamined.frx":2562
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecialExamined.frx":31D4
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecialExamined.frx":3E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecialExamined.frx":4AB8
            Key             =   "IMG7"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   1270
      ButtonWidth     =   1455
      ButtonHeight    =   1270
      Style           =   1
      ImageList       =   "imgTbrS"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��ǩ��ӡ"
            Key             =   "tbLAB"
            ImageIndex      =   4
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbLabPreview"
                  Text            =   "Ԥ��"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tblabPrint"
                  Text            =   "��ӡ"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�嵥��ӡ"
            Key             =   "tbList"
            ImageIndex      =   5
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbListPreview"
                  Text            =   "Ԥ��"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbListPrint"
                  Text            =   "��ӡ"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "-"
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����鿴"
            Key             =   "tbViewRequest"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�ؼ����"
            Key             =   "tbAcceptSpeExam"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�ؼ����"
            Key             =   "tbEndSpeExam"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Frame framSpeExam 
      Caption         =   "�ؼ��¼"
      Height          =   7215
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   9975
      Begin VB.Frame FramCheck 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5760
         TabIndex        =   6
         Top             =   6840
         Width           =   4095
         Begin VB.CheckBox chkYSQ 
            Caption         =   "������"
            Height          =   255
            Left            =   1080
            TabIndex        =   9
            Top             =   -7
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkYJS 
            Caption         =   "�ѽ���"
            Height          =   180
            Left            =   2160
            TabIndex        =   8
            Top             =   30
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkYWC 
            Caption         =   "�����"
            Height          =   180
            Left            =   3120
            TabIndex        =   7
            Top             =   30
            Width           =   855
         End
      End
      Begin VB.OptionButton optFenZi 
         Caption         =   "���Ӳ���"
         Height          =   255
         Left            =   3360
         TabIndex        =   2
         Tag             =   "2"
         Top             =   0
         Width           =   1095
      End
      Begin VB.OptionButton optTeShu 
         Caption         =   "����Ⱦɫ"
         Height          =   255
         Left            =   2160
         TabIndex        =   1
         Tag             =   "1"
         Top             =   0
         Width           =   1095
      End
      Begin VB.OptionButton optMianYi 
         Caption         =   "�����黯"
         Height          =   255
         Left            =   960
         TabIndex        =   0
         Tag             =   "0"
         Top             =   0
         Value           =   -1  'True
         Width           =   1095
      End
      Begin zl9PACSWork.ucFlexGrid ufgData 
         Height          =   6015
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   10610
         GridRows        =   21
         IsBtnNextCell   =   0   'False
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         HeadFontCharset =   134
         HeadFontWeight  =   400
         DataFontCharset =   134
         DataFontWeight  =   400
         ExtendLastCol   =   -1  'True
      End
      Begin VB.Label labRecordInf 
         Caption         =   "��ǰ����Ŀ����0    ��ǰ������Ŀ����0"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   6840
         Width           =   5295
      End
   End
End
Attribute VB_Name = "frmPatholSpecialExamined"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IWorkMenu

Private Const C_INT_MIANYI As Integer = 0
Private Const C_INT_TESHU As Integer = 1
Private Const C_INT_FENZI As Integer = 2

Private Const M_STR_HINT_NoSelectData As String = "��Ч�ļ�����ݣ���ѡ����Ҫִ�еļ���¼��"
Private Const M_STR_MODULE_MENU_TAG As String = "�ؼ�"

Private WithEvents zlReport As zl9Report.clsReport
Attribute zlReport.VB_VarHelpID = -1

Private mlngModule As Long
Private mstrPrivs As String              'ģ��Ȩ��
Private mlngCurDeptId As Long          '��ǰ����
Private mobjOwner As Object

Private mlngAdviceID As Long
Private mlngSendNo As Long
Private mblnMoved As Boolean
Private mlngStudyState As Long

Private mrecStudyInf As TStudyStateInf
Private mblnReadOnly As Boolean

Private mblnAutoAcceptOfAfterPrint As Boolean '��ӡ���Զ�����

Private mlngTmpAdviceId As Long
Private mlngTmpSendNo As Long


Private mObjActiveMenuBar As CommandBar
Private mbytFontSize As Byte '�ֺ�    9--С����    12--������

Private mblnRefreshState As Boolean

Private mKeyCode As Long
Private mKeyShift As Long



'��ȡ�˵��ӿڶ���
Property Get zlMenu() As IWorkMenu
    Set zlMenu = Me
End Property


Public Sub NotificationRefresh()
'֪ͨˢ��
    mblnRefreshState = False
End Sub


'�ӿ�ʵ�ֲ���*********************************************************************************

Public Function IWorkMenu_zlIsModuleMenu(ByVal objControlMenu As XtremeCommandBars.ICommandBarControl) As Boolean
'�жϲ˵��Ƿ����ڸ�ģ��˵�
    IWorkMenu_zlIsModuleMenu = IIf(objControlMenu.Category = M_STR_MODULE_MENU_TAG, True, False)
End Function


Public Sub IWorkMenu_zlCreateMenu(objMenuBar As Object)
'����Ӱ���¼��Ӧ�Ĳ˵�
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    Dim cbrControl As CommandBarControl
    
    Dim str3DFuncs() As String
    Dim i As Long
    Dim lng3DFunc As Long
    
    
    Set mObjActiveMenuBar = objMenuBar.ActiveMenuBar

    If Not HasMenu(objMenuBar, conMenu_PatholSpeExam) Then
        Set cbrMenuBar = mObjActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_PatholSpeExam, "�ؼ�(&T)", GetPatholMenuIndex(objMenuBar) + 1, False)
        cbrMenuBar.ID = conMenu_PatholSpeExam
        cbrMenuBar.Category = M_STR_MODULE_MENU_TAG
        
        
        With cbrMenuBar.CommandBar
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_PatholSpeExam_LAB, "��ǩ��ӡ(&B)", "", 1, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_PatholSpeExam_PreviewLAB, "Ԥ��(V)", "", 1, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_PatholSpeExam_PrintLab, "��ӡ(P)", "", 1, False)
                
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_PatholSpeExam_List, "�嵥��ӡ(&L)", "", 1, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_PatholSpeExam_PreviewList, "Ԥ��(V)", "", 1, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_PatholSpeExam_PrintList, "��ӡ(P)", "", 1, False)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholSpeExam_RequestView, "����鿴(&Q)", "", 1, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholSpeExam_Accept, "�ؼ����(&R)", "", 1, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholSpeExam_Finish, "�ؼ����(&F)", "", 1, False)
        End With
    End If
End Sub


Public Sub IWorkMenu_zlCreateToolBar(objToolBar As Object)
'����������
    Exit Sub
End Sub

Public Sub IWorkMenu_zlClearMenu()
'����������Ĳ˵�
    Exit Sub
End Sub


Public Sub IWorkMenu_zlClearToolBar()
'��������Ĺ�����
    Exit Sub
End Sub


Public Sub IWorkMenu_zlExecuteMenu(ByVal lngMenuId As Long)
'���ݲ˵�IDִ�ж�Ӧ����
    Dim objCbrControl As XtremeCommandBars.CommandBarControl
    
    Select Case lngMenuId
        Case conMenu_PatholSpeExam_PreviewLAB
            Call PrintSpeExamLabel(False)
            
        Case conMenu_PatholSpeExam_PrintLab
            Call PrintSpeExamLabel(True)
            
        Case conMenu_PatholSpeExam_PreviewList
            Call PrintWorkList(False)
            
        Case conMenu_PatholSpeExam_PrintList
            Call PrintWorkList(True)
            
        Case conMenu_PatholSpeExam_RequestView
            Call ShowSpeExamRequest
            
        Case conMenu_PatholSpeExam_Accept
            Call SpeExamined_Accept
            
        Case conMenu_PatholSpeExam_Finish
            Call SpeExamined_Sure
    End Select
End Sub


Public Sub IWorkMenu_zlUpdateMenu(ByVal control As XtremeCommandBars.ICommandBarControl)
'���²˵�
    Dim blnIsAllowSpeExam As Boolean

    If Not Me.Visible Then
        control.Enabled = False
        Exit Sub
    End If
    
    blnIsAllowSpeExam = CheckPopedom(mstrPrivs, "�����黯") Or CheckPopedom(mstrPrivs, "���Ӳ���") Or CheckPopedom(mstrPrivs, "����Ⱦɫ") And mlngAdviceID > 0
    
    Select Case control.ID
        Case conMenu_PatholSpeExam_LAB
            control.Enabled = blnIsAllowSpeExam And mrecStudyInf.strPatholNumber <> ""
            
        Case conMenu_PatholSpeExam_List
            control.Enabled = blnIsAllowSpeExam And mrecStudyInf.strPatholNumber <> ""
            
        Case conMenu_PatholSpeExam_RequestView
            control.Enabled = blnIsAllowSpeExam And mrecStudyInf.strPatholNumber <> ""
            
        Case conMenu_PatholSpeExam_Accept
            control.Enabled = blnIsAllowSpeExam And Not mblnReadOnly
            
        Case conMenu_PatholSpeExam_Finish
            control.Enabled = blnIsAllowSpeExam And Not mblnReadOnly
    End Select
End Sub


Public Sub IWorkMenu_zlPopupMenu(objPopup As XtremeCommandBars.ICommandBar)
'�����Ҽ��˵�
    Exit Sub
End Sub

Public Sub IWorkMenu_zlRefreshSubMenu(objMenuBar As Object)
'ˢ�µ������Ӳ˵�
    Exit Sub
End Sub
'*************************************************************************************************


Private Function CreateModuleMenu(objMenuControl As CommandBarControls, _
    ByVal lngType As XTPControlType, ByVal lngID As Long, ByVal strCaption As String, _
    Optional strToolTip As String = "", Optional lngIconId As Long = 0, Optional blnStartGroup As Boolean = False) As CommandBarControl
'������ģ���ڵĲ˵�
    
    Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption)
    CreateModuleMenu.ID = lngID '������ﲻָ��id�����ܽ���Щ�˵���ӵ��Ҽ��˵���
    
    If lngIconId <> 0 Then CreateModuleMenu.IconId = lngIconId
    If blnStartGroup Then CreateModuleMenu.BeginGroup = True
    If strToolTip <> "" Then CreateModuleMenu.ToolTipText = strToolTip
    
    CreateModuleMenu.Category = M_STR_MODULE_MENU_TAG
End Function


Public Sub zlInitModule(ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngDepartId As Long, Optional owner As Object = Nothing)
'��ʼ��ģ�����
    mlngModule = lngModule
    mstrPrivs = strPrivs
    mlngCurDeptId = lngDepartId
    
    If Not owner Is Nothing Then Set mobjOwner = owner
End Sub


Public Sub zlUpdateAdviceInf(ByVal lngAdviceID As Long, ByVal lngSendNO As Long, _
    ByVal lngStudyState As Long, ByVal blnMoved As Boolean)
'����ҽ����Ϣ
    mlngAdviceID = lngAdviceID
    mlngSendNo = lngSendNO
    mblnMoved = blnMoved
    mlngStudyState = lngStudyState
    mblnReadOnly = False
    mblnRefreshState = True
    
    '���ݱ�ת��ʱ��û��Ȩ��ʱ��״̬Ϊָ��״̬ʱ����ģ��Ϊֻ��
    If blnMoved Or blnMoved Or lngStudyState = 6 Or lngStudyState = 5 Or lngStudyState = 0 Or lngStudyState = 1 Or lngStudyState = -2 Then
        mblnReadOnly = True
    End If

End Sub

Public Sub zlRefreshFace(Optional blnForceRefresh As Boolean = False)
'ˢ�½�������
    If (mlngTmpAdviceId = mlngAdviceID And mlngTmpSendNo = mlngSendNo And mblnRefreshState) And Not blnForceRefresh Then Exit Sub
    
    mlngTmpAdviceId = mlngAdviceID
    mlngTmpSendNo = mlngSendNo
    mblnRefreshState = True
    
    If mlngAdviceID <= 0 Then
        Call ConfigSpeExamFace(False, "ҽ��ID��Ч���顣")
        Exit Sub
    End If
    
   Call GetPatholStudyState(mlngAdviceID, mrecStudyInf)
        
    '�����ʾ����
    Call ufgData.ClearListData
       
    If mrecStudyInf.strPatholNumber = "" Then
        Call ConfigSpeExamFace(False, "�ü����δ������Ч�Ĳ���ţ���ȷ�ϸü���Ƿ��ѱ����ա�")
        Exit Sub
    Else
        Call ConfigSpeExamFace(True)
    End If
    
    Call ConfigSpeExamType(mrecStudyInf.lngPatholAdviceId)
    
    Call QuerySpeExamData(mrecStudyInf.lngPatholAdviceId)
    
    Call RefreshSpeExamInf
    
    Call ConfigPopedom(mblnReadOnly)
    
'    If Not (owner Is Nothing) Then
'        Call Me.Show(1, owner)
'    End If
End Sub

Public Sub zlRefresh(lngAdviceID As Long, ByVal blnReadOnly As Boolean, _
    strPrivs As String, ByVal blnMoved As Boolean, _
    ByVal lngCurDepartmentId As Long, Optional owner As Form = Nothing)
    
    If lngAdviceID <= 0 Then
        Call ConfigSpeExamFace(False, "ҽ��ID��Ч���顣")
        Exit Sub
    End If
    
'    If mlngCurAdviceId = lngAdviceID Then
'        Call ConfigSpeExamFace(False, "�ü����δ������Ч�Ĳ���ţ���ȷ�ϸü���Ƿ��ѱ����ա�")
'        Exit Sub
'    End If
    
    mlngAdviceID = lngAdviceID
    mstrPrivs = strPrivs
    mblnMoved = blnMoved
    mlngCurDeptId = lngCurDepartmentId
    
    Call GetPatholStudyState(lngAdviceID, mrecStudyInf)
        
    '�����ʾ����
    Call ufgData.ClearListData
       
    If mrecStudyInf.strPatholNumber = "" Then
        Call ConfigSpeExamFace(False, "�ü����δ������Ч�Ĳ���ţ���ȷ�ϸü���Ƿ��ѱ����ա�")
        Exit Sub
    Else
        Call ConfigSpeExamFace(True)
    End If
    
    Call ConfigSpeExamType(mrecStudyInf.lngPatholAdviceId)
    
    Call QuerySpeExamData(mrecStudyInf.lngPatholAdviceId)
    
    Call RefreshSpeExamInf
    
    Call ConfigPopedom(blnReadOnly)
    
    If Not (owner Is Nothing) Then
        Call Me.Show(1, owner)
    End If
End Sub




Private Sub RefreshSpeExamInf()
'ˢ����Ƭ��¼����
    Dim i As Long
    Dim lngNeedCount As Long
    Dim lngTotal As Long
    
    lngNeedCount = 0
    lngTotal = 0
    
    
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsNullRow(i) Then
        
            
            lngTotal = lngTotal + 1
                
            If ufgData.Text(i, gstrSpeExam_��ǰ״̬) <> "�����" Then
                lngNeedCount = lngNeedCount + 1
            End If
        End If
    Next i
    
    labRecordInf.Caption = "��ǰ����Ŀ����" & lngTotal & "    ��ǰ������Ŀ����" & lngNeedCount
    
End Sub


Private Sub ConfigPopedom(ByVal blnIsReadOnly As Boolean)
'����Ȩ��
    Dim blnIsAllowSpeExam As Boolean
    
    blnIsAllowSpeExam = CheckPopedom(mstrPrivs, "�����黯") Or CheckPopedom(mstrPrivs, "���Ӳ���") Or CheckPopedom(mstrPrivs, "����Ⱦɫ")
    
    tbrMain.Buttons("tbAcceptSpeExam").Enabled = blnIsAllowSpeExam And Not blnIsReadOnly
    tbrMain.Buttons("tbEndSpeExam").Enabled = blnIsAllowSpeExam And Not blnIsReadOnly
    
    tbrMain.Buttons("tbViewRequest").Enabled = blnIsAllowSpeExam
    tbrMain.Buttons("tbLAB").Enabled = blnIsAllowSpeExam
    tbrMain.Buttons("tbList").Enabled = blnIsAllowSpeExam
    
    ufgData.ReadOnly = blnIsReadOnly
    
    optMianYi.Enabled = CheckPopedom(mstrPrivs, "�����黯")
    optTeShu.Enabled = CheckPopedom(mstrPrivs, "����Ⱦɫ")
    optFenZi.Enabled = CheckPopedom(mstrPrivs, "���Ӳ���")
End Sub

Private Sub ConfigSpeExamFace(ByVal blnIsValid As Boolean, Optional ByVal strHintInf As String = "")
'�����ؼ����
    tbrMain.Buttons("tbAcceptSpeExam").Enabled = blnIsValid
    tbrMain.Buttons("tbEndSpeExam").Enabled = blnIsValid
    
    tbrMain.Buttons("tbViewRequest").Enabled = blnIsValid
    tbrMain.Buttons("tbLAB").Enabled = blnIsValid
    tbrMain.Buttons("tbList").Enabled = blnIsValid

    optFenZi.Enabled = blnIsValid
    optMianYi.Enabled = blnIsValid
    optTeShu.Enabled = blnIsValid

    chkYSQ.Enabled = blnIsValid
    chkYJS.Enabled = blnIsValid
    chkYWC.Enabled = blnIsValid
    
    If blnIsValid Then
        Call ufgData.CloseHintInf
    Else
        Call ufgData.ShowHintInf(strHintInf)
        labRecordInf.Caption = ""
    End If
End Sub


Private Function GetSelectSpeExamType()
'ȡ��ѡ����ؼ�����
    If optMianYi.value Then GetSelectSpeExamType = C_INT_MIANYI
    If optTeShu.value Then GetSelectSpeExamType = C_INT_TESHU
    If optFenZi.value Then GetSelectSpeExamType = C_INT_FENZI
End Function


Private Sub AdjustFace()
    '�������沼��
    framSpeExam.Left = 0
    framSpeExam.Top = tbrMain.Top + tbrMain.Height + 120
    framSpeExam.Width = Me.Width - 0
    framSpeExam.Height = Me.Height - tbrMain.Height - 240
    
    
    ufgData.Left = 120
    ufgData.Top = 280 + IIf(mbytFontSize = 9, 0, 120)
    ufgData.Width = framSpeExam.Width - 240
    ufgData.Height = framSpeExam.Height - labRecordInf.Height - 600
    
    labRecordInf.Left = 120
    labRecordInf.Top = framSpeExam.Height - labRecordInf.Height - 120
    
    
    '����FrameCheckλ��
    
     FramCheck.Top = framSpeExam.Height - labRecordInf.Height - 120
     FramCheck.Left = framSpeExam.Width - FramCheck.Width - 200
     
     chkYJS.Top = 0
     chkYSQ.Top = 0
     chkYWC.Top = 0
End Sub

Private Sub ConfigSpeExamType(ByVal strPatholNum As String)
'���õ�ǰ�ؼ�����
    Dim lngType As Long
    
    lngType = GetCurSpeExamType(strPatholNum)
    
    Select Case lngType
        Case C_INT_MIANYI
            optMianYi.value = True
        Case C_INT_TESHU
            optTeShu.value = True
        Case C_INT_FENZI
            optFenZi.value = True
    End Select

End Sub


Private Sub InitSpeExamList()
'��ʼ���ؼ��б�
    Dim strTemp As String
    
    ufgData.IsKeepRows = True
    ufgData.GridRows = glngMaxRowCount
    '�����и�
    ufgData.RowHeightMin = glngStandardRowHeight
    
    ufgData.IsCopyMode = True
    
    '�ж����ݿ�������Ƿ������� �����ȡ���ݿ����  û�������Ĭ��
    strTemp = zlDatabase.GetPara("�ؼ���Ϣ�б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
     
    If strTemp = "" Then
        ufgData.ColNames = gstrSpeExamCols
    Else
        ufgData.ColNames = strTemp
    End If
    
    ufgData.DefaultColNames = gstrSpeExamCols
    ufgData.ColConvertFormat = gstrSpeExamConvertFormat
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo errHandle
    Call ExecuteTbrOperation(Button.Key)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub tbrMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error GoTo errHandle
    Call ExecuteTbrOperation(ButtonMenu.Key)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ExecuteTbrOperation(ByVal strButtonKey As String)
    Dim strKey As String
    
    strKey = UCase(strButtonKey)
    
    Select Case strKey
        Case UCase("tbLab"), UCase("tbLabPreview")
            'Ԥ����ǩ
            Call PrintSpeExamLabel(False)
            
        Case UCase("tbLabPrint")
            '��ӡ��ǩ
            Call PrintSpeExamLabel(True)
        
        Case UCase("tbList"), UCase("tbListPreview")
            'Ԥ���嵥
            Call PrintWorkList(False)
        
        Case UCase("tbListPrint")
            '��ӡ�嵥
            Call PrintWorkList(True)
        
        Case UCase("tbViewRequest")
            '�鿴����
            Call ShowSpeExamRequest
        
        Case UCase("tbAcceptSpeExam")
            '�ؼ����
            Call SpeExamined_Accept
        
        Case UCase("tbEndSpeExam")
            '�ؼ����
            Call SpeExamined_Sure
            
    End Select
End Sub


Private Sub ufgData_OnColFormartChange()
 '�����б����
     zlDatabase.SetPara "�ؼ���Ϣ�б�����", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
End Sub


Private Function GetCurSpeExamType(ByVal lngPatholAdviceId As Long) As Long
'ȡ�õ�ǰ�ؼ�����
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    
    
    strSql = "select ���߹���,���ӹ���,��Ⱦ���� from ��������Ϣ where ����ҽ��ID=[1]"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatholAdviceId)
    
    If rsData.RecordCount > 0 Then
    
        If Nvl(rsData!���߹���) > 0 Then
            GetCurSpeExamType = 0
            If CheckPopedom(mstrPrivs, "�����黯") Then Exit Function
        End If
        
        If Nvl(rsData!��Ⱦ����) > 0 Then
            GetCurSpeExamType = 1
            If CheckPopedom(mstrPrivs, "����Ⱦɫ") Then Exit Function
        End If
        
        If Nvl(rsData!���ӹ���) > 0 Then
            GetCurSpeExamType = 2
            If CheckPopedom(mstrPrivs, "���Ӳ���") Then Exit Function
        End If
    End If
    
    
    
    If CheckPopedom(mstrPrivs, "�����黯") Then
        GetCurSpeExamType = 0
        Exit Function
    End If
    
    If CheckPopedom(mstrPrivs, "����Ⱦɫ") Then
        GetCurSpeExamType = 1
        Exit Function
    End If
    
    If CheckPopedom(mstrPrivs, "���Ӳ���") Then
        GetCurSpeExamType = 2
        Exit Function
    End If
End Function


Private Sub QuerySpeExamData(ByVal lngPatholAdviceId As Long)
'�����ؼ����ݣ����������黯�����Ӳ�������Ⱦɫ��
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    strSql = "select to_Number(a.ID) as ID,a.�Ŀ�ID,to_Number(b.���) as ���,b.�걾����,to_Number(a.����ID) as ����ID,to_Number(a.����ID) as ����ID,c.��������,to_Number(a.��������) as ��������,to_Number(a.��ǰ״̬) as ��ǰ״̬,a.��Ŀ���, d.����ʱ��, a.���ʱ��,a.�ؼ�ҽʦ,to_Number(a.�嵥״̬) as �嵥״̬,to_Number(a.�ؼ�����) as �ؼ�����,to_Number(a.�ؼ�ϸĿ) as �ؼ�ϸĿ" & _
            " from �����ؼ���Ϣ a, ����ȡ����Ϣ b, ��������Ϣ c, ����������Ϣ d " & _
            " where a.�Ŀ�id=b.�Ŀ�id and a.����id=c.����id and a.����ID=d.����ID and b.����ҽ��ID=[1] and (�ؼ�����=-1"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    If CheckPopedom(mstrPrivs, "�����黯") Then strSql = strSql & " or �ؼ�����=0"
    If CheckPopedom(mstrPrivs, "����Ⱦɫ") Then strSql = strSql & " or �ؼ�����=1"
    If CheckPopedom(mstrPrivs, "���Ӳ���") Then strSql = strSql & " or �ؼ�����=2"
        
    strSql = strSql & ") order by �ؼ�����,��ǰ״̬,���,ID"
            
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatholAdviceId)
    
    '��ȡ��Ӧ�ؼ���Ŀ���� ����������
    Call FilterData
End Sub



Public Sub ReSetFormFontSize(ByVal bytFontSize As Byte)
'����:�������ù���վ����������С
On Error GoTo errHandle

    Dim objCtrl As control
    Dim CtlFont As StdFont
    Dim strFontType  As String
    
    
    
    mbytFontSize = bytFontSize
    
    optMianYi.Left = optMianYi.Left + IIf(optMianYi.FontSize = bytFontSize, 0, IIf(bytFontSize = 9, -300, 300))
    optTeShu.Left = optTeShu.Left + IIf(optTeShu.FontSize = bytFontSize, 0, IIf(bytFontSize = 9, -500, 500))
    optFenZi.Left = optFenZi.Left + IIf(optFenZi.FontSize = bytFontSize, 0, IIf(bytFontSize = 9, -700, 700))
    
    chkYSQ.Left = chkYSQ.Left + IIf(chkYSQ.FontSize = bytFontSize, 0, IIf(bytFontSize = 9, 500, -500))
    chkYJS.Left = chkYJS.Left + IIf(chkYSQ.FontSize = bytFontSize, 0, IIf(bytFontSize = 9, 300, -300))
    
    Set CtlFont = New StdFont
    Me.FontSize = bytFontSize
    strFontType = IIf(IsUseClearType = True, "΢���ź�", "����")
    
    CtlFont.Name = strFontType
    CtlFont.Size = bytFontSize
    
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("Frame")
            If objCtrl.Name = "FramCheck" Then
                objCtrl.Height = TextHeight("��") * 1.7
            End If
            objCtrl.Font.Size = bytFontSize
            objCtrl.Font.Name = strFontType
        Case UCase("TabStrip") 'ҳ��ؼ�
            objCtrl.Font.Size = bytFontSize
            objCtrl.Font.Name = strFontType
        Case UCase("Label")
            objCtrl.FontSize = bytFontSize
            objCtrl.FontName = strFontType
            objCtrl.Height = TextHeight("��") + 150
        Case UCase("vsFlexGrid")
            objCtrl.Cell(flexcpFont, 0, 0, 0, objCtrl.DataGrid.Cols - 1) = CtlFont
            objCtrl.Font = CtlFont
            objCtrl.RowHeight(0) = TextHeight("��") + 150
         Case UCase("ucFlexGrid")
            objCtrl.DataGrid.Cell(flexcpFont, 0, 0, 0, objCtrl.DataGrid.Cols - 1) = CtlFont
            objCtrl.DataGrid.Font = CtlFont
            objCtrl.DataGrid.RowHeight(0) = TextHeight("��") + 150
        Case UCase("ComboBox")
            objCtrl.FontSize = bytFontSize
            objCtrl.FontName = strFontType
        Case UCase("OptionButton")
            objCtrl.FontSize = bytFontSize
            objCtrl.FontName = strFontType
            objCtrl.Width = TextWidth("����" & objCtrl.Caption)
            objCtrl.Height = TextHeight("��") + 100
        Case UCase("listBox")
            objCtrl.Font = CtlFont
        Case UCase("CheckBox")
            objCtrl.FontSize = bytFontSize
            objCtrl.FontName = strFontType
            objCtrl.Width = TextWidth("����" & objCtrl.Caption)
        Case UCase("DTPicker")
            objCtrl.Font.Size = bytFontSize
            objCtrl.FontN.ame = strFontType
            objCtrl.Width = TextWidth("2012-01-01 23:59:59") * 1.25
            objCtrl.Height = TextHeight("��") * 1.5
        Case UCase("textBox")
          objCtrl.FontSize = bytFontSize
          objCtrl.FontName = strFontType
        Case UCase("ReportControl")
            Set objCtrl.PaintManager.CaptionFont = CtlFont
            Set objCtrl.PaintManager.TextFont = CtlFont
            objCtrl.Redraw
        Case UCase("DockingPane")
            Set objCtrl.PaintManager.CaptionFont = CtlFont
        Case UCase("CommandBars")
            Set objCtrl.Options.Font = CtlFont
        Case UCase("TabControl")
            Set objCtrl.PaintManager.Font = CtlFont
        Case UCase("CommandButton")
            objCtrl.FontName = strFontType
            objCtrl.FontSize = bytFontSize
        End Select
    Next
    
    Call AdjustFace
    
    Exit Sub
errHandle:
End Sub



Private Sub FilterData()
'��������
     Dim strFilter As String
     Dim lngCurSpeExamType As Long
     
    If ufgData.AdoData Is Nothing Then Exit Sub
            
    If optMianYi.value Then lngCurSpeExamType = 0
    If optTeShu.value Then lngCurSpeExamType = 1
    If optFenZi.value Then lngCurSpeExamType = 2
            
    '�жϵ�ǰ״̬�����ݸ�ѡ����ʾ����
    If chkYSQ.value <> 0 Then
        If strFilter <> "" Then strFilter = strFilter & " or "
        strFilter = strFilter & "(��ǰ״̬=0 and �ؼ�����=" & lngCurSpeExamType & ")"
    End If
    
    If chkYJS.value <> 0 Then
         If strFilter <> "" Then strFilter = strFilter & " or "
         strFilter = strFilter & "(��ǰ״̬=1 and �ؼ�����=" & lngCurSpeExamType & ")"
    End If
    
    If chkYWC.value <> 0 Then
         If strFilter <> "" Then strFilter = strFilter & " or "
         strFilter = strFilter & "(��ǰ״̬=2 and �ؼ�����=" & lngCurSpeExamType & ")"
    End If
    
    If strFilter = "" Then
         strFilter = "(��ǰ״̬=9 and �ؼ�����=" & lngCurSpeExamType & ")"
    End If
    
    ufgData.AdoData.Filter = strFilter
    
    'ˢ������
    Call ufgData.RefreshData

    Call RefreshSpeExamInf
    
End Sub

Private Sub chkYSQ_Click()
On Error GoTo errHandle
    '���ù������ݷ���
    Call FilterData
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkYJS_Click()
On Error GoTo errHandle

    Call FilterData
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkYWC_Click()
On Error GoTo errHandle

    Call FilterData
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub SpeExamined_Accept()
'�ؼ����
    Dim strSql As String
    Dim i As Long
    Dim blnIsExit As Boolean
    
    blnIsExit = False
    Select Case GetSelectSpeExamType
        Case C_INT_MIANYI
            If mrecStudyInf.lngMianYiStep <> TExecuteStep.NeedDo And mrecStudyInf.lngMianYiStep <> TExecuteStep.AcceptDo Then blnIsExit = True
            
        Case C_INT_FENZI
            If mrecStudyInf.lngFenZiStep <> TExecuteStep.NeedDo And mrecStudyInf.lngFenZiStep <> TExecuteStep.AcceptDo Then blnIsExit = True
            
        Case C_INT_TESHU
            If mrecStudyInf.lngTeRanStep <> TExecuteStep.NeedDo And mrecStudyInf.lngTeRanStep <> TExecuteStep.AcceptDo Then blnIsExit = True
    End Select
    
    '���ؼ�׶Σ����ܽ��н���
    If blnIsExit Then
        
        Call MsgBoxD(Me, "��δ�����ؼ�׶Σ����ܽ����ؼ�ȷ�ϲ�����", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    
    If Not CheckAllowAccept Then
        Call MsgBoxD(Me, "û����Ҫ���н��ܵ��ؼ���Ŀ��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strSql = "Zl_�����ؼ�_����('" & mrecStudyInf.lngPatholAdviceId & "'," & GetSelectSpeExamType & ",'" & UserInfo.���� & "')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Select Case GetSelectSpeExamType
        Case C_INT_MIANYI
            mrecStudyInf.lngMianYiStep = TExecuteStep.AcceptDo
        Case C_INT_FENZI
            mrecStudyInf.lngFenZiStep = TExecuteStep.AcceptDo
        Case C_INT_TESHU
            mrecStudyInf.lngTeRanStep = TExecuteStep.AcceptDo
    End Select
    
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsEmptyKey(i) Then
            If ufgData.Text(i, gstrSpeExam_��ǰ״̬) = "������" Then
                Call ufgData.SyncText(i, gstrSpeExam_��ǰ״̬, "�ѽ���", True)
                Call ufgData.SyncText(i, gstrSpeExam_�ؼ�ҽʦ, UserInfo.����, True)
            End If
        End If
    Next i
    
    Call MsgBoxD(Me, "�ѽ���" & Decode(GetSelectSpeExamType, 0, "�����黯", 1, "����Ⱦɫ", 2, "���Ӳ���") & "��顣", vbOKOnly, Me.Caption)
End Sub


Private Function CheckAllowAccept() As Boolean
    Dim i As Long
    
    CheckAllowAccept = False
    
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsEmptyKey(i) Then
            If Trim(ufgData.Text(i, gstrSpeExam_��ǰ״̬)) = "������" Then
                CheckAllowAccept = True
                Exit Function
            End If
        End If
    Next i
End Function



Private Sub SpeExamined_Sure()
'�ؼ�ȷ��
    Dim strSql As String
    Dim i As Long
    Dim dtServicesTime As Date
    Dim blnIsExit As Boolean
    
    blnIsExit = False
    Select Case GetSelectSpeExamType
        Case C_INT_MIANYI
            If mrecStudyInf.lngMianYiStep <> TExecuteStep.NeedDo And mrecStudyInf.lngMianYiStep <> TExecuteStep.AcceptDo Then blnIsExit = True
            
        Case C_INT_FENZI
            If mrecStudyInf.lngFenZiStep <> TExecuteStep.NeedDo And mrecStudyInf.lngFenZiStep <> TExecuteStep.AcceptDo Then blnIsExit = True
            
        Case C_INT_TESHU
            If mrecStudyInf.lngTeRanStep <> TExecuteStep.NeedDo And mrecStudyInf.lngTeRanStep <> TExecuteStep.AcceptDo Then blnIsExit = True
    End Select
    
    '���ؼ�׶Σ����ܽ���ȷ��
    If blnIsExit Then
        
        Call MsgBoxD(Me, "��δ�����ؼ�׶Σ����ܽ����ؼ�ȷ�ϲ�����", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    
    If Not CheckAllowSpeExamSure Then
        Call MsgBoxD(Me, "�ѽ��ܵ��ؼ���Ŀ�����δ��ȫ¼�룬���ܽ���ȷ�ϡ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
'    '������Ŀ���
'    Call SpeExamined_Save

    dtServicesTime = zlDatabase.Currentdate
    
    strSql = "Zl_�����ؼ�_ȷ��('" & mrecStudyInf.lngPatholAdviceId & "'," & GetSelectSpeExamType & "," & To_Date(dtServicesTime) & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Select Case GetSelectSpeExamType
        Case C_INT_MIANYI
            mrecStudyInf.lngMianYiStep = TExecuteStep.AlreadDo
        Case C_INT_FENZI
            mrecStudyInf.lngFenZiStep = TExecuteStep.AlreadDo
        Case C_INT_TESHU
            mrecStudyInf.lngTeRanStep = TExecuteStep.AlreadDo
    End Select
    
    
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsEmptyKey(i) Then
            If ufgData.Text(i, gstrSpeExam_��ǰ״̬) = "�ѽ���" Then
                Call ufgData.SyncText(i, gstrSpeExam_��ǰ״̬, "�����", True)
                Call ufgData.SyncText(i, gstrSpeExam_���ʱ��, dtServicesTime, True)
            End If
            
            If ufgData.Text(i, gstrSpeExam_��ǰ״̬) = "������" Then
                Select Case GetSelectSpeExamType
                    Case C_INT_MIANYI
                        mrecStudyInf.lngMianYiStep = TExecuteStep.NeedDo
                    Case C_INT_FENZI
                        mrecStudyInf.lngFenZiStep = TExecuteStep.NeedDo
                    Case C_INT_TESHU
                        mrecStudyInf.lngTeRanStep = TExecuteStep.NeedDo
                End Select
            End If
        End If
    Next i
    
    '�����ؼ�ȷ���¼�
    Call SendMsgToMainWindow(Me, wetSpeExamSure, mlngAdviceID)
    
    Call MsgBoxD(Me, "����ɶ�" & Decode(GetSelectSpeExamType, 0, "�����黯", 1, "����Ⱦɫ", 2, "���Ӳ���") & "����ȷ�ϡ�", vbOKOnly, Me.Caption)
End Sub


Public Function CheckAllowSpeExamSure() As Boolean
'�Ƿ������ؼ�ȷ��(������������������Ŀ���Ϊ�յ���Ŀ���ܽ���ȷ��)
    Dim i As Long
    
    CheckAllowSpeExamSure = True
    
'    For i = 1 To ufgData.GridRows - 1
'        If Not ufgData.KeyEmpty(i) Then
'            If Trim(ufgData.Text(i, gstrSpeExam_��Ŀ���)) = "" And ufgData.Text(i, gstrSpeExam_��ǰ״̬) = "�ѽ���" Then
'                CheckAllowSpeExamSure = False
'                Exit Function
'            End If
'        End If
'    Next i
End Function






Private Sub PrintSpeExamLabel(Optional ByVal blnIsPrint As Boolean = True)
'��ӡ�ؼ���Ŀ��ǩ
    Dim i As Long
    Dim j As Long
    Dim strValue(5) As String
    
    j = 0
    strValue(0) = "0": strValue(1) = "0": strValue(2) = "0": strValue(3) = "0": strValue(4) = "0": strValue(5) = "0"
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsEmptyKey(i) Then
            If zlCommFun.ActualLen(strValue(j)) > 2000 Then
                j = j + 1
                strValue(j) = ""
            End If
    
            If strValue(j) <> "" Then strValue(j) = strValue(j) & ","
    
            strValue(j) = strValue(j) & ufgData.KeyValue(i)
        End If
    Next i
    
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_11", Me, "��ĿID1=" & strValue(0), "��ĿID2=" & strValue(1), "��ĿID3=" & strValue(2), "��ĿID4=" & strValue(3), "��ĿID5=" & strValue(4), "��ĿID6=" & strValue(5), IIf(blnIsPrint, 2, 1))
End Sub


Private Sub PrintSelectSpeExamLabel(Optional ByVal blnIsPrint As Boolean = True)
'��ӡѡ��ĲĿ��ǩ
On Error GoTo errHandle
    Dim strValue(5) As String
    
    
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ��ӡ���ؼ��¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.IsNullRow(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ҫ��ӡ���ؼ��¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strValue(0) = ufgData.KeyValue(ufgData.SelectionRow)
    strValue(1) = "0": strValue(2) = "0": strValue(3) = "0": strValue(4) = "0": strValue(5) = "0"

    
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_11", Me, "��ĿID1=" & strValue(0), "��ĿID2=" & strValue(1), "��ĿID3=" & strValue(2), "��ĿID4=" & strValue(3), "��ĿID5=" & strValue(4), "��ĿID6=" & strValue(5), IIf(blnIsPrint, 2, 1))
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub SpeExamined_Save()
'�����ؼ���Ŀ
    Dim i As Long
    Dim strSql As String
    
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsEmptyKey(i) Then
            If ufgData.Text(i, gstrSpeExam_��ǰ״̬) = "�ѽ���" Then
                strSql = "Zl_�����ؼ�_��Ŀ¼��(" & ufgData.KeyValue(i) & ",null)"
        
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            End If
        End If
    Next i
    
End Sub

Private Sub cmdSave_Click()
'������Ŀ���
On Error GoTo errHandle
    If Not CheckAllowSpeExamSure Then
        Call MsgBoxD(Me, "��Ŀ�����δ��ȫ¼�룬���ܽ��б��档", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call SpeExamined_Save
    
    Call MsgBoxD(Me, "��Ŀ����ѱ��档", vbOKOnly, Me.Caption)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub PrintWorkList(Optional ByVal blnIsPrint As Boolean = True)
'��ӡ�ؼ칤���б�
    Dim i As Long
    Dim j As Long
    Dim strValue(5) As String
    
    j = 0
    strValue(0) = "0": strValue(1) = "0": strValue(2) = "0": strValue(3) = "0": strValue(4) = "0": strValue(5) = "0"
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsEmptyKey(i) Then
            If zlCommFun.ActualLen(strValue(j)) > 2000 Then
                j = j + 1
                strValue(j) = ""
            End If
    
            If strValue(j) <> "" Then strValue(j) = strValue(j) & ","
    
            strValue(j) = strValue(j) & ufgData.KeyValue(i)
        End If
    Next i
    
    '�����嵥�Ĵ�ӡ��ʹ�ô�����Ԥ���ķ�ʽ
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_10", Me, "��ĿID=" & strValue(0), "��ĿID1=" & strValue(1), "��ĿID2=" & strValue(2), "��ĿID3=" & strValue(3), "��ĿID4=" & strValue(4), "��ĿID5=" & strValue(5), IIf(blnIsPrint, 2, 1))
    
End Sub


Private Sub ShowSpeExamRequest()
'��ʾ�ؼ�����
Dim frmRequestView As New frmPatholRequisition_View
On Error GoTo errFree
    Call frmRequestView.ShowRequestViewWind(mrecStudyInf.lngPatholAdviceId, GetSelectSpeExamType, mblnMoved, Me)
errFree:
    Call Unload(frmRequestView)
    Set frmRequestView = Nothing
End Sub

Private Sub Form_Initialize()
    mKeyCode = -1
    mKeyShift = -1
    
    Set zlReport = New zl9Report.clsReport
    mblnAutoAcceptOfAfterPrint = False
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    '��ʼ���ؼ��б�
    Call InitSpeExamList
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub Form_Resize()
On Error Resume Next
    Call AdjustFace
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Set zlReport = Nothing
End Sub

'Private Sub vfgData_KeyDown(KeyCode As Integer, Shift As Integer)
'    mKeyCode = KeyCode
'    mKeyShift = Shift
'End Sub
'
'Private Sub vfgData_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
'    mKeyCode = KeyCode
'    mKeyShift = Shift
'End Sub

Private Function GetRomanAscii(ByVal lngNum As Long) As Integer
'������ת�����������ֵ�ascii
    GetRomanAscii = Decode(lngNum, 49, -23823, 50, -23822, 51, 23821, 52, -23820, 53, -23819, 54, -23818, 55, -23817, 56, -23816, 57, -23815)
End Function


'
'Private Sub vfgData_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'    '��С���̵�*�滻Ϊ%����
'    If KeyAscii = 42 And mKeyShift <> 1 Then KeyAscii = 37
'
''    If mKeyShift = 2 Then
''        If KeyAscii >= 49 And KeyAscii <= 57 Then KeyAscii = GetRomanAscii(KeyAscii)
''    End If
'End Sub
'
'Private Sub vfgData_KeyUp(KeyCode As Integer, Shift As Integer)
'    mKeyCode = -1
'    mKeyShift = -1
'End Sub

'Private Sub vfgData_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
'    mKeyCode = -1
'    mKeyShift = -1
'End Sub

'Private Sub vfgData_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
''����Ŀ����⣬�����в�����༭
'    If Col <> mvfgSpeExam.GetColumnIndex(gstrSpeExam_��Ŀ���) Then Cancel = True
'End Sub

Private Sub UpdateWorkListPrintState()
'�ڴ�ӡ�󣬸��¹����嵥�Ĵ�ӡ״̬
    Dim strSql As String
    Dim i As Long
        
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsEmptyKey(i) Then
            strSql = "Zl_�����ؼ�_�嵥��ӡ(" & ufgData.KeyValue(i) & ")"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)

            Call ufgData.SyncText(i, gstrSpeExamWork_�嵥״̬, "�Ѵ�ӡ", True)
        End If
    Next i
End Sub



Private Sub OptFenZi_Click()
'����ָ���ؼ����͵�����
On Error GoTo errHandle
     Call FilterData
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub optMianYi_Click()
'����ָ���ؼ����͵�����
On Error GoTo errHandle
    Call FilterData
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub OptTeShu_Click()
'����ָ���ؼ����͵�����
On Error GoTo errHandle
    Call FilterData
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ShowAntibodyInf(ByVal lngAntibodyRow As Long)
'��ʾ������ϸ��Ϣ
    Dim frmAntibodyInf As New frmPatholRequisition_AntibodyInf
    On Error GoTo errFree
        Call frmAntibodyInf.ShowAntibodyInf(ufgData.Text(lngAntibodyRow, gstrSpeExam_����ID), Me)
errFree:
    Call Unload(frmAntibodyInf)
    Set frmAntibodyInf = Nothing
    
End Sub



Private Sub ufgData_OnCellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo errHandle
    Call ShowAntibodyInf(Row)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub ufgData_OnColsNameReSet()
On Error GoTo errHandle

    Call QuerySpeExamData(mrecStudyInf.lngPatholAdviceId)
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnStartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'ֻ���ؼ���Ŀ�����պ󣬲��ܽ��б༭
    Dim strState As String
    
    strState = ufgData.Text(Row, gstrSpeExam_��ǰ״̬)
    
    If Col = ufgData.GetColIndex(gstrSpeExam_��Ŀ���) Then
        Cancel = IIf(Trim(strState) = "" Or strState = "������", True, False)
        
        If Cancel Then
            Call MsgBoxD(Me, "�ؼ���Ŀδ�����ܣ����ܽ���¼�롣", vbOKOnly, Me.Caption)
        End If
    End If
End Sub

Private Sub zlReport_AfterPrint(ByVal ReportNum As String)
On Error GoTo errHandle
    '��������ؼ��嵥��ӡ����ֱ���˳�
    If ReportNum <> "ZL1_PATHOLSPEEXAM_01" Then Exit Sub
    
    Call UpdateWorkListPrintState
    
    If mblnAutoAcceptOfAfterPrint Then
    '��ӡ���Զ�����
        Call SpeExamined_Accept
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

