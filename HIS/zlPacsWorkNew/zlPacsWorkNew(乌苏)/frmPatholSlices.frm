VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmPatholSlices 
   Caption         =   "������Ƭ"
   ClientHeight    =   8955
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   10665
   Icon            =   "frmPatholSlices.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   10665
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.ImageList imgTbrS 
      Left            =   8415
      Top             =   765
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
            Picture         =   "frmPatholSlices.frx":000C
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSlices.frx":0C7E
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSlices.frx":18F0
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSlices.frx":2562
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSlices.frx":31D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSlices.frx":3E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSlices.frx":4AB8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame framSlices 
      Caption         =   "��Ƭ��¼"
      Height          =   7215
      Left            =   240
      TabIndex        =   1
      Top             =   795
      Width           =   9975
      Begin VB.Frame FramCheck 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   6120
         TabIndex        =   3
         Top             =   6840
         Width           =   3735
         Begin VB.CheckBox chkYWC 
            Caption         =   "�����"
            Height          =   180
            Left            =   2760
            TabIndex        =   6
            Top             =   30
            Width           =   855
         End
         Begin VB.CheckBox chkYJS 
            Caption         =   "�ѽ���"
            Height          =   180
            Left            =   1800
            TabIndex        =   5
            Top             =   30
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkWCL 
            Caption         =   "δ����"
            Height          =   255
            Left            =   720
            TabIndex        =   4
            Top             =   0
            Value           =   1  'Checked
            Width           =   855
         End
      End
      Begin zl9PACSWork.ucFlexGrid ufgData 
         Height          =   6255
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   11033
         DefaultCols     =   ""
         GridRows        =   21
         BackColor       =   12648447
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         Editable        =   0
         HeadFontCharset =   134
         HeadFontWeight  =   400
         DataFontCharset =   134
         DataFontWeight  =   400
         ExtendLastCol   =   -1  'True
      End
      Begin VB.Label labRecordInf 
         Caption         =   "��ǰ����Ƭ����0    ��ǰ����Ƭ����0"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   6840
         Width           =   4695
      End
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10665
      _ExtentX        =   18812
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
            Caption         =   "��Ƭ����"
            Key             =   "tbAcceptSlices"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��Ƭ���"
            Key             =   "tbEndSlices"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPatholSlices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IWorkMenu

Private Const M_STR_HINT_NoSelectData As String = "��Ч�ļ�����ݣ���ѡ����Ҫִ�еļ���¼��"
Private Const M_STR_MODULE_MENU_TAG As String = "��Ƭ"


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

Private mrecStudy As TStudyStateInf
Private mblnReadOnly As Boolean

Private mlngTmpAdviceId As Long
Private mlngTmpSendNo As Long

Private mblnAutoAcceptOfAfterPrint As Boolean
Private mbytFontSize As Byte '�ֺ�    9--С����    12--������


Private mObjActiveMenuBar As CommandBar

Private mblnRefreshState As Boolean


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

    If Not HasMenu(objMenuBar, conMenu_PatholSlices) Then
        Set cbrMenuBar = mObjActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_PatholSlices, "��Ƭ(&L)", GetPatholMenuIndex(objMenuBar) + 1, False)
        cbrMenuBar.ID = conMenu_PatholSlices
        cbrMenuBar.Category = M_STR_MODULE_MENU_TAG
                
        With cbrMenuBar.CommandBar
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_PatholSlices_LAB, "��ǩ��ӡ(&B)", "", 1, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_PatholSlices_PreviewLAB, "Ԥ��(V)", "", 1, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_PatholSlices_PrintLAB, "��ӡ(P)", "", 1, False)
                
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_PatholSlices_List, "�嵥��ӡ(&T)", "", 1, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_PatholSlices_PreviewList, "Ԥ��(V)", "", 1, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_PatholSlices_PrintList, "��ӡ(P)", "", 1, False)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholSlices_RequestView, "����鿴(&Q)", "", 1, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholSlices_Accept, "��Ƭ����(&R)", "", 1, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholSlices_Finish, "��Ƭ���(&F)", "", 1, False)
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
        Case conMenu_PatholSlices_PreviewLAB
            Call PrintSlicesLabel(False)
            
        Case conMenu_PatholSlices_PrintLAB
            Call PrintSlicesLabel(True)
            
        Case conMenu_PatholSlices_PreviewList
            Call PrintWorkList(False)
            
        Case conMenu_PatholSlices_PrintList
            Call PrintWorkList(True)
            
        Case conMenu_PatholSlices_RequestView
            Call ShowSlicesRequest
            
        Case conMenu_PatholSlices_Accept
            Call Slices_Accept
            
        Case conMenu_PatholSlices_Finish
            Call Slices_Sure
    End Select
End Sub


Public Sub IWorkMenu_zlUpdateMenu(ByVal control As XtremeCommandBars.ICommandBarControl)
'���²˵�
    Dim blnIsAllowSlices As Boolean
    
    If Not Me.Visible Then
        control.Enabled = False
        Exit Sub
    End If
    
    blnIsAllowSlices = CheckPopedom(mstrPrivs, "������Ƭ") And mlngAdviceID > 0
    
    Select Case control.ID
        Case conMenu_PatholSlices_LAB
            control.Enabled = blnIsAllowSlices And mrecStudy.strPatholNumber <> ""
            
        Case conMenu_PatholSlices_List
            control.Enabled = blnIsAllowSlices And mrecStudy.strPatholNumber <> ""
            
        Case conMenu_PatholSlices_RequestView
            control.Enabled = blnIsAllowSlices And mrecStudy.strPatholNumber <> ""
            
        Case conMenu_PatholSlices_Accept
            control.Enabled = blnIsAllowSlices And Not mblnReadOnly
            
        Case conMenu_PatholSlices_Finish
            control.Enabled = blnIsAllowSlices And Not mblnReadOnly
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
    If blnMoved Or lngStudyState = 6 Or lngStudyState = 5 Or lngStudyState = 0 Or lngStudyState = 1 Or lngStudyState = -2 Then
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
        Call ConfigSlicesFace(False, "ҽ��ID��Ч���顣")
        Exit Sub
    End If
    
    Call GetPatholStudyState(mlngAdviceID, mrecStudy)
    
    If Trim(mrecStudy.strPatholNumber) = "" Then
        Call ConfigSlicesFace(False, "�ü����δ������Ч�Ĳ���ţ���ȷ�ϸü���Ƿ��ѱ����ա�")
        
        Exit Sub
    Else
        Call ConfigSlicesFace(True)
    End If

    
    '��ȡ��Ƭ����
    Call LoadSlicesData
    
    'ˢ�²Ŀ�����
    Call RefreshSilcesCount
    
    Call ConfigPopedom(mblnReadOnly)
    
'    If Not (owner Is Nothing) Then
'        Call Me.Show(1, owner)
'    End If
End Sub


Public Sub zlRefresh(ByVal lngAdviceID As Long, ByVal blnReadOnly As Boolean, _
    ByVal strPrivs As String, ByVal blnMoved As Boolean, _
    ByVal lngCurDepartmentId As Long, Optional owner As Form = Nothing)
'ˢ��ȡ��ģ��
'���ͬʱ��ȡ�Ĺ��ܣ������ȡ�ļ�¼����Ƭ��Ҫˢ��
'    If lngAdviceID = mlngCurAdviceId Then  Exit Sub
        
    If lngAdviceID <= 0 Then
        Call ConfigSlicesFace(False, "ҽ��ID��Ч���顣")
        Exit Sub
    End If
    
    mlngAdviceID = lngAdviceID              'ҽ��ID
    mstrPrivs = strPrivs                    'ִ��Ȩ��
    mblnMoved = blnMoved                    '�Ƿ�ת��
    mlngCurDeptId = lngCurDepartmentId      '���ű��
    
   

    Call GetPatholStudyState(mlngAdviceID, mrecStudy)
    
    If Trim(mrecStudy.strPatholNumber) = "" Then
        Call ConfigSlicesFace(False, "�ü����δ������Ч�Ĳ���ţ���ȷ�ϸü���Ƿ��ѱ����ա�")
        
        Exit Sub
    Else
        Call ConfigSlicesFace(True)
    End If

    
    '��ȡ��Ƭ����
    Call LoadSlicesData
    
    'ˢ�²Ŀ�����
    Call RefreshSilcesCount
    
    Call ConfigPopedom(blnReadOnly)
    
    If Not (owner Is Nothing) Then
        Call Me.Show(1, owner)
    End If
End Sub


Private Sub RefreshSilcesCount()
'ˢ����Ƭ��¼����
    Dim i As Long
    Dim lngRecord As Long
    Dim lngTotal As Long
    Dim lngSlices As Long
    
    lngTotal = 0
    lngSlices = 0
    
    
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsNullRow(i) Then
            
            lngTotal = lngTotal + Val(ufgData.Text(i, gstrSlices_��Ƭ��))
            
            If ufgData.Text(i, gstrSlices_��ǰ״̬) <> "�����" Then
                lngSlices = lngSlices + Val(ufgData.Text(i, gstrSlices_��Ƭ��))
            End If
        End If
    Next i
    
    labRecordInf.Caption = "��ǰ����Ƭ����" & lngTotal & "    ��ǰ����Ƭ����" & lngSlices
    
End Sub

Private Sub ConfigPopedom(ByVal blnIsReadOnly As Boolean)
'����Ȩ��
    Dim blnIsAllowSlices As Boolean
    
    blnIsAllowSlices = CheckPopedom(mstrPrivs, "������Ƭ")
    
    tbrMain.Buttons("tbAcceptSlices").Enabled = blnIsAllowSlices And Not blnIsReadOnly
    tbrMain.Buttons("tbEndSlices").Enabled = blnIsAllowSlices And Not blnIsReadOnly

    
    tbrMain.Buttons("tbLAB").Enabled = blnIsAllowSlices
    tbrMain.Buttons("tbList").Enabled = blnIsAllowSlices
    tbrMain.Buttons("tbViewRequest").Enabled = blnIsAllowSlices
    
    ufgData.ReadOnly = blnIsReadOnly
End Sub


Private Sub AdjustFace()
    '�������沼��
    framSlices.Left = 0
    framSlices.Top = tbrMain.Top + tbrMain.Height + 120
    framSlices.Width = Me.Width - 0
    framSlices.Height = Me.Height - tbrMain.Height - 240
    
    ufgData.Left = 120
    ufgData.Top = 240 + IIf(mbytFontSize = 9, 0, 120)
    ufgData.Width = framSlices.Width - 240
    ufgData.Height = framSlices.Height - labRecordInf.Height - 480
    
    labRecordInf.Left = 120
    labRecordInf.Top = framSlices.Height - labRecordInf.Height - 120 + IIf(mbytFontSize = 9, 0, 85)

    
    '����FrameCheckλ��
     FramCheck.Top = framSlices.Height - labRecordInf.Height - 120 + IIf(mbytFontSize = 9, 0, 70)
     FramCheck.Left = framSlices.Width - FramCheck.Width - 200
     
     chkWCL.Top = 0
     chkYJS.Top = 0
     chkYWC.Top = 0
End Sub

Public Sub ReSetFormFontSize(ByVal bytFontSize As Byte)
'����:�������ù���վ����������С
On Error GoTo errHandle

    Dim objCtrl As control
    Dim CtlFont As StdFont
    Dim strFontType  As String
    
    '���ƶ��ؼ�λ��
    mbytFontSize = bytFontSize
    
    '����������
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



Private Sub LoadSlicesData()
'��ȡ��Ƭ��Ϣ
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select a.ID,a.�Ŀ�ID,b.���,b.ȡ��λ��, b.�걾����,a.��Ƭ��,a.��Ƭ����, a.��Ƭ��ʽ,a.��Ƭʱ��,a.��Ƭ�� as ��Ƭ��ʦ,a.��ǰ״̬,a.�嵥״̬" & _
            " from ������Ƭ��Ϣ a, ����ȡ����Ϣ b " & _
            " where a.�Ŀ�id=b.�Ŀ�id and b.ȷ��״̬=1 and b.����ҽ��ID = [1] order by a.��ǰ״̬,b.���,a.ID"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mrecStudy.lngPatholAdviceId)
    
    Call FilterData

End Sub

Private Sub FilterData()
'��������
     Dim strFilter As String
    
    '�жϵ�ǰ״̬�����ݸ�ѡ����ʾ����
    If chkWCL.value <> 0 Then
        If strFilter = "" Then
            strFilter = "��ǰ״̬=0"
        Else
             strFilter = strFilter & " or " & "��ǰ״̬=0"
        End If
        
    End If
    
    If chkYJS.value <> 0 Then
        If strFilter = "" Then
            strFilter = "��ǰ״̬=1"
        Else
             strFilter = strFilter & " or " & "��ǰ״̬=1"
        End If
    End If
    
    If chkYWC.value <> 0 Then
        If strFilter = "" Then
            strFilter = "��ǰ״̬=2"
        Else
             strFilter = strFilter & " or " & "��ǰ״̬=2"
        End If
    End If
    
     If strFilter = "" Then
            strFilter = "��ǰ״̬=9"
    End If
    
    ufgData.AdoData.Filter = strFilter
    'ˢ������
    Call ufgData.RefreshData

    Call RefreshSilcesCount
End Sub

Private Sub chkWCL_Click()
On Error Resume Next
    Call FilterData

End Sub

Private Sub chkYJS_Click()
On Error Resume Next
    Call FilterData

End Sub

Private Sub chkYWC_Click()
On Error Resume Next
    Call FilterData

End Sub



Private Sub InitSlicesList()
'��ʼ����Ƭ�б�
    Dim strTemp As String
    
    ufgData.IsKeepRows = True
    ufgData.GridRows = glngMaxRowCount
    '�����и�
    ufgData.RowHeightMin = glngStandardRowHeight
    
     '�ж����ݿ�������Ƿ������� �����ȡ���ݿ����  û�������Ĭ��
    strTemp = zlDatabase.GetPara("������Ƭ�б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
     
    If strTemp = "" Then
        ufgData.ColNames = gstrSlicesCols
    Else
        ufgData.ColNames = strTemp
    End If
    
    ufgData.DefaultColNames = gstrSlicesCols
    ufgData.ColConvertFormat = gstrSlicesConvertFormat
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
        Case UCase("tbLAB"), UCase("tbLABPreview")
            'Ԥ����ǩ
            Call PrintSlicesLabel(False)
            
        Case UCase("tbLABPrint")
            '��ӡ��ǩ
            Call PrintSlicesLabel(True)
        
        Case UCase("tbList"), UCase("tbListPreview")
            'Ԥ���嵥
            Call PrintWorkList(False)
            
        Case UCase("tbListPrint")
            '��ӡ�嵥
            Call PrintWorkList(True)
            
        Case UCase("tbAcceptSlices")
            '��Ƭ����
            Call Slices_Accept
            
        Case UCase("tbEndSlices")
            '��Ƭ���
            Call Slices_Sure
            
        Case UCase("tbViewRequest")
            '�鿴���뵥
            ShowSlicesRequest
    End Select
End Sub

Private Sub ufgData_OnColFormartChange()
'�رմ���ʱ�����б�����
    zlDatabase.SetPara "������Ƭ�б�����", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
End Sub




Private Sub ConfigSlicesFace(ByVal blnIsValid As Boolean, Optional ByVal strHintInf As String = "")
'�����ؼ����

    tbrMain.Buttons("tbAcceptSlices").Enabled = blnIsValid
    tbrMain.Buttons("tbEndSlices").Enabled = blnIsValid
    tbrMain.Buttons("tbLAB").Enabled = blnIsValid
    tbrMain.Buttons("tbList").Enabled = blnIsValid
    tbrMain.Buttons("tbViewRequest").Enabled = blnIsValid
    
    
    chkWCL.Enabled = blnIsValid
    chkYJS.Enabled = blnIsValid
    chkYWC.Enabled = blnIsValid
    
    If blnIsValid Then
        Call ufgData.CloseHintInf
    Else
        Call ufgData.ShowHintInf(strHintInf)
        labRecordInf.Caption = ""
    End If
End Sub


Private Sub Slices_Accept()
'��Ƭ����
    Dim strSql As String
    Dim i As Long
    
    '����Ƭ�׶Σ����ܽ��н���
    If mrecStudy.lngSlicesStep <> TExecuteStep.NeedDo And mrecStudy.lngSlicesStep <> TExecuteStep.AcceptDo Then
        Call MsgBoxD(Me, "��δ������Ƭ�׶Σ����ܽ�����Ƭ���ܲ�����", vbOKOnly, Me.Caption)
        Exit Sub
    End If
       
    
    If Not HasNeedState("δ����") Then
        Call MsgBoxD(Me, "û����Ҫ���н��ܵ���Ƭ��Ϣ����ȷ���Ƿ����δ�������Ƭ��Ϣ��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strSql = "Zl_������Ƭ_����('" & mrecStudy.lngPatholAdviceId & "','" & UserInfo.���� & "')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    mrecStudy.lngSlicesStep = 2
    
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsEmptyKey(i) Then
            If ufgData.Text(i, gstrSlices_��ǰ״̬) = "δ����" Then
                ufgData.Text(i, gstrSlices_��ǰ״̬) = "�ѽ���"
                ufgData.Text(i, gstrSlices_��Ƭ��) = UserInfo.����
            End If
        End If
    Next i
    
    Call MsgBoxD(Me, "�ѽ�����Ƭ��", vbOKOnly, Me.Caption)
End Sub


Private Function HasNeedState(ByVal strState As String) As Boolean
'�ж��Ƿ���Ҫ���к���
    Dim i As Long
    
    HasNeedState = False
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsEmptyKey(i) Then
            If ufgData.Text(i, gstrSlices_��ǰ״̬) = strState Then
                HasNeedState = True
                Exit Function
            End If
        End If
    Next i
End Function


Private Sub Slices_Sure()
'��Ƭȷ��
    Dim strSql As String
    Dim i As Long
    Dim j As Long
    Dim lngSlicesCount As Long
    Dim strTemp As String
    Dim rsData As ADODB.Recordset
    Dim dtServicesTime As Date
    
    '����Ƭ�׶Σ����ܽ���ȷ��
    If mrecStudy.lngSlicesStep <> TExecuteStep.NeedDo And mrecStudy.lngSlicesStep <> TExecuteStep.AcceptDo Then
        Call MsgBoxD(Me, "��δ������Ƭ�׶Σ����ܽ�����Ƭȷ�ϲ�����", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If Not HasNeedState("�ѽ���") Then
        Call MsgBoxD(Me, "û����Ҫ����ȷ�ϵ���Ƭ��Ϣ����ȷ���Ƿ�����ѱ����ܵ���Ƭ��Ϣ��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    dtServicesTime = zlDatabase.Currentdate
    
    strSql = "Zl_������Ƭ_ȷ��('" & mrecStudy.lngPatholAdviceId & "'," & To_Date(dtServicesTime) & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    mrecStudy.lngSlicesStep = 3
    
    For i = 1 To ufgData.GridRows - 1
    
        If Not ufgData.IsEmptyKey(i) Then
            If ufgData.Text(i, gstrSlices_��ǰ״̬) = "�ѽ���" Then
                ufgData.Text(i, gstrSlices_��ǰ״̬) = "�����"
                ufgData.Text(i, gstrSlices_��Ƭʱ��) = dtServicesTime
            End If
            
            If ufgData.Text(i, gstrSlices_��ǰ״̬) = "δ����" Then
                mrecStudy.lngSlicesStep = 1
            End If
        End If
        
    Next i
    
    '������Ƭȷ���¼�
    Call SendMsgToMainWindow(Me, wetSlicesSure, mlngAdviceID)
    
    Call MsgBoxD(Me, "��ȷ����Ƭ��", vbOKOnly, Me.Caption)
End Sub





Private Sub PrintSlicesLabel(Optional ByVal blnIsPrint As Boolean = True)
'��ӡ�ؼ���Ŀ��ǩ
    Dim i As Long
    Dim j As Long
    Dim strValue(5) As String
    
    Dim strSliceId As String
    Dim k As Long
    Dim lngCount As Long
    
    j = 0
    strValue(0) = "0": strValue(1) = "0": strValue(2) = "0": strValue(3) = "0": strValue(4) = "0": strValue(5) = "0"
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsEmptyKey(i) Then
            If zlCommFun.ActualLen(strValue(j)) > 2000 Then
                j = j + 1
                strValue(j) = ""
            End If
            
            strSliceId = ufgData.KeyValue(i)
            lngCount = Val(ufgData.Text(i, gstrSlices_��Ƭ��))
    
            If strValue(j) <> "" Then strValue(j) = strValue(j) & ","
    
            strValue(j) = strValue(j) & strSliceId
            
            If lngCount > 1 Then
                For k = 1 To lngCount - 1
                    strValue(j) = strValue(j) & "," & strSliceId
                Next k
            End If
            
        End If
    Next i
    
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_09", Me, "��ƬID1=" & strValue(0), "��ƬID2=" & strValue(1), "��ƬID3=" & strValue(2), "��ƬID4=" & strValue(3), "��ƬID5=" & strValue(4), "��ƬID6=" & strValue(5), IIf(blnIsPrint, 2, 1))
End Sub



Private Sub PrintSelectSlicesLabel(Optional ByVal blnIsPrint As Boolean = True)
'��ӡѡ��ĲĿ��ǩ
On Error GoTo errHandle
    Dim strValue(5) As String
    Dim strSliceId As String
    Dim i As Long
    Dim lngCount As Long
    
    
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ��ӡ����Ƭ��¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.IsNullRow(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ҫ��ӡ����Ƭ��¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strSliceId = ufgData.KeyValue(ufgData.SelectionRow)
    lngCount = Val(ufgData.Text(ufgData.SelectionRow, gstrSlices_��Ƭ��))
    
    strValue(0) = strSliceId
    If lngCount > 1 Then
    '����Ƭ������1ʱ���򴫵���ͬ������ID
        For i = 1 To lngCount - 1
            strValue(0) = strValue(0) & "," & strSliceId
        Next i
    End If
    
    strValue(1) = "0": strValue(2) = "0": strValue(3) = "0": strValue(4) = "0": strValue(5) = "0"

    
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_09", Me, "��ƬID1=" & strValue(0), "��ƬID2=" & strValue(1), "��ƬID3=" & strValue(2), "��ƬID4=" & strValue(3), "��ƬID5=" & strValue(4), "��ƬID6=" & strValue(5), IIf(blnIsPrint, 2, 1))
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub PrintWorkList(Optional ByVal blnIsPrint As Boolean = True)
'��ӡ��Ƭ�����б�
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
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_08", Me, "��ƬID1=" & strValue(0), "��ƬID2=" & strValue(1), "��ƬID3=" & strValue(2), "��ƬID4=" & strValue(3), "��ƬID5=" & strValue(4), "��ƬID6=" & strValue(5), IIf(blnIsPrint, 2, 1))
    
End Sub


Private Sub ShowSlicesRequest()
'��ʾ��Ƭ����
Dim frmRequestView As New frmPatholRequisition_View
On Error GoTo errFree
    Call frmRequestView.ShowRequestViewWind(mrecStudy.lngPatholAdviceId, 3, mblnMoved, Me)
errFree:
    Call Unload(frmRequestView)
    Set frmRequestView = Nothing
End Sub

Private Sub Form_Initialize()
    Set zlReport = New zl9Report.clsReport
    mblnAutoAcceptOfAfterPrint = False
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    '��ʼ����Ƭ��ʾ�б�
    Call InitSlicesList

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call AdjustFace
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    
    Set zlReport = Nothing
End Sub


Private Sub UpdateWorkListPrintState()
'�ڴ�ӡ�󣬸��¹����嵥�Ĵ�ӡ״̬
    Dim strSql As String
    Dim i As Long
        
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsEmptyKey(i) Then
            strSql = "Zl_������Ƭ_�嵥��ӡ(" & ufgData.KeyValue(i) & ")"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)

            Call ufgData.SyncText(i, gstrSlices_�嵥״̬, "�Ѵ�ӡ", True)
        End If
    Next i
End Sub


Private Sub ufgData_OnColsNameReSet()
On Error GoTo errHandle

    Call LoadSlicesData
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub zlReport_AfterPrint(ByVal ReportNum As String)
On Error GoTo errHandle
    '���������Ƭ�嵥��ӡ����ֱ���˳�
    If ReportNum <> "ZL1_PATHOLSLICES_01" Then Exit Sub
    
    Call UpdateWorkListPrintState
    
    If mblnAutoAcceptOfAfterPrint Then
    '��ӡ���Զ�����
        Call Slices_Accept
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

