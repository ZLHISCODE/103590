VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPatholSpecimen 
   Caption         =   "�걾����"
   ClientHeight    =   8805
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   10140
   Icon            =   "frmPatholSpecimen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   10140
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.ImageList imgTbrS 
      Left            =   1845
      Top             =   795
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecimen.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecimen.frx":0C7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecimen.frx":18F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecimen.frx":2562
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecimen.frx":31D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecimen.frx":3E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecimen.frx":4AB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecimen.frx":572A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecimen.frx":639C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholSpecimen.frx":700E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "��ʷ���ռ�¼"
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   5475
      Width           =   9855
      Begin RichTextLib.RichTextBox txtHistoryRecord 
         Height          =   1815
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   3201
         _Version        =   393217
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"frmPatholSpecimen.frx":7788
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�ͼ�걾��¼"
      Height          =   4455
      Left            =   30
      TabIndex        =   1
      Top             =   675
      Width           =   9855
      Begin VB.ListBox lstPartment 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2790
         ItemData        =   "frmPatholSpecimen.frx":7825
         Left            =   7320
         List            =   "frmPatholSpecimen.frx":7827
         TabIndex        =   7
         Top             =   1080
         Width           =   2415
      End
      Begin VB.ComboBox cbxSpecimentPartment 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   2415
      End
      Begin zl9PACSWork.ucFlexGrid ufgData 
         Height          =   3735
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   6588
         DefaultCols     =   ""
         GridRows        =   21
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         HeadFontCharset =   134
         HeadFontWeight  =   400
         DataFontCharset =   134
         DataFontWeight  =   400
         ExtendLastCol   =   -1  'True
      End
      Begin VB.Label labSpecimenName 
         Caption         =   "�걾��λ����ѡ��"
         Height          =   255
         Left            =   7440
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label labRecordInf 
         Caption         =   "�걾������0"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   4080
         Width           =   3375
      End
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   240
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPatholSpecimen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IWorkMenu


Private Const M_STR_HINT_NoSelectData As String = "��Ч�ļ�����ݣ���ѡ����Ҫִ�еļ���¼��"
Private Const M_STR_MODULE_MENU_TAG As String = "�걾"

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

Private mblnReadOnly As Boolean

Private mlngTmpAdviceId As Long
Private mlngTmpSendNo As Long


Private mrsSpecimenPartData As ADODB.Recordset


Private mObjActiveMenuBar As CommandBar

Private mblnRefreshState As Boolean
Private mbytFontSize As Byte '�ֺ�    9--С����    12--������
Private mstrFormats As String 'rtf��ʽ�����ڸı��ֺ�
Private mblLordingOrRefreshing As Boolean '�Ƿ����ڼ��ػ���ˢ��

Private mblnShowSentInfo As Boolean    '�Ƿ�������ʾ�ͼ���Ϣ



'��ȡ�˵��ӿڶ���
Property Get zlMenu() As IWorkMenu
    Set zlMenu = Me
End Property


Public Sub NotificationRefresh()
'֪ͨˢ��
    mblnRefreshState = False
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrorHand
    Select Case control.ID
        Case conMenu_PatholSpecimen_PreviewLab, conMenu_PatholSpecimen_LAB
            Call PrintSpecimenLabel(False)
            
        Case conMenu_PatholSpecimen_PrintLab
            Call PrintSpecimenLabel(True)
            
        Case conMenu_PatholSpecimen_PreviewAccept, conMenu_PatholSpecimen_ACP
            Call PrintAcceptNotification(False)
            
        Case conMenu_PatholSpecimen_PrintAccept
            Call PrintAcceptNotification(True)
        
        Case conMenu_PatholSpecimen_Get
            '�����Զ���ȡ��Ϣ����
            Call AutoGetSpecimenInf
            
        Case conMenu_PatholSpecimen_Del
            'ɾ���걾
            Call DelSelectionSpecimen
            
        Case conMenu_PatholSpecimen_Save
            '����걾
            Call SaveCurSpecimenInf
            
        Case conMenu_PatholSpecimen_Accept
            '�걾����
            Call SpecimenAccept
            
        Case conMenu_PatholSpecimen_Reject
            '�걾����
            Call SpecimenReject
            
        Case conMenu_PatholSpecimen_Cancel
            '�걾����
            Call CancelSelectionSpecimen
    End Select
    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrorHand
    Dim blnIsAllowAccept As Boolean
    
    blnIsAllowAccept = CheckPopedom(mstrPrivs, "�걾����") And mlngAdviceID > 0
    
    Select Case control.ID
        Case conMenu_PatholSpecimen_ACP, conMenu_PatholSpecimen_LAB, conMenu_PatholSpecimen_PreviewLab, _
        conMenu_PatholSpecimen_PreviewAccept, conMenu_PatholSpecimen_PrintLab, conMenu_PatholSpecimen_PrintAccept
            control.Enabled = blnIsAllowAccept
                   
        Case conMenu_PatholSpecimen_Del, conMenu_PatholSpecimen_Save, conMenu_PatholSpecimen_Accept, _
        conMenu_PatholSpecimen_Reject, conMenu_PatholSpecimen_Cancel

            control.Enabled = blnIsAllowAccept And Not mblnReadOnly
    End Select
    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Resume
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
    
    If Not HasMenu(objMenuBar, conMenu_PatholSpecimen) Then
        Set cbrMenuBar = mObjActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_PatholSpecimen, "�걾(&P)", GetPatholMenuIndex(objMenuBar) + 1, False)
        cbrMenuBar.ID = conMenu_PatholSpecimen
        cbrMenuBar.Category = M_STR_MODULE_MENU_TAG
        
        With cbrMenuBar.CommandBar
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_PatholSpecimen_LAB, "��ǩ��ӡ(&L)", "", 1, False)
            With cbrControl
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PatholSpecimen_PreviewLab, "Ԥ��(V)", "", 1, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PatholSpecimen_PrintLab, "��ӡ(P)", "", 1, False)
            End With
                
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_PatholSpecimen_ACP, "ƾ����ӡ(&A)", "", 1, False)
            With cbrControl
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PatholSpecimen_PreviewAccept, "Ԥ��(V)", "", 1, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PatholSpecimen_PrintAccept, "��ӡ(P)", "", 1, False)
            End With
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholSpecimen_Get, "�걾��ȡ(&G)", "", 1, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholSpecimen_Del, "�걾ɾ��(&D)", "", 1, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholSpecimen_Save, "�걾����(&S)", "", 1, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholSpecimen_Accept, "�걾����(&R)", "", 0, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholSpecimen_Reject, "�걾����(&J)", "", 0, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholSpecimen_Cancel, "�걾����(&H)", "", 0, True)
            
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
        Case conMenu_PatholSpecimen_PreviewLab      'Ԥ����ǩ
            Call PrintSpecimenLabel(False)
            
        Case conMenu_PatholSpecimen_PrintLab        '��ӡ��ǩ
            Call PrintSpecimenLabel(True)
            
        Case conMenu_PatholSpecimen_PreviewAccept    '���յ�Ԥ��
            Call PrintAcceptNotification(False)
            
        Case conMenu_PatholSpecimen_PrintAccept     '���յ���ӡ
            Call PrintAcceptNotification(True)
        
        Case conMenu_PatholSpecimen_Get             '�Ŀ���ȡ
            Call AutoGetSpecimenInf
            
        Case conMenu_PatholSpecimen_Del             'ɾ��ѡ��ı걾
            Call DelSelectionSpecimen
            
        Case conMenu_PatholSpecimen_Save            '���浱ǰ�걾��Ϣ
            Call SaveCurSpecimenInf
            
        Case conMenu_PatholSpecimen_Accept          '�걾����
            Call SpecimenAccept
            
        Case conMenu_PatholSpecimen_Reject          '�걾����
            Call SpecimenReject
            
        Case conMenu_PatholSpecimen_Cancel        '�걾����
            Call CancelSelectionSpecimen
    End Select
End Sub


Public Sub IWorkMenu_zlUpdateMenu(ByVal control As XtremeCommandBars.ICommandBarControl)
'���²˵�
    Dim blnIsAllowAccept As Boolean

    If Not Me.Visible Then
        control.Enabled = False
        Exit Sub
    End If
    
    blnIsAllowAccept = CheckPopedom(mstrPrivs, "�걾����") And mlngAdviceID > 0
    
    Select Case control.ID
        Case conMenu_PatholSpecimen_ACP, conMenu_PatholSpecimen_LAB, conMenu_PatholSpecimen_PreviewLab, _
        conMenu_PatholSpecimen_PreviewAccept, conMenu_PatholSpecimen_PrintLab, conMenu_PatholSpecimen_PrintAccept
            control.Enabled = blnIsAllowAccept
                   
        Case conMenu_PatholSpecimen_Del, conMenu_PatholSpecimen_Save, conMenu_PatholSpecimen_Accept, _
        conMenu_PatholSpecimen_Reject, conMenu_PatholSpecimen_Cancel
            control.Enabled = blnIsAllowAccept And Not mblnReadOnly
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
    If blnMoved Or lngStudyState = 6 Or lngStudyState = -2 Or lngStudyState = 5 Then
        mblnReadOnly = True
    End If

End Sub

Public Sub ReSetFormFontSize(ByVal bytFontSize As Byte)
'����:�������ù���վ����������С
On Error GoTo ErrHandle

    Dim objCtrl As control
    Dim CtlFont As StdFont
    Dim strFontType  As String
    
    
    mbytFontSize = bytFontSize
    
    
    Set CtlFont = New StdFont
    Me.FontSize = bytFontSize
    strFontType = IIf(IsUseClearType = True, "΢���ź�", "����")
    
    CtlFont.Name = strFontType
    CtlFont.Size = bytFontSize
    
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("Frame")
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
            objCtrl.FontSize = bytFontSize
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
        Case UCase("richtextbox")
        
            If bytFontSize = C_INT_FONTSISE_SMALL Then
                mstrFormats = "{\rtf1\ansi\ansicpg936\deff0\deflang1033\deflangfe2052{\fonttbl{\f0\fnil\fcharset134 \'cb\'ce\'cc\'e5;}}" & _
                        "{\colortbl ;\red255\green104\blue104;\red19\green164\blue251;}" & _
                        "{\*\generator Msftedit 5.41.21.2509;}\viewkind4\uc1\sl276\slmult1\lang2052\b\f0\fs18 "
            ElseIf bytFontSize = C_INT_FONTSISE_MEDIUM Then
                mstrFormats = "{\rtf1\ansi\ansicpg936\deff0\deflang1033\deflangfe2052{\fonttbl{\f0\fnil\fcharset134 \'cb\'ce\'cc\'e5;}}" & _
                        "{\colortbl ;\red255\green104\blue104;\red19\green164\blue251;}" & _
                        "{\*\generator Msftedit 5.41.21.2509;}\viewkind4\uc1\sl276\slmult1\lang2052\b\f0\fs24 "
            ElseIf bytFontSize = C_INT_FONTSISE_BIG Then
                mstrFormats = "{\rtf1\ansi\ansicpg936\deff0\deflang1033\deflangfe2052{\fonttbl{\f0\fnil\fcharset134 \'cb\'ce\'cc\'e5;}}" & _
                        "{\colortbl ;\red255\green104\blue104;\red19\green164\blue251;}" & _
                        "{\*\generator Msftedit 5.41.21.2509;}\viewkind4\uc1\sl276\slmult1\lang2052\b\f0\fs30 "
            End If
            
            txtHistoryRecord.Text = ""
            Call LoadSpecimenAcceptOrRejectHistoryData
        End Select
    Next
    
    Call AdjustFace
    
    Exit Sub
ErrHandle:
End Sub




Public Sub zlRefreshFace(Optional blnForceRefresh As Boolean = False)
'ˢ�½�������
    Dim lngNewAdviceId As Long

    If (mlngTmpAdviceId = mlngAdviceID And mlngTmpSendNo = mlngSendNo And mblnRefreshState) And Not blnForceRefresh Then Exit Sub
    
    lngNewAdviceId = mlngAdviceID
    mblnRefreshState = True
    
    If mlngTmpAdviceId <> mlngAdviceID And mlngTmpAdviceId > 0 Then
        '�ж�ȡ���Ƿ���Ҫ����ȷ��
        If IsNeedSaveSpecimen Then
            If MsgBoxD(Me, "��δ��¼��ı걾���б��棬�Ƿ���Ҫ���棿", vbYesNo, Me.Caption) = vbYes Then
                mlngAdviceID = mlngTmpAdviceId
                
                Call SaveCurSpecimenInf
            End If
        End If
    End If
        
    mlngAdviceID = lngNewAdviceId
    mlngTmpAdviceId = lngNewAdviceId
    mlngTmpSendNo = mlngSendNo
    
    If mlngAdviceID <= 0 Then
        Call ConfigSpecimenFace(False, "ҽ��ID��Ч���顣")
        Exit Sub
    Else
        Call ConfigSpecimenFace(True)
    End If
    
    mblLordingOrRefreshing = True
    '����걾����
    Call LoadSpecimenData
    
    
    '��ȡ�걾���ռ�¼
    txtHistoryRecord.Text = ""
    Call LoadSpecimenAcceptOrRejectHistoryData
    
    'ˢ�±걾����
    Call RefreshSpecimenCount
    
    Call ConfigPopedom(mblnReadOnly)
    mblLordingOrRefreshing = False
'    If Not (owner Is Nothing) Then
'        Call Me.Show(1, owner)
'    End If
    
End Sub

Public Sub zlRefresh(lngAdviceID As Long, ByVal blnReadOnly As Boolean, _
    strPrivs As String, ByVal blnMoved As Boolean, _
    ByVal lngCurDepartmentId As Long, Optional owner As Form = Nothing)
    
    If lngAdviceID <= 0 Then
        Call ConfigSpecimenFace(False, "ҽ��ID��Ч���顣")
        Exit Sub
    Else
        Call ConfigSpecimenFace(True)
    End If
    
    
    If lngAdviceID <> mlngAdviceID And mlngAdviceID > 0 Then
        '�ж�ȡ���Ƿ���Ҫ����ȷ��
        If IsNeedSaveSpecimen Then
            If MsgBoxD(Me, "��δ�Ա걾���б���������Ƿ���Ҫ���棿", vbYesNo, Me.Caption) = vbYes Then
                Call SaveCurSpecimenInf
            End If
        End If
    End If
    
    
'    If mlngCurAdviceId = lngAdviceID Then Exit Sub
    
    mlngAdviceID = lngAdviceID
    mstrPrivs = strPrivs
    mblnMoved = blnMoved
    mlngCurDeptId = lngCurDepartmentId
    
'    mlngStudyProcedure = GetStudyProcedure
    
    mblLordingOrRefreshing = True
    '����걾����
    Call LoadSpecimenData
    
    
    '��ȡ�걾���ռ�¼
    txtHistoryRecord.Text = ""
    Call LoadSpecimenAcceptOrRejectHistoryData
    
    'ˢ�±걾����
    Call RefreshSpecimenCount
    
    Call ConfigPopedom(blnReadOnly)
    mblLordingOrRefreshing = False
    If Not (owner Is Nothing) Then
        Call Me.Show(1, owner)
    End If
End Sub


Public Sub LoadSpecimenAcceptOrRejectHistoryData()
    Dim strSql As String
    Dim rsHistory As ADODB.Recordset
    Dim strRecord As String
    Dim lngStart As Long
    Dim strFormats As String
    
    strSql = "select �ͼ쵥λ,�ͼ����,�ͼ���,�ͼ�����,��ϵ��ʽ,�Ǽ���,����״̬,����ԭ��,֪ͨ��,��ע from �����ͼ���Ϣ where ҽ��ID=[1] and" _
               & " �ͼ�����<>to_date('1000/10/10 10:10:10','yyyy/mm/dd hh24:mi:ss')  order by �ͼ����� "
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set rsHistory = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
    
    If rsHistory.RecordCount <= 0 Then Exit Sub
               
    strFormats = mstrFormats
    txtHistoryRecord.Text = ""
    While Not rsHistory.EOF
        If Val(Nvl(rsHistory!����״̬)) = 1 Then
            strRecord = Nvl(rsHistory!�ͼ�����) & "����[ " & Nvl(rsHistory!�ͼ���) & " ]��[ " & Nvl(rsHistory!�ͼ쵥λ) & Nvl(rsHistory!�ͼ����) & " ]�ͼ�ı걾�ѱ�[ " & Nvl(rsHistory!�Ǽ���) & " ]���ա�"
            
            strFormats = strFormats & "\cf2 " & strRecord & "\par"
        Else
            strRecord = Nvl(rsHistory!�ͼ�����) & "����[ " & Nvl(rsHistory!�ͼ���) & " ]��[ " & Nvl(rsHistory!�ͼ쵥λ) & Nvl(rsHistory!�ͼ����) & " ]�ͼ�ı걾�ѱ�[ " & Nvl(rsHistory!�Ǽ���) & " ]���ա���֪ͨ[ " & Nvl(rsHistory!֪ͨ��) & " ] ��ϵ��ʽ[ " & Nvl(rsHistory!��ϵ��ʽ) & " ]"
            
            strFormats = strFormats & "\cf1 " & strRecord & "\par"
        End If
        
        rsHistory.MoveNext
    Wend
    
    txtHistoryRecord.SelRTF = strFormats & "}"
End Sub


Private Sub ConfigPopedom(ByVal blnIsReadOnly As Boolean)
'����Ȩ��
    Dim blnIsAllowAccept As Boolean
    
    blnIsAllowAccept = CheckPopedom(mstrPrivs, "�걾����")
    
    ufgData.ReadOnly = blnIsReadOnly
    
    
    lstPartment.Enabled = blnIsAllowAccept
    cbxSpecimentPartment.Enabled = blnIsAllowAccept
    
    If blnIsReadOnly Then
        cbxSpecimentPartment.BackColor = Me.BackColor
        lstPartment.BackColor = Me.BackColor
    Else
        cbxSpecimentPartment.BackColor = vbWhite
        lstPartment.BackColor = vbWhite
    End If

End Sub



Private Sub LoadSpecimenData()
'��ȡ���յı걾��Ϣ
    Dim strSql As String
    Dim rsSpecimen As ADODB.Recordset
    
    strSql = "select �걾ID,�ͼ�ID,�걾����,�걾����,�ɼ���λ,����,�������,���λ��,ԭ�б��,��������,��ע,case when nvl(�ͼ�ID,0)<=0 then 'δ����' else '�Ѻ���' end as ����״̬ " & _
             "from ����걾��Ϣ where ҽ��id=[1] order by �걾����,�������,��������,�걾ID"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
    
    Call ufgData.RefreshData
End Sub



Private Sub AdjustFace()
    '�������沼��
    Frame1.Left = 0
    If mbytFontSize = C_INT_FONTSISE_SMALL Then
        Frame1.Top = 800
    ElseIf mbytFontSize = C_INT_FONTSISE_MEDIUM Then
        Frame1.Top = 850
    Else
        Frame1.Top = 900
    End If
    Frame1.Width = Me.Width - 0
    Frame1.Height = Me.Height - Frame2.Height - 1000
    
    ufgData.Left = 120
    ufgData.Top = 240 + IIf(mbytFontSize = C_INT_FONTSISE_SMALL, 0, 120)
    ufgData.Width = Frame1.Width - lstPartment.Width - 360
    ufgData.Height = Frame1.Height - labRecordInf.Height - 480
    
    labRecordInf.Left = 120
    labRecordInf.Top = Frame1.Height - labRecordInf.Height - 120 + IIf(mbytFontSize = C_INT_FONTSISE_SMALL, 0, 85)


    
    '����frame2������
    Frame2.Left = 0
    Frame2.Top = Frame1.Top + Frame1.Height + 120
    Frame2.Width = Frame1.Width
    
    txtHistoryRecord.Left = 120
    txtHistoryRecord.Top = 240 + IIf(mbytFontSize = C_INT_FONTSISE_SMALL, 0, 120)
    txtHistoryRecord.Width = Frame2.Width - 240
    txtHistoryRecord.Height = Frame2.Height - 360 + IIf(mbytFontSize = C_INT_FONTSISE_SMALL, 0, -120)
    
    labSpecimenName.Left = ufgData.Left + ufgData.Width + 120
    labSpecimenName.Top = ufgData.Top + IIf(mbytFontSize = C_INT_FONTSISE_SMALL, 0, 120)

    cbxSpecimentPartment.Left = labSpecimenName.Left
    cbxSpecimentPartment.Top = labSpecimenName.Top + labSpecimenName.Height + 120

    lstPartment.Left = labSpecimenName.Left
    lstPartment.Top = cbxSpecimentPartment.Top + cbxSpecimentPartment.Height + 120
    lstPartment.Height = ufgData.Height - labSpecimenName.Height - cbxSpecimentPartment.Height - 240
    
    
End Sub




Private Sub RefreshSpecimenCount()
    'ˢ�±걾����
    Dim Count As Long
    Dim lngTotal As Long
    Dim i As Long
    
    lngTotal = 0
    
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.RowHidden(i) Then
            If Not ufgData.IsNullRow(i) Then
                lngTotal = lngTotal + Val(ufgData.Text(i, gSpecimen_����))
            End If
        End If
    Next i
    
    labRecordInf.Caption = "�걾������" & lngTotal
End Sub




Private Function IsNeedSaveSpecimen() As Boolean
'�Ƿ���Ҫȡ��ȷ��
    Dim i As Long
    
    IsNeedSaveSpecimen = False
    
    For i = 1 To ufgData.GridRows - 1
        If ufgData.RowState(i) = TDataRowState.Add And Not ufgData.RowHidden(i) Then
            IsNeedSaveSpecimen = True
            Exit For
        End If
    Next i
End Function



Private Sub ConfigSpecimenFace(ByVal blnIsValid As Boolean, Optional ByVal strHintInf As String = "")
'���ú��ս���
    
    lstPartment.Enabled = blnIsValid
    cbxSpecimentPartment.Enabled = blnIsValid
    
    If blnIsValid Then
        Call ufgData.CloseHintInf
        
        cbxSpecimentPartment.BackColor = Me.BackColor
        lstPartment.BackColor = Me.BackColor
    Else
        Call ufgData.ShowHintInf(strHintInf)
        labRecordInf.Caption = ""
        
        cbxSpecimentPartment.BackColor = vbWhite
        lstPartment.BackColor = vbWhite
    End If
End Sub



Private Sub cbxSpecimentPartment_Click()
On Error GoTo ErrHandle
    Dim i As Integer
    Dim strSql As String
    
    '���ListBox
    lstPartment.Clear
    
    If Trim(cbxSpecimentPartment.Text) <> "" Then
        mrsSpecimenPartData.Filter = "�걾��λ='" & cbxSpecimentPartment.Text & "'"
    Else
        mrsSpecimenPartData.Filter = ""
    End If
    
    
    While Not mrsSpecimenPartData.EOF
       '���ؾ�����걾����
        lstPartment.AddItem Nvl(mrsSpecimenPartData!�걾����)       '�Ƶ���һ������
        mrsSpecimenPartData.MoveNext
    Wend
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub AutoGetSpecimenInf()
'�Զ���ҽ������ȡ�걾��Ϣ
    Dim strSql As String
    Dim lngRow As Long
    Dim i As Long
    Dim blnFind As Boolean
    Dim objCheck As CheckState
    Dim strSpecimenType As String
    Dim lngImgIndex As Long
    Dim rsAdviceRecord As ADODB.Recordset
    
    '�Ѿ����յı걾���ܽ�����Ϣ��ȡ
    If Not Val(ufgData.Text(1, gSpecimen_�ͼ�ID)) <= 0 Then
        Call MsgBoxD(Me, "�걾�ѱ����գ����ܽ�����ȡ��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strSql = "select a.�걾��λ as �걾����,a.��鷽��,b.�걾��λ,b.�걾���� from ����ҽ����¼ a,������걾 b where a.�걾��λ=b.�걾����(+) and ���ID=[1]"
    Set rsAdviceRecord = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
    
    If rsAdviceRecord.RecordCount <= 0 Then Exit Sub
    
    Do While Not rsAdviceRecord.EOF
        
        For i = 1 To ufgData.GridRows - 1
            If (ufgData.Text(i, gstrMaterial_�걾����) = Nvl(rsAdviceRecord!�걾����)) And Not ufgData.RowHidden(i) Then GoTo continue
        Next i
        
        lngRow = ufgData.GetNullRowIndex
        
        '�����ֵ
        ufgData.Text(lngRow, gSpecimen_�걾����) = Nvl(rsAdviceRecord!�걾����)
        
        If Trim(ufgData.Text(lngRow, gSpecimen_�걾����)) = "" Then ufgData.Text(lngRow, gSpecimen_�걾����) = "0-�����걾"
        
        Call ufgData.GetFieldDisplayText(gSpecimen_�걾����, Val(Nvl(rsAdviceRecord!�걾����)), blnFind, objCheck, strSpecimenType, lngImgIndex)
        ufgData.Text(lngRow, gSpecimen_�걾����) = Val(Nvl(rsAdviceRecord!�걾����)) & "-" & strSpecimenType
        
        ufgData.Text(lngRow, gSpecimen_�ɼ���λ) = Nvl(rsAdviceRecord!�걾��λ)
        ufgData.Text(lngRow, gSpecimen_����) = 1
        ufgData.Text(lngRow, gSpecimen_�������) = Decode(Nvl(rsAdviceRecord!��鷽��), "�걾", 0, "����", 1, "��Ƭ", 2, "��Ƭ", 3, "����", 4) & "-" & Nvl(rsAdviceRecord!��鷽��)
          
          
        Call ufgData.DataGrid.Select(lngRow, ufgData.GetColIndex(gSpecimen_�걾����))
        Call ufgData.DataGrid.EditCell
        
continue:
        rsAdviceRecord.MoveNext
    Loop
    
End Sub


Private Sub CancelSelectionSpecimen()
'���˱걾

    Dim Row As Integer
    Dim ID As String
    Dim strSql As String
    
    If ufgData.ShowingRowCount <= 0 Then Exit Sub
    
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫɾ���ı걾��¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.IsNullRow(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ҫɾ���ı걾��¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Row = ufgData.SelectionRow
    ID = ufgData.Text(Row, gSpecimen_�걾ID)
    
    If ufgData.IsSelectionRow = False Then
        Call MsgBoxD(Me, "��ѡ����Ҫ���˵���Ŀ��")
        Exit Sub
    End If
    
    
    If CheckAllowUpdateSpecimen(ID) = False Then
        Call MsgBoxD(Me, "�ñ걾�����˶�Ӧ��ȡ�ļ�¼�����ܽ��л��ˣ�����ɾ����Ӧ��ȡ����Ϣ��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strSql = "Zl_����걾_�˻�(" & ID & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Call ufgData.SyncData(Row, gSpecimen_��������, Null, True)
    Call ufgData.SyncData(Row, gSpecimen_�ͼ�ID, Null, True)
    Call ufgData.SyncText(Row, gSpecimen_����״̬, "δ����", True)
                
End Sub


Private Sub DelSelectionSpecimen()
'ɾ��ѡ��ı걾
    If ufgData.ShowingRowCount <= 0 Then Exit Sub
    
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫɾ���ı걾��¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.IsNullRow(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ҫɾ���ı걾��¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '�Ѿ����յı걾���ܽ���ɾ����
    If Not Val(ufgData.Text(ufgData.SelectionRow, gSpecimen_�ͼ�ID)) <= 0 Then
        Call MsgBoxD(Me, "�걾�ѱ����ղ��ܽ���ɾ��,���Ƚ��л��˴���", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If MsgBoxD(Me, "ȷ��Ҫɾ��ѡ��ı걾������", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    
    'ɾ����
    Call ufgData.DelCurRow
    
    '����ɾ���ı걾���ݣ��ѱ����յı걾���ܽ���ɾ�������պ��ͼ�ID��Ϊ�գ�
    Call SaveSpecimenData(False, True)
    
    'ˢ�±걾����
    Call RefreshSpecimenCount
End Sub


'Private Sub cmdReload_Click()
'On Error GoTo errHandle
'    '�ָ��б�����
'    mclsVFGSpecimen.RestoreList
'
'    'ˢ�±걾����
'    Call RefreshSpecimenCount
'
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'End Sub



Private Function GetPatholNum(ByVal lngAdviceID As Long) As String
'��ȡ����ŵ������Ϣ
    Dim strSql As String
    Dim rsPatholNum As ADODB.Recordset
    
    
    GetPatholNum = ""
    strSql = "select ����� from ��������Ϣ where ҽ��id=[1]"
    
    Set rsPatholNum = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceID)
    
    If rsPatholNum.RecordCount <= 0 Then Exit Function
    
    GetPatholNum = Nvl(rsPatholNum!�����)
End Function


Private Function CheckNewSpecimenInf() As Boolean
'����Ƿ����µ���Ҫ���յı걾��Ϣ
    Dim i As Long
    
    CheckNewSpecimenInf = False
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsNullRow(i) And ufgData.IsEmptyKey(i) Then
            CheckNewSpecimenInf = True
            Exit Function
        End If
    Next i
End Function


Private Sub SpecimenAccept()
'�걾����
    Dim blnValid As Boolean
    
    '�걾����
    If ufgData.ShowingDataRowCount <= 0 Then
        Call MsgBoxD(Me, "û���ҵ���Ҫ���յı걾��Ϣ������걾�Ƿ���ȷ¼�롣", vbOKOnly, Me.Caption)
        Exit Sub
    End If

'    '�ж��Ƿ�����Ҫ���յı걾��Ϣ
'    If Not CheckNewSpecimenInf() Then
'        Call MsgBoxD(Me, "û���ҵ���Ҫ�ٴκ��յı걾��Ϣ������걾�Ƿ���ȷ¼�롣", vbOKOnly, Me.Caption)
'        Exit Sub
'    End If
    
    
    blnValid = Not ufgData.IsErrColorWithList
    If Not blnValid Then
        Call MsgBoxD(Me, "��⵽�걾�б������Ч���ݣ���ȷ���Ƿ���ȷ������¼�룬����ɫ����ǵĵ�Ԫ��Ϊ��¼���ݡ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If


    Dim blnIsSucceed As Boolean
    blnIsSucceed = frmPatholSpecimen_AcceptOrReject.ShowAcceptOrRejectSpecimenWindow(mlngAdviceID, _
                                    mlngCurDeptId, txtHistoryRecord, False, Me, mstrPrivs, mblnShowSentInfo)
    
    If blnIsSucceed Then

        '���º���״̬
        Call UpdateAcceptState
    
        '����ִ�к����¼�
        Call SendMsgToMainWindow(Me, wetSpecimenAccept, mlngAdviceID, GetPatholNum(mlngAdviceID))
        
        Call ufgData.SetMenuState(False)
    End If
    
'    'ˢ�±걾����
'    Call RefreshSpecimenCount

End Sub


Private Sub UpdateAcceptState()
'���º���״̬
    Dim i As Long
    
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsNullRow(i) Then
            ufgData.Text(i, gSpecimen_����״̬) = "�Ѻ���"
        End If
    Next i
End Sub


Private Sub PrintSpecimenLabel(Optional ByVal blnIsPrint As Boolean = True)
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
    
    
    '�ж� �Ƿ��ӡ����ֵ ���ֵΪ0 ���ʾû�й�ѡ �Ƿ�ֱ�Ӵ�ӡ����
    '��󸽼Ӳ���:0=ȱʡֵ,�ɲ���,��ʾ����(������Ԥ��),1=ֱ�ӵ�Ԥ��,2=ֱ�Ӵ�ӡ,3-�����Excel
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_12", Me, "�걾ID1=" & strValue(0), "�걾ID2=" & strValue(1), "�걾ID3=" & strValue(2), "�걾ID4=" & strValue(3), "�걾ID5=" & strValue(4), "�걾ID6=" & strValue(5), IIf(blnIsPrint, 2, 1))
End Sub


Private Sub PrintSelectSpecimenLabel(Optional ByVal blnIsPrint As Boolean = True)
'��ӡѡ��ĲĿ��ǩ
On Error GoTo ErrHandle
    Dim strValue(5) As String
    
    
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ��ӡ�ı걾��¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.IsNullRow(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ҫ��ӡ�ı걾��¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strValue(0) = ufgData.KeyValue(ufgData.SelectionRow)
    strValue(1) = "0": strValue(2) = "0": strValue(3) = "0": strValue(4) = "0": strValue(5) = "0"

    '��󸽼Ӳ���:0=ȱʡֵ,�ɲ���,��ʾ����(������Ԥ��),1=ֱ�ӵ�Ԥ��,2=ֱ�Ӵ�ӡ,3-�����Excel
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_12", Me, "�걾ID1=" & strValue(0), "�걾ID2=" & strValue(1), "�걾ID3=" & strValue(2), "�걾ID4=" & strValue(3), "�걾ID5=" & strValue(4), "�걾ID6=" & strValue(5), IIf(blnIsPrint, 2, 1))
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub PrintAcceptNotification(Optional ByVal blnIsPrint As Boolean = True)
'��ӡ�걾����֪ͨ��
    '��󸽼Ӳ���:0=ȱʡֵ,�ɲ���,��ʾ����(������Ԥ��),1=ֱ�ӵ�Ԥ��,2=ֱ�Ӵ�ӡ,3-�����Excel
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_13", Me, "ҽ��ID=" & mlngAdviceID, IIf(blnIsPrint, 2, 1))
End Sub



Private Sub SpecimenReject()
'���ձ걾
    Dim blnIsSucceed As Boolean
    blnIsSucceed = frmPatholSpecimen_AcceptOrReject.ShowAcceptOrRejectSpecimenWindow(mlngAdviceID, "", txtHistoryRecord, True, Me, mstrPrivs, mblnShowSentInfo)

    If blnIsSucceed Then
        '�������ִ���¼�...

    End If

End Sub



Public Sub SaveSpecimenData(ByVal blnSetFocus As Boolean, Optional ByVal blnIsSaveOnlyDel As Boolean = False)
'����걾����
'blnSetFocus:��������һ�ν��㣬����109548,���������ʱ����ΪTRUE
    Dim i As Long
    Dim strSql As String
    Dim rsReturn As ADODB.Recordset
    Dim lngSpecimenID As Long
    Dim dtServicesTime As String
    
    If blnSetFocus Then ufgData.SetFocus
    
    For i = 1 To ufgData.GridRows - 1
        If ufgData.RowState(i) = TDataRowState.Add And Not blnIsSaveOnlyDel Then
                dtServicesTime = zlDatabase.Currentdate
                
                '��ӱ걾����
                strSql = "select Zl_����걾_����([1],[2],[3],[4],[5],[6],[7],[8],[9],[10]) as ����ֵ from dual"
                
                    Set rsReturn = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                            mlngAdviceID, _
                            ufgData.Text(i, gSpecimen_�걾����), _
                            Val(ufgData.Text(i, gSpecimen_�걾����)), _
                            ufgData.Text(i, gSpecimen_�ɼ���λ), _
                            Val(ufgData.Text(i, gSpecimen_����)), _
                            Val(ufgData.Text(i, gSpecimen_�������)), _
                            ufgData.Text(i, gSpecimen_ԭ�б��), _
                            ufgData.Text(i, gSpecimen_���λ��), _
                            CDate(dtServicesTime), _
                            ufgData.Text(i, gSpecimen_��ע) _
                            )
                            
                    If rsReturn.RecordCount <= 0 Then
                        Call err.Raise(0, "SaveSpecimenData", "δ�ɹ���ȡ������ı걾ID,����ʧ�ܡ�")
                        Exit Sub
                    End If
                    
                    lngSpecimenID = rsReturn!����ֵ
                    
                    '���������ı걾ID
                    ufgData.Text(i, gSpecimen_�걾ID) = lngSpecimenID
                    ufgData.Text(i, gSpecimen_��������) = dtServicesTime
                    ufgData.Text(i, gSpecimen_����״̬) = "δ����"
                    
            ElseIf ufgData.RowState(i) = TDataRowState.Del Then
                'ɾ���걾�����걾�����պ󣬲�����ɾ��
                strSql = "Zl_����걾_ɾ��(" & Val(ufgData.KeyValue(i)) & ")"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            ElseIf ufgData.RowState(i) = TDataRowState.Update And Not blnIsSaveOnlyDel Then
                '���±걾
                lngSpecimenID = Val(ufgData.KeyValue(i))
        
                strSql = "Zl_����걾_����(" & lngSpecimenID & ",'" & _
                        ufgData.Text(i, gSpecimen_�걾����) & "'," & _
                        Val(ufgData.Text(i, gSpecimen_�걾����)) & ",'" & _
                        ufgData.Text(i, gSpecimen_�ɼ���λ) & "'," & _
                        Val(ufgData.Text(i, gSpecimen_����)) & "," & _
                        Val(ufgData.Text(i, gSpecimen_�������)) & ",'" & _
                        ufgData.Text(i, gSpecimen_ԭ�б��) & "','" & _
                        ufgData.Text(i, gSpecimen_���λ��) & "','" & _
                        ufgData.Text(i, gSpecimen_��ע) & "')"
        
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        End If
        
        '������״̬
        ufgData.RowState(i) = TDataRowState.Normal
    Next i
End Sub

Private Sub SaveCurSpecimenInf()
'���浱ǰ�걾��Ϣ
    Dim blnValid As Boolean
    
    '�걾����
    If ufgData.ShowingDataRowCount <= 0 Then
        Call MsgBoxD(Me, "û���ҵ���Ҫ����ı걾��Ϣ������걾�Ƿ���ȷ¼�롣", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    blnValid = Not ufgData.IsErrColorWithList
    If Not blnValid Then
        Call MsgBoxD(Me, "��⵽�걾�б��д�����Ч���ݣ���ȷ���Ƿ���ȷ������¼�룬����ɫ����ǵĵ�Ԫ��Ϊ��¼���ݡ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call SaveSpecimenData(True)
    
    Call SendMsgToMainWindow(Me, wetSpecimenSave, mlngAdviceID)
    
    Call MsgBoxD(Me, "�����ѳɹ����档", vbOKOnly, Me.Caption)
    
'    'ˢ�±걾����
'    Call RefreshSpecimenCount
End Sub


Private Sub Form_Initialize()
    Set zlReport = New zl9Report.clsReport
End Sub


Private Sub LoadSpecimenPart()
'���ر걾��鲿λ
    Dim i As Integer
    Dim strSql As String
    Dim rsSpecimenPart As ADODB.Recordset
    
    
    strSql = "select �걾����,����,�걾��λ,�걾���� from ������걾 order by �걾��λ,�걾����"
    Set mrsSpecimenPartData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    
    strSql = "select distinct �걾��λ from ������걾 order by �걾��λ"
    Set rsSpecimenPart = zlDatabase.OpenSQLRecord(strSql, Me.Caption)

    cbxSpecimentPartment.Clear
    Call cbxSpecimentPartment.AddItem("")
    
    While Not rsSpecimenPart.EOF
       '��ӱ걾��鲿λ
       cbxSpecimentPartment.AddItem Nvl(rsSpecimenPart!�걾��λ) '�Ƶ���һ������
       rsSpecimenPart.MoveNext
    Wend
    
    If cbxSpecimentPartment.ListCount > 0 Then cbxSpecimentPartment.ListIndex = 0
End Sub


Private Sub Form_Load()
On Error GoTo ErrHandle
   Dim strTemp As String
   
   '��������
    ufgData.GridRows = glngStandardRowCount
    '�����и�
    ufgData.RowHeightMin = glngStandardRowHeight
    ufgData.IsCopyMode = True
    
    Set mrsSpecimenPartData = Nothing
    '�ж����ݿ�������Ƿ������� �����ȡ���ݿ����  û�������Ĭ��
    strTemp = zlDatabase.GetPara("�걾�����б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
    mblnShowSentInfo = Val(zlDatabase.GetPara("¼����Ժ��Ϣ", glngSys, G_LNG_PATHOLSYS_NUM, 0)) '�Ƿ���ʾ�ͼ���Ϣ
    
    ufgData.DefaultColNames = gstrSpecimenCols

    If strTemp = "" Then
        ufgData.ColNames = gstrSpecimenCols
    Else
        ufgData.ColNames = strTemp
    End If
    
    ufgData.ColConvertFormat = gstrSpecimenConvertFormat
    
    Call InitCommandBars
    '���ر걾��鲿λ
    Call LoadSpecimenPart
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnColFormartChange()
'�����б�����
    zlDatabase.SetPara "�걾�����б�����", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Call AdjustFace
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Set zlReport = Nothing
    Set mrsSpecimenPartData = Nothing
End Sub


Private Sub lstPartment_DblClick()
On Error GoTo ErrHandle
    Dim strPartSuperadd As String
    Dim strSpeciPartName As String
    Dim strSpeciName As String
    Dim blnFind As Boolean
    Dim objCheck As CheckState
    Dim lngImgIndex As Long
    Dim strSpecimenType As String

    
    If Not CheckAllowUpdateSpecimen(Val(ufgData.KeyValue(ufgData.SelectionRow))) Then
        Call MsgBoxD(Me, "�ü���ѽ���ȡ�ģ����ܽ��б༭��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '�жϵ�ǰѡ���ǲ�����ͷ���� ���Զ�ѡ����һ�У����� ������
    If ufgData.SelectionRow = 0 Then
        Call ufgData.EditNextCell(1)
    End If
    
    strSpeciPartName = ufgData.Text(ufgData.SelectionRow, gSpecimen_�ɼ���λ)
    strSpeciName = ufgData.Text(ufgData.SelectionRow, gSpecimen_�걾����)
    
    mrsSpecimenPartData.Filter = "�걾����='" & lstPartment.Text & "'"
    
     '�ɼ���λ�ж�
    If strSpeciPartName <> "" And strSpeciPartName <> Nvl(mrsSpecimenPartData!�걾��λ) Then
        Call MsgBoxD(Me, "�ɼ���λ��һ�£����ܽ����޸ġ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    
    ufgData.Text(ufgData.SelectionRow, gSpecimen_�ɼ���λ) = Nvl(mrsSpecimenPartData!�걾��λ)
    
    Call ufgData.GetFieldDisplayText(gSpecimen_�걾����, Val(Nvl(mrsSpecimenPartData!�걾����)), blnFind, objCheck, strSpecimenType, lngImgIndex)
    ufgData.Text(ufgData.SelectionRow, gSpecimen_�걾����) = Val(Nvl(mrsSpecimenPartData!�걾����)) & "-" & strSpecimenType
    
    '���������ͬ�����ƣ������
    If strSpeciName Like "*" & lstPartment.Text & "*" Then
        Call ufgData_OnAfterEdit(ufgData.DataGrid.Row, ufgData.DataGrid.Col)
        
        Call ufgData.DataGrid.Select(ufgData.SelectionRow, ufgData.GetColIndex(gSpecimen_�걾����))
        Call ufgData.DataGrid.EditCell
        Exit Sub
    End If
    
    
    '�걾�����ж�,������׷�� ����������
    If strSpeciName <> "" Then
        strPartSuperadd = strSpeciName & "," & lstPartment.Text
    Else
        strPartSuperadd = lstPartment.Text
    End If
    
    ufgData.Text(ufgData.SelectionRow, gSpecimen_�걾����) = strPartSuperadd
    
    Call ufgData_OnAfterEdit(ufgData.DataGrid.Row, ufgData.DataGrid.Col)
    
    Call ufgData.DataGrid.Select(ufgData.SelectionRow, ufgData.GetColIndex(gSpecimen_�걾����))
    Call ufgData.DataGrid.EditCell
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnAfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim iCol As Long
    Dim i As Long
    Dim iCount As Long
    Dim strNewSpecimenName As String
    Dim lngCode As String
    Dim blnFind As Boolean
    Dim objCheck As CheckState
    Dim lngImgIndex As Long
    Dim strSpecimenType As String
    
    If ufgData.IsNullRow(Row) Then
        ufgData.RowState(Row) = TDataRowState.Normal
        Call ufgData.SetRowColor(Row, ufgData.BackColor)
        
        Exit Sub
    End If
    
    If Col = ufgData.GetColIndex(gSpecimen_�걾����) Then
    '���걾�����Ƿ��ظ�
        lngCode = ufgData.Text(Row, gSpecimen_�걾����)
        If lngCode <> "" Then
            mrsSpecimenPartData.Filter = "����='" & lngCode & "'"
            If mrsSpecimenPartData.RecordCount > 0 Then
                ufgData.Text(Row, gSpecimen_�걾����) = mrsSpecimenPartData!�걾����
                ufgData.Text(Row, gSpecimen_�ɼ���λ) = mrsSpecimenPartData!�걾��λ
                
                Call ufgData.GetFieldDisplayText(gSpecimen_�걾����, Val(Nvl(mrsSpecimenPartData!�걾����)), blnFind, objCheck, strSpecimenType, lngImgIndex)
                ufgData.Text(Row, gSpecimen_�걾����) = Val(Nvl(mrsSpecimenPartData!�걾����)) & "-" & strSpecimenType
            End If
        End If
    
        strNewSpecimenName = ufgData.CheckEquateValue(Row, Col)
        If strNewSpecimenName <> "" Then
            Call MsgBoxD(Me, "�걾���� [" & ufgData.Text(Row, gSpecimen_�걾����) & "]�Ѿ����ڡ�", vbOKOnly, Me.Caption)
            
            ufgData.Text(Row, gSpecimen_�걾����) = strNewSpecimenName
        End If
    End If
    
    '���δ¼��걾���ƣ�����ʾ����ɫ
    iCol = ufgData.GetColIndex(gSpecimen_�걾����)
    
    ufgData.CellColor(Row, iCol) = IIf(ufgData.Text(Row, gSpecimen_�걾����) = "", ufgData.ErrCellColor, ufgData.BackColor)
       
    
    '���δ¼��걾���ͣ�����ʾ����ɫ
    iCol = ufgData.GetColIndex(gSpecimen_�걾����)
    
    ufgData.CellColor(Row, iCol) = IIf(ufgData.Text(Row, gSpecimen_�걾����) = "", ufgData.ErrCellColor, ufgData.BackColor)
    
    
    '���δ¼��걾����������ʾ����ɫ
    iCol = ufgData.GetColIndex(gSpecimen_����)
    
    ufgData.CellColor(Row, iCol) = IIf(Val(ufgData.Text(Row, gSpecimen_����)) <= 0, ufgData.ErrCellColor, ufgData.BackColor)
    
    
    
    '���δ¼����ϣ�����ʾ����ɫ
    iCol = ufgData.GetColIndex(gSpecimen_�������)
    
    ufgData.CellColor(Row, iCol) = IIf(ufgData.Text(Row, gSpecimen_�������) = "", ufgData.ErrCellColor, ufgData.BackColor)
    
    
    '���걾�����ı�ʱ��ˢ�±걾��������ʾ
    If Col = ufgData.GetColIndex(gSpecimen_����) Then
        Call RefreshSpecimenCount
    End If
End Sub

'Private Function CheckIsMaterials(ByVal lngSpecimenID As Long) As Boolean
''����Ƿ����ȡ�Ĵ���
'
'    Dim strSql As String
'    Dim rsData As ADODB.Recordset
'
'    CheckIsMaterials = False
'
'    If lngSpecimenID <= 0 Then Exit Function
'
'    strSql = "select �Ŀ�ID from ����ȡ����Ϣ where  �걾ID=[1]"
'    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngSpecimenID)
'
'    If rsData.RecordCount > 0 Then CheckAllowUpdateSpecimen = False
'
'End Function


Private Function CheckAllowUpdateSpecimen(ByVal lngSpecimenID As Long) As Boolean
'����Ƿ��������
'δ��Ƭ�ĲĿ���ɽ��и���,ͨ����鲡����Ƭ��Ϣ�����жϲĿ��Ƿ�����Ƭ(�����ǰ״̬��Ϊ0��������Ƭ)

    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    CheckAllowUpdateSpecimen = True
    
    If lngSpecimenID <= 0 Then Exit Function
    
    strSql = "select �Ŀ�ID from ����ȡ����Ϣ where  �걾ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngSpecimenID)
    
    If rsData.RecordCount > 0 Then CheckAllowUpdateSpecimen = False
End Function




Private Sub ufgData_OnColsNameReSet()
On Error GoTo ErrHandle

    Call LoadSpecimenData
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnSelChange()
    If mblLordingOrRefreshing Then Exit Sub
    If ufgData.SelectionRow = 0 Then Exit Sub
    If Not CheckAllowUpdateSpecimen(Val(ufgData.KeyValue(ufgData.SelectionRow))) Then
        Call ufgData.SetMenuState(False)
    Else
        Call ufgData.SetMenuState(True)
    End If
End Sub


Private Sub ufgData_OnStartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'�������ڶ�ȡ�������е�����
    Dim dtServices As Date
    
    '�������̴���1ʱ��˵���ѽ���ȡ�Ĳ���
    If Not CheckAllowUpdateSpecimen(Val(ufgData.KeyValue(ufgData.SelectionRow))) Then
        Cancel = True
        Call MsgBoxD(Me, "�ü���ѽ���ȡ�ģ����ܽ��б༭��", vbOKOnly, Me.Caption)
    End If
        
    
    'If Col = ufgData.vfgHelper.GetColumnIndex(gSpecimen_����) And Row > 0 Then
        If Val(ufgData.Text(Row, gSpecimen_����)) <= 0 Then ufgData.Text(Row, gSpecimen_����) = "1"
    '    Exit Sub
    'End If
    
    
    'If Col = ufgData.vfgHelper.GetColumnIndex(gSpecimen_�걾����) And Row > 0 Then
        If Trim(ufgData.Text(Row, gSpecimen_�걾����)) = "" Then ufgData.Text(Row, gSpecimen_�걾����) = "0-�����걾"
    '    Exit Sub
    'End If
    
    'If Col = ufgData.vfgHelper.GetColumnIndex(gSpecimen_�������) And Row > 0 Then
        If Trim(ufgData.Text(Row, gSpecimen_�������)) = "" Then ufgData.Text(Row, gSpecimen_�������) = "0-�걾"
    '    Exit Sub
    'End If
End Sub

Private Sub InitCommandBars()
On Error GoTo errH
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim cbrPopControl As CommandBarControl
    Dim intTMP As Integer
    Dim cbrEdit As CommandBarEdit
                                
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    Set Me.cbrMain.Icons = zlCommFun.GetPubIcons
    
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.Visible = False
    
    '�ɼ�����������
    Set cbrToolBar = Me.cbrMain.Add("������", xtpBarTop)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = True
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls

        Set cbrControl = .Add(xtpControlSplitButtonPopup, conMenu_PatholSpecimen_LAB, "��ǩ��ӡ"): cbrControl.IconId = 5001: cbrControl.ToolTipText = "��ǩ��ӡ"
            With cbrControl
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PatholSpecimen_PreviewLab, "Ԥ��", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PatholSpecimen_PrintLab, "��ӡ", "", 0, False)
            End With
        Set cbrControl = .Add(xtpControlSplitButtonPopup, conMenu_PatholSpecimen_ACP, "ƾ����ӡ"): cbrControl.IconId = 5002: cbrControl.ToolTipText = "ƾ����ӡ"
            With cbrControl
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PatholSpecimen_PreviewAccept, "Ԥ��", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PatholSpecimen_PrintAccept, "��ӡ", "", 0, False)
            End With

        Set cbrControl = .Add(xtpControlButton, conMenu_PatholSpecimen_Get, "��ȡ�걾"): cbrControl.IconId = 5003: cbrControl.ToolTipText = "��ȡ�걾"
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholSpecimen_Del, "ɾ���걾"): cbrControl.IconId = 5004: cbrControl.ToolTipText = "ɾ���걾"
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholSpecimen_Save, "����걾"): cbrControl.IconId = 5005: cbrControl.ToolTipText = "����걾"
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholSpecimen_Accept, "���ձ걾"): cbrControl.IconId = 5006: cbrControl.ToolTipText = "���ձ걾"
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholSpecimen_Reject, "���ձ걾"): cbrControl.IconId = 5007: cbrControl.ToolTipText = "���ձ걾"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholSpecimen_Cancel, "�걾����"): cbrControl.IconId = 5019: cbrControl.ToolTipText = "�걾����"
        cbrControl.BeginGroup = True
        
        
        

    End With
    Exit Sub
errH:
End Sub
