VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{84865D89-6B2D-42E2-98C7-18F4206945F5}#2.0#0"; "zl9PacsControl.ocx"
Begin VB.Form frmPatholMaterials 
   Caption         =   "ȡ�ĵǼ�"
   ClientHeight    =   9405
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   10560
   Icon            =   "frmPatholMaterials.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9405
   ScaleWidth      =   10560
   StartUpPosition =   3  '����ȱʡ
   Begin zl9PacsControl.ucSplitter ucSplitter1 
      Height          =   135
      Left            =   0
      TabIndex        =   11
      Top             =   4320
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   238
      MousePointer    =   7
      SplitType       =   0
      SplitLevel      =   3
      StartDistance   =   840
      Control1Name    =   "picMaterial"
      Control2Name    =   "picPane1"
   End
   Begin VB.PictureBox picMaterial 
      BorderStyle     =   0  'None
      Height          =   3480
      Left            =   0
      ScaleHeight     =   3480
      ScaleWidth      =   10560
      TabIndex        =   12
      Top             =   840
      Width           =   10560
      Begin VB.Frame framMaterial 
         Caption         =   "ȡ�ļ�¼"
         Height          =   3495
         Left            =   105
         TabIndex        =   13
         Top             =   15
         Width           =   10080
         Begin VB.CommandButton cmdAutoInputMaterials 
            Caption         =   "¼��ȡ����Ϣ(&W)"
            Height          =   400
            Left            =   9600
            TabIndex        =   20
            ToolTipText     =   "��ȡ�ļ�¼��Ϣ¼�뵽�޼�������"
            Top             =   2880
            Width           =   1575
         End
         Begin VB.TextBox txtPos 
            Height          =   375
            Left            =   3000
            TabIndex        =   15
            ToolTipText     =   "����������ȡ�ĺ�ʣ��걾����ŵ�λ�á�"
            Top             =   2880
            Width           =   2535
         End
         Begin VB.ComboBox cbxSpecimenProcess 
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
            ItemData        =   "frmPatholMaterials.frx":000C
            Left            =   7320
            List            =   "frmPatholMaterials.frx":0022
            TabIndex        =   14
            Text            =   "���汣��"
            Top             =   2880
            Width           =   2295
         End
         Begin zl9PACSWork.ucFlexGrid ufgData 
            Height          =   2535
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   4471
            DefaultCols     =   ""
            GridRows        =   51
            IsCopyAdoMode   =   0   'False
            IsEjectConfig   =   -1  'True
            HeadFontCharset =   134
            HeadFontWeight  =   400
            DataFontCharset =   134
            DataFontWeight  =   400
            ExtendLastCol   =   -1  'True
         End
         Begin VB.Label labInf 
            Caption         =   "ʣ����λ�ã�"
            Height          =   255
            Left            =   1680
            TabIndex        =   19
            Top             =   3000
            Width           =   1695
         End
         Begin VB.Label labRecordInf 
            Caption         =   "�Ŀ�������0"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   120
            TabIndex        =   18
            Top             =   3000
            Width           =   1335
         End
         Begin VB.Label labSpecimenProcess 
            Caption         =   "�걾��������"
            Height          =   255
            Left            =   6000
            TabIndex        =   17
            Top             =   3000
            Width           =   1815
         End
      End
   End
   Begin MSComctlLib.ImageList imgTbrS 
      Left            =   9885
      Top             =   540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholMaterials.frx":005E
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholMaterials.frx":0CD0
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholMaterials.frx":1942
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholMaterials.frx":25B4
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholMaterials.frx":3226
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholMaterials.frx":3E98
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholMaterials.frx":4B0A
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholMaterials.frx":577C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picPane1 
      BorderStyle     =   0  'None
      Height          =   4950
      Left            =   0
      ScaleHeight     =   4950
      ScaleWidth      =   10560
      TabIndex        =   6
      Top             =   4455
      Width           =   10560
      Begin VB.Frame framWordEdit 
         Height          =   3135
         Left            =   3255
         TabIndex        =   9
         Top             =   -450
         Width           =   9855
         Begin zl9PACSWork.WordInputText wtDescription 
            Height          =   2895
            Left            =   600
            TabIndex        =   0
            Top             =   0
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   5106
            DepartId        =   0
         End
      End
      Begin TabDlg.SSTab tsFilter 
         Height          =   330
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   9885
         _ExtentX        =   17436
         _ExtentY        =   582
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   4
         TabHeight       =   520
         TabMaxWidth     =   2822
         WordWrap        =   0   'False
         TabCaption(0)   =   "�޼�����(&D)"
         TabPicture(0)   =   "frmPatholMaterials.frx":63EE
         Tab(0).ControlEnabled=   0   'False
         Tab(0).ControlCount=   0
         TabCaption(1)   =   "�Ѹƹ���(&C)"
         TabPicture(1)   =   "frmPatholMaterials.frx":640A
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).ControlCount=   0
      End
      Begin VB.Frame framDecalin 
         Height          =   3135
         Left            =   15
         TabIndex        =   8
         Top             =   480
         Width           =   9855
         Begin VB.CommandButton cmdChange 
            Caption         =   "�� ��(&G)"
            Height          =   400
            Left            =   5880
            TabIndex        =   3
            Top             =   2520
            Width           =   1215
         End
         Begin VB.CommandButton cmdDecalin 
            Caption         =   "�� ��(&T)"
            Height          =   400
            Left            =   4560
            TabIndex        =   2
            Top             =   2520
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "�� ��(&R)"
            Height          =   400
            Left            =   7200
            TabIndex        =   4
            Top             =   2520
            Width           =   1215
         End
         Begin VB.CommandButton cmdSucceed 
            Caption         =   "�� ��(&O)"
            Height          =   400
            Left            =   8520
            TabIndex        =   5
            Top             =   2520
            Width           =   1215
         End
         Begin zl9PACSWork.ucFlexGrid ufgDecalin 
            Height          =   2535
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   4471
            DefaultCols     =   ""
            GridRows        =   21
            IsCopyAdoMode   =   0   'False
            IsEjectConfig   =   -1  'True
            HeadFontCharset =   134
            HeadFontWeight  =   400
            DataFontCharset =   134
            DataFontWeight  =   400
         End
      End
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   1270
      ButtonWidth     =   1455
      ButtonHeight    =   1270
      Style           =   1
      ImageList       =   "imgTbrS"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��ǩ��ӡ"
            Key             =   "tbLAB"
            ImageIndex      =   7
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
            Caption         =   "-"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�鿴����"
            Key             =   "tbViewRequest"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��ȡ�Ŀ�"
            Key             =   "tbGetMaterials"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ɾ���Ŀ�"
            Key             =   "tbDelMaterials"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����Ŀ�"
            Key             =   "tbSaveMaterials"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ȷ��ȡ��"
            Key             =   "tbSureMaterials"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPatholMaterials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IWorkMenu

Private Const mMustColor As Long = &HC0C0FF

Private Const M_STR_HINT_NoSelectData As String = "��Ч�ļ�����ݣ���ѡ����Ҫִ�еļ���¼��"
Private Const M_STR_MODULE_MENU_TAG As String = "ȡ��"


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

Private mStrTemp As String

Private mlngTmpAdviceId As Long
Private mlngTmpSendNo As Long

Private mrecStudy As TStudyStateInf
Attribute mrecStudy.VB_VarHelpID = -1
Private mblnReadOnly As Boolean

Private mObjActiveMenuBar As CommandBar

Private mblnRefreshState As Boolean
Private mbytFontSize As Byte '�ֺ�    9--С����    12--������



'��ȡ�˵��ӿڶ���
Property Get zlMenu() As IWorkMenu
    Set zlMenu = Me
End Property


Public Sub NotificationRefresh()
'֪ͨˢ��
    mblnRefreshState = False
End Sub




Private Sub cmdAutoInputMaterials_Click()
'�ھ޼������п���¼��ȡ�ļ�¼��Ϣ
    Dim i As Integer, j As Integer
    Dim strTemp As String, strMaterials As String
On Error GoTo errHandle
    
    For i = 1 To ufgData.DataGrid.Rows - 1
        If Not ufgData.IsNullRow(i) Then
            For j = 0 To ufgData.DataGrid.Cols - 1
                If Not ufgData.DataGrid.ColHidden(j) And ufgData.Text(0, ufgData.GetColName(j)) <> "��" Then
                    strTemp = strTemp & ", " & ufgData.Text(0, ufgData.GetColName(j)) & ":" & ufgData.Text(i, ufgData.GetColName(j))
                End If
            Next
            
            If strTemp <> "" Then
                strMaterials = strMaterials & Mid(strTemp, 3) & vbCrLf
                strTemp = ""
            End If
        End If
    Next
    
    If strMaterials <> "" Then wtDescription.WordText = strMaterials & wtDescription.WordText
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Resize()
On Error GoTo errHandle
    Call ucSplitter1.RePaint(False)
Exit Sub
errHandle:
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
    
    If Not HasMenu(objMenuBar, conMenu_PatholMaterial) Then
        Set cbrMenuBar = mObjActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_PatholMaterial, "ȡ��(&I)", GetPatholMenuIndex(objMenuBar) + 1, False)
        cbrMenuBar.ID = conMenu_PatholMaterial
        cbrMenuBar.Category = M_STR_MODULE_MENU_TAG
        
        
        With cbrMenuBar.CommandBar
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholMaterial_PreviewAll, "��ǩԤ��(&V)", "", 0, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholMaterial_PrintAll, "��ǩ��ӡ(&P)", "", 0, False)
'            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholMaterial_PreviewSingle, "Ԥ��ѡ�б�ǩ(&E)", "", 0, False)
'            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholMaterial_PrintSingle, "��ӡѡ�б�ǩ(&I)", "", 0, False)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholMaterial_RequestView, "����鿴(&R)", "", 0, True)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholMaterial_Get, "�Ŀ���ȡ(&G)", "", 0, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholMaterial_Del, "�Ŀ�ɾ��(&D)", "", 0, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholMaterial_Save, "�Ŀ鱣��(&S)", "", 0, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholMaterial_Sure, "ȷ��ȡ��(&U)", "", 0, False)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholMaterial_Decalcification, "�Ѹ�(&F)", "", 0, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholMaterial_ChangeVat, "����(&A)", "", 0, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholMaterial_CancelVat, "����(&C)", "", 0, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PahtolMaterial_Finish, "���(&F)", "", 0, False)
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
        Case conMenu_PatholMaterial_PreviewAll      'Ԥ������ȡ�ı�ǩ
            Call PrintMaterialLabel(False)
            
        Case conMenu_PatholMaterial_PrintAll        '��ӡ����ȡ�ı�ǩ
            Call PrintMaterialLabel(True)
            
        Case conMenu_PatholMaterial_PreviewSingle   'Ԥ��ѡ�б�ǩ
            Call PrintSelectMaterialLabel(False)
            
        Case conMenu_PatholMaterial_PrintSingle     '��ӡѡ�б�ǩ
            Call PrintSelectMaterialLabel(True)
            
        Case conMenu_PatholMaterial_RequestView     '����鿴
            Call ShowMaterialRequest
            
        Case conMenu_PatholMaterial_Get             '�Ŀ���ȡ
            Call MaterialGet
            
        Case conMenu_PatholMaterial_Del             'ɾ��ѡ�вĿ�
            Call DelSelectionMaterial
            
        Case conMenu_PatholMaterial_Save            '���浱ǰȡ����Ϣ
            Call SaveCurMaterialInf
        
        Case conMenu_PatholMaterial_Sure            'ȷ��ȡ��
            Call SureCurMaterialInf
            
        Case conMenu_PatholMaterial_Decalcification '�Ѹ�
            Call Decalcification
            
        Case conMenu_PatholMaterial_ChangeVat       '����
            Call ChangeVat
            
        Case conMenu_PatholMaterial_CancelVat       '����
            Call CancelVat
            
        Case conMenu_PahtolMaterial_Finish          '���
            Call Finish
    End Select
End Sub


Public Sub IWorkMenu_zlUpdateMenu(ByVal control As XtremeCommandBars.ICommandBarControl)
'���²˵�
    Dim blnIsAllowMaterial As Boolean
    
    If Not Me.Visible Then
        control.Enabled = False
        Exit Sub
    End If
    
    blnIsAllowMaterial = CheckPopedom(mstrPrivs, "����ȡ��") And mlngAdviceID > 0
    
    Select Case control.ID
        Case conMenu_PatholMaterial_PreviewAll
            control.Enabled = blnIsAllowMaterial And mrecStudy.strPatholNumber <> ""
            
        Case conMenu_PatholMaterial_PrintAll
            control.Enabled = blnIsAllowMaterial And mrecStudy.strPatholNumber <> ""
            
        Case conMenu_PatholMaterial_PreviewSingle
            control.Enabled = blnIsAllowMaterial And mrecStudy.strPatholNumber <> ""
            
        Case conMenu_PatholMaterial_PrintSingle
            control.Enabled = blnIsAllowMaterial And mrecStudy.strPatholNumber <> ""
            
        Case conMenu_PatholMaterial_RequestView
            control.Enabled = blnIsAllowMaterial And mrecStudy.strPatholNumber <> ""
        
        Case conMenu_PatholMaterial_Get
            control.Enabled = blnIsAllowMaterial And Not mblnReadOnly
            
        Case conMenu_PatholMaterial_Del
            control.Enabled = blnIsAllowMaterial And Not mblnReadOnly
            
        Case conMenu_PatholMaterial_Save
            control.Enabled = blnIsAllowMaterial And Not mblnReadOnly
            
        Case conMenu_PatholMaterial_Sure
            control.Enabled = blnIsAllowMaterial And Not mblnReadOnly
            
        Case conMenu_PatholMaterial_Decalcification
            control.Enabled = blnIsAllowMaterial And Not mblnReadOnly
            
        Case conMenu_PatholMaterial_ChangeVat
            control.Enabled = blnIsAllowMaterial And Not mblnReadOnly
            
        Case conMenu_PatholMaterial_CancelVat
            control.Enabled = blnIsAllowMaterial And Not mblnReadOnly
            
        Case conMenu_PahtolMaterial_Finish
            control.Enabled = blnIsAllowMaterial And Not mblnReadOnly
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
    Dim lngNewAdviceId As Long
    
    
    If (mlngTmpAdviceId = mlngAdviceID And mlngTmpSendNo = mlngSendNo And mblnRefreshState) And Not blnForceRefresh Then Exit Sub
    
    lngNewAdviceId = mlngAdviceID
    mblnRefreshState = True
    
    If mlngTmpAdviceId <> mlngAdviceID And mlngTmpAdviceId > 0 Then
        '�ж�ȡ���Ƿ���Ҫ����ȷ��
        If IsNeedMaterialSure Then
            If MsgBoxD(Me, "��δ�Լ�� [" & mrecStudy.strPatholNumber & "] ����ȡ��ȷ�ϣ��Ƿ���ִ��ȷ�ϣ�", vbYesNo, Me.Caption) = vbYes Then
                mlngAdviceID = mlngTmpAdviceId
                
                Call ExecuteTbrOperation("tbSureMaterials")
            End If
        End If
    End If
    
    mlngAdviceID = lngNewAdviceId
    mlngTmpAdviceId = lngNewAdviceId
    mlngTmpSendNo = mlngSendNo
    
    If lngNewAdviceId <= 0 Then
        Call ConfigMaterialFace(False, "ҽ��ID��Ч���顣")
        Exit Sub
    Else
        Call ConfigMaterialFace(True)
    End If
    
    Call LoadReportModule
    
    Call GetPatholStudyState(mlngAdviceID, mrecStudy)
    
    
    If Trim(mrecStudy.strPatholNumber) = "" Then
        Call ConfigMaterialFace(False, "�ü����δ������Ч�Ĳ���ţ���ȷ�ϸü���Ƿ��ѱ����ա�")
        
'        If Not (mobjOwner Is Nothing) Then
'            Call MsgBoxD(Me, "�ü����δ������Ч�Ĳ���ţ���ȷ�ϸü���Ƿ��ѱ����ա�", vbOKOnly, Me.Caption)
'        End If
        
        Exit Sub
    Else
        Call ConfigMaterialFace(True)
    End If
    
    '�ж� ������˼������ �� "ϸ��" "����ʯ��"���������Ѹƽ���
    tsFilter.TabVisible(1) = IIf(mrecStudy.lngStudyType = 2 Or mrecStudy.lngStudyType = 5, False, True)
    
    '����ȡ���б�
    Call ConfigMaterialList(mrecStudy.lngStudyType)
    '�����Ѹ��б�
    Call ConfigDecalinList

    
    '��������
    Call ConfigGridInput(mrecStudy.lngStudyType)
    
        
    '��ȡ�Ŀ��¼
    Call LoadMaterialData(mrecStudy.lngPatholAdviceId)
    
    '��ȡ�ѸƼ�¼
    Call LoadDecalinData(mlngAdviceID)
    
    '��ȡ�޼�����
    Call LoadDescriptionInf(mrecStudy.lngPatholAdviceId)
    
    
    'ˢ�²Ŀ�����
    Call RefreshMaterialCount
    
    Call ConfigPopedom(mblnReadOnly)
    
'    If Not (owner Is Nothing) Then
'        Call Me.Show(1, owner)
'    End If
End Sub

Public Sub zlRefresh(ByVal lngAdviceID As Long, ByVal blnReadOnly As Boolean, _
    ByVal strPrivs As String, ByVal blnMoved As Boolean, _
    ByVal lngCurDepartmentId As Long, Optional owner As Form = Nothing)
'ˢ��ȡ��ģ��
    If lngAdviceID <= 0 Then
        Call ConfigMaterialFace(False, "ҽ��ID��Ч���顣")
        Exit Sub
    Else
        Call ConfigMaterialFace(True)
    End If
    
    If lngAdviceID <> mlngAdviceID And mlngAdviceID > 0 Then
        '�ж�ȡ���Ƿ���Ҫ����ȷ��
        If IsNeedMaterialSure Then
            If MsgBoxD(Me, "��δ�Լ�� [" & mrecStudy.strPatholNumber & "] ����ȡ��ȷ�ϣ��Ƿ���ִ��ȷ�ϣ�", vbYesNo, Me.Caption) = vbYes Then
                Call ExecuteTbrOperation("tbSureMaterials")
            End If
        End If
    End If
    
    
    mlngAdviceID = lngAdviceID
    mstrPrivs = strPrivs
    mblnMoved = blnMoved
    mlngCurDeptId = lngCurDepartmentId
    
    Call LoadReportModule
    
    Call GetPatholStudyState(lngAdviceID, mrecStudy)
    
    
    If Trim(mrecStudy.strPatholNumber) = "" Then
        Call ConfigMaterialFace(False, "�ü����δ������Ч�Ĳ���ţ���ȷ�ϸü���Ƿ��ѱ����ա�")
        
        If Not (owner Is Nothing) Then
            Call MsgBoxD(Me, "�ü����δ������Ч�Ĳ���ţ���ȷ�ϸü���Ƿ��ѱ����ա�", vbOKOnly, Me.Caption)
        End If
        
        Exit Sub
    Else
        Call ConfigMaterialFace(True)
    End If
    
    '�ж� ������˼������ �� "ϸ��" "����ʯ��"���������Ѹƽ���
    tsFilter.TabVisible(1) = IIf(mrecStudy.lngStudyType = 2 Or mrecStudy.lngStudyType = 5, False, True)
    
    '����ȡ���б�
    Call ConfigMaterialList(mrecStudy.lngStudyType)
    '�����Ѹ��б�
    Call ConfigDecalinList

    
    '��������
    Call ConfigGridInput(mrecStudy.lngStudyType)

    
    
    '��ȡ�Ŀ��¼
    Call LoadMaterialData(mrecStudy.lngPatholAdviceId)
    
    '��ȡ�ѸƼ�¼
    Call LoadDecalinData(mlngAdviceID)
    
    '��ȡ�޼�����
    Call LoadDescriptionInf(mrecStudy.lngPatholAdviceId)
    
    
    'ˢ�²Ŀ�����
    Call RefreshMaterialCount
    
    Call ConfigPopedom(blnReadOnly)
    
    If Not (owner Is Nothing) Then
        Call Me.Show(1, owner)
    End If
End Sub


Private Sub ConfigPopedom(ByVal blnIsReadOnly As Boolean)
'����Ȩ��
    Dim blnIsAllowMaterial As Boolean
    
    blnIsAllowMaterial = CheckPopedom(mstrPrivs, "����ȡ��")
    
    tbrMain.Buttons("tbGetMaterials").Enabled = blnIsAllowMaterial And Not blnIsReadOnly
    tbrMain.Buttons("tbDelMaterials").Enabled = blnIsAllowMaterial And Not blnIsReadOnly
    tbrMain.Buttons("tbSaveMaterials").Enabled = blnIsAllowMaterial And Not blnIsReadOnly
    tbrMain.Buttons("tbSureMaterials").Enabled = blnIsAllowMaterial And Not blnIsReadOnly
    tbrMain.Buttons("tbLAB").Enabled = blnIsAllowMaterial And Not blnIsReadOnly
    tbrMain.Buttons("tbViewRequest").Enabled = blnIsAllowMaterial And Not blnIsReadOnly
    
    cmdDecalin.Enabled = blnIsAllowMaterial And Not blnIsReadOnly
    cmdChange.Enabled = blnIsAllowMaterial And Not blnIsReadOnly
    cmdCancel.Enabled = blnIsAllowMaterial And Not blnIsReadOnly
    cmdSucceed.Enabled = blnIsAllowMaterial And Not blnIsReadOnly
    
    
    txtPos.Locked = blnIsReadOnly
    txtPos.BackColor = IIf(blnIsReadOnly, Me.BackColor, vbWhite)
    
    cbxSpecimenProcess.Enabled = blnIsAllowMaterial And Not blnIsReadOnly
    
    wtDescription.ReadOnly = blnIsReadOnly
    
    ufgData.ReadOnly = blnIsReadOnly
    ufgDecalin.ReadOnly = blnIsReadOnly
End Sub



Private Sub ConfigMaterialFace(ByVal blnIsValid As Boolean, Optional ByVal strHintInf As String = "")
'����ȡ�Ľ���
    tbrMain.Buttons("tbGetMaterials").Enabled = blnIsValid
    tbrMain.Buttons("tbDelMaterials").Enabled = blnIsValid
    tbrMain.Buttons("tbSaveMaterials").Enabled = blnIsValid
    tbrMain.Buttons("tbSureMaterials").Enabled = blnIsValid
    tbrMain.Buttons("tbLAB").Enabled = blnIsValid
    tbrMain.Buttons("tbViewRequest").Enabled = blnIsValid
    
    cmdDecalin.Enabled = blnIsValid
    cmdChange.Enabled = blnIsValid
    cmdCancel.Enabled = blnIsValid
    cmdSucceed.Enabled = blnIsValid
    
    txtPos.Enabled = blnIsValid
    txtPos.BackColor = IIf(Not blnIsValid, Me.BackColor, vbWhite)
    
    cbxSpecimenProcess.Enabled = blnIsValid
    cbxSpecimenProcess.BackColor = IIf(Not blnIsValid, Me.BackColor, vbWhite)
    
    If blnIsValid Then
        Call ufgData.CloseHintInf
        Call ufgDecalin.CloseHintInf
    Else
        Call ufgData.ShowHintInf(strHintInf)
        Call ufgDecalin.ShowHintInf(strHintInf)
        
        wtDescription.WordText = ""
        txtPos.Text = ""
    End If
End Sub


Private Function IsNeedMaterialSure() As Boolean
'�Ƿ���Ҫȡ��ȷ��
    Dim i As Long
    
    IsNeedMaterialSure = False
    
    For i = 1 To ufgData.GridRows - 1
        If ufgData.RowState(i) = TDataRowState.Add And Not ufgData.RowHidden(i) Then
            IsNeedMaterialSure = True
            Exit For
        End If
    Next i
End Function


Private Sub ConfigMaterialList(ByVal lngStudyType As Long)
'���òĿ���ʾ�б�
    Dim strTemp As String
    
    '��������
    ufgData.GridRows = glngStandardRowCount
    '�����и�
    ufgData.RowHeightMin = glngStandardRowHeight
    
    ufgData.ColConvertFormat = gstrMaterialConvertFormat
    
    Select Case lngStudyType
    Case 0, 3, 4, 5
        '�ж����ݿ�������Ƿ������� �����ȡ���ݿ����  û�������Ĭ��
        strTemp = zlDatabase.GetPara("����ȡ���б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
         
        If strTemp = "" Then
            ufgData.ColNames = gstrNormalMaterialCols
        Else
            ufgData.ColNames = strTemp
        End If
        
        ufgData.DefaultColNames = gstrNormalMaterialCols
    Case 1
        strTemp = zlDatabase.GetPara("����ȡ���б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
         
        If strTemp = "" Then
            ufgData.ColNames = gstrIceMaterialCols
        Else
            ufgData.ColNames = strTemp
        End If
        
        ufgData.DefaultColNames = gstrIceMaterialCols
    Case 2
        
        strTemp = zlDatabase.GetPara("ϸ��ȡ���б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
         
        If strTemp = "" Then
            ufgData.ColNames = gstrCellMaterialCols
        Else
            ufgData.ColNames = strTemp
        End If
        
        ufgData.DefaultColNames = gstrCellMaterialCols
    End Select

End Sub

Private Sub ConfigDecalinList()
'�����Ѹ���ʾ�б�
    ufgDecalin.ColConvertFormat = gstrDecalinConvertFormat
    
    '��������
    ufgDecalin.GridRows = glngStandardRowCount
    '�����и�
    ufgDecalin.RowHeightMin = glngStandardRowHeight
    
    Dim strTemp As String
    
    '�ж����ݿ�������Ƿ������� �����ȡ���ݿ����  û�������Ĭ��
    strTemp = zlDatabase.GetPara("�����Ѹ��б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
     
    If strTemp = "" Then
        '��ʼ���걾��ʾ�б�
        ufgDecalin.ColNames = gstrDecalinCols
    Else
        ufgDecalin.ColNames = strTemp
    End If
    
    '��ֹ�Ҽ������б����ô���
    ufgDecalin.IsEjectConfig = False
    ufgDecalin.DefaultColNames = gstrDecalinCols
    
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
            Call PrintMaterialLabel(False)

        Case UCase("tbLabPrint")
            '��ӡ��ǩ
            Call PrintMaterialLabel(True)
            
        Case UCase("tbGetMaterials")
            '�Ŀ���ȡ
            Call MaterialGet
            
        Case UCase("tbDelMaterials")
            'ɾ���Ŀ�
            Call DelSelectionMaterial
            
        Case UCase("tbSaveMaterials")
            '����Ŀ�
            Call SaveCurMaterialInf
            
        Case UCase("tbSureMaterials")
            'ȷ��ȡ��
            Call SureCurMaterialInf
            
        Case UCase("tbViewRequest")
            '�鿴����
            Call ShowMaterialRequest
            
    End Select
End Sub

Private Sub ufgData_OnColFormartChange()
'���ݲ�ͬ��ȡ�����ͱ��治ͬ����ͷ����

    Select Case mrecStudy.lngStudyType
        Case 0, 3, 4, 5
        
            zlDatabase.SetPara "����ȡ���б�����", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
            
        Case 1
        
            zlDatabase.SetPara "����ȡ���б�����", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
            
        Case 2
        
            zlDatabase.SetPara "ϸ��ȡ���б�����", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
            
    End Select

End Sub


Private Sub ufgData_OnColsNameReSet()
On Error GoTo errHandle

    '��������
    Call ConfigGridInput(mrecStudy.lngStudyType)
    '��ȡ�Ŀ��¼
    Call LoadMaterialData(mrecStudy.lngPatholAdviceId)
    'ˢ�²Ŀ�����
    Call RefreshMaterialCount
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgDecalin_OnColFormartChange()
'�����Ѹ��б����
 zlDatabase.SetPara "�����Ѹ��б�����", ufgDecalin.GetColsString(ufgDecalin), glngSys, G_LNG_PATHOLSYS_NUM

End Sub

Private Sub ConfigGridInput(ByVal lngStudyType As Long)
'���������б�
    Dim strSql As String
    Dim strUsers As String
    Dim strSpecimenName As String
    Dim rsData As ADODB.Recordset
    
    
    '��ȡ��ȡҽʦ
    strSql = "select a.���� from ��Ա�� a, ������Ա b where a.id=b.��ԱID and b.����ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.����ID)
    
    ufgData.ComboxListFormat(ufgData.GetColIndex(gstrMaterial_��ȡҽʦ)) = " "
    If rsData.RecordCount > 0 Then
        strUsers = ""
        While Not rsData.EOF
            If Trim(strUsers) <> "" Then strUsers = strUsers & "|"
            
            strUsers = strUsers & Nvl(rsData!����)
            
            rsData.MoveNext
        Wend
        
        ufgData.ComboxListFormat(ufgData.GetColIndex(gstrMaterial_��ȡҽʦ)) = strUsers
        ufgData.ComboxListFormat(ufgData.GetColIndex(gstrMaterial_��ȡҽʦ)) = " |" & strUsers
        ufgDecalin.ComboxListFormat(ufgDecalin.GetColIndex(gstrDecalin_����Ա)) = strUsers
    End If
    
    
    '��ȡ�걾����
    strSql = "select �걾ID, �걾����,������� from ����걾��Ϣ where �ͼ�ID>0 and ҽ��ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
    
    ufgData.ComboxListFormat(ufgData.GetColIndex(gstrMaterial_�걾����)) = " "
    ufgDecalin.ComboxListFormat(ufgDecalin.GetColIndex(gstrDecalin_�걾����)) = " "
    If rsData.RecordCount > 0 Then
        strSpecimenName = ""
        
        While Not rsData.EOF
            '�ж������������� ���� ���� ��������� ���� ���� ��������� �걾�Ĳ���������
            If Nvl(rsData!�������) = 1 And mrecStudy.lngStudyType = 3 Or Nvl(rsData!�������) = 0 Then
                If Trim(strSpecimenName) <> "" Then strSpecimenName = strSpecimenName & "|"
                strSpecimenName = strSpecimenName & "#" & Nvl(rsData!�걾ID) & ";" & Nvl(rsData!�걾����)
            End If
            
            rsData.MoveNext
        Wend
        
        ufgData.ComboxListFormat(ufgData.GetColIndex(gstrMaterial_�걾����)) = strSpecimenName
        
        '�������Ϊ ��ϸ����������ʯ�����������û���Ѹƹ��� �����ü����Ѹ��б�ı걾���ƣ������桱����������ʬ�족������ż��ء�
        If lngStudyType = 0 Or lngStudyType = 1 Or lngStudyType = 3 Or lngStudyType = 4 Then
            ufgDecalin.ComboxListFormat(ufgDecalin.GetColIndex(gstrDecalin_�걾����)) = strSpecimenName
        End If
    End If
End Sub



Public Sub ReSetFormFontSize(ByVal bytFontSize As Byte)
'����:�������ù���վ����������С
 On Error GoTo errHandle
 
    Dim objCtrl As control
    Dim CtlFont As StdFont
    Dim strFontType  As String
    
    '���ƶ��ؼ�λ��
    mbytFontSize = bytFontSize

    
    '�����������С
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
            objCtrl.Height = TextHeight("��") + 114
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
            objCtrl.Width = TextWidth("�޹�" & objCtrl.Caption)
        Case UCase("CheckBox")
            objCtrl.FontSize = bytFontSize
            objCtrl.FontName = strFontType
            objCtrl.Width = TextWidth("�޹�" & objCtrl.Caption)
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
    
    Call picMaterial_Resize
    Call picPane1_Resize
    
    Exit Sub
errHandle:
End Sub

'Private Sub LoadMaterialParameter()
'    wtDescription.ModuleHeight = zlDatabase.GetPara("Material_ModuleHeight", glngSys, glngModul, 0)
'    wtDescription.WordWidth = zlDatabase.GetPara("Material_WordWidth", glngSys, glngModul, 0)
'
'    mlngMaterialListHeight = zlDatabase.GetPara("Material_ListHeight", glngSys, glngModul, 0)
'End Sub


'Private Sub SaveMaterialParameter()
'    Call zlDatabase.SetPara("Material_ModuleHeight", wtDescription.ModuleHeight, glngSys, glngModul, True)
'    Call zlDatabase.SetPara("Material_WordWidth", wtDescription.WordWidth, glngSys, glngModul, True)
'    Call zlDatabase.SetPara("Material_ListHeight", mlngMaterialListHeight, glngSys, glngModul, True)
'
'End Sub

Private Sub SwitchWork(ByVal blnIsChangeDescription As Boolean)
'�л�����ҳ��
    framDecalin.Visible = Not blnIsChangeDescription
    framWordEdit.Visible = blnIsChangeDescription
End Sub


Private Sub LoadDecalinData(ByVal lngAdviceID As Long)
'�����Ѹ���Ϣ
    Dim strSql As String
    Dim rsDecalin As ADODB.Recordset
    
    strSql = "select a.ID,a.�걾ID,b.�걾����,a.��ʼʱ��, case when a.����ʱ�� / 60 < 1 then '0' else '' end || to_char(a.����ʱ�� / 60) as ����ʱ��, a.��ʼʱ�� + a.����ʱ��/60/24 as ����ʱ��, a.��ǰ�״�,a.���״̬,a.����Ա" & _
                " from �����Ѹ���Ϣ a, ����걾��Ϣ b " & _
                " where a.�걾id = b.�걾id and b.ҽ��ID =[1] order by a.���״̬, a.��ʼʱ��,a.Id"
    
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgDecalin.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceID)
    
    Call ufgDecalin.RefreshData
End Sub


Private Sub LoadMaterialData(ByVal lngPatholAdviceId As Long)
'����Ŀ��¼��Ϣ
    Dim strSql As String
    Dim rsMaterial As ADODB.Recordset
    
    strSql = "select a.�Ŀ�ID,a.���, a.�걾ID as �걾����,a.�Ƿ��Ѹ�,a.����ID,a.�Ƿ�����, a.ȷ��״̬,case when a.����ID>0 then '��ȡ��' else '����ȡ��' end as ȡ������, a.�걾ID,a.ȡ��λ��,a.��״,a.��ɫ,a.����,a.�걾��,a.������,b.��Ƭ��,a.�Ƿ����,a.��ȡҽʦ,a.��ȡҽʦ,a.��¼ҽʦ,a.ȡ��ʱ�� " & _
                " from ����ȡ����Ϣ a,������Ƭ��Ϣ b" & _
                " where a.����ҽ��ID =[1] and a.�Ŀ�ID = b.�Ŀ�ID and (b.����ID is null or a.����ID=b.����ID) " & _
                " union all " & _
                "  select a.�Ŀ�ID,a.���, a.�걾ID as �걾����,a.�Ƿ��Ѹ�,a.����ID, a.�Ƿ�����,a.ȷ��״̬,'�������' as ȡ������, " & _
                "  a.�걾ID,a.ȡ��λ��,a.��״,a.��ɫ,a.����,a.�걾��,a.������,0 as ��Ƭ��,a.�Ƿ����,a.��ȡҽʦ,a.��ȡҽʦ,'' as ��¼ҽʦ,a.ȡ��ʱ�� " & _
                "  from ����ȡ����Ϣ a,��������Ϣ b, ����걾��Ϣ c " & _
                "  Where a.����ҽ��ID = [1] And a.����ҽ��ID = b.����ҽ��ID And a.�걾ID = c.�걾ID And c.������� = 1 And b.������� = 3"

    strSql = "select * from (" & strSql & ") order by ȡ������,���"

'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatholAdviceId)
    
    Call ufgData.RefreshData
End Sub


Private Sub LoadDescriptionInf(ByVal lngPatholAdviceId As Long)
'����޼���������Ϣ
    Dim strSql As String
    Dim rsDescription As ADODB.Recordset
    
    strSql = "select �޼�����,ʣ��λ��,�������� from ��������Ϣ where ����ҽ��ID=[1]"
    
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set rsDescription = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatholAdviceId)
    
    wtDescription.WordText = ""
    txtPos.Text = ""
    
    If rsDescription.RecordCount <= 0 Then Exit Sub
    
    wtDescription.WordText = Nvl(rsDescription("�޼�����").value)
    txtPos.Text = Nvl(rsDescription("ʣ��λ��").value)
    cbxSpecimenProcess.Text = Nvl(rsDescription("��������").value)
End Sub


Private Function Decalin_Start(ByVal lngDecalinRowIndex As Long) As String
'��ʼ�Ѹ�
    Dim strSql As String
    Dim lngTimeLen As Long
    Dim dtEndTime As Date
    Dim rsDecalin As ADODB.Recordset
    
    Decalin_Start = ""
    
    strSql = "select Zl_�����Ѹ�_��ʼ([1],[2],[3],[4]) as ����ֵ from dual"
    
    lngTimeLen = Fix(Val(ufgDecalin.Text(lngDecalinRowIndex, gstrDecalin_����ʱ��)) * 60)
    dtEndTime = CDate(ufgDecalin.Text(lngDecalinRowIndex, gstrDecalin_��ʼʱ��))
    
    Set rsDecalin = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                                ufgDecalin.Text(lngDecalinRowIndex, gstrDecalin_�걾����), _
                                                dtEndTime, _
                                                lngTimeLen, _
                                                ufgDecalin.Text(lngDecalinRowIndex, gstrDecalin_����Ա) _
                                                )
                                                
    If rsDecalin.RecordCount <= 0 Then
        Decalin_Start = "�Ѹ�ִ��ʧ�ܣ�δ�ܷ�����Ч���Ѹ�ID��"
        Exit Function
    End If
    
    
    '�����Ѹ���ʾ�б�
    ufgDecalin.RowState(lngDecalinRowIndex) = TDataRowState.Normal
    
    ufgDecalin.Text(lngDecalinRowIndex, gstrDecalin_ID) = rsDecalin!����ֵ
    ufgDecalin.Text(lngDecalinRowIndex, gstrDecalin_����ʱ��) = DateAdd("n", lngTimeLen, dtEndTime)
    ufgDecalin.Text(lngDecalinRowIndex, gstrDecalin_��ǰ�״�) = 1
    ufgDecalin.Text(lngDecalinRowIndex, gstrDecalin_����Ա) = UserInfo.����
    ufgDecalin.Text(lngDecalinRowIndex, gstrDecalin_��ǰ״̬) = "δ���"
    
End Function



Private Sub Decalin_Succed()
'����Ѹ�
    Dim strSql As String
    Dim lngDecalinId As Long
    
    lngDecalinId = ufgDecalin.KeyValue(ufgDecalin.SelectionRow)
    
    strSql = "Zl_�����Ѹ�_���(" & lngDecalinId & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    '�����Ѹ���ʾ�б�
    ufgDecalin.Text(ufgDecalin.SelectionRow, gstrDecalin_��ǰ״̬) = "�����"
End Sub



Private Sub Decalin_Change(ByVal dtStart As Date, ByVal lngTimeLen As Double)
'�Ѹƻ���
    Dim strSql As String
    Dim lngDecalinId As Long
    
    lngDecalinId = ufgDecalin.KeyValue(ufgDecalin.SelectionRow)
    
    strSql = "Zl_�����Ѹ�_����(" & lngDecalinId & "," & To_Date(dtStart) & "," & Fix(lngTimeLen * 60) & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    '�����Ѹ���ʾ�б�
    ufgDecalin.Text(ufgDecalin.SelectionRow, gstrDecalin_��ǰ�״�) = Val(ufgDecalin.Text(ufgDecalin.SelectionRow, gstrDecalin_��ǰ�״�)) + 1
    ufgDecalin.Text(ufgDecalin.SelectionRow, gstrDecalin_��ʼʱ��) = dtStart
    ufgDecalin.Text(ufgDecalin.SelectionRow, gstrDecalin_����ʱ��) = Format$(lngTimeLen, "0.0")
    ufgDecalin.Text(ufgDecalin.SelectionRow, gstrDecalin_����ʱ��) = DateAdd("n", lngTimeLen * 60, dtStart)
End Sub



Private Sub Decalin_Cancel()
'�����Ѹ�
    Dim strSql As String
    Dim lngDecalinId As Long

    lngDecalinId = Val(ufgDecalin.KeyValue(ufgDecalin.SelectionRow))

    If Trim(lngDecalinId) > 0 Then
        strSql = "Zl_�����Ѹ�_����(" & lngDecalinId & ")"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    End If

    'ɾ���Ѹ���ʾ�б�
    Call ufgDecalin.DelCurRow
End Sub

Private Sub CancelVat()
'�����Ѹ�

    If ufgData.ShowingDataRowCount <= 0 Then Exit Sub
    
    If Not ufgDecalin.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ�����ѸƵļ�¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgDecalin.IsNullRow(ufgDecalin.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ҫ�����ѸƵļ�¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    
    '�жϵ�ǰ��¼�Ƿ��Ѿ�����Ѹƣ�����ɵ��Ѹ������ܽ��г���
    If ufgDecalin.Text(ufgDecalin.SelectionRow, gstrDecalin_��ǰ״̬) = "�����" Then
        Call MsgBoxD(Me, "�ñ걾������Ѹƣ����ܽ��г�����", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    
    If MsgBoxD(Me, "ȷ��Ҫɾ����ǰδ��ɵ��Ѹ�������", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    Call Decalin_Cancel
    
    Call ConfigDecalcificationBut
End Sub

Private Sub cmdCancel_Click()
On Error GoTo errHandle
    Call CancelVat
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ChangeVat()
'����
    Dim frmChangeInput As frmPatholMaterials_Change
    
    If Not ufgDecalin.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ���׵ļ�¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgDecalin.IsNullRow(ufgDecalin.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ҫ���׵ļ�¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '�жϵ�ǰ��¼�Ƿ��Ѿ���ʼ�Ѹ�
    If ufgDecalin.IsEmptyKey(ufgDecalin.SelectionRow) Then
        Call MsgBoxD(Me, "�ñ걾��δ��ʼ�Ѹƣ�����ִ�л��ײ���������ִ���Ѹơ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgDecalin.Text(ufgDecalin.SelectionRow, gstrDecalin_��ǰ״̬) = "�����" Then
        Call MsgBoxD(Me, "�Ѹ���������ɣ����ܽ��л��ײ�����", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Set frmChangeInput = New frmPatholMaterials_Change
    
On Error GoTo errFree
    
    Call frmChangeInput.ShowChangeWindow(Me)
        
    If Not frmChangeInput.IsSure Then Exit Sub
    
    '����
    Call Decalin_Change(frmChangeInput.StartTime, frmChangeInput.TimeLen)
errFree:
    Unload frmChangeInput
    Set frmChangeInput = Nothing
End Sub


Private Sub cmdChange_Click()
On Error GoTo errHandle
    Call ChangeVat
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Decalcification()
'�Ѹ�

    Dim strErr As String
    Dim blnValid As Boolean
    
    If Not ufgDecalin.IsSelectionRow Then
        Call MsgBoxD(Me, "��¼����Ҫ�ѸƵ������Ϣ��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgDecalin.IsNullRow(ufgDecalin.SelectionRow) Then
        Call MsgBoxD(Me, "��¼����Ҫ�ѸƵ������Ϣ��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '�жϵ�ǰ��¼�Ƿ��Ѿ���ʼ�Ѹ�
    If Not ufgDecalin.IsEmptyKey(ufgDecalin.SelectionRow) Then
        Call MsgBoxD(Me, "�ñ걾�ѿ�ʼִ���ѸƲ�����", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    
    '��������Ƿ���Ч
    blnValid = Not ufgDecalin.IsErrColorWithRow(ufgDecalin.SelectionRow)
    If Not blnValid Then
        Call MsgBoxD(Me, "��⵽�Ѹ��б������Ч���ݣ���ȷ���Ƿ���ȷ������¼�룬����ɫ����ǵĵ�Ԫ��Ϊ��¼���ݡ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    
    '��ʼ�Ѹ�
    strErr = Decalin_Start(ufgDecalin.SelectionRow)
    If Trim(strErr) <> "" Then
        Call MsgBoxD(Me, strErr, vbOKOnly, Me.Caption)
    End If
    
    
    Call ConfigDecalcificationBut
    
End Sub


Private Sub cmdDecalin_Click()
On Error GoTo errHandle
    Call Decalcification
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub DelSelectionMaterial()
'ɾ��ѡ�еĲĿ�

    If ufgData.ShowingRowCount <= 0 Then Exit Sub
    
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫɾ���ĲĿ��¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.IsNullRow(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ҫɾ���ĲĿ��¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '�жϸòĿ��Ƿ�����Ƭ
    If Not ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        If Not CheckAllowUpdate(ufgData.KeyValue(ufgData.SelectionRow)) Then
            Call MsgBoxD(Me, "�òĿ��¼��ִ����Ƭ�������ܽ���ɾ����", vbOKOnly, Me.Caption)
            Exit Sub
        End If
    End If
    
    '�жϸòĿ��Ƿ���������Ƭ
    If Not ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        If CheckRequisitionSlices(ufgData.KeyValue(ufgData.SelectionRow)) Then
            Call MsgBoxD(Me, "�òĿ��¼��������Ƭ�������ܽ���ɾ����", vbOKOnly, Me.Caption)
            Exit Sub
        End If
    End If
    
    '�жϸòĿ��Ƿ��������ؼ�
    If Not ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        If CheckRequisitionSpeExam(ufgData.KeyValue(ufgData.SelectionRow)) Then
            Call MsgBoxD(Me, "�òĿ��¼�������ؼ촦�����ܽ���ɾ����", vbOKOnly, Me.Caption)
            Exit Sub
        End If
    End If
    
    If MsgBoxD(Me, "ȷ��Ҫɾ��ѡ��ĲĿ�������", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    'ɾ����
    Call ufgData.DelCurRow
    
    '����ɾ���ĲĿ�����
    Call SaveMaterialData(True)
    
    'ˢ�²Ŀ�����
    Call RefreshMaterialCount
End Sub

Private Sub RefreshMaterialCount()
    'ˢ�²Ŀ�����
    Dim lngTotal As Long
    Dim i As Long
    
    lngTotal = 0
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.RowHidden(i) Then
            If Not ufgData.IsNullRow(i) Then
                
                Select Case mrecStudy.lngStudyType
                    Case StudyType.stNormal, StudyType.stMeet, StudyType.stAutopsy, StudyType.stIce, StudyType.stSpeed
                        lngTotal = lngTotal + Val(ufgData.Text(i, gstrMaterial_������))
                    Case StudyType.stCell
                        lngTotal = lngTotal + Val(ufgData.Text(i, gstrMaterial_ϸ������))
                End Select
            End If
        End If
    Next i
    
    labRecordInf.Caption = "�Ŀ�������" & lngTotal
End Sub


Private Sub AutoGetMaterialInf()
'�Զ���ȡ�Ŀ���Ϣ
    Dim strSql As String
    Dim rsSpeciman As ADODB.Recordset
    Dim lngRow As Long
    Dim i As Long
    Dim dtServicesTime As Date
    Dim strComboboxText As String
    
    strSql = "select a.�걾ID,a.�걾����,a.�������,b.Ĭ�ϱ걾��,b.Ĭ����Ƭ�� from ����걾��Ϣ a,������걾 b where a.�걾���� = b.�걾����(+) and a.ҽ��ID=[1]"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    Set rsSpeciman = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
    
    If rsSpeciman.RecordCount <= 0 Then Exit Sub
    
    lngRow = 0
    dtServicesTime = zlDatabase.Currentdate
    
    Do While Not rsSpeciman.EOF
        
        '������ղ��ϲ�Ϊ�걾���ͣ��򲻼������ݣ�ֱ������
        If Nvl(rsSpeciman!�������) <> 0 Then GoTo continue
        
        For i = 1 To ufgData.GridRows - 1
            If (ufgData.Text(i, gstrMaterial_�걾����) = rsSpeciman("�걾ID").value Or ufgData.Text(i, gstrMaterial_�걾����) = rsSpeciman("�걾����").value) _
                And Not ufgData.RowHidden(i) Then GoTo continue
            
            If ufgData.IsNullRow(i) And Not ufgData.RowHidden(i) Then
                lngRow = i
                Exit For
            End If
        Next i
        If lngRow = 0 Then Exit Do
        
        ufgData.Text(lngRow, gSpecimen_�걾����) = Nvl(rsSpeciman!�걾ID)
        
        ufgData.Text(lngRow, gstrMaterial_��Ƭ��) = IIf(Nvl(rsSpeciman!Ĭ����Ƭ��) = "", 1, Nvl(rsSpeciman!Ĭ����Ƭ��))
        
        '�ж�����������Ϊ2���ȡĬ�ϱ걾��
        If mrecStudy.lngStudyType = 2 Then
            ufgData.Text(lngRow, gstrMaterial_�걾��) = Nvl(rsSpeciman!Ĭ�ϱ걾��)
        End If

        '�ڲĿ���ȡ�в���Ҫȡ��ʱ��,ȡ��ʱ����ȷ��ȡ�ĺ�Ų���
        'ufgData.Text(lngRow, gstrMaterial_ȡ��ʱ��) = dtServicesTime
        
        If mrecStudy.lngStudyType <> StudyType.stCell And mrecStudy.lngStudyType <> StudyType.stIce Then
            ufgData.Text(lngRow, gstrMaterial_������) = "1"
            If ufgData.Text(lngRow, gstrMaterial_�Ƿ�����) = "" Then ufgData.Text(lngRow, gstrMaterial_�Ƿ�����) = "1-��"
            
        Else
            If mrecStudy.lngStudyType = StudyType.stIce Then
                ufgData.Text(lngRow, gstrMaterial_�Ƿ����) = "0-��"
                ufgData.Text(lngRow, gstrMaterial_������) = "0"
            Else
                ufgData.Text(lngRow, gstrMaterial_ϸ������) = "0"
            End If
            
            If ufgData.Text(lngRow, gstrMaterial_�Ƿ�����) = "" Then ufgData.Text(lngRow, gstrMaterial_�Ƿ�����) = "0-��"
        End If
        
        If Not IsDate(ufgData.Text(lngRow, gstrMaterial_ȡ��ʱ��)) Then
            ufgData.Text(lngRow, gstrMaterial_ȡ��ʱ��) = dtServicesTime
        End If
        
        
        If ufgData.Text(lngRow, gstrMaterial_��ȡҽʦ) = "" Then
            If lngRow - 1 > 0 Then
                If ufgData.Text(lngRow - 1, gstrMaterial_��ȡҽʦ) <> "" Then
                    ufgData.Text(lngRow, gstrMaterial_��ȡҽʦ) = ufgData.Text(lngRow - 1, gstrMaterial_��ȡҽʦ)
                End If
            End If
            
            If ufgData.Text(lngRow, gstrMaterial_��ȡҽʦ) = "" Then
                strComboboxText = ufgData.DataGrid.ColComboList(ufgData.GetColIndex(gstrMaterial_��ȡҽʦ))
                
                If strComboboxText <> "" Then
                    If InStr(strComboboxText, "|") > 0 Then
                        strComboboxText = Mid(strComboboxText, 1, InStr(strComboboxText, "|") - 1)
                    End If
                    ufgData.Text(lngRow, gstrMaterial_��ȡҽʦ) = strComboboxText
                    
                End If
            End If
        End If
        
        '���µ�ǰ��״̬Ϊ���
        ufgData.RowState(lngRow) = TDataRowState.Add
        
        Call ufgData_OnAfterEdit(lngRow, ufgData.GetColIndex(gSpecimen_�걾����))
        
continue:
        rsSpeciman.MoveNext
    Loop
    
    '��ʾ�û�
    If ufgData.ShowingDataRowCount = 0 Then
        Call MsgBoxD(Me, "���ղ��ϲ�Ϊ�걾���ͣ������Զ���ȡ�Ŀ���Ϣ��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
End Sub


Private Sub MaterialGet()
'��ȡ�Ŀ�
    '��ȡ�Ľ׶Σ����ܽ���ȷ��
    If mrecStudy.lngMaterialStep <> TExecuteStep.NeedDo Then
        Call MsgBoxD(Me, "��δ����ȡ�Ľ׶Σ������Զ���ȡ�Ŀ���Ϣ��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '�Զ���ȡ�Ŀ���Ϣ
    Call AutoGetMaterialInf
    
    'ˢ�²Ŀ�����
    Call RefreshMaterialCount
End Sub


Private Sub PrintMaterialLabel(Optional ByVal blnIsPrint As Boolean = True)
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
    
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_01", Me, "�Ŀ�ID1=" & strValue(0), "�Ŀ�ID2=" & strValue(1), "�Ŀ�ID3=" & strValue(2), "�Ŀ�ID4=" & strValue(3), "�Ŀ�ID5=" & strValue(4), "�Ŀ�ID6=" & strValue(5), IIf(blnIsPrint, 2, 1))
End Sub



Private Sub PrintSelectMaterialLabel(Optional ByVal blnIsPrint As Boolean = True)
'��ӡѡ��ĲĿ��ǩ
On Error GoTo errHandle
    Dim strValue(5) As String
    
    
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ��ӡ�ĲĿ��¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.IsNullRow(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ҫ��ӡ�ĲĿ��¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strValue(0) = ufgData.KeyValue(ufgData.SelectionRow)
    strValue(1) = "0": strValue(2) = "0": strValue(3) = "0": strValue(4) = "0": strValue(5) = "0"

    
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_01", Me, "�Ŀ�ID1=" & strValue(0), "�Ŀ�ID2=" & strValue(1), "�Ŀ�ID3=" & strValue(2), "�Ŀ�ID4=" & strValue(3), "�Ŀ�ID5=" & strValue(4), "�Ŀ�ID6=" & strValue(5), IIf(blnIsPrint, 2, 1))
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ShowMaterialRequest()
'��ʾ�ؼ�����
Dim frmRequestView As New frmPatholRequisition_View
On Error GoTo errFree
    Call frmRequestView.ShowRequestViewWind(mrecStudy.lngPatholAdviceId, 4, mblnMoved, Me)
errFree:
    Call Unload(frmRequestView)
    Set frmRequestView = Nothing
End Sub


'Private Sub CmdRefresh_Click()
'On Error GoTo errHandle
'    '�ָ��б�����
'    Call ufgData.RefreshData
'
'    'ˢ�²Ŀ�����
'    Call RefreshMaterialCount
'
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'End Sub

Private Sub SaveCurMaterialInf()
'���浱ǰȡ����Ϣ

    Dim blnValid As Boolean
    
    '�Ŀ鱣��
    If ufgData.ShowingDataRowCount <= 0 Then
        Call MsgBoxD(Me, "û���ҵ���Ҫ����ĲĿ���Ϣ����¼��Ŀ����ݡ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    blnValid = Not ufgData.IsErrColorWithList
    If Not blnValid Then
        Call MsgBoxD(Me, "��⵽ȡ���б��д�����Ч���ݣ���ȷ����������Ƿ���ȷ������¼�룬����ɫ����ǵĵ�Ԫ��Ϊ��¼���ݡ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call SaveMaterialData
    
    Call SendMsgToMainWindow(Me, wetMaterialSave, mlngAdviceID)
    
    Call MsgBoxD(Me, "�����ѳɹ����档", vbOKOnly, Me.Caption)
    
    'ˢ�²Ŀ�����
    Call RefreshMaterialCount
    
End Sub


Private Sub Finish()
'����Ѹ�

    If ufgDecalin.ShowingRowCount <= 0 Then Exit Sub

    If Not ufgDecalin.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ����ѸƵļ�¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgDecalin.IsNullRow(ufgDecalin.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ҫ����ѸƵļ�¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '�жϵ�ǰ��¼�Ƿ��Ѿ���ʼ�Ѹ�
    If ufgDecalin.IsEmptyKey(ufgDecalin.SelectionRow) Then
        Call MsgBoxD(Me, "�ñ걾��δ��ʼ�Ѹƣ�����ִ�иò���������ִ���Ѹơ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call Decalin_Succed
End Sub

Private Sub cmdSucceed_Click()
On Error GoTo errHandle
    Call Finish

Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Function GetRequisitionId(Optional ByVal lngRequisitionType As Long = 4) As Long
'��ȡ����ID
'lngRequisitionType:Ĭ��Ϊ4����ʾ��ȡ��
'�����������ͻ�ȡ��δִ�е�����ID
'���û�в�ȡ�����¼���򷵻ؿ�����ID

    Dim strSql As String
    Dim rsRequisition As ADODB.Recordset
    
    GetRequisitionId = -1
    
    strSql = "select ����ID from ����������Ϣ where ����״̬=0 and ����ҽ��ID=[1] and ��������=[2]"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set rsRequisition = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mrecStudy.lngPatholAdviceId, lngRequisitionType)
    
    If rsRequisition.RecordCount > 0 Then GetRequisitionId = rsRequisition("����ID").value
End Function



Private Function CheckAllowUpdate(ByVal strMaterialId As String) As Boolean
'����Ƿ��������
'δ��Ƭ�ĲĿ���ɽ��и���,ͨ����鲡����Ƭ��Ϣ�����жϲĿ��Ƿ�����Ƭ(�����ǰ״̬��Ϊ0��������Ƭ)

    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    CheckAllowUpdate = True
    
    If Trim(strMaterialId) = "" Then Exit Function
    
    strSql = "select ID from ������Ƭ��Ϣ where  �Ŀ�ID=[1] and ��ǰ״̬<>0"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strMaterialId)
    
    If rsData.RecordCount > 0 Then CheckAllowUpdate = False
End Function

Private Function CheckRequisitionSlices(ByVal strMaterialId As String) As Boolean
'����Ƿ���������Ƭ
'��ִ����Ƭ����ĲĿ�ȡ�Ľ��治��ɾ���Ŀ�

    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    CheckRequisitionSlices = False
    
    If Trim(strMaterialId) = "" Then Exit Function
    
    strSql = "select ����ID from ������Ƭ��Ϣ where ����ID is not null and �Ŀ�ID=[1]"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strMaterialId)
    
    If rsData.RecordCount > 0 Then CheckRequisitionSlices = True
End Function

Private Function CheckRequisitionSpeExam(ByVal strMaterialId As String) As Boolean
'����Ƿ��������ؼ�
'�������ؼ촦��ĲĿ�ȡ�Ľ��治��ɾ���Ŀ�

    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    CheckRequisitionSpeExam = False
    
    If Trim(strMaterialId) = "" Then Exit Function
    
    strSql = "select ����ID from �����ؼ���Ϣ where  �Ŀ�ID=[1]"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strMaterialId)
    
    If rsData.RecordCount > 0 Then CheckRequisitionSpeExam = True
End Function


Public Sub SureMaterialData()
'ȷ��ȡ������
    Dim strSql As String
    Dim i As Long
    
    strSql = "Zl_����ȡ��_ȷ��('" & mrecStudy.lngPatholAdviceId & "')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsNullRow(i) Then
            ufgData.Text(i, gstrMaterial_ȷ��״̬) = "��ȷ��"
        End If
    Next i
    
    
    mrecStudy.lngSlicesStep = 1
    mrecStudy.lngMaterialStep = 2
End Sub


Private Sub SplitMaterialNumber(ByVal strDataValue As String, ByRef strID As String, ByRef strSeq As String)
'�ֽ���ִ�й��̷��صĲĿ����
    Dim lngFind As Long
    
    lngFind = InStr(strDataValue, "-")
    
    If lngFind <= 0 Then Exit Sub
    
    strID = Mid(strDataValue, 1, lngFind - 1)
    strSeq = Mid(strDataValue, lngFind + 1, 18)
End Sub


Public Sub SaveMaterialData(Optional ByVal blnIsSaveOnlyDel As Boolean = False)
'------------------------------------------------------------------------------
'blnIsSaveOnlyDel:�Ƿ��������ɾ��������
'------------------------------------------------------------------------------


'ȡ��ȷ�ϱ���
'���û���µĲĿ飬��ֻ����޼�������ʣ��λ��
'�������Ƭ�����ܽ��и��²���


    Dim i As Long
    Dim strSql As String
    Dim rsResult As ADODB.Recordset
    Dim lngRequisitionId As Long
    Dim dtSerivcesTime As Date
    Dim strNewId As String
    Dim strNewSeq As String
    Dim lngCount As Long
    
    
    '��ȡ��ȡ������ID�����û�в�ȡ�����¼���򷵻ؿ�����ID
    lngRequisitionId = GetRequisitionId
    
    For i = 1 To ufgData.GridRows - 1
        If ufgData.RowState(i) = TDataRowState.Add And Not blnIsSaveOnlyDel Then
            
            dtSerivcesTime = ufgData.Text(i, gstrMaterial_ȡ��ʱ��) 'zlDatabase.Currentdate
            
            '���ȡ�ļ�¼
            Select Case mrecStudy.lngStudyType
                Case StudyType.stNormal, StudyType.stMeet, StudyType.stAutopsy, StudyType.stSpeed
                    strSql = "select Zl_����ȡ��_����([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14])  as ����ֵ from dual"
                    
                    '�����������ʾ�ڽ����У����Ծ����¼������Ϊ׼
                    lngCount = 1
                    If Not ufgData.DataGrid.ColHidden(ufgData.GetColIndex(gstrMaterial_������)) Then
                        lngCount = Val(ufgData.Text(i, gstrMaterial_������))
                    End If
                    
                    Set rsResult = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                                            mrecStudy.lngPatholAdviceId, _
                                                            IIf(lngRequisitionId <= 0, "", lngRequisitionId), _
                                                            ufgData.Text(i, gstrMaterial_�걾����), _
                                                            ufgData.DisplayText(i, gstrMaterial_�걾����), _
                                                            ufgData.Text(i, gstrMaterial_ȡ��λ��), _
                                                            ufgData.Text(i, gstrMaterial_��״), _
                                                            lngCount, _
                                                            Val(ufgData.Text(i, gstrMaterial_��Ƭ��)), _
                                                            ufgData.Text(i, gstrMaterial_��ȡҽʦ), _
                                                            ufgData.Text(i, gstrMaterial_��ȡҽʦ), _
                                                            Val(ufgData.Text(i, gstrMaterial_�Ƿ�����)), _
                                                            Val(ufgData.Text(i, gstrMaterial_�Ƿ��Ѹ�)), _
                                                            UserInfo.����, _
                                                            CDate(dtSerivcesTime))
                                                            
                                                            
                Case StudyType.stIce
                    strSql = "select Zl_����ȡ��_����([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14])  as ����ֵ from dual"
                    
                    '�����������ʾ�ڽ����У����Ծ����¼������Ϊ׼
                    lngCount = 1
                    If Not ufgData.DataGrid.ColHidden(ufgData.GetColIndex(gstrMaterial_������)) Then
                        lngCount = Val(ufgData.Text(i, gstrMaterial_������))
                    End If
                    
                    Set rsResult = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                                            mrecStudy.lngPatholAdviceId, _
                                                            IIf(lngRequisitionId <= 0, "", lngRequisitionId), _
                                                            ufgData.Text(i, gstrMaterial_�걾����), _
                                                            ufgData.DisplayText(i, gstrMaterial_�걾����), _
                                                            ufgData.Text(i, gstrMaterial_ȡ��λ��), _
                                                            ufgData.Text(i, gstrMaterial_��״), _
                                                            Val(ufgData.Text(i, gstrMaterial_�Ƿ�����)), _
                                                            Val(ufgData.Text(i, gstrMaterial_�Ƿ����)), _
                                                            lngCount, _
                                                            Val(ufgData.Text(i, gstrMaterial_��Ƭ��)), _
                                                            ufgData.Text(i, gstrMaterial_��ȡҽʦ), _
                                                            ufgData.Text(i, gstrMaterial_��ȡҽʦ), _
                                                            UserInfo.����, _
                                                            CDate(dtSerivcesTime))
                Case StudyType.stCell
                    strSql = "select Zl_����ȡ��_ϸ��([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14]) as ����ֵ from dual"
                    
                    Set rsResult = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                                            mrecStudy.lngPatholAdviceId, _
                                                            IIf(lngRequisitionId <= 0, "", lngRequisitionId), _
                                                            ufgData.Text(i, gstrMaterial_�걾����), _
                                                            ufgData.DisplayText(i, gstrMaterial_�걾����), _
                                                            ufgData.Text(i, gstrMaterial_��ɫ), _
                                                            ufgData.Text(i, gstrMaterial_����), _
                                                            ufgData.Text(i, gstrMaterial_�걾��), _
                                                            Val(ufgData.Text(i, gstrMaterial_��Ƭ��)), _
                                                            Val(ufgData.Text(i, gstrMaterial_�Ƿ�����)), _
                                                            Val(ufgData.Text(i, gstrMaterial_ϸ������)), _
                                                            ufgData.Text(i, gstrMaterial_��ȡҽʦ), _
                                                            ufgData.Text(i, gstrMaterial_��ȡҽʦ), _
                                                            UserInfo.����, _
                                                            CDate(dtSerivcesTime))
            End Select
            
            
            If rsResult.RecordCount <= 0 Then
                Call err.Raise(0, "SaveMaterialData", "δ�ɹ���ȡ������ĲĿ��,����ʧ�ܡ�")
                Exit Sub
            End If
            
            Call SplitMaterialNumber(rsResult("����ֵ").value, strNewId, strNewSeq)
            
            '���²Ŀ��б�
            ufgData.Text(i, gstrMaterial_�Ŀ�ID) = strNewId
            ufgData.Text(i, gstrMaterial_�Ŀ��) = strNewSeq
            ufgData.Text(i, gstrMaterial_��¼ҽʦ) = UserInfo.����
            ufgData.Text(i, gstrMaterial_ȡ��ʱ��) = dtSerivcesTime
            ufgData.Text(i, gstrMaterial_ȡ������) = IIf(lngRequisitionId > 0, "��ȡ��", "����ȡ��")
            ufgData.Text(i, gstrMaterial_ȷ��״̬) = "δȷ��"
            
        ElseIf ufgData.RowState(i) = TDataRowState.Update And Not blnIsSaveOnlyDel Then
            dtSerivcesTime = ufgData.Text(i, gstrMaterial_ȡ��ʱ��)
            
            '����ȡ�ļ�¼
            Select Case mrecStudy.lngStudyType
                Case StudyType.stNormal, StudyType.stMeet, StudyType.stAutopsy, StudyType.stSpeed
                    strSql = "Zl_����ȡ��_�������('" & ufgData.KeyValue(i) & "','" & _
                                                        ufgData.Text(i, gstrMaterial_ȡ��λ��) & "','" & _
                                                        ufgData.Text(i, gstrMaterial_��״) & "'," & _
                                                        Val(ufgData.Text(i, gstrMaterial_������)) & "," & _
                                                        Val(ufgData.Text(i, gstrMaterial_��Ƭ��)) & ",'" & _
                                                        ufgData.Text(i, gstrMaterial_��ȡҽʦ) & "','" & _
                                                        ufgData.Text(i, gstrMaterial_��ȡҽʦ) & "'," & _
                                                        Val(ufgData.Text(i, gstrMaterial_�Ƿ�����)) & "," & _
                                                        Val(ufgData.Text(i, gstrMaterial_�Ƿ��Ѹ�)) & "," & _
                                                        To_Date(dtSerivcesTime) & ")"
                    
                    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
                                                            
                                                            
                Case StudyType.stIce
                    strSql = "Zl_����ȡ��_��������('" & ufgData.KeyValue(i) & "','" & _
                                                        ufgData.Text(i, gstrMaterial_ȡ��λ��) & "','" & _
                                                        ufgData.Text(i, gstrMaterial_��״) & "'," & _
                                                        Val(ufgData.Text(i, gstrMaterial_�Ƿ����)) & "," & _
                                                        Val(ufgData.Text(i, gstrMaterial_�Ƿ�����)) & "," & _
                                                        Val(ufgData.Text(i, gstrMaterial_������)) & "," & _
                                                        Val(ufgData.Text(i, gstrMaterial_��Ƭ��)) & ",'" & _
                                                        ufgData.Text(i, gstrMaterial_��ȡҽʦ) & "','" & _
                                                        ufgData.Text(i, gstrMaterial_��ȡҽʦ) & "'," & _
                                                        To_Date(dtSerivcesTime) & ")"
                    
                    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
                    
                Case StudyType.stCell
                    strSql = "Zl_����ȡ��_ϸ������('" & ufgData.KeyValue(i) & "','" & _
                                                        ufgData.Text(i, gstrMaterial_��ɫ) & "','" & _
                                                        ufgData.Text(i, gstrMaterial_����) & "','" & _
                                                        ufgData.Text(i, gstrMaterial_�걾��) & "'," & _
                                                        Val(ufgData.Text(i, gstrMaterial_��Ƭ��)) & "," & _
                                                        Val(ufgData.Text(i, gstrMaterial_�Ƿ�����)) & "," & _
                                                        Val(ufgData.Text(i, gstrMaterial_ϸ������)) & ",'" & _
                                                        ufgData.Text(i, gstrMaterial_��ȡҽʦ) & "','" & _
                                                        ufgData.Text(i, gstrMaterial_��ȡҽʦ) & "'," & _
                                                        To_Date(dtSerivcesTime) & ")"
                    
                    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            End Select
        
            
        ElseIf ufgData.RowState(i) = TDataRowState.Del Then
            'ɾ��ȡ�ļ�¼
            If Trim(ufgData.KeyValue(i)) <> "" Then
                strSql = "Zl_����ȡ��_ɾ��('" & ufgData.KeyValue(i) & "')"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            End If
        End If
        
        
        '������״̬
        ufgData.RowState(i) = TDataRowState.Normal
    Next i
    
    
    '����ȡ����Ϣ(�޼�������ʣ��λ�õ�)
    strSql = "Zl_����ȡ��_��Ϣ����('" & mrecStudy.lngPatholAdviceId & "','" & wtDescription.WordText & "','" & txtPos.Text & "','" & cbxSpecimenProcess.Text & "')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    

End Sub


Private Sub SureCurMaterialInf()
'ȷ�ϵ�ǰȡ����Ϣ
    Dim i As Integer
    Dim blnValid As Boolean
    
    '������Ѹ�δ��ɵģ����ܽ���ȷ��
    If Not ufgDecalin.IsNullRow(i) And ufgDecalin.RowState(i) <> Del Then
        For i = 1 To ufgDecalin.DataGrid.Rows - 1
            If ufgDecalin.Text(i, gstrDecalin_��ǰ״̬) <> "" And ufgDecalin.Text(i, gstrDecalin_��ǰ״̬) <> "�����" Then
                Call MsgBoxD(Me, "�����Ѹ�δ��ɣ����ܽ���ȡ��ȷ�ϲ�����", vbOKOnly, Me.Caption)
                Exit Sub
            End If
        Next
    End If
    
    '��ȡ�Ľ׶Σ����ܽ���ȷ��
    If mrecStudy.lngMaterialStep <> TExecuteStep.NeedDo Then
        Call MsgBoxD(Me, "��δ����ȡ�Ľ׶Σ����ܽ���ȡ��ȷ�ϲ�����", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '�Ŀ鱣��
    If ufgData.ShowingDataRowCount <= 0 Then
        Call MsgBoxD(Me, "û���ҵ���Ҫȷ�ϵĲĿ���Ϣ����¼��Ŀ����ݡ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    blnValid = Not ufgData.IsErrColorWithList
    If Not blnValid Then
        Call MsgBoxD(Me, "��⵽ȡ���б��д�����Ч���ݣ���ȷ����������Ƿ���ȷ������¼�룬����ɫ����ǵĵ�Ԫ��Ϊ��¼���ݡ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '����Ŀ��¼(�ȱ���Ŀ��¼)
    Call SaveMaterialData
    
    'ȷ��ȡ��
    Call SureMaterialData
    
    '�����¼�
    Call SendMsgToMainWindow(Me, wetMaterialSure, mlngAdviceID)
    
    Call MsgBoxD(Me, "�����ȡ��ȷ�ϡ�", vbOKOnly, Me.Caption)

End Sub


Private Sub Form_Initialize()
    Dim strRegPath As String
    Set zlReport = New zl9Report.clsReport
    
    strRegPath = "����ģ��\" & App.ProductName & "\" & Me.Name
    picMaterial.Height = Val(GetSetting("ZLSOFT", strRegPath, "MaterialListHeight", picMaterial.Height))
End Sub


Private Sub LoadReportModule()
'���뱨��ģ��
    Dim strLinkClassName As String
    
    If mlngCurDeptId = wtDescription.CurDepartId Then Exit Sub
    
    strLinkClassName = zlDatabase.GetPara("�޼�����ģ��", glngSys, glngModul, "")
    
    wtDescription.ModuleName = strLinkClassName
    wtDescription.CurDepartId = mlngCurDeptId
    
    Call wtDescription.LoadWordModel
End Sub


Private Sub Form_Load()
On Error GoTo errHandle
    
    Call SwitchWork(True)
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Dim strRegPath As String
    
'    '�������
'    Call SaveMaterialParameter
    strRegPath = "����ģ��\" & App.ProductName & "\" & Me.Name
    Call SaveSetting("ZLSOFT", strRegPath, "MaterialListHeight", picMaterial.Height)
    
    Set zlReport = Nothing
End Sub



Private Sub picMaterial_Resize()
On Error Resume Next
    framMaterial.Left = 0
    framMaterial.Top = 0
    framMaterial.Width = picMaterial.Width
    framMaterial.Height = picMaterial.Height
    
    txtPos.Height = IIf(mbytFontSize = 9, 315, 330)
    txtPos.Top = framMaterial.Height - txtPos.Height - 120 + IIf(mbytFontSize = 9, 0, 30)
    
    
    ufgData.Left = 120
    ufgData.Top = 240 + IIf(mbytFontSize = 9, 0, 120)
    ufgData.Width = framMaterial.Width - 240
    ufgData.Height = framMaterial.Height - txtPos.Height - 480

    labRecordInf.Left = 120
    labRecordInf.Top = txtPos.Top + 60
    
    labInf.Left = labRecordInf.Left + labRecordInf.Width + 400
    labInf.Top = labRecordInf.Top
    
    txtPos.Left = labInf.Left + labInf.Width
    
    labSpecimenProcess.Left = txtPos.Left + txtPos.Width + 900
    labSpecimenProcess.Top = labInf.Top
    
    cbxSpecimenProcess.Left = labSpecimenProcess.Left + labSpecimenProcess.Width - 150
    cbxSpecimenProcess.Top = txtPos.Top
    
    cmdAutoInputMaterials.Left = cbxSpecimenProcess.Left + cbxSpecimenProcess.Width + 400
    cmdAutoInputMaterials.Top = cbxSpecimenProcess.Top
End Sub


Private Sub picPane1_Resize()
On Error Resume Next

    tsFilter.Left = 0
    tsFilter.Top = 0
    tsFilter.Width = picPane1.Width

    '�Ѹƹ���----------------------------------------------------
    framDecalin.Left = 0
    framDecalin.Top = tsFilter.Top + tsFilter.Height + 10
    framDecalin.Width = picPane1.Width
    framDecalin.Height = picPane1.Height - tsFilter.Height - 120

    ufgDecalin.Left = 120
    ufgDecalin.Top = 240
    ufgDecalin.Width = framDecalin.Width - 240
    ufgDecalin.Height = framDecalin.Height - cmdSucceed.Height - 480

    cmdSucceed.Left = framDecalin.Width - cmdSucceed.Width - 120
    cmdSucceed.Top = framDecalin.Height - cmdSucceed.Height - 120

    cmdCancel.Left = cmdSucceed.Left - cmdCancel.Width - 120
    cmdCancel.Top = cmdSucceed.Top

    cmdChange.Left = cmdCancel.Left - cmdChange.Width - 120
    cmdChange.Top = cmdSucceed.Top

    cmdDecalin.Left = cmdChange.Left - cmdDecalin.Width - 120
    cmdDecalin.Top = cmdSucceed.Top



    '�޼�����----------------------------------------------------
    framWordEdit.Left = 0
    framWordEdit.Top = tsFilter.Top + tsFilter.Height + 10
    framWordEdit.Width = picPane1.Width
    framWordEdit.Height = picPane1.Height - tsFilter.Height - 120


    wtDescription.Left = 0
    wtDescription.Top = 0
    wtDescription.Width = framWordEdit.Width
    wtDescription.Height = framWordEdit.Height
End Sub

Private Sub tsFilter_Click(PreviousTab As Integer)
On Error GoTo errHandle
    Call SwitchWork(IIf(tsFilter.Tab = 0, True, False))
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnChangeEdit()
'����Ǳ걾���Ƹı� ��ִ�������¸��༭��
    If ufgData.DataGrid.Col = ufgData.GetColIndex(gstrMaterial_�걾����) Then
        ufgData.EditNextCellWithCurRow
    End If

End Sub


Private Sub ufgData_OnAfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim iCol As Long
    Dim i As Long
    Dim iCount As Long
    Dim strNewSpecimenName As String
    Dim strSql As String
    Dim rsInspectionSpe As ADODB.Recordset
    Dim rsMaterislType As ADODB.Recordset
    
    '�жϸ����ǲ��ǻ�����ղ�����ȷ�ϣ�����ִ��ɾ��
    If ufgData.DataGrid.Cell(flexcpTextDisplay, Row, ufgData.GetColIndex(gstrMaterial_ȷ��״̬)) <> "��ȷ��" And ufgData.DataGrid.Cell(flexcpTextDisplay, Row, ufgData.GetColIndex(gstrMaterial_ȡ������)) <> "�������" Then
        '�жϸ����Ƿ����أ������ʾ�����Ա�ɾ�� ��ִ��ɾ��*
        If Not ufgData.DataGrid.RowHidden(Row) Then
            strSql = "select ������� from ����걾��Ϣ where �걾����=[1] and ҽ��ID=[2]"
            Set rsMaterislType = zlDatabase.OpenSQLRecord(strSql, Me.Caption, ufgData.DataGrid.Cell(flexcpTextDisplay, Row, ufgData.GetColIndex(gstrMaterial_�걾����)), mlngAdviceID)
            '�ж����ݼ��Ƿ������ݣ�û����ִ��ɾ��
            If rsMaterislType.RecordCount > 0 Then
                '�жϲ�������Ƿ�Ϊ���� ����ɾ��������
                If Nvl(rsMaterislType!�������) = 1 Then
                    Call MsgBoxD(Me, "�ñ걾�Ĳ�������Ϊ���飬���ܽ���ȡ�Ĳ�����", vbOKOnly, Me.Caption)
    
                    'ɾ��������
                    Call ufgData.DelCurRow
                    Exit Sub
                End If
            End If
        End If
    End If

    If ufgData.IsNullRow(Row) Then
        ufgData.RowState(Row) = TDataRowState.Normal
        
        Call ufgData.SetRowColor(Row, ufgData.BackColor)
        Exit Sub
    End If
    
    If ufgData.Text(Row, gstrMaterial_��Ƭ��) = "" Or ufgData.Text(Row, gstrMaterial_�걾��) = "" Then
        '�����Ƭ�����߱걾����һ��Ϊ����ִ�в�ѯ
        strSql = "select �걾����,Ĭ�ϱ걾��,Ĭ����Ƭ�� from ������걾"
        Set rsInspectionSpe = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    End If
    
    
      '�����Ƭ��Ϊ����ִ�ж�ȡ����  ��Ϊ��������
    If ufgData.Text(Row, gstrMaterial_��Ƭ��) = "" Then
        If rsInspectionSpe.RecordCount > 0 Then
            rsInspectionSpe.MoveFirst
            Do While Not rsInspectionSpe.EOF
                '�����ʾ�ı걾������ ������걾�е�ƥ�����ȡĬ����Ƭ��
                If ufgData.DataGrid.Cell(flexcpTextDisplay, Row, ufgData.GetColIndex(gstrMaterial_�걾����)) = Nvl(rsInspectionSpe!�걾����) Then
                    ufgData.Text(Row, gstrMaterial_��Ƭ��) = Nvl(rsInspectionSpe!Ĭ����Ƭ��)
                    Exit Do
                Else
                    ufgData.Text(Row, gstrMaterial_��Ƭ��) = 1
                End If
                rsInspectionSpe.MoveNext
            Loop
        End If
    End If

    
    Select Case mrecStudy.lngStudyType
        Case StudyType.stNormal, StudyType.stMeet, StudyType.stAutopsy, StudyType.stSpeed  '����,����,ʬ����
        
            If Val(ufgData.Text(Row, gstrMaterial_������)) < 1 And ufgData.Text(Row, gstrMaterial_�Ƿ�����) = "1-��" Then
                ufgData.Text(Row, gstrMaterial_������) = "1"
            End If
        
            '�����������ʾ����������������������¼�룬�����޸ĵ�Ԫ����ɫ
            If Not ufgData.DataGrid.ColHidden(ufgData.GetColIndex(gstrMaterial_������)) Then
                
                '���δ¼����������������ʾ����ɫ
                iCol = ufgData.GetColIndex(gstrMaterial_������)
                
                ufgData.CellColor(Row, iCol) = IIf(Val(ufgData.Text(Row, gstrMaterial_������)) <= 0, ufgData.ErrCellColor, ufgData.BackColor)
            End If
            
        Case StudyType.stIce  '�������
        
            '���������Ϊ����Ĭ�ϵ���0  ��Ϊ��������
            If ufgData.Text(Row, gstrMaterial_������) = "" Then
                ufgData.Text(Row, gstrMaterial_������) = "0"
            End If
        
            If Val(ufgData.Text(Row, gstrMaterial_������)) < 1 And ufgData.Text(Row, gstrMaterial_�Ƿ�����) = "1-��" Then
                ufgData.Text(Row, gstrMaterial_������) = "1"
            End If
            
            '�����������ʾ����������������������¼�룬�����޸ĵ�Ԫ����ɫ
            If Not ufgData.DataGrid.ColHidden(ufgData.GetColIndex(gstrMaterial_������)) Then
                If InStr(ufgData.Text(Row, gstrMaterial_�Ƿ����), "��") > 0 Then
                
                    '���δ¼����������������ʾ����ɫ(�������Ҳ��Ҫ¼����������ȷ����Ƭ����)
                    iCol = ufgData.GetColIndex(gstrMaterial_������)
                    
                    ufgData.CellColor(Row, iCol) = IIf(Val(ufgData.Text(Row, gstrMaterial_������)) <= 0, ufgData.ErrCellColor, ufgData.BackColor)
                Else
                    iCol = ufgData.GetColIndex(gstrMaterial_������)
                    
                    ufgData.CellColor(Row, iCol) = ufgData.BackColor
                End If
            End If
            
            
            '�����������ʾ�˱��࣬��������¼�룬�����޸ĵ�Ԫ����ɫ
            If Not ufgData.DataGrid.ColHidden(ufgData.GetColIndex(gstrMaterial_�Ƿ����)) Then
                '���δ¼���Ƿ���࣬����ʾ����ɫ
                iCol = ufgData.GetColIndex(gstrMaterial_�Ƿ����)
            
                ufgData.CellColor(Row, iCol) = IIf(ufgData.Text(Row, gstrMaterial_�Ƿ����) = "", ufgData.ErrCellColor, ufgData.BackColor)
            End If
        
        Case StudyType.stCell  'ϸ�����
        
            '���ϸ������Ϊ����Ĭ�ϵ���0  ��Ϊ��������
            If ufgData.Text(Row, gstrMaterial_ϸ������) = "" Then
                ufgData.Text(Row, gstrMaterial_ϸ������) = "0"
            End If
            
            If Val(ufgData.Text(Row, gstrMaterial_ϸ������)) < 1 And ufgData.Text(Row, gstrMaterial_�Ƿ�����) = "1-��" Then
                ufgData.Text(Row, gstrMaterial_ϸ������) = "1"
            End If
            
            
             '����걾��Ϊ����ִ�ж�ȡ����  ��Ϊ��������
            If ufgData.Text(Row, gstrMaterial_�걾��) = "" Then
                If rsInspectionSpe.RecordCount > 0 Then
                    rsInspectionSpe.MoveFirst
                    Do While Not rsInspectionSpe.EOF
                        '�����ʾ�ı걾������ ������걾�е�ƥ�����ȡĬ����Ƭ��
                        If ufgData.DataGrid.Cell(flexcpTextDisplay, Row, ufgData.GetColIndex(gstrMaterial_�걾����)) = rsInspectionSpe("�걾����").value Then
                            ufgData.Text(Row, gstrMaterial_�걾��) = rsInspectionSpe("Ĭ�ϱ걾��").value
                            Exit Do
                        End If
                        rsInspectionSpe.MoveNext
                    Loop
                End If
            End If
            
        
            '���δ¼��걾��������ʾ����ɫ
            iCol = ufgData.GetColIndex(gstrMaterial_�걾��)
            
            ufgData.CellColor(Row, iCol) = IIf(Val(ufgData.Text(Row, gstrMaterial_�걾��)) <= 0, ufgData.ErrCellColor, ufgData.BackColor)
                    
    End Select
    
    
    '���δ¼��걾���ƣ�����ʾ����ɫ
    iCol = ufgData.GetColIndex(gstrMaterial_�걾����)

    ufgData.CellColor(Row, iCol) = IIf(ufgData.Text(Row, gstrMaterial_�걾����) = "", ufgData.ErrCellColor, ufgData.BackColor)
                 
    
    '���δ¼����ȡҽʦ������ʾ����ɫ
    iCol = ufgData.GetColIndex(gstrMaterial_��ȡҽʦ)
    
    ufgData.CellColor(Row, iCol) = IIf(ufgData.Text(Row, gstrMaterial_��ȡҽʦ) = "", ufgData.ErrCellColor, ufgData.BackColor)
    
    
    '���δ¼����ȡҽʦ������ʾ����ɫ
    iCol = ufgData.GetColIndex(gstrMaterial_ȡ��ʱ��)
    
    ufgData.CellColor(Row, iCol) = IIf(Not IsDate(ufgData.Text(Row, gstrMaterial_ȡ��ʱ��)), ufgData.ErrCellColor, ufgData.BackColor)

    '���Ϊ�������Ƭ��������༭
    If mrecStudy.lngStudyType = stMeet Then
        ufgData.Text(Row, gstrMaterial_��Ƭ��) = ""
    End If

       
End Sub

Private Sub ufgData_OnStartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim dtServicesTime As Date
    Dim strComboboxText As String
    Dim strSql As String
    Dim rsInspectionSpe As ADODB.Recordset

    
    '�ж��Ƿ���������µ�ȡ��
    If ufgData.IsNullRow(Row) Then
        If mrecStudy.lngMaterialStep <> TExecuteStep.NeedDo Then
            Cancel = True
            Call MsgBoxD(Me, "��ȡ�Ľ׶Σ����ܽ����µ�ȡ�Ĳ����������롣", vbOKOnly, Me.Caption)
        
            Exit Sub
        End If
    End If

    
    '�ж��Ƿ��������
    If Not CheckAllowUpdate(ufgData.KeyValue(Row)) Then
        Cancel = True
        Call MsgBoxD(Me, "�òĿ��¼��ִ����Ƭ�������ܽ��и��¡�", vbOKOnly, Me.Caption)
        
        Exit Sub
    End If
    
    
    If Not IsDate(ufgData.Text(Row, gstrMaterial_ȡ��ʱ��)) Then
        ufgData.Text(Row, gstrMaterial_ȡ��ʱ��) = zlDatabase.Currentdate
    End If
    
    
    If Row > 0 Then
    
            '���ر걾����
            'If Col = ufgData.vfgHelper.GetColumnIndex(gstrMaterial_�걾����) Then
                If ufgData.Text(Row, gstrMaterial_�걾����) = "" Then
                    If Row - 1 > 0 Then
                        If ufgData.Text(Row - 1, gstrMaterial_�걾����) <> "" Then
                            ufgData.Text(Row, gstrMaterial_�걾����) = ufgData.Text(Row - 1, gstrMaterial_�걾����)
                        End If
                    End If
                    
                    If ufgData.Text(Row, gstrMaterial_�걾����) = "" Then
                        strComboboxText = ufgData.DataGrid.ColComboList(ufgData.GetColIndex(gstrMaterial_�걾����))
                        
                        If strComboboxText <> "" Then
                            If InStr(strComboboxText, ";") > 0 Then
                                strComboboxText = Mid(strComboboxText, 1, InStr(strComboboxText, ";") - 1)
                            End If
                            ufgData.Text(Row, gstrMaterial_�걾����) = Mid(strComboboxText, InStr(strComboboxText, "#") + 1, 255)
                            
                        End If
                    End If
                End If
    
    
        If mrecStudy.lngStudyType = StudyType.stIce Then
        
            If Not ufgData.DataGrid.ColHidden(ufgData.GetColIndex(gstrMaterial_�Ƿ����)) Then
                If ufgData.Text(Row, gstrMaterial_�Ƿ����) = "" Then ufgData.Text(Row, gstrMaterial_�Ƿ����) = "0-��"
            End If
            
            
            
            '������Ǳ��࣬��������������б༭
            'If Col = ufgData.vfgHelper.GetColumnIndex(gstrMaterial_������) Then
            If Not ufgData.DataGrid.ColHidden(ufgData.GetColIndex(gstrMaterial_������)) Then
                If InStr(ufgData.Text(Row, gstrMaterial_�Ƿ����), "��") > 0 Then
                    If Val(ufgData.Text(Row, gstrMaterial_������)) <= 0 Then ufgData.Text(Row, gstrMaterial_������) = "1"
                Else
                    If Col = ufgData.GetColIndex(gstrMaterial_������) Then Cancel = True
                End If
            End If
            'End If
            If ufgData.Text(Row, gstrMaterial_�Ƿ�����) = "" Then ufgData.Text(Row, gstrMaterial_�Ƿ�����) = "0-��"
            
        ElseIf mrecStudy.lngStudyType = StudyType.stCell Then
            'If Col = ufgData.vfgHelper.GetColumnIndex(gstrMaterial_ϸ������) Then
            'ufgData.Text(Row, gstrMaterial_ϸ������) = "0"
            
            If UCase(ufgData.DisplayText(Row, gstrMaterial_�걾����)) = UCase("TCT") Then
                ufgData.Text(Row, gstrMaterial_�걾��) = "20ml"
            ElseIf ufgData.DisplayText(Row, gstrMaterial_�걾����) Like "*̵*" Then
                ufgData.Text(Row, gstrMaterial_�걾��) = "1.5ml"
            End If
            
            If ufgData.Text(Row, gstrMaterial_�Ƿ�����) = "" Then ufgData.Text(Row, gstrMaterial_�Ƿ�����) = "0-��"
        Else
            '����������
            'If Col = ufgData.vfgHelper.GetColumnIndex(gstrMaterial_������) Then
                If Not ufgData.DataGrid.ColHidden(ufgData.GetColIndex(gstrMaterial_������)) Then
                    If Val(ufgData.Text(Row, gstrMaterial_������)) <= 0 Then ufgData.Text(Row, gstrMaterial_������) = "1"
                End If
                
                If ufgData.Text(Row, gstrMaterial_�Ƿ�����) = "" Then ufgData.Text(Row, gstrMaterial_�Ƿ�����) = "1-��"
       
            
            'End If
        End If
        

        
        'If Col = ufgData.vfgHelper.GetColumnIndex(gstrMaterial_��ȡҽʦ) Then
            If ufgData.Text(Row, gstrMaterial_��ȡҽʦ) = "" Then
                If Row - 1 > 0 Then
                    If ufgData.Text(Row - 1, gstrMaterial_��ȡҽʦ) <> "" Then
                        ufgData.Text(Row, gstrMaterial_��ȡҽʦ) = ufgData.Text(Row - 1, gstrMaterial_��ȡҽʦ)
                    End If
                End If
                
                If ufgData.Text(Row, gstrMaterial_��ȡҽʦ) = "" Then
                    strComboboxText = ufgData.DataGrid.ColComboList(ufgData.GetColIndex(gstrMaterial_��ȡҽʦ))
                    
                    If strComboboxText <> "" Then
                        If InStr(strComboboxText, "|") > 0 Then
                            strComboboxText = Mid(strComboboxText, 1, InStr(strComboboxText, "|") - 1)
                        End If
                        ufgData.Text(Row, gstrMaterial_��ȡҽʦ) = strComboboxText
                        
                    End If
                End If
            End If
        'End If

         
    End If
End Sub

Private Sub ufgDecalin_OnAfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim iCol As Long
    Dim i As Long
    Dim iCount As Long
    Dim strNewSpecimenName As String

    If ufgDecalin.IsNullRow(Row) Then
        ufgDecalin.RowState(Row) = TDataRowState.Normal
        Call ufgDecalin.SetRowColor(Row, ufgDecalin.BackColor)
        
        Exit Sub
    End If



    '���δ¼��걾���ƣ�����ʾ����ɫ
    iCol = ufgDecalin.GetColIndex(gstrDecalin_�걾����)
    
    ufgDecalin.CellColor(Row, iCol) = IIf(ufgDecalin.Text(Row, gstrDecalin_�걾����) = "", ufgDecalin.ErrCellColor, ufgDecalin.BackColor)
           


    '���δ¼�뿪ʼʱ�䣬����ʾ����ɫ
    iCol = ufgDecalin.GetColIndex(gstrDecalin_��ʼʱ��)
    
    ufgDecalin.CellColor(Row, iCol) = IIf(Not IsDate(ufgDecalin.Text(Row, gstrDecalin_��ʼʱ��)), ufgDecalin.ErrCellColor, ufgDecalin.BackColor)
          
    

    '���δ¼������ʱ��������ʾ����ɫ
    iCol = ufgDecalin.GetColIndex(gstrDecalin_����ʱ��)
    
    ufgDecalin.CellColor(Row, iCol) = IIf(Val(ufgDecalin.Text(Row, gstrDecalin_����ʱ��)) <= 0, ufgDecalin.ErrCellColor, ufgDecalin.BackColor)
    
    
    '���δ¼�����Ա������ʾ����ɫ
    iCol = ufgDecalin.GetColIndex(gstrDecalin_����Ա)
    
    ufgDecalin.CellColor(Row, iCol) = IIf(ufgDecalin.Text(Row, gstrDecalin_����Ա) = "", ufgDecalin.ErrCellColor, ufgDecalin.BackColor)
    
End Sub


Private Sub ConfigDecalcificationBut()
'�����Ѹư�ť
    If Not ufgDecalin.IsSelectionRow Then
        cmdDecalin.Enabled = False
        cmdChange.Enabled = False
        cmdCancel.Enabled = False
        cmdSucceed.Enabled = False
    End If
    
    cmdDecalin.Enabled = Not ufgDecalin.IsNullRow(ufgDecalin.SelectionRow) And ufgDecalin.IsEmptyKey(ufgDecalin.SelectionRow)
    cmdChange.Enabled = Not ufgDecalin.IsEmptyKey(ufgDecalin.SelectionRow)
    cmdCancel.Enabled = Not ufgDecalin.IsNullRow(ufgDecalin.SelectionRow)
    cmdSucceed.Enabled = Not ufgDecalin.IsEmptyKey(ufgDecalin.SelectionRow)
End Sub



Private Sub ufgDecalin_OnClick()
On Error GoTo errHandle
    Call ConfigDecalcificationBut
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub






Private Sub ufgDecalin_OnStartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim dtServicesTime As Date
    
    If Col = ufgDecalin.GetColIndex(gstrDecalin_��ʼʱ��) And Row > 0 Then
        
        dtServicesTime = zlDatabase.Currentdate
        ufgDecalin.Text(Row, gstrDecalin_��ʼʱ��) = dtServicesTime
        
        Exit Sub
    End If
End Sub


