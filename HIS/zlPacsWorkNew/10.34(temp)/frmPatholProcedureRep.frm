VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPatholProcedureRep 
   Caption         =   "�����ؼ챨��"
   ClientHeight    =   10035
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   11385
   Icon            =   "frmPatholProcedureRep.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10035
   ScaleWidth      =   11385
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ImageList imgTbrS 
      Left            =   9045
      Top             =   540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholProcedureRep.frx":000C
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholProcedureRep.frx":0C7E
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholProcedureRep.frx":18F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholProcedureRep.frx":2562
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholProcedureRep.frx":31D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholProcedureRep.frx":3E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholProcedureRep.frx":4AB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholProcedureRep.frx":572A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholProcedureRep.frx":639C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   8040
      Left            =   30
      ScaleHeight     =   8040
      ScaleWidth      =   11325
      TabIndex        =   1
      Top             =   780
      Width           =   11325
      Begin zl9PacsControl.ucSplitter ucSplitter1 
         Height          =   8040
         Left            =   3255
         TabIndex        =   21
         Top             =   0
         Width           =   85
         _ExtentX        =   159
         _ExtentY        =   14182
         BackColor       =   -2147483633
         SplitWidth      =   85
         SplitLevel      =   3
         Control1Name    =   "picWordModule"
         Control2Name    =   "picReportEdit"
      End
      Begin VB.PictureBox picWordModule 
         BorderStyle     =   0  'None
         Height          =   8040
         Left            =   0
         ScaleHeight     =   8040
         ScaleWidth      =   3255
         TabIndex        =   17
         Top             =   0
         Width           =   3255
         Begin VB.Frame framWord 
            Height          =   7215
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   3255
            Begin zl9PACSWork.WordInputModule wimWord 
               Height          =   4335
               Left            =   120
               TabIndex        =   19
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   7646
               CurDepartId     =   0
            End
            Begin zl9PACSWork.ucFlexGrid ufgData 
               Height          =   2415
               Left            =   120
               TabIndex        =   20
               Top             =   4680
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   4260
               GridRows        =   21
               BackColor       =   12648447
               IsCopyAdoMode   =   0   'False
               IsEjectConfig   =   -1  'True
               Editable        =   0
               HeadFontCharset =   134
               HeadFontWeight  =   400
               DataFontCharset =   134
               DataFontWeight  =   400
            End
         End
      End
      Begin VB.PictureBox picReportEdit 
         BorderStyle     =   0  'None
         Height          =   8040
         Left            =   3340
         ScaleHeight     =   8040
         ScaleWidth      =   7980
         TabIndex        =   2
         Top             =   0
         Width           =   7985
         Begin VB.Frame framReport 
            Height          =   7455
            Left            =   45
            TabIndex        =   3
            Top             =   0
            Width           =   7695
            Begin VB.ComboBox cbxReportType 
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
               ItemData        =   "frmPatholProcedureRep.frx":700E
               Left            =   1020
               List            =   "frmPatholProcedureRep.frx":7010
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   240
               Width           =   1545
            End
            Begin VB.ComboBox cbxSpecimenName 
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
               Left            =   6000
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   240
               Width           =   1545
            End
            Begin VB.ComboBox cbxReportSub 
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
               Left            =   3600
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   240
               Width           =   2025
            End
            Begin RichTextLib.RichTextBox txtAdvice 
               Height          =   1815
               Left            =   120
               TabIndex        =   4
               Top             =   3360
               Width           =   7335
               _ExtentX        =   12938
               _ExtentY        =   3201
               _Version        =   393217
               BorderStyle     =   0
               Enabled         =   -1  'True
               ScrollBars      =   2
               AutoVerbMenu    =   -1  'True
               TextRTF         =   $"frmPatholProcedureRep.frx":7012
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin RichTextLib.RichTextBox txtResult 
               Height          =   2055
               Left            =   120
               TabIndex        =   6
               Top             =   960
               Width           =   7335
               _ExtentX        =   12938
               _ExtentY        =   3625
               _Version        =   393217
               BorderStyle     =   0
               ScrollBars      =   2
               AutoVerbMenu    =   -1  'True
               TextRTF         =   $"frmPatholProcedureRep.frx":70AF
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin zl9PACSWork.ReportImage rpImage 
               Height          =   1935
               Left            =   120
               TabIndex        =   7
               Top             =   5280
               Width           =   7335
               _ExtentX        =   12938
               _ExtentY        =   3413
               ShowPhotoCount  =   3
               BackColor       =   4210752
            End
            Begin MSComCtl2.DTPicker dtpReportTime 
               Height          =   300
               Left            =   5640
               TabIndex        =   10
               Top             =   3050
               Width           =   1785
               _ExtentX        =   3149
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd HH:mm"
               Format          =   94044163
               CurrentDate     =   40646.4399652778
            End
            Begin VB.Label labReportType 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "�������ͣ�"
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   120
               TabIndex        =   16
               Top             =   300
               Width           =   900
            End
            Begin VB.Label labSpecimenName 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "�걾����"
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   5280
               TabIndex        =   15
               Top             =   300
               Width           =   720
            End
            Begin VB.Line Line1 
               X1              =   110
               X2              =   7440
               Y1              =   650
               Y2              =   650
            End
            Begin VB.Label labResult 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "�������"
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   120
               TabIndex        =   14
               Top             =   720
               Width           =   900
            End
            Begin VB.Label labAdvice 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "������ϣ�"
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   120
               TabIndex        =   13
               Top             =   3120
               Width           =   900
            End
            Begin VB.Label labReportSub 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "�������"
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   2640
               TabIndex        =   12
               Top             =   300
               Width           =   900
            End
            Begin VB.Label labReportTime 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "����ʱ�䣺"
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   4800
               TabIndex        =   11
               Top             =   3090
               Width           =   900
            End
         End
      End
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   1270
      ButtonWidth     =   1455
      ButtonHeight    =   1270
      Style           =   1
      ImageList       =   "imgTbrS"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�����ӡ"
            Key             =   "tbReport"
            ImageIndex      =   2
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbReportPreview"
                  Text            =   "Ԥ��"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbReportPrint"
                  Text            =   "��ӡ"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�������"
            Key             =   "tbAuditing"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "���泷��"
            Key             =   "tbCancelReport"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "-"
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�������"
            Key             =   "tbClearContext"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��Ŀ¼��"
            Key             =   "tbInputProject"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��������"
            Key             =   "tbNewReport"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ɾ������"
            Key             =   "tbDelReport"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "���汨��"
            Key             =   "tbSaveReport"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPatholProcedureRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IWorkMenu

Private Const M_REPORTSTATE_NORMAL As Long = 0  'δ��ӡ
Private Const M_REPORTSTATE_VIEW As Long = 1    '����
Private Const M_REPORTSTATE_CANCEL As Long = 2  '�ѳ���
Private Const M_REPORTSTATE_PRINT As Long = 3   '�Ѵ�ӡ

Private Const M_STR_HINT_NoSelectData As String = "��Ч�ļ�����ݣ���ѡ����Ҫִ�еļ���¼��"
Private Const M_STR_MODULE_MENU_TAG As String = "����"


Dim WithEvents zlReport As zl9Report.clsReport
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
Private mrecStudyInf As TStudyStateInf

Private mlngTmpAdviceId As Long
Private mlngTmpSendNo As Long


Private mSelMiniImg As DicomImage

Private mintShowPhotoNumber As Integer
Private mintCurImgIndex As Integer
Private strCurTempReportPath As String
Private mblnEditState As Boolean

Private mCurEditText As RichTextBox


Private mblnIsAllowSpeExam As Boolean
Private mblnIsAllowWriteReport As Boolean

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


Private Sub Form_Resize()
On Error GoTo errHandle
    picBack.Left = 0
    picBack.Top = tbrMain.Top + tbrMain.Height
    picBack.Width = Me.ScaleWidth
    picBack.Height = Me.ScaleHeight - tbrMain.Height - tbrMain.Top

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


    If Not HasMenu(objMenuBar, conMenu_PatholProRep) Then
        Set cbrMenuBar = mObjActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_PatholProRep, "���̱���(&O)", GetPatholMenuIndex(objMenuBar) + 1, False)
        cbrMenuBar.ID = conMenu_PatholProRep
        cbrMenuBar.Category = M_STR_MODULE_MENU_TAG
                
        With cbrMenuBar.CommandBar
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholProRep_Preview, "����Ԥ��(&V)", "", 1, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholProRep_Print, "�����ӡ(&P)", "", 1, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholProRep_Already, "�������(&A)", "", 1, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholProRep_Back, "���泷��(&C)", "", 1, False)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholProRep_Clear, "�������(&R)", "", 1, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholProRep_Input, "�ؼ���Ŀ¼��(&I)", "", 1, False)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholProRep_New, "��������(&N)", "", 1, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholProRep_Del, "ɾ������(&D)", "", 1, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholProRep_Save, "���汨��(&S)", "", 1, False)
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
        Case conMenu_PatholProRep_Preview
            Call PrintCurProcedureRep(False)
            
        Case conMenu_PatholProRep_Print
            Call PrintCurProcedureRep(True)
            
        Case conMenu_PatholProRep_Already
            Call UpdateCurProcedureRepState(M_REPORTSTATE_VIEW)
            
        Case conMenu_PatholProRep_Back
            Call UpdateCurProcedureRepState(M_REPORTSTATE_CANCEL)
            
        Case conMenu_PatholProRep_Clear
            Call ClearReportContext
            
        Case conMenu_PatholProRep_Input
            Call GetSpeExamResult
            
        Case conMenu_PatholProRep_New
            Call NewProcedureRep
            
        Case conMenu_PatholProRep_Del
            Call DelCurProcedureRep
            
        Case conMenu_PatholProRep_Save
            Call SaveCurProcedureRep
    End Select
End Sub


Public Sub IWorkMenu_zlUpdateMenu(ByVal control As XtremeCommandBars.ICommandBarControl)
'���²˵�
    Dim blnIsPopedom As Boolean
    
    If Not Me.Visible Then
        control.Enabled = False
        Exit Sub
    End If
    
    Select Case Val(cbxReportType.Text)
        Case 0:
            blnIsPopedom = CheckPopedom(mstrPrivs, "��������")
        Case 1
            blnIsPopedom = CheckPopedom(mstrPrivs, "���߱���")
        Case 2, 3
            blnIsPopedom = CheckPopedom(mstrPrivs, "���ӱ���")
        Case 4
            blnIsPopedom = CheckPopedom(mstrPrivs, "��Ⱦ����")
    End Select
    
    
    
    Select Case control.ID
        Case conMenu_PatholProRep_Preview
            control.Enabled = (mblnIsAllowSpeExam Or mblnIsAllowWriteReport) And mlngAdviceID > 0
            
        Case conMenu_PatholProRep_Print
            control.Enabled = (mblnIsAllowSpeExam Or mblnIsAllowWriteReport) And mlngAdviceID > 0
            
        Case conMenu_PatholProRep_Already
            control.Enabled = mblnIsAllowWriteReport And Not mblnReadOnly And mlngAdviceID > 0
            
        Case conMenu_PatholProRep_Back
            control.Enabled = mblnIsAllowWriteReport And Not mblnReadOnly And mlngAdviceID > 0
            
        Case conMenu_PatholProRep_Input
            control.Enabled = Not (mblnReadOnly Or GetCurRepAllowAuditing) And blnIsPopedom And Val(cbxReportType.Text) > 0 And mlngAdviceID > 0
            
        Case conMenu_PatholProRep_Clear
            control.Enabled = Not (mblnReadOnly Or GetCurRepAllowAuditing) And blnIsPopedom And mlngAdviceID > 0
            
        Case conMenu_PatholProRep_New
            control.Enabled = Not (mblnReadOnly Or GetCurRepAllowAuditing) And blnIsPopedom And mlngAdviceID > 0
            
        Case conMenu_PatholProRep_Del
            control.Enabled = Not (mblnReadOnly Or GetCurRepAllowAuditing) And blnIsPopedom And mlngAdviceID > 0
            
        Case conMenu_PatholProRep_Save
            control.Enabled = Not (mblnReadOnly Or GetCurRepAllowAuditing) And blnIsPopedom And mlngAdviceID > 0
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
    
    Call ClearReportContext
    
    If mlngAdviceID <= 0 Then
        Call ConfigProcedureReportFace(False, "ҽ��ID��Ч���顣")
        Exit Sub
    End If
    
    Call LoadReportModule
        
    mblnIsAllowSpeExam = CheckPopedom(mstrPrivs, "��������") Or CheckPopedom(mstrPrivs, "���߱���") Or CheckPopedom(mstrPrivs, "���ӱ���") Or CheckPopedom(mstrPrivs, "��Ⱦ����")
    mblnIsAllowWriteReport = CheckPopedom(mstrPrivs, "�����ؼ챨�����")
    
    
    Call GetPatholStudyState(mlngAdviceID, mrecStudyInf)

    
    If mrecStudyInf.strPatholNumber = "" Then
        Call ConfigProcedureReportFace(False, "�ü����δ������Ч�Ĳ���ţ���ȷ�ϸü���Ƿ��ѱ����ա�")
        Call rpImage.ReInit
        
        Exit Sub
    Else
        Call ConfigReportType
        Call LoadReportSub(Val(cbxReportType.Text))
        Call ConfigProcedureReportFace(True)
        
 
        '���뱨��ͼ��
        Call rpImage.LoadReportImages(mlngAdviceID, mblnMoved, Me)
        '���ñ걾¼��
        Call ConfigSpecimenName(mlngAdviceID)
        '��ȡ���̱����¼
        Call LoadProcedureRepData(mblnReadOnly)
    End If

    Call EnableReportWithSpeExamType(Val(cbxReportType.Text), mblnReadOnly, GetCurRepAllowAuditing)
'    If Not (owner Is Nothing) Then
'        Call Me.Show(1, owner)
'    End If
End Sub


Public Sub zlRefresh(lngAdviceID As Long, ByVal blnReadOnly As Boolean, _
    strPrivs As String, ByVal blnMoved As Boolean, _
    ByVal lngCurDepartmentId As Long, Optional owner As Form = Nothing)
'ˢ������
    Call ClearReportContext
    
        
    If lngAdviceID <= 0 Then
        Call ConfigProcedureReportFace(False, "ҽ��ID��Ч���顣")
        Exit Sub
    End If

'    If mlngCurAdviceId = lngAdviceID Then Exit Sub
    
    mlngAdviceID = lngAdviceID
    mstrPrivs = strPrivs
    mblnMoved = blnMoved
    mlngCurDeptId = lngCurDepartmentId
'    mblnReadOnly = blnReadOnly
    
    Call LoadReportModule
        
    mblnIsAllowSpeExam = CheckPopedom(mstrPrivs, "��������") Or CheckPopedom(mstrPrivs, "���߱���") Or CheckPopedom(mstrPrivs, "���ӱ���") Or CheckPopedom(mstrPrivs, "��Ⱦ����")
    mblnIsAllowWriteReport = CheckPopedom(mstrPrivs, "�����ؼ챨�����")
    
    
    Call GetPatholStudyState(lngAdviceID, mrecStudyInf)

    
    If mrecStudyInf.strPatholNumber = "" Then
        Call ConfigProcedureReportFace(False, "�ü����δ������Ч�Ĳ���ţ���ȷ�ϸü���Ƿ��ѱ����ա�")
        Call rpImage.ReInit
        
        Exit Sub
    Else
        Call ConfigReportType
        Call LoadReportSub(Val(cbxReportType.Text))
        Call ConfigProcedureReportFace(True)
        
 
        '���뱨��ͼ��
        Call rpImage.LoadReportImages(mlngAdviceID, mblnMoved, Me)
        '���ñ걾¼��
        Call ConfigSpecimenName(mlngAdviceID)
        '��ȡ���̱����¼
        Call LoadProcedureRepData(blnReadOnly)
    End If

    Call EnableReportWithSpeExamType(Val(cbxReportType.Text), blnReadOnly, GetCurRepAllowAuditing)

    
    If Not (owner Is Nothing) Then
        Call Me.Show(1, owner)
    End If
End Sub



Private Function GetCurRepAllowAuditing() As Boolean
'�жϵ�ǰ���̱����Ƿ��������
    GetCurRepAllowAuditing = False
    
    If ufgData.ShowingDataRowCount <= 0 Then Exit Function
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then Exit Function
    
    GetCurRepAllowAuditing = ("�Ѵ�ӡ,�Ѳ���" Like "*" & ufgData.Text(ufgData.SelectionRow, gstrProcedureRep_��ǰ״̬) & "*")
End Function



Private Sub ConfigProcedureReportFace(ByVal blnIsValid As Boolean, Optional ByVal strHintInf As String = "")
'�����ؼ����
    Call EnableReportWithSpeExamType(Val(cbxReportType.Text), Not blnIsValid, Not blnIsValid)
    
    If blnIsValid Then
        Call ufgData.CloseHintInf
    Else
'        cmdPrint.Enabled = blnIsValid
'        cbxReportSub.Enabled = blnIsValid
'        dtpReportTime.Enabled = blnIsValid
    
        Call ufgData.ShowHintInf(strHintInf)
    End If
    
    tbrMain.Buttons("tbReport").Enabled = blnIsValid
    cbxReportSub.Enabled = blnIsValid
    dtpReportTime.Enabled = blnIsValid
End Sub



Private Sub InitProcedureRepList()
'��ʼ�����̱����б�
    Dim strTemp As String
    
    ufgData.IsKeepRows = False


     '�ж����ݿ�������Ƿ������� �����ȡ���ݿ����  û�������Ĭ��
    strTemp = zlDatabase.GetPara("���̱����б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
     
    If strTemp = "" Then
        ufgData.ColNames = gstrProcedureRepCols
    Else
        ufgData.ColNames = strTemp
    End If
    
    '��ֹ�Ҽ������б����ô���
    ufgData.IsEjectConfig = False
    '��������
    ufgData.GridRows = glngStandardRowCount
    '�����и�
    ufgData.RowHeightMin = glngStandardRowHeight

    ufgData.DefaultColNames = gstrProcedureRepCols
    ufgData.ColConvertFormat = gstrProcedureRepConvertFormat
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
        Case UCase("tbReport"), UCase("tbReportPreview")
            'Ԥ�����̱���
            Call PrintCurProcedureRep(False)
            
        Case UCase("tbReportPrint")
            '��ӡ���̱���
            Call PrintCurProcedureRep(True)
            
        Case UCase("tbAuditing")
            '���̱������
            Call UpdateCurProcedureRepState(M_REPORTSTATE_VIEW)
            
        Case UCase("tbCancelReport")
            '���ع��̱���
            Call UpdateCurProcedureRepState(M_REPORTSTATE_CANCEL)
            
        Case UCase("tbClearContext")
            '���¼������
            Call ClearReportContext
            
        Case UCase("tbInputProject")
            '��Ŀ¼��
            Call GetSpeExamResult
            
        Case UCase("tbNewReport")
            '��������
            Call NewProcedureRep
            
        Case UCase("tbDelReport")
            'ɾ������
            Call DelCurProcedureRep
            
        Case UCase("tbSaveReport")
            '���汨��
            Call SaveCurProcedureRep
            
    End Select
End Sub

Private Sub ufgData_OnColFormartChange()
'�رմ���ʱ�����б�����
    zlDatabase.SetPara "���̱����б�����", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
End Sub


Private Sub LoadProcedureRepData(ByVal blnReadOnly As Boolean)
'��ȡ���̱�������
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim blnIsAllowWriteReport As Boolean
    
    blnIsAllowWriteReport = CheckPopedom(mstrPrivs, "�����ؼ챨�����")
    
    strSql = "select ID,�걾����,��������,��������,�����,������,����ͼ��,����ҽʦ,��������,��ǰ״̬,��ע from ������̱��� where ����ҽ��ID=[1] and (��������=-1"
    
    If CheckPopedom(mstrPrivs, "��������") Or blnIsAllowWriteReport Then strSql = strSql & " or ��������=0"
    If CheckPopedom(mstrPrivs, "���߱���") Or blnIsAllowWriteReport Then strSql = strSql & " or ��������=1"
    If CheckPopedom(mstrPrivs, "���ӱ���") Or blnIsAllowWriteReport Then strSql = strSql & " or ��������=2"
    If CheckPopedom(mstrPrivs, "��Ⱦ����") Or blnIsAllowWriteReport Then strSql = strSql & " or ��������=3"
    
    strSql = strSql & ")"
    
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mrecStudyInf.lngPatholAdviceId)

    Call ufgData.RefreshData
    
    If ufgData.ShowingDataRowCount >= 1 Then
        Call LoadReportContext(1)
    
        Call EnableReportWithSpeExamType(Val(cbxReportType.Text), blnReadOnly, GetCurRepAllowAuditing)
    End If
End Sub



Private Sub EnableReportWithSpeExamType(ByVal lngSpeExamType As Long, ByVal blnStudyFinal As Boolean, _
    ByVal blnRepAuditing As Boolean, Optional ByVal blnShowHint As Boolean = True)
'���ñ���ı༭״̬
    Dim blnIsPopedom As Boolean
    
    Select Case lngSpeExamType
        Case 0:
            blnIsPopedom = CheckPopedom(mstrPrivs, "��������")
        Case 1
            blnIsPopedom = CheckPopedom(mstrPrivs, "���߱���")
        Case 2, 3
            blnIsPopedom = CheckPopedom(mstrPrivs, "���ӱ���")
        Case 4
            blnIsPopedom = CheckPopedom(mstrPrivs, "��Ⱦ����")
    End Select
    
    tbrMain.Buttons("tbInputProject").Enabled = blnIsPopedom And lngSpeExamType > 0 And Not (blnStudyFinal Or blnRepAuditing)
    
    cbxReportType.Enabled = mblnIsAllowSpeExam And Not blnStudyFinal
    cbxSpecimenName.Enabled = mblnIsAllowSpeExam And Not blnStudyFinal
    
    txtResult.Locked = Not blnIsPopedom Or (blnStudyFinal Or blnRepAuditing)
    txtAdvice.Locked = Not blnIsPopedom Or (blnStudyFinal Or blnRepAuditing)
    
    txtResult.BackColor = IIf(Not (blnStudyFinal Or blnRepAuditing) And blnIsPopedom, vbWhite, Me.BackColor)
    txtAdvice.BackColor = IIf(Not (blnStudyFinal Or blnRepAuditing) And blnIsPopedom, vbWhite, Me.BackColor)
    
    txtResult.Enabled = Not (blnStudyFinal Or blnRepAuditing) And blnIsPopedom
    txtResult.Enabled = Not (blnStudyFinal Or blnRepAuditing) And blnIsPopedom
    
    rpImage.Enable = Not (blnStudyFinal Or blnRepAuditing) And blnIsPopedom
    
    tbrMain.Buttons("tbDelReport").Enabled = Not (blnStudyFinal Or blnRepAuditing) And blnIsPopedom
    tbrMain.Buttons("tbSaveReport").Enabled = Not (blnStudyFinal Or blnRepAuditing) And blnIsPopedom
    tbrMain.Buttons("tbClearContext").Enabled = Not (blnStudyFinal Or blnRepAuditing) And blnIsPopedom
    tbrMain.Buttons("tbNewReport").Enabled = Not (blnStudyFinal) And blnIsPopedom

    tbrMain.Buttons("tbAuditing").Enabled = mblnIsAllowWriteReport And Not (blnStudyFinal) And Not blnRepAuditing
    tbrMain.Buttons("tbCancelReport").Enabled = mblnIsAllowWriteReport And Not (blnStudyFinal) And blnRepAuditing
    
    tbrMain.Buttons("tbReport").Enabled = mblnIsAllowSpeExam Or mblnIsAllowWriteReport
    
    If Not blnIsPopedom And Not mblnIsAllowWriteReport And blnShowHint Then 'And blnOldEditState And Not mblnIsAllowWriteReport
'        Call MsgBoxD(Me, "���߱��༭���౨���Ȩ�ޡ�", vbOKOnly, Me.Caption)
    End If
End Sub


Private Sub LoadReportType()
'���뱨������
    Dim lngIndex As Long
    
    Call cbxReportType.Clear
    
    Call cbxReportType.AddItem("0-��������")
    Call cbxReportType.AddItem("1-���߱���")
    Call cbxReportType.AddItem("2-���ӱ���")
    Call cbxReportType.AddItem("3-��Ⱦ����")
    
    cbxReportType.ListIndex = 1
End Sub


Private Sub LoadReportSub(ByVal lngReportType As Long)
'���뱨������
    cbxReportSub.Clear
    
'    Call cbxReportSub.AddItem("")
    
    If lngReportType = 1 Then
        Call cbxReportSub.AddItem("1-����(����)")
        Call cbxReportSub.AddItem("2-����(��ҩ��ҩ)")
        
        cbxReportSub.ListIndex = 1
    ElseIf lngReportType = 2 Then
        Call cbxReportSub.AddItem("1-����(ӫ��)")  '��Ӧ 3
        Call cbxReportSub.AddItem("2-����(��ͨ)")  '��Ӧ 4
        
        cbxReportSub.ListIndex = 0
    End If
    
End Sub



Public Sub ReSetFormFontSize(ByVal bytFontSize As Byte)
'����:�������ù���վ����������С
On Error GoTo errHandle

    Dim objCtrl As control
    Dim CtlFont As StdFont
    Dim strFontType  As String
    
    
    mbytFontSize = bytFontSize
    
    cbxReportType.Left = cbxReportType.Left + IIf(cbxReportType.FontSize = bytFontSize, 0, IIf(bytFontSize = 9, -100, 100))
    cbxReportSub.Left = cbxReportSub.Left + IIf(cbxReportSub.FontSize = bytFontSize, 0, IIf(bytFontSize = 9, -100, 100))
    cbxSpecimenName.Left = cbxSpecimenName.Left + IIf(cbxSpecimenName.FontSize = bytFontSize, 0, IIf(bytFontSize = 9, -100, 100))
    
    
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
            objCtrl.Font.Name = strFontType
            objCtrl.Width = TextWidth("2012-01-01 23:59:59") * 1.25
            objCtrl.Height = TextHeight("��") * 1.7
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
    
    Call picReportEdit_Resize
    
    Exit Sub
errHandle:
End Sub






Private Sub ConfigReportType()
'���õ�ǰ���ı�������
    '�жϼ�����ͣ�����Ǳ�����飬��Ĭ������Ϊ������鱨�棬
    If mrecStudyInf.lngStudyType = 1 Then
        If CheckPopedom(mstrPrivs, "��������") Then
            cbxReportType.ListIndex = 0
            Exit Sub
        End If
    End If
    
    
    If mrecStudyInf.lngMianYiStep > 0 Then
        If CheckPopedom(mstrPrivs, "���߱���") Then
            cbxReportType.ListIndex = 1
            Exit Sub
        End If
    End If
 
    
    If mrecStudyInf.lngTeRanStep > 0 Then
        If CheckPopedom(mstrPrivs, "��Ⱦ����") Then
            cbxReportType.ListIndex = 3
            Exit Sub
        End If
    End If
    
    If mrecStudyInf.lngFenZiStep > 0 Then
        If CheckPopedom(mstrPrivs, "���ӱ���") Then
            cbxReportType.ListIndex = 2
            Exit Sub
        End If
    End If
    
    
    
    
    If CheckPopedom(mstrPrivs, "���߱���") Then
        cbxReportType.ListIndex = 1
        Exit Sub
    End If
    
    
    If CheckPopedom(mstrPrivs, "���ӱ���") Then
        cbxReportType.ListIndex = 2
        Exit Sub
    End If
    
    
    If CheckPopedom(mstrPrivs, "��Ⱦ����") Then
        cbxReportType.ListIndex = 3
        Exit Sub
    End If
End Sub


Private Sub ConfigSpecimenName(ByVal lngAdviceID As String)
'��ȡ�걾����
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select �걾���� from ����걾��Ϣ where ҽ��ID=[1] and �ͼ�ID > 0"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceID)
    
    Call cbxSpecimenName.Clear
    
    If rsData.RecordCount < 0 Then Exit Sub

    Call cbxSpecimenName.AddItem("")
    While Not rsData.EOF
        Call cbxSpecimenName.AddItem(Nvl(rsData!�걾����))
                
        rsData.MoveNext
    Wend

    cbxSpecimenName.ListIndex = 0
End Sub


Private Sub ShowReportImageWindow()
'
    Dim frmImage As New frmPatholProcedureRep_Image
    On Error GoTo errFree
        Call frmImage.ShowImageWindow(mlngAdviceID, mblnMoved, Me)
errFree:
    Call Unload(frmImage)
    Set frmImage = Nothing
End Sub


Private Sub cmdAddRepImage_Click()
On Error GoTo errHandle
    Call ShowReportImageWindow
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Function GetSelectReportImgs() As String
'��ȡѡ��ı���ͼ��
    Dim i As Long
    Dim j As Long
    Dim objLabs As DicomLabels
    Dim strUids As String
    
    strUids = ""
    For i = 1 To rpImage.dcmViewer.Images.Count
        Set objLabs = rpImage.dcmViewer.Images(i).Labels
        
        For j = 1 To objLabs.Count
            If objLabs(j).Tag = rpImage.SelectTag Then
                If Not objLabs(j).Transparent Then
                    If strUids <> "" Then strUids = strUids & ";"
                    strUids = strUids & rpImage.dcmViewer.Images(i).InstanceUID
                End If
            End If
        Next j
    Next i
    
    GetSelectReportImgs = strUids
End Function



Private Function GetReportTypeValue(ByVal strCode As String) As String
'��ȡ��������
    Dim blnFind As Boolean
    Dim chkState As CheckState
    Dim strValue As String
    Dim lngImgIndex As Long
    
    Call ufgData.GetFieldDisplayText(gstrProcedureRep_��������, strCode, blnFind, chkState, strValue, lngImgIndex)
    GetReportTypeValue = IIf(blnFind, strValue, strCode)
End Function

Private Function GetReportSubValue(ByVal strCode As String) As String
'��ȡ��������
    Dim blnFind As Boolean
    Dim chkState As CheckState
    Dim strValue As String
    Dim lngImgIndex As Long
    
    Call ufgData.GetFieldDisplayText(gstrProcedureRep_��������, strCode, blnFind, chkState, strValue, lngImgIndex)
    GetReportSubValue = IIf(blnFind, strValue, strCode)
End Function



Private Function GetReportTypeCode(ByVal strValue As String) As String
'��ȡ��������
    Dim blnFind As Boolean
    Dim strCode As String
    
    strCode = ufgData.GetFieldDataValue(gstrProcedureRep_��������, strValue, blnFind)
    GetReportTypeCode = IIf(blnFind, strCode, strValue)
End Function


Private Function GetReportSubCode(ByVal strValue As String) As String
'��ȡ��������
    Dim blnFind As Boolean
    Dim strCode As String
    
    strCode = ufgData.GetFieldDataValue(gstrProcedureRep_��������, strValue, blnFind)
    GetReportSubCode = IIf(blnFind, strCode, strValue)
End Function


Private Sub NewProcedureRep()
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim strRepImages As String
    Dim lngNewRow As Long
    Dim dtServicesTime As Date
    Dim lngSpeexamDetails As Long
    
    
    lngSpeexamDetails = 0
    
    '��ȡ��ǰ�ؼ�ϸĿ
    Select Case Val(cbxReportType.Text)
        Case 1
            lngSpeexamDetails = Val(cbxReportSub.Text)
        Case 3
            lngSpeexamDetails = 0
        Case 2
            lngSpeexamDetails = IIf(Val(cbxReportSub.Text) > 0, Val(cbxReportSub.Text) + 2, 0)
    End Select
    
    
    strRepImages = GetSelectReportImgs()
    dtServicesTime = dtpReportTime.value  ' zlDatabase.Currentdate
    
    strSql = "select Zl_������̱���_����([1],[2],[3],[4],[5],[6],[7],[8],[9],[10]) as ����ֵ from dual"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                            mrecStudyInf.lngPatholAdviceId, _
                                            cbxSpecimenName.Text, _
                                            Val(cbxReportType.Text), _
                                            lngSpeexamDetails, _
                                            txtAdvice.Text, _
                                            txtResult.Text, _
                                            UserInfo.����, _
                                            CDate(dtServicesTime), _
                                            strRepImages, _
                                            "")
                                            
    If rsData.RecordCount <= 0 Then
        Call err.Raise(0, "NewProcedureRep", "δ�ɹ���ȡ������ı���ID,����ʧ�ܡ�")
        Exit Sub
    End If
    
    '�����̱��������б�
    lngNewRow = ufgData.NewRow ' ufgData.GetNullRowIndex
    
    ufgData.Text(lngNewRow, gstrProcedureRep_ID) = rsData!����ֵ
    ufgData.Text(lngNewRow, gstrProcedureRep_����ͼ��) = strRepImages
    ufgData.Text(lngNewRow, gstrProcedureRep_�걾����) = cbxSpecimenName.Text
    ufgData.Text(lngNewRow, gstrProcedureRep_��������) = GetReportTypeValue(Val(cbxReportType.Text))
    ufgData.Text(lngNewRow, gstrProcedureRep_��������) = GetReportSubValue(IIf(Val(cbxReportType.Text) = 1, Val(cbxReportSub.Text), IIf(Val(cbxReportType.Text) = 2, Val(cbxReportSub.Text) + 2, 0)))
    ufgData.Text(lngNewRow, gstrProcedureRep_�����) = txtResult.Text
    ufgData.Text(lngNewRow, gstrProcedureRep_������) = txtAdvice.Text
    ufgData.Text(lngNewRow, gstrProcedureRep_������) = UserInfo.����
    ufgData.Text(lngNewRow, gstrProcedureRep_��������) = dtServicesTime
    ufgData.Text(lngNewRow, gstrProcedureRep_��ǰ״̬) = "δ��ӡ"
'    ufgData.text(lngNewRow, gstrProcedureRep_��ע)=txtMemo.Text

'    Call ufgData.LocateRow(lngNewRow)
    Call ufgData_OnSelChange

    'Call MsgBoxD(Me, Decode(Val(cbxReportType.Text), 0, "����", 1, "����", 2, "����", 3, "��Ⱦ", "") & "�����ѳɹ���ӡ�", vbOKOnly, Me.Caption)
End Sub



Private Sub LoadReportContext(ByVal lngRow As Long)
'���뱨������
    Dim i As Long
    Dim strRepImages As String
    Dim lngReportSub As Long
    
    txtResult.Text = ufgData.Text(lngRow, gstrProcedureRep_�����)
    txtAdvice.Text = ufgData.Text(lngRow, gstrProcedureRep_������)
    
    dtpReportTime.value = ufgData.Text(lngRow, gstrProcedureRep_��������)
    
    '��ȡ�걾����
    For i = 0 To cbxSpecimenName.ListCount - 1
        If cbxSpecimenName.list(i) = ufgData.Text(lngRow, gstrProcedureRep_�걾����) Then
            cbxSpecimenName.ListIndex = i
            Exit For
        End If
    Next i
    
    '��ȡ��������
    cbxReportType.ListIndex = GetReportTypeCode(ufgData.Text(lngRow, gstrProcedureRep_��������))
    
    '��ȡ��������
    lngReportSub = GetReportSubCode(ufgData.Text(lngRow, gstrProcedureRep_��������))
    cbxReportSub.ListIndex = IIf(lngReportSub > 2, lngReportSub - 2, lngReportSub) - 1
    
    '����ͼ���ѡ��״̬
    strRepImages = ufgData.Text(lngRow, gstrProcedureRep_����ͼ��)
    
    For i = 1 To rpImage.dcmViewer.Images.Count
        If InStr(1, strRepImages, rpImage.dcmViewer.Images(i).InstanceUID) > 0 Then
            rpImage.ItemSelected(i) = True
        Else
            rpImage.ItemSelected(i) = False
        End If
    Next i
End Sub


Private Sub ClearReportContext()
'�������༭������
    txtResult.Text = ""
    txtAdvice.Text = ""
    
    If cbxSpecimenName.ListCount > 0 Then cbxSpecimenName.ListIndex = 0
    
    Call rpImage.ClearSelected
    
End Sub



Private Sub cbxReportType_Click()
On Error GoTo errHandle
'    mblnEditState = True
    Call LoadReportSub(Val(cbxReportType.Text))
    Call LoadReportModule(True)
    
    Call EnableReportWithSpeExamType(Val(cbxReportType.Text), mblnReadOnly, GetCurRepAllowAuditing)
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cbxSpecimenName_Click()
    mblnEditState = True
End Sub


Private Sub SaveCurProcedureRep()
'������̱������
    Dim strSql As String
    Dim strSelectRpImages As String
    Dim lngSpeexamDetails As Long
    Dim dtServicesTime As Date
    
    If Not ufgData.IsSelectionRow Then Exit Sub
    
    If ufgData.IsNullRow(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "��ǰ�����¼���ܽ��б��棬�볢�ԡ��������桱��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        If MsgBoxD(Me, "�Ƿ���Ҫ����µ�" & Decode(Val(cbxReportType.Text), 0, "����", 1, "����", 2, "����", 3, "��Ⱦ", "") & "���棿", vbYesNo, Me.Caption) = vbYes Then
            Call NewProcedureRep
    
            Call MsgBoxD(Me, Decode(Val(cbxReportType.Text), 0, "����", 1, "����", 2, "����", 3, "��Ⱦ", "") & "�����ѳɹ���ӡ�", vbOKOnly, Me.Caption)
        End If
        
        Exit Sub
    End If
    
    lngSpeexamDetails = 0
    
    '��ȡ��ǰ�ؼ�ϸĿ
    Select Case Val(cbxReportType.Text)
        Case 1
            lngSpeexamDetails = Val(cbxReportSub.Text)
        Case 3
            lngSpeexamDetails = 0
        Case 2
            lngSpeexamDetails = IIf(Val(cbxReportSub.Text) > 0, Val(cbxReportSub.Text) + 2, 0)
    End Select
    
    
    dtServicesTime = dtpReportTime.value
    strSelectRpImages = GetSelectReportImgs()
    
    strSql = "Zl_������̱���_����(" & ufgData.KeyValue(ufgData.SelectionRow) & ",'" & _
                                        cbxSpecimenName.Text & "'," & _
                                        Val(cbxReportType.Text) & "," & _
                                        lngSpeexamDetails & ",'" & _
                                        txtAdvice.Text & "','" & _
                                        txtResult.Text & "'," & _
                                        To_Date(dtServicesTime) & ",'" & _
                                        strSelectRpImages & "',Null)"
                                        
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    mblnEditState = False
    
    '���������б�
    ufgData.Text(ufgData.SelectionRow, gstrProcedureRep_����ͼ��) = strSelectRpImages
    ufgData.Text(ufgData.SelectionRow, gstrProcedureRep_�걾����) = cbxSpecimenName.Text
    ufgData.Text(ufgData.SelectionRow, gstrProcedureRep_��������) = GetReportTypeValue(Val(cbxReportType.Text))
    ufgData.Text(ufgData.SelectionRow, gstrProcedureRep_��������) = GetReportSubValue(IIf(Val(cbxReportType.Text) = 1, Val(cbxReportSub.Text), IIf(Val(cbxReportType.Text) = 2, Val(cbxReportSub.Text) + 2, 0)))
    ufgData.Text(ufgData.SelectionRow, gstrProcedureRep_�����) = txtResult.Text
    ufgData.Text(ufgData.SelectionRow, gstrProcedureRep_������) = txtAdvice.Text
    ufgData.Text(ufgData.SelectionRow, gstrProcedureRep_��������) = dtServicesTime
    ufgData.Text(ufgData.SelectionRow, gstrProcedureRep_��ǰ״̬) = "δ��ӡ"
'    ufgData.Text(ufgData.SelectionRow, gstrProcedureRep_��ע)= txtMemo.Text

    Call MsgBoxD(Me, "�����ѱ��档", vbOKOnly, Me.Caption)
End Sub



Private Sub DelCurProcedureRep()
'ɾ�����̱���
    Dim strSql As String
    
    If Not ufgData.IsSelectionRow Then Exit Sub
    
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ҫɾ����" & Decode(Val(cbxReportType.Text), 0, "����", 1, "����", 2, "����", 3, "��Ⱦ", "") & "���档", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If MsgBoxD(Me, "ȷ��Ҫɾ����ѡ���" & Decode(Val(cbxReportType.Text), 0, "����", 1, "����", 2, "����", 3, "��Ⱦ", "") & "������ɾ���󱨸潫���ָܻ���", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    strSql = "Zl_������̱���_ɾ��(" & ufgData.KeyValue(ufgData.SelectionRow) & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Call ufgData.DelRow(ufgData.SelectionRow, False)
    
    
    '������������̱��棬�������������̱��棬�����������
    If Not ufgData.IsSelectionRow Then Exit Sub
    
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        Call ClearReportContext
        
        Exit Sub
    End If
    
    Call LoadReportContext(ufgData.SelectionRow)

End Sub


Private Function GetSubReportFormat(ByVal strReportFmt As String, ByVal strRepTag As String) As String
'���ݱ���tag��ȡ��ʽ����
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    GetSubReportFormat = ""
    
    strSql = "select b.��� from zlreports a, zlrptfmts b " & _
                " where a.id = b.����id and a.���=upper([1]) and b.˵�� like [2]"
                
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strReportFmt, "%" & strRepTag & "%")
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    GetSubReportFormat = rsData!���
    
End Function


Private Sub SetDcmLabesVisible(dcmImage As DicomImage, ByVal blnVisible As Boolean)
    Dim i As Long
    
    For i = 1 To dcmImage.Labels.Count
        dcmImage.Labels(i).Visible = blnVisible
    Next i
End Sub


Private Function GetReportImageFile() As String
'ȡ�ñ���ͼ���ļ�
    Dim i As Long
    Dim strImageFiles As String
    Dim objCurDcmImage As DicomImage
    
    For i = 1 To rpImage.dcmViewer.Images.Count
        If rpImage.ItemSelected(i) Then
            Set objCurDcmImage = rpImage.dcmViewer.Images(i)
           
            '����lab��ǩ
            Call SetDcmLabesVisible(objCurDcmImage, False)
            
            Call objCurDcmImage.FileExport(strCurTempReportPath & objCurDcmImage.InstanceUID & ".jpg", "JPG")
            
            '��ʾ��ǩ
            Call SetDcmLabesVisible(objCurDcmImage, True)
            
            If strImageFiles <> "" Then strImageFiles = strImageFiles & ";"
            strImageFiles = strImageFiles & strCurTempReportPath & objCurDcmImage.InstanceUID & ".jpg"
        End If
    Next i
    
    GetReportImageFile = strImageFiles
End Function


Private Sub PrintCurProcedureRep(Optional ByVal blnIsPrint As Boolean = True)
'��ӡ���̱���
    Dim lngReportType As Long
    Dim lngReportSub As Long
    Dim lngReportID As Long
    Dim strReportFormat As String
    Dim strSubFormat As String
    Dim lngSelectImgCount As Long
    Dim strReportImgFiles As String
    Dim aryImageFiles() As String
    
    If Not ufgData.IsSelectionRow Then Exit Sub
    
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ҫ��ӡ�ı����¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    lngReportID = ufgData.Text(ufgData.SelectionRow, gstrProcedureRep_ID)
    lngReportType = GetReportTypeCode(ufgData.Text(ufgData.SelectionRow, gstrProcedureRep_��������))
    lngReportSub = GetReportSubCode(ufgData.Text(ufgData.SelectionRow, gstrProcedureRep_��������))
    
    Select Case lngReportType
        Case 0 '��������
            strReportFormat = "ZL1_Inside_1294_03"
        Case 1 '�����黯
            If lngReportSub = 1 Then
                '����
                strReportFormat = "ZL1_Inside_1294_15"
            ElseIf lngReportSub = 2 Then
                '��ҩ��ҩ
                strReportFormat = "ZL1_Inside_1294_04"
            End If
        Case 2 '���Ӳ���
            If lngReportSub = 3 Then
                'ӫ��
                strReportFormat = "ZL1_Inside_1294_05"
            ElseIf lngReportSub = 4 Then
                '��ͨ
                strReportFormat = "ZL1_Inside_1294_06"
            End If
        Case 3 '����Ⱦɫ
            strReportFormat = "ZL1_Inside_1294_14"
    End Select
    
    
    lngSelectImgCount = rpImage.SelectedCount()
    
    strSubFormat = GetSubReportFormat(strReportFormat, lngSelectImgCount & "��")
    
    If strSubFormat <> "" Then strSubFormat = "ReportFormat=" & strSubFormat
    
    strReportImgFiles = GetReportImageFile()
    
    aryImageFiles = Split(strReportImgFiles & ";;;;;;;;", ";")
    
    Call zlReport.ReportOpen(gcnOracle, 100, strReportFormat, Me, strSubFormat, _
                            "�����=" & mrecStudyInf.strPatholNumber & "", "���̱���ID=" & lngReportID, _
                            "pic1=" & aryImageFiles(0), _
                            "pic2=" & aryImageFiles(1), _
                            "pic3=" & aryImageFiles(2), _
                            "pic4=" & aryImageFiles(3), _
                            "pic5=" & aryImageFiles(4), _
                            "pic6=" & aryImageFiles(5), _
                            "pic7=" & aryImageFiles(6), _
                            "pic8=" & aryImageFiles(7), IIf(blnIsPrint, 2, 1))
End Sub

Private Sub GetSpeExamResult()
'��ȡ�ؼ���
If mCurEditText Is Nothing Then Exit Sub
If mCurEditText.Locked Then Exit Sub

Dim frmResultGet As New frmPatholResultGet
On Error GoTo errFree
    Select Case Val(cbxReportType.Text)
        Case 1  '���߽��
            Call frmResultGet.ShowResultGetWind(mrecStudyInf.lngPatholAdviceId, 0, mstrPrivs, Me)
        Case 2  '���ӽ��
            Call frmResultGet.ShowResultGetWind(mrecStudyInf.lngPatholAdviceId, 2, mstrPrivs, Me)
        Case 3  '��Ⱦ���
            Call frmResultGet.ShowResultGetWind(mrecStudyInf.lngPatholAdviceId, 1, mstrPrivs, Me)
    End Select
    
    If frmResultGet.IsOk Then
        mCurEditText.SelText = frmResultGet.txtResult.Text
    End If
    
errFree:
    Call Unload(frmResultGet)
    Set frmResultGet = Nothing
End Sub



Private Sub UpdateCurProcedureRepState(ByVal lngRPState As Long)
'���¹��̱���״̬
    Dim strSql As String
    Dim strRPState As String
    
    If Not ufgData.IsSelectionRow Then Exit Sub
        
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ҫ���иò����Ĺ��̱��档", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If lngRPState = M_REPORTSTATE_CANCEL Then
        If MsgBoxD(Me, "ȷ��Ҫ���ظñ�����", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    End If
    
    strSql = "Zl_������̱���_״̬(" & ufgData.KeyValue(ufgData.SelectionRow) & "," & lngRPState & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    'ͬ�������б�״̬��ʾ
    strRPState = ""
    Select Case lngRPState
        Case M_REPORTSTATE_NORMAL
            strRPState = "δ��ӡ"
        Case M_REPORTSTATE_VIEW
            strRPState = "�Ѳ���"
            
            Call EnableReportWithSpeExamType(Val(cbxReportType.Text), mblnReadOnly, True)
        Case M_REPORTSTATE_CANCEL
            strRPState = "�ѳ���"
            
            Call EnableReportWithSpeExamType(Val(cbxReportType.Text), mblnReadOnly, False)
        Case M_REPORTSTATE_PRINT
            strRPState = "�Ѵ�ӡ"
            
            Call EnableReportWithSpeExamType(Val(cbxReportType.Text), mblnReadOnly, True)
    End Select
    
    ufgData.Text(ufgData.SelectionRow, gstrProcedureRep_��ǰ״̬) = strRPState
End Sub


Private Sub Form_Initialize()
    Set mCurEditText = txtResult
    Set zlReport = New zl9Report.clsReport
    
    mblnEditState = False
    
    
    strCurTempReportPath = IIf(Len(App.Path) > 3, App.Path & "\", App.Path) & "TmpReportImg\"
    
    '���Ŀ¼���ڣ���ɾ����ʱ����Ŀ¼
    If Dir(strCurTempReportPath, vbDirectory) <> "" Then
      Call mdlDir.DeleteFolder(strCurTempReportPath)
    End If
    
    '�ж���ʱ����Ŀ¼�Ƿ���ڣ��粹�����򴴽�
    If Dir(strCurTempReportPath, vbDirectory) = "" Then
        Call MkDir(strCurTempReportPath)
    End If
End Sub

Private Sub LoadReportModule(Optional blnRefresh As Boolean = False)
'���뱨��ģ��
    Dim strLinkClassName As String
    
    If mlngCurDeptId = wimWord.CurDepartId And Not blnRefresh Then Exit Sub
    
    Select Case Val(cbxReportType.Text)
        Case 0
            strLinkClassName = zlDatabase.GetPara("���汨��ģ��", glngSys, glngModul, "")
        Case 1
            strLinkClassName = zlDatabase.GetPara("���߱���ģ��", glngSys, glngModul, "")
        Case 2
            strLinkClassName = zlDatabase.GetPara("���ӱ���ģ��", glngSys, glngModul, "")
        Case 3
            strLinkClassName = zlDatabase.GetPara("��Ⱦ����ģ��", glngSys, glngModul, "")
    End Select
    
    wimWord.ModuleName = strLinkClassName
    wimWord.CurDepartId = mlngCurDeptId
    
    Call wimWord.LoadWordModel
End Sub


Private Sub Form_Load()
On Error GoTo errHandle

    dtpReportTime.value = zlDatabase.Currentdate
    
    '��ʼ���б�
    Call InitProcedureRepList
    
    '���뱨������
    Call LoadReportType
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set zlReport = Nothing
End Sub

Private Sub picReportEdit_Resize()
On Error Resume Next
    Dim lngAvgHeight As Long

    framReport.Left = 0
    framReport.Top = 0
    framReport.Width = picReportEdit.Width
    framReport.Height = picReportEdit.Height - 120
    
    lngAvgHeight = Round((framReport.Height - cbxReportType.Height - 120 * 10) / 3)
    
    labReportType.Left = 120
    labReportType.Top = 300
    
    cbxReportType.Left = labReportType.Left + labReportType.Width + 120
    cbxReportType.Top = 240
    
    labReportSub.Left = cbxReportType.Left + cbxReportType.Width + 240
    labReportSub.Top = labReportType.Top
    
    cbxReportSub.Left = labReportSub.Left + labReportSub.Width + 120
    cbxReportSub.Top = 240
    
    labSpecimenName.Left = cbxReportSub.Left + cbxReportSub.Width + 240
    labSpecimenName.Top = labReportType.Top
    
    cbxSpecimenName.Left = labSpecimenName.Left + labSpecimenName.Width + 120
    cbxSpecimenName.Top = 240
    
    Line1.X1 = 120
    Line1.X2 = framReport.Width - 120
    
    Line1.Y1 = IIf(mbytFontSize = 9, 650, 695)
    Line1.Y2 = IIf(mbytFontSize = 9, 650, 695)
    
    labResult.Left = 120
    labResult.Top = 720
    
    txtResult.Left = 120
    txtResult.Top = labResult.Top + labResult.Height + 60
    txtResult.Width = framReport.Width - 240
    txtResult.Height = lngAvgHeight
    
    labAdvice.Left = 120
    labAdvice.Top = txtResult.Top + txtResult.Height + 120
    
    txtAdvice.Left = 120
    txtAdvice.Top = labAdvice.Top + labAdvice.Height + 60
    txtAdvice.Width = framReport.Width - 240
    txtAdvice.Height = lngAvgHeight - 260
    
    labReportTime.Left = txtAdvice.Width - dtpReportTime.Width - labReportTime.Width
    labReportTime.Top = labAdvice.Top
    
    dtpReportTime.Left = labReportTime.Left + labReportTime.Width + 120
    dtpReportTime.Top = labReportTime.Top - 60

    
    rpImage.Left = 120
    rpImage.Top = txtAdvice.Top + txtAdvice.Height + 120
    rpImage.Width = framReport.Width - 240
    rpImage.Height = lngAvgHeight
End Sub


Private Sub picWordModule_Resize()
On Error Resume Next
    framWord.Left = 0
    framWord.Top = 0
    framWord.Width = picWordModule.Width
    framWord.Height = picWordModule.Height - 120
    
    wimWord.Left = 120
    wimWord.Top = 240
    wimWord.Width = framWord.Width - 240
    wimWord.Height = Round(framWord.Height / 3 * 2) - 240
    
    ufgData.Left = 120
    ufgData.Top = wimWord.Top + wimWord.Height + 120
    ufgData.Width = framWord.Width - 240
    ufgData.Height = Round(framWord.Height / 3) - 240
End Sub



Private Sub rpImage_SelectedChange(ByVal lngSelectedIndex As Long, ByVal blnSelected As Boolean)
    mblnEditState = True
End Sub

Private Sub txtAdvice_Change()
    mblnEditState = True
End Sub



Private Sub txtMemo_Change()
    mblnEditState = True
End Sub

Private Sub txtAdvice_GotFocus()
    Set mCurEditText = txtAdvice
End Sub

Private Sub txtResult_Change()
    mblnEditState = True
End Sub

Private Sub txtResult_GotFocus()
    Set mCurEditText = txtResult
End Sub

Private Sub ufgData_OnSelChange()
'���뱨������
On Error GoTo errHandle

    If ufgData.ShowingRowCount <= 1 Or Not ufgData.IsSelectionRow Then
'        tbrMain.Buttons("tbReport").Enabled = False
'        tbrMain.Buttons("tbAuditing").Enabled = False
'        tbrMain.Buttons("tbCancelReport").Enabled = False
        
        Exit Sub
    End If
    

    Call ClearReportContext
    Call EnableReportWithSpeExamType(Val(cbxReportType.Text), mblnReadOnly, GetCurRepAllowAuditing, False)
    
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then Exit Sub
    
    Call LoadReportContext(ufgData.SelectionRow)
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub wimWord_OnWordDbClickEvent(ByVal strWord As String)
'����ʾ�
On Error GoTo errHandle
    If Not mCurEditText.Locked Then mCurEditText.SelText = strWord
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub zlReport_AfterPrint(ByVal ReportNum As String)
On Error GoTo errHandle
    
    '��ӡ�󱣴��Ѵ�ӡ�ı���
    If mblnEditState Then Call SaveCurProcedureRep
    
    '�޸ĵ�ǰ����״̬
    Call UpdateCurProcedureRepState(M_REPORTSTATE_PRINT)
    
    '��ӡ������༭
    Call EnableReportWithSpeExamType(Val(cbxReportType.Text), mblnReadOnly, True)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
