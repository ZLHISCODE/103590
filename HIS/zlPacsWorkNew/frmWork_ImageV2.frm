VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWork_ImageV2 
   BorderStyle     =   0  'None
   Caption         =   "Ӱ���Ƭ"
   ClientHeight    =   7530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picView 
      Height          =   2775
      Left            =   360
      ScaleHeight     =   2715
      ScaleWidth      =   4755
      TabIndex        =   1
      Top             =   4560
      Width           =   4815
      Begin DicomObjects.DicomViewer DViewer 
         Height          =   2055
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   2415
         _Version        =   262147
         _ExtentX        =   4260
         _ExtentY        =   3625
         _StockProps     =   35
         BackColor       =   -2147483636
      End
   End
   Begin MSComctlLib.ListView lvwSeq 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2990
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "iLsTree32"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwImage 
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2990
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "iLsTree32"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   480
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane DkpMain 
      Bindings        =   "frmWork_ImageV2.frx":0000
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmWork_ImageV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IWorkMenuV2


Private Const M_STR_HINT_NoSelectData As String = "��Ч�ļ�����ݣ���ѡ����Ҫִ�еļ���¼��"
Private Const M_STR_MODULE_MENU_TAG As String = "Ӱ��"

Private mlngModule As Long
Private mstrPrivs As String
Private mlngDepartId As Long
Private mObjOwner As Object

Private mobjStudyInfo As clsStudyInfo
Private mObjNotify As IEventNotify


Private mlngCurImageCount As Long

'Private mintImageLocation As Integer            '��¼ͼ�����ڵ�λ�ã�0���������ݿ⣻1���������ݿ⣻2���������ݿ�ͼ���ϴ�����ƽ̨
Private mblnAutoOpenViewer As Boolean           '�Ƿ��Զ��򿪹�Ƭ����ADViewer

Private mstrImageLevel As String                'Ӱ�������ȼ���
Private mintImageLevel As Integer               'Ӱ�������ж�
Private mintViewHistoryImageDays As Integer     '�Զ�����ʷͼ������

Private mblnShowPic As Boolean
Private mblnAddImage As Boolean                 '�Ƿ�׷��ͼ��

Private iCurImageIndex As Integer
Private mintSelectAllSeq As Integer                 '0--��״̬��1--ѡ��ȫ�����У�2--��ѡ��ȫ������
Private mintSelectAllImg As Integer                 '0--��״̬��1--ѡ��ȫ��ͼ��2--��ѡ��ȫ��ͼ��

Private mblnObserve As Boolean    '�Ƿ��й�Ƭ����Ȩ��   true��  false��

Private mblnUse3D As Boolean
Private mstr3DExeDir As String
Private mstr3DPara As String
Private mstr3DFunctions As String
Private mbln3DAutoDecompress As Boolean
Private mObjActiveMenuBar As CommandBars

Private mbyrFontState As Byte '����״̬�������ж��Ƿ�����ؼ�λ��

Private mobjPacsCore As zl9PacsCore.clsViewer

Private mblnIsRefreshStudy As Boolean
Private mblnIsHistoryMode As Boolean

'����ID
Property Get DeptId() As Long
    DeptId = mlngDepartId
End Property

'�����Ϣ
Property Get StudyInfo() As clsStudyInfo
    Set StudyInfo = mobjStudyInfo
End Property

Property Set StudyInfo(value As clsStudyInfo)
    Set mobjStudyInfo = value
    
    mblnIsRefreshStudy = False
End Property


Property Get AdviceId() As Long
    AdviceId = mobjStudyInfo.lngAdviceId
End Property

'��ȡ��Ҫʹ�õ��ⲿ����
Property Get PacsCore() As Object
    Set PacsCore = mobjPacsCore
End Property

Property Set PacsCore(value As Object)
    Set mobjPacsCore = value
End Property


'��ȡ�˵��ӿڶ���
Property Get zlMenu() As IWorkMenuV2
    Set zlMenu = Me
End Property



'�ӿ�ʵ�ֲ���**************************************************************************************************


Public Function IWorkMenuV2_zlBaseMenuID() As Long
End Function

Public Function IWorkMenuV2_zlExecuteCmd(ByVal lngCmdType As Long)
'ִ�в˵�����

End Function

Public Function IWorkMenuV2_zlGetModuleMenuId() As Long
'��ȡӰ��˵��Ĳ˵�ID
    IWorkMenuV2_zlGetModuleMenuId = conMenu_Img_Group
End Function



Public Function IWorkMenuV2_zlIsModuleMenu(ByVal strModuleName As String, objControlMenu As XtremeCommandBars.ICommandBarControl) As Boolean
'�жϲ˵��Ƿ����ڸ�ģ��˵�
    IWorkMenuV2_zlIsModuleMenu = IIf(objControlMenu.Category = M_STR_MODULE_MENU_TAG, True, False)
End Function


Public Sub IWorkMenuV2_zlCreateMenu(ByVal strModuleName As String, objMenuBar As Object)
'����Ӱ���¼��Ӧ�Ĳ˵�
    
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim objGC As CommandBarControl
    
    Dim str3DFuncs() As String
    Dim i As Long
    Dim lng3DFunc As Long
    
    Set mObjActiveMenuBar = objMenuBar
    
    'ɾ��Ӱ�������Ӳ˵�
    Set cbrMenuBar = objMenuBar.FindControl(, conMenu_ManagePopup)
    Set cbrControl = cbrMenuBar.CommandBar.FindControl(, conMenu_Manage_ImageQuality, , True)
    If Not cbrControl Is Nothing Then
        Call cbrControl.Delete
    End If

    Set cbrMenuBar = objMenuBar.FindControl(, conMenu_ManagePopup)
    With cbrMenuBar.CommandBar
        '����Ӱ�������˵�
        If CheckPopedom(mstrPrivs, "Ӱ���ʿ�") Then
            Set objGC = cbrMenuBar.CommandBar.FindControl(, conMenu_Manage_GChannel, , True)
            
            If objGC Is Nothing Then
                Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_Manage_ImageQuality, "Ӱ������", "", 0, False, .Controls.Count - 1)
            Else
                Set cbrControl = CreateModuleMenu(objGC.Parent.Controls, xtpControlPopup, conMenu_Manage_ImageQuality, "Ӱ������", "", 0, False, objGC.Index - 1)
            End If
            
            Call CreateSubordinateMenuTools(mstrImageLevel, cbrControl)
        End If
    End With
    
    If Not HasMenu(objMenuBar, conMenu_Img_Group) Then
        Set cbrMenuBar = mObjActiveMenuBar.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_Img_Group, "Ӱ��", 3, False)
        cbrMenuBar.ID = conMenu_Img_Group
        cbrMenuBar.Category = M_STR_MODULE_MENU_TAG
        
        
        With cbrMenuBar.CommandBar
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Img_Look, "Ӱ���Ƭ", "", 8111, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Img_Contrast, "Ӱ��Ա�", "", 8112, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Img_Look3D, "3D��Ƭ", "", 8115, False)
            
            '���������ά�ؽ����ܣ��򴴽���Ӧ�˵�
            If mblnUse3D = True Then
                Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_Img_3D, "��ά�ؽ�")  '.Add(xtpControlPopup, conMenu_Img_3D, "��ά�ؽ�"): cbrControl.ID = conMenu_Img_3D
                    If mstr3DFunctions <> "" Then
                        str3DFuncs = Split(mstr3DFunctions, ",")
                        For i = 1 To UBound(str3DFuncs)
                            lng3DFunc = Val(str3DFuncs(i))
                            If lng3DFunc >= 1 And lng3DFunc <= 6 Then
                                Select Case lng3DFunc
                                    Case 1
                                        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Img_3D_VA, "�ݻ��ؽ�")
                                    Case 2
                                        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Img_3D_MPR, "MPR")
                                    Case 3
                                        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Img_3D_MMPR, "MMPR")
                                    Case 4
                                        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Img_3D_VE, "�����ڿ���")
                                    Case 5
                                        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Img_3D_SA, "�����ؽ�")
                                    Case 6
                                        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Img_3D_PF, "��ע����")
                                End Select
                            End If
                        Next i
                    End If
            End If
             
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Img_Delete, "Ӱ��ȫɾ", "", 8113, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Img_Query, "Ӱ���ȡ(Q/R)", "", 8111, False)
            
            If gblnUseXinWangView = True Then
                '�ж��Ƿ�ʹ���°��Ƭ
                'Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_FilmPrevew, "��ƬԤ��", "", 0, True)
                Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_FilmPrint, "��Ƭ��ӡ", "", 0, False)
                'Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_FilmDelete, "��Ƭɾ��", "", 0, False)
            End If
            
            If CheckPopedom(mstrPrivs, "��ʦִ��") Then
                Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_TechDoctorExecute, "��ʦִ��", "ָ����ǰ���ļ�鼼ʦ", 807, True)
            End If
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_ChangeDevice, "����Ӱ������", "", 3203, False)
            
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_Show, "Ӱ����ʾ", "��ʾ��ǰ����Ӱ������ͼ", 3061, True): cbrControl.Style = xtpButtonIconAndCaption
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_Expend_AllCollapse, "ȫѡ����", "ѡ�е�ǰ��������", 3010, False): cbrControl.Style = xtpButtonIconAndCaption
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_Expend_AllExpend, "ȫ������", "���ѡ�е�ǰ��������", 3004, False): cbrControl.Style = xtpButtonIconAndCaption
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_SelectAllImages, "ȫѡͼ��", "ѡ�е�ǰ����ͼ��", 227, False): cbrControl.Style = xtpButtonIconAndCaption
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_UnSelectAllImages, "ȫ��ͼ��", "���ѡ�е�ǰ����ͼ��", 229, False): cbrControl.Style = xtpButtonIconAndCaption
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_ReverseSelectImages, "��ѡͼ��", "����ѡ������ͼ��", 3012, False): cbrControl.Style = xtpButtonIconAndCaption
                
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_DeleteSeries, "ɾ��ѡ������", "ɾ��ѡ�������", 0, True): cbrControl.Style = xtpButtonIconAndCaption
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_DeleteImage, "ɾ��ѡ��ͼ��", "ɾ��ѡ���ͼ��", 0, False): cbrControl.Style = xtpButtonIconAndCaption
                
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_DevSet, "Ӱ���豸����", "ˢ�µ�ǰ����ͼ������", 181, True): cbrControl.Style = xtpButtonIconAndCaption
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_RefreshImg, "ˢ��ͼ��", "ˢ�µ�ǰ����ͼ������", 791, True): cbrControl.Style = xtpButtonIconAndCaption
        End With
    End If
End Sub

Public Sub IWorkMenuV2_zlCreateToolBar(ByVal strModuleName As String, objToolBar As Object)
'����������
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrLogOut As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim lngIndex As Long
        
    Dim str3DFuncs() As String
    Dim i As Long
    Dim lng3DFunc As Long
    
    'ɾ��Ӱ������������
    Set cbrControl = objToolBar.FindControl(, conMenu_Manage_ImageQuality)
    If Not cbrControl Is Nothing Then Call cbrControl.Delete

    '����Ӱ������������
    If CheckPopedom(mstrPrivs, "Ӱ���ʿ�") Then
        Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_ImageQuality, "Ӱ������", "Ӱ������", 3061, False, mObjActiveMenuBar.FindControl(, conMenu_Manage_Result).Index + 1)
        Call CreateSubordinateMenuTools(mstrImageLevel, cbrControl)
    End If
    
    If HasMenu(objToolBar, conMenu_Img_Look) Then Exit Sub
    
    Set cbrLogOut = objToolBar.FindControl(, conMenu_img_ContrastView)
    
    lngIndex = 2
    If Not cbrLogOut Is Nothing Then lngIndex = cbrLogOut.Index
'
'    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_Img_Look, "��Ƭ", "Ӱ���Ƭ", 8111, True, lngIndex + 1)
'    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_Img_Contrast, "�Ա�", "��Ƭ�Ա�", 8112, False, lngIndex + 2)
    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_Img_Look3D, "3D��Ƭ", "3D��Ƭ", 8115, False, lngIndex + 1)
    
    If CheckPopedom(mstrPrivs, "��ʦִ��") Then
        Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_Manage_TechDoctorExecute, "��ʦִ��", "ָ����ǰ���ļ�鼼ʦ", 807, False, lngIndex + 2)
    End If
    
    '���������ά�ؽ����ܣ��򴴽���Ӧ�˵�
    If mblnUse3D = True Then
        Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButtonPopup, conMenu_Img_3D, "��ά", "��ά�ؽ�", 8115, False, lngIndex + 3)
            If mstr3DFunctions <> "" Then
                str3DFuncs = Split(mstr3DFunctions, ",")
                For i = 1 To UBound(str3DFuncs)
                    lng3DFunc = Val(str3DFuncs(i))
                    If lng3DFunc >= 1 And lng3DFunc <= 6 Then
                        Select Case lng3DFunc
                            Case 1
                                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Img_3D_VA, "�ݻ��ؽ�")
                            Case 2
                                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Img_3D_MPR, "MPR")
                            Case 3
                                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Img_3D_MMPR, "MMPR")
                            Case 4
                                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Img_3D_VE, "�����ڿ���")
                            Case 5
                                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Img_3D_SA, "�����ؽ�")
                            Case 6
                                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Img_3D_PF, "��ע����")
                        End Select
                    End If
                Next i
            End If
    End If
    
    If gblnUseXinWangView = True Then
        '�ж��Ƿ�ʹ���°��Ƭ
        Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_Manage_FilmPrint, "��Ƭ��ӡ", "", 3202, False, lngIndex + 4)
    End If
End Sub

Private Sub CreateSubordinateMenuTools(ByVal strImageLevel As String, ByVal cbrControl As CommandBarControl)
'�����¼��˵��͹�����
    Dim cbrPopControl As CommandBarControl
    Dim intTxtLen As Integer
    Dim i As Integer
    
    intTxtLen = Len(strImageLevel) - Len(Replace(strImageLevel, ",", "")) + 1
    For i = 1 To 4
        If i <= intTxtLen Then
            If Trim(Split(strImageLevel, ",")(i - 1)) <> "" Then
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, _
                    Decode(i, 1, conMenu_Manage_ImageFirst, 2, conMenu_Manage_ImageSecond, 3, conMenu_Manage_ImageThird, 4, conMenu_Manage_ImageFourth), Trim(Split(strImageLevel, ",")(i - 1)), "", 0, False)
            End If
        End If
    Next i
End Sub

Public Sub IWorkMenuV2_zlClearMenu(ByVal strModuleName As String)
'����������Ĳ˵�
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    
    If mObjActiveMenuBar Is Nothing Then Exit Sub
    
    Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Img_Group)
    If Not cbrControl Is Nothing Then Call cbrControl.Delete
    
    'ɾ��Ӱ�������Ӳ˵�
    Set cbrMenuBar = mObjActiveMenuBar.FindControl(, conMenu_ManagePopup)
    Set cbrControl = cbrMenuBar.CommandBar.FindControl(, conMenu_Manage_ImageQuality)
    If Not cbrControl Is Nothing Then Call cbrControl.Delete
End Sub


Public Sub IWorkMenuV2_zlClearToolBar(ByVal strModuleName As String)
'��������Ĺ�����
    Dim cbrControl As CommandBarControl
    
    If mObjActiveMenuBar Is Nothing Then Exit Sub
    
    Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Img_Look)
    If Not cbrControl Is Nothing Then Call cbrControl.Delete
    
    Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Img_Contrast)
    If Not cbrControl Is Nothing Then Call cbrControl.Delete
    
    Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Img_Look3D)
    If Not cbrControl Is Nothing Then Call cbrControl.Delete
    
    Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Manage_TechDoctorExecute)
    If Not cbrControl Is Nothing Then Call cbrControl.Delete
    
    Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Img_3D)
    If Not cbrControl Is Nothing Then Call cbrControl.Delete
    
    If gblnUseXinWangView = True Then
    '�ж��Ƿ�ʹ���°��Ƭ
        Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Manage_FilmPrint)
        If Not cbrControl Is Nothing Then Call cbrControl.Delete
    End If
    
    'ɾ��Ӱ������������
    Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Manage_ImageQuality)
    If Not cbrControl Is Nothing Then Call cbrControl.Delete
End Sub

Public Sub IWorkMenuV2_zlExecuteMenu(ByVal strModuleName As String, ByVal lngMenuId As Long)
'���ݲ˵�IDִ�ж�Ӧ����
    
    Select Case lngMenuId
        Case conMenu_Img_Look                           '��Ƭ
            Call Menu_Img_��Ƭ
        Case conMenu_Img_Contrast                       '�Աȹ�Ƭ
            Call Menu_Img_�Աȹ�Ƭ
        Case conMenu_Img_Look3D                         '3D��Ƭ
            Call Menu_Img_3D��Ƭ
        Case conMenu_Manage_FilmPrint                   '��Ƭ��ӡ
            Call Menu_Manage_FilmPrint
        Case conMenu_Img_3D_MMPR                        '��ά�ؽ���MMPR
            Call sub��ά�ؽ�("MMPR")
        Case conMenu_Img_3D_MPR                         '��ά�ؽ���MPR
            Call sub��ά�ؽ�("MPR")
        Case conMenu_Img_3D_PF                          '��ά�ؽ�,��ע����
            Call sub��ά�ؽ�("PF")
        Case conMenu_Img_3D_SA                          '��ά�ؽ��������ؽ�
            Call sub��ά�ؽ�("SA")
        Case conMenu_Img_3D_VA                          '��ά�ؽ����ݻ��ؽ�
            Call sub��ά�ؽ�("VA")
        Case conMenu_Img_3D_VE                          '��ά�ؽ��������ڿ���
            Call sub��ά�ؽ�("VE")
        Case conMenu_Img_Delete                         'ͼ��ɾ��
            Call Menu_Img_ͼ��ɾ��
        Case conMenu_Img_Query                          '���豸��ȡͼ��
            Call Menu_Img_��ȡͼ��
        Case conMenu_Manage_TechDoctorExecute           '��ʦִ��
            Call Menu_Img_��ʦִ��
        Case conMenu_Manage_ChangeDevice                '����Ӱ������
            Call Menu_Img_��������豸
        Case conMenu_View_Show          '��ʾͼ��
            mblnShowPic = Not mblnShowPic
            Call zlMenuClick("Ӱ����ʾ")
        Case conMenu_View_Expend_AllCollapse    'ȫѡ����
            Call zlMenuClick("ȫѡ����")
        Case conMenu_View_Expend_AllExpend      'ȫ������
            Call zlMenuClick("ȫ������")
        Case conMenu_Manage_SelectAllImages     'ȫѡͼ��
            Call zlMenuClick("ȫѡͼ��")
        Case conMenu_Manage_UnSelectAllImages   'ȫ��ͼ��
            Call zlMenuClick("ȫ��ͼ��")
        Case conMenu_Manage_ReverseSelectImages '��ѡͼ��
            Call zlMenuClick("��ѡͼ��")
        Case conMenu_Manage_DeleteSeries        'ɾ������
            Call zlMenuDeleteImageClick(lngMenuId)
        Case conMenu_Manage_DeleteImage         'ɾ��ͼ��
            Call zlMenuDeleteImageClick(lngMenuId)
        Case conMenu_Cap_DevSet                 'Ӱ���豸����
            frmPACSImageDeviceSetup.Show vbModal, mObjOwner
        Case conMenu_Manage_RefreshImg          'ˢ��ͼ��
            Call zlRefreshFace(mobjStudyInfo, True)
        Case conMenu_Manage_ImageFirst, conMenu_Manage_ImageSecond, conMenu_Manage_ImageThird, conMenu_Manage_ImageFourth
            Call Menu_Manage_Ӱ������(lngMenuId, mstrImageLevel)
    End Select
End Sub

Public Sub IWorkMenuV2_zlUpdateMenu(ByVal strModuleName As String, Control As XtremeCommandBars.ICommandBarControl)
'���²˵�
    If mobjStudyInfo Is Nothing Then
        Control.Enabled = False
        Exit Sub
    End If
    
    Select Case Control.ID
        Case conMenu_Img_Look       '��Ƭ����ǰ�����ͼ�񣬻����ǻ�������ʷ��飬����Թ�Ƭ
            Control.Enabled = (mobjStudyInfo.strStudyUID <> "") Or mintViewHistoryImageDays > 0
            
            
        Case conMenu_Img_Contrast   '��Ƭ�Աȣ�ֻ������PACS����ʾ
            Control.Enabled = (mobjStudyInfo.strStudyUID <> "")
            Control.Visible = IIf(mobjStudyInfo.intImageLocation = 0, True, False)
        
        Case conMenu_Img_Look3D     '3D��Ƭ��ֻ������PACS����ʾ
            Control.Enabled = mobjStudyInfo.strStudyUID <> "" And mlngCurImageCount >= 50
            Control.Visible = (mobjStudyInfo.intImageLocation <> 0)
            
        Case conMenu_Manage_FilmPrint                    '��Ƭ��ӡ
            Control.Visible = CheckPopedom(mstrPrivs, "��Ƭ�����ӡ")
            
        Case conMenu_Img_3D         '��ά�ؽ�
            If CheckPopedom(mstrPrivs, "��ά�ؽ�����") And mblnUse3D = True Then
                Control.Visible = True
            Else
                Control.Visible = False
            End If
            
            If Control.Visible = True Then Control.Enabled = mobjStudyInfo.strStudyUID <> ""
            
        Case conMenu_Img_Delete '���ͼ��ͼ������ƽ̨������ʾ��ť
            If Not CheckPopedom(mstrPrivs, "���ͼ��") Or (mobjStudyInfo.intImageLocation = 2) Then
                Control.Visible = False
            Else
                Control.Visible = True
            End If
            
            If Control.Visible = True Then Control.Enabled = mobjStudyInfo.strStudyUID <> ""
            
            Control.Enabled = Not mblnIsHistoryMode
            
        Case conMenu_Img_Query ',��ȡͼ��ֻ������PACS����ʾ
            If (Not CheckPopedom(mstrPrivs, "���ͼ��")) Or (mobjStudyInfo.intImageLocation <> 0) Then
                Control.Visible = False
            Else
                Control.Visible = True
            End If
            
            If Control.Visible Then Control.Enabled = mobjStudyInfo.intStep > 1
            
        Case conMenu_Manage_ChangeDevice    '����Ӱ���豸����
                If mobjStudyInfo.strImgType = "CR" Or _
                    mobjStudyInfo.strImgType = "DR" Or _
                    mobjStudyInfo.strImgType = "DX" Or _
                    mobjStudyInfo.strImgType = "RF" Then
                    Control.Enabled = True
                Else
                    Control.Enabled = False
                End If
                
                Control.Enabled = Control.Enabled And Not mblnIsHistoryMode
                
        Case conMenu_Manage_TechDoctorExecute   '��ʦִ��
            If mobjStudyInfo.blnIsTechincalSure Then Control.Caption = "��ʦȡ��" Else Control.Caption = "��ʦִ��"
            
            If mobjStudyInfo.intStep >= 2 And mobjStudyInfo.intStep < 5 Then
                Control.Enabled = True
                
                If mobjStudyInfo.blnIsTechincalSure Then
                    Control.Enabled = UserInfo.���� = mobjStudyInfo.strDoDoctor Or CheckPopedom(mstrPrivs, "ȡ����ʦִ��")
                End If
            Else
                Control.Enabled = False
            End If
            
            Control.Enabled = Control.Enabled And Not mblnIsHistoryMode
            
        Case conMenu_Manage_DeleteSeries    'ɾ��ѡ������
            Control.Enabled = lvwSeq.ListItems.Count > 0 And Me.Visible And Not mblnIsHistoryMode
        Case conMenu_Manage_DeleteImage     'ɾ��ѡ��ͼ��
            Control.Enabled = lvwImage.ListItems.Count > 0 And Me.Visible And Not mblnIsHistoryMode
        Case conMenu_View_Show, conMenu_View_Expend_AllCollapse, conMenu_View_Expend_AllExpend  'ͼ����ʾ��ȫѡ���У�ȫ������
            Control.Enabled = lvwSeq.ListItems.Count > 0 And Me.Visible
            Control.Visible = (mobjStudyInfo.intImageLocation = 0) And Me.Visible
            Control.Checked = Me.cbrMain.FindControl(, Control.ID).Checked
        Case conMenu_Manage_SelectAllImages, conMenu_Manage_UnSelectAllImages, conMenu_Manage_ReverseSelectImages   'ȫѡͼ��ȫ��ͼ�񣬷�ѡͼ��
            Control.Enabled = lvwImage.ListItems.Count > 0 And Me.Visible
            Control.Visible = (mobjStudyInfo.intImageLocation = 0) And Me.Visible
            Control.Checked = Me.cbrMain.FindControl(, Control.ID).Checked
        Case conMenu_Img_Group, conMenu_Img_Query, conMenu_View_Refresh, conMenu_Cap_DevSet 'Ӱ��Ӱ���ȡ��ͼ��ˢ��
            Control.Enabled = True
        Case conMenu_Manage_ImageFirst, conMenu_Manage_ImageSecond, conMenu_Manage_ImageThird, conMenu_Manage_ImageFourth, conMenu_Manage_ImageQuality
            If Not CheckPopedom(mstrPrivs, "Ӱ���ʿ�") Or mintImageLevel = 0 Then
                Control.Visible = False
            ElseIf (mobjStudyInfo.intStep >= 3 And mobjStudyInfo.intStep <= 5) Or mobjStudyInfo.intStep = -1 Then
                Control.Visible = True
                Control.Enabled = mobjStudyInfo.strStudyUID <> ""
            Else
                Control.Visible = True
                Control.Enabled = False
            End If
    End Select
End Sub

Public Sub IWorkMenuV2_zlPopupMenu(ByVal strModuleName As String, objPopup As XtremeCommandBars.ICommandBar)
'�����Ҽ��˵�
    Dim objControl As CommandBarControl
    Dim objMenuControl As CommandBarControl
    
    For Each objMenuControl In mObjActiveMenuBar.ActiveMenuBar.Controls
        If objMenuControl.ID = conMenu_Img_Group And objMenuControl.type = xtpControlPopup Then
            For Each objControl In objMenuControl.CommandBar.Controls
                If objControl.ID = conMenu_Img_Look Or _
                   objControl.ID = conMenu_Img_Contrast Or _
                   objControl.ID = conMenu_Img_Look3D Or _
                   objControl.ID = conMenu_Img_Delete Or _
                   objControl.ID = conMenu_Img_Query Or _
                   objControl.ID = conMenu_Manage_TechDoctorExecute Or _
                   objControl.ID = conMenu_Manage_ChangeDevice Or _
                   objControl.ID = conMenu_Manage_FilmPrint Then
                   
                    objControl.Copy objPopup
                    
                End If
            Next
        End If
    Next
End Sub

Public Sub IWorkMenuV2_zlRefreshSubMenu(ByVal strModuleName As String, objMenuBar As Object)
'ˢ�µ������Ӳ˵�
    Exit Sub
End Sub


'**********************************************************************************************************************

Private Function CreateModuleMenu(objMenuControl As CommandBarControls, _
    ByVal lngType As XTPControlType, ByVal lngID As Long, ByVal strCaption As String, _
    Optional strToolTip As String = "", Optional lngIconId As Long = 0, Optional blnStartGroup As Boolean = False, Optional ByVal lngIndex As Long = -1) As CommandBarControl
'������ģ���ڵĲ˵�
    
    If lngIndex >= 0 Then
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption, lngIndex)
    Else
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption)
    End If
    
    CreateModuleMenu.ID = lngID '������ﲻָ��id�����ܽ���Щ�˵���ӵ��Ҽ��˵���
    
    If lngIconId <> 0 Then CreateModuleMenu.iconid = lngIconId
    If blnStartGroup Then CreateModuleMenu.BeginGroup = True
    If strToolTip <> "" Then CreateModuleMenu.ToolTipText = strToolTip
    
    CreateModuleMenu.Category = M_STR_MODULE_MENU_TAG
End Function


Public Sub zlInitModule(objNotify As IEventNotify, ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngDepartId As Long)
    mlngModule = lngModule
    mstrPrivs = strPrivs
    mlngDepartId = lngDepartId
    
    Set mObjNotify = objNotify
    
    If Not objNotify Is Nothing Then Set mObjOwner = objNotify.Owner
    
    mblnUse3D = Val(zlDatabase.GetPara("������ά�ؽ�", glngSys, lngModule, 0))
    mstr3DExeDir = zlDatabase.GetPara("3D����·��", glngSys, lngModule, "")
    mstr3DPara = zlDatabase.GetPara("3D����", glngSys, lngModule, "")
    mstr3DFunctions = zlDatabase.GetPara("3D����", glngSys, lngModule, "")
    mbln3DAutoDecompress = Val(zlDatabase.GetPara("3D�Զ���ѹ��", glngSys, lngModule, 0))
    mstrImageLevel = nvl(GetDeptPara(mlngDepartId, "Ӱ�������ȼ�", "��,��"))
    mintImageLevel = Val(GetDeptPara(mlngDepartId, "Ӱ�������ж�", "0"))
    mintViewHistoryImageDays = Val(GetDeptPara(mlngDepartId, "�Զ�����ʷͼ������", 0))
    If mintViewHistoryImageDays > 15 Or mintViewHistoryImageDays <= 0 Then
        mintViewHistoryImageDays = 1
    End If
End Sub



Public Function zlRefreshFace(objStudyInfo As clsStudyInfo, _
    Optional blnForceRefresh As Boolean = False, Optional ByVal blnIsHistory As Boolean = False) As Boolean
On Error GoTo DBError

    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    mblnIsHistoryMode = blnIsHistory
       
    If Not mobjStudyInfo Is Nothing And Not objStudyInfo Is Nothing Then
        If mblnIsRefreshStudy = True Then    '�Ѿ�ˢ�¹�����Ҫ�жϼ����Ϣ��ͬ�Լ��Ƿ����ǿ��ˢ��
            If mobjStudyInfo.IsEquals(objStudyInfo) And blnForceRefresh = False Then Exit Function
        Else
            blnForceRefresh = True
        End If
    End If
        
    
    mblnShowPic = False

    Set mobjStudyInfo = objStudyInfo

    'ת����Ӱ���ܱ��汨��
    If mobjStudyInfo.blnMoved Then
        mstrPrivs = Replace(mstrPrivs, "ͼ���������", "")
        mstrPrivs = Replace(mstrPrivs, "ͼ���ע����", "")
        mstrPrivs = Replace(mstrPrivs, "���ͼ��", "")
    End If
    
    Call ShowSeqImg
    
    mblnIsRefreshStudy = True
    
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub Menu_Manage_Ӱ������(ByVal lngID As Long, ByVal strImageLevel As String)
On Error GoTo errhandle
    Dim strSQL As String
    Dim strResult As String
    Dim strGrades() As String

    If mobjStudyInfo.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, frmPacsMain.Caption
        Exit Sub
    End If
    
    If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_QUALITYTAG, 0, mobjStudyInfo.lngAdviceId)
    
    Select Case lngID
        Case conMenu_Manage_ImageFirst
            strResult = 1
        Case conMenu_Manage_ImageSecond
            strResult = 2
        Case conMenu_Manage_ImageThird
            strResult = 3
        Case conMenu_Manage_ImageFourth
            strResult = 4
    End Select

    strSQL = "Zl_Ӱ������_Update(" & mobjStudyInfo.lngAdviceId & ",'" & strResult & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "Ӱ������")
    
    If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_QUALITYTAG, 1, mobjStudyInfo.lngAdviceId, strResult)
Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Menu_Img_��Ƭ()
On Error GoTo errhandle
    
    If mobjStudyInfo.lngAdviceId = 0 Then
        MsgBoxD mObjOwner, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    'ˢ�½���
    Call zlRefreshFace(mobjStudyInfo, False)
    
    Call zlMenuClick("Ӱ����")
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Img_3D��Ƭ()
On Error GoTo errhandle
    
    If mobjStudyInfo.lngAdviceId = 0 Then
        MsgBoxD mObjOwner, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    'ˢ�½���
    Call zlRefreshFace(mobjStudyInfo, False)
    
    Call zlMenuClick("Ӱ��3D��Ƭ")
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Img_�Աȹ�Ƭ()
On Error GoTo errhandle
    
    If mobjStudyInfo.lngAdviceId = 0 Then
        MsgBoxD mObjOwner, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    'ˢ�½���
    Call zlRefreshFace(mobjStudyInfo, False)
    
    Call zlMenuClick("Ӱ��Ա�")
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_Manage_FilmPrint()
'��Ƭ��ӡ
On Error GoTo errhandle
    Dim blnPrintResult As Boolean
    
    '�ж��Ƿ������Ӧ����Ȩ��
    If Not CheckPopedom(mstrPrivs, "��Ƭ�����ӡ") Then
        MsgBoxD Me, "�����߱���Ƭ��ӡȨ�ޣ�����ϵ����Ա��", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_XWFILMPRINT, 0, mobjStudyInfo.lngAdviceId)
    
    blnPrintResult = XWShowFilmPrintWind(mobjStudyInfo.lngAdviceId, Me)
    
    If blnPrintResult = True Then
        '���ͽ�Ƭ��ӡ��Ϣ����������
        If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_XWFILMPRINT, 1, mobjStudyInfo.lngAdviceId)
    End If
    
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_Img_ͼ��ɾ��()
On Error GoTo errhandle
    Dim rsTemp As ADODB.Recordset
    Dim blnIsCancel As Boolean
    
    If mobjStudyInfo.lngAdviceId = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    '���ͼ������ƽ̨mintImageLocation=2������ʾ��ͼ��ɾ������ť��������ɾ��ͼ��
    If mobjStudyInfo.intImageLocation = 1 Then
        'ͼ��������PACS������ArchiveManagerɾ��ͼ��
        Call subXWShowArchiveManager(1)
    ElseIf mobjStudyInfo.intImageLocation = 0 Then    'ͼ��������PACS
        If Not mObjNotify Is Nothing Then
            Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 0, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, blnIsCancel)
            If blnIsCancel Then Exit Sub
        End If
        
        Call zlRefreshFace(mobjStudyInfo, False)
            
        gstrSQL = "select ���UID from Ӱ�����¼ where ҽ��ID =[1] and  ���ͺ� = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���UID", mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo)
        
        If rsTemp.EOF Then Exit Sub
            
        If MsgBoxD(Me, "�Ƿ�ȷ��Ҫɾ���ü�������Ӱ��", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub

        
        'ɾ��Ӱ���ļ���Ŀ¼
        RemoveCheckImages mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo
        
        gstrSQL = "ZL_Ӱ����_PhotoDelete(" & mobjStudyInfo.lngAdviceId & "," & mobjStudyInfo.lngSendNo & ")"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            
        Call ClearListData
        
        '������һ������Ҳ��ɾ����,Ӧ��ˢ���б�
        If lvwSeq.ListItems.Count = 0 Then
            If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 1, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, -1, hwnd)
        Else
            If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 1, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, , hwnd)
        End If
    End If
Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub ReSetFormFontSize(ByVal bytFontSize As Byte)
'����:�������ù���վ����������С
    
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    
    Me.FontSize = bytFontSize
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("ListView")
            objCtrl.Font.Size = bytFontSize
            objCtrl.Font.Name = "΢���ź�"
        Case UCase("TabStrip") 'ҳ��ؼ�
            objCtrl.Font.Size = bytFontSize
        Case UCase("Label")
            objCtrl.FontSize = bytFontSize
            objCtrl.Height = TextHeight("��") + 20
        Case UCase("vsFlexGrid")
            objCtrl.FontSize = bytFontSize
        Case UCase("ucFlexGrid")
            objCtrl.DataGrid.Cell(flexcpFontSize, 0, 0, 0, objCtrl.DataGrid.Cols - 1) = bytFontSize
            objCtrl.DataGrid.FontSize = bytFontSize
        Case UCase("ComboBox")
            objCtrl.FontSize = bytFontSize
        Case UCase("OptionButton")
            objCtrl.FontSize = bytFontSize
            objCtrl.Width = TextWidth("�޹�" & objCtrl.Caption)
        Case UCase("CheckBox")
            objCtrl.FontSize = bytFontSize
            objCtrl.Width = TextWidth("�޹�" & objCtrl.Caption)
        Case UCase("DTPicker")
            objCtrl.Font.Size = bytFontSize
            objCtrl.Width = TextWidth("2012-01-01 23:59:59") * 1.25
            objCtrl.Height = TextHeight("��") * 1.5
        Case UCase("textBox")
          objCtrl.FontSize = bytFontSize
        Case UCase("ReportControl")
            Set CtlFont = objCtrl.PaintManager.CaptionFont
            CtlFont.Size = bytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
            
            Set CtlFont = objCtrl.PaintManager.TextFont
            CtlFont.Size = bytFontSize
            Set objCtrl.PaintManager.TextFont = CtlFont
            objCtrl.Redraw
        Case UCase("DockingPane")
            Set CtlFont = objCtrl.PaintManager.CaptionFont
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = bytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
        Case UCase("CommandBars")
            Set CtlFont = objCtrl.options.Font
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = bytFontSize
            Set objCtrl.options.Font = CtlFont
        Case UCase("TabControl")
            Set CtlFont = objCtrl.PaintManager.Font
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = bytFontSize
            Set objCtrl.PaintManager.Font = CtlFont
        Case UCase("CommandButton")
            objCtrl.FontSize = bytFontSize
        End Select
    Next
    
    Call lvwSeq.Refresh
    
End Sub




Private Sub ClearListData()
'ɾ�������б��е�����
    lvwSeq.ListItems.Clear
    lvwImage.ListItems.Clear
    DViewer.Images.Clear
End Sub


Private Sub Menu_Img_��ȡͼ��()
On Error GoTo errhandle
    Dim strImageDeviceNumber As String, rsTemp As ADODB.Recordset

    If mobjStudyInfo.lngAdviceId = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_GETIMAGE, 0, mobjStudyInfo.lngAdviceId)
    
    Call zlRefreshFace(mobjStudyInfo, False)
    
    strImageDeviceNumber = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPACSImageDeviceSetup", "Ĭ��Ӱ���豸", "")
    
    'û��Ĭ���豸ʱ����
    If strImageDeviceNumber = "" Then
        If MsgBoxD(Me, "û������Ĭ��Ӱ�����豸���Ƿ��������ã�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        Else
            frmPACSImageDeviceSetup.Show vbModal, Me
            strImageDeviceNumber = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPACSImageDeviceSetup", "Ĭ��Ӱ���豸", "")
            If strImageDeviceNumber = "" Then Exit Sub
        End If
    End If
    
    gstrSQL = "select �豸��,�豸��, IP��ַ,�˿ں�,����AE,�豸AE from Ӱ���豸Ŀ¼ where �豸�� = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(Mid(strImageDeviceNumber, 2)))
    
    '��Ĭ���豸��ɾ������������
    If rsTemp.EOF = True Then
        MsgBoxD Me, "Ĭ���豸�ѱ�ɾ�������������ã�", vbInformation, gstrSysName
        frmPACSImageDeviceSetup.Show vbModal, Me
        Exit Sub
    End If
        
    '���ж��豸��AE���˿��Ƿ���ȷ������,δ���ú�����ʾ���˳�
    If IsNull(rsTemp("�˿ں�")) Or IsNull(rsTemp("�豸AE")) Or IsNull(rsTemp("����AE")) Then
        MsgBoxD Me, "�뵽��Ӱ���豸Ŀ¼��ģ���У�����Q/R��ѯʹ�õ��豸�˿ںţ��豸AE�ͱ���AE��", vbInformation, Me.Caption
        Exit Sub
    End If
    
    frmPACSGetDeviceImage.ShowMe Me, rsTemp("IP��ַ"), rsTemp("�˿ں�"), rsTemp("�豸��"), rsTemp("����AE"), rsTemp("�豸AE"), mobjStudyInfo.lngAdviceId
        
        
    If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_GETIMAGE, 1, mobjStudyInfo.lngAdviceId)
Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Menu_Img_��ʦִ��()
'��ʦִ�л�ʦȡ��
On Error GoTo errhandle
    Dim strSQL As String
    Dim intResult As Integer '0-ȡ����1-ִ��
        
    If mobjStudyInfo.lngAdviceId = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_TECHDO, 0, mobjStudyInfo.lngAdviceId)
    
    If mobjStudyInfo.blnIsTechincalSure Then     '��ʦȡ��
        strSQL = "Zl_Ӱ��ʦִ��('" & UserInfo.���� & "'," & mobjStudyInfo.lngAdviceId & ",1)"
        Call zlDatabase.ExecuteProcedure(strSQL, "��ʦȡ��")
        
        mobjStudyInfo.blnIsTechincalSure = False
        
        intResult = 0
    Else
        If mobjStudyInfo.strDoDoctor <> UserInfo.���� Then
            If Not MsgBoxD(Me, "��ǰ��Ա��ָ���ļ�鼼ʦ����ͬ," & vbCrLf & "ȷ��Ҫ����ִ����", vbYesNo, "��ʦִ��") = vbNo Then
                strSQL = "Zl_Ӱ��ʦִ��('" & UserInfo.���� & "'," & mobjStudyInfo.lngAdviceId & ")"
                Call zlDatabase.ExecuteProcedure(strSQL, "��ʦִ��")
                
                mobjStudyInfo.blnIsTechincalSure = True
                mobjStudyInfo.strDoDoctor = UserInfo.����
                
                intResult = 1
            End If
        Else
            strSQL = "Zl_Ӱ��ʦִ��('" & UserInfo.���� & "'," & mobjStudyInfo.lngAdviceId & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, "��ʦִ��")
            
            mobjStudyInfo.blnIsTechincalSure = True
            mobjStudyInfo.strDoDoctor = UserInfo.����
            
            intResult = 1
        End If
    End If
    
    If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_TECHDO, 1, mobjStudyInfo.lngAdviceId, intResult)

    Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Menu_Img_��������豸()
On Error GoTo errhandle
    Dim strModality As String
    Dim rResult As VbMsgBoxResult
    Dim strSQL As String
    
    If mobjStudyInfo.lngAdviceId = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_CHANGEDEVICE, 0, mobjStudyInfo.lngAdviceId)
    
    frmChangeDevice.ShowMe UCase(mobjStudyInfo.strImgType), Me
    strModality = frmChangeDevice.strDeviceType
    
    If strModality <> "" Then
        strSQL = "Zl_Ӱ����_Ӱ�����(" & mobjStudyInfo.lngAdviceId & "," & mobjStudyInfo.lngSendNo & ",'" & strModality & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    
    If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_CHANGEDEVICE, 1, mobjStudyInfo.lngAdviceId)
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub sub��ά�ؽ�(strCommand As String)
    Dim strImageDir As String
    
    Call zlRefreshFace(mobjStudyInfo, False)
    
    '��֯��ά�ؽ���Ҫ��ͼ��
    strImageDir = ZLfun3DImgProcess(mbln3DAutoDecompress)
    If strImageDir <> "" Then
        Call sub3DProcess(strCommand, strImageDir)
    End If
End Sub


Private Sub sub3DProcess(strCommand As String, strImageDir As String)
On Error GoTo errhandle
    Dim str3DCommand As String
    
    '��֯��ά�ؽ����
    str3DCommand = mstr3DExeDir & " " & mstr3DPara & " " & strCommand & " " & strImageDir
    
    Shell str3DCommand
    
errhandle:
End Sub


'ִ�в˵�����
Public Sub zlMenuClick(mnuClick As String)
    
    mblnAddImage = False
    Select Case mnuClick
        Case "Ӱ����"
            DViewer_DblClick
        Case "Ӱ��Ա�"
            mblnAddImage = True
            DViewer_DblClick
        Case "Ӱ��3D��Ƭ"
            Call Open3DViewer(mobjStudyInfo.lngAdviceId, Me, mobjStudyInfo.blnMoved)
        Case "Ӱ����ʾ"
            If Not lvwImage.SelectedItem Is Nothing Then ShowLvwImage lvwImage.SelectedItem
        Case "ȫѡ����"
            If mintSelectAllSeq = 0 Or mintSelectAllSeq = 2 Then
                mintSelectAllSeq = 1
            ElseIf mintSelectAllSeq = 1 Then
                mintSelectAllSeq = 0
            End If
            Call subSetMenuState
            SelectAllSeq True
        Case "ȫ������"
            If mintSelectAllSeq = 0 Or mintSelectAllSeq = 1 Then
                mintSelectAllSeq = 2
            ElseIf mintSelectAllSeq = 2 Then
                mintSelectAllSeq = 0
            End If
            Call subSetMenuState
            SelectAllSeq False
        Case "ȫѡͼ��"
            If mintSelectAllImg = 0 Or mintSelectAllImg = 2 Then
                mintSelectAllImg = 1
            ElseIf mintSelectAllImg = 1 Then
                mintSelectAllImg = 0
            End If
            Call subSetMenuState
            SelectAllImg True
        Case "ȫ��ͼ��"
            If mintSelectAllImg = 0 Or mintSelectAllImg = 1 Then
                mintSelectAllImg = 2
            ElseIf mintSelectAllImg = 2 Then
                mintSelectAllImg = 0
            End If
            Call subSetMenuState
            SelectAllImg False
        Case "��ѡͼ��"
            Dim i As Integer
            With lvwImage
                For i = 1 To .ListItems.Count
                    .ListItems(i).Checked = Not .ListItems(i).Checked
                Next
            End With
            Call WriteSelectdImages(lvwImage.tag)
    End Select
End Sub

Private Sub subSetMenuState()
On Error Resume Next
    If mblnShowPic Then
        Me.cbrMain.FindControl(, conMenu_View_Show).Checked = True
    Else
        Me.cbrMain.FindControl(, conMenu_View_Show).Checked = False
    End If
    
    If mintSelectAllSeq = 0 Then            '0--��״̬
        Me.cbrMain.FindControl(, conMenu_View_Expend_AllCollapse).Checked = False
        Me.cbrMain.FindControl(, conMenu_View_Expend_AllExpend).Checked = False
    ElseIf mintSelectAllSeq = 1 Then        '1--ѡ��ȫ������
        Me.cbrMain.FindControl(, conMenu_View_Expend_AllCollapse).Checked = True
        Me.cbrMain.FindControl(, conMenu_View_Expend_AllExpend).Checked = False
    ElseIf mintSelectAllSeq = 2 Then        '2--��ѡ��ȫ������
        Me.cbrMain.FindControl(, conMenu_View_Expend_AllCollapse).Checked = False
        Me.cbrMain.FindControl(, conMenu_View_Expend_AllExpend).Checked = True
    End If
    
    If mintSelectAllImg = 0 Then            '0--��״̬
        Me.cbrMain.FindControl(, conMenu_Manage_SelectAllImages).Checked = False
        Me.cbrMain.FindControl(, conMenu_Manage_UnSelectAllImages).Checked = False
    ElseIf mintSelectAllImg = 1 Then        '1--ѡ��ȫ��ͼ��
        Me.cbrMain.FindControl(, conMenu_Manage_SelectAllImages).Checked = True
        Me.cbrMain.FindControl(, conMenu_Manage_UnSelectAllImages).Checked = False
    ElseIf mintSelectAllImg = 2 Then        '2--��ѡ��ȫ��ͼ��
        Me.cbrMain.FindControl(, conMenu_Manage_SelectAllImages).Checked = False
        Me.cbrMain.FindControl(, conMenu_Manage_UnSelectAllImages).Checked = True
    End If
    
    
End Sub

Private Sub SelectAllSeq(ByVal blnSelect As Boolean)
    Dim i As Integer
    With lvwSeq
        For i = 1 To .ListItems.Count
            .ListItems(i).Checked = blnSelect
        Next

        'ͼ��������PACS
        If mobjStudyInfo.intImageLocation = 0 Then
            If Not lvwSeq.SelectedItem Is Nothing Then
                ShowImageList lvwSeq.SelectedItem
            Else
                ShowImageList Nothing
            End If
        End If
    End With
End Sub

Private Sub SelectAllImg(ByVal blnSelect As Boolean)
    Dim i As Integer
    With lvwImage
        For i = 1 To .ListItems.Count
            .ListItems(i).Checked = blnSelect
        Next
    End With
    Call WriteSelectdImages(lvwImage.tag)
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_View_Show          '��ʾͼ��
            mblnShowPic = Not mblnShowPic
            Control.Checked = mblnShowPic
            Call zlMenuClick("Ӱ����ʾ")
        Case conMenu_View_Expend_AllCollapse    'ȫѡ����
            Call zlMenuClick("ȫѡ����")
        Case conMenu_View_Expend_AllExpend      'ȫ������
            Call zlMenuClick("ȫ������")
        Case conMenu_Manage_SelectAllImages     'ȫѡͼ��
            Call zlMenuClick("ȫѡͼ��")
        Case conMenu_Manage_UnSelectAllImages   'ȫ��ͼ��
            Call zlMenuClick("ȫ��ͼ��")
        Case conMenu_Manage_ReverseSelectImages '��ѡͼ��
            Call zlMenuClick("��ѡͼ��")
        Case conMenu_View_Refresh
            Call zlRefreshFace(mobjStudyInfo, True)
        Case conMenu_Manage_DeleteSeries        'ɾ������
            Call zlMenuDeleteImageClick(Control.ID)
        Case conMenu_Manage_DeleteImage         'ɾ��ͼ��
            Call zlMenuDeleteImageClick(Control.ID)
    End Select
End Sub

Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
Exit Sub
    Select Case Control.ID
        Case conMenu_View_Expend_AllCollapse, conMenu_View_Expend_AllExpend   'ȫѡ���У�ȫ�����У�

            Control.Enabled = lvwSeq.ListItems.Count > 0
            Control.Checked = False
            
        Case conMenu_Manage_SelectAllImages, conMenu_Manage_UnSelectAllImages, conMenu_Manage_ReverseSelectImages 'ȫѡͼ��ȫ��ͼ�񣬷�ѡͼ��
            Control.Enabled = lvwSeq.ListItems.Count > 0
            Control.Visible = (mobjStudyInfo.intImageLocation = 0)
            Control.Checked = False
            
        Case conMenu_View_Show
            Control.Enabled = lvwSeq.ListItems.Count > 0
            Control.Visible = (mobjStudyInfo.intImageLocation = 0)
            Control.Checked = mblnShowPic
            
        Case conMenu_Manage_ImageInterval   'ͼ����
            Control.Visible = (mobjStudyInfo.intImageLocation = 0)
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = lvwSeq.hwnd
    ElseIf Item.ID = 2 Then
        Item.Handle = lvwImage.hwnd
    ElseIf Item.ID = 3 Then
        Item.Handle = picView.hwnd
    End If
End Sub

Private Sub DViewer_DblClick()
'��ʾ��Ƭվ
    Dim strSerials As String, strSeqUID As String
    Dim Item As MSComctlLib.ListItem
    Dim intImageInverval As Integer
    Dim strImages As String
    Dim rsTemp As ADODB.Recordset
    Dim strFtpURL As String
    
    On Error GoTo CallError
       
    'ͼ�����������ݿ����
    If mobjStudyInfo.intImageLocation = 1 Or mobjStudyInfo.intImageLocation = 2 Then
        strSerials = ""
        
        If mobjStudyInfo.intImageLocation = 1 Then
            If lvwSeq.SelectedItem Is Nothing Then Exit Sub '��ǰ���û��ͼ�񣬾��˳�
            
            
            For Each Item In lvwSeq.ListItems
                strSeqUID = Mid(Item.Key, 2)
                If Item.Checked Then
                    'ֻ�е�ǰ���б���ѡ�ˣ�����ѡ��ɲ���ͼ�����ȫ��ͼ�󣬲Ŵ򿪸�����
                    strSerials = strSerials & ",'" & strSeqUID & "'"
                End If
            Next
            
            strSerials = Mid(strSerials, 2)
        End If
        
        If gblnXWLog = True Then
            Call WriteCommLog("DViewer_DblClick", "����OpenViewer�ӿ�", "���в���Ϊ��" & strSerials)
        End If
    
        Call OpenViewer(1, Nothing, mobjStudyInfo.lngAdviceId, False, Me, strSerials)
        
        Exit Sub
    Else
        'ͼ��������FTP����
        If gblnUseXinWangView = True Then
            '������ϰ汾�����ݣ���ʹ����������Ƭϵͳ����ֱ�Ӵ���Զ��Ŀ¼�ļ���
        
            If lvwSeq.SelectedItem Is Nothing Then Exit Sub '��ǰ���û��ͼ�񣬾��˳�
            
            Set rsTemp = GetStudyImageData(mobjStudyInfo.lngAdviceId, mobjStudyInfo.blnMoved)

            strImages = ""
            For Each Item In lvwSeq.ListItems
                strSeqUID = Mid(Item.Key, 2)
                If Item.Checked Then
                    'ֻ�е�ǰ���б���ѡ�ˣ�����ѡ��ɲ���ͼ�����ȫ��ͼ�󣬲Ŵ򿪸�����
                    rsTemp.Filter = "����UID='" & strSeqUID & "'"
                    While Not rsTemp.EOF
                        If nvl(rsTemp!�豸��1) <> "" Then
                            strFtpURL = "\\" & nvl(rsTemp!Host1) & "\" & gstrImageShareDir & nvl(rsTemp!Root1) & nvl(rsTemp!Url)
                        Else
                            strFtpURL = "\\" & nvl(rsTemp!Host2) & "\" & gstrImageShareDir & nvl(rsTemp!Root2) & nvl(rsTemp!Url)
                        End If

                        If strImages <> "" Then strImages = strImages & "[;]"

                        strFtpURL = Replace(strFtpURL, "//", "/")
                        strImages = strImages & Replace(strFtpURL, "/", "\")

                        rsTemp.MoveNext
                    Wend
                End If
            Next

            '��Զ��Ŀ¼�ļ����жԱȹ�Ƭ
            Call OEMViewOpen(0, strImages, 0, mobjStudyInfo.strImgType)
            
            Exit Sub
        End If
    End If
    
    '--------------------�������ִ�����Exit Sub
    
    'ͼ�����������ݿ⣬ʹ�ù�Ƭվ��ͼ��
    '�ж��Ƿ�򿪵�ǰͼ�������ǰ���û��ͼ��������һ�μ���ͼ��
    If lvwSeq.SelectedItem Is Nothing Then
        Call OpenLatestImage(Me, mobjPacsCore, mobjStudyInfo, mintViewHistoryImageDays)
    Else
        '�����ǡ�����UID1|1-3;5-27;33-100+����UID2|ȫ����,ȫ����ʾ��ȫ��ͼ��
        strImages = ""
        strSerials = ""
        For Each Item In lvwSeq.ListItems
            strSeqUID = Mid(Item.Key, 2)
            If Item.Checked Then
                'ֻ�е�ǰ���б���ѡ�ˣ�����ѡ��ɲ���ͼ�����ȫ��ͼ�󣬲Ŵ򿪸�����
                If Item.SubItems(1) <> "" Then          'Ϊ�ձ�ʾû��ѡ���κ�ͼ��
                    strSerials = strSerials & ",'" & strSeqUID & "'"
                    If strImages = "" Then
                        strImages = strSeqUID & "|" & Item.SubItems(1)
                    Else
                        strImages = strImages & "+" & strSeqUID & "|" & Item.SubItems(1)
                    End If
                End If
            End If
        Next
        If Len(strSerials) = 0 Then         'û��ѡ���κ�����,��Ĭ�ϴ򿪵�ǰ���е�ͼ��
            strSerials = ",'" & Mid(lvwSeq.SelectedItem.Key, 2) & "'"
            If lvwSeq.SelectedItem.SubItems(1) <> "" Then
                strImages = Mid(lvwSeq.SelectedItem.Key, 2) & "|" & lvwSeq.SelectedItem.SubItems(1)
            Else
                strImages = Mid(lvwSeq.SelectedItem.Key, 2) & "|ȫ��"
            End If
        End If
        
        strSerials = Mid(strSerials, 2)
        
        intImageInverval = Val(Me.cbrMain.FindControl(, conMenu_Manage_ImageInterval, , True).Text)
        
        OpenViewer 1, mobjPacsCore, mobjStudyInfo.lngAdviceId, mblnAddImage, Me, strSerials, mobjStudyInfo.blnMoved, intImageInverval, strImages
    End If
    Exit Sub
CallError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub DViewer_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim i As Integer
    If Button <> 1 Then Exit Sub
    
    With DViewer
        i = .ImageIndex(X, Y)
        If i > 0 And i <= .Images.Count And i <> iCurImageIndex Then
            .Images(iCurImageIndex).BorderColour = vbWhite
            .Images(i).BorderColour = vbRed
            iCurImageIndex = i
        End If
    End With
End Sub

Private Sub Form_Activate()
On Error GoTo errhandle
    If Me.tag = "Loading" Then Me.tag = ""
        
errhandle:
End Sub

Private Sub Form_Load()
    Dim objFileSystem As New Scripting.FileSystemObject
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim Pane1 As Pane
    Dim strRegPath As String
    
    '��ȡ���ز���
    strRegPath = "����ģ��\" & App.ProductName & "\frmPacsImg"
    mintSelectAllSeq = Val(GetSetting("ZLSOFT", strRegPath, "SelectAllSeq", 0))
    mintSelectAllImg = Val(GetSetting("ZLSOFT", strRegPath, "SelectAllImg", 0))
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOfficeXP
    Set Me.cbrMain.Icons = zlCommFun.GetPubIcons
    With Me.cbrMain.options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = False
        '.SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.Visible = False
    
    Set cbrToolBar = Me.cbrMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Show, "Ӱ����ʾ")
            cbrControl.iconid = 3061: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "��ʾ��ǰ����Ӱ������ͼ"
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "ȫѡ����")
            cbrControl.iconid = 3010: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "ѡ�е�ǰ��������"
            cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Expend_AllExpend, "ȫ������")
            cbrControl.iconid = 3004: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "���ѡ�е�ǰ��������"
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_SelectAllImages, "ȫѡͼ��")
            cbrControl.iconid = 227: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "ѡ�е�ǰ����ͼ��"
            cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_UnSelectAllImages, "ȫ��ͼ��")
        cbrControl.iconid = 229: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "���ѡ�е�ǰ����ͼ��"
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ReverseSelectImages, "��ѡͼ��")
        cbrControl.iconid = 3012: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "����ѡ������ͼ��"
        Set cbrControl = .Add(xtpControlComboBox, conMenu_Manage_ImageInterval, "ͼ����")
            cbrControl.ToolTipText = "���ô�ͼ��ʱ��ͼ��֮��ļ������"
            cbrControl.AddItem "0"
            cbrControl.AddItem "2"
            cbrControl.AddItem "3"
            cbrControl.AddItem "4"
            cbrControl.AddItem "5"
            cbrControl.AddItem "7"
            cbrControl.AddItem "10"
            cbrControl.ListIndex = 0
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��")
            cbrControl.iconid = 791: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "ˢ�µ�ǰ����ͼ������": cbrControl.flags = xtpFlagRightAlign
    End With
        
    Call subSetMenuState
    
    '�жϵ�ǰ�û��Ƿ���� ��Ƭվ�Ļ���Ȩ��
    mblnObserve = CheckPopedom(";" & GetPrivFunc(glngSys, 1289) & ";", "����")

    With dkpMain
        .SetCommandBars Me.cbrMain
        .options.UseSplitterTracker = False 'ʵʱ�϶�
        .options.ThemedFloatingFrames = False
        .options.AlphaDockingContext = True
        .options.HideClient = True
        Set Pane1 = .CreatePane(1, 0, 300, DockTopOf, Nothing)
            Pane1.Handle = lvwSeq.hwnd
            Pane1.options = PaneNoCaption Or PaneNoCloseable
            
        Set Pane1 = .CreatePane(2, 0, 300, DockBottomOf, Pane1)
            Pane1.Handle = lvwImage.hwnd
            Pane1.options = PaneNoCaption Or PaneNoCloseable
            
        Set Pane1 = .CreatePane(3, 0, 400, DockBottomOf, Nothing)
            Pane1.Handle = picView.hwnd
            Pane1.options = PaneNoCaption Or PaneNoCloseable
    End With
    
    dkpMain.LoadStateFromString GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, "")
    Call RestoreWinState(Me, App.ProductName)
    
'    gblnUseXinWangView = IIf(RegOpenKey(HKEY_CURRENT_USER, "\Software\Silver\Silver Pacs", lngKey) = 0&, True, False) 'IIf(InStr(GetPrivFunc(glngSys, G_LNG_XWPACSVIEW_MODULE), "����") > 0, True, False)
    
   '�����RIS����վ���������������ݿ⣬��ȡ����
    If gblnUseXinWangView Then
        '    ���Ͻػ���Ϣ��hook
'        plngXWPreWndProc = XWHook(mobjOwner.hWnd)
        
        Call XWDBServerOpen
        
        mblnAutoOpenViewer = (Val(zlDatabase.GetPara("XW�Զ��򿪹�Ƭվ", glngSys, G_LNG_XWPACSVIEW_MODULE, 1)) = 1)
        If mblnAutoOpenViewer = True Then
            Call XWADViewerStart
        End If
    End If
End Sub

Private Sub ShowSeqList()
'-----------------------------------------------------------------------------------------
'���ܣ���ѯ�������
'��������
'���أ���
'-----------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim tmpItem As MSComctlLib.ListItem
    Dim strCurKey As String
    
    On Error GoTo DBError
    If Not lvwSeq.SelectedItem Is Nothing Then strCurKey = lvwSeq.SelectedItem.Key
    
    With lvwSeq
        
        If .ColumnHeaders.Count <> 7 Then
            With .ColumnHeaders
                .Clear
                .Add , , "Ӱ�����", 2000
                .Add , , "��ͼ��", 2000
                .Add , , "����", 1000
                .Add , , "���к�", 1000
                .Add , , "ͼ����", 1000
                .Add , , "˵��", 2500
                .Add , , "�ɼ�ʱ��", 2500
            End With
            .ListItems.Add , , "Temp"
        End If
        
        .ListItems.Clear
    End With
    
    strSQL = "Select A.����UID,A.���к�,A.��������,A.�ɼ�ʱ��,B.Ӱ�����,B.����," & _
        " B.���UID,Sum(1) As ͼ���� " & _
        "From Ӱ�������� A,Ӱ�����¼ B,Ӱ����ͼ�� D " & _
        "Where A.���UID=B.���UID  And A.����UID=D.����UID And B.ҽ��ID= [1]  And B.���ͺ�= [2] " & _
        "Group By A.����UID,A.���к�,A.��������,A.�ɼ�ʱ��,B.Ӱ�����,B.����,B.���UID " & _
        "Order By B.Ӱ�����,B.����,A.���к�"
        
    If mobjStudyInfo.blnMoved Then
        strSQL = Replace(strSQL, "Ӱ��������", "HӰ��������")
        strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
        strSQL = Replace(strSQL, "Ӱ����ͼ��", "HӰ����ͼ��")
    End If
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo)
   
    lvwSeq.tag = ""
    If Not rsTmp.EOF Then
        lvwSeq.tag = nvl(rsTmp("���UID"))
        Do While Not rsTmp.EOF
            
            Set tmpItem = lvwSeq.ListItems.Add(, "_" & rsTmp("����UID"), rsTmp("Ӱ�����"))
            With tmpItem
                If mintSelectAllImg = 0 Or mintSelectAllImg = 1 Then
                    .SubItems(1) = "ȫ��"
                Else
                    .SubItems(1) = ""
                End If
                
                .SubItems(2) = nvl(rsTmp("����"))
                .SubItems(3) = nvl(rsTmp("���к�"))
                .SubItems(4) = nvl(rsTmp("ͼ����"), 0)
                .SubItems(5) = nvl(rsTmp("��������"))
                .SubItems(6) = nvl(rsTmp("�ɼ�ʱ��"), date)
                
                If .Key = strCurKey Then .Selected = True
            End With
            rsTmp.MoveNext
        Loop
    End If

    If lvwSeq.Sorted = True Then
        Call lvwSeqSort(lvwSeq.SortKey)
    End If
    
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowImageList(ByVal Item As MSComctlLib.ListItem)
'-----------------------------------------------------------------------------------------
'���ܣ���ѯ�������
'��������
'���أ���
'-----------------------------------------------------------------------------------------
    Dim strSeriesUID As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim tmpItem As MSComctlLib.ListItem
    Dim strCurKey As String
    Dim strOpenImages As String
    Dim ImagesArray() As String
    Dim iSegment As Integer
    Dim iStart As Integer
    Dim iEnd As Integer
    Dim iSegCount As Integer
    
    If Not lvwImage.SelectedItem Is Nothing Then strCurKey = lvwImage.SelectedItem.Key
    With lvwImage
        With .ColumnHeaders
            .Clear
            .Add , , "ͼ���", 2000
            .Add , , "ͼ������", 6000
        End With
        .ListItems.Add , , "Temp"
        .ListItems.Clear
    End With
    
    If Item Is Nothing Then
        Exit Sub
    End If
    
    On Error GoTo err
    strOpenImages = Item.SubItems(1)
    If strOpenImages <> "ȫ��" And strOpenImages <> "" Then
        ImagesArray = Split(strOpenImages, ";")
        iSegment = 0
        iSegCount = UBound(ImagesArray)
        iStart = Split(ImagesArray(iSegment), "-")(0)
        iEnd = Split(ImagesArray(iSegment), "-")(1)
    End If
    strSeriesUID = Mid(Item.Key, 2)
    strSQL = "Select ͼ���,ͼ������,ͼ��UID From Ӱ����ͼ�� Where ����UID = [1] Order By ͼ���"
    If mobjStudyInfo.blnMoved Then
        strSQL = Replace(strSQL, "Ӱ����ͼ��", "HӰ����ͼ��")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡͼ����Ϣ", strSeriesUID)
    
    lvwImage.tag = ""
    If Not rsTmp.EOF Then
        lvwImage.tag = strSeriesUID
        Do While Not rsTmp.EOF
            Set tmpItem = lvwImage.ListItems.Add(, rsTmp("ͼ��UID"), rsTmp("ͼ���"))
            With tmpItem
                .SubItems(1) = nvl(rsTmp("ͼ������"))
                If strOpenImages = "ȫ��" Then
                    tmpItem.Checked = True
                ElseIf strOpenImages = "" Then
                    tmpItem.Checked = False
                Else
                    If rsTmp("ͼ���") >= iStart And rsTmp("ͼ���") <= iEnd Then
                        '��������������Ҫѡ�е�
                        tmpItem.Checked = True
                    ElseIf rsTmp("ͼ���") > iEnd Then
                        '���ڱ�����ֹ���룬��κż�1 �����µ�����ʼ�������ֹ����
                        iSegment = iSegment + 1
                        If iSegment > iSegCount Then
                            tmpItem.Checked = False
                        Else
                            iStart = Split(ImagesArray(iSegment), "-")(0)
                            iEnd = Split(ImagesArray(iSegment), "-")(1)
                            If rsTmp("ͼ���") >= iStart And rsTmp("ͼ���") <= iEnd Then
                                tmpItem.Checked = True
                            Else
                                tmpItem.Checked = False
                            End If
                        End If
                    Else
                        'С�ڱ�����ʼ���룬��ѡ��
                        tmpItem.Checked = False
                    End If
                End If
                If .Key = strCurKey Then .Selected = True
            End With
            rsTmp.MoveNext
        Loop
    End If
    
    DViewer.Images.Clear: iCurImageIndex = 0
    
    If lvwImage.ListItems.Count >= 1 Then
        Call ShowLvwImage(lvwImage.ListItems(1))
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strRegPath As String
    
    Set mObjOwner = Nothing
    Set mObjActiveMenuBar = Nothing
    Set mobjPacsCore = Nothing
    
    strRegPath = "����ģ��\" & App.ProductName & "\frmPacsImg"
    SaveSetting "ZLSOFT", strRegPath, "SelectAllSeq", mintSelectAllSeq
    SaveSetting "ZLSOFT", strRegPath, "SelectAllImg", mintSelectAllImg
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
    Call SaveWinState(Me, App.ProductName)
    
    '�����RIS����վ����Ͽ����������ݿ������
    If gblnUseXinWangView Then
        '    ж��hook
'        XWUnhook mobjOwner.hWnd, plngXWPreWndProc
        
        Call XWDBServerClose
        Call XWADViewerExit
    End If
End Sub

Private Sub lvwImage_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call WriteSelectdImages(lvwImage.tag)
End Sub

Private Sub lvwImage_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    If Item.Checked <> Item.Selected Then
        Item.Checked = Item.Selected
        Call WriteSelectdImages(lvwImage.tag)
    End If
    Call ShowLvwImage(Item)
End Sub

Private Sub ShowLvwImage(ByVal Item As MSComctlLib.ListItem)
    Dim strImageUID As String
    
    If mblnShowPic = False Then
        DViewer.Images.Clear
        Exit Sub
    End If
    
    On Error GoTo DBError
    strImageUID = Item.Key
    '��ȡͼ��DViewer��
    GetAllImages Me, DViewer, mobjStudyInfo.blnMoved, 3, 0, lvwImage.tag, 1, 1, False, "", strImageUID

    If DViewer.Images.Count > 0 Then
        iCurImageIndex = 1
    Else
        iCurImageIndex = 0
    End If
    Exit Sub
DBError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwImage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lvwImage.ListItems.Count >= 1 And Button = 2 Then Call ShowPopupImage(False)
End Sub

Private Sub lvwSeq_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
     Call lvwSeqSort(ColumnHeader.Index - 1)
End Sub

Private Sub lvwSeqSort(intSortKey As Integer)
    Dim i As Integer
    
    lvwSeq.Sorted = False
    lvwSeq.SortKey = intSortKey
    lvwSeq.SortOrder = IIf(lvwSeq.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    
    '����ֵ�͵���������
    If intSortKey = 3 Or intSortKey = 4 Then
        For i = 1 To lvwSeq.ListItems.Count
            lvwSeq.ListItems(i).SubItems(intSortKey) = Format(lvwSeq.ListItems(i).SubItems(intSortKey), "0000000000")
        Next i
        lvwSeq.Sorted = True
        For i = 1 To lvwSeq.ListItems.Count
            lvwSeq.ListItems(i).SubItems(intSortKey) = Val(lvwSeq.ListItems(i).SubItems(intSortKey))
        Next i
    Else
        lvwSeq.Sorted = True
    End If
End Sub

Private Sub lvwSeq_DblClick()
    If Not mblnObserve Then Exit Sub
    If lvwSeq.SelectedItem Is Nothing Then Exit Sub
    DViewer_DblClick
End Sub

Private Sub lvwSeq_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    'ͼ��������PACS����֧�ֶ����е�ѡ��
    If mobjStudyInfo.intImageLocation = 0 Then
        lvwSeq.SelectedItem = Item
        Call ShowImageList(Item)
    End If
End Sub

Private Sub lvwSeq_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'ͼ��������PACS����֧�ֶ����е�ѡ��
On Error GoTo errhandle
    If mobjStudyInfo.intImageLocation = 0 Then
        If Item.Checked <> Item.Selected Then
            Item.Checked = Item.Selected
        End If
        Call ShowImageList(Item)
    Else
        mlngCurImageCount = Item.SubItems(3)
    End If
        
    Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub lvwSeq_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ''ͼ��������PACS����֧��ɾ�����еĵ����˵�
    If mobjStudyInfo.intImageLocation = 0 And lvwSeq.ListItems.Count >= 1 And Button = 2 Then
        Call ShowPopupImage(True)
    End If
End Sub

Private Sub picView_Resize()
On Error GoTo errhandle
    Dim iCols As Integer, iRows As Integer
    
    With DViewer
        .Left = 0: .Top = 0
        .Width = picView.ScaleWidth: .Height = picView.ScaleHeight
        
        If .Images.Count > 0 Then
            ResizeRegion .Images.Count, .Width, .Height, iRows, iCols
            .MultiColumns = iCols: .MultiRows = iRows
        End If
    End With
errhandle:
End Sub

Public Function ZLfun3DImgProcess(blnAutoDecompress As Boolean) As String
'------------------------------------------------
'���ܣ���ά�ؽ�Ԥ�����ƶ���ǰ��ѡ�����е�ͼ��
'������ blnAutoDecompress -- True,���غ��ѹ����False ��ֱ�����ز�������
'���أ�ͼ���ƶ���Ŀ��Ŀ¼������ƶ�ʧ���򷵻ؿ�
'------------------------------------------------

    Dim strSeriesUID As String
    Dim Item As MSComctlLib.ListItem
    Dim iSeriesCount As Integer
    
    On Error GoTo CallError
    If lvwSeq.SelectedItem Is Nothing Then
        MsgBoxD Me, "��ѡ��һ�����н�����ά�ؽ���", vbInformation, Me.Caption
        ZLfun3DImgProcess = ""
        Exit Function
    End If
    
    iSeriesCount = 0
    For Each Item In lvwSeq.ListItems
        If Item.Checked Then
            iSeriesCount = iSeriesCount + 1
            strSeriesUID = Mid(Item.Key, 2)
        End If
    Next
    
    '�ж��Ƿ�ֻ�ж�����б�ѡ����ά�ؽ�һ��ֻ�ܴ���һ������
    If iSeriesCount <> 1 Then
        MsgBoxD Me, "��ѡ��һ�����н�����ά�ؽ���ÿ���ؽ�ֻ��ѡ��һ��ϵ�С�", vbInformation, Me.Caption
        ZLfun3DImgProcess = ""
        Exit Function
    End If
    
    '�ƶ�ָ������UID��ͼ��
    ZLfun3DImgProcess = funMove3DImage(strSeriesUID, mobjStudyInfo.blnMoved, blnAutoDecompress)
    Exit Function
CallError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ZLfun3DImgProcess = ""
End Function

Private Function funMove3DImage(strSeriesUID As String, blnMoved As Boolean, blnDecompress As Boolean) As String
'------------------------------------------------
'���ܣ���һ�����е�ͼ���ƶ���3D��ʱĿ¼�У��ȴ���ά�ؽ�����ĵ���
'������
'       strSeriesUID -- ͼ�������UID
'       blnMoved -- ͼ���Ƿ�ת��
'       blnDecompress -- ����ͼ����Ƿ��ѹ����True����ѹ����False�����غ�������
'���أ�ͼ���ƶ���Ŀ��Ŀ¼������ƶ�ʧ���򷵻ؿ�
'------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim struFtpTag As TFtpConTag
    Dim lngResult As Long
    
    Dim str3DCachePath As String
    Dim strTmpFile As String
    Dim strImageFullPath As String
    Dim dcmImages As New DicomImages
    
    strSQL = "Select A.ͼ���,D.FTP�û��� As User1,D.FTP���� As Pwd1," & _
        "D.IP��ַ As Host1,'/'||D.FtpĿ¼||'/' As Root1," & _
        "Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID As ͼ��Ŀ¼,A.ͼ��UID,d.�豸�� as �豸��1, " & _
        "E.FTP�û��� As User2,E.FTP���� As Pwd2," & _
        "E.IP��ַ As Host2,'/'||E.FtpĿ¼||'/' As Root2," & _
        "e.�豸�� as �豸��2,C.���UID,B.����UID " & _
        "From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
        "Where A.����UID=B.����UID And B.���UID=C.���UID And C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) "
    If blnMoved Then
        strSQL = Replace(strSQL, "Ӱ����ͼ��", "HӰ����ͼ��")
        strSQL = Replace(strSQL, "Ӱ��������", "HӰ��������")
        strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
    End If

    On Error GoTo DBError
    strSQL = strSQL & "And A.����UID= [1] Order By A.ͼ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡͼ��", strSeriesUID)
    
    If rsTmp.RecordCount > 0 Then
        
        '��������Ŀ¼,3Dͼ��Ŀ¼��ǰ׺"App.Path & "\TmpImage\3D"+��������+���UID+����UID
        str3DCachePath = FormatFilePath(GetAppRootPath() & "\Apply\TmpImage\3D\" & Replace(nvl(rsTmp("ͼ��Ŀ¼")), "/", "\") & "\" & strSeriesUID & "\")
        strImageFullPath = FormatFilePath(GetAppRootPath() & "\Apply\TmpImage\" & Replace(nvl(rsTmp("ͼ��Ŀ¼")), "/", "\") & "\")
        MkLocalDir str3DCachePath

        On Error GoTo DBError
        
        Do While Not rsTmp.EOF
            '���3DĿ¼��û��ͼ���ټ�鱾�ػ���Ŀ¼������ٴ�FTP����ͼ��
            If blnDecompress Then
                '����Զ���ѹ�����򱾵�ͼ��Ŀ¼�ļ�����Ҫ�޸�
                strTmpFile = str3DCachePath & "3DTemp"
            Else
                strTmpFile = str3DCachePath & nvl(rsTmp("ͼ��UID"))
            End If
            
            If Dir(strTmpFile) = vbNullString Then  '��ͼ������Ҫ���κβ���
                If Dir(strImageFullPath & nvl(rsTmp("ͼ��UID"))) = vbNullString Then
                    '���ػ���ͼ�񲻴��ڣ����ȡFTPͼ��
                    '����FTP����
                    struFtpTag = FtpTagInstance(nvl(rsTmp("Host1")), _
                                                nvl(rsTmp("User1")), _
                                                nvl(rsTmp("Pwd1")), _
                                                nvl(rsTmp("Root1")) & nvl(rsTmp("ͼ��Ŀ¼")))
                    
                    If Trim(struFtpTag.Ip) = "" Then
                        struFtpTag = FtpTagInstance(nvl(rsTmp("Host2")), _
                                                    nvl(rsTmp("User2")), _
                                                    nvl(rsTmp("Pwd2")), _
                                                    nvl(rsTmp("Root2")) & nvl(rsTmp("ͼ��Ŀ¼")))
                    End If
                    
                    lngResult = FtpDownload(struFtpTag, nvl(rsTmp!ͼ��UID), strTmpFile, False)
                    If lngResult = frAbort Then Exit Function

                Else
                '���ع�Ƭ������ͼ����ڣ�ֱ�Ӹ��Ƶ�3DĿ¼
                    FileCopy strImageFullPath & nvl(rsTmp("ͼ��UID")), strTmpFile
                End If
                
                '����Զ���ѹ��������Ѿ����غõ���ʱ�ļ�����ѹ�����ٱ���
                If blnDecompress Then
                    dcmImages.ReadFile strTmpFile
                    dcmImages(1).WriteFile str3DCachePath & nvl(rsTmp("ͼ��UID")), True
                    dcmImages.Clear
                    Kill strTmpFile
                End If
            End If
            rsTmp.MoveNext
        Loop
    End If
    
    funMove3DImage = str3DCachePath
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
    funMove3DImage = ""
End Function

Private Sub ShowSeqImg()
On Error GoTo err
    '����ͼ�����ڵ�PACSλ�ã����ò�ͬ�Ĺ�������ʾ�����б�
    If mobjStudyInfo.intImageLocation = 1 Or mobjStudyInfo.intImageLocation = 2 Then  'ͼ��������PACS������ƽ̨
    
        If dkpMain.Panes(3).Closed = False Then dkpMain.Panes(3).Closed = True
        If dkpMain.Panes(2).Closed = False Then dkpMain.Panes(2).Closed = True
        
'        lvwImage.Visible = False
'        lvwSeq.Visible = False
        
        Call showXWSeq

        lvwImage.ListItems.Clear
        lvwImage.ColumnHeaders.Clear
        
        DViewer.Images.Clear
    Else
    
        If dkpMain.Panes(3).Closed Then dkpMain.Panes(3).Closed = False
        If dkpMain.Panes(2).Closed Then dkpMain.Panes(2).Closed = False
        
        Call ShowSeqList     '��ʾ����
        
        If lvwSeq.SelectedItem Is Nothing Then
            DViewer.Images.Clear
            Call ShowImageList(Nothing)
        ElseIf mintSelectAllSeq = 0 Then
            lvwSeq_ItemClick lvwSeq.SelectedItem
        ElseIf mintSelectAllSeq = 1 Then
            SelectAllSeq True
        ElseIf mintSelectAllSeq = 2 Then
            SelectAllSeq False
        End If
        
        If lvwImage.SelectedItem Is Nothing Then
            DViewer.Images.Clear
        Else
            ShowLvwImage lvwImage.SelectedItem
        End If
    End If
    
    lvwImage.Enabled = IIf(mobjStudyInfo.intImageLocation = 0, True, False)
    lvwImage.HideColumnHeaders = IIf(mobjStudyInfo.intImageLocation = 0, False, True)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub WriteSelectdImages(strSeriesUID As String)
    Dim i As Integer
    Dim j As Integer
    Dim strOpenImages As String
    Dim blnSelectAll As Boolean
    Dim blnSelectNone As Boolean
    Dim iStart As Integer
    Dim iEnd As Integer
    Dim iSegment As Integer
    
    blnSelectNone = True
    blnSelectAll = True
    For j = 1 To lvwImage.ListItems.Count
        If lvwImage.ListItems(j).Checked = True Then
            blnSelectNone = False
            '��ʼ��¼����
            If iStart <> 0 Then
                iEnd = lvwImage.ListItems(j).Text
            Else
                iStart = lvwImage.ListItems(j).Text
                iEnd = lvwImage.ListItems(j).Text
            End If
        Else
            blnSelectAll = False
            '������¼����
            If iStart <> 0 Then
                iSegment = iSegment + 1
                If strOpenImages = "" Then
                    strOpenImages = iStart & "-" & iEnd
                Else
                    strOpenImages = strOpenImages & ";" & iStart & "-" & iEnd
                End If
                iStart = 0
                iEnd = 0
            End If
        End If
    Next j
    If iStart <> 0 Then
        iSegment = iSegment + 1
        If strOpenImages = "" Then
            strOpenImages = iStart & "-" & iEnd
        Else
            strOpenImages = strOpenImages & ";" & iStart & "-" & iEnd
        End If
    End If
    If blnSelectAll = True Then
        strOpenImages = "ȫ��"
    End If
    If blnSelectNone = True Then
        strOpenImages = ""
    End If
    
    For i = 1 To lvwSeq.ListItems.Count
        If lvwSeq.ListItems(i).Key = "_" & strSeriesUID Then
            lvwSeq.ListItems(i).ListSubItems(1) = strOpenImages
        End If
    Next i
End Sub

Private Sub ShowPopupImage(blnIsSeries As Boolean)
'------------------------------------------------
'���ܣ���������Ҽ������˵�
'������ blnIsSeries -- True ���в˵���False ͼ��˵�
'------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrToolBar As CommandBar
Dim cbrToolPopup As CommandBarPopup
    
    If Not CheckPopedom(mstrPrivs, "���ͼ��") Then Exit Sub
    If mobjStudyInfo.blnMoved Then Exit Sub
    If mblnIsHistoryMode Then Exit Sub
    
    '����Ҽ������˵�
    Set cbrToolBar = cbrMain.Add("����Ҽ�", xtpBarPopup)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        If blnIsSeries Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Manage_DeleteSeries, "ɾ������")
         Else
            Set cbrControl = .Add(xtpControlButton, conMenu_Manage_DeleteImage, "ɾ��ͼ��")
         End If
    End With
    cbrToolBar.Visible = True
    cbrToolBar.ShowPopup
End Sub

Private Sub zlMenuDeleteImageClick(lngControlID As Long)
'------------------------------------------------
'���ܣ�ɾ����ǰѡ�е�ͼ��
'������ lngControlID -- ��ťID
'���ܣ�ɾ��ͼ��
'------------------------------------------------
    Dim i As Integer
    Dim blImgDeleted As Boolean '�Ƿ���ͼ��ɾ��--true ��
    Dim blnIsCancel As Boolean
    
    On Error GoTo err
    blImgDeleted = False
    
    If MsgBoxD(Me, "ȷ��Ҫɾ�����й�ѡ�е�ͼ����", vbOKCancel, "ɾ��ͼ��") = vbCancel Then Exit Sub
    
    If Not mObjNotify Is Nothing Then
        Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 0, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, blnIsCancel)
        If blnIsCancel Then Exit Sub
    End If
    
    If lngControlID = conMenu_Manage_DeleteImage Then
        'ɾ����ǰ��ѡ��ͼ��
        For i = 1 To lvwImage.ListItems.Count
            If lvwImage.ListItems(i).Checked = True Then
                If DeleteImages(Me, 1, lvwImage.ListItems(i).Key, "") = False Then
                    If i = 1 Then
                        Exit Sub
                    Else
                        Exit For
                    End If
                End If
                
                blImgDeleted = True
            End If
        Next
    ElseIf lngControlID = conMenu_Manage_DeleteSeries Then
        'ɾ����ǰ��ѡ������
        For i = 1 To lvwSeq.ListItems.Count
            If lvwSeq.ListItems(i).Checked = True Then
                If DeleteImages(Me, 2, "", Mid(lvwSeq.ListItems(i).Key, 2)) = False Then
                    If i = 1 Then
                        Exit Sub
                    Else
                        Exit For
                    End If
                End If
                
                blImgDeleted = True
            End If
        Next
    End If
        
    'ˢ���б���ʾ
    Call ShowSeqImg
    
    '������һ������Ҳ��ɾ����,Ӧ��ˢ���б�
    If lvwSeq.ListItems.Count = 0 Then
        If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 1, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, -1, hwnd)
    Else
        If Not mObjNotify Is Nothing Then Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 1, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, , hwnd)
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub showXWSeq()
'------------------------------------------------
'���ܣ���ʾ����PACS��ͼ������
'��������
'���أ���
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim tmpItem As MSComctlLib.ListItem
    Dim lngImgCount As Long
    
    On Error GoTo err
    
    If gcnXWDBServer.State <> adStateOpen Then Exit Sub
    
    With lvwSeq
        If .ColumnHeaders.Count <> 6 Then
            With .ColumnHeaders
                .Clear
                .Add , , "Ӱ�����", 2000
                .Add , , "����", 1000
                .Add , , "���к�", 1000
                .Add , , "ͼ����", 900
                .Add , , "˵��", 4000
                .Add , , "�ɼ�ʱ��", 2500
            End With
            .ListItems.Add , , "Temp"
        End If
        .ListItems.Clear
    End With
    
    strSQL = "select F_SER_ID as SERIES����,F_STU_ID as Study����,F_SER_NO as ���к�,F_COUNT_IMG as ͼ����, F_SER_DATE as ��������,F_SER_TIME as ����ʱ��, " _
                & " F_SER_CONTEXT as ��������,F_MODALITY as Ӱ������,F_STU_NO as ҽ��ID from V_OEM_SERIES where F_STU_NO ='" & mobjStudyInfo.lngAdviceId & "' order by F_SER_NO"
    Set rsTemp = gcnXWDBServer.Execute(strSQL)
    
    lngImgCount = 0
    lvwSeq.tag = ""
    If Not rsTemp.EOF Then
        Do While Not rsTemp.EOF
            lngImgCount = lngImgCount + nvl(rsTemp!ͼ����, 0)
            Set tmpItem = lvwSeq.ListItems.Add(, "_" & rsTemp!SERIES����, rsTemp!Ӱ������)
            With tmpItem
                .SubItems(1) = nvl(rsTemp!Study����)
                .SubItems(2) = nvl(rsTemp!���к�)
                .SubItems(3) = nvl(rsTemp!ͼ����)
                .SubItems(4) = nvl(rsTemp!��������)
                .SubItems(5) = Replace(nvl(rsTemp!��������, date), ".", "-") + " " + nvl(rsTemp!����ʱ��, time)
                .Checked = True
            End With
            rsTemp.MoveNext
        Loop
    End If
        
    Exit Sub
    
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

