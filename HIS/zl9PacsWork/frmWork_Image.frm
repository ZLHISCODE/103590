VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWork_Image 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
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
      Bindings        =   "frmWork_Image.frx":0000
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmWork_Image"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IWorkMenu

Private Const M_STR_HINT_NoSelectData As String = "��Ч�ļ�����ݣ���ѡ����Ҫִ�еļ���¼��"
Private Const M_STR_MODULE_MENU_TAG As String = "Ӱ��"

Private mlngModule As Long
Private mstrPrivs As String
Private mlngDepartId As Long
Private mobjOwner As Object

'ͨ�������ģ����治��ʾʱ����Ҫ��ˢ�½����еı������ݣ����Ƕ���ĳЩ���ܣ����Ƭ�ȣ�����Ҫ��ִ�о��幦��ʱ��ˢ�½�������
'�����Ҫ������ˢ�º�ҽ����Ϣ�Ĵ��ݶ���Ϊ������ͬ�ķ����ֱ�ΪzlUpdateAdviceInf��zlRefreshFace

Private mlngTmpAdviceId As Long
Private mlngTmpSendNo As Long

Private mlngAdviceID As Long
Private mlngSendNo As Long
Private mblnMoved As Boolean

Private mlngCurImageCount As Long
Private mstrStudyUID As String
Private mstrModalityType As String
Private mlngStudyState As Long
Private mlngStudyHistoryCount As Long           '��ʷ������

Private mintImageLocation As Integer            '��¼ͼ�����ڵ�λ�ã�0���������ݿ⣻1���������ݿ⣻2���������ݿ�ͼ���ϴ�����ƽ̨
Private mblnAutoOpenViewer As Boolean           '�Ƿ��Զ��򿪹�Ƭ����ADViewer

Private mstrImageLevel As String                'Ӱ�������ȼ���
Private mintImageLevel As Integer               'Ӱ�������ж�
Private mcboStudyHistory As ComboBox            '��ʷ���
Private mintViewHistoryImageDays As Integer     '�Զ�����ʷͼ������

Private mblnShowPic As Boolean
Private mblnAddImage As Boolean                 '�Ƿ�׷��ͼ��

Private mblnLocalizerBackward As Boolean        '��λƬ����
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
Private mblnExamineDoctorVerify As Boolean '��ʦȷ��
Private mstrExamineDoctorName As String    '��ʦ����

Private mobjPacsCore As zl9PacsCore.clsViewer

Private mblnRefreshState As Boolean


'��ȡ��Ҫʹ�õ��ⲿ����
Property Get PacsCore() As Object
    Set PacsCore = mobjPacsCore
End Property

Property Set PacsCore(value As Object)
    Set mobjPacsCore = value
End Property


'��ȡ�˵��ӿڶ���
Property Get zlMenu() As IWorkMenu
    Set zlMenu = Me
End Property


Public Sub NotificationRefresh()
'֪ͨˢ��
    mblnRefreshState = False
End Sub


'�ӿ�ʵ�ֲ���**************************************************************************************************



Public Function IWorkMenu_zlGetModuleMenuId() As Long
'��ȡӰ��˵��Ĳ˵�ID
    IWorkMenu_zlGetModuleMenuId = conMenu_Img_Group
End Function



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
    
    Set mObjActiveMenuBar = objMenuBar
    
    'ɾ��Ӱ�������Ӳ˵�
    Set cbrMenuBar = objMenuBar.FindControl(, conMenu_ManagePopup)
    Set cbrControl = cbrMenuBar.CommandBar.FindControl(, conMenu_Manage_ImageQuality)
    If Not cbrControl Is Nothing Then Call cbrControl.Delete

    Set cbrMenuBar = objMenuBar.FindControl(, conMenu_ManagePopup)
    With cbrMenuBar.CommandBar
        '����Ӱ�������˵�
        If CheckPopedom(mstrPrivs, "Ӱ���ʿ�") Then
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_Manage_ImageQuality, "Ӱ������", "", 0, False, cbrMenuBar.CommandBar.FindControl(, conMenu_Manage_GChannel).Index - 1)
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
             
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_TechDoctorExecute, "��ʦִ��", "ָ����ǰ���ļ�鼼ʦ", 807, True)
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

Public Sub IWorkMenu_zlCreateToolBar(objToolBar As Object)
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
    
    Set cbrLogOut = objToolBar.FindControl(, conMenu_Manage_InQueue)
    
    lngIndex = 4
    If Not cbrLogOut Is Nothing Then lngIndex = cbrLogOut.Index

    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_Img_Look, "��Ƭ", "Ӱ���Ƭ", 8111, True, lngIndex + 1)
    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_Img_Contrast, "�Ա�", "��Ƭ�Ա�", 8112, False, lngIndex + 2)
    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_Img_Look3D, "3D��Ƭ", "3D��Ƭ", 8115, False, lngIndex + 3)
    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_Manage_TechDoctorExecute, "��ʦִ��", "ָ����ǰ���ļ�鼼ʦ", 807, False, lngIndex + 4)
    
    '���������ά�ؽ����ܣ��򴴽���Ӧ�˵�
    If mblnUse3D = True Then
        Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButtonPopup, conMenu_Img_3D, "��ά", "��ά�ؽ�", 8115, False, lngIndex + 5)
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
        Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_Manage_FilmPrint, "��Ƭ��ӡ", "", 3202, False, lngIndex + 6)
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

Public Sub IWorkMenu_zlClearMenu()
'����������Ĳ˵�
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    
    Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Img_Group)
    If Not cbrControl Is Nothing Then Call cbrControl.Delete
    
    'ɾ��Ӱ�������Ӳ˵�
    Set cbrMenuBar = mObjActiveMenuBar.FindControl(, conMenu_ManagePopup)
    Set cbrControl = cbrMenuBar.CommandBar.FindControl(, conMenu_Manage_ImageQuality)
    If Not cbrControl Is Nothing Then Call cbrControl.Delete
End Sub


Public Sub IWorkMenu_zlClearToolBar()
'��������Ĺ�����
    Dim cbrControl As CommandBarControl
    
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

Public Sub IWorkMenu_zlExecuteMenu(ByVal lngMenuId As Long)
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
            frmPACSImageDeviceSetup.Show vbModal, mobjOwner
        Case conMenu_Manage_RefreshImg          'ˢ��ͼ��
            Call zlRefreshFace(True)
        Case conMenu_Manage_ImageFirst, conMenu_Manage_ImageSecond, conMenu_Manage_ImageThird, conMenu_Manage_ImageFourth
            Call Menu_Manage_Ӱ������(lngMenuId, mstrImageLevel)
    End Select
End Sub

Public Sub IWorkMenu_zlUpdateMenu(ByVal control As XtremeCommandBars.ICommandBarControl)
'���²˵�
    Select Case control.ID
        Case conMenu_Img_Look       '��Ƭ����ǰ�����ͼ�񣬻����ǻ�������ʷ��飬����Թ�Ƭ
            control.Enabled = (mstrStudyUID <> "") Or (mlngStudyHistoryCount > 1)
        Case conMenu_Img_Contrast   '��Ƭ�Աȣ�ֻ������PACS����ʾ
            control.Enabled = (mstrStudyUID <> "") Or (mlngStudyHistoryCount > 1)
            control.Visible = IIf(mintImageLocation = 0, True, False)
        
        Case conMenu_Img_Look3D     '3D��Ƭ��ֻ������PACS����ʾ
            control.Enabled = mstrStudyUID <> "" And mlngCurImageCount >= 50
            control.Visible = (mintImageLocation <> 0)
            
        Case conMenu_Manage_FilmPrint                    '��Ƭ��ӡ
            control.Visible = CheckPopedom(mstrPrivs, "��Ƭ�����ӡ")
            
        Case conMenu_Img_3D         '��ά�ؽ�
            If CheckPopedom(mstrPrivs, "��ά�ؽ�����") And mblnUse3D = True Then
                control.Visible = True
            Else
                control.Visible = False
            End If
            
            If control.Visible = True Then control.Enabled = mstrStudyUID <> ""
            
        Case conMenu_Img_Delete '���ͼ��ͼ������ƽ̨������ʾ��ť
            If Not CheckPopedom(mstrPrivs, "���ͼ��") Or (mintImageLocation = 2) Then
                control.Visible = False
            Else
                control.Visible = True
            End If
            
            If control.Visible = True Then control.Enabled = mstrStudyUID <> ""
            
        Case conMenu_Img_Query ',��ȡͼ��ֻ������PACS����ʾ
            If (Not CheckPopedom(mstrPrivs, "���ͼ��")) Or (mintImageLocation <> 0) Then
                control.Visible = False
            Else
                control.Visible = True
            End If
            
            If control.Visible Then control.Enabled = mlngStudyState > 1
            
        Case conMenu_Manage_ChangeDevice    '����Ӱ���豸����
                If mstrModalityType = "CR" Or _
                    mstrModalityType = "DR" Or _
                    mstrModalityType = "DX" Or _
                    mstrModalityType = "RF" Then
                    control.Enabled = True
                Else
                    control.Enabled = False
                End If
        Case conMenu_Manage_TechDoctorExecute   '��ʦִ��
            If mblnExamineDoctorVerify Then control.Caption = "��ʦȡ��" Else control.Caption = "��ʦִ��"
            
            If mlngStudyState >= 2 And mlngStudyState < 5 Then
                control.Enabled = True
                
                If mblnExamineDoctorVerify Then
                    control.Enabled = UserInfo.���� = mstrExamineDoctorName Or CheckPopedom(mstrPrivs, "ȡ����ʦִ��")
                End If
            Else
                control.Enabled = False
            End If
            
        Case conMenu_Manage_DeleteSeries    'ɾ��ѡ������
            control.Enabled = lvwSeq.ListItems.Count > 0 And Me.Visible
        Case conMenu_Manage_DeleteImage     'ɾ��ѡ��ͼ��
            control.Enabled = lvwImage.ListItems.Count > 0 And Me.Visible
        Case conMenu_View_Show, conMenu_View_Expend_AllCollapse, conMenu_View_Expend_AllExpend  'ͼ����ʾ��ȫѡ���У�ȫ������
            control.Enabled = lvwSeq.ListItems.Count > 0 And Me.Visible
            control.Visible = (mintImageLocation = 0) And Me.Visible
            control.Checked = Me.cbrMain.FindControl(, control.ID).Checked
        Case conMenu_Manage_SelectAllImages, conMenu_Manage_UnSelectAllImages, conMenu_Manage_ReverseSelectImages   'ȫѡͼ��ȫ��ͼ�񣬷�ѡͼ��
            control.Enabled = lvwImage.ListItems.Count > 0 And Me.Visible
            control.Visible = (mintImageLocation = 0) And Me.Visible
            control.Checked = Me.cbrMain.FindControl(, control.ID).Checked
        Case conMenu_Img_Group, conMenu_Img_Query, conMenu_View_Refresh, conMenu_Cap_DevSet 'Ӱ��Ӱ���ȡ��ͼ��ˢ��
            control.Enabled = True
        Case conMenu_Manage_ImageFirst, conMenu_Manage_ImageSecond, conMenu_Manage_ImageThird, conMenu_Manage_ImageFourth, conMenu_Manage_ImageQuality
            If Not CheckPopedom(mstrPrivs, "Ӱ���ʿ�") Or mintImageLevel = 0 Then
                control.Visible = False
            ElseIf (mlngStudyState >= 3 And mlngStudyState <= 5) Or mlngStudyState = -1 Then
                control.Visible = True
                control.Enabled = mstrStudyUID <> ""
            Else
                control.Visible = True
                control.Enabled = False
            End If
    End Select
End Sub

Public Sub IWorkMenu_zlPopupMenu(objPopup As XtremeCommandBars.ICommandBar)
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

Public Sub IWorkMenu_zlRefreshSubMenu(objMenuBar As Object)
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
    
    If lngIconId <> 0 Then CreateModuleMenu.IconId = lngIconId
    If blnStartGroup Then CreateModuleMenu.BeginGroup = True
    If strToolTip <> "" Then CreateModuleMenu.ToolTipText = strToolTip
    
    CreateModuleMenu.Category = M_STR_MODULE_MENU_TAG
End Function


Public Sub zlInitModule(ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngDepartId As Long, Optional owner As Object = Nothing)
    mlngModule = lngModule
    mstrPrivs = strPrivs
    mlngDepartId = lngDepartId
    
    If Not owner Is Nothing Then Set mobjOwner = owner
    
    mblnUse3D = Val(zlDatabase.GetPara("������ά�ؽ�", glngSys, lngModule, 0))
    mstr3DExeDir = zlDatabase.GetPara("3D����·��", glngSys, lngModule, "")
    mstr3DPara = zlDatabase.GetPara("3D����", glngSys, lngModule, "")
    mstr3DFunctions = zlDatabase.GetPara("3D����", glngSys, lngModule, "")
    mbln3DAutoDecompress = Val(zlDatabase.GetPara("3D�Զ���ѹ��", glngSys, lngModule, 0))
    mstrImageLevel = NVL(GetDeptPara(mlngDepartId, "Ӱ�������ȼ�", "��,��"))
    mintImageLevel = Val(GetDeptPara(mlngDepartId, "Ӱ�������ж�", "0"))
    mintViewHistoryImageDays = Val(GetDeptPara(mlngDepartId, "�Զ�����ʷͼ������", 1))
    If mintViewHistoryImageDays > 15 Or mintViewHistoryImageDays <= 0 Then
        mintViewHistoryImageDays = 1
    End If
End Sub


Public Sub zlUpdateAdviceInf(ByVal lngAdviceID As Long, ByVal lngSendNO As Long, ByVal lngStudyState As Long, ByVal blnMoved As Boolean)
'ͬ�����ҽ����Ϣ
    mlngAdviceID = lngAdviceID
    mlngSendNo = lngSendNO
    mblnMoved = blnMoved
    mlngStudyState = lngStudyState
    mblnRefreshState = True
    
    Call GetPacsStudyInf(lngAdviceID, lngSendNO, blnMoved)
End Sub

Public Sub zlUpdateOtherInf(cboStudyHistory As Object, blnIsTechincalSure As Boolean, strTechincalDoctor As String)
    '��ʦ�Ƿ�ȷ��
    If blnIsTechincalSure = True Then
        mblnExamineDoctorVerify = True
        mstrExamineDoctorName = strTechincalDoctor
    Else
        mblnExamineDoctorVerify = False
        mstrExamineDoctorName = ""
    End If
    
    '�Զ�����ʷͼ��
    Set mcboStudyHistory = cboStudyHistory
    mlngStudyHistoryCount = mcboStudyHistory.ListCount
End Sub

Public Function zlRefreshFace(Optional blnForceRefresh As Boolean = False) As Boolean
On Error GoTo DBError

    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
       
    
    If (mlngTmpAdviceId = mlngAdviceID And mlngTmpSendNo = mlngSendNo And mblnRefreshState) And Not blnForceRefresh Then Exit Function
        
    mlngTmpAdviceId = mlngAdviceID
    mlngTmpSendNo = mlngSendNo
    mblnRefreshState = True
    
    mblnShowPic = False


    'ת����Ӱ���ܱ��汨��
    If mblnMoved Then
        mstrPrivs = Replace(mstrPrivs, "ͼ���������", "")
        mstrPrivs = Replace(mstrPrivs, "ͼ���ע����", "")
        mstrPrivs = Replace(mstrPrivs, "���ͼ��", "")
    End If
    
    strSql = "select ID ,����ID,������,����ֵ from Ӱ�����̲��� where ����ID =  " & _
             "(Select ִ�в���ID From ����ҽ������ Where ҽ��ID =[1])"
             
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
    While Not rsTemp.EOF
        If rsTemp!������ = "��λƬ����" Then mblnLocalizerBackward = NVL(rsTemp!����ֵ)
        rsTemp.MoveNext
    Wend
    
    Call ShowSeqImg
    
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
End Function


Private Sub GetPacsStudyInf(ByVal lngAdviceID As Long, ByVal lngSendNO As Long, ByVal blnMoved As Boolean)
'��ȡ�����Ϣ
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select a.���UID,a.Ӱ�����,a.ͼ��λ�� from Ӱ�����¼ a where a.ҽ��ID=[1] and a.���ͺ�=[2]"
    
    If blnMoved Then
        strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "", lngAdviceID, lngSendNO)
    
    '����Ĭ��ֵ
    mstrStudyUID = ""
    mstrModalityType = ""
    mintImageLocation = 0
    
    If rsData.RecordCount <= 0 Then Exit Sub
        
    mstrStudyUID = NVL(rsData!���UID)
    mstrModalityType = NVL(rsData!Ӱ�����)
    mintImageLocation = NVL(rsData!ͼ��λ��, 0)
End Sub

Private Sub Menu_Manage_Ӱ������(ByVal lngID As Long, ByVal strImageLevel As String)
On Error GoTo errHandle
    Dim strSql As String
    Dim strResult As String
    Dim strGrades() As String

    If mlngAdviceID <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, frmPacsMain.Caption
        Exit Sub
    End If
    
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

    strSql = "Zl_Ӱ������_Update(" & mlngAdviceID & ",'" & strResult & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "Ӱ������")
    
    Call SendMsgToMainWindow(Me, wetImageQuality, mlngAdviceID, strResult)
Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_Img_��Ƭ()
On Error GoTo errHandle
    
    If mlngAdviceID = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    'ˢ�½���
    Call zlRefreshFace
    
    Call zlMenuClick("Ӱ����")
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Img_3D��Ƭ()
On Error GoTo errHandle
    
    If mlngAdviceID = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    'ˢ�½���
    Call zlRefreshFace
    
    Call zlMenuClick("Ӱ��3D��Ƭ")
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Img_�Աȹ�Ƭ()
On Error GoTo errHandle
    
    If mlngAdviceID = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    'ˢ�½���
    Call zlRefreshFace
    
    Call zlMenuClick("Ӱ��Ա�")
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_Manage_FilmPrint()
'��Ƭ��ӡ
On Error GoTo errHandle
    Dim blnPrintResult As Boolean
    
    '�ж��Ƿ������Ӧ����Ȩ��
    If Not CheckPopedom(mstrPrivs, "��Ƭ�����ӡ") Then
        MsgBoxD Me, "�����߱���Ƭ��ӡȨ�ޣ�����ϵ����Ա��", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    blnPrintResult = XWShowFilmPrintWind(mlngAdviceID, Me)
    
    If blnPrintResult = True Then
        '���ͽ�Ƭ��ӡ��Ϣ����������
        Call SendMsgToMainWindow(Me, wetPrintFilm, mlngAdviceID)
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_Img_ͼ��ɾ��()
On Error GoTo errHandle
    Dim rsTemp As ADODB.Recordset
    
    If mlngAdviceID = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    '���ͼ������ƽ̨mintImageLocation=2������ʾ��ͼ��ɾ������ť��������ɾ��ͼ��
    If mintImageLocation = 1 Then
        'ͼ��������PACS������ArchiveManagerɾ��ͼ��
        Call subXWShowArchiveManager(1)
    ElseIf mintImageLocation = 0 Then   'ͼ��������PACS
        Call zlRefreshFace
            
        gstrSQL = "select ���UID from Ӱ�����¼ where ҽ��ID =[1] and  ���ͺ� = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���UID", mlngAdviceID, mlngSendNo)
        
        If rsTemp.EOF Then Exit Sub
            
        If MsgBoxD(Me, "�Ƿ�ȷ��Ҫɾ���ü�������Ӱ��", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub

        
        'ɾ��Ӱ���ļ���Ŀ¼
        RemoveCheckImages mlngAdviceID, mlngSendNo
        
        gstrSQL = "ZL_Ӱ����_PhotoDelete(" & mlngAdviceID & "," & mlngSendNo & ")"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            
        Call ClearListData
        Call SendMsgToMainWindow(Me, wetDelImg, 0)

        Call SendMsgToMainWindow(Me, wetDelAllImg, mlngAdviceID)
    End If
Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub



Public Sub ReSetFormFontSize(ByVal bytFontSize As Byte)
'����:�������ù���վ����������С
    
    Dim objCtrl As control
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
            Set CtlFont = objCtrl.Options.Font
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = bytFontSize
            Set objCtrl.Options.Font = CtlFont
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
On Error GoTo errHandle
    Dim strImageDeviceNumber As String, rsTemp As ADODB.Recordset

    If mlngAdviceID = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    Call zlRefreshFace
    
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
        MsgBoxD Me, "�뵽��Ӱ���豸Ŀ¼��ģ���У�����Q/R��ѯʹ�õ��豸�˿ںţ��豸AE�ͱ���AE��"
        Exit Sub
    End If
    
    frmPACSGetDeviceImage.ShowMe Me, rsTemp("IP��ַ"), rsTemp("�˿ں�"), rsTemp("�豸��"), rsTemp("����AE"), rsTemp("�豸AE"), mlngAdviceID
        
    Call SendMsgToMainWindow(Me, wetGetImg, mlngAdviceID)
Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_Img_��ʦִ��()
'��ʦִ�л�ʦȡ��
On Error GoTo errHandle
    Dim strSql As String
    Dim intResult As Integer '0-ȡ����1-ִ��
        
    If mlngAdviceID = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If mblnExamineDoctorVerify Then     '��ʦȡ��
        strSql = "Zl_Ӱ��ʦִ��('" & UserInfo.���� & "'," & mlngAdviceID & ",1)"
        Call zlDatabase.ExecuteProcedure(strSql, "��ʦȡ��")
        
        mblnExamineDoctorVerify = False
        
        intResult = 0
    Else
        If mstrExamineDoctorName <> UserInfo.���� Then
            If Not MsgBoxD(Me, "��ǰ��Ա��ָ���ļ�鼼ʦ����ͬ," & vbCrLf & "ȷ��Ҫ����ִ����", vbYesNo, "��ʦִ��") = vbNo Then
                strSql = "Zl_Ӱ��ʦִ��('" & UserInfo.���� & "'," & mlngAdviceID & ")"
                Call zlDatabase.ExecuteProcedure(strSql, "��ʦִ��")
                
                mblnExamineDoctorVerify = True
                mstrExamineDoctorName = UserInfo.����
                
                intResult = 1
            End If
        Else
            strSql = "Zl_Ӱ��ʦִ��('" & UserInfo.���� & "'," & mlngAdviceID & ")"
            Call zlDatabase.ExecuteProcedure(strSql, "��ʦִ��")
            
            mblnExamineDoctorVerify = True
            mstrExamineDoctorName = UserInfo.����
            
            intResult = 1
        End If
    End If
    
    Call SendMsgToMainWindow(Me, wetTechDo, mlngAdviceID, intResult)

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub



Private Sub Menu_Img_��������豸()
On Error GoTo errHandle
    Dim strModality As String
    Dim rResult As VbMsgBoxResult
    Dim strSql As String
    
    If mlngAdviceID = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    frmChangeDevice.ShowMe UCase(mstrModalityType), Me
    strModality = frmChangeDevice.strDeviceType
    
    If strModality <> "" Then
        strSql = "Zl_Ӱ����_Ӱ�����(" & mlngAdviceID & "," & mlngSendNo & ",'" & strModality & "')"
        zlDatabase.ExecuteProcedure strSql, Me.Caption
    End If
    

    Call SendMsgToMainWindow(Me, wetChangeImgType, mlngAdviceID)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub sub��ά�ؽ�(strCommand As String)
    Dim strImageDir As String
    
    Call zlRefreshFace
    
    '��֯��ά�ؽ���Ҫ��ͼ��
    strImageDir = ZLfun3DImgProcess(mbln3DAutoDecompress)
    If strImageDir <> "" Then
        Call sub3DProcess(strCommand, strImageDir)
    End If
End Sub


Private Sub sub3DProcess(strCommand As String, strImageDir As String)
On Error GoTo errHandle
    Dim str3DCommand As String
    
    '��֯��ά�ؽ����
    str3DCommand = mstr3DExeDir & " " & mstr3DPara & " " & strCommand & " " & strImageDir
    
    Shell str3DCommand
    
errHandle:
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
            Call Open3DViewer(mlngAdviceID, Me, mblnMoved)
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
        If mintImageLocation = 0 Then
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

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    Select Case control.ID
        Case conMenu_View_Show          '��ʾͼ��
            mblnShowPic = Not mblnShowPic
            control.Checked = mblnShowPic
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
            Call zlRefreshFace(True)
        Case conMenu_Manage_DeleteSeries        'ɾ������
            Call zlMenuDeleteImageClick(control.ID)
        Case conMenu_Manage_DeleteImage         'ɾ��ͼ��
            Call zlMenuDeleteImageClick(control.ID)
    End Select
End Sub

Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    Select Case control.ID
        Case conMenu_View_Expend_AllCollapse, conMenu_View_Expend_AllExpend   'ȫѡ���У�ȫ�����У�

            control.Enabled = lvwSeq.ListItems.Count > 0
            control.Checked = False
            
        Case conMenu_Manage_SelectAllImages, conMenu_Manage_UnSelectAllImages, conMenu_Manage_ReverseSelectImages 'ȫѡͼ��ȫ��ͼ�񣬷�ѡͼ��
            control.Enabled = lvwSeq.ListItems.Count > 0
            control.Visible = (mintImageLocation = 0)
            control.Checked = False
            
        Case conMenu_View_Show
            control.Enabled = lvwSeq.ListItems.Count > 0
            control.Visible = (mintImageLocation = 0)
            control.Checked = mblnShowPic
            
        Case conMenu_Manage_ImageInterval   'ͼ����
            control.Visible = (mintImageLocation = 0)
    End Select
End Sub

Private Sub DkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = lvwSeq.hWnd
    ElseIf Item.ID = 2 Then
        Item.Handle = lvwImage.hWnd
    ElseIf Item.ID = 3 Then
        Item.Handle = picView.hWnd
    End If
End Sub

Private Sub DViewer_DblClick()
'��ʾ��Ƭվ
    Dim strSerials As String, strSeqUID As String
    Dim Item As MSComctlLib.ListItem
    Dim intImageInverval As Integer
    Dim strImages As String
    Dim rsTemp As ADODB.Recordset
    Dim strFtpUrl As String
    
    On Error GoTo CallError
       
    'ͼ�����������ݿ����
    If mintImageLocation = 1 Or mintImageLocation = 2 Then
        strSerials = ""
        
        If mintImageLocation = 1 Then
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
    
        Call OpenViewer(1, Nothing, mlngAdviceID, False, Me, strSerials)
        Exit Sub
    Else
        'ͼ��������FTP����
        If gblnUseXinWangView = True Then
            '������ϰ汾�����ݣ���ʹ����������Ƭϵͳ����ֱ�Ӵ���Զ��Ŀ¼�ļ���
        
            If lvwSeq.SelectedItem Is Nothing Then Exit Sub '��ǰ���û��ͼ�񣬾��˳�
            
            Set rsTemp = GetStudyImageData(mlngAdviceID, mblnMoved)

            strImages = ""
            For Each Item In lvwSeq.ListItems
                strSeqUID = Mid(Item.Key, 2)
                If Item.Checked Then
                    'ֻ�е�ǰ���б���ѡ�ˣ�����ѡ��ɲ���ͼ�����ȫ��ͼ�󣬲Ŵ򿪸�����
                    rsTemp.Filter = "����UID='" & strSeqUID & "'"
                    While Not rsTemp.EOF
                        If NVL(rsTemp!�豸��1) <> "" Then
                            strFtpUrl = "\\" & NVL(rsTemp!Host1) & "\" & gstrImageShareDir & NVL(rsTemp!Root1) & NVL(rsTemp!Url)
                        Else
                            strFtpUrl = "\\" & NVL(rsTemp!Host2) & "\" & gstrImageShareDir & NVL(rsTemp!Root2) & NVL(rsTemp!Url)
                        End If

                        If strImages <> "" Then strImages = strImages & "[;]"

                        strFtpUrl = Replace(strFtpUrl, "//", "/")
                        strImages = strImages & Replace(strFtpUrl, "/", "\")

                        rsTemp.MoveNext
                    Wend
                End If
            Next

            '��Զ��Ŀ¼�ļ����жԱȹ�Ƭ
            Call OEMViewOpen(0, strImages, 0, mstrModalityType)
            
            Exit Sub
        End If
    End If
    
    '--------------------�������ִ�����Exit Sub
    
    'ͼ�����������ݿ⣬ʹ�ù�Ƭվ��ͼ��
    '�ж��Ƿ�򿪵�ǰͼ�������ǰ���û��ͼ��������һ�μ���ͼ��
    If lvwSeq.SelectedItem Is Nothing Then
        Call OpenLatestImage
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
        
        OpenViewer 1, mobjPacsCore, mlngAdviceID, mblnAddImage, Me, strSerials, mblnMoved, mblnLocalizerBackward, intImageInverval, strImages
    End If
    Exit Sub
CallError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub OpenLatestImage()
'���������������һ��ͼ�����û���������������һ��ͼ����ʾ����б���û�ѡ��
    Dim strSql As String
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim strOrderIDs As String
    Dim lngOrderID As Long
        
    lngOrderID = 0
    If mcboStudyHistory.ListCount <= 1 Then Exit Sub
    
    For i = 0 To mcboStudyHistory.ListCount - 1
        strOrderIDs = strOrderIDs & "," & mcboStudyHistory.ItemData(i)
    Next i
    
    strOrderIDs = Mid(strOrderIDs, 2)
    
    strSql = "select A.ҽ��ID from ����ҽ������ A, Ӱ�����¼ B where A.ҽ��ID = B.ҽ��ID and  B.���UID is not null " _
            & " and  �״�ʱ�� >=sysdate-" & mintViewHistoryImageDays & " and a.ҽ��ID in (" & strOrderIDs & ")"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "�Զ��򿪽���ͼ��")
    If rsTemp.RecordCount >= 1 Then
        lngOrderID = rsTemp!ҽ��ID
    Else
    
        strSql = "select A.ҽ��ID as ID,�״�ʱ�� as ���ʱ��,ҽ������, Ӱ����� from ����ҽ������ A, Ӱ�����¼ B ,����ҽ����¼ C" & _
                 " where A.ҽ��ID = B.ҽ��ID and C.ID=A.ҽ��ID and B.���UID is not null and a.ҽ��ID in (" & strOrderIDs & ") order by �״�ʱ�� desc"
        
        Set rsTemp = zlDatabase.ShowSelect(Me, strSql, 0, "���ͼ��", True, "", "", False, False, False, Screen.Width / 2, Screen.Height / 2)
        If Not rsTemp Is Nothing Then
            If rsTemp.RecordCount >= 1 Then
                lngOrderID = rsTemp!ID
            End If
        End If
    End If
    
    If lngOrderID <> 0 Then
        OpenViewer 1, mobjPacsCore, lngOrderID, mblnAddImage, Me, "", mblnMoved, mblnLocalizerBackward
    End If
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
On Error GoTo errHandle
    If Me.tag = "Loading" Then Me.tag = ""
        
errHandle:
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
    With Me.cbrMain.Options
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
            cbrControl.IconId = 3061: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "��ʾ��ǰ����Ӱ������ͼ"
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "ȫѡ����")
            cbrControl.IconId = 3010: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "ѡ�е�ǰ��������"
            cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Expend_AllExpend, "ȫ������")
            cbrControl.IconId = 3004: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "���ѡ�е�ǰ��������"
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_SelectAllImages, "ȫѡͼ��")
            cbrControl.IconId = 227: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "ѡ�е�ǰ����ͼ��"
            cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_UnSelectAllImages, "ȫ��ͼ��")
        cbrControl.IconId = 229: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "���ѡ�е�ǰ����ͼ��"
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ReverseSelectImages, "��ѡͼ��")
        cbrControl.IconId = 3012: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "����ѡ������ͼ��"
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
            cbrControl.IconId = 791: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "ˢ�µ�ǰ����ͼ������": cbrControl.flags = xtpFlagRightAlign
    End With
        
    Call subSetMenuState
    
    '�жϵ�ǰ�û��Ƿ���� ��Ƭվ�Ļ���Ȩ��
    mblnObserve = CheckPopedom(";" & GetPrivFunc(glngSys, 1289) & ";", "����")

    With DkpMain
        .SetCommandBars Me.cbrMain
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.ThemedFloatingFrames = False
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
        Set Pane1 = .CreatePane(1, 0, 300, DockTopOf, Nothing)
            Pane1.Handle = lvwSeq.hWnd
            Pane1.Options = PaneNoCaption Or PaneNoCloseable
            
        Set Pane1 = .CreatePane(2, 0, 300, DockBottomOf, Pane1)
            Pane1.Handle = lvwImage.hWnd
            Pane1.Options = PaneNoCaption Or PaneNoCloseable
            
        Set Pane1 = .CreatePane(3, 0, 400, DockBottomOf, Nothing)
            Pane1.Handle = picView.hWnd
            Pane1.Options = PaneNoCaption Or PaneNoCloseable
    End With
    
    DkpMain.LoadStateFromString GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(DkpMain), DkpMain.Name, "")
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
    Dim strSql As String, rsTmp As New ADODB.Recordset
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
    
    strSql = "Select A.����UID,A.���к�,A.��������,A.�ɼ�ʱ��,B.Ӱ�����,B.����," & _
        " B.���UID,Sum(1) As ͼ���� " & _
        "From Ӱ�������� A,Ӱ�����¼ B,Ӱ����ͼ�� D " & _
        "Where A.���UID=B.���UID  And A.����UID=D.����UID And B.ҽ��ID= [1]  And B.���ͺ�= [2] " & _
        "Group By A.����UID,A.���к�,A.��������,A.�ɼ�ʱ��,B.Ӱ�����,B.����,B.���UID " & _
        "Order By B.Ӱ�����,B.����,A.���к�"
        
    If mblnMoved Then
        strSql = Replace(strSql, "Ӱ��������", "HӰ��������")
        strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
        strSql = Replace(strSql, "Ӱ����ͼ��", "HӰ����ͼ��")
    End If
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID, mlngSendNo)
   
    lvwSeq.tag = ""
    If Not rsTmp.EOF Then
        lvwSeq.tag = NVL(rsTmp("���UID"))
        Do While Not rsTmp.EOF
            
            Set tmpItem = lvwSeq.ListItems.Add(, "_" & rsTmp("����UID"), rsTmp("Ӱ�����"))
            With tmpItem
                If mintSelectAllImg = 0 Or mintSelectAllImg = 1 Then
                    .SubItems(1) = "ȫ��"
                Else
                    .SubItems(1) = ""
                End If
                
                .SubItems(2) = NVL(rsTmp("����"))
                .SubItems(3) = NVL(rsTmp("���к�"))
                .SubItems(4) = NVL(rsTmp("ͼ����"), 0)
                .SubItems(5) = NVL(rsTmp("��������"))
                .SubItems(6) = NVL(rsTmp("�ɼ�ʱ��"), date)
                
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
    Dim strSql As String
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
    strSql = "Select ͼ���,ͼ������,ͼ��UID From Ӱ����ͼ�� Where ����UID = [1] Order By ͼ���"
    If mblnMoved Then
        strSql = Replace(strSql, "Ӱ����ͼ��", "HӰ����ͼ��")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡͼ����Ϣ", strSeriesUID)
    
    lvwImage.tag = ""
    If Not rsTmp.EOF Then
        lvwImage.tag = strSeriesUID
        Do While Not rsTmp.EOF
            Set tmpItem = lvwImage.ListItems.Add(, rsTmp("ͼ��UID"), rsTmp("ͼ���"))
            With tmpItem
                .SubItems(1) = NVL(rsTmp("ͼ������"))
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
    
    strRegPath = "����ģ��\" & App.ProductName & "\frmPacsImg"
    SaveSetting "ZLSOFT", strRegPath, "SelectAllSeq", mintSelectAllSeq
    SaveSetting "ZLSOFT", strRegPath, "SelectAllImg", mintSelectAllImg
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(DkpMain), DkpMain.Name, DkpMain.SaveStateToString)
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
    GetAllImages Me, DViewer, mblnMoved, 3, 0, lvwImage.tag, 1, 1, False, "", strImageUID

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
    If mintImageLocation = 0 Then
        lvwSeq.SelectedItem = Item
        Call ShowImageList(Item)
    End If
End Sub

Private Sub lvwSeq_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'ͼ��������PACS����֧�ֶ����е�ѡ��
On Error GoTo errHandle
    If mintImageLocation = 0 Then
        If Item.Checked <> Item.Selected Then
            Item.Checked = Item.Selected
        End If
        Call ShowImageList(Item)
    Else
        mlngCurImageCount = Item.SubItems(3)
    End If
        
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub lvwSeq_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ''ͼ��������PACS����֧��ɾ�����еĵ����˵�
    If mintImageLocation = 0 And lvwSeq.ListItems.Count >= 1 And Button = 2 Then
        Call ShowPopupImage(True)
    End If
End Sub

Private Sub picView_Resize()
On Error GoTo errHandle
    Dim iCols As Integer, iRows As Integer
    
    With DViewer
        .Left = 0: .Top = 0
        .Width = picView.ScaleWidth: .Height = picView.ScaleHeight
        
        If .Images.Count > 0 Then
            ResizeRegion .Images.Count, .Width, .Height, iRows, iCols
            .MultiColumns = iCols: .MultiRows = iRows
        End If
    End With
errHandle:
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
        MsgBoxD Me, "��ѡ��һ�����н�����ά�ؽ���"
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
        MsgBoxD Me, "��ѡ��һ�����н�����ά�ؽ���ÿ���ؽ�ֻ��ѡ��һ��ϵ�С�"
        ZLfun3DImgProcess = ""
        Exit Function
    End If
    
    '�ƶ�ָ������UID��ͼ��
    ZLfun3DImgProcess = funMove3DImage(strSeriesUID, mblnMoved, blnAutoDecompress)
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
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    Dim str3DCachePath As String
    Dim strTmpFile As String
    Dim strImageFullPath As String
    Dim dcmImages As New DicomImages
    
    strSql = "Select A.ͼ���,D.FTP�û��� As User1,D.FTP���� As Pwd1," & _
        "D.IP��ַ As Host1,'/'||D.FtpĿ¼||'/' As Root1," & _
        "Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID As ͼ��Ŀ¼,A.ͼ��UID,d.�豸�� as �豸��1, " & _
        "E.FTP�û��� As User2,E.FTP���� As Pwd2," & _
        "E.IP��ַ As Host2,'/'||E.FtpĿ¼||'/' As Root2," & _
        "e.�豸�� as �豸��2,C.���UID,B.����UID " & _
        "From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
        "Where A.����UID=B.����UID And B.���UID=C.���UID And C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) "
    If blnMoved Then
        strSql = Replace(strSql, "Ӱ����ͼ��", "HӰ����ͼ��")
        strSql = Replace(strSql, "Ӱ��������", "HӰ��������")
        strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
    End If

    On Error GoTo DBError
    strSql = strSql & "And A.����UID= [1] Order By A.ͼ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡͼ��", strSeriesUID)
    
    If rsTmp.RecordCount > 0 Then
        
        '��������Ŀ¼,3Dͼ��Ŀ¼��ǰ׺"App.Path & "\TmpImage\3D"+��������+���UID+����UID
        str3DCachePath = App.Path & "\TmpImage\3D\" & Replace(NVL(rsTmp("ͼ��Ŀ¼")), "/", "\") & "\" & strSeriesUID & "\"
        strImageFullPath = App.Path & "\TmpImage\" & Replace(NVL(rsTmp("ͼ��Ŀ¼")), "/", "\") & "\"
        MkLocalDir str3DCachePath

        On Error GoTo DBError
        
        Do While Not rsTmp.EOF
            '���3DĿ¼��û��ͼ���ټ�鱾�ػ���Ŀ¼������ٴ�FTP����ͼ��
            If blnDecompress Then
                '����Զ���ѹ�����򱾵�ͼ��Ŀ¼�ļ�����Ҫ�޸�
                strTmpFile = str3DCachePath & "3DTemp"
            Else
                strTmpFile = str3DCachePath & NVL(rsTmp("ͼ��UID"))
            End If
            
            If Dir(strTmpFile) = vbNullString Then  '��ͼ������Ҫ���κβ���
                If Dir(strImageFullPath & NVL(rsTmp("ͼ��UID"))) = vbNullString Then
                    '���ػ���ͼ�񲻴��ڣ����ȡFTPͼ��
                    '����FTP����
                    If rsTmp("�豸��1") <> vbNullString And Inet1.hConnection = 0 Then
                        If Inet1.FuncFtpConnect(NVL(rsTmp("Host1")), NVL(rsTmp("User1")), NVL(rsTmp("Pwd1"))) = 0 Then
                            If rsTmp("�豸��2") <> vbNullString Then
                                If Inet2.FuncFtpConnect(NVL(rsTmp("Host2")), NVL(rsTmp("User2")), NVL(rsTmp("Pwd2"))) = 0 Then
                                    MsgBoxD Me, "FTP�����������ӣ������������á�"
                                    funMove3DImage = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                    If Inet1.FuncDownloadFile(NVL(rsTmp("Root1")) & rsTmp("ͼ��Ŀ¼"), strTmpFile, rsTmp("ͼ��UID")) <> 0 Then
                        '���豸��1��ȡͼ��ʧ�ܣ�����豸��2��ȡͼ��
                        If rsTmp("�豸��2") <> vbNullString Then
                            If Inet2.hConnection = 0 Then Inet2.FuncFtpConnect NVL(rsTmp("Host2")), NVL(rsTmp("User2")), NVL(rsTmp("Pwd2"))
                            Call Inet2.FuncDownloadFile(NVL(rsTmp("Root2")) & rsTmp("ͼ��Ŀ¼"), strTmpFile, rsTmp("ͼ��UID"))
                        End If
                    End If
                Else
                '���ع�Ƭ������ͼ����ڣ�ֱ�Ӹ��Ƶ�3DĿ¼
                    FileCopy strImageFullPath & NVL(rsTmp("ͼ��UID")), strTmpFile
                End If
                
                '����Զ���ѹ��������Ѿ����غõ���ʱ�ļ�����ѹ�����ٱ���
                If blnDecompress Then
                    dcmImages.ReadFile strTmpFile
                    dcmImages(1).WriteFile str3DCachePath & NVL(rsTmp("ͼ��UID")), True
                    dcmImages.Clear
                    Kill strTmpFile
                End If
            End If
            rsTmp.MoveNext
        Loop
    End If
    Inet1.FuncFtpDisConnect
    Inet2.FuncFtpDisConnect
    funMove3DImage = str3DCachePath
    Exit Function
DBError:
    '�Ͽ�FTP����
    Inet1.FuncFtpDisConnect
    Inet2.FuncFtpDisConnect
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    funMove3DImage = ""
End Function

Private Sub ShowSeqImg()
On Error GoTo err
    '����ͼ�����ڵ�PACSλ�ã����ò�ͬ�Ĺ�������ʾ�����б�
    If mintImageLocation = 1 Or mintImageLocation = 2 Then  'ͼ��������PACS������ƽ̨
        
        lvwImage.Visible = False
        lvwSeq.Visible = False
        
        Call showXWSeq

        lvwImage.ListItems.Clear
        lvwImage.ColumnHeaders.Clear
        
        DViewer.Images.Clear
    Else
    
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
    
    lvwImage.Enabled = IIf(mintImageLocation = 0, True, False)
    lvwImage.HideColumnHeaders = IIf(mintImageLocation = 0, False, True)
    
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
    If mblnMoved Then Exit Sub
    
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
    
    On Error GoTo err
    blImgDeleted = False
    
    If MsgBoxD(Me, "ȷ��Ҫɾ�����й�ѡ�е�ͼ����", vbOKCancel, "ɾ��ͼ��") = vbCancel Then Exit Sub
    
    If lngControlID = conMenu_Manage_DeleteImage Then
        'ɾ����ǰ��ѡ��ͼ��
        For i = 1 To lvwImage.ListItems.Count
            If lvwImage.ListItems(i).Checked = True Then
                Call DeleteImages(Me, 1, lvwImage.ListItems(i).Key, "")
                blImgDeleted = True
            End If
        Next
    ElseIf lngControlID = conMenu_Manage_DeleteSeries Then
        'ɾ����ǰ��ѡ������
        For i = 1 To lvwSeq.ListItems.Count
            If lvwSeq.ListItems(i).Checked = True Then
                Call DeleteImages(Me, 2, "", Mid(lvwSeq.ListItems(i).Key, 2))
                blImgDeleted = True
            End If
        Next
    End If
        
    If blImgDeleted Then '��ͼ��ɾ��
        Call SendMsgToMainWindow(Me, wetDelImg, 0)
    End If
    
    'ˢ���б���ʾ
    Call ShowSeqImg
    
    '������һ������Ҳ��ɾ����,Ӧ��ˢ���б�
    If lvwSeq.ListItems.Count = 0 Then
        Call SendMsgToMainWindow(Me, wetDelAllImg, mlngAdviceID)
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
    Dim strSql As String
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
    
    strSql = "select F_SER_ID as SERIES����,F_STU_ID as Study����,F_SER_NO as ���к�,F_COUNT_IMG as ͼ����, F_SER_DATE as ��������,F_SER_TIME as ����ʱ��, " _
                & " F_SER_CONTEXT as ��������,F_MODALITY as Ӱ������,F_STU_NO as ҽ��ID from V_OEM_SERIES where F_STU_NO ='" & mlngAdviceID & "' order by F_SER_NO"
    Set rsTemp = gcnXWDBServer.Execute(strSql)
    
    lngImgCount = 0
    lvwSeq.tag = ""
    If Not rsTemp.EOF Then
        Do While Not rsTemp.EOF
            lngImgCount = lngImgCount + NVL(rsTemp!ͼ����, 0)
            Set tmpItem = lvwSeq.ListItems.Add(, "_" & rsTemp!SERIES����, rsTemp!Ӱ������)
            With tmpItem
                .SubItems(1) = NVL(rsTemp!Study����)
                .SubItems(2) = NVL(rsTemp!���к�)
                .SubItems(3) = NVL(rsTemp!ͼ����)
                .SubItems(4) = NVL(rsTemp!��������)
                .SubItems(5) = Replace(NVL(rsTemp!��������, date), ".", "-") + " " + NVL(rsTemp!����ʱ��, time)
                .Checked = True
            End With
            rsTemp.MoveNext
        Loop
    End If
        
    Exit Sub
    
err:
    If ErrCenter() = 1 Then Resume
End Sub

