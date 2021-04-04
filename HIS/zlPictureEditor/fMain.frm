VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fMain 
   Caption         =   "λͼ�༭��"
   ClientHeight    =   7050
   ClientLeft      =   2190
   ClientTop       =   4665
   ClientWidth     =   9780
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   9780
   Begin VB.PictureBox picFinal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   6300
      ScaleHeight     =   915
      ScaleWidth      =   960
      TabIndex        =   3
      Top             =   3735
      Visible         =   0   'False
      Width           =   960
   End
   Begin zlPictureEditor.ucCanvas Canvas 
      Height          =   3885
      Left            =   495
      TabIndex        =   2
      Top             =   360
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   6853
   End
   Begin zlPictureEditor.Progress Progress 
      Height          =   375
      Left            =   4500
      TabIndex        =   0
      Top             =   5220
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   6720
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   4780
            MinWidth        =   4039
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1058
            MinWidth        =   1058
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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
   Begin XtremeCommandBars.ImageManager ImageManager 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "fMain.frx":038A
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   7065
      Top             =   1035
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'-- λͼ����
Public WithEvents DIBFilter As cDIBFilter   ' DIB �˾�����(24 bpp)
Attribute DIBFilter.VB_VarHelpID = -1
Public WithEvents DIBDither As cDIBDither   ' DIB ��������(1, 4, 8 bpp)
Attribute DIBDither.VB_VarHelpID = -1
Public DIBPal               As New cDIBPal  ' DIB ��ɫ����� (1, 4, 8 bpp)
Public DIBSave              As New cDIBSave ' Save ���� (BMP)  (1, 4, 8, 24 bpp)
Attribute DIBSave.VB_VarHelpID = -1
Public DIBbpp               As Byte         ' ��ǰ��ɫ���

'-- ������������
Private Const m_UNDO_LEVELS As Long = 25    ' ��� Undo ��Ŀ
Private m_AppID             As Long         ' ����ID (gfrmMain.hwnd)
Private m_UndoPos           As Long         ' ��ǰ Undo λ��
Private m_UndoMax           As Long         ' ���ɳ�����Ŀ
Private m_Temp              As String       ' ��ʱ�ļ���

'-- �Ի���
Private m_LastFilter        As Integer      ' ���ʹ�õ��˾� (�˾������)
Private m_LastFilename      As String       ' ��ǰ�ļ�
Private m_LastPath          As String       ' ��ǰ·��
Private m_Saved             As Boolean      ' �ѱ���
Private m_FileExt           As String       ' ��ǰ�ļ���չ��
Private m_DialogPreview     As Boolean      ' �Ի���: ��ʾԤ��
Private m_DialogFitMode     As Boolean      ' �Ի���: �ʺ�ģʽ
Private m_DialogJPEGquality As Integer      ' �Ի���: JPEG ���� (0-100)

'-- GDI+
Private m_GDIpToken         As Long         ' ���ڹر� GDI+

Private cbp�ļ� As CommandBarPopup
Private cbp�༭ As CommandBarPopup
Private cbp���� As CommandBarPopup
Private cbp��ɫ As CommandBarPopup
Private cbp�˾� As CommandBarPopup
Private cbp��ͼ As CommandBarPopup
Private cbp���� As CommandBarPopup

Private Bar��׼ As CommandBar

Private mblnModeless As Boolean

Public Event pOK(ByRef FinalPicture As StdPicture, ByVal lngWidth As Long, ByVal lngHeight As Long)    '���棬�����޸ĺ����ʱͼƬ·����JPEGͼƬ��
Public Event pCancel()                      'ȡ�����˳�

'################################################################################################################
'## ���ܣ�  �򿪲���ʾ���༭������
'##
'## ������  srcPic      :In     ԴͼƬ
'##         frmParent   :In     ������
'##         blnModeless :In     �Ƿ��Ƿ�ģ̬��Ĭ��Ϊ��ģ̬
'################################################################################################################
Public Sub ShowMe(ByRef srcPic As StdPicture, Optional ByRef frmParent As Object, Optional ByVal blnModeless As Boolean = True)
    If srcPic = 0 Then
        Unload Me
    Else
        '-- Create DIB
        DoEvents
        Screen.MousePointer = vbHourglass
        Call pvSetDIBPicture(srcPic)
        Screen.MousePointer = vbNormal

        '-- Reset Undo/Redo and save first Undo
        Call pvClearAllDIB
        Call pvSaveUndoDIB
        '-- Save info
        m_LastFilename = "[�ڲ�ͼƬ]"
        Call RefreshFileInfo
    End If
    mblnModeless = blnModeless
    Me.Show IIf(blnModeless, vbModeless, vbModal), frmParent
End Sub

Private Sub InitMenus()
    Dim i As Long, j As Long
'    '## ����λ�ûָ�
'    Me.Left = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "MainLeft", (Screen.Width - 12000) / 2)
'    Me.Top = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "MainTop", (Screen.Height - 9000) / 2)
'    Me.Width = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "MainWidth", 12000)
'    Me.Height = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "MainHeight", 9000)

    '## �˵���ʼ��
    Dim cbpPopup As CommandBarPopup                     '��ʱ����
    Dim cbpPopupSub As CommandBarPopup                  '��ʱ����
    Dim objControl As CommandBarControl                 '�������ؼ�
    Dim objCustControl As CommandBarControlCustom       '�Զ���ؼ�
    Dim Combo As CommandBarComboBox                     '������������ؼ�
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBars.Icons = ImageManager.Icons
    
    CommandBars.ActiveMenuBar.Title = "�˵���"
    
    Set cbp�ļ� = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "�ļ�(&F)")
    With cbp�ļ�.CommandBar.Controls
        .Add xtpControlButton, ID_FILE_OPEN, "��(&O)..."
        .Add xtpControlButton, ID_FILE_SAVE, "����(&S)"
        .Add xtpControlButton, ID_FILE_SAVEAS, "���Ϊ(&A)..."
        
        Set objControl = .Add(xtpControlButton, ID_FILE_PRINT, "��ӡ(&P)...")
        objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, ID_FILE_EXIT, "�˳�(&Q)")
        objControl.BeginGroup = True
    End With
    
    Set cbp�༭ = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "�༭(&E)")
    With cbp�༭.CommandBar.Controls
        .Add xtpControlButton, ID_EDIT_UNDO, "����(&U)"
        .Add xtpControlButton, ID_EDIT_REDO, "����(&R)"
        
        Set objControl = .Add(xtpControlButton, ID_EDIT_COPY, "����(&C)")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_EDIT_PASTE, "ճ��(&P)"
        
        Set objControl = .Add(xtpControlButton, ID_EDIT_SIZE, "����ͼ��ߴ�(&S)...")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_EDIT_ORIENT, "������������(&O)..."
    
        Set objControl = .Add(xtpControlButton, ID_EDIT_SCROLLMODE, "��ģʽ(&L)")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_EDIT_CROPMODE, "����ģʽ(&O)"
    End With
    
    Set cbp���� = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "����(&Z)")
    With cbp����.CommandBar.Controls
        .Add xtpControlButton, ID_ZOOM_IN, "�Ŵ�(&I)"
        .Add xtpControlButton, ID_ZOOM_OUT, "��С(&O)"
        .Add xtpControlButton, ID_ZOOM_11, "ʵ�ʳߴ�(&A)"
        
        Set objControl = .Add(xtpControlButton, ID_ZOOM_FIT, "�ʺϴ���(&F)")
        objControl.BeginGroup = True
    End With
    
    Set cbp��ɫ = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "��ɫ(&C)")
    With cbp��ɫ.CommandBar.Controls
        .Add xtpControlButton, ID_COLOR_BLACKWHITE, "�Ҷ�-�ڰ�"
        .Add xtpControlButton, ID_COLOR_GREYS16, "�Ҷ�-16ɫ"
        .Add xtpControlButton, ID_COLOR_GREYS256, "�Ҷ�-256ɫ"
        
        Set objControl = .Add(xtpControlButton, ID_COLOR_COLOR2, "��ɫ-2ɫ")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_COLOR_COLOR16, "��ɫ-16ɫ"
        .Add xtpControlButton, ID_COLOR_COLOR256, "��ɫ-256ɫ"
        
        Set objControl = .Add(xtpControlButton, ID_COLOR_TRUECOLOR, "���ɫ")
        objControl.BeginGroup = True
    End With
    
    Set cbp�˾� = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "�˾�(&L)")
    With cbp�˾�.CommandBar.Controls
        .Add xtpControlButton, ID_ADJUST_BRIGHT, "����(&B)"
        .Add xtpControlButton, ID_ADJUST_CONTRAST, "�Աȶ�(&C)"
        .Add xtpControlButton, ID_ADJUST_SITUATION, "���Ͷ�(&S)"
    
        Set cbpPopup = .Add(xtpControlPopup, 0, "��ɫ(&C)")
        cbpPopup.BeginGroup = True
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_COLOR1, "�Ҷ�(&G)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_COLOR2, "��ƬЧ��(&N)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_COLOR3, "����Ƭ(&O)"
        
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_FILTER_COLOR4, "��ɫ���(&C)...")
        objControl.BeginGroup = True
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_COLOR5, "�滻 &HS..."
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_COLOR6, "�滻 &L..."
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_COLOR7, "�ع����(&M)..."
    
        Set cbpPopup = .Add(xtpControlPopup, 0, "������(&D)")
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_DEF1, "ģ��(&B)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_DEF2, "�ữ(&F)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_DEF3, "��(&S)"
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_FILTER_DEF4, "��ɢ(&D)")
        objControl.BeginGroup = True
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_DEF5, "���ػ�(&P)"
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_FILTER_DEF6, "ȥ��(&K)")
        objControl.BeginGroup = True
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_DEF7, "��һ��ȥ��(&M)"
    
        Set cbpPopup = .Add(xtpControlPopup, 0, "��Ե(&E)")
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_EDGES1, "������Ե(&C)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_EDGES2, "����Ч��(&E)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_EDGES3, "īˮ����(&O)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_EDGES4, "�滭(&R)"
    
        Set cbpPopup = .Add(xtpControlPopup, 0, "����(&S)")
        cbpPopup.BeginGroup = True
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_SPECIAL1, "����(&N)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_SPECIAL2, "ɨ����(&S)"
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_FILTER_SPECIAL3, "����(&D)")
        objControl.BeginGroup = True
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_SPECIAL4, "��ʴ(&E)"
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_FILTER_SPECIAL5, "����(&T)...")
        objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, ID_ADJUST_FILTERBROWSER, "�����˾�(&I)...")
        objControl.BeginGroup = True
    End With

    Set cbp��ͼ = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "��ͼ(&V)")
    With cbp��ͼ.CommandBar.Controls
        .Add xtpControlButton, ID_VIEW_TOOLBARLIST, "�������б�"
        .Add xtpControlButton, ID_VIEW_PANORAMIC, "����ͼ(&V)"
        .Add xtpControlButton, ID_VIEW_PROPERTY, "����(&P)"
    End With
    
    Set cbp���� = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "����(&H)")
    With cbp����.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, ID_HELP_CONTENT, "��������(&H)")
        objControl.BeginGroup = True
        Set cbpPopupSub = .Add(xtpControlPopup, 0, "&Web�ϵ�ҽҵ")
        objControl.BeginGroup = True
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_HELP_ONLINE, "ҽҵ����(&H)"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_HELP_CONTACT, "���ͷ���(&M)"
        Set objControl = .Add(xtpControlButton, ID_HELP_ABOUT, "����(&A)...")
        objControl.BeginGroup = True
    End With
    
    '## ��������ʼ��
    
    Set Bar��׼ = CommandBars.Add("��׼", xtpBarTop)
    With Bar��׼.Controls
        Set objControl = .Add(xtpControlButton, ID_FILE_OPEN, "��")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_FILE_SAVE, "����"
        .Add xtpControlButton, ID_FILE_PRINT, "��ӡ"
        
        Set objControl = .Add(xtpControlButton, ID_EDIT_UNDO, "����")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_EDIT_REDO, "����"
        
        Set objControl = .Add(xtpControlButton, ID_EDIT_SCROLLMODE, "��ģʽ")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_EDIT_CROPMODE, "����ģʽ"
        
        Set objControl = .Add(xtpControlButton, ID_ZOOM_IN, "�Ŵ�")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_ZOOM_OUT, "��С"
        .Add xtpControlButton, ID_ZOOM_11, "ʵ�ʳߴ�"
        
        Set objControl = .Add(xtpControlButton, ID_ZOOM_FIT, "�ʺϴ���")
        objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, ID_EDIT_SIZE, "���ڳߴ�")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_EDIT_ORIENT, "���ڷ���"
    
        Set objControl = .Add(xtpControlButton, ID_FILTER_DEF3, "��")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_FILTER_DEF2, "�ữ"
        .Add xtpControlButton, ID_FILTER_DEF6, "ȥ��"
    
        Set objControl = .Add(xtpControlButton, ID_ADJUST_BRIGHT, "����")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_ADJUST_CONTRAST, "�Աȶ�"
        .Add xtpControlButton, ID_ADJUST_SITUATION, "���Ͷ�"
    
        Set objControl = .Add(xtpControlButton, ID_ADJUST_FILTERBROWSER, "�����˾�...")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_FILTER_SPECIAL5, "����..."
        Set objControl = .Add(xtpControlButton, ID_FILE_SAVE, "����(&S)")
        objControl.BeginGroup = True
        objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, ID_FILE_EXIT, "�˳�(&Q)")
        objControl.Style = xtpButtonIconAndCaption
    End With
        
    CommandBars.KeyBindings.Add FCONTROL, Asc("O"), ID_FILE_OPEN
    CommandBars.KeyBindings.Add FCONTROL, Asc("S"), ID_FILE_SAVE
    CommandBars.KeyBindings.Add FCONTROL, Asc("A"), ID_FILE_SAVEAS
    CommandBars.KeyBindings.Add FCONTROL, Asc("P"), ID_FILE_PRINT
    CommandBars.KeyBindings.Add FCONTROL, Asc("Z"), ID_EDIT_UNDO
    CommandBars.KeyBindings.Add FCONTROL, Asc("Y"), ID_EDIT_REDO
    CommandBars.KeyBindings.Add FCONTROL, Asc("C"), ID_EDIT_COPY
    CommandBars.KeyBindings.Add FCONTROL, Asc("V"), ID_EDIT_PASTE
    CommandBars.KeyBindings.Add FCONTROL, Asc("Q"), ID_FILE_EXIT
    CommandBars.KeyBindings.Add FALT, Asc("Q"), ID_FILE_EXIT
    
    CommandBars.KeyBindings.Add 0, VK_F1, ID_HELP_CONTENT
    CommandBars.KeyBindings.Add 0, VK_F2, ID_ADJUST_BRIGHT
    CommandBars.KeyBindings.Add 0, VK_F3, ID_ADJUST_CONTRAST
    CommandBars.KeyBindings.Add 0, VK_F4, ID_ADJUST_SITUATION
    CommandBars.KeyBindings.Add 0, VK_F6, ID_VIEW_PANORAMIC
    CommandBars.KeyBindings.Add 0, VK_F8, ID_VIEW_PROPERTY
    CommandBars.KeyBindings.Add 0, VK_F12, ID_ADJUST_FILTERBROWSER
    
    '��ʾ��չ��ť
    CommandBars.Options.ShowExpandButtonAlways = True
    CommandBars.EnableCustomization (True)
    CommandBars.Options.UseDisabledIcons = True
End Sub

'################################################################################################################
'## ���ܣ�  ��Ӱ�ť
'################################################################################################################
Public Function AddButton(Controls As CommandBarControls, ControlType As XTPControlType, ID As Long, Caption As String, Optional BeginGroup As Boolean = False, Optional DescriptionText As String = "", Optional ButtonStyle As XTPButtonStyle = xtpButtonAutomatic, Optional Category As String = "Controls") As CommandBarControl
    Dim Control As CommandBarControl
    Set Control = Controls.Add(ControlType, ID, Caption)
    
    Control.BeginGroup = BeginGroup
    Control.DescriptionText = DescriptionText
    Control.Style = ButtonStyle
    Control.Category = Category

    Set AddButton = Control
End Function

Private Sub Canvas_DIBProgressStart()
    BeginShowProgress
End Sub

Private Sub CommandBars_Customization(ByVal Options As XtremeCommandBars.ICustomizeOptions)
    Dim Controls As CommandBarControls
    Set Controls = CommandBars.DesignerControls
    
    If (Controls.Count = 0) Then
        AddButton Controls, xtpControlButton, ID_FILE_OPEN, "��", , "��", xtpButtonAutomatic, "�ļ�"
        AddButton Controls, xtpControlButton, ID_FILE_SAVE, "����", , "����", xtpButtonAutomatic, "�ļ�"
        AddButton Controls, xtpControlButton, ID_FILE_SAVEAS, "���Ϊ", , "���Ϊ", xtpButtonAutomatic, "�ļ�"
        AddButton Controls, xtpControlButton, ID_FILE_PRINT, "��ӡ", , "��ӡ", xtpButtonAutomatic, "�ļ�"
        AddButton Controls, xtpControlButton, ID_FILE_EXIT, "�˳�", , "�˳�", xtpButtonAutomatic, "�ļ�"

        AddButton Controls, xtpControlButton, ID_EDIT_UNDO, "����", , "����", xtpButtonAutomatic, "�༭"
        AddButton Controls, xtpControlButton, ID_EDIT_REDO, "����", , "����", xtpButtonAutomatic, "�༭"
        AddButton Controls, xtpControlButton, ID_EDIT_COPY, "����", , "����", xtpButtonAutomatic, "�༭"
        AddButton Controls, xtpControlButton, ID_EDIT_PASTE, "ճ��", , "ճ��", xtpButtonAutomatic, "�༭"
        AddButton Controls, xtpControlButton, ID_EDIT_SIZE, "����ͼ��ߴ�...", , "����ͼ��ߴ�...", xtpButtonAutomatic, "�༭"
        AddButton Controls, xtpControlButton, ID_EDIT_ORIENT, "������������...", , "������������...", xtpButtonAutomatic, "�༭"
        AddButton Controls, xtpControlButton, ID_EDIT_SCROLLMODE, "��ģʽ", , "��ģʽ", xtpButtonAutomatic, "�༭"
        AddButton Controls, xtpControlButton, ID_EDIT_CROPMODE, "����ģʽ", , "����ģʽ", xtpButtonAutomatic, "�༭"

        AddButton Controls, xtpControlButton, ID_ZOOM_IN, "�Ŵ�", , "�Ŵ�", xtpButtonAutomatic, "����"
        AddButton Controls, xtpControlButton, ID_ZOOM_OUT, "��С", , "��С", xtpButtonAutomatic, "����"
        AddButton Controls, xtpControlButton, ID_ZOOM_11, "ʵ�ʳߴ�", , "ʵ�ʳߴ�", xtpButtonAutomatic, "����"
        AddButton Controls, xtpControlButton, ID_ZOOM_FIT, "�ʺϴ���", , "�ʺϴ���", xtpButtonAutomatic, "����"

        AddButton Controls, xtpControlButton, ID_COLOR_BLACKWHITE, "�Ҷ�-�ڰ�", , "�Ҷ�-�ڰ�", xtpButtonAutomatic, "��ɫ"
        AddButton Controls, xtpControlButton, ID_COLOR_GREYS16, "�Ҷ�-16ɫ", , "�Ҷ�-16ɫ", xtpButtonAutomatic, "��ɫ"
        AddButton Controls, xtpControlButton, ID_COLOR_GREYS256, "�Ҷ�-256ɫ", , "�Ҷ�-256ɫ", xtpButtonAutomatic, "��ɫ"
        AddButton Controls, xtpControlButton, ID_COLOR_COLOR2, "��ɫ-2ɫ", , "��ɫ-2ɫ", xtpButtonAutomatic, "��ɫ"
        AddButton Controls, xtpControlButton, ID_COLOR_COLOR16, "��ɫ-16ɫ", , "��ɫ-16ɫ", xtpButtonAutomatic, "��ɫ"
        AddButton Controls, xtpControlButton, ID_COLOR_COLOR256, "��ɫ-256ɫ", , "��ɫ-256ɫ", xtpButtonAutomatic, "��ɫ"
        AddButton Controls, xtpControlButton, ID_COLOR_TRUECOLOR, "���ɫ", , "���ɫ", xtpButtonAutomatic, "��ɫ"

        AddButton Controls, xtpControlButton, ID_ADJUST_BRIGHT, "����", , "����", xtpButtonAutomatic, "����"
        AddButton Controls, xtpControlButton, ID_ADJUST_CONTRAST, "�Աȶ�", , "�Աȶ�", xtpButtonAutomatic, "����"
        AddButton Controls, xtpControlButton, ID_ADJUST_SITUATION, "���Ͷ�", , "���Ͷ�", xtpButtonAutomatic, "����"
        AddButton Controls, xtpControlButton, ID_ADJUST_FILTERBROWSER, "�����˾�...", , "�����˾�...", xtpButtonAutomatic, "����"

        AddButton Controls, xtpControlButton, ID_FILTER_COLOR1, "�Ҷ�", , "�Ҷ�", xtpButtonAutomatic, "�˾�����ɫ"
        AddButton Controls, xtpControlButton, ID_FILTER_COLOR2, "��ƬЧ��", , "��ƬЧ��", xtpButtonAutomatic, "�˾�����ɫ"
        AddButton Controls, xtpControlButton, ID_FILTER_COLOR3, "����Ƭ", , "����Ƭ", xtpButtonAutomatic, "�˾�����ɫ"
        AddButton Controls, xtpControlButton, ID_FILTER_COLOR4, "��ɫ���...", , "��ɫ���...", xtpButtonAutomatic, "�˾�����ɫ"
        AddButton Controls, xtpControlButton, ID_FILTER_COLOR5, "�滻 HS...", , "�滻 HS...", xtpButtonAutomatic, "�˾�����ɫ"
        AddButton Controls, xtpControlButton, ID_FILTER_COLOR6, "�滻 L...", , "�滻 L...", xtpButtonAutomatic, "�˾�����ɫ"
        AddButton Controls, xtpControlButton, ID_FILTER_COLOR7, "�ع����...", , "�ع����...", xtpButtonAutomatic, "�˾�����ɫ"

        AddButton Controls, xtpControlButton, ID_FILTER_DEF1, "ģ��", , "ģ��", xtpButtonAutomatic, "�˾���������"
        AddButton Controls, xtpControlButton, ID_FILTER_DEF2, "�ữ", , "�ữ", xtpButtonAutomatic, "�˾���������"
        AddButton Controls, xtpControlButton, ID_FILTER_DEF3, "��", , "��", xtpButtonAutomatic, "�˾���������"
        AddButton Controls, xtpControlButton, ID_FILTER_DEF4, "��ɢ", , "��ɢ", xtpButtonAutomatic, "�˾���������"
        AddButton Controls, xtpControlButton, ID_FILTER_DEF5, "���ػ�", , "���ػ�", xtpButtonAutomatic, "�˾���������"
        AddButton Controls, xtpControlButton, ID_FILTER_DEF6, "ȥ��", , "ȥ��", xtpButtonAutomatic, "�˾���������"
        AddButton Controls, xtpControlButton, ID_FILTER_DEF7, "��һ��ȥ��", , "��һ��ȥ��", xtpButtonAutomatic, "�˾���������"

        AddButton Controls, xtpControlButton, ID_FILTER_EDGES1, "������Ե", , "������Ե", xtpButtonAutomatic, "�˾�����Ե"
        AddButton Controls, xtpControlButton, ID_FILTER_EDGES2, "����Ч��", , "����Ч��", xtpButtonAutomatic, "�˾�����Ե"
        AddButton Controls, xtpControlButton, ID_FILTER_EDGES3, "īˮ����", , "īˮ����", xtpButtonAutomatic, "�˾�����Ե"
        AddButton Controls, xtpControlButton, ID_FILTER_EDGES4, "�滭", , "�滭", xtpButtonAutomatic, "�˾�����Ե"

        AddButton Controls, xtpControlButton, ID_FILTER_SPECIAL1, "����", , "����", xtpButtonAutomatic, "�˾�������"
        AddButton Controls, xtpControlButton, ID_FILTER_SPECIAL2, "ɨ����", , "ɨ����", xtpButtonAutomatic, "�˾�������"
        AddButton Controls, xtpControlButton, ID_FILTER_SPECIAL3, "����", , "����", xtpButtonAutomatic, "�˾�������"
        AddButton Controls, xtpControlButton, ID_FILTER_SPECIAL4, "��ʴ", , "��ʴ", xtpButtonAutomatic, "�˾�������"
        AddButton Controls, xtpControlButton, ID_FILTER_SPECIAL5, "����...", , "����...", xtpButtonAutomatic, "�˾�������"

        AddButton Controls, xtpControlButton, ID_VIEW_TOOLBARLIST, "�������б�", , "�������б�", xtpButtonAutomatic, "��ͼ"
        AddButton Controls, xtpControlButton, ID_VIEW_PANORAMIC, "����ͼ", , "����ͼ", xtpButtonAutomatic, "��ͼ"
        AddButton Controls, xtpControlButton, ID_VIEW_PROPERTY, "����", , "����", xtpButtonAutomatic, "��ͼ"

        AddButton Controls, xtpControlButton, ID_HELP_CONTENT, "��������", , "��������", xtpButtonAutomatic, "����"
        AddButton Controls, xtpControlButton, ID_HELP_ONLINE, "ҽҵ����", , "ҽҵ����", xtpButtonAutomatic, "����"
        AddButton Controls, xtpControlButton, ID_HELP_CONTACT, "���ͷ���", , "���ͷ���", xtpButtonAutomatic, "����"
        AddButton Controls, xtpControlButton, ID_HELP_ABOUT, "����...", , "����...", xtpButtonAutomatic, "����"
    End If
End Sub

Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case ID_FILE_OPEN
        '��
        DoFileMenu 0
    Case ID_FILE_SAVE
        '����
        DoFileMenu 1
    Case ID_FILE_SAVEAS
        '���Ϊ
        DoFileMenu 2
    Case ID_FILE_PRINT
        '��ӡ
        DoFileMenu 3
    Case ID_FILE_EXIT
        '�˳�
        DoFileMenu 4
    Case ID_EDIT_UNDO
        '����
        DoEditMenu 0
    Case ID_EDIT_REDO
        '����
        DoEditMenu 1
    Case ID_EDIT_COPY
        '����
        DoEditMenu 2
    Case ID_EDIT_PASTE
        'ճ��
        DoEditMenu 3
    Case ID_EDIT_SIZE
        '�����ߴ�
        DoEditMenu 4
    Case ID_EDIT_ORIENT
        '��������
        DoEditMenu 5
    Case ID_EDIT_SCROLLMODE
        '��ģʽ
        DoEditMenu 6
    Case ID_EDIT_CROPMODE
        '����ģʽ
        DoEditMenu 7
    Case ID_ZOOM_IN
        '�Ŵ�
        Bar��׼.FindControl(, ID_ZOOM_FIT).Checked = False
        DoZoomMenu 0
    Case ID_ZOOM_OUT
        '��С
        Bar��׼.FindControl(, ID_ZOOM_FIT).Checked = False
        DoZoomMenu 1
    Case ID_ZOOM_11
        '1:1
        Bar��׼.FindControl(, ID_ZOOM_FIT).Checked = False
        DoZoomMenu 2
    Case ID_ZOOM_FIT
        '�ʺ�
        DoZoomMenu 3
    Case ID_COLOR_BLACKWHITE
        '�Ҷ�-�ڰ�
        DoColorMenu 0
    Case ID_COLOR_GREYS16
        '�Ҷ�-16ɫ
        DoColorMenu 1
    Case ID_COLOR_GREYS256
        '�Ҷ�-256ɫ
        DoColorMenu 2
    Case ID_COLOR_COLOR2
        '��ɫ-2ɫ
        DoColorMenu 3
    Case ID_COLOR_COLOR16
        '��ɫ-16ɫ
        DoColorMenu 4
    Case ID_COLOR_COLOR256
        '��ɫ-256ɫ
        DoColorMenu 5
    Case ID_COLOR_TRUECOLOR
        '���ɫ
        DoColorMenu 6
    Case ID_ADJUST_BRIGHT
        '����
        DoAdjustMenu 0
    Case ID_ADJUST_CONTRAST
        '�Աȶ�
        DoAdjustMenu 1
    Case ID_ADJUST_SITUATION
        '���Ͷ�
        DoAdjustMenu 2
    Case ID_ADJUST_FILTERBROWSER
        '�˾������
        DoAdjustMenu 3
    Case ID_FILTER_COLOR1
        '��ɫ���Ҷ�
        DoFilterColorMenu 0
    Case ID_FILTER_COLOR2
        '��ɫ����ƬЧ��
        DoFilterColorMenu 1
    Case ID_FILTER_COLOR3
        '��ɫ��ˮ�ۻ�
        DoFilterColorMenu 2
    Case ID_FILTER_COLOR4
        '��ɫ����ɫ���
        DoFilterColorMenu 3
    Case ID_FILTER_COLOR5
        '��ɫ���滻 HS...
        DoFilterColorMenu 4
    Case ID_FILTER_COLOR6
        '��ɫ���滻 L...
        DoFilterColorMenu 5
    Case ID_FILTER_COLOR7
        '��ɫ��λ��
        DoFilterColorMenu 6
    Case ID_FILTER_DEF1
        '�����ȣ�ģ��
        DoFilterDefMenu 0
    Case ID_FILTER_DEF2
        '�����ȣ��ữ
        DoFilterDefMenu 1
    Case ID_FILTER_DEF3
        '�����ȣ���
        DoFilterDefMenu 2
    Case ID_FILTER_DEF4
        '�����ȣ���ɢ
        DoFilterDefMenu 3
    Case ID_FILTER_DEF5
        '�����ȣ����ػ�
        DoFilterDefMenu 4
    Case ID_FILTER_DEF6
        '�����ȣ�ȥ�
        DoFilterDefMenu 5
    Case ID_FILTER_DEF7
        '�����ȣ���һ��ȥ�
        DoFilterDefMenu 6
    Case ID_FILTER_EDGES1
        '��Ե������
        DoFilterEdgesMenu 0
    Case ID_FILTER_EDGES2
        '��Ե������
        DoFilterEdgesMenu 1
    Case ID_FILTER_EDGES3
        '��Ե����ͼ
        DoFilterEdgesMenu 2
    Case ID_FILTER_EDGES4
        '��Ե����Ŀ
        DoFilterEdgesMenu 3
    Case ID_FILTER_SPECIAL1
        '���⣭����
        DoFilterSpecialMenu 0
    Case ID_FILTER_SPECIAL2
        '���⣭ɨ����
        DoFilterSpecialMenu 1
    Case ID_FILTER_SPECIAL3
        '���⣭����
        DoFilterSpecialMenu 2
    Case ID_FILTER_SPECIAL4
        '���⣭��ʴ
        DoFilterSpecialMenu 3
    Case ID_FILTER_SPECIAL5
        '���⣭����...
        DoFilterSpecialMenu 4
    Case ID_VIEW_PANORAMIC
        '����ͼ
        DoViewMenu 1
    Case ID_VIEW_PROPERTY
        '����
        DoViewMenu 2
    Case ID_HELP_CONTENT
        '��������
        ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
    Case ID_HELP_CONTACT
        '���ͷ���
        Call zlMailTo(Me.hwnd)
    Case ID_HELP_ONLINE
        '����ҽҵ
        Call zlHomePage(Me.hwnd)
    Case ID_HELP_ABOUT
        '����...
        ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
    End Select
End Sub

Private Sub CommandBars_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub CommandBars_Resize()
    On Error Resume Next
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    CommandBars.GetClientRect Left, Top, Right, Bottom
    If Right >= Left And Bottom >= Top Then
        Canvas.Move Left, Top, Right - Left, Bottom - Top
    Else
        Canvas.Move 0, 0, 0, 0
    End If
End Sub

Private Sub CommandBars_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim bEnbBPP As Boolean
    Dim bEnbDIB As Boolean
    Dim lIdx    As Long
    
    bEnbBPP = (DIBbpp = 24)             '�������ɫ����
    bEnbDIB = (Canvas.DIB.hDIB <> 0)    '����DIB����
    
    Select Case Control.ID
    Case ID_FILE_OPEN
        '��
    Case ID_FILE_SAVE
        '����
        Control.Enabled = bEnbDIB
    Case ID_FILE_SAVEAS
        '���Ϊ
        Control.Enabled = bEnbDIB
    Case ID_FILE_PRINT
        '��ӡ
        Control.Enabled = bEnbDIB
    Case ID_FILE_EXIT
        '�˳�
    Case ID_EDIT_UNDO
        '����
        Control.Enabled = (m_UndoPos > 1)
    Case ID_EDIT_REDO
        '����
        Control.Enabled = (m_UndoPos <> m_UndoMax)
    Case ID_EDIT_COPY
        '����
        Control.Enabled = (Canvas.DIB.hDIB <> 0)
    Case ID_EDIT_PASTE
        'ճ��
        Control.Enabled = (Clipboard.GetFormat(vbCFBitmap))
    Case ID_EDIT_SIZE
        '�����ߴ�
        Control.Enabled = bEnbDIB
    Case ID_EDIT_ORIENT
        '��������
        Control.Enabled = bEnbDIB
    Case ID_EDIT_SCROLLMODE
        '��ģʽ
        Control.Enabled = (bEnbDIB)
        Control.Checked = (Canvas.WorkMode = [cnvScrollMode])
    Case ID_EDIT_CROPMODE
        '����ģʽ
        Control.Enabled = (bEnbDIB)
        Control.Checked = (Canvas.WorkMode = [cnvCropMode])
    Case ID_ZOOM_IN
        '�Ŵ�
    Case ID_ZOOM_OUT
        '��С
    Case ID_ZOOM_11
        '1:1
    Case ID_ZOOM_FIT
        '�ʺ�
        Control.Checked = (Canvas.FitMode = True)
    Case ID_COLOR_BLACKWHITE
        '�Ҷ�-�ڰ�
        Control.Enabled = bEnbDIB
    Case ID_COLOR_GREYS16
        '�Ҷ�-16ɫ
        Control.Enabled = bEnbDIB
    Case ID_COLOR_GREYS256
        '�Ҷ�-256ɫ
        Control.Enabled = bEnbDIB
    Case ID_COLOR_COLOR2
        '��ɫ-2ɫ
        Control.Enabled = bEnbDIB
    Case ID_COLOR_COLOR16
        '��ɫ-16ɫ
        Control.Enabled = bEnbDIB
    Case ID_COLOR_COLOR256
        '��ɫ-256ɫ
        Control.Enabled = bEnbDIB
    Case ID_COLOR_TRUECOLOR
        '���ɫ
        Control.Enabled = bEnbDIB
    Case ID_ADJUST_BRIGHT
        '����
        Control.Enabled = bEnbBPP
    Case ID_ADJUST_CONTRAST
        '�Աȶ�
        Control.Enabled = bEnbBPP
    Case ID_ADJUST_SITUATION
        '���Ͷ�
        Control.Enabled = bEnbBPP
    Case ID_ADJUST_FILTERBROWSER
        '�˾������
        Control.Enabled = bEnbBPP
    Case ID_FILTER_COLOR1
        '��ɫ���Ҷ�
        Control.Enabled = bEnbBPP
    Case ID_FILTER_COLOR2
        '��ɫ����ƬЧ��
        Control.Enabled = bEnbBPP
    Case ID_FILTER_COLOR3
        '��ɫ��ˮ�ۻ�
        Control.Enabled = bEnbBPP
    Case ID_FILTER_COLOR4
        '��ɫ����ɫ���
        Control.Enabled = bEnbBPP
    Case ID_FILTER_COLOR5
        '��ɫ���滻 HS...
        Control.Enabled = bEnbBPP
    Case ID_FILTER_COLOR6
        '��ɫ���滻 L...
        Control.Enabled = bEnbBPP
    Case ID_FILTER_COLOR7
        '��ɫ��λ��
        Control.Enabled = bEnbBPP
    Case ID_FILTER_DEF1
        '�����ȣ�ģ��
        Control.Enabled = bEnbBPP
    Case ID_FILTER_DEF2
        '�����ȣ��ữ
        Control.Enabled = bEnbBPP
    Case ID_FILTER_DEF3
        '�����ȣ���
        Control.Enabled = bEnbBPP
    Case ID_FILTER_DEF4
        '�����ȣ���ɢ
        Control.Enabled = bEnbBPP
    Case ID_FILTER_DEF5
        '�����ȣ����ػ�
        Control.Enabled = bEnbBPP
    Case ID_FILTER_DEF6
        '�����ȣ�ȥ�
        Control.Enabled = bEnbBPP
    Case ID_FILTER_DEF7
        '�����ȣ���һ��ȥ�
        Control.Enabled = bEnbBPP
    Case ID_FILTER_EDGES1
        '��Ե������
        Control.Enabled = bEnbBPP
    Case ID_FILTER_EDGES2
        '��Ե������
        Control.Enabled = bEnbBPP
    Case ID_FILTER_EDGES3
        '��Ե����ͼ
        Control.Enabled = bEnbBPP
    Case ID_FILTER_EDGES4
        '��Ե����Ŀ
        Control.Enabled = bEnbBPP
    Case ID_FILTER_SPECIAL1
        '���⣭����
        Control.Enabled = bEnbBPP
    Case ID_FILTER_SPECIAL2
        '���⣭ɨ����
        Control.Enabled = bEnbBPP
    Case ID_FILTER_SPECIAL3
        '���⣭����
        Control.Enabled = bEnbBPP
    Case ID_FILTER_SPECIAL4
        '���⣭��ʴ
        Control.Enabled = bEnbBPP
    Case ID_FILTER_SPECIAL5
        '���⣭����...
        Control.Enabled = bEnbBPP
    Case ID_VIEW_PANORAMIC
        '����ͼ
        Control.Checked = gfPanView.Visible
    Case ID_VIEW_PROPERTY
        '����
        Control.Enabled = bEnbDIB
    Case ID_HELP_CONTENT
        '��������
    Case ID_HELP_CONTACT
        '���ͷ���
    Case ID_HELP_ONLINE
        '����ҽҵ
    Case ID_HELP_ABOUT
        '����...
    End Select
End Sub

Private Sub DIBDither_ProgressStart()
    BeginShowProgress
End Sub

Private Sub DIBFilter_ProgressStart()
    BeginShowProgress
End Sub

Private Sub Form_Activate()
    If Me.Enabled And Me.Visible Then Me.SetFocus
End Sub

'========================================================================================
' ������
'========================================================================================
Private Sub Form_Load()
    Dim GpInput As GdiplusStartupInput
    '-- ���� GDI+ Dll
    GpInput.GdiplusVersion = 1
    If (mGDIpEx.GdiplusStartup(m_GDIpToken, GpInput) <> [OK]) Then
        Call MsgBox("���� GDI+ �����޷�����λͼ�༭������ GDI+ DLL �Ƿ���ڻ����𻵣�", vbInformation + vbOKOnly)
        Call Unload(Me)
        Exit Sub
    End If
    '-- �˵���ʼ��
    Call InitMenus
    '-- �ָ�����
    Call mSettings.LoadMainSettings
    
    '-- Initial zoom = 100%
    stbThis.Panels(3).Text = "100%"
    
    '-- Initialize 'evented' objects
    Set DIBFilter = New cDIBFilter
    Set DIBDither = New cDIBDither
    
    '-- Get App. ID and <Temp> path (Undo/Redo temp. files)
    m_AppID = Me.hwnd
    m_Temp = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))

'    '-- Hook wheel for zooming
'    Call mHook.HookWheel(Me.hwnd)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim sRet As VbMsgBoxResult
    
    If (Canvas.DIB.hDIB <> 0 And Not m_Saved And m_UndoPos > 1) Then
    
        sRet = MsgBox("�Ƿ����˳�֮ǰ����ͼƬ��", vbYesNoCancel Or vbInformation)
        Select Case sRet
            Case vbYes    '-- Save
                Call DoFileMenu(1)
                Cancel = 0
            Case vbNo     '-- Don't save
                Cancel = 0
            Case vbCancel '-- Cancel
                Cancel = 1
        End Select
    End If
    If (Cancel = 0) Then Call Unload(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Canvas.DIB.Destroy
    '-- Save settings
    Call mSettings.SaveMainSettings
    
    '-- Delete temp. files
    Call pvClearAllDIB
    
    ' Unload the GDI+ Dll
    Call mGDIpEx.GdiplusShutdown(m_GDIpToken)

    '-- Free objects
    Set DIBFilter = Nothing
    Set DIBDither = Nothing
    Set DIBPal = Nothing
    Set DIBSave = Nothing
    
    RaiseEvent pCancel
End Sub

'========================================================================================
' Processing
'========================================================================================

Public Sub Canvas_DIBProgress(ByVal p As Long)
    Progress.Value = CDbl(p) / Progress.Max
End Sub

Public Sub Canvas_DIBProgressEnd()

    '-- Progress end
    Progress.Value = 0
    Progress.Cls
    Progress.Visible = False
    Call gfPanView.Repaint
    '-- DIB processed (-> 24bpp: Size changed, orientation changed)
    DIBbpp = 24
    Call pvSetPalMode(DIBbpp)
    stbThis.Panels(2).Text = Canvas.DIB.Width & "��" & Canvas.DIB.Height & "��" & DIBbpp & "bpp"
    
    '-- Save Undo
    Call pvSaveUndoDIB
End Sub

Public Sub DIBFilter_Progress(ByVal p As Long)
    Progress.Value = CDbl(p) / Progress.Max
End Sub

Public Sub DIBFilter_ProgressEnd()

    '-- Progress end
    Progress.Value = 0
    Progress.Cls
    Progress.Visible = False
    Call Canvas.Repaint
    Call gfPanView.Repaint
    '-- If not previewing (Filter browse box), save Undo
    If (gfFilter.Previewing = False And gfTexturize.Previewing = False) Then Call pvSaveUndoDIB
End Sub

Public Sub DIBDither_Progress(ByVal p As Long)
    Progress.Value = CDbl(p) / Progress.Max
End Sub

Public Sub DIBDither_ProgressEnd()

    '-- Progress end
    Progress.Value = 0
    Progress.Cls
    Progress.Visible = False
    Call Canvas.Repaint
    Call gfPanView.Repaint
    '-- Save Undo
    Call pvSaveUndoDIB
End Sub

Private Sub RefreshFileInfo()
    '-- Compact path to Textfile panel width
    If (m_LastFilename <> "[δ����]") Then
        Dim strTemp As String
        strTemp = Mid(m_LastFilename, InStrRev(m_LastFilename, "\") + 1)
        stbThis.Panels(1).Text = "�ļ���: " & strTemp ' CompactPath(Me.hDC, m_LastFilename, Info.TextFileWidth)
      Else
        stbThis.Panels(1).Text = "�ļ���: [δ����]"
    End If
End Sub

Private Sub DoFileMenu(Index As Integer)
    '�ļ��˵���ִ���¼�
    Dim fDlg     As New fDialogEx
    Dim sRet     As String
    Dim bSuccess As Boolean
    Select Case Index
        Case 0 '-- Open...
            '-- Show Open Dialog
            sRet = GetFileName(m_LastPath, "����֧�ֵ��ļ���ʽ|*.bmp;*.gif;*.jpg;*.png;*.tif|λͼ��ʽ�ļ� (*.bmp)|*.bmp|GIF ��ʽ�ļ� (*.gif)|*.gif|JPEG ��ʽ�ļ� (*.jpg)|*.jpg|PNG ��ʽ�ļ� (*.png)|*.png|TIFF ��ʽ�ļ� (*.tif)|*.tif", 0, "��...", True, fDlg)
            If (sRet <> vbNullString) Then
                '-- Get last path
                m_LastPath = sRet
                '-- Create DIB
                DoEvents
                Screen.MousePointer = vbHourglass
                Call pvSetDIBPicture(pvGetStdPicture(sRet, bSuccess))
                Screen.MousePointer = vbNormal
                
                If (bSuccess) Then
                    '-- Reset Undo/Redo and save first Undo
                    Call pvClearAllDIB
                    Call pvSaveUndoDIB
                    '-- Save info
                    m_LastFilename = sRet
                    Call RefreshFileInfo
                End If
            End If
        Case 1 '-- Save
'            If (m_LastFilename = "[δ����]" Or (FileFound(pvExtToBMP(m_LastFilename)) And pvExtToBMP(m_LastFilename) <> m_LastFilename)) Then
'                '-- Save as...
'                Call Unload(fDlg)
'                Set fDlg = Nothing
'                Call DoFileMenu(2)
'              Else
'                '-- Save as BMP
'                DoEvents
'                Call DIBSave.Save_BMP(pvExtToBMP(m_LastFilename), Canvas.DIB, DIBPal, DIBDither, DIBbpp)
'                '-- Saved flag
'                m_Saved = True
'                '-- Save info
'                m_LastFilename = pvExtToBMP(m_LastFilename)
'                Call RefreshFileInfo
'            End If
            Dim strFileName As String
            strFileName = m_Temp & "\R" & m_AppID & ".jpg"
            Call pvCorrectExt(strFileName)
            Call mGDIpEx.SaveDIB(gfrmMain.Canvas.DIB, strFileName, [ImageJPEG], 100)         '100%��ͼƬ����
            If FileFound(strFileName) Then
                Set picFinal.Picture = LoadPicture(strFileName)
                RaiseEvent pOK(picFinal.Picture, picFinal.Width, picFinal.Height)
                m_Saved = True
                Err = 0: On Error Resume Next: Kill strFileName
            Else
                MsgBox "����ʧ�ܣ�", vbOKOnly + vbInformation, "zlPictureEditor"
            End If
        Case 2 '-- Save as...
            '-- Show Open Dialog
            sRet = GetFileName(m_LastFilename, "Bitmap (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif", 0, "���Ϊ...", False, fDlg)
            If (sRet <> vbNullString) Then
                '-- Missing ext.?
                Call pvCorrectExt(sRet)
                '-- Save...
                DoEvents
                Select Case m_FileExt
                    Case "*.bmp" '-- BMP
                        Call DIBSave.Save_BMP(sRet, Canvas.DIB, DIBPal, DIBDither, DIBbpp)
                    Case "*.gif" '-- GIF
                        Call mGDIpEx.SaveDIB(gfrmMain.Canvas.DIB, sRet, [ImageGIF])
                    Case "*.jpg" '-- JPEG
                        Call mGDIpEx.SaveDIB(gfrmMain.Canvas.DIB, sRet, [ImageJPEG], fDlg.txtJPEGQuality)
                    Case "*.png" '-- PNG
                        Call mGDIpEx.SaveDIB(gfrmMain.Canvas.DIB, sRet, [ImagePNG])
                    Case "*.tif" '-- TIFF
                        Call mGDIpEx.SaveDIB(gfrmMain.Canvas.DIB, sRet, [ImageTIFF])
                End Select
                '-- Saved flag
                m_Saved = True
                '-- Save info
                If (m_FileExt = "*.bmp") Then
                    m_LastFilename = sRet
                End If
                Call RefreshFileInfo
            End If
        Case 3 '-- Print...
            If (Printers.Count) Then
                Call gfPrint.Show(vbModal, Me)
              Else
                Call MsgBox("�Բ���û�а�װ��ӡ����", vbExclamation)
            End If
        Case 4 '-- Exit
            Call Unload(Me)
    End Select
    
    Call Unload(fDlg)
    Set fDlg = Nothing
End Sub

Private Sub mnuEditTop_Click()
End Sub

Private Sub DoEditMenu(Index As Integer)
    Dim sRet As VbMsgBoxResult
    Select Case Index
        Case 0 '-- Undo
            Call Undo
        Case 1 '-- Redo
            Call Redo
        Case 2 '-- Copy
            If (Canvas.DIB.hDIB <> 0) Then
                Call Canvas.DIB.CopyToClipboard
            End If
        Case 3 '-- Paste
            If (Clipboard.GetFormat(vbCFBitmap)) Then
                '-- Something there ?
                If (Canvas.DIB.hDIB <> 0 And Not m_Saved And m_UndoPos > 1) Then
                    '-- Ask for save
                    sRet = MsgBox("�Ƿ���ճ��ǰ����ı䣿", vbYesNoCancel Or vbInformation)
                    Select Case sRet
                        Case vbYes    '-- Save
                            Call DoFileMenu(1)
                        Case vbNo     '-- Ignore
                        Case vbCancel '-- Exit
                            Exit Sub
                    End Select
                End If
                '-- Initialize DIB
                Call pvSetDIBPicture(Clipboard.GetData(vbCFBitmap))
                '-- Reset Undo/Redo and save first Undo
                Call pvClearAllDIB
                Call pvSaveUndoDIB
                '-- [δ����] image
                m_LastFilename = "[δ����]"
            End If
        Case 4 '-- Resize
            Call gfResize.Show(vbModal, Me)
        Case 5 '-- Orientation
            Call gfOrientation.Show(vbModal, Me)
        Case 6 '-- Scroll mode
            Canvas.WorkMode = [cnvScrollMode]
            Call Canvas.Repaint
        Case 7 '-- Crop mode
            Canvas.WorkMode = [cnvCropMode]
    End Select
End Sub

Public Sub DoZoomMenu(Index As Integer)
    Select Case Index
        Case 0 '-- Zoom +
            Canvas.Zoom = Canvas.Zoom + IIf(Canvas.Zoom < 25, 1, 0)
            Canvas.FitMode = False
        Case 1 '-- Zoom -
            Canvas.Zoom = Canvas.Zoom - IIf(Canvas.Zoom > 1, 1, 0)
            Canvas.FitMode = False
        Case 2 '-- 1 : 1
            Canvas.Zoom = 1
            Canvas.FitMode = False
        Case 3 '-- Best fit
            Canvas.FitMode = True
    End Select
    Call Canvas.Resize
    stbThis.Panels(3).Text = Format(Canvas.Zoom, "0%")
End Sub

Private Sub DoColorMenu(Index As Integer)
    Dim sDIB As New cDIB
    Dim bfW As Long, bfH As Long
    Dim bfx As Long, bfy As Long
    
    Select Case Index
        Case 0  '-- Black and White (Stucki)
            DIBbpp = 1
            Call DIBPal.CreateBlackAndWhite
            Call DIBDither.DitherToBlackAndWhite(Canvas.DIB, 16)
        Case 1  '-- 16 greys
            DIBbpp = 4
            Call DIBPal.CreateGreys_016
            Call DIBDither.DitherToGreyPalette(Canvas.DIB, DIBPal, True)
        Case 2  '-- 256 greys
            DIBbpp = 8
            Call DIBPal.CreateGreys_256
            Call DIBDither.DitherToGreyPalette(Canvas.DIB, DIBPal, False)
        Case 3  '-- 2 colors
            DIBbpp = 1
            Call DIBPal.CreateBlackAndWhite
            Call DIBDither.DitherToBlackAndWhite(Canvas.DIB, 16)
        Case 4  '-- 16 colors
            DIBbpp = 4
            If (DIBPal.IsGreyScale) Then
                Call DIBPal.CreateGreys_016
                Call DIBDither.DitherToGreyPalette(Canvas.DIB, DIBPal, True)
              Else
                '-- Strecth to fit 150x150 (This will speed up all this)
                Call Canvas.DIB.GetBestFitInfo(150, 150, bfx, bfy, bfW, bfH)
                Call sDIB.Create(bfW, bfH)
                Call sDIB.LoadDIBBlt(Canvas.DIB)
                '-- Create optimal palette and dither.
                '   I don't know why these weight coeffs. work well...
                '   wChannel = f(Lchannel)
                '   wR = 1/(3-0.222) = 0.360
                '   wG = 1/(3-0.707) = 0.436
                '   wB = 1/(3-0.071) = 0.341
                Screen.MousePointer = vbHourglass
                Call DIBPal.CreateOptimal(sDIB, 16, 8, 0.36, 0.436, 0.341)
                Screen.MousePointer = vbNormal
                Call DIBDither.DitherToColorPalette(Canvas.DIB, DIBPal, True)
            End If
        Case 5  '-- 256 colors
            DIBbpp = 8
                If (DIBPal.IsGreyScale) Then
                Call DIBPal.CreateGreys_256
                Call DIBDither.DitherToGreyPalette(Canvas.DIB, DIBPal, False)
              Else
                '-- Strecth to fit 150x150 (This will speed up all this)
                Call Canvas.DIB.GetBestFitInfo(150, 150, bfx, bfy, bfW, bfH)
                Call sDIB.Create(bfW, bfH)
                Call sDIB.LoadDIBBlt(Canvas.DIB)
                '-- Create optimal palette and dither
                Screen.MousePointer = vbHourglass
                Call DIBPal.CreateOptimal(sDIB, 256, 8, 1, 1, 1)
                Screen.MousePointer = vbNormal
                Call DIBDither.DitherToColorPalette(Canvas.DIB, DIBPal, True)
            End If
        Case 6  '-- True color (24bpp)
            DIBbpp = 24
            Call DIBPal.Clear
            Call DIBDither.DitherToTrueColor(Canvas.DIB)
    End Select
    '-- Refresh
    Call Canvas.Repaint
    '-- Select current mode
    Call pvSetPalMode(DIBbpp)
    '-- Update info
    stbThis.Panels(2).Text = Canvas.DIB.Width & "��" & Canvas.DIB.Height & "��" & DIBbpp & "bpp"
End Sub

Private Sub DoAdjustMenu(Index As Integer)
    Select Case Index
    Case 0 '-- Brightness...
        Call gfFilter.Initialize(fltBrightness)
        Call gfFilter.Show(vbModal, Me)
    Case 1 '-- Contrast...
        Call gfFilter.Initialize(fltContrast)
        Call gfFilter.Show(vbModal, Me)
    Case 2 '-- Saturation...
        Call gfFilter.Initialize([fltSaturation])
        Call gfFilter.Show(vbModal, Me)
    Case 3 '-- Filter browser...
        Call gfFilter.Initialize(m_LastFilter)
        Call gfFilter.Show(vbModal, Me)
    End Select
    Call Canvas.Repaint
End Sub

Private Sub DoFilterColorMenu(Index As Integer)
    Select Case Index
    Case 0 '-- Greys
        Call DIBFilter.Greys(Canvas.DIB)
    Case 1 '-- Negative
        Call DIBFilter.Negative(Canvas.DIB)
    Case 2 '-- Sepia
        Call DIBFilter.Colorize(Canvas.DIB, 0.5, 0.25)
    Case 3 '-- Colorize...
        Call gfFilter.Initialize([fltColorize])
        Call gfFilter.Show(vbModal, Me)
    Case 4 '-- Replace HS...
        Call gfFilter.Initialize([fltReplaceHS])
        Call gfFilter.Show(vbModal, Me)
    Case 5 '-- Replace L...
        Call gfFilter.Initialize([fltReplaceL])
        Call gfFilter.Show(vbModal, Me)
    Case 6 '-- Shift...
        Call gfFilter.Initialize(fltShift)
        Call gfFilter.Show(vbModal, Me)
    End Select
    
    Call Canvas.Repaint
End Sub

Private Sub DoFilterDefMenu(Index As Integer)
    Select Case Index
    Case 0 '-- Blur
        Call DIBFilter.Blur(Canvas.DIB)
    Case 1 '-- Soften
        Call DIBFilter.Soften(Canvas.DIB)
    Case 2 '-- Sharpen
        Call DIBFilter.Sharpen(Canvas.DIB)
    Case 3 '-- Diffuse
        Call DIBFilter.Diffuse(Canvas.DIB)
    Case 4 '-- Pixelize
        Call DIBFilter.Pixelize(Canvas.DIB)
    Case 5 '-- Despeckle
        Call DIBFilter.Despeckle(Canvas.DIB)
    Case 6 '-- Despeckle more
        Call DIBFilter.DespeckleMore(Canvas.DIB)
    End Select
    
    Call Canvas.Repaint
End Sub

Private Sub DoFilterEdgesMenu(Index As Integer)
    Select Case Index
    Case 0 '-- Contour
        Call DIBFilter.Contour(Canvas.DIB)
    Case 1 '-- Emboss
        Call DIBFilter.Emboss(Canvas.DIB)
    Case 2 '-- Outline
        Call DIBFilter.Outline(Canvas.DIB)
    Case 3 '-- Relieve
        Call DIBFilter.Relieve(Canvas.DIB)
    End Select
    
    Call Canvas.Repaint
End Sub

Private Sub DoFilterSpecialMenu(Index As Integer)
    Select Case Index
    Case 0 '-- Noise
        Call DIBFilter.Noise(Canvas.DIB)
    Case 1 '-- Scanlines
        Call DIBFilter.Scanlines(Canvas.DIB)
    Case 2 '-- Dilate (Max.R.F. - 4N)
        Call DIBFilter.RankFilterMaximum(Canvas.DIB)
    Case 3 '-- Erode (Min.R.F. - 4N)
        Call DIBFilter.RankFilterMinimum(Canvas.DIB)
    Case 4 '-- Texturize...
        Call gfTexturize.Show(vbModal, Me)
    End Select
    
    Call Canvas.Repaint
End Sub

Private Sub DoViewMenu(Index As Integer)
    Select Case Index
    Case 1 '-- Panoramic view
        If gfPanView.Visible = False Then
            Call gfPanView.Show(IIf(mblnModeless, vbModeless, vbModal), Me)
          Else
            Call gfPanView.Hide
        End If
    Case 2 '-- Properties...
        Call gfProperties.Show(vbModal, Me)
    End Select
End Sub

'========================================================================================
' ���� �����¼�����������
'========================================================================================

Private Sub Canvas_Resize()
    Call gfPanView.Repaint
End Sub

Private Sub Canvas_Scroll()
    Call gfPanView.Repaint
End Sub

Private Sub Canvas_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim scrHMax As Long, scrVMax As Long
  Dim scrHPos As Long, scrVPos As Long

  Dim bScroll As Boolean
  Dim lInc    As Long
    
    With Canvas
        Select Case KeyCode
            Case vbKeyAdd      '{NumPad +}
                Call DoZoomMenu(0)
            Case vbKeySubtract '{NumPad -}
                Call DoZoomMenu(1)
            Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
                Call .GetScrollInfo(scrHMax, scrVMax, scrHPos, scrVPos)
                bScroll = True
        End Select
                    
        If (bScroll) Then
            lInc = 10 * Canvas.Zoom
            Select Case KeyCode
                Case vbKeyUp    '{Cursor Up}
                    If (scrVPos > 0) Then
                        Call .SetScrollInfo(scrHPos, scrVPos - lInc)
                      Else
                        Call .SetScrollInfo(scrHPos, 0)
                    End If
                Case vbKeyDown  '{Cursor Down}
                    If (scrVPos < scrVMax) Then
                        Call .SetScrollInfo(scrHPos, scrVPos + lInc)
                      Else
                        Call .SetScrollInfo(scrHPos, scrVMax)
                    End If
                Case vbKeyLeft  '{Cursor Left}
                    If (scrHPos > 0) Then
                        Call .SetScrollInfo(scrHPos - lInc, scrVPos)
                      Else
                        Call .SetScrollInfo(0, scrVPos)
                    End If
                Case vbKeyRight '{Cursor Right}
                    If (scrHPos < scrHMax) Then
                        Call .SetScrollInfo(scrHPos + lInc, scrVPos)
                      Else
                        Call .SetScrollInfo(scrHMax, scrVPos)
                    End If
            End Select
            Call gfPanView.Repaint
        End If
        
        Call Canvas.Repaint
    End With
End Sub

Private Sub Canvas_Crop()
    '-- Change to True color mode
    DIBbpp = 24
    Call pvSetPalMode(DIBbpp)
    
    '-- Update Info and Progress
    With Canvas.DIB
        stbThis.Panels(2).Text = .Width & "��" & .Height & "��" & DIBbpp & "bpp"
        Progress.Max = .Height
    End With
    
    '-- Refresh Panoramic view
    Call gfPanView.Repaint
    
    '-- Save Undo
    Call pvSaveUndoDIB
End Sub

'========================================================================================
' DIB/��ɫ�� ��ʼ��
'========================================================================================

Private Function pvGetStdPicture(ByVal sFilename As String, bSuccess As Boolean) As StdPicture
    On Error Resume Next
    If (pvGetExt(sFilename) = "png" Or pvGetExt(sFilename) = "tif") Then
        '-- Use GDI+ loading
        Set pvGetStdPicture = mGDIpEx.LoadPictureEx(sFilename)
      Else
        '-- Use VB LoadPicture
        Set pvGetStdPicture = LoadPicture(sFilename)
    End If
    
    '-- Is there an image ?
    bSuccess = Not (pvGetStdPicture Is Nothing)
    
    If (bSuccess = False) Then
        '-- Nothing loaded
        Call MsgBox("����ͼƬʱ�����������", vbExclamation)
    End If
    
    On Error GoTo 0
End Function
    
Private Sub pvSetDIBPicture(Image As StdPicture)
  Static lstW As Long
  Static lstH As Long

    If (Not Picture Is Nothing) Then

        '-- Save last DIB dimensions
        lstW = Canvas.DIB.Width
        lstH = Canvas.DIB.Height
        
        '-- Clear palette
        Call DIBPal.Clear
        
        '-- Create 32bpp DIB section from std. picture.
        '   Case source <=8bpp, palette saved in DIBPal, palette indexes in DIBDither.
        '   Return value: source color depth / 0 = Err.
        DIBbpp = Canvas.DIB.CreateFromStdPicture(Image, DIBPal, DIBDither)
        
        '-- Select current depth mode
        Call pvSetPalMode(DIBbpp)
        
        '-- Remove Crop rectangle and resize canvas
        Call Canvas.RemoveCropRectangle
        With Canvas.DIB
            If (lstW <> .Width Or lstH <> .Height) Then
                Call Canvas.Resize
              Else
                Call Canvas.Repaint
            End If
        End With
        
        '-- Refresh panoramic view
        Call gfPanView.Repaint
        
        '-- Set progress bar max value
        Progress.Max = Canvas.DIB.Height
        
        '-- Show image info: Size + bpp
        stbThis.Panels(2).Text = Canvas.DIB.Width & "��" & Canvas.DIB.Height & "��" & DIBbpp & "bpp"
        stbThis.Panels(3).Text = Format(Canvas.Zoom, "0%")
    End If
End Sub

Private Sub pvSetPalMode(ByVal bpp As Long)
  Dim lIdxNew As Long
  Dim lIdxOld As Long
    
    Select Case bpp
        Case 1  '-- 2 colors / Black and White
            lIdxNew = IIf(DIBPal.IsGreyScale, 0, 4)
        Case 4  '-- 16 colors / 16 greys
            lIdxNew = IIf(DIBPal.IsGreyScale, 1, 5)
        Case 8  '-- 256 colors / 256 greys
            lIdxNew = IIf(DIBPal.IsGreyScale, 2, 6)
        Case 24 '-- True color
            lIdxNew = 8
        Case Else
            Exit Sub
    End Select
'
'    For lIdxOld = 0 To 8
'        mnuColors(lIdxOld).Checked = False
'    Next lIdxOld
'    mnuColors(lIdxNew).Checked = True
    
End Sub

'========================================================================================
' ����/���� ����
'========================================================================================

Public Sub Undo()
    Dim sPath As String
    If (m_UndoPos > 1) Then
        '-- Get path
        sPath = m_Temp & "\b" & m_AppID & Format(m_UndoPos - 2, "000") & ".dat"
        '-- Load Undo DIB
        Call pvSetDIBPicture(LoadPicture(sPath))
        '-- Refresh Panoramic view
        Call gfPanView.Repaint
    
        If (m_UndoPos > 0) Then
            m_UndoPos = m_UndoPos - 1
        End If
    End If
End Sub

Public Sub Redo()
    Dim sPath As String
    If (m_UndoPos < m_UndoMax) Then
        '-- Get path
        sPath = m_Temp & "\b" & m_AppID & Format(m_UndoPos, "000") & ".dat"
        '-- Load Redo DIB
        Call pvSetDIBPicture(LoadPicture(sPath))
        '-- Refresh Panoramic view
        Call gfPanView.Repaint
    
        m_UndoPos = m_UndoPos + 1
        If (m_UndoPos > m_UndoMax) Then
            m_UndoMax = m_UndoPos
        End If
    End If
End Sub

Private Sub pvClearAllDIB()
    
    '-- Delete all temp. files
    On Error Resume Next
       Kill m_Temp & "\b" & m_AppID & "*.dat"
    On Error GoTo 0
    
    '-- Reset 'counters'
    m_UndoPos = 0
    m_UndoMax = 0
End Sub

Private Sub pvSaveUndoDIB()
    Dim lIdx  As Long
    Dim sPath As String
    
    '-- Get path
    sPath = m_Temp & "\b" & m_AppID & Format(m_UndoPos, "000") & ".dat"
    '-- Save DIB
    With gfrmMain
        Call .DIBSave.Save_BMP(sPath, .Canvas.DIB, .DIBPal, .DIBDither, .DIBbpp)
    End With
    '-- Saved flag
    m_Saved = False
    If (m_UndoMax - m_UndoPos > 0) Then
        On Error Resume Next
        For lIdx = m_UndoPos + 1 To m_UndoMax
            Kill m_Temp & "\b" & m_AppID & Format(lIdx, "000") & ".dat"
        Next lIdx
        On Error GoTo 0
    End If

    If (m_UndoPos < m_UNDO_LEVELS) Then
        m_UndoPos = m_UndoPos + 1
        m_UndoMax = m_UndoPos
      Else
        Call pvRotateUndoFiles
    End If
End Sub

Private Sub pvRotateUndoFiles()

  Dim bOldName As String
  Dim bNewName As String
  Dim lIdx     As Long

    On Error Resume Next
    '-- Kill first
    Kill m_Temp & "\b" & m_AppID & "000.dat"
    '-- 'Rotate' the others (Move up 1)
    For lIdx = 1 To m_UNDO_LEVELS
        bOldName = m_Temp & "\b" & m_AppID & Format(lIdx - 0, "000") & ".dat"
        bNewName = m_Temp & "\b" & m_AppID & Format(lIdx - 1, "000") & ".dat"
        Name bOldName As bNewName
    Next lIdx
    On Error GoTo 0
End Sub

Private Function pvExtToBMP(ByVal sFilename As String) As String
    pvExtToBMP = Left$(sFilename, Len(sFilename) - 3) & "bmp"
End Function

Private Function pvGetExt(ByVal sFilename As String) As String
    pvGetExt = Right$(sFilename, 3)
End Function

Private Function pvCorrectExt(sFilename As String)
    If (Right$(sFilename, 4) <> Right$(m_FileExt, 4)) Then
        sFilename = sFilename & Right$(m_FileExt, 4)
    End If
End Function

'========================================================================================
' ȫ������ (����)
'========================================================================================

Public Property Let LastFilterID(ByVal FilterID As fltIDCts)
    m_LastFilter = FilterID
End Property

Public Property Get LastFilename() As String
    LastFilename = m_LastFilename
End Property

Public Property Let LastFilename(ByVal sLastFilename As String)
    m_LastFilename = sLastFilename
End Property

Public Property Get LastPath() As String
    LastPath = m_LastPath
End Property

Public Property Let LastPath(ByVal sLastPath As String)
    m_LastPath = sLastPath
End Property

Public Property Get FileExt() As String
    FileExt = m_FileExt
End Property

Public Property Let FileExt(ByVal sFileExt As String)
    m_FileExt = sFileExt
End Property

Public Property Get DialogPreview() As Boolean
    DialogPreview = m_DialogPreview
End Property

Public Property Let DialogPreview(ByVal bShow As Boolean)
    m_DialogPreview = bShow
End Property

Public Property Get DialogFitMode() As Boolean
    DialogFitMode = m_DialogFitMode
End Property

Public Property Let DialogFitMode(ByVal bEnable As Boolean)
    m_DialogFitMode = bEnable
End Property

Public Property Get DialogJPEGquality() As Integer
    DialogJPEGquality = m_DialogJPEGquality
End Property

Public Property Let DialogJPEGquality(ByVal iValue As Integer)
    m_DialogJPEGquality = iValue
End Property

Private Sub BeginShowProgress()
    '��������λ
    On Error Resume Next
    With Progress
        .Left = stbThis.Panels(4).Left + Screen.TwipsPerPixelX * 2
        .Top = stbThis.Top + (stbThis.Height - .Height) / 2 + Screen.TwipsPerPixelY * 2
        .Width = stbThis.Panels(4).Width - Screen.TwipsPerPixelX * 4
        .Height = stbThis.Height - Screen.TwipsPerPixelY * 4
        .Visible = True: Me.Refresh
    End With
End Sub

