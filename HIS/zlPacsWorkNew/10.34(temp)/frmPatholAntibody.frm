VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPatholAntibody 
   Caption         =   "����ά��"
   ClientHeight    =   7755
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   10155
   Icon            =   "frmPatholAntibody.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   10155
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picFeedback 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2985
      ScaleWidth      =   9705
      TabIndex        =   5
      Top             =   4200
      Width           =   9735
      Begin zl9PACSWork.ucFlexGrid ufgFeedback 
         Height          =   2175
         Left            =   0
         TabIndex        =   6
         Top             =   720
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   3836
         IsKeepRows      =   0   'False
         BackColor       =   12648447
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         Editable        =   0
         HeadFontCharset =   134
         HeadFontWeight  =   400
         DataFontCharset =   134
         DataFontWeight  =   400
      End
      Begin VB.Label labSubTitle 
         Caption         =   "���巴����¼"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.Line linFlag 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         X1              =   0
         X2              =   9840
         Y1              =   360
         Y2              =   360
      End
   End
   Begin VB.PictureBox picData 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   120
      ScaleHeight     =   3585
      ScaleWidth      =   9705
      TabIndex        =   0
      Top             =   480
      Width           =   9735
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   1080
         TabIndex        =   1
         ToolTipText     =   "���ݿ������ƽ��п��ٶ�λ��"
         Top             =   120
         Width           =   1695
      End
      Begin TabDlg.SSTab tsFilter 
         Height          =   330
         Left            =   0
         TabIndex        =   2
         Top             =   600
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   582
         _Version        =   393216
         Tabs            =   4
         Tab             =   3
         TabsPerRow      =   4
         TabHeight       =   520
         TabMaxWidth     =   2822
         WordWrap        =   0   'False
         TabCaption(0)   =   "���п���(0)"
         TabPicture(0)   =   "frmPatholAntibody.frx":179A
         Tab(0).ControlEnabled=   0   'False
         Tab(0).ControlCount=   0
         TabCaption(1)   =   "���ڿ���(0)"
         TabPicture(1)   =   "frmPatholAntibody.frx":17B6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "��������(0)"
         TabPicture(2)   =   "frmPatholAntibody.frx":17D2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         TabCaption(3)   =   "���ÿ���(0)"
         TabPicture(3)   =   "frmPatholAntibody.frx":17EE
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).ControlCount=   0
      End
      Begin zl9PACSWork.ucFlexGrid ufgData 
         Height          =   2655
         Left            =   0
         TabIndex        =   3
         Top             =   960
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   4683
         IsKeepRows      =   0   'False
         BackColor       =   12648447
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         Editable        =   0
         HeadFontCharset =   134
         HeadFontWeight  =   400
         DataFontCharset =   134
         DataFontWeight  =   400
      End
      Begin VB.Label labFind 
         AutoSize        =   -1  'True
         Caption         =   "���ٲ��ң�"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "���ݿ������ƽ��п��ٶ�λ"
         Top             =   120
         Width           =   900
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   7395
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatholAntibody.frx":180A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11033
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   1200
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPatholAntibody.frx":209E
      Left            =   480
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPatholAntibody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngAntibodyLowCount As Long
Private mstrPrivs As String

Private mblnDataModifyState As Boolean
Private mblnFeedbackModifyState As Boolean

'�˵�����ö�ٶ���
Private Enum TMenuType
    mtAntibodyAdd = 1
    mtAntibodyMod = 2
    mtAntibodyDel = 3
    mtAntibodyStatus = 4
    mtFeedbackAdd = 5
    mtFeedbackMod = 6
    mtFeedbackDel = 7
End Enum

'�������ݵ���ʾ����
Private Enum TAntibodyDataShowType
    stAll = 0   '���п���
    stOverdue = 1 '���ڿ���
    stLow = 2 '��������
    stDisable = 3 '���ÿ���
End Enum

Public Sub ShowAntibodyManageWind(ByVal strPrivs As String, Optional owner As Form = Nothing)
'��ʾ���������
    mstrPrivs = strPrivs
    
    Call ConfigPopedom
    
    Call Me.Show(1, owner)
End Sub

Private Sub ConfigPopedom()
'����Ȩ��
    mblnDataModifyState = InStr(mstrPrivs, "�������") > 0
    mblnFeedbackModifyState = InStr(mstrPrivs, "���巴��") > 0
End Sub

Private Sub InitAntibodyList()
'��ʼ��������ʾ��
    Dim strTemp As String

     '�ж����ݿ�������Ƿ������� �����ȡ���ݿ����  û�������Ĭ��
    strTemp = zlDatabase.GetPara("������Ϣ�б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
     
    If strTemp = "" Then
        ufgData.ColNames = gstrAntibodyCols
    Else
        ufgData.ColNames = strTemp
    End If
    
     '��������
    ufgData.GridRows = glngStandardRowCount
    '�����и�
    ufgData.RowHeightMin = glngStandardRowHeight
    ufgData.DefaultColNames = gstrAntibodyCols
    ufgData.ColConvertFormat = gstrAntibodyConvertFormat
    ufgData.IsShowPopupMenu = False
End Sub

Private Sub InitFeedbackList()
'��ʼ�����巴����ʾ��
    Dim strTemp As String
    
     '�ж����ݿ�������Ƿ������� �����ȡ���ݿ����  û�������Ĭ��
    strTemp = zlDatabase.GetPara("���巴���б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
     
    If strTemp = "" Then
        ufgFeedback.ColNames = gstrAntibodyFeedbackCols
    Else
        ufgFeedback.ColNames = strTemp
    End If
    
     '��������
    'ufgFeedback.GridRows = glngStandardRowCount
    
    '��ֹ�Ҽ������б����ô���
    ufgFeedback.IsEjectConfig = False
    
    '�����и�
    ufgFeedback.RowHeightMin = glngStandardRowHeight
    
    ufgFeedback.DefaultColNames = gstrAntibodyFeedbackCols
    ufgFeedback.ColConvertFormat = gstrAntibodyFeedbackConvertFormat
    ufgFeedback.IsShowPopupMenu = False
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo ErrorHand
    
    Select Case control.ID
        Case TMenuType.mtAntibodyAdd                '��������
            Menu_Antibody_Add
            
        Case TMenuType.mtAntibodyMod                '�޸Ŀ���
            Menu_Antibody_Mod
            
        Case TMenuType.mtAntibodyDel                'ɾ������
            Menu_Antibody_Del
            
        Case TMenuType.mtAntibodyStatus
            If InStr(control.Caption, "���ÿ���") > 0 Then
                Menu_Antibody_Enable                '���ÿ���
                control.Caption = "���ÿ���"
            Else
                Menu_Antibody_Disable               '���ÿ���
                control.Caption = "���ÿ���"
            End If
        
        Case TMenuType.mtFeedbackAdd                '��������
            Menu_Feedback_Add
            
        Case TMenuType.mtFeedbackMod                '�޸ķ���
            Menu_Feedback_Mod
            
        Case TMenuType.mtFeedbackDel                'ɾ������
            Menu_Feedback_Del
            
        Case conMenu_File_Exit                      '�˳�
            Call Menu_File_Exit
        
        '---------------------------�鿴----------------
        Case conMenu_View_ToolBar_Button            '������
            Call Menu_View_ToolBar_Button_click(control)

        Case conMenu_View_ToolBar_Text              '��ť����
            Call Menu_View_ToolBar_Text_click(control)

        Case conMenu_View_StatusBar                 '״̬��
            Call Menu_View_StatusBar_click(control)
            
'--------------------------����-----------------
        Case conMenu_Help_Help
            Call Menu_Help_Help_click

        Case conMenu_Help_Web_Forum
            Call Menu_Help_Web_Forum_click

        Case conMenu_Help_Web_Home
            Call Menu_Help_Web_Home_click

        Case conMenu_Help_Web_Mail
            Call Menu_Help_Web_Mail_click

        Case conMenu_Help_About
            Call Menu_Help_About_click
    End Select
    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_File_Exit()
    Unload Me
End Sub

Private Sub Menu_Help_About_click()
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
End Sub

Private Sub Menu_Help_Web_Mail_click()
    zlMailTo hWnd
End Sub

Private Sub Menu_Help_Web_Home_click()
    zlHomePage hWnd
End Sub

Private Sub Menu_Help_Web_Forum_click()
    Call zlWebForum(Me.hWnd)
End Sub

Private Sub Menu_View_ToolBar_Button_click(ByVal control As XtremeCommandBars.ICommandBarControl)
Dim i As Integer
    For i = 2 To cbrMain.Count
        Me.cbrMain(i).Visible = Not Me.cbrMain(i).Visible
    Next

    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_View_ToolBar_Text_click(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrorHand
    Dim i As Integer, cbrControl As CommandBarControl
    Dim intStyle As Integer

    For i = 2 To cbrMain.Count
        If Me.cbrMain(i).Controls.Count >= 1 Then
            intStyle = Me.cbrMain(i).Controls(i).Style
            If intStyle = xtpButtonIconAndCaption Then
                intStyle = xtpButtonIcon
                Me.cbrMain(i).ShowTextBelowIcons = False
            Else
                intStyle = xtpButtonIconAndCaption
                Me.cbrMain(i).ShowTextBelowIcons = True
            End If
        End If
        
        For Each cbrControl In Me.cbrMain(i).Controls
            cbrControl.Style = intStyle
        Next
    Next
    
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
    
    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_View_StatusBar_click(ByVal control As XtremeCommandBars.ICommandBarControl)
    Me.stbThis.Visible = Not Me.stbThis.Visible
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_Help_Help_click()
    '���ܣ����ð�������
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cbrMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible = True Then Bottom = stbThis.Height
End Sub

Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim blnHasDataRecord As Boolean
    Dim blnHasFeedBackRecord As Boolean
    
    On Error GoTo ErrorHand
    
    '���û�м�¼����û��ѡ���У��˵��͹������򲻿���
    blnHasDataRecord = ufgData.IsSelectionRow
    blnHasFeedBackRecord = ufgFeedback.IsSelectionRow
    
    Select Case control.ID
        Case TMenuType.mtAntibodyAdd
            control.Enabled = mblnDataModifyState
            
        Case TMenuType.mtAntibodyMod
            control.Enabled = mblnDataModifyState And blnHasDataRecord
        
        Case TMenuType.mtAntibodyDel
            control.Enabled = mblnDataModifyState And blnHasDataRecord
            
        Case TMenuType.mtAntibodyStatus
            If control.Parent.type = xtpControlPopup Then
                control.Caption = IIf(ufgData.CurText("ʹ��״̬") = "ʹ����", "���ÿ���(&I)", "���ÿ���(&S)")
            Else
                control.Caption = IIf(ufgData.CurText("ʹ��״̬") = "ʹ����", "���ÿ���", "���ÿ���")
            End If
            
            control.IconId = IIf(ufgData.CurText("ʹ��״̬") = "ʹ����", 3006, 3009)
            
            control.Enabled = mblnDataModifyState And blnHasDataRecord
            control.Enabled = Not control.Enabled
            control.Enabled = Not control.Enabled
            
        Case TMenuType.mtFeedbackAdd
            control.Enabled = mblnDataModifyState And blnHasDataRecord
            
        Case TMenuType.mtFeedbackMod
            control.Enabled = mblnFeedbackModifyState And blnHasFeedBackRecord
        
        Case TMenuType.mtFeedbackDel
            control.Enabled = mblnFeedbackModifyState And blnHasFeedBackRecord
            
    End Select
    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub PicData_Resize()
    On Error Resume Next
    
    tsFilter.Left = 120
    tsFilter.Top = 120
    
    labFind.Left = tsFilter.Left + tsFilter.Width + 240
    labFind.Top = 160
    
    txtFind.Left = labFind.Left + labFind.Width
    txtFind.Top = labFind.Top - 40
    
    ufgData.Left = 120
    ufgData.Top = tsFilter.Top + tsFilter.Height
    ufgData.Height = picData.Height - ufgData.Top - 60
    ufgData.Width = picData.Width - 240
End Sub

Private Sub picFeedback_Resize()
    On Error Resume Next
    
    linFlag.X1 = 0
    linFlag.X2 = picFeedback.Width
    linFlag.Y1 = 200
    linFlag.Y2 = 200
    
    labSubTitle.Top = 110
    
    ufgFeedback.Left = 120
    ufgFeedback.Top = labSubTitle.Top + labSubTitle.Height + 40
    ufgFeedback.Height = picFeedback.Height - ufgFeedback.Top
    ufgFeedback.Width = picFeedback.Width - 240
End Sub

Private Sub ufgData_OnColFormartChange()
'���������б�����
    zlDatabase.SetPara "������Ϣ�б�����", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
End Sub

Private Sub ufgData_OnColsNameReSet()
On Error GoTo errHandle
   Call LoadAntibodyData(0)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'�����Ҽ��˵�
On Error GoTo errHandle
    If Button = 2 Then
        Dim objPopup As CommandBar
        Dim objControl As CommandBarControl

        Set cbrMain.Icons = zlCommFun.GetPubIcons
        Set objPopup = cbrMain.Add("�Ҽ��˵�", xtpBarPopup)
        With objPopup.Controls
            Set objControl = .Add(xtpControlButton, TMenuType.mtAntibodyAdd, "��������(&A)"): objControl.IconId = 4112
            Set objControl = .Add(xtpControlButton, TMenuType.mtAntibodyMod, "�޸Ŀ���(&U)"): objControl.IconId = 4113
            Set objControl = .Add(xtpControlButton, TMenuType.mtAntibodyDel, "ɾ������(&D)"): objControl.IconId = 4114
            
            Set objControl = .Add(xtpControlButton, TMenuType.mtAntibodyStatus, "���ÿ���(&S)"): objControl.IconId = 3009
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, TMenuType.mtFeedbackAdd, "��������(&N)"): objControl.IconId = 4010
            objControl.BeginGroup = True
        End With
        objPopup.ShowPopup
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnSelChange()
    ufgData_OnClick
End Sub

Private Sub ufgFeedback_OnClick()
    mblnDataModifyState = False
    mblnFeedbackModifyState = True
End Sub

Private Sub ufgFeedback_OnColFormartChange()
'���������б�����
    zlDatabase.SetPara "���巴���б�����", ufgFeedback.GetColsString(ufgFeedback), glngSys, G_LNG_PATHOLSYS_NUM
End Sub

Private Sub LoadFeedbackData(ByVal lngAntibodyId As Long)
'���뿹�巴������
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select ID,�ο������,ʵ������,�������,��������,����ҽ��,����ʱ�� from �����巴�� where ����ID=[1] order by id"
    Set ufgFeedback.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, ufgData.KeyValue(lngAntibodyId))
    
    Call ufgFeedback.RefreshData
End Sub

'Private Function GetTestTypeValue(ByVal lngTestType As Long) As String
''��ȡʵ�����͵�ȡֵ
'    GetTestTypeValue = ""
'
'    Select Case lngTestType
'        Case 0
'            GetTestTypeValue = "�����黯"
'        Case 1
'            GetTestTypeValue = "����Ⱦɫ"
'        Case 2
'            GetTestTypeValue = "���Ӳ���"
'        Case 3
'            GetTestTypeValue = "����"
'    End Select
'End Function

'
'Private Sub FillFeedbackDataToList(rsData As ADODB.Recordset)
''��ӿ��巴����¼
'
'    '���������ǰ����ʾ�����������˳�
'    If rsData.AbsolutePosition >= vfgData.Rows Then Exit Sub
'
'    Call mclsVFGFeedback.SetTextWithColName(rsData.AbsolutePosition, gstrAntibodyFeedback_ID, Nvl(rsData!ID))
'    Call mclsVFGFeedback.SetTextWithColName(rsData.AbsolutePosition, gstrAntibodyFeedback_�ο������, Nvl(rsData!�ο������))
'    Call mclsVFGFeedback.SetTextWithColName(rsData.AbsolutePosition, gstrAntibodyFeedback_ʵ������, GetTestTypeValue(Val(Nvl(rsData!ʵ������))))
'    Call mclsVFGFeedback.SetTextWithColName(rsData.AbsolutePosition, gstrAntibodyFeedback_��������, Nvl(rsData!��������))
'    Call mclsVFGFeedback.SetTextWithColName(rsData.AbsolutePosition, gstrAntibodyFeedback_�������, Nvl(rsData!�������))
'    Call mclsVFGFeedback.SetTextWithColName(rsData.AbsolutePosition, gstrAntibodyFeedback_����ҽ��, Nvl(rsData!����ҽ��))
'    Call mclsVFGFeedback.SetTextWithColName(rsData.AbsolutePosition, gstrAntibodyFeedback_����ʱ��, Format(Nvl(rsData!����ʱ��), gstrFullDateTimeFormat))
'
'End Sub

Private Sub LoadAntibodyData(iShowType As TAntibodyDataShowType)
'��ȡ��������
    Dim strSql As String
    Dim rsAntibody As ADODB.Recordset
    
    Select Case iShowType
        Case TAntibodyDataShowType.stAll:
            strSql = "select ����ID,��������,ʹ���˷�,�����˷�,��������,��Ч��,��������,��¡��,���ö���,������,Ӧ�����,�Ǽ���,�Ǽ�ʱ��,ʹ��״̬,��ע from ��������Ϣ order by ����ID"
        Case TAntibodyDataShowType.stOverdue:
            strSql = "select ����ID,��������,ʹ���˷�,�����˷�,��������,��Ч��,��������,��¡��,���ö���,������,Ӧ�����,�Ǽ���,�Ǽ�ʱ��,ʹ��״̬,��ע from ��������Ϣ where ��������<sysdate order by ����ID"
        Case TAntibodyDataShowType.stLow:
            strSql = "select ����ID,��������,ʹ���˷�,�����˷�,��������,��Ч��,��������,��¡��,���ö���,������,Ӧ�����,�Ǽ���,�Ǽ�ʱ��,ʹ��״̬,��ע from ��������Ϣ where ʹ���˷�-�����˷� < " & mlngAntibodyLowCount & " order by ����ID"
        Case TAntibodyDataShowType.stDisable:
            strSql = "select ����ID,��������,ʹ���˷�,�����˷�,��������,��Ч��,��������,��¡��,���ö���,������,Ӧ�����,�Ǽ���,�Ǽ�ʱ��,ʹ��״̬,��ע from ��������Ϣ where ʹ��״̬=0 order by ����ID"
    End Select
    
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    Call ufgData.RefreshData
End Sub

'Private Sub FillAntibodyDataToList(rsData As ADODB.Recordset)
'    '��ӿ����¼
'
'    '���������ǰ����ʾ�����������˳�
'    If rsData.AbsolutePosition >= vfgData.Rows Then Exit Sub
'
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_����ID, Nvl(rsData!����ID))
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_��������, Nvl(rsData!��������))
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_ʹ���˷�, Val(Nvl(rsData!ʹ���˷�)))
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_�����˷�, Val(Nvl(rsData!�����˷�)))
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_��������, Format(Nvl(rsData!��������), gstrDateFormat))
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_��Ч��, Val(Nvl(rsData!��Ч��)) & "��")
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_��������, Format(Nvl(rsData!��������), gstrDateFormat))
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_��¡��, IIf(Val(Nvl(rsData!��¡��)) = 1, "����¡", "���¡"))
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_���ö���, Nvl(rsData!���ö���))
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_������, Nvl(rsData!������))
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_Ӧ�����, Nvl(rsData!Ӧ�����))
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_ʹ��״̬, IIf(Val(Nvl(rsData!ʹ��״̬)) = 0, "�ѽ���", "ʹ����"))
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_�Ǽ���, Nvl(rsData!�Ǽ���))
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_�Ǽ�ʱ��, Format(Nvl(rsData!�Ǽ�ʱ��), gstrFullDateTimeFormat))
'    Call mclsVFGAntibody.SetTextWithColName(rsData.AbsolutePosition, gstrAntibody_��ע, Nvl(rsData!��ע))
'
'End Sub

Private Sub RefreshAntibodyCount()
'ˢ�¿���������ʾ
    Dim strSql As String
    Dim rsAntibodyCount As ADODB.Recordset
    
    strSql = "select " & _
             " (select count(1)  from ��������Ϣ) as ����, " & _
             " (select count(1)  from ��������Ϣ where ʹ��״̬=0) as ����, " & _
             " (select count(1)  from ��������Ϣ where (ʹ���˷�-�����˷�) < " & mlngAntibodyLowCount & ") as ����, " & _
             " (select count(1)  from ��������Ϣ where �������� < sysdate) as ���� " & _
             " from dual"
             
    Set rsAntibodyCount = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    If rsAntibodyCount.RecordCount <= 0 Then
        tsFilter.TabCaption(0) = "���п���(0)"
        tsFilter.TabCaption(1) = "���ڿ���(0)"
        tsFilter.TabCaption(2) = "��������(0)"
        tsFilter.TabCaption(3) = "���ÿ���(0)"
    Else
        tsFilter.TabCaption(0) = "���п���(" & Val(Nvl(rsAntibodyCount!����)) & ")"
        tsFilter.TabCaption(1) = "���ڿ���(" & Val(Nvl(rsAntibodyCount!����)) & ")"
        tsFilter.TabCaption(2) = "��������(" & Val(Nvl(rsAntibodyCount!����)) & ")"
        tsFilter.TabCaption(3) = "���ÿ���(" & Val(Nvl(rsAntibodyCount!����)) & ")"
    End If
End Sub

Private Sub Menu_Antibody_Add()
On Error GoTo errHandle
    Dim blnOk As Boolean
    
    blnOk = ShowUpdateWindow(True)
    
    If blnOk Then RefreshAntibodyCount
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function CheckAntibodyIsUsed(lngAntibodyId As Long) As Boolean
'��鿹���Ƿ��Ѿ�ʹ��
    Dim strSql As String
    Dim rsAntibody As ADODB.Recordset
    
    CheckAntibodyIsUsed = False
    
    strSql = "select 1 from �����ؼ���Ϣ where ����ID=[1]"
    Set rsAntibody = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAntibodyId)
    
    If rsAntibody.RecordCount > 0 Then CheckAntibodyIsUsed = True
End Function

Private Sub DelAntibodyData(lngAntibodyId As Long)
'ɾ���������ݼ�¼
    Dim strSql As String
    
    strSql = "zl_������_ɾ��(" & lngAntibodyId & ")"
    
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
End Sub

Private Function ShowUpdateFeedbackWindow(Optional ByVal isNew As Boolean = False) As Boolean
    Dim frmUpdate As New frmPatholAntibody_FeedbackUpdate
    On Error GoTo errFree
        If isNew Then
            ShowUpdateFeedbackWindow = frmUpdate.ShowAddAntibodyFeedback(Val(ufgData.KeyValue(ufgData.SelectionRow)), ufgFeedback, Me)
        Else
            ShowUpdateFeedbackWindow = frmUpdate.ShowUpdateAntibodyFeedback(ufgFeedback, Me)
        End If
errFree:
    Unload frmUpdate
    Set frmUpdate = Nothing
End Function

Private Sub Menu_Feedback_Add()
On Error GoTo errHandle
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ���з����Ŀ����¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call ShowUpdateFeedbackWindow(True)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Antibody_Del()
On Error GoTo errHandle
    '��Ҫ�жϵ�ǰ�����Ƿ��Ѿ�ʹ�ù����Ѿ�ʹ�õĿ��岻��ִ��ɾ��
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫɾ���Ŀ����¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Dim lngCurAntibodyId As Long

    lngCurAntibodyId = ufgData.KeyValue(ufgData.SelectionRow)
    
    If CheckAntibodyIsUsed(lngCurAntibodyId) Then
        Call MsgBoxD(Me, "�����ѱ�ʹ�ò���ɾ����", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If MsgBoxD(Me, "ȷ��Ҫɾ����ǰѡ��Ŀ����¼��", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    Call DelAntibodyData(lngCurAntibodyId)
    
    '��ն�Ӧ�Ŀ��巴����¼
    Call ufgFeedback.ClearListData
    
    Call ufgData.DelRow(ufgData.SelectionRow, False)
    
    Call RefreshAntibodyCount
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub EnableOrDisableAntibody(lngAntibodyId As Long, blnIsEnable As Boolean)
'���û����ÿ���
    Dim strSql As String
    
    strSql = "Zl_������_ʹ��״̬(" & lngAntibodyId & "," & IIf(blnIsEnable, 1, 0) & ")"
                                   
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
End Sub

Private Sub DelFeedbackData(lngFeedbackId As Long)
'ɾ���������ݼ�¼
    Dim strSql As String
    
    strSql = "Zl_�����巴��_ɾ��(" & lngFeedbackId & ")"
    
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
End Sub

Private Sub Menu_Feedback_Del()
'ɾ�����巴����¼
On Error GoTo errHandle

    If Not ufgFeedback.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫɾ���ķ�����¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Dim lngCurFeedbackId As Long
    
    lngCurFeedbackId = ufgFeedback.KeyValue(ufgFeedback.SelectionRow)
    
    If MsgBoxD(Me, "ȷ��Ҫɾ����ǰѡ��Ŀ����¼��", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    'ɾ�����巴������
    Call DelFeedbackData(lngCurFeedbackId)

    Call ufgFeedback.DelRow(ufgFeedback.SelectionRow, False)
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Antibody_Disable()
On Error GoTo errHandle
    
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ���õĿ����¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If

    If ufgData.Text(ufgData.SelectionRow, gstrAntibody_ʹ��״̬) = "�ѽ���" Then
        Call MsgBoxD(Me, "�����ѱ����á�", vbOKOnly, Me.Caption)
        Exit Sub
    End If

    Dim lngCurAntibodyId As Long

    lngCurAntibodyId = ufgData.KeyValue(ufgData.SelectionRow)
    Call EnableOrDisableAntibody(lngCurAntibodyId, False)

    '����������ʾ�б�
    ufgData.Text(ufgData.SelectionRow, gstrAntibody_ʹ��״̬) = "�ѽ���"

    Call RefreshAntibodyCount
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Antibody_Enable()
On Error GoTo errHandle

    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ���õĿ����¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If

    If ufgData.Text(ufgData.SelectionRow, gstrAntibody_ʹ��״̬) = "ʹ����" Then
        Call MsgBoxD(Me, "�����Ѵ���ʹ���С�", vbOKOnly, Me.Caption)
        Exit Sub
    End If

    Dim lngCurAntibodyId As Long

    lngCurAntibodyId = ufgData.KeyValue(ufgData.SelectionRow)
    Call EnableOrDisableAntibody(lngCurAntibodyId, True)

    '����������ʾ�б�
    ufgData.Text(ufgData.SelectionRow, gstrAntibody_ʹ��״̬) = "ʹ����"

    Call RefreshAntibodyCount
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function ShowUpdateWindow(Optional ByVal isNew As Boolean = False) As Boolean
    Dim frmUpdate As New frmPatholAntibody_AntiUpdate
    On Error GoTo errFree
        If isNew Then
            ShowUpdateWindow = frmUpdate.ShowAddAntibodyWindow(ufgData, Me)
        Else
            ShowUpdateWindow = frmUpdate.ShowUpdateAntibodyWindow(ufgData, Me)
        End If
errFree:
    Unload frmUpdate
    Set frmUpdate = Nothing
    
End Function

Private Sub Menu_Antibody_Mod()
'�������
On Error GoTo errHandle
    
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ���µĿ����¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If

    Dim blnOk As Boolean

    blnOk = ShowUpdateWindow(False)

    If blnOk Then RefreshAntibodyCount

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Feedback_Mod()
On Error GoTo errHandle

    If Not ufgFeedback.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ���µķ�����¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If

    Call ShowUpdateFeedbackWindow(False)

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Initialize()
    mlngAntibodyLowCount = 3
'    tsFilter.Tab = 0
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
'    InitDebugObject 1294, Me, "zlhis", "his"
    Call InitCommandBars
    
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitFace
    '�б��ʼ��
    Call InitAntibodyList
    Call InitFeedbackList
    
'    Call LoadAntibodyData(stAll)
    tsFilter.Tab = 0 '�Ը����Ը�ֵʱ���ᴥ�������¼�
    
    Call RefreshAntibodyCount
    
    '���ѡ���˵�һ�У����Զ�������������
    If ufgData.IsSelectionRow And Trim(ufgData.KeyValue(ufgData.SelectionRow)) <> "" Then
        Call LoadFeedbackData(Val(ufgData.SelectionRow))
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Exit Sub
End Sub

Private Sub InitFace()
'��ʼ�����ܽ���
    Dim Pane1 As Pane, Pane2 As Pane

    With dkpMain
        .CloseAll
        .SetCommandBars cbrMain
        .Options.HideClient = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With

    Set Pane1 = dkpMain.CreatePane(1, 0, Round(Me.Height * 3 / 5), DockTopOf)
    Pane1.Title = "�ײͼ�¼"
    Pane1.Handle = picData.hWnd
    Pane1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Pane1.MinTrackSize.Height = 100
    
    Set Pane2 = dkpMain.CreatePane(2, 0, Round(Me.Height * 2 / 5), DockBottomOf, Pane1)
    Pane2.Title = "������ϸ"
    Pane2.Handle = picFeedback.hWnd
    Pane2.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Pane2.MinTrackSize.Height = 100
End Sub

Private Sub InitCommandBars()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    
    '���ò˵����͹��������
    With cbrMain.Options
        .ShowExpandButtonAlways = False                         '�����ڹ������Ҳ���ʾѡ�ť,��ʹ�������㹻��
        .ToolBarAccelTips = True                                '��ʾ��ť��ʾ
        .AlwaysShowFullMenus = False                            '�����õĲ˵���������
        .UseFadedIcons = False                                  'ͼ����ʾΪ��ɫЧ��
        .IconsWithShadow = True                                 '���ָ�������ͼ����ʾ��ӰЧ��
        .UseDisabledIcons = True                                '��������ť����ʱͼ����ʾΪ������ʽ
        .LargeIcons = True                                      '��������ʾΪ��ͼ��
        .SetIconSize True, 24, 24                               '���ô�ͼ��ĳߴ�
        .SetIconSize False, 16, 16                              '����Сͼ��ĳߴ�
    End With
    With cbrMain
        .VisualTheme = xtpThemeOffice2003                       '���ÿؼ���ʾ���
        .EnableCustomization False                              '�Ƿ������Զ�������
        Set .Icons = zlCommFun.GetPubIcons                      '���ù�����ͼ��ؼ�
    End With

    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '�˵�����
'Begin------------------------�༭�˵�--------------------------------------Ĭ�Ͽɼ�
    cbrMain.ActiveMenuBar.Title = "�˵�"
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)")
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&Q)")
        cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)")
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAntibodyAdd, "��������(&A)"): cbrControl.IconId = 4112
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAntibodyMod, "�޸Ŀ���(&U)"): cbrControl.IconId = 4113
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAntibodyDel, "ɾ������(&D)"): cbrControl.IconId = 4114
        
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAntibodyStatus, "���ÿ���(&S)"): cbrControl.IconId = 3009
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtFeedbackAdd, "��������(&N)"): cbrControl.IconId = 4010
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtFeedbackMod, "�޸ķ���(&U)"): cbrControl.IconId = 3003
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtFeedbackDel, "ɾ������(&C)"): cbrControl.IconId = 4008
    End With
    
    'Begin----------------------�鿴�˵�--------------------------------------
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(V)")
    With cbrMenuBar.CommandBar
        Set cbrControl = .Controls.Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(T)")
        cbrControl.ID = conMenu_View_ToolBar
            With cbrControl.CommandBar '�����˵�
                Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(0)"): cbrPopControl.Checked = True
                Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(1)"): cbrPopControl.Checked = True
            End With
        Set cbrControl = .Controls.Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(S)"): cbrControl.Checked = True
    End With

    'Begin----------------------�����˵�--------------------------------------Ĭ�Ͽɼ�
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(H)")
    With cbrMenuBar.CommandBar
        Set cbrControl = .Controls.Add(xtpControlButton, conMenu_Help_Help, "��������(M)")
        Set cbrControl = .Controls.Add(xtpControlButtonPopup, conMenu_Help_Web, "WEB�ϵ�����(W)")
            With cbrControl.CommandBar
                Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Forum, "������̳(0)")
                Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Home, "������ҳ(1)")
                Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(2)")
            End With
        Set cbrControl = .Controls.Add(xtpControlButton, conMenu_Help_About, "���ڡ�(A)")
    End With
    '---------------------����������------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAntibodyAdd, "��������"): cbrControl.IconId = 4112
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAntibodyMod, "�޸Ŀ���"): cbrControl.IconId = 4113
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAntibodyDel, "ɾ������"): cbrControl.IconId = 4114
        
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAntibodyStatus, "���ÿ���"): cbrControl.IconId = 3009
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtFeedbackAdd, "��������"): cbrControl.IconId = 4010
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtFeedbackMod, "�޸ķ���"): cbrControl.IconId = 3003
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtFeedbackDel, "ɾ������"): cbrControl.IconId = 4008
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
        cbrControl.BeginGroup = True
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub tsFilter_Click(PreviousTab As Integer)
    Dim i As Integer
On Error GoTo errHandle
    If PreviousTab = tsFilter.Tab Then Exit Sub
    
    Select Case tsFilter.Tab
        Case 0: '���п���
            Call LoadAntibodyData(stAll)
        Case 1: '���ڿ���
            Call LoadAntibodyData(stOverdue)
        Case 2: '��������
            Call LoadAntibodyData(stLow)
        Case 3: '���ÿ���
            Call LoadAntibodyData(stDisable)
    End Select
    
    For i = 1 To ufgData.DataGrid.Rows - 1
        ufgData.Text(i, "��Ч��") = ufgData.Text(i, "��Ч��") & "��"
    Next
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub txtFind_Change()
On Error GoTo errHandle
    Dim lngFindIndex As Long
    
    If Trim(txtFind.Text) = "" Then Exit Sub
    
    lngFindIndex = ufgData.FindRowIndex(txtFind.Text, gstrAntibody_��������)
    
    If lngFindIndex > 0 Then Call ufgData.LocateRow(lngFindIndex)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub txtFind_GotFocus()
On Error Resume Next
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)
End Sub

Private Sub ufgData_OnClick()
On Error GoTo errHandle
    mblnDataModifyState = True
    mblnFeedbackModifyState = False
    
    ufgFeedback.ClearListData
    If ufgData.GridRows <= 1 Then Exit Sub
    If Not ufgData.IsSelectionRow Then Exit Sub
    
    Call LoadFeedbackData(ufgData.SelectionRow)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnDblClick()
On Error GoTo errHandle
    Dim blnOk As Boolean
    
    If ufgData.GridRows <= 1 Then Exit Sub
    If ufgData.MouseRowIndex <= 0 Then Exit Sub
        

    blnOk = ShowUpdateWindow(False)
    If blnOk Then RefreshAntibodyCount
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgFeedback_OnDblClick()
On Error GoTo errHandle
    
    If ufgFeedback.GridRows <= 1 Then Exit Sub
    If ufgFeedback.MouseRowIndex <= 0 Then Exit Sub
        
    Call ShowUpdateFeedbackWindow(False)
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgFeedback_OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'�����Ҽ��˵�
On Error GoTo errHandle
    If Button = 2 Then
        Dim objPopup As CommandBar
        Dim objControl As CommandBarControl

        Set cbrMain.Icons = zlCommFun.GetPubIcons
        Set objPopup = cbrMain.Add("�Ҽ��˵�", xtpBarPopup)
        With objPopup.Controls
            Set objControl = .Add(xtpControlButton, TMenuType.mtFeedbackMod, "�޸ķ���(&U)"): objControl.IconId = 3003
            Set objControl = .Add(xtpControlButton, TMenuType.mtFeedbackDel, "ɾ������(&C)"): objControl.IconId = 4008
        End With
        objPopup.ShowPopup
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
