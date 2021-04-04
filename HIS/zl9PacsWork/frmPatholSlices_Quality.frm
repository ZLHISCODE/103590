VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmPatholSlices_Quality 
   Caption         =   "��Ƭ����"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   Icon            =   "frmPatholSlices_Quality.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   11280
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txtSlideNum 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2670
      TabIndex        =   3
      Top             =   6765
      Width           =   1260
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   6180
      Left            =   420
      ScaleHeight     =   6180
      ScaleWidth      =   9735
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   345
      Width           =   9735
      Begin zl9PACSWork.ucFlexGrid ufgData 
         Height          =   5145
         Left            =   375
         TabIndex        =   1
         Top             =   825
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   9075
         DefaultCols     =   ""
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         HeadFontSize    =   10.5
         HeadFontCharset =   134
         HeadFontWeight  =   400
         DataFontSize    =   10.5
         DataFontCharset =   134
         DataFontWeight  =   400
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   7515
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatholSlices_Quality.frx":038A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6165
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
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
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPatholSlices_Quality"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngAdviceID As Long       'ҽ��ID

Private mrecStudyInf As TStudyStateInf
Private mblnCurModifyState As Boolean
Private mblnAllowEditState As Boolean

Private Enum TMenuType
    mtFile = 1          '�ļ�
      mtSave = 2        '����
      mtCancel = 3      '����
      mtQuit = 4        '�˳�
    
    mtEdit = 5          '�༭
      mtModify = 6      '
      
    mtApplyAll = 7      'Ӧ�õ�����
    mtClear = 8         '���
    
    mtFind = 9          '����
    mtPlace = 10        'ռλ
End Enum


Public Sub ShowSlideEvaluateWindow(ByVal lngAdviceID As Long, ByVal lngStudyStep As Long, _
    ByVal strPrivs As String, owner As Object)

    '���Է���
'    InitDebugObject 1290, Me, "zlhis", "HIS"
    
        '����ҽ��ID��ȡ����ŵ����״̬��Ϣ
    Call GetPatholStudyState(lngAdviceID, mrecStudyInf)
    
    mblnAllowEditState = CheckPopedom(strPrivs, "��������") And lngStudyStep < 6
    
    Call Me.Show(1, owner)
End Sub

Private Sub InitQualityList()
'��ʼ����Ƭ�����б�
   
    ufgData.IsEjectConfig = False
    ufgData.IsShowPopupMenu = False
    ufgData.IsKeepRows = False
    ufgData.IsCopyMode = True
    
    ufgData.RowHeightMin = 315
    ufgData.ColNames = gstrSlicesQualityCols
    
    
    ufgData.ColConvertFormat = gstrSlicesQualityConvertFormat
    ufgData.DataGrid.ExtendLastCol = True
    
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
'ִ�н��湦��
Dim strResult As String

On Error GoTo errHandle
    Select Case control.ID
    
        Case TMenuType.mtCancel
            Call CancelModify       '�����޸�
            
'        Case TMenuType.mtModify
'            Call ModifyEvaluate     '�޸�����
            
        Case TMenuType.mtClear
            Call ClearEvaluate      '�������
            
        Case TMenuType.mtApplyAll
            Call ApplyAll           'Ӧ�õ�����
            
        Case TMenuType.mtSave
            Call SaveEvaluate       '������������
            
        Case TMenuType.mtFind
            Call FindData           '��������
                        
        
'---------------------------�鿴----------------
        Case conMenu_View_ToolBar_Button '������
            Call Menu_View_ToolBar_Button_click(control)
        Case conMenu_View_ToolBar_Text '��ť����
            Call Menu_View_ToolBar_Text_click(control)
        Case conMenu_View_ToolBar_Size '��ͼ��
            Call Menu_View_ToolBar_Size_click(control)
        Case conMenu_View_StatusBar '״̬��
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
            
        Case TMenuType.mtQuit   '�˳�
            Call Unload(Me)
    End Select
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub FindData()
'���Ҳ�Ƭ����
    Dim lngFindRowIndex As Long
    
    If Trim(txtSlideNum.Text) = "" Then Exit Sub
    
    lngFindRowIndex = ufgData.FindRowIndex(Trim(txtSlideNum.Text), "�����")
    
    If lngFindRowIndex >= 1 Then
        Call ufgData.LocateRow(lngFindRowIndex)
        
        If ufgData.Text(lngFindRowIndex, "��Ƭ����") = "" And mblnAllowEditState Then
            ufgData.DataGrid.TextMatrix(lngFindRowIndex, ufgData.GetColIndex("��Ƭ����")) = "��"
            
            mblnCurModifyState = True
        End If
        
'        Call ufgData.EditNextCellWithCurRow(False)
    End If
    
    Call zlControl.TxtSelAll(txtSlideNum)
End Sub


Private Sub CancelModify()
'�����޸�
    ufgData.DataGrid.Row = ufgData.DataGrid.Row
    
    Call ufgData.RestoreList(False)
    Call ufgData.RefreshReadColColor
    
    mblnCurModifyState = False
    'Call ConfigFaceEditState(False)
End Sub

'Private Sub ModifyEvaluate()
''�޸�����
'    Call ConfigFaceEditState(True)
'End Sub

Private Sub ClearEvaluate()
'�������
    Dim i As Long
    
    ufgData.DataGrid.Row = ufgData.DataGrid.Row
    For i = 1 To ufgData.GridRows - 1
        '���ﲻ��ʹ��ufgdata��text���Ը�ֵ����Ϊ�����Ի�����е�flexcpData���ԣ�ʹ��ȡ��ʱ���ָܻ�����
        ufgData.DataGrid.TextMatrix(i, ufgData.GetColIndex("��Ƭ����")) = ""
        ufgData.DataGrid.TextMatrix(i, ufgData.GetColIndex("������")) = ""
        ufgData.DataGrid.TextMatrix(i, ufgData.GetColIndex("��������")) = ""
    Next i
    
    mblnCurModifyState = True
End Sub

Private Sub SaveEvaluate()
'��������
    Dim i As Long
    Dim strSql As String
    Dim strQuality As String
    Dim dtCurDate As Date
    
    ufgData.DataGrid.Row = ufgData.DataGrid.Row
    dtCurDate = zlDatabase.Currentdate
    
    'ѭ������������������
    For i = 1 To ufgData.GridRows - 1
        strQuality = Trim(ufgData.Text(i, "��Ƭ����"))
        
        strSql = "Zl_����Ƭ��Ϣ_��������(" & CLng(Val(ufgData.KeyValue(i))) & _
                                            ",'" & strQuality & _
                                            "','" & UserInfo.���� & "')"
                                       
        zlDatabase.ExecuteProcedure strSql, "��Ƭ��������"

        
        ufgData.Text(i, "��Ƭ����") = strQuality '����flexcpdata���ݣ��Ա���г����ָ�
        ufgData.Text(i, "������") = IIf(strQuality = "", "", UserInfo.����)
        ufgData.Text(i, "��������") = IIf(strQuality = "", "", Format(dtCurDate, "yyyy-mm-dd"))
    Next i
    
    mblnCurModifyState = False
'    Call ConfigFaceEditState(False)
End Sub

Private Sub ApplyAll()
'Ӧ�õ�����
    Dim strCurValue As String
    Dim i As Long
    
    ufgData.DataGrid.Row = ufgData.DataGrid.Row
    strCurValue = ufgData.CurText("��Ƭ����")
    
    If strCurValue = "" Then
        Call MsgBoxD(Me, "��ǰ��¼δ���ò�Ƭ����������Ӧ�õ����������С�", vbOKOnly)
        Exit Sub
    End If
    
    For i = 1 To ufgData.GridRows - 1
        ufgData.DataGrid.TextMatrix(i, ufgData.GetColIndex("��Ƭ����")) = strCurValue
    Next i
    
    mblnCurModifyState = True
End Sub


Private Sub Menu_Help_Web_Mail_click()
On Error GoTo errHandle
    zlMailTo hWnd
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_About_click()
On Error GoTo errHandle
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Help_click()
'���ܣ����ð�������
On Error GoTo errHandle
    ShowHelp App.ProductName, Me.hWnd, Me.Name
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Web_Forum_click()
On Error GoTo errHandle
    Call zlWebForum(Me.hWnd)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Web_Home_click()
On Error GoTo errHandle
    zlHomePage hWnd
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_View_StatusBar_click(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Me.stbThis.Visible = Not Me.stbThis.Visible
    control.Checked = Not control.Checked
    picBack.Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    
    Me.cbrMain.RecalcLayout
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_View_ToolBar_Button_click(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Dim i As Integer
    
    For i = 2 To cbrMain.Count
        Me.cbrMain(i).Visible = Not Me.cbrMain(i).Visible
    Next

    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_View_ToolBar_Size_click(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Me.cbrMain.Options.LargeIcons = Not Me.cbrMain.Options.LargeIcons
    control.Checked = Not control.Checked
    
    Me.cbrMain.RecalcLayout
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_View_ToolBar_Text_click(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
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
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub cbrMain_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
On Error Resume Next
    picBack.Left = Left
    picBack.Top = Top
    picBack.Width = Right - Left
    picBack.Height = Bottom - Top - IIf(stbThis.Visible, stbThis.Height, 0)
End Sub


Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
'���²˵��Ͱ�ť��ʾ
On Error Resume Next
    
    Select Case control.ID

        Case TMenuType.mtSave, TMenuType.mtCancel ', TMenuType.mtClear, mtApplyAll
            control.Enabled = mblnCurModifyState And mblnAllowEditState


'        Case TMenuType.mtModify
'            control.Enabled = Not mblnCurModifyState And ufgData.GridRows > 1

        Case TMenuType.mtApplyAll, TMenuType.mtClear
            control.Enabled = ufgData.GridRows > 1 And mblnAllowEditState
    End Select
    
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    
    Call RestoreWinState(Me, App.ProductName)
    
    mblnCurModifyState = False
    
    Call InitCommandBars
    
    '��ʼ����Ƭ��ʾ�б�
    Call InitQualityList
    
    Call LoadSlideData(mrecStudyInf.lngPatholAdviceId)

    '�����ǰ�������޸ģ��򽫽����޸�Ϊֻ���鿴״̬
    If Not mblnAllowEditState Then
        Call ConfigFaceEditState(False)
    End If
    
    stbThis.Panels(3).Text = "�����ˣ�" & UserInfo.����
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub InitCommandBars()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    
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
        .VisualTheme = xtpThemeOffice2003                      '���ÿؼ���ʾ���
        .EnableCustomization False                             '�Ƿ������Զ�������
        Set .Icons = zlCommFun.GetPubIcons                     '���ù�����ͼ��ؼ�
    End With

    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '�˵�����
'Begin------------------------�༭�˵�--------------------------------------Ĭ�Ͽɼ�
    cbrMain.ActiveMenuBar.Title = "�˵�"
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, TMenuType.mtFile, "�ļ�(&F)")
    
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtSave, "����(&S)"): cbrControl.IconId = 3091
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtCancel, "����(&C)"): cbrControl.IconId = 3565
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtQuit, "�˳�(&Q)"): cbrControl.IconId = 2613: cbrControl.BeginGroup = True
    
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, TMenuType.mtEdit, "�༭(&E)")
    
    'Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtModify, "����(&M)"): cbrControl.IconId = 3003
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtApplyAll, "Ӧ�õ�����(&A)"): cbrControl.IconId = 3002: cbrControl.BeginGroup = True
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtClear, "�������(&C)"): cbrControl.IconId = 4008: cbrControl.BeginGroup = True
    
    'Begin----------------------�鿴�˵�--------------------------------------
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(V)")
    Call CreateViewAndHelpMenu(cbrMenuBar, Nothing)
    
    'Begin----------------------�����˵�--------------------------------------Ĭ�Ͽɼ�
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(H)")
    Call CreateViewAndHelpMenu(Nothing, cbrMenuBar)
    
    
    
    '---------------------����������------------------------------------------
        
            
    Set cbrToolBar = Me.cbrMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True
        
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtSave, "����", "��������"): cbrControl.IconId = 3091
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtCancel, "����", "�����޸�"): cbrControl.IconId = 3565
    
    'Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtModify, "����", "��Ƭ����"): cbrControl.IconId = 3003
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtClear, "���", "�������"): cbrControl.IconId = 4008
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtApplyAll, "Ӧ�õ�����", "����������������Ϊ��ͬ"): cbrControl.IconId = 3002: cbrControl.BeginGroup = True

    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtQuit, "�˳�", "�˳�"): cbrControl.IconId = 2613: cbrControl.BeginGroup = True
    cbrControl.BeginGroup = True
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '�������ϽǶ�λ����
    Set cbrControl = cbrMain.ActiveMenuBar.Controls.Add(xtpControlLabel, TMenuType.mtPlace, "����ţ�")
        cbrControl.ID = TMenuType.mtPlace
        cbrControl.flags = xtpFlagRightAlign
        cbrControl.IconId = 1
        
    Set cbrCustom = cbrMain.ActiveMenuBar.Controls.Add(xtpControlCustom, TMenuType.mtPlace, "�����")
        cbrCustom.Handle = txtSlideNum.hWnd
        cbrCustom.flags = xtpFlagRightAlign
        cbrCustom.Style = xtpButtonIconAndCaption
        
    Set cbrControl = cbrMain.ActiveMenuBar.Controls.Add(xtpControlButton, TMenuType.mtFind, " �� λ(&L) ")
        cbrControl.ID = TMenuType.mtFind
        cbrControl.flags = xtpFlagRightAlign
End Sub


Private Sub ConfigFaceEditState(ByVal blnIsEdit As Boolean)
    Dim lngColIndex As Long
    
    lngColIndex = ufgData.GetColIndex("��Ƭ����")
    mblnCurModifyState = blnIsEdit
    
    ufgData.ReadOnly = Not blnIsEdit
    
    If ufgData.GridRows <= 1 Then Exit Sub
    
    If blnIsEdit Then
        ufgData.DataGrid.Cell(flexcpBackColor, 1, ufgData.GetColIndex("��Ƭ����"), ufgData.DataGrid.Rows - 1, ufgData.GetColIndex("��Ƭ����")) = &H80000005
    Else
        ufgData.DataGrid.Cell(flexcpBackColor, 1, ufgData.GetColIndex("��Ƭ����"), ufgData.DataGrid.Rows - 1, ufgData.GetColIndex("��Ƭ����")) = &H8000000F
    End If

End Sub


Private Sub LoadSlideData(ByVal lngPatholAdviceId As Long)
On Error GoTo errHandle
    Dim i As Integer
    Dim strSql As String
    Dim rsData As ADODB.Recordset

    '��ѯ��Դ����Ƭ������ɲ�Ƭ��Ϣ
    strSql = "select a.ID,a.��Դ����, a.��ԴId, a.�Ŀ�ID,a.�����,a.��Ƭ����,a.������,a.��������, d.�걾����,c.ȡ��λ��, c.��� as �Ŀ��," & _
                     "decode(b.��Ƭ����,0,'����',1,'����',2,'ϸ��',3,'����',4,'����',5,'����','') as ��Ƭ���� " & _
                     "from  ����Ƭ��Ϣ a,������Ƭ��Ϣ b, ����ȡ����Ϣ c, ����걾��Ϣ d " & _
                     "where a.��Դid = b.id and b.�Ŀ�ID = c.�Ŀ�id and c.�걾Id=d.�걾Id and a.��Դ����=0 and b.��ǰ״̬=2 and a.����ҽ��ID =[1]"
    
    strSql = strSql & vbCrLf & " union all " & vbCrLf
    
    '��ѯ��Դ���ؼ������ɲ�Ƭ��Ϣ
    strSql = "select * from (" & strSql & " select a.ID,a.��Դ����, a.��ԴId, a.�Ŀ�ID,a.�����,a.��Ƭ����,a.������,a.��������, d.�걾����,c.ȡ��λ��,c.��� as �Ŀ��," & _
                     "decode(b.�ؼ�����,0,'����',1,'��Ⱦ',2,'����','') as ��Ƭ���� " & _
                     "from  ����Ƭ��Ϣ a,�����ؼ���Ϣ b, ����ȡ����Ϣ c, ����걾��Ϣ d " & _
                     "where a.��Դid = b.id and b.�Ŀ�ID = c.�Ŀ�id and c.�걾Id=d.�걾Id and a.��Դ����=1 and b.��ǰ״̬=2 and a.����ҽ��ID =[1] ) order by ����� "
    
                     
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ȡ��Ƭ��Ϣ", lngPatholAdviceId)
    
    Call ufgData.ClearListData
    If rsData.RecordCount < 1 Then Exit Sub
    
    Set ufgData.AdoData = rsData
    Call ufgData.RefreshData
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errHandle
    Call SaveWinState(Me, App.ProductName)
errHandle:
End Sub

Private Sub picBack_Resize()
On Error Resume Next
   
    ufgData.Left = 40
    ufgData.Top = 0
    ufgData.Width = picBack.ScaleWidth - 80
    ufgData.Height = picBack.ScaleHeight - 20
End Sub


Private Sub txtSlideNum_KeyPress(KeyAscii As Integer)
On Error GoTo errHandle
    If KeyAscii = 13 Then
        Call FindData
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnChangeEdit()
    mblnCurModifyState = True
End Sub

Private Sub ufgData_OnDblClick()
On Error GoTo errHandle
    Call ufgData.EditNextCellWithCurRow(False)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

