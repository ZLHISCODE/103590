VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmPatholRequisition 
   Caption         =   "�������"
   ClientHeight    =   6825
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   10350
   Icon            =   "frmPatholRequisition.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   10350
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picRequest 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   240
      ScaleHeight     =   3255
      ScaleWidth      =   9615
      TabIndex        =   3
      Top             =   360
      Width           =   9615
      Begin zl9PACSWork.ucFlexGrid ufgRequest 
         Height          =   2895
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   5106
         GridRows        =   21
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
   End
   Begin VB.PictureBox picRequestContext 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   240
      ScaleHeight     =   3015
      ScaleWidth      =   9615
      TabIndex        =   1
      Top             =   3480
      Width           =   9615
      Begin zl9PACSWork.ucFlexGrid ufgContext 
         Height          =   2415
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   4260
         GridRows        =   21
         IsKeepRows      =   0   'False
         BackColor       =   12648447
         IsEnterNextCell =   0   'False
         IsBtnNextCell   =   0   'False
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         Editable        =   1
         HeadFontCharset =   134
         HeadFontWeight  =   400
         DataFontCharset =   134
         DataFontWeight  =   400
      End
      Begin VB.Label labSubTitle 
         AutoSize        =   -1  'True
         Caption         =   "������Ŀ"
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
         Height          =   180
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Width           =   780
      End
      Begin VB.Line linSubTitle 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         X1              =   0
         X2              =   9840
         Y1              =   240
         Y2              =   240
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6465
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatholRequisition.frx":179A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11377
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
      Left            =   600
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpRequest 
      Bindings        =   "frmPatholRequisition.frx":202E
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPatholRequisition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnSpecialExamState As Boolean
Private mblnSpecialExam As Boolean
Private mblnSlices As Boolean
Private mblnSlicesState As Boolean

Private mblnReqSpecialExam As Boolean
Private mblnReqSlices As Boolean
Private mblnReqGet As Boolean
Private mblnDelReq As Boolean

Private mlngCurAdviceId As Long
Private mstrPrivs As String
Private mblnMoved As Boolean
Private mlngRequestType As Long

Private mlngCurDepartmentId As Long

Private mrecStudyInf As TStudyStateInf

Private Enum TMenuType
    mtReqGet = 1         '��ȡ����
    mtReqSlices = 2      '��Ƭ����
    mtReqSpecialExam = 3 '�ؼ�����
    mtDelReq = 4         '��������
    
    mtAddSpeExamPro = 5  '���
    mtNewSE = 6          '����
    mtRedoSE = 7         '����
    mtDelSE = 8          'ɾ��
    
'    mtAddSlicesPro = 9   '����
'    mtDelSlicesPro = 10  'ɾ��
End Enum

Public blnIsUpdate As Boolean


Public Sub zlRefresh(lngAdviceID As Long, ByVal blnReadOnly As Boolean, strPrivs As String, ByVal blnMoved As Boolean, _
    ByVal lngCurDepartmentId As Long, Optional owner As Form = Nothing)
On Error GoTo errHandle

    If lngAdviceID <= 0 Then
        Call ConfigRequisitionFace(False, "ҽ��ID��Ч���顣")
        Exit Sub
    End If
    
'    If mlngCurAdviceId = lngAdviceId Then Exit Sub

    mlngCurAdviceId = lngAdviceID
    mstrPrivs = strPrivs
    mblnMoved = blnMoved
    mlngCurDepartmentId = lngCurDepartmentId
    blnIsUpdate = False
    
    Call GetPatholStudyState(lngAdviceID, mrecStudyInf)
    
   
    If mrecStudyInf.strPatholNumber = "" Then
        Call ConfigRequisitionFace(False, "�ü����δ������Ч�Ĳ���ţ���ȷ�ϸü���Ƿ��ѱ����ա�")
        
        If Not (owner Is Nothing) Then
            Call MsgBoxD(Me, "�ü����δ������Ч�Ĳ���ţ���ȷ�ϸü���Ƿ��ѱ����ա�", vbOKOnly, Me.Caption)
        End If
        
        Exit Sub
    Else
        '��ȡ������Ϣ
        Call LoadRequestInf(mrecStudyInf.lngPatholAdviceId)
        
        '����������ϸ
        Call ufgRequest_OnClick
        
        Call ConfigRequisitionFace(True)
    End If
    
    Call ConfigPopedom(blnReadOnly)
    
    If Not (owner Is Nothing) Then
        Call Me.Show(1, owner)
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ConfigPopedom(ByVal blnIsReadOnly As Boolean)
'����Ȩ��
    Dim blnSpeExamPopedom As Boolean
    Dim blnSlicesPopedom As Boolean
    Dim blnMaterialPopedom As Boolean
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    blnSpeExamPopedom = CheckPopedom(mstrPrivs, "�ؼ�����")
    blnSlicesPopedom = CheckPopedom(mstrPrivs, "��Ƭ����")
    blnMaterialPopedom = CheckPopedom(mstrPrivs, "��ȡ����")
    
    mblnReqSpecialExam = blnSpeExamPopedom And Not blnIsReadOnly
    mblnReqSlices = blnSlicesPopedom And Not blnIsReadOnly
    mblnReqGet = blnMaterialPopedom And Not blnIsReadOnly
    mblnDelReq = (blnSpeExamPopedom Or blnSlicesPopedom Or blnMaterialPopedom) And Not blnIsReadOnly
    
    mblnSpecialExam = blnSpeExamPopedom And Not blnIsReadOnly
    
    mblnSlices = blnSlicesPopedom And Not blnIsReadOnly
    
    ufgRequest.ReadOnly = blnIsReadOnly
    ufgContext.ReadOnly = blnIsReadOnly
    
    '�õ���Ƭ״̬�Ӷ��������� ���ؼ����롱�͡���Ƭ���롱 ������ť
    strSql = "select distinct ��ǰ״̬ from ��������Ϣ a,������Ƭ��Ϣ b where a.����ҽ��id = b.����ҽ��id and a.ҽ��id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "�õ���Ƭ״̬", mlngCurAdviceId)
    
    If rsTemp.RecordCount < 1 Then
        mblnReqSlices = False
        mblnReqSpecialExam = False
        Exit Sub
    End If
    
    mblnReqSlices = IIf(Nvl(rsTemp!��ǰ״̬, 0) = 2, True, False)
    mblnReqSpecialExam = IIf(Nvl(rsTemp!��ǰ״̬, 0) = 2, True, False)

End Sub



Private Sub ConfigRequisitionFace(ByVal blnIsValid As Boolean, Optional ByVal strHintInf As String = "")
'�����������
    mblnReqSpecialExam = blnIsValid
    mblnReqSlices = blnIsValid
    mblnReqGet = blnIsValid
    mblnDelReq = blnIsValid

    mblnSpecialExam = blnIsValid

    mblnSlices = blnIsValid
    
    If blnIsValid Then
        Call ufgRequest.CloseHintInf
        Call ufgContext.CloseHintInf
    Else
        Call ufgRequest.ShowHintInf(strHintInf)
        Call ufgContext.ShowHintInf(strHintInf)
    End If
End Sub


Private Sub InitFace()
'��ʼ�����沼��
    Dim Pane1 As Pane, Pane2 As Pane

    With dkpRequest
        .CloseAll
        .SetCommandBars cbrMain
        .Options.HideClient = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With

    Set Pane1 = dkpRequest.CreatePane(1, 0, Round(Me.Height * 3 / 5), DockTopOf, Nothing)
    Pane1.Title = "�����¼"
    Pane1.Handle = picRequest.hWnd
    Pane1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Pane1.MinTrackSize.Width = 50
    Pane1.MinTrackSize.Height = 100

    Set Pane2 = dkpRequest.CreatePane(2, 0, Round(Me.Height * 2 / 5), DockBottomOf, Pane1)
    Pane2.Title = "������ϸ"
    Pane2.Handle = picRequestContext.hWnd
    Pane2.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Pane2.MinTrackSize.Width = 50
    Pane2.MinTrackSize.Height = 100
End Sub


Private Sub InitRequisitionList()
'��ʼ�������б�
    Dim strTemp As String
    
    '�ж����ݿ�������Ƿ������� �����ȡ���ݿ����  û�������Ĭ��
    strTemp = zlDatabase.GetPara("��������б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
     
    If strTemp = "" Then
        ufgRequest.ColNames = gstrRequisitionCols
    Else
        ufgRequest.ColNames = strTemp
    End If
    
    '��������
    ufgRequest.GridRows = glngStandardRowCount
    '�����и�
    ufgRequest.RowHeightMin = glngStandardRowHeight
    
    ufgRequest.DefaultColNames = gstrRequisitionCols
    ufgRequest.ColConvertFormat = gstrRequisitionConvertFormat
    ufgRequest.IsShowPopupMenu = False
End Sub

Private Sub InitRequestContextList(ByVal lngRequestType As Long)
'��ʼ��������Ŀ��ϸ�б�
    Dim strTemp As String
    
    mlngRequestType = lngRequestType
    
    Select Case lngRequestType
        Case 0, 1, 2
        
            '�ж����ݿ�������Ƿ������� �����ȡ���ݿ����  û�������Ĭ��
            strTemp = zlDatabase.GetPara("�ؼ������б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
             
            If strTemp = "" Then
                ufgContext.ColNames = gstrRequest_SpeExam_Cols
            Else
                ufgContext.ColNames = strTemp
            End If
                   '��ֹ�Ҽ������б����ô���
            ufgContext.IsEjectConfig = False
            ufgContext.DefaultColNames = gstrRequest_SpeExam_Cols
            ufgContext.ColConvertFormat = gstrRequest_SpeExamConvertFormat
            
        Case 3
            
            '�ж����ݿ�������Ƿ������� �����ȡ���ݿ����  û�������Ĭ��
            strTemp = zlDatabase.GetPara("��Ƭ�����б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
             
            If strTemp = "" Then
                ufgContext.ColNames = gstrRequest_Slices_Cols
            Else
                ufgContext.ColNames = strTemp
            End If
                   '��ֹ�Ҽ������б����ô���
            ufgContext.IsEjectConfig = False
            ufgContext.DefaultColNames = gstrRequest_Slices_Cols
            ufgContext.ColConvertFormat = gstrRequest_SlicesConvertFormat
        Case 4, 5
            
            '�ж����ݿ�������Ƿ������� �����ȡ���ݿ����  û�������Ĭ��
            strTemp = zlDatabase.GetPara("��ȡ�����б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
             
            If strTemp = "" Then
                ufgContext.ColNames = gstrRequest_Material_Cols
            Else
                ufgContext.ColNames = strTemp
            End If
                   '��ֹ�Ҽ������б����ô���
            ufgContext.IsEjectConfig = False
            ufgContext.DefaultColNames = gstrRequest_Material_Cols
            '��������
            ufgContext.GridRows = glngStandardRowCount
            '�����и�
            ufgContext.RowHeightMin = glngStandardRowHeight
        
            ufgContext.ColConvertFormat = gstrRequest_MaterialConvertFormat
    End Select
    
    ufgContext.IsShowPopupMenu = False
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim blnHasRequestRecord As Boolean
    Dim blnHasContextRecord As Boolean
    
On Error GoTo ErrorHand
    blnHasRequestRecord = ufgRequest.IsSelectionRow
    blnHasContextRecord = ufgContext.IsSelectionRow
    
    Select Case control.ID
        Case TMenuType.mtReqGet                       '��ȡ����
            Call Menu_Edit_ReqGet
        
        Case TMenuType.mtReqSlices                    '��Ƭ����
            Call Menu_Edit_ReqSlices
        
        Case TMenuType.mtReqSpecialExam               '�ؼ�����
            Call Menu_Edit_ReqSpecialExam
        
        Case TMenuType.mtDelReq                       '��������
            Call Menu_Edit_DelReq
            
        Case TMenuType.mtAddSpeExamPro                '���
            If mblnSpecialExamState And mblnSpecialExam Then
                Call Menu_Edit_AddSpeExamPro
            ElseIf mblnSlicesState And mblnSlices Then
                Call Menu_Edit_AddSlicesPro
            End If
        
        Case TMenuType.mtNewSE                        '����
            Call Menu_Edit_NewSE
        
        Case TMenuType.mtRedoSE                       '����
            Call Menu_Edit_RedoSE
        
        Case TMenuType.mtDelSE                        'ɾ��
            If mblnSpecialExamState And mblnSpecialExam And blnHasContextRecord Then
                Call Menu_Edit_DelSE
            ElseIf mblnSlicesState And mblnSlices And blnHasContextRecord Then
                Call Menu_Edit_DelSlicesPro
            End If
            
'        Case TMenuType.mtAddSlicesPro                 '����
'            Call Menu_Edit_AddSlicesPro
'
'        Case TMenuType.mtDelSlicesPro                 'ɾ��
'            Call Menu_Edit_DelSlicesPro

        Case conMenu_File_Exit                        '�˳�
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
    Dim blnHasRequestRecord As Boolean
    Dim blnHasContextRecord As Boolean
    
On Error GoTo ErrorHand
    blnHasRequestRecord = ufgRequest.IsSelectionRow
    blnHasContextRecord = ufgContext.IsSelectionRow
    
    Select Case control.ID
        Case TMenuType.mtReqGet                         '��ȡ����
            control.Enabled = mblnReqGet
        
        Case TMenuType.mtReqSlices                      '��Ƭ����
            control.Enabled = mblnReqSlices
        
        Case TMenuType.mtReqSpecialExam                 '�ؼ�����
            control.Enabled = mblnReqSpecialExam
        
        Case TMenuType.mtDelReq                         '��������
            control.Enabled = mblnDelReq
            
        Case TMenuType.mtAddSpeExamPro                  '���
            control.Enabled = (mblnSpecialExamState And mblnSpecialExam) Or (mblnSlicesState And mblnSlices)
        
        Case TMenuType.mtNewSE                          '����
            control.Enabled = mblnSpecialExamState And mblnSpecialExam
        
        Case TMenuType.mtRedoSE                         '����
            control.Enabled = mblnSpecialExamState And mblnSpecialExam And blnHasContextRecord
        
        Case TMenuType.mtDelSE                          'ɾ��
            control.Enabled = ((mblnSpecialExamState And mblnSpecialExam) Or (mblnSlicesState And mblnSlices)) And blnHasContextRecord
            
'        Case TMenuType.mtAddSlicesPro                   '���
'            control.Enabled = mblnSlicesState And mblnSlices
'
'        Case TMenuType.mtDelSlicesPro                   'ɾ��
'            control.Enabled = mblnSlicesState And mblnSlices And blnHasContextRecord

    End Select
    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgContext_OnColFormartChange()
    '����ı�ʱ�����б�����
    
    Select Case mlngRequestType
        Case 0, 1, 2
        
            zlDatabase.SetPara "�ؼ������б�����", ufgContext.GetColsString(ufgContext), glngSys, G_LNG_PATHOLSYS_NUM
            
        Case 3
        
            zlDatabase.SetPara "��Ƭ�����б�����", ufgContext.GetColsString(ufgContext), glngSys, G_LNG_PATHOLSYS_NUM
           
        Case 4, 5
            
            zlDatabase.SetPara "��ȡ�����б�����", ufgContext.GetColsString(ufgContext), glngSys, G_LNG_PATHOLSYS_NUM
            
    End Select

End Sub

Private Sub ufgRequest_OnColFormartChange()
'����ı�ʱ�����б�����
     zlDatabase.SetPara "��������б�����", ufgRequest.GetColsString(ufgRequest), glngSys, G_LNG_PATHOLSYS_NUM
     
End Sub


Public Sub ChangeControlFace(ByVal lngRequestType As Long)
'�ı���ƽ���
    mblnSpecialExamState = IIf("0,1,2" Like "*" & lngRequestType & "*", True, False)
    
    mblnSlicesState = IIf("3" Like "*" & lngRequestType & "*", True, False)
End Sub


Private Sub ShowSpecialExamRequestWindow()
'��ʾ�ؼ����봰��
    Dim frmSpeExamRequest As New frmPatholRequisition_SpeExam
    On Error GoTo errFree
      
        
        '��ʾ�ؼ����봰��
        Call frmSpeExamRequest.ShowSpeExamRequestWindow(ufgRequest, ufgContext, mrecStudyInf.lngPatholAdviceId, -1, Me)
        
        blnIsUpdate = frmSpeExamRequest.blnIsOk
errFree:
    Call Unload(frmSpeExamRequest)
    Set frmSpeExamRequest = Nothing
    
End Sub


Private Sub AddSpeExamProject(ByVal lngCurRequestId As Long, Optional ByVal blnIsBuZuo As Boolean = False)
'����ؼ���Ŀ
    Dim frmSpeExamRequest As New frmPatholRequisition_SpeExam
    On Error GoTo errFree
    
        '��ʾ�ؼ����봰��
        Call frmSpeExamRequest.ShowSpeExamRequestWindow(ufgRequest, ufgContext, mrecStudyInf.lngPatholAdviceId, lngCurRequestId, Me, blnIsBuZuo)
        
        blnIsUpdate = frmSpeExamRequest.blnIsOk
errFree:
    Call Unload(frmSpeExamRequest)
    Set frmSpeExamRequest = Nothing
End Sub


Private Sub Menu_Edit_AddSlicesPro()
'������Ƭ������Ŀ
On Error GoTo errHandle
    If Not ufgRequest.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ�������������¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgRequest.IsEmptyKey(ufgRequest.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ч�������¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '��ʾ��Ƭ��������
    Call ShowSlicesRequestWindow(ufgRequest.KeyValue(ufgRequest.SelectionRow))
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Edit_AddSpeExamPro()
'����ؼ���Ŀ
On Error GoTo errHandle
    If Not ufgRequest.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ�������������¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgRequest.IsEmptyKey(ufgRequest.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ч�������¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '�����ؼ���Ŀ
    Call AddSpeExamProject(ufgRequest.KeyValue(ufgRequest.SelectionRow), False)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub DelSpeExamProject(ByVal lngSpeExamRow As Long)
'ɾ���ؼ���Ŀ
    Dim strSql As String
    
    strSql = "Zl_��������_�ؼ���Ŀ_ɾ��(" & ufgContext.KeyValue(lngSpeExamRow) & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Call ufgContext.RemoveRow(lngSpeExamRow)
End Sub


Private Function CheckAllowDelSpeExam(ByVal lngSpeExamRow As Long) As Boolean
'�ж��Ƿ�����ɾ���ؼ���Ŀ
    CheckAllowDelSpeExam = IIf(ufgContext.Text(lngSpeExamRow, gstrRequest_SpeExam_��ǰ״̬) = "������", True, False)
    
    If Not CheckAllowDelSpeExam Then
        Call MsgBoxD(Me, "���ؼ���Ŀ�ѱ����ܻ���ɣ�����ִ��ɾ��������", vbOKOnly, Me.Caption)
    End If
    
End Function


Private Sub CancelRequest(ByVal lngRequestRow As Long)
'��������
    Dim strSql As String
    
    strSql = "Zl_��������_ɾ��(" & Val(ufgRequest.KeyValue(lngRequestRow)) & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Call ufgRequest.RemoveRow(lngRequestRow)
End Sub



Private Function CheckAllowDelRequest(ByVal lngRequestRow As Long)
'��������Ƿ�����ɾ��
    CheckAllowDelRequest = IIf(ufgRequest.Text(lngRequestRow, gstrRequisition_��ǰ״̬) = "������", True, False)
    
    If Not CheckAllowDelRequest Then
        Call MsgBoxD(Me, "��������Ŀ�ѱ����ܻ�ִ�У�����ɾ����", vbOKOnly, Me.Caption)
    End If
End Function


Private Sub Menu_Edit_DelReq()
'ɾ������
On Error GoTo errHandle
    If Not ufgRequest.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫɾ���������¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgRequest.IsEmptyKey(ufgRequest.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ч�������¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If Not CheckAllowDelRequest(ufgRequest.SelectionRow) Then Exit Sub
    
    If MsgBoxD(Me, "ȷ��Ҫɾ����������Ŀ��", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    'ɾ���ؼ���Ŀ
    Call CancelRequest(ufgRequest.SelectionRow)
    
    If ufgRequest.ShowingDataRowCount > 0 Then
        Call ufgRequest.DataGrid.Select(ufgRequest.GridRows - 1, 0)
    End If
    
    Call ufgRequest_OnClick
    
    
    blnIsUpdate = True
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Edit_DelSE()
'ɾ���ؼ���Ŀ
On Error GoTo errHandle
    If Not ufgContext.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫɾ�����ؼ���Ŀ��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgContext.IsEmptyKey(ufgContext.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ч���ؼ���Ŀ��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If Not CheckAllowDelSpeExam(ufgContext.SelectionRow) Then Exit Sub
    
    If MsgBoxD(Me, "ȷ��Ҫɾ�����ؼ���Ŀ��", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    'ɾ���ؼ���Ŀ
    Call DelSpeExamProject(ufgContext.SelectionRow)
    
    blnIsUpdate = True
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Function CheckAllowDelSlices(ByVal lngSlicesRow As Long) As Boolean
'����Ƿ�����ɾ����Ƭ��Ŀ
    CheckAllowDelSlices = IIf(ufgContext.Text(lngSlicesRow, gstrRequest_Slices_��ǰ״̬) = "������", True, False)
    
    If Not CheckAllowDelSlices Then
        Call MsgBoxD(Me, "����Ƭ��Ŀ�ѱ����ܻ���ɣ�����ִ��ɾ��������", vbOKOnly, Me.Caption)
    End If
End Function


Private Sub DelSlicesProject(ByVal lngSlicesRow As Long)
'ɾ����Ƭ��Ŀ
    Dim strSql As String
    
    strSql = "Zl_��������_��Ƭ��Ŀ_ɾ��(" & ufgContext.KeyValue(lngSlicesRow) & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Call ufgContext.RemoveRow(lngSlicesRow)
End Sub


Private Sub Menu_Edit_DelSlicesPro()
'ɾ����Ƭ��Ŀ
On Error GoTo errHandle
    If Not ufgContext.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫɾ������Ƭ��Ŀ��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgContext.IsEmptyKey(ufgContext.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ч����Ƭ��Ŀ��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If Not CheckAllowDelSpeExam(ufgContext.SelectionRow) Then Exit Sub
    
    If MsgBoxD(Me, "ȷ��Ҫɾ������Ƭ��Ŀ��", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    'ɾ���ؼ���Ŀ
    Call DelSlicesProject(ufgContext.SelectionRow)
    
    blnIsUpdate = True
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_Edit_NewSE()
'������������µ��ؼ���Ŀ��ͬ��ֻ������������Ϊ����
On Error GoTo errHandle
    If Not ufgRequest.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ�������������¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgRequest.IsEmptyKey(ufgRequest.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ч�������¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '�����ؼ���Ŀ
    Call AddSpeExamProject(ufgRequest.KeyValue(ufgRequest.SelectionRow), True)
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub





Private Function GetRedoCount(ByVal strMaterialId As String, ByVal strAntibodyName As String)
'��ȡ�ؼ���Ŀ����������
    Dim i As Long
    Dim lngCount As Long
    
    lngCount = 0
    For i = 1 To ufgContext.GridRows - 1
        If Not ufgContext.IsEmptyKey(i) Then
            If Val(ufgContext.Text(i, gstrRequest_SpeExam_�Ŀ��)) = Val(strMaterialId) And _
                UCase(ufgContext.Text(i, gstrRequest_SpeExam_��������)) = UCase(strAntibodyName) Then
                If lngCount < GetNumber(ufgContext.Text(i, gstrRequest_SpeExam_��������)) Then lngCount = GetNumber(ufgContext.Text(i, gstrRequest_SpeExam_��������))
            End If
        End If
    Next i
    
    GetRedoCount = lngCount

End Function


Private Sub RedoSpeExamProject(ByVal lngSpeExamRow As Long)
'��Ŀ����
    Dim lngNewRow As Long
    Dim lngRedoCount As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select Zl_��������_�ؼ���Ŀ_����([1]) as ����ֵ from dual"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(ufgContext.KeyValue(lngSpeExamRow)))
    
    If rsData.RecordCount <= 0 Then
        Call err.Raise(0, "RedoSpeExamProject", "δ�ɹ���ȡ��������ؼ���ĿID,����ʧ�ܡ�")
        Exit Sub
    End If
    
    '�����ؼ��¼���б�
    lngNewRow = ufgContext.NewRow
    
    lngRedoCount = GetRedoCount(ufgContext.Text(lngSpeExamRow, gstrRequest_SpeExam_�Ŀ��), _
                                ufgContext.Text(lngSpeExamRow, gstrRequest_SpeExam_��������))
    
    '����������
    Call ufgContext.CopyRowData(lngSpeExamRow, lngNewRow)
    
    ufgContext.Text(lngNewRow, gstrRequest_SpeExam_ID) = rsData!����ֵ
    ufgContext.Text(lngNewRow, gstrRequest_SpeExam_��������) = "��" & lngRedoCount + 1 & "������"
    ufgContext.Text(lngNewRow, gstrRequest_SpeExam_��ǰ״̬) = "������"
    ufgContext.Text(lngNewRow, gstrRequest_SpeExam_������) = ""
    ufgContext.Text(lngNewRow, gstrRequest_SpeExam_���ʱ��) = ""
    ufgContext.Text(lngNewRow, gstrRequest_SpeExam_��Ŀ���) = ""
    
    Call ufgContext.LocateRow(lngNewRow)
End Sub



Private Function GetAntibodyUseCount(ByVal lngAntibodyId As String) As Long
'��ȡ����Ŀ��÷���
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    GetAntibodyUseCount = 0
    
    strSql = "select ʹ���˷�-�����˷� as �����˷� from ��������Ϣ where ����ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAntibodyId)
    
    If rsData.RecordCount > 0 Then GetAntibodyUseCount = Val(Nvl(rsData!�����˷�))
End Function

Private Sub Menu_Edit_RedoSE()
'�ؼ���Ŀ����
On Error GoTo errHandle
    If Not ufgContext.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ�������ؼ���Ŀ��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgContext.IsEmptyKey(ufgContext.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ч���ؼ���Ŀ��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgContext.Text(ufgContext.SelectionRow, gstrRequest_SpeExam_��ǰ״̬) <> "�����" Then
        Call MsgBoxD(Me, "����Ŀ��δ��ɣ����ܽ���������", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If GetAntibodyUseCount(Val(ufgContext.Text(ufgContext.SelectionRow, gstrRequest_SpeExam_����ID))) <= 0 Then
        If MsgBoxD(Me, "���� [" & ufgContext.Text(ufgContext.SelectionRow, gstrRequest_SpeExam_��������) & "] ���޿����˷ݣ��Ƿ������Ӹ���Ŀ��", vbYesNo, Me.Caption) <> vbYes Then
            Exit Sub
        End If
    End If
    
    '�ؼ���Ŀ����
    Call RedoSpeExamProject(ufgContext.SelectionRow)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then
    Resume
    End If
End Sub


Private Sub ShowSlicesRequestWindow(Optional ByVal lngCurRequestId As Long = -1)
'��ʾ��Ƭ���봰��
Dim frmSlicesRequest As New frmPatholRequisition_Slices
On Error GoTo errFree
    
    Call frmSlicesRequest.ShowSlicesRequestWindow(ufgRequest, ufgContext, mrecStudyInf.lngPatholAdviceId, lngCurRequestId, mlngRequestType, Me)
    
    blnIsUpdate = frmSlicesRequest.blnIsOk
errFree:
    Call Unload(frmSlicesRequest)
    Set frmSlicesRequest = Nothing
End Sub


Private Sub ShowSupMaterialRequestWindow()
'��ʾ��ȡ���봰��
    Dim frmSupMateriasRequest As New frmPatholRequisition_SupMaterial
    On Error GoTo errFree
        
        Call frmSupMateriasRequest.ShowSupMaterialWindow(ufgRequest, ufgContext, mrecStudyInf.lngPatholAdviceId, Me)
        
        blnIsUpdate = frmSupMateriasRequest.blnIsOk
errFree:
    Call Unload(frmSupMateriasRequest)
    Set frmSupMateriasRequest = Nothing
    
End Sub



Private Sub Menu_Edit_ReqGet()
'��ȡ������
On Error GoTo errHandle
    If Not CheckAllowNewRequest(TRequestType.rtMaterial) Then Exit Sub
     
    Call ShowSupMaterialRequestWindow
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Edit_ReqSlices()
On Error GoTo errHandle
    '��ʾ��Ƭ����
    If Not CheckAllowNewRequest(TRequestType.rtSlices) Then Exit Sub

    Call ShowSlicesRequestWindow
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Function CheckAllowNewRequest(ByVal rtRequestType As TRequestType) As Boolean
'�ж��Ƿ���������µ�����
    Dim strRequestType As String
    Dim i As Integer
    
    CheckAllowNewRequest = True


    strRequestType = "Null"
    '���������������ͬ��������δ���ʱ�����ܽ�������
    Select Case rtRequestType
        Case TRequestType.rtMianyi
            strRequestType = "�����黯"
        Case TRequestType.rtTeran
            strRequestType = "����Ⱦɫ"
        Case TRequestType.rtFenzi
            strRequestType = "���Ӳ���"
        Case TRequestType.rtSlices
            strRequestType = "����Ƭ"
        Case TRequestType.rtMaterial
            strRequestType = "��ȡ��"
    End Select
    
    For i = 1 To ufgRequest.GridRows - 1
        If ufgRequest.Text(i, gstrRequisition_��ǰ״̬) = "������" _
            And ufgRequest.Text(i, gstrRequisition_��������) Like "*" & strRequestType & "*" Then
            CheckAllowNewRequest = False
            Exit For
        End If
    Next i

    If Not CheckAllowNewRequest Then
        Call MsgBoxD(Me, "�ü�������δ��ɵ����룬����ִ�дβ�����", vbOKOnly, Me.Caption)
    End If
End Function


Private Sub Menu_Edit_ReqSpecialExam()
On Error GoTo errHandle
    '��ʾ�ؼ�����
    If Not CheckAllowNewRequest(-1) Then Exit Sub
    
    Call ShowSpecialExamRequestWindow
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Form_Load()
On Error GoTo errHandle
    Call InitCommandBars
    
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitFace
    
    '��ʼ�������б�
    Call InitRequisitionList
    
    Call ChangeControlFace(-1)
    
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
    With cbrMenuBar.CommandBar
        Set cbrControl = .Controls.Add(xtpControlButtonPopup, conMenu_View_ToolBar, "����"): cbrControl.IconId = 3903
        With cbrControl.CommandBar '�����˵�
            Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtReqGet, "��ȡ����(&G)"): cbrPopControl.IconId = 10016
            Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtReqSlices, "��Ƭ����(&S)"): cbrPopControl.IconId = 10017
            Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtReqSpecialExam, "�ؼ�����(&T)"): cbrPopControl.IconId = 10018
        End With
        Set cbrControl = .Controls.Add(xtpControlButton, TMenuType.mtDelReq, "��������(&D)"): cbrControl.IconId = 3565
        
        Set cbrControl = .Controls.Add(xtpControlButton, TMenuType.mtAddSpeExamPro, "���(&A)"): cbrControl.IconId = 4010
        cbrControl.BeginGroup = True
        Set cbrControl = .Controls.Add(xtpControlButton, TMenuType.mtDelSE, "ɾ��(&C)"): cbrControl.IconId = 4008
        Set cbrControl = .Controls.Add(xtpControlButton, TMenuType.mtNewSE, "����(&N)"): cbrControl.IconId = 3082
        Set cbrControl = .Controls.Add(xtpControlButton, TMenuType.mtRedoSE, "����(&U)"): cbrControl.IconId = 3945
        
'        Set cbrControl = .Controls.Add(xtpControlButton, TMenuType.mtAddSlicesPro, "����(&A)"): cbrControl.IconId = 4112
'        cbrControl.BeginGroup = True
'        Set cbrControl = .Controls.Add(xtpControlButton, TMenuType.mtDelSlicesPro, "ɾ��(&C)"): cbrControl.IconId = 4114
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
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "����"): cbrControl.IconId = 3903
        With cbrControl.CommandBar '�����˵�
            Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtReqGet, "��ȡ����(&G)"): cbrPopControl.IconId = 10016
            Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtReqSlices, "��Ƭ����(&S)"): cbrPopControl.IconId = 10017
            Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtReqSpecialExam, "�ؼ�����(&T)"): cbrPopControl.IconId = 10018
        End With
        
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtDelReq, "��������"): cbrControl.IconId = 3565
        
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAddSpeExamPro, "���"): cbrControl.IconId = 4010
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtDelSE, "ɾ��"): cbrControl.IconId = 4008
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtNewSE, "����"): cbrControl.IconId = 3082
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtRedoSE, "����"): cbrControl.IconId = 3945
        
'        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAddSlicesPro, "����"): cbrControl.IconId = 4112
'        cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, TMenuType.mtDelSlicesPro, "ɾ��"): cbrControl.IconId = 4114
        
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

Private Sub picRequest_Resize()
On Error Resume Next
    '�������沼��
    ufgRequest.Left = 120
    ufgRequest.Top = 0
    ufgRequest.Width = picRequest.Width - 240
    ufgRequest.Height = picRequest.Height - 60
End Sub


Private Sub picRequestContext_Resize()
On Error Resume Next
    '����picRequestContext������
    labSubTitle.Left = 120
    labSubTitle.Top = 120
    
    linSubTitle.X1 = 0
    linSubTitle.Y1 = labSubTitle.Top + 90
    linSubTitle.X2 = picRequestContext.Width
    linSubTitle.Y2 = labSubTitle.Top + 90
    
    ufgContext.Left = 120
    ufgContext.Top = 400
    ufgContext.Width = picRequestContext.Width - 240
    ufgContext.Height = picRequestContext.Height - 360
End Sub


Private Sub LoadRequestInf(ByVal lngPatholAdviceId As Long)
'����������Ϣ
    Dim strSql As String
    
    strSql = "select ����ID,������,��������,����״̬,����ϸĿ,����ʱ��,����״̬,��������,���ʱ�� from ����������Ϣ where ����ҽ��ID=[1] and (��������=-1 "
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    If CheckPopedom(mstrPrivs, "�ؼ�����") Then strSql = strSql & " or ��������<=2"
    If CheckPopedom(mstrPrivs, "��Ƭ����") Then strSql = strSql & " or ��������=3"
    If CheckPopedom(mstrPrivs, "��ȡ����") Then strSql = strSql & " or ��������=4"
    
    strSql = strSql & ")order by ��������,����ʱ��"
    
    Set ufgRequest.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatholAdviceId)
    
    Call ufgRequest.RefreshData
End Sub


Private Sub LoadSpeExamRequestContext(ByVal lngRequestId As Long)
'��ȡ�ؼ���������
    Dim strSql As String
    
    strSql = "select a.ID,a.�Ŀ�ID,b.���,b.�걾����,c.����ID, b.�걾����,c.��������,a.��������,a.��ǰ״̬,a.��Ŀ���,a.���ʱ��,a.�ؼ�ҽʦ " & _
                " from �����ؼ���Ϣ a, ����ȡ����Ϣ b, ��������Ϣ c " & _
                " where a.�Ŀ�id = b.�Ŀ�id and a.����id=c.����id and a.����id=[1] order by a.��������, a.�Ŀ�ID, c.��������"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgContext.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngRequestId)
    
    Call ufgContext.RefreshData
End Sub




Private Sub LoadSlicesRequestContext(ByVal lngRequestId As Long)
'��ȡ��Ƭ��������
    Dim strSql As String
    
    strSql = "select a.ID,a.�Ŀ�ID,b.���,b.�걾����,a.��Ƭ����,a.��Ƭ��ʽ,a.��Ƭ��,a.��ǰ״̬,a.��Ƭʱ��,a.��Ƭ�� " & _
            " from ������Ƭ��Ϣ a, ����ȡ����Ϣ b " & _
            " where a.�Ŀ�id=b.�Ŀ�id and a.����id=[1] order by a.��ǰ״̬, b.�걾����,a.�Ŀ�ID"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgContext.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngRequestId)
    
    Call ufgContext.RefreshData
End Sub




Private Sub LoadSupMaterialRequestContext(ByVal lngRequestId As Long)
'��ȡȡ�ĵ��������
    Dim strSql As String
    
    strSql = "select �Ŀ�ID,���,�걾����,�걾��,������,ȡ��ʱ��,��ȡҽʦ,��ȡҽʦ,��¼ҽʦ " & _
            " from ����ȡ����Ϣ where  ����id=[1] order by ȡ��ʱ��"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgContext.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngRequestId)
    
    Call ufgContext.RefreshData
End Sub


Private Sub ShowAntibodyInf(ByVal lngAntibodyRow As Long)
'��ʾ������ϸ��Ϣ
    Dim frmAntibodyInf As New frmPatholRequisition_AntibodyInf
    On Error GoTo errFree
        Call frmAntibodyInf.ShowAntibodyInf(ufgContext.Text(lngAntibodyRow, gstrRequest_SpeExam_����ID), Me)
errFree:
    Call Unload(frmAntibodyInf)
    Set frmAntibodyInf = Nothing
    
End Sub



Private Sub ufgContext_OnCellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo errHandle
    Call ShowAntibodyInf(Row)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub ufgRequest_OnClick()
'��ȡ��������
On Error GoTo errHandle
    Dim strRequestType As String
    
    '���������Ŀ��ϸ
    Call ufgContext.ClearListData
    
    If Not ufgRequest.IsSelectionRow Then Exit Sub
    If ufgRequest.IsEmptyKey(ufgRequest.SelectionRow) Then Exit Sub
    
    strRequestType = ufgRequest.Text(ufgRequest.SelectionRow, gstrRequisition_��������)
    
    Select Case strRequestType
        Case "�����黯", "���Ӳ���", "����Ⱦɫ"
        
            Call InitRequestContextList(0)
            Call ChangeControlFace(0)
            
            '��ȡ�ؼ���Ŀ��ϸ
            Call LoadSpeExamRequestContext(ufgRequest.KeyValue(ufgRequest.SelectionRow))
            
            Case "����Ƭ", "����", "����", "����", "��Ƭ", "��Ⱦ", "��Ƭ"
            
            Call InitRequestContextList(3)
            Call ChangeControlFace(3)
             
            '��ȡ��Ƭ��Ŀ��ϸ
            Call LoadSlicesRequestContext(ufgRequest.KeyValue(ufgRequest.SelectionRow))

        Case "��ȡ��", "��ȡ��"
            
            Call InitRequestContextList(4)
            Call ChangeControlFace(4)
            
            '��ȡȡ����Ŀ��ϸ
            Call LoadSupMaterialRequestContext(ufgRequest.KeyValue(ufgRequest.SelectionRow))
            
    End Select
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgRequest_OnColsNameReSet()
On Error GoTo errHandle

   '��ȡ������Ϣ
    Call LoadRequestInf(mrecStudyInf.lngPatholAdviceId)
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ufgRequest_OnSelChange()
    ufgRequest_OnClick
End Sub
