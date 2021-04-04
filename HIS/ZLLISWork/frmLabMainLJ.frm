VERSION 5.00
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmLabMainLJ 
   Caption         =   "�ʿز�ѯ"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11760
   Icon            =   "frmLabMainLJ.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   11760
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cbo��Ŀ 
      Height          =   300
      Left            =   1860
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   3045
   End
   Begin C1Chart2D8.Chart2D chtCopy 
      Height          =   435
      Left            =   1515
      TabIndex        =   1
      Top             =   660
      Visible         =   0   'False
      Width           =   765
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   1349
      _ExtentY        =   767
      _StockProps     =   0
      ControlProperties=   "frmLabMainLJ.frx":058A
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   360
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmLabMainLJ.frx":0BE9
      Left            =   975
      Top             =   330
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmLabMainLJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmChartLJ As New frmQCChartLJ                 'LJ����ͼ����
Private mfrmQCTodayReport As New frmQCTodayReport       '��дʧ�ؼ�¼
Private mlngSampleID As Long                            '�걾ID
Private mstrQCID As String                              '�ʿ�ƷID
Private mlngMachineID As Long                           '����ID
Private mlngResult As Long                              '��ͨ���ID
Private mEditMode As Integer                            '�༭ģʽ 0=�Ǳ༭ 1=���ڱ༭
Private mstrPigeonhole As String                        '�鵵��
Private mstrReportMan As String                         '������
Private mstrStart As String, mstrEnd As String

Private Sub cbo��Ŀ_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim strStartDate As String, strEndDate As String
    Dim strDateSpace As String
    Dim strNowDate As String
    
    If Me.cbo��Ŀ.ListCount = 0 Then Exit Sub
    
    '�õ�ǰʱ��
    gstrSql = "select nvl(����ʱ��,sysdate) as ����ʱ�� from ����걾��¼ where id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngSampleID)
    If rsTmp.EOF = True Then
        MsgBox "û���ҵ���Ӧ�ı걾!", vbInformation, gstrSysName: Exit Sub
    End If
    strNowDate = Nvl(rsTmp("����ʱ��"))
    
    strStartDate = Format(getMonthFirst(CDate(strNowDate)), "yyyy-mm-dd"): strEndDate = Format(getMonthLast(CDate(strNowDate)), "yyyy-mm-dd")

    mstrStart = strStartDate
    mstrEnd = strEndDate
    '-----------------------------------------------------------------------------------------------------------------------
    '�õ��ʿ�Ʒ
    mstrQCID = ""
'    gstrSql = "Select M.ID, '' As ѡ��, M.���� , M.���� || ', ˮƽ:' || M.ˮƽ As �ʿ�Ʒ, M.ˮƽ" & vbNewLine & _
'            "From �����ʿ�Ʒ M, �����ʿ�Ʒ��Ŀ I, �����ʿؾ�ֵ X, ( Select Distinct ����id From �����ʿؼ�¼ Where �걾id = [1] ) Y " & vbNewLine & _
'            "Where M.ID = I.�ʿ�Ʒid And I.�ʿ�Ʒid = X.�ʿ�Ʒid And I.��Ŀid = X.��Ŀid And M.����id = Y.����id And I.��Ŀid = [2] And" & vbNewLine & _
'            "      X.�ڼ� = [3]" & vbNewLine & _
'            "Order By M.��ʼ����, M.ˮƽ"
'    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngSampleID, CLng(Me.cbo��Ŀ.ItemData(Me.cbo��Ŀ.ListIndex)), _
                strDateSpace)
    gstrSql = "Select Distinct M.ID, '' As ѡ��, M.����, M.���� || ', ˮƽ:' || M.ˮƽ As �ʿ�Ʒ, M.ˮƽ, M.��ʼ����" & vbNewLine & _
        "From �����ʿ�Ʒ M, �����ʿ�Ʒ��Ŀ I, �����ʿؾ�ֵ X, (Select Distinct ����id From �����ʿؼ�¼ Where �걾id = [1]) Y" & vbNewLine & _
        "Where M.ID = I.�ʿ�Ʒid And I.�ʿ�Ʒid = X.�ʿ�Ʒid And I.��Ŀid = X.��Ŀid And M.����id = Y.����id And I.��Ŀid = [2] And" & vbNewLine & _
        "      To_Date([3], 'YYYY-MM-DD') Between X.��ʼ���� And Nvl(X.��������, M.��������) and " & vbNewLine & _
        "      To_Date([4], 'YYYY-MM-DD') Between X.��ʼ���� And Nvl(X.��������, M.��������)" & vbNewLine & _
        "Order By M.��ʼ����, M.ˮƽ"
    
    gstrSql = "Select Id,ѡ��,����,�ʿ�Ʒ,ˮƽ,min(��ʼ����) As ��ʼ����,Min(��������) As ��������" & vbNewLine & _
            "From (" & vbNewLine & _
            "Select M.ID, '' As ѡ��, M.���� , M.���� || ', ˮƽ:' || M.ˮƽ As �ʿ�Ʒ, M.ˮƽ, to_Char(X.��ʼ����,'yy-MM-dd') as ��ʼ����,to_char(Nvl(X.��������, M.��������),'yy-MM-dd')  as ��������" & vbNewLine & _
            "From �����ʿ�Ʒ M, �����ʿ�Ʒ��Ŀ I, �����ʿؾ�ֵ X, (Select Distinct ����id From �����ʿؼ�¼ Where �걾id = [1]) Y" & vbNewLine & _
            "Where M.ID = I.�ʿ�Ʒid And I.�ʿ�Ʒid = X.�ʿ�Ʒid And I.��Ŀid = X.��Ŀid And M.����id = Y.����ID And I.��Ŀid = [2] And" & vbNewLine & _
            "      (To_Date([3], 'yyyy-MM-dd') Between X.��ʼ���� And Nvl(X.��������, M.��������))" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select M.ID, '' As ѡ��, M.���� , M.���� || ', ˮƽ:' || M.ˮƽ As �ʿ�Ʒ, M.ˮƽ, to_Char(X.��ʼ����,'yy-MM-dd') as ��ʼ����,to_char(Nvl(X.��������, M.��������),'yy-MM-dd')  as ��������" & vbNewLine & _
            "From �����ʿ�Ʒ M, �����ʿ�Ʒ��Ŀ I, �����ʿؾ�ֵ X, (Select Distinct ����id From �����ʿؼ�¼ Where �걾id = [1]) Y" & vbNewLine & _
            "Where M.ID = I.�ʿ�Ʒid And I.�ʿ�Ʒid = X.�ʿ�Ʒid And I.��Ŀid = X.��Ŀid And M.����id = Y.����ID And I.��Ŀid = [2] And" & vbNewLine & _
            "        (  (X.��ʼ���� Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd'))" & vbNewLine & _
            "         Or" & vbNewLine & _
            "          (nvl(X.��������,Sysdate) Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd')+1-1/24*60*60)" & vbNewLine & _
            "         )" & vbNewLine & _
            "       )" & vbNewLine & _
            "Group By      Id,ѡ��,����,�ʿ�Ʒ,ˮƽ" & vbNewLine & _
            "Order By �ʿ�Ʒ,ˮƽ"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngSampleID, CLng(Me.cbo��Ŀ.ItemData(Me.cbo��Ŀ.ListIndex)), _
                CStr(Format(mstrStart, "yyyy-mm-dd")), CStr(Format(mstrEnd, "yyyy-mm-dd")))
                
    strDateSpace = ""
    Do Until rsTmp.EOF
        mstrQCID = mstrQCID & "," & Val(Nvl(rsTmp("ID")))
        strDateSpace = strDateSpace & ";" & Val("" & rsTmp("ID")) & "=" & Format("" & rsTmp("��ʼ����"), "yyyy-MM-dd") & "," & Format("" & rsTmp("��������"), "yyyy-MM-dd")
        rsTmp.MoveNext
    Loop
    mstrQCID = Mid(mstrQCID, 2)
    If strDateSpace <> "" Then strDateSpace = Mid(strDateSpace, 2)
    '-----------------------------------------------------------------------------------------------------------------------
    mfrmChartLJ.zlRefresh mstrQCID, cbo��Ŀ.ItemData(cbo��Ŀ.ListIndex), Format(strStartDate, "yyyy-mm-dd"), _
                        Format(strEndDate, "yyyy-mm-dd"), strDateSpace
    gstrSql = "select ID from ������ͨ��� where ����걾id = [1] and ������Ŀid = [2] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngSampleID, cbo��Ŀ.ItemData(cbo��Ŀ.ListIndex))
    If rsTmp.EOF = False Then mlngResult = rsTmp("ID")
    mfrmQCTodayReport.zlRefresh mlngResult
    
    gstrSql = "select ������, �鵵�� from �����ʿر��� where ���id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngResult)
    If rsTmp.EOF = False Then mstrPigeonhole = Trim(Nvl(rsTmp("�鵵��"))): mstrReportMan = Trim(Nvl(rsTmp("������")))
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As CommandBarControl
    Dim lngQC As Long
    Dim rsTmp As New ADODB.Recordset
    
    If Me.Visible = False Then Exit Sub
    
    On Error GoTo errH
    
    Select Case Control.ID
        Case conMenu_File_PrintSet                                  '��ӡ����
            Call zlPrintSet
        Case conMenu_File_Print                                     '��ӡ����ͼ
            Call mfrmChartLJ.ChartPrint: Call PrintQC_LJ(True)
        Case conMenu_Edit_Leave_Post                                '������ͼ
            Call mfrmChartLJ.ChartSaveAs
        Case conMenu_Edit_MarkMap                                   '���ƿ���ͼ
            Call mfrmChartLJ.ChartCopy
        Case conMenu_File_Exit                                      '�˳�
            Unload Me
        Case conMenu_View_ToolBar_Button                            '��׼��ť
            Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
            Me.cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text                              '�ı���ǩ
             For Each cbrControl In Me.cbsThis(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            Me.cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size                              '��ͼ��
            Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
            Me.cbsThis.RecalcLayout
            
        Case conMenu_Edit_Save                                      '����
            mlngResult = mfrmQCTodayReport.zlEditSave()
            If mlngResult <> 0 Then
                mfrmQCTodayReport.zlRefresh mlngResult
                mEditMode = 0
            End If
            
        Case conMenu_Edit_Untread                                   'ȡ��
            mfrmQCTodayReport.zlEditCancel
            mEditMode = 0
        
        Case conMenu_Edit_Adjust                                    '����
            Call mfrmQCTodayReport.ZlEditStart(mlngResult)
            mEditMode = 1
        
        Case conMenu_Edit_Archive                                   '�鵵
            gstrSql = "select �鵵�� from �����ʿر��� where ���id = [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngResult)
            If rsTmp.EOF = False Then
                If Nvl(rsTmp("�鵵��")) = "" Then
                    If MsgBox("���Ҫ����ǰʧ�ر���鵵��", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Sub
                    gstrSql = "Zl_�����ʿر���_Archive(" & mlngResult & ",0)"
                    zlDatabase.ExecuteProcedure gstrSql, Me.Caption
                    mstrPigeonhole = gstrDBUser
                Else
                    If MsgBox("��ʧ�ر����Ѿ��鵵�����ȡ���鵵��", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Sub
                    gstrSql = "Zl_�����ʿر���_Archive(" & mlngResult & ",1)"
                    zlDatabase.ExecuteProcedure gstrSql, Me.Caption
                    mstrPigeonhole = ""
                End If
            End If
            Call mfrmQCTodayReport.zlRefresh(mlngResult)
            
        Case conMenu_View_Refresh                                   'ˢ��
            Call cbo��Ŀ_Click
        Case conMenu_Tool_Analyse                                   'ʧ�ؼ���
            If InStr(mstrQCID, ",") > 0 Then
                lngQC = Mid(mstrQCID, 1, InStr(mstrQCID, ",") - 1)
            Else
                lngQC = mstrQCID
            End If
            frmQCCompute.ShowMe Me, mlngMachineID, cbo��Ŀ.ItemData(cbo��Ŀ.ListIndex), zlDatabase.Currentdate, lngQC
        Case conMenu_Tool_Define                                    '���¶�ֵ
            If InStr(mstrQCID, ",") > 0 Then
                lngQC = Mid(mstrQCID, 1, InStr(mstrQCID, ",") - 1)
            Else
                lngQC = mstrQCID
            End If
            frmQCRedefine.ShowMe Me, mlngMachineID, cbo��Ŀ.ItemData(cbo��Ŀ.ListIndex), zlDatabase.Currentdate, lngQC
        Case conMenu_Help_Web                                       'WEB�ϵ�����
            Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web_Home                                  '��ҳ
            Call zlHomePage(Me.hwnd)
        Case conMenu_Help_Web_Mail                                  '���ͷ���
            Call zlMailTo(Me.hwnd)
        Case conMenu_Help_About                                     '����
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    End Select
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsThis_Resize()
    If Me.Visible = True Then
        Me.dkpMan.RecalcLayout
    End If
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_View_ToolBar_Button
            Control.Checked = Me.cbsThis(2).Visible
        Case conMenu_View_ToolBar_Text
            Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
        Case conMenu_View_ToolBar_Size
            Control.Checked = Me.cbsThis.Options.LargeIcons
        Case conMenu_Edit_Adjust
            Control.Enabled = (mEditMode = 0 And mstrPigeonhole = "")
        Case conMenu_Edit_Archive
            Control.Enabled = (mEditMode = 0 And mstrReportMan <> "")
        Case conMenu_Edit_Save, conMenu_Edit_Untread
            Control.Enabled = (mEditMode = 1)
    End Select
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = mfrmChartLJ.hwnd
    Case 2
        Item.Handle = mfrmQCTodayReport.hwnd
    End Select
End Sub

Private Sub dkpMan_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Top = 500
End Sub

Private Sub dkpMan_Resize()
    Me.cbsThis.RecalcLayout
End Sub

Private Sub Form_Load()
    Dim cbrControl As CommandBarControl, cbrMenuBar As CommandBarControl, cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    '-----------------------------------------------------
    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, False)
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '�˵�����
    Me.cbsThis.ActiveMenuBar.Title = "�˵�"
'    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ����ͼ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Leave_Post, "������ͼ(&S)..."): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "���ƿ���ͼ(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "����(&R)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive, "�鵵(&T)")
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", -1, False)
    cbrMenuBar.ID = xtpControlPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Analyse, "ʧ�ؼ���(&Y)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Define, "���¶�ֵ(&N)")
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With

    Set cbrControl = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "��Ŀ")
    cbrControl.Flags = xtpFlagRightAlign
    Set cbrCustom = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlCustom, 0, "��Ŀ")
    cbrCustom.Handle = Me.cbo��Ŀ.hwnd: cbrCustom.Flags = xtpFlagRightAlign
    
    '�����
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("C"), conMenu_Edit_MarkMap
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_ESCAPE, conMenu_Edit_Untread
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
    End With

    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        'conMenu_Edit_Save
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Leave_Post, "���Ϊ"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive, "�鵵"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Analyse, "ʧ�ؼ���"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Define, "���¶�ֵ")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '����ͣ������
    Dim panThis As Pane, panChild As Pane

    With Me.dkpMan
        Set panThis = .CreatePane(1, 700, 1000, DockBottomOf, Nothing)
        panThis.Title = "�ʿ�ͼ��"
        panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        
        Set panChild = .CreatePane(2, 300, 1000, DockRightOf, panThis)
        panChild.Title = "ʧ�ؼ�¼"
        panChild.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    End With

    Set mfrmChartLJ = New frmQCChartLJ
    Set mfrmQCTodayReport = New frmQCTodayReport
    
    '����ָ�
'    Call RestoreWinState(Me, App.ProductName)

    '�õ���������Ŀ

    gstrSql = "Select Distinct B.ID, B.����, B.������, B.Ӣ���� " & vbNewLine & _
                " From ������ͨ��� A, ����������Ŀ B, �����ʿ�Ʒ��Ŀ C  " & vbNewLine & _
                " Where A.������Ŀid = B.ID And A.������Ŀid = C.��Ŀid And a.����걾ID = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngSampleID)
    With rsTmp
        Me.cbo��Ŀ.Clear
        Do While Not .EOF
            Me.cbo��Ŀ.AddItem !���� & ", " & !������ & "/" & !Ӣ����
            Me.cbo��Ŀ.ItemData(Me.cbo��Ŀ.NewIndex) = !ID
            .MoveNext
        Loop
        If Me.cbo��Ŀ.ListCount = 0 Then MsgBox "��δ��������ʿ�Ʒ���ã�", vbInformation, gstrSysName
        If Me.cbo��Ŀ.ListCount > 0 Then
            Me.cbo��Ŀ.ListIndex = 0
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Sub ShowMe(lngSampleID As Long, objfrm As Object, lngMachineID As Long)
    mlngSampleID = lngSampleID
    mlngMachineID = lngMachineID
    Me.Show , objfrm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrQCID = ""
    mlngSampleID = 0
    Unload mfrmChartLJ
    Set mfrmChartLJ = Nothing
End Sub
Private Function getMonthFirst(dtNow As Date) As String
    '����         �õ����µĵ�һ��
    '����         dtNow ��������
    
    getMonthFirst = Format(dtNow, "YYYY-MM")
    getMonthFirst = getMonthFirst & "-01"
    
End Function
Private Function getMonthLast(dtNow As Date) As String
    '����         �õ����µ����һ��
    '����         dtNow ��������
    Dim strYear As String
    Dim strMonth As String
    strYear = Format(dtNow, "YYYY")
    strMonth = Format(dtNow, "MM")
    If CInt(strMonth) = 12 Then strMonth = "00": strYear = CInt(strYear) + 1
    getMonthLast = Format(CDate(strYear & "-" & CInt(strMonth) + 1 & "-01") - 1, "yyyy-mm-dd")
    
End Function

Private Sub PrintQC_LJ(blnPrintMode As Boolean)
    '��ӡ��Ԥ��LJ�ʿ�ͼ
    '����           intPrintMode =1 ��ӡ =2 Ԥ��
    
    Dim rsTmp As New ADODB.Recordset
    Dim strPrintType As String                  '��Ӧ�ĵ���
    Dim strQCID As String                       '�ʿ�ƷID���ܻ�����","�ָ��Ķ��ID
    Dim lngQCID As Long                         '�����ʿ�ƷID
    Dim lngItemID As String                     '��ĿID
    Dim lngMachine As Long                      '����ID
    Dim intloop As Integer                      'ѭ���ִ�
    Dim intReportCount As Integer               'Ҫ��ӡ��ͼ����
    Dim intPrintType As Integer
    
    intReportCount = mfrmChartLJ.ChartPrint
    
    
    On Error GoTo errH
    
    strPrintType = "ZL1_INSIDE_1209_1"
    
    gstrSql = "Select b.w, b.h " & vbNewLine & _
                " From Zlreports a, Zlrptitems b" & vbNewLine & _
                " Where a.Id = b.����id And a.��� = [1] And b.���� = '�ʿ�ͼ'"
                
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strPrintType)
    'û���ҵ�ʱ�˳�
    If rsTmp.EOF Then
        MsgBox "�ڵ��ݶ�����û�ж���<�ʿ�ͼ>,���ڵ����ж���һ����Ϊ<�ʿ�ͼ>��ͼ���!", vbQuestion, Me.Caption
        Exit Sub
    End If
    
    For intloop = 0 To intReportCount - 1
        With Me.chtCopy
            .Load App.path & "\QC_Tmp" & intloop
            Kill App.path & "\QC_Tmp" & intloop
            .Width = Nvl(rsTmp("w"), 1280 * Screen.TwipsPerPixelX)
            .Height = Nvl(rsTmp("h"), 500 * Screen.TwipsPerPixelY)
            .Header.Text = ""
            .ChartLabels.RemoveAll
            .ChartArea.Location.Top = -5
            .ChartArea.Location.Height = .ChartArea.Location.Height + 15
            If intPrintType = 3 Then
                .ChartArea.Location.Left = 30
            End If
            .SaveImageAsJpeg App.path & "\QC" & intloop & ".jpg", 1000, False, False, False
        End With
    Next
    
    '�õ��ʿ�ƷID
    lngQCID = mfrmChartLJ.ZLGetLJ_QCID
    strQCID = mfrmChartLJ.ZLGetLJ_QCIDStr
    
    
    '�õ���ĿID
    If Me.cbo��Ŀ.ListCount = 0 Then Exit Sub
    lngItemID = CLng(Me.cbo��Ŀ.ItemData(Me.cbo��Ŀ.ListIndex))
    lngMachine = mlngMachineID
    
    If Dir(App.path & "\QC0.jpg") <> "" Then
        Call ReportOpen(gcnOracle, glngSys, strPrintType, Me, "�ʿ�ͼ=" & App.path & "\QC0.jpg", _
        "�ʿ�ƷID=" & lngQCID, "��ĿID=" & lngItemID, "��ʼ����=" & mstrStart, "��������=" & mstrEnd, _
        "����ID=" & lngMachine, "�ʿ�Ʒ��=" & IIf(strQCID = "", "0", strQCID), _
        "�ʿ�ͼ1=" & App.path & "\QC1.jpg", "�ʿ�ͼ2=" & App.path & "\QC2.jpg", _
        IIf(blnPrintMode, 2, 1))
    End If
    
    Kill App.path & "\QC*.jpg"
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

