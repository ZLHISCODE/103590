VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.0#0"; "zlRichEditor.ocx"
Begin VB.Form frmEPRModelsMan 
   Caption         =   "�������İ�"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11535
   Icon            =   "frmEPRModelsMan.frx":0000
   ScaleHeight     =   7275
   ScaleWidth      =   11535
   StartUpPosition =   1  '����������
   Begin zlRichEditor.Editor EdoTmp 
      Height          =   375
      Left            =   5865
      TabIndex        =   7
      Top             =   4965
      Visible         =   0   'False
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   661
   End
   Begin VB.PictureBox PicModels 
      Appearance      =   0  'Flat
      BackColor       =   &H00E7CFBA&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6720
      Left            =   45
      ScaleHeight     =   6720
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   495
      Width           =   5355
      Begin VB.TextBox txtSeek 
         Height          =   270
         Left            =   4170
         TabIndex        =   3
         ToolTipText     =   "�����س��������Ʋ��ң���������붨λ��"
         Top             =   165
         Width           =   1170
      End
      Begin MSComctlLib.ListView lvwModels 
         Height          =   5490
         Left            =   30
         TabIndex        =   1
         Top             =   750
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   9684
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.CheckBox chklevel 
         BackColor       =   &H00E7CFBA&
         Caption         =   "ȫԺͨ��"
         Height          =   225
         Index           =   0
         Left            =   75
         TabIndex        =   4
         Top             =   195
         Width           =   1035
      End
      Begin VB.CheckBox chklevel 
         BackColor       =   &H00E7CFBA&
         Caption         =   "����ͨ��"
         Height          =   225
         Index           =   1
         Left            =   1170
         TabIndex        =   5
         Top             =   195
         Width           =   1035
      End
      Begin VB.CheckBox chklevel 
         BackColor       =   &H00E7CFBA&
         Caption         =   "����ʹ��"
         Height          =   225
         Index           =   2
         Left            =   2265
         TabIndex        =   6
         Top             =   195
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E7CFBA&
         Caption         =   "���ƹ���"
         Height          =   165
         Left            =   3390
         TabIndex        =   2
         Top             =   210
         Width           =   780
      End
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmEPRModelsMan.frx":6852
      Left            =   2235
      Top             =   15
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmEPRModelsMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mModelInfo As frmEPRModelsInfo
Private mModelcontent As frmEPRModelsContent
Private mbytMode As Byte '��ǰģʽ 0-���� 1-���� 2-�޸�
Private mstrPrivs As String 'Ȩ��
Private mlngPatiId As Long, mlngPageId As Long, mlngDeptId As Long
Private mblnApply As Boolean '�Ƿ�Ӧ��ѡ������
Public Function Showfrm(ByVal parentfrm As Object, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptId As Long, ByVal strPrivs As String) As Boolean
    mblnApply = False
    mstrPrivs = strPrivs
    mlngPatiId = lngPatiID: mlngPageId = lngPageId: mlngDeptId = lngDeptId
    Me.Show 1, parentfrm
    Showfrm = mblnApply
End Function
Private Sub InitFace()
'��ʼ���沼��
Dim PaneTmp As Pane
    With Me.dkpMain
        .SetCommandBars cbrMain
        .Options.HideClient = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    Set PaneTmp = dkpMain.CreatePane(1, 550, 580, DockLeftOf, Nothing)
    PaneTmp.Title = "���İ��б�"
    PaneTmp.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Set PaneTmp = dkpMain.CreatePane(2, 550, 0, DockRightOf, Nothing)
    PaneTmp.MaxTrackSize.Height = 200
    PaneTmp.Title = "���İ��༭"
    PaneTmp.Options = PaneNoCaption
    Set PaneTmp = dkpMain.CreatePane(3, 550, 0, DockBottomOf, PaneTmp)
    PaneTmp.Title = "���İ�����"
    PaneTmp.Options = PaneNoCaption
    
End Sub
Private Sub InitCommandButton()
'���ܣ���ʼ��������
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbrMain.VisualTheme = xtpThemeOffice2003
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbrMain.EnableCustomization False
    cbrMain.ActiveMenuBar.Visible = False
    Set cbrMain.Icons = zlCommFun.GetPubIcons
    
    '���ɹ�����
    Set objBar = cbrMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "����")
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Apply, "Ӧ��ѡ�����İ�"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyTo, "���İ����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next
    
    '�ȼ���:ע�ⲻ�ܺ�ϵͳ���ı��༭�ȼ���ͻ
    With cbrMain.KeyBindings
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify
        .Add FCONTROL, vbKeyS, conMenu_Edit_Save
        .Add FALT, vbKeyX, conMenu_File_Exit
    End With
End Sub
Private Sub Initlvw()
    With lvwModels.ColumnHeaders
        .Clear
        .Add , "_ID", "ID", 0
        .Add , "_���", "���", 800
        .Add , "_����", "����", 2000
        .Add , "_����", "����", 600
        .Add , "_˵��", "˵��", 1800
        .Add , "_ͨ�ü�", "ͨ�ü�", 1000
        .Add , "_����ID", "����ID", 0
        .Add , "_��ԱID", "��ԱID", 0
        .Add , "_����", "����", 800
        .Add , "_��Ա", "��Ա", 800
        
    End With
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
'mbytMode As Byte '��ǰģʽ 0-���� 1-���� 2-�޸�
    Select Case Control.ID
        Case conMenu_Edit_NewItem ' "����")
            mbytMode = 1: Call mModelInfo.zlRefresh("", mstrPrivs)
            PicModels.Enabled = False
            Me.dkpMain.FindPane(2).Select
            mModelInfo.Enabled = True
        Case conMenu_Edit_Modify '"�޸�")
            mbytMode = 2: mModelInfo.zlEditStart '������ѡ��ʱ��ˢ��
            PicModels.Enabled = False
            Me.dkpMain.FindPane(2).Select
            mModelInfo.Enabled = True
        Case conMenu_Edit_Delete '"ɾ��")
            zlDelModels
        Case conMenu_Edit_Untread 'ȡ��
            mbytMode = 0: mModelInfo.zlEndEdit: InitData
            PicModels.Enabled = True
        Case conMenu_Edit_Save '"����")
            If mModelInfo.zlSaveData Then mbytMode = 0: InitData: PicModels.Enabled = True
        Case conMenu_Tool_Apply '"Ӧ��ѡ�����İ�")
            Call zlModelsApply
        Case conMenu_Edit_ApplyTo '"���İ����")
            Call mModelcontent.zlRefresh(lvwModels.SelectedItem.Text, mstrPrivs, 1)
        Case conMenu_Help_Help '"����")
            
        Case conMenu_File_Exit
            Unload Me
    End Select
End Sub

Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim l As Long 'mbytMode As Byte '��ǰģʽ 0-���� 1-���� 2-�޸�
    l = lvwModels.ListItems.Count

    Select Case Control.ID
        Case conMenu_Edit_NewItem ' "����")
            Control.Enabled = InStr(mstrPrivs, "�������İ�����") > 0 And mbytMode = 0
        Case conMenu_Edit_Modify '"�޸�")
            Control.Enabled = InStr(mstrPrivs, "�������İ�����") > 0 And mbytMode = 0 And l > 0
        Case conMenu_Edit_Delete '"ɾ��")
            Control.Enabled = InStr(mstrPrivs, "�������İ�����") > 0 And mbytMode = 0 And l > 0
        Case conMenu_Edit_Untread 'ȡ��  ȡ��������ȡ���޸�
            Control.Enabled = InStr(mstrPrivs, "�������İ�����") > 0 And (mbytMode = 1 Or mbytMode = 2)
        Case conMenu_Edit_Save '"����")
            Control.Enabled = InStr(mstrPrivs, "�������İ�����") > 0 And (mbytMode = 1 Or mbytMode = 2)
        Case conMenu_Tool_Apply '"Ӧ��ѡ�����İ�")
            Control.Enabled = mbytMode = 0 And l > 0
        Case conMenu_Edit_ApplyTo '"���İ����")
            Control.Enabled = InStr(mstrPrivs, "�������İ�����") > 0 And mbytMode = 0 And l > 0
        Case conMenu_Help_Help '"����")
        
        Case conMenu_File_Exit
        
    End Select
End Sub
Private Sub zlModelsApply()
Dim arrSQL() As Variant, i As Integer, blnTran As Boolean, Doc As New cEPRDocument, rsTemp As ADODB.Recordset, strFileIDS As String
    On Error GoTo ErrHandle
'    arrSQL = Array()

    With mModelcontent.lvwModelContent
        For i = 1 To .ListItems.Count                       '��鲡���ļ�
            If .ListItems(i).Checked = True Then
                strFileIDS = strFileIDS & "," & .ListItems(i).SubItems(11)
            End If
        Next
        If strFileIDS <> "" Then strFileIDS = Mid(strFileIDS, 2)
        gstrSQL = "Select A.�������� From ���Ӳ�����¼ A,Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) B Where A.����ID=[2] and A.��ҳID=[3] AND A.�ļ�ID=B.Column_Value"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strFileIDS, mlngPatiId, mlngPageId)
        If Not rsTemp.EOF Then
            If MsgBox("��ǰѡ�еĲ����ļ��� [" & NVL(rsTemp!��������) & "] �Ѿ���д�����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                Exit Sub
            End If
        End If
        
        For i = 1 To .ListItems.Count
            If .ListItems(i).Checked = True Then
                Call Doc.InitEPRDoc(cprEM_����, cprET_�������༭, .ListItems(i).SubItems(11), cprPF_סԺ, mlngPatiId, mlngPageId, , mlngDeptId)
                Call Doc.ImportEPRDemo(EdoTmp, .ListItems(i).Tag, True)
                Call Doc.SaveEPRDoc(EdoTmp)
'                ReDim Preserve arrSQL(UBound(arrSQL) + 1)'�����ݿ���̲���,��ͨ������ʵ��,�ٶȻ����
'                arrSQL(UBound(arrSQL)) = "Zl_������������_Apply(" & .ListItems(i).Tag & "," & mlngPatiID & "," & mlngPageID & "," & mlngdeptID & ",'" & gstrUserName & "')"
            End If
        Next
    End With

'    gcnOracle.BeginTrans '--------------------------д������
'    blnTran = True
'    For i = 0 To UBound(arrSQL)
'        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "")
'    Next
'    gcnOracle.CommitTrans: blnTran = False
    mblnApply = True
    Unload Me
    Exit Sub
ErrHandle:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub zlDelModels()
    On Error GoTo ErrHandle
    gstrSQL = "zl_�������İ�_Delete(" & lvwModels.SelectedItem.Text & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    Call InitData
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub chklevel_Click(Index As Integer)
Dim i As Integer, blnOnly As Boolean
    For i = 0 To chklevel.UBound
        If chklevel(i).Enabled Then
            If chklevel(i).Value = vbChecked Then
                blnOnly = True: Exit For 'ֻҪ�б�ѡ�м��˳�
            End If
        End If
    Next
    
    If blnOnly = False Then chklevel(Index).Value = vbChecked '��֤ʼ����һ����ѡ��
    Call InitData
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case 1
            Item.Handle = PicModels.hwnd
        Case 2
            Item.Handle = mModelInfo.hwnd
            Item.MaxTrackSize.Height = 100
            Item.MinTrackSize.Height = 100
        Case 3
            Item.Handle = mModelcontent.hwnd
    End Select
            Item.Selected = True
End Sub

Private Sub Form_Load()
    If mModelInfo Is Nothing Then Set mModelInfo = New frmEPRModelsInfo
    If mModelcontent Is Nothing Then Set mModelcontent = New frmEPRModelsContent
    mbytMode = 0
    InitFace
    InitCommandButton
    Initlvw
    RefreshModels
    RestoreWinState Me, App.ProductName
End Sub
Private Sub RefreshModels()
    If InStr(mstrPrivs, "���˲�������") <= 0 Then chklevel(2).Enabled = False: chklevel(2).Value = False
    If InStr(mstrPrivs, "���Ҳ�������") <= 0 Then chklevel(1).Enabled = False: chklevel(1).Value = False
    If InStr(mstrPrivs, "ȫԺ��������") <= 0 Then chklevel(0).Enabled = False: chklevel(0).Value = False
    InitData
End Sub
Private Sub InitData()
    Dim rsTemp As ADODB.Recordset, objItem As ListItem, lngID As Long
    On Error GoTo ErrHandle
    If lvwModels.ListItems.Count > 0 Then
        lngID = lvwModels.SelectedItem.Text
    End If
    
    gstrSQL = ""
    If chklevel(0).Value = vbChecked Then gstrSQL = "A.ͨ�ü�=0" 'ȫԺͨ��
    If chklevel(1).Value = vbChecked Then gstrSQL = gstrSQL & IIf(gstrSQL = "", "", " or ") & "(A.ͨ�ü�=1 and A.����ID=[1])" '����ͨ��
    If chklevel(2).Value = vbChecked Then gstrSQL = gstrSQL & IIf(gstrSQL = "", "", " or ") & "(A.ͨ�ü�=2 and A.��ԱID=[2])" '����ʹ��
    If chklevel(0).Value = vbChecked And chklevel(1).Value = vbChecked And chklevel(2).Value = vbChecked Then gstrSQL = "" ''ȫѡ
    
    If gstrSQL = "" Then '����Ȩ�޼�����
        If chklevel(0).Enabled Then gstrSQL = "A.ͨ�ü�=0"
        If chklevel(1).Enabled Then gstrSQL = gstrSQL & IIf(gstrSQL = "", "", " OR ") & "(A.ͨ�ü�=1 and A.����ID=[1])"
        If chklevel(2).Enabled Then gstrSQL = gstrSQL & IIf(gstrSQL = "", "", " OR ") & "(A.ͨ�ü�=2 and A.��ԱID=[2])"
    End If
    
    gstrSQL = "select A.ID,A.���,A.����,A.����,A.˵��,A.ͨ�ü�,A.����ID,A.��ԱID,B.���� ����,C.���� " & _
                " from �������İ� A,���ű� B,��Ա�� C " & _
                " where A.����ID=B.ID AND A.��ԱID=C.ID " & IIf(gstrSQL = "", "", " AND (" & gstrSQL & ")")
    If Trim(txtSeek.Text) <> "" Then
        gstrSQL = gstrSQL & " And " & zlCommFun.GetLike("A", "����", Trim(txtSeek))
    End If
    gstrSQL = gstrSQL & " Order by A.ͨ�ü�,A.���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngDeptId, glngUserId)
    lvwModels.ListItems.Clear
    With rsTemp
        Do Until .EOF
            Set objItem = lvwModels.ListItems.Add(, "_" & !ID, !ID)
                objItem.SubItems(1) = !���
                objItem.SubItems(2) = !����
                objItem.SubItems(3) = NVL(!����)
                objItem.SubItems(4) = NVL(!˵��)
                objItem.SubItems(5) = Decode(NVL(!ͨ�ü�, 0), 0, "ȫԺͨ��", 1, "����ͨ��", "����ʹ��")
                objItem.SubItems(6) = NVL(!����ID, 0)
                objItem.SubItems(7) = NVL(!��ԱID, 0)
                objItem.SubItems(8) = NVL(!����, 0)
                objItem.SubItems(9) = NVL(!����, 0)
            If !ID = lngID Then
                objItem.Selected = True
            End If
            .MoveNext
        Loop
    End With
    If lvwModels.ListItems.Count > 0 Then
        If lvwModels.SelectedItem Is Nothing Then
            lvwModels.ListItems(1).Selected = True
        End If
        lvwModels_ItemClick lvwModels.SelectedItem
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not mModelInfo Is Nothing Then Unload mModelInfo: Set mModelInfo = Nothing
    If Not mModelcontent Is Nothing Then Unload mModelcontent: Set mModelcontent = Nothing
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwModels_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call mModelInfo.zlRefresh(Item.Text & "|" & Item.SubItems(1) & "|" & Item.SubItems(2) & "|" & Item.SubItems(3) & "|" & Item.SubItems(4) & "|" & Item.SubItems(5) & "|" & Item.SubItems(6) & "|" & Item.SubItems(7), mstrPrivs) 'Ȩ�����ھ���ͨ�ü����޸ģ�ֻ����ӵ�м��ڱ䶯
    Call mModelcontent.zlRefresh(Item.Text, mstrPrivs, 0)
End Sub

Private Sub PicModels_Resize()
On Error Resume Next
    With lvwModels
        .Top = chklevel(0).Top + chklevel(0).Height + 100
        .Width = PicModels.Width
        .Left = 0
        .Height = PicModels.Height - .Top
    End With
End Sub
Private Sub txtSeek_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call InitData
    ElseIf InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyAscii))) > 0 Then '���붨λ
        Dim i As Integer, strtmp As String
        If txtSeek.SelLength > 0 Then
            strtmp = ""
        Else
            strtmp = txtSeek.Text
        End If
        For i = 1 To lvwModels.ListItems.Count
            If UCase(lvwModels.ListItems(i).SubItems(3)) Like UCase(Trim(strtmp)) & UCase(Chr(KeyAscii)) & "*" Then
                lvwModels.ListItems(i).Selected = True: lvwModels_ItemClick lvwModels.SelectedItem: Exit Sub
            End If
        Next
    End If
End Sub
