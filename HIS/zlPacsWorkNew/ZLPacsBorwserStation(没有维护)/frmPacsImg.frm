VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmPACSImg 
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
      Bindings        =   "frmPacsImg.frx":0000
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPACSImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngAdviceID As Long, mlngSendNo As Long
Private mblnShowPic As Boolean
Private mstrPrivs As String
Private mblnMoved As Boolean
Private mblnAddImage As Boolean                 '�Ƿ�׷��ͼ��
Private mShowPhotoNumber As Integer
Private mblnLocalizerBackward As Boolean        '��λƬ����
Private iCurImageIndex As Integer
Public pobjPacsCore As zl9PacsCore.clsViewer
Private mintSelectAllSeq As Integer                 '0--��״̬��1--ѡ��ȫ�����У�2--��ѡ��ȫ������
Private mintSelectAllImg As Integer                 '0--��״̬��1--ѡ��ȫ��ͼ��2--��ѡ��ȫ��ͼ��

Public Function zlRefresh(ByVal lngAdviceID As Long, ByVal lngSendNO As Long, ByVal strPrivs As String, _
    Optional ByVal blnMoved As Boolean = False, Optional ByVal blnRefresh As Boolean = False) As Boolean
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo DBError
    If mlngAdviceID = lngAdviceID And mlngSendNo = lngSendNO And Not blnRefresh Then Exit Function
    mblnMoved = blnMoved
    mblnShowPic = False
    mlngAdviceID = lngAdviceID
    mlngSendNo = lngSendNO
    mstrPrivs = strPrivs

    'ת����Ӱ���ܱ��汨��
    If mblnMoved Then
        mstrPrivs = Replace(mstrPrivs, "ͼ���������", "")
        mstrPrivs = Replace(mstrPrivs, "ͼ���ע����", "")
        mstrPrivs = Replace(mstrPrivs, "���ͼ��", "")
    End If
    
    mShowPhotoNumber = 15
    strSQL = "select ID ,����ID,������,����ֵ from Ӱ�����̲��� where ����ID =  " & _
             "(Select ִ�в���ID From ����ҽ������ Where ҽ��ID =[1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngAdviceID)
    While Not rsTemp.EOF
        Select Case rsTemp!������
        Case "�����ʾ����ͼ��"
            mShowPhotoNumber = Abs(Nvl(rsTemp!����ֵ, 15))
        Case "��λƬ����"
            mblnLocalizerBackward = Nvl(rsTemp!����ֵ)
        End Select
        rsTemp.MoveNext
    Wend
    
    Call ShowSeqImg
    
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'ִ�в˵�����
Public Sub zlMenuClick(mnuClick As String)
    
    mblnAddImage = False
    Select Case mnuClick
        Case "Ӱ����"
            DViewer_DblClick
        Case "Ӱ��Ա�"
            mblnAddImage = True
            DViewer_DblClick
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
            Call WriteSelectdImages(lvwImage.Tag)
    End Select
End Sub

Private Sub subSetMenuState()
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
        If Not lvwSeq.SelectedItem Is Nothing Then
            ShowImageList lvwSeq.SelectedItem
        Else
            ShowImageList Nothing
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
    Call WriteSelectdImages(lvwImage.Tag)
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
            Call zlRefresh(mlngAdviceID, mlngSendNo, mstrPrivs, mblnMoved, True)
    End Select
End Sub

Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    Select Case control.ID
        Case conMenu_View_Expend_AllCollapse, conMenu_View_Expend_AllExpend, conMenu_Manage_SelectAllImages, _
             conMenu_Manage_UnSelectAllImages, conMenu_Manage_ReverseSelectImages
            control.Enabled = lvwSeq.ListItems.Count > 0
        Case conMenu_View_Show
            control.Enabled = lvwSeq.ListItems.Count > 0
            control.Checked = mblnShowPic
    End Select
End Sub

Private Sub DkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = lvwSeq.Hwnd
    ElseIf Item.ID = 2 Then
        Item.Handle = lvwImage.Hwnd
    ElseIf Item.ID = 3 Then
        Item.Handle = picView.Hwnd
    End If
End Sub

Private Sub DViewer_DblClick()
'��ʾ��Ƭվ
    Dim strSerials As String, strSeqUID As String
    Dim Item As MSComctlLib.ListItem
    Dim intImageInverval As Integer
    Dim strImages As String
    
    On Error GoTo CallError
    If lvwSeq.SelectedItem Is Nothing Then Exit Sub
    
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
    If Len(strSerials) = 0 Then         'û��ѡ���κ�����,��Ĭ�ϴ򿪸����е�ȫ��ͼ��
        strSerials = ",'" & Mid(lvwSeq.SelectedItem.Key, 2) & "'"
        strImages = Mid(lvwSeq.SelectedItem.Key, 2) & "|ȫ��"
    End If
    
    strSerials = Mid(strSerials, 2)
    
    intImageInverval = Val(Me.cbrMain.FindControl(, conMenu_Manage_ImageInterval, , True).Text)

    OpenViewer pobjPacsCore, mlngAdviceID, mblnAddImage, Me, strSerials, mblnMoved, mblnLocalizerBackward, intImageInverval, strImages
    Exit Sub
CallError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub DViewer_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim i As Integer
    If Button <> 1 Then Exit Sub
    
    With DViewer
        i = .ImageIndex(x, y)
        If i > 0 And i <= .Images.Count And i <> iCurImageIndex Then
            .Images(iCurImageIndex).BorderColour = vbWhite
            .Images(i).BorderColour = vbRed
            iCurImageIndex = i
        End If
    End With
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If Me.Tag = "Loading" Then
        Me.Tag = ""
        
    End If
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
            cbrControl.IconId = 825: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "��ʾ��ǰ����Ӱ������ͼ"
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
            cbrControl.IconId = 791: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "ˢ�µ�ǰ����ͼ������": cbrControl.Flags = xtpFlagRightAlign
    End With
        
    Call subSetMenuState
       
    With dkpMain
        .SetCommandBars Me.cbrMain
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.ThemedFloatingFrames = False
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
        Set Pane1 = .CreatePane(1, 0, 300, DockTopOf, Nothing)
            Pane1.Handle = lvwSeq.Hwnd
            Pane1.Options = PaneNoCaption Or PaneNoCloseable
            
        Set Pane1 = .CreatePane(2, 0, 300, DockBottomOf, Pane1)
            Pane1.Handle = lvwImage.Hwnd
            Pane1.Options = PaneNoCaption Or PaneNoCloseable
            
        Set Pane1 = .CreatePane(3, 0, 400, DockBottomOf, Nothing)
            Pane1.Handle = picView.Hwnd
            Pane1.Options = PaneNoCaption Or PaneNoCloseable
    End With
    dkpMain.LoadStateFromString GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, "")
    Call RestoreWinState(Me, App.ProductName)
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
        With .ColumnHeaders
            .Clear
            .Add , , "Ӱ�����", 2000
            .Add , , "��ͼ��", 2000
            .Add , , "����", 800, 1
            .Add , , "���к�", 800, 1
            .Add , , "ͼ����", 800, 1
            .Add , , "˵��", 2500
            .Add , , "�ɼ�ʱ��", 1800
        End With
        .ListItems.Add , , "Temp"
        .ListItems.Clear
    End With
    
    strSQL = "Select A.����UID,A.���к�,A.��������,A.�ɼ�ʱ��,B.Ӱ�����,B.����," & _
        " B.���UID,Sum(1) As ͼ���� " & _
        "From Ӱ�������� A,Ӱ�����¼ B,Ӱ����ͼ�� D " & _
        "Where B.ҽ��ID= [1]  And B.���ͺ�= [2] And A.���UID=B.���UID  And A.����UID=D.����UID " & _
        "Group By A.����UID,A.���к�,A.��������,A.�ɼ�ʱ��,B.Ӱ�����,B.����,B.���UID " & _
        "Order By B.Ӱ�����,B.����,A.���к�"
    If mblnMoved Then
        strSQL = Replace(strSQL, "Ӱ��������", "HӰ��������")
        strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
        strSQL = Replace(strSQL, "Ӱ����ͼ��", "HӰ����ͼ��")
    End If
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngAdviceID, mlngSendNo)
   
    lvwSeq.Tag = ""
    If Not rsTmp.EOF Then
        lvwSeq.Tag = Nvl(rsTmp("���UID"))
        Do While Not rsTmp.EOF
            
            Set tmpItem = lvwSeq.ListItems.Add(, "_" & rsTmp("����UID"), rsTmp("Ӱ�����"))
            With tmpItem
                If mintSelectAllImg = 0 Or mintSelectAllImg = 1 Then
                    .SubItems(1) = "ȫ��"
                Else
                    .SubItems(1) = ""
                End If
                
                .SubItems(2) = Nvl(rsTmp("����"))
                .SubItems(3) = Nvl(rsTmp("���к�"))
                .SubItems(4) = Nvl(rsTmp("ͼ����"), 0)
                .SubItems(5) = Nvl(rsTmp("��������"))
                .SubItems(6) = Nvl(rsTmp("�ɼ�ʱ��"), Date)
                
                If .Key = strCurKey Then .Selected = True
            End With
            rsTmp.MoveNext
        Loop
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
    If mblnMoved Then
        strSQL = Replace(strSQL, "Ӱ����ͼ��", "HӰ����ͼ��")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡͼ����Ϣ", strSeriesUID)
    
    lvwImage.Tag = ""
    If Not rsTmp.EOF Then
        lvwImage.Tag = strSeriesUID
        Do While Not rsTmp.EOF
            Set tmpItem = lvwImage.ListItems.Add(, rsTmp("ͼ��UID"), rsTmp("ͼ���"))
            With tmpItem
                .SubItems(1) = Nvl(rsTmp("ͼ������"))
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
    Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strRegPath As String
    
    strRegPath = "����ģ��\" & App.ProductName & "\frmPacsImg"
    SaveSetting "ZLSOFT", strRegPath, "SelectAllSeq", mintSelectAllSeq
    SaveSetting "ZLSOFT", strRegPath, "SelectAllImg", mintSelectAllImg
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lvwImage_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call WriteSelectdImages(lvwImage.Tag)
End Sub

Private Sub lvwImage_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    If Item.Checked <> Item.Selected Then
        Item.Checked = Item.Selected
        Call WriteSelectdImages(lvwImage.Tag)
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
    GetAllImages DViewer, mblnMoved, 3, 0, lvwImage.Tag, 1, 1, False, "", strImageUID

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

Private Sub lvwSeq_DblClick()
    If lvwSeq.SelectedItem Is Nothing Then Exit Sub
    DViewer_DblClick
End Sub

Private Sub lvwSeq_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    lvwSeq.SelectedItem = Item
    Call ShowImageList(Item)
End Sub

Private Sub lvwSeq_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked <> Item.Selected Then
        Item.Checked = Item.Selected
    End If
    Call ShowImageList(Item)
End Sub

Private Sub picView_Resize()
    Dim iCols As Integer, iRows As Integer
    
    On Error Resume Next
    With DViewer
        .Left = 0: .Top = 0
        .Width = picView.ScaleWidth: .Height = picView.ScaleHeight
        
        If .Images.Count > 0 Then
            ResizeRegion .Images.Count, .Width, .Height, iRows, iCols
            .MultiColumns = iCols: .MultiRows = iRows
        End If
    End With
End Sub

Public Function ZLfun3DImgProcess() As String
'------------------------------------------------
'���ܣ���ά�ؽ�Ԥ�����ƶ���ǰ��ѡ�����е�ͼ��
'��������
'���أ�ͼ���ƶ���Ŀ��Ŀ¼������ƶ�ʧ���򷵻ؿ�
'------------------------------------------------

    Dim strSeriesUID As String
    Dim Item As MSComctlLib.ListItem
    Dim iSeriesCount As Integer
    
    On Error GoTo CallError
    If lvwSeq.SelectedItem Is Nothing Then
        MsgBox "��ѡ��һ�����н�����ά�ؽ���"
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
        MsgBox "��ѡ��һ�����н�����ά�ؽ���ÿ���ؽ�ֻ��ѡ��һ��ϵ�С�"
        ZLfun3DImgProcess = ""
        Exit Function
    End If
    
    '�ƶ�ָ������UID��ͼ��
    ZLfun3DImgProcess = funMove3DImage(strSeriesUID, mblnMoved)
    Exit Function
CallError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ZLfun3DImgProcess = ""
End Function

Private Function funMove3DImage(strSeriesUID As String, blnMoved As Boolean) As String
'------------------------------------------------
'���ܣ���һ�����е�ͼ���ƶ���3D��ʱĿ¼�У��ȴ���ά�ؽ�����ĵ���
'������
'       lngAdviceID --  ҽ��ID
'       strSeriesUID -- ͼ�������UID
'       blnMoved -- ͼ���Ƿ�ת��
'���أ�ͼ���ƶ���Ŀ��Ŀ¼������ƶ�ʧ���򷵻ؿ�
'------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    Dim str3DCachePath As String
    Dim strTmpFile As String
    Dim strImageFullPath As String
    
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
        str3DCachePath = App.Path & "\TmpImage\3D\" & Replace(Nvl(rsTmp("ͼ��Ŀ¼")), "/", "\") & "\" & strSeriesUID & "\"
        strImageFullPath = App.Path & "\TmpImage\" & Replace(Nvl(rsTmp("ͼ��Ŀ¼")), "/", "\") & "\"
        MkLocalDir str3DCachePath

        On Error GoTo DBError
        
        Do While Not rsTmp.EOF
            '���3DĿ¼��û��ͼ���ټ�鱾�ػ���Ŀ¼������ٴ�FTP����ͼ��
            strTmpFile = str3DCachePath & Nvl(rsTmp("ͼ��UID"))
            If Dir(strTmpFile) = vbNullString Then  '��ͼ������Ҫ���κβ���
                If Dir(strImageFullPath & Nvl(rsTmp("ͼ��UID"))) = vbNullString Then
                    '���ػ���ͼ�񲻴��ڣ����ȡFTPͼ��
                    '����FTP����
                    If rsTmp("�豸��1") <> vbNullString And Inet1.hConnection = 0 Then
                        If Inet1.FuncFtpConnect(Nvl(rsTmp("Host1")), Nvl(rsTmp("User1")), Nvl(rsTmp("Pwd1"))) = 0 Then
                            If rsTmp("�豸��2") <> vbNullString Then
                                If Inet2.FuncFtpConnect(Nvl(rsTmp("Host2")), Nvl(rsTmp("User2")), Nvl(rsTmp("Pwd2"))) = 0 Then
                                    MsgBox "FTP�����������ӣ������������á�"
                                    funMove3DImage = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                    If Inet1.FuncDownloadFile(Nvl(rsTmp("Root1")) & rsTmp("ͼ��Ŀ¼"), strTmpFile, rsTmp("ͼ��UID")) <> 0 Then
                        '���豸��1��ȡͼ��ʧ�ܣ�����豸��2��ȡͼ��
                        If rsTmp("�豸��2") <> vbNullString Then
                            If Inet2.hConnection = 0 Then Inet2.FuncFtpConnect Nvl(rsTmp("Host2")), Nvl(rsTmp("User2")), Nvl(rsTmp("Pwd2"))
                            Call Inet2.FuncDownloadFile(Nvl(rsTmp("Root2")) & rsTmp("ͼ��Ŀ¼"), strTmpFile, rsTmp("ͼ��UID"))
                        End If
                    End If
                Else
                '���ع�Ƭ������ͼ����ڣ�ֱ�Ӹ��Ƶ�3DĿ¼
                    FileCopy strImageFullPath & Nvl(rsTmp("ͼ��UID")), strTmpFile
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
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    funMove3DImage = ""
End Function

Private Sub ShowSeqImg()
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
