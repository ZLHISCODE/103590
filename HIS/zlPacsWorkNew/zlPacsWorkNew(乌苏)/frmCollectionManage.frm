VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmCollectionManage 
   Caption         =   "�ղع���"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   345
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCollectionManage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   12000
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   120
      ScaleHeight     =   5775
      ScaleWidth      =   12015
      TabIndex        =   0
      Top             =   1080
      Width           =   12015
      Begin zl9PacsControl.ucSplitter ucSplitter 
         Height          =   5775
         Left            =   3255
         TabIndex        =   4
         Top             =   0
         Width           =   90
         _ExtentX        =   159
         _ExtentY        =   10186
         BackColor       =   -2147483644
         SplitWidth      =   90
         SplitLevel      =   3
         Con1MinSize     =   2250
         Con2MinSize     =   2430
         Control1Name    =   "PicTvw"
         Control2Name    =   "PicData"
      End
      Begin VB.PictureBox PicData 
         BorderStyle     =   0  'None
         Height          =   5775
         Left            =   3345
         ScaleHeight     =   5775
         ScaleWidth      =   8670
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   8670
         Begin VSFlex8Ctl.VSFlexGrid vfgCollectionData 
            Height          =   4455
            Left            =   1560
            TabIndex        =   3
            Top             =   600
            Width           =   5535
            _cx             =   9763
            _cy             =   7858
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   3
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   7
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
      End
      Begin VB.PictureBox PicTvw 
         BorderStyle     =   0  'None
         Height          =   5775
         Left            =   0
         ScaleHeight     =   5775
         ScaleWidth      =   3255
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   3255
         Begin MSComctlLib.ImageList imgList 
            Left            =   2640
            Top             =   120
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   3
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCollectionManage.frx":6852
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCollectionManage.frx":6BEC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCollectionManage.frx":6F86
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.TreeView tvwCollectionType 
            Height          =   5295
            Left            =   240
            TabIndex        =   6
            Top             =   1320
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   9340
            _Version        =   393217
            Indentation     =   494
            LabelEdit       =   1
            Style           =   7
            ImageList       =   "imgList"
            Appearance      =   1
         End
         Begin VB.Image imgTree 
            Height          =   720
            Index           =   3
            Left            =   2040
            Picture         =   "frmCollectionManage.frx":7320
            Top             =   1680
            Width           =   720
         End
         Begin VB.Image imgTree 
            Height          =   720
            Index           =   1
            Left            =   480
            Picture         =   "frmCollectionManage.frx":DB72
            Top             =   1680
            Width           =   720
         End
         Begin VB.Image imgTree 
            Height          =   720
            Index           =   2
            Left            =   1320
            Picture         =   "frmCollectionManage.frx":143C4
            Top             =   1680
            Width           =   720
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   7560
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCollectionManage.frx":1AC16
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14288
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
   Begin XtremeCommandBars.ImageManager imgPopup 
      Left            =   1320
      Top             =   360
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmCollectionManage.frx":1B4AA
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   480
      Top             =   360
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCollectionManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strSql As String
Private mstrNodeKey As String
Private mstrNodeName As String
Private mobjNode As Node
Private mrsTvwData As ADODB.Recordset
Private mobjSourceNode As Object
'�˵�
Private Enum popMenus
    conMenu_Edit_Add = 100
    conMenu_Edit_Rename = 101
    conMenu_Edit_Del = 102
    conMenu_Edit_DelColl = 103
    conMenu_Edit_Share = 104
End Enum

Public Sub ShowCollectionManageWind(Optional owner As Form = Nothing)
'��ʾ�ղع�����
    Call Me.Show(1, owner)
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo ErrorHand
    
    Select Case control.ID
        Case conMenu_Edit_Add
            Call Menu_Edit_Add
            
        Case conMenu_Edit_Rename
            Call Menu_Edit_Rename
            
        Case conMenu_Edit_Del
            Call Menu_Edit_Del
            
        Case conMenu_Edit_DelColl
            Call Menu_Edit_DelColl
            
        Case conMenu_Edit_Share
            Call Menu_Edit_Share(control)

        Case conMenu_File_Exit
            Call Menu_File_Exit
            
        Case conMenu_View_Refresh
            Call LoadTreeView
'---------------------------�鿴----------------
        Case conMenu_View_ToolBar_Button '������
            Call Menu_View_ToolBar_Button_click(control)

        Case conMenu_View_ToolBar_Text '��ť����
            Call Menu_View_ToolBar_Text_click(control)

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
    End Select
    Call Form_Resize
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

Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo ErrorHand
    If tvwCollectionType.SelectedItem Is Nothing Then Exit Sub
    control.Enabled = True
    Select Case control.ID
        Case conMenu_Edit_Add
            
            
        Case conMenu_Edit_Rename
            If tvwCollectionType.SelectedItem.Text = "�ղ����" Then control.Enabled = False
            
        Case conMenu_Edit_Del
            If tvwCollectionType.SelectedItem.Text = "�ղ����" Then control.Enabled = False
            
        Case conMenu_Edit_DelColl
            
            
        Case conMenu_Edit_Share
            If tvwCollectionType.SelectedItem.Text = "�ղ����" Then
                control.Enabled = False
                Exit Sub
            End If
            
            control.Caption = IIf(tvwCollectionType.SelectedItem.Tag = 3, "ȡ������", "���ù���")
            If control.Caption = "ȡ������" Then control.ToolTipText = "ȡ�����ղ�Ϊ����"
        Case conMenu_File_Exit
            
        
    End Select
    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo errHand
    
    InitCommandBars
    '����TreeView����
    Call LoadTreeView
     
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub InitVfgData()
    Dim i As Integer
'��ʼ�����ݿؼ�
    With vfgCollectionData
        .Clear
        .FixedRows = 1
        .Cols = 13
        .ColWidth(0) = 500
        .ColWidth(1) = 900
        .ColWidth(2) = 900
        .ColWidth(3) = 900
        .ColWidth(4) = 1000
        .ColWidth(5) = 700
        .ColWidth(6) = 700
        .ColWidth(7) = 2500
        .ColWidth(8) = 3000
        .ColWidth(9) = 1000
        .ColWidth(10) = 1200
        .ColWidth(11) = 1200
        .TextMatrix(0, 1) = "����"
        .TextMatrix(0, 2) = "�����"
        .TextMatrix(0, 3) = "סԺ��"
        .TextMatrix(0, 4) = "����"
        .TextMatrix(0, 5) = "�Ա�"
        .TextMatrix(0, 6) = "����"
        .TextMatrix(0, 7) = "ҽ������"
        .TextMatrix(0, 8) = "��λ����"
        .TextMatrix(0, 9) = "����ҽ��"
        .TextMatrix(0, 10) = "����ʱ��"
        .TextMatrix(0, 11) = "�ղ�ʱ��"
        .TextMatrix(0, 12) = "����ID"
        For i = 0 To 12
            .ColAlignment(i) = flexAlignLeftCenter
        Next
        .AllowSelection = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .ColHidden(12) = True
        .AllowUserResizing = flexResizeColumns
    End With

End Sub

Private Sub LoadTreeView()
'����TreeView���ݷ���
    Dim strCurrKey As String
    Dim rsTemp As ADODB.Recordset
    Dim dtServicesTime As Date
    Dim strSql As String
err = 0: On Error GoTo errHand

    strSql = "select id from Ӱ���ղ���� where �ղ����='�ղ����' "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    '������ݿ���û�ж����ڵ����ݣ�����붥���ڵ�����
    If rsTemp.RecordCount <= 0 Then
         '��ǰ������ʱ��
        dtServicesTime = zlDatabase.Currentdate
        
        strSql = "select Zl_Ӱ���ղ����_����([1],[2],[3],[4],[5]) as ����ֵ from dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                "", _
                                "�ղ����", _
                                0, _
                                "", _
                                dtServicesTime)
    
    End If

    strSql = "select ID,�ϼ�ID,�ղ����,�Ƿ��� from Ӱ���ղ���� where ������= '" & UserInfo.���� & "' or ������ is null Start With �ϼ�id Is Null Connect By Prior ID = �ϼ�id"
    Set mrsTvwData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    With mrsTvwData
        Me.tvwCollectionType.Nodes.Clear
        
        If Not tvwCollectionType.SelectedItem Is Nothing Then strCurrKey = tvwCollectionType.SelectedItem.Key
        Do While Not .EOF
            If IsNull(!�ϼ�ID) Then
                Set mobjNode = Me.tvwCollectionType.Nodes.Add(, , "_" & Nvl(!ID), Nvl(!�ղ����), IIf(!�Ƿ��� = 0, 1, 3), IIf(Nvl(!�Ƿ���) = 0, 2, 3))
                mobjNode.Tag = IIf(Nvl(!�Ƿ���) = 0, 2, 3)
            Else
                Set mobjNode = Me.tvwCollectionType.Nodes.Add("_" & Nvl(!�ϼ�ID), tvwChild, "_" & Nvl(!ID), Nvl(!�ղ����), IIf(Nvl(!�Ƿ���) = 0, 1, 3), IIf(Nvl(!�Ƿ���) = 0, 2, 3))
                mobjNode.Tag = IIf(Nvl(!�Ƿ���) = 0, 2, 3)
            End If
            
            mobjNode.Sorted = True
            mobjNode.Expanded = True
            If strCurrKey = mobjNode.Key Then mobjNode.Selected = True
            .MoveNext
        Loop
    End With
    
    '�������ʱ�Զ�ѡ������� �Ҳ��ղص�����
err = 0: On Error GoTo 0
    If Me.tvwCollectionType.Nodes.Count > 0 Then
        If tvwCollectionType.SelectedItem Is Nothing Then Me.tvwCollectionType.Nodes(1).Selected = True
        tvwCollectionType.SelectedItem.EnsureVisible
        Call tvwCollectionType_NodeClick(Me.tvwCollectionType.SelectedItem)
    End If

    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub tvwCollectionType_AfterLabelEdit(Cancel As Integer, NewString As String)
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errHand
    tvwCollectionType.DragIcon = imgTree(1).Picture
    If NewString = "" Then NewString = tvwCollectionType.SelectedItem.Text
    If tvwCollectionType.SelectedItem.Text = NewString Then Exit Sub
    '�ж��޸������Ƿ��ظ�
    strSql = "select �ղ���� from Ӱ���ղ���� where ������= [1] and �ղ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.����, NewString)
    
    If rsTemp.RecordCount > 0 Then
        Call MsgBoxD(Me, "�ղ������ظ���", vbOKOnly, Me.Caption)
        '�ղ������ظ�����ԭ��������
        NewString = tvwCollectionType.SelectedItem.Text
        tvwCollectionType.SelectedItem.Selected = True
        Exit Sub
    End If
    
    strSql = "Zl_Ӱ���ղ����_����(" & Val(Mid(tvwCollectionType.SelectedItem.Key, 2)) & ",'" & _
                        Decode(Trim(NewString), "", tvwCollectionType.SelectedItem.Text, Trim(NewString)) & "',2)"

    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tvwCollectionType_DragDrop(Source As control, X As Single, Y As Single)
    Dim strCurParent As String
    Dim lngNodesKey As Long
    
On Error Resume Next
    If mobjSourceNode Is Nothing Then Exit Sub
    strCurParent = mobjSourceNode.Parent.Text
    If strCurParent = "" Then Exit Sub
    If Not (tvwCollectionType.DropHighlight Is Nothing) Then
        Set mobjSourceNode.Parent = tvwCollectionType.DropHighlight
        Set tvwCollectionType.DropHighlight = Nothing
        
        If strCurParent <> mobjSourceNode.Parent.Text Then
            strSql = "Zl_Ӱ���ղ����_���·���(" & Mid(mobjSourceNode.Key, 2) & _
                                          "," & Val(Mid(tvwCollectionType.SelectedItem.Parent.Key, 2)) & ")"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        End If
    End If
    
    Set mobjSourceNode = Nothing
    
End Sub

Private Sub tvwCollectionType_DragOver(Source As control, X As Single, Y As Single, State As Integer)
    Dim objNode As Node
    Dim objTargetNode As Object
    
    On Error GoTo errHand

    If mobjSourceNode Is Nothing Then Exit Sub
    
    Set objNode = tvwCollectionType.HitTest(X, Y)
    
    If objNode Is objTargetNode Then Exit Sub
    Set objTargetNode = objNode
    
    Set tvwCollectionType.DropHighlight = objTargetNode
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub tvwCollectionType_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHand
    If KeyCode = vbKeyF2 Then
        If tvwCollectionType.SelectedItem.Text <> "�ղ����" Then tvwCollectionType.StartLabelEdit
    ElseIf KeyCode = vbKeyF5 Then
        Call LoadTreeView
    End If
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub tvwCollectionType_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo errHand
    If tvwCollectionType.HitTest(X, Y) Is Nothing Then Exit Sub
    Set mobjSourceNode = tvwCollectionType.HitTest(X, Y)
    tvwCollectionType.SelectedItem = tvwCollectionType.HitTest(X, Y)
    If mobjSourceNode.Text = "�ղ����" Then Set mobjSourceNode = Nothing
    'ˢ�°�ť״̬
    cbrMain.RecalcLayout
    If tvwCollectionType.HitTest(X, Y).Text <> "�ղ����" Then tvwCollectionType_NodeClick mobjSourceNode
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub tvwCollectionType_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If mobjSourceNode Is Nothing Then Exit Sub
    
    tvwCollectionType.DragIcon = IIf(tvwCollectionType.SelectedItem.Tag = 2, imgTree(2).Picture, imgTree(3).Picture)
    If Button = vbLeftButton Then
        Set tvwCollectionType.SelectedItem = mobjSourceNode
        tvwCollectionType.Drag vbBeginDrag
    End If
    
    Exit Sub
End Sub

Private Sub tvwCollectionType_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'�����Ҽ��˵�
On Error GoTo ErrHandle
    If Button = 2 Then
        Dim objPopup As CommandBar
        Dim objControl As CommandBarControl

        Set cbrMain.Icons = imgPopup.Icons
        Set objPopup = cbrMain.Add("�Ҽ��˵�", xtpBarPopup)
        With objPopup.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Add, "�������(&A)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Rename, "������(&U)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Del, "ɾ�����(&D)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Share, "���ù���(&S)")
        End With
        objPopup.ShowPopup
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub tvwCollectionType_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo errHand
'����ѡ�нڵ��µ�����
    Dim rsCollectionData As ADODB.Recordset
    Dim rsIsShare As ADODB.Recordset
    Dim i As Long
    Dim strSql As String
    Dim strAdviceTemp As String
    
    On Error GoTo errHand
    Set mobjNode = Node
    '�õ��ڵ��Key
    mstrNodeKey = Mid(Node.Key, 2)
    
    strSql = "select distinct e.id,nvl(c.����,d.����) ����,nvl(c.����,d.����) ����,nvl(c.�Ա�,d.�Ա�) �Ա�,a.ҽ������,a.����ҽ��,a.����ʱ��,c.�����,c.סԺ��,d.����,f.�Ƿ���,e.�ղ�ʱ�� " & _
            " from ����ҽ����¼ a,����ҽ������ b,������Ϣ c,Ӱ�����¼ d,Ӱ���ղ����� e,Ӱ���ղ���� f" & _
            " where a.id = b.ҽ��id and b.ҽ��ID=d.ҽ��ID(+)" & _
            " and a.����ID=c.����ID and a.���id is null" & _
            " and b.ҽ��id = e.ҽ��id and e.�ղ�id = f.id and f.������='" & UserInfo.���� & _
            "' and f.id in (select distinct id from Ӱ���ղ���� start with id = " & mstrNodeKey & " connect by prior id=�ϼ�id) order by e.id"

    Set rsCollectionData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)

    With vfgCollectionData
        .Clear
        
        If rsCollectionData.RecordCount = 0 Then
            .Rows = 1
        Else
            .Rows = rsCollectionData.RecordCount + 1
        End If
        
        '��ʼ��������ʾ�ؼ�
        Call InitVfgData
        
        For i = 1 To rsCollectionData.RecordCount
        
            strAdviceTemp = Nvl(rsCollectionData!ҽ������)
            
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Nvl(rsCollectionData!����)
            .TextMatrix(i, 2) = Nvl(rsCollectionData!�����)
            .TextMatrix(i, 3) = Nvl(rsCollectionData!סԺ��)
            .TextMatrix(i, 4) = Nvl(rsCollectionData!����)
            .TextMatrix(i, 5) = Nvl(rsCollectionData!�Ա�)
            .TextMatrix(i, 6) = Nvl(rsCollectionData!����)
            
            '����ҽ�����ݵ�ҽ�����ֺͲ�λ��������
            .TextMatrix(i, 7) = Mid(strAdviceTemp, 1, InStr(strAdviceTemp, ":") - 1)
            .TextMatrix(i, 8) = Mid(strAdviceTemp, InStr(strAdviceTemp, ":") + 1, Len(strAdviceTemp))
            
            .TextMatrix(i, 9) = Nvl(rsCollectionData!����ҽ��)
            .TextMatrix(i, 10) = Format(Nvl(rsCollectionData!����ʱ��), "yyyy-mm-dd")
            .TextMatrix(i, 11) = Format(Nvl(rsCollectionData!�ղ�ʱ��), "yyyy-mm-dd")
            .TextMatrix(i, 12) = Nvl(rsCollectionData!ID)

            If Not rsCollectionData.EOF Then rsCollectionData.MoveNext
        Next
    End With
    
    stbThis.Panels(2).Text = "��ǰ�ղ�������� " & rsCollectionData.RecordCount & " ���ղ�"
    
    cbrMain.RecalcLayout
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Menu_Edit_Add()
'�����ղ�����
On Error GoTo errHand
    Dim lngNodesKey As Long
    Dim rsTemp As ADODB.Recordset
    Dim dtServicesTime As Date
    Dim strSql As String
      
    '��ǰ������ʱ��
    dtServicesTime = zlDatabase.Currentdate

    strSql = "select Zl_Ӱ���ղ����_����([1],[2],[3],[4],[5]) as ����ֵ from dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                Val(Mid(tvwCollectionType.SelectedItem.Key, 2)), _
                                "�½��ڵ�" & GetNextIndex, _
                                0, _
                                UserInfo.����, _
                                dtServicesTime)

    If rsTemp.RecordCount > 0 Then lngNodesKey = Nvl(rsTemp!����ֵ)
    
    '��treeView�ؼ�����������ڵ�
    Set mobjNode = Me.tvwCollectionType.Nodes.Add("_" & Val(Mid(tvwCollectionType.SelectedItem.Key, 2)), tvwChild, "_" & lngNodesKey, "�½��ڵ�" & GetNextIndex, 1, 2)
    mobjNode.Selected = True
    mobjNode.Tag = 2
    tvwCollectionType.StartLabelEdit
                
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetNextIndex() As String
    Dim i                   As Integer
    Dim j                   As Integer
    Dim strName()           As String
    Dim strTemp             As String

    On Error GoTo ErrorHand
    
    mstrNodeName = ""
    Call GetAllNode(mobjNode.Root)
    strName() = Split(mstrNodeName, "|")
    For i = 0 To UBound(strName()) - 1
        For j = i + 1 To UBound(strName()) - 1
            If CInt(Mid(strName(i), 5)) > CInt(Mid(strName(j), 5)) Then
                strTemp = strName(i)
                strName(i) = strName(j)
                strName(j) = strTemp
            End If
        Next
    Next
    For i = 0 To UBound(strName()) - 1
        If "�½��ڵ�" & i + 1 <> strName(i) Then
            GetNextIndex = i + 1
            Exit Function
        End If
    Next
    GetNextIndex = i + 1
    Exit Function
ErrorHand:
    GetNextIndex = i + 1
End Function

Private Sub GetAllNode(ByVal Node As MSComctlLib.Node)
    Dim objNode As Node
    
    If Node.Children > 0 Then
        Set objNode = Node.Child
        Do While Not objNode Is Nothing
            If InStr(objNode.Text, "�½��ڵ�") > 0 Then
                If objNode.Text <> "�½��ڵ�" Then mstrNodeName = mstrNodeName & objNode.Text & "|"
            End If
            Call GetAllNode(objNode)
            Set objNode = objNode.Next
        Loop
    End If
End Sub

Private Sub Menu_Edit_Del()
'ɾ���ղ�����(����ɾ��)
On Error GoTo errHand
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset

    If tvwCollectionType.SelectedItem.Children <> 0 Then
        Call MsgBoxD(Me, "���������������ͣ�����ɾ����", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If vfgCollectionData.Rows > 1 Then
        If MsgBoxD(Me, "ȷ��ɾ����(ɾ�������ͻ�ɾ���ղ���Ϣ)", vbOKCancel, Me.Caption) = 2 Then Exit Sub
    End If
    
    strSql = "Zl_Ӱ���ղ����_ɾ��(" & Val(Mid(tvwCollectionType.SelectedItem.Key, 2)) & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    'ˢ��
    tvwCollectionType_NodeClick tvwCollectionType.SelectedItem.Parent
    '��TreeView�ؼ���ɾ��ѡ�нڵ�
    tvwCollectionType.Nodes.Remove (tvwCollectionType.SelectedItem.Key)
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Menu_Edit_Rename()
'�����ղ�����
On Error GoTo errHand
Dim strSql As String
Dim strCurNodeText As String
    
    Set mobjNode = Me.tvwCollectionType.SelectedItem
    strCurNodeText = mobjNode.Text
    mobjNode.Selected = True
    tvwCollectionType.DragIcon = imgTree(1).Picture
    tvwCollectionType.StartLabelEdit
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Edit_Share(ByVal control As XtremeCommandBars.ICommandBarControl)
'�����ղع���״̬
On Error GoTo errHand
Dim strSql As String

    'ֻ���¹���״̬
    strSql = "Zl_Ӱ���ղ����_����(" & Val(Mid(tvwCollectionType.SelectedItem.Key, 2)) & ",null," & IIf(control.Caption = "ȡ������", 0, 1) & ")"

    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    tvwCollectionType.SelectedItem.SelectedImage = IIf(control.Caption = "ȡ������", 2, 3)
    tvwCollectionType.SelectedItem.Image = IIf(control.Caption = "ȡ������", 1, 3)
    tvwCollectionType.SelectedItem.Tag = IIf(control.Caption = "ȡ������", 2, 3)
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub Menu_Edit_DelColl()
'ɾ���ղؼ�����
On Error GoTo errHand
Dim strSql As String
Dim i As Integer

    If Me.vfgCollectionData.SelectedRows = 0 Then
       Call MsgBoxD(Me, "����ѡ��Ҫɾ�����ղ����ݡ�", vbOKOnly, Me.Caption)
       Exit Sub
    End If
    
    If MsgBoxD(Me, "��ȷ��Ҫɾ����ѡ����ղ�������", vbOKCancel, Me.Caption) = 2 Then Exit Sub
    With vfgCollectionData
        For i = 0 To .SelectedRows - 1
            strSql = "Zl_Ӱ���ղ�����_ɾ��(" & .TextMatrix(.SelectedRow(0), 12) & ")"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            'ɾ��VSFlexGrid����
            vfgCollectionData.RemoveItem (vfgCollectionData.SelectedRow(0))
        Next
    End With
     
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
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
        .VisualTheme = xtpThemeOffice2003                      '���ÿؼ���ʾ���
        .EnableCustomization False                              '�Ƿ������Զ�������
        Set .Icons = imgPopup.Icons                           '���ù�����ͼ��ؼ�
    End With

    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '�˵�����
'Begin------------------------�༭�˵�--------------------------------------Ĭ�Ͽɼ�
    cbrMain.ActiveMenuBar.Title = "�˵�"
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)")
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_File_Exit, "�˳�(&Q)")
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)")
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Add, "�������(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Rename, "������(&U)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Del, "ɾ�����(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_DelColl, "ɾ���ղ�(&M)")
        cbrControl.BeginGroup = True
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
        Set cbrControl = .Controls.Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(R)")
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
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, conMenu_Edit_Add, "�������", "�������")
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, conMenu_Edit_Rename, "������", "������")
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, conMenu_Edit_Del, "ɾ�����", "ɾ�����")
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, conMenu_Edit_Share, "���ù���")
    cbrControl.ToolTipText = "�����ղ�����Ϊ����"
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, conMenu_Edit_DelColl, "ɾ���ղ�", "ɾ���ղ�")
    cbrControl.BeginGroup = True
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, conMenu_View_Refresh, "ˢ��", "ˢ��")
    cbrControl.BeginGroup = True
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, conMenu_File_Exit, "�˳�", "�˳�")
    cbrControl.BeginGroup = True
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    
    cbrMain.GetClientRect lngLeft, lngTop, lngRight, lngBottom
    
    picMain.Left = 0
    picMain.Top = lngTop
    picMain.Width = Me.ScaleWidth
    If stbThis.Visible = True Then
        picMain.Height = Me.ScaleHeight - picMain.Top - stbThis.Height
    Else
        picMain.Height = Me.ScaleHeight - picMain.Top
    End If
    
    '����ı�,�����û��ؼ��Զ���������
    ucSplitter.RePaint
End Sub

Private Sub PicTvw_Resize()
    On Error Resume Next
    
    tvwCollectionType.Top = 0
    tvwCollectionType.Left = 60
    tvwCollectionType.Height = PicTvw.Height
    tvwCollectionType.Width = PicTvw.Width - 60
End Sub

Private Sub PicData_Resize()
    On Error Resume Next
    
    vfgCollectionData.Top = 0
    vfgCollectionData.Left = 0
    vfgCollectionData.Height = PicData.Height
    vfgCollectionData.Width = PicData.Width - 60
End Sub

Private Sub vfgCollectionData_Click()
    On Error GoTo errHand
    tvwCollectionType.HideSelection = False
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
End Sub




