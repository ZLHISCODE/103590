VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmPatholMeal 
   Caption         =   "�ײ�ά��"
   ClientHeight    =   8190
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   11295
   Icon            =   "frmPatholMeal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   11295
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picDatas 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   240
      ScaleHeight     =   3735
      ScaleWidth      =   10455
      TabIndex        =   6
      Top             =   360
      Width           =   10455
      Begin VB.Frame framMeals 
         Caption         =   "�ײͼ�¼"
         Height          =   3495
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   10215
         Begin VB.ComboBox cboMealClass 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   320
            Width           =   1935
         End
         Begin zl9PACSWork.ucFlexGrid ufgMeal 
            Height          =   2415
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   4260
            IsKeepRows      =   0   'False
            DisCellColor    =   16777215
            IsCopyAdoMode   =   0   'False
            IsEjectConfig   =   -1  'True
            Editable        =   0
            HeadFontCharset =   134
            HeadFontWeight  =   400
            DataFontCharset =   134
            DataFontWeight  =   400
         End
         Begin VB.Label lblMealClass 
            Caption         =   "�ײ����"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   975
         End
      End
   End
   Begin VB.PictureBox picMealLink 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   240
      ScaleHeight     =   3735
      ScaleWidth      =   10455
      TabIndex        =   0
      Top             =   4080
      Width           =   10455
      Begin VB.Frame framAntibody 
         Caption         =   "������ϸ"
         Height          =   3495
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   10215
         Begin VB.CheckBox chkRowFilter 
            Caption         =   "ֻ��ʾ��ѡ��"
            Height          =   180
            Left            =   240
            TabIndex        =   3
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   3480
            TabIndex        =   2
            ToolTipText     =   "���ݿ������ƽ��п��ٶ�λ��"
            Top             =   300
            Width           =   1695
         End
         Begin zl9PACSWork.ucFlexGrid ufgMealLink 
            Height          =   2535
            Left            =   120
            TabIndex        =   4
            Top             =   840
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   4471
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
         Begin VB.Label lblFind 
            Caption         =   "���ٲ��ң�"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2400
            TabIndex        =   5
            ToolTipText     =   "���ݿ������ƽ��п��ٶ�λ��"
            Top             =   360
            Width           =   1095
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   11
      Top             =   7830
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatholMeal.frx":179A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13044
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
   Begin XtremeDockingPane.DockingPane dkpMeal 
      Bindings        =   "frmPatholMeal.frx":202E
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPatholMeal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrPrivs As String
Private mblnEdit As Boolean
Private mblnIsUpdate As Boolean
Private mblnCurModifyState As Boolean

Public Sub ShowMealWindow(ByVal strPrivs As String, owner As Form)
'��ʾ�ײ�ά������
    mstrPrivs = strPrivs
    
    Call ConfigPopedom
    
    Call Me.Show(1, owner)
End Sub


Private Sub ConfigPopedom()
'����Ȩ��
    Dim blnIsAllowMeal As Boolean
    
    blnIsAllowMeal = CheckPopedom(mstrPrivs, "�ײ�ά��")

    mblnCurModifyState = blnIsAllowMeal
    ufgMealLink.Enabled = False
    ufgMealLink.DataGrid.Enabled = True
    ufgMealLink.DataGrid.BackColor = &H8000000F
    mblnIsUpdate = False
End Sub

Private Sub InitMealList()
'��ʼ���ײ���ʾ�б�
    Dim strTemp As String
    
     '�ж����ݿ�������Ƿ������� �����ȡ���ݿ����  û�������Ĭ��
    strTemp = zlDatabase.GetPara("�����ײ��б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
    ufgMeal.DefaultColNames = gstrAntibodyMealCols
     
    If strTemp = "" Then
        ufgMeal.ColNames = gstrAntibodyMealCols
    Else
        ufgMeal.ColNames = strTemp
    End If
    
    ufgMeal.IsCopyMode = True
    '��ֹ�Ҽ������б����ô���
    ufgMeal.IsEjectConfig = False
    '��������
    ufgMeal.GridRows = glngStandardRowCount
    '�����и�
    ufgMeal.RowHeightMin = glngStandardRowHeight
    ufgMeal.ColConvertFormat = gstrAntibodyMealConvertFormat
    ufgMeal.IsShowPopupMenu = False
End Sub

Private Sub InitMealLinkList()
'��ʼ���ײ���ϸ�б�
    Dim strTemp As String
    
     '�ж����ݿ�������Ƿ������� �����ȡ���ݿ����  û�������Ĭ��
    strTemp = zlDatabase.GetPara("�ײ���ϸ�б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
    ufgMealLink.DefaultColNames = gstrAntibodyMealLinkCols
     
    If strTemp = "" Then
        ufgMealLink.ColNames = gstrAntibodyMealLinkCols
    Else
        ufgMealLink.ColNames = strTemp
    End If
    
    '��ֹ�Ҽ������б����ô���
    ufgMealLink.IsEjectConfig = False
      '��������
    ufgMealLink.GridRows = glngStandardRowCount
    '�����и�
    ufgMealLink.RowHeightMin = glngStandardRowHeight
    ufgMealLink.ColConvertFormat = gstrAntibodyMealLinkConvertFormat
    ufgMealLink.IsShowPopupMenu = False
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo ErrorHand
    
    Select Case control.ID
        Case conMenu_PatholMeal_Save              '����
            Call Menu_PatholMeal_Save

        Case conMenu_PatholMeal_Cancel            'ȡ��
            Call Menu_PatholMeal_Cancel

        Case conMenu_PatholMeal_AddRecord         '����
            Call Menu_PatholMeal_AddMeal
            
        Case conMenu_PatholMeal_ModRecord         '�޸�
            Call Menu_PatholMeal_ModMeal
            
        Case conMenu_PatholMeal_DelRecord         'ɾ��
            Call Menu_PatholMeal_DelRecord
            
        Case conMenu_PatholMeal_UpRow             '����
            Call Menu_PatholMeal_UpRow
            
        Case conMenu_PatholMeal_DownRow           '����
            Call Menu_PatholMeal_DownRow
            
        Case conMenu_File_Exit                    '�˳�
            Call Menu_File_Exit
        
        '---------------------------�鿴----------------
        Case conMenu_View_ToolBar_Button          '������
            Call Menu_View_ToolBar_Button_click(control)

        Case conMenu_View_ToolBar_Text            '��ť����
            Call Menu_View_ToolBar_Text_click(control)

        Case conMenu_View_StatusBar               '״̬��
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

Private Sub Menu_PatholMeal_Save()
    '����ǰ�����ô��ڱ༭״̬�ĵ�Ԫ��ʧȥ���㣬�����ⲻ�������ֵ
    ufgMeal.DataGrid.Col = 5
    ufgMeal.DataGrid.Row = ufgMeal.SelectionRow
    ufgMeal.DataGrid.SetFocus
    
    '����ײ������Ƿ�Ϊ��
    If ufgMeal.Text(ufgMeal.SelectionRow, gstrAntibodyMeal_�ײ�����) = "" Then
        MsgBoxD Me, "δ��ͨ����֤��ԭ�����ײ����Ʋ���Ϊ�գ�", vbExclamation, Me.Caption
        ufgMeal.LocateRow ufgMeal.SelectionRow
        ufgMeal.DataGrid.EditCell
        Exit Sub
    Else
        '����ײ������Ƿ��ظ�
        If Not CheckMealName Then Exit Sub
    End If
    
    Call Menu_PatholMeal_SaveMeal
    Call Menu_PatholMeal_SureSelected
    Call LoadMealClass

    ufgMeal.DataGrid.Cell(flexcpText, ufgMeal.SelectionRow, 5) = Format(ufgMeal.DataGrid.Cell(flexcpText, ufgMeal.SelectionRow, 5), "yyyy-mm-dd")

    If ufgMeal.AdoData.RecordCount > 0 Then Call ConfigMealLink(Val(ufgMeal.KeyValue(ufgMeal.SelectionRow)))
    
    mblnEdit = False
    mblnIsUpdate = False
    cboMealClass.Enabled = True
    lblMealClass.Enabled = True
    
    If ufgMealLink.DataGrid.Row > 0 Then ufgMealLink.LocateRow (1)
    ufgMealLink.Enabled = False
    ufgMealLink.DataGrid.Enabled = True
    ufgMealLink.DataGrid.BackColor = &H8000000F
    ufgMealLink.ReadOnly = False
    
    stbThis.Panels(2).Text = "��ǰ�ײ�����Ϊ��" & ufgMeal.DataGrid.Cell(flexcpText, ufgMeal.SelectionRow, 2)
End Sub

Private Function CheckMealName() As Boolean
'����ײ������Ƿ��ظ�
    Dim i As Integer
    
    CheckMealName = False
    For i = 1 To ufgMeal.GridRows - 1
        If Not ufgMeal.RowState(i) = TDataRowState.Del Then
            If Not mblnIsUpdate Then
                If (Not ufgMeal.RowState(i) = TDataRowState.Add) And (Not ufgMeal.RowHidden(i)) Then
                    If ufgMeal.Text(i, gstrAntibodyMeal_�ײ�����) = ufgMeal.DataGrid.Cell(flexcpText, ufgMeal.SelectionRow, 2) Then
                        MsgBoxD Me, "���������ظ���", vbExclamation, Me.Caption

                        ufgMeal.LocateRow ufgMeal.SelectionRow
                        ufgMeal.DataGrid.EditCell
                        Exit Function
                    End If
                End If
            Else
                If Not ufgMeal.SelectionRow = i And (Not ufgMeal.RowHidden(i)) Then
                    If ufgMeal.Text(i, gstrAntibodyMeal_�ײ�����) = ufgMeal.DataGrid.Cell(flexcpText, ufgMeal.SelectionRow, 2) Then
                        MsgBoxD Me, "���������ظ���", vbExclamation, Me.Caption

                        ufgMeal.LocateRow ufgMeal.SelectionRow
                        ufgMeal.DataGrid.EditCell
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
    
    CheckMealName = True
End Function

Private Sub Menu_PatholMeal_Cancel()
    Dim i As Integer
    
    'ȡ��ǰ�����ô��ڱ༭״̬�ĵ�Ԫ��ʧȥ���㣬�����ָܻ���ǰ�е�Ԫ���������Ϣ
    ufgMeal.DataGrid.Col = 5
    ufgMeal.DataGrid.Row = ufgMeal.SelectionRow
    ufgMeal.DataGrid.SetFocus
    
    mblnEdit = False
    mblnIsUpdate = False
    cboMealClass.Enabled = True
    lblMealClass.Enabled = True
    mblnCurModifyState = True
    
    If ufgMeal.CurKeyValue = "" Then ufgMeal.DelCurRow False
    
    ufgMealLink.HeadCheckValue = False
    
    If ufgMeal.AdoData.RecordCount > 0 Then
        Call ufgMeal.RestoreCurRowText
        Call ufgMeal.LocateRow(ufgMeal.SelectionRow)
        Call ConfigMealLink(Val(ufgMeal.KeyValue(ufgMeal.SelectionRow)))
    End If
    
    For i = 1 To ufgMeal.AdoData.RecordCount - 1
        ufgMeal.DataGrid.TextMatrix(i, 5) = Format(ufgMeal.DataGrid.Cell(flexcpText, i, 5), "yyyy-mm-dd")
    Next
    If ufgMealLink.DataGrid.Row > 0 Then ufgMealLink.LocateRow (1)
    ufgMealLink.Enabled = False
    ufgMealLink.DataGrid.Enabled = True
    ufgMealLink.DataGrid.BackColor = &H8000000F
    ufgMealLink.ReadOnly = False
    
    stbThis.Panels(2).Text = "��ǰ�ײ�����Ϊ��" & ufgMeal.DataGrid.Cell(flexcpText, ufgMeal.SelectionRow, 2)
End Sub

Private Sub cbrMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible = True Then Bottom = stbThis.Height
End Sub

Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim blnHasRecord As Boolean
    
On Error GoTo errHandle
    blnHasRecord = ufgMeal.IsSelectionRow
    
    Select Case control.ID
        Case conMenu_PatholMeal_Save
            control.Enabled = (Not mblnCurModifyState) And blnHasRecord
            
        Case conMenu_PatholMeal_Cancel
            control.Enabled = (Not mblnCurModifyState) And blnHasRecord
            
        Case conMenu_PatholMeal_AddRecord
            control.Enabled = mblnCurModifyState Or (Not blnHasRecord)
            
        Case conMenu_PatholMeal_ModRecord
            control.Enabled = mblnCurModifyState And blnHasRecord

        Case conMenu_PatholMeal_DelRecord
            control.Enabled = mblnCurModifyState And blnHasRecord

        Case conMenu_PatholMeal_UpRow
            control.Enabled = mblnEdit

        Case conMenu_PatholMeal_DownRow
            control.Enabled = mblnEdit

        Case conMenu_File_Exit
            
    End Select
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgMeal_OnBeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '�ڷ��������޸ĵ�����£�������༭��Ԫ��
    If Not mblnEdit Then Cancel = True
End Sub

Private Sub ufgMeal_OnBeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If Not mblnCurModifyState And OldRow <> NewRow Then Cancel = True
End Sub

Private Sub ufgMeal_OnColFormartChange()
 '�����б����
    zlDatabase.SetPara "�����ײ��б�����", ufgMeal.GetColsString(ufgMeal), glngSys, G_LNG_PATHOLSYS_NUM
End Sub

Private Sub ufgMeal_OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'�����Ҽ��˵�
On Error GoTo errHandle
    If Button = 2 Then
        Dim objPopup As CommandBar
        Dim objControl As CommandBarControl

        Set cbrMain.Icons = zlCommFun.GetPubIcons
        Set objPopup = cbrMain.Add("�Ҽ��˵�", xtpBarPopup)
        With objPopup.Controls
            Set objControl = .Add(xtpControlButton, conMenu_PatholMeal_AddRecord, "�����ײ�(&A)")
            Set objControl = .Add(xtpControlButton, conMenu_PatholMeal_ModRecord, "�޸��ײ�(&M)")
            Set objControl = .Add(xtpControlButton, conMenu_PatholMeal_DelRecord, "ɾ���ײ�(&D)")
            
            Set objControl = .Add(xtpControlButton, conMenu_PatholMeal_Save, "����(&S)"): objControl.IconId = 3091
            Set objControl = .Add(xtpControlButton, conMenu_PatholMeal_Cancel, "����(&R)"): objControl.IconId = 3565
        End With
        objPopup.ShowPopup
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgMealLink_OnCheckChanging(ByVal Row As Long, ByVal Col As Long, AllowChange As Boolean)
    If Not ufgMealLink.ReadOnly Then AllowChange = False
End Sub

Private Sub ufgMealLink_OnColFormartChange()
    zlDatabase.SetPara "�ײ���ϸ�б�����", ufgMealLink.GetColsString(ufgMealLink), glngSys, G_LNG_PATHOLSYS_NUM
End Sub

Private Sub InitFace()
'��ʼ�����ܽ���
    Dim Pane1 As Pane, Pane2 As Pane

    With dkpMeal
        .CloseAll
        .SetCommandBars cbrMain
        .Options.HideClient = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With

    Set Pane1 = dkpMeal.CreatePane(1, 0, Round(Me.Width / 2), DockLeftOf)
    Pane1.Title = "�ײͼ�¼"
    Pane1.Handle = picDatas.hWnd
    Pane1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Pane1.MinTrackSize.Width = 50

    Set Pane2 = dkpMeal.CreatePane(2, 0, Round(Me.Width / 2), DockRightOf)
    Pane2.Title = "������ϸ"
    Pane2.Handle = picMealLink.hWnd
    Pane2.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Pane2.MinTrackSize.Width = 50
End Sub

Private Sub LoadMealData()
'�����ײ�����
    Dim i As Integer
    Dim strSql As String
    Dim rsMeal As ADODB.Recordset
    
    strSql = "select �ײ�ID,�ײ�����,�ײ����,�ײ�˵��,����ʱ��,������ from �����ײ���Ϣ"
      
    Set ufgMeal.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    Call ufgMeal.RefreshData
    For i = 1 To ufgMeal.AdoData.RecordCount
        ufgMeal.DataGrid.Cell(flexcpText, i, 5) = Format(ufgMeal.DataGrid.Cell(flexcpText, i, 5), "yyyy-mm-dd")
    Next
End Sub

Private Sub LoadAntibodyData()
'��ȡ�������ݣ��ų����õĿ��壩
    Dim strSql As String
    Dim rsAntibody As ADODB.Recordset

    strSql = "select '' as ����ID,����ID,��������,��¡��,���ö���,������,Ӧ�����,��ע, '' as ����˳�� from ��������Ϣ where ʹ��״̬ = 1 order by ����ID"
      
    Set ufgMealLink.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    Call ufgMealLink.RefreshData
End Sub

Private Sub cboMealClass_Click()
On Error GoTo errHandle
    '�����ײ���Ϣ
    
    If cboMealClass.Text = "" Then
        ufgMeal.AdoData.Filter = ""
    Else
        ufgMeal.AdoData.Filter = "�ײ����='" & cboMealClass.Text & "'"
    End If
    
    Call ufgMeal.RefreshData
    
    If ufgMeal.DataGrid.Row <= 0 Then Exit Sub
    Call ConfigMealLink(Val(ufgMeal.KeyValue(ufgMeal.DataGrid.Row)))
    
    stbThis.Panels(2).Text = "��ǰ�ײ�����Ϊ��" & ufgMeal.DataGrid.Cell(flexcpText, ufgMeal.SelectionRow, 2)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkRowFilter_Click()
On Error GoTo errHandle
    If chkRowFilter.value = 1 Then
        Call ufgMealLink.ShowCheckRows
    Else
        Call ufgMealLink.ShowAllRows
    End If
Exit Sub
errHandle:
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

Private Sub Menu_PatholMeal_DelRecord()
'ɾ���ײ�
On Error GoTo errHandle
    If Not ufgMeal.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ����ɾ�����ײͼ�¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If MsgBoxD(Me, "ȷ��Ҫɾ�����ײ���", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    Call ufgMeal.DelRow(ufgMeal.SelectionRow, False)
    
    Call SaveMealData(True)
    
    If ufgMeal.ShowingDataRowCount <= 0 Then
        '��û���ײ�����ʱ������ײ͹���
        Call ufgMealLink.ClearCellCheck(ufgMealLink.GetColIndexWithRowCheck)
        
        Call ReinitMealLinkData
    Else
        '������һ�ײ͹���
        Call ConfigMealLink(Val(ufgMeal.KeyValue(ufgMeal.SelectionRow)))
    End If
    
    If ufgMeal.IsSelectionRow Then
        If ufgMeal.IsEmptyKey(ufgMeal.SelectionRow) Then
            Call ConfigButState(False)
        End If
    End If
    
    Call LoadMealClass
    
    stbThis.Panels(2).Text = "��ǰ�ײ�����Ϊ��" & ufgMeal.DataGrid.Cell(flexcpText, ufgMeal.SelectionRow, 2)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_PatholMeal_DownRow()
On Error GoTo errHandle
    Call ufgMealLink.MoveDown(ufgMealLink.SelectionRow)
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub SaveMealData(Optional ByVal blnIsSaveOnlyDel As Boolean = False)
'�����ײ�����
    Dim i As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim dtServicesTime As Date
    
    For i = 1 To ufgMeal.GridRows - 1
        Select Case ufgMeal.RowState(i)
            Case IIf(blnIsSaveOnlyDel, -1, TDataRowState.Add)
                dtServicesTime = zlDatabase.Currentdate
                
                '����µ��ײ�
                strSql = "select Zl_�����ײ�_����([1],[2],[3],[4],[5]) as ����ֵ from dual"
                
                Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                                    ufgMeal.Text(i, gstrAntibodyMeal_�ײ�����), _
                                                    ufgMeal.Text(i, gstrAntibodyMeal_�ײ����), _
                                                    ufgMeal.Text(i, gstrAntibodyMeal_�ײ�˵��), _
                                                    CDate(Format(dtServicesTime, "yyyy-mm-dd")), _
                                                    UserInfo.����)
                
                If rsData.RecordCount <= 0 Then
                    Call err.Raise(0, "SaveMealData", "δ�ɹ���ȡ��������ײ�ID,����ʧ�ܡ�")
                    Exit Sub
                End If
                
                ufgMeal.Text(i, gstrAntibodyMeal_�ײ�ID) = rsData!����ֵ
                ufgMeal.Text(i, gstrAntibodyMeal_������) = UserInfo.����
                ufgMeal.Text(i, gstrAntibodyMeal_����ʱ��) = dtServicesTime
                
                ufgMeal.SyncRowDataToAdo i
            Case TDataRowState.Del
                'ɾ���ײ�(�ἶ��ɾ����������)
                
                strSql = "Zl_�����ײ�_ɾ��(" & Val(ufgMeal.KeyValue(i)) & ")"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
                
                If ufgMeal.ShowingDataRowCount <= 0 Then
                    '���ѡ��
                    Call ufgMealLink.ClearCellCheck(ufgMealLink.GetColIndexWithRowCheck)
                    '������ID����Ϊ��
                    Call ReinitMealLinkData
                End If
                
                ufgMeal.SyncRowDataToAdo i
            Case IIf(blnIsSaveOnlyDel, -1, TDataRowState.Update)
                '�����ײ�
                strSql = "Zl_�����ײ�_����(" & Val(ufgMeal.KeyValue(i)) & ",'" & _
                                            ufgMeal.Text(i, gstrAntibodyMeal_�ײ�����) & "','" & _
                                            ufgMeal.Text(i, gstrAntibodyMeal_�ײ����) & "','" & _
                                            ufgMeal.Text(i, gstrAntibodyMeal_�ײ�˵��) & "')"
                
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
                
                ufgMeal.SyncRowDataToAdo i
        End Select
        
        '������״̬
        ufgMeal.RowState(i) = TDataRowState.Normal
    Next i
  
End Sub

Private Sub SaveMealLinkData(ByVal lngMealId As Long)
'�����ײͶ�Ӧ�Ŀ�������
    Dim i As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim lngMealLinkId As Long
    Dim lngAntibodyOrder As Long
    
    lngAntibodyOrder = 0
    
    'ɾ���ײ͹���
    strSql = "Zl_�����ײ͹���_ɾ��(" & lngMealId & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        
    For i = 1 To ufgMealLink.GridRows - 1

        If ufgMealLink.GetRowCheck(i) Then
            '�ж��Ƿ��й���ID,���û�У��������ӹ������������������
            strSql = "select Zl_�����ײ͹���_����([1],[2],[3]) as ����ֵ from dual"
            Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngMealId, Val(ufgMealLink.KeyValue(i)), lngAntibodyOrder)
            
            If rsData.RecordCount <= 0 Then
                Call err.Raise(0, "SaveMealLinkData", "δ�ɹ���ȡ��������ײͿ������ID,����ʧ�ܡ�")
                Exit Sub
            End If
            
            '���ù���ID
            ufgMealLink.Text(i, gstrAntibodyMealLink_����ID) = rsData!����ֵ
            
            lngAntibodyOrder = lngAntibodyOrder + 1
        Else
            ufgMealLink.Text(i, gstrAntibodyMealLink_����ID) = ""
        End If
        
'        If ufgMealLink.RowState(i) = TDataRowState.Update Then
'            lngMealLinkId = Val(ufgMealLink.Text(i, gstrAntibodyMealLink_����ID))
'
'            If ufgMealLink.GetRowChecked(i) Then
'                '�ж��Ƿ��й���ID,���û�У��������ӹ������������������
'                If lngMealLinkId <= 0 Then
'                    strSQL = "select Zl_�����ײ͹���_����([1],[2]) as ����ֵ from dual"
'                    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngMealId, Val(ufgMealLink.GetKeyValue(i)))
'
'                    If rsData.RecordCount <= 0 Then
'                        Call err.Raise(0, "SaveMealLinkData", "δ�ɹ���ȡ��������ײͿ������ID,����ʧ�ܡ�")
'                        Exit Sub
'                    End If
'
'                    '���ù���ID
'                    Call ufgMealLink.SetText(i, gstrAntibodyMealLink_����ID, rsData!����ֵ)
'                End If
'            Else
'                '�ж��Ƿ��й���ID,����У���ɾ�����������û����������
'                If lngMealLinkId > 0 Then
'                    strSQL = "Zl_�����ײ͹���_ɾ��1(" & lngMealLinkId & ")"
'                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
'
'                    '����ײ͹���ID
'                    Call ufgMealLink.SetText(i, gstrAntibodyMealLink_����ID, "")
'                End If
'            End If
'
'            '�ָ���״̬
'            ufgMealLink.RowState(i) = TDataRowState.Normal
'        End If
    Next i
End Sub

Private Sub Menu_PatholMeal_AddMeal()
    Dim i As Integer
    
On Error GoTo errHandle
    mblnEdit = True
    mblnIsUpdate = False
    cboMealClass.Enabled = False
    lblMealClass.Enabled = False
    ufgMealLink.ReadOnly = True
    
    For i = 1 To ufgMealLink.DataGrid.Rows - 1
        ufgMealLink.SetRowCheck i, False
    Next
    ufgMealLink.Enabled = True
    ufgMealLink.DataGrid.BackColor = vbWhite
    
    ufgMeal.Editable = flexEDKbd
    ufgMeal.NewRow
    ufgMeal.DataGrid.EditCell

    mblnCurModifyState = False
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
            
Private Sub Menu_PatholMeal_ModMeal()
On Error GoTo errHandle
    mblnEdit = True
    mblnIsUpdate = True
    cboMealClass.Enabled = False
    lblMealClass.Enabled = False
    ufgMealLink.Enabled = True
    ufgMealLink.DataGrid.BackColor = vbWhite
    ufgMealLink.ReadOnly = True
    
    ufgMeal.Editable = flexEDKbd
    ufgMeal.LocateRow ufgMeal.SelectionRow
    ufgMeal.DataGrid.EditCell

    mblnCurModifyState = False
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_PatholMeal_SaveMeal()
'�����ײ���Ϣ
On Error GoTo errHandle
    Dim blnValid As Boolean
    
    blnValid = Not ufgMeal.IsErrColorWithList
    If Not blnValid Then
        Call MsgBoxD(Me, "��⵽�ײ��б��д�����Ч���ݣ���ȷ����������Ƿ���ȷ������¼�룬����ɫ����ǵĵ�Ԫ��Ϊ��¼���ݡ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '�����ײ���Ϣ
    Call SaveMealData
    
    If ufgMeal.IsSelectionRow Then
        If Not ufgMeal.IsEmptyKey(ufgMeal.SelectionRow) Then
            Call ConfigButState(True)
        End If
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_PatholMeal_SureSelected()
'ȷ�Ϲ���ѡ��
On Error GoTo errHandle

    If Not ufgMeal.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ������Ӧ���ײ���Ϣ��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call SaveMealLinkData(Val(ufgMeal.KeyValue(ufgMeal.SelectionRow)))
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_PatholMeal_UpRow()
On Error GoTo errHandle
    Call ufgMealLink.MoveUp(ufgMealLink.SelectionRow)

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
'    InitDebugObject 1294, Me, "zlhis", "his"
    Call InitCommandBars

    Call RestoreWinState(Me, App.ProductName)
    
    Call InitFace
    
    '��ʼ���б�
    Call InitMealList
    Call InitMealLinkList
    
    '��������
    Call LoadMealData
    Call LoadAntibodyData
    Call LoadMealClass
    
    '���ѡ���˵�һ�У����Զ�������������
    If ufgMeal.IsSelectionRow And Trim(ufgMeal.KeyValue(ufgMeal.SelectionRow)) <> "" Then
        Call ConfigMealLink(Val(ufgMeal.KeyValue(ufgMeal.SelectionRow)))
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub LoadMealClass()
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim strMealClass As String
    
    strSql = "select distinct �ײ���� from �����ײ���Ϣ"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    cboMealClass.Clear
    cboMealClass.AddItem ""
    
    While Not rsData.EOF
        If Nvl(rsData!�ײ����) <> "" Then
            cboMealClass.AddItem Nvl(rsData!�ײ����)
            strMealClass = strMealClass & "|" & Nvl(rsData!�ײ����)
        End If
        rsData.MoveNext
    Wend
    
    ufgMeal.ComboxListFormat(ufgMeal.GetColIndex(gstrAntibodyMeal_�ײ����)) = strMealClass
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
        Set .Icons = zlCommFun.GetPubIcons                          '���ù�����ͼ��ؼ�
    End With

    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '�˵�����
'Begin------------------------�༭�˵�--------------------------------------Ĭ�Ͽɼ�
    cbrMain.ActiveMenuBar.Title = "�˵�"
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)")
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_Save, "����(&S)")
        cbrControl.IconId = 3091
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_Cancel, "����(&R)")
        cbrControl.IconId = 3565
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&Q)")
        cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)")
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_AddRecord, "�����ײ�(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_ModRecord, "�޸��ײ�(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_DelRecord, "ɾ���ײ�(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_UpRow, "����(&U)")
        cbrControl.IconId = 21802
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_DownRow, "����(&D)")
        cbrControl.IconId = 21801
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
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_Save, "����")
        cbrControl.IconId = 3091
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_Cancel, "����")
        cbrControl.IconId = 3565
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_AddRecord, "�����ײ�")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_ModRecord, "�޸��ײ�")
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_DelRecord, "ɾ���ײ�")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_UpRow, "����")
        cbrControl.IconId = 21802
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_PatholMeal_DownRow, "����")
        cbrControl.IconId = 21801
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

Private Sub picDatas_Resize()
'�����ײͽ��沼��
On Error Resume Next
    framMeals.Left = 120
    framMeals.Top = 120
    framMeals.Width = picDatas.Width - 120
    framMeals.Height = picDatas.Height - 240
    
    ufgMeal.Left = 120
    ufgMeal.Top = lblMealClass.Top + 360
    ufgMeal.Width = framMeals.Width - 240
    ufgMeal.Height = framMeals.Height - lblMealClass.Height - lblMealClass.Top - 240
End Sub

Private Sub ConfigMealLink(ByVal lngMealId As Long)
'�����ײ͹���(��ѯ���ײ������Ŀ��壬Ȼ���ٿ����б�Ķ�Ӧchecked������ΪTrue)
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select ID, ����ID,����˳��  from �����ײ͹��� where �ײ�ID=[1] order by ����˳��"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngMealId)
    
    '���ѡ��
    Call ufgMealLink.ClearCellCheck(ufgMealLink.GetColIndexWithRowCheck())
    
    '������ID����Ϊ��
    Call ReinitMealLinkData
        
    If rsData.RecordCount <= 0 Then Exit Sub

    Do While Not rsData.EOF
        Call SetMealAntibodyLink(Val(Nvl(rsData!����ID)), Val(Nvl(rsData!ID)), Val(Nvl(rsData!����˳��)))
        
        rsData.MoveNext
    Loop
    
'    Call ufgMealLink.Sort(ufgMealLink.vfgHelper.GetColumnIndex(ufgMealLink.vfgHelper.CheckColName))
    Call ufgMealLink.Sort(ufgMealLink.GetColIndex(gstrAntibodyMealLink_����˳��))
End Sub

Private Sub SetMealAntibodyLink(ByVal lngAntibodyId As Long, ByVal lngMealLinkId As Long, ByVal lngAntibodyOrder As Long)
'�����ײͿ������
    Dim i As Long
    
    ufgMealLink.ReadOnly = True
    For i = 1 To ufgMealLink.GridRows - 1
        If Val(ufgMealLink.KeyValue(i)) = lngAntibodyId Then
            ufgMealLink.Text(i, gstrAntibodyMealLink_����ID) = lngMealLinkId
            ufgMealLink.Text(i, gstrAntibodyMealLink_����˳��) = String(4 - Len("" & lngAntibodyOrder & ""), "0") & lngAntibodyOrder
            Call ufgMealLink.SetRowCheck(i, True)
            ufgMealLink.ReadOnly = False
            Exit Sub
        End If
    Next i
    
End Sub

Private Sub ReinitMealLinkData()
'���³�ʼ���ײ͹�������
    Dim i As Long

    For i = 1 To ufgMealLink.GridRows - 1
        ufgMealLink.Text(i, gstrAntibodyMealLink_����ID) = ""
        ufgMealLink.Text(i, gstrAntibodyMealLink_����˳��) = "9999"
        ufgMealLink.RowState(i) = TDataRowState.Normal
    Next i
End Sub

Private Sub picMealLink_Resize()
'�����ײ���ϸ���沼��
On Error Resume Next
    framAntibody.Left = 120
    framAntibody.Top = 120
    framAntibody.Width = picMealLink.Width - 240
    framAntibody.Height = picMealLink.Height - 240
    
    ufgMealLink.Left = 120
    ufgMealLink.Top = chkRowFilter.Top + 360
    ufgMealLink.Width = framAntibody.Width - 240
    ufgMealLink.Height = framAntibody.Height - lblFind.Height - lblFind.Top - 240
End Sub

Private Sub txtFind_Change()
On Error GoTo errHandle
    Dim lngFindIndex As Long
    
    If Trim(txtFind.Text) = "" Then Exit Sub
    
    lngFindIndex = ufgMealLink.FindRowIndex(txtFind.Text, gstrAntibodyMealLink_��������)
    
    If lngFindIndex > 0 Then Call ufgMealLink.LocateRow(lngFindIndex)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub txtFind_GotFocus()
On Error Resume Next
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)
End Sub

Private Sub ufgMeal_OnAfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim iCol As Long
    Dim i As Long
    
    If ufgMeal.IsNullRow(Row) Then
        ufgMeal.RowState(Row) = TDataRowState.Normal
        Call ufgMeal.SetRowColor(Row, ufgMeal.BackColor)
        
        Exit Sub
    End If
        
    '���δ¼��걾���ƣ�����ʾ����ɫ
    iCol = ufgMeal.GetColIndex(gstrAntibodyMeal_�ײ�����)
    
    ufgMeal.CellColor(Row, iCol) = IIf(ufgMeal.Text(Row, gstrAntibodyMeal_�ײ�����) = "", ufgMeal.ErrCellColor, ufgMeal.BackColor)
End Sub

Private Sub ConfigButState(ByVal blnEnable As Boolean)
    mblnCurModifyState = blnEnable
End Sub

Private Sub ufgMeal_OnClick()
On Error GoTo errHandle
    If Not mblnCurModifyState Then
        ufgMeal.Editable = flexEDKbd
        ufgMeal.DataGrid.EditCell
        
        Exit Sub
    End If
    
    If ufgMeal.ShowingDataRowCount <= 0 Then
        Call ufgMealLink.ClearCellCheck(ufgMealLink.GetColIndexWithRowCheck)
        Call ReinitMealLinkData
        Call ConfigButState(False)
        
        Exit Sub
    End If

    If ufgMeal.MouseRowIndex <= 0 Then Exit Sub
    
    If Trim(ufgMeal.KeyValue(ufgMeal.MouseRowIndex)) = "" Then
        Call ufgMealLink.ClearCellCheck(ufgMealLink.GetColIndexWithRowCheck)
        Call ReinitMealLinkData
        Call ConfigButState(False)
        
        Exit Sub
    End If

    '���ù���
    Call ConfigMealLink(Val(ufgMeal.KeyValue(ufgMeal.MouseRowIndex)))
    Call ConfigButState(True)
    
    If chkRowFilter.value = 1 Then
        Call ufgMealLink.ShowCheckRows
    End If
    stbThis.Panels(2).Text = "��ǰ�ײ�����Ϊ��" & ufgMeal.DataGrid.Cell(flexcpText, ufgMeal.SelectionRow, 2)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgMealLink_OnClick()
On Error GoTo errHandle
    stbThis.Panels(2).Text = "��ǰ��������Ϊ��" & ufgMealLink.DataGrid.Cell(flexcpText, ufgMealLink.SelectionRow, 3)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgMeal_OnStartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row > 0 Then
        If ufgMeal.Text(Row, gstrAntibodyMeal_�ײ����) = "" And cboMealClass.Text <> "" Then
            ufgMeal.Text(Row, gstrAntibodyMeal_�ײ����) = cboMealClass.Text
        End If
    End If
End Sub

Private Sub ufgMealLink_OnAfterEdit(ByVal Row As Long, ByVal Col As Long)
    '�����޸�Ϊ����״̬
    If ufgMealLink.IsComboboxCol(Col) Then
        ufgMealLink.RowState(Row) = TDataRowState.Update
    End If
End Sub


Private Sub ufgMealLink_OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'�����Ҽ��˵�
On Error GoTo errHandle
    If Button = 2 Then
        Dim objPopup As CommandBar
        Dim objControl As CommandBarControl

        Set cbrMain.Icons = zlCommFun.GetPubIcons
        Set objPopup = cbrMain.Add("�Ҽ��˵�", xtpBarPopup)
        With objPopup.Controls
            Set objControl = .Add(xtpControlButton, conMenu_PatholMeal_UpRow, "����(&U)"): objControl.IconId = 21802
            Set objControl = .Add(xtpControlButton, conMenu_PatholMeal_DownRow, "����(&D)"): objControl.IconId = 21801
        End With
        objPopup.ShowPopup
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
