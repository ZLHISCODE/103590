VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmPatholSpecialExamined_WorkPrint 
   Caption         =   "�ؼ���������"
   ClientHeight    =   8280
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   11685
   Icon            =   "frmPatholSpecialExamined_WorkPrint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   11685
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   120
      ScaleHeight     =   7215
      ScaleWidth      =   11415
      TabIndex        =   2
      Top             =   480
      Width           =   11415
      Begin VB.OptionButton optAll 
         Caption         =   "�� ��"
         Height          =   180
         Left            =   6480
         TabIndex        =   21
         Top             =   1080
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optXiMu2 
         Caption         =   "��ҩ��ҩ"
         Height          =   180
         Left            =   5160
         TabIndex        =   20
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton optXiMu1 
         Caption         =   "�� ��"
         Height          =   180
         Left            =   3960
         TabIndex        =   19
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Frame framFilter 
         Height          =   735
         Left            =   0
         TabIndex        =   8
         Top             =   120
         Width           =   11415
         Begin VB.CheckBox chkMoney 
            Caption         =   "�Ʒ�"
            Height          =   255
            Left            =   9120
            TabIndex        =   12
            Top             =   280
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.TextBox txtStartPatholNum 
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
            Left            =   5880
            TabIndex        =   11
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtEndPatholNum 
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
            Left            =   7600
            TabIndex        =   10
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdFilter 
            Caption         =   "�� ѯ(&Q)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   9960
            TabIndex        =   9
            Top             =   200
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker dtpStartRequisition 
            Height          =   300
            Left            =   1080
            TabIndex        =   13
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd hh:mm"
            Format          =   97058819
            CurrentDate     =   40679.0594097222
         End
         Begin MSComCtl2.DTPicker dtpEndRequisition 
            Height          =   300
            Left            =   3120
            TabIndex        =   14
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   97058819
            CurrentDate     =   40679.0594097222
         End
         Begin VB.Label Label2 
            Caption         =   "��"
            Height          =   255
            Left            =   7360
            TabIndex        =   18
            Top             =   285
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "��"
            Height          =   255
            Left            =   2920
            TabIndex        =   17
            Top             =   285
            Width           =   255
         End
         Begin VB.Label Label7 
            Caption         =   "����ʱ�䣺"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   280
            Width           =   975
         End
         Begin VB.Label labPatholNum 
            Caption         =   "����ţ�"
            Height          =   255
            Left            =   5160
            TabIndex        =   15
            Top             =   285
            Width           =   720
         End
      End
      Begin VB.Frame framSpeExam 
         Caption         =   "��Ƭ��������"
         Height          =   5055
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   11295
         Begin zl9PACSWork.ucFlexGrid ufgData 
            Height          =   4695
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   8281
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
      Begin VB.CheckBox chkYSQ 
         Caption         =   "������"
         Height          =   255
         Left            =   4560
         TabIndex        =   5
         ToolTipText     =   "��ʾ��Ƭ״̬Ϊ��δ��������Ƭ��¼��"
         Top             =   6600
         Width           =   855
      End
      Begin VB.CheckBox chkYJS 
         Caption         =   "�ѽ���"
         Height          =   180
         Left            =   5520
         TabIndex        =   4
         ToolTipText     =   "��ʾ��Ƭ״̬Ϊ���ѽ��ܡ�����Ƭ��¼��"
         Top             =   6630
         Width           =   855
      End
      Begin VB.CheckBox chkYWC 
         Caption         =   "�����"
         Height          =   180
         Left            =   6480
         TabIndex        =   3
         ToolTipText     =   "��ʾ��Ƭ״̬Ϊ��δ��ɡ�����Ƭ��¼��"
         Top             =   6630
         Width           =   855
      End
      Begin XtremeSuiteControls.TabControl tabFilter 
         Height          =   375
         Left            =   -120
         TabIndex        =   22
         Top             =   960
         Width           =   11445
         _Version        =   589884
         _ExtentX        =   20188
         _ExtentY        =   661
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picTag 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   7920
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatholSpecialExamined_WorkPrint.frx":179A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13732
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
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPatholSpecialExamined_WorkPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim WithEvents zlReport As zl9Report.clsReport
Attribute zlReport.VB_VarHelpID = -1

Private mlngPatholAdivceId As Long
Private mufgGrid As ucFlexGrid
Private mstrPrivs As String

Private mblnAutoAcceptOfAfterPrint As Boolean '��ӡ���Զ�����


Public blnIsOk As Boolean


Private mlngFilterTabIndex As Long

Private Enum TMenuType
    mtLab = 1
    mtLabView = 10
    mtLabPrint = 11
    
    mtWork = 2
    mtWorkView = 20
    mtWorkPrint = 21
    
    mtAccept = 3
    mtComplete = 4
    mtchkYSQ = 5
    mtchkYJS = 6
    mtchkYWC = 7
End Enum

Public Sub ShowWorkPrint(ufgGrid As ucFlexGrid, ByVal lngPatholAdivceId As Long, _
    ByVal lngCurSpeExamType As Long, ByVal strPrivs As String, owner As Form)
'��ʾ�����嵥��ӡ����
    Set mufgGrid = ufgGrid

    mlngPatholAdivceId = lngPatholAdivceId
    mstrPrivs = strPrivs
    blnIsOk = False

    '���õ�ǰ�ؼ�����
'    Call ConfigSpeExamType(lngCurSpeExamType)
    
    Call ConfigSpeExamPopedom
    
'    '�����ؼ�����
'    If lngPatholAdivceId > 0 Then
'        Call LoadSpeExamData(lngPatholAdivceId, lngCurSpeExamType)
'    End If
    
    'ˢ������
    Call RefreshSilcesCount
    
    Call Me.Show(1, owner)
    
End Sub


'Private Sub ConfigSpeExamType(ByVal strCurSpeExamType As String)
''���õ�ǰ�ؼ�����
'    Dim i As Long
'
'    For i = 0 To tabFilter.ItemCount - 1
'        If tabFilter(i).Tag Like "*" & strCurSpeExamType & "*" Then
'            tabFilter(i).Selected = True
'            Exit Sub
'        End If
'    Next i
'End Sub


Private Sub ConfigSpeExamPopedom()
'�����ؼ�Ȩ�ޣ�����û��Ȩ�޵ı�ǩ
    Dim blnIsPopedom As Boolean
    
    blnIsPopedom = CheckPopedom(mstrPrivs, "�����黯")
    tabFilter(0).Visible = blnIsPopedom
    
    blnIsPopedom = CheckPopedom(mstrPrivs, "����Ⱦɫ")
    tabFilter(1).Visible = blnIsPopedom
    
    blnIsPopedom = CheckPopedom(mstrPrivs, "���Ӳ���")
    tabFilter(2).Visible = blnIsPopedom
End Sub


Private Sub InitSpeExamWorkList()
'��ʼ���ؼ칤���嵥��ʾ�б�
'    ufgData.DataGrid.MergeCells = flexMergeRestrictRows
    Dim strTemp As String
    
        '��������
    ufgData.GridRows = glngStandardRowCount
    '�����и�
    ufgData.RowHeightMin = glngStandardRowHeight
    
    '�ж����ݿ�������Ƿ������� �����ȡ���ݿ����  û�������Ĭ��
    strTemp = zlDatabase.GetPara("�ؼ��嵥�б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
     
    If strTemp = "" Then
        ufgData.ColNames = gstrSpeExamWorkCols
    Else
        ufgData.ColNames = strTemp
    End If

    ufgData.DefaultColNames = gstrSpeExamWorkCols
    ufgData.ColConvertFormat = gstrSpeExamWorkConvertFormat
    ufgData.IsShowPopupMenu = False
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo ErrorHand
    
    Select Case control.ID
        Case TMenuType.mtLabView                    '��ǩԤ��
            Call Menu_File_LabView(control)
            
        Case TMenuType.mtLabPrint                   '��ǩ��ӡ
            Call Menu_File_LabPrint(control)
        
        Case TMenuType.mtWorkView                   '��ǩԤ��
            Call Menu_File_WorkView(control)
        
        Case TMenuType.mtWorkPrint                  '�嵥��ӡ
            Call Menu_File_WorkPrint(control)
        
        Case TMenuType.mtAccept                     '�ؼ����
            Call Menu_Edit_Accept
        
        Case TMenuType.mtComplete                   '�ؼ����
            Call Menu_Edit_Complate
        
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
On Error Resume Next
    blnIsOk = False
    Call Unload(Me)
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

Private Sub cbrMain_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
On Error Resume Next
    picBack.Left = Left
    picBack.Top = Top
    picBack.Width = Right - Left
    picBack.Height = Bottom - Top
End Sub

Private Sub chkYJS_Click()
On Error GoTo errHandle
    If Not tabFilter.Visible Then Exit Sub
    
    Call FilterSpeExamData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkYSQ_Click()
On Error GoTo errHandle
    If Not tabFilter.Visible Then Exit Sub
    
    Call FilterSpeExamData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkYWC_Click()
On Error GoTo errHandle
    If Not tabFilter.Visible Then Exit Sub
    
    Call FilterSpeExamData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub optAll_Click()
On Error GoTo errHandle
    If Not tabFilter.Visible Then Exit Sub
    
    Call FilterSpeExamData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub optXiMu1_Click()
On Error GoTo errHandle
    If Not tabFilter.Visible Then Exit Sub
    
    Call FilterSpeExamData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub optXiMu2_Click()
On Error GoTo errHandle
    If Not tabFilter.Visible Then Exit Sub
    
    Call FilterSpeExamData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub tabFilter_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error GoTo errHandle
    Call ConfigSpeexamDetail(Item.Tag)
    
    If Not tabFilter.Visible Then Exit Sub
    
    Call FilterSpeExamData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Exit Sub
End Sub


Private Sub ConfigSpeexamDetail(ByVal strSpeExamTag As String)
'�����ؼ�ϸĿ

    optXiMu1.Visible = True
    optXiMu2.Visible = True
    optAll.Visible = True
            
    Select Case Val(strSpeExamTag)
        Case 0
            optXiMu1.Caption = "�� ��"
            optXiMu2.Caption = "��ҩ��ҩ"
            
        Case 1
            optXiMu1.Visible = False
            optXiMu2.Visible = False
            optAll.Visible = False
            
        Case 2
            optXiMu1.Caption = "ӫ�����"
            optXiMu2.Caption = "��ͨ����"

    End Select
End Sub



Private Sub ufgData_OnColFormartChange()
 '�����б����
     zlDatabase.SetPara "�ؼ��嵥�б�����", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
End Sub


Private Sub RefreshSilcesCount()
'ˢ����Ƭ��¼����
    Dim i As Long
    Dim lngFinishCount As Long
    Dim lngNeedCount As Long

    lngFinishCount = 0
    lngNeedCount = 0


    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.RowHidden(i) Then
            If Not ufgData.IsNullRow(i) Then

                If ufgData.Text(i, gstrSlices_��ǰ״̬) <> "�����" Then
                    lngNeedCount = lngNeedCount + 1
                Else
                    lngFinishCount = lngFinishCount + 1
                End If
            End If
        End If
    Next i

    stbThis.Panels(2).Text = "�������Ŀ����" & lngFinishCount & "    ������Ŀ����" & lngNeedCount
    
End Sub



Private Sub LoadSpeExamData(ByVal lngPatholAdivceId As Long, Optional ByVal lngSpeExamType As Long = -1)
'�����ؼ�����
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select a.Id,c.�������,c.�����,c.����ҽ��ID, e.����, a.�Ŀ�ID, b.���, b.�걾����, a.�ؼ�����, a.����id,  d.��������, a.��������,a.��ǰ״̬,a.�嵥״̬ " & _
                " from �����ؼ���Ϣ a, ����ȡ����Ϣ b, ��������Ϣ c, ��������Ϣ d, ����ҽ����¼ e " & _
                " Where a.�Ŀ�id = b.�Ŀ�id And b.����ҽ��ID = c.����ҽ��ID And c.ҽ��ID = e.ID And a.����id = d.����id " & _
                " and c.����ҽ��ID=[1] " & IIf(lngSpeExamType >= 0, " and �ؼ�����=[2]", "") & " and a.��ǰ״̬ <> 2 and a.�嵥״̬=0" & _
                " order by a.�ؼ�����,b.�����,a.��ǰ״̬,b.���,a.Id "
    'If mblnMoved Then strSql = GetMovedDataSql(strSql)
                
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatholAdivceId, lngSpeExamType)
    
    Call ufgData.RefreshData
End Sub


Private Sub GetSpeExamData()
'���ݹ���������ѯ�ؼ�����
    Dim strSql As String
    Dim strPatholNumQuery As String
    Dim rsData As ADODB.Recordset
    
    
    strPatholNumQuery = ""
    If Trim(txtStartPatholNum.Text) <> "" And Trim(txtEndPatholNum.Text) <> "" Then
        strPatholNumQuery = " and (REGEXP_SUBSTR(upper(c.�����), '[[:alpha:]]+') >=REGEXP_SUBSTR(upper([3]),'[[:alpha:]]+') and to_number(REGEXP_SUBSTR(upper(c.�����), '[[:digit:]]+')) >=to_number(REGEXP_SUBSTR(upper([3]),  '[[:digit:]]+'))) "
        strPatholNumQuery = strPatholNumQuery & " and  (REGEXP_SUBSTR(upper(c.�����), '[[:alpha:]]+') <=REGEXP_SUBSTR(upper([4]),'[[:alpha:]]+') and to_number(REGEXP_SUBSTR(upper(c.�����),  '[[:digit:]]+')) <=to_number(REGEXP_SUBSTR(upper([4]), '[[:digit:]]+'))) "
    ElseIf Trim(txtStartPatholNum.Text) <> "" Then
        strPatholNumQuery = " and upper(c.�����)=upper([3]) "
    ElseIf Trim(txtStartPatholNum.Text) <> "" Then
        strPatholNumQuery = " and upper(c.�����) =upper([4]) "
    End If
    
    
    strSql = "select * from (select /*+ Rule*/ a.Id,c.�������,c.�����,c.����ҽ��ID, e.����, a.�Ŀ�ID,b.���, b.�걾����, a.�ؼ�����,a.�ؼ�ϸĿ, a.����id,  d.��������, a.��������,a.��ǰ״̬,a.�嵥״̬,f.����״̬, " & _
                " (select count(*) from ����ҽ������ X,������ü�¼ Y where X.��¼����=Y.��¼���� and X.no = Y.no and Y.��¼״̬=0 and X.ҽ��Id=c.ҽ��ID) as ����, f.����ʱ��, a.���ʱ��" & _
                " from �����ؼ���Ϣ a, ����ȡ����Ϣ b, ��������Ϣ c, ��������Ϣ d, ����ҽ����¼ e ,����������Ϣ f " & _
                " Where a.�Ŀ�id = b.�Ŀ�id And b.����ҽ��ID = c.����ҽ��ID And c.ҽ��ID = e.ID And a.����id = d.����id and a.����ID=f.����ID and f.����ʱ�� between [1] and [2] " & _
                IIf(strPatholNumQuery <> "", strPatholNumQuery, "") & ")" & _
                IIf(chkMoney.value <> 0, " where ����״̬<>1 and ����<=0 ", "") & _
                " order by �ؼ�����,�����, �ؼ�ϸĿ,��ǰ״̬,���,ID "
    'If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
                
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                            CDate(dtpStartRequisition.value), _
                                            CDate(dtpEndRequisition.value), _
                                            txtStartPatholNum.Text, _
                                            txtEndPatholNum.Text)
                                            
                                                                    
    Call FilterSpeExamData
End Sub


Private Sub FilterSpeExamData()
'���˲�ѯ��������Ƭ����
    Dim strFilters As String
    Dim strStudyTypeFilter As String
    
    strFilters = ""
    
    
    strStudyTypeFilter = ""
    
    '�ؼ�ϸĿ��0-�ޣ�1-����2-��ҩ��ҩ��3-ӫ����ӣ�4-��ͨ����
    Select Case Val(tabFilter.Selected.Tag)
        Case 0
            'optXiMu1��ʾ�Ƿ�ѡ�����߼���
            If optXiMu1.value Then
                strStudyTypeFilter = "�ؼ�����=0 and �ؼ�ϸĿ=1"
            ElseIf optXiMu2.value Then
                strStudyTypeFilter = "�ؼ�����=0 and �ؼ�ϸĿ=2"
            Else
                strStudyTypeFilter = "�ؼ�����=0"
            End If
            
            
        Case 1
            strStudyTypeFilter = "�ؼ�����=1"
            
        Case 2
            'optXiMu1��ʾ�Ƿ�ѡ��ӫ�����
            If optXiMu1.value Then
                strStudyTypeFilter = "�ؼ�����=2 and �ؼ�ϸĿ=3"
            ElseIf optXiMu2.value Then
                strStudyTypeFilter = "�ؼ�����=2 and �ؼ�ϸĿ=4"
            Else
                strStudyTypeFilter = "�ؼ�����=2"
            End If
    End Select
    
        
    If chkYSQ.value <> 0 Then
        strFilters = "(" & IIf(strStudyTypeFilter = "", "", strStudyTypeFilter & " and ") & " ��ǰ״̬=0)"
    End If
    
    If chkYJS.value <> 0 Then
        If strFilters <> "" Then strFilters = strFilters & " or "
        strFilters = strFilters & "(" & IIf(strStudyTypeFilter = "", "", strStudyTypeFilter & " and ") & " ��ǰ״̬=1)"
    End If
    
    If chkYWC.value <> 0 Then
        If strFilters <> "" Then strFilters = strFilters & " or "
        strFilters = strFilters & "(" & IIf(strStudyTypeFilter = "", "", strStudyTypeFilter & " and ") & " ��ǰ״̬=2)"
    End If
    
    '�������״̬������ѡ������ʾ��ǰ�ؼ����������м�¼
    If chkYSQ.value = 0 And chkYJS.value = 0 And chkYWC.value = 0 Then
        strFilters = "(" & IIf(strStudyTypeFilter = "", "", strStudyTypeFilter & " and ") & " ��ǰ״̬=0)" & " or " & _
                     "(" & IIf(strStudyTypeFilter = "", "", strStudyTypeFilter & " and ") & " ��ǰ״̬=1)" & " or " & _
                     "(" & IIf(strStudyTypeFilter = "", "", strStudyTypeFilter & " and ") & " ��ǰ״̬=2)"
    End If
    
    If ufgData.AdoData Is Nothing Then Exit Sub
    
    ufgData.AdoData.Filter = strFilters
    ufgData.RefreshData
    
    Call RefreshSilcesCount
End Sub


Private Sub picBack_Resize()
'�������ڲ���
    On Error Resume Next
    
    framFilter.Left = 120
    framFilter.Top = 0
    framFilter.Width = picBack.Width - 240
    
    tabFilter.Left = 120
    tabFilter.Top = framFilter.Top + framFilter.Height
    tabFilter.Width = picBack.Width - 240
    
    optXiMu1.Left = dtpEndRequisition.Left + 120
    optXiMu1.Top = tabFilter.Top + 90
    
    optXiMu2.Left = optXiMu1.Left + optXiMu1.Width + 120
    optXiMu2.Top = optXiMu1.Top
    
    optAll.Left = optXiMu2.Left + optXiMu2.Width + 120
    optAll.Top = optXiMu1.Top

    chkYSQ.Left = optAll.Left + optAll.Width + 720
    chkYSQ.Top = optAll.Top - 20
    
    chkYJS.Left = chkYSQ.Left + chkYSQ.Width + 240
    chkYJS.Top = chkYSQ.Top + 40
    
    chkYWC.Left = chkYJS.Left + chkYJS.Width + 240
    chkYWC.Top = chkYJS.Top
    
    framSpeExam.Left = 120
    framSpeExam.Top = tabFilter.Top + tabFilter.Height
    framSpeExam.Width = picBack.Width - 240
    framSpeExam.Height = picBack.Height - framFilter.Height - tabFilter.Height - 60
    
    ufgData.Left = 120
    ufgData.Top = 240
    ufgData.Width = framSpeExam.Width - 240
    ufgData.Height = framSpeExam.Height - 300
End Sub



Private Sub SpeExamBatAccept()
'�ؼ���������
    Dim i As Long
    Dim curPatholAdviceID As String
    Dim strSql As String
    Dim blnUpdateCallWind As Boolean
    
    blnUpdateCallWind = False
    
    For i = 1 To ufgData.GridRows - 1
        '���ѡ�еļ�飬�Ž��н���
        If ufgData.GetCellCheckState(i, ufgData.GetColIndexWithRowCheck()) Then
            If curPatholAdviceID <> ufgData.Text(i, gstrSpeExamWork_����ҽ��ID) Then
                curPatholAdviceID = ufgData.Text(i, gstrSpeExamWork_����ҽ��ID)
                
                strSql = "Zl_�����ؼ�_����(" & Val(curPatholAdviceID) & "," & Val(tabFilter.Selected.Tag) & ",'" & UserInfo.���� & "')"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            End If
            
            '���µ�ǰ�б�״̬
            If ufgData.Text(i, gstrSpeExamWork_��ǰ״̬) = "������" Then
                ufgData.Text(i, gstrSpeExamWork_��ǰ״̬) = "�ѽ���"
            End If
            
            '���µ��ý����б�״̬
            If Val(curPatholAdviceID) = mlngPatholAdivceId Then
                blnUpdateCallWind = True
            End If
        End If
    Next i
    
    If blnUpdateCallWind And Not (mufgGrid Is Nothing) Then
        For i = 1 To mufgGrid.GridRows - 1
            If mufgGrid.Text(i, gstrSpeExam_��ǰ״̬) = "������" Then
                Call mufgGrid.SyncText(i, gstrSpeExam_��ǰ״̬, "�ѽ���", True)
                Call mufgGrid.SyncText(i, gstrSpeExam_�ؼ�ҽʦ, UserInfo.����, True)
            End If
        Next i
    End If
End Sub




Private Sub SpeExamBatSure()
'�ؼ���������
    Dim i As Long
    Dim curPatholAdviceID As String
    Dim strSql As String
    Dim blnUpdateCallWind As Boolean
    Dim dtServicesTime As Date
    
    dtServicesTime = zlDatabase.Currentdate
    blnUpdateCallWind = False
    
    For i = 1 To ufgData.GridRows - 1
        '���ѡ�еļ�飬�Ž��н���
        If ufgData.GetCellCheckState(i, ufgData.GetColIndexWithRowCheck()) Then
            If curPatholAdviceID <> ufgData.Text(i, gstrSpeExamWork_����ҽ��ID) Then
                curPatholAdviceID = ufgData.Text(i, gstrSpeExamWork_����ҽ��ID)
                
                strSql = "Zl_�����ؼ�_ȷ��(" & Val(curPatholAdviceID) & "," & Val(tabFilter.Selected.Tag) & "," & To_Date(dtServicesTime) & ")"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            End If
            
            '���µ�ǰ�б�״̬
            If ufgData.Text(i, gstrSpeExamWork_��ǰ״̬) = "�ѽ���" Then
                ufgData.Text(i, gstrSpeExamWork_��ǰ״̬) = "�����"
            End If
            
            '���µ��ý����б�״̬
            If Val(curPatholAdviceID) = mlngPatholAdivceId Then
                blnUpdateCallWind = True
            End If
        End If
    Next i
    
    If blnUpdateCallWind And Not (mufgGrid Is Nothing) Then
        For i = 1 To mufgGrid.GridRows - 1
            If mufgGrid.Text(i, gstrSpeExam_��ǰ״̬) = "�ѽ���" Then
                Call mufgGrid.SyncText(i, gstrSpeExam_��ǰ״̬, "�����", True)
                Call mufgGrid.SyncText(i, gstrSpeExam_�ؼ�ҽʦ, UserInfo.����, True)
            End If
        Next i
    End If
End Sub



'Private Sub cbxSpeExamType_Click()
'On Error GoTo errHandle
'    Dim blnIsAllowFilter As Boolean
'
'    blnIsAllowFilter = False
'    Select Case Val(tabFilter.Selected.Tag)
'        Case 0
'            blnIsAllowFilter = CheckPopedom(mstrPrivs, "�����黯")
'
'        Case 1
'            blnIsAllowFilter = CheckPopedom(mstrPrivs, "����Ⱦɫ")
'
'        Case 2
'            blnIsAllowFilter = CheckPopedom(mstrPrivs, "���Ӳ���")
'
'    End Select
'
'    cmdFilter.Enabled = blnIsAllowFilter
'
'    If Not blnIsAllowFilter Then
'        Call MsgBoxD(Me, "���߱���ѯ���ؼ��������ݵ�Ȩ�ޡ�", vbOKOnly, Me.Caption)
'    End If
'Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'End Sub




Private Function CheckAllowSureOrAccept(Optional ByVal blnIsSure As Boolean = True) As Boolean
'�ж��Ƿ���Ҫ���к���
    Dim i As Long
    
    CheckAllowSureOrAccept = False
    For i = 1 To ufgData.GridRows - 1
        If ufgData.GetRowCheck(i) = True And (ufgData.Text(i, gstrSlices_��ǰ״̬) = IIf(blnIsSure, "�ѽ���", "������")) Then
            CheckAllowSureOrAccept = True
            Exit Function
        End If
    Next i
End Function


Private Sub Menu_Edit_Accept()
'�ؼ����
On Error GoTo errHandle
    If Not CheckAllowSureOrAccept(False) Then
        Call MsgBoxD(Me, "û����Ҫ���н��ܵ��ؼ���Ŀ��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    
    Call SpeExamBatAccept
    
    blnIsOk = True
    
    Call MsgBoxD(Me, "����ɶ�ѡ�м��Ľ��ܴ���", vbOKOnly, Me.Caption)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_Edit_Complate()
'�ؼ�ȷ��
On Error GoTo errHandle
    If Not CheckAllowSureOrAccept(True) Then
        Call MsgBoxD(Me, "û����Ҫ������ɵ��ؼ���Ŀ��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    
    Call SpeExamBatSure
    
    blnIsOk = True
    
    Call MsgBoxD(Me, "����ɶ���ѡ������ɴ���", vbOKOnly, Me.Caption)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdFilter_Click()
On Error GoTo errHandle
    Call GetSpeExamData
    
    Call RefreshSilcesCount
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub PrintSpeExamLabel(ByVal cbrControl As CommandBarControl)
'��ӡ�ؼ���Ŀ��ǩ
    Dim i As Long
    Dim j As Long
    Dim strValue(5) As String
    Dim bytStyle As Byte
    
    j = 0
    strValue(0) = "0": strValue(1) = "0": strValue(2) = "0": strValue(3) = "0": strValue(4) = "0": strValue(5) = "0"
    For i = 1 To ufgData.GridRows - 1
        If ufgData.GetCellCheckState(i, ufgData.GetColIndexWithRowCheck()) Then
            If zlCommFun.ActualLen(strValue(j)) > 2000 Then
                j = j + 1
                strValue(j) = ""
            End If

            If strValue(j) <> "" Then strValue(j) = strValue(j) & ","

            strValue(j) = strValue(j) & ufgData.KeyValue(i)
        End If
    Next i
    
    If cbrControl.ID = TMenuType.mtLabView Then
        bytStyle = 1
    Else
        bytStyle = 2
    End If
    
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_11", Me, "��ĿID1=" & strValue(0), "��ĿID2=" & strValue(1), "��ĿID3=" & strValue(2), "��ĿID4=" & strValue(3), "��ĿID5=" & strValue(4), "��ĿID6=" & strValue(5), bytStyle)
End Sub

Private Sub Menu_File_LabView(ByVal cbrControl As CommandBarControl)
'��ǩԤ��
On Error GoTo errHandle
    Call PrintSpeExamLabel(cbrControl)
    
    blnIsOk = True
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_File_LabPrint(ByVal cbrControl As CommandBarControl)
'��ǩ��ӡ
On Error GoTo errHandle
    Call PrintSpeExamLabel(cbrControl)
    
    blnIsOk = True
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub PrintWorkList(ByVal cbrControl As CommandBarControl)
'��ӡ�ؼ칤���б�
    Dim i As Long
    Dim j As Long
    Dim strValue(5) As String
    Dim bytStyle As Byte
    
    j = 0
    strValue(0) = "0": strValue(1) = "0": strValue(2) = "0": strValue(3) = "0": strValue(4) = "0": strValue(5) = "0"
    For i = 1 To ufgData.GridRows - 1
        If ufgData.GetCellCheckState(i, ufgData.GetColIndexWithRowCheck()) Then
            If zlCommFun.ActualLen(strValue(j)) > 2000 Then
                j = j + 1
                strValue(j) = ""
            End If

            If strValue(j) <> "" Then strValue(j) = strValue(j) & ","

            strValue(j) = strValue(j) & ufgData.KeyValue(i)
        End If
    Next i
    
    If cbrControl.ID = TMenuType.mtWorkView Then
        bytStyle = 1
    Else
        bytStyle = 2
    End If
    
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_10", Me, "��ĿID=" & strValue(0), "��ĿID1=" & strValue(1), "��ĿID2=" & strValue(2), "��ĿID3=" & strValue(3), "��ĿID4=" & strValue(4), "��ĿID5=" & strValue(5), bytStyle)
    
End Sub

Private Sub Menu_File_WorkView(ByVal cbrControl As CommandBarControl)
'Ԥ���ؼ칤���嵥
On Error GoTo errHandle
    
    Call PrintWorkList(cbrControl)
    
    blnIsOk = True
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_File_WorkPrint(ByVal cbrControl As CommandBarControl)
'��ӡ�ؼ칤���嵥
On Error GoTo errHandle
    
    Call PrintWorkList(cbrControl)
    
    blnIsOk = True
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Form_Initialize()
    Set zlReport = New zl9Report.clsReport
    mblnAutoAcceptOfAfterPrint = False
End Sub


Private Sub InitFilterPage()
    With tabFilter
        .RemoveAll
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .PaintManager.ColorSet.ButtonSelected = &HFFC0C0
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True
        .RemoveAll
        



        .InsertItem 0, "�����黯", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "0-�����黯"
'        .Item(tabFilter.ItemCount - 1).Visible = true
'        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
        
        .InsertItem 1, "����Ⱦɫ", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "1-����Ⱦɫ"
        
        .InsertItem 2, "���Ӳ���", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "2-���Ӳ���"

    End With
    
    tabFilter.Item(mlngFilterTabIndex).Selected = True
End Sub


Private Sub LoadFilterParameter()
    mlngFilterTabIndex = Val(zlDatabase.GetPara("�ؼ���������ҳ��", glngSys, glngModul, 0))
    chkYSQ.value = Val(zlDatabase.GetPara("�ؼ�����������", glngSys, glngModul, 1))
    chkYJS.value = Val(zlDatabase.GetPara("�ؼ������ѽ���", glngSys, glngModul, 0))
    chkYWC.value = Val(zlDatabase.GetPara("�ؼ����������", glngSys, glngModul, 0))
End Sub


Private Sub Form_Load()
On Error GoTo errHandle
    Dim curDate As Date
    
    Call InitCommandBars
    
    Call RestoreWinState(Me, App.ProductName)
    
    Call LoadFilterParameter
    
    Call InitFilterPage
    
    Call InitSpeExamWorkList
    
    curDate = zlDatabase.Currentdate
    dtpStartRequisition.value = Format(curDate, "yyyy-mm-dd 00:00")
    dtpEndRequisition.value = Format(curDate, "yyyy-mm-dd 23:59")
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
        Set cbrControl = .Add(xtpControlPopup, TMenuType.mtLab, "��ǩ"): cbrControl.IconId = 9023
            With cbrControl.CommandBar '�����˵�
                Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtLabView, "Ԥ��(0)"): cbrPopControl.IconId = 102
                Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtLabPrint, "��ӡ(1)"): cbrPopControl.IconId = 103
            End With
        Set cbrControl = .Add(xtpControlPopup, TMenuType.mtWork, "�嵥"): cbrControl.IconId = 3031
            With cbrControl.CommandBar '�����˵�
                Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtWorkView, "Ԥ��(0)"): cbrPopControl.IconId = 102
                Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtWorkPrint, "��ӡ(1)"): cbrPopControl.IconId = 103
            End With
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&Q)")
        cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)")
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAccept, "�ؼ۽���(&R)"): cbrControl.IconId = 747
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtComplete, "�ؼ����(&S)"): cbrControl.IconId = 3200
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
        Set cbrControl = .Add(xtpControlPopup, TMenuType.mtLab, "��ǩ"): cbrControl.IconId = 9023
            With cbrControl.CommandBar '�����˵�
                Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtLabView, "Ԥ��(0)"): cbrPopControl.IconId = 102
                Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtLabPrint, "��ӡ(1)"): cbrPopControl.IconId = 103
            End With
        Set cbrControl = .Add(xtpControlPopup, TMenuType.mtWork, "�嵥"): cbrControl.IconId = 3031
            With cbrControl.CommandBar '�����˵�
                Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtWorkView, "Ԥ��(0)"): cbrPopControl.IconId = 102
                Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtWorkPrint, "��ӡ(1)"): cbrPopControl.IconId = 103
            End With
            
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAccept, "�ؼ����"): cbrControl.IconId = 747
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtComplete, "�ؼ����"): cbrControl.IconId = 3200
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
        cbrControl.BeginGroup = True
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
End Sub

Private Sub SaveFilterParameter()
    Call zlDatabase.SetPara("�ؼ���������ҳ��", tabFilter.Selected.Index, glngSys, glngModul)
    Call zlDatabase.SetPara("�ؼ�����������", chkYSQ.value, glngSys, glngModul)
    Call zlDatabase.SetPara("�ؼ������ѽ���", chkYJS.value, glngSys, glngModul)
    Call zlDatabase.SetPara("�ؼ����������", chkYWC.value, glngSys, glngModul)
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    
    Call SaveFilterParameter
    
    Set zlReport = Nothing
End Sub




Private Sub UpdateWorkListPrintState()
'�ڴ�ӡ�󣬸��¹����嵥�Ĵ�ӡ״̬
    Dim strSql As String
    Dim i As Long
    Dim strPrintIds As String
        
    strPrintIds = ""
    For i = 1 To ufgData.GridRows - 1
        If ufgData.GetCellCheckState(i, ufgData.GetColIndexWithRowCheck()) Then
            strPrintIds = strPrintIds & "," & ufgData.KeyValue(i)

            strSql = "Zl_�����ؼ�_�嵥��ӡ(" & ufgData.KeyValue(i) & ")"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)

            ufgData.Text(i, gstrSpeExamWork_�嵥״̬) = "�Ѵ�ӡ"
        End If
    Next i

    '���µ�ǰ�����ؼ�״̬
    If Trim(strPrintIds) <> "" And Not (mufgGrid Is Nothing) Then
        strPrintIds = strPrintIds & ","

        For i = 1 To mufgGrid.GridRows - 1
            If UCase(strPrintIds) Like "*," & UCase(mufgGrid.KeyValue(i)) & ",*" Then

                Call mufgGrid.SyncText(i, gstrSpeExam_�嵥״̬, "�Ѵ�ӡ", True)
            End If
        Next i
    End If
End Sub






Private Sub ufgData_OnColsNameReSet()
On Error GoTo errHandle

   If ufgData.DataGrid.Rows > 1 Then Call GetSpeExamData
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub zlReport_AfterPrint(ByVal ReportNum As String)
'�嵥�Ѵ�ӡ
On Error GoTo errHandle
    Call UpdateWorkListPrintState
    
    If mblnAutoAcceptOfAfterPrint Then
        Call SpeExamBatAccept
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
