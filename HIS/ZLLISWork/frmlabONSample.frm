VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.Form frmlabONSample 
   Caption         =   "�걾����"
   ClientHeight    =   6150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11580
   Icon            =   "frmlabONSample.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   2535
      Left            =   3270
      TabIndex        =   0
      Top             =   1470
      Width           =   3255
      _Version        =   589884
      _ExtentX        =   5741
      _ExtentY        =   4471
      _StockProps     =   0
      BorderStyle     =   1
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      ShowItemsInGroups=   -1  'True
      AutoColumnSizing=   0   'False
   End
   Begin VB.PictureBox picFilter 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      ScaleHeight     =   585
      ScaleWidth      =   11265
      TabIndex        =   1
      Top             =   570
      Width           =   11295
      Begin VB.OptionButton optSave 
         Caption         =   "������"
         Height          =   195
         Index           =   0
         Left            =   8580
         TabIndex        =   4
         Top             =   210
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optSave 
         Caption         =   "�ѱ���"
         Height          =   195
         Index           =   1
         Left            =   9570
         TabIndex        =   3
         Top             =   210
         Width           =   1125
      End
      Begin VB.ComboBox cboMachine 
         Height          =   300
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   150
         Width           =   2025
      End
      Begin MSComCtl2.DTPicker DtpBegin 
         Height          =   285
         Left            =   4260
         TabIndex        =   5
         Top             =   165
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   99745795
         CurrentDate     =   39198
      End
      Begin MSComCtl2.DTPicker DTPEnd 
         Height          =   285
         Left            =   6420
         TabIndex        =   6
         Top             =   150
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   99745795
         CurrentDate     =   39198
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Left            =   6120
         TabIndex        =   9
         Top             =   210
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "���ʱ��:"
         Height          =   180
         Left            =   3390
         TabIndex        =   8
         Top             =   210
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "��������:"
         Height          =   180
         Left            =   210
         TabIndex        =   7
         Top             =   210
         Width           =   810
      End
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlabONSample.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlabONSample.frx":68BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlabONSample.frx":6E58
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlabONSample.frx":73F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlabONSample.frx":798C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlabONSample.frx":E1EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlabONSample.frx":14A50
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlabONSample.frx":1B2B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlabONSample.frx":21B14
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlabONSample.frx":28376
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   720
      Top             =   150
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmlabONSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mCol           '�б�
    ѡ�� = 0
    �걾��
    �걾����
    ��������
    ������Դ
    ���ʱ��
    �����
    ������
    ����ʱ��
    ����λ��
    ���滷��
    �걾id
End Enum

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_Manage_ThingModi               'ȫѡ
            Call RptSelect(Me.rptList.Records, True)
            Me.rptList.Populate
        Case conMenu_Manage_ThingDel                'ȫ��
            Call RptSelect(Me.rptList.Records, False)
            Me.rptList.Populate
        Case conMenu_Edit_Import                    '����
            Call SaveData
        Case conMenu_LIS_Cancel                     'ȡ������
            Call SaveData(1)
        Case conMenu_View_Refresh                   'ˢ��
            Call RefreshData
        Case conMenu_File_Exit                      '�˳�
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    Dim Column As ReportColumn
    Dim strSQL As String
    Dim rsTmp As New adodb.Recordset
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrthis.VisualTheme = xtpThemeOffice2003
    Me.cbrthis.Icons = zlCommFun.GetPubIcons
    With Me.cbrthis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbrthis.EnableCustomization False
    Me.cbrthis.ActiveMenuBar.Visible = False
    

    '�����
    With Me.cbrthis.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_Select
        .Add FCONTROL, Asc("Z"), conMenu_Edit_DeSelect
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F2, conMenu_Edit_Audit
        .Add FCONTROL, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F5, conMenu_View_Refresh
    End With

    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbrthis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ThingModi, "ȫѡ"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ThingDel, "ȫ��")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Import, "�洢"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_Cancel, "ȡ���洢")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�"): cbrControl.BeginGroup = True
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '�����
    With Me.cbrthis.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_Select
        .Add FCONTROL, Asc("Z"), conMenu_Edit_DeSelect
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F2, conMenu_Edit_Audit
        .Add FCONTROL, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F5, conMenu_View_Refresh
    End With
    
    DtpBegin = Now
    DTPEnd = Now
    
    On Error GoTo errH
    
    strSQL = "select id,����,���� from �������� "
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption)
    cboMachine.Clear
    cboMachine.AddItem "��������"
    cboMachine.ItemData(cboMachine.NewIndex) = 0
    Do Until rsTmp.EOF
        cboMachine.AddItem rsTmp("����") & "-" & rsTmp("����")
        cboMachine.ItemData(cboMachine.NewIndex) = rsTmp("ID")
        rsTmp.MoveNext
    Loop
    cboMachine.ListIndex = 0
    
    With Me.rptList.Columns
        
        rptList.AllowColumnRemove = False
        rptList.ShowItemsInGroups = False
        
        With rptList.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
        rptList.SetImageList ImgList
        Set Column = .Add(mCol.ѡ��, "ѡ��", 18, False): Column.Icon = 0
        Set Column = .Add(mCol.�걾��, "�걾��", 80, True)
        Set Column = .Add(mCol.��������, "��������", 65, True)
        Set Column = .Add(mCol.�걾����, "�걾����", 65, True)
        Set Column = .Add(mCol.���ʱ��, "���ʱ��", 80, True)
        Set Column = .Add(mCol.�����, "�����", 65, True)
        Set Column = .Add(mCol.������, "������", 65, True)
        Set Column = .Add(mCol.����ʱ��, "����ʱ��", 100, True)
        Set Column = .Add(mCol.����λ��, "����λ��", 65, True)
        Set Column = .Add(mCol.���滷��, "���滷��", 65, True)
        Set Column = .Add(mCol.�걾id, "�걾id", 65, True): Column.Visible = False
        Me.rptList.Populate
    End With
    Call RefreshData
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    With picFilter
        .Top = 460
        .Left = 10
        .Width = Me.ScaleWidth - 25
    End With
    With rptList
        .Top = picFilter.Top + picFilter.Height + 20
        .Left = 10
        .Width = Me.ScaleWidth - 25
        .Height = Me.ScaleHeight - .Top - 20
    End With
End Sub

Private Sub RefreshData()
    '����           ˢ������
    Dim rsTmp As New adodb.Recordset
    Dim Record As ReportRecord
    Dim intLoop As Integer, lngLoop As Long
    Dim cbrControl As CommandBarControl                 '�ı���ǩ
    Dim lngMachineID As Long                            '����ID
        
    On Error GoTo errH
    
    If DtpBegin.Value > DTPEnd.Value Then
        MsgBox "��ʼ���ڲ��ܴ��ڽ������ڣ�", vbInformation, gstrSysName
        DtpBegin.SetFocus
        Exit Sub
    End If
    
    gstrSql = "select /*+ RULE */ DISTINCT B.���ID AS ID,A.ҽ��id,F.���ͺ�,0 AS ѡ��," & _
             " Decode(A.����id, Null, " & vbCrLf & _
               " to_Char(Trunc(A.�걾���/10000)+1,'0000')|| '-'||to_Char(MOD(A.�걾���,10000),'0000'), to_number(A.�걾���)) As �걾��, " & _
             "A.�걾����," & _
             "TO_CHAR(A.���ʱ��,'MM-DD HH24:MI') AS ���ʱ��," & _
             "A.�����," & _
             "A.������," & _
             "TO_CHAR(B.����ʱ��,'MM-DD HH24:MI') AS ����ʱ��," & _
             "B.����ҽ�� AS ������," & _
             "C.���� AS �������," & _
             "E.���� AS ִ�п���," & _
             "A.id as �걾ID, " & _
             "B.����id, " & _
             "D.���� AS ��������,0 As ת��,Decode(A.�걾���,1,'��','') As ����, " & _
             "decode(a.���ʱ��,Null,'��','��') as �Ƿ����, " & _
             "Decode(a.����״̬, 1, '������', 2, '�Ѽ���') As ִ��״̬, " & _
             "Decode(a.�Ƿ���, 1, '', '����ʧ��') As ����, a.��ӡ����,a.΢����걾, " & _
             "a.����,a.�걾���,a.����ID,a.������Դ,a.Ӥ��,b.��������ID,a.������,b.��ҳID,������,����ʱ��,����λ��,���滷��  " & _
        "from ����걾��¼ A, ����ҽ����¼ B, ���ű� C, �������� D,���ű� E,����ҽ������ F,������Ϣ G " & _
        " WHERE A.ҽ��ID = B.���ID(+) AND B.��������ID = C.ID(+) AND B.ID=F.ҽ��id(+) AND " & _
             "A.����ID = D.ID(+) AND B.ִ�п���id = E.ID AND A.����״̬ IN (1,2) AND a.����ID = G.����ID and ����� is not null and ������ is null  "
    
    gstrSql = gstrSql & " and ����ʱ�� between [1] and [2] "
    
    If optSave(0).Value = True Then
        gstrSql = gstrSql & " and ������ is null "
        Set cbrControl = cbrthis.FindControl(, conMenu_Edit_Import, True, True)
        cbrControl.Enabled = True
        Set cbrControl = cbrthis.FindControl(, conMenu_LIS_Cancel, True, True)
        cbrControl.Enabled = False
    Else
        gstrSql = gstrSql & " and ������ is not null   "
        Set cbrControl = cbrthis.FindControl(, conMenu_Edit_Import, True, True)
        cbrControl.Enabled = False
        Set cbrControl = cbrthis.FindControl(, conMenu_LIS_Cancel, True, True)
        cbrControl.Enabled = True
    End If
    
    If cboMachine.ListIndex >= 0 Then
        If Val(cboMachine.ItemData(cboMachine.ListIndex)) > 0 Then
            gstrSql = gstrSql & " and ����ID = [3] "
            lngMachineID = cboMachine.ItemData(cboMachine.ListIndex)
        End If
    End If
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, CDate(Format(Me.DtpBegin, "yyyy-mm-dd 00:00:00")), _
                                         CDate(Format(Me.DTPEnd, "yyyy-mm-dd 23:23:59")), lngMachineID)
    
    
    
    With Me.rptList
        .Records.DeleteAll
        .Populate
        Do Until rsTmp.EOF
            Set Record = .Records.Add
            .Populate
            For intLoop = 0 To .Columns.Count
                Record.AddItem ""
            Next
            Record.Item(mCol.ѡ��).HasCheckbox = True
            Record.Item(mCol.�걾id).Value = Nvl(rsTmp("�걾ID"))
            Record.Item(mCol.�걾��).Value = Val(Nvl(rsTmp("�걾���")))
            Record.Item(mCol.�걾��).Caption = Trim(Nvl(rsTmp("�걾��")))
            Record.Item(mCol.�걾����).Value = Nvl(rsTmp("�걾����"))
            Record.Item(mCol.��������).Value = Nvl(rsTmp("����"))
            Record.Item(mCol.�����).Value = Nvl(rsTmp("�����"))
            Record.Item(mCol.���ʱ��).Value = Nvl(rsTmp("���ʱ��"))
            Record.Item(mCol.������).Value = Nvl(rsTmp("������"))
            Record.Item(mCol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
            Record.Item(mCol.����λ��).Value = Nvl(rsTmp("����λ��"))
            Record.Item(mCol.���滷��).Value = Nvl(rsTmp("���滷��"))
            Record.Item(mCol.�걾id).Value = Nvl(rsTmp("�걾id"))
            rsTmp.MoveNext
        Loop
        .Populate
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub RptSelect(Records As ReportRecords, blTrue As Boolean)
    '����                           ѡ���ȡ��ѡ��
    '����                           Records = �б����
    '                               blTrue  True = ѡ�� False = ȡ��ѡ��
    Dim intLoop As Integer
    
    For intLoop = 0 To Records.Count - 1
        Records(intLoop).Item(mCol.ѡ��).Checked = blTrue
    Next
End Sub

Private Sub optSave_Click(Index As Integer)
    Call RefreshData
End Sub

Private Sub rptList_MouseDown(Button As Integer, Shift As Integer, x As Long, Y As Long)
    Dim hitColumn As ReportColumn
    Dim Record As ReportRecord
    Dim blSelect As Boolean

    With Me.rptList
        Set hitColumn = .HitTest(x, Y).Column
        If Not hitColumn Is Nothing Then
            If hitColumn.Caption = "ѡ��" And .HitTest(x, Y).ht = xtpHitTestHeader Then
                If .Records.Count > 0 Then blSelect = Not .Records(0).Item(mCol.ѡ��).Checked
                For Each Record In .Records
                    Record.Item(mCol.ѡ��).Checked = blSelect
                Next
            End If
        End If
        .Populate
    End With
End Sub

Private Sub SaveData(Optional intType As Integer)
    '����           ����
    '����           intType 0=�걾���� 1=�걾ȡ������
    Dim strVal As String
    Dim astrVal() As String
    Dim strIDs As String
    Dim strSQL As String
    Dim intLoop As Integer
    
    On Error GoTo errH
    
    If CheckSel = False Then
        MsgBox "��һ���걾��û��ѡ���ܱ���!", vbInformation, "����걾"
        Exit Sub
    End If
    If intType = 0 Then
        '����
        strVal = frmlabONSampleEdit.ShowMe(Me)
        If strVal = "" Then Exit Sub
        
        '��ʼ��֯����
        astrVal = Split(strVal, "|")
        With Me.rptList
            For intLoop = 0 To .Records.Count - 1
                If .Records(intLoop).Item(mCol.ѡ��).Checked = True Then
                    strIDs = strIDs & "," & .Records(intLoop).Item(mCol.�걾id).Value
                End If
            Next
        End With
        
        If strIDs <> "" Then
            strIDs = Mid(strIDs, 2)
            If strIDs <> "" Then
                '����
                strSQL = "ZL_����걾����_edit(0,'" & strIDs & "','" & astrVal(0) & "','" & astrVal(1) & "','" & astrVal(2) & "')"
                zldatabase.ExecuteProcedure strSQL, "����걾"
            End If
        End If
    Else
        'ȡ������
        With Me.rptList
            For intLoop = 0 To .Records.Count - 1
                If .Records(intLoop).Item(mCol.ѡ��).Checked = True Then
                    strIDs = strIDs & "," & .Records(intLoop).Item(mCol.�걾id).Value
                End If
            Next
        End With
        If strIDs <> "" Then
            strIDs = Mid(strIDs, 2)
            If strIDs <> "" Then
                '����
                strSQL = "ZL_����걾����_edit(1,'" & strIDs & "','','','')"
                zldatabase.ExecuteProcedure strSQL, "����걾"
            End If
        End If
    End If
    Call RefreshData
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Function CheckSel() As Boolean
    '����           ��鵱ǰ�б��Ƿ���ѡ��ļ�¼
    Dim intLoop As Integer
    With Me.rptList
        For intLoop = 0 To .Records.Count - 1
            If .Records(intLoop).Item(mCol.ѡ��).Checked = True Then
                CheckSel = True
                Exit Function
            End If
        Next
    End With
End Function

