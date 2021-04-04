VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPatholSlices_WorkPrint 
   Caption         =   "��Ƭ��������"
   ClientHeight    =   7470
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   12240
   Icon            =   "frmPatholSlices_WorkPrint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   12240
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picTag 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11520
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   18
      Top             =   5880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   600
      ScaleHeight     =   6735
      ScaleWidth      =   11175
      TabIndex        =   0
      Top             =   360
      Width           =   11175
      Begin VB.Frame framFilter 
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   11055
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
            Left            =   8160
            TabIndex        =   9
            Top             =   240
            Width           =   1455
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
            Left            =   6360
            TabIndex        =   8
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
            Left            =   9720
            TabIndex        =   7
            Top             =   200
            Width           =   1215
         End
         Begin VB.OptionButton optMaterialTime 
            Caption         =   "ȡ��"
            Height          =   255
            Left            =   360
            TabIndex        =   6
            Top             =   165
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optRequisitionTime 
            Caption         =   "����"
            Height          =   255
            Left            =   360
            TabIndex        =   5
            Top             =   390
            Width           =   735
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Left            =   3720
            TabIndex        =   10
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   99155971
            CurrentDate     =   40679.726087963
         End
         Begin MSComCtl2.DTPicker dtpStart 
            Height          =   300
            Left            =   1560
            TabIndex        =   11
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   99155971
            CurrentDate     =   40679.0594097222
         End
         Begin VB.Label Label3 
            Caption         =   "��"
            Height          =   255
            Left            =   7920
            TabIndex        =   15
            Top             =   285
            Width           =   255
         End
         Begin VB.Label Label2 
            Caption         =   "��         ʱ�䣺"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   280
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "��"
            Height          =   255
            Left            =   3480
            TabIndex        =   13
            Top             =   285
            Width           =   255
         End
         Begin VB.Label labPatholNum 
            Caption         =   "����ţ�"
            Height          =   255
            Left            =   5640
            TabIndex        =   12
            Top             =   285
            Width           =   720
         End
      End
      Begin VB.CheckBox chkYWC 
         Caption         =   "�����"
         Height          =   180
         Left            =   10080
         TabIndex        =   3
         ToolTipText     =   "��ʾ��Ƭ״̬Ϊ��δ��ɡ�����Ƭ��¼��"
         Top             =   5910
         Width           =   855
      End
      Begin VB.CheckBox chkYJS 
         Caption         =   "�ѽ���"
         Height          =   180
         Left            =   9120
         TabIndex        =   2
         ToolTipText     =   "��ʾ��Ƭ״̬Ϊ���ѽ��ܡ�����Ƭ��¼��"
         Top             =   5910
         Width           =   855
      End
      Begin VB.CheckBox chkWCL 
         Caption         =   "δ����"
         Height          =   255
         Left            =   8160
         TabIndex        =   1
         ToolTipText     =   "��ʾ��Ƭ״̬Ϊ��δ��������Ƭ��¼��"
         Top             =   5880
         Width           =   855
      End
      Begin XtremeSuiteControls.TabControl tabFilter 
         Height          =   375
         Left            =   0
         TabIndex        =   16
         Top             =   840
         Width           =   11085
         _Version        =   589884
         _ExtentX        =   19553
         _ExtentY        =   661
         _StockProps     =   64
      End
      Begin zl9PACSWork.ucFlexGrid ufgData 
         Height          =   4215
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   7435
         DefaultCols     =   ""
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   19
      Top             =   7110
      Width           =   12240
      _ExtentX        =   21590
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatholSlices_WorkPrint.frx":179A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14711
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
Attribute VB_Name = "frmPatholSlices_WorkPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents zlReport As zl9Report.clsReport
Attribute zlReport.VB_VarHelpID = -1

Private mlngPatholAdviceId As Long
Private mufgParGrid As ucFlexGrid

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
    mtchkWCL = 5
    mtchkYJS = 6
    mtchkYWC = 7
End Enum

Private Sub RefreshSilcesCount()
'ˢ����Ƭ��¼����
    Dim i As Long
    Dim lngRecord As Long
    Dim lngTotal As Long
    Dim lngSlices As Long
    
    On Error GoTo errH
    
    lngTotal = 0
    lngSlices = 0
    
    
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.RowHidden(i) Then
            If Not ufgData.IsNullRow(i) Then

                lngTotal = lngTotal + Val(ufgData.Text(i, gstrSlices_��Ƭ��))

                If ufgData.Text(i, gstrSlices_��ǰ״̬) <> "�����" Then
                    lngSlices = lngSlices + Val(ufgData.Text(i, gstrSlices_��Ƭ��))
                End If
            End If
        End If
    Next i
    
    stbThis.Panels(2).Text = "��Ƭ������" & lngTotal & "    ����Ƭ����" & lngSlices
    Exit Sub
errH:
    Call MsgBoxD(Me, err.Description, vbOKOnly, Me.Caption)
End Sub



Public Sub ShowWorkPrint(ufgGrid As ucFlexGrid, ByVal lngPatholAdviceId As Long, owner As Form)
'��ʾ�����嵥��ӡ����
    Set mufgParGrid = ufgGrid

    mlngPatholAdviceId = lngPatholAdviceId
    blnIsOk = False
        
'    '���뵱ǰ�����Ƭ����
'    If Trim(lngPatholAdviceId) > 0 Then
'        Call LoadSpecifySlicesData
'    End If
    
    Call RefreshSilcesCount
    
    Call Me.Show(1, owner)
    
End Sub


Private Sub GetSlicesData()
    Dim strSql As String
    Dim strPatholNumQuery As String


    strPatholNumQuery = ""
    If Trim(txtStartPatholNum.Text) <> "" And Trim(txtEndPatholNum.Text) <> "" Then
        strPatholNumQuery = " and (REGEXP_SUBSTR(upper(c.�����), '[[:alpha:]]+') >=REGEXP_SUBSTR(upper([3]),'[[:alpha:]]+') and to_number(REGEXP_SUBSTR(upper(c.�����), '[[:digit:]]+')) >=to_number(REGEXP_SUBSTR(upper([3]),  '[[:digit:]]+'))) "
        strPatholNumQuery = strPatholNumQuery & " and  (REGEXP_SUBSTR(upper(c.�����), '[[:alpha:]]+') <=REGEXP_SUBSTR(upper([4]),'[[:alpha:]]+') and to_number(REGEXP_SUBSTR(upper(c.�����),  '[[:digit:]]+')) <=to_number(REGEXP_SUBSTR(upper([4]), '[[:digit:]]+'))) "
    ElseIf Trim(txtStartPatholNum.Text) <> "" Then
        strPatholNumQuery = " and upper(c.�����)=upper([3]) "
    ElseIf Trim(txtStartPatholNum.Text) <> "" Then
        strPatholNumQuery = " and upper(c.�����) =upper([4]) "
    End If
    
    
    strSql = "select a.Id,c.�����,c.����ҽ��ID, e.����, g.���� as �ű�����, a.�Ŀ�ID,b.���,b.ȡ��λ��, d.�걾����, d.�걾����, a.��Ƭ����, a.��Ƭ��ʽ,a.��Ƭ��,b.ȡ��ʱ��, a.��ǰ״̬,a.�嵥״̬ " & _
                " from ������Ƭ��Ϣ a, ����ȡ����Ϣ b, ��������Ϣ c, ����걾��Ϣ d, ����ҽ����¼ e ,���������� g" & _
                IIf(optRequisitionTime.value, ",����������Ϣ f ", "") & _
                " Where a.�Ŀ�id = b.�Ŀ�id And b.����ҽ��ID = c.����ҽ��ID and b.ȷ��״̬=1 And c.ҽ��ID = e.ID And b.�걾id = d.�걾id and c.�������ID=g.ID" & _
                IIf(optMaterialTime.value, " and b.ȡ��ʱ�� between [1] and [2]", "") & _
                IIf(optRequisitionTime.value, " and a.����ID=f.����ID and f.����ʱ�� between [1] and [2]", "") & _
                IIf(strPatholNumQuery <> "", strPatholNumQuery, "") & _
                " order by c.�����,a.��ǰ״̬,b.���,a.Id "
    'If mblnMoved Then strSql = GetMovedDataSql(strSql)
                

    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                            CDate(dtpStart.value), _
                                            CDate(dtpEnd.value), _
                                            txtStartPatholNum.Text, _
                                            txtEndPatholNum.Text)
                                            
                                                                    
    Call FilterSlicesData
End Sub



Private Sub FilterSlicesData()
'���˲�ѯ��������Ƭ����
    Dim strFilters As String
    Dim strStudyTypeFilter As String
    
    strFilters = ""
    
    
    strStudyTypeFilter = ""
    Select Case tabFilter.Selected.tag
        Case "����"
            strStudyTypeFilter = ""
        Case Else
            strStudyTypeFilter = "�ű�����=" & "'" & tabFilter.Selected.tag & "'"
    End Select
    
        
    If chkWCL.value <> 0 Then
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
    If chkWCL.value = 0 And chkYJS.value = 0 And chkYWC.value = 0 Then
        strFilters = "(" & IIf(strStudyTypeFilter = "", "", strStudyTypeFilter & " and ") & " ��ǰ״̬=0)" & " or " & _
                     "(" & IIf(strStudyTypeFilter = "", "", strStudyTypeFilter & " and ") & " ��ǰ״̬=1)" & " or " & _
                     "(" & IIf(strStudyTypeFilter = "", "", strStudyTypeFilter & " and ") & " ��ǰ״̬=2)"
    End If
    
    If ufgData.AdoData Is Nothing Then Exit Sub
    
    ufgData.AdoData.Filter = strFilters
    ufgData.RefreshData
    
    Call RefreshSilcesCount
End Sub



Private Sub LoadSpecifySlicesData()
'����ָ���Ĳ�����Ƭ����
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select a.ID, a.Id,c.�����,c.����ҽ��ID, f.���� as �ű����� ,e.����, c.�������ID, a.�Ŀ�ID,b.���,b.ȡ��λ��, d.�걾����, d.�걾����, a.��Ƭ����, a.��Ƭ��ʽ, a.��Ƭ��,a.��ǰ״̬,a.�嵥״̬ " & _
                 " from ������Ƭ��Ϣ a, ����ȡ����Ϣ b, ��������Ϣ c, ����걾��Ϣ d, ����ҽ����¼ e,���������� f " & _
                " Where a.�Ŀ�id = b.�Ŀ�id And b.����ҽ��ID = c.����ҽ��ID And c.ҽ��ID = e.ID And b.�걾id = d.�걾id " & _
                " and c.����ҽ��ID=[1] and a.��ǰ״̬ <> 2 and c.�������ID=f.ID" & _
                " order by �ű�����,�����,�Ŀ�ID,��ǰ״̬ "
    'If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngPatholAdviceId)
    
    Call ufgData.RefreshData
End Sub


Private Sub InitSlicesWorkList()
'��ʼ����Ƭ�嵥��ʾ��
    Dim strTemp As String
    
    '��������
    ufgData.GridRows = glngStandardRowCount
    '�����и�
    ufgData.RowHeightMin = glngStandardRowHeight
	ufgData.DefaultColNames = gstrSlicesWorkCols
    
    '�ж����ݿ�������Ƿ������� �����ȡ���ݿ����  û�������Ĭ��
    strTemp = zlDatabase.GetPara("������Ƭ�б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
     
    If strTemp = "" Then
        '��ʼ���걾��ʾ�б�
        ufgData.ColNames = gstrSlicesWorkCols
    Else
        ufgData.ColNames = strTemp
    End If
    
    
    ufgData.ColConvertFormat = gstrSlicesWorkConvertFormat
    ufgData.IsShowPopupMenu = False
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Select Case control.ID
        Case TMenuType.mtLabView                    '��ǩԤ��
            Call Menu_File_LabView(control)
        
        Case TMenuType.mtLabPrint                   '��ǩ��ӡ
            Call Menu_File_LabPrint(control)
            
        Case TMenuType.mtWorkView                   '�嵥Ԥ��
            Call Menu_File_WorkView(control)
        
        Case TMenuType.mtWorkPrint                  '�嵥��ӡ
            Call Menu_File_WorkPrint(control)
        
        Case TMenuType.mtAccept                     '��Ƭ����
            Call Menu_Edit_Accept
        
        Case TMenuType.mtComplete                   '��Ƭ���
            Call Menu_Edit_Complete

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
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_File_Exit()
    blnIsOk = False
    Me.Hide
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
    picMain.Left = Left
    picMain.Top = Top
    picMain.Width = Right - Left
    picMain.Height = Bottom - Top
End Sub

Private Sub picMain_Resize()
On Error Resume Next
    framFilter.Left = 120
    framFilter.Top = 0
    framFilter.Width = picMain.Width - 240
    
    tabFilter.Left = 120
    tabFilter.Top = framFilter.Top + framFilter.Height
    tabFilter.Width = picMain.Width - chkWCL.Width * 5 - 720
    
    chkWCL.Left = tabFilter.Width
    chkWCL.Top = tabFilter.Top + 40
    
    chkYJS.Left = chkWCL.Left + chkWCL.Width + 240
    chkYJS.Top = chkWCL.Top + 40
    
    chkYWC.Left = chkYJS.Left + chkYJS.Width + 240
    chkYWC.Top = chkYJS.Top
    
    ufgData.Left = 120
    ufgData.Top = tabFilter.Top + tabFilter.Height
    ufgData.Width = picMain.Width - 240
    ufgData.Height = picMain.Height - framFilter.Height - tabFilter.Height
End Sub

Private Sub ufgData_OnColFormartChange()
'�����б�����
    zlDatabase.SetPara "������Ƭ�б�����", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
End Sub

Private Sub SlicesBatAccept()
'�ؼ���������
    Dim i As Long
    Dim curPatholAdviceID As String
    Dim strSql As String
    Dim blnUpdateCallWind As Boolean
    
    blnUpdateCallWind = False
    
    For i = 1 To ufgData.GridRows - 1
        '���ѡ�еļ�飬�Ž��н���
        If ufgData.GetCellCheckState(i, ufgData.GetColIndexWithRowCheck()) Then
            If curPatholAdviceID <> ufgData.Text(i, gstrSlicesWork_����ҽ��ID) Then
                curPatholAdviceID = ufgData.Text(i, gstrSlicesWork_����ҽ��ID)
                
                strSql = "Zl_������Ƭ_����(" & Val(curPatholAdviceID) & ",'" & UserInfo.���� & "')"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            End If
            
            '���µ�ǰ�б�״̬
            If ufgData.Text(i, gstrSlicesWork_��ǰ״̬) = "δ����" Then
                ufgData.Text(i, gstrSlicesWork_��ǰ״̬) = "�ѽ���"
            End If
            
            '���µ��ý����б�״̬
            If Val(curPatholAdviceID) = mlngPatholAdviceId Then
                blnUpdateCallWind = True
            End If
        End If
    Next i
    
    If blnUpdateCallWind And Not (mufgParGrid Is Nothing) Then
        For i = 1 To mufgParGrid.GridRows - 1
            If mufgParGrid.Text(i, gstrSlicesWork_��ǰ״̬) = "δ����" Then
                mufgParGrid.Text(i, gstrSlices_��ǰ״̬) = "�ѽ���"
                mufgParGrid.Text(i, gstrSlices_��Ƭ��) = UserInfo.����
            End If
        Next i
    End If
End Sub




Private Sub SlicesBatSure()
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
            If curPatholAdviceID <> ufgData.Text(i, gstrSlicesWork_����ҽ��ID) Then
                curPatholAdviceID = ufgData.Text(i, gstrSlicesWork_����ҽ��ID)
                
                strSql = "Zl_������Ƭ_ȷ��(" & Val(curPatholAdviceID) & "," & zlStr.To_Date(dtServicesTime) & ")"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            End If
            
            '���µ�ǰ�б�״̬
            If ufgData.Text(i, gstrSlicesWork_��ǰ״̬) = "�ѽ���" Then
                ufgData.Text(i, gstrSlicesWork_��ǰ״̬) = "�����"
            End If
            
            '���µ��ý����б�״̬
            If Val(curPatholAdviceID) = mlngPatholAdviceId Then
                blnUpdateCallWind = True
            End If
        End If
    Next i
    
    If blnUpdateCallWind And Not (mufgParGrid Is Nothing) Then
        For i = 1 To mufgParGrid.GridRows - 1
            If mufgParGrid.Text(i, gstrSlicesWork_��ǰ״̬) = "�ѽ���" Then
                mufgParGrid.Text(i, gstrSlices_��ǰ״̬) = "�����"
                mufgParGrid.Text(i, gstrSlices_��Ƭ��) = UserInfo.����
            End If
        Next i
    End If
End Sub



Private Function CheckAllowSureOrAccept(Optional ByVal blnIsSure As Boolean = True) As Boolean
'�ж��Ƿ���Ҫ���к���
    Dim i As Long
    
    CheckAllowSureOrAccept = False
    For i = 1 To ufgData.GridRows - 1
        If ufgData.GetRowCheck(i) = True And (ufgData.Text(i, gstrSlices_��ǰ״̬) = IIf(blnIsSure, "�ѽ���", "δ����")) Then
            CheckAllowSureOrAccept = True
            Exit Function
        End If
    Next i
End Function


Private Sub chkWCL_Click()
On Error GoTo ErrHandle
    If Not tabFilter.Visible Then Exit Sub
    
    Call FilterSlicesData
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkYJS_Click()
On Error GoTo ErrHandle
    If Not tabFilter.Visible Then Exit Sub
    
    Call FilterSlicesData
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkYWC_Click()
On Error GoTo ErrHandle
    If Not tabFilter.Visible Then Exit Sub
    
    Call FilterSlicesData
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Edit_Accept()
'��Ƭ����
On Error GoTo ErrHandle
    If Not CheckAllowSureOrAccept(False) Then
        Call MsgBoxD(Me, "������Ҫ���ܵ���Ƭ��Ϣ��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call SlicesBatAccept
    
    blnIsOk = True
    
    Call MsgBoxD(Me, "����ɶ���ѡ���Ľ��ܴ���", vbOKOnly, Me.Caption)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Edit_Complete()
'��Ƭ���
On Error GoTo ErrHandle
    If Not CheckAllowSureOrAccept(True) Then
        Call MsgBoxD(Me, "������Ҫ��ɵ���Ƭ��Ϣ��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call SlicesBatSure
    
    blnIsOk = True
    
    Call MsgBoxD(Me, "����ɶ���ѡ������Ƭ����", vbOKOnly, Me.Caption)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdFilter_Click()
On Error GoTo ErrHandle
    Call GetSlicesData
    
    Call RefreshSilcesCount
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub PrintSlicesLabel(ByVal cbrControl As CommandBarControl)
'��ӡԤ���ؼ���Ŀ��ǩ
    Dim i As Long
    Dim j As Long
    Dim strValue(5) As String
    
    Dim strSliceId As String
    Dim k As Long
    Dim lngCount As Long
    Dim bytStyle As Byte
    
    j = 0
    strValue(0) = "0": strValue(1) = "0": strValue(2) = "0": strValue(3) = "0": strValue(4) = "0": strValue(5) = "0"
    For i = 1 To ufgData.GridRows - 1
        If ufgData.GetCellCheckState(i, ufgData.GetColIndexWithRowCheck()) Then
            If zlCommFun.ActualLen(strValue(j)) > 2000 Then
                j = j + 1
                strValue(j) = ""
            End If

            strSliceId = ufgData.KeyValue(i)
            lngCount = Val(ufgData.Text(i, gstrSlices_��Ƭ��))
    
            If strValue(j) <> "" Then strValue(j) = strValue(j) & ","
    
            strValue(j) = strValue(j) & strSliceId
            
            If lngCount > 1 Then
                For k = 1 To lngCount - 1
                    strValue(j) = strValue(j) & "," & strSliceId
                Next k
            End If
        End If
    Next i
    
    If cbrControl.ID = TMenuType.mtLabView Then
        bytStyle = 1
    Else
        bytStyle = 2
    End If
    
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_09", Me, "��ƬID1=" & strValue(0), "��ƬID2=" & strValue(1), "��ƬID3=" & strValue(2), "��ƬID4=" & strValue(3), "��ƬID5=" & strValue(4), "��ƬID6=" & strValue(5), bytStyle)
End Sub

Private Sub Menu_File_LabView(ByVal cbrControl As CommandBarControl)
'��ǩԤ��
On Error GoTo ErrHandle
    Call PrintSlicesLabel(cbrControl)
    
    blnIsOk = True
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_File_LabPrint(ByVal cbrControl As CommandBarControl)
'��ǩ��ӡ
On Error GoTo ErrHandle
    Call PrintSlicesLabel(cbrControl)
    
    blnIsOk = True
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub




Private Sub PrintWorkList(ByVal cbrControl As CommandBarControl)
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
    
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_08", Me, "��ƬID1=" & strValue(0), "��ƬID2=" & strValue(1), "��ƬID3=" & strValue(2), "��ƬID4=" & strValue(3), "��ƬID5=" & strValue(4), "��ƬID6=" & strValue(5), bytStyle)
    
End Sub

Private Sub Menu_File_WorkView(ByVal cbrControl As CommandBarControl)
On Error GoTo ErrHandle
    
    Call PrintWorkList(cbrControl)
    
    blnIsOk = True
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_File_WorkPrint(ByVal cbrControl As CommandBarControl)
On Error GoTo ErrHandle
    
    Call PrintWorkList(cbrControl)
    
    blnIsOk = True
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Initialize()
    Set zlReport = New zl9Report.clsReport
    
    mblnAutoAcceptOfAfterPrint = False
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
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAccept, "��Ƭ����(&R)"): cbrControl.IconId = 747
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtComplete, "��Ƭ���(&S)"): cbrControl.IconId = 3200
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
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAccept, "��Ƭ����"): cbrControl.IconId = 747
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtComplete, "��Ƭ���"): cbrControl.IconId = 3200
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
        cbrControl.BeginGroup = True
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
End Sub

Private Sub InitFilterPage()
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim i As Long
    
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
        
        strSql = "select ID,���� from ����������"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ò�������")

        If rsData.RecordCount > 0 Then
            
            rsData.MoveFirst
        
            For i = 0 To rsData.RecordCount - 1
                If NVL(rsData!����, "  ") <> "  " Then
                    .InsertItem i, rsData!����, picTag.hWnd, 0
                    .Item(tabFilter.ItemCount - 1).tag = rsData!����
                End If
                rsData.MoveNext
            Next
            
            .InsertItem rsData.RecordCount, "��  ��", picTag.hWnd, 0
            .Item(tabFilter.ItemCount - 1).tag = "����"
            
        End If
        
    End With
    
    tabFilter.Item(mlngFilterTabIndex).Selected = True
End Sub


Private Sub LoadFilterParameter()
    mlngFilterTabIndex = Val(zlDatabase.GetPara("��Ƭ��������ҳ��", glngSys, glngModul, 0))
    chkWCL.value = Val(zlDatabase.GetPara("��Ƭ����δ����", glngSys, glngModul, 1))
    chkYJS.value = Val(zlDatabase.GetPara("��Ƭ�����ѽ���", glngSys, glngModul, 0))
    chkYWC.value = Val(zlDatabase.GetPara("��Ƭ���������", glngSys, glngModul, 0))
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Dim curDate As Date
    
    Call InitCommandBars
    
    Call RestoreWinState(Me, App.ProductName)
    
    Call LoadFilterParameter
    
    Call InitFilterPage
    
    '��ʼ�������б�
    Call InitSlicesWorkList
    
    curDate = zlDatabase.Currentdate
    
    dtpStart.value = Format(curDate - 1, "yyyy-mm-dd 00:00")
    dtpEnd.value = Format(curDate - 1, "yyyy-mm-dd 23:59")
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub SaveFilterParameter()
    Call zlDatabase.SetPara("��Ƭ��������ҳ��", tabFilter.Selected.Index, glngSys, glngModul)
    Call zlDatabase.SetPara("��Ƭ����δ����", chkWCL.value, glngSys, glngModul)
    Call zlDatabase.SetPara("��Ƭ�����ѽ���", chkYJS.value, glngSys, glngModul)
    Call zlDatabase.SetPara("��Ƭ���������", chkYWC.value, glngSys, glngModul)
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    
    Call SaveFilterParameter
    
    Set zlReport = Nothing
End Sub

Private Sub UpdateSlicesPrintState()
'�ڴ�ӡ�󣬽��ܴ�ӡ������Ƭ����
    Dim strSql As String
    Dim i As Long
    Dim strPrintIds As String
        
    strPrintIds = ""
    For i = 1 To ufgData.GridRows - 1
        If ufgData.GetCellCheckState(i, ufgData.GetColIndexWithRowCheck()) Then
            strPrintIds = strPrintIds & "," & ufgData.KeyValue(i)
            
            strSql = "Zl_������Ƭ_�嵥��ӡ(" & ufgData.KeyValue(i) & ")"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            
            ufgData.Text(i, gstrSlices_�嵥״̬) = "�Ѵ�ӡ"
        End If
    Next i
    
    '���µ�ǰ������Ƭ��¼״̬
    If Trim(strPrintIds) <> "" And Not (mufgParGrid Is Nothing) Then
        strPrintIds = strPrintIds & ","

        For i = 1 To mufgParGrid.GridRows - 1
            If UCase(strPrintIds) Like "*," & UCase(mufgParGrid.KeyValue(i)) & ",*" Then

                mufgParGrid.Text(i, gstrSpeExam_�嵥״̬) = "�Ѵ�ӡ"
            End If
        Next i
    End If
End Sub



Private Sub tabFilter_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error GoTo ErrHandle
    If Not tabFilter.Visible Then Exit Sub
    
    Call FilterSlicesData
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnColsNameReSet()
On Error GoTo ErrHandle

    If ufgData.DataGrid.Rows > 1 Then Call GetSlicesData
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub zlReport_AfterPrint(ByVal ReportNum As String)
'�嵥�Ѵ�ӡ
On Error GoTo ErrHandle
    '���������Ƭ�嵥��ӡ����ֱ���˳�
    If ReportNum <> "ZL1_PATHOLSLICES_01" Then Exit Sub
    
    Call UpdateSlicesPrintState
    
    '��ӡ���Զ�����
    If mblnAutoAcceptOfAfterPrint Then
        Call SlicesBatAccept
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

