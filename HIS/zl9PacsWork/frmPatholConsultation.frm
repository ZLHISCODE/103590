VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmPatholConsultation 
   Caption         =   "�������"
   ClientHeight    =   7980
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   9135
   Icon            =   "frmPatholConsultation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   9135
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6975
      Left            =   0
      ScaleHeight     =   6975
      ScaleWidth      =   9135
      TabIndex        =   0
      Top             =   480
      Width           =   9135
      Begin VB.Frame framRequisition 
         Caption         =   "�����¼"
         Height          =   6855
         Left            =   240
         TabIndex        =   1
         Top             =   0
         Width           =   8655
         Begin VB.TextBox txtAdvice 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   5400
            Width           =   8175
         End
         Begin VB.TextBox txtResult 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   3840
            Width           =   8175
         End
         Begin zl9PACSWork.ucFlexGrid ufgData 
            Height          =   3135
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   5530
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
         Begin VB.Label labAdvice 
            Caption         =   "���������"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   5160
            Width           =   1095
         End
         Begin VB.Label labResult 
            Caption         =   "��������"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   3480
            Width           =   1095
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   7620
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatholConsultation.frx":179A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9234
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
Attribute VB_Name = "frmPatholConsultation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mlngCurAdviceId As Long     '��ǰҽ��ID
Private mstrPrivs As String         '��ǰȨ�޴�
Private mblnMoved As Boolean        '�Ƿ�ת��

Private mlngCurDepartmentId As Long

Private mrecStudyInf As TStudyStateInf
Private mblnIsDoFeedback As Boolean

Private mblnDataModifyState As Boolean
Private mblnViewState As Boolean
Private mblnFeedBackState As Boolean

Private Enum TMenuType
    mtFeedback = 1      '����
    mtCancle = 2        '����
    mtView = 3          '����
    
    mtAddCon = 4        '��ӻ���
    mtDelCon = 5        'ɾ������
End Enum

Public Sub zlRefresh(lngAdviceID As Long, ByVal blnReadOnly As Boolean, _
    strPrivs As String, ByVal blnMoved As Boolean, _
    ByVal lngCurDepartmentId As Long, ByVal blnIsDoFeedback As Boolean, Optional owner As Form = Nothing)
'�������

    If lngAdviceID <= 0 Then
        Call ConfigConsultationFace(False, "ҽ��ID��Ч���顣")
        Exit Sub
    End If
    
'    If mlngCurAdviceId = lngAdviceId Then Exit Sub

    mlngCurAdviceId = lngAdviceID
    mstrPrivs = strPrivs
    mblnMoved = blnMoved
    mlngCurDepartmentId = lngCurDepartmentId
    mblnIsDoFeedback = blnIsDoFeedback
    
    
    '���ô��ڱ���
    If mblnIsDoFeedback Then
        Me.Caption = "�������-����"
    Else
        Me.Caption = "�������-����"
    End If
    
    Call GetPatholStudyState(lngAdviceID, mrecStudyInf)
        
   
    If mrecStudyInf.strPatholNumber = "" Then
        Call ConfigConsultationFace(False, "�ü����δ������Ч�Ĳ���ţ���ȷ�ϸü���Ƿ��ѱ����ա�")
        
        If Not (owner Is Nothing) Then
            Call MsgBoxD(Me, "�ü����δ������Ч�Ĳ���ţ���ȷ�ϸü���Ƿ��ѱ����ա�", vbOKOnly, Me.Caption)
        End If
        
        Exit Sub
    Else
        Call ConfigConsultationFace(True)
    End If
    
    '�����������
    Call LoadConsultationData
    
    '����Ȩ��
    Call ConfigPopedom(blnReadOnly)
    
    '�����ӵ���ߣ��򵯳�����
    If Not (owner Is Nothing) Then
        Call Me.Show(0, owner)
    End If
End Sub


Private Sub ConfigPopedom(ByVal blnIsReadOnly As Boolean)
'����Ȩ��
    Dim blnIsAllowConRequest As Boolean
    Dim blnIsAllowConFeedback As Boolean
    
    blnIsAllowConRequest = CheckPopedom(mstrPrivs, "��������")
    blnIsAllowConFeedback = CheckPopedom(mstrPrivs, "���ﷴ��")
    
    mblnDataModifyState = blnIsAllowConRequest And Not mblnIsDoFeedback And Not blnIsReadOnly
    mblnViewState = blnIsAllowConRequest And Not mblnIsDoFeedback And Not blnIsReadOnly
    mblnFeedBackState = blnIsAllowConFeedback And mblnIsDoFeedback And Not blnIsReadOnly

    ufgData.ReadOnly = blnIsReadOnly
End Sub


Private Sub ConfigConsultationFace(ByVal blnIsValid As Boolean, Optional ByVal strHintInf As String = "")
'�����ؼ����

    mblnDataModifyState = blnIsValid
    mblnViewState = blnIsValid
    mblnFeedBackState = blnIsValid
    
    If blnIsValid Then
        Call ufgData.CloseHintInf
    Else
        Call ufgData.ShowHintInf(strHintInf)
    End If
End Sub



Private Sub LoadConsultationData()
'����������ݵ��б�
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select Id, ����ҽʦ, ���ﵥλ, ����ҽʦ, ��������, ����ʱ��, ��ֹʱ��, �������,��Ͻ��,������,��ע,��ǰ״̬, ���ʱ�� " & _
            " from ���������Ϣ where ����ҽ��ID=[1] order by ��ǰ״̬,����ʱ��,��ֹʱ��,���ʱ��"
            
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
        
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mrecStudyInf.lngPatholAdviceId)
    
    Call ufgData.RefreshData
    
    If ufgData.ShowingDataRowCount > 0 Then
        Call LoadConContext(1)
    End If
End Sub



Private Sub InitConsultationList()
'��ʼ��������ʾ�б�
    Dim strTemp As String
    
   
    
    '�ж����ݿ�������Ƿ������� �����ȡ���ݿ����  û�������Ĭ��
    strTemp = zlDatabase.GetPara("��������б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
    ufgData.DefaultColNames = gstrConsultationCols
     
    If strTemp = "" Then
        ufgData.ColNames = gstrConsultationCols
    Else
        ufgData.ColNames = strTemp
    End If
     '��������
    ufgData.GridRows = glngStandardRowCount
    '�����и�
    ufgData.RowHeightMin = glngStandardRowHeight
    ufgData.ColConvertFormat = gstrConsultationConvertFormat
    ufgData.IsShowPopupMenu = False
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrorHand
    Select Case control.ID
        Case TMenuType.mtFeedback                   '����
            Call Menu_Edit_Feedback
        
        Case TMenuType.mtCancle                     '����
            Call Menu_Edit_Cancel
        
        Case TMenuType.mtView                       '����
            Call Menu_Edit_View
            
        Case TMenuType.mtAddCon                     '��ӻ���
            Call Menu_Edit_AddCon
        
        Case TMenuType.mtDelCon                     'ɾ������
            Call Menu_Edit_DelCon
        
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

Private Sub cbrMain_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
On Error Resume Next
    picBack.Left = Left
    picBack.Top = Top
    picBack.Width = Right - Left
    picBack.Height = Bottom - Top
End Sub

Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrorHand
    Select Case control.ID
        Case TMenuType.mtFeedback               '����
            control.Enabled = mblnFeedBackState
            
        Case TMenuType.mtCancle                 '����
            control.Enabled = mblnDataModifyState And ufgData.IsSelectionRow
            
        Case TMenuType.mtView                   '����
            control.Enabled = mblnViewState And ufgData.IsSelectionRow
            
        Case TMenuType.mtAddCon                 '��ӻ���
            control.Enabled = mblnDataModifyState
        
        Case TMenuType.mtDelCon                 'ɾ������
            control.Enabled = mblnDataModifyState
            
    End Select
    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Exit Sub
End Sub

Private Sub ufgData_OnColFormartChange()
'�����б����
     zlDatabase.SetPara "��������б�����", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
End Sub

Private Sub picBack_Resize()
'�������沼��
On Error Resume Next
    framRequisition.Left = 0
    framRequisition.Top = 60
    framRequisition.Width = picBack.Width
    framRequisition.Height = picBack.Height - 60
    
    ufgData.Left = 120
    ufgData.Top = 240
    ufgData.Width = framRequisition.Width - 240
    ufgData.Height = framRequisition.Height - txtResult.Height - txtAdvice.Height - labResult.Height * 2 - 840
    
    labResult.Left = 120
    labResult.Top = ufgData.Top + ufgData.Height + 240
    
    txtResult.Left = 120
    txtResult.Top = labResult.Top + labResult.Height + 60
    txtResult.Width = ufgData.Width
    
    labAdvice.Left = 120
    labAdvice.Top = txtResult.Top + txtResult.Height + 120
    
    txtAdvice.Left = 120
    txtAdvice.Top = labAdvice.Top + labAdvice.Height + 60
    txtAdvice.Width = ufgData.Width
End Sub


Private Sub ShowNewConsultationWindow()
'��ʾ������������
Dim frmConsultation As New frmPatholConsultation_New
On Error GoTo errFree
    Call frmConsultation.ShowConsultationWindow(ufgData, mrecStudyInf.lngPatholAdviceId, mlngCurDepartmentId, Me)
errFree:
    Call Unload(frmConsultation)
    Set frmConsultation = Nothing
End Sub


Private Sub Menu_Edit_AddCon()
'��ӻ���
On Error GoTo errHandle

'    If mlngStudyProcedure <> TStudyProcedure.spDiagnose Then
'        Call MsgBoxD(Me, "��ǰ����ִ�й��̴��ڷ���Ͻ׶Σ����ܽ��л������롣", vbOKOnly, Me.Caption)
'        Exit Sub
'    End If
    
    Call ShowNewConsultationWindow
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Function CheckAllowUpdateConsultation(ByVal lngConsultationRow As Long) As Boolean
'����Ƿ�������»����¼
    CheckAllowUpdateConsultation = IIf(ufgData.Text(lngConsultationRow, gstrConsultation_��ǰ״̬) <> "�Ѳ���" And ufgData.Text(lngConsultationRow, gstrConsultation_��ǰ״̬) <> "�ѷ���", True, False)
    If Not CheckAllowUpdateConsultation Then
        Call MsgBoxD(Me, "�û����ѷ������Ѳ��ģ�����ִ�д˲�����", vbOKOnly, Me.Caption)
        Exit Function
    End If
    
    CheckAllowUpdateConsultation = IIf(ufgData.Text(lngConsultationRow, gstrConsultation_����ҽʦ) = UserInfo.����, True, False)
    If Not CheckAllowUpdateConsultation Then
        Call MsgBoxD(Me, "�û���ֻ��������ҽʦ [" & ufgData.Text(lngConsultationRow, gstrConsultation_����ҽʦ) & "] �����޸ġ�", vbOKOnly, Me.Caption)
        Exit Function
    End If
End Function


Private Sub DelConsultationData(ByVal lngConsultationRow As Long)
'ɾ�������¼
    Dim strSql As String
    
    strSql = "Zl_�������_ɾ��(" & ufgData.KeyValue(lngConsultationRow) & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Call ufgData.DelRow(lngConsultationRow, False)
End Sub


Private Sub CancelConsultationFinish(ByVal lngConsultationRow As Long)
'�����������
    Dim strSql As String
    
    strSql = "Zl_�������_״̬(" & ufgData.KeyValue(lngConsultationRow) & ",1)"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    ufgData.Text(lngConsultationRow, gstrConsultation_��ǰ״̬) = "�ѳ���"

End Sub


Private Function CheckAllowCancelConsultation(ByVal lngConsultationRow As Long) As Boolean
'����Ƿ�����������
    CheckAllowCancelConsultation = IIf(ufgData.Text(lngConsultationRow, gstrConsultation_����ҽʦ) = UserInfo.����, True, False)
    If Not CheckAllowCancelConsultation Then
        Call MsgBoxD(Me, "�û���ֻ��������ҽʦ [" & ufgData.Text(lngConsultationRow, gstrConsultation_����ҽʦ) & "] �����޸ġ�", vbOKOnly, Me.Caption)
        Exit Function
    End If
End Function



Private Sub Menu_Edit_Cancel()
'�������
On Error GoTo errHandle
    If Not ufgData.IsSelectionRow Then Exit Sub
    
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ҫ�����Ļ����¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '����Ƿ�������
    If Not CheckAllowCancelConsultation(ufgData.SelectionRow) Then Exit Sub
    
    If MsgBoxD(Me, "ȷ��Ҫ�����û����¼��", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    '�����������
    Call CancelConsultationFinish(ufgData.SelectionRow)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Edit_DelCon()
'ɾ�������¼
On Error GoTo errHandle
    If Not ufgData.IsSelectionRow Then Exit Sub
    
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ҫɾ���Ļ����¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If Not CheckAllowUpdateConsultation(ufgData.SelectionRow) Then Exit Sub
    
    If MsgBoxD(Me, "ȷ��Ҫɾ���û����¼��", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    'ɾ�������¼
    Call DelConsultationData(ufgData.SelectionRow)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub ShowConsultationFeedback(ByVal lngConsultationRow As Long)
'��ʾ���ﷴ������
Dim frmFeedback As New frmPatholConsultation_Feedback
On Error GoTo errFree
    Call frmFeedback.ShowFeedbackWindow(ufgData, Val(ufgData.KeyValue(lngConsultationRow)), mlngCurDepartmentId, Me)
    
    
errFree:
'    ���ﴰ�ڵ���ʾʹ�÷�ģ̬���ڣ�������ﲻ�ܽ����ͷ�
'    Call Unload(frmFeedback)
'    Set frmFeedback = Nothing
End Sub


Private Sub Menu_Edit_Feedback()
'���ﷴ��
On Error GoTo errHandle
    If Not ufgData.IsSelectionRow Then Exit Sub
    
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ҫ�����Ļ����¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '�жϵ�ǰ�����û��Ƿ�Ϊ�ü�¼�Ļ���ҽ��
    If UserInfo.���� <> ufgData.Text(ufgData.SelectionRow, gstrConsultation_����ҽʦ) Then
        Call MsgBoxD(Me, "��ǰ�û���ü�¼�Ļ���ҽʦ��ͬ�����ܷ�������ѡ�����ҽʦ���� [ " & UserInfo.���� & "] �Ļ����¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call ShowConsultationFeedback(ufgData.SelectionRow)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Function CheckAllowViewConsultation(ByVal lngConsultationRow As Long)
'����Ƿ�����鿴�����¼
    CheckAllowViewConsultation = IIf(ufgData.Text(lngConsultationRow, gstrConsultation_��ǰ״̬) <> "������", True, False)
    
    If Not CheckAllowViewConsultation Then
        Call MsgBoxD(Me, "�û�����δ���������ܲ��ġ�", vbOKOnly, Me.Caption)
        Exit Function
    End If
    
    CheckAllowViewConsultation = IIf(ufgData.Text(lngConsultationRow, gstrConsultation_����ҽʦ) = UserInfo.����, True, False)
    If Not CheckAllowViewConsultation Then
        Call MsgBoxD(Me, "�û���ֻ��������ҽʦ [" & ufgData.Text(lngConsultationRow, gstrConsultation_����ҽʦ) & "] �����޸ġ�", vbOKOnly, Me.Caption)
        Exit Function
    End If
End Function



Private Sub ShowFeedbackViewWindow(ByVal lngConsultationRow As Long)
'��ʾ���ﷴ������
'Dim frmFeedbackView As New frmPatholConsultation_Feedback
'On Error GoTo errFree
'    Call frmFeedbackView.ShowFeedbackViewWindow(ufgData, Me)
    
    '�޸Ļ����¼״̬
    Call ViewConsultation(lngConsultationRow)
'errFree:
''    Call Unload(frmFeedback)
''    Set frmFeedback = Nothing
End Sub



Private Sub ViewConsultation(ByVal lngConsultationRow As Long)
'�������
    Dim strSql As String

    strSql = "Zl_�������_״̬(" & ufgData.KeyValue(lngConsultationRow) & ",3)"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)

    ufgData.Text(lngConsultationRow, gstrConsultation_��ǰ״̬) = "�Ѳ���"
End Sub


Private Sub Menu_Edit_View()
'�鿴����
On Error GoTo errHandle
    If Not ufgData.IsSelectionRow Then Exit Sub
    
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ҫ�鿴�Ļ����¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '����Ƿ�����鿴�����¼
    If Not CheckAllowViewConsultation(ufgData.SelectionRow) Then Exit Sub
    
    Call ShowFeedbackViewWindow(ufgData.SelectionRow)
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub Form_Load()
On Error GoTo errHandle
    Call InitCommandBars
    
    Call RestoreWinState(Me, App.ProductName)
    
    '�ô���ʹ�õķ�ģʽ������ʾ�������Ҫ��ǰ
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3 '�������ö�
    
    Call InitConsultationList
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
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtFeedback, "����(&F)"): cbrControl.IconId = 9022
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtCancle, "����(&C)"): cbrControl.IconId = 3014
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtView, "����(&S)"): cbrControl.IconId = 225

        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAddCon, "��ӻ���(&A)"): cbrControl.IconId = 4112
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtDelCon, "ɾ������(&D)"): cbrControl.IconId = 4114
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
'        Set cbrControl = .Controls.Add(xtpControlButton, conMenu_Help_About, "���ڡ�(A)")
    End With
    '---------------------����������------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtFeedback, "����"): cbrControl.IconId = 9022
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtCancle, "����"): cbrControl.IconId = 3014
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtView, "����"): cbrControl.IconId = 225

        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAddCon, "��ӻ���"): cbrControl.IconId = 4112
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtDelCon, "ɾ������"): cbrControl.IconId = 4114
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
        cbrControl.BeginGroup = True
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
End Sub

Private Sub ClearConContext()
'���������ʾ����
    txtAdvice.Text = ""
    txtResult.Text = ""
End Sub


Private Sub LoadConContext(ByVal lngRow As Long)
'���ػ��ﱨ������
    txtResult.Text = ufgData.Text(lngRow, gstrConsultation_��Ͻ��)
    txtAdvice.Text = ufgData.Text(lngRow, gstrConsultation_������)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
End Sub


Private Sub ufgData_OnClick()
On Error GoTo errHandle
    If Not ufgData.IsSelectionRow Then Exit Sub
    
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        Call ClearConContext
    Else
        Call LoadConContext(ufgData.SelectionRow)
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub ufgData_OnColsNameReSet()
On Error GoTo errHandle

    Call LoadConsultationData
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnDblClick()
'���Ļ��ﷴ��
On Error GoTo errHandle
    Call ViewFeedback
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ViewFeedback()
'��ʾ���ﷴ������
Dim frmFeedbackView As New frmPatholConsultation_Feedback
On Error GoTo errFree
    If Not ufgData.IsSelectionRow Then Exit Sub
    
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ҫ�鿴�Ļ����¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call frmFeedbackView.ShowFeedbackViewWindow(ufgData, Me)

errFree:
'    Call Unload(frmFeedback)
'    Set frmFeedback = Nothing
End Sub


