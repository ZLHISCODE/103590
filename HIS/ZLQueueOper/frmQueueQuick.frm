VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmQueueQuick 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ŶӺ���"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5265
   Icon            =   "frmQueueQuick.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   5265
   Begin VB.Frame frmLine 
      Height          =   30
      Left            =   0
      TabIndex        =   9
      Top             =   960
      Width           =   5295
   End
   Begin VB.PictureBox picCurPatient 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   5295
      TabIndex        =   7
      Top             =   480
      Width           =   5295
      Begin VB.Label lblCurrentPatient 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   600
         TabIndex        =   8
         Top             =   120
         Width           =   135
      End
      Begin VB.Image imgCurrent 
         Height          =   255
         Left            =   120
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.PictureBox picNextPatient 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   5295
      TabIndex        =   5
      Top             =   960
      Width           =   5295
      Begin VB.Label lblNextPatient 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   6
         Top             =   0
         Width           =   105
      End
      Begin VB.Image imgNext 
         Height          =   255
         Left            =   120
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDeptInfo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2640
      ScaleHeight     =   375
      ScaleWidth      =   2655
      TabIndex        =   2
      Top             =   0
      Width           =   2655
      Begin VB.ComboBox cboQueue 
         Height          =   300
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   60
         Width           =   1695
      End
      Begin VB.Label lblPeople 
         AutoSize        =   -1  'True
         Caption         =   "�ࣺxx��"
         Height          =   180
         Left            =   1800
         TabIndex        =   4
         Top             =   120
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList imgIcon 
      Left            =   5160
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQueueQuick.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQueueQuick.frx":D0B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picPassed 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   720
      ScaleHeight     =   375
      ScaleWidth      =   4575
      TabIndex        =   0
      Top             =   1320
      Width           =   4575
      Begin VB.PictureBox picHide 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   135
         TabIndex        =   10
         Top             =   0
         Width           =   135
      End
      Begin XtremeCommandBars.CommandBars cbrPassed 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.Label lblPassed 
      AutoSize        =   -1  'True
      Caption         =   "�ѹ��ţ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   780
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Bindings        =   "frmQueueQuick.frx":13916
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmQueueQuick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mobjQueueList As ReportControl
Private mobjCallList As ReportControl
Private mlngState As Long   '��ǰִ��״̬��0-�ȴ���1-˳����2-�غ���3-����
Private mlngCurID As Long   '��ǰ���в���ID
Private mlngCurIndex As Long    '��ǰ��������
Private mblnCallNext As Boolean  '��ǰ���к�����󣬲��ܼ�������
Private mlngCurQueue As Long  '��ǰ����
Private mstrCurDept As String   '��ǰ����

Private Const M_STR_PRODUCENAME = "ZL9PACSWork"

Public Event DoUnload()
Public Event SetPatientFocus(ByVal blnQueueList As Boolean, ByVal lngID As Long, ByVal strValue As String)
Public Event DoExecute(ByVal lngType As Long, ByVal lngID As Long, ByVal strValue As String, ByRef blnResult As Boolean)
Public Event IsQueueing(ByVal lngID As Long, ByRef blnResult As Boolean)

Public Function ShowQueueQuick(ower As Object, objQueueList As ReportControl, objCallList As ReportControl, strQueryNames As String)
    Me.Show , ower
    Me.ZOrder
    Call RefreshQueueQuick(objQueueList, objCallList, strQueryNames)
End Function

Public Function RefreshQueueQuick(objQueueList As ReportControl, objCallList As ReportControl, strQueryNames As String)
    Dim i As Long
    Dim arrQueueName() As String
    
    Set mobjQueueList = objQueueList
    Set mobjCallList = objCallList
    
    cboQueue.Clear
    
    If Len(strQueryNames) = 0 Then Exit Function
    
    arrQueueName = Split(strQueryNames, ",")
    
    For i = 0 To UBound(arrQueueName)
        If Len(arrQueueName(i)) > 0 Then
            cboQueue.AddItem Split(arrQueueName(i) & "-", "-")(1), i
        End If
    Next
    
    If UBound(arrQueueName) >= 0 Then
        mstrCurDept = Split(arrQueueName(0), "-")(0)
    Else
        mstrCurDept = ""
    End If
    
    cboQueue.ListIndex = mlngCurQueue
    
    mlngState = 0
    
    Call RefreshByQueueName
    Call RefreshPassedNum
End Function

Private Sub RefreshByQueueName()
    Dim lngNumIndex As Long
    Dim lngNameIndex As Long
    Dim lngSexIndex As Long
    Dim lngAgeIndex As Long
    Dim lngPositionIndex As Long
    Dim lngCurIndex As Long
    Dim lngQueueName As Long
    Dim lngIdIndex As Long
    
    mlngCurIndex = -1
    mlngCurID = 0
    
    lngNumIndex = GetColIndex("�ŶӺ���", mobjQueueList)
    lngNameIndex = GetColIndex("��������", mobjQueueList)
    lngIdIndex = GetColIndex("ID", mobjQueueList)
    lngSexIndex = GetColIndex("�Ա�", mobjQueueList)
    lngAgeIndex = GetColIndex("����", mobjQueueList)
    lngPositionIndex = GetColIndex("�����Ŀ", mobjQueueList)
    lngQueueName = GetColIndex("��������", mobjQueueList)
    '�ȴ����ж����ж�������Ϊ�����Ҷ��С�ʱ��ǰ����һ���ո��
    lngCurIndex = GetWillRowIndex("��������", IIf(cboQueue.Text = "���Ҷ���", " " & cboQueue.Text, mstrCurDept & "-" & cboQueue.Text), mobjQueueList)

    lblCurrentPatient.Caption = ""
    lblCurrentPatient.Caption = ""
    lblNextPatient.Caption = ""
    lblNextPatient.ToolTipText = ""
        
    '��������Ϣ���ص�����
    If lngCurIndex >= 0 Then
        With mobjQueueList
            If GetNextPatient(lngCurIndex) Then
                '��ʾ��ǰ������Ϣ
                lblCurrentPatient.Caption = "(" & Val(.Rows(lngCurIndex).Record(lngNumIndex).value) & ")" & .Rows(lngCurIndex).Record(lngNameIndex).value & " " & .Rows(lngCurIndex).Record(lngSexIndex).value & " " & .Rows(lngCurIndex).Record(lngAgeIndex).value & " " & .Rows(lngCurIndex).Record(lngPositionIndex).value
                lblCurrentPatient.ToolTipText = lblCurrentPatient.Caption
                mlngCurIndex = lngCurIndex
                
                mlngCurID = Val(.Rows(lngCurIndex).Record(lngIdIndex).value)
                RaiseEvent SetPatientFocus(True, mlngCurID, IIf(cboQueue.Text = "���Ҷ���", " " & cboQueue.Text, mstrCurDept & "-" & cboQueue.Text))
                
                '��ʾ��һ��������Ϣ
                lngCurIndex = lngCurIndex + 1
                If lngCurIndex < .Rows.Count Then
                    If GetNextPatient(lngCurIndex) Then
                        If .Rows(lngCurIndex).Record(lngQueueName).value = IIf(cboQueue.Text = "���Ҷ���", " " & cboQueue.Text, mstrCurDept & "-" & cboQueue.Text) Then
                            lblNextPatient.Caption = "(" & Val(.Rows(lngCurIndex).Record(lngNumIndex).value) & ")" & .Rows(lngCurIndex).Record(lngNameIndex).value & " " & .Rows(lngCurIndex).Record(lngSexIndex).value & " " & .Rows(lngCurIndex).Record(lngAgeIndex).value & " " & .Rows(lngCurIndex).Record(lngPositionIndex).value
                            lblNextPatient.ToolTipText = lblNextPatient.Caption
                        End If
                    End If
                End If
            End If
        End With
    End If
    
    Call RefreshInfoShow
    Call RefreshSurPlusPeople
End Sub

Private Sub RefreshInfoShow()
    
    If Len(lblCurrentPatient.Caption) = 0 Then
        imgCurrent.Visible = False
        With lblCurrentPatient.Font
            .Bold = False
            .Size = 10
        End With
        lblCurrentPatient.Caption = "��ǰ���С�" & cboQueue.Text & "��û�пɺ��еĻ���"
        lblCurrentPatient.ToolTipText = ""
    Else
        imgCurrent.Visible = True
        With lblCurrentPatient.Font
            .Bold = True
            .Size = 11
        End With
    End If

    If lblCurrentPatient.Width + imgCurrent.Width < picCurPatient.Width - 100 Then
        imgCurrent.Left = picCurPatient.Width / 2 - (imgCurrent.Width + lblCurrentPatient.Width + 100) / 2
    Else
        imgCurrent.Left = 0
    End If

    lblCurrentPatient.Left = imgCurrent.Left + imgCurrent.Width + 100
    
    imgCurrent.Top = picCurPatient.Height / 2 - imgCurrent.Height / 2
    lblCurrentPatient.Top = picCurPatient.Height / 2 - lblCurrentPatient.Height / 2
    
    If Len(lblNextPatient.Caption) = 0 Then
        imgNext.Visible = False
    Else
        imgNext.Visible = True
    End If
    
    If lblNextPatient.Width + imgNext.Width < picNextPatient.Width - 100 Then
        imgNext.Left = picNextPatient.Width / 2 - (lblNextPatient.Width + imgNext.Width + 100) / 2
    Else
        imgNext.Left = 0
    End If
    lblNextPatient.Left = imgNext.Left + imgNext.Width + 100
    
    
    imgNext.Top = picNextPatient.Height / 2 - imgNext.Height / 2
    lblNextPatient.Top = picNextPatient.Height / 2 - lblNextPatient.Height / 2
End Sub


'��ȡ��һ���ɺ��в��˵�λ��
Private Function GetNextPatient(ByRef lngIndex As Long) As Boolean
    Dim i As Long
    Dim lngStateIndex As Long
    Dim lngIdIndex As Long
    Dim blnCheckResult As Boolean
    
    lngStateIndex = GetColIndex("�Ŷ�״̬", mobjQueueList)
    lngIdIndex = GetColIndex("ID", mobjQueueList)
     
    GetNextPatient = False
    
    With mobjQueueList
        For i = lngIndex To .Rows.Count - 1
            If .Rows(i).GroupRow Then Exit For
            If .Rows(i).Record(lngStateIndex).value = "�Ŷ���" Then
                RaiseEvent IsQueueing(.Rows(i).Record(lngIdIndex).value, blnCheckResult)
                If blnCheckResult Then
                    lngIndex = i
                    GetNextPatient = True
                    Exit For
                End If
            End If
        Next
    End With
End Function

'��ǰ����ʣ������
Private Sub RefreshSurPlusPeople()
    Dim lngCurIndex As Long
    Dim i As Long
    Dim lngSurPlusPeople As Long
    Dim lngStateIndex As Long
    
    lngSurPlusPeople = 0

    '�ȴ����ж����ж�������Ϊ�����Ҷ��С�ʱ��ǰ����һ���ո��
    lngCurIndex = GetWillRowIndex("��������", IIf(cboQueue.Text = "���Ҷ���", " " & cboQueue.Text, mstrCurDept & "-" & cboQueue.Text), mobjQueueList)
    lngStateIndex = GetColIndex("�Ŷ�״̬", mobjQueueList)
    
    If lngCurIndex >= 0 Then
        With mobjQueueList
            For i = lngCurIndex To .Rows.Count - 1
                If .Rows(i).GroupRow Then Exit For
                If .Rows(i).Record(lngStateIndex).value = "�Ŷ���" Then
                    lngSurPlusPeople = lngSurPlusPeople + 1
                End If
            Next
        End With
    End If
    
    lblPeople.Caption = "�ࣺ" & lngSurPlusPeople & "��"
    
    If mlngState > 0 And lngSurPlusPeople = 1 Or lngSurPlusPeople = 0 Then
        mblnCallNext = False
    Else
        mblnCallNext = True
    End If
End Sub

Private Sub cboQueue_Click()
    On Error GoTo errHandle
    
    mlngCurQueue = cboQueue.ListIndex
    Call RefreshByQueueName
    Call RefreshPassedNum
    
    mlngState = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strCurQueue As String
    Dim blnResult As Boolean
    
    On Error GoTo errHandle
    
    strCurQueue = IIf(cboQueue.Text = "���Ҷ���", " " & cboQueue.Text, mstrCurDept & "-" & cboQueue.Text)
    
    blnResult = False
    Select Case Control.Id
        Case conMenu_Queue_CallNext    '˳��
            Call RefreshByQueueName
            
            RaiseEvent DoExecute(1, mlngCurID, strCurQueue, blnResult)
            
            If Not blnResult Then
                Exit Sub
            End If
            
            If mlngState > 0 And mlngState < 3 Then
                Call RefreshPassedNum
            End If
            
            mlngState = 1
        Case conMenu_Queue_Broadcast    '�غ�
            
            RaiseEvent DoExecute(2, mlngCurID, strCurQueue, blnResult)

            If Not blnResult Then
                Exit Sub
            End If
            
            mlngState = 2
            
        Case conMenu_Queue_RecDiagnose  '����
        
            RaiseEvent DoExecute(3, mlngCurID, strCurQueue, blnResult)

            If Not blnResult Then
                Exit Sub
            End If
            
            mlngState = 3
    End Select
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo errHandle
    
    Select Case Control.Id
        Case conMenu_Queue_CallNext '˳��
            Control.Enabled = mblnCallNext
            
        Case conMenu_Queue_Broadcast  '�غ�
            Control.Enabled = mlngState > 0
            
        Case conMenu_Queue_RecDiagnose  '����
            Control.Enabled = mlngState > 0
    End Select
    
    If mlngState > 0 Then
        lblCurrentPatient.ForeColor = &H8000&
    Else
        lblCurrentPatient.ForeColor = &H0&
    End If

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cbrPassed_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnResult  As Boolean

    blnResult = False
    RaiseEvent DoExecute(3, Val(Control.Category), IIf(cboQueue.Text = "���Ҷ���", " " & cboQueue.Text, mstrCurDept & "-" & cboQueue.Text), blnResult)
    
    If Not blnResult Then
        Exit Sub
    End If
    Call RefreshPassedNum
End Sub

Private Sub Form_Load()
    Call InitPosition
    Call InitCommandBars
    Call SetFont

    imgCurrent.Picture = imgIcon.ListImages(1).Picture
    imgCurrent.ToolTipText = "��ǰ���л���"
    imgNext.Picture = imgIcon.ListImages(2).Picture
    imgNext.ToolTipText = "��һλ�����еĻ���"
End Sub


Private Sub SetFont()
    With lblCurrentPatient.Font
        .Name = "����"
        .Bold = True
        .Size = 11
    End With
    
    With lblNextPatient.Font
        .Name = "����"
        .Bold = False
        .Size = 10
    End With
End Sub

Private Sub InitCommandBars()
'���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    
    Set Me.cbrMain.Icons = zlCommFun.GetPubIcons 'imgPublic.Icons
    
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 16, 16
    End With

    Me.cbrMain.ActiveMenuBar.Visible = False
    
    '����������
    Set cbrToolBar = Me.cbrMain.Add("������", xtpBarTop)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    cbrToolBar.ContextMenuPresent = False
    
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_CallNext, "˳��"): cbrControl.IconId = 744: cbrControl.ToolTipText = "˳��"
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_RecDiagnose, "����"): cbrControl.IconId = 3009: cbrControl.ToolTipText = "����"
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Broadcast, "�غ�"): cbrControl.IconId = 745: cbrControl.ToolTipText = "�غ�"
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next

    cbrPassed.VisualTheme = xtpThemeOfficeXP
    
    Set cbrPassed.Icons = zlCommFun.GetPubIcons 'imgPublic.Icons
    
    With cbrPassed.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 16, 16
    End With

    cbrPassed.EnableCustomization False
    cbrPassed.ActiveMenuBar.Visible = False
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    picCurPatient.Top = Me.ScaleTop + 400
'    picCurPatient.Height = (Me.ScaleHeight - 800 - 30) / 2
    
    frmLine.Top = picCurPatient.Top + picCurPatient.Height
    
    picNextPatient.Top = frmLine.Top + frmLine.Height
    picNextPatient.Height = picCurPatient.Height
    
    picHide.Left = 0
    picHide.Top = 0
    picHide.Height = picPassed.Height
    picHide.Width = 100
End Sub

Private Sub InitPosition()
    Me.Left = Val(GetSetting("ZLSOFT", "����ģ��\" & M_STR_PRODUCENAME & "\�Ŷӽк�", "X", Screen.Width - Me.Width))
    Me.Top = Val(GetSetting("ZLSOFT", "����ģ��\" & M_STR_PRODUCENAME & "\�Ŷӽк�", "Y", Screen.Height - Me.Height * 1.5))
End Sub

Private Sub Form_Unload(Cancel As Integer)

    SaveSetting "ZLSOFT", "����ģ��\" & M_STR_PRODUCENAME & "\�Ŷӽк�", "X", Me.Left
    SaveSetting "ZLSOFT", "����ģ��\" & M_STR_PRODUCENAME & "\�Ŷӽк�", "Y", Me.Top
    
    RaiseEvent DoUnload
End Sub

Public Sub UnloadMe()
    
    Unload Me
End Sub

'�������Ű�ť
Private Sub RefreshPassedNum()
    Dim cbrToolBar As CommandBar
    Dim cbrControl As CommandBarControl
    Dim lngNumIndex As Long
    Dim lngNameIndex As Long
    Dim lngCurIndex As Long
    Dim lngIdIndex As Long
    Dim lngStateIndex As Long
    Dim lngCount As Long
    Dim strInfo As String
    
    lngCount = 0
    
    lngNumIndex = GetColIndex("�ŶӺ���", mobjCallList)
    lngNameIndex = GetColIndex("��������", mobjCallList)
    lngStateIndex = GetColIndex("�Ŷ�״̬", mobjCallList)
    lngIdIndex = GetColIndex("ID", mobjCallList)
    lngCurIndex = GetWillRowIndex("��������", IIf(cboQueue.Text = "���Ҷ���", " " & cboQueue.Text, mstrCurDept & "-" & cboQueue.Text), mobjCallList)
    
    cbrPassed.DeleteAll
    
    Set cbrToolBar = cbrPassed.Add("����", xtpBarTop)
    cbrToolBar.EnableDocking xtpFlagHideWrap
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.ContextMenuPresent = False
    
    If lngCurIndex >= 0 Then
        For i = lngCurIndex To mobjCallList.Rows.Count - 1
            If mobjCallList.Rows(i).GroupRow Then Exit For
            
            If mlngCurID <> mobjCallList.Rows(i).Record(lngIdIndex).value And mobjCallList.Rows(i).Record(lngStateIndex).value <> "������" Then
                strInfo = "(" & mobjCallList.Rows(i).Record(lngNumIndex).value & ")" & mobjCallList.Rows(i).Record(lngNameIndex).value
                If LenB(strInfo) > 40 Then
                    strInfo = Mid(strInfo, 1, 16) & "..."
                End If
                Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, conMenu_Queue_Passed * 10# + 1, strInfo): cbrControl.IconId = 3009
                cbrControl.ToolTipText = "(" & mobjCallList.Rows(i).Record(lngNumIndex).value & ")" & mobjCallList.Rows(i).Record(lngNameIndex).value & "  ����"
                cbrControl.Category = mobjCallList.Rows(i).Record(lngIdIndex).value
                cbrControl.Style = xtpButtonIconAndCaption
                cbrControl.BeginGroup = True
                lngCount = lngCount + 1
            End If
        Next
    End If
    
    If lngCount = 0 Then
        Me.Height = 2160 - picPassed.Height
        cbrToolBar.Visible = False
    Else
        Me.Height = 2160
        cbrToolBar.Visible = True
    End If
End Sub

