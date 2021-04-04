VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmMessageRead 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����Ϣ"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1890
   Icon            =   "frmMessageRead.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   1890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer timMessage 
      Interval        =   1000
      Left            =   630
      Top             =   1140
   End
   Begin VB.Image imgTipClose 
      Height          =   270
      Left            =   180
      Picture         =   "frmMessageRead.frx":6852
      Top             =   1800
      Width           =   810
   End
   Begin VB.Image imgMail 
      Height          =   480
      Left            =   210
      Picture         =   "frmMessageRead.frx":7084
      Top             =   450
      Width           =   480
   End
   Begin VB.Image imgClose 
      Height          =   240
      Left            =   840
      Picture         =   "frmMessageRead.frx":7D4E
      Top             =   225
      Width           =   240
   End
   Begin XtremeSuiteControls.PopupControl popMsg 
      Left            =   1305
      Top             =   1410
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   4
   End
End
Attribute VB_Name = "frmMessageRead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnIcon As Boolean     'ͼ���Ѿ���ʾ
Dim mdatLast As Date        '���һ����ʾ֪ͨͼ���ʱ��
Private mlngPeriod As Long   '��Ϣ�������

'------------------------------------------------------------------------------------------
'Popupcontrol �ؼ���ʾ�õ���API
Const IDOK = 1
Const IDCLOSE = 2
Const IDSITE = 3

Private Sub Form_Load()
    mdatLast = zlDatabase.Currentdate()
    
    If Val(zlDatabase.GetPara("��¼����ʼ���Ϣ")) = 1 Then
        'ֻҪ��δ����Ϣ������
        mdatLast = CDate("1900-01-01")
        
        mlngPeriod = Val(zlDatabase.GetPara("�ʼ���Ϣ�������"))
        If (mlngPeriod < 10 Or mlngPeriod > 300) And mlngPeriod <> 0 Then mlngPeriod = 60
    End If
    mblnIcon = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'ֻ���X���д�������ˡ�����MouseMove��һ��ͨ���¼�
    '����¼���������ƶ������,X��ֵΪ"1E00";����¼��ǰ�����Ҽ������,X��ֵΪ"1E3C"......
    If Hex(X) <> "1E0F" Then Exit Sub '����������
    
    If gblnMessageShow = False Then
        '����Ϣ�շ�����
        With frmMessageManager
            On Error Resume Next
            .Show , gfrmMain
            If Err.Number <> 0 Then
                ShowWindow .hWnd, SW_SHOWNORMAL
                BringWindowToTop .hWnd
                SetActiveWindow .hWnd
            End If
        End With
    Else
        With frmMessageManager
            .mlngIndexPre = -1
            If .mlngIndex = 1 Then
               .FillList
            Else
                .lblICO_MouseDown 1, 1, 0, 0, 0
            End If
            On Error Resume Next
            .Show , gfrmMain
            If Err.Number <> 0 Then
                ShowWindow .hWnd, SW_SHOWNORMAL
                BringWindowToTop .hWnd
                SetActiveWindow .hWnd
            End If
        End With
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call RemoveIcon
End Sub

Private Sub popMsg_ItemClick(ByVal item As XtremeSuiteControls.IPopupControlItem)
    
    Dim frmTmp As Form, lngStyle As Long
    On Error Resume Next
    If item.Id = IDCLOSE Or item.Id = IDOK Then
        popMsg.Close
    End If

    If item.Id = IDSITE Then
        popMsg.Close
        frmMessageManager.Show , gfrmMain
        If Err.Number = 401 Then
            Unload frmMessageManager
            Call PopShow(1, "���ڲ��ܴ���Ϣ�Ķ����壬���ȹرյ�ǰ���壡")
        ElseIf Err.Number <> 0 Then
            If ErrCenter() = 1 Then
                Resume
            End If
        End If
    End If

End Sub

Private Sub timMessage_Timer()
    Static lngTimes As Long
    
    lngTimes = lngTimes + 1
    If lngTimes >= mlngPeriod Then
        lngTimes = 0
        Call UpdateNotify
    End If
End Sub

Public Sub UpdateNotify()
'����֪ͨ��Ϣ
    Dim rsTemp As New ADODB.Recordset
    Dim datTemp As Date
    Dim str���� As String
    
    On Error GoTo ErrH
    '����δ����Ϣ�����ֵ
    gstrSQL = "select A.ʱ��,A.���� From zlmessages A, (select max(A.ʱ��) as ʱ�� " & _
              "  from zlmessages A, zlmsgstate B " & _
              "  where A.ID=B.��ϢID and B.����=2 and B.ɾ��=0 and substr(B.״̬,1,1)='0' and B.�û�=[1]) B where A.ʱ��=B.ʱ�� "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, gstrDbUser)
    
    If IsNull(rsTemp("ʱ��")) Then
        Call RemoveIcon
    Else
        Do Until rsTemp.EOF
            datTemp = rsTemp("ʱ��")
            str���� = "" & rsTemp!����
            rsTemp.MoveNext
        Loop
        If datTemp > mdatLast Then
            '��ʾ�����ʼ�����
            If gblnMessageShow = True Then
                If frmMessageManager.mlngIndex = 1 Then
                    frmMessageManager.mlngIndexPre = -1
                    frmMessageManager.FillList    'ֱ��ˢ���б�
                End If
            End If
            If Len(str����) > 20 Then str���� = Mid(str����, 1, 17) & "..."
            Call PopShow(0, str����)
            Call AddIcon
            mdatLast = datTemp        '�����ʱ����Ϊ���һ�ε�
        ElseIf datTemp < mdatLast Then
            '��ʾ���ʼ��Ѿ�����
            Call RemoveIcon
            
        End If
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub AddIcon()
    Dim t As NOTIFYICONDATA
    
    If mblnIcon = False Then
        t.cbSize = Len(t)
        t.hWnd = Me.hWnd   '�¼�����������
        t.uId = 1&
        t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        t.ucallbackMessage = WM_MOUSEMOVE
        t.hIcon = Me.Icon
        t.szTip = "�����µġ�δ�򿪹�����Ϣ����" & Chr$(0)
        Shell_NotifyIcon NIM_ADD, t
        Beep
    End If
    mblnIcon = True
End Sub

Private Sub RemoveIcon()
    Dim t As NOTIFYICONDATA
    t.cbSize = Len(t)
    t.hWnd = Me.hWnd   '�¼�����������
    t.uId = 1&
    Shell_NotifyIcon NIM_DELETE, t
    mblnIcon = False
End Sub

Private Sub PopShow(ByVal lngType As Long, ByVal strMsg As String)
    '��ʾ��Ϣ��ʾ
    On Error Resume Next
    
    popMsg.Animation = 2 'pop������ֶ��� 0-None(��) 1-Fade(����)��2-Slide(����) 3-Unfold(չ��)
    popMsg.AnimateDelay = 500 '������ʱ ms
    popMsg.ShowDelay = 5000 '��ʾ��ʱ ms
    popMsg.Transparency = 200 '͸����
    
    If lngType = 0 Then
        SetOffice2003Theme (strMsg)
    Else
        SetToolTipTheme (strMsg)
    End If
    popMsg.Show
End Sub

Private Sub SetOffice2003Theme(ByVal str���� As String)
    '��ʾ������Ϣ��ʾ
    Dim item As PopupControlItem
    popMsg.RemoveAllItems
    
    Set item = popMsg.AddItem(50, 27, 200, 45, "�����µ���Ϣ��")
    item.Bold = True
    item.Hyperlink = False
    
    Set item = popMsg.AddItem(12, 20, 12, 27, "")
    item.SetIcon Me.imgMail.Picture.Handle, xtpPopupItemIconNormal
    item.IconIndex = 1
    
    Set item = popMsg.AddItem(50, 45, 270, 95, str����)
    item.TextColor = RGB(0, 61, 178)
    item.Id = IDSITE
    item.TextAlignment = DT_LEFT Or DT_WORDBREAK
    
    Set item = popMsg.AddItem(250, 10, 266, 26, "")
    item.SetIcon Me.imgClose.Picture.Handle, xtpPopupItemIconNormal
    item.Id = IDCLOSE
    item.Button = True
    
    popMsg.VisualTheme = xtpPopupThemeOffice2003
    popMsg.SetSize 270, 100

End Sub


Sub SetToolTipTheme(ByVal strTip As String)
    '��ʾ��ʾ
    Dim item As PopupControlItem
    
    popMsg.RemoveAllItems
    
    Set item = popMsg.AddItem(0, 0, 220, 90, "", RGB(255, 255, 225), 0)
    
    Set item = popMsg.AddItem(20, 30, 200, 100, strTip)
    item.TextAlignment = DT_CENTER Or DT_WORDBREAK
    item.Hyperlink = False
    
    Set item = popMsg.AddItem(5, 0, 170, 25, "��ʾ")
    item.TextAlignment = DT_SINGLELINE Or DT_VCENTER
    item.Bold = True
    item.Hyperlink = False
    
    Set item = popMsg.AddItem(220 - 20, 2, 220 - 2, 2 + 18, "")
    item.SetIcons Me.imgTipClose.Picture.Handle, 0, xtpPopupItemIconNormal Or xtpPopupItemIconSelected Or xtpPopupItemIconPressed
    item.IconIndex = 0
    item.Id = IDCLOSE
   
    popMsg.VisualTheme = xtpPopupThemeCustom
    popMsg.SetSize 220, 90

End Sub

