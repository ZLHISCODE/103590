VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmDockSeat 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12090
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   12090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox PicTab 
      Height          =   5670
      Left            =   330
      ScaleHeight     =   5610
      ScaleWidth      =   11355
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1095
      Width           =   11415
      Begin VB.PictureBox picSeats 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5985
         Left            =   180
         ScaleHeight     =   5955
         ScaleWidth      =   10800
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   165
         Width           =   10830
         Begin zl9Transfusion.udSeat ctlSeat 
            Height          =   2310
            Index           =   0
            Left            =   165
            TabIndex        =   4
            Top             =   300
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   4075
         End
      End
      Begin VB.VScrollBar vsbSeat 
         Height          =   5955
         Left            =   11100
         TabIndex        =   2
         Top             =   240
         Width           =   260
      End
   End
   Begin MSComctlLib.TabStrip TabSeat 
      Height          =   6675
      Left            =   135
      TabIndex        =   0
      Top             =   495
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   11774
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgRpt 
      Left            =   3420
      Top             =   105
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeat.frx":0000
            Key             =   "move"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeat.frx":059A
            Key             =   "��ִ��"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeat.frx":0B34
            Key             =   "no"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeat.frx":10CE
            Key             =   "����ִ��"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeat.frx":1668
            Key             =   "yes"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeat.frx":1C02
            Key             =   "Calling"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsSub 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmDockSeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Public Event Activate() '���Ѽ���ʱ
Public Event RequestRefresh() 'Ҫ��������ˢ��
Public Event StatusTextUpdate(ByVal Text As String) 'Ҫ�����������״̬������

Private Const subMenu_Add = 101
Private Const subMenu_Modify = 102
Private Const subMenu_Delete = 103

Private Const subMenu_View = 200
Private Const subMenu_Icon = 201
Private Const subMenu_List = 202
Private Const subMenu_Report = 203

Private Const subMenu_Clear = 300 '���ռ�õ���λ
Private Const subMenu_SetSeating = 400 '������λ

Private mcurSeatings As Seatings        '��λ��¼��

Public lng����ID As Long                '�����崫������ ���ڰ�����λ
Public objPati As cPatient

Private mSourceItem As String           '����ʱ��Դ��λ
Private mObjItem As String              '����ʱ��Ŀ����λ
Private mcbsMain As CommandBars         '�������õĹ�����

Private mSelectKey As String            '��ǰѡ�����λ
Private mSelectIndex As Integer         '��ǰѡ�����λ������

Private mSelectType As String           '��ǰѡ��ķ���ҳ

Private mblnFormResize As Boolean           '�����Ƿ��ڱ仯���仯ʱ��ˢ����λ
Private mlngMax As Long
    
Public Sub zlRefresh(ByVal curSeatings As Seatings)
     
    Dim intIndex As Integer
    Dim curSeating As Seating
    Dim strType As String
    Dim strTmp As String
    Set mcurSeatings = Nothing
    mSourceItem = ""
    mObjItem = ""
    Set mcurSeatings = curSeatings
    
    strType = ""
    TabSeat.Tabs.Clear
    
    For Each curSeating In mcurSeatings
        With curSeating
            
            '�ӷ���
            
            strTmp = "" & .����
            If strTmp = "" Then strTmp = "��ͨ��λ"
            
            If strType = "" Then
                TabSeat.Tabs.Add , strTmp, strTmp
                strType = strTmp
            ElseIf InStr("," & strType & ",", "," & strTmp & ",") <= 0 Then
                TabSeat.Tabs.Add , strTmp, strTmp
                strType = strType & "," & strTmp
            End If
             
        End With
    Next
    
    If strType = "" Then
        TabSeat.Tabs.Add , "��ͨ��λ", "��ͨ��λ"
    End If
    
    mSelectIndex = -1
    mSelectKey = ""
    
    If mSelectType <> "" Then
        On Error Resume Next
        TabSeat.Tabs(mSelectType).Selected = True
        Call TabSeat_Click
    Else

        TabSeat.Tabs("��ͨ��λ").Selected = True
        Call TabSeat_Click
    End If
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As CommandBar)
    If CommandBar.Parent Is Nothing Then Exit Sub
    
    Select Case CommandBar.Parent.ID
    Case Else
    End Select
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
     '#
    Dim StrKey As String, strObjKey As String
    
    Select Case Control.ID
'        Case conMenu_Edit_Seat_Icon
'            lvwSeating(mintActiveLvw).View = lvwIcon
'        Case conMenu_Edit_Seat_Report
'            lvwSeating(mintActiveLvw).View = lvwReport
'        Case conMenu_Edit_Seat_List
'            lvwSeating(mintActiveLvw).View = lvwList
        Case conMenu_Edit_Seat_Add
            If frmSeatingMana.SeatingMana(0, mcurSeatings, 0, "", Me, mSelectType) Then
                RaiseEvent RequestRefresh
            End If
        Case conMenu_Edit_Seat_Modify
            Call ModiSeat
        Case conMenu_Edit_Seat_Delete
            StrKey = mSelectKey
            If mcurSeatings.Delete(StrKey) Then
                RaiseEvent RequestRefresh
            End If
        Case conMenu_Edit_Seat_Set
            '������λ
            If lng����ID <> 0 And Not objPati Is Nothing Then
                StrKey = mSelectKey
                If MsgBox("�Ƿ���[" & objPati.���� & "]��[" & StrKey & "]��λ", vbQuestion + vbOKCancel + vbDefaultButton2, gstrSysName) = vbOK Then
                    If mcurSeatings.SetSeating(lng����ID, objPati.�Һŵ�, StrKey) Then
                        RaiseEvent RequestRefresh
                    End If
                End If
            End If
        Case conMenu_Edit_Seat_Clear
        '���ռ�õ���λ
            StrKey = mSelectKey
            If mcurSeatings.Clear(StrKey) Then
                RaiseEvent RequestRefresh
            End If
        Case conMenu_Edit_Seat_Swap
            '����λ
            StrKey = mSelectKey
            strObjKey = frmSeatingSwap.ObjectKey(StrKey, mcurSeatings, Me)
            If strObjKey <> "" Then
                If mcurSeatings.SwapSeating(StrKey, strObjKey) Then
                    RaiseEvent RequestRefresh
                End If
            End If
    End Select
End Sub

Private Sub ModiSeat()
    '�޸���λ
    Dim StrKey As String
    If mSelectKey <> "" Then
        StrKey = mSelectKey
        If ctlSeat(mSelectIndex).PatiName = "" Then
            If frmSeatingMana.SeatingMana(1, mcurSeatings, 0, StrKey, Me) Then
                RaiseEvent RequestRefresh
            End If
        End If
    End If

End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
 
    
    Select Case Control.ID
        Case conMenu_Edit_Seat_Modify, conMenu_Edit_Seat_Delete
        
            Control.Enabled = InStr(";" & gstrPrivs & ";", ";" & "��λ����" & ";") > 0
            
            If Control.Enabled Then
                If mSelectIndex = -1 Then
                    Control.Enabled = False
                ElseIf Not (ctlSeat.UBound >= mSelectIndex And ctlSeat.lBound <= mSelectIndex) Then
                    Control.Enabled = False
                ElseIf ctlSeat(mSelectIndex).PatiName <> "" Then
                    Control.Enabled = False
                End If
            End If
            
        
        Case conMenu_Edit_Seat_Add
            Control.Enabled = InStr(";" & gstrPrivs & ";", ";" & "��λ����" & ";") > 0
'        Case conMenu_Edit_Seat_Icon
'            Control.Checked = lvwSeating(mintActiveLvw).View = lvwIcon
'        Case conMenu_Edit_Seat_List
'            Control.Checked = lvwSeating(mintActiveLvw).View = lvwList
'        Case conMenu_Edit_Seat_Report
'            Control.Checked = lvwSeating(mintActiveLvw).View = lvwReport
        Case conMenu_Edit_Seat_Set
        
            Control.Enabled = InStr(";" & gstrPrivs & ";", ";" & "��λ����" & ";") > 0

            If Control.Enabled Then
                If mSelectIndex = -1 Then
                    Control.Enabled = False
                ElseIf Not (ctlSeat.UBound >= mSelectIndex And ctlSeat.lBound <= mSelectIndex) Then
                    Control.Enabled = False
                ElseIf Not (lng����ID <> 0 And ctlSeat(mSelectIndex).Stat = 0) Then
                    Control.Enabled = False
                End If
            End If
        Case conMenu_Edit_Seat_Clear
            Control.Enabled = InStr(";" & gstrPrivs & ";", ";" & "��λ����" & ";") > 0
            If Control.Enabled Then
                If mSelectIndex = -1 Then
                    Control.Enabled = False
                ElseIf Not (ctlSeat.UBound >= mSelectIndex And ctlSeat.lBound <= mSelectIndex) Then
                    Control.Enabled = False
                ElseIf ctlSeat(mSelectIndex).PatiName = "" Then
                    Control.Enabled = False
                End If
                
            End If
        Case conMenu_Edit_Seat_Swap
            Control.Enabled = InStr(";" & gstrPrivs & ";", ";" & "��λ����" & ";") > 0

            If Control.Enabled Then
                If mSelectIndex = -1 Then
                    Control.Enabled = False
                ElseIf Not (ctlSeat.UBound >= mSelectIndex And ctlSeat.lBound <= mSelectIndex) Then
                    Control.Enabled = False
                ElseIf ctlSeat(mSelectIndex).PatiName = "" Then
                    Control.Enabled = False
                End If
            End If
    End Select
End Sub

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal cbsMain As CommandBars, ByVal int���� As Integer)
    '������Ҫ���ʼ���������ϵĲ˵�
    Dim objMenu As CommandBarPopup, objViewMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl

    '������Ŀ�Ĳ˵�:���ڹ���˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set mcbsMain = cbsMain
    Set mcbsMain.Icons = zlCommFun.GetPubIcons
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "��λ����(&S)", objMenu.Index + 1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Set, "������λ(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Clear, "�����λ(&C)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Swap, "������λ(&W)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Add, "������λ(&A)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Modify, "�޸���λ(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Delete, "ɾ����λ(&D)")
        
'        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_View_Seat, "��λͼ��")
'        objPopup.ID = conMenu_Edit_Seat_View: objPopup.BeginGroup = True
'        With objPopup.CommandBar.Controls
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_View_GBed, "��ͨ��λ")
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_View_RBed, "ռ�ô�λ")
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_View_YBed, "ά����λ")
'
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_View_Gseat, "��ͨ��λ")
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_View_Rseat, "ռ����λ")
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_View_Yseat, "ά����λ")
'        End With
    End With
    
'    Set objViewMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
'    If objViewMenu Is Nothing Then
'        With objMenu.CommandBar.Controls
'            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Seat_View, "�鿴(&V)")
'            objPopup.ID = conMenu_Edit_Seat_View: objPopup.BeginGroup = True
'            With objPopup.CommandBar.Controls
'                Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Icon, "ͼ�귽ʽ(&I)")
'                Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_List, "�б�ʽ(&L)")
'                Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Report, "����ʽ(&R)")
'            End With
'        End With
'    Else
'        With objViewMenu.CommandBar.Controls
'            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Seat_View, "�鿴��ʽ(&V)")
'            objPopup.ID = conMenu_Edit_Seat_View: objPopup.BeginGroup = True
'            With objPopup.CommandBar.Controls
'                Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Icon, "ͼ��(&I)")
'                Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_List, "�б�(&L)")
'                Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Report, "����(&R)")
'            End With
'        End With
'
'    End If
    '����������:���ļ�������˵������ť֮��ʼ����
    '-----------------------------------------------------
    Set objBar = cbsMain(2)
    For Each objControl In objBar.Controls '�����ǰ������һ��Control
        
        If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
            Set objControl = objBar.Controls(objControl.Index - 1): Exit For
        End If
    Next
    With objBar.Controls
        Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_Seat, "��λ", objControl.Index + 1)
        objPopup.ID = conMenu_Edit_Seat: objPopup.BeginGroup = True
        
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Set, "������λ")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Clear, "�����λ")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Swap, "������λ")
            
        End With
        
        
    End With
End Sub

Private Sub cbsSub_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsSub.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    'lngBottom = lngBottom - stbCutline.Height
 
    mblnFormResize = True
    
    
    Me.TabSeat.Left = lngLeft
    Me.TabSeat.Top = lngTop
    Me.TabSeat.Width = lngRight - lngLeft
    Me.TabSeat.Height = lngBottom - lngTop
    mblnFormResize = False
    Call picTab_Resize
    
    Me.Refresh
End Sub

Private Sub ctlSeat_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Dim strObjKey As String
    
    
    If TypeName(Source) = "udSeat" Then
        '����
         
        
        If Not mSourceItem = "" And Not mObjItem = "" Then
            strObjKey = frmSeatingSwap.ObjectKey(mSourceItem, mcurSeatings, Me, mObjItem)
            If strObjKey <> "" Then
                If mcurSeatings.SwapSeating(mSourceItem, strObjKey) = True Then
                    RaiseEvent RequestRefresh
                End If
            End If
             
        End If
    End If

    If TypeName(Source) = "ReportControl" And lng����ID <> 0 Then
        '����
         
        'lvwSeating(Index).MousePointer = ccDefault
        If Not mObjItem = "" And Not objPati Is Nothing Then
            If MsgBox("�Ƿ���[" & objPati.���� & "]��[" & mcurSeatings(mObjItem).��� & "]��λ", vbQuestion + vbOKCancel + vbDefaultButton2, gstrSysName) = vbOK Then
                Call mcurSeatings.SetSeating(lng����ID, objPati.�Һŵ�, mObjItem)
                RaiseEvent RequestRefresh
            End If
        End If
    End If
    
End Sub

Private Sub ctlSeat_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)

    Dim objOver As ListItem
    
    Source.DragIcon = imgRpt.ListImages("move").Picture
    If TypeName(Source) = "udSeat" Then
         
        
        If ctlSeat(Index).Stat = 0 Then
                '��������λ
                
            Set Source.DragIcon = imgRpt.ListImages("yes").Picture
            mObjItem = ctlSeat(Index).Key
        Else
            Set Source.DragIcon = imgRpt.ListImages("no").Picture
            mObjItem = ""
        End If
 
    End If
    If TypeName(Source) = "ReportControl" Then
        
        If ctlSeat(Index).Stat = 0 And lng����ID <> 0 Then

            Set Source.DragIcon = imgRpt.ListImages("yes").Picture
            mObjItem = ctlSeat(Index).Key
        Else
            Set Source.DragIcon = imgRpt.ListImages("no").Picture
            mObjItem = ""
        End If
    End If
    
    If State = 1 Then
        '�뿪
        'ctlSeat(Index).GridColor = vbBlue
    Else
        'ctlSeat(Index).GridColor = vbMagenta
    End If
    
End Sub

Private Sub ctlSeat_GotFocus(Index As Integer)
    
    ctlSeat(Index).GridColor = vbRed
    mSelectIndex = Index
    mSelectKey = ctlSeat(Index).Key
    
    RaiseEvent StatusTextUpdate(mSelectKey)
End Sub

Private Sub ctlSeat_LostFocus(Index As Integer)
    ctlSeat(Index).GridColor = vbBlue
End Sub

Private Sub ctlSeat_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        '��λ���ܲ���
        If ctlSeat(Index).PatiName = "" Then
            ctlSeat(Index).Drag vbCancel
            Exit Sub
        End If
        mSourceItem = ctlSeat(Index).Key
        
        Set ctlSeat(Index).DragIcon = imgRpt.ListImages("move").Picture
        ctlSeat(Index).Drag vbBeginDrag
    End If
    
End Sub

Private Sub ctlSeat_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    If Button = 2 Then
        Set objPopup = mcbsMain.ActiveMenuBar.FindControl(, conMenu_EditPopup)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Private Sub Form_Load()
   cbsSub.ActiveMenuBar.Visible = False
End Sub

Private Sub Form_Resize()
    Call cbsSub_Resize
End Sub

Private Sub picSeats_Resize()
    
    Dim iCount As Integer, iRow As Integer
    Dim lngCurLeft As Long, lngCurTop As Long
    Dim lngCarHeight As Long, lngCarWidth As Long '��Ƭ����
    Dim lngSplitWidth  As Long  '���
    
    On Error Resume Next
    
    If mblnFormResize Then Exit Sub
    With picSeats
        
        lngCarHeight = ctlSeat(ctlSeat.lBound).Height
        lngCarWidth = ctlSeat(ctlSeat.lBound).Width
        lngSplitWidth = 35
        
        '-- ��Ƭ
        For iCount = ctlSeat.lBound To ctlSeat.UBound
        
            If iCount = ctlSeat.lBound Then
                '��һ����Ƭ
                ctlSeat(iCount).Left = .ScaleTop + lngSplitWidth
                ctlSeat(iCount).Top = .ScaleTop + lngSplitWidth
                
                lngCurLeft = ctlSeat(iCount).Left
                lngCurTop = ctlSeat(iCount).Top
                iRow = 0
            Else
                '֮��Ŀ�Ƭ���ݵ�һ����Ƭ��λ�����У�������Ȼ���
                lngCurLeft = lngCurLeft + lngCarWidth + lngSplitWidth
                If lngCurLeft + lngCarWidth > picSeats.ScaleWidth Then
                    iRow = iRow + 1
                    lngCurLeft = ctlSeat(ctlSeat.lBound).Left
                    lngCurTop = lngCarHeight * iRow + lngSplitWidth * iRow + lngSplitWidth
                    
                End If
                ctlSeat(iCount).Left = lngCurLeft
                ctlSeat(iCount).Top = lngCurTop
            End If
        Next
        mblnFormResize = True
        If ctlSeat(ctlSeat.UBound).Top + lngCarHeight + lngSplitWidth > Me.PicTab.ScaleHeight Then
            picSeats.Height = ctlSeat(ctlSeat.UBound).Top + lngCarHeight + lngSplitWidth
        End If
        mblnFormResize = False
    End With
    
    '--- ������
    vsbSeat.Top = Me.PicTab.ScaleTop
    vsbSeat.Left = Me.PicTab.ScaleWidth - vsbSeat.Width
    vsbSeat.Height = Me.PicTab.ScaleHeight
    vsbSeat.Max = (picSeats.Height - Me.PicTab.ScaleHeight) / Screen.TwipsPerPixelX  'ת��Ϊ����Ϊ��λ
    Call ShowSize(0)
    vsbSeat.Value = 0
    

End Sub


Private Sub ShowSize(Optional lngTop As Single = 0, Optional lngLeft As Single = 0)
    '����:��ʾ������λ
    picSeats.Left = lngLeft
    picSeats.Top = lngTop
    
    Me.Refresh
End Sub

Private Sub picTab_Resize()
    Dim lngCurLeft As Long
    Dim lngCarWidth As Long
    Dim lngCarHeight As Long
    Dim lngSplitWidth As Long
    Dim iRow As Integer
    Dim i As Integer
    On Error Resume Next
    mblnFormResize = True
    Me.PicTab.Move Me.TabSeat.ClientLeft, Me.TabSeat.ClientTop, Me.TabSeat.ClientWidth, Me.TabSeat.ClientHeight
    
    Me.picSeats.Left = Me.PicTab.ScaleLeft
    Me.picSeats.Width = Me.PicTab.ScaleWidth - Me.vsbSeat.Width
    Me.picSeats.Top = Me.PicTab.ScaleTop
    
    lngCarWidth = ctlSeat(ctlSeat.lBound).Width
    lngCarHeight = ctlSeat(ctlSeat.lBound).Height

    lngSplitWidth = 15
    iRow = 0
    For i = 1 To mlngMax
        lngCurLeft = lngSplitWidth + lngCarWidth * i + lngSplitWidth + i
        If lngCurLeft > Me.picSeats.ScaleWidth Then
            iRow = iRow + 1
        End If
    Next
    picSeats.Height = lngSplitWidth + iRow + 1 * lngCarHeight + lngSplitWidth * iRow + 1

    If picSeats.Height < Me.PicTab.ScaleHeight Then
        picSeats.Height = Me.PicTab.ScaleHeight
    End If
    
    mblnFormResize = False
    
    Call picSeats_Resize
End Sub

Private Sub TabSeat_Click()
    Dim i As Integer, iCur As Integer, iMax As Integer
        
    '--װ����λ
    mSelectType = TabSeat.SelectedItem.Key
    
    Dim curSeating As Seating
    iMax = 0
    For Each curSeating In mcurSeatings
        If IIf(curSeating.���� = "", "��ͨ��λ", curSeating.����) = mSelectType Then
            iMax = iMax + 1
        End If
    Next
    
    iCur = ctlSeat.UBound + 1
    If iMax <> iCur Then
        If iMax > iCur Then
            For i = iCur To iMax - 1
                'If ctlSeat.UBound = i - 1 Then Exit For
                    Load ctlSeat(i)
            Next
        Else
            For i = iMax To iCur - 1
                If i = ctlSeat.lBound Then
                    ctlSeat(i).Visible = False
                Else
                    Unload ctlSeat(i)
                End If
                
            Next
        End If
    End If
    
    i = 0
    For Each curSeating In mcurSeatings
        If IIf(curSeating.���� = "", "��ͨ��λ", curSeating.����) = mSelectType Then
            ctlSeat(i).SeatNo = curSeating.���
            ctlSeat(i).PatiName = curSeating.���� '& " " & curSeating.���g
            ctlSeat(i).SeatType = curSeating.����
            ctlSeat(i).Sex = "" & curSeating.�Ա�
            ctlSeat(i).Diagnosis = curSeating.���
            ctlSeat(i).Start = curSeating.��ʼʱ�� '& " " & curSeating.�����T
            ctlSeat(i).Stat = curSeating.״̬
            ctlSeat(i).Key = curSeating.Key
            
            ctlSeat(i).GridColor = vbBlue
            ctlSeat(i).GridWidth = 1
            ctlSeat(i).Visible = True
            i = i + 1
        End If
    Next
    
    mlngMax = iMax
    
    Call picTab_Resize
End Sub

Private Sub vsbSeat_Change()
    Call ShowSize(-vsbSeat.Value * 15#)
End Sub

Private Sub vsbSeat_Scroll()
    Call ShowSize(-vsbSeat.Value * 15#)
End Sub
