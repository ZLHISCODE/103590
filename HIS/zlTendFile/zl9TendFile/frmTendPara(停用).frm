VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTendPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��¼��ѡ��"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4920
   Icon            =   "frmTendPara.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCanc 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3735
      TabIndex        =   17
      Top             =   3720
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2580
      TabIndex        =   16
      Top             =   3720
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   3510
      Index           =   1
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   4800
      Begin VB.PictureBox picControl 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1845
         Left            =   1950
         ScaleHeight     =   1845
         ScaleWidth      =   2295
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1485
         Visible         =   0   'False
         Width           =   2295
         Begin VB.CommandButton cmdUnVisible 
            Height          =   315
            Left            =   1815
            Picture         =   "frmTendPara.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   23
            TabStop         =   0   'False
            ToolTipText     =   "ȡ��"
            Top             =   1500
            Width           =   450
         End
         Begin VB.PictureBox PicColorCollect 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1350
            Left            =   60
            Picture         =   "frmTendPara.frx":0596
            ScaleHeight     =   1350
            ScaleWidth      =   2160
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   90
            Width           =   2160
            Begin VB.Shape shpValue 
               BorderColor     =   &H00C56A31&
               FillColor       =   &H00FF8080&
               Height          =   270
               Left            =   0
               Top             =   0
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Shape shpBorder 
               BorderColor     =   &H00C56A31&
               FillColor       =   &H00FF8080&
               Height          =   270
               Left            =   1890
               Top             =   1080
               Visible         =   0   'False
               Width           =   270
            End
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   200
            Left            =   90
            ScaleHeight     =   165
            ScaleWidth      =   165
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   1575
            Width           =   200
         End
         Begin VB.Label lblColor 
            Caption         =   "&HFFFFFF"
            Height          =   195
            Left            =   405
            TabIndex        =   22
            Top             =   1575
            UseMnemonic     =   0   'False
            Width           =   1365
         End
      End
      Begin VB.PictureBox picLineColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1950
         ScaleHeight     =   180
         ScaleWidth      =   2265
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "���ѡ����ɫ"
         Top             =   1275
         Width           =   2295
      End
      Begin VB.CheckBox chk 
         Caption         =   "Ԥ������ӡʱͬһҳ��ͬ������ʾһ��"
         Height          =   180
         Index           =   4
         Left            =   300
         TabIndex        =   13
         Top             =   2580
         Width           =   3540
      End
      Begin VB.ComboBox cboOperSing 
         Height          =   300
         ItemData        =   "frmTendPara.frx":0D0C
         Left            =   1950
         List            =   "frmTendPara.frx":0D0E
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   540
         Width           =   2670
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1755
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   9
         Text            =   "1"
         Top             =   1590
         Width           =   420
      End
      Begin VB.ComboBox cboNodule 
         Height          =   300
         ItemData        =   "frmTendPara.frx":0D10
         Left            =   1950
         List            =   "frmTendPara.frx":0D12
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   900
         Width           =   2670
      End
      Begin VB.CheckBox chk 
         Caption         =   "ֻ�ڵ�ǰҳ����ʾ��ҳ���ݣ�������ҳ����ʾ��"
         Height          =   180
         Index           =   2
         Left            =   300
         TabIndex        =   14
         Top             =   2865
         Width           =   4215
      End
      Begin VB.CheckBox chk 
         Caption         =   "סԺ����ͬһʱ����Ҫ��¼��ݻ����ļ�"
         Height          =   180
         Index           =   0
         Left            =   300
         TabIndex        =   11
         Top             =   1965
         Width           =   3645
      End
      Begin VB.CheckBox chk 
         Caption         =   "�����ļ�ҳ�밴�ļ�˳����"
         Height          =   180
         Index           =   3
         Left            =   300
         TabIndex        =   15
         Top             =   3165
         Width           =   3135
      End
      Begin VB.CheckBox chk 
         Caption         =   "Ԥ������ӡʱǩ������ʾǩ��ͼƬ"
         Height          =   180
         Index           =   1
         Left            =   300
         TabIndex        =   12
         Top             =   2280
         Width           =   3645
      End
      Begin VB.ComboBox cboSinger 
         Height          =   300
         ItemData        =   "frmTendPara.frx":0D14
         Left            =   1950
         List            =   "frmTendPara.frx":0D16
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   180
         Width           =   2670
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   0
         Left            =   2160
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1575
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txt(0)"
         BuddyDispid     =   196622
         BuddyIndex      =   0
         OrigLeft        =   2175
         OrigTop         =   930
         OrigRight       =   2430
         OrigBottom      =   1200
         Max             =   30
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblLineColor 
         AutoSize        =   -1  'True
         Caption         =   "С���ʶ��ɫ"
         Height          =   180
         Left            =   810
         TabIndex        =   18
         Top             =   1290
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʿ��ǩ������ʾģʽ"
         Height          =   180
         Index           =   3
         Left            =   105
         TabIndex        =   3
         Top             =   600
         Width           =   1800
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����¼�볬����ǰ        ��Ļ����¼����"
         Height          =   180
         Index           =   2
         Left            =   300
         TabIndex        =   8
         Top             =   1635
         Width           =   3600
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "С��ȱʡ��ʶ"
         Height          =   180
         Index           =   1
         Left            =   810
         TabIndex        =   5
         Top             =   975
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǩģʽ"
         Height          =   180
         Index           =   0
         Left            =   1170
         TabIndex        =   1
         Top             =   240
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmTendPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmMain As Object
Private mblnOK As Boolean
Private mstrPrivs As String

Private mvarColor As OLE_COLOR
Private Const tomAutoColor As Long = -9999997
'�趨һ�����岶����꣬���������������Ϣ�������ô���
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
'ȡ����겶��
Private Declare Function ReleaseCapture Lib "user32" () As Long

Public Function ShowPara(ByVal frmMain As Object, ByVal strPrivs As String) As Boolean
    Dim intLoop As Integer
    Dim strTmp As String
    Dim strSQL As String, strPar As String
    Dim curDate As Date, intDay As Integer
    Dim intStart As Integer
    
    mblnOK = False
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
    
    '��ʼ���µ����
    '------------------------------------------------------------------------------------------------------------------
    
    '43588,������,2012-09-13,��Ӽ�¼����ǩģʽ
    cboSinger.Clear
    cboSinger.AddItem "0-Ƹ��ְ��+��ǩȨ��"
    cboSinger.AddItem "1-��ǩȨ��"
    
    cboNodule.Clear
    cboNodule.AddItem "0-������"
    cboNodule.AddItem "1-���»����߱�ʶ"
    cboNodule.AddItem "2-����ֵ�·���˫���߱�ʶ"
    cboNodule.AddItem "3-�Ϸ������߱�ʶ"
    '72664:������,2014-07-18,���С���ʶ
    cboNodule.AddItem "4-����ֵ�·��������߱�ʶ"
    
    '58414,������,2013-04-10,��ӻ�ʿ��ǩ������ʾģʽ
    cboOperSing.Clear
    cboOperSing.AddItem "0-��������ʾ"
    cboOperSing.AddItem "1-������ʾ"
    cboOperSing.AddItem "2-��β����ʾ"
    cboOperSing.AddItem "3-β����ʾ"
    
    '43588:������,2012-09-13,��Ӽ�¼����ǩģʽ
    strTmp = Val(zlDatabase.GetPara("��¼����ǩģʽ", glngSys, 1255, "0", Array(cboSinger, lbl(0)), InStr(mstrPrivs, "����ѡ������") > 0))
    If Val(strTmp) >= 0 And Val(strTmp) <= 1 Then
        cboSinger.ListIndex = CInt(Val(strTmp))
    Else
        cboSinger.ListIndex = 0
    End If
    
    strTmp = zlDatabase.GetPara("С��ȱʡ��ʽ", glngSys, 1255, "0", Array(cboNodule, lbl(1)), InStr(mstrPrivs, "����ѡ������") > 0)
    If Val(strTmp) >= 0 And Val(strTmp) <= 4 Then
        cboNodule.ListIndex = Val(strTmp)
    Else
        cboNodule.ListIndex = 0
    End If
    
    '58414,������,2013-04-10,��ӻ�ʿ��ǩ������ʾģʽ
    strTmp = Val(zlDatabase.GetPara("��ʿ��ǩ������ʾģʽ", glngSys, 1255, "2", Array(cboOperSing, lbl(3)), InStr(mstrPrivs, "����ѡ������") > 0))
    If Val(strTmp) >= 0 And Val(strTmp) <= 3 Then
        cboOperSing.ListIndex = CInt(Val(strTmp))
    Else
        cboOperSing.ListIndex = 2
    End If
    
    txt(0).Text = Val(zlDatabase.GetPara("����¼�뻤����������", glngSys, 1255, "1", Array(txt(0), lbl(2)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(0).Value = Val(zlDatabase.GetPara("��Ӧ��ݻ����ļ�", glngSys, 1255, "0", Array(chk(0)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(1).Value = Val(zlDatabase.GetPara("��¼��ǩ������ʾ��ʽ", glngSys, 1255, "0", Array(chk(1)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(2).Value = Val(zlDatabase.GetPara("��ҳ����ֻ��ʾ�ڵ�һҳ", glngSys, 1255, "0", Array(chk(2)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(3).Tag = 1
    chk(3).Value = Val(zlDatabase.GetPara("�����ļ�ҳ�����", glngSys, 1255, "0", Array(chk(3)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(3).Tag = 0
    '64583:������,2013-09-22,Ԥ������ӡʱͬһҳ��ͬ������ʾ��ʽ:���;һ��
    chk(4).Value = Val(zlDatabase.GetPara("��¼��������ʾ��ʽ", glngSys, 1255, "0", Array(chk(4)), InStr(mstrPrivs, "����ѡ������") > 0))
    '68739:������,2014-1-2,���"С���ʶ��ɫ"
    picLineColor.BackColor = Val(zlDatabase.GetPara("С���ʶ��ɫ", glngSys, 1255, "255", Array(lblLineColor), InStr(mstrPrivs, "����ѡ������") > 0))
    
    Me.Show 1, mfrmMain
    ShowPara = mblnOK
    
End Function

Private Sub cboNodule_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub


Private Sub cboOperSing_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cboSinger_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chk_Click(Index As Integer)
    Dim strInfo As String
    If Not Index = 3 Then Exit Sub
    If Val(chk(Index).Tag) = 1 Then chk(Index).Tag = 0: Exit Sub
    strInfo = "�˲���ֱ��Ӱ���ż�¼�������ļ���ҳ���Ź��򣬶���������������ļ���ҳ�뽫���յ�����Ĺ����š�"
    strInfo = strInfo & vbCrLf & "1�����˼�¼���ļ�����С�ڵ���1�������������µļ�¼���ļ���"
    strInfo = strInfo & vbCrLf & "2���������м�¼���ļ�֮�������˺ϲ���ӡ������ȡ����ĳ�ݱ��ϲ����ļ���"
    strInfo = strInfo & vbCrLf & "�������Ƿ������"
    If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        chk(Index).Tag = 1
        chk(Index).Value = IIf(chk(Index).Value = 0, 1, 0)
    End If
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub


Private Sub cmdCanc_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim intStart As Integer
    Dim strTmp As String
    Dim lngColor As Long
    
    '43588:������,2012-09-13,��Ӽ�¼����ǩģʽ
    Call zlDatabase.SetPara("��¼����ǩģʽ", Val(cboSinger.ListIndex), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("С��ȱʡ��ʽ", Val(cboNodule.ListIndex), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("����¼�뻤����������", Val(txt(0).Text), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("��Ӧ��ݻ����ļ�", chk(0).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("��¼��ǩ������ʾ��ʽ", chk(1).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("��ҳ����ֻ��ʾ�ڵ�һҳ", chk(2).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("�����ļ�ҳ�����", chk(3).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    '58414,������,2013-04-10,��ӻ�ʿ��ǩ������ʾģʽ
    Call zlDatabase.SetPara("��ʿ��ǩ������ʾģʽ", Val(cboOperSing.ListIndex), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    '64583:������,2013-09-22,Ԥ������ӡʱͬһҳ��ͬ������ʾ��ʽ:���;һ��
    Call zlDatabase.SetPara("��¼��������ʾ��ʽ", chk(4).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    '68739:������,2014-1-2,���"С���ʶ��ɫ"
    Call zlDatabase.SetPara("С���ʶ��ɫ", Val(picLineColor.BackColor), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    
    
    mblnOK = True
    
    Unload Me
End Sub

Private Sub cmdUnVisible_Click()
    picControl.Visible = False
    If picLineColor.Enabled And picLineColor.Visible Then picLineColor.SetFocus
End Sub

Private Sub picLineColor_Click()
    picControl.Top = picLineColor.Top + picLineColor.Height
    picControl.Left = picLineColor.Left
    picControl.Visible = True
    Call SetCOLOR(Val(picLineColor.BackColor))
End Sub

Private Sub txt_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txt(Index))
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub PicColorCollect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > 0 And X < PicColorCollect.ScaleWidth And Y > 0 And Y < PicColorCollect.ScaleHeight Then
        SetCapture PicColorCollect.hWnd
        shpBorder.Visible = True
    Else
        ReleaseCapture
        shpBorder.Visible = False
    End If

    Dim lRow As Long, lCol As Long, lX As Long, lY As Long
    lRow = Y \ (18 * Screen.TwipsPerPixelY)
    lCol = X \ (18 * Screen.TwipsPerPixelX)
    lX = ((lCol) * 18 + 4) * Screen.TwipsPerPixelX
    lY = ((lRow) * 18 + 4) * Screen.TwipsPerPixelY
    shpBorder.Move lCol * 18 * Screen.TwipsPerPixelX, lRow * 18 * Screen.TwipsPerPixelY, 270, 270
    
    If PicColorCollect.Point(lX, lY) = -1 Then Exit Sub
    picColor.BackColor = PicColorCollect.Point(lX, lY)
    Select Case CStr(Hex(picColor.BackColor))
    Case "0"
        lblColor = "��ɫ"
    Case "3399"
        lblColor = "��ɫ"
    Case "3333"
        lblColor = "���ɫ"
    Case "3300"
        lblColor = "����"
    Case "663300"
        lblColor = "����"
    Case "800000"
        lblColor = "����"
    Case "993333"
        lblColor = "����"
    Case "333333"
        lblColor = "��ɫ-80%"
    Case "80"
        lblColor = "���"
    Case "66FF"
        lblColor = "��ɫ"
    Case "8080"
        lblColor = "���"
    Case "8000"
        lblColor = "��ɫ"
    Case "808000"
        lblColor = "��ɫ"
    Case "FF0000"
        lblColor = "��ɫ"
    Case "996666"
        lblColor = "��-��"
    Case "808080"
        lblColor = "��ɫ-50%"
    Case "FF"
        lblColor = "��ɫ"
    Case "99FF"
        lblColor = "ǳ��ɫ"
    Case "CC99"
        lblColor = "���ɫ"
    Case "669933"
        lblColor = "����"
    Case "CCCC33"
        lblColor = "ˮ��ɫ"
    Case "FF6633"
        lblColor = "ǳ��"
    Case "800080"
        lblColor = "������"
    Case "999999"
        lblColor = "��ɫ-40%"
    Case "FF00FF"
        lblColor = "�ۺ�"
    Case "CCFF"
        lblColor = "��ɫ"
    Case "FFFF"
        lblColor = "��ɫ"
    Case "FF00"
        lblColor = "����"
    Case "FFFF00"
        lblColor = "����"
    Case "FFCC00"
        lblColor = "����"
    Case "663399"
        lblColor = "÷��"
    Case "C0C0C0"
        lblColor = "��ɫ-25%"
    Case "CC99FF"
        lblColor = "õ���"
    Case "99CCFF"
        lblColor = "��ɫ"
    Case "99FFFF"
        lblColor = "ǳ��"
    Case "CCFFCC"
        lblColor = "ǳ��"
    Case "FFFFCC"
        lblColor = "ǳ����"
    Case "FFCC99"
        lblColor = "����"
    Case "FF99CC"
        lblColor = "����"
    Case "FFFFFF"
        lblColor = "��ɫ"
    Case Else
        lblColor = "&H" & CStr(Hex(picColor.BackColor))
    End Select
End Sub

Private Sub PicColorCollect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lRow As Long, lCol As Long, lX As Long, lY As Long
    lRow = Y \ (18 * Screen.TwipsPerPixelY)
    lCol = X \ (18 * Screen.TwipsPerPixelX)
    lX = ((lCol) * 18 + 4) * Screen.TwipsPerPixelX
    lY = ((lRow) * 18 + 4) * Screen.TwipsPerPixelY
    
    '��ָ����ɫ��ͼ
    picLineColor.BackColor = picColor.BackColor
    picControl.Visible = False
    If picLineColor.Enabled And picLineColor.Visible Then picLineColor.SetFocus
End Sub

Private Sub SetCOLOR(vData As OLE_COLOR)
    mvarColor = vData
    Dim lRow As Long, lCol As Long
    shpValue.Visible = True
    Select Case CStr(Hex(vData))
    Case "0"
        lblColor = "��ɫ"
        lRow = 0
        lCol = 0
    Case "3399"
        lblColor = "��ɫ"
        lRow = 0
        lCol = 1
    Case "3333"
        lblColor = "���ɫ"
        lRow = 0
        lCol = 2
    Case "3300"
        lblColor = "����"
        lRow = 0
        lCol = 3
    Case "663300"
        lblColor = "����"
        lRow = 0
        lCol = 4
    Case "800000"
        lblColor = "����"
        lRow = 0
        lCol = 5
    Case "993333"
        lblColor = "����"
        lRow = 0
        lCol = 6
    Case "333333"
        lblColor = "��ɫ-80%"
        lRow = 0
        lCol = 7
    Case "80"
        lblColor = "���"
        lRow = 1
        lCol = 0
    Case "66FF"
        lblColor = "��ɫ"
        lRow = 1
        lCol = 1
    Case "8080"
        lblColor = "���"
        lRow = 1
        lCol = 2
    Case "8000"
        lblColor = "��ɫ"
        lRow = 1
        lCol = 3
    Case "808000"
        lblColor = "��ɫ"
        lRow = 1
        lCol = 4
    Case "FF0000"
        lblColor = "��ɫ"
        lRow = 1
        lCol = 5
    Case "996666"
        lblColor = "��-��"
        lRow = 1
        lCol = 6
    Case "808080"
        lblColor = "��ɫ-50%"
        lRow = 1
        lCol = 7
    Case "FF"
        lblColor = "��ɫ"
        lRow = 2
        lCol = 0
    Case "99FF"
        lblColor = "ǳ��ɫ"
        lRow = 2
        lCol = 1
    Case "CC99"
        lblColor = "���ɫ"
        lRow = 2
        lCol = 2
    Case "669933"
        lblColor = "����"
        lRow = 2
        lCol = 3
    Case "CCCC33"
        lblColor = "ˮ��ɫ"
        lRow = 2
        lCol = 4
    Case "FF6633"
        lblColor = "ǳ��"
        lRow = 2
        lCol = 5
    Case "800080"
        lblColor = "������"
        lRow = 2
        lCol = 6
    Case "999999"
        lblColor = "��ɫ-40%"
        lRow = 2
        lCol = 7
    Case "FF00FF"
        lblColor = "�ۺ�"
        lRow = 3
        lCol = 0
    Case "CCFF"
        lblColor = "��ɫ"
        lRow = 3
        lCol = 1
    Case "FFFF"
        lblColor = "��ɫ"
        lRow = 3
        lCol = 2
    Case "FF00"
        lblColor = "����"
        lRow = 3
        lCol = 3
    Case "FFFF00"
        lblColor = "����"
        lRow = 3
        lCol = 4
    Case "FFCC00"
        lblColor = "����"
        lRow = 3
        lCol = 5
    Case "663399"
        lblColor = "÷��"
        lRow = 3
        lCol = 6
    Case "C0C0C0"
        lblColor = "��ɫ-25%"
        lRow = 3
        lCol = 7
    Case "CC99FF"
        lblColor = "õ���"
        lRow = 4
        lCol = 0
    Case "99CCFF"
        lblColor = "��ɫ"
        lRow = 4
        lCol = 1
    Case "99FFFF"
        lblColor = "ǳ��"
        lRow = 4
        lCol = 2
    Case "CCFFCC"
        lblColor = "ǳ��"
        lRow = 4
        lCol = 3
    Case "FFFFCC"
        lblColor = "ǳ����"
        lRow = 4
        lCol = 4
    Case "FFCC99"
        lblColor = "����"
        lRow = 4
        lCol = 5
    Case "FF99CC"
        lblColor = "����"
        lRow = 4
        lCol = 6
    Case "FFFFFF"
        lblColor = "��ɫ"
        lRow = 4
        lCol = 7
    Case Else
        lblColor = "&H" & CStr(Hex(picColor.BackColor))
    End Select
    shpBorder.Visible = False
    shpValue.Move lCol * 18 * Screen.TwipsPerPixelX, lRow * 18 * Screen.TwipsPerPixelY, 270, 270
    shpValue.Visible = True
    If vData = tomAutoColor Or vData = -1 Then
    
    Else
        picColor.BackColor = vData
    End If
End Sub

