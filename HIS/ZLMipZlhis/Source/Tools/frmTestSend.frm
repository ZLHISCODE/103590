VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmTestSend 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   3
      Left            =   1095
      ScaleHeight     =   240
      ScaleWidth      =   2190
      TabIndex        =   6
      Top             =   1020
      Width           =   2220
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   4
         Left            =   -30
         TabIndex        =   7
         Text            =   "cbo"
         Top             =   -30
         Width           =   2265
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Index           =   2
      Left            =   990
      ScaleHeight     =   2655
      ScaleWidth      =   5085
      TabIndex        =   4
      Top             =   1920
      Width           =   5085
      Begin RichTextLib.RichTextBox rtbSend 
         Height          =   1905
         Left            =   135
         TabIndex        =   5
         Top             =   45
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   3360
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"frmTestSend.frx":0000
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   5850
      ScaleHeight     =   240
      ScaleWidth      =   2190
      TabIndex        =   2
      Top             =   375
      Width           =   2220
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   -30
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   -30
         Width           =   2265
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   2415
      ScaleHeight     =   240
      ScaleWidth      =   2205
      TabIndex        =   0
      Top             =   300
      Width           =   2235
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   -30
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   -30
         Width           =   2265
      End
   End
   Begin XtremeCommandBars.ImageManager ImageManager1 
      Left            =   6495
      Top             =   1605
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmTestSend.frx":009D
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   90
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmTestSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mobjMipModule As Object

Private mblnStarted As Boolean

Private Sub cbo_Click(Index As Integer)
    Dim rs As ADODB.Recordset
    Dim strSQL As String
        
    Select Case Index
    Case 0
        With cbo(1)
            .Clear
            strSQL = "Select ���,���� From zlPrograms Where ϵͳ=[1]"
            Set rs = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "", cbo(0).ItemData(cbo(0).ListIndex))
            If rs.EOF = False Then
                Do While Not rs.EOF
                    .AddItem rs("���").Value & "-" & rs("����").Value
                    .ItemData(.NewIndex) = rs("���").Value
                    
                    rs.MoveNext
                Loop
            End If
            .ListIndex = 0
        End With

    End Select
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
    Case conMenu_Edit_Reuse
        If mblnStarted = False Then
        
            Set mobjMipModule = CreateObject("zl9ComLib.clsMipModule")
            Call mobjMipModule.InitMessage(cbo(0).ItemData(cbo(0).ListIndex), cbo(1).ItemData(cbo(1).ListIndex), "")
            Call gobjComLib.AddMipModule(mobjMipModule)
            mblnStarted = True
        Else
            If Not (mobjMipModule Is Nothing) Then
                mobjMipModule.CloseMessage
                Call gobjComLib.DelMipModule(mobjMipModule)
                Set mobjMipModule = Nothing
            End If
            mblnStarted = False
        End If
    Case conMenu_Edit_Send
        
        If Not (mobjMipModule Is Nothing) Then
            If mobjMipModule.CommitMessage(UCase(cbo(4).Text), rtbSend.Text) = False Then
                MsgBox "����ʧ��"
            End If
        End If
        
    End Select
    
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long
    Dim lngTop  As Long
    Dim lngRight  As Long
    Dim lngBottom  As Long

    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    
    '���������ؼ�Resize����
    picBack(2).Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
End Sub


Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
    Case 1
        Control.Enabled = mblnStarted
    Case conMenu_Edit_Reuse
        
        Control.Caption = IIf(mblnStarted = True, "�ǳ�ģ��", "����ģ��")
                
    Case conMenu_Edit_Send
        Control.Enabled = mblnStarted
    End Select
End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    Call InitCommandBar
        
    '---------------------------------
    With cbo(0)
        .Clear
        
        strSQL = "Select ���,���� From zlsystems"
        Set rs = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "")
        If rs.EOF = False Then
            Do While Not rs.EOF
                .AddItem rs("���").Value & "-" & rs("����").Value
                .ItemData(.NewIndex) = rs("���").Value
                
                rs.MoveNext
            Loop
        End If
        .ListIndex = 0
    End With
    Call cbo_Click(1)

    
    Call gobjComLib.zlControl.CboLocate(cbo(0), Val(GetSetting(App.ProductName, "����", "����ϵͳ��", "100")), True)
    Call gobjComLib.zlControl.CboLocate(cbo(1), Val(GetSetting(App.ProductName, "����", "����ģ���", "0")), True)

    
    cbo(4).Text = GetSetting(App.ProductName, "��Ϣ", "������Ϣ���", "")
    rtbSend.Text = GetSetting(App.ProductName, "��Ϣ", "������Ϣ����", "")
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call SaveSetting(App.ProductName, "����", "����ϵͳ��", cbo(0).ItemData(cbo(0).ListIndex))
    Call SaveSetting(App.ProductName, "����", "����ģ���", cbo(1).ItemData(cbo(1).ListIndex))
   
    
    Call SaveSetting(App.ProductName, "��Ϣ", "������Ϣ����", rtbSend.Text)
    Call SaveSetting(App.ProductName, "��Ϣ", "������Ϣ���", cbo(4).Text)
    
    If mblnStarted = True Then
        If Not (mobjMipModule Is Nothing) Then
            mobjMipModule.CloseMessage
            Call gobjComLib.DelMipModule(mobjMipModule)
            Set mobjMipModule = Nothing
        End If
        mblnStarted = False
    End If

End Sub


Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objExtendedBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim strList As String
    Dim strListName() As String
    Dim i As Long
    Dim blnChck As Boolean
    Dim strTmp As String
    
    '��ʼ����
    '------------------------------------------------------------------------------------------------------------------
    Call CommandBarInit(cbsMain)
    cbsMain.VisualTheme = xtpThemeWhidbey
    cbsMain.Options.LargeIcons = False
    Set cbsMain.Icons = Me.ImageManager1.Icons
    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ

    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    
    '����������:������������
    '------------------------------------------------------------------------------------------------------------------

    Set objBar = cbsMain.Add("��׼", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
        
    Set objControl = NewToolBar(objBar, xtpControlLabel, 1, "")
    Set objControl = NewToolBar(objBar, xtpControlLabel, 0, "����ϵͳ")
    
    Set cbrCustom = NewToolBar(objBar, xtpControlCustom, 0, "")
    cbrCustom.Handle = picBack(0).hWnd
    
    Set objControl = NewToolBar(objBar, xtpControlLabel, 0, "����ģ��")

    Set cbrCustom = NewToolBar(objBar, xtpControlCustom, 0, "")
    cbrCustom.Handle = picBack(1).hWnd

    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Reuse, "����ģ��")
    objControl.IconId = 2
    
    Set objControl = NewToolBar(objBar, xtpControlLabel, 0, "��Ϣ��ʶ", True)
    Set cbrCustom = NewToolBar(objBar, xtpControlCustom, 0, "")
    cbrCustom.Handle = picBack(3).hWnd
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Send, "������Ϣ", True)
    objControl.IconId = 3
End Function

Private Sub picBack_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 2
        rtbSend.Move 15, 15, picBack(Index).Width - 30, picBack(Index).Height - 30
    End Select
End Sub
