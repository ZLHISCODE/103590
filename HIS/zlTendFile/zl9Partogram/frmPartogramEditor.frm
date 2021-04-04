VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPartogramEditor 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�������ݱ༭"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11415
   Icon            =   "frmPartogramEditor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picPrompt 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1410
      ScaleHeight     =   285
      ScaleWidth      =   9645
      TabIndex        =   2
      Top             =   7710
      Width           =   9645
      Begin VB.Label lblPrompt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   15
         TabIndex        =   3
         Top             =   60
         Width           =   10500
      End
   End
   Begin zl9Partogram.usrPartogramEditor usrPartogramEditor 
      Height          =   6045
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   9255
      _extentx        =   16325
      _extenty        =   10663
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7560
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPartogramEditor.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17224
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmPartogramEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
'�������
'######################################################################################################################
Public mblnDoctorStation As Boolean
Public mstrPrivs As String
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mintӤ�� As Integer
Private mlngFileID As Long

Private mblnChange As Boolean                           '�޸ı�־
Private mblnSign As Boolean                             'ǩ����־
Private mblnArchive As Boolean                          '�鵵��־
Private mlng���� As Long
Private mblnSave As Boolean
Private mbytFontSize As Byte

'��ȡ��ǰѡ�е�Ӥ����
Public Property Get FileNumIndex() As Long
    FileNumIndex = mlng����
End Property

Public Property Let FileNumIndex(FileIndex As Long)
    mlng���� = FileIndex
End Property

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_Save  '����
        Call SaveData
    Case conMenu_Edit_Transf_Cancle  'ȡ��
        Call CancelData
    Case conMenu_Tool_Sign  'ǩ��
        Call SignData(False)
    Case conMenu_Tool_SignAuditAffirm
        Call SignData(True)
    Case conMenu_Tool_SignEarse  'ȡ��ǩ��
        Call UnSignData(False)
    Case conMenu_Tool_SignAuditCancel
        Call UnSignData(True)
    Case conMenu_File_Exit
        Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Call Form_Resize
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_Save  '����
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "����ͼ��ͼ") > 0)
        Control.Enabled = mblnChange And Not mblnArchive And Control.Visible
    Case conMenu_Edit_Transf_Cancle  'ȡ��
        Control.Visible = Not mblnDoctorStation And Not gblnMoved
        Control.Enabled = mblnChange And Not mblnArchive And Control.Visible
    Case conMenu_Tool_Sign, conMenu_Tool_SignEarse, conMenu_Tool_SignAuditAffirm, conMenu_Tool_SignAuditCancel 'ǩ��
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "����ͼ��ͼ") > 0)
        Control.Enabled = Not mblnArchive And Not mblnChange And Control.Visible
    End Select
End Sub


Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    Call Me.cbsMain.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    
    If Me.WindowState = vbMinimized Then Exit Sub
    Err = 0: On Error Resume Next
    With usrPartogramEditor
        .Left = 0
        .Top = lngScaleTop
        .Width = lngScaleRight
        .Height = lngScaleBottom - lngScaleTop - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    With picPrompt
        .Visible = stbThis.Visible
        .Top = stbThis.Top + 50
        .Height = stbThis.Height - 100
        .Left = stbThis.Panels(2).Left + 50
        .Width = stbThis.Panels(2).Width - 100
    End With
    With lblPrompt
        .Visible = stbThis.Visible
        .Width = picPrompt.Width
        .Height = TextHeight("��")
        .Top = (picPrompt.Height - .Height) \ 2
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-20 15:15:00
    '����:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim bytFontSize As Byte
    bytFontSize = mbytFontSize
    
    Me.FontSize = bytFontSize
    Me.FontName = "����"
    
    Set CtlFont = cbsMain.Options.Font
    If CtlFont Is Nothing Then
        Set CtlFont = Me.Font
    End If
    CtlFont.Size = bytFontSize
    Set cbsMain.Options.Font = CtlFont
    lblPrompt.FontSize = bytFontSize
    cbsMain.RecalcLayout
End Sub

Public Function ShowMe(ByVal frmParent As Object, ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, _
    ByVal lngDeptID As Long, ByVal intBaby As Integer, ByVal strPrivs As String, _
    Optional ByVal blnEdit As Boolean, Optional ByVal bytSize As Byte = 0) As Boolean
    '******************************************************************************************************************
    '���ܣ� ��ʾ�����¼�ļ�����
    '������ frmParent           �ϼ��������
    '       lngFileID           �����ļ���ʽ���
    '       lngPatiID           ����id
    '       lngPageID           ��ҳid
    '       intBaby             Ӥ����־
    '���أ� ��
    '******************************************************************************************************************
'    Dim bln������ As Boolean
    
    Err = 0
    On Error GoTo errHand
    mlngFileID = lngFileID
    mlng����ID = lngPatiID
    mlng��ҳID = lngPageId
    mintӤ�� = 0
    mlng���� = 1
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    If Not frmParent Is Nothing Then mlng���� = frmParent.FileNumIndex
    mblnSave = False
    mstrPrivs = strPrivs
    Me.WindowState = 2
    Call MainDefMenus
    Call ReSetFontSize
    
    Call usrPartogramEditor.ShowMe(Me, lngFileID, lngPatiID, lngPageId, lngDeptID, intBaby, strPrivs, blnEdit, bytSize)
    
    '������ʾ
    If frmParent Is Nothing Then
        Me.Show vbModal
    Else
        Me.Show vbModal, frmParent
    End If
    
    If Not frmParent Is Nothing Then frmParent.FileNumIndex = mlng����
    ShowMe = mblnSave
    
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub MainDefMenus()
    Dim cbrControl As Object
    Dim cbrToolBar As Object
    
    CommandBarsGlobalSettings.App = App
    Set Me.cbsMain.Icons = zlCommFun.GetPubIcons
    With Me.cbsMain.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 24, 24
        .LargeIcons = True
    End With
    
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    cbsMain.ActiveMenuBar.Visible = False
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Sign, "��¼ǩ��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "ȡ��ǩ��"): cbrControl.IconId = 229
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignAuditAffirm, "�ϼ���ǩ"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignAuditCancel, "ȡ����ǩ")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�"): cbrControl.BeginGroup = True
    End With
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
End Sub

Private Sub usrPartogramEditor_AfterDataChanged(ByVal blnChange As Boolean)
    mblnChange = blnChange
End Sub

Private Sub usrPartogramEditor_AfterDataSave(ByVal blnSave As Boolean)
    mblnSave = blnSave
End Sub

Private Sub usrPartogramEditor_AfterFileIndex(ByVal lngFileIndex As Long)
    mlng���� = lngFileIndex
End Sub

Private Sub usrPartogramEditor_AfterPartogramInfo(ByVal lngFlieId As Long, ByVal lngFileIndex As Long, ByVal lngFileFormatID As Long, ByVal rsPartogram As ADODB.Recordset)
    If Not frmPartogramInfo.ShowMe(Me, lngFlieId, lngFileIndex, lngFileFormatID, rsPartogram, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))) Then Exit Sub
    'ˢ��Ҫ����Ϣ
    Call usrPartogramEditor.zlRefresh(False)
    mblnSave = True
End Sub

Private Sub usrPartogramEditor_AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean, ByVal blnSign As Boolean, ByVal blnArchive As Boolean)
    mblnArchive = blnArchive
    mblnSign = blnSign
    lblPrompt.Caption = strInfo
    lblPrompt.ForeColor = IIf(blnImportant, &HFF&, &H80000008)
End Sub

Public Function SaveData() As Boolean
    SaveData = usrPartogramEditor.SaveME
End Function

Public Function CancelData() As Boolean
    CancelData = usrPartogramEditor.CancelMe
End Function

Public Sub SignData(blnVerify As Boolean)
    Call usrPartogramEditor.SignMe(blnVerify)
End Sub

Public Sub UnSignData(blnVerify As Boolean)
    Call usrPartogramEditor.UnSignMe(blnVerify)
End Sub
