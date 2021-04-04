VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmDockInTend_Data 
   BorderStyle     =   0  'None
   Caption         =   "����ҳ��"
   ClientHeight    =   5175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   10680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picPrompt 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7575
      ScaleHeight     =   195
      ScaleWidth      =   2235
      TabIndex        =   3
      Top             =   690
      Width           =   2235
      Begin VB.Label lblPrompt 
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   45
         TabIndex        =   4
         Top             =   0
         Width           =   2175
      End
   End
   Begin VB.PictureBox picNone 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2085
      Left            =   7920
      ScaleHeight     =   2085
      ScaleWidth      =   2055
      TabIndex        =   1
      Top             =   1770
      Visible         =   0   'False
      Width           =   2055
      Begin VB.Label lblNote 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "û������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1275
      End
   End
   Begin XtremeSuiteControls.TabControl tbcData 
      Height          =   5115
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      _Version        =   589884
      _ExtentX        =   12938
      _ExtentY        =   9022
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmDockInTend_Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mobjParent As Object
Private WithEvents mfrmDockInTend_File As frmDockInTend_File
Attribute mfrmDockInTend_File.VB_VarHelpID = -1
'Private WithEvents mfrmDockInTendData As frmDockInTendData

Public Event Activate()
Public Event SelPageChange()     'ҳ��仯ʱ֪ͨ����������Ӧ����
Public Event AfterDataChanged(ByVal blnChange As Boolean)
Public Event AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean, ByVal blnSign As Boolean, ByVal blnArchive As Boolean)
Public Event ISChartArchive(ByVal blnArchive As Boolean)
Public Event zlRefreshViewFile()
Public Event StartTimer(ByVal blnStart As Boolean)

Private mbytFontSize As Byte

Public Sub InitData(ByVal objParent As Object, ByVal strPrivs As String)
    Set mobjParent = objParent
    mstrPrivs = strPrivs
End Sub

Private Sub Form_Activate()
    RaiseEvent Activate
End Sub

Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-19 15:16
    '����:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize
End Sub

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-19 15:16
    '����:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim CtlFont As StdFont
    Dim objCtrl As Control
    Dim bytSize As Byte
    bytSize = IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))
    
    Call mfrmDockInTend_File.ReSetFontSize(bytSize)
    
    Me.FontSize = mbytFontSize
    
    Set CtlFont = tbcData.PaintManager.Font
    If CtlFont Is Nothing Then
        Set CtlFont = Me.Font
    End If
    CtlFont.Size = mbytFontSize
    Set tbcData.PaintManager.Font = CtlFont
    tbcData.PaintManager.Layout = xtpTabLayoutAutoSize
    
    lblPrompt.FontSize = mbytFontSize
    Call Form_Resize
End Sub

Private Sub Form_Load()
    With Me.tbcData
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .ShowIcons = True
            .DisableLunaColors = False
            .Position = xtpTabPositionBottom 'xtpTabPositionTop
        End With
        
        Set mfrmDockInTend_File = New frmDockInTend_File
        Call mfrmDockInTend_File.InitData(Me, mstrPrivs)
'
'        Set mfrmDockInTendData = New frmDockInTendData
'        Call mfrmDockInTendData.InitData(Me, mstrPrivs)
        
        .InsertItem(0, "������", picNone.hWnd, 0).Tag = "_������"
        .InsertItem(1, "�ļ�����", mfrmDockInTend_File.hWnd, 0).Tag = "_�ļ�����"
        '.InsertItem(2, "�����б�", mfrmDockInTendData.hwnd, 0).Tag = "_�����б�"
        
        Call zlInit
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    tbcData.Move 10, 10, Me.Width - 20, Me.Height - 20
    lblNote.Move (Me.Width - lblNote.Width) / 2, (Me.Height - lblNote.Height) / 2
    picPrompt.Height = 240
    picPrompt.Move 360 + TextWidth("���ļ����� "), tbcData.Top + tbcData.Height - picPrompt.Height - 80, tbcData.Width - picPrompt.Left
    lblPrompt.Height = TextHeight("��")
    lblPrompt.Width = picPrompt.Width
    lblPrompt.Top = (picPrompt.Height - lblPrompt.Height) \ 2
End Sub

Public Sub zlInit()
    tbcData.Item(0).Visible = True
    tbcData.Item(1).Visible = False
    'tbcData.Item(2).Visible = False
    tbcData.Item(0).Selected = True
End Sub

Public Function zlRefreshTend(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal intBaby As Integer, ByVal lngDeptID As Long, ByVal blnEdit As Boolean, _
    Optional ByVal blnDoctorStation As Boolean, Optional ByVal lngKey As Long, Optional bytSel As Byte, Optional ByVal intCurveReSize As Integer = 0) As Long
    'bytSel:0-���µ�;1-��¼��
    
    Call zlInit
    If lngKey = 0 Then Exit Function
    
    tbcData.Item(0).Visible = False
    tbcData.Item(1).Visible = True
    'tbcData.Item(2).Visible = True
    tbcData.Item(1).Selected = True
    Call mfrmDockInTend_File.zlRefresh(lngPatiID, lngPageId, intBaby, lngDeptID, blnEdit, blnDoctorStation, bytSel, lngKey, intCurveReSize)
End Function

Public Function zlRefreshEPR(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, ByVal blnEdit As Boolean, _
    Optional ByVal blnDoctorStation As Boolean, Optional ByVal lngKey As Long, Optional bytEdit As Byte) As Long
    
    Call zlInit
    If lngKey = 0 Then Exit Function
    
    tbcData.Item(0).Visible = False
    tbcData.Item(1).Visible = True
    'tbcData.Item(2).Visible = False
    tbcData.Item(1).Selected = True
    Call mfrmDockInTend_File.zlRefresh(lngPatiID, lngPageId, 0, lngDeptID, blnEdit, blnDoctorStation, 2, lngKey)
End Function

Public Sub zlViewAnimalHeat(ByVal strPara As String, ByVal bytMode As Byte, ByVal strPrivs As String, ByVal bytSize As Byte)
    Dim blnOK As Boolean
    
    blnOK = mfrmDockInTend_File.zlViewAnimalHeat(strPara, bytMode, strPrivs, bytSize)
    If blnOK = True Then
       RaiseEvent zlRefreshViewFile
    End If
End Sub

Public Sub zlViewCaveData(ByVal intDataEditor As Integer)
    Call mfrmDockInTend_File.zlViewCaveData(intDataEditor)
End Sub

Public Sub zlViewFile(ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, _
    ByVal intBaby As Integer, ByVal blnChildForm As Boolean, ByVal strPrivs As String, ByVal blnEdit As Boolean, ByVal bytSize As Byte)
    Dim blnOK As Boolean
    blnOK = mfrmDockInTend_File.zlViewFile(lngFileID, lngPatiID, lngPageId, lngDeptID, intBaby, blnChildForm, strPrivs, blnEdit, bytSize)
    If blnOK = True Then
       RaiseEvent zlRefreshViewFile
    End If
End Sub
'
'Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
'    Call mfrmDockInTendData.zlExecuteCommandBars(Control)
'End Sub
'
'Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
'    Call mfrmDockInTendData.zlUpdateCommandBars(Control)
'End Sub

Public Sub zlViewpartogram(ByVal strPara As String, ByVal bytMode As Byte, ByVal strPrivs As String, ByVal bytSize As Byte)
    Call mfrmDockInTend_File.zlViewpartogram(strPara, bytMode, strPrivs, bytSize)
End Sub

Public Sub zlViewpartogramEditor(ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, _
    ByVal intBaby As Integer, ByVal strPrivs As String, ByVal bytSize As Byte)
    Call mfrmDockInTend_File.zlViewpartogramEditor(lngFileID, lngPatiID, lngPageId, lngDeptID, intBaby, strPrivs, bytSize)
End Sub

Public Function zlPrintTendFile(ByVal bytKind As Byte, ByVal bytMode As Byte) As Long
    zlPrintTendFile = mfrmDockInTend_File.zlPrintTendFile(bytKind, bytMode)
End Function

Public Sub SignMarker()
    Call mfrmDockInTend_File.SignMarker
End Sub

Public Sub zlSaveDocument(blnSave As Boolean)
    Call mfrmDockInTend_File.SaveData(blnSave)
End Sub

Public Sub zlSignDocument(blnOK As Boolean, blnVerify As Boolean, blnExchange As Boolean)
    Call mfrmDockInTend_File.SignData(blnOK, blnVerify, blnExchange)
End Sub

Public Sub zlArchiveDocument(blnOK As Boolean)
    Call mfrmDockInTend_File.ArchiveData(blnOK)
End Sub

Public Sub ViewReSetFontSize(ByVal intSEL As Integer, ByVal bytSize As Byte)
    Call mfrmDockInTend_File.ViewReSetFontSize(intSEL, bytSize)
End Sub

Private Sub mfrmDockInTend_File_AfterDataChanged(ByVal blnChange As Boolean)
    RaiseEvent AfterDataChanged(blnChange)
End Sub

Private Sub mfrmDockInTend_File_AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean, ByVal blnSign As Boolean, ByVal blnArchive As Boolean)
    lblPrompt.Caption = strInfo
    lblPrompt.ForeColor = IIf(blnImportant, &HFF&, &H80000008)
    RaiseEvent AfterRowColChange(strInfo, blnImportant, blnSign, blnArchive)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmDockInTend_File
'    Unload mfrmDockInTendData
End Sub

Private Sub mfrmDockInTend_File_ISChartArchive(ByVal blnArchive As Boolean)
    lblPrompt.Caption = ""
    lblPrompt.ForeColor = &H80000008
    RaiseEvent ISChartArchive(blnArchive)
End Sub

Public Sub BulkPrintDocument(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, ByVal intBaby As Integer)
    Call mfrmDockInTend_File.BulkPrintDocument(lngPatiID, lngPageId, lngDeptID, intBaby)
End Sub

Private Sub mfrmDockInTend_File_StartTimer(ByVal blnStart As Boolean)
    RaiseEvent StartTimer(blnStart)
End Sub
