VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmDockInTend_Data 
   BorderStyle     =   0  'None
   Caption         =   "数据页面"
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
   StartUpPosition =   3  '窗口缺省
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
         Caption         =   "没有数据"
         BeginProperty Font 
            Name            =   "宋体"
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
Public Event SelPageChange()     '页面变化时通知主程序做相应处理
Public Event AfterDataChanged(ByVal blnChange As Boolean)
Public Event AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean, ByVal blnSign As Boolean, ByVal blnArchive As Boolean)

Public Sub InitData(ByVal objParent As Object, ByVal strPrivs As String)
    Set mobjParent = objParent
    mstrPrivs = strPrivs
End Sub

Private Sub Form_Activate()
    RaiseEvent Activate
End Sub

Private Sub Form_Load()
    With Me.tbcData
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .ShowIcons = True
            .DisableLunaColors = False
            .Position = xtpTabPositionTop
        End With
        
        Set mfrmDockInTend_File = New frmDockInTend_File
        Call mfrmDockInTend_File.InitData(Me, mstrPrivs)
'
'        Set mfrmDockInTendData = New frmDockInTendData
'        Call mfrmDockInTendData.InitData(Me, mstrPrivs)
        
        .InsertItem(0, "无数据", picNone.hwnd, 0).Tag = "_无数据"
        .InsertItem(1, "文件内容", mfrmDockInTend_File.hwnd, 0).Tag = "_文件内容"
        '.InsertItem(2, "数据列表", mfrmDockInTendData.hwnd, 0).Tag = "_数据列表"
        
        Call zlInit
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    tbcData.Move 10, 10, Me.Width - 20, Me.Height - 20
    lblNote.Move (Me.Width - lblNote.Width) / 2, (Me.Height - lblNote.Height) / 2
End Sub

Public Sub zlInit()
    tbcData.Item(0).Visible = True
    tbcData.Item(1).Visible = False
    'tbcData.Item(2).Visible = False
    tbcData.Item(0).Selected = True
End Sub

Public Function zlRefreshTend(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal intBaby As Integer, ByVal lngDeptID As Long, ByVal blnEdit As Boolean, _
    Optional ByVal blnDoctorStation As Boolean, Optional ByVal lngKey As Long, Optional bytSel As Byte) As Long
    'bytSel:0-体温单;1-记录单
    
    Call zlInit
    If lngKey = 0 Then Exit Function
    
    tbcData.Item(0).Visible = False
    tbcData.Item(1).Visible = True
    'tbcData.Item(2).Visible = True
    tbcData.Item(1).Selected = True
    Call mfrmDockInTend_File.zlRefresh(lngPatiID, lngPageId, intBaby, lngDeptID, blnEdit, blnDoctorStation, bytSel, lngKey)
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

Public Sub zlViewAnimalHeat(ByVal strPara As String, ByVal bytMode As Byte, ByVal strPrivs As String)
    Call mfrmDockInTend_File.zlViewAnimalHeat(strPara, bytMode, strPrivs)
End Sub

Public Sub zlViewFile(ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, _
    ByVal intBaby As Integer, ByVal blnChildForm As Boolean, ByVal strPrivs As String, ByVal blnEdit As Boolean)
    Call mfrmDockInTend_File.zlViewFile(lngFileID, lngPatiID, lngPageId, lngDeptID, intBaby, blnChildForm, strPrivs, blnEdit)
End Sub
'
'Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
'    Call mfrmDockInTendData.zlExecuteCommandBars(Control)
'End Sub
'
'Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
'    Call mfrmDockInTendData.zlUpdateCommandBars(Control)
'End Sub

Public Function zlPrintDocument(ByVal bytKind As Byte, ByVal bytMode As Byte) As Long
    zlPrintDocument = mfrmDockInTend_File.zlPrintDocument(bytKind, bytMode)
End Function

Public Sub zlSaveDocument(blnSave As Boolean)
    Call mfrmDockInTend_File.SaveData(blnSave)
End Sub

Public Sub zlSignDocument(blnOk As Boolean, blnVerify As Boolean)
    Call mfrmDockInTend_File.SignData(blnOk, blnVerify)
End Sub

Public Sub zlArchiveDocument(blnOk As Boolean)
    Call mfrmDockInTend_File.ArchiveData(blnOk)
End Sub

Private Sub mfrmDockInTend_File_AfterDataChanged(ByVal blnChange As Boolean)
    RaiseEvent AfterDataChanged(blnChange)
End Sub

Private Sub mfrmDockInTend_File_AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean, ByVal blnSign As Boolean, ByVal blnArchive As Boolean)
    RaiseEvent AfterRowColChange(strInfo, blnImportant, blnSign, blnArchive)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmDockInTend_File
'    Unload mfrmDockInTendData
End Sub
