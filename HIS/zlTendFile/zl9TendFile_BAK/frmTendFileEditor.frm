VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTendFileEditor 
   BackColor       =   &H00FFFFFF&
   Caption         =   "护理记录查阅"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11415
   Icon            =   "frmTendFileEditor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   Begin zl9TendFile.usrTendFileEditor usrTendFileEditor 
      Height          =   6045
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   10663
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
            Picture         =   "frmTendFileEditor.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17224
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmTendFileEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
'窗体变量
'######################################################################################################################
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mint婴儿 As Integer
Private mlngFileID As Long
Private mblnChildForm As Boolean
Private mblnStartUp As Boolean

Public WithEvents zlEvent_Print As zl9TendFilePrint.zlPrintMethod
Attribute zlEvent_Print.VB_VarHelpID = -1
Public Event zlAfterPrint(ByVal lngFileID As Long)
Public Event AfterDataChanged(ByVal blnChange As Boolean)
Public Event AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean, ByVal blnSign As Boolean, ByVal blnArchive As Boolean)

Private Sub Form_Load()
    mblnStartUp = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With usrTendFileEditor
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChildForm = False Then Call SaveWinState(Me, App.ProductName)
End Sub

Public Sub ShowMe(ByVal frmParent As Form, ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, _
    ByVal lngDeptID As Long, ByVal intBaby As Integer, ByVal blnChildForm As Boolean, ByVal strPrivs As String, _
    Optional ByVal blnEdit As Boolean)
    '******************************************************************************************************************
    '功能： 显示护理记录文件内容
    '参数： frmParent           上级窗体对象
    '       lngFileID           护理文件格式句柄
    '       lngPatiID           病人id
    '       lngPageID           主页id
    '       intBaby             婴儿标志
    '返回： 无
    '******************************************************************************************************************
'    Dim bln护理级别 As Boolean
    
    Err = 0
    On Error GoTo errHand
    mlngFileID = lngFileID
    mblnChildForm = blnChildForm
    mlng病人ID = lngPatiID
    mlng主页ID = lngPageId
    mint婴儿 = intBaby
    
    If mblnChildForm Then
        If mblnStartUp Then
            Call FormSetCaption(Me, False, False)
            
            stbThis.Visible = Not mblnChildForm
            mblnStartUp = False
        End If
    Else
        Me.WindowState = 2
        blnEdit = False
    End If
    
    Call usrTendFileEditor.ShowMe(Me, lngFileID, lngPatiID, lngPageId, lngDeptID, intBaby, strPrivs, blnEdit)
    
    '窗体显示
    If blnChildForm = False Then
        If frmParent Is Nothing Then
            Me.Show vbModal
        Else
            Me.Show vbModal, frmParent
        End If
        Unload Me
        Exit Sub
    End If
    
    Exit Sub
    
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub usrTendFileEditor_AfterDataChanged(ByVal blnChange As Boolean)
    RaiseEvent AfterDataChanged(blnChange)
End Sub

Private Sub usrTendFileEditor_AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean, ByVal blnSign As Boolean, ByVal blnArchive As Boolean)
    RaiseEvent AfterRowColChange(strInfo, blnImportant, blnSign, blnArchive)
End Sub

Private Sub zlEvent_Print_zlAfterPrint()
    RaiseEvent zlAfterPrint(mlngFileID)
End Sub

Public Function SaveData() As Boolean
    SaveData = usrTendFileEditor.SaveME
End Function

Public Function CancelData() As Boolean
    CancelData = usrTendFileEditor.CancelMe
End Function

Public Sub SignData(blnVerify As Boolean)
    Call usrTendFileEditor.SignMe(blnVerify)
End Sub

Public Sub UnSignData(blnVerify As Boolean)
    Call usrTendFileEditor.UnSignMe(blnVerify)
End Sub

Public Sub ArchiveData()
    Call usrTendFileEditor.ArchiveMe
End Sub

Public Sub UnArchiveData()
    Call usrTendFileEditor.UnArchiveMe
End Sub

Public Function zlPrintTend(Optional ByVal bytMode As Byte = 2, Optional ByVal strPrintDeviceName As String) As Boolean
    '1-预览,2-打印
    
    Select Case bytMode
    Case 1
        Call zlRptPrint(2, strPrintDeviceName)
    Case 2
        Call zlRptPrint(1, strPrintDeviceName)
    Case 3
        Call zlRptPrint(3, strPrintDeviceName)
    End Select
End Function

Private Sub zlRptPrint(ByVal bytMode As Byte, Optional ByVal strPrintDeviceName As String)
    Dim objPrint As New zlPrintTends, objAppRow As zlTabAppRow
    Dim lngWidth As Long
    
    If zlEvent_Print Is Nothing Then
        Set zlEvent_Print = VBA.GetObject("", "zl9TendFilePrint.zlPrintMethod")
    End If
    
    Call zlEvent_Print.InitPrint(gcnOracle, gstrDBUser)
    If strPrintDeviceName = "" Then
        bytMode = zlEvent_Print.zlPrintAsk(mlng病人ID, mlng主页ID, mint婴儿, mlngFileID)
    Else
        SaveSetting "ZLSOFT", "公共模块\zl9PrintMode\Default", "DeviceName", strPrintDeviceName
        bytMode = zlEvent_Print.zlPrintAsk(mlng病人ID, mlng主页ID, mint婴儿, mlngFileID, True)
    End If
    
    If bytMode <> 0 Then zlEvent_Print.zlPrintOrViewTends (strPrintDeviceName <> ""), bytMode
End Sub
