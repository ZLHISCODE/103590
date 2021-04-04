VERSION 5.00
Begin VB.Form frmDockInTend_File 
   BorderStyle     =   0  'None
   Caption         =   "文件内容"
   ClientHeight    =   4410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picFile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3435
      Left            =   90
      ScaleHeight     =   3435
      ScaleWidth      =   6915
      TabIndex        =   0
      Top             =   360
      Width           =   6915
   End
End
Attribute VB_Name = "frmDockInTend_File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'######################################################################################################################

Private mintSel As Integer          '当前选中状态
Private mfrmTendBody As Object
Private WithEvents mfrmTendFile As frmTendFileEditor
Attribute mfrmTendFile.VB_VarHelpID = -1

Private mobjParent As Object
Private mblnFirst As Boolean
Private mstrPrivs As String                             '当前使用者对本程序(1255)的权限串
Private mlngPatiId As Long                              '病人id
Private mlngPageId As Long                              '主页id
Private mlngDeptId As Long                              '当前操作科室id，如病人科室和当前科室不一致，则不能操作归档外的功能
Private mintBaby As Integer
Private mblnEdit As Boolean                             '是否允许操作，通常由上级程序根据当前操作科室是否当前病人病区决定。
Private mblnDoctorStation As Boolean

Private rsTemp As New ADODB.Recordset
Private mfrmMain As Object
Private mblnTendArchive As Boolean

Private Enum enuSEL
    体温单
    记录单
End Enum

Public Event Activate()
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event AfterDataChanged(ByVal blnChange As Boolean)
Public Event AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean, ByVal blnSign As Boolean, ByVal blnArchive As Boolean)

Private Sub Form_Activate()
    If mblnFirst Then
        mfrmTendBody.Show
        mfrmTendFile.Show
        mblnFirst = False
    End If
End Sub

Private Sub Form_Load()
    mblnFirst = True
    '创建体温部件与护理记录单
    If Not CreateBodyEditor Then Exit Sub
    Set mfrmTendBody = gobjBodyEditor.GetNewTendBody
    Set mfrmTendFile = New frmTendFileEditor
    '将体温部件设置为无标题栏的子窗体
    Call FormSetCaption(mfrmTendBody, False, False)
    Call mfrmTendBody.zlInit
    Load mfrmTendBody
    Load mfrmTendFile
    '设置其父窗体
    Call SetParent(mfrmTendBody.hwnd, picFile.hwnd)
    Call SetParent(mfrmTendFile.hwnd, picFile.hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmTendBody Is Nothing Then Unload mfrmTendBody
    Unload mfrmTendFile
End Sub

Public Sub InitData(ByVal objParent As Object, ByVal strPrivs As String)
    mstrPrivs = strPrivs
    Set mobjParent = objParent
End Sub

Public Function zlRefresh(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal intBaby As Integer, ByVal lngDeptID As Long, ByVal blnEdit As Boolean, _
    Optional ByVal blnDoctorStation As Boolean, Optional ByVal intSel As Integer, Optional ByVal lngKey As Long) As Long
    
    mfrmTendBody.Visible = (intSel = 体温单)
    mfrmTendFile.Visible = (intSel = 记录单)
    
    Select Case intSel
    Case 0
        Call mfrmTendBody.zlRefresh(Me, lngPatiID & ";" & lngPageId & ";" & lngDeptID & ";" & lngKey & ";0;" & blnEdit & ";" & intBaby, mstrPrivs)
    Case 1
        Call mfrmTendFile.ShowMe(Nothing, lngKey, lngPatiID, lngPageId, lngDeptID, intBaby, True, mstrPrivs, blnEdit)
    End Select
End Function

Private Sub Form_Resize()
    On Error Resume Next
    
    picFile.Move 10, 10, Me.Width - 20, Me.Height - 20
End Sub

Private Sub mfrmTendFile_AfterDataChanged(ByVal blnChange As Boolean)
    RaiseEvent AfterDataChanged(blnChange)
End Sub

Private Sub mfrmTendFile_AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean, ByVal blnSign As Boolean, ByVal blnArchive As Boolean)
    RaiseEvent AfterRowColChange(strInfo, blnImportant, blnSign, blnArchive)
End Sub

Private Sub picFile_Resize()
    On Error Resume Next
    
    mfrmTendBody.Move 0, 0, picFile.Width, picFile.Height
    mfrmTendFile.Move 0, 0, picFile.Width, picFile.Height
End Sub

Public Sub zlViewAnimalHeat(ByVal strPara As String, ByVal bytMode As Byte, ByVal strPrivs As String)
    Call gobjBodyEditor.GetNewTendBody.ShowEdit(Me, strPara, bytMode, strPrivs)
End Sub

Public Sub zlViewFile(ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, _
    ByVal intBaby As Integer, ByVal blnChildForm As Boolean, ByVal strPrivs As String, ByVal blnEdit As Boolean)
    Call frmTendFileEditor.ShowMe(Me, lngFileID, lngPatiID, lngPageId, lngDeptID, intBaby, blnChildForm, strPrivs, blnEdit)
End Sub

Public Function zlPrintDocument(ByVal bytKind As Byte, Optional ByVal bytMode As Byte = 2) As Long
    '功能:  打印体温单;bytMode：2-打印
    Dim strSQL As String
    
    If bytKind = 1 Then
        '体温单(返回值:1-成功;2-打印)
        zlPrintDocument = mfrmTendBody.zlPrintBody(bytMode)
    Else
        '护理记录单
        Call mfrmTendFile.zlPrintTend(bytMode)
    End If
End Function

Public Sub SaveData(blnSave As Boolean)
    If blnSave Then
        blnSave = mfrmTendFile.SaveData
    Else
        blnSave = mfrmTendFile.CancelData
    End If
End Sub

Public Sub SignData(blnOK As Boolean, blnVerify As Boolean)
    If blnOK Then
        Call mfrmTendFile.SignData(blnVerify)
    Else
        Call mfrmTendFile.UnSignData(blnVerify)
    End If
End Sub

Public Sub ArchiveData(blnOK As Boolean)
    If blnOK Then
        Call mfrmTendFile.ArchiveData
    Else
        Call mfrmTendFile.UnArchiveData
    End If
End Sub

