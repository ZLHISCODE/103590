VERSION 5.00
Begin VB.Form frmCaseTendEditForSinglePerson 
   BorderStyle     =   0  'None
   Caption         =   "单病人多时点快速录入"
   ClientHeight    =   5175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9120
   Icon            =   "frmCaseTendEditForSinglePerson.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmCaseTendEditForSinglePerson.frx":000C
   ScaleHeight     =   5175
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   Begin zlRichEPR.usrTendEditor usrTendEditor1 
      Height          =   4515
      Left            =   -30
      TabIndex        =   0
      Top             =   60
      Width           =   8565
      _extentx        =   4895
      _extenty        =   3307
   End
End
Attribute VB_Name = "frmCaseTendEditForSinglePerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event AfterDataChanged()
Public Event AfterArchiveChanged()
Public Event AfterRefresh()
Public Event AfterSelChange(ByVal lngCert As Long, ByVal strCertLevel As String)
Public Event DbClick(ByVal strData As String)
Public Event AfterRowColChange(ByVal strInfo As String)

Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-18 15:16
    '问题:51746
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    Call usrTendEditor1.ReSetFontSize(bytSize)
End Sub

Public Sub SetEditable(ByVal blnEditable As Boolean)
    usrTendEditor1.mblnEditable = blnEditable
End Sub

Public Function GetCopyData() As String
    GetCopyData = usrTendEditor1.GetCopyData
End Function

Public Function IsPigeonhole() As Boolean
    IsPigeonhole = usrTendEditor1.IsPigeonhole
End Function

Private Sub Form_Resize()
    On Error Resume Next
    usrTendEditor1.Left = 0
    usrTendEditor1.Top = 0
    usrTendEditor1.Width = Me.ScaleWidth
    usrTendEditor1.Height = Me.ScaleHeight
End Sub

Private Sub usrTendEditor1_AfterArchiveChanged()
    RaiseEvent AfterArchiveChanged
End Sub

Private Sub usrTendEditor1_AfterDataChanged()
    RaiseEvent AfterDataChanged
End Sub

Private Sub usrTendEditor1_AfterRefresh()
    RaiseEvent AfterRefresh
End Sub

Private Sub usrTendEditor1_AfterRowColChange(ByVal strInfo As String)
    RaiseEvent AfterRowColChange(strInfo)
End Sub

Private Sub usrTendEditor1_AfterSelChange(ByVal lngCert As Long, ByVal strCertLevel As String)
    RaiseEvent AfterSelChange(lngCert, strCertLevel)
End Sub

Private Sub usrTendEditor1_DbClick(ByVal strData As String)
    RaiseEvent DbClick(strData)
End Sub

Public Sub ShowMe(ByVal frmParent As Form, ByVal lngPatiID As Long, ByVal lngPageId As Long, lngDeptId As Long, _
    Optional ByVal intBaby As Integer = 0, Optional ByVal byt护理级别 As Byte = 3, Optional ByVal strPrivs As String, _
    Optional ByVal blnCancel As Boolean = False, Optional ByVal blnEditable As Boolean = True)

    Call usrTendEditor1.ShowMe(frmParent, lngPatiID, lngPageId, lngDeptId, intBaby, byt护理级别, strPrivs, blnCancel, blnEditable)
End Sub

Public Sub ArchiveMe()
    Call usrTendEditor1.ArchiveMe
End Sub

Public Sub UnArchiveMe()
    Call usrTendEditor1.UnArchiveMe
End Sub

Public Sub SignMe()
    Call usrTendEditor1.SignMe
End Sub

Public Sub UnSignMe()
    Call usrTendEditor1.UnSignMe
End Sub

Public Sub SignMarker()
    Call usrTendEditor1.SignMarker
End Sub

Public Function SaveME() As Boolean
    SaveME = usrTendEditor1.SaveME
End Function

Public Function CancelMe() As Boolean
    CancelMe = usrTendEditor1.CancelMe
End Function
