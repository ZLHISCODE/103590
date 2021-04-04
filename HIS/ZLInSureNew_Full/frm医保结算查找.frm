VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm医保结算查找 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "结算过滤"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   Icon            =   "frm医保结算查找.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1800
      TabIndex        =   7
      Top             =   2100
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3000
      TabIndex        =   8
      Top             =   2100
      Width           =   1100
   End
   Begin VB.Frame fraScope 
      Caption         =   "时间范围"
      Height          =   1815
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   4155
      Begin VB.TextBox txt医保号 
         Height          =   300
         Left            =   1830
         TabIndex        =   6
         Top             =   1350
         Width           =   2115
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   1830
         TabIndex        =   4
         Top             =   870
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   19791875
         CurrentDate     =   36279
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1830
         TabIndex        =   2
         Top             =   390
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   19791875
         CurrentDate     =   36279
         MinDate         =   2
      End
      Begin VB.Label lbl医保号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医保号(&A)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   960
         TabIndex        =   5
         Top             =   1410
         Width           =   810
      End
      Begin VB.Label lblTimeStop 
         AutoSize        =   -1  'True
         Caption         =   "结束时间(&E)"
         Height          =   180
         Left            =   780
         TabIndex        =   3
         Top             =   930
         Width           =   990
      End
      Begin VB.Label lblTimeStart 
         AutoSize        =   -1  'True
         Caption         =   "开始时间(&B)"
         Height          =   180
         Left            =   780
         TabIndex        =   1
         Top             =   450
         Width           =   990
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Left            =   150
         Picture         =   "frm医保结算查找.frx":000C
         Top             =   420
         Width           =   480
      End
   End
End
Attribute VB_Name = "frm医保结算查找"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnOK As Boolean
Dim mdatBegin As Date, mdatEnd As Date
Private mstrCard As String

Private Sub dtpBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then dtpEnd.SetFocus
End Sub

Private Sub dtpEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdOK.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If dtpBegin.Value > dtpEnd.Value Then
        MsgBox "开始时间大于结束时间了。", vbExclamation, gstrSysName
        Exit Sub
    End If
    mdatBegin = dtpBegin.Value
    mdatEnd = dtpEnd.Value
    mstrCard = Trim(UCase(txt医保号.Text))
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnOK = False
End Sub

Public Function GetTimeScope(datBegin As Date, datEnd As Date, strCard As String, ByVal frmOwner As Form, _
                Optional ByVal strCaption As String, Optional blnStrict As Boolean = True) As Boolean
                
    If strCaption <> "" Then
        frm医保结算查找.Caption = strCaption
    End If
    
    dtpBegin.Value = datBegin
    dtpEnd.Value = datEnd
    txt医保号.Text = strCard
    
    If blnStrict = True Then
        '严格限制日期
        dtpBegin.MaxDate = CDate(Format(zldatabase.Currentdate, "yyyy-MM-dd") & " 23:59:59")
        dtpEnd.MaxDate = dtpBegin.MaxDate
    End If
    frm医保结算查找.Show vbModal, frmOwner
    
    GetTimeScope = mblnOK
    If mblnOK = True Then
        datBegin = mdatBegin
        datEnd = mdatEnd
        strCard = mstrCard
    End If
End Function


