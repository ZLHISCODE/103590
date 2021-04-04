VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMedicalStationCallBack 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "体检复查随访"
   ClientHeight    =   2280
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6015
   Icon            =   "frmMedicalStationCallBack.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txt 
      Height          =   300
      Left            =   2295
      TabIndex        =   3
      Top             =   1635
      Width           =   1680
   End
   Begin VB.CheckBox chk2 
      Caption         =   "需要随访(&2)"
      Height          =   195
      Left            =   930
      TabIndex        =   2
      Top             =   1695
      Width           =   1320
   End
   Begin VB.PictureBox picConver 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2340
      ScaleHeight     =   240
      ScaleWidth      =   1605
      TabIndex        =   8
      Top             =   1245
      Width           =   1605
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4755
      TabIndex        =   4
      Top             =   210
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4755
      TabIndex        =   5
      Top             =   690
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   3570
      Left            =   4635
      TabIndex        =   7
      Top             =   -315
      Width           =   30
   End
   Begin VB.CheckBox chk 
      Caption         =   "需要复查(&1)"
      Height          =   180
      Left            =   930
      TabIndex        =   0
      Top             =   1275
      Width           =   1305
   End
   Begin MSComCtl2.DTPicker dtp 
      Height          =   300
      Left            =   2295
      TabIndex        =   1
      Top             =   1215
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "yyyy年MM月dd日"
      Format          =   164888579
      CurrentDate     =   38358
   End
   Begin VB.Label Label1 
      Caption         =   "个月"
      Height          =   195
      Left            =   4020
      TabIndex        =   9
      Top             =   1695
      Width           =   405
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "根据体检病人的体检情况，确定该病人是否需要复查，如果需要，请在下面“需要复查”前画上√并输入复查时间。"
      Height          =   540
      Index           =   0
      Left            =   915
      TabIndex        =   6
      Top             =   285
      Width           =   3600
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   210
      Picture         =   "frmMedicalStationCallBack.frx":000C
      Top             =   255
      Width           =   480
   End
End
Attribute VB_Name = "frmMedicalStationCallBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnOK As Boolean
Private mstrDate As String
Private mlngTime As Long

Public Function ShowEdit(ByVal frmMain As Object, ByRef strDate As String, ByRef lngTime As Long) As Boolean
    
    mblnOK = False
    
    mstrDate = strDate
    mlngTime = lngTime
    
    If mstrDate <> "" Then
        chk.Value = 1
        dtp.Value = Format(mstrDate, dtp.CustomFormat)
    Else
        chk.Value = 0
        dtp.Value = Format(DateAdd("d", 30, CDate(zlDatabase.Currentdate)), dtp.CustomFormat)
    End If
    
    If mlngTime > 0 Then
        chk2.Value = 1
        txt.Text = mlngTime
    Else
        chk2.Value = 0
        txt.Text = ""
    End If
    
    Me.Show 1, frmMain
    
    If mblnOK Then
        strDate = mstrDate
        lngTime = mlngTime
    End If
    
    ShowEdit = mblnOK
End Function

Private Sub chk_Click()
    
    dtp.Enabled = (chk.Value = 1)
    picConver.Visible = Not dtp.Enabled
    
End Sub

Private Sub chk_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk2_Click()
    txt.Enabled = (chk2.Value = 1)
End Sub

Private Sub chk2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If chk.Value = 1 Then
        
        mstrDate = Format(dtp.Value, "yyyy-MM-dd")
        
        If mstrDate <= Format(zlDatabase.Currentdate, "yyyy-MM-dd") Then
            ShowSimpleMsg "复查日期必须大于当前日期！"
            Exit Sub
        End If
        
    Else
        mstrDate = ""
    End If
    
    mlngTime = 0
    
    If chk2.Value = 1 Then
        If Val(txt.Text) <= 0 Then
            ShowSimpleMsg "随访期限必须大于等于1个月！"
            Exit Sub
        End If
        
        mlngTime = Val(txt.Text)
    End If
    
    mblnOK = True
    
    Unload Me
End Sub

Private Sub dtp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
