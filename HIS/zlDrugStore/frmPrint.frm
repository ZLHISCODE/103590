VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "续打瓶签"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3765
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   3765
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtEndNO 
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   960
      TabIndex        =   5
      Top             =   2040
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancle 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2400
      TabIndex        =   4
      Top             =   2040
      Width           =   1100
   End
   Begin VB.TextBox txtNO 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker DtpPrint 
      Height          =   300
      Left            =   1080
      TabIndex        =   0
      Top             =   1440
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   114098179
      CurrentDate     =   39998
   End
   Begin VB.Label lblEndNo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "结束序号"
      Height          =   180
      Left            =   240
      TabIndex        =   7
      Top             =   930
      Width           =   720
   End
   Begin VB.Label lblNo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "开始序号"
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   337
      Width           =   720
   End
   Begin VB.Label lblTimeBegin 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "打印时间"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   1500
      Width           =   720
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim rsTemp As Recordset
    Dim strSQL As String
    
    strSQL = "select id from 输液配药记录 where 打印时间=[1] and 打印序号>=[2] and 打印序号<=[3]"
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "", DtpPrint.Value, Val(txtNO.Text), IIf(txtEndNO.Text = "", 10000, Val(txtEndNO.Text)))
    
    Do While Not rsTemp.EOF
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1345_1", Me, _
                        "配药ID=" & rsTemp!Id, _
                        "PrintEmpty=0", 2)
        rsTemp.MoveNext
    Loop
End Sub

Public Sub ShowMe(ByVal frmParent As Object)
    Me.Show 1, frmParent
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    On Error GoTo errHandle
    
    strSQL = "select max(打印时间) 打印时间 from 输液配药记录"
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "Form_Load")
    
    If rsTemp.EOF Then
        DtpPrint.Value = Now
    Else
        If NVL(rsTemp!打印时间) = "" Then
            DtpPrint.Value = Now
        Else
            DtpPrint.Value = rsTemp!打印时间
        End If
    End If
    
    
    Exit Sub
errHandle:
    
End Sub

Private Sub txtEndNO_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtEndNO_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtEndNO.Text) Then
        MsgBox "打印序号请录入数字！", vbInformation + vbOKOnly, gstrSysName
        Cancel = True
    End If
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtNO_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtNO.Text) Then
        MsgBox "打印序号请录入数字！", vbInformation + vbOKOnly, gstrSysName
        Cancel = True
    End If
End Sub
