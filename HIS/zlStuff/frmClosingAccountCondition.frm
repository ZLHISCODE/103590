VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClosingAccountCondition 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "手工结存设置"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4440
   Icon            =   "frmClosingAccountCondition.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2040
      TabIndex        =   1
      Top             =   1800
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3235
      TabIndex        =   2
      Top             =   1800
      Width           =   1100
   End
   Begin VB.Frame fraConditiom 
      Caption         =   "期末日期选择"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.OptionButton opt指定时间 
         Caption         =   "指定时间"
         Height          =   180
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton opt当前时间 
         Caption         =   "当前时间"
         Height          =   180
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Left            =   1560
         TabIndex        =   5
         Top             =   900
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
         Format          =   194510851
         CurrentDate     =   36901
      End
   End
End
Attribute VB_Name = "frmClosingAccountCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng库房ID As Long
Private mblnSelect As Boolean
Private mstr结存时间 As String

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Public Function GetCondition(frmMain As Form, ByVal lng库房ID As Long, ByRef str结存时间) As Boolean
    '选择当前时间，返回str结存时间=""；选择指定时间，返回str结存时间为具体时间；
    'GetCondition：true-结存；false-取消结存
    mlng库房ID = lng库房ID
    mblnSelect = False
    
    Me.Show 1, frmMain
    
    str结存时间 = mstr结存时间
    GetCondition = mblnSelect
    
End Function

Private Sub CmdSave_Click()
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    mstr结存时间 = IIf(opt指定时间.Value = True, Format(dtpDate.Value, "yyyy-MM-dd hh:mm:ss"), "")
    
    If opt指定时间.Value = True Then '指定时间要检查时间是否大于上期期末时间
        gstrSQL = " Select Max(期末日期) 上期末日期, Max(期末日期) + 1 / 24 / 60 / 60 期初日期 From 材料结存记录 Where 库房id = [1] And 取消人 Is Null "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "", mlng库房ID)
        
        If rsTemp.EOF = True Or IsNull(rsTemp!上期末日期) = True Then
            MsgBox "该库房没有结存记录，请先初始化！", vbInformation, gstrSysName
            mblnSelect = False
        Else
            If mstr结存时间 < Format(rsTemp!期初日期, "yyyy-MM-dd hh:mm:ss") Then
                MsgBox "指定时间必须大于上期期末日期（" & Format(rsTemp!上期末日期, "yyyy-MM-dd hh:mm:ss") & "）！", vbInformation, gstrSysName
                dtpDate.SetFocus
                Exit Sub
            End If
            
            If mstr结存时间 > Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss") Then
                MsgBox "指定时间不能大于当前系统时间！", vbInformation, gstrSysName
                dtpDate.SetFocus
                Exit Sub
            End If
            
            mblnSelect = True
        End If
    Else '选择当前时间不用校验
        mblnSelect = True
    End If
    
    Unload Me
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()

    dtpDate.Value = Format(zlDatabase.Currentdate, dtpDate.CustomFormat)
    dtpDate.Enabled = opt指定时间.Value = True
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub opt当前时间_Click()
    dtpDate.Enabled = opt指定时间.Value = True
End Sub

Private Sub opt指定时间_Click()
    dtpDate.Enabled = opt指定时间.Value = True
End Sub
