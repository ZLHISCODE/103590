VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTendRecalCulation 
   Caption         =   "记录单重算"
   ClientHeight    =   5115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5220
   Icon            =   "frmTendRecalCulation.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   5220
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picImg 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   2010
      Picture         =   "frmTendRecalCulation.frx":6852
      ScaleHeight     =   240
      ScaleWidth      =   270
      TabIndex        =   15
      Top             =   3795
      Width           =   270
   End
   Begin VB.CheckBox chkPrint 
      Caption         =   "是否重算已打印文件"
      Height          =   180
      Left            =   255
      TabIndex        =   14
      Top             =   3480
      Width           =   1920
   End
   Begin VB.PictureBox picImg 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   2235
      Picture         =   "frmTendRecalCulation.frx":6DDC
      ScaleHeight     =   240
      ScaleWidth      =   270
      TabIndex        =   13
      Top             =   3465
      Width           =   270
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   300
      Left            =   3465
      TabIndex        =   10
      Top             =   4110
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   168886273
      CurrentDate     =   43335
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   300
      Left            =   1080
      TabIndex        =   9
      Top             =   4110
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   168886273
      CurrentDate     =   43335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   300
      Left            =   3900
      TabIndex        =   8
      Top             =   4590
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   300
      Left            =   2715
      TabIndex        =   7
      Top             =   4575
      Width           =   900
   End
   Begin VB.CheckBox chkOut 
      Caption         =   "出院病人时间范围"
      Height          =   300
      Left            =   255
      TabIndex        =   6
      Top             =   3765
      Width           =   1785
   End
   Begin VB.ListBox lstMain 
      Appearance      =   0  'Flat
      Height          =   1710
      Left            =   255
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   1620
      Width           =   4530
   End
   Begin VB.OptionButton opt 
      Caption         =   "全院重算"
      Height          =   495
      Index           =   2
      Left            =   3750
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.OptionButton opt 
      Caption         =   "按科室重算"
      Height          =   375
      Index           =   1
      Left            =   1995
      TabIndex        =   3
      Top             =   1140
      Width           =   1335
   End
   Begin VB.OptionButton opt 
      Caption         =   "按病区重算"
      Height          =   255
      Index           =   0
      Left            =   255
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblBeginTime 
      AutoSize        =   -1  'True
      Caption         =   "开始时间："
      Height          =   180
      Left            =   255
      TabIndex        =   12
      Top             =   4170
      Width           =   900
   End
   Begin VB.Label lblEndTime 
      AutoSize        =   -1  'True
      Caption         =   "～结束时间："
      Height          =   180
      Left            =   2445
      TabIndex        =   11
      Top             =   4170
      Width           =   1080
   End
   Begin VB.Label lblWay 
      Caption         =   "重算方式："
      Height          =   255
      Left            =   255
      TabIndex        =   1
      Top             =   765
      Width           =   1095
   End
   Begin VB.Label lblTitle 
      Caption         =   "    根据不同病区，不同科室的需求，指定本次重算的范围。勾选重算可能会占用你几分钟时间！"
      Height          =   585
      Left            =   255
      TabIndex        =   0
      Top             =   180
      Width           =   4500
   End
End
Attribute VB_Name = "frmTendRecalCulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrspart As New ADODB.Recordset
Private mrsRoom As New ADODB.Recordset
Private mstr科室ID As String

Private Sub chkOut_Click()
    dtpBegin.Enabled = chkOut.Value
    dtpEnd.Enabled = chkOut.Value
End Sub

Private Sub cmdCancel_Click()
    mstr科室ID = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    Dim strID As String
    Dim strTime As String
    If lstMain.ListCount > 0 Then
        For i = 0 To lstMain.ListCount - 1
            If lstMain.Selected(i) Then
                strID = strID & "," & lstMain.ItemData(i)
            End If
        Next
    End If
    If strID <> "" Then strID = Mid(strID, 2)
    mstr科室ID = strID
    
    mstr科室ID = mstr科室ID & "|" & chkPrint.Value
    
    If chkOut.Value = Checked Then
        strTime = dtpBegin.Value & "'" & dtpEnd.Value
        mstr科室ID = mstr科室ID & "|" & strTime
    End If
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrspart = Nothing
    Set mrsRoom = Nothing
End Sub

Private Sub opt_Click(Index As Integer)
    Dim i As Integer
    Select Case Index
        Case 0
            lstMain.Clear
            If mrspart.RecordCount > 0 Then mrspart.MoveFirst
            i = 0
            With mrspart
                Do While Not .EOF
                    Me.lstMain.AddItem !名称, i
                    Me.lstMain.ItemData(i) = Val(!ID & "")
                    i = i + 1
                    .MoveNext
                Loop
            End With
        Case 1
            lstMain.Clear
            If mrsRoom.RecordCount > 0 Then mrsRoom.MoveFirst
            i = 0
            With mrsRoom
                Do While Not .EOF
                    Me.lstMain.AddItem !名称, i
                    Me.lstMain.ItemData(i) = Val(!ID & "")
                    i = i + 1
                    .MoveNext
                Loop
            End With
        Case Else
            lstMain.Clear
            Me.lstMain.AddItem "全院"
            Me.lstMain.ItemData(i) = -1
    End Select
End Sub

 
Public Function ShowEditor(ByVal FileId As Long) As String
    Dim strSQL As String
    
    On Error GoTo errHand
    
    strSQL = "Select Distinct b.名称, b.Id From 病人护理文件 A, 部门表 B Where a.科室id = b.Id And 格式id = [1]"
    Set mrspart = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, FileId)
    
    strSQL = " Select Distinct b.名称,b.id " & vbNewLine & _
        " From 病人护理文件 A, 部门表 B, 病区科室对应 C " & vbNewLine & _
        " Where a.科室id = c.科室id And c.病区id = b.Id And 格式id =[1]"
    Set mrsRoom = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, FileId)
    dtpBegin.Value = zlDatabase.Currentdate - 3
    dtpEnd.Value = zlDatabase.Currentdate
    Me.Show 1
    ShowEditor = mstr科室ID
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub picImg_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    If Index = 0 Then
        strInfo = "是否将已打印的记录单数据清空，勾选将清空打印记录，需重新打印。"
    ElseIf Index = 1 Then
        strInfo = "出院病人时间范围，重算这个时间段范围出院的病人。"
    End If
    Call zlCommFun.ShowTipInfo(picImg(Index).hWnd, strInfo, True)
End Sub
