VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmCheckCourseCondition 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "生成盘点表"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "FrmCheckCourseCondition.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox Cbo盘点时间 
      Height          =   300
      Left            =   1185
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2085
      Width           =   2910
   End
   Begin VB.CheckBox chk只针对盘点单中的卫材进行盘点 
      Caption         =   "只针对盘点单中的卫材进行盘点(&Z)"
      Height          =   225
      Left            =   1185
      TabIndex        =   9
      Top             =   2490
      Width           =   3240
   End
   Begin VB.CheckBox chk自动删除汇总后的盘点单 
      Caption         =   "自动删除汇总后的盘点单(&S)"
      Height          =   225
      Left            =   1185
      TabIndex        =   10
      Top             =   2745
      Width           =   2895
   End
   Begin VB.ComboBox cbo库房 
      Height          =   300
      Left            =   1185
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1020
      Width           =   2895
   End
   Begin VB.Frame fra 
      Height          =   105
      Index           =   1
      Left            =   -270
      TabIndex        =   15
      Top             =   3015
      Width           =   6660
   End
   Begin VB.Frame fra 
      Height          =   105
      Index           =   0
      Left            =   -75
      TabIndex        =   14
      Top             =   765
      Width           =   6285
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   120
      TabIndex        =   13
      Top             =   3240
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3945
      TabIndex        =   12
      Top             =   3240
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2790
      TabIndex        =   11
      Top             =   3240
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   300
      Left            =   1185
      TabIndex        =   4
      Top             =   1365
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   114491395
      CurrentDate     =   38552
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   300
      Left            =   1185
      TabIndex        =   6
      Top             =   1695
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   114491395
      CurrentDate     =   38552
   End
   Begin VB.Label lblInfor 
      Caption         =   "根据盘点记录单生成盘存表；以下条件中的“开始时间、结束时间”主要用于过滤出盘点时间。"
      ForeColor       =   &H8000000D&
      Height          =   525
      Left            =   795
      TabIndex        =   0
      Top             =   390
      Width           =   4215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "结束时间(&J)"
      Height          =   180
      Left            =   135
      TabIndex        =   5
      Top             =   1785
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "开始时间(&K)"
      Height          =   180
      Left            =   135
      TabIndex        =   3
      Top             =   1425
      Width           =   990
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "盘点时间(&P)"
      Height          =   180
      Left            =   135
      TabIndex        =   7
      Top             =   2130
      Width           =   990
   End
   Begin VB.Label lbl库房 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "库房(&D)"
      Height          =   180
      Left            =   495
      TabIndex        =   1
      Top             =   1080
      Width           =   630
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   135
      Picture         =   "FrmCheckCourseCondition.frx":000C
      Top             =   195
      Width           =   480
   End
End
Attribute VB_Name = "FrmCheckCourseCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnSelect As Boolean
Private mlng库房id As Long
Private mstr盘点时间 As String
Private mfrmMain As Form
Private mcllNO As Collection
Private mbln盘点单 As Boolean '只针对盘点单中的药品进行盘点
Private mstr盘点单号 As String
Private mbln删除盘点单 As Boolean
Private Const mlngModule = 1719


Private Sub GetCheckCard()
    Dim rsTemp As New ADODB.Recordset
    Dim str盘点时间 As String
    Dim strNo As String
    Dim n As Long
    
    On Error GoTo ErrHandle
    CmdSave.Enabled = False
    gstrSQL = "" & _
        "   Select Distinct 频次 盘点时间,No From 药品收发记录" & _
        "   Where 单据 = 23 And NVL(外观,' ')<>'1' And 库房ID+0=[1] " & _
        "           And 填制日期 Between [2] And [3] " & _
        "   Order by 频次"
        
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[读取盘点时间]", cbo库房.ItemData(cbo库房.ListIndex), dtpStartDate.Value, dtpEndDate.Value)
    
    Cbo盘点时间.Clear
    Set mcllNO = New Collection
    With rsTemp
        Do While Not .EOF
            If Format(CDate(!盘点时间), "yyyy-mm-dd hh:mm:ss") = str盘点时间 Then
                strNo = IIf(strNo = "", "'" & !NO & "'", strNo & "," & "'" & !NO & "'")
                mcllNO.Remove (str盘点时间)
                mcllNO.Add strNo, str盘点时间
            Else
                str盘点时间 = Format(CDate(!盘点时间), "yyyy-mm-dd hh:mm:ss")
                strNo = "'" & !NO & "'"
                mcllNO.Add strNo, str盘点时间
                Cbo盘点时间.AddItem str盘点时间
            End If
            .MoveNext
        Loop
    End With
    
    If Cbo盘点时间.ListCount <> 0 Then
        CmdSave.Enabled = True
        Cbo盘点时间.ListIndex = 0
    Else
        CmdSave.Enabled = False
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub cbo库房_Click()
'    Dim rsTmp As New ADODB.Recordset
'    CmdSave.Enabled = False
'    '装入盘点时间
'    gstrSQL = "" & _
'        "   Select Distinct 频次 盘点时间 From 药品收发记录" & _
'        "   Where 单据 = 23 And 库房ID=[1]" & _
'        "       And 频次 Not In (" & _
'        "               Select Distinct 频次 From 药品收发记录" & _
'        "               Where 单据=22 And 库房ID=[1]" & _
'        "               And 审核人 Is Not Null And Mod(记录状态,3)=1 And 频次 Is Not Null)" & _
'        "    Order by 频次"
'
'    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "-读取盘点时间", cbo库房.ItemData(cbo库房.ListIndex))
'
'    With rsTmp
'        Cbo盘点时间.Clear
'        Do While Not .EOF
'            Cbo盘点时间.AddItem !盘点时间
'            .MoveNext
'        Loop
'    End With
'
'    rsTmp.Close
'    If Cbo盘点时间.ListCount <> 0 Then
'        CmdSave.Enabled = True
'        Cbo盘点时间.ListIndex = 0
'    End If

    Call GetCheckCard
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub CmdSave_Click()
    mlng库房id = cbo库房.ItemData(cbo库房.ListIndex)
    mstr盘点时间 = Cbo盘点时间
    
    mbln盘点单 = (chk只针对盘点单中的卫材进行盘点.Value = 1)
    mbln删除盘点单 = (chk自动删除汇总后的盘点单.Value = 1)
    mstr盘点单号 = mcllNO.Item(Cbo盘点时间.Text)
    
    frmCheckCard.txtStock.Caption = cbo库房.Text
    frmCheckCard.txtStock.Tag = mlng库房id
    frmCheckCard.txtCheckDate = mstr盘点时间
    frmCheckCard.CmdSave.Enabled = False
    frmCheckCard.CmdCancel.Enabled = False
    mblnSelect = True
    Unload Me
End Sub

 

Private Sub dtpEndDate_Change()
    If dtpEndDate.Value < dtpStartDate.Value Then
        dtpStartDate.Value = dtpEndDate.Value
    End If
    
    Call GetCheckCard
End Sub


Private Sub dtpStartDate_Change()
    If dtpEndDate.Value < dtpStartDate.Value Then
        dtpEndDate.Value = Format(dtpStartDate.Value, "yyyy-mm-dd") & " 23:59:59"
    End If
    Call GetCheckCard
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    Dim mblnSelectStock As String, mintLoop As Integer
    
    mblnSelectStock = IIf(Val(zlDatabase.GetPara("是否选择库房", glngSys, mlngModule, "0")) = 1, 1, 0)
    dtpStartDate.CustomFormat = "yyyy-MM-dd HH:mm:ss"
    dtpEndDate.CustomFormat = "yyyy-MM-dd HH:mm:ss"
    dtpEndDate.Value = Format(Sys.Currentdate, "yyyy-mm-dd 23:59:59")
    dtpStartDate.Value = Format(Sys.Currentdate, "yyyy-mm-dd") & " 00:00:00"
    
    
    '装入库房
    With mfrmMain.cboStock
        cbo库房.Clear
        For mintLoop = 0 To .ListCount - 1
            cbo库房.AddItem .List(mintLoop)
            cbo库房.ItemData(cbo库房.NewIndex) = .ItemData(mintLoop)
        Next
        cbo库房.ListIndex = .ListIndex
    End With
    If InStr(1, mfrmMain.mstrPrivs, "所有库房") <> 0 Then
        If mblnSelectStock = "0" Then
            cbo库房.Enabled = False
        Else
            cbo库房.Enabled = True
        End If
    Else
        cbo库房.Enabled = False
    End If
End Sub

Public Function GetCondition(frmMain As Form, _
            ByRef lng库房ID As Long, ByRef str盘点时间 As String, _
            ByRef str盘点单号 As String, ByRef bln只统计盘点单卫材 As Boolean, _
            ByRef bln删除盘点单 As Boolean) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------
    '功能:获取按盘点记录单汇总的相关条件
    '入参:frmMain-父窗口
    '出参:
    '       lng库房ID-库房ID
    '       str盘点时间-盘点时间格式为yyyy-mm-dd hh24:mi:ss
    '       str盘点单号-盘点单据号,以'NO','NO'分隔
    '       bln只统计盘点单物资-只统计盘点中存在的单物资
    '       bln删除盘点单-删除本次所汇总的盘点单中的物资,条件是在str盘点单号中的盘点记录单
    '返回:按了确定为true,否则为False
    '------------------------------------------------------------------------------------------------------------------------------
    mblnSelect = False
    Set mfrmMain = frmMain
    Me.Show vbModal, frmMain
    GetCondition = mblnSelect
    If mblnSelect = False Then Exit Function
    lng库房ID = mlng库房id
    str盘点时间 = mstr盘点时间
    str盘点单号 = mstr盘点单号
    
    bln只统计盘点单卫材 = mbln盘点单
    bln删除盘点单 = mbln删除盘点单
End Function

 
