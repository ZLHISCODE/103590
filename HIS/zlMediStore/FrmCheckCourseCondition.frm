VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCheckCourseCondition 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "生成盘点表"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4035
   Icon            =   "FrmCheckCourseCondition.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2730
      TabIndex        =   7
      Top             =   2775
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1560
      TabIndex        =   6
      Top             =   2775
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   90
      TabIndex        =   8
      Top             =   2775
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "条件"
      Height          =   2535
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   3765
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   300
         Left            =   960
         TabIndex        =   12
         Top             =   735
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   159514627
         CurrentDate     =   38552
      End
      Begin VB.CheckBox chk自动删除汇总后的盘点单 
         Caption         =   "自动删除汇总后的盘点单"
         Height          =   225
         Left            =   690
         TabIndex        =   9
         Top             =   2205
         Width           =   2895
      End
      Begin VB.CheckBox chk只针对盘点单中的药品进行盘点 
         Caption         =   "只针对盘点单中的药品进行盘点"
         Height          =   225
         Left            =   690
         TabIndex        =   5
         Top             =   1950
         Width           =   2895
      End
      Begin VB.ComboBox Cbo盘点时间 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1545
         Width           =   2475
      End
      Begin VB.ComboBox cbo库房 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   330
         Width           =   2475
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   300
         Left            =   960
         TabIndex        =   13
         Top             =   1140
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   159514627
         CurrentDate     =   38552
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "开始时间"
         Height          =   180
         Left            =   180
         TabIndex        =   11
         Top             =   795
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "结束时间"
         Height          =   180
         Left            =   180
         TabIndex        =   10
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label lbl库房 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "库房"
         Height          =   180
         Left            =   540
         TabIndex        =   1
         Top             =   390
         Width           =   360
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "盘点时间"
         Height          =   180
         Left            =   165
         TabIndex        =   3
         Top             =   1605
         Width           =   720
      End
   End
End
Attribute VB_Name = "FrmCheckCourseCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnSelect As Boolean
Private mlng库房ID As Long
Private mstr盘点时间 As String
Private mbln盘点单 As Boolean '只针对盘点单中的药品进行盘点
Private mfrmMain As Form
Private mstr盘点单号 As String
Private mbln删除盘点单 As Boolean
Private mcolCheckCourseCard As Collection     '记录每个盘点时间对应的盘点单号
Private Sub GetCheckCard()
    Dim rsTmp As New ADODB.Recordset
    Dim str盘点时间 As String
    Dim strNo As String
    Dim n As Long
    
    On Error GoTo errHandle
    CmdSave.Enabled = False
    '装入盘点时间
    gstrSQL = "Select Distinct 频次 盘点时间,No From 药品收发记录" & _
              " Where 单据 = 14 And NVL(外观,' ')<>'1' And 库房ID+0=[1] " & _
              " And 填制日期 Between [2] And [3] " & _
              " Order by 频次"
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[读取盘点时间]", cbo库房.ItemData(cbo库房.ListIndex), dtpStartDate.Value, dtpEndDate.Value)
    
    Set mcolCheckCourseCard = New Collection
    
    Cbo盘点时间.Clear
    With rsTmp
        Do While Not .EOF
            If Format(CDate(!盘点时间), "yyyy-mm-dd hh:mm:ss") = str盘点时间 Then
                strNo = IIf(strNo = "", "'" & !NO & "'", strNo & "," & "'" & !NO & "'")
                mcolCheckCourseCard.Remove (str盘点时间)
                mcolCheckCourseCard.Add strNo, str盘点时间
            Else
                str盘点时间 = Format(CDate(!盘点时间), "yyyy-mm-dd hh:mm:ss")
                strNo = "'" & !NO & "'"
                mcolCheckCourseCard.Add strNo, str盘点时间
                Cbo盘点时间.AddItem str盘点时间
            End If
            .MoveNext
        Loop
    End With
    
    rsTmp.Close
    If Cbo盘点时间.ListCount <> 0 Then
        CmdSave.Enabled = True
        Cbo盘点时间.ListIndex = 0
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo库房_Click()
    Call GetCheckCard
End Sub






Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    mlng库房ID = cbo库房.ItemData(cbo库房.ListIndex)
    mstr盘点时间 = Cbo盘点时间
    mbln盘点单 = (chk只针对盘点单中的药品进行盘点.Value = 1)
    mbln删除盘点单 = (chk自动删除汇总后的盘点单.Value = 1)
    mstr盘点单号 = mcolCheckCourseCard.Item(Cbo盘点时间.Text)
    frmNewCheckCard.txtStock.Caption = cbo库房.Text
    frmNewCheckCard.txtStock.Tag = mlng库房ID
    frmNewCheckCard.txtCheckDate = mstr盘点时间
'    frmCheckCard.CmdSave.Enabled = False
'    frmCheckCard.CmdCancel.Enabled = False
    
    mblnSelect = True
    Unload Me
End Sub

Private Sub dtpEndDate_Change()
    Call GetCheckCard
End Sub


Private Sub dtpStartDate_Change()
    Call GetCheckCard
End Sub


Private Sub Form_Load()
    Dim mblnSelectStock As String, mintLoop As Integer
    
    mblnSelectStock = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\药品盘点管理", "库房", "0")
    
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

Public Function GetCondition(FrmMain As Form, ByRef lng库房ID As Long, ByRef str盘点单号 As String, ByRef bln盘点单 As Boolean, ByRef bln删除盘点单) As Boolean
    mblnSelect = False
    Set mfrmMain = FrmMain
    Me.Show vbModal, FrmMain
    GetCondition = mblnSelect
    
    If mblnSelect = False Then Exit Function
    lng库房ID = mlng库房ID
    bln盘点单 = mbln盘点单
    str盘点单号 = mstr盘点单号
    bln删除盘点单 = mbln删除盘点单
End Function
