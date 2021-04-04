VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSquareCardFilter 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picFilter 
      BorderStyle     =   0  'None
      Height          =   4590
      Left            =   60
      ScaleHeight     =   4590
      ScaleWidth      =   3885
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   3885
      Begin VB.CheckBox chk含停用 
         Caption         =   "包含停用的卡"
         Height          =   180
         Left            =   645
         TabIndex        =   18
         Top             =   3810
         Width           =   1425
      End
      Begin VB.CommandButton cmd刷新 
         Caption         =   "过滤(&F)"
         Height          =   390
         Left            =   2700
         TabIndex        =   20
         Top             =   4185
         Width           =   1050
      End
      Begin VB.TextBox txtEdit 
         Height          =   330
         Index           =   2
         Left            =   630
         TabIndex        =   17
         Top             =   3285
         Width           =   3105
      End
      Begin VB.TextBox txtEdit 
         Height          =   315
         Index           =   1
         Left            =   2430
         TabIndex        =   8
         Top             =   615
         Width           =   1290
      End
      Begin VB.TextBox txtEdit 
         Height          =   315
         Index           =   0
         Left            =   630
         TabIndex        =   7
         Top             =   630
         Width           =   1290
      End
      Begin VB.ComboBox cbo类型 
         Height          =   300
         Left            =   630
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   135
         Width           =   3090
      End
      Begin VB.CheckBox chkDate 
         Caption         =   "按回收日期过滤"
         Height          =   180
         Index           =   1
         Left            =   45
         TabIndex        =   13
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CheckBox chkDate 
         Caption         =   "按发卡日期过滤"
         Height          =   180
         Index           =   0
         Left            =   45
         TabIndex        =   10
         Top             =   1245
         Value           =   1  'Checked
         Width           =   1605
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   315
         Index           =   0
         Left            =   630
         TabIndex        =   11
         Top             =   1500
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   183828483
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   315
         Index           =   0
         Left            =   2430
         TabIndex        =   12
         Top             =   1500
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   183828483
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   315
         Index           =   1
         Left            =   630
         TabIndex        =   14
         Top             =   2310
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   183828483
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   315
         Index           =   1
         Left            =   2430
         TabIndex        =   15
         Top             =   2310
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   183828483
         CurrentDate     =   37007
      End
      Begin VB.TextBox txtEdit 
         Height          =   330
         Index           =   3
         Left            =   630
         TabIndex        =   16
         Top             =   2865
         Width           =   3105
      End
      Begin VB.Label lblEDIT 
         AutoSize        =   -1  'True
         Caption         =   "领卡人"
         Height          =   180
         Index           =   3
         Left            =   15
         TabIndex        =   19
         Top             =   3360
         Width           =   540
      End
      Begin VB.Label lbl至 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Index           =   0
         Left            =   2025
         TabIndex        =   9
         Top             =   690
         Width           =   180
      End
      Begin VB.Label lblEDIT 
         AutoSize        =   -1  'True
         Caption         =   "卡号"
         Height          =   180
         Index           =   2
         Left            =   195
         TabIndex        =   6
         Top             =   705
         Width           =   360
      End
      Begin VB.Label lblEDIT 
         AutoSize        =   -1  'True
         Caption         =   "卡类型"
         Height          =   180
         Index           =   0
         Left            =   15
         TabIndex        =   4
         Top             =   180
         Width           =   540
      End
      Begin VB.Label lbl至 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Index           =   3
         Left            =   2040
         TabIndex        =   3
         Top             =   1545
         Width           =   180
      End
      Begin VB.Label lbl至 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Index           =   4
         Left            =   2040
         TabIndex        =   2
         Top             =   2370
         Width           =   180
      End
      Begin VB.Label lblEDIT 
         AutoSize        =   -1  'True
         Caption         =   "发卡人"
         Height          =   180
         Index           =   1
         Left            =   15
         TabIndex        =   1
         Top             =   2940
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmSquareCardFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Option Explicit
Private mArrFilter As Variant
Private mstrPrivs As String, mlngModule As Long
Private Enum mtxtIdx
    idx_开始卡号 = 0
    idx_结束卡号 = 1
    idx_领卡人 = 2
    idx_发卡人 = 3
End Enum
'--------------------------------------------------------------------------------------------------------
Public Event zlRefreshCon(ByVal arrFilter As Variant)

Public Sub Init条件(ByVal lngModul As Long, ByVal strPrivs As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关条件
    '编制:刘兴洪
    '日期:2009-11-18 14:48:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mstrPrivs = strPrivs: mlngModule = lngModul
    Call InitData
End Sub

Public Property Get GetFilterCon() As Variant
    Call GetFilter
    Set GetFilterCon = mArrFilter
End Property

Public Sub ReActionFilter()
    '重新缴活过滤
     cmd刷新_Click
End Sub

Private Function GetFilter() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取条件信息
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-11-18 14:27:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllFilter As Collection, strReg As String
    
    '基本查询条件
    Set cllFilter = New Collection
    If cbo类型.ListIndex = 0 Then
        cllFilter.Add "所有", "卡类型"
    Else
        cllFilter.Add zlstr.NeedName(cbo类型.Text), "卡类型"
    End If
    cllFilter.Add Array(Trim(txtEdit(mtxtIdx.idx_开始卡号).Text), Trim(txtEdit(mtxtIdx.idx_结束卡号).Text)), "卡号范围"
    
    cllFilter.Add Trim(txtEdit(mtxtIdx.idx_领卡人).Text), "领卡人"
    cllFilter.Add Trim(txtEdit(mtxtIdx.idx_发卡人).Text), "发卡人"
    
    If chkDate(0).value = 1 Then
        cllFilter.Add Array(Format(dtpStartDate(0).value, "yyyy-mm-dd") & " 00:00:00", Format(dtpEndDate(0).value, "yyyy-mm-dd") & " 23:59:59"), "发卡时间"
    Else
        cllFilter.Add Array("1901-01-01", "1901-01-01"), "发卡时间"
    End If
    If chkDate(1).value = 1 Then
        cllFilter.Add Array(Format(dtpStartDate(1).value, "yyyy-mm-dd") & " 00:00:00", Format(dtpEndDate(1).value, "yyyy-mm-dd") & " 23:59:59"), "回收时间"
    Else
        cllFilter.Add Array("1901-01-01", "1901-01-01"), "回收时间"
    End If
    cllFilter.Add IIf(chk含停用.value = 1, 1, 0), "包含停用卡"
    Set mArrFilter = cllFilter
End Function

Private Sub cmd刷新_Click()
    If chkDate(0).value = 0 And chkDate(1).value = 0 Then
        ShowMsgbox "必需确定一个时间范围，请检查！"
        Exit Sub
    End If
    Call GetFilter
    RaiseEvent zlRefreshCon(mArrFilter)
End Sub

Private Sub InitData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '编制:刘兴洪
    '日期:2009-11-18 14:37:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    
    dtpEndDate(0).MaxDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    dtpEndDate(1).MaxDate = dtpEndDate(0).MaxDate

    dtpEndDate(0).value = dtpEndDate(0).MaxDate
    dtpEndDate(1).value = dtpEndDate(0).MaxDate
    
    dtpStartDate(0).value = Format(DateAdd("d", -1, zlDatabase.Currentdate), "yyyy-mm-dd")  '缺省按7天找
    dtpStartDate(1).value = dtpStartDate(0).value
    
    On Error GoTo errHandle
    '加载卡类型:如果没有按缺省标志,以所有卡为准,否则按缺省为准
    strSql = "Select 编码, 名称, 缺省面额, 缺省折扣, 缺省标志 From 消费卡类型"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    With cbo类型
        .Clear
        .AddItem "所有卡"
        .ListIndex = .NewIndex
        Do While Not rsTemp.EOF
            .AddItem Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称)
            If Val(Nvl(rsTemp!缺省标志)) = 1 Then .ListIndex = .NewIndex
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
 
Private Sub chkDate_Click(Index As Integer)
    Select Case Index
    Case 0
        If chkDate(Index).value = 0 Then
           If chkDate(1).value = 0 Then chkDate(1).value = 1
        End If
    Case 1
        If chkDate(Index).value = 0 Then
           If chkDate(0).value = 0 Then chkDate(0).value = 1
        End If
    End Select
    dtpStartDate(Index).Enabled = chkDate(Index).value = 1
    dtpEndDate(Index).Enabled = chkDate(Index).value = 1
End Sub

Private Sub chkDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub dtpEndDate_Change(Index As Integer)
     If dtpEndDate(Index).value > dtpStartDate(Index).MaxDate Then dtpEndDate(Index).value = dtpStartDate(Index).MaxDate
    If dtpEndDate(Index).value < dtpStartDate(Index).value Then
        dtpStartDate(Index).value = dtpEndDate(Index).value
    End If
End Sub

Private Sub dtpEndDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub dtpStartDate_Change(Index As Integer)
    If dtpStartDate(Index).value > dtpEndDate(Index).MaxDate Then dtpStartDate(Index).value = dtpEndDate(Index).MaxDate
    If dtpEndDate(Index).value < dtpStartDate(Index).value Then
        dtpEndDate(Index).value = dtpStartDate(Index).value
    End If
End Sub

Private Sub dtpStartDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    End If
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With picFilter
        .Left = ScaleLeft
        .Top = ScaleTop
        .Height = ScaleHeight
    End With
End Sub

Private Sub picFilter_Resize()
    Err = 0: On Error Resume Next
    With picFilter
         cmd刷新.Left = .ScaleLeft + .ScaleWidth - cmd刷新.Width - 50
    End With
End Sub

Private Sub txtEdit_Change(Index As Integer)
    txtEdit(Index).Tag = ""
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txtEdit(Index).Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If txtEdit(Index).Text = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    Select Case Index
    Case mtxtIdx.idx_发卡人
        If Select人员选择器(Me, txtEdit(Index), Trim(txtEdit(Index).Text)) = False Then
            Exit Sub
        End If
    Case mtxtIdx.idx_领卡人
        If Select人员选择器(Me, txtEdit(Index), Trim(txtEdit(Index).Text)) = False Then
            Exit Sub
        End If
    Case Else
        '由于卡号不知长度,所以无法补位
    End Select
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
    Case mtxtIdx.idx_开始卡号, mtxtIdx.idx_结束卡号
        '小写字母转换为大写
        If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
            KeyAscii = KeyAscii - Asc("a") + Asc("A")
        End If
    End Select
End Sub
