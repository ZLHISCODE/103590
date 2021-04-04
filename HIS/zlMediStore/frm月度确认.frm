VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm月度确认 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "月度确认"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   Icon            =   "frm月度确认.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd清除 
      Caption         =   "清除(&L)"
      Height          =   350
      Left            =   150
      TabIndex        =   9
      ToolTipText     =   "清除上一个月的确认记录"
      Top             =   2010
      Width           =   945
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3210
      TabIndex        =   8
      Top             =   2010
      Width           =   945
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2190
      TabIndex        =   7
      Top             =   2010
      Width           =   945
   End
   Begin VB.Frame Frame1 
      Caption         =   "确认本月度的终止日期(&A)"
      Height          =   1755
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   4155
      Begin VB.ComboBox cbo当前月份 
         Height          =   300
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1605
      End
      Begin MSComCtl2.DTPicker dtp开始日期 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   780
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
         Format          =   114491395
         CurrentDate     =   38148
      End
      Begin MSComCtl2.DTPicker dtp结束日期 
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         Top             =   1230
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
         Format          =   114491395
         CurrentDate     =   38148
      End
      Begin VB.Label lbl月份 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "月份(&M)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   720
         TabIndex        =   1
         Top             =   420
         Width           =   630
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "结束日期(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   5
         Top             =   1297
         Width           =   990
      End
      Begin VB.Label lbl开始日期 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "开始日期(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   3
         Top             =   847
         Width           =   990
      End
   End
End
Attribute VB_Name = "frm月度确认"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstr缺省结束日期 As String
Private mlng库房ID As Long
Private mblnStart As Boolean

Private Sub cbo当前月份_Click()
    Dim str开始日期 As String, str结束日期 As String
    
    Call GetPeriod(Me.cbo当前月份.Text, str开始日期, str结束日期)
    '设置结束日期和开始日期
    '结束日期的最小时间为开始时间加1天
    Me.Dtp结束日期.Value = Format(str结束日期, "yyyy年MM月dd日 HH:mm:ss")
    Me.Dtp开始日期.Value = Format(str开始日期, "yyyy年MM月dd日 HH:mm:ss")
    Me.Dtp开始日期.Enabled = (Me.cbo当前月份.ListCount > 2)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    On Error GoTo ErrHand
    
    '开始日期必须小于结束日期
    If Not (Format(Me.Dtp结束日期.Value, "yyyy-MM-dd HH:mm:ss") > Format(Me.Dtp开始日期.Value, "yyyy-MM-dd HH:mm:ss")) Then
        MsgBox "结束日期必须大于开始时间！", vbInformation, gstrSysName
        Me.Dtp结束日期.SetFocus
        Exit Sub
    End If
    If Format(Me.Dtp结束日期.Value, "yyyy-MM-dd HH:mm:ss") > Format(Sys.Currentdate, "yyyy-MM-dd HH:mm:ss") Then
        MsgBox "结束日期不能大于当前日期！", vbInformation, gstrSysName
        Me.Dtp结束日期.SetFocus
        Exit Sub
    End If
    
    gstrSQL = "zl_库房确认记录_UPDATE(" & mlng库房ID & ",'" & Me.cbo当前月份.Text & "'" & _
        ",to_date('" & Format(Me.Dtp开始日期.Value, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss')" & _
        ",to_date('" & Format(Me.Dtp结束日期.Value, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss')" & _
        ")"
    Call zlDataBase.ExecuteProcedure(gstrSQL, "保存库房确认记录")
    
    Unload Me
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmd清除_Click()
    On Error GoTo ErrHand
    
    If MsgBox("你确定要清除最后一个月的确认记录吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    gstrSQL = "zl_库房确认记录_Back(" & mlng库房ID & ")"
    Call zlDataBase.ExecuteProcedure(gstrSQL, "清除最后一次的确认记录")
    Unload Me
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Activate()
    If Not mblnStart Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim str月份 As String
    Dim blnInit As Boolean
    Dim rsTemp As New ADODB.Recordset
    '第一次可以由操作员设置开始日期、选择月份，缺省的开始日期和结束日期为期间表的开始日期和结束日期
    '以后只能设置
    '   1、结束日期，结束日期缺省为当天，最大选择到当天，可以向前选择，但不能小于等于开始日期
    '   2、月份不能选择，只能是大于最后一个月份
    On Error GoTo errHandle
    mblnStart = False
    '缺省装入上月、本月两个月份
    gstrSQL = "" & _
        " Select MAX(月份) 月份" & _
        " From 库房确认记录" & _
        " WHERE 库房ID=[1] And 性质=1"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[取最大月份]", mlng库房ID)
        
    If Not IsNull(rsTemp!月份) Then
        blnInit = False
        str月份 = rsTemp!月份   '最大月份=上月
    Else
        blnInit = True
    End If
    
    '装入期间表（第一次则装入所有期间，否则装入上月及以后期间）
    gstrSQL = "Select 期间 As 月份,开始日期,终止日期 " & _
            " From 期间表" & _
            " Where " & IIf(blnInit, "1=1", " 期间>=[1] And Rownum<3") & _
            "" & _
            " Order by 期间"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[读取期间表]", str月份)
    
    Me.cbo当前月份.Clear
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
    Else
        Exit Sub
    End If
    
    Do While Not rsTemp.EOF
        Me.cbo当前月份.AddItem rsTemp!月份
        rsTemp.MoveNext
    Loop
    
    '读取上次结束日期
    Call LocateCbo(str月份)
    
    mblnStart = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ShowEditor(ByVal lng库房ID As Long, Optional ByVal str结束日期 As String = "")
    mlng库房ID = lng库房ID
    mstr缺省结束日期 = str结束日期
    Me.Show 1
End Sub

Private Function LocateCbo(ByVal strInput As String) As Boolean
    Dim intItem As Integer, intItems As Integer
    '定位到当前月份
    LocateCbo = True
    Me.cbo当前月份.ListIndex = 0
    If strInput = "" Then Exit Function
    intItems = Me.cbo当前月份.ListCount - 1
    For intItem = 0 To intItems
        If Me.cbo当前月份.Text = strInput Then
            Me.cbo当前月份.ListIndex = intItem
            Exit Function
        End If
    Next
    LocateCbo = False
End Function

Private Sub GetPeriod(ByVal str月份 As String, str开始日期 As String, str结束日期 As String)
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    '读取指定月份的开始日期和终止日期
    gstrSQL = "Select to_char(开始时间,'yyyy-MM-DD hh24:mi:ss') As 开始日期,to_char(终止时间,'yyyy-MM-DD hh24:mi:ss') As 结束日期" & _
            " From 库房确认记录" & _
            " Where 库房ID=[1] And 月份=[2] And 性质=1"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[读取指定月份的开始日期和终止日期]", mlng库房ID, str月份)
    
    If Not rsTemp.EOF Then
        str开始日期 = rsTemp!开始日期
        str结束日期 = rsTemp!结束日期
    Else
        '试着去取上个月的终止日期做为本次开始日期，如果还是取不到，说明是第一次运行
        Call GetStartDate(str月份, str开始日期)
        If mstr缺省结束日期 <> "" And mstr缺省结束日期 > str开始日期 Then
            str结束日期 = mstr缺省结束日期
        Else
            str结束日期 = Format(Sys.Currentdate(), "yyyy-MM-dd HH:mm:ss")
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetStartDate(ByVal str月份 As String, str开始日期 As String)
    Dim rsTemp As New ADODB.Recordset
    '根据上个月的终止日期得到本月开始日期
    On Error GoTo errHandle
    gstrSQL = "Select to_char(max(终止时间),'yyyy-MM-DD hh24:mi:ss') As 开始日期" & _
            " From 库房确认记录" & _
            " Where 库房ID=[1] And 月份<[2] And 性质=1"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[根据上个月的终止日期得到本月开始日期]", mlng库房ID, str月份)
    
    If Not IsNull(rsTemp!开始日期) Then
        str开始日期 = DateAdd("s", 1, rsTemp!开始日期)
    Else
        str开始日期 = "2004-01-01 00:00:00"
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
