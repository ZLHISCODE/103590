VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHistSql 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "历史Sql语句"
   ClientHeight    =   1320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7905
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1320
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox chkSess 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "只查看当前会话"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6240
      TabIndex        =   6
      Top             =   165
      Width           =   1575
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "查找(&F)"
      Height          =   345
      Left            =   5040
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dtpStart 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   135
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy/MM/dd HH:mm"
      Format          =   215220227
      CurrentDate     =   42961
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   315
      Left            =   3120
      TabIndex        =   2
      Top             =   135
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy/MM/dd HH:mm"
      Format          =   215220227
      CurrentDate     =   42961
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfSQL 
      Height          =   540
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   4995
      _cx             =   8811
      _cy             =   952
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   15921906
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   8421504
      GridColorFixed  =   8421504
      TreeColor       =   0
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   280
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "时间范围"
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   202
      Width           =   720
   End
   Begin VB.Label lblTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "至"
      Height          =   180
      Left            =   2880
      TabIndex        =   4
      Top             =   202
      Width           =   180
   End
End
Attribute VB_Name = "frmHistSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngSid As Long
Private mlngSerial As Long
Private mstrUser As String  '通过是否传入用户名判断是弹出窗体还是嵌入窗体,弹出窗体无须传入用户名

Public Sub SetUser(ByVal strUser As String)
    '此方法用于查询用户的历史SQL时传入用户名
    mstrUser = strUser
End Sub

Public Sub SetSid(ByVal lngSid As Long, ByVal lngSerial As Long)
    '此方法用户有阻塞会话调用,要求传入 sid+Serial#
    mlngSid = lngSid
    mlngSerial = lngSerial
End Sub

Public Sub ShowMe()
    Me.Show
End Sub

Private Sub cmdFind_Click()
    LoadSQL
End Sub

Private Sub Form_Load()
    Dim strCol As String
    
    dtpStart.value = date
    dtpEnd.value = date + 1
    
    strCol = "  ,500,1;Sid,1200,1;用户,1005,1;机器名,1980,1;程序名称,1365,1;执行时间,1650,1;等待时间,1005,1;等待事件,1005,1;" & _
                "阻塞会话,1200,1;SQL_ID,800,1;SQL文本,1500,1"
    
    InitTable vsfSQL, strCol
    vsfSQL.Rows = 1
    vsfSQL.TextMatrix(0, 1) = "Sid,Serial#"
    
    If mlngSid <> 0 Then
        LoadSQL
    End If
    
    chkSess.Visible = Not mstrUser = ""
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    vsfSQL.Width = Me.ScaleWidth - vsfSQL.Left - 120
    vsfSQL.Height = Me.ScaleHeight - vsfSQL.Top - 60
End Sub


Private Sub LoadSQL()
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim i As Long, strSID As String
    Dim dtStart As Date, dtEnd As Date
    
    On Error GoTo errH
    
    dtStart = CDate(Format(dtpStart.value, "yyyy-MM-dd hh:mm:ss"))
    dtEnd = CDate(Format(dtpEnd.value, "yyyy-MM-dd hh:mm:ss"))
    strSQL = "Select b.Session_Id || ',' || b.Session_Serial# As ""Sid"", b.Machine, c.Username, b.Wait_Time," & vbNewLine & _
                "       decode(b.Blocking_Session || ',' ||b.Blocking_Session_Serial#,',','',b.Blocking_Session || ',' ||b.Blocking_Session_Serial#) As ""Blocking_Session"", " & vbNewLine & _
                "b.Program, d.Sql_Id, d.Sql_Text," & vbNewLine & _
                "       To_Char(b.sql_exec_start, 'yyyy/mm/dd hh24:mi') sql_exec_start,Event" & vbNewLine & _
                "From Dba_Hist_Snapshot A, Dba_Hist_Active_Sess_History B, Dba_Users C, Dba_Hist_Sqltext D" & vbNewLine & _
                "Where a.Snap_Id = b.Snap_Id And a.Dbid = b.Dbid  And b.User_Id = c.User_Id" & vbNewLine & _
                "And b.Sql_Id(+) = d.Sql_Id And b.Sql_Id Is Not Null" & vbNewLine & _
                "And a.Begin_Interval_Time between [1] and [2] " & vbNewLine & _
                IIf(mstrUser <> "", "And c.Username = [3] ", " And b.Session_Id =[4] And  b.Session_Serial# =[5]") & vbNewLine & _
                IIf(chkSess.value = 1, " And b.Session_Id =[4] And  b.Session_Serial# =[5]", "") & vbNewLine & _
                "Order By b.Session_Id, b.Session_Serial#, b.Sample_Time Desc"
                
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取历史Sql", dtStart, dtEnd, mstrUser, mlngSid, mlngSerial)
    
    If rsTmp.RecordCount = 0 Then
        vsfSQL.Rows = 1
        Exit Sub
    End If
    
    With vsfSQL
        .Redraw = flexRDNone
        .MergeCol(.ColIndex("Sid-Serial#")) = True
        .MergeCells = flexMergeRestrictAll
        
        .OutlineBar = flexOutlineBarSimpleLeaf
        .SubtotalPosition = flexSTAbove
        .Rows = 1: i = 1
        .Rows = rsTmp.RecordCount + 1
        .ComboList = "..."
        'Sid",Serial#,1200,1;用户,1000,1;机器名,1000,1;程序名称,1000,1;记录时间,1500,1;等待时间,1000,1; 阻塞会话,1200,1;SQL_ID,800,1;SQL文本,2000,1
        Do While Not rsTmp.EOF
            .IsSubtotal(i) = True
            If strSID = rsTmp!Sid & "" Then
                .RowOutlineLevel(i) = 2
            Else
                strSID = rsTmp!Sid & ""
                .RowOutlineLevel(i) = 1
            End If
            .TextMatrix(i, .ColIndex("Sid")) = rsTmp!Sid & ""
            .TextMatrix(i, .ColIndex("用户")) = rsTmp!USERNAME & ""
            .TextMatrix(i, .ColIndex("机器名")) = rsTmp!Machine & ""
            .TextMatrix(i, .ColIndex("程序名称")) = rsTmp!Program & ""
            .TextMatrix(i, .ColIndex("执行时间")) = rsTmp!sql_exec_start & ""
            .TextMatrix(i, .ColIndex("等待时间")) = rsTmp!Wait_Time & ""
            .TextMatrix(i, .ColIndex("等待事件")) = rsTmp!Event & ""
            .TextMatrix(i, .ColIndex("阻塞会话")) = rsTmp!Blocking_Session & ""
            .TextMatrix(i, .ColIndex("SQL_ID")) = rsTmp!Sql_Id & ""
            .TextMatrix(i, .ColIndex("SQL文本")) = rsTmp!sql_text & ""
            i = i + 1
            rsTmp.MoveNext
        Loop
        
        .Redraw = flexRDDirect
    End With
    
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
End Sub

Private Sub FindWindowAndSetActive(ByVal FrmObj As Form)
    Dim LngTargetHdl As Long
    '--如果该窗体已经打开,则激活它(这样,窗体的大小不会发生变化)--zyb
    LngTargetHdl = FindWindow(vbNullString, FrmObj.Caption)
    If LngTargetHdl <> 0 Then
        If IsIconic(LngTargetHdl) Then
            Call ShowWindow(LngTargetHdl, 9)            '还原指定窗体为原大小
        End If
        Call SetActiveWindow(LngTargetHdl)
    End If
End Sub

Private Sub vsfSQL_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfSQL
        If .TextMatrix(Row, Col) <> "" And (Col = .ColIndex("SQL文本") Or (Col = .ColIndex("阻塞会话") And mstrUser <> "")) Then
            Cancel = False
        Else
            Cancel = True
        End If
    End With
End Sub

Private Sub vsfSQL_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    With vsfSQL
        If .TextMatrix(NewRow, NewCol) <> "" And (NewCol = .ColIndex("SQL文本") Or (NewCol = .ColIndex("阻塞会话") And mstrUser <> "")) Then
            '显示按钮条件判断:
            '1.sql文本不为空
            '2.阻塞会话不为空的且非弹出窗体(通过是否传入SID判断,没有传入User,说明是弹出窗体)
            .ComboList() = "..."
            .FocusRect = flexFocusSolid
        Else
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vsfSQL_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim lngSid As Long, lngSerial As Long
    
    With vsfSQL
        Select Case Col
        
        Case .ColIndex("SQL文本")
            frmHistSqlPlan.ShowMe .TextMatrix(Row, .ColIndex("SQL_ID"))
        Case .ColIndex("阻塞会话")
            lngSid = Split(.TextMatrix(Row, Col), ",")(0)
            lngSerial = Split(.TextMatrix(Row, Col), ",")(1)
            frmHistSqlParent.ShowMe lngSid, lngSerial
        End Select
    End With
End Sub

Private Sub vsfSQL_DblClick()
    With vsfSQL
        If .Rows = 1 Then Exit Sub
        If .Rows = 0 Then Exit Sub
        
        If .Col <> .ColIndex("阻塞会话") Or .TextMatrix(.Row, .ColIndex("阻塞会话")) = "" Then
            frmHistSqlPlan.ShowMe .TextMatrix(.Row, .ColIndex("SQL_ID"))
        End If
    End With
End Sub
