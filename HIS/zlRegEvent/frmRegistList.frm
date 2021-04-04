VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmRegistList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "出诊表"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8955
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRegistList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "退出(&E)"
      Height          =   390
      Left            =   7065
      TabIndex        =   7
      ToolTipText     =   "热键:F2"
      Top             =   5775
      Width           =   1350
   End
   Begin VB.Frame fraInfo 
      Caption         =   "号源信息"
      Height          =   1455
      Left            =   90
      TabIndex        =   0
      Top             =   105
      Width           =   8760
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         Caption         =   "是否建档:"
         Height          =   240
         Left            =   5040
         TabIndex        =   10
         Top             =   675
         Width           =   1080
      End
      Begin VB.Label lblControlDays 
         AutoSize        =   -1  'True
         Caption         =   "预约天数:"
         Height          =   240
         Left            =   5040
         TabIndex        =   9
         Top             =   1035
         Width           =   1080
      End
      Begin VB.Label lblControl 
         AutoSize        =   -1  'True
         Caption         =   "排班模式:"
         Height          =   240
         Left            =   255
         TabIndex        =   8
         Top             =   1035
         Width           =   1080
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Caption         =   "号类:"
         Height          =   240
         Left            =   2370
         TabIndex        =   5
         Top             =   315
         Width           =   600
      End
      Begin VB.Label lblDoc 
         Caption         =   "医生:"
         Height          =   240
         Left            =   255
         TabIndex        =   4
         Top             =   675
         Width           =   2070
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "项目:"
         Height          =   240
         Left            =   5040
         TabIndex        =   3
         Top             =   315
         Width           =   600
      End
      Begin VB.Label lblDept 
         Caption         =   "科室:"
         Height          =   240
         Left            =   2370
         TabIndex        =   2
         Top             =   675
         Width           =   2565
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         Caption         =   "号码:"
         Height          =   240
         Left            =   255
         TabIndex        =   1
         Top             =   315
         Width           =   600
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   3960
      Left            =   90
      TabIndex        =   6
      Top             =   1680
      Width           =   8760
      _cx             =   15452
      _cy             =   6985
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   16185078
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRegistList.frx":058A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      Editable        =   0
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
End
Attribute VB_Name = "frmRegistList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng号源Id As Long
Private mblnUnload As Boolean

Public Sub ShowMe(frmMain As Object, ByVal lng号源Id As Long)
    mlng号源Id = lng号源Id
    Me.Show vbModal, frmMain
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If mblnUnload Then
        mblnUnload = False
        Unload Me
    End If
    If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
End Sub

Private Sub Form_Load()
    Dim strSQL As String, strCurDate As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    strSQL = "Select 号类, 号码, b.名称 As 科室, c.名称 As 项目, a.医生姓名, a.是否建病案, a.排班方式, a.预约天数" & vbNewLine & _
            "From 临床出诊号源 A, 部门表 B, 收费项目目录 C" & vbNewLine & _
            "Where a.科室id = b.Id And a.项目id = c.Id And a.Id = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng号源Id)
    If rsTemp.EOF Then
        MsgBox "无法确认号源信息,读取出诊表失败!", vbInformation, gstrSysName
        mblnUnload = True
        Exit Sub
    Else
        lblNO.Caption = "号码:" & rsTemp!号码
        lblType.Caption = "号类:" & rsTemp!号类
        lblDept.Caption = "科室:" & rsTemp!科室
        lblItem.Caption = "项目:" & rsTemp!项目
        lblDoc.Caption = "医生:" & rsTemp!医生姓名
        lblPati.Caption = "是否建档:" & IIf(Val(Nvl(rsTemp!是否建病案)) = 0, "否", "是")
        Select Case Val(Nvl(rsTemp!排班方式))
        Case 0
            lblControl.Caption = "排班模式:固定排班"
        Case 1
            lblControl.Caption = "排班模式:按月排班"
        Case 2
            lblControl.Caption = "排班模式:按周排班"
        End Select
        lblControlDays.Caption = "预约天数:" & Val(Nvl(rsTemp!预约天数, gint预约天数)) & "天"
    End If
    
    strSQL = "Select 出诊日期,上班时段,开始时间,终止时间,已挂数,限号数,已约数,限约数 From 临床出诊记录 Where 号源ID=[1] And 出诊日期 >= Trunc(Sysdate) Order By 出诊日期,开始时间"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng号源Id)
    With vsfList
        .Clear 1
        .Rows = 1
        Do While Not rsTemp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = Format(rsTemp!出诊日期, "yyyy-mm-dd")
            .TextMatrix(.Rows - 1, 1) = rsTemp!上班时段 & "(" & Format(rsTemp!开始时间, "hh:mm") & "-" & Format(rsTemp!终止时间, "hh:mm") & ")"
            .TextMatrix(.Rows - 1, 2) = Nvl(rsTemp!已挂数, 0) & "/" & Nvl(rsTemp!限号数, "∞")
            .TextMatrix(.Rows - 1, 3) = Nvl(rsTemp!已约数, 0) & "/" & Nvl(rsTemp!限约数, "∞")
            rsTemp.MoveNext
        Loop
        .MergeCol(0) = True
        .AutoSize 0, .Cols - 1
    End With
End Sub
