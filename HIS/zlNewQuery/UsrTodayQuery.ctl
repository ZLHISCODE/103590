VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl UsrTodayQuery 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7755
   KeyPreview      =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   7755
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1050
      Top             =   4710
   End
   Begin VB.PictureBox picTitle 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   60
      ScaleHeight     =   465
      ScaleWidth      =   7665
      TabIndex        =   0
      Top             =   30
      Width           =   7665
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   3
         Left            =   3120
         TabIndex        =   10
         Top             =   15
         Width           =   900
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "特诊"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   2
         Left            =   2040
         TabIndex        =   9
         Top             =   15
         Width           =   900
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "专科"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   15
         Width           =   900
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "普通"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   1
         Left            =   1080
         TabIndex        =   2
         Top             =   15
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "专家"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   5
         Left            =   5040
         TabIndex        =   3
         Top             =   15
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   741
         Caption         =   "所有科室上班时间"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   10
         Left            =   6840
         TabIndex        =   7
         Top             =   15
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "帮助"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   4
         Left            =   4080
         TabIndex        =   11
         Top             =   15
         Width           =   900
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "急诊"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   6
         Left            =   6360
         TabIndex        =   12
         Top             =   15
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   741
         Caption         =   "挂号科室上班时间"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
   End
   Begin zl9NewQuery.ctlButton UsrCmd 
      Height          =   420
      Index           =   7
      Left            =   4410
      TabIndex        =   4
      Top             =   570
      Visible         =   0   'False
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   741
      Caption         =   "重读数据"
      BackColor       =   16777215
      FontSize        =   10.5
      TextAligment    =   0
   End
   Begin zl9NewQuery.ctlButton UsrCmd 
      Height          =   420
      Index           =   8
      Left            =   5415
      TabIndex        =   5
      Top             =   570
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   741
      Caption         =   "上翻"
      BackColor       =   16777215
      FontSize        =   10.5
      TextAligment    =   0
   End
   Begin zl9NewQuery.ctlButton UsrCmd 
      Height          =   420
      Index           =   9
      Left            =   6630
      TabIndex        =   6
      Top             =   570
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   741
      Caption         =   "下翻"
      BackColor       =   16777215
      FontSize        =   10.5
      TextAligment    =   0
   End
   Begin VSFlex8Ctl.VSFlexGrid msfResult 
      Height          =   3285
      Left            =   555
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1800
      Width           =   4740
      _cx             =   1993679017
      _cy             =   1993676450
      Appearance      =   0
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
      BackColorFixed  =   15199202
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16633516
      ForeColorSel    =   16711680
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   16761024
      GridColorFixed  =   16761024
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
   Begin MSComctlLib.ImageList ilsImage 
      Left            =   6960
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrTodayQuery.ctx":0000
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrTodayQuery.ctx":039A
            Key             =   "up"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrTodayQuery.ctx":0734
            Key             =   "down"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrTodayQuery.ctx":0ACE
            Key             =   "menu1"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrTodayQuery.ctx":0E68
            Key             =   "menu2"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrTodayQuery.ctx":20EA
            Key             =   "menu3"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrTodayQuery.ctx":2484
            Key             =   "menu4"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrTodayQuery.ctx":281E
            Key             =   "time"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrTodayQuery.ctx":2BB8
            Key             =   "patient"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrTodayQuery.ctx":2F52
            Key             =   "back"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrTodayQuery.ctx":32EC
            Key             =   "unselect"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrTodayQuery.ctx":3686
            Key             =   "select"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrTodayQuery.ctx":3A20
            Key             =   "next"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrTodayQuery.ctx":3DBA
            Key             =   "finish"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrTodayQuery.ctx":4154
            Key             =   "clear"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrTodayQuery.ctx":44EE
            Key             =   "help"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "UsrTodayQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mvarCurPos As Long
Private mvarRows As Long

Private mvarInternal As Long
Private mvarLong As Long

Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)

Public Sub InitLoad()
    '初始化进入
    
    UsrCmd(0).Picture = ilsImage.ListImages("unselect")
    UsrCmd(1).Picture = ilsImage.ListImages("unselect")
    UsrCmd(2).Picture = ilsImage.ListImages("unselect")
    UsrCmd(3).Picture = ilsImage.ListImages("unselect")
    UsrCmd(4).Picture = ilsImage.ListImages("unselect")
    UsrCmd(5).Picture = ilsImage.ListImages("unselect")
    UsrCmd(6).Picture = ilsImage.ListImages("unselect")
    UsrCmd(7).Picture = ilsImage.ListImages("refresh")
    UsrCmd(8).Picture = ilsImage.ListImages("up")
    UsrCmd(9).Picture = ilsImage.ListImages("down")
    UsrCmd(10).Picture = ilsImage.ListImages("help")
    
    UsrCmd(7).Enabled = False
    UsrCmd(8).Enabled = False
    
    UsrCmd(5).Visible = IIf(Val(GetPara("今日就诊可查询科室上班时间", "0")) = 1, True, False)
    UsrCmd(6).Visible = IIf(Val(GetPara("今日就诊可查询科室上班时间", "0")) = 1, True, False)

    
    tmr.Enabled = False
        mvarLong = Val(GetPara("今日就诊刷新间隔", "5")) * 1000
        tmr.Enabled = IIf(mvarLong = 0, False, True)
    mvarInternal = 0
    
    Call UsrCmd_CommandClick(0)
End Sub

Private Sub msfResult_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub tmr_Timer()
    mvarInternal = mvarInternal + 1000
    If mvarInternal < mvarLong Then Exit Sub
    mvarInternal = 0
    '刷新
    Call UsrCmd_CommandClick(7)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
        
    Call ResizeControl(picTitle, 0, 0, UserControl.Width, picTitle.Height)
    
    UsrCmd(9).Left = UserControl.Width - UsrCmd(8).Width - 30
    Call ResizeControl(UsrCmd(8), UsrCmd(9).Left - UsrCmd(8).Width - 30, UsrCmd(9).Top, UsrCmd(8).Width, UsrCmd(8).Height)
    Call ResizeControl(UsrCmd(7), UsrCmd(8).Left - UsrCmd(7).Width - 30, UsrCmd(9).Top, UsrCmd(7).Width, UsrCmd(7).Height)
    Call ResizeControl(UsrCmd(10), picTitle.ScaleWidth - UsrCmd(10).Width - 60, UsrCmd(0).Top, UsrCmd(10).Width, UsrCmd(10).Height)
    
    Call ResizeControl(UsrCmd(0), UsrCmd(0).Left, UsrCmd(0).Top, UsrCmd(5).Width, UsrCmd(5).Height)
    Call ResizeControl(UsrCmd(1), UsrCmd(0).Left + UsrCmd(0).Width + 60, UsrCmd(0).Top, UsrCmd(5).Width, UsrCmd(5).Height)
    Call ResizeControl(UsrCmd(2), UsrCmd(1).Left + UsrCmd(1).Width + 60, UsrCmd(0).Top, UsrCmd(5).Width, UsrCmd(5).Height)
    Call ResizeControl(UsrCmd(3), UsrCmd(2).Left + UsrCmd(2).Width + 60, UsrCmd(0).Top, UsrCmd(5).Width, UsrCmd(5).Height)
    Call ResizeControl(UsrCmd(4), UsrCmd(3).Left + UsrCmd(3).Width + 60, UsrCmd(0).Top, UsrCmd(5).Width, UsrCmd(5).Height)
    Call ResizeControl(UsrCmd(5), UsrCmd(4).Left + UsrCmd(4).Width + 320, UsrCmd(0).Top, UsrCmd(5).Width, UsrCmd(5).Height)
    Call ResizeControl(UsrCmd(6), UsrCmd(5).Left + UsrCmd(5).Width + 60, UsrCmd(0).Top, UsrCmd(5).Width, UsrCmd(5).Height)
    
    Call ResizeControl(msfResult, 0, UsrCmd(7).Top + UsrCmd(7).Height + 60, UserControl.Width, UserControl.Height - UsrCmd(7).Top - UsrCmd(7).Height - 60)
    
End Sub

Private Sub UsrCmd_CommandClick(Index As Integer)
    Dim i As Integer
       
    UsrCmd(0).Picture = ilsImage.ListImages("unselect")
    UsrCmd(1).Picture = ilsImage.ListImages("unselect")
    UsrCmd(2).Picture = ilsImage.ListImages("unselect")
    UsrCmd(3).Picture = ilsImage.ListImages("unselect")
    UsrCmd(4).Picture = ilsImage.ListImages("unselect")
    UsrCmd(5).Picture = ilsImage.ListImages("unselect")
    UsrCmd(6).Picture = ilsImage.ListImages("unselect")
    
    Select Case Index
    Case 0, 1, 2, 3, 4, 5, 6            '挂号安排、专家门诊、上班查询
        Call DrawMsfHeader(Index)
        mvarCurPos = 1
        mvarRows = 0
        
        UsrCmd(0).State = 0
        UsrCmd(1).State = 0
        UsrCmd(2).State = 0
        UsrCmd(3).State = 0
        UsrCmd(4).State = 0
        UsrCmd(5).State = 0
        UsrCmd(6).State = 0
        UsrCmd(Index).State = -1
                        
        Select Case Index
        Case 0              '普通
            Call LoadRegPlan(0)
        Case 1              '专家
            Call LoadRegPlan(1)
        Case 2              '专科
            Call LoadRegPlan(2)
        Case 3              '特诊
            Call LoadRegPlan(3)
        Case 4              '急诊
            Call LoadRegPlan(4)
        Case 5              '上班查询
            Call LoadDeptWorkTime
        Case 6
            Call LoadDeptWorkTime(True)
        End Select
        
        UsrCmd(Index).Picture = ilsImage.ListImages("select")
        Call EnablePageButton(msfResult, mvarCurPos, mvarRows, UsrCmd(8), UsrCmd(9))
    
    Case 7
        For i = 0 To 6
            If UsrCmd(i).State = -1 Then
                Call UsrCmd_CommandClick(i)
                UsrCmd(i).Picture = ilsImage.ListImages("select")
            End If
        Next
    Case 8, 9
        Call TurnToPage(msfResult, IIf(Index = 7, -1, 1), mvarCurPos)
        Call EnablePageButton(msfResult, mvarCurPos, mvarRows, UsrCmd(8), UsrCmd(9))
    Case 10
        Call frmHelp.ShowHelp(Me, -3, UserControl.Width, UserControl.Height)
    End Select
    
End Sub

Private Sub DrawMsfHeader(ByVal bytMode As Byte)
'功能:根据所选功能作出相应的表格
'参数:bytMode       所选取功能索引
    
    With msfResult
        .Rows = 2
        .Cols = 0
        ClearSpecRowCol msfResult, 1, Array()
                
        Select Case bytMode
        Case 0, 1, 2, 3, 4  '
            Call AddColumn(msfResult, "科室", 900, 1)
            Call AddColumn(msfResult, "项目", 2100, 1)
            Call AddColumn(msfResult, "医生", 850, 1)
            Call AddColumn(msfResult, "限号", 500, 7)
            Call AddColumn(msfResult, "已挂", 0, 7)
            Call AddColumn(msfResult, "周日", 500, 4)
            Call AddColumn(msfResult, "周一", 500, 4)
            Call AddColumn(msfResult, "周二", 500, 4)
            Call AddColumn(msfResult, "周三", 500, 4)
            Call AddColumn(msfResult, "周四", 500, 4)
            Call AddColumn(msfResult, "周五", 500, 4)
            Call AddColumn(msfResult, "周六", 500, 4)
            Call AddColumn(msfResult, "应诊诊室", 1900, 1)
            Call AddColumn(msfResult, "", 1200, 1)
            Call CalcAutoColWidth(msfResult, 12)
'        Case 1      '
'            Call AddColumn(msfResult, "科室", 900, 1)
'            Call AddColumn(msfResult, "项目", 2100, 1)
'            Call AddColumn(msfResult, "医生", 850, 1)
'            Call AddColumn(msfResult, "限号", 500, 7)
'            Call AddColumn(msfResult, "已挂", 0, 7)
'            Call AddColumn(msfResult, "周日", 500, 4)
'            Call AddColumn(msfResult, "周一", 500, 4)
'            Call AddColumn(msfResult, "周二", 500, 4)
'            Call AddColumn(msfResult, "周三", 500, 4)
'            Call AddColumn(msfResult, "周四", 500, 4)
'            Call AddColumn(msfResult, "周五", 500, 4)
'            Call AddColumn(msfResult, "周六", 500, 4)
'            Call AddColumn(msfResult, "应诊诊室", 1900, 1)
'            Call AddColumn(msfResult, "", 1200, 1)
'            Call CalcAutoColWidth(msfResult, 12)
        Case 5, 6    '
            Call AddColumn(msfResult, "科室", 1500, 1)
            Call AddColumn(msfResult, "序号", 600, 1)
            Call AddColumn(msfResult, "周日", 1200, 4)
            Call AddColumn(msfResult, "周一", 1200, 4)
            Call AddColumn(msfResult, "周二", 1200, 4)
            Call AddColumn(msfResult, "周三", 1200, 4)
            Call AddColumn(msfResult, "周四", 1200, 4)
            Call AddColumn(msfResult, "周五", 1200, 4)
            Call AddColumn(msfResult, "周六", 1200, 4)
            Call AddColumn(msfResult, "", 1200, 1)
            Call CalcAutoColWidth(msfResult, 0)
        End Select
        
    End With
End Sub

Private Sub LoadRegPlan(Optional ByVal Index As Integer = 0)
'功能:将所有安排装入到mshPlan
'返回值:装入成功返回True,否则返回False
    Dim i As Long
    Dim vDay As Long
    
    vDay = Format(zlDatabase.Currentdate, "w") + 4
    
    msfResult.Rows = 2
    ClearSpecRowCol msfResult, 1, Array()
    
    On Error GoTo errH
    
    Select Case Index
    Case 0              '普通
        gstrSQL = "Select A.ID,A.号码,A.科室ID,C.名称 as 科室,A.项目ID,B.名称 as 项目," & _
            " A.医生姓名,A.医生ID,F.限号数,周日, 周一, 周二, 周三, 周四, 周五, 周六" & _
            " From 挂号安排 A,收费项目目录 B,部门表 C,挂号安排限制 F" & _
            " Where A.项目ID = B.ID AND B.类别='1' And A.科室ID = C.ID And A.号类='普通' And 项目特性<>1 And " & GetNodeCheckSQL("b.站点") & " And " & GetNodeCheckSQL("c.站点") & " " & _
            "     And a.id=F.安排ID(+)  And Decode(To_Char(Sysdate, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) =F.限制项目(+) " & vbNewLine & _
            "     And A.停用日期 Is Null And Not Exists(Select 1 From 挂号安排停用状态 Where 安排ID=A.ID and Sysdate between 开始停止时间 and 结束停止时间 )" & _
            "     And sysDate Between Nvl(a.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And Nvl(a.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
            " Order by C.名称"
    Case 1              '专家
        gstrSQL = "Select A.ID,A.号码,A.科室ID,C.名称 as 科室,A.项目ID,B.名称 as 项目," & _
            " A.医生姓名,A.医生ID,F.限号数,周日, 周一, 周二, 周三, 周四, 周五, 周六" & _
            " From 挂号安排 A,收费项目目录 B,部门表 C,挂号安排限制 F " & _
            " Where A.项目ID = B.ID AND B.类别='1' And A.科室ID = C.ID And A.号类='专家' And 项目特性<>1 And " & GetNodeCheckSQL("b.站点") & " And " & GetNodeCheckSQL("c.站点") & " " & _
            "      And a.id=F.安排ID(+)  And Decode(To_Char(Sysdate, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) =F.限制项目(+) " & vbNewLine & _
            "     And A.停用日期 Is Null And Not Exists(Select 1 From 挂号安排停用状态 Where 安排ID=A.ID and Sysdate between 开始停止时间 and 结束停止时间 )" & _
            "     And sysDate Between Nvl(a.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And Nvl(a.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
            " Order by C.名称"
    Case 2              '专科
        gstrSQL = "Select A.ID,A.号码,A.科室ID,C.名称 as 科室,A.项目ID,B.名称 as 项目," & _
            " A.医生姓名,A.医生ID,F.限号数,周日, 周一, 周二, 周三, 周四, 周五, 周六" & _
            " From 挂号安排 A,收费项目目录 B,部门表 C,挂号安排限制 F" & _
            " Where A.项目ID = B.ID AND B.类别='1' And A.科室ID = C.ID And A.号类='专科' And 项目特性<>1 And " & GetNodeCheckSQL("b.站点") & " And " & GetNodeCheckSQL("c.站点") & " " & _
            "     And a.id=F.安排ID(+)  And Decode(To_Char(Sysdate, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) =F.限制项目(+) " & vbNewLine & _
            "     And A.停用日期 Is Null And Not Exists(Select 1 From 挂号安排停用状态 Where 安排ID=A.ID and Sysdate between 开始停止时间 and 结束停止时间 )" & _
            "     And sysDate Between Nvl(a.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And Nvl(a.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
            " Order by C.名称"
    Case 3              '特诊
        gstrSQL = "Select A.ID,A.号码,A.科室ID,C.名称 as 科室,A.项目ID,B.名称 as 项目," & _
            " A.医生姓名,A.医生ID,F.限号数,周日, 周一, 周二, 周三, 周四, 周五, 周六" & _
            " From 挂号安排 A,收费项目目录 B,部门表 C,挂号安排限制 F" & _
            " Where A.项目ID = B.ID AND B.类别='1' And A.科室ID = C.ID And A.号类='特诊' And 项目特性<>1 And " & GetNodeCheckSQL("b.站点") & " And " & GetNodeCheckSQL("c.站点") & " " & _
            "     And a.id=F.安排ID(+)  And Decode(To_Char(Sysdate, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) =F.限制项目(+) " & vbNewLine & _
            "     And A.停用日期 Is Null And Not Exists(Select 1 From 挂号安排停用状态 Where 安排ID=A.ID and Sysdate between 开始停止时间 and 结束停止时间 )" & _
            "     And sysDate Between Nvl(a.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And Nvl(a.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
            " Order by C.名称"
    Case 4              '急诊
        gstrSQL = "Select A.ID,A.号码,A.科室ID,C.名称 as 科室,A.项目ID,B.名称 as 项目," & _
            " A.医生姓名,A.医生ID,F.限号数,周日, 周一, 周二, 周三, 周四, 周五, 周六" & _
            " From 挂号安排 A,收费项目目录 B,部门表 C,挂号安排限制 F" & _
            " Where A.项目ID = B.ID AND B.类别='1' And A.科室ID = C.ID And B.项目特性 = 1 And " & GetNodeCheckSQL("b.站点") & " And " & GetNodeCheckSQL("c.站点") & " " & _
            "     And a.id=F.安排ID(+)  And Decode(To_Char(Sysdate, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) =F.限制项目(+) " & vbNewLine & _
            "     And A.停用日期 Is Null And Not Exists(Select 1 From 挂号安排停用状态 Where 安排ID=A.ID and Sysdate between 开始停止时间 and 结束停止时间 )" & _
            "     And sysDate Between Nvl(a.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And Nvl(a.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
            " Order by C.名称"
    End Select
    
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "今日就诊")
        
    With msfResult
        .Redraw = False
        i = 1
        If gRs.BOF = False Then
            While Not gRs.EOF
                .TextMatrix(i, 0) = gRs!科室
                .TextMatrix(i, 1) = gRs!项目
                .TextMatrix(i, 2) = IIf(IsNull(gRs!医生姓名), "", gRs!医生姓名)
                .TextMatrix(i, 3) = IIf(IsNull(gRs!限号数), "", gRs!限号数)
                '.TextMatrix(i, 4) = GetHaveRegist(IIf(IsNull(gRs!科室ID), 0, gRs!科室ID), IIf(IsNull(gRs!项目ID), 0, gRs!项目ID), IIf(IsNull(gRs!医生ID), 0, gRs!医生ID), IIf(IsNull(gRs!医生姓名), "医生", gRs!医生姓名))
                .TextMatrix(i, 5) = IIf(IsNull(gRs!周日), "", gRs!周日)
                .TextMatrix(i, 6) = IIf(IsNull(gRs!周一), "", gRs!周一)
                .TextMatrix(i, 7) = IIf(IsNull(gRs!周二), "", gRs!周二)
                .TextMatrix(i, 8) = IIf(IsNull(gRs!周三), "", gRs!周三)
                .TextMatrix(i, 9) = IIf(IsNull(gRs!周四), "", gRs!周四)
                .TextMatrix(i, 10) = IIf(IsNull(gRs!周五), "", gRs!周五)
                .TextMatrix(i, 11) = IIf(IsNull(gRs!周六), "", gRs!周六)
                .TextMatrix(i, 12) = Read应诊诊室(gRs!ID)
                i = i + 1
                .Rows = i + 1
                gRs.MoveNext
            Wend
        End If
        If .Rows > 2 Then .Rows = .Rows - 1
        SetDefaultDate msfResult, vDay
        mvarRows = msfResult.Rows - 1
        .Redraw = True
    End With
    msfResult.Rows = msfResult.Rows + 50
    
    ShowPlan = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadDeptWorkTime(Optional ByVal bln只显示可挂号科室 As Boolean = False)
    Dim i As Long
    Dim curRow As Long
    Dim startRow As Long
    Dim vSvrLastDept As String
    Dim vOrder As Long
    Dim vOrderDate As Long
    Dim vDay As Long
    
    vDay = Format(zlDatabase.Currentdate, "w") + 1
    On Error GoTo errH
    
    If bln只显示可挂号科室 Then
        gstrSQL = "Select B.名称,A.星期,A.开始时间,A.终止时间 from 部门安排 A,部门表 B where A.部门ID=B.ID And B.id In (Select Distinct 科室ID As ID From 挂号安排) order by A.部门ID,A.星期"
    Else
        gstrSQL = "Select B.名称,A.星期,A.开始时间,A.终止时间 from 部门安排 A,部门表 B where A.部门ID=B.ID order by A.部门ID,A.星期"
    End If
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "今日就诊")
            
    With msfResult
        .MergeCells = flexMergeFree
        .MergeCol(0) = True
        .Redraw = False
        i = 1
        vOrderDate = -1
        If gRs.BOF = False Then
            While Not gRs.EOF
                If vSvrLastDept <> gRs!名称 Then
                   vSvrLastDept = gRs!名称
                   vOrder = 1
                   vOrderDate = -1
                   vOrderDate = IIf(IsNull(gRs!星期), -1, gRs!星期)
                   .Rows = .Rows + 1
                   i = .Rows - 1
                   curRow = i
                   startRow = i
                Else
                    If vOrderDate <> IIf(IsNull(gRs!星期), -1, gRs!星期) Then
                        vOrderDate = IIf(IsNull(gRs!星期), -1, gRs!星期)
                        curRow = startRow
                        vOrder = 1
                    Else
                        If curRow + vOrder > .Rows - 1 Then
                            .Rows = .Rows + 1
                            vOrder = vOrder + 1
                            i = .Rows - 1
                            curRow = i
                        Else
                            curRow = curRow + 1
                        End If
                    End If
                End If

                .RowData(curRow) = 1
                .TextMatrix(curRow, 0) = gRs!名称
                .TextMatrix(curRow, vOrderDate + 2) = IIf(IsNull(gRs!开始时间), "", Format(gRs!开始时间, "HH:MM")) & "-" & IIf(IsNull(gRs!终止时间), "", Format(gRs!终止时间, "HH:MM"))

                gRs.MoveNext
            Wend
            If .Rows > 2 Then .RemoveItem 1
            SetDefaultDate msfResult, vDay
            mvarRows = .Rows - 1
        End If
        
        vSvrLastDept = ""
        vOrder = 0
        For i = 1 To .Rows - 1
            If vSvrLastDept <> .TextMatrix(i, 0) Then
                vSvrLastDept = .TextMatrix(i, 0)
                vOrder = 1
            Else
                vOrder = vOrder + 1
            End If
            If .RowData(i) = 1 Then .TextMatrix(i, 1) = vOrder
        Next
        
        .Rows = .Rows + 50
        .Redraw = True
    End With
        
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Read应诊诊室(ByVal lngID As Long) As String
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If lngID = 0 Then Exit Function
        Set rsTmp = GetRs挂号诊室
        If rsTmp Is Nothing Then
            strSQL = "Select 号表ID,门诊诊室 From 挂号安排诊室 Where 号表ID=[1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "今日就诊", lngID)
        End If
        rsTmp.Filter = "号表ID=" & lngID
        If rsTmp.RecordCount = 0 Then Exit Function
        Do While Not rsTmp.EOF
            Read应诊诊室 = Read应诊诊室 & ";" & rsTmp!门诊诊室
            rsTmp.MoveNext
        Loop
        Read应诊诊室 = Mid(Read应诊诊室, 2)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetDefaultDate(msf As Object, ByVal vDay As Long)
    Dim svrRow As Long
    Dim svrCol As Long
    Dim i As Long
                
    msf.Redraw = False
    svrRow = msf.Row
    svrCol = msf.Col
    msf.Col = vDay
    For i = 0 To msf.Rows - 1
        msf.Row = i
        msfResult.CellForeColor = &HFF0000
    Next
    msf.Row = svrRow
    msf.Col = svrCol
    msf.Redraw = True
End Sub

Private Sub UsrCmd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Function GetHaveRegist(ByVal lng科室ID As Long, ByVal lng项目ID As Long, _
    ByVal lng医生id As Long, ByVal str医生姓名 As String, ByVal str号码 As String) As String
    Dim rs As New ADODB.Recordset
    On Error Resume Next
    GetHaveRegist = ""
    gstrSQL = "" & _
    "   Select 已挂数 from 病人挂号汇总  " & _
    "   Where 科室ID=[1] And 项目ID=[2] and Nvl(医生ID,0)=[3] and Nvl(医生姓名,'医生')=[4] " & _
    "       And 日期=Trunc(Sysdate) And (号码 =[4] or 号码 is null )"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "今日就诊", lng科室ID, lng项目ID, lng医生id, str医生姓名, str号码)
    If rs.BOF = False Then GetHaveRegist = IIf(IsNull(rs!已挂数), "", rs!已挂数)
    CloseRecord rs
    If Val(GetHaveRegist) = 0 Then GetHaveRegist = ""
    
End Function

Public Property Let Enabled(ByVal vData As Boolean)
    UserControl.Enabled = vData
End Property
