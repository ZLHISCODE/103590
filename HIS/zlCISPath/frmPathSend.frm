VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Begin VB.Form frmPathSend 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "生成路径项目"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11775
   Icon            =   "frmPathSend.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   11775
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdPati 
      Height          =   240
      Left            =   7560
      Picture         =   "frmPathSend.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "选择婴儿"
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox lstPati 
      Appearance      =   0  'Flat
      Height          =   1080
      ItemData        =   "frmPathSend.frx":6948
      Left            =   5160
      List            =   "frmPathSend.frx":6955
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   16
      Top             =   4080
      Visible         =   0   'False
      Width           =   2295
   End
   Begin zlRichEditor.Editor edtEditor 
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11775
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   8025
      Width           =   11775
      Begin VB.CommandButton cmdMergeStep 
         Caption         =   "合并路径阶段选择(&S)"
         Height          =   350
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   9360
         TabIndex        =   14
         Top             =   120
         Width           =   1100
      End
      Begin MSComCtl2.DTPicker dtpAdviceTime 
         Height          =   300
         Left            =   7320
         TabIndex        =   11
         Top             =   145
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   190513155
         CurrentDate     =   41129.5916666667
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   10560
         TabIndex        =   8
         Top             =   120
         Width           =   1100
      End
      Begin VB.Label lblAdviceTime 
         Caption         =   "医嘱缺省开始时间"
         Height          =   180
         Left            =   5760
         TabIndex        =   10
         Top             =   205
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   11760
         Y1              =   45
         Y2              =   45
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   11760
         Y1              =   30
         Y2              =   30
      End
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   11775
      TabIndex        =   3
      Top             =   0
      Width           =   11775
      Begin VB.Label lblFont 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Height          =   180
         Left            =   3960
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "当前时间阶段："
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblDate 
         BackStyle       =   0  'Transparent
         Caption         =   "当前时间："
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   120
         Width           =   4455
      End
      Begin VB.Label lblPhase 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmPathSend.frx":6971
         Height          =   615
         Left            =   2520
         TabIndex        =   4
         Top             =   360
         Width           =   6855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   10000
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   11760
         Y1              =   945
         Y2              =   945
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   195
         Picture         =   "frmPathSend.frx":6A0B
         Top             =   45
         Width           =   720
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsItem 
      Height          =   3405
      Index           =   0
      Left            =   30
      TabIndex        =   2
      Top             =   1950
      Width           =   11655
      _cx             =   20558
      _cy             =   6006
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      BackColorAlternate=   16777215
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   7
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   320
      ColWidthMin     =   0
      ColWidthMax     =   8000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathSend.frx":7293
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
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
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPhase 
      Height          =   705
      Left            =   30
      TabIndex        =   0
      Top             =   1200
      Width           =   11655
      _cx             =   20558
      _cy             =   1244
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
      BackColor       =   15597549
      ForeColor       =   -2147483640
      BackColorFixed  =   15597549
      ForeColorFixed  =   15724768
      BackColorSel    =   45056
      ForeColorSel    =   16777215
      BackColorBkg    =   15597549
      BackColorAlternate=   15597549
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   32768
      FloodColor      =   192
      SheetBorder     =   15724768
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   2
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   450
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathSend.frx":7438
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
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
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.TabStrip tabBranch 
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   870
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "主路径"
            Key             =   "_0"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsItem 
      Height          =   2355
      Index           =   1
      Left            =   30
      TabIndex        =   12
      Top             =   5640
      Width           =   11655
      _cx             =   20558
      _cy             =   4154
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      BackColorAlternate=   16777215
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   7
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   320
      ColWidthMin     =   0
      ColWidthMax     =   8000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathSend.frx":74CD
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
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
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lblMerge 
      Caption         =   "合并路径:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   5400
      Width           =   11415
   End
End
Attribute VB_Name = "frmPathSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngFun '0-生成路径，1-补充生成(不能选择阶段),2-查看路径阶段定义,3-重新生成医嘱

Private mPP As TYPE_PATH_Pati
Private mPati As TYPE_Pati

Private mint场合 As Integer  '0-医生站调用,1-护士站调用
Private mlng时间进度 As Integer 'mlngFun=0时传入，1=下一阶段提前至今天,2=下一阶段提前至明天,-1=下一阶段延后（继续当前阶段）,0=正常

Private mlng项目ID As Long       '重新生成的项目ID
Private mlng执行ID As Long       '重新生成的路径执行ID

Private mlng病人阶段ID As Long   '当前选择的阶段(查看时)，或病人当前阶段（生成时）
Private mlng天数 As Long         '当前应该生成的天数(实际天数)
Private mdat时间 As Date         '当前应该进入的日期(生成路径项目及医嘱的日期和时间)
Private mlng补录间隔 As Long
Private mlng提前天数 As Long         '传入时应该生成的天数(用于下一阶段提前时)
Private mlng路径医嘱天数 As Long   '路径医嘱生成超前天数
Private mblnIsHaveBranch As Boolean  '是否存在分支路径
Private mdatDur As Date          '路径生成时间
Private mrsMerge As ADODB.Recordset
Private mstrMerge As String      '已经选择的合并路径阶段：分支1:阶段ID1,分支2:阶段ID2........
Private mstrMergeStep As String  '已经选择合并路径阶段,用于生成：合并路径记录ID1:阶段ID1,合并路径记录ID2:阶段ID2........
Private mclsMipModule As zl9ComLib.clsMipModule ' 消息平台对象

Private mstrBaby As String '婴儿姓名串,例：马云,马克思,...
Private mfrmParent As Object
Private mrsPhase As ADODB.Recordset
Private mcol As Collection
Private mEditType As Collection
Private mlngMergeCount As Long   '合并路径数

Private Enum 执行方式
    T0无需执行 = 0
    T1每天必须 = 1
    T2至少一次 = 2
    T3必要时 = 3
    T4必须且仅一次 = 4
End Enum

Private Enum TYPE_Func
    Func生成路径 = 0
    Func补充生成 = 1
    Func查看路径 = 2
    Func重新生成 = 3
End Enum

Private mblnOK As Boolean

Public Function ShowMe(frmParent As Object, ByVal lngFun As Long, ByVal int场合 As Integer, t_pati As TYPE_Pati, t_pp As TYPE_PATH_Pati, _
    ByVal lng病人阶段ID As Long, ByVal lng天数 As Long, Optional ByVal lng项目ID As Long, Optional ByVal lng执行ID As Long, _
    Optional ByVal lng时间进度 As Long, Optional ByRef objMip As Object, Optional ByVal bln提前 As Boolean = False, Optional ByVal strSQLPhase As String) As Boolean
'参数：lng项目ID,lng执行ID=重新生成医嘱时才需传入
'      lng时间进度=mlngFun=0时传入，1=下一阶段提前,2-下一阶段提前至明天,-1=下一阶段延后（继续当前阶段）,0=正常
'      bln提前=true :提前生成路径,False-非提前生成
'     strSQLPhase-护理路径生成时传人 SQL语句,字段列：阶段ID,日期,天数
    Set mfrmParent = frmParent
    mlngFun = lngFun
    mint场合 = int场合
    mlng项目ID = lng项目ID
    mlng执行ID = lng执行ID
    
    mPati = t_pati
    mPP = t_pp
    mlng病人阶段ID = lng病人阶段ID  '缺省选中当前阶段
    mlng天数 = lng天数
    mlng提前天数 = lng天数
    mlng时间进度 = lng时间进度
    If bln提前 Then
        '提前生成
        mdatDur = DateAdd("d", 1, CDate(Format(mPP.当前日期, "YYYY-MM-DD 00:00:00")))
    Else
        mdatDur = zlDatabase.Currentdate
    End If
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    Set mrsPhase = GetPhase(mPP.路径ID, mPP.版本号, mlng病人阶段ID, mPP.当前阶段分支ID, mlng天数, , strSQLPhase)
    If mrsPhase.RecordCount = 0 Then
        MsgBox "当前时间(第" & lng天数 & "天)没有适用的路径阶段，不能生成路径项目。" & vbCrLf & "可能是病人入院天数超过了标准住院日，或者没有后续的时间阶段。", vbInformation, gstrSysName
        Exit Function
    End If
    
    mblnOK = False
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Function GetPhase(ByVal lng路径ID As Long, ByVal lng版本号 As Long, ByVal lng当前阶段ID As Long, ByVal lng当前阶段分支ID As Long, _
        ByVal lng天数 As Long, Optional ByVal lng合并路径记录ID As Long, Optional ByVal strPhase As String) As ADODB.Recordset
'功能：读取当前时间可用的阶段
'参数：lng合并路径记录ID =合并路径记录ID
'     strPhase -护理路径生成时传人
    Dim strSql As String, strIF As String, str阶段分类 As String
    Dim rsTmp As ADODB.Recordset, datPathIn As Date, lng时间进度 As Long
    Dim lng理论天数 As Long, lng序号 As Long
    Dim strMainIF As String
    Dim strSubSQL As String
    
    If mlngFun = 2 Then '查看阶段定义的项目
        strSql = "Select a.Id, Nvl(a.父id,0) as 父id, a.序号, a.名称, a.说明,a.开始天数, a.结束天数, a.分类,NVL(a.分支ID,0) AS 分支ID" & vbNewLine & _
                "From 临床路径阶段 A" & vbNewLine & _
                "Where a.路径id = [1] And a.版本号 = [2] And a.id = [4]" & vbNewLine & _
                "Order by 序号"
    Else
        datPathIn = GetPatiInPath(mPati, mPP.病人路径ID)
        If lng路径ID = mPP.路径ID Then
            mdat时间 = DateAdd("d", lng天数 - 1, datPathIn)  '当天应该生成的日期
        
            strSql = "Select To_number(Trunc(Sysdate)-Trunc([1])) 补录间隔 From Dual"
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取补录间隔", mdat时间)
            mlng补录间隔 = Val("" & rsTmp!补录间隔)
        End If
        If mlngFun = 0 Then
            If mint场合 = 1 Then
                '护士站生成可选阶段（与医生站不同阶段信息是从病人路径执行中取）
                strSql = "Select b.日期,b.天数,a.Id, Nvl(a.父id,0) as 父id, a.序号, a.名称, a.说明,a.开始天数, a.结束天数, a.分类,NVL(a.分支ID,0) AS 分支ID From (" & strPhase & ") B, 临床路径阶段 A Where a.Id = b.阶段ID order by b.排序 "
                On Error GoTo errH
                Set GetPhase = zlDatabase.OpenSQLRecord(strSql, "获取适用阶段", strPhase)
                Exit Function
            ElseIf mint场合 = 0 Then '医生站
                If mlng时间进度 = -1 Then    '延后时继续当前阶段
                    strIF = " And a.id = [4]"
                Else
                    If mPP.当前阶段ID <> 0 Then
                        If mPP.原路径ID = lng路径ID Or lng合并路径记录ID <> 0 Then
                            lng序号 = GetPhaseNO(IIf(lng合并路径记录ID <> 0, lng当前阶段ID, mPP.当前阶段ID))
                        Else
                            '如果跳转了路径，则新路径可能该病人以前用过，可用的阶段序号应大于等于上次用过的最大阶段序号。
                            lng序号 = GetLastPhaseNO(mPP.病人路径ID, lng路径ID)
                        End If
                    End If
                    
                    If mlng时间进度 = 1 Or mlng时间进度 = 2 Then
                        '时间进度=2,提前至明天,
                        '提前时显示下一天的可用阶段(只能是当前阶段的后续阶段。当天还有后续阶段不再使用，这种不应在评估时作为提前)
                        lng理论天数 = GetMustDay(mPP.病人路径ID, lng天数, , lng合并路径记录ID)
                        strIF = " And Decode(a.分支ID,Null,NVL(d.序号,a.序号),NVL(d.序号,a.序号)+NVL(E.序号,c.序号))>[6] "
                    Else
                        '理论天数(根据它来决定可选阶段)
                        If mPP.当前阶段ID <> 0 Then
                            lng理论天数 = GetMustDay(mPP.病人路径ID, lng天数, , lng合并路径记录ID)
                            
                            '之前可能有提前执行过的阶段的时间范围在当前天数内，要排除那些阶段。路径跳转时不检查
                            strIF = " And Decode(a.分支ID,Null,NVL(d.序号,a.序号),NVL(d.序号,a.序号)+NVL(E.序号,c.序号))>=[6] "
                        Else
                            lng理论天数 = lng天数
                        End If
                        
                         '同一天有多个阶段时，当前阶段及分支不能再用,如果是进入下一天了，则说明没有相同天数的阶段
                        If lng天数 = mPP.当前天数 Then
                            strIF = strIF & " And Nvl(a.父id,a.id) <> " & IIf(mPP.阶段父ID <> 0, "[8]", "[4]")
                        End If
                    End If
                    
                    '如果是分支路径，则加上前一阶段的序号
                    If lng当前阶段分支ID <> 0 Then
                        strSql = "Select 前一阶段ID From 临床路径分支 Where ID=[1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取分支路径前一阶段ID", lng当前阶段分支ID)
                        If rsTmp.RecordCount > 0 Then lng序号 = lng序号 + GetPhaseNO(Val(rsTmp!前一阶段ID & ""))
                    End If
                    
                    str阶段分类 = Get阶段分类(mPP.病人路径ID)
                    If str阶段分类 <> "" Then
                        strIF = strIF & " And (a.父id is Null Or a.父id is Not Null And a.分类 = [5])"
                    End If
                    '如果当前阶段已经是分支路径，则只能继续走当前分支,否则判断当前阶段是否存在分支
                    If lng当前阶段分支ID <> 0 Then
                        strIF = strIF & " And a.分支ID=[7]"
                    Else
                        strIF = strIF & " And (a.分支ID is Null or a.分支ID In(Select ID From 临床路径分支 B Where a.路径id=b.路径id and a.版本号=b.版本号 And b.前一阶段ID=[4]))"
                    End If
                    
                    strMainIF = strIF
                    
                    strIF = strIF & " And (a.开始天数 Is Null Or [3] Between a.开始天数 And Nvl(a.结束天数,a.开始天数) "
                    '合并路径提前只能提前一个阶段
                    If (mlng时间进度 = 1 Or mlng时间进度 = 2) And lng合并路径记录ID = 0 Then
                        strIF = strIF & " Or a.开始天数 >= [3])"
                    Else
                        strIF = strIF & ")"
                    End If
                End If
            End If
        Else
            strIF = " And a.id = [4]"
        End If
      
        strSql = "Select a.Id, Nvl(a.父id,0) as 父id, a.序号, a.名称, a.说明,a.开始天数, a.结束天数, a.分类,NVL(a.分支ID,0) AS 分支ID" & vbNewLine & _
                "From 临床路径阶段 A,临床路径分支 B,临床路径阶段 C,临床路径阶段 D,临床路径阶段 E " & strSubSQL & vbNewLine & _
                "Where a.分支id=b.id(+) and b.前一阶段id=c.id(+) And a.父ID=d.id(+) And c.父id=e.id(+) and a.路径id = [1] And a.版本号 = [2]" & _
                strIF & vbNewLine & " Order by NVL(d.序号,a.序号)"
 
    End If
    On Error GoTo errH
    Set GetPhase = zlDatabase.OpenSQLRecord(strSql, "获取适用阶段", lng路径ID, lng版本号, lng理论天数, lng当前阶段ID, str阶段分类, lng序号, lng当前阶段分支ID, mPP.阶段父ID, mPP.病人路径ID)
    
    If (mlng时间进度 = 1 Or mlng时间进度 = 2) And GetPhase.RecordCount = 0 Then
        '阶段提前时，如果当前阶段有多天，则按当前天数取不到下一阶段，直接取序号大于当前阶段的下一阶段
        strSql = "Select * From (Select a.Id, Nvl(a.父id,0) as 父id, a.序号, a.名称, a.说明,a.开始天数, a.结束天数, a.分类,NVL(a.分支ID,0) AS 分支ID" & vbNewLine & _
                "From 临床路径阶段 A,临床路径分支 B,临床路径阶段 C,临床路径阶段 D,临床路径阶段 E" & vbNewLine & _
                "Where a.分支id=b.id(+) and b.前一阶段id=c.id(+) And a.父ID=d.id(+)  And c.父id=e.id(+) and a.路径id = [1] And a.版本号 = [2]" & _
                strMainIF & vbNewLine & " Order by NVL(d.序号,a.序号)) Where Rownum=1"
        Set GetPhase = zlDatabase.OpenSQLRecord(strSql, "获取适用阶段", lng路径ID, lng版本号, lng理论天数, lng当前阶段ID, str阶段分类, lng序号, lng当前阶段分支ID)
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckMergeSend(ByVal lng合并路径记录ID As Long) As Boolean
'功能：判断合并路径是否是第一次生成（从评估信息里面查，因为执行里面可能未生成合并路径的项目）
    Dim strSql As String, rsTmp As Recordset
    
    On Error GoTo errH
    strSql = "Select Count(1) as 个数 From 病人合并路径评估 Where 合并路径记录ID=[1] And 路径记录ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng合并路径记录ID, mPP.病人路径ID)
    CheckMergeSend = rsTmp!个数 = 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdMergeStep_Click()
'功能：合并路径阶段选择
    Dim lngDay As Long, lngEOF As Long
    
    If frmPathMergeStep.ShowMe(mfrmParent, mrsMerge, mlngMergeCount, mstrMerge) = True Then
        If mrsMerge.RecordCount > 0 Then
            mrsMerge.MoveFirst
            vsItem(1).Rows = vsItem(1).FixedRows
            mstrMergeStep = ""
            '求出首要路径的当前天数和现在生成天数相差多少(首要路径延后或提前，合并路径也一样)
            lngDay = mlng天数 - mPP.当前天数
            Do While Not mrsMerge.EOF
                lngEOF = mrsMerge.AbsolutePosition
                Call LoadItem(Val(mrsMerge!ID & ""), vsItem(1), Val(mrsMerge!路径ID & ""), Val(mrsMerge!版本号 & ""), Val(mrsMerge!当前天数 & "") + lngDay, Val(mrsMerge!合并路径记录ID & ""))
                mrsMerge.AbsolutePosition = lngEOF
                mrsMerge.MoveNext
            Loop
        End If
    End If
End Sub

Private Sub cmdPati_Click()
    Dim i As Long
    Dim arrtmp As Variant
    Dim lngW As Long
    Dim lngH As Long
    Dim strTmp As String
    Dim strSelect As String
            
    lstPati.Visible = True
    lblFont.FontSize = lblPhase.FontSize
    With lstPati
        .Clear
        strTmp = "病人本人|" & mstrBaby
        arrtmp = Split(strTmp, "|")
        For i = LBound(arrtmp) To UBound(arrtmp)
            lblFont.Caption = arrtmp(i)
            If lngW < lblFont.Width Then lngW = lblFont.Width
           .AddItem arrtmp(i)
        Next
        lngH = (i - 1) * 210 + 240
        If lngH > 1080 Then lngH = 1080
        lngW = lngW + 700
        If lngW > 2500 Then lngW = 2500
    End With

    With vsItem(Val(cmdPati.Tag))
        strSelect = .TextMatrix(.Row, .Col)
        For i = 0 To lstPati.ListCount - 1
            If InStr("|" & strSelect & "|", "|" & lstPati.List(i) & "|") > 0 Then
                lstPati.Selected(i) = True
            End If
        Next
        If lngW < .ColWidth(mcol("婴儿")) Then lngW = .ColWidth(mcol("婴儿"))
        lstPati.Move .Left + .ColPos(.Col), .Top + .RowPos(.Row) + .RowHeight(.Row) + 30, lngW, lngH
    End With
    Call lstPati.SetFocus
End Sub

Private Sub Form_Load()
    
    If mlngFun <> 2 Then vsItem(0).Editable = flexEDKbdMouse: vsItem(1).Editable = flexEDKbdMouse
    
    Call LoadBranch
    Call LoadPhase
    
    mlng路径医嘱天数 = Val(zlDatabase.GetPara("路径医嘱生成超前天数", glngSys, p临床路径应用, "1"))
    '医嘱缺省时间默认取当前时间
    dtpAdviceTime.Value = mdatDur
    
    If vsPhase.Cols = 1 And tabBranch.Tabs.count = 1 Then
        vsPhase.Visible = False
        lblPhase.Caption = vsPhase.TextMatrix(0, 0) & vbCrLf & vsPhase.Cell(flexcpData, 0, 0)
                
        vsItem(0).Top = vsPhase.Top
        vsItem(0).Height = IIf(lblMerge.Visible, lblMerge.Top - 50, picBottom.Top) - vsItem(0).Top
    Else
        lblNote.Visible = False
        lblPhase.Left = lblNote.Left
    
        If Grid.HScrollVisible(vsPhase) Then
            '横向滚动条
            vsPhase.Height = 1000
            vsItem(0).Height = vsItem(0).Height - (vsPhase.Top + vsPhase.Height - vsItem(0).Top + 120)
            vsItem(0).Top = vsPhase.Top + vsPhase.Height + 60
        Else
            If vsPhase.Rows = 1 Then
                vsPhase.RowHeightMax = vsPhase.Height
                vsPhase.RowHeight(0) = vsPhase.Height
            End If
        End If
    End If
        
    If mlngFun = 2 Then
        Me.Caption = "查看阶段定义的项目"
        lblDate.Visible = False
        
        cmdOK.Visible = False
        cmdCancel.Caption = "退出(&X)"
        mstrBaby = ""
    Else
        lblDate.Caption = "生成路径项目日期：" & Format(mdat时间, "yyyy-MM-dd") & ",第" & mlng天数 & "天"
        
        If mlng补录间隔 > 0 Then
            lblDate.Caption = lblDate.Caption & "(" & mlng补录间隔 & "天前)"
            lblDate.ForeColor = vbRed
        End If
        mstrBaby = GetBabyRegList
    End If
    If mlngFun <> 0 Then cmdMergeStep.Visible = False
    
    Call InitItem
    
    If mlngFun = 2 Then '查看时只显示分类 , 项目内容
        Me.Width = vsItem(0).Width + 360
        cmdCancel.Left = vsItem(0).Left + vsItem(0).Width - 1200
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 150
    End If
    
    Set mEditType = New Collection
    Call LoadMerge
    Call LoadItem(Val(vsPhase.ColData(vsPhase.Col)), vsItem(0), mPP.路径ID, mPP.版本号, mlng天数)
    
   
    If vsItem(0).Rows = 1 And vsItem(1).Rows = 1 Then
        vsItem(0).Rows = 2
        vsItem(1).Rows = 2
        vsItem(0).TextMatrix(1, mcol("项目内容")) = "没有至少执行一次或可选性的路径项目"
        cmdOK.Enabled = False
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsPhase = Nothing
    Set mcol = Nothing
    Set mEditType = Nothing
    mstrMerge = ""
    Set mclsMipModule = Nothing
End Sub

Private Sub lstPati_ItemCheck(Item As Integer)
    Dim i As Long
    Dim strList As String
    strList = ""
    For i = 0 To lstPati.ListCount - 1
        If lstPati.Selected(i) Then
            strList = strList & "|" & lstPati.List(i)
        End If
    Next
    If strList <> "" Then
        strList = Mid(strList, 2)
    Else
        strList = lstPati.List(0)  '缺省选中病人本人,不允许为空
    End If
    
    vsItem(Val(cmdPati.Tag)).TextMatrix(vsItem(Val(cmdPati.Tag)).Row, vsItem(Val(cmdPati.Tag)).Col) = strList
End Sub

Private Sub lstPati_KeyPress(KeyAscii As Integer)
    If lstPati.Visible Then
        If KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Then
            lstPati.Visible = False
        End If
    End If
End Sub

Private Sub lstPati_LostFocus()
    lstPati.Visible = False
End Sub

Private Sub tabBranch_Click()
    Call LoadPhase
    If vsPhase.Cols = 1 And tabBranch.Tabs.count = 1 Then
        vsPhase.Visible = False
        lblPhase.Caption = vsPhase.TextMatrix(0, 0) & vbCrLf & vsPhase.Cell(flexcpData, 0, 0)
        vsItem(0).Top = vsPhase.Top
        vsItem(0).Height = IIf(lblMerge.Visible, lblMerge.Top - 50, picBottom.Top) - vsItem(0).Top
    Else
        vsPhase.Visible = True
        vsItem(0).Top = vsPhase.Top + vsPhase.Height + 45
        vsItem(0).Height = IIf(lblMerge.Visible, lblMerge.Top - 50, picBottom.Top) - vsItem(0).Top
    End If
End Sub

Private Sub vsItem_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    With vsItem(Index)
        If cmdPati.Visible Then
            cmdPati.Move .Left + .ColPos(.Col) + .ColWidth(.Col) - 255, .Top + .RowPos(.Row) + 15, 255, 240
            lstPati.Visible = False
        End If
    End With
End Sub

Private Sub vsItem_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    With vsItem(Index)
        If .Col = mcol("婴儿") And cmdPati.Visible Then
            cmdPati.Move .Left + .ColPos(.Col) + .ColWidth(.Col) - 255, .Top + .RowPos(.Row) + 15, 255, 240
            lstPati.Visible = False
        End If
    End With
End Sub

Private Sub vsItem_Click(Index As Integer)
    With vsItem(Index)
        If lstPati.Visible Then lstPati.Visible = False
    End With
End Sub

Private Sub vsItem_DblClick(Index As Integer)
    Dim lng项目ID As Long
    
    If vsItem(Index).Col = mcol("项目内容") Then
        lng项目ID = Val(vsItem(Index).TextMatrix(vsItem(Index).Row, mcol("ID")))
        If lng项目ID <> 0 Then
            Call frmPathItemEdit.ShowView(mfrmParent, lng项目ID)
        End If
    End If
    
End Sub

Private Sub vsItem_GotFocus(Index As Integer)
    vsItem(Index).ForeColorSel = vbWhite
    vsItem(Index).BackColorSel = &H8000000D
End Sub

Private Sub vsItem_LostFocus(Index As Integer)
    vsItem(Index).ForeColorSel = vbBlack
    vsItem(Index).BackColorSel = vbWhite
End Sub

Private Sub vsItem_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call ResultEnterNextCell(vsItem(Index))
    End If
End Sub

Private Sub ResultEnterNextCell(vsthis As VSFlexGrid)
    With vsthis
        If .Col <= .Cols - 1 Then
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
                .ShowCell .Row, .Col
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub

Private Sub vsItem_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    
    With vsItem(Index)
        If Col = mcol("重选") Then
            If .Cell(flexcpChecked, Row, mcol("重选")) = 1 Then
                If .Cell(flexcpChecked, Row, mcol("选择")) <> 1 Then .Cell(flexcpChecked, Row, mcol("选择")) = 1
            End If
        ElseIf Col = mcol("选择") Then
            If .Cell(flexcpChecked, Row, mcol("选择")) = 2 Then
                '未选择时取消重选
                If .Cell(flexcpChecked, Row, mcol("重选")) = 1 Then .Cell(flexcpChecked, Row, mcol("重选")) = 2
                
                '未选择时弹出变异原因选择
                If mlngFun = 0 Then
                    If .RowData(Row) = 执行方式.T1每天必须 Then
                        Call vsItem_CellButtonClick(Index, Row, mcol("变异原因"))
                    End If
                End If
                
            ElseIf .Cell(flexcpChecked, Row, mcol("选择")) = 1 Then
                If .RowData(Row) = 执行方式.T1每天必须 Then
                '选择时，清除变异原因
                    .TextMatrix(Row, mcol("变异原因")) = ""
                    .Cell(flexcpData, Row, mcol("变异原因")) = ""
                End If
            End If
        ElseIf Col = mcol("全选") Then
            If .Cell(flexcpChecked, Row, mcol("全选")) = 2 Then
                For i = Row To .Rows - 1
                    '每天生成的取消需要填写原因，所以不取消勾选，要取消只能一个一个取消
                    If .TextMatrix(i, mcol("分类")) <> .TextMatrix(Row, mcol("分类")) Then Exit For
                    If Not (.RowData(i) = 执行方式.T0无需执行 Or .RowData(i) = 执行方式.T1每天必须) Then
                        If .Cell(flexcpChecked, i, mcol("选择")) = 1 Then .Cell(flexcpChecked, i, mcol("选择")) = 2
                    End If
                Next
                For i = Row - 1 To .FixedRows Step -1
                    If .TextMatrix(i, mcol("分类")) <> .TextMatrix(Row, mcol("分类")) Then Exit For
                    If Not (.RowData(i) = 执行方式.T0无需执行 Or .RowData(i) = 执行方式.T1每天必须) Then
                        If .Cell(flexcpChecked, i, mcol("选择")) = 1 Then .Cell(flexcpChecked, i, mcol("选择")) = 2
                    End If
                Next
                
            ElseIf .Cell(flexcpChecked, Row, mcol("全选")) = 1 Then
                For i = Row To .Rows - 1
                    If .TextMatrix(i, mcol("分类")) <> .TextMatrix(Row, mcol("分类")) Then Exit For
                    If .Cell(flexcpChecked, i, mcol("选择")) = 2 Then
                        .Cell(flexcpChecked, i, mcol("选择")) = 1
                        If .RowData(i) = 执行方式.T1每天必须 Then
                        '选择时，清除变异原因
                            .TextMatrix(i, mcol("变异原因")) = ""
                            .Cell(flexcpData, i, mcol("变异原因")) = ""
                        End If
                    End If
                Next
                For i = Row - 1 To .FixedRows Step -1
                    If .TextMatrix(i, mcol("分类")) <> .TextMatrix(Row, mcol("分类")) Then Exit For
                    If .Cell(flexcpChecked, i, mcol("选择")) = 2 Then
                        .Cell(flexcpChecked, i, mcol("选择")) = 1
                        If .RowData(i) = 执行方式.T1每天必须 Then
                        '选择时，清除变异原因
                            .TextMatrix(i, mcol("变异原因")) = ""
                            .Cell(flexcpData, i, mcol("变异原因")) = ""
                        End If
                    End If
                Next
            End If
        End If
    End With
End Sub


Private Sub vsItem_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsItem(Index)
        If NewRow >= .FixedRows And Me.Visible Then
            If mlngFun = 0 Then
                If NewCol = mcol("变异原因") Then
                    '未选择时，可设置或选择变异原因
                    If .RowData(NewRow) = 执行方式.T1每天必须 And .Cell(flexcpChecked, NewRow, mcol("选择")) = 2 Then
                        .ColComboList(mcol("变异原因")) = "..."
                    Else
                        .ColComboList(mcol("变异原因")) = ""
                    End If
                End If
            End If
            cmdPati.Visible = False: lstPati.Visible = False
            If NewCol = mcol("婴儿") And mlngFun <> 2 Then
                If mstrBaby <> "" Then
                    cmdPati.Visible = True
                    cmdPati.Enabled = True
                    cmdPati.Move .Left + .ColPos(NewCol) + .ColWidth(NewCol) - 255, .Top + .RowPos(NewRow) + 15, 255, 240
                    cmdPati.Tag = Index
                    If .RowData(NewRow) = 执行方式.T0无需执行 Then cmdPati.Enabled = False
                End If
            End If
        End If
    End With
End Sub

Private Sub vsItem_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsItem(Index)
        If Col = mcol("选择") Then
            '生成路径时，每天生成的，可以不选，但要求输变异原因
            If .RowData(Row) = 执行方式.T0无需执行 Or mlngFun <> 0 And .RowData(Row) = 执行方式.T1每天必须 Then
                Cancel = True
            End If
        ElseIf Col = mcol("重选") Then
            If Val(.Cell(flexcpChecked, Row, Col)) = 0 Then Cancel = True
        ElseIf Col = mcol("全选") Then
            If Val(.Cell(flexcpChecked, Row, Col)) = 0 Then Cancel = True
        ElseIf Col = mcol("婴儿") Then
            Cancel = True
        ElseIf Col = mcol("变异原因") Then
            If .ColComboList(mcol("变异原因")) = "" Then Cancel = True
        Else
            Cancel = True
        End If
    End With
End Sub

'暂不支持变异原因的输入
'Private Sub vsItem_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'    If Col = mcol("变异原因") Then
'        Dim rsTmp As ADODB.Recordset, strSQL As String, strInput As String
'        Dim vPoint As POINTAPI, blnCancel As Boolean
'
'        vPoint = GetCoordPos(vsItem(0).EditWindow, vsItem(0).CellTop, vsItem(0).CellLeft)
'        strInput = gstrLike & vsItem(0).EditText & "%"
'        strSQL = "Select b.名称 As 分类, a.编码, a.名称, a.简码" & vbNewLine & _
'                "From 变异常见原因 A, 变异常见原因 B" & vbNewLine & _
'                "Where a.末级 = 1 And a.上级 = b.编码 and a.性质=1 And (a.名称 like [1] or 简码 like [1] or 编码 like [1]" & vbNewLine & _
'                "order by b.名称"
'        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "变异常见原因", True, False, "请选择", False, False, False, vPoint.X, vPoint.Y, vsItem(0).EditWindow, blnCancel, False, True, strInput)
'        If rsTmp Is Nothing Then
'            If Not blnCancel Then
'                Cancel = True
'                MsgBox "系统没有初始变异常见原因，请与系统管理员联系。", vbInformation, gstrSysName
'                Exit Sub
'            End If
'        Else
'            vsItem(0).TextMatrix(Row, Col) = rsTmp!名称
'            vsItem(0).Cell(flexcpData, Row, Col) = CStr(rsTmp!编码)
'        End If
'    End If
'End Sub

Private Sub vsItem_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    With vsItem(Index)
        If Col = mcol("变异原因") Then
            Dim strSql As String, blnCancel As Boolean
            Dim rsTmp As ADODB.Recordset
                    
            strSql = "Select b.名称 as 分类,a.编码 as ID,a.编码,a.名称,a.简码 From 变异常见原因 a,变异常见原因 b" & _
                    " Where a.性质=1 And a.末级=1 And a.上级=b.编码 And b.末级=0 " & _
                    " Order by 分类,a.编码"
            
            Set rsTmp = zlDatabase.ShowSelect(Me, strSql, 0, "变异常见原因", True, , , True, True, True, _
                     Me.Left + .Left + .ColPos(Col), Me.Top + .Top + .RowPos(Row) + .RowHeight(Row) * 2, .RowHeight(Row), blnCancel, False, True)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "系统没有初始变异常见原因，请与系统管理员联系。", vbInformation, gstrSysName
                    Exit Sub
                End If
            Else
                .TextMatrix(Row, Col) = rsTmp!名称
                .Cell(flexcpData, Row, Col) = CStr(rsTmp!编码)
            End If
        End If
    End With
End Sub

Private Sub vsPhase_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lng阶段ID As Long
    Dim str天数 As String

    If OldCol <> NewCol And Me.Visible = True And NewCol >= 0 And NewRow >= 0 And vsPhase.Redraw = flexRDDirect Then
        lng阶段ID = vsPhase.ColData(NewCol)
        
        If mint场合 = 1 And mlngFun = 0 Then
            str天数 = vsPhase.TextMatrix(1, NewCol)  '格式：日期（第n天）
            str天数 = Val(Mid(str天数, InStr(str天数, "(第") + 2))
            mrsPhase.Filter = "ID=" & lng阶段ID & " and 天数= " & str天数
            mdat时间 = CDate(mrsPhase!日期 & "")
            mlng天数 = Val(mrsPhase!天数 & "")
            lblDate.Caption = "生成路径项目日期：" & Format(mdat时间, "yyyy-MM-dd") & ",第" & mlng天数 & "天"
            mlng提前天数 = mlng天数
        Else
            mrsPhase.Filter = "ID=" & lng阶段ID
            mlng天数 = (mrsPhase!开始天数 & "")
        End If
        Call LoadItem(lng阶段ID, vsItem(0), mPP.路径ID, mPP.版本号, mlng天数)
                
    End If
End Sub

Private Sub LoadBranch()
'功能：加载分支路径
    Dim i As Long, j As Long, strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    If Not (mint场合 = 1 And mlngFun = 0) Then
        If mrsPhase.RecordCount > 0 Then
            Do While Not mrsPhase.EOF
                If Val(mrsPhase!分支ID & "") <> 0 Then
                    strTmp = strTmp & "," & mrsPhase!分支ID
                End If
                mrsPhase.MoveNext
            Loop
            strTmp = strTmp & ","
            mrsPhase.MoveFirst
        End If
    
        strSql = "Select ID,名称 From 临床路径分支 Where 前一阶段ID=[3] And 路径ID=[1] And 版本号=[2]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取分支信息", mPP.路径ID, mPP.版本号, mPP.当前阶段ID)
        If rsTmp.RecordCount > 0 Then
            Do While Not rsTmp.EOF
                '可能分支路径的阶段中没有适合当前天数的阶段
                If InStr(strTmp, "," & rsTmp!ID & ",") > 0 Then
                    tabBranch.Tabs.Add , "_" & rsTmp!ID, "分支:" & rsTmp!名称
                End If
                rsTmp.MoveNext
            Loop
            mblnIsHaveBranch = True
        End If
    End If
    If tabBranch.Tabs.count = 1 Then
        tabBranch.Visible = False
        mblnIsHaveBranch = False
        vsPhase.Top = tabBranch.Top
        vsItem(0).Top = vsPhase.Top + vsPhase.Height + 45
        vsItem(0).Height = vsItem(0).Height + tabBranch.Height - 15
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadPhase()
'功能：加载可选择的阶段,如果病人的当前时间阶段仍然可用，则选中，否则缺省为第一个
    Dim i As Long, j As Long, str阶段分类 As String
    Dim rsNode As ADODB.Recordset

    With vsPhase
        .Clear
        .Redraw = flexRDNone
        .Col = -1
        If mint场合 = 1 And mlngFun = 0 Then
            mrsPhase.Filter = ""  '护士场合生成路径
            .Cols = mrsPhase.RecordCount
        Else
            mrsPhase.Filter = IIf(mblnIsHaveBranch, "分支ID=" & Mid(tabBranch.SelectedItem.Key, 2), "")
            .Cols = mrsPhase.RecordCount
            str阶段分类 = Get阶段分类(0, mPP.当前阶段ID)
            If mlngFun = 0 And mlng时间进度 <> -1 Then '补充生成、重新生成时、下一阶段延后（继续当前阶段），只有当前阶段的记录
                mrsPhase.Filter = "父ID<>0 " & IIf(mblnIsHaveBranch, " And 分支ID=" & Mid(tabBranch.SelectedItem.Key, 2), "")
                If mrsPhase.RecordCount > 0 Then    '有备用分支
                    Set rsNode = mrsPhase.Clone
                    .Rows = 2
                    .MergeRow(0) = True
                Else
                    .Rows = 1
                End If
                mrsPhase.Filter = "父ID=0" & IIf(mblnIsHaveBranch, " And 分支ID=" & Mid(tabBranch.SelectedItem.Key, 2), "")
            End If
        End If
        
        For i = 0 To .Cols - 1
            .ColWidth(i) = 2000
            .ColAlignment(i) = flexAlignCenterCenter
            .TextMatrix(0, i) = mrsPhase!名称
            .Cell(flexcpData, 0, i) = CStr(IIf(IsNull(mrsPhase!分类), "", "分类：" & mrsPhase!分类 & " ") & mrsPhase!说明)
            .ColData(i) = Val(mrsPhase!ID)
            
            
            If mint场合 = 1 And mlngFun = 0 Then
                If .ColData(i) & "_" & mrsPhase!天数 = mlng病人阶段ID & "_" & mlng天数 Then .Col = i
                .MergeCol(i) = True
                .TextMatrix(1, i) = mrsPhase!日期 & " " & "(第" & mrsPhase!天数 & "天)"
            Else
                If .ColData(i) = mlng病人阶段ID Then .Col = i
                If Not rsNode Is Nothing Then
                    rsNode.Filter = "父ID=" & mrsPhase!ID
                    If rsNode.RecordCount = 0 Then
                         .MergeCol(i) = True
                         .TextMatrix(1, i) = mrsPhase!名称
                    Else
                         .TextMatrix(1, i) = "缺省"
                         .ColWidth(i) = 1000
                        For j = 1 To rsNode.RecordCount
                            i = i + 1
                             .ColWidth(i) = 1000
                             .ColAlignment(i) = flexAlignCenterCenter
                            .TextMatrix(0, i) = mrsPhase!名称 '第一行设置相同内容用于合并
                            .TextMatrix(1, i) = IIf(IsNull(rsNode!说明), "分支" & j, "" & rsNode!说明)
                            .Cell(flexcpData, 1, i) = CStr(IIf(IsNull(rsNode!分类), "", "分类：" & rsNode!分类 & " ") & rsNode!说明)
                            
                            .ColData(i) = Val(rsNode!ID)
                            If .ColData(i) = mlng病人阶段ID Then
                                .Col = i
                            ElseIf .Col = 0 And str阶段分类 <> "" Then
                                If str阶段分类 = "" & rsNode!分类 Then .Col = i
                            End If
                            rsNode.MoveNext
                        Next
                    End If
                End If
            End If
            mrsPhase.MoveNext
        Next
        If .Col < 0 Then .Col = 0
        mrsPhase.Filter = "ID=" & Val(.ColData(.Col))
        .Redraw = True
        vsPhase_AfterRowColChange -1, -1, .Row, .Col
    End With
End Sub

Private Sub LoadItem(lng阶段ID As Long, objVsg As VSFlexGrid, ByVal lng路径ID As Long, ByVal lng版本号 As Long, ByVal lng天数 As Long, Optional ByVal lng合并路径记录ID As Long)
'功能：加载当前阶段的路径项目
'参数：objVsg，需要加载的表格（首要路径或合并路径）,如果是加载合并路径，则在后面添加，不清空列表
    Dim i As Long, j As Long, blnFocus As Boolean, bln长嘱 As Boolean
    Dim rsTmp As ADODB.Recordset, strSql As String, strIDs As String, strTmp As String
    Dim str可选长嘱 As String, lngOld相关ID As Long, bln多组长嘱 As Boolean
    Dim rsAdvice As ADODB.Recordset, rsFile As ADODB.Recordset
    Dim lngRow As Long
    Dim strFilter As String, blnEnd As Boolean
    Dim str诊疗项目IDs As String
    Dim lng首要路径阶段ID As Long
    Dim strNewTmp As String
     
    If mlngFun = 1 Then '补充生成，无需执行的不显示，当天执行过的不能重复生成,只执行一次的当前阶段已执行则不显示
        strSql = " And a.执行方式<>0 And Not Exists(Select 1 From 病人路径执行 c " & _
                "Where c.路径记录id = [4] And c.阶段id = [7] And c.项目id = a.id And (c.天数 = [5] and a.执行方式<>4 or a.执行方式=4))"
        lng首要路径阶段ID = mPP.当前阶段ID
    ElseIf mlngFun = 3 Then '重新生成
        strSql = " And a.ID = [6]"
    Else
        strSql = " And (a.执行方式<>4 or a.执行方式=4 And Not Exists(Select 1 From 病人路径执行 c " & _
                "Where c.路径记录id = [4] And c.阶段id = [7] And c.项目id = a.id))"
        If objVsg.Index = 0 Then
            lng首要路径阶段ID = lng阶段ID
        Else
            lng首要路径阶段ID = Val(vsPhase.ColData(vsPhase.Col))
        End If
    End If
    '连接“临床路径分类”，只是为了按分类排序'保存时再检查，是否为当天阶段的最后一天，至少执行一次的项目是否选择
    strSql = "Select a.分类, a.ID, a.项目内容, a.执行方式, a.图标id, a.内容要求" & vbNewLine & _
        "From 临床路径项目 A, 临床路径分类 B" & vbNewLine & _
        "Where a.分类 = b.名称 And a.路径id = b.路径id And a.版本号 = b.版本号 And a.路径id = [1] And a.版本号 = [2] And a.阶段id = [3] And NVL(a.分支ID,0)=nvl(b.分支id,0)" & vbNewLine & _
        Decode(mint场合, 0, " And NVL(a.生成者,1) = 1 ", 1, " And a.生成者 = 2 ") & vbNewLine & _
        strSql & vbNewLine & _
        IIf(mlngFun = 3, "", "Order By b.序号, a.项目序号")
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng路径ID, lng版本号, lng阶段ID, mPP.病人路径ID, mlng天数, mlng项目ID, lng首要路径阶段ID)
    
    With objVsg
        .Redraw = flexRDNone
        If objVsg.Index = 0 Then
            .Rows = .FixedRows
            .Rows = .FixedRows + rsTmp.RecordCount
            lngRow = 1
        Else
            lngRow = .Rows
            If lngRow = 2 Then
                If .TextMatrix(1, mcol("ID")) = "" Then lngRow = 1
            End If
            .Rows = lngRow + rsTmp.RecordCount
        End If
        '由于固定列合并不能影响后面的列，所以新增了一列判断按分类合并全选列
        .MergeCells = flexMergeRestrictAll
        .MergeCol(mcol("分类")) = True
        .MergeCol(mcol("分类值")) = True
        .MergeCol(mcol("全选")) = True
        '判断是否是最后一天，如果是合并路径表格才计算
        If objVsg.Index = 1 Then
            strFilter = mrsMerge.Filter
            If lng天数 = 0 Then lng天数 = 1
            mrsMerge.Filter = "ID=" & lng阶段ID
            If mrsMerge.RecordCount > 0 Then
                With mrsMerge
                    If Not IsNull(!开始天数) Then
                        If IsNull(!结束天数) Then
                            blnEnd = (Val(!开始天数) = lng天数)
                        Else
                            blnEnd = (Val(!结束天数) = lng天数)
                        End If
                    End If
                End With
                mrsMerge.Filter = IIf(strFilter = "0", 0, strFilter)
                '记录下加载的合并路径阶段
                mstrMergeStep = mstrMergeStep & "," & lng合并路径记录ID & ":" & lng阶段ID
            End If
        End If
        For i = lngRow To rsTmp.RecordCount + lngRow - 1
            .TextMatrix(i, mcol("ID")) = rsTmp!ID
            strIDs = strIDs & "," & rsTmp!ID
            .TextMatrix(i, mcol("分类")) = rsTmp!分类
            .TextMatrix(i, mcol("分类值")) = rsTmp!分类
            .TextMatrix(i, mcol("项目内容")) = rsTmp!项目内容
            If mlngFun <> 2 Then .TextMatrix(i, mcol("内容要求")) = Val("" & rsTmp!内容要求)
            .TextMatrix(i, mcol("执行方式")) = Decode(rsTmp!执行方式, 0, "无", 1, "每天", 2, "至少一次", 3, "必要时", 4, "必须一次")
            .RowData(i) = Val(rsTmp!执行方式)
            If objVsg.Index = 1 And blnEnd And mlngFun <> 2 Then
                .TextMatrix(i, mcol("是否最后一天")) = "1"
                .TextMatrix(i, mcol("阶段ID")) = lng阶段ID
                .TextMatrix(i, mcol("合并路径记录ID")) = lng合并路径记录ID
            End If
            
            If mlngFun <> 2 Then
                Select Case rsTmp!执行方式
                    Case 执行方式.T0无需执行
                        .TextMatrix(i, mcol("选择")) = " "
                        .Cell(flexcpBackColor, i, mcol("选择")) = &H8000000F
                    Case 执行方式.T1每天必须
                        .Cell(flexcpChecked, i, mcol("选择")) = 1
                        .Cell(flexcpPictureAlignment, i, mcol("选择")) = flexPicAlignCenterCenter
                        If mlngFun <> 0 Then .Cell(flexcpBackColor, i, mcol("选择")) = &H8000000F
                        '生成时，不设置为灰色背景，因为可以不选择生成（须录入变异原因）
                   Case Else
                        If mlngFun = 3 Then '重选时，选择列未显示，自动勾上
                            .Cell(flexcpChecked, i, mcol("选择")) = 1
                        Else
                            .Cell(flexcpChecked, i, mcol("选择")) = IIf(rsTmp.RecordCount = 1, 1, 2)
                            .Cell(flexcpPictureAlignment, i, mcol("选择")) = flexPicAlignCenterCenter
                        End If
                End Select
            End If
            
                        
            If Not IsNull(rsTmp!图标ID) Then
                Call .Select(i, mcol("项目内容"))
                .CellPictureAlignment = flexPicAlignRightCenter 'flexPicAlignLeftCenter
                .CellPicture = GetPathIcon(rsTmp!图标ID)
            End If
            
            If mstrBaby <> "" Then
                .TextMatrix(i, mcol("婴儿")) = "病人本人"
            End If
            
            If (rsTmp!执行方式 = 执行方式.T3必要时) And blnFocus = False Then
                Call .Select(i, mcol("选择"))
                blnFocus = True
            End If
            rsTmp.MoveNext
        Next
        
        strIDs = Mid(strIDs, 2)
        '加载项目对应的医嘱
        str可选长嘱 = ""
        Set rsAdvice = GetAdvice(strIDs)
        If rsAdvice.RecordCount > 0 Then
            For i = .FixedRows To .Rows - 1
                rsAdvice.Filter = "路径项目ID=" & ZVal(Val(.TextMatrix(i, mcol("ID"))))
                strTmp = "": bln长嘱 = False: bln多组长嘱 = False: lngOld相关ID = 0: str诊疗项目IDs = ""
                For j = 1 To rsAdvice.RecordCount
                    strTmp = strTmp & "," & rsAdvice!医嘱内容ID
                    str诊疗项目IDs = str诊疗项目IDs & "," & rsAdvice!诊疗项目ID
                    If rsAdvice!期效 = 0 Then
                        bln长嘱 = True    '一般情况，同一项目的医嘱期效相同,如果有混用的情况，只要有长嘱都算
                        
                        If mlngFun <> 2 Then
                            If j > 1 And bln多组长嘱 = False Then
                                If lngOld相关ID <> Val(rsAdvice!相关id) Then bln多组长嘱 = True
                            End If
                            lngOld相关ID = rsAdvice!相关id
                        End If
                    End If
                    rsAdvice.MoveNext
                Next
                If strTmp <> "" Then
                    If bln多组长嘱 Then
                        If .TextMatrix(i, mcol("内容要求")) = "1" Then str可选长嘱 = str可选长嘱 & "," & .TextMatrix(i, mcol("ID"))
                    End If
                    .TextMatrix(i, mcol("医嘱内容ID")) = Mid(strTmp, 2)
                    If mlngFun <> 2 Then .TextMatrix(i, mcol("诊疗项目ID")) = Mid(str诊疗项目IDs, 2)
                    .TextMatrix(i, mcol("项目内容")) = .TextMatrix(i, mcol("项目内容")) & " ……"
                    .TextMatrix(i, mcol("长嘱")) = IIf(bln长嘱, 1, 0)
                    If bln长嘱 Then .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HC00000      '蓝色
                    
                    '前次已产生的长嘱自动勾选
                    If bln长嘱 And Not (.RowData(i) = 执行方式.T1每天必须 Or .RowData(i) = 执行方式.T0无需执行) Then
                        Set rsTmp = GetLastAdvice(.TextMatrix(i, mcol("ID")))
                        If rsTmp.RecordCount > 0 Then
                            .Cell(flexcpChecked, i, mcol("选择")) = 1
                        End If
                    End If
                End If
            Next
        End If
        
        '对于上次已生成的，存在多个可选医嘱的，设置该行为可选择状态
        If mlngFun <> 2 Then
            If str可选长嘱 <> "" Then
                bln多组长嘱 = False
                Set rsAdvice = GetLastAdvice(str可选长嘱)
                For i = .FixedRows To .Rows - 1
                    rsAdvice.Filter = "项目id=" & .TextMatrix(i, mcol("ID"))
                    If rsAdvice.RecordCount = 0 Then
                        .TextMatrix(i, mcol("重选")) = " "
                        .Cell(flexcpBackColor, i, mcol("重选")) = &H8000000F
                    Else
                        .Cell(flexcpChecked, i, mcol("重选")) = IIf(mlngFun = 3, 1, 2)  '重新生成医嘱，重选的项目，自动勾上
                        .Cell(flexcpPictureAlignment, i, mcol("重选")) = flexPicAlignCenterCenter
                        If mlngFun = 3 Then
                            .Cell(flexcpBackColor, i, mcol("重选")) = &H8000000F
                        Else
                            .Editable = flexEDKbdMouse
                        End If
                        If bln多组长嘱 = False Then bln多组长嘱 = True
                    End If
                Next
                '没有一行记录可重选时，隐藏该列
                If bln多组长嘱 = False Then
                    .ColHidden(mcol("重选")) = True
                Else
                    If .ColHidden(mcol("重选")) Then .ColHidden(mcol("重选")) = False
                End If
            Else
                .ColHidden(mcol("重选")) = True
            End If
        End If
        
        '加载项目对应的病历文件
        If mlngFun <> 3 Then
            Set rsFile = GetFile(strIDs)
            If rsFile.RecordCount > 0 Then
                strIDs = ""
                For i = .FixedRows To .Rows - .FixedRows
                    rsFile.Filter = "路径项目ID=" & ZVal(Val(.TextMatrix(i, mcol("ID"))))
                    strTmp = "": strNewTmp = "" '记录新版电子病历ID
                    For j = 1 To rsFile.RecordCount
                        If rsFile!文件ID & "" <> "" Then
                            strTmp = strTmp & "," & rsFile!文件ID
                            If InStr(strIDs & ",", "," & rsFile!文件ID & ",") = 0 Then
                                On Error Resume Next
                                mEditType.Add Val(rsFile!保留), "C" & rsFile!文件ID
                                On Error GoTo errH
                            End If
                        Else
                            strNewTmp = strNewTmp & "," & rsFile!原型ID
                        End If
                        rsFile.MoveNext
                    Next
                    If strTmp <> "" Or strNewTmp <> "" Then
                        strIDs = strIDs & strTmp    '同一个路径项目的文件ID不会重，所以放到第二层循环中
                        .TextMatrix(i, mcol("文件ID")) = IIf(strTmp <> "", Mid(strTmp, 2), "") & "|" & IIf(strNewTmp <> "", Mid(strNewTmp, 2), "")
                        .TextMatrix(i, mcol("项目内容")) = .TextMatrix(i, mcol("项目内容")) & " …"
                    End If
                    
                Next
            End If
        End If
        If .Rows = .FixedRows Then .Rows = .Rows + 1
        .Redraw = True
    End With
               
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitMergeRs()
    If Not mrsMerge Is Nothing Then
        If mrsMerge.State = 1 Then mrsMerge.Close
    End If
    Set mrsMerge = New ADODB.Recordset
    
    mrsMerge.Fields.Append "ID", adBigInt
    mrsMerge.Fields.Append "父id", adBigInt, , adFldIsNullable
    mrsMerge.Fields.Append "序号", adBigInt, , adFldIsNullable
    mrsMerge.Fields.Append "名称", adVarChar, 100, adFldIsNullable
    mrsMerge.Fields.Append "说明", adVarChar, 200, adFldIsNullable
    mrsMerge.Fields.Append "开始天数", adBigInt, , adFldIsNullable
    mrsMerge.Fields.Append "结束天数", adBigInt, , adFldIsNullable
    mrsMerge.Fields.Append "分类", adVarChar, 50, adFldIsNullable
    mrsMerge.Fields.Append "分支ID", adBigInt, , adFldIsNullable
    mrsMerge.Fields.Append "路径ID", adBigInt, , adFldIsNullable
    mrsMerge.Fields.Append "版本号", adBigInt, , adFldIsNullable
    mrsMerge.Fields.Append "路径名称", adVarChar, 200, adFldIsNullable
    mrsMerge.Fields.Append "当前阶段ID", adBigInt, , adFldIsNullable
    mrsMerge.Fields.Append "当前天数", adBigInt, , adFldIsNullable
    mrsMerge.Fields.Append "合并路径记录ID", adBigInt, , adFldIsNullable
    
    mrsMerge.CursorLocation = adUseClient
    mrsMerge.LockType = adLockOptimistic
    mrsMerge.CursorType = adOpenStatic
    mrsMerge.Open
End Sub

Private Sub LoadMerge()
'功能：加载合并路径项目
    Dim i As Long
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim strTmp As String
    Dim rsMerge As ADODB.Recordset
    Dim lngDay As Long
    
    strSql = "Select a.id,b.名称,a.版本号,b.ID as 路径ID,NVL(a.当前天数,0) as 当前天数,a.当前阶段ID,c.分支ID as 当前阶段分支ID " & _
            " From 病人合并路径 A,临床路径目录 B,临床路径阶段 C " & _
            " Where a.路径ID=b.id And a.当前阶段ID = c.ID(+) And a.结束时间 is null And a.首要路径记录ID=[1] order by a.导入时间"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
    mlngMergeCount = rsTmp.RecordCount
    Call InitMergeRs
    vsItem(1).Rows = vsItem(1).FixedRows
    mstrMergeStep = ""
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            strTmp = strTmp & "," & rsTmp!名称
            '求出首要路径的当前天数和现在生成天数相差多少(首要路径延后或提前，合并路径也一样)
            lngDay = mlng天数 - mPP.当前天数
            '获取合并路径阶段
            Set rsMerge = GetPhase(Val(rsTmp!路径ID & ""), Val(rsTmp!版本号 & ""), Val(rsTmp!当前阶段ID & ""), Val(rsTmp!当前阶段分支ID & ""), Val(rsTmp!当前天数 & "") + lngDay, Val(rsTmp!ID & ""))
            If rsMerge.RecordCount > 0 Then
                Do While Not rsMerge.EOF
                    mrsMerge.AddNew
                    mrsMerge!ID = rsMerge!ID
                    mrsMerge!父ID = rsMerge!父ID
                    mrsMerge!序号 = rsMerge!序号
                    mrsMerge!名称 = rsMerge!名称
                    mrsMerge!说明 = rsMerge!说明
                    mrsMerge!开始天数 = rsMerge!开始天数
                    mrsMerge!结束天数 = rsMerge!结束天数
                    mrsMerge!分类 = rsMerge!分类
                    mrsMerge!分支ID = rsMerge!分支ID
                    mrsMerge!路径ID = rsTmp!路径ID
                    mrsMerge!版本号 = rsTmp!版本号
                    mrsMerge!路径名称 = rsTmp!名称
                    mrsMerge!当前阶段ID = rsTmp!当前阶段ID
                    mrsMerge!当前天数 = rsTmp!当前天数
                    mrsMerge!合并路径记录ID = rsTmp!ID
                    mrsMerge.Update
                    rsMerge.MoveNext
                Loop
                rsMerge.MoveFirst
                Call LoadItem(Val(rsMerge!ID & ""), vsItem(1), Val(rsTmp!路径ID & ""), Val(rsTmp!版本号 & ""), Val(rsTmp!当前天数 & "") + lngDay, Val(rsTmp!ID & ""))
                mstrMerge = mstrMerge & "," & rsMerge!分支ID & ":" & Val(rsMerge!ID & "")
            End If
            rsTmp.MoveNext
        Loop
        mstrMerge = Mid(mstrMerge, 2)
        lblMerge.Caption = lblMerge.Caption & Mid(strTmp, 2)
        '如果当前没有可用合并路径阶段,则隐藏
        If mrsMerge.RecordCount = 0 Then
            lblMerge.Visible = False
            vsItem(1).Visible = False
            cmdMergeStep.Visible = False
            vsItem(0).Height = vsItem(1).Top + vsItem(1).Height - vsItem(0).Top
        End If
    Else
        '没有合并路径，则隐藏按钮和表格
        lblMerge.Visible = False
        vsItem(1).Visible = False
        cmdMergeStep.Visible = False
        vsItem(0).Height = vsItem(1).Top + vsItem(1).Height - vsItem(0).Top
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitItem()
'功能: 初始化路径项目表头
    Dim strcol As String, arrHead As Variant
    Dim i As Long
    
    If mlngFun = 2 Then
        strcol = "分类,1200,4;分类值;全选,450,4;项目内容,5950,1;执行方式;选择;婴儿;ID;医嘱内容ID;长嘱;文件ID"
    Else
        strcol = "分类,1200,4;分类值;全选" & IIf(mlngFun <> Func重新生成, ",450,4", "") & ";项目内容," & IIf(mstrBaby = "", 5950, 4950) & ",1;执行方式,900,1" & _
                ";选择" & IIf(mlngFun <> Func重新生成, ",500,4", "") & _
                ";重选,500,4;婴儿" & IIf(mstrBaby = "", "", ",1100,1") & _
                ";ID;医嘱内容ID;长嘱;文件ID;内容要求;变异原因" & IIf(mlngFun = 0, ",1800,4", "") & ";是否最后一天;阶段ID;合并路径记录ID;诊疗项目ID;重复项目"
    End If
    arrHead = Split(strcol, ";")
    Set mcol = New Collection
   
    With vsItem(0)
        .Redraw = flexRDNone
        .Clear
        .FixedRows = 1: .FixedCols = 1
        .Cols = UBound(arrHead) + 1
        .Rows = .FixedRows
        
        For i = 0 To UBound(arrHead)
            mcol.Add i, Split(arrHead(i), ",")(0)
            .TextMatrix(.FixedRows - 1, i) = Split(arrHead(i), ",")(0)
            
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(i) = False
                .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(i) = True
                .ColWidth(i) = 0
            End If
        Next
        .Redraw = True
    End With
    With vsItem(1)
        .Redraw = flexRDNone
        .Clear
        .FixedRows = 1: .FixedCols = 1
        .Cols = UBound(arrHead) + 1
        .Rows = .FixedRows
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, i) = Split(arrHead(i), ",")(0)
            
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(i) = False
                .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(i) = True
                .ColWidth(i) = 0
            End If
        Next
        .Redraw = True
    End With
End Sub

Private Sub vsPhase_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If vsPhase.MouseCol >= 0 And vsPhase.MouseRow >= 0 Then
        Dim strInfo As String
        strInfo = Trim(vsPhase.Cell(flexcpData, vsPhase.MouseRow, vsPhase.MouseCol))
        Call zlCommFun.ShowTipInfo(vsPhase.Hwnd, strInfo)
    End If
End Sub


Private Function GetBabyRegList() As String
'功能：读取病人的婴儿姓名列表
'参数：
'返回："姓名1,姓名2,姓名3…
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select 序号,婴儿姓名 From 病人新生儿记录 Where 病人ID=[1] And 主页ID=[2] Order by 序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetBabyRegList", mPati.病人ID, mPati.主页ID)
    
    strSql = ""
    Do While Not rsTmp.EOF
        strSql = IIf(strSql = "", "", strSql & "|") & "婴儿:" & Nvl(Replace(rsTmp!婴儿姓名, "|", "_"))
        rsTmp.MoveNext
    Loop
    GetBabyRegList = strSql
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetBabyIndex(strtxt As String) As String
'功能：根据当前行的内容返回婴儿序号
    Dim i As Long, j As Long
    Dim arrtmp As Variant
    Dim arrBaby As Variant
    Dim strBaby As String
    
    If mstrBaby <> "" Then
        arrtmp = Split("病人本人|" & mstrBaby, "|")
        arrBaby = Split(strtxt, "|")
        For j = 0 To UBound(arrBaby)
            For i = 0 To UBound(arrtmp)
                If arrtmp(i) = arrBaby(j) Then
                    strBaby = strBaby & "|" & i
                    Exit For
                End If
            Next
        Next
        GetBabyIndex = Mid(strBaby, 2)
    Else
        GetBabyIndex = "0"  '没有婴儿缺省取病人本人
    End If
End Function

Private Function UnExecutedOfPhase(ByVal lng项目ID As Long) As Boolean
'功能：检查指定的项目在当前阶段是否执行过
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select 1 From 病人路径执行 Where 路径记录ID = [1] And 阶段ID = [2] And 项目ID = [3]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, Val(vsPhase.ColData(vsPhase.Col)), lng项目ID)
    UnExecutedOfPhase = rsTmp.RecordCount = 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, j As Long, blnEnd As Boolean, strIDs As String, strAdviceOfItem As String
    Dim arrSQL As Variant, arrBaby As Variant, DatCurr As Date
    Dim strTmp As String, strBaby As String, strBB As String, lng天数 As Long
    Dim rsTmp As ADODB.Recordset, rsLastAdvice As ADODB.Recordset
    Dim rsUsed As ADODB.Recordset   '重新生成时校对但未作废的医嘱
    Dim blnHave As Boolean, blnHaveDoc As Boolean
    Dim strLAdivceOfItem As String  '上次生成路径项目及长嘱ID
    Dim strLAdvices As String       '上次生成的长嘱ID
    Dim str项目IDs As String, str重选项目IDs As String, str医嘱IDs As String
    Dim k As Long, n As Long, strAgain As String
    Dim colItem As New Collection
    Dim strAgaignTmp As String
    Dim str路径项目IDs As String   '路径生成时中医修改了的配方的，且超出了允许修改配方的比例的项目，对应的变异原因：项目ID1|变异编码1,项目2|变异编码2····
    Dim colPathItems As New Collection
    
    arrSQL = Array()
    '1.检查必须执行一次的项目
    For k = 0 To vsItem.count - 1
        With vsItem(k)
            If mlngFun = 0 Then
                If k = 0 Then
                    With mrsPhase
                        .Filter = "ID=" & vsPhase.ColData(vsPhase.Col)
                        If Not IsNull(!开始天数) Then
                            If IsNull(!结束天数) Then
                                blnEnd = (Val(!开始天数) = mlng天数)
                            Else
                                blnEnd = (Val(!结束天数) = mlng天数)
                            End If
                        End If
                    End With
                End If
                '合并路径由于可能是多个阶段项目，所有是否最后一天，根据列：是否最后一天来判断
                If blnEnd Or k = 1 Then
                    For i = 1 To .Rows - 1
                        If k = 0 Or .TextMatrix(i, mcol("是否最后一天")) = "1" Then
                            If .RowData(i) = 执行方式.T2至少一次 Or .RowData(i) = 执行方式.T4必须且仅一次 Then
                                If .Cell(flexcpChecked, i, mcol("选择")) = 2 Then
                                    If UnExecutedOfPhase(Val(.TextMatrix(i, mcol("ID")))) Then
                                        .Cell(flexcpChecked, i, mcol("选择")) = 1
                                        .Row = i
                                        blnHave = True
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
                If k = vsItem.count - 1 And blnHave Then
                    MsgBox "本阶段至少或必须产生一次的项目没有选择，系统已自动选择，请检查确认后继续。", vbInformation, gstrSysName
                    Exit Sub
                End If
                
                '每天生成的项目，如果没有选择，则必须输入变异原因
                For i = 1 To .Rows - 1
                    If .RowData(i) = 执行方式.T1每天必须 Then
                        If .Cell(flexcpChecked, i, mcol("选择")) = 2 Then
                            If .TextMatrix(i, mcol("变异原因")) = "" Then
                                MsgBox "必须生成的项目，如果选择不生成，则要求必须选择变异原因。", vbInformation, gstrSysName
                                If .Visible And .Enabled Then .SetFocus
                                .Select i, mcol("变异原因")
                                Exit Sub
                            End If
                        End If
                    End If
                Next
            End If
        End With
        blnEnd = False
    Next
    
    '2.获取路径项目对应的医嘱
    If mPP.当前天数 > 0 Then
        For k = 0 To vsItem.count - 1
            With vsItem(k)
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpChecked, i, mcol("选择")) = 1 Then
                        If Val(.TextMatrix(i, mcol("长嘱"))) = 1 Then
                            '如果选择了重选医嘱，则这些该项目对应的医嘱要显示出来，在医嘱编辑界面，允许选择和以前一样的，
                            '医嘱下达界面保存时不新产生医嘱记录的数据，但是要将以前的医嘱ID收集来填“病人路径医嘱”
                            If .Cell(flexcpChecked, i, mcol("重选")) = 1 Then
                                str重选项目IDs = str重选项目IDs & "," & .TextMatrix(i, mcol("ID"))
                            Else
                                str项目IDs = str项目IDs & "," & .TextMatrix(i, mcol("ID"))
                            End If
                        End If
                    End If
                Next
            End With
        Next
        If str重选项目IDs <> "" Then
            str重选项目IDs = Mid(str重选项目IDs, 2)
            Set rsLastAdvice = GetLastAdvice(str重选项目IDs)    '上次生成了医嘱的记录集，用于传入到医嘱下达窗体，保存时检查是否生成新的医嘱
        End If
        If str项目IDs <> "" Then
            str项目IDs = Mid(str项目IDs, 2)
            Set rsTmp = GetLastAdvice(str项目IDs) '如果阶段不同，本次的项目ID和前次的不一样,只是名称相同
            
            str项目IDs = ""
            strLAdivceOfItem = ""
            strLAdvices = ""
            For i = 1 To rsTmp.RecordCount
                '前一次生成了的就不用重复产生医嘱，但要将以前的医嘱ID收集来填“病人路径医嘱”
                strLAdivceOfItem = strLAdivceOfItem & "," & rsTmp!项目ID & ":" & rsTmp!病人医嘱id
                strLAdvices = strLAdvices & "," & rsTmp!病人医嘱id
                '收集前一次已生成了长期医嘱的项目ID
                If InStr("," & str项目IDs & ",", "," & rsTmp!项目ID & ",") = 0 Then
                    str项目IDs = str项目IDs & "," & rsTmp!项目ID
                End If
                rsTmp.MoveNext
            Next
            strLAdivceOfItem = Mid(strLAdivceOfItem, 2)
            strLAdvices = Mid(strLAdvices, 2)
            str项目IDs = str项目IDs & ","   '本次不用再生成医嘱的路径项目ID
        End If
    End If
    '91635 重新生成医嘱：
    '1）情况一：重新生成项目对应的医嘱中存在已经校对但未作废的医嘱记录时,允许用户生成其他医嘱,但校对未作废的医嘱保持不变。
    '2）情况二：重新生成项目对应的医嘱都是未校对的,则删除该项目对应的所有医嘱,重新产生医嘱记录数据。
    '重新生成时,已经校对但未作废的医嘱记录集,用于传人到医嘱下达窗体,保存时检查是否生成新的医嘱
    If mlngFun = Func重新生成 Then
        Set rsUsed = GetUsedAdvice(mlng执行ID, mlng项目ID)
        If rsUsed.RecordCount > 0 Then
            If rsLastAdvice Is Nothing Then
                Set rsLastAdvice = rsUsed
            Else
                For i = 1 To rsUsed.RecordCount
                    rsLastAdvice.Filter = "项目ID =" & rsUsed!项目ID & " And 组ID = " & rsUsed!组ID
                    If rsLastAdvice.RecordCount = 0 Then
                        rsLastAdvice.AddNew
                        rsLastAdvice!项目ID = rsUsed!项目ID
                        rsLastAdvice!病人医嘱id = rsUsed!病人医嘱id
                        rsLastAdvice!组ID = rsUsed!组ID
                        rsLastAdvice!诊疗项目ID = rsUsed!诊疗项目ID
                        rsLastAdvice.Update
                    End If
                    rsUsed.MoveNext
                Next
            End If
        End If
    End If
    
    strIDs = ""
    For k = 0 To vsItem.count - 1
        With vsItem(k)
            '必须生成的，如果选择不生成医嘱（选择了变异原因），则要生成路径项目
            If mlngFun = 0 Then
                For i = 1 To .Rows - 1
                    If .RowData(i) = 执行方式.T1每天必须 Then
                        If .Cell(flexcpChecked, i, mcol("选择")) = 2 Then
                            If InStr("," & str项目IDs & ",", "," & .TextMatrix(i, mcol("ID")) & ",") = 0 Then
                                str项目IDs = str项目IDs & "," & .TextMatrix(i, mcol("ID"))
                            End If
                        End If
                    End If
                Next
            End If
            
            '产生要生成医嘱的项目ID串，前次已生成长嘱的项目不用再生成,必须生成的而选择了不生成的不用生成
            For i = 1 To .Rows - 1
                .TextMatrix(i, mcol("重复项目")) = ""
                If .Cell(flexcpChecked, i, mcol("选择")) = 1 Then
                    If InStr(str项目IDs, "," & .TextMatrix(i, mcol("ID")) & ",") = 0 Then
                        strTmp = Trim(.TextMatrix(i, mcol("医嘱内容ID")))
                        If strTmp <> "" Then
                            '项目内容相同且对应的医嘱也相同，则不重复生成
                            strAgaignTmp = Trim(.TextMatrix(i, mcol("诊疗项目ID")))
                            '手录医嘱不判断重复
                            If InStr(strAgain & vbCrLf, vbCrLf & strAgaignTmp & vbCrLf) = 0 Or strAgaignTmp = "" Then
                                strAgain = strAgain & vbCrLf & strAgaignTmp
                                If strAgaignTmp <> "" Then
                                    colItem.Add .TextMatrix(i, mcol("ID")) & vbCrLf & .TextMatrix(i, mcol("项目内容")), strAgaignTmp
                                End If
                                arrBaby = Split(GetBabyIndex(.TextMatrix(i, mcol("婴儿"))), "|")
                                For n = LBound(arrBaby) To UBound(arrBaby)
                                    strBB = arrBaby(n) & ":" & .TextMatrix(i, mcol("ID"))
                                    If InStr(strTmp, ",") = 0 Then
                                         strBaby = strTmp & ":" & strBB
                                     Else
                                         strBaby = Replace(strTmp, ",", ":" & strBB & ",") & ":" & strBB
                                     End If
                                     strIDs = strIDs & "," & strBaby
                                Next
                            Else
                                .TextMatrix(i, mcol("重复项目")) = colItem(strAgaignTmp)
                            End If
                        End If
                    End If
                    If blnHaveDoc = False Then
                        If Trim(.TextMatrix(i, mcol("文件ID"))) <> "" Then blnHaveDoc = True
                    End If
                End If
            Next
        End With
    Next
    strIDs = Mid(strIDs, 2) '医嘱内容ID:婴儿序号:路径项目ID,...，例：227:0:38,335:1:69
    
    If blnHaveDoc Then
        If InStr(GetInsidePrivs(p住院病历管理), ";病历书写;") = 0 Then
            MsgBox "你没有病历书写的权限，不能生成包含病历的路径项目。", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
    End If
    
    
    '产生医嘱的缺省开始执行时间
    DatCurr = mdatDur
        
    If strIDs <> "" Then    '全是无需执行的项目时不产生医嘱，但要产生路径执行项目
        If InStr(GetInsidePrivs(p住院医嘱下达), ";医嘱下达;") = 0 Then
            MsgBox "你没有医嘱下达的权限，不能生成包含医嘱的路径项目。", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
        '检查时间
        If Format(DatCurr, "YYYY-MM-DD") > Format(dtpAdviceTime.Value, "YYYY-MM-DD") Or Format(dtpAdviceTime.Value, "YYYY-MM-DD") > Format(DatCurr + mlng路径医嘱天数, "YYYY-MM-DD") Then
            MsgBox "临床路径的医嘱必须在当前日期和提前的天数之间，当前允许提前" & mlng路径医嘱天数 & "天。", vbInformation, gstrSysName
            If dtpAdviceTime.Enabled And dtpAdviceTime.Visible Then dtpAdviceTime.SetFocus
            Exit Sub
        End If
        
        Me.Hide
        If gobjKernel.ShowAdviceEdit(mfrmParent, mint场合, 1, mPati.病人ID, mPati.主页ID, strIDs, CDate(dtpAdviceTime.Value), arrSQL, strAdviceOfItem, rsLastAdvice, DatCurr, str路径项目IDs, mclsMipModule) = False Then
            Unload Me
            Exit Sub
        End If
        '如果中医配方修改的味数超过设置的标准比例，并填了变异原因，则补上变异原因
        If str路径项目IDs <> "" Then
            '如果一个项目有两个变异原因，则取第一个
            On Error Resume Next
            For i = 0 To UBound(Split(str路径项目IDs, ","))
                colPathItems.Add Split(Split(str路径项目IDs, ",")(i), "|")(1), "_" & Split(Split(str路径项目IDs, ",")(i), "|")(0)
            Next
            For i = 1 To vsItem(0).Rows - 1
                strTmp = ""
                If vsItem(0).TextMatrix(i, mcol("ID")) & "" <> "" Then
                    strTmp = colPathItems("_" & vsItem(0).TextMatrix(i, mcol("ID")))
                    If strTmp <> "" Then
                        vsItem(0).Cell(flexcpData, i, mcol("变异原因")) = strTmp
                    End If
                End If
            Next
            For i = 1 To vsItem(1).Rows - 1
                strTmp = ""
                If vsItem(1).TextMatrix(i, mcol("ID")) & "" <> "" Then
                    strTmp = colPathItems("_" & vsItem(1).TextMatrix(i, mcol("ID")))
                    If strTmp <> "" Then
                        vsItem(1).Cell(flexcpData, i, mcol("变异原因")) = strTmp
                    End If
                End If
            Next
            On Error GoTo 0
        End If
    End If
    
    str医嘱IDs = ""  '记录待停止的长嘱ID
    If (mlngFun = 0 Or mlngFun = 3) Then
        If InStr(GetInsidePrivs(p住院医嘱下达), ";医嘱停止;") > 0 Then
            Set rsTmp = GetLastAdvice(, "," & strLAdvices & ",")
            For i = 1 To rsTmp.RecordCount
                str医嘱IDs = str医嘱IDs & "," & rsTmp!病人医嘱id
                rsTmp.MoveNext
            Next
            str医嘱IDs = Mid(str医嘱IDs, 2)
        End If
    End If
    
    '需要继续产生的长嘱的病人路径医嘱
    If strLAdivceOfItem <> "" Then
        If strAdviceOfItem = "" Then
            strAdviceOfItem = strLAdivceOfItem
        Else
            strAdviceOfItem = strAdviceOfItem & "," & strLAdivceOfItem
        End If
    End If
    '重新生成时收集校对未停用的医嘱Id来填"病人路径医嘱"
    If mlngFun = Func重新生成 Then
        rsUsed.Filter = ""
        For i = 1 To rsUsed.RecordCount
            If i = 1 And strAdviceOfItem = "" Then
                strAdviceOfItem = mlng项目ID & ":" & rsUsed!病人医嘱id
            Else
                If InStr("," & strAdviceOfItem & ",", "," & mlng项目ID & ":" & rsUsed!病人医嘱id & ",") = 0 Then '避免重复添加（有可能医嘱下达界面已经产生了该数据）
                    strAdviceOfItem = strAdviceOfItem & "," & mlng项目ID & ":" & rsUsed!病人医嘱id
                End If
            End If
            rsUsed.MoveNext
        Next
    End If
    Call SaveData(arrSQL, strAdviceOfItem, lng天数)
    '修正医嘱序号
    Call ModifyAdviceSerialNum
    
    '生成路径后，检查是否有需要停止的长嘱(上次有，但本次没有的长嘱)
    '-----------------------------------------------------------------------------
    If str医嘱IDs <> "" Then
        '如果本次生成没有这些长嘱，则需要停止
         strIDs = GetShouldStopAdvice(str医嘱IDs, lng天数)
         If strIDs <> "" Then
            Me.Hide
            Call gobjKernel.ShowAdviceOperate(mfrmParent, mint场合, mPati.病人ID, mPati.主页ID, mPati.病区ID, strIDs, DateAdd("s", 1, DatCurr), mclsMipModule)
            
            Call CheckStopAdvice(mPati.病人ID, mPati.主页ID, strIDs)
            '医生没有停止的长嘱，需要生成为路径外项目
            If strIDs <> "" Then
            Call AddOutPathItem(strIDs, 1, mPati.病人ID, mPati.主页ID)
         End If
         End If
         
    End If
    
    mblnOK = True
    Unload Me
End Sub

Private Sub ModifyAdviceSerialNum()
'功能：重新整理医嘱序号
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    Screen.MousePointer = 11
    strSql = "Select Count(*) as Num From (Select 序号,Count(ID) From 病人医嘱记录 Where 病人ID=[1] And 主页ID=[2] Having Count(ID)>1 Group by 序号)"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "查询病人医嘱数量", mPati.病人ID, mPati.主页ID)
    
    If rsTmp.EOF Then Screen.MousePointer = 0: Exit Sub
    
    If Nvl(rsTmp!Num, 0) = 0 Then Screen.MousePointer = 0: Exit Sub
    
    strSql = "ZL_病人医嘱记录_更新序号(NULL,NULL," & mPati.病人ID & "," & mPati.主页ID & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "修正医嘱序号")
    
    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetPreSendData(ByRef lng阶段ID As Long, ByRef dat日期 As Date)
'功能：根据当前阶段和日期返回上一次生成路径项目的阶段和日期
    Dim strSql As String, rsTmp As ADODB.Recordset

    On Error GoTo errH
    strSql = "Select 阶段id, 日期, 天数" & vbNewLine & _
             "From 病人路径执行" & vbNewLine & _
             "Where 路径记录id = [1] And 登记时间 = (Select Max(登记时间)" & vbNewLine & _
             "                             From 病人路径执行" & vbNewLine & _
             "                             Where 路径记录id = [1] And 登记时间 <  (Select Min(登记时间) 登记时间" & vbNewLine & _
             "                                    From 病人路径执行 A" & vbNewLine & _
             "                                    Where a.路径记录id = [1] And a.阶段id = [2] And a.日期 = [3])" & vbNewLine & _
             "                             ) And Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
    If rsTmp.RecordCount > 0 Then    '本次生成是第一次生成的情况无记录
        lng阶段ID = rsTmp!阶段ID
        dat日期 = rsTmp!日期
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function GetLastAdvice(Optional ByVal strIDsOfLA As String, Optional ByVal strLAdvices As String) As ADODB.Recordset
'功能：获取路径项目在最近一次生成的长期医嘱的项目记录集
'参数：strIDsOfLA=当前选择的含有长期医嘱的项目ID，
'    :strLAdvices=取上次的所有有效的长嘱(已校对，已重整，已暂停，已启用),排除特殊医嘱（术前，术后，病危，病重，记录入出量）,护理等级
'返回：1.返回的路径项目ID是本次生成的路径项目的ID;2-病人医嘱Id是本次生成需要停止的长嘱ID
    Dim strSql As String
    Dim lng阶段ID As Long, dat日期 As Date, lng天数 As Long
    
    '找当前正在执行的长期医嘱
    If strIDsOfLA <> "" Then
        '存在与当前项目同名的，且上一次执行（前一天或同一天）生成了长期医嘱的(且未作废或停止的)，则本次不重复生成
        '项目id是主键，已确定了路径id及版本号
        strSql = "Select /*+ rule*/ f.id as 项目id, b.病人医嘱id,Nvl(d.相关id,d.id) 组ID,d.诊疗项目ID" & vbNewLine & _
            "From 病人路径执行 A, 病人路径医嘱 B, 病人医嘱记录 D, 临床路径项目 E, 临床路径项目 F" & vbNewLine & _
            IIf(InStr(strIDsOfLA, ",") > 0, ",(Select Column_Value As 项目id From Table(f_Num2list([1]))) C Where c.项目id = f.Id ", " Where f.id = [1]") & vbNewLine & _
            "     And f.项目内容 = e.项目内容 And e.Id = a.项目id And a.路径记录id = [2] And" & vbNewLine & _
            "      a.阶段id = [3] And a.日期 = [4] And a.Id = b.路径执行id And b.病人医嘱id = d.Id And d.医嘱期效 = 0 And d.医嘱状态 Not In(4,8,9)" & vbNewLine & _
            " Group By f.Id, b.病人医嘱id, Nvl(d.相关id, d.Id), d.诊疗项目id, d.序号 " & _
            " Order by d.序号"
    Else
        '排除护理等级（与医嘱停止界面保持一致）
        strSql = "Select b.病人医嘱id" & vbNewLine & _
                "From 病人路径执行 A, 病人路径医嘱 B, 病人医嘱记录 C,诊疗项目目录 D" & vbNewLine & _
                "Where a.路径记录id = [2] And a.阶段id = [3] And a.日期 = [4] And a.Id = b.路径执行id And b.病人医嘱id = c.Id And c.医嘱期效 = 0 And" & vbNewLine & _
                "      c.医嘱状态 In (3, 5, 6, 7) And C.诊疗项目ID=D.ID And Not(D.类别='H' and D.操作类型='1' And D.执行频率=2) " & _
                "   And Not(D.类别='Z' And D.操作类型 IN('4','14', '9', '10', '12')) And instr( '" & strLAdvices & "',','|| b.病人医嘱ID||',')=0"
    End If
    On Error GoTo errH
    If mlngFun = 0 Then
        lng阶段ID = mPP.当前阶段ID
        dat日期 = CDate(mPP.当前日期)
    Else
    '补充生成，重新生成时，取前次生成的阶段和日期
        Call GetPreSendData(lng阶段ID, dat日期)
    End If
    Set GetLastAdvice = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strIDsOfLA, mPP.病人路径ID, lng阶段ID, dat日期)
        
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetUsedAdvice(ByVal lng执行ID As Long, ByVal lng项目ID As Long) As ADODB.Recordset
'功能:重新生成时,返回当前项目中校对但未作废的医嘱记录
    Dim strSql As String
    
    strSql = "Select [1] As 项目id, a.病人医嘱id, Nvl(b.相关id, b.Id) As 组id, b.诊疗项目id" & vbNewLine & _
            "From 病人路径医嘱 A, 病人医嘱记录 B" & vbNewLine & _
            "Where a.病人医嘱id = b.Id And a.路径执行id = [2] And b.医嘱状态 > 1 And b.医嘱状态 <> 4" & vbNewLine & _
            "Order By b.序号"
    On Error GoTo errH
    
    Set GetUsedAdvice = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng项目ID, lng执行ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub GetLastEvaluate(strLastVariation As String, str审核人 As String, str评估人 As String)
'功能：获得最后一次评估的信息
    Dim strSql As String, rsTmp As Recordset
    
    On Error GoTo errH
    strSql = "Select 变异原因,变异审核人,评估人 From 病人路径评估 Where 路径记录ID=[1] And 天数=[2] And 阶段ID=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前天数, mPP.当前阶段ID)
    If rsTmp.RecordCount > 0 Then
        strLastVariation = rsTmp!变异原因 & ""
        str审核人 = rsTmp!变异审核人 & ""
        str评估人 = rsTmp!评估人 & ""
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetFirstType()
'功能：如果一个项目都没有，则从数据库中取第一个分类
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    On Error GoTo errH
    strSql = "Select 名称 from 临床路径分类 where 路径ID=[1] and 版本号=[2] and NVL(分支ID,0)=[3] And 序号=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "取第一个分类", mPP.路径ID, mPP.版本号, Val(Mid(tabBranch.SelectedItem.Key, 2)))
    If rsTmp.RecordCount > 0 Then GetFirstType = rsTmp!名称 & ""
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SaveData(ByVal arrSQL As Variant, ByVal strAdviceOfItem As String, ByRef lng天数 As Long)
'功能：保存路径项目
'参数：strAdviceOfItem=路径项目与医嘱ID的对应,例：38:1983,69:1978
    Dim colSQL As New Collection, colDoc As New Collection, blnTrans As Boolean, colNewDoc As New Collection
    Dim strSql As String, i As Long, j As Long, l As Long, k As Long, lngBaby As Long
    Dim strDate As String, strAddDate As String, strAdviceIDs As String, strFileIDs As String, strFileID As String
    Dim str病人病历IDs As String, strEMRID As String, lng病历ID As Long, strVariation As String, strBaby As String
    Dim strFileIDsTmp As String, strFiles As String
    Dim arrItem As Variant, lng序号 As Long
    Dim blnIsSend As Boolean   '判断用户是否勾选了项目
    Dim strLastVariation As String
    Dim str审核人 As String
    Dim str评估人 As String
    Dim varFilter As Variant
    Dim AddDate As Date
    Dim strFirstType As String
    Dim strMergeStep As String
    Dim strEPR As String
    Dim blnAgain As Boolean
    Dim strAgaignTmp As String
    Dim strAgain As String
    Dim colItemName As New Collection
    Dim blnDef As Boolean
    Dim strPara As String, strParaTmp As String
    Dim rsTmp As ADODB.Recordset
    Dim arrtmp As Variant
    
    AddDate = zlDatabase.Currentdate
    strAddDate = "To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    arrItem = Split(strAdviceOfItem, ",")
    
    '如果一个项目都没有，则从数据库中取第一个分类
    If vsItem(0).TextMatrix(1, mcol("分类")) = "" Then
        strFirstType = GetFirstType
    Else
        strFirstType = vsItem(0).TextMatrix(1, mcol("分类"))
    End If

    strMergeStep = Mid(mstrMergeStep, 2)
    lng天数 = mlng天数
    
    strDate = "To_Date('" & Format(mdat时间, "yyyy-MM-dd") & "','YYYY-MM-DD')"
    'mrsPhase调用前已执行filter
    '提前进度
    If mlng时间进度 = 2 Then
        k = mlng提前天数
    Else
        k = mlng提前天数 + 1
    End If
     '判断当前选择的阶段开始天数是否等于传入天数，否则中间的天数用未生成任何项目来处理
    If lng天数 > k Then
        varFilter = mrsPhase.Filter
        Call GetLastEvaluate(strLastVariation, str审核人, str评估人)
        For i = k To lng天数 - 1
            '生成
            mrsPhase.Filter = "开始天数=" & i & " And 父ID = 0" & IIf(mblnIsHaveBranch, " And 分支ID=" & Mid(tabBranch.SelectedItem.Key, 2), "")
            If Not mrsPhase.EOF Then
                strSql = "Zl_病人路径生成_Insert(1," & mPati.病人ID & "," & mPati.主页ID & ",NULL," & mPati.科室ID & "," & _
                        mPP.病人路径ID & "," & mrsPhase!ID & _
                        "," & strDate & "," & mlng提前天数 & _
                        ",'" & strFirstType & "',Null" & _
                        ",Null,Null,Null,'" & UserInfo.姓名 & "',To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & _
                        "','YYYY-MM-DD HH24:MI:SS'),'未生成任何项目',Null,'已经执行|1" & vbTab & "已经执行',Null,Null,'',1" & _
                        ",Null,'" & strMergeStep & "')"
                colSQL.Add strSql, "C" & colSQL.count + 1
                
                '登记时间加一秒，是为了取消生成时取上一个阶段的ID。
                AddDate = AddDate + 1 / 24 / 60 / 60
                '评估
                strSql = "Zl_病人路径评估_Insert(1," & mPP.病人路径ID & "," & mrsPhase!ID & _
                        "," & strDate & "," & mlng提前天数 & ",'" & _
                        str评估人 & "',1,'','" & UserInfo.姓名 & "','" & str审核人 & "','" & strLastVariation & "',1,Null,Null" & ",Null,1" & ")"
                        
                colSQL.Add strSql, "C" & colSQL.count + 1
            End If
        Next
        mrsPhase.Filter = varFilter
    End If
        
    For k = 0 To vsItem.count - 1
        With vsItem(k)
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, mcol("选择")) = 1 Or .Cell(flexcpChecked, i, mcol("选择")) = 2 And mlngFun = 0 And .RowData(i) = 执行方式.T1每天必须 Then
                    strBaby = GetBabyIndex(.TextMatrix(i, mcol("婴儿")))
                    
                    strAdviceIDs = ""
                    str病人病历IDs = ""
                    strFileIDs = ""
                    strVariation = ""
                    blnAgain = False
                    strFileIDsTmp = ""
                    blnDef = False
                    
                    If .Cell(flexcpChecked, i, mcol("选择")) = 1 Then
                        If Val(.TextMatrix(i, mcol("医嘱内容ID"))) <> 0 Then
                            strAgaignTmp = Trim(.TextMatrix(i, mcol("诊疗项目ID")))
                            If InStr(strAgain & vbCrLf, vbCrLf & strAgaignTmp & vbCrLf) = 0 Then
                                For j = 0 To UBound(arrItem)
                                    If Split(arrItem(j), ":")(0) = .TextMatrix(i, mcol("ID")) Then  '路径项目ID
                                        strAdviceIDs = strAdviceIDs & "," & Split(arrItem(j), ":")(1)  '医嘱ID
                                    ElseIf .TextMatrix(i, mcol("重复项目")) <> "" Then
                                        '如果有重复项目，如果项目名称相同，则不重复生成相同项目，如果名称不同，则生成项目指向相同医嘱
                                        If Split(arrItem(j), ":")(0) = Split(.TextMatrix(i, mcol("重复项目")), vbCrLf)(0) Then
                                            If .TextMatrix(i, mcol("项目内容")) = Split(.TextMatrix(i, mcol("重复项目")), vbCrLf)(1) Then
                                                blnAgain = True
                                            Else
                                                strAdviceIDs = strAdviceIDs & "," & Split(arrItem(j), ":")(1)  '医嘱ID
                                            End If
                                        End If
                                    End If
                                Next
                            Else
                                '处理继承的项目，如果有重复的则只生成一个项目
                                If .TextMatrix(i, mcol("项目内容")) = colItemName("C" & strAgaignTmp) Then
                                    blnAgain = True
                                Else
                                    For j = 0 To UBound(arrItem)
                                        If Split(arrItem(j), ":")(0) = .TextMatrix(i, mcol("ID")) Then  '路径项目ID
                                            strAdviceIDs = strAdviceIDs & "," & Split(arrItem(j), ":")(1)  '医嘱ID
                                        ElseIf .TextMatrix(i, mcol("重复项目")) <> "" Then
                                            '如果有重复项目，如果项目名称相同，则不重复生成相同项目，如果名称不同，则生成项目指向相同医嘱
                                            If Split(arrItem(j), ":")(0) = Split(.TextMatrix(i, mcol("重复项目")), vbCrLf)(0) Then
                                                If .TextMatrix(i, mcol("项目内容")) = Split(.TextMatrix(i, mcol("重复项目")), vbCrLf)(1) Then
                                                    blnAgain = True
                                                Else
                                                    strAdviceIDs = strAdviceIDs & "," & Split(arrItem(j), ":")(1)  '医嘱ID
                                                End If
                                            End If
                                        End If
                                    Next
                                End If
                                blnDef = True
                            End If
                            If Not blnAgain And Not blnDef Then
                                strAgain = strAgain & vbCrLf & strAgaignTmp
                                colItemName.Add .TextMatrix(i, mcol("项目内容")), "C" & strAgaignTmp
                            End If
                            strAdviceIDs = Mid(strAdviceIDs, 2)
                            
                            '如果中药配方修改过后填写了变异原因，则保存到当条项目中
                            If .Cell(flexcpData, i, mcol("变异原因")) <> "" Then
                                strVariation = .Cell(flexcpData, i, mcol("变异原因"))
                            End If
                        End If
                        
                        strEPR = Trim(.TextMatrix(i, mcol("文件ID")))     '可能有多个
                        If strEPR <> "" Then
                           strFiles = Split(strEPR, "|")(0)  '旧版
                           strPara = Split(strEPR, "|")(1)  '新版
                           strEPR = ""
                           If strFiles <> "" Then
                                arrtmp = Split(strBaby, "|")
                                For l = LBound(arrtmp) To UBound(arrtmp)
                                    strEMRID = "": strFileIDsTmp = ""
                                    For j = 0 To UBound(Split(strFiles, ","))
                                        strFileID = Split(strFiles, ",")(j)
                                         '病历始终不生成重复的，一个病历文件只生成一个
                                        lngBaby = CLng(arrtmp(l) & "")
                                        If InStr(strEPR & ",", "," & lngBaby & "_" & strFileID & ",") = 0 Then
                                            lng病历ID = zlDatabase.GetNextId("电子病历记录")
                                            strEMRID = strEMRID & "," & lng病历ID
                                            colDoc.Add lng病历ID & ":" & lngBaby & ":" & mEditType("C" & strFileID), "C" & (colDoc.count + 1)
                                            strFileIDsTmp = strFileIDsTmp & "," & strFileID
                                            strEPR = strEPR & "," & lngBaby & "_" & strFileID
                                        End If
                                    Next
                                    str病人病历IDs = str病人病历IDs & "|" & Mid(strEMRID, 2)
                                    strFileIDs = strFileIDs & "|" & Mid(strFileIDsTmp, 2)
                                Next
                                str病人病历IDs = Mid(str病人病历IDs, 2)
                                strFileIDs = Mid(strFileIDs, 2)
                            End If
                            If strPara <> "" And Not gobjEmr Is Nothing Then '新版病历
                                If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then Set gobjEmr = Nothing
                                If Not gobjEmr Is Nothing Then
                                    strParaTmp = "": strFileIDsTmp = ""
                                    For j = 0 To UBound(Split(strPara, ","))
                                        strParaTmp = "<parameter><antetypeid>" & Split(strPara, ",")(j) & "</antetypeid><patient>" & mPati.病人ID & "</patient></parameter>"
                                        '记录集包含字段：原型ID,任务ID,生成时间,起始时间,终止时间；
                                        On Error Resume Next
                                        Set rsTmp = gobjEmr.MakeBeforTask(strParaTmp)
                                        Err.Clear: On Error GoTo 0
                                        If rsTmp.State <> adStateClosed Then
                                            If rsTmp.RecordCount = 1 Then
                                                strFileIDsTmp = strFileIDsTmp & "," & rsTmp!任务ID
                                            End If
                                        End If
                                    Next
                                    strPara = Mid(strFileIDsTmp, 2)
                                    colNewDoc.Add strPara, "C" & (colNewDoc.count + 1) '记录返回的任务ID,避免事务提交失败,删除生成成功的新版病历
                                End If
                            End If
                            If str病人病历IDs & strPara = "" Then blnAgain = True
                        End If
                    Else
                        strVariation = .Cell(flexcpData, i, mcol("变异原因"))
                    End If
                    
                    If Not blnAgain Then
                        lng序号 = lng序号 + 1
                        strSql = "Zl_病人路径生成_Insert(" & lng序号 & "," & mPati.病人ID & "," & mPati.主页ID & ",'" & strBaby & "'," & mPati.科室ID & "," & _
                            mPP.病人路径ID & "," & mrsPhase!ID & _
                            "," & strDate & "," & mlng提前天数 & _
                            ",'" & .TextMatrix(i, mcol("分类")) & "'," & .TextMatrix(i, mcol("ID")) & _
                            ",'" & strAdviceIDs & "','" & strFileIDs & "','" & str病人病历IDs & "'" & _
                            ",'" & UserInfo.姓名 & "',To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),Null,Null,Null,Null,Null,'" & _
                            strVariation & "',Null,Null,'" & strMergeStep & "'," & ZVal(Val(.TextMatrix(i, mcol("合并路径记录ID")))) & "," & ZVal(Val(.TextMatrix(i, mcol("阶段ID")))) & ",0," & IIf(mint场合 = 0, 1, 2) & ",'" & strPara & "')"
                        colSQL.Add strSql, "C" & colSQL.count + 1
                        blnIsSend = True
                    End If
                End If
            Next
        End With
    Next
    '如果没有勾选任何项目，则生成一条特殊的项目：未生成任何项目
    If Not blnIsSend Then
        If mlngFun = 0 Then
            lng序号 = lng序号 + 1
            strSql = "Zl_病人路径生成_Insert(" & lng序号 & "," & mPati.病人ID & "," & mPati.主页ID & ",NULL," & mPati.科室ID & "," & _
                    mPP.病人路径ID & "," & mrsPhase!ID & _
                    "," & strDate & "," & mlng提前天数 & _
                    ",'" & strFirstType & "',Null" & _
                    ",Null,Null,Null,'" & UserInfo.姓名 & "',To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'未生成任何项目',Null,'已经执行|1" & vbTab & "已经执行',Null,Null,'',Null" & _
                    ",Null,'" & strMergeStep & "',NULL,NULL,0," & IIf(mint场合 = 0, 1, 2) & ")"
            colSQL.Add strSql, "C" & colSQL.count + 1
        End If
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        If mlngFun = 3 Then
            strSql = "Zl_病人路径生成_Delete(" & mlng执行ID & ",1)"
            zlDatabase.ExecuteProcedure strSql, Me.Caption
        End If
        
        '1.先产生医嘱,因为病人路径医嘱有外键
        For i = 0 To UBound(arrSQL)
            zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
        Next
        '2.产生病人路径数据，以及病历文件数据
        For i = 1 To colSQL.count
            Call zlDatabase.ExecuteProcedure(colSQL("C" & i), Me.Caption)
        Next
        '3.产生病历文件RTF数据
        For i = 1 To colDoc.count
            arrItem = Split(colDoc("C" & i), ":")
            If arrItem(2) = 0 Or arrItem(2) = 1 Then     '全文编辑方式的病历
                lng病历ID = (arrItem(0))
                Call ReadRTFData(lng病历ID, edtEditor)
                Call SaveRTFData(lng病历ID, mPati.病人ID, mPati.主页ID, Val(arrItem(1)), edtEditor)
            End If
        Next
    gcnOracle.CommitTrans: blnTrans = False
    Call ZLHIS_CIS_001(mclsMipModule, mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID)
 
    Exit Sub
errH:
    If blnTrans Then
        gcnOracle.RollbackTrans
        '--删除产出的新版病历
        If Not gobjEmr Is Nothing Then
            If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then Set gobjEmr = Nothing
            If Not gobjEmr Is Nothing Then
                For i = 1 To colNewDoc.count
                    strPara = "<parameter><taskid>" & colNewDoc("C" & i) & "</taskid></parameter>"
                    On Error Resume Next
                    Call gobjEmr.DeleteTask(strPara)
                    Err.Clear: On Error GoTo 0
                Next
            End If
        End If
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetShouldStopAdvice(ByVal strIDs As String, ByVal lng天数 As Long) As String
'功能：获取当前应当停止的长期医嘱（上一次执行中存在，但本次执行中不存在）
'参数：strIDs=最后一次执行的长期医嘱ID
'      lng天数=本次生成的天数
'      返回：长期医嘱ID
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long
    
    On Error GoTo errH
    strSql = "Select /*+ rule*/ Column_Value As 病人医嘱id" & vbNewLine & _
            "From Table(f_Num2list([1])) " & vbNewLine & _
            "Minus" & vbNewLine & _
            "Select b.病人医嘱id" & vbNewLine & _
            "From 病人路径执行 A, 病人路径医嘱 B" & vbNewLine & _
            "Where a.路径记录id = [2] And a.阶段id = [3] And a.天数 = [4] And a.Id = b.路径执行id"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strIDs, mPP.病人路径ID, Val(mrsPhase!ID), lng天数)
    For i = 1 To rsTmp.RecordCount
        GetShouldStopAdvice = GetShouldStopAdvice & "," & rsTmp!病人医嘱id
        rsTmp.MoveNext
    Next
    GetShouldStopAdvice = Mid(GetShouldStopAdvice, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function MakePathAdivceRS() As ADODB.Recordset
    Set MakePathAdivceRS = New ADODB.Recordset
    MakePathAdivceRS.Fields.Append "路径项目ID", adBigInt
    MakePathAdivceRS.Fields.Append "原医嘱ID", adBigInt
    
    MakePathAdivceRS.Fields.Append "路径项目分类", adVarChar, 50, adFldIsNullable
    MakePathAdivceRS.Fields.Append "医嘱IDS", adLongVarWChar, 4000, adFldIsNullable
    MakePathAdivceRS.CursorLocation = adUseClient
    MakePathAdivceRS.LockType = adLockOptimistic
    MakePathAdivceRS.CursorType = adOpenStatic
    MakePathAdivceRS.Open
End Function

Private Sub CheckStopAdvice(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByRef strUnStopIDs As String)
'功能:
'参数:
'strUnStopIDs-未停止的医嘱ID（一组医嘱的所有ID）返回要添加的路径外项目
'lng当前阶段ID-当前阶段ID
    Dim rsUnStop As ADODB.Recordset
    Dim rsPath As ADODB.Recordset
    Dim rsPathAdvice As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset

    Dim strSql As String
    Dim i As Long, j As Long
    Dim k As Long

    Dim lng病人路径Id  As Long
    Dim lng阶段ID As Long
    Dim lng天数 As Long
    Dim lngPos As Long
    Dim strDate As String
    Dim strTag As String
    Dim str相关ID As String
    Dim AddDate As Date
    Dim colSQL As New Collection
    Dim blnTrans As Boolean
    Dim str医嘱ID As String
    
    On Error GoTo errH
    strSql = "Select b.Id" & vbNewLine & _
    " From 病人医嘱记录 B" & vbNewLine & _
    " Where b.Id in (select Column_Value As 病人医嘱id From Table(f_Num2list([1]))) And b.停嘱时间 Is Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strUnStopIDs)

    '获取未停止的长嘱ID
    For i = 1 To rsTmp.RecordCount
        str医嘱ID = str医嘱ID & "," & rsTmp!ID
        rsTmp.MoveNext
    Next
    strUnStopIDs = Mid(str医嘱ID, 2)
    If strUnStopIDs = "" Then Exit Sub
    
    '获取当前路径：路径记录ID,当前阶段Id,当前日期，天数
    strSql = "Select a.路径记录id, a.当前阶段id, a.当前天数, b.日期 " & vbNewLine & _
             "From (Select a.Id As 路径记录id, a.当前阶段id, a.当前天数, Max(b.Id) 执行id" & vbNewLine & _
             "       From 病人临床路径 A, 病人路径执行 B" & vbNewLine & _
             "       Where a.病人id = [1] And a.主页id = [2] And a.Id = b.路径记录id And b.阶段id = a.当前阶段id And b.天数 = a.当前天数" & vbNewLine & _
             "       Group By a.Id, a.当前阶段id, a.当前天数) A, 病人路径执行 B" & vbNewLine & _
             "Where a.执行id = b.Id"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng病人ID, lng主页ID)

    If rsTmp.RecordCount = 1 Then
        lng病人路径Id = Val(rsTmp!路径记录ID)
        lng阶段ID = Val(rsTmp!当前阶段ID)
        strDate = "To_Date('" & Format(rsTmp!日期, "yyyy-MM-dd") & "','YYYY-MM-DD')"
        lng天数 = Val(rsTmp!当前天数)
    Else
        Exit Sub
    End If

    strSql = "select a.ID, a.相关ID, b.类别, a.诊疗项目ID, b.操作类型" & vbNewLine & _
            "  from 病人医嘱记录 a, 诊疗项目目录 b" & vbNewLine & _
            " where a.诊疗项目ID = b.id" & vbNewLine & _
            "   and a.id in (Select Column_Value As 病人医嘱id" & vbNewLine & _
            "                  From Table(f_Num2list([1])))"


    Set rsUnStop = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strUnStopIDs)

    strSql = "Select c.ID, c.相关ID,c.诊疗项目id,a.id as 路径项目ID,a.分类 as 路径项目分类 " & vbNewLine & _
            "From 临床路径项目 a, 临床路径医嘱 b, 路径医嘱内容 c" & vbNewLine & _
            "where a.id = b.路径项目id" & vbNewLine & _
            "   and b.医嘱内容id = c.id" & vbNewLine & _
            "   and a.阶段id = [1]" & vbNewLine & _
            "   and c.期效 = 0"

    Set rsPath = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng阶段ID)

    strTag = ""
    Set rsPathAdvice = Nothing
    For i = 1 To rsUnStop.RecordCount
        lngPos = rsUnStop.AbsolutePosition
        If Val(rsUnStop!相关id & "") = 0 And Not (rsUnStop!类别 & "" = "E" And rsUnStop!操作类型 & "" = "2") Or InStr(",5,6,", "," & rsUnStop!类别 & ",") > 0 Then
            '在一并给药中增加一行时，只检查和设置当前行，因为路径外项目可能和路径内项目一并给药
            If InStr(",5,6,", "," & rsUnStop!类别 & ",") > 0 Then
                '药品单个药进行匹配 65982
                rsUnStop.Filter = "ID=" & rsUnStop!ID
                str相关ID = Val(rsUnStop!相关id & "")
            Else
                rsUnStop.Filter = "ID=" & rsUnStop!ID & " Or 相关ID=" & rsUnStop!ID
                str相关ID = Val(rsUnStop!ID & "")
            End If
            '药品不含给药途径、用法、煎法，输血不含途径,检验不含采集方式，手术不含附加手术、麻醉，检查不含部位方法
            If Not (rsUnStop!类别 & "" = "E" And InStr(",2,3,4,6,", "," & rsUnStop!操作类型 & ",") > 0) _
                And Not (InStr(",G,F,D,", "," & rsUnStop!类别 & ",") > 0 And Val(rsUnStop!相关id & "") <> 0) Then
                rsPath.Filter = ""
                For j = 1 To rsPath.RecordCount
                    If Nvl(rsPath!诊疗项目ID, 0) = Nvl(rsUnStop!诊疗项目ID, 0) Then '长期医嘱统一处理
                        '路径内项目
                        If InStr("," & strTag & ",", "," & str相关ID & ",") = 0 Then
                            rsUnStop.Filter = "相关ID=" & str相关ID & " OR ID =" & str相关ID
                            If InStr(",5,6,", "," & rsUnStop!类别 & ",") > 0 Then
                                strTag = strTag & "," & rsUnStop!相关id
                            Else
                                strTag = strTag & "," & rsUnStop!ID
                            End If
                            
                            If rsPathAdvice Is Nothing Then Set rsPathAdvice = MakePathAdivceRS
                            rsPathAdvice.Filter = "路径项目ID = " & rsPath!路径项目ID
                            
                            For k = 1 To rsUnStop.RecordCount
                                rsPathAdvice.Filter = "路径项目ID = " & rsPath!路径项目ID
                                If rsPathAdvice.RecordCount = 0 Then
                                    rsPathAdvice.AddNew
                                    rsPathAdvice!路径项目ID = rsPath!路径项目ID & ""
                                    rsPathAdvice!路径项目分类 = rsPath!路径项目分类 & ""
                                    rsPathAdvice!医嘱IDs = rsUnStop!ID & ""
                                Else
                                    rsPathAdvice!医嘱IDs = rsPathAdvice!医嘱IDs & "," & rsUnStop!ID
                                End If
                                rsPathAdvice.Update
                                '从未停止的长嘱中移除
                                strUnStopIDs = Replace("," & strUnStopIDs & ",", "," & rsUnStop!ID & ",", ",")
                                If Left(strUnStopIDs, 1) = "," Then strUnStopIDs = Mid(strUnStopIDs, 2)
                                If Right(strUnStopIDs, 1) = "," Then strUnStopIDs = Mid(strUnStopIDs, 1, Len(strUnStopIDs) - 1)
                                rsUnStop.MoveNext
                            Next
                        End If
                        Exit For
                    End If
                    rsPath.MoveNext
                Next
            End If
        End If
        rsUnStop.Filter = ""
        rsUnStop.AbsolutePosition = lngPos
        rsUnStop.MoveNext
    Next
    
    If rsPathAdvice Is Nothing Then Exit Sub
    rsPathAdvice.Filter = ""
    AddDate = zlDatabase.Currentdate
    For j = 1 To rsPathAdvice.RecordCount
        strSql = "Zl_病人路径生成_Insert(1," & lng病人ID & "," & lng主页ID & ",NULL,0," & lng病人路径Id & "," & lng阶段ID & _
            "," & strDate & "," & lng天数 & ",'" & rsPathAdvice!路径项目分类 & "'," & rsPathAdvice!路径项目ID & ",'" & rsPathAdvice!医嘱IDs & "',Null,Null" & _
            ",'" & UserInfo.姓名 & "'," & "To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & "'',NULL,NULL,NULL,NULL,'',1)"
            
        colSQL.Add strSql, "C" & colSQL.count + 1
        '登记时间加一秒，是为了取消生成时取上一个阶段的ID。
        AddDate = AddDate + 1 / 24 / 60 / 60
        rsPathAdvice.MoveNext
    Next
  
    gcnOracle.BeginTrans: blnTrans = True
    For i = 1 To colSQL.count
        Call zlDatabase.ExecuteProcedure(colSQL("C" & i), "路径生成")
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
