VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Begin VB.Form frmPathSendOut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "生成路径项目"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11790
   Icon            =   "frmPathSendOut.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   11790
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdPati 
      Height          =   240
      Left            =   7560
      Picture         =   "frmPathSendOut.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "选择婴儿"
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox lstPati 
      Appearance      =   0  'Flat
      Height          =   1080
      ItemData        =   "frmPathSendOut.frx":6948
      Left            =   5160
      List            =   "frmPathSendOut.frx":6955
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   12
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
      Height          =   645
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   11790
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5175
      Width           =   11790
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   9360
         TabIndex        =   11
         Top             =   240
         Width           =   1100
      End
      Begin MSComCtl2.DTPicker dtpAdviceTime 
         Height          =   300
         Left            =   7320
         TabIndex        =   10
         Top             =   270
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   111673347
         CurrentDate     =   41129.5916666667
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   10560
         TabIndex        =   8
         Top             =   240
         Width           =   1100
      End
      Begin VB.Label lblAdviceTime 
         Caption         =   "医嘱缺省开始时间"
         Height          =   180
         Left            =   5760
         TabIndex        =   9
         Top             =   330
         Width           =   1575
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
      ScaleWidth      =   11790
      TabIndex        =   3
      Top             =   0
      Width           =   11790
      Begin VB.Label lblFont 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Height          =   180
         Left            =   3960
         TabIndex        =   14
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
         Caption         =   "以下是可以进入的时间阶段，请根据情况选择即将进入的时间阶段。"
         Height          =   615
         Left            =   2640
         TabIndex        =   4
         Top             =   360
         Width           =   9015
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   195
         Picture         =   "frmPathSendOut.frx":6971
         Top             =   45
         Width           =   720
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsItem 
      Height          =   3405
      Left            =   0
      TabIndex        =   2
      Top             =   1720
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
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   320
      ColWidthMin     =   0
      ColWidthMax     =   8000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathSendOut.frx":71F9
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
      Top             =   960
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
      FormatString    =   $"frmPathSendOut.frx":7395
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
End
Attribute VB_Name = "frmPathSendOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mPP                 As TYPE_PATH_Pati
Private mPati               As TYPE_Pati
Private mfrmParent          As Object

Private mlng时间进度        As Integer                  'mlngFun=0时传入，1=下一阶段提前至今天,2=下一阶段延后至今天,-1=下一阶段延后（继续当前阶段）,0=正常
Private mstrBaby            As String                   '婴儿姓名串,例：马云,马克思,...
Private mblnOK              As Boolean

Private mlngFun             As Long                     '0-生成路径，1-补充生成(不能选择阶段),2-查看路径阶段定义,3-重新生成医嘱
Private mlng项目ID          As Long                     '重新生成的项目ID
Private mlng执行ID          As Long                     '重新生成的路径执行ID
Private mlng病人阶段ID      As Long                     '当前选择的阶段(查看时)，或病人当前阶段（生成时）
Private mlng提前天数        As Long                     '传入时应该生成的天数(用于下一阶段提前时)
Private mlng路径医嘱天数    As Long                     '路径医嘱生成超前天数
Private mlng天数            As Long                     '当前应该生成的天数(实际天数)

Private mdatDur             As Date                     '路径生成时间
Private mcol                As Collection
Private mEditType           As Collection

Private mrsPhase            As ADODB.Recordset          '可用阶段
Private mclsMipModule       As zl9ComLib.clsMipModule   '消息平台对象

Private Enum 执行方式
    T0无需执行 = 0
    T1必须生成 = 1
    T3必要时 = 3
End Enum

Private Enum TYPE_Func
    Func生成路径 = 0
    Func补充生成 = 1
    Func查看路径 = 2
    Func重新生成 = 3
End Enum

Public Function ShowMe(frmParent As Object, ByVal lngFun As Long, t_pati As TYPE_Pati, t_pp As TYPE_PATH_Pati, _
                        ByVal lng病人阶段ID As Long, ByVal lng天数 As Long, Optional ByVal lng项目ID As Long, _
                        Optional ByVal lng执行ID As Long, Optional ByVal lng时间进度 As Long, Optional ByVal bln提前 As Boolean = False) As Boolean
'参数：lng项目ID,lng执行ID=重新生成医嘱时才需传入
'      lng时间进度= mlngFun=0时传入，1=下一阶段提前,2-下一阶段提前至明天,-1=下一阶段延后（继续当前阶段）,0=正常
'      bln提前=true :提前生成路径,False-非提前生成
    Set mfrmParent = frmParent
    mlngFun = lngFun
    mlng项目ID = lng项目ID
    mlng执行ID = lng执行ID
    
    mPati = t_pati
    mPP = t_pp
    mlng病人阶段ID = lng病人阶段ID      '缺省选中当前阶段
    
    mlng天数 = lng天数
    
    mlng提前天数 = lng天数
    mlng时间进度 = lng时间进度
    If bln提前 Then                     '提前生成
        mdatDur = DateAdd("d", 1, CDate(Format(mPP.当前日期, "YYYY-MM-DD 00:00:00")))
    Else
        mdatDur = zlDatabase.Currentdate
    End If
    
    Set mrsPhase = GetPhase(mPP.路径ID, mPP.版本号, mlng病人阶段ID, mlng天数)
    
    mblnOK = False
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Function GetPhase(ByVal lng路径ID As Long, ByVal lng版本号 As Long, ByVal lng当前阶段ID As Long, ByVal lng天数 As Long) As ADODB.Recordset
'功能：读取当前时间可用的阶段
'1；2-7；8-12；13-19；20-30
'可用阶段：当前就诊天数对应的阶段；如果当前时间为第一天，则只显示第一阶段，如果
    Dim strSql As String, strIF As String, str阶段分类 As String
    Dim rsTmp As ADODB.Recordset, datPathIn As Date, lng时间进度 As Long
    Dim lng序号 As Long
    Dim strMainIF As String
    
    If mlngFun = 2 Then         '查看阶段定义的项目
        strSql = " Select a.Id,Nvl(a.父id,0) as 父id,a.序号,a.名称,a.说明,a.开始天数,a.结束天数,a.分类" & vbNewLine & _
                 " From 门诊路径阶段 A" & vbNewLine & _
                 " Where a.路径id = [1] And a.版本号 = [2] And a.id = [4]" & vbNewLine & _
                 " Order by 序号"
    Else
        datPathIn = GetPatiInPathOut(mPP.病人路径ID)                                            '获取病人的进入路径的开始时间
        
        If mlngFun = 0 Then
            If mlng时间进度 = -1 Then                     '延后时继续当前阶段
                strIF = " And a.id = [4]"
            Else
                If mPP.当前阶段ID <> 0 Then
                    lng序号 = GetPhaseNOOut(mPP.当前阶段ID)
                End If
                
                If mlng时间进度 = 1 Or mlng时间进度 = 2 Then
                    strIF = " And NVL(d.序号,a.序号)>[6] "
                Else
                    If mPP.当前阶段ID <> 0 Then
                        '之前可能有提前执行过的阶段的时间范围在当前天数内，要排除那些阶段，路径跳转时不检查。
                        strIF = " And NVL(d.序号,a.序号)>=[6] "
                    End If
                    
                     '同一天有多个阶段时，当前阶段及分支不能再用,如果是进入下一天了，则说明没有相同天数的阶段
                    If lng天数 = mPP.当前天数 Then
                        strIF = strIF & " And Nvl(a.父id,a.id) <> " & IIf(mPP.阶段父ID <> 0, "[7]", "[4]")
                    End If
                End If
                
                str阶段分类 = Get阶段分类Out(mPP.病人路径ID)
                If str阶段分类 <> "" Then
                    strIF = strIF & " And (a.父id is Null Or a.父id is Not Null And a.分类 = [5])"
                End If

                strMainIF = strIF
                
                'strIF = strIF & " And (a.开始天数 Is Null Or [3] Between a.开始天数 And Nvl(a.结束天数,a.开始天数) " & ")"
            End If
        Else
            strIF = " And a.id = [4]"
        End If
      
        strSql = " Select a.Id, Nvl(a.父id, 0) As 父id, a.序号, a.名称, a.说明, a.开始天数, a.结束天数, a.分类" & vbNewLine & _
                 " From 门诊路径阶段 A, 门诊路径阶段 D" & vbNewLine & _
                 " Where a.父id = d.Id(+) And a.路径id = [1] And a.版本号 = [2]" & vbNewLine & _
                   strIF & " Order By Nvl(d.序号, a.序号)"
    End If
    On Error GoTo errH
    Set GetPhase = zlDatabase.OpenSQLRecord(strSql, "获取适用阶段", lng路径ID, lng版本号, lng天数, lng当前阶段ID, str阶段分类, lng序号, mPP.阶段父ID)
    
    If (mlng时间进度 = 1 Or mlng时间进度 = 2) And GetPhase.RecordCount = 0 Then
        '阶段提前时，如果当前阶段有多天，则按当前天数取不到下一阶段，直接取序号大于当前阶段的下一阶段
        strSql = " Select * From (Select a.Id, Nvl(a.父id,0) as 父id, a.序号, a.名称, a.说明,a.开始天数, a.结束天数, a.分类" & vbNewLine & _
                 " From 门诊路径阶段 A,门诊路径阶段 D " & vbNewLine & _
                 " Where a.父ID=d.id(+) and a.路径id = [1] And a.版本号 = [2]" & _
                  strMainIF & vbNewLine & " Order by NVL(d.序号,a.序号)) Where Rownum=1"
        Set GetPhase = zlDatabase.OpenSQLRecord(strSql, "获取适用阶段", lng路径ID, lng版本号, lng天数, lng当前阶段ID, str阶段分类, lng序号)
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

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
            If lngW < lblFont.Width Then
                lngW = lblFont.Width
            End If
           .AddItem arrtmp(i)
        Next
        lngH = (i - 1) * 210 + 240
        If lngH > 1080 Then lngH = 1080
        lngW = lngW + 700
        If lngW > 2500 Then lngW = 2500
    End With

    With vsItem
        strSelect = .TextMatrix(.Row, .Col)
        For i = 0 To lstPati.ListCount - 1
            If InStr("|" & strSelect & "|", "|" & lstPati.List(i) & "|") > 0 Then
                lstPati.Selected(i) = True
            End If
        Next
        If lngW < .ColWidth(mcol("婴儿")) Then
            lngW = .ColWidth(mcol("婴儿"))
        End If
        lstPati.Move .Left + .ColPos(.Col), .Top + .RowPos(.Row) + .RowHeight(.Row) + 30, lngW, lngH
    End With
    Call lstPati.SetFocus
End Sub

Private Sub Form_Load()
    If mlngFun <> 2 Then
        vsItem.Editable = flexEDKbdMouse
    End If
    
    vsItem.Top = vsPhase.Top + vsPhase.Height + 45
    
    Call LoadPhase                          '加载可选择的阶段
    
    mlng路径医嘱天数 = Val(zlDatabase.GetPara("路径医嘱生成超前天数", glngSys, P门诊路径应用, "1"))
    
    If vsPhase.Cols = 1 Then
        vsPhase.Visible = False
        lblPhase.Caption = vsPhase.TextMatrix(0, 0) & vbCrLf & vsPhase.Cell(flexcpData, 0, 0)
                
        vsItem.Top = vsPhase.Top
        vsItem.Height = picBottom.Top - vsItem.Top
    Else
        lblNote.Visible = False
        lblPhase.Left = lblNote.Left
    
        If Grid.HScrollVisible(vsPhase) Then
            '横向滚动条
            vsPhase.Height = 1000
            vsItem.Height = vsItem.Height - (vsPhase.Top + vsPhase.Height - vsItem.Top + 120)
            vsItem.Top = vsPhase.Top + vsPhase.Height + 60
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
        dtpAdviceTime.Visible = False
        lblAdviceTime.Visible = False
    Else
        '医嘱缺省时间默认取当前时间
        dtpAdviceTime.Value = mdatDur
        lblDate.Caption = "生成路径项目日期：" & Format(mdatDur, "yyyy-MM-dd") & ",路径第" & mlng天数 & "天"
        mstrBaby = GetBabyRegList
    End If
    
    Call InitItem
    
    If mlngFun = 2 Then                                         '查看时只显示分类 , 项目内容
        Me.Width = vsItem.Width + 360
        cmdCancel.Left = vsItem.Left + vsItem.Width - 1200
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 150
    End If
    
    Set mEditType = New Collection
    Call LoadItem(Val(vsPhase.ColData(vsPhase.Col)), vsItem, mPP.路径ID, mPP.版本号)
    
    If vsItem.Rows = 1 Then
        vsItem.Rows = 2
        vsItem.TextMatrix(1, mcol("项目内容")) = "没有必须执行或可选性的路径项目"
        cmdOK.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsPhase = Nothing
    Set mcol = Nothing
    Set mEditType = Nothing
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
        strList = lstPati.List(0)                           '缺省选中病人本人,不允许为空
    End If
    
    vsItem.TextMatrix(vsItem.Row, vsItem.Col) = strList
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

Private Sub vsItem_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    With vsItem
        If cmdPati.Visible Then
            cmdPati.Move .Left + .ColPos(.Col) + .ColWidth(.Col) - 255, .Top + .RowPos(.Row) + 15, 255, 240
            lstPati.Visible = False
        End If
    End With
End Sub

Private Sub vsItem_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    With vsItem
        If .Col = mcol("婴儿") And cmdPati.Visible Then
            cmdPati.Move .Left + .ColPos(.Col) + .ColWidth(.Col) - 255, .Top + .RowPos(.Row) + 15, 255, 240
            lstPati.Visible = False
        End If
    End With
End Sub

Private Sub vsItem_Click()
    With vsItem
        If lstPati.Visible Then
            lstPati.Visible = False
        End If
    End With
End Sub

Private Sub vsItem_DblClick()
    Dim lng项目ID As Long
    
    If vsItem.Col = mcol("项目内容") Then
        lng项目ID = Val(vsItem.TextMatrix(vsItem.Row, mcol("ID")))
        If lng项目ID <> 0 Then
            Call frmPathItemEditOut.ShowView(mfrmParent, lng项目ID)
        End If
    End If
End Sub

Private Sub vsItem_GotFocus()
    vsItem.ForeColorSel = vbWhite
    vsItem.BackColorSel = &H8000000D
End Sub

Private Sub vsItem_LostFocus()
    vsItem.ForeColorSel = vbBlack
    vsItem.BackColorSel = vbWhite
End Sub

Private Sub vsItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call ResultEnterNextCell(vsItem)
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

Private Sub vsItem_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    
    With vsItem
        If Col = mcol("选择") Then
            If .Cell(flexcpChecked, Row, mcol("选择")) = 2 Then
                '未选择时弹出变异原因选择
                If mlngFun = 0 Then
                    If .RowData(Row) = 执行方式.T1必须生成 Then
                        Call vsItem_CellButtonClick(Row, mcol("变异原因"))
                    End If
                End If
            ElseIf .Cell(flexcpChecked, Row, mcol("选择")) = 1 Then
                If .RowData(Row) = 执行方式.T1必须生成 Then
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
                    If Not (.RowData(i) = 执行方式.T0无需执行 Or .RowData(i) = 执行方式.T1必须生成) Then
                        If .Cell(flexcpChecked, i, mcol("选择")) = 1 Then
                            .Cell(flexcpChecked, i, mcol("选择")) = 2
                        End If
                    End If
                Next
                For i = Row - 1 To .FixedRows Step -1
                    If .TextMatrix(i, mcol("分类")) <> .TextMatrix(Row, mcol("分类")) Then Exit For
                    If Not (.RowData(i) = 执行方式.T0无需执行 Or .RowData(i) = 执行方式.T1必须生成) Then
                        If .Cell(flexcpChecked, i, mcol("选择")) = 1 Then
                            .Cell(flexcpChecked, i, mcol("选择")) = 2
                        End If
                    End If
                Next
                
            ElseIf .Cell(flexcpChecked, Row, mcol("全选")) = 1 Then
                For i = Row To .Rows - 1
                    If .TextMatrix(i, mcol("分类")) <> .TextMatrix(Row, mcol("分类")) Then Exit For
                    If .Cell(flexcpChecked, i, mcol("选择")) = 2 Then
                        .Cell(flexcpChecked, i, mcol("选择")) = 1
                        If .RowData(i) = 执行方式.T1必须生成 Then
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
                        If .RowData(i) = 执行方式.T1必须生成 Then
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

Private Sub vsItem_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsItem
        If NewRow >= .FixedRows And Me.Visible Then
            If mlngFun = 0 Then
                If NewCol = mcol("变异原因") Then
                    '未选择时，可设置或选择变异原因
                    If .RowData(NewRow) = 执行方式.T1必须生成 And .Cell(flexcpChecked, NewRow, mcol("选择")) = 2 Then
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
                    If .RowData(NewRow) = 执行方式.T0无需执行 Then
                        cmdPati.Enabled = False
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub vsItem_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsItem
        If Col = mcol("选择") Then
            '生成路径时，每天生成的，可以不选，但要求输变异原因
            If .RowData(Row) = 执行方式.T0无需执行 Or mlngFun <> 0 And .RowData(Row) = 执行方式.T1必须生成 Then
                Cancel = True
            End If
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

Private Sub vsItem_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strSql As String, blnCancel As Boolean
    Dim rsTmp As ADODB.Recordset
            
    With vsItem
        If Col = mcol("变异原因") Then
            strSql = "Select b.名称 as 分类,a.编码 as ID,a.编码,a.名称,a.简码 From 门诊变异常见原因 a,门诊变异常见原因 b" & _
                    " Where a.性质=1 And a.末级=1 And a.上级=b.编码 And b.末级=0 " & _
                    " Order by 分类,a.编码"
            
            Set rsTmp = zlDatabase.ShowSelect(Me, strSql, 0, "门诊变异常见原因", True, , , True, True, True, _
                     Me.Left + .Left + .ColPos(Col), Me.Top + .Top + .RowPos(Row) + .RowHeight(Row) * 2, .RowHeight(Row), blnCancel, False, True)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "系统没有初始门诊变异常见原因，请与系统管理员联系。", vbInformation, gstrSysName
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
        mrsPhase.Filter = "ID=" & lng阶段ID
        mlng天数 = (mrsPhase!开始天数 & "")
        Call LoadItem(lng阶段ID, vsItem, mPP.路径ID, mPP.版本号)
    End If
End Sub

Private Sub LoadPhase()
'功能：加载可选择的阶段,如果病人的当前时间阶段仍然可用，则选中，否则缺省为第一个
    Dim i As Long, j As Long, str阶段分类 As String
    Dim rsNode As ADODB.Recordset

    With vsPhase
        .Clear
        .Redraw = flexRDNone
        .Col = -1
        mrsPhase.Filter = ""
        .Cols = mrsPhase.RecordCount
        str阶段分类 = Get阶段分类Out(0, mPP.当前阶段ID)
        If mlngFun = 0 And mlng时间进度 <> -1 Then '补充生成、重新生成时、下一阶段延后（继续当前阶段），只有当前阶段的记录
            mrsPhase.Filter = "父ID<>0 "
            If mrsPhase.RecordCount > 0 Then    '有备用分支
                Set rsNode = mrsPhase.Clone
                .Rows = 2
                .MergeRow(0) = True
            Else
                .Rows = 1
            End If
            mrsPhase.Filter = "父ID=0"
        End If
    
        For i = 0 To .Cols - 1
            .ColWidth(i) = 2000
            .ColAlignment(i) = flexAlignCenterCenter
            .TextMatrix(0, i) = mrsPhase!名称
            .Cell(flexcpData, 0, i) = CStr(IIf(IsNull(mrsPhase!分类), "", "分类：" & mrsPhase!分类 & " ") & mrsPhase!说明)
            .ColData(i) = Val(mrsPhase!ID)
            
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
                        .TextMatrix(0, i) = mrsPhase!名称           '第一行设置相同内容用于合并
                        .TextMatrix(1, i) = IIf(IsNull(rsNode!说明), "分支" & j, "" & rsNode!说明)
                        .Cell(flexcpData, 1, i) = CStr(IIf(IsNull(rsNode!分类), "", "分类：" & rsNode!分类 & " ") & rsNode!说明)
                        
                        .ColData(i) = Val(rsNode!ID)
                        If .ColData(i) = mlng病人阶段ID Then
                            .Col = i
                        ElseIf .Col = 0 And str阶段分类 <> "" Then
                            If str阶段分类 = "" & rsNode!分类 Then
                                .Col = i
                            End If
                        End If
                        rsNode.MoveNext
                    Next
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

Private Sub LoadItem(lng阶段ID As Long, objVsg As VSFlexGrid, ByVal lng路径ID As Long, ByVal lng版本号 As Long)
'功能：加载当前阶段的路径项目
'参数：objVsg，需要加载的表格
    Dim i As Long, j As Long, blnFocus As Boolean
    Dim rsTmp As ADODB.Recordset, strSql As String, strIDs As String, strTmp As String
    Dim rsAdvice As ADODB.Recordset, rsFile As ADODB.Recordset
    Dim lngRow As Long
    Dim str诊疗项目IDs As String
    Dim lng首要路径阶段ID As Long
    Dim strNewTmp As String
     
    If mlngFun = 1 Then '补充生成，无需执行的不显示，当天执行过的不能重复生成,只执行一次的当前阶段已执行则不显示
        strSql = " And a.执行方式<>0 And Not Exists(Select 1 From 病人门诊路径执行 c " & _
                 " Where c.路径记录id = [4] And c.阶段id = [7] And c.项目id = a.id And (c.天数 = [5] and a.执行方式<>4 or a.执行方式=4))"
        lng首要路径阶段ID = mPP.当前阶段ID
    ElseIf mlngFun = 3 Then '重新生成
        strSql = " And a.ID = [6]"
    Else
        strSql = " And (a.执行方式<>4 or a.执行方式=4 And Not Exists(Select 1 From 病人门诊路径执行 c " & _
                 " Where c.路径记录id = [4] And c.阶段id = [7] And c.项目id = a.id))"
        lng首要路径阶段ID = lng阶段ID
    End If
    '连接“门诊路径分类”，只是为了按分类排序'保存时再检查，是否为当天阶段的最后一天，至少执行一次的项目是否选择
    strSql = " Select a.分类, a.ID, a.项目内容, a.执行方式, a.图标id, a.内容要求" & vbNewLine & _
             " From 门诊路径项目 A, 门诊路径分类 B" & vbNewLine & _
             " Where a.分类 = b.名称 And a.路径id = b.路径id And a.版本号 = b.版本号 And a.路径id = [1] And a.版本号 = [2] And a.阶段id = [3] " & vbNewLine & _
               strSql & IIf(mlngFun = 3, "", "Order By b.序号, a.项目序号")
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng路径ID, lng版本号, lng阶段ID, mPP.病人路径ID, mlng天数, mlng项目ID, lng首要路径阶段ID)
    
    With objVsg
        .Redraw = flexRDNone
        .Rows = .FixedRows
        .Rows = .FixedRows + rsTmp.RecordCount
        lngRow = 1
        '由于固定列合并不能影响后面的列，所以新增了一列判断按分类合并全选列
        .MergeCells = flexMergeRestrictAll
        .MergeCol(mcol("分类")) = True
        .MergeCol(mcol("分类值")) = True
        .MergeCol(mcol("全选")) = True

        For i = lngRow To rsTmp.RecordCount + lngRow - 1
            .TextMatrix(i, mcol("ID")) = rsTmp!ID
            strIDs = strIDs & "," & rsTmp!ID
            .TextMatrix(i, mcol("分类")) = rsTmp!分类
            .TextMatrix(i, mcol("分类值")) = rsTmp!分类
            .TextMatrix(i, mcol("项目内容")) = rsTmp!项目内容
            
            If mlngFun <> 2 Then
                .TextMatrix(i, mcol("内容要求")) = Val("" & rsTmp!内容要求)
            End If
            
            .TextMatrix(i, mcol("执行方式")) = Decode(rsTmp!执行方式, 0, "无", 1, "必须", 3, "必要时")
            .RowData(i) = Val(rsTmp!执行方式)
            
            If mlngFun <> 2 Then
                Select Case rsTmp!执行方式
                    Case 执行方式.T0无需执行
                        .TextMatrix(i, mcol("选择")) = " "
                        .Cell(flexcpBackColor, i, mcol("选择")) = &H8000000F
                    Case 执行方式.T1必须生成
                        .Cell(flexcpChecked, i, mcol("选择")) = 1
                        .Cell(flexcpPictureAlignment, i, mcol("选择")) = flexPicAlignCenterCenter
                        If mlngFun <> 0 Then
                            .Cell(flexcpBackColor, i, mcol("选择")) = &H8000000F
                        End If
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
        Set rsAdvice = GetAdviceOut(strIDs)
        If rsAdvice.RecordCount > 0 Then
            For i = .FixedRows To .Rows - 1
                rsAdvice.Filter = "路径项目ID=" & ZVal(Val(.TextMatrix(i, mcol("ID"))))
                strTmp = ""
                str诊疗项目IDs = ""
                
                For j = 1 To rsAdvice.RecordCount
                    strTmp = strTmp & "," & rsAdvice!医嘱内容ID
                    str诊疗项目IDs = str诊疗项目IDs & "," & rsAdvice!诊疗项目ID
                    rsAdvice.MoveNext
                Next
                If strTmp <> "" Then
                    .TextMatrix(i, mcol("医嘱内容ID")) = Mid(strTmp, 2)
                    If mlngFun <> 2 Then
                        .TextMatrix(i, mcol("诊疗项目ID")) = Mid(str诊疗项目IDs, 2)
                    End If
                    .TextMatrix(i, mcol("项目内容")) = .TextMatrix(i, mcol("项目内容")) & " ……"
                End If
            Next
        End If
        
        '加载项目对应的病历文件
        If mlngFun <> 3 Then
            Set rsFile = GetFile(strIDs, 1)
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
        If .Rows = .FixedRows Then
            .Rows = .Rows + 1
        End If
        .Redraw = True
    End With
               
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
        strcol = "分类,1200,4;分类值;全选,450,4;项目内容,5950,1;执行方式;选择;婴儿;ID;医嘱内容ID;文件ID"
    Else
        strcol = "分类,1200,4;分类值;全选" & IIf(mlngFun <> Func重新生成, ",450,4", "") & ";项目内容," & IIf(mstrBaby = "", 5950, 4950) & ",1;执行方式,900,1" & _
                ";选择" & IIf(mlngFun <> Func重新生成, ",500,4", "") & _
                ";婴儿" & IIf(mstrBaby = "", "", ",1100,1") & _
                ";ID;医嘱内容ID;文件ID;内容要求;变异原因" & IIf(mlngFun = 0, ",1800,4", "") & ";是否最后一天;阶段ID;诊疗项目ID;重复项目"
    End If
    arrHead = Split(strcol, ";")
    Set mcol = New Collection
   
    With vsItem
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
'返回："姓名1,姓名2,姓名3…
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select 序号,婴儿姓名 From 病人新生儿记录 Where 病人ID=[1] And 主页ID=[2] Order by 序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetBabyRegList", mPati.病人ID, mPati.挂号ID)
    
    strSql = ""
    Do While Not rsTmp.EOF
        strSql = IIf(strSql = "", "", strSql & "|") & "婴儿:" & NVL(Replace(rsTmp!婴儿姓名, "|", "_"))
        rsTmp.MoveNext
    Loop
    GetBabyRegList = strSql
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
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
        GetBabyIndex = "0"                  '没有婴儿缺省取病人本人
    End If
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
    Dim rsUsed As ADODB.Recordset                   '重新生成时校对但未作废的医嘱
    Dim blnHaveDoc As Boolean
    Dim str项目IDs As String
    Dim k As Long, n As Long, strAgain As String
    Dim colItem As New Collection
    Dim strAgaignTmp As String
    Dim str路径项目IDs As String                    '路径生成时中医修改了的配方的，且超出了允许修改配方的比例的项目，对应的变异原因：项目ID1|变异编码1,项目2|变异编码2····
    Dim colPathItems As New Collection
    
    arrSQL = Array()
    '1.检查必须执行一次的项目
    With vsItem
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

            '必须生成的项目，如果没有选择，则必须输入变异原因
            For i = 1 To .Rows - 1
                If .RowData(i) = 执行方式.T1必须生成 Then
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
    
    '重新生成医嘱：重新生成项目对应的医嘱为新开状态；删除该项目对应的所有医嘱,重新产生医嘱记录数据;存在发送未作废的医嘱必须要求作废。
    If mlngFun = Func重新生成 Then
        Set rsLastAdvice = GetUsedAdvice(mlng执行ID, mlng项目ID)
    End If
    
    strIDs = ""
    With vsItem
        '必须生成的，如果选择不生成医嘱（选择了变异原因），则要生成路径项目
        If mlngFun = 0 Then
            For i = 1 To .Rows - 1
                If .RowData(i) = 执行方式.T1必须生成 Then
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
                    If Trim(.TextMatrix(i, mcol("文件ID"))) <> "" Then
                        blnHaveDoc = True
                    End If
                End If
            End If
        Next
    End With
    strIDs = Mid(strIDs, 2)             '医嘱内容ID:婴儿序号:路径项目ID,...，例：227:0:38,335:1:69
    
    If blnHaveDoc Then
        If InStr(GetInsidePrivs(p住院病历管理), ";病历书写;") = 0 Then
            MsgBox "你没有病历书写的权限，不能生成包含病历的路径项目。", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
    End If
    
    '产生医嘱的缺省开始执行时间
    DatCurr = mdatDur
    
    If strIDs <> "" Then    '全是无需执行的项目时不产生医嘱，但要产生路径执行项目
        If InStr(GetInsidePrivs(p门诊医嘱下达), ";医嘱下达;") = 0 Then
            MsgBox "你没有医嘱下达的权限，不能生成包含医嘱的路径项目。", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
        '检查时间
        If Format(DatCurr, "YYYY-MM-DD") > Format(dtpAdviceTime.Value, "YYYY-MM-DD") Or Format(dtpAdviceTime.Value, "YYYY-MM-DD") > Format(DatCurr + mlng路径医嘱天数, "YYYY-MM-DD") Then
            MsgBox "门诊临床路径的医嘱必须在当前日期和提前的天数之间，当前允许提前" & mlng路径医嘱天数 & "天。", vbInformation, gstrSysName
            If dtpAdviceTime.Enabled And dtpAdviceTime.Visible Then
                dtpAdviceTime.SetFocus
            End If
            Exit Sub
        End If
        
        Me.Hide
        If gobjKernel.ShowOutAdviceEdit(mfrmParent, 0, 1, mPati.病人ID, mPati.挂号NO, strIDs, CDate(dtpAdviceTime.Value), arrSQL, strAdviceOfItem, rsLastAdvice, DatCurr, str路径项目IDs, mclsMipModule, , mPati.科室ID) = False Then
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
            For i = 1 To vsItem.Rows - 1
                strTmp = ""
                If vsItem.TextMatrix(i, mcol("ID")) & "" <> "" Then
                    strTmp = colPathItems("_" & vsItem.TextMatrix(i, mcol("ID")))
                    If strTmp <> "" Then
                        vsItem.Cell(flexcpData, i, mcol("变异原因")) = strTmp
                    End If
                End If
            Next
            On Error GoTo 0
        End If
    End If
    
    Call SaveData(arrSQL, strAdviceOfItem, lng天数)
    '修正医嘱序号
    Call ModifyAdviceSerialNum
    mblnOK = True
    Unload Me
End Sub

Private Sub ModifyAdviceSerialNum()
'功能：重新整理医嘱序号
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String

    On Error GoTo errH
    Screen.MousePointer = 11
    strSql = "Select Count(*) as Num From (Select 序号,Count(ID) From 病人医嘱记录 Where 病人ID=[1] And 挂号单=[2] Having Count(ID)>1 Group by 序号)"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "查询病人医嘱数量", mPati.病人ID, mPati.挂号NO)

    If rsTmp.EOF Then
        Screen.MousePointer = 0
        Exit Sub
    End If

    If NVL(rsTmp!Num, 0) = 0 Then
        Screen.MousePointer = 0
        Exit Sub
    End If

    strSql = "ZL_病人医嘱记录_更新序号(NULL,NULL," & mPati.病人ID & ",'" & mPati.挂号NO & "')"
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

Private Function GetUsedAdvice(ByVal lng执行ID As Long, ByVal lng项目ID As Long) As ADODB.Recordset
'功能:重新生成时,返回当前项目对应的有效医嘱记录
    Dim strSql As String
    
    strSql = " Select [1] as 项目ID, a.病人医嘱id, Nvl(b.相关id, b.Id) As 组id, b.诊疗项目id " & vbNewLine & _
             " From 病人门诊路径医嘱 A, 病人医嘱记录 B" & vbNewLine & _
             " Where a.病人医嘱id = b.Id And a.路径执行id = [2] " & vbNewLine & _
             " Order By b.序号"
    On Error GoTo errH
    
    Set GetUsedAdvice = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng项目ID, lng执行ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetLastEvaluate(strLastVariation As String, str审核人 As String, str评估人 As String)
'功能：获得最后一次评估的信息
    Dim strSql As String, rsTmp As Recordset
    
    On Error GoTo errH
    strSql = "Select 变异原因,变异审核人,评估人 From 病人门诊路径评估 Where 路径记录ID=[1] And 天数=[2] And 阶段ID=[3]"
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
    strSql = "Select 名称 from 门诊路径分类 Where 路径ID=[1] and 版本号=[2] and 序号=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "取第一个分类", mPP.路径ID, mPP.版本号)
    If rsTmp.RecordCount > 0 Then
        GetFirstType = rsTmp!名称 & ""
    End If
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
    Dim strEPR As String
    Dim blnAgain As Boolean
    Dim strAgaignTmp As String
    Dim strAgain As String
    Dim colItemName As New Collection
    Dim blnDef As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim arrtmp As Variant
    
    AddDate = zlDatabase.Currentdate
    strAddDate = "To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    arrItem = Split(strAdviceOfItem, ",")
    
    '如果一个项目都没有，则从数据库中取第一个分类
    If vsItem.TextMatrix(1, mcol("分类")) = "" Then
        strFirstType = GetFirstType
    Else
        strFirstType = vsItem.TextMatrix(1, mcol("分类"))
    End If

    lng天数 = mlng天数
    
    strDate = "To_Date('" & Format(mdatDur, "yyyy-MM-dd") & "','YYYY-MM-DD')"
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
            mrsPhase.Filter = "开始天数=" & i & " And 父ID = 0"
            If Not mrsPhase.EOF Then
                strSql = "Zl_病人门诊路径生成_Insert(1," & mPati.病人ID & "," & mPati.挂号ID & ",NULL," & mPati.科室ID & "," & _
                        mPP.病人路径ID & "," & mrsPhase!ID & _
                        "," & strDate & "," & mlng提前天数 & _
                        ",'" & strFirstType & "',Null" & _
                        ",Null,Null,Null,'" & UserInfo.姓名 & "',To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & _
                        "','YYYY-MM-DD HH24:MI:SS'),'未生成任何项目','已经执行|1" & vbTab & "已经执行')"
                colSQL.Add strSql, "C" & colSQL.count + 1
                
                '登记时间加一秒，是为了取消生成时取上一个阶段的ID。
                AddDate = AddDate + 1 / 24 / 60 / 60
                '评估
                strSql = "Zl_病人门诊路径评估_Insert(1," & mPP.病人路径ID & "," & mrsPhase!ID & _
                        "," & strDate & "," & mlng提前天数 & ",'" & _
                        str评估人 & "',1,'','" & UserInfo.姓名 & "','" & str审核人 & "','" & strLastVariation & "',1,Null,1)"
                        
                colSQL.Add strSql, "C" & colSQL.count + 1
            End If
        Next
        mrsPhase.Filter = varFilter
    End If
        
    With vsItem
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, mcol("选择")) = 1 Or .Cell(flexcpChecked, i, mcol("选择")) = 2 And mlngFun = 0 And .RowData(i) = 执行方式.T1必须生成 Then
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

                        If str病人病历IDs = "" Then
                            blnAgain = True
                        End If
                    End If
                Else
                    strVariation = .Cell(flexcpData, i, mcol("变异原因"))
                End If
                
                If Not blnAgain Then
                    lng序号 = lng序号 + 1
                    strSql = "Zl_病人门诊路径生成_Insert(" & lng序号 & "," & mPati.病人ID & "," & mPati.挂号ID & ",'" & strBaby & "'," & mPati.科室ID & "," & _
                        mPP.病人路径ID & "," & mrsPhase!ID & "," & strDate & "," & mlng提前天数 & ",'" & .TextMatrix(i, mcol("分类")) & "'," & .TextMatrix(i, mcol("ID")) & _
                        ",'" & strAdviceIDs & "','" & strFileIDs & "','" & str病人病历IDs & "'" & _
                        ",'" & UserInfo.姓名 & "',To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),Null,Null,Null,Null,'" & strVariation & "')"
                    colSQL.Add strSql, "C" & colSQL.count + 1
                    blnIsSend = True
                End If
            End If
        Next
    End With
    '如果没有勾选任何项目，则生成一条特殊的项目：未生成任何项目
    If Not blnIsSend Then
        If mlngFun = 0 Then
            lng序号 = lng序号 + 1
            strSql = "Zl_病人门诊路径生成_Insert(" & lng序号 & "," & mPati.病人ID & "," & mPati.挂号ID & ",NULL," & mPati.科室ID & "," & _
                    mPP.病人路径ID & "," & mrsPhase!ID & "," & strDate & "," & mlng提前天数 & ",'" & strFirstType & "',Null" & _
                    ",Null,Null,Null,'" & UserInfo.姓名 & "',To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'未生成任何项目','已经执行|1" & vbTab & "已经执行')"
            colSQL.Add strSql, "C" & colSQL.count + 1
        End If
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        If mlngFun = 3 Then
            strSql = "Zl_病人门诊路径生成_Delete(" & mlng执行ID & ",1)"
            Debug.Print strSql & vbCrLf
            zlDatabase.ExecuteProcedure strSql, Me.Caption
        End If
        
        '1.先产生医嘱,因为病人路径医嘱有外键
        For i = 0 To UBound(arrSQL)
            Debug.Print CStr(arrSQL(i)) & vbCrLf
            zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
        Next
        '2.产生病人路径数据，以及病历文件数据
        For i = 1 To colSQL.count
            Debug.Print colSQL("C" & i) & vbCrLf
            Call zlDatabase.ExecuteProcedure(colSQL("C" & i), Me.Caption)
        Next
        '3.产生病历文件RTF数据
        For i = 1 To colDoc.count
            arrItem = Split(colDoc("C" & i), ":")
            If arrItem(2) = 0 Or arrItem(2) = 1 Then     '全文编辑方式的病历
                lng病历ID = (arrItem(0))
                Call ReadRTFData(lng病历ID, edtEditor)
                Call SaveRTFData(lng病历ID, mPati.病人ID, mPati.挂号ID, Val(arrItem(1)), edtEditor, 1)
            End If
        Next
    gcnOracle.CommitTrans: blnTrans = False

    Exit Sub
errH:
    If blnTrans Then
        gcnOracle.RollbackTrans
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
