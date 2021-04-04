VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmDiagEdit 
   BackColor       =   &H00EFF0E0&
   Caption         =   "诊断选择及编辑"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10755
   Icon            =   "frmDiagEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleMode       =   0  'User
   ScaleWidth      =   10974.49
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picZY 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   360
      ScaleHeight     =   3855
      ScaleWidth      =   9615
      TabIndex        =   5
      Top             =   480
      Width           =   9615
      Begin VSFlex8Ctl.VSFlexGrid vsDiagZY 
         Height          =   3675
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   9495
         _cx             =   16748
         _cy             =   6482
         Appearance      =   1
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
         BackColorFixed  =   14811105
         ForeColorFixed  =   -2147483630
         BackColorSel    =   4210752
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDiagEdit.frx":6852
         ScrollTrack     =   -1  'True
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
   End
   Begin VB.PictureBox picXY 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   120
      ScaleHeight     =   3615
      ScaleWidth      =   9615
      TabIndex        =   3
      Top             =   600
      Width           =   9615
      Begin VSFlex8Ctl.VSFlexGrid vsDiagXY 
         Height          =   3465
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   9495
         _cx             =   16748
         _cy             =   6112
         Appearance      =   1
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
         BackColorFixed  =   14811105
         ForeColorFixed  =   -2147483630
         BackColorSel    =   4210752
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   9
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDiagEdit.frx":6A47
         ScrollTrack     =   -1  'True
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
   End
   Begin VB.Frame fraInput 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   350
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   7920
      Begin VB.OptionButton optInput 
         BackColor       =   &H00EFF0E0&
         Caption         =   "根据诊断标准输入(&1)"
         ForeColor       =   &H00004000&
         Height          =   180
         Index           =   0
         Left            =   3840
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   37
         Value           =   -1  'True
         Width           =   2010
      End
      Begin VB.OptionButton optInput 
         BackColor       =   &H00EFF0E0&
         Caption         =   "根据疾病编码输入(&2)"
         ForeColor       =   &H00004000&
         Height          =   180
         Index           =   1
         Left            =   5880
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   37
         Width           =   2010
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      Height          =   650
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   10755
      TabIndex        =   8
      Top             =   5580
      Width           =   10755
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   8520
         TabIndex        =   10
         Top             =   150
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   7200
         TabIndex        =   9
         Top             =   150
         Width           =   1100
      End
      Begin VB.Image imgButtonNew 
         Height          =   240
         Left            =   720
         Picture         =   "frmDiagEdit.frx":6C4E
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgButtonDel 
         Height          =   240
         Left            =   0
         Picture         =   "frmDiagEdit.frx":71D8
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin XtremeSuiteControls.TabControl tbcMain 
      Height          =   4095
      Left            =   60
      TabIndex        =   7
      Top             =   360
      Width           =   9735
      _Version        =   589884
      _ExtentX        =   17171
      _ExtentY        =   7223
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmDiagEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'入口参数
Private mlng病人ID As Long
Private mlng就诊ID As Long
Private mint病人来源 As Integer
Private mlngCur标识 As Long
Private mlng科室ID As Long
Private mstr开单人 As String
Private mstr诊断IDs As String
Private mstr诊断s As String
Private mlng组医嘱ID As String
Private mblnOK As Boolean
'其他变量
Private mint中医诊断输入 As Integer
Private mint西医诊断输入 As Integer
Private mstrPrivs As String
Private mblnChange As Boolean
Private mlngPathState As Long
Private mlngDiagnosisType As Long
Private mstrPathDiag As String
Private mblnIsPathOutTime As Boolean
Private mstr性别 As String
Private mblnReturn As Boolean
Private mint简码 As Integer
Private mstr诊断输入 As String
Private mstrLike As String
Private mint险类 As Integer
Private mstr挂号单 As String
Private mbln手术 As Boolean '病人是否进行过手术
Private mlng损伤中毒 As Long
Private mlng病理诊断 As Long
Private mbln中医 As Boolean
Private mbytSize As Byte

Private Const M_LNG_P住院医生站 = 1261
Private Const M_LNG_P门诊医生站 = 1260
Private Const M_LNG_SYS = 100
Private Const ColorUnEditCell = &H8000000B  '灰蓝色
Private mrsAdvice As ADODB.Recordset

Private Enum COL诊断情况
    col诊断类型 = 0
    col关联 = 1
    col诊断编码 = 2
    Col诊断描述 = 3
    col中医证候 = 4
    col发病时间 = 5
    col备注 = 6
    col入院病情 = 7
    col出院情况 = 8
    col是否未治 = 9
    col是否疑诊 = 10
    col增加 = 11
    colDel = 12
    col诊断ID = 13
    col疾病ID = 14
    col类型 = 15 '1-西医门诊诊断;2-西医入院诊断;3-出院诊断(其他诊断);5-院内感染;6-病理诊断;7-损伤中毒码;10-并发症
    
    colZY疑诊 = 9
    colZY增加 = 10
    colZYDel = 11
    colzy诊断ID = 12
    colzy疾病ID = 13
    colzy证候ID = 14
    colzy类型 = 15
End Enum

Public Function ShowMe(frmParent As Object, ByVal lng标识ID As Long, ByVal lng病人ID As Long, ByVal lng就诊ID As Long, ByVal int病人来源 As Integer, ByVal lng开单科室ID As Long, ByVal str开单人 As String, _
                    ByRef str诊断IDs As String, ByRef str诊断S As String, ByVal bytSize As Byte, Optional ByVal lng医嘱组ID As Long) As Boolean
'参数：lng病人ID=病人ID
'      lng就诊ID=住院:主页ID,门诊：挂号单ID
'      int病人来源=1-门诊，2-住院
'      lng开单科室ID=病人所在科室，诊断使用
'      lng标识ID =用于区分各个申请单的标识，用于保存相应的诊断
'      str开单人=操作员姓名，诊断登记人
'      str诊断IDs=该申请单相关的诊断ID,多个诊断时诊断ID以逗号分割
'      bytSize=0-9号字体，1-12号字体
'返回： ShowDiagEdit= 是确定还是取消
'       str诊断S=返回诊断描述字符串，供申请单使用
'       str诊断IDs=该申请单选择的相关的诊断ID,多个诊断时诊断ID以逗号分割
    mlng病人ID = lng病人ID
    mlng就诊ID = lng就诊ID
    mint病人来源 = int病人来源
    mlngCur标识 = lng标识ID
    mlng科室ID = lng开单科室ID
    mstr开单人 = str开单人
    mstr诊断IDs = str诊断IDs
    mstr诊断s = str诊断S
    mlng组医嘱ID = lng医嘱组ID
    mbytSize = bytSize
    mstrPrivs = gobjComLib.GetPrivFunc(M_LNG_SYS, IIf(mint病人来源 = 2, M_LNG_P住院医生站, M_LNG_P门诊医生站))
    Show 1, frmParent

    str诊断IDs = mstr诊断IDs
    str诊断S = mstr诊断s
    ShowMe = mblnOK

End Function

Private Sub cmdCancel_Click()
    If vsDiagXY.Tag = "" Or vsDiagZY.Tag = "" And vsDiagZY.Visible Then
        If MsgBox("退出后你所做的修改将不会生效。是否退出？", vbYesNo + vbDefaultButton2 + vbInformation, Me.Caption) = vbYes Then
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub cmdOK_Click()
    If CheckData() Then
        Call SaveData
        mblnOK = True
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    '住院首页相关
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lngColWidth As Long
    
    On Error GoTo errH
    
    mblnOK = False
    mlngPathState = -1
    mstrLike = IIf(Val(gobjComLib.zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    
    strSQL = "Select A.险类,Nvl(A.路径状态,-1) 路径状态" & _
        " From 病案主页 A" & _
        " Where A.病人ID=[1] And A.主页ID=[2]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng就诊ID)
    If Not rsTmp.EOF Then
        mint险类 = NVL(rsTmp!险类, 0)
        'mlngPathState=-1:未导入,0-不符合导入条件，1-执行中，2-正常结束，3-变异结束
        mlngPathState = Val(rsTmp!路径状态 & "")
    End If
    
    strSQL = "Select 1 From 病人手麻记录  A Where  A.病人ID=[1] And A.主页ID=[2] "
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng就诊ID)
    mbln手术 = Not rsTmp.EOF
    '病人信息部份
    '---------------------------------------------------------------
    
    strSQL = "Select 性别 From 病人信息 Where 病人ID=[1]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
    mstr性别 = NVL(rsTmp!性别)
    If mint病人来源 <> 2 Then
        strSQL = "Select NO From 病人挂号记录 Where 病人id = [1] And ID = [2]"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng就诊ID)
        If Not rsTmp.EOF Then
            mstr挂号单 = rsTmp!NO & ""
        End If
    End If
    '诊断输入方式
    mstr诊断输入 = gobjComLib.zlDatabase.GetPara(65, M_LNG_SYS, , "11")
    mint简码 = Val(gobjComLib.zlDatabase.GetPara("简码方式")) '简码匹配方式：0-拼音,1-五笔
    mlng损伤中毒 = Val(gobjComLib.zlDatabase.GetPara("损伤中毒检查", M_LNG_SYS, M_LNG_P住院医生站, 2) & "")
    mlng病理诊断 = Val(gobjComLib.zlDatabase.GetPara("病理诊断检查", M_LNG_SYS, M_LNG_P住院医生站, 2) & "")
    
    If mlngPathState <> -1 Then
        '只处理首页中输入的诊断，以前没填的，缺省当作来自于“西医入院诊断”
        strSQL = "Select Nvl(诊断类型,2) as 诊断类型,NVL(疾病ID,0) As 疾病ID,NVL(诊断ID,0) as 诊断ID,状态 From 病人临床路径 Where 病人ID=[1] And 主页ID=[2] And (诊断来源 = 3 or 诊断来源 is null) Order By 导入时间"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng就诊ID)
        If rsTmp.RecordCount > 0 Then
            mlngDiagnosisType = rsTmp!诊断类型
            '如果有多条路径，则取第一条的状态
            If rsTmp.RecordCount >= 2 Then mlngPathState = Val(rsTmp!状态 & "")
            rsTmp.MoveNext
            Do While Not rsTmp.EOF
                mstrPathDiag = mstrPathDiag & "," & rsTmp!诊断类型 & "|" & rsTmp!疾病ID & "|" & rsTmp!诊断ID
                rsTmp.MoveNext
            Loop
            mstrPathDiag = Mid(mstrPathDiag, 2)
        Else
            mlngDiagnosisType = 0
        End If
        '完成路径的时间是否比出院诊断记录时间大()取第一条路径
        If mlngPathState = 2 Then
            strSQL = "Select Sign(Nvl(a.结束时间, Null)-Nvl(b.记录日期, Sysdate)) As 判断" & vbNewLine & _
                    "From 病人临床路径 A, (Select 病人id, 主页id, 记录日期 From 病人诊断记录 Where 记录来源 = 3 And 诊断次序 = 1 And 诊断类型 = [3]) B" & vbNewLine & _
                    " Where a.病人id = b.病人id(+) And a.主页id = b.主页id(+) And a.病人ID=[1] And A.主页ID=[2]" & _
                    " and a.导入时间=(Select Min(导入时间) From 病人临床路径 Where 病人ID=[1] and 主页ID=[2])"
            Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng就诊ID, IIf(mlngDiagnosisType > 10, 13, 3))
            If rsTmp.RecordCount > 0 Then
                mblnIsPathOutTime = Val(rsTmp!判断 & "") = 1
            Else
                mblnIsPathOutTime = False
            End If
        End If
    End If
    
    strSQL = "Select 1 From 部门性质说明 Where 工作性质='中医科' And 部门ID=[1]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng科室ID)
    mbln中医 = Not rsTmp.EOF


    If mbln中医 Then
        tbcMain.PaintManager.Color = xtpTabColorOffice2003
        tbcMain.PaintManager.ColorSet.ControlFace = &HEFF0E0
        Call tbcMain.InsertItem(0, "西医诊断", picXY.hwnd, 0)
        Call tbcMain.InsertItem(1, "中医诊断", picZY.hwnd, 0)
        tbcMain(0).Selected = True
    Else
        tbcMain.Enabled = False
        tbcMain.Visible = False
        vsDiagZY.Visible = False
        picZY.Visible = False
        vsDiagZY.Enabled = False
        If mint病人来源 = 2 Then
            If Val(gobjComLib.zlDatabase.GetPara("西医诊断输入", M_LNG_SYS, M_LNG_P住院医生站, 0, Array(optInput(0), optInput(1)), InStr(mstrPrivs, "参数设置") > 0)) = 0 Then
                optInput(0).value = True
            Else
                optInput(1).value = True
            End If
        ElseIf mint病人来源 = 1 Then
            optInput(Val(gobjComLib.zlDatabase.GetPara("门诊诊断输入", M_LNG_SYS, M_LNG_P门诊医生站, 0, Array(optInput(0), optInput(1)), InStr(mstrPrivs, "参数设置") > 0))).value = True
        End If
    End If
    
    If mstr诊断IDs = "" Then
        With grsDiagConn
            .Filter = "标识ID=" & mlngCur标识
            .Sort = "诊断ID"
            Do While Not .EOF
                mstr诊断IDs = mstr诊断IDs & IIf(mstr诊断IDs = "", "", ",") & !诊断ID
                .MoveNext
            Loop
        End With
    End If
    
    Call LoadData
    
    Call SetVSColHidden
    
    If mint病人来源 = 2 Then
        vsDiagXY.ColWidth(Col诊断描述) = vsDiagXY.ColWidth(Col诊断描述) - 1000
        Me.Width = Me.Width + 1500
        Me.Height = Me.Height + 927
    Else
        vsDiagXY.ColWidth(col发病时间) = vsDiagXY.ColWidth(col发病时间) + 400
        vsDiagZY.ColWidth(col发病时间) = vsDiagZY.ColWidth(col发病时间) + 400
        vsDiagXY.ColWidth(Col诊断描述) = vsDiagXY.ColWidth(Col诊断描述) - 200
        vsDiagZY.ColWidth(Col诊断描述) = vsDiagZY.ColWidth(Col诊断描述) - 200
    End If
    
    If mbytSize = 0 Then
        Me.Width = Me.Width - 2000
        Me.Height = Me.Height - 1236
        vsDiagXY.ColWidth(Col诊断描述) = vsDiagXY.ColWidth(Col诊断描述) + 600
        vsDiagZY.ColWidth(Col诊断描述) = vsDiagZY.ColWidth(Col诊断描述) + 1000
    End If
    
    If Not mbln中医 Then
        Me.Width = Me.Width + 400
    End If
    
    Call SetPublicFontSize(mbytSize)
    Call gobjComLib.zlControl.VSFSetFontSize(vsDiagXY, IIf(mbytSize = 0, 9, 12))
    Call gobjComLib.zlControl.VSFSetFontSize(vsDiagZY, IIf(mbytSize = 0, 9, 12))
    lngColWidth = 270 '防止列宽过大时新增删除按钮出现黑色阴影
    vsDiagXY.ColWidth(colDel) = lngColWidth
    vsDiagXY.ColWidth(col增加) = lngColWidth
    vsDiagZY.ColWidth(colZYDel) = lngColWidth
    vsDiagZY.ColWidth(colZY增加) = lngColWidth
    Exit Sub
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub Form_Resize()
    If Me.Width > 20000 Then
        Me.Width = 20000
    End If
    If Me.Height > 12000 Then
        Me.Height = 12000
    End If
    
    If Me.Width < 6000 Then
        Me.Width = 6000
    End If
    
    If Me.Height < 5000 Then
        Me.Height = 5000
    End If
    fraInput.Width = Me.Width
    fraInput.Left = 0

    tbcMain.Top = fraInput.Top + fraInput.Height - IIf(mbln中医, 120, 180)
    tbcMain.Height = picBottom.Top - tbcMain.Top
    tbcMain.Width = Me.Width - tbcMain.Left - 100
    If mbln中医 Then
        tbcMain.Top = tbcMain.Top - 200
        fraInput.Left = tbcMain.Left + IIf(mbytSize = 0, 1840, 2320)
    End If
    optInput(1).Left = Me.Width - fraInput.Left - optInput(1).Width - 240
    optInput(0).Left = optInput(1).Left - optInput(0).Width - 100
    
    picZY.Top = tbcMain.Top + 210
    picZY.Height = picBottom.Top - picZY.Top
    picZY.Width = tbcMain.Width - 180
    
    picXY.Top = picZY.Top
    picXY.Height = picZY.Height
    picXY.Width = picZY.Width
    
    '确定取消按钮位置设置
    If mbytSize = 1 Then
        cmdCancel.Top = 90
        cmdOK.Top = 90
    End If
    cmdCancel.Left = Me.Width - cmdCancel.Width - 360
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 180
End Sub

Private Sub picXY_Resize()
    vsDiagXY.Height = picXY.Height - 300
    vsDiagXY.Width = picXY.Width
End Sub


Private Sub picZY_Resize()
    vsDiagZY.Height = picZY.Height - 300
    vsDiagZY.Width = picZY.Width
End Sub

Private Sub SetVSColHidden()
'功能设置VS中列的可见性
    With vsDiagXY
        .ColHidden(col诊断类型) = mint病人来源 = 1
        .ColHidden(col中医证候) = True
        .ColHidden(col入院病情) = mint病人来源 = 1
        .ColHidden(col出院情况) = mint病人来源 = 1
        .ColHidden(col是否未治) = mint病人来源 = 1
        .ColHidden(col发病时间) = mint病人来源 <> 1
    End With
    
    With vsDiagZY
        .ColHidden(col诊断类型) = mint病人来源 = 1
        .ColHidden(col入院病情) = mint病人来源 = 1
        .ColHidden(col出院情况) = mint病人来源 = 1
        .ColHidden(col发病时间) = mint病人来源 <> 1
        .ColHidden(colZY疑诊) = mint病人来源 <> 1
    End With
End Sub



Private Sub tbcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mint病人来源 = 2 Then
        If Item.Index = 0 Then
            If Val(gobjComLib.zlDatabase.GetPara("西医诊断输入", M_LNG_SYS, M_LNG_P住院医生站, 0, Array(optInput(0), optInput(1)), InStr(mstrPrivs, "参数设置") > 0)) = 0 Then
                optInput(0).value = True
            Else
                optInput(1).value = True
            End If

        Else
            If Val(gobjComLib.zlDatabase.GetPara("中医诊断输入", M_LNG_SYS, M_LNG_P住院医生站, 0, Array(optInput(0), optInput(1)), InStr(mstrPrivs, "参数设置") > 0)) = 0 Then
                optInput(0).value = True
            Else
                optInput(1).value = True
            End If
        End If
    ElseIf mint病人来源 = 1 Then
        optInput(Val(gobjComLib.zlDatabase.GetPara("门诊诊断输入", M_LNG_SYS, M_LNG_P门诊医生站, 0, Array(optInput(0), optInput(1)), InStr(mstrPrivs, "参数设置") > 0))).value = True
    End If
    Call Form_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mint病人来源 = 2 Then
        Call gobjComLib.zlDatabase.SetPara("西医诊断输入", IIf(optInput(0).value, 0, 1), M_LNG_SYS, M_LNG_P住院医生站, InStr(mstrPrivs, "参数设置") > 0)
        Call gobjComLib.zlDatabase.SetPara("中医诊断输入", IIf(optInput(1).value, 0, 1), M_LNG_SYS, M_LNG_P住院医生站, InStr(mstrPrivs, "参数设置") > 0)
    ElseIf mint病人来源 = 1 Then
        Call gobjComLib.zlDatabase.SetPara("门诊诊断输入", IIf(optInput(1).value, 0, 1), M_LNG_SYS, M_LNG_P门诊医生站, InStr(mstrPrivs, "参数设置") > 0)
    End If
End Sub

Private Sub vsDiagXY_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsDiagXY
        If Col = col出院情况 Then
            '主要处理非回车离开:不用ComboIndex,取消编辑时不对
            .TextMatrix(Row, Col) = NeedName(.TextMatrix(Row, Col))
            If Not XYCellEditable(Row, col是否未治) Then
                .TextMatrix(Row, col是否未治) = ""
            End If
            .Tag = ""
        End If
        If Col = Col诊断描述 Then
            ' .EditText = "" 排除单元格有内容并按回车的状况
            If .EditText = "" And .Cell(flexcpData, Row, Col) <> "" Then
                '在调用vsDiagXY_KeyDown(vbKeyDelete, 0)点是可以删除当前行，点否则恢复原始数据
                .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                Call vsDiagXY_KeyDown(vbKeyDelete, 0)
            End If
        End If
        Call vsDiagXY_AfterRowColChange(-1, -1, .Row, .Col)
        '判断是否做了修改
        If vsDiagXY.Tag = "未修改" And Col <> col关联 Then
            vsDiagXY.Tag = ""
        End If
    End With
    
End Sub

Private Sub vsDiagXY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Long

    With vsDiagXY
        '清除图片
        For i = .FixedRows To .Rows - 1
            If Not .Cell(flexcpPicture, i, col增加) Is Nothing Then
                Set .Cell(flexcpPicture, i, col增加) = Nothing
            End If
            If Not .Cell(flexcpPicture, i, colDel) Is Nothing Then
               Set .Cell(flexcpPicture, i, colDel) = Nothing
            End If
        Next

        If Not XYCellEditable(NewRow, NewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
            Set .CellButtonPicture = Nothing

            If NewCol = Col诊断描述 Then
                .ComboList = "..."
            ElseIf NewCol = col出院情况 Then
                .ComboList = .ColData(NewCol)
            ElseIf NewCol = col入院病情 Then
                If .TextMatrix(NewRow, 0) = "出院诊断" Or .TextMatrix(NewRow, 0) = "其他诊断" Or .TextMatrix(NewRow, 0) = "" Then
                    .ComboList = "有|临床未确定|情况不明|无"
                Else
                    .ComboList = ""
                    .FocusRect = flexFocusLight
                End If
            ElseIf NewCol = col增加 Then
                .ComboList = "..."
                .FocusRect = flexFocusNone
                Set .CellButtonPicture = imgButtonNew.Picture
            ElseIf NewCol = colDel Then
                .ComboList = "..."
                .FocusRect = flexFocusNone
                Set .CellButtonPicture = imgButtonDel.Picture
            Else
                .ComboList = ""
            End If
        End If
        If NewRow >= .FixedRows Then
            '显示图片
            If NewCol <> col增加 And .TextMatrix(NewRow, Col诊断描述) <> "" And .TextMatrix(NewRow, 0) <> "出院诊断" Then
                Set .Cell(flexcpPicture, NewRow, col增加) = imgButtonNew.Picture
            End If
            '显示图片
            If NewCol <> colDel And .RowData(NewRow) & "" = "" Then
                Set .Cell(flexcpPicture, NewRow, colDel) = imgButtonDel.Picture
            End If
        End If
    End With
End Sub

Private Sub vsDiagXY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = col增加 Then Cancel = True
End Sub

Private Sub vsDiagXY_Click()
    With vsDiagXY
        If (.MouseCol = col增加 Or .MouseCol = colDel) And .MouseRow >= .FixedRows Then
            If .MouseCol = col增加 Then
                If .TextMatrix(.MouseRow, Col诊断描述) = "" Or .TextMatrix(.MouseRow, 0) = "出院诊断" Then Exit Sub
            End If

            .Select .MouseRow, .MouseCol
            Call vsDiagXY_CellButtonClick(.MouseRow, .MouseCol)
        End If
    End With
End Sub
Private Sub vsDiagXY_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    With vsDiagXY
        If Col = col出院情况 Then
            '定位到匹配项
            For i = 0 To .ComboCount - 1
                If NeedName(.ComboItem(i)) = .TextMatrix(Row, Col) Then
                    .ComboIndex = i: Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub vsDiagXY_DblClick()
    Call vsDiagXY_KeyPress(32)
    '设置为已修改
    If vsDiagXY.Col = col是否未治 Or vsDiagXY.Col = col是否疑诊 Then
        If vsDiagXY.Tag = "未修改" Then vsDiagXY.Tag = ""
    End If
End Sub

Private Sub vsDiagXY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, j As Long

    With vsDiagXY
        If KeyCode = vbKeyF4 Then
            If .Col = Col诊断描述 Then
                Call gobjComLib.zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, Col诊断描述) <> "" Then
                If .RowData(.Row) & "" <> "" Then Exit Sub
                If GetAdviceIDByDiag(Val(.Cell(flexcpData, .Row, col是否疑诊) & "")) <> "" Then Exit Sub
                
                If mlngPathState = 1 Or mlngPathState = 2 Then
                    If .TextMatrix(.Row, col诊断类型) = "入院诊断" And mlngDiagnosisType = 2 Or .TextMatrix(.Row, col诊断类型) = "门诊诊断" And mlngDiagnosisType = 1 Then
                        If .TextMatrix(.Row, col诊断类型) <> .TextMatrix(.Row - 1, col诊断类型) Then
                            '首要诊断不允许改
                            Exit Sub
                        End If
                    End If
                End If
                '合并路径
                If Not CheckMergePath(mlng病人ID, mlng就诊ID, Val(.TextMatrix(.Row, col类型)), Val(.TextMatrix(.Row, col疾病ID))) Then Exit Sub
                '两条路径以上
                If mstrPathDiag <> "" And mlngPathState > 0 Then
                    If InStr("," & mstrPathDiag & ",", "," & .TextMatrix(.Row, col类型) & "|" & Val(.TextMatrix(.Row, col疾病ID)) & "|" & Val(.TextMatrix(.Row, col诊断ID)) & ",") > 0 Then
                        '导入诊断不允许该
                        Exit Sub
                    End If
                End If
                If mlngPathState = 2 And mblnIsPathOutTime Then
                    If .TextMatrix(.Row, col诊断类型) = "出院诊断" And mlngDiagnosisType <= 2 Then
                        '正常完成的出院诊断不允许改
                        Exit Sub
                    End If
                End If
                If MsgBox("确实要清除该行诊断信息吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    i = Val(.TextMatrix(.Row, col类型))
                    .Cell(flexcpText, .Row, .FixedCols, .Row, .Cols - 1) = ""
                    .Cell(flexcpData, .Row, .FixedCols, .Row, .Cols - 1) = Empty
                    .TextMatrix(.Row, col类型) = i

                    '下面的同类诊断数据上移
                    If .TextMatrix(.Row, col诊断类型) = "" Then
                        .RemoveItem .Row
                    Else
                        If .Row + 1 <= .Rows - 1 Then
                            If .TextMatrix(.Row + 1, col诊断类型) = "" Then
                                '下一行为无标题的增加行时，数据才上移，否则当前行为有标题时只清空行
                                For i = .Row + 1 To .Rows - 1
                                    If Val(.TextMatrix(i, col类型)) = Val(.TextMatrix(.Row, col类型)) Then
                                        For j = .FixedCols To .Cols - 1
                                            .TextMatrix(i - 1, j) = .TextMatrix(i, j)
                                            .Cell(flexcpData, i - 1, j) = .Cell(flexcpData, i, j)
                                        Next
                                        .Cell(flexcpText, i, .FixedCols, i, .Cols - 1) = ""
                                        .Cell(flexcpData, i, .FixedCols, i, .Cols - 1) = Empty
                                        .TextMatrix(i, col类型) = Val(.TextMatrix(.Row, col类型))
                                        .RowData(i - 1) = .RowData(i)
                                        .RowData(i) = Empty
                                        If i = .Rows - 1 Then
                                            If .TextMatrix(i, col诊断类型) = "" Then .RemoveItem i
                                            Exit For
                                        ElseIf Val(.TextMatrix(i + 1, col类型)) <> Val(.TextMatrix(i, col类型)) Then
                                            If .TextMatrix(i, col诊断类型) = "" Then .RemoveItem i
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                    .Tag = ""
                End If
            ElseIf .TextMatrix(.Row, col诊断类型) = "" Then
                .RemoveItem .Row
            End If
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsDiagXY_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsDiagXY_KeyPress(KeyAscii As Integer)
    With vsDiagXY
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call XYEnterNextCell
        ElseIf KeyAscii = 32 And (.Col = col是否未治 Or .Col = col是否疑诊) Then
            If XYCellEditable(.Row, .Col) Then
                KeyAscii = 0
                If .Col = col是否疑诊 Then
                    .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = "", "？", "")
                ElseIf .Col = col是否未治 Then
                    .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = "", "√", "")
                End If
                .Tag = ""
            End If
        Else
            If .Col = Col诊断描述 Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsDiagXY_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" '使按钮状态进入输入状态
                End If
            End If
        End If
    End With
End Sub

Private Sub vsDiagXY_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsDiagXY.EditSelStart = 0
    vsDiagXY.EditSelLength = gobjComLib.zlCommFun.ActualLen(vsDiagXY.EditText)
End Sub

Private Sub vsDiagXY_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not XYCellEditable(Row, Col) Then
        Cancel = True
    ElseIf Col = col是否未治 Or Col = col是否疑诊 Then
        Cancel = True '不直接编辑
    End If
End Sub

Private Sub vsDiagXY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim str性别 As String, lngRow As Long

    With vsDiagXY
        If Col = Col诊断描述 Then
            If optInput(0).value Then
                '按诊断输入:西医部份，一个诊断可能属于多个分类
                Set rsTmp = gobjComLib.zlDatabase.ShowILLSelect(Me, "1", mlng科室ID, , True, False)
            Else
                '7-损伤中毒：Y-损伤中毒的外部原因；6-病理诊断：M-肿瘤形态学编码；其它诊断：D-ICD-10疾病编码
                Set rsTmp = gobjComLib.zlDatabase.ShowILLSelect(Me, Decode(Val(.TextMatrix(Row, col类型)), 7, "'Y'", 6, "'M,D'", "'D'"), mlng科室ID, mstr性别, True)
            End If
            If Not rsTmp Is Nothing Then
                .Tag = ""
                Call XYSetDiagInput(Row, rsTmp)
                Call XYEnterNextCell
            End If
        ElseIf Col = col增加 Then
            lngRow = Row + 1: .AddItem "", lngRow
            .TextMatrix(lngRow, col类型) = .TextMatrix(Row, col类型)
            .Cell(flexcpBackColor, lngRow, col诊断编码) = ColorUnEditCell      '灰蓝色
            
            .Row = lngRow: .Col = Col诊断描述
            .ShowCell .Row, .Col
        ElseIf Col = colDel Then
            Call vsDiagXY_KeyDown(vbKeyDelete, 0)
        End If
    End With
End Sub

Private Sub vsDiagXY_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True

        With vsDiagXY
            If Col = col出院情况 Then
                KeyAscii = 0
                If .ComboIndex <> -1 Then
                    '此时.TextMatrix尚未更新,所以取ComboItem
                    .TextMatrix(Row, Col) = NeedName(.ComboItem(.ComboIndex))
                    If Not XYCellEditable(Row, col是否未治) Then
                        .TextMatrix(Row, col是否未治) = ""
                    End If
                    Call XYEnterNextCell
                    .Tag = ""
                End If
            End If
        End With
    Else
        mblnReturn = False
    End If
End Sub

Private Sub XYSetDiagInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'功能：处理西医诊断项目的输入
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    Dim bln分化程度 As Boolean

    With vsDiagXY
        If Not rsInput Is Nothing Then
            For i = 1 To rsInput.RecordCount
                If i > 1 Then
                    '损伤中毒选择多条时的处理
                    If lngRow = .Rows - 1 Then
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, col类型) = .TextMatrix(lngRow, col类型)
                    End If
                    '确定当前显示行
                    If Val(.TextMatrix(lngRow + 1, col类型)) = Val(.TextMatrix(lngRow, col类型)) Then
                        For j = lngRow + 1 To .Rows - 1
                            If Val(.TextMatrix(j, col类型)) = Val(.TextMatrix(lngRow, col类型)) Then
                                lngRow = j
                                If .TextMatrix(j, Col诊断描述) = "" Then Exit For
                            Else
                                Exit For
                            End If
                        Next
                        If .TextMatrix(lngRow, Col诊断描述) <> "" Then
                            lngRow = lngRow + 1: .AddItem "", lngRow
                            .TextMatrix(lngRow, col类型) = .TextMatrix(lngRow - 1, col类型)
                        End If
                    Else
                        lngRow = lngRow + 1: .AddItem "", lngRow
                        .TextMatrix(lngRow, col类型) = .TextMatrix(lngRow - 1, col类型)
                    End If
                End If
                .TextMatrix(lngRow, col关联) = 1
                .TextMatrix(lngRow, col诊断编码) = "" & rsInput!编码
                .TextMatrix(lngRow, Col诊断描述) = "" & rsInput!名称
                
                .Cell(flexcpData, lngRow, Col诊断描述) = .TextMatrix(lngRow, Col诊断描述)

                '根据诊断确定疾病,或根据疾病确定诊断
                If optInput(0).value Then
                    .TextMatrix(lngRow, col诊断ID) = rsInput!项目ID
                    .TextMatrix(lngRow, col疾病ID) = ""
                    strSQL = "Select 疾病ID as ID From 疾病诊断对照 Where 诊断ID=[1]"
                Else
                    .TextMatrix(lngRow, col疾病ID) = rsInput!项目ID
                    .TextMatrix(lngRow, col诊断ID) = ""
                    strSQL = "Select 诊断ID as ID From 疾病诊断对照 Where 疾病ID=[1]"
                End If
                On Error GoTo errH
                Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!项目ID))
                If Not rsTmp.EOF Then
                    If optInput(0).value Then
                        .TextMatrix(lngRow, col疾病ID) = NVL(rsTmp!id)
                    Else
                        .TextMatrix(lngRow, col诊断ID) = NVL(rsTmp!id)
                    End If
                End If

                rsInput.MoveNext
            Next
        Else
            .TextMatrix(lngRow, col诊断编码) = ""
            .TextMatrix(lngRow, Col诊断描述) = .EditText
            .Cell(flexcpData, lngRow, Col诊断描述) = .TextMatrix(lngRow, Col诊断描述)
            .TextMatrix(lngRow, col诊断ID) = ""
            .TextMatrix(lngRow, col疾病ID) = ""
        End If

        .Cell(flexcpForeColor, 1, col是否疑诊, .Rows - 1, col是否疑诊) = vbRed
    End With
    Exit Sub
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub vsDiagXY_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim str性别 As String, int诊断输入 As Integer
    Dim strInput As String, vPoint As POINTAPI

    With vsDiagXY
        If Col = Col诊断描述 Then
            '.Cell(flexcpData, Row, Col) <> ""排除空行回车
            If .EditText = "" And .Cell(flexcpData, Row, Col) <> "" Then
                .EditText = ""
                .Tag = ""
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call XYEnterNextCell
            ElseIf .TextMatrix(Row, col诊断编码) <> "" And .Cell(flexcpData, Row, Col) <> "" And .EditText Like "*" & .Cell(flexcpData, Row, Col) & "*" Then
                '判断加了前缀后的名称是否存在其他的诊断编码
                strInput = UCase(.EditText)
                strSQL = GetSQL(0, strInput, str性别)
                On Error GoTo errH
                Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, strInput, _
                        Decode(Val(.TextMatrix(Row, col类型)), 7, "'Y'", 6, "'M,D'", "'D'"), str性别, mint简码 + 1)
                If rsTmp.RecordCount <> 1 Then
                    '允许在标准的名称前后输入附加信息
                    .TextMatrix(Row, Col诊断描述) = .EditText
                    .Tag = ""
                Else
                    Call XYSetDiagInput(Row, rsTmp)
                    .EditText = .Text
                    .Tag = ""
                End If
                '不处理.Cell(flexcpData, Row, Col)，以便修改内容时再次使用like判断
            Else
                If Val(.TextMatrix(Row, col类型)) = 1 Then
                    int诊断输入 = Val(Mid(mstr诊断输入, 1, 1))
                Else
                    int诊断输入 = Val(Mid(mstr诊断输入, 2, 1))
                End If
                If int诊断输入 = 0 Then int诊断输入 = 1

                strInput = UCase(.EditText)
                strSQL = GetSQL(0, strInput, str性别)
                If int诊断输入 = 1 And gobjComLib.zlCommFun.IsCharChinese(strInput) Then
                    '损伤中毒码：Y-损伤中毒的外部原因；病理诊断允许：M-肿瘤形态学编码；其它诊断：D-ICD-10疾病编码
                    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", mstrLike & strInput & "%", _
                        Decode(Val(.TextMatrix(Row, col类型)), 7, "'Y'", 6, "'M,D'", "'D'"), str性别, mint简码 + 1)
                    If rsTmp.EOF Then
                        Set rsTmp = Nothing
                    ElseIf rsTmp.RecordCount > 1 Then
                        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, strInput, _
                        Decode(Val(.TextMatrix(Row, col类型)), 7, "'Y'", 6, "'M,D'", "'D'"), str性别, mint简码 + 1)
                        If rsTmp.RecordCount <> 1 Then Set rsTmp = Nothing '自由录入时有多个匹配不进行选择
                    End If
                    .Tag = ""
                    Call XYSetDiagInput(Row, rsTmp)
                    .EditText = .Text
                    If mblnReturn And rsTmp Is Nothing Then Call XYEnterNextCell '不是自由录入时，暂不跳到下一行，因为可能还要改描述内容
                Else
                    vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = gobjComLib.zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIf(optInput(0).value, "疾病诊断", "疾病编码"), _
                        False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                        strInput & "%", mstrLike & strInput & "%", Decode(Val(.TextMatrix(Row, col类型)), 7, "'Y'", 6, "'M,D'", "'D'"), str性别, mint简码 + 1)
                    If blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                        Cancel = True
                    Else
                        '检查诊断输入方式
                        If rsTmp Is Nothing And ((int诊断输入 = 2 Or int诊断输入 = 3 And mint险类 <> 0)) Then
                            MsgBox "没有找到与输入匹配的内容。", vbInformation, gstrSysName
                            Cancel = True
                        Else
                            .Tag = ""
                            Call XYSetDiagInput(Row, rsTmp): .EditText = .Text
                            'If mblnReturn Then Call XYEnterNextCell    '暂不跳到下一行，因为可能还要改描述内容
                        End If
                    End If
                End If
            End If
            mblnReturn = False
        ElseIf Col = col发病时间 Then
            If .EditText <> "" Then
                strInput = GetFullDate(.EditText)
                If IsDate(strInput) Then
                    .EditText = Format(strInput, "yyyy-MM-dd HH:mm")
                    .Tag = ""
                Else
                    MsgBox "请输入正确的发病时间，例如：""2012-12-21 00:00""。"
                    Cancel = True
                End If
            End If
        End If
    End With
    Exit Sub
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub vsDiagZY_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsDiagZY
        If Col = col出院情况 Then
            .TextMatrix(Row, Col) = NeedName(.TextMatrix(Row, Col))
            .Tag = ""
        End If
        If Col = Col诊断描述 Then
            ' .EditText = "" 排除单元格有内容并按回车的状况
            If .EditText = "" And .Cell(flexcpData, Row, Col) <> "" Then
                '在调用vsDiagZY_KeyDown(vbKeyDelete, 0)点是可以删除当前行，点否则恢复原始数据
                .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                Call vsDiagZY_KeyDown(vbKeyDelete, 0)
            End If
        End If
        Call vsDiagZY_AfterRowColChange(-1, -1, .Row, .Col)
        If Col <> col关联 Then .Tag = ""
    End With
End Sub

Private Sub vsDiagZY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Long

    With vsDiagZY

        '清除图片
        For i = .FixedRows To .Rows - 1
            If Not .Cell(flexcpPicture, i, colZY增加) Is Nothing Then
                Set .Cell(flexcpPicture, i, colZY增加) = Nothing
            End If
            If Not .Cell(flexcpPicture, i, colZYDel) Is Nothing Then
               Set .Cell(flexcpPicture, i, colZYDel) = Nothing
            End If
        Next

        If Not ZYCellEditable(NewRow, NewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
            Set .CellButtonPicture = Nothing

            If NewCol = Col诊断描述 Then
                .ComboList = "..."
            ElseIf NewCol = col中医证候 Then
                If .TextMatrix(NewRow, Col诊断描述) = "" Then
                    .ComboList = ""
                    .FocusRect = flexFocusLight
                Else
                    .ComboList = "..."
                End If
            ElseIf NewCol = col出院情况 Then
                .ComboList = .ColData(NewCol)
            ElseIf NewCol = col入院病情 Then
                If .TextMatrix(NewRow, colzy类型) = "13" Then
                    .ComboList = "有|临床未确定|情况不明|无"
                Else
                    .ComboList = ""
                    .FocusRect = flexFocusLight
                End If
            ElseIf NewCol = colZY增加 Then
                .ComboList = "..."
                .FocusRect = flexFocusNone
                Set .CellButtonPicture = imgButtonNew.Picture
            ElseIf NewCol = colZYDel Then
                .ComboList = "..."
                .FocusRect = flexFocusNone
                Set .CellButtonPicture = imgButtonDel.Picture
            Else
                .ComboList = ""
            End If
        End If
        If NewRow >= .FixedRows Then
            '显示图片
            If NewCol <> colZY增加 And .TextMatrix(NewRow, Col诊断描述) <> "" And .TextMatrix(NewRow, 0) <> "主要诊断" Then
                Set .Cell(flexcpPicture, NewRow, colZY增加) = imgButtonNew.Picture
            End If
            '显示图片
            If NewCol <> colZYDel And .RowData(NewRow) & "" = "" Then
                Set .Cell(flexcpPicture, NewRow, colZYDel) = imgButtonDel.Picture
            End If
        End If
    End With
End Sub

Private Sub vsDiagZY_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = colZY疑诊 Then Cancel = True
End Sub

Private Sub vsDiagZY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = colZY增加 Then Cancel = True
End Sub

Private Sub vsDiagZY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim str性别 As String, lngRow As Long
    Dim blnCancle As Boolean

    With vsDiagZY
        If Col = Col诊断描述 Then
            If optInput(1).value Then
                '按诊断输入:中医部份，一个诊断可能属于多个分类
                Set rsTmp = gobjComLib.zlDatabase.ShowILLSelect(Me, "2", mlng科室ID, , True, False)
            Else
                'B-中医疾病编码
                Set rsTmp = gobjComLib.zlDatabase.ShowILLSelect(Me, "B", mlng科室ID, mstr性别, True)
            End If
            If Not rsTmp Is Nothing Then
                .Tag = ""
                Call ZYSetDiagInput(Row, rsTmp)
                Call ZYEnterNextCell
            End If
        ElseIf Col = col中医证候 Then
            If optInput(1).value Then
                '按诊断输入:先查是否有对应
                If Not Set中医证候(Row, Val(.TextMatrix(Row, colzy诊断ID))) Then
                    Set rsTmp = gobjComLib.zlDatabase.ShowILLSelect(Me, "Z", mlng科室ID, mstr性别, True)
                Else
                    Exit Sub
                End If
            Else
                'Z-中医疾病编码
                Set rsTmp = gobjComLib.zlDatabase.ShowILLSelect(Me, "Z", mlng科室ID, mstr性别, True)
            End If
            If Not rsTmp Is Nothing Then
                .Tag = ""
                Call Set中医证候(Row, 0, rsTmp)
                Call ZYEnterNextCell
            End If
        ElseIf Col = colZY增加 Then
            lngRow = Row + 1: .AddItem "", lngRow
            .TextMatrix(lngRow, colzy类型) = .TextMatrix(Row, colzy类型)
            .Cell(flexcpBackColor, lngRow, col诊断编码) = ColorUnEditCell      '灰蓝色
            .Row = lngRow: .Col = Col诊断描述
            .ShowCell .Row, .Col
        ElseIf Col = colZYDel Then
            Call vsDiagZY_KeyDown(vbKeyDelete, 0)
        End If
    End With
End Sub

Private Sub vsDiagZY_Click()
    With vsDiagZY
        If (.MouseCol = colZY增加 Or .MouseCol = colZYDel) And .MouseRow >= .FixedRows Then
            If .MouseCol = colZY增加 Then
                If .TextMatrix(.MouseRow, Col诊断描述) = "" Or .TextMatrix(.MouseRow, 0) = "主要诊断" Then Exit Sub
            End If

            .Select .MouseRow, .MouseCol
            Call vsDiagZY_CellButtonClick(.MouseRow, .MouseCol)
        End If
    End With
End Sub

Private Sub vsDiagZY_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    With vsDiagZY
        If Col = col出院情况 Then
            '定位到匹配项
            For i = 0 To .ComboCount - 1
                If NeedName(.ComboItem(i)) = .TextMatrix(Row, Col) Then
                    .ComboIndex = i: Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub vsDiagZY_DblClick()
    Call vsDiagZY_KeyPress(32)
End Sub

Private Sub vsDiagZY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, j As Long

    With vsDiagZY
        If KeyCode = vbKeyF4 Then
            If .Col = Col诊断描述 Then
                Call gobjComLib.zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, Col诊断描述) <> "" Then
                If .RowData(.Row) & "" <> "" Then Exit Sub
                If GetAdviceIDByDiag(Val(.Cell(flexcpData, .Row, col是否疑诊) & "")) <> "" Then Exit Sub
                If mlngPathState = 1 Or mlngPathState = 2 Then
                    If .TextMatrix(.Row, col诊断类型) = "入院诊断" And mlngDiagnosisType = 12 Or .TextMatrix(.Row, col诊断类型) = "门诊诊断" And mlngDiagnosisType = 11 Then
                        If .TextMatrix(.Row, col诊断类型) <> .TextMatrix(.Row - 1, col诊断类型) Then
                            '首要诊断不允许改
                            Exit Sub
                        End If
                    End If
                End If
                '合并路径
                If Not CheckMergePath(mlng病人ID, mlng就诊ID, Val(.TextMatrix(.Row, colzy类型)), Val(.TextMatrix(.Row, colzy疾病ID))) Then Exit Sub
                '两条路径以上
                If mstrPathDiag <> "" And mlngPathState > 0 Then
                    If InStr("," & mstrPathDiag & ",", "," & .TextMatrix(.Row, colzy类型) & "|" & Val(.TextMatrix(.Row, col疾病ID)) & "|" & Val(.TextMatrix(.Row, col诊断ID)) & ",") > 0 Then
                        '导入诊断不允许该
                        Exit Sub
                    End If
                End If
                If mlngPathState = 2 And mblnIsPathOutTime Then
                    If .TextMatrix(.Row, col诊断类型) = "主要诊断" And mlngDiagnosisType > 10 Then
                        '正常完成的出院诊断不允许改
                        Exit Sub
                    End If
                End If
                If MsgBox("确实要清除该行诊断信息吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    i = Val(.TextMatrix(.Row, colzy类型))
                    .Cell(flexcpText, .Row, .FixedRows, .Row, .Cols - 1) = ""
                    .Cell(flexcpData, .Row, .FixedRows, .Row, .Cols - 1) = Empty
                    .TextMatrix(.Row, colzy类型) = i

                    '下面的同类诊断数据上移
                    If .TextMatrix(.Row, col诊断类型) = "" Then
                        .RemoveItem .Row
                    Else
                        If .Row + 1 <= .Rows - 1 Then
                            If .TextMatrix(.Row + 1, col诊断类型) = "" Then
                                '下一行为无标题的增加行时，数据才上移，否则当前行为有标题时只清空行
                                For i = .Row + 1 To .Rows - 1
                                    If Val(.TextMatrix(i, colzy类型)) = Val(.TextMatrix(.Row, colzy类型)) Then
                                        For j = .FixedCols To .Cols - 1
                                            .TextMatrix(i - 1, j) = .TextMatrix(i, j)
                                            .Cell(flexcpData, i - 1, j) = .Cell(flexcpData, i, j)
                                        Next
                                        .Cell(flexcpText, i, .FixedCols, i, .Cols - 1) = ""
                                        .Cell(flexcpData, i, .FixedCols, i, .Cols - 1) = Empty
                                        .TextMatrix(i, colzy类型) = Val(.TextMatrix(.Row, colzy类型))
                                        .RowData(i - 1) = .RowData(i)
                                        .RowData(i) = Empty
                                        
                                        If i = .Rows - 1 Then
                                            If .TextMatrix(i, col诊断类型) = "" Then .RemoveItem i
                                            Exit For
                                        ElseIf Val(.TextMatrix(i + 1, colzy类型)) <> Val(.TextMatrix(i, colzy类型)) Then
                                            If .TextMatrix(i, col诊断类型) = "" Then .RemoveItem i
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                    .Tag = ""
                End If
            ElseIf .TextMatrix(.Row, col诊断类型) = "" Then
                .RemoveItem .Row
            End If
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsDiagZY_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsDiagZY_KeyPress(KeyAscii As Integer)
    With vsDiagZY
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call ZYEnterNextCell
        ElseIf KeyAscii = 32 And (.Col = colZY疑诊) Then
            If ZYCellEditable(.Row, .Col) Then
                KeyAscii = 0
                If .Col = colZY疑诊 Then
                    .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = "", "？", "")
                End If
            End If
        Else
            If .Col = Col诊断描述 Or .Col = col中医证候 Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsDiagZY_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" '使按钮状态进入输入状态
                End If
            End If
        End If
    End With
End Sub

Private Sub vsDiagZY_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True

        With vsDiagZY
            If Col = col出院情况 Then
                KeyAscii = 0

                '此时.TextMatrix尚未更新,所以取ComboItem
                .TextMatrix(Row, Col) = NeedName(.ComboItem(.ComboIndex))
                .Tag = ""
                Call ZYEnterNextCell
            End If
        End With
    Else
        mblnReturn = False
    End If
End Sub

Private Sub vsDiagZY_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsDiagZY.EditSelStart = 0
    vsDiagZY.EditSelLength = gobjComLib.zlCommFun.ActualLen(vsDiagZY.EditText)
End Sub

Private Sub vsDiagZY_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not ZYCellEditable(Row, Col) Then
        Cancel = True
    ElseIf Col = colZY疑诊 Then
        Cancel = True '不直接编辑
    End If
End Sub

Private Function GetSQL(ByVal intType As Integer, ByVal strInput As String, ByRef str性别 As String, Optional ByVal strOtherInfo As String) As String
'功能：获得查询西医诊断的SQL
'参数：intType:获取的SQL类型,0-西医诊断，1-中医诊断，2-手术操作
'    strInput-查询条件，str性别--病人的性别
'   strOtherInfo:中医诊断-疾病编码种类
'返回：strsql--查询诊断的SQL
    Dim strSQL As String

    If mstr性别 Like "*男*" Then
        str性别 = "男"
    ElseIf mstr性别 Like "*女*" Then
        str性别 = "女"
    End If

    Select Case intType
        Case 0 '西医诊断
            If optInput(0).value Then
            '按诊断输入:西医部份，一个诊断可能属于多个分类
                If gobjComLib.zlCommFun.IsCharChinese(strInput) Then
                    strSQL = "B.名称 Like [2]" '输入汉字时只匹配名称
                Else
                    strSQL = "A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2]"
                End If
                strSQL = _
                    " Select Distinct A.ID,A.ID as 项目ID,A.编码,A.名称,A.说明,A.编者" & _
                    " From 疾病诊断目录 A,疾病诊断别名 B" & _
                    " Where A.ID=B.诊断ID And A.类别=1" & _
                    " And B.码类=[5] And (" & strSQL & ")" & _
                    " Order by A.编码"
            Else
                If gobjComLib.zlCommFun.IsCharChinese(strInput) Then
                    strSQL = "名称 Like [2]" '输入汉字时只匹配名称
                Else
                    strSQL = "编码 Like [1] Or 名称 Like [2] Or " & IIf(mint简码 = 0, "简码", "五笔码") & " Like [2]"
                End If
                strSQL = _
                    " Select ID,ID as 项目ID,编码,附码,名称," & IIf(mint简码 = 0, "简码", "五笔码 as 简码") & ",说明" & _
                    " From 疾病编码目录 Where Instr([3],类别)>0 And (" & strSQL & ")" & _
                    IIf(str性别 <> "", " And (性别限制=[4] Or 性别限制 is NULL)", "") & _
                    " And (撤档时间 is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Order by 编码"
            End If

        Case 1 '中医诊断
            If optInput(0).value And strOtherInfo <> "Z" Then
                '按诊断输入:中医部份，一个诊断可能属于多个分类
                If gobjComLib.zlCommFun.IsCharChinese(strInput) Then
                    strSQL = "B.名称 Like [2]" '输入汉字时只匹配名称
                Else
                    strSQL = "A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2]"
                End If
                strSQL = _
                    " Select Distinct A.ID,A.ID as 项目ID,A.编码,A.名称,A.说明,A.编者" & _
                    " From 疾病诊断目录 A,疾病诊断别名 B" & _
                    " Where A.ID=B.诊断ID And A.类别=2" & _
                    " And B.码类=[4] And (" & strSQL & ")" & _
                    " Order by A.编码"
            Else
                'B-中医疾病编码
                If gobjComLib.zlCommFun.IsCharChinese(strInput) Then
                    strSQL = "名称 Like [2]" '输入汉字时只匹配名称
                Else
                    strSQL = "编码 Like [1] Or 名称 Like [2] Or " & IIf(mint简码 = 0, "简码", "五笔码") & " Like [2]"
                End If
                strSQL = _
                    " Select ID,ID as 项目ID,编码,附码,名称," & IIf(mint简码 = 0, "简码", "五笔码 as 简码") & ",说明" & _
                    " From 疾病编码目录" & _
                    " Where 类别='" & IIf(strOtherInfo = "", "B", strOtherInfo) & "' And (" & strSQL & ")" & _
                    IIf(str性别 <> "", " And (性别限制=[3] Or 性别限制 is NULL)", "") & _
                    " And (撤档时间 is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Order by 编码"
            End If
    End Select
    GetSQL = strSQL
End Function

Private Sub vsDiagZY_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    Dim str性别 As String, int诊断输入 As Integer

    With vsDiagZY
        If Col = Col诊断描述 Or Col = col中医证候 Then
            '.Cell(flexcpData, Row, Col) <> ""排除空行回车
            If .EditText = "" And .Cell(flexcpData, Row, Col) <> "" Then
                .EditText = ""
                '中医症候则清除备份数据
                If Col = col中医证候 Then
                    .Cell(flexcpData, Row, Col) = ""
                End If
                .Tag = ""
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call ZYEnterNextCell
            ElseIf Col = Col诊断描述 And .TextMatrix(Row, col诊断编码) <> "" And .Cell(flexcpData, Row, Col) <> "" And .EditText Like "*" & .Cell(flexcpData, Row, Col) & "*" Then
                strInput = UCase(.EditText)
                strSQL = GetSQL(1, strInput, str性别)
                On Error GoTo errH
                Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, strInput, str性别, mint简码 + 1)
                If rsTmp.RecordCount = 1 Then
                    Call ZYSetDiagInput(Row, rsTmp):
                    .EditText = .Text
                Else
                    '允许在标准的名称前后输入附加信息
                    .TextMatrix(Row, Col诊断描述) = .EditText
                End If
                .Tag = ""
                '不处理.Cell(flexcpData, Row, Col)，以便修改内容时再次使用like判断
            Else
                If Val(.TextMatrix(Row, colzy类型)) = 11 Then
                    int诊断输入 = Val(Mid(mstr诊断输入, 1, 1))
                Else
                    int诊断输入 = Val(Mid(mstr诊断输入, 2, 1))
                End If
                If int诊断输入 = 0 Then int诊断输入 = 1

                strInput = UCase(.EditText)
                strSQL = GetSQL(1, strInput, str性别, IIf(Col = Col诊断描述, "B", "Z"))
                If Col = Col诊断描述 Then
                    If int诊断输入 = 1 And gobjComLib.zlCommFun.IsCharChinese(strInput) Then
                        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", mstrLike & strInput & "%", str性别, mint简码 + 1)
                        If rsTmp.EOF Then
                            Set rsTmp = Nothing
                        ElseIf rsTmp.RecordCount > 1 Then
                            Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, strInput, str性别, mint简码 + 1)
                            If rsTmp.RecordCount <> 1 Then Set rsTmp = Nothing '自由录入时有多个匹配不进行选择
                        End If
                        .Tag = ""
                        Call ZYSetDiagInput(Row, rsTmp): .EditText = .Text
                        If mblnReturn And rsTmp Is Nothing Then Call ZYEnterNextCell '不是自由录入时，暂不跳到下一行，因为可能还要改描述内容
                    Else
                        vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                        Set rsTmp = gobjComLib.zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIf(optInput(0).value, "疾病诊断", "疾病编码"), False, "", "", False, False, True, _
                            vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%", str性别, mint简码 + 1)
                        If blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                            Cancel = True
                        Else
                            '检查诊断输入方式
                            If rsTmp Is Nothing And ((int诊断输入 = 2 Or int诊断输入 = 3 And mint险类 <> 0)) Then
                                MsgBox "没有找到与输入匹配的内容。", vbInformation, gstrSysName
                                Cancel = True
                            Else
                                .Tag = ""
                                Call ZYSetDiagInput(Row, rsTmp): .EditText = .Text
                                'If mblnReturn Then Call ZYEnterNextCell '暂不跳到下一行，因为可能还要改描述内容
                            End If
                        End If
                    End If
                ElseIf Col = col中医证候 Then
                    If optInput(0).value Then
                        '按诊断输入:先查是否有对应
                        If Set中医证候(Row, Val(.TextMatrix(Row, colzy诊断ID))) Then
                            mblnReturn = False
                            Exit Sub
                        End If
                    End If
                    vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = gobjComLib.zlDatabase.ShowSQLSelect(Me, strSQL, 0, "中医证候", False, "", "", False, False, True, _
                        vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%", str性别, mint简码 + 1)
                    If blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                        Cancel = True
                    Else
                        '检查诊断输入方式
                        If rsTmp Is Nothing Then
                            MsgBox "没有找到与输入匹配的内容。", vbInformation, gstrSysName
                            Cancel = True
                        Else
                            .Tag = ""
                            Call Set中医证候(Row, 0, rsTmp)
                        End If
                    End If
                End If
            End If
            mblnReturn = False
        ElseIf Col = col发病时间 Then
            If .EditText <> "" Then
                strInput = GetFullDate(.EditText)
                If IsDate(strInput) Then
                    .EditText = Format(strInput, "yyyy-MM-dd HH:mm")
                    .Tag = ""
                Else
                    MsgBox "请输入正确的发病时间，例如：""2012-12-21 00:00""。"
                    Cancel = True
                End If
            End If
        End If
    End With
    Exit Sub
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub LoadData()
    Dim bln首页诊断 As Boolean
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim i As Long, lngRow As Long, j As Long
    Dim str治疗结果 As String
    Dim str诊断Id As String
    
    On Error GoTo errH
    '西医诊断
    '--------------------------------------------------------------
    '判断首页是否填过诊断
    strSQL = "Select 1 From 病人诊断记录 Where 病人ID=[1] And 主页ID=[2] And 记录来源=3  And RowNum<2"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng就诊ID)
    bln首页诊断 = rsTmp.RecordCount > 0
    If Not bln首页诊断 And mint病人来源 = 2 Then
        strTmp = " And a.记录来源 IN(1,2,3,4) "
    Else
        strTmp = " and a.记录来源=3 "
    End If

    '缺省表格初始化
    With vsDiagXY
        .ColData(col出院情况) = Get治疗结果
        '1-西医门诊诊断;2-西医入院诊断;3-出院诊断(其他诊断);5-院内感染;6-病理诊断;7-损伤中毒码;10-并发症
        .TextMatrix(1, col类型) = 1
        If mint病人来源 = 2 Then
            .TextMatrix(2, col类型) = 2
            .TextMatrix(3, col类型) = 3
            .TextMatrix(4, col类型) = 3
            .TextMatrix(5, col类型) = 5
            .TextMatrix(6, col类型) = 10
            .TextMatrix(7, col类型) = 6
            .TextMatrix(8, col类型) = 7
        Else
            .Rows = .FixedRows + 1
        End If
    End With

    '读取各种来源的诊断
    strSQL = "Select a.备注,a.ID,a.病人ID,a.主页ID,a.医嘱ID,a.记录来源,a.诊断次序,a.编码序号,a.病历ID,a.诊断类型,a.疾病ID,a.入院病情," & _
        " a.诊断ID,a.证候ID,a.诊断描述,a.出院情况,a.是否未治,a.是否疑诊,a.记录日期,a.记录人,a.取消时间,a.取消人,a.病例ID, b.编码 As 疾病编码, c.编码 As 诊断编码,A.发病时间 " & _
        " From 病人诊断记录 A, 疾病编码目录 B, 疾病诊断目录 C" & _
        " Where a.疾病id = b.Id(+) And a.诊断id = c.Id(+)  And a.诊断类型 IN(1,2,3,5,6,7,10,21)" & _
        strTmp & _
        " And a.取消时间 is Null And a.病人ID=[1] And a.主页ID=[2]" & _
        " Order by a.诊断类型,a.记录来源 Desc,a.诊断次序,a.ID"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng就诊ID)
    If Not rsTmp.EOF Then
        With vsDiagXY
            If mint病人来源 = 2 Then
                strSQL = "1,2,3,5,6,7,10"
            Else
                 strSQL = "1"
            End If
            For i = 0 To UBound(Split(strSQL, ","))
                rsTmp.Filter = "记录来源=3 And 诊断类型=" & Split(strSQL, ",")(i)
                If mint病人来源 = 2 Then
                    If rsTmp.EOF Then
                        rsTmp.Filter = "记录来源=2 And 诊断类型=" & Split(strSQL, ",")(i)
                    End If
                    If rsTmp.EOF Then
                        rsTmp.Filter = "记录来源=1 And 诊断类型=" & Split(strSQL, ",")(i)
                    End If
                    If rsTmp.EOF Then
                        rsTmp.Filter = "记录来源=4 And 诊断类型=" & Split(strSQL, ",")(i)
                    End If
                End If
                Do While Not rsTmp.EOF
                    '确定当前显示行
                    lngRow = .FindRow(CStr(Split(strSQL, ",")(i)), , col类型)
                    For j = lngRow To .Rows - 1
                        If Val(.TextMatrix(j, col类型)) = Val(Split(strSQL, ",")(i)) Then
                            lngRow = j
                            If .TextMatrix(j, Col诊断描述) = "" Then Exit For
                        Else
                            Exit For
                        End If
                    Next
                    
                    If .TextMatrix(lngRow, Col诊断描述) <> "" Then
                        lngRow = lngRow + 1: .AddItem "", lngRow
                        .TextMatrix(lngRow, col类型) = Split(strSQL, ",")(i)
                    End If
                    
                    If InStr("," & mstr诊断IDs & ",", "," & rsTmp!id & ",") > 0 Then
                        .TextMatrix(lngRow, col关联) = 1
                    End If
                    
                    str诊断Id = str诊断Id & "," & rsTmp!id
                    
                    If IsNull(rsTmp!诊断描述) Then
                        .TextMatrix(lngRow, col诊断编码) = ""
                        .TextMatrix(lngRow, Col诊断描述) = ""
                    Else
                        If Mid(rsTmp!诊断描述, 1, 1) <> "(" Or (Val(rsTmp!诊断ID & "") = 0 And Val(rsTmp!疾病ID & "") = 0) Then '中医的诊断描述后面加了（候症），所以只判断第一个字符
                            '由于疾病编码和诊断可以对应，如果两个都不为空的时候，先判断疾病编码，先取疾病编码
                            If Val(rsTmp!疾病ID & "") <> 0 Then
                                .TextMatrix(lngRow, col诊断编码) = NVL(rsTmp!疾病编码)
                            ElseIf Val(rsTmp!诊断ID & "") <> 0 Then
                                .TextMatrix(lngRow, col诊断编码) = NVL(rsTmp!诊断编码)
                            Else
                                .TextMatrix(lngRow, col诊断编码) = ""
                            End If
                            .TextMatrix(lngRow, Col诊断描述) = rsTmp!诊断描述
                        Else
                            .TextMatrix(lngRow, col诊断编码) = Mid(rsTmp!诊断描述, 2, InStr(rsTmp!诊断描述, ")") - 2)
                            .TextMatrix(lngRow, Col诊断描述) = Mid(rsTmp!诊断描述, InStr(rsTmp!诊断描述, ")") + 1)
                        End If
                    End If
                    If Not IsNull(rsTmp!疾病ID) Or Not IsNull(rsTmp!诊断ID) Then
                        .Cell(flexcpData, lngRow, Col诊断描述) = Get诊断描述(Val("" & rsTmp!诊断ID), Val("" & rsTmp!疾病ID))    '获取原始名称以便修改时判断
                    Else
                        .Cell(flexcpData, lngRow, Col诊断描述) = .TextMatrix(lngRow, Col诊断描述)
                    End If
                    If mint病人来源 = 1 Then
                        .TextMatrix(lngRow, col发病时间) = Format(rsTmp!发病时间 & "", "YYYY-MM-DD HH:mm")
                    Else
                        .TextMatrix(lngRow, col出院情况) = NVL(rsTmp!出院情况)
                        .TextMatrix(lngRow, col入院病情) = NVL(rsTmp!入院病情)
                        .TextMatrix(lngRow, col是否未治) = IIf(NVL(rsTmp!是否未治, 0) = 1, "√", "")
                    End If
                    
                    .TextMatrix(lngRow, col备注) = NVL(rsTmp!备注)
                    .Cell(flexcpData, lngRow, col是否疑诊) = Val(rsTmp!id & "")
                    .TextMatrix(lngRow, col是否疑诊) = IIf(NVL(rsTmp!是否疑诊, 0) = 1, "？", "")
                    .TextMatrix(lngRow, col诊断ID) = NVL(rsTmp!诊断ID, 0)
                    .TextMatrix(lngRow, col疾病ID) = NVL(rsTmp!疾病ID, 0)
                    rsTmp.MoveNext
                Loop
            Next
        End With
    End If

    vsDiagXY.Cell(flexcpForeColor, 1, col是否疑诊, vsDiagXY.Rows - 1, col是否疑诊) = vbRed
    lngRow = GetRow(3)
    If lngRow <> -1 Then
        vsDiagXY.Cell(flexcpBackColor, GetRow(3), vsDiagXY.FixedRows, GetRow(3), vsDiagXY.Cols - 1) = &HC0FFC0
    End If
    vsDiagXY.Cell(flexcpBackColor, 1, col诊断编码, vsDiagXY.Rows - 1, col诊断编码) = ColorUnEditCell      '灰蓝色
    vsDiagXY.Row = 1: vsDiagXY.Col = Col诊断描述
    Call vsDiagXY_AfterRowColChange(-1, -1, vsDiagXY.Row, vsDiagXY.Col)
    vsDiagXY.Tag = "未修改"
    '中医诊断
    '---------------------------------------------------------------
    If mbln中医 Then
        '缺省表格初始化
        With vsDiagZY
            '11-中医门诊诊断;12-中医入院诊断;13-中医出院诊断(主要诊断、其它诊断)
            .ColData(col出院情况) = str治疗结果
            .TextMatrix(1, colzy类型) = 11
            If mint病人来源 = 2 Then
                .TextMatrix(2, colzy类型) = 12
                .TextMatrix(3, colzy类型) = 13
                .TextMatrix(4, colzy类型) = 13
            Else
                .Rows = .FixedRows + 1
            End If
        End With
        If Not bln首页诊断 And mint病人来源 = 2 Then
            strTmp = " And a.记录来源 IN(1,2,3,4) "
        Else
            strTmp = " and a.记录来源=3 "
        End If
    
        '读取各种来源的诊断
        strSQL = "Select a.备注, a.Id, a.病人id, a.主页id, a.医嘱id, a.记录来源, a.诊断次序, a.编码序号, a.病历id, a.诊断类型,a.入院病情," & _
            " a.疾病id, a.诊断id, a.证候id, a.诊断描述,a.出院情况, a.是否未治, a.是否疑诊, a.记录日期, a.记录人, a.取消时间," & _
            " a.取消人, a.病例id, b.编码 As 疾病编码, c.编码 As 诊断编码,d.编码 as 证候编码 ,A.发病时间 From 病人诊断记录 A, 疾病编码目录 B, 疾病诊断目录 C,疾病编码目录 D" & _
            " Where a.疾病id = b.Id(+) And a.诊断id = c.Id(+) And a.证候ID=d.ID(+) And a.诊断类型 IN(11,12,13)" & _
            strTmp & _
            " And 取消时间 Is Null And 病人ID=[1] And 主页ID=[2]" & _
            " Order by a.诊断类型,a.记录来源 Desc,a.诊断次序,a.编码序号,a.ID"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng就诊ID)
        strTmp = ""
        If Not rsTmp.EOF Then
            With vsDiagZY
                If mint病人来源 = 2 Then
                    strSQL = "11,12,13"
                Else
                     strSQL = "11"
                End If
                
                For i = 0 To UBound(Split(strSQL, ","))
                    
                    rsTmp.Filter = "记录来源=3 And 诊断类型=" & Split(strSQL, ",")(i)
                    If mint病人来源 = 2 Then
                        If rsTmp.EOF Then
                            rsTmp.Filter = "记录来源=2 And 诊断类型=" & Split(strSQL, ",")(i)
                        End If
                        If rsTmp.EOF Then
                            rsTmp.Filter = "记录来源=1 And 诊断类型=" & Split(strSQL, ",")(i)
                        End If
                        If rsTmp.EOF Then
                            rsTmp.Filter = "记录来源=4 And 诊断类型=" & Split(strSQL, ",")(i)
                        End If
                    End If
                    Do While Not rsTmp.EOF
                        '确定当前显示行
                        lngRow = .FindRow(CStr(Split(strSQL, ",")(i)), , colzy类型)
                        For j = lngRow To .Rows - 1
                            If Val(.TextMatrix(j, colzy类型)) = Val(Split(strSQL, ",")(i)) Then
                                lngRow = j
                                If .TextMatrix(j, Col诊断描述) = "" Then Exit For
                            Else
                                Exit For
                            End If
                        Next
                        
                        If .TextMatrix(lngRow, Col诊断描述) <> "" Then
                            lngRow = lngRow + 1: .AddItem "", lngRow
                            .TextMatrix(lngRow, colzy类型) = Split(strSQL, ",")(i)
                        End If
                        
                        If InStr("," & mstr诊断IDs & ",", "," & rsTmp!id & ",") > 0 Then
                            .TextMatrix(lngRow, col关联) = 1
                        End If
                        
                        str诊断Id = str诊断Id & "," & rsTmp!id
                        
                        If IsNull(rsTmp!诊断描述) Then
                            .TextMatrix(lngRow, col诊断编码) = ""
                            .TextMatrix(lngRow, Col诊断描述) = ""
                        Else
                            If Mid(rsTmp!诊断描述, 1, 1) <> "(" Or (Val(rsTmp!诊断ID & "") = 0 And Val(rsTmp!疾病ID & "") = 0) Then     '中医的诊断描述后面加了（候症），所以只判断第一个字符
                                '由于疾病编码和诊断可以对应，如果两个都不为空的时候，先判断疾病编码，先取疾病编码
                                If Val(rsTmp!疾病ID & "") <> 0 Then
                                    .TextMatrix(lngRow, col诊断编码) = NVL(rsTmp!疾病编码)
                                ElseIf Val(rsTmp!诊断ID & "") <> 0 Then
                                    .TextMatrix(lngRow, col诊断编码) = NVL(rsTmp!诊断编码)
                                Else
                                    .TextMatrix(lngRow, col诊断编码) = ""
                                End If
                                .TextMatrix(lngRow, Col诊断描述) = rsTmp!诊断描述
                            Else
                                .TextMatrix(lngRow, col诊断编码) = Mid(rsTmp!诊断描述, 2, InStr(rsTmp!诊断描述, ")") - 2)
                                .TextMatrix(lngRow, Col诊断描述) = Mid(rsTmp!诊断描述, InStr(rsTmp!诊断描述, ")") + 1)
                            End If
                        End If

                        .TextMatrix(lngRow, col备注) = NVL(rsTmp!备注)
                       .Cell(flexcpData, lngRow, colZY疑诊) = Val(rsTmp!id & "")
                       .Cell(flexcpData, lngRow, col诊断编码) = .TextMatrix(lngRow, col诊断编码)
                        .TextMatrix(lngRow, colzy诊断ID) = NVL(rsTmp!诊断ID, 0)
                        .TextMatrix(lngRow, colzy疾病ID) = NVL(rsTmp!疾病ID, 0)
                        .TextMatrix(lngRow, colzy证候ID) = NVL(rsTmp!证候id, 0)
                        If mint病人来源 = 1 Then
                            .TextMatrix(lngRow, colZY疑诊) = IIf(NVL(rsTmp!是否疑诊, 0) = 1, "？", "")
                            .TextMatrix(lngRow, col发病时间) = Format(rsTmp!发病时间 & "", "YYYY-MM-DD HH:mm")
                        Else
                            .TextMatrix(lngRow, col出院情况) = NVL(rsTmp!出院情况)
                            .TextMatrix(lngRow, col入院病情) = NVL(rsTmp!入院病情)
                        End If
                        '取证候名称
                        If InStr(.TextMatrix(lngRow, Col诊断描述), "(") > 0 And InStr(.TextMatrix(lngRow, Col诊断描述), ")") > 0 Then
                            strTmp = Mid(.TextMatrix(lngRow, Col诊断描述), InStrRev(.TextMatrix(lngRow, Col诊断描述), "(") + 1)
                            strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
                            '先取证候
                            .TextMatrix(lngRow, col中医证候) = strTmp
                            '去掉诊断描述的证候
                            .TextMatrix(lngRow, Col诊断描述) = Mid(.TextMatrix(lngRow, Col诊断描述), 1, InStrRev(.TextMatrix(lngRow, Col诊断描述), "(") - 1)
                        Else
                           .TextMatrix(lngRow, col中医证候) = ""
                        End If
                        '自由录入诊断的诊断描述，需要去掉证候，因此此句代码后移
                        If Not IsNull(rsTmp!疾病ID) Or Not IsNull(rsTmp!诊断ID) Then
                            .Cell(flexcpData, lngRow, Col诊断描述) = Get诊断描述(Val("" & rsTmp!诊断ID), Val("" & rsTmp!疾病ID))    '获取原始名称以便修改时判断
                        Else
                            .Cell(flexcpData, lngRow, Col诊断描述) = .TextMatrix(lngRow, Col诊断描述)
                        End If
                        rsTmp.MoveNext
                    Loop
                Next
            End With
        End If
        vsDiagZY.Cell(flexcpForeColor, vsDiagZY.FixedRows, colZY疑诊, vsDiagZY.Rows - 1, colZY疑诊) = vbRed
        lngRow = GetRow(13)
        If lngRow <> -1 Then
            vsDiagZY.Cell(flexcpBackColor, GetRow(13), vsDiagZY.FixedRows, GetRow(13), vsDiagZY.Cols - 1) = &HC0FFC0
        End If
        vsDiagZY.Cell(flexcpBackColor, 1, col诊断编码, vsDiagZY.Rows - 1, col诊断编码) = ColorUnEditCell      '灰蓝色
        vsDiagZY.Row = 1: vsDiagZY.Col = Col诊断描述
        Call vsDiagZY_AfterRowColChange(-1, -1, vsDiagXY.Row, vsDiagXY.Col)
        vsDiagZY.Tag = "未修改"
    End If
    '保存诊断医嘱关系
    If str诊断Id <> "" Then
        str诊断Id = Mid(str诊断Id, 2)
        
        strSQL = "Select /*+ RULE*/" & vbNewLine & _
                " F_List2str(Cast(Collect(A.医嘱id || '') As T_Strlist)) As 医嘱ids, A.诊断id" & vbNewLine & _
                "From 病人诊断医嘱 A, 病人医嘱记录 B" & vbNewLine & _
                "Where A.诊断id In (Select Column_Value From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist))) And A.医嘱id = B.Id And" & vbNewLine & _
                "      B.医嘱状态 <> -1 And B.医嘱状态 <> 4" & vbNewLine & _
                "Group By A.诊断id"
                
        Set mrsAdvice = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str诊断Id)
        
        With vsDiagXY
            For i = .FixedRows To .Rows - 1
                If Val(.Cell(flexcpData, i, col是否疑诊) & "") > 0 Then
                    .RowData(i) = GetAdviceIDByDiag(Val(.Cell(flexcpData, i, col是否疑诊) & ""))
                End If
            Next
        End With
        
        If mbln中医 Then
            With vsDiagZY
                For i = .FixedRows To .Rows - 1
                    If Val(.Cell(flexcpData, i, colZY疑诊) & "") > 0 Then
                        .RowData(i) = GetAdviceIDByDiag(Val(.Cell(flexcpData, i, colZY疑诊) & ""))
                    End If
                Next
            End With
        End If
    End If
    Exit Sub
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub SaveData()
   Dim arrSQL As Variant
   Dim intIdx As Integer
   Dim i As Long
   Dim str诊断描述  As String
   Dim datCurDate As Date
   Dim blnTrans As Boolean
   Dim lngID As Long
   Dim blnChange As Boolean
   Dim str关联医嘱ID As String
   
    mstr诊断IDs = ""
    mstr诊断s = ""
    arrSQL = Array()
    datCurDate = gobjComLib.zlDatabase.Currentdate
    blnChange = vsDiagXY.Tag = ""
    '西医诊断
    If blnChange Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_DELETE(" & mlng病人ID & "," & mlng就诊ID & ",3,NULL,'1,2,3,5,6,7,10')"
    End If
    With vsDiagXY
        intIdx = 0
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, Col诊断描述)) <> "" Then
                If Trim(.TextMatrix(i, col诊断编码)) = "" Then
                    str诊断描述 = .TextMatrix(i, Col诊断描述)
                Else
                    str诊断描述 = "(" & .TextMatrix(i, col诊断编码) & ")" & .TextMatrix(i, Col诊断描述)
                End If
                lngID = Val(.Cell(flexcpData, i, col是否疑诊))
                str关联医嘱ID = ""
                If Not mrsAdvice Is Nothing Then
                    mrsAdvice.Filter = "诊断ID=" & lngID
                    If Not mrsAdvice.EOF Then
                        mrsAdvice.MoveFirst
                        str关联医嘱ID = mrsAdvice!医嘱IDs
                    End If
                End If
                
                If Val(.TextMatrix(i, col关联)) <> 0 Then
                    If lngID = 0 Then lngID = gobjComLib.zlDatabase.GetNextId("病人诊断记录")
                    mstr诊断IDs = mstr诊断IDs & "," & lngID
                    mstr诊断s = mstr诊断s & "," & str诊断描述
                End If
                If blnChange Then
                    If Val(.TextMatrix(i, col类型)) <> Val(.TextMatrix(i - 1, col类型)) Then intIdx = 0
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1): intIdx = intIdx + 1
                    arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_INSERT(" & mlng病人ID & "," & mlng就诊ID & ",3,NULL," & _
                        Val(.TextMatrix(i, col类型)) & "," & ZVal(.TextMatrix(i, col疾病ID)) & "," & ZVal(.TextMatrix(i, col诊断ID)) & "," & _
                        "NULL,'" & str诊断描述 & "','" & NeedName(.TextMatrix(i, col出院情况)) & "'," & _
                        IIf(.TextMatrix(i, col是否未治) = "", 0, 1) & "," & IIf(.TextMatrix(i, col是否疑诊) = "", 0, 1) & "," & _
                        "To_Date('" & Format(datCurDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                        IIf(str关联医嘱ID = "", "Null,", "'" & str关联医嘱ID & "',") & intIdx & ",'" & .TextMatrix(i, col备注) & "','" & .TextMatrix(i, col入院病情) & "',Null,'" & mstr开单人 & "'," & IIf(lngID = 0, "Null", lngID) & ")"
                End If
            End If
        Next
    End With
 
    '中医诊断
    If vsDiagZY.Visible Then
        blnChange = vsDiagZY.Tag = ""
        If blnChange Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_DELETE(" & mlng病人ID & "," & mlng就诊ID & ",3,NULL,'11,12,13')"
        End If
        With vsDiagZY
            intIdx = 0
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, Col诊断描述)) <> "" Then
                    If Trim(.TextMatrix(i, col诊断编码)) = "" Then
                        str诊断描述 = .TextMatrix(i, Col诊断描述) & IIf(.TextMatrix(i, col中医证候) <> "", "(" & .TextMatrix(i, col中医证候) & ")", "")
                    Else
                        str诊断描述 = "(" & .TextMatrix(i, col诊断编码) & ")" & .TextMatrix(i, Col诊断描述) & IIf(.TextMatrix(i, col中医证候) <> "", "(" & .TextMatrix(i, col中医证候) & ")", "")
                    End If
                    lngID = Val(.Cell(flexcpData, i, colZY疑诊))
                    str关联医嘱ID = ""
                    If Not mrsAdvice Is Nothing Then
                        mrsAdvice.Filter = "诊断ID=" & lngID
                        If Not mrsAdvice.EOF Then
                            mrsAdvice.MoveFirst
                            str关联医嘱ID = mrsAdvice!医嘱IDs
                        End If
                    End If
                    If Val(.TextMatrix(i, col关联)) <> 0 Then
                        If lngID = 0 Then lngID = gobjComLib.zlDatabase.GetNextId("病人诊断记录")
                        mstr诊断IDs = mstr诊断IDs & "," & lngID
                        mstr诊断s = mstr诊断s & "," & str诊断描述
                    End If
                    If blnChange Then
                        If Val(.TextMatrix(i, colzy类型)) <> Val(.TextMatrix(i - 1, colzy类型)) Then intIdx = 0
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1): intIdx = intIdx + 1
                        arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_INSERT(" & mlng病人ID & "," & mlng就诊ID & ",3,NULL," & _
                            Val(.TextMatrix(i, colzy类型)) & "," & ZVal(.TextMatrix(i, colzy疾病ID)) & "," & ZVal(.TextMatrix(i, colzy诊断ID)) & "," & _
                            ZVal(.TextMatrix(i, colzy证候ID)) & ",'" & str诊断描述 & "','" & NeedName(.TextMatrix(i, col出院情况)) & "'," & _
                            "NULL,NULL,To_Date('" & Format(datCurDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                            IIf(str关联医嘱ID = "", "Null,", "'" & str关联医嘱ID & "',") & intIdx & ",'" & .TextMatrix(i, col备注) & "','" & .TextMatrix(i, col入院病情) & "',Null,'" & mstr开单人 & "'," & IIf(lngID = 0, "Null", lngID) & ")"
                    End If
                End If
            Next
        End With
    End If
    
    If mstr诊断IDs <> "" Then mstr诊断IDs = Mid(mstr诊断IDs, 2)
    If mstr诊断s <> "" Then mstr诊断s = Mid(mstr诊断s, 2)
    
    If vsDiagXY.Tag = "" Or vsDiagZY.Tag = "" And vsDiagZY.Visible Then
        Screen.MousePointer = 11
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call gobjComLib.zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next
        gcnOracle.CommitTrans: blnTrans = False
        
        
        On Error GoTo 0
        Screen.MousePointer = 0
    End If
    
    Exit Sub
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Function Get治疗结果() As String
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo errH
    strSQL = "Select 编码,名称,简码 From 治疗结果 Order by 编码"
    Call gobjComLib.zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)

    strSQL = ""
    Do While Not rsTmp.EOF
        strSQL = strSQL & "|" & rsTmp!编码 & "-" & rsTmp!名称
        rsTmp.MoveNext
    Loop
    If strSQL = "" Then
        Get治疗结果 = "1-治愈|2-好转|3-未愈|4-死亡|5-其他"
    Else
        Get治疗结果 = Mid(strSQL, 2)
    End If
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Private Function NeedName(strList As String) As String
'说明:1-strList以()或[]分割编码与名称时，必须以[编码]或(编码)开头,编码必须为数字或字母
'     2-分隔符有优先级：回车符(Chr(13)）> - > [] > ()

    '优先判断以回车符分割
    If InStr(strList, Chr(13)) > 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, Chr(13)) + 1))
        Exit Function
    End If
    '以[]分割
    If InStr(strList, "]") > 0 And InStr(strList, "-") = 0 And Left(LTrim(strList), 1) = "[" Then
        If gobjComLib.zlCommFun.IsNumOrChar(Mid(strList, 2, InStr(strList, "]") - 2)) Then
            NeedName = LTrim(Mid(strList, InStr(strList, "]") + 1))
            Exit Function
        End If
    End If
    '以()分割
    If InStr(strList, ")") > 0 And InStr(strList, "-") = 0 And Left(LTrim(strList), 1) = "(" Then
        If gobjComLib.zlCommFun.IsNumOrChar(Mid(strList, 2, InStr(strList, ")") - 2)) Then
            NeedName = LTrim(Mid(strList, InStr(strList, ")") + 1))
            Exit Function
        End If
    End If
    '以-分割
    NeedName = LTrim(Mid(strList, InStr(strList, "-") + 1))
    
End Function

Private Function XYCellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    With vsDiagXY
        '隐藏列不可编辑
        If .ColHidden(lngCol) Then Exit Function
        
        If lngCol = col关联 Then
            If Trim(.TextMatrix(lngRow, Col诊断描述)) = "" Then
                Exit Function
            End If
        Else
            If .RowData(lngRow) & "" <> "" Then Exit Function
        End If
        
        If lngCol = Col诊断描述 And (mlngPathState = 1 Or mlngPathState = 2) Then
            If .TextMatrix(lngRow, col诊断类型) = "入院诊断" And mlngDiagnosisType = 2 Or .TextMatrix(lngRow, col诊断类型) = "门诊诊断" And mlngDiagnosisType = 1 Then
                If .TextMatrix(lngRow, Col诊断描述) <> "" And .TextMatrix(lngRow, col诊断类型) <> .TextMatrix(lngRow - 1, col诊断类型) Then
                    '首要诊断不允许改
                    Exit Function
                End If
            End If
            '合并路径
            If Not CheckMergePath(mlng病人ID, mlng就诊ID, Val(.TextMatrix(lngRow, col类型)), Val(.TextMatrix(lngRow, col疾病ID))) Then Exit Function
        End If
        If lngCol = Col诊断描述 Then
            '两条路径以上
            If mstrPathDiag <> "" And mlngPathState > 0 Then
                If InStr("," & mstrPathDiag & ",", "," & .TextMatrix(.Row, col类型) & "|" & Val(.TextMatrix(.Row, col疾病ID)) & "|" & Val(.TextMatrix(.Row, col诊断ID)) & ",") > 0 Then
                    '导入诊断不允许该
                    Exit Function
                End If
            End If
        End If
        If lngCol = Col诊断描述 And mlngPathState = 2 And mblnIsPathOutTime Then
            If .TextMatrix(.Row, col诊断类型) = "出院诊断" And mlngDiagnosisType <= 2 Then
                '正常完成的出院诊断不允许改
                Exit Function
            End If
        End If
        '必须先输入诊断
        If .TextMatrix(lngRow, Col诊断描述) = "" Then
            If lngCol = col出院情况 Or lngCol = col备注 Or lngCol = col是否未治 Or lngCol = col是否疑诊 Or lngCol = col增加 Or lngCol = col发病时间 Then
                Exit Function
            End If
        End If
        If lngCol = col诊断编码 Then Exit Function
        
        If lngCol = col增加 Then
            If Val(.TextMatrix(lngRow, col类型)) = 3 Then
                If .TextMatrix(lngRow, col诊断类型) = "出院诊断" Then Exit Function
            End If
        End If
        
        '出院诊断和院内感染允许输入出院情况(因为可能院内感染在出院时已经好转或治愈了)
        If Val(.TextMatrix(lngRow, col类型)) = 3 Or Val(.TextMatrix(lngRow, col类型)) = 5 Or Val(.TextMatrix(lngRow, col类型)) = 10 Then
            '出院诊断必须依次输入(尚未输入时)
            If .TextMatrix(lngRow, Col诊断描述) = "" And Val(.TextMatrix(lngRow, col类型)) = 3 Then
                If Val(.TextMatrix(lngRow - 1, col类型)) = 3 And .TextMatrix(lngRow - 1, Col诊断描述) = "" Then
                    Exit Function
                End If
            End If

            '出院情况为"其他"时才可以设置是否未治
            If .TextMatrix(lngRow, col出院情况) <> "其他" And lngCol = col是否未治 Then
                Exit Function
            End If
        ElseIf lngCol = col出院情况 Or lngCol = col是否未治 Then
            Exit Function
        End If
        
        '入院病情只能在出院诊断和其他诊断行填写
        If lngCol = col入院病情 Then
            If .TextMatrix(lngRow, col类型) <> "3" Then
                Exit Function
            End If
        End If
    End With
    XYCellEditable = True
End Function

Private Function CheckMergePath(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lngDiagType As Long, ByVal lngDiag As Long) As Boolean
'功能：检查合并路径对应的诊断不能修改
'参数：lngDiagType：诊断类型,lngDiag=疾病ID
    Dim strSQL As String, rsTmp As Recordset
    
    On Error GoTo errH
    If lngDiag = 0 Or lngDiagType = 0 Then CheckMergePath = True: Exit Function
    strSQL = "Select 诊断类型,疾病ID From 病人合并路径 Where 病人ID=[1] And 主页ID=[2]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lng病人ID, lng主页ID)
    Do While Not rsTmp.EOF
        If lngDiagType = Val(rsTmp!诊断类型 & "") And lngDiag = Val(rsTmp!疾病ID & "") Then
            Exit Function
        End If
        rsTmp.MoveNext
    Loop
    CheckMergePath = True
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Private Sub XYEnterNextCell()
    Dim i As Long, j As Long
    
    With vsDiagXY
        '从下一单元开始循环搜索
        For i = .Row To .Rows - 1
            For j = IIf(i = .Row, .Col + 1, Col诊断描述) To col增加
                If XYCellEditable(i, j) And .ColWidth(j) <> 0 Then Exit For
            Next
            If j <= col增加 Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
            .ShowCell .Row, .Col
        Else
            Call gobjComLib.zlCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub

Private Function Decode(ParamArray arrPar() As Variant) As Variant
'功能：模拟Oracle的Decode函数
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Private Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Private Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
'功能：得控件中指定坐标在屏幕中的位置(Twip)
    Dim vPoint As POINTAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

Private Function ZYCellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    With vsDiagZY
        '隐藏列不可编辑
        If .ColHidden(lngCol) Then Exit Function

        If lngCol = col关联 Then
            If Trim(.TextMatrix(lngRow, Col诊断描述)) = "" Then
                Exit Function
            End If
        Else
            If .RowData(lngRow) & "" <> "" Then Exit Function
        End If
        
        If lngCol = Col诊断描述 And (mlngPathState = 1 Or mlngPathState = 2) Then
            If .TextMatrix(lngRow, col诊断类型) = "入院诊断" And mlngDiagnosisType = 12 Or .TextMatrix(lngRow, col诊断类型) = "门诊诊断" And mlngDiagnosisType = 11 Then
                If .TextMatrix(lngRow, Col诊断描述) <> "" And .TextMatrix(lngRow, col诊断类型) <> .TextMatrix(lngRow - 1, col诊断类型) Then
                    '首要诊断不允许改
                    Exit Function
                End If
            End If
            '合并路径
            If Not CheckMergePath(mlng病人ID, mlng就诊ID, Val(.TextMatrix(lngRow, colzy类型)), Val(.TextMatrix(lngRow, colzy疾病ID))) Then Exit Function
        End If
        
        If lngCol = Col诊断描述 Then
            '两条路径以上
            If mstrPathDiag <> "" And mlngPathState > 0 Then
                If InStr("," & mstrPathDiag & ",", "," & .TextMatrix(.Row, colzy类型) & "|" & Val(.TextMatrix(.Row, col疾病ID)) & "|" & Val(.TextMatrix(.Row, col诊断ID)) & ",") > 0 Then
                    '导入诊断不允许该
                    Exit Function
                End If
            End If
        End If
        If lngCol = Col诊断描述 And mlngPathState = 2 And mblnIsPathOutTime Then
            If .TextMatrix(.Row, col诊断类型) = "主要诊断" And mlngDiagnosisType > 10 Then
                '正常完成的出院诊断不允许改
                Exit Function
            End If
        End If
        '必须先输入诊断
        If .TextMatrix(lngRow, Col诊断描述) = "" Then
            If lngCol = col出院情况 Or lngCol = col备注 Or lngCol = colZY增加 Or lngCol = col发病时间 Or lngCol = colZY疑诊 Then Exit Function
        End If
        If lngCol = col诊断编码 Then Exit Function
        
        If lngCol = colZY增加 Then
            If Val(.TextMatrix(lngRow, colzy类型)) = 13 Then
                If .TextMatrix(lngRow, col诊断类型) = "主要诊断" Then Exit Function
            End If
        End If
        
        If Val(.TextMatrix(lngRow, colzy类型)) = 13 Then
            '出院诊断必须依次输入(尚未输入时)
            If .TextMatrix(lngRow, Col诊断描述) = "" Then
                If Val(.TextMatrix(lngRow - 1, colzy类型)) = 13 And .TextMatrix(lngRow - 1, Col诊断描述) = "" Then
                    Exit Function
                End If
            End If
        ElseIf lngCol = col出院情况 Then
            '非出院诊断时不允许输入
            If Val(.TextMatrix(lngRow, colzy类型)) <> 13 Then Exit Function
        End If
        '入院病情只能在主要诊断和其他诊断行填写
        If lngCol = col入院病情 Then
            If .TextMatrix(lngRow, colzy类型) <> "13" Then
                Exit Function
            End If
        End If
        '必须先输诊断再输证候
        If lngCol = col中医证候 Then
            If .TextMatrix(lngRow, Col诊断描述) = "" Then Exit Function
        End If
    End With
    ZYCellEditable = True
End Function

Private Sub ZYSetDiagInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'功能：处理中医诊断项目的输入
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI
    Dim i As Long, j As Long
    Dim strTmp As String
    
    With vsDiagZY
        If Not rsInput Is Nothing Then
            For i = 1 To rsInput.RecordCount
                If i > 1 Then
                    '其他诊断选择多条时的处理
                    If lngRow = .Rows - 1 Then
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, colzy类型) = .TextMatrix(lngRow, colzy类型)
                    End If
                    '确定当前显示行
                    If Val(.TextMatrix(lngRow + 1, colzy类型)) = Val(.TextMatrix(lngRow, colzy类型)) Then
                        For j = lngRow + 1 To .Rows - 1
                            If Val(.TextMatrix(j, colzy类型)) = Val(.TextMatrix(lngRow, colzy类型)) Then
                                lngRow = j
                                If .TextMatrix(j, Col诊断描述) = "" Then Exit For
                            Else
                                Exit For
                            End If
                        Next
                        If .TextMatrix(lngRow, Col诊断描述) <> "" Then
                            lngRow = lngRow + 1: .AddItem "", lngRow
                            .TextMatrix(lngRow, colzy类型) = .TextMatrix(lngRow - 1, colzy类型)
                        End If
                    Else
                        lngRow = lngRow + 1: .AddItem "", lngRow
                        .TextMatrix(lngRow, colzy类型) = .TextMatrix(lngRow - 1, colzy类型)
                    End If
                End If
                
                If InStr(.TextMatrix(lngRow, Col诊断描述), "(") > 0 And InStr(.TextMatrix(lngRow, Col诊断描述), ")") > 0 Then
                    strTmp = Mid(.TextMatrix(lngRow, Col诊断描述), InStrRev(.TextMatrix(lngRow, Col诊断描述), "("))
                End If
                
                .TextMatrix(lngRow, col关联) = 1
                .TextMatrix(lngRow, col诊断编码) = "" & rsInput!编码
                .TextMatrix(lngRow, Col诊断描述) = "" & rsInput!名称 & strTmp
                .Cell(flexcpData, lngRow, Col诊断描述) = .TextMatrix(lngRow, Col诊断描述)
                                
                
                '根据诊断确定疾病,或根据疾病确定诊断
                If optInput(0).value Then
                    .TextMatrix(lngRow, colzy诊断ID) = rsInput!项目ID
                    .TextMatrix(lngRow, colzy疾病ID) = ""
                    strSQL = "Select 疾病ID as ID From 疾病诊断对照 Where 诊断ID=[1]"
                Else
                    .TextMatrix(lngRow, colzy疾病ID) = rsInput!项目ID
                    .TextMatrix(lngRow, colzy诊断ID) = ""
                    strSQL = "Select 诊断ID as ID From 疾病诊断对照 Where 疾病ID=[1]"
                End If
                Set rsTmp = New ADODB.Recordset
                On Error GoTo errH
                Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!项目ID))
                If Not rsTmp.EOF Then
                    If optInput(0).value Then
                        .TextMatrix(lngRow, colzy疾病ID) = NVL(rsTmp!id)
                    Else
                        .TextMatrix(lngRow, colzy诊断ID) = NVL(rsTmp!id)
                    End If
                End If
                
                '中医根据疾病诊断参考取证候
                Call Set中医证候(lngRow, Val(.TextMatrix(lngRow, colzy诊断ID)))
                
                rsInput.MoveNext
            Next
        Else
            .TextMatrix(lngRow, col诊断编码) = ""
            .TextMatrix(lngRow, Col诊断描述) = .EditText
            .Cell(flexcpData, lngRow, Col诊断描述) = .TextMatrix(lngRow, Col诊断描述)
            .TextMatrix(lngRow, colzy诊断ID) = ""
            .TextMatrix(lngRow, colzy疾病ID) = ""
            .TextMatrix(lngRow, colzy证候ID) = ""
        End If
        .Cell(flexcpForeColor, .FixedRows, colZY疑诊, .Rows - 1, colZY疑诊) = vbRed
    End With
    Exit Sub
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub ZYEnterNextCell()
    Dim i As Long, j As Long
    
    With vsDiagZY
        '从下一单元开始循环搜索
        For i = .Row To .Rows - 1
            For j = IIf(i = .Row, .Col + 1, Col诊断描述) To colZY增加
                If ZYCellEditable(i, j) And .ColWidth(j) <> 0 Then Exit For
            Next
            If j <= colZY增加 Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
            .ShowCell .Row, .Col
        Else
            Call gobjComLib.zlCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub

Private Function Set中医证候(ByVal lngRow As Long, ByVal lng诊断ID As Long, Optional ByVal rsInput As Recordset) As Boolean
'功能：中医根据疾病诊断参考取证候
'参数：rsInput-如果不为空，则输出指定的中药证候记录集
'返回：是否有对应关系
    Dim rsTmp As Recordset
    Dim strSQL As String
    Dim blnCancel As Boolean
    Dim vPoint As POINTAPI
    Dim strTmp As String
    
    With vsDiagZY
        '去掉已有的证候
        If InStr(.TextMatrix(lngRow, Col诊断描述), "(") > 0 And InStr(.TextMatrix(lngRow, Col诊断描述), ")") > 0 Then
            strTmp = Mid(.TextMatrix(lngRow, Col诊断描述), 1, InStrRev(.TextMatrix(lngRow, Col诊断描述), "(") - 1)
        Else
            strTmp = .TextMatrix(lngRow, Col诊断描述)
        End If
        If rsInput Is Nothing Then
            If lng诊断ID <> 0 Then
                strSQL = "Select Distinct a.证候序号 as ID,a.证候ID,a.证候名称,b.编码 as 证候编码" & _
                    " From 疾病诊断参考 A,疾病编码目录 B" & _
                    " Where a.证候ID=b.ID(+) And a.诊断ID=[1] And a.证候名称 is Not NULL" & _
                    " Order by a.证候序号"
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = Nothing
                Set rsTmp = gobjComLib.zlDatabase.ShowSQLSelect(Me, strSQL, 0, "中医证候", False, "", "", False, False, True, _
                    vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, lng诊断ID)
                If Not rsTmp Is Nothing Then
                    .TextMatrix(lngRow, colzy证候ID) = NVL(rsTmp!证候id)
                    If Not IsNull(rsTmp!证候名称) Then
                        .TextMatrix(lngRow, Col诊断描述) = strTmp
                        .Cell(flexcpData, lngRow, Col诊断描述) = .TextMatrix(lngRow, Col诊断描述)
                        .TextMatrix(lngRow, col中医证候) = NVL(rsTmp!证候名称)
                        .Cell(flexcpData, lngRow, col中医证候) = .TextMatrix(lngRow, col中医证候)
                        If .EditText <> "" Then .EditText = .TextMatrix(lngRow, col中医证候)
                        mblnChange = True
                        .Tag = ""
                    End If
                    Set中医证候 = True
                Else
                    If blnCancel Then
                        Set中医证候 = True
                        If .EditText <> "" Then .EditText = .Cell(flexcpData, lngRow, col中医证候)
                    Else
                        Set中医证候 = False
                    End If
                End If
            Else
                Set中医证候 = False
            End If
        Else
            .TextMatrix(lngRow, colzy证候ID) = NVL(rsInput!项目ID)
            .TextMatrix(lngRow, Col诊断描述) = strTmp
            .Cell(flexcpData, lngRow, Col诊断描述) = .TextMatrix(lngRow, Col诊断描述)
            .TextMatrix(lngRow, col中医证候) = NVL(rsInput!名称)
            .Cell(flexcpData, lngRow, col中医证候) = .TextMatrix(lngRow, col中医证候)
            If .EditText <> "" Then .EditText = .TextMatrix(lngRow, col中医证候)
        End If
    End With
End Function

Private Function Get诊断描述(ByVal lng诊断ID As Long, ByVal lng疾病ID As Long) As String
'功能：根据诊断ID或疾病ID获取字典表中的名称（病人诊断记录中的名称可以是修改后的,允许加前缀或后缀），以便再次修改时判断
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    If lng诊断ID <> 0 Then
        strSQL = "Select 名称 From 疾病诊断目录 Where ID = [1]"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lng诊断ID)
        If rsTmp.RecordCount > 0 Then Get诊断描述 = "" & rsTmp!名称
    ElseIf lng疾病ID <> 0 Then
        strSQL = "Select 名称 From 疾病编码目录 Where ID = [1]"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lng疾病ID)
        If rsTmp.RecordCount > 0 Then Get诊断描述 = "" & rsTmp!名称
    End If
    
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Private Function GetRow(ByVal lng诊断类型 As Long) As Long
'功能：返回指定诊断类型的第一诊断行
    If InStr(",11,12,13,", "," & lng诊断类型 & ",") > 0 Then
        GetRow = vsDiagZY.FindRow(CStr(lng诊断类型), , colzy类型)
    Else
        GetRow = vsDiagXY.FindRow(CStr(lng诊断类型), , col类型)
    End If
End Function

Private Function ZVal(ByVal varValue As Variant) As String
'功能：将0零转换为"NULL"串,在生成SQL语句时用
    ZVal = IIf(Val(varValue) = 0, "NULL", Val(varValue))
End Function

Private Function GetAdviceIDByDiag(ByVal lng诊断ID As Long) As String
'功能：根据诊断ID获取诊断相关医嘱ID
    Dim strTmp As String, str医嘱IDs As String
    Dim lngPos As Long
    If Not mrsAdvice Is Nothing Then
        mrsAdvice.Filter = "诊断ID=" & lng诊断ID
        If Not mrsAdvice.EOF Then
            mrsAdvice.MoveFirst
            str医嘱IDs = mrsAdvice!医嘱IDs
            lngPos = InStr(str医嘱IDs, mlng组医嘱ID & "")
            If str医嘱IDs = mlng组医嘱ID & "" Then
            '关联医嘱为当前医嘱，不做处理，返回空串
            ElseIf lngPos <= 0 Then
            '当前医嘱未关联当前诊断
                strTmp = str医嘱IDs
            Else
            '医嘱ID串包含逗号的情况，可通过字符串替换。
                If lngPos = 1 Then
                '当前医嘱在开头位置
                    strTmp = Replace(str医嘱IDs, mlng组医嘱ID & ",", "")
                Else
                '当前医嘱在非开头位置
                    strTmp = Replace(str医嘱IDs, "," & mlng组医嘱ID, "")
                End If
            End If
        End If
    End If
    
    With grsDiagConn
        .Filter = "诊断ID=" & lng诊断ID
        .Sort = "标识ID"
        Do While Not .EOF
            If Val(!标识ID & "") <> mlngCur标识 Then
                strTmp = strTmp & "," & !标识ID
            End If
            .MoveNext
        Loop
    End With
    
    GetAdviceIDByDiag = strTmp
End Function

Private Function CheckData() As Boolean
    Dim i As Long
    Dim j As Long
    Dim curDate As Date
    
    curDate = gobjComLib.zlDatabase.Currentdate
    
    With vsDiagXY
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, Col诊断描述) <> "" And .TextMatrix(i - 1, Col诊断描述) = "" _
                And Val(.TextMatrix(i, col类型)) = Val(.TextMatrix(i - 1, col类型)) Then
                .Row = i - 1: .Col = Col诊断描述
                Call ShowMessage(vsDiagXY, "请依次输入诊断信息。")
                Exit Function
            End If
            
            If Trim(.TextMatrix(i, Col诊断描述)) <> "" Then
                If gobjComLib.zlCommFun.ActualLen(.TextMatrix(i, Col诊断描述)) > 200 Then
                    .Row = i: .Col = Col诊断描述
                    Call ShowMessage(vsDiagXY, IIf(.TextMatrix(i, col诊断类型) = "", "出院诊断", .TextMatrix(i, col诊断类型)) & "内容太长，只允许200个字符或100个汉字。")
                    Exit Function
                End If
                If .TextMatrix(i, col发病时间) <> "" And Not .ColHidden(col发病时间) Then
                    If Format(curDate, "YYYY-MM-DD HH:mm") < Format(.TextMatrix(i, col发病时间), "YYYY-MM-DD HH:mm") Then
                         .Row = i: .Col = col发病时间
                        Call ShowMessage(vsDiagXY, "发病时间应该早于当前时间。")
                        Exit Function
                    End If
                End If
                If gobjComLib.zlCommFun.ActualLen(.TextMatrix(i, col备注)) > 50 Then
                    .Row = i: .Col = col备注
                    Call ShowMessage(vsDiagXY, """" & .TextMatrix(i, Col诊断描述) & """的备注内容太长，只允许50个字符或25个汉字。")
                    Exit Function
                End If
                If Val(.TextMatrix(i, col类型)) = 5 Then     '院内感染
                    If .TextMatrix(i, col出院情况) = "" And Not .ColHidden(col出院情况) Then
                        .Row = i: .Col = col出院情况
                        If ShowMessage(vsDiagXY, "院内感染的出院情况没有填写，是否继续？", True) = vbNo Then Exit Function
                    End If
                End If
                If Val(.TextMatrix(i, col类型)) = 3 Then
                    If .TextMatrix(i, col出院情况) = "" And Not .ColHidden(col出院情况) Then
                        .Row = i: .Col = col出院情况
                        Call ShowMessage(vsDiagXY, "请填写出院诊断的出院情况。")
                        Exit Function
                    ElseIf Val(.TextMatrix(i - 1, col类型)) <> 3 And InStr(.TextMatrix(i, col出院情况), "其他") > 0 And mbln手术 And Not .ColHidden(col出院情况) Then
                        .Row = i: .Col = col出院情况
                        If ShowMessage(vsDiagXY, "该病人进行了手术，但出院情况选择为其他。是否继续？", True) = vbNo Then Exit Function
                    ElseIf Val(.TextMatrix(i - 1, col类型)) = 3 And InStr(.TextMatrix(GetRow(3), col出院情况), "死亡") = 0 And InStr(.TextMatrix(i, col出院情况), "死亡") > 0 And Not .ColHidden(col出院情况) Then
                        .Row = i: .Col = col出院情况
                        Call ShowMessage(vsDiagXY, "主要诊断的出院情况不为死亡，但其它诊断的出院情况却为死亡。")
                        Exit Function
                    ElseIf .TextMatrix(i, col诊断类型) = "出院诊断" And Not .ColHidden(col出院情况) Then
                        If mlng损伤中毒 <> 0 Then
                            '主要诊断需要有损伤的外部原因
                            If InStr("ST", Left(.TextMatrix(i, col诊断编码), 1)) > 0 And Left(.TextMatrix(i, col诊断编码), 1) <> "" Then
                                '需要损伤中毒外部原因
                                If .TextMatrix(GetRow(7), Col诊断描述) = "" Then
                                    If Not vsDiagZY.Visible Then
                                        .Row = GetRow(7): .Col = Col诊断描述
                                        If mlng损伤中毒 = 1 Then
                                            Call ShowMessage(vsDiagXY, "请填写损伤中毒的原因。")
                                            Exit Function
                                        Else
                                            If ShowMessage(vsDiagXY, "没有填写损伤中毒的原因,是否继续？", True) = vbNo Then Exit Function
                                        End If
                                    End If
                                End If
                            Else
                                If .TextMatrix(GetRow(7), Col诊断描述) <> "" Then
                                    .Row = GetRow(7): .Col = Col诊断描述
                                    If mlng损伤中毒 = 1 Then
                                        Call ShowMessage(vsDiagXY, "不能填写损伤中毒的原因。")
                                        Exit Function
                                    Else
                                        If ShowMessage(vsDiagXY, "出院诊断与损伤中毒的原因不符,是否继续？", True) = vbNo Then Exit Function
                                    End If
                                End If
                            End If
                        End If
                        If mlng病理诊断 <> 0 Then
                            '主要诊断需要填写病理诊断的外部原因
                            If InStr("CD", Left(.TextMatrix(i, col诊断编码), 1)) > 0 And Left(.TextMatrix(i, col诊断编码), 1) <> "" Then
                                '需要病理诊断的外部原因
                                If .TextMatrix(GetRow(6), Col诊断描述) = "" Then
                                    If Not vsDiagZY.Visible Then
                                        .Row = GetRow(6): .Col = Col诊断描述
                                        If mlng病理诊断 = 1 Then
                                            Call ShowMessage(vsDiagXY, "请填写病理诊断。")
                                            Exit Function
                                        Else
                                            If ShowMessage(vsDiagXY, "没有填写病理诊断,是否继续？", True) = vbNo Then Exit Function
                                        End If
                                    End If
                                End If
                            Else
                                If .TextMatrix(GetRow(6), Col诊断描述) <> "" Then
                                    .Row = GetRow(6): .Col = Col诊断描述
                                    If mlng病理诊断 = 1 Then
                                        Call ShowMessage(vsDiagXY, "不能填写病理诊断。")
                                        Exit Function
                                    Else
                                        If ShowMessage(vsDiagXY, "出院诊断与病理诊断不符,是否继续？", True) = vbNo Then Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    For j = GetRow(3) To .Rows - 1
                        If Val(.TextMatrix(j, col类型)) = 3 Then
                            If j <> i And .TextMatrix(j, Col诊断描述) <> "" Then
                                If .TextMatrix(j, Col诊断描述) = .TextMatrix(i, Col诊断描述) Then
                                    .Row = i: .Col = Col诊断描述
                                    Call ShowMessage(vsDiagXY, "发现存在两行相同的出院诊断信息。")
                                    Exit Function
                                ElseIf Val(.TextMatrix(i, col疾病ID)) <> 0 Then
                                    If Val(.TextMatrix(j, col疾病ID)) = Val(.TextMatrix(i, col疾病ID)) Then
                                        .Row = i: .Col = Col诊断描述
                                        Call ShowMessage(vsDiagXY, "发现存在两行相同的出院诊断信息。")
                                        Exit Function
                                    End If
                                ElseIf Val(.TextMatrix(i, col诊断ID)) <> 0 Then
                                    If Val(.TextMatrix(j, col诊断ID)) = Val(.TextMatrix(i, col诊断ID)) Then
                                        .Row = i: .Col = Col诊断描述
                                        Call ShowMessage(vsDiagXY, "发现存在两行相同的出院诊断信息。")
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        Next
    End With
        
    If vsDiagZY.Visible Then
        With vsDiagZY
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, Col诊断描述) <> "" And .TextMatrix(i - 1, Col诊断描述) = "" _
                    And Val(.TextMatrix(i, colzy类型)) = Val(.TextMatrix(i - 1, colzy类型)) Then
                    .Row = i - 1: .Col = Col诊断描述
                    Call ShowMessage(vsDiagZY, "请依次输入诊断信息。")
                    Exit Function
                End If
            
                If Trim(.TextMatrix(i, Col诊断描述)) <> "" Then
                    If gobjComLib.zlCommFun.ActualLen(.TextMatrix(i, Col诊断描述)) > 200 Then
                        .Row = i: .Col = Col诊断描述
                        Call ShowMessage(vsDiagZY, IIf(.TextMatrix(i, col诊断类型) = "", "出院诊断", .TextMatrix(i, col诊断类型)) & "内容太长，只允许200个字符或100个汉字。")
                        Exit Function
                    End If
                    If .TextMatrix(i, col发病时间) <> "" And Not .ColHidden(col发病时间) Then
                        If Format(curDate, "YYYY-MM-DD HH:mm") < Format(.TextMatrix(i, col发病时间), "YYYY-MM-DD HH:mm") Then
                             .Row = i: .Col = col发病时间
                            Call ShowMessage(vsDiagXY, "发病时间应该早于当前时间。")
                            Exit Function
                        End If
                    End If
                    If gobjComLib.zlCommFun.ActualLen(.TextMatrix(i, col备注)) > 50 Then
                        .Row = i: .Col = col备注
                        Call ShowMessage(vsDiagZY, """" & .TextMatrix(i, Col诊断描述) & """的备注内容太长，只允许50个字符或25个汉字。")
                        Exit Function
                    End If
                    If Val(.TextMatrix(i, colzy类型)) = 13 Then
                        If .TextMatrix(i, col出院情况) = "" And Not .ColHidden(col出院情况) Then
                            .Row = i: .Col = col出院情况
                            Call ShowMessage(vsDiagZY, "请填写出院诊断的出院情况。")
                            Exit Function
                        ElseIf Val(.TextMatrix(i - 1, colzy类型)) = 13 And InStr(.TextMatrix(GetRow(13), col出院情况), "死亡") = 0 And InStr(.TextMatrix(i, col出院情况), "死亡") > 0 And Not .ColHidden(col出院情况) Then
                            .Row = i: .Col = col出院情况
                            Call ShowMessage(vsDiagZY, "主要诊断的出院情况不为死亡，但其它诊断的出院情况却为死亡。")
                            Exit Function
                        End If
                        
                        For j = GetRow(13) To .Rows - 1
                            If j <> i And .TextMatrix(j, Col诊断描述) <> "" Then
                                If .TextMatrix(j, Col诊断描述) = .TextMatrix(i, Col诊断描述) Then
                                    .Row = i: .Col = Col诊断描述
                                    Call ShowMessage(vsDiagZY, "发现存在两行相同的出院诊断信息。")
                                    Exit Function
                                ElseIf Val(.TextMatrix(i, colzy疾病ID)) <> 0 Then
                                    If Val(.TextMatrix(j, colzy疾病ID)) = Val(.TextMatrix(i, colzy疾病ID)) Then
                                        .Row = i: .Col = Col诊断描述
                                        Call ShowMessage(vsDiagZY, "发现存在两行相同的出院诊断信息。")
                                        Exit Function
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            Next
        End With
    End If
    CheckData = True
End Function

Private Function ShowMessage(objTmp As Object, ByVal strMsg As String, Optional ByVal blnAsk As Boolean) As VbMsgBoxResult
'功能：显示提示信息并定位在输入项目上
    Dim lngColor As Long
    
    lngColor = objTmp.CellBackColor: objTmp.CellBackColor = &HC0C0FF
    Call objTmp.ShowCell(objTmp.Row, objTmp.Col)
    If Not blnAsk Then
        MsgBox strMsg, vbInformation, gstrSysName
    Else
        ShowMessage = MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
    End If

    objTmp.CellBackColor = lngColor
    If objTmp.Enabled And objTmp.Visible Then objTmp.SetFocus
    Me.Refresh
End Function

Private Function SetPublicFontSize(ByVal bytSize As Byte, Optional ByVal strOther As String)
'功能：设置窗体及所有控件的字体大小
'参数：
'      bytSize:设置为9号字体,0:设置为9号字体,1,设置为12号字体
'      strOther:不进行字体设置的控件父容器的集合,格式为：容器名字1,容器名字2,容器名字3,....
'说明：1.如果涉及到VsFlexGrid等表格控件，需要根据所在的环境重新调整列宽和行高
'      2.如果存在未列出的其他控件或自定义控件,需要用特定方法指定字体大小及相关处理的，需另外单独设置

    Dim objCtrol As Control
    Dim CtlFont As StdFont
    Dim i As Long, lngOldSize As Long
    Dim lngFontSize As Long
    Dim dblRate As Double
    Dim blnDo As Boolean
    Dim strContainer As String
    
    lngFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Me.FontSize = lngFontSize
    strOther = "," & strOther & ","
    blnDo = False
        
    For Each objCtrol In Me.Controls
        Select Case TypeName(objCtrol)
            Case "TabStrip", "Label", "ComboBox", "ListView", "OptionButton", "CheckBox", "DTPicker", "TextBox", "ReportControl", _
                "DockingPane", "CommandBars", "TabControl", "CommandButton", "Frame", "RichTextBox", "MaskEdBox", "IDKind"
                blnDo = True
            Case Else
                blnDo = False
        End Select
        
        If strOther <> ",," And blnDo Then
            '对于CommandBars用户自定义控件读取objCtrol.Container会出错
            strContainer = ""
            On Error Resume Next
            strContainer = objCtrol.Container.Name
            Err.Clear: On Error GoTo 0
            If InStr(1, strOther, "," & strContainer & ",") > 0 Then
                 blnDo = False
            End If
        End If
        
        If blnDo Then
            Select Case TypeName(objCtrol)
                Case "TabStrip"
                        objCtrol.Font.Size = lngFontSize
                Case "Label"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Height = Me.TextHeight("字") + 20
                        'Label宽度需要自行调整
               Case "ComboBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                Case "ListView"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        For i = 1 To objCtrol.ColumnHeaders.count
                            objCtrol.ColumnHeaders(i).Width = objCtrol.ColumnHeaders(i).Width * dblRate
                        Next
                Case "OptionButton"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = Me.TextWidth("字体" & objCtrol.Caption)
                        objCtrol.Height = objCtrol.Height * dblRate
                Case "CheckBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                Case "DTPicker"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = Me.TextWidth("2012-01-01    ")
                        objCtrol.Height = Me.TextHeight("字") + IIf(bytSize = 0, 100, 120)
                Case "TextBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                        objCtrol.Height = Me.TextHeight("字")
                Case "MaskEdBox"
                        objCtrol.FontSize = lngFontSize
                        objCtrol.Width = Me.TextWidth(objCtrol.Mask)
                        objCtrol.Height = Me.TextHeight("字")
                Case "DockingPane"
                        Set CtlFont = objCtrol.PaintManager.CaptionFont
                        If CtlFont Is Nothing Then '控件初始加载时CtlFont为nothing
                            Set CtlFont = Me.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.CaptionFont = CtlFont
                        
                        Set CtlFont = objCtrol.TabPaintManager.Font
                        If CtlFont Is Nothing Then '控件初始加载时CtlFont为nothing
                            Set CtlFont = Me.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.TabPaintManager.Font = CtlFont
        
                        Set CtlFont = objCtrol.PanelPaintManager.Font
                        If CtlFont Is Nothing Then '控件初始加载时CtlFont为nothing
                            Set CtlFont = Me.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PanelPaintManager.Font = CtlFont
                Case "CommandBars"
                        Set CtlFont = objCtrol.Options.Font
                        If CtlFont Is Nothing Then '控件初始加载时CtlFont为nothing
                            Set CtlFont = Me.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.Options.Font = CtlFont
                Case "TabControl"
                        Set CtlFont = objCtrol.PaintManager.Font
                        If CtlFont Is Nothing Then  '控件初始加载时CtlFont为nothing
                            Set CtlFont = Me.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.Font = CtlFont
                        objCtrol.PaintManager.Layout = xtpTabLayoutAutoSize
                Case "CommandButton"
                        lngOldSize = objCtrol.FontSize
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.FontSize = lngFontSize
                        objCtrol.Width = dblRate * objCtrol.Width
                        objCtrol.Height = dblRate * objCtrol.Height
                Case "Frame"
                        objCtrol.FontSize = lngFontSize
                Case "IDKind"
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = dblRate * objCtrol.Width
                        objCtrol.Height = dblRate * objCtrol.Height
            End Select
        End If
    Next
End Function

Private Function GetFullDate(ByVal strText As String, Optional blnTime As Boolean = True) As String
'功能：根据输入的日期简串,返回完整的日期串(yyyy-MM-dd[ HH:mm])
'参数：blnTime=是否处理时间部份
    Dim curDate As Date, strTmp As String
    
    If strText = "" Then Exit Function
    curDate = gobjComLib.zlDatabase.Currentdate
    strTmp = strText
    
    If InStr(strTmp, "-") > 0 Or InStr(strTmp, "/") Or InStr(strTmp, ":") > 0 Then
        '输入串中包含日期分隔符
        If IsDate(strTmp) Then
            strTmp = Format(strTmp, "yyyy-MM-dd HH:mm")
            If Right(strTmp, 5) = "00:00" And InStr(strText, ":") = 0 Then
                '只输入了日期部份
                strTmp = Mid(strTmp, 1, 11) & Format(curDate, "HH:mm")
            ElseIf Left(strTmp, 10) = "1899-12-30" Then
                '只输入了时间部份
                strTmp = Format(curDate, "yyyy-MM-dd") & Right(strTmp, 6)
            End If
        Else
            '输入非法日期,返回原内容
            strTmp = strText
        End If
    Else
        '不包含日期分隔符
        If Len(strTmp) <= 2 Then
            '当作输入dd
            strTmp = Format(strTmp, "00")
            strTmp = Format(curDate, "yyyy-MM") & "-" & strTmp & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 4 Then
            '当作输入MMdd
            strTmp = Format(strTmp, "0000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 6 Then
            '当作输入yyMMdd
            strTmp = Format(strTmp, "000000")
            strTmp = Format(Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & "-" & Right(strTmp, 2), "yyyy-MM-dd") & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 8 Then
            '当作输入MMddHHmm
            strTmp = Format(strTmp, "00000000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & " " & Mid(strTmp, 5, 2) & ":" & Right(strTmp, 2)
            If Not IsDate(strTmp) Then
                '当作输入yyyyMMdd
                strTmp = Format(strText, "00000000")
                strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
            End If
        Else
            '当作输入yyyyMMddHHmm
            strTmp = Format(strTmp, "000000000000")
            strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & Mid(strTmp, 9, 2) & ":" & Right(strTmp, 2)
        End If
    End If
    
    If IsDate(strTmp) And Not blnTime Then
        strTmp = Format(strTmp, "yyyy-MM-dd")
    End If
    GetFullDate = strTmp
End Function
