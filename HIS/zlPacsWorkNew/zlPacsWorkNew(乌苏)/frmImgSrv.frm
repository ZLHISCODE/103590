VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmImgSrv 
   BorderStyle     =   0  'None
   Caption         =   "影像接收服务"
   ClientHeight    =   6390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fraReceiveSet 
      ForeColor       =   &H00FF0000&
      Height          =   6315
      Left            =   90
      TabIndex        =   6
      Top             =   0
      Width           =   11310
      Begin VB.ComboBox cboDevice 
         Height          =   300
         ItemData        =   "frmImgSrv.frx":0000
         Left            =   1380
         List            =   "frmImgSrv.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   345
         Width           =   1815
      End
      Begin VB.ComboBox cboEncode 
         Height          =   300
         ItemData        =   "frmImgSrv.frx":0004
         Left            =   1380
         List            =   "frmImgSrv.frx":0014
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Frame FraAuto 
         Caption         =   "自动匹配设置"
         Height          =   2130
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   10995
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1140
            Left            =   5670
            TabIndex        =   16
            Top             =   690
            Width           =   4065
            Begin VB.OptionButton optMatch 
               Caption         =   "按 ""医嘱ID"" 匹配"
               Height          =   195
               Index           =   2
               Left            =   690
               TabIndex        =   19
               ToolTipText     =   "按医嘱ID将病人和接收的影像进行匹配"
               Top             =   780
               Width           =   3300
            End
            Begin VB.OptionButton optMatch 
               Caption         =   "按 ""检查号"" 匹配"
               Height          =   195
               Index           =   0
               Left            =   690
               TabIndex        =   18
               ToolTipText     =   "按检查号将病人和接收的影像进行匹配"
               Top             =   240
               Width           =   3300
            End
            Begin VB.OptionButton optMatch 
               Caption         =   "按 ""病人标识号(门诊/住院号)"" 匹配"
               Height          =   195
               Index           =   1
               Left            =   690
               TabIndex        =   17
               ToolTipText     =   "按病人标识号将病人和接收的影像进行匹配"
               Top             =   510
               Width           =   3300
            End
            Begin VB.Label lblDataItem 
               Caption         =   "数据库项目"
               Height          =   885
               Left            =   90
               TabIndex        =   20
               Top             =   150
               Width           =   225
            End
            Begin VB.Line Line5 
               X1              =   345
               X2              =   510
               Y1              =   585
               Y2              =   585
            End
            Begin VB.Line Line6 
               X1              =   510
               X2              =   510
               Y1              =   315
               Y2              =   890
            End
            Begin VB.Line Line7 
               X1              =   510
               X2              =   630
               Y1              =   315
               Y2              =   315
            End
            Begin VB.Line Line8 
               X1              =   525
               X2              =   630
               Y1              =   870
               Y2              =   870
            End
         End
         Begin VB.OptionButton optImgMatch 
            Caption         =   "Accession Number"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   10
            Top             =   1155
            Width           =   1740
         End
         Begin VB.OptionButton optImgMatch 
            Caption         =   "Patient Name"
            Height          =   255
            Index           =   2
            Left            =   1080
            TabIndex        =   9
            Top             =   1425
            Width           =   1740
         End
         Begin VB.OptionButton optImgMatch 
            Caption         =   "Patient ID"
            Height          =   255
            Index           =   0
            Left            =   1080
            TabIndex        =   8
            Top             =   885
            Width           =   1740
         End
         Begin VB.ComboBox cboMatchOther 
            Height          =   300
            ItemData        =   "frmImgSrv.frx":0045
            Left            =   8550
            List            =   "frmImgSrv.frx":004F
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   330
            Width           =   1785
         End
         Begin VB.CheckBox chkMatchStudyUID 
            Caption         =   "启用 ""检查UID"" 匹配"
            Height          =   300
            Left            =   120
            TabIndex        =   2
            Top             =   330
            Width           =   2100
         End
         Begin VB.CheckBox chkImageType 
            Caption         =   "根据图像类型拆分序列"
            Height          =   300
            Left            =   4170
            TabIndex        =   3
            Top             =   330
            Width           =   2130
         End
         Begin VB.Line Line4 
            X1              =   930
            X2              =   1035
            Y1              =   1545
            Y2              =   1545
         End
         Begin VB.Line Line3 
            X1              =   915
            X2              =   1035
            Y1              =   990
            Y2              =   990
         End
         Begin VB.Line Line2 
            X1              =   915
            X2              =   915
            Y1              =   990
            Y2              =   1565
         End
         Begin VB.Line Line1 
            X1              =   735
            X2              =   900
            Y1              =   1260
            Y2              =   1260
         End
         Begin VB.Label lblImgItem 
            Caption         =   "图像项目"
            Height          =   690
            Left            =   480
            TabIndex        =   11
            Top             =   930
            Width           =   225
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "辅助匹配(&A)"
            Height          =   180
            Left            =   7545
            TabIndex        =   4
            ToolTipText     =   "该参数针对[数据库项目]按""病人标识号""/""检查号""匹配有效"
            Top             =   390
            Width           =   990
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgList 
         Height          =   1290
         Left            =   5010
         TabIndex        =   14
         Top             =   225
         Width           =   6150
         _cx             =   10848
         _cy             =   2275
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
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
      Begin VB.Line Line12 
         X1              =   4545
         X2              =   4710
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line11 
         X1              =   4710
         X2              =   4710
         Y1              =   570
         Y2              =   1145
      End
      Begin VB.Line Line10 
         X1              =   4710
         X2              =   4830
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Line Line9 
         X1              =   4725
         X2              =   4830
         Y1              =   1125
         Y2              =   1125
      End
      Begin VB.Label lblRoute 
         Caption         =   "自动转发设置"
         Height          =   1080
         Left            =   4335
         TabIndex        =   15
         Top             =   330
         Width           =   225
      End
      Begin VB.Label LblCmp 
         AutoSize        =   -1  'True
         Caption         =   "压缩方式(&Y)"
         Height          =   180
         Left            =   240
         TabIndex        =   1
         Top             =   1260
         Width           =   990
      End
      Begin VB.Label lblSave 
         AutoSize        =   -1  'True
         Caption         =   "存储设备(&F)"
         Height          =   180
         Left            =   240
         TabIndex        =   0
         Top             =   405
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmImgSrv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngSrvID As Long

'自动路由的参数和常量
Private str自动路由目的地 As String
Private str自动路由压缩方式 As String
Private str自动路由目录结构 As String
Private Const AR不压缩 = "不压缩"
Private Const AR现有压缩 = "按当前方式压缩"
Private Const AR检查级别 = "检查级别(默认)"
Private Const AR序列级别 = "序列级别(3D)"


Public Sub ShowRefresh(ByVal SrvID As Long)
    mlngSrvID = SrvID
    If mlngSrvID = 0 Then
        fraReceiveSet.Caption = "上方列表中所选服务尚未保存，不能进行设置！"
        fraReceiveSet.Enabled = False
    Else
        fraReceiveSet.Caption = ""
        fraReceiveSet.Enabled = True
    End If
    RefreshPara
End Sub

Private Sub RefreshPara()
Dim rsTemp As New ADODB.Recordset, i As Integer
    gstrSQL = "select 服务ID,参数名称 ,参数值 from 影像DICOM服务参数 where 服务ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取参数", mlngSrvID)
    InitvfgList
    cboDevice.ListIndex = -1
    cboEncode.ListIndex = -1
    chkImageType.value = False
    chkMatchStudyUID.value = False
    cboMatchOther.ListIndex = 0
    str自动路由目的地 = ""
    str自动路由压缩方式 = ""
    str自动路由目录结构 = ""
    
    Do Until rsTemp.EOF
        Select Case rsTemp!参数名称
            Case "存储设备"
                Call SeekIndexWithNo(cboDevice, Nvl(rsTemp!参数值), True)
            Case "压缩方式"
                Call SeekIndex(cboEncode, Nvl(rsTemp!参数值), True)
            Case "按检查UID匹配"
                chkMatchStudyUID.value = rsTemp!参数值
            Case "按类型拆分序列"
                chkImageType.value = rsTemp!参数值
            Case "匹配图像项目"
                optImgMatch(Nvl(rsTemp!参数值, 0)) = True
            Case "匹配数据库项目"
                optMatch(Nvl(rsTemp!参数值, 0)) = True
            Case "消息转发" '组成形式 "目的地1|目的地2---" 消息是UDP消息,将来可能为工作站开发工具以实现自动传输,无需查看时才提取
                Call FillBlRoute("消息转发", Nvl(rsTemp!参数值), "", "")
            Case "自动路由"
                str自动路由目的地 = Nvl(rsTemp!参数值)
            Case "自动路由压缩方式"
                str自动路由压缩方式 = Nvl(rsTemp!参数值)
            Case "自动路由目录结构"
                str自动路由目录结构 = Nvl(rsTemp!参数值)
            Case "存储过滤方式"
                Call SeekIndexWithNo(cboMatchOther, Nvl(rsTemp!参数值, 0), True)
        End Select
        rsTemp.MoveNext
    Loop
    
    '填写自动路由参数
    If str自动路由目的地 <> "" Then
        Call FillBlRoute("自动路由", str自动路由目的地, str自动路由压缩方式, str自动路由目录结构)
    End If
    
    '根据图像类型拆分序列”这个参数，只是针对某些CT使用
    gstrSQL = "select 影像类别 from 影像DICOM服务对 A,影像设备目录 B WHERE A.服务ID=[1] AND A.设备号=B.设备号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取参数", mlngSrvID)
    If Not rsTemp.EOF Then
        If UCase(rsTemp!影像类别) <> "CT" Then
            chkImageType.value = 0
            chkImageType.Visible = False
        End If
    End If
End Sub

Private Sub FillBlRoute(ByVal strType As String, ByVal strData As String, ByVal strPara1 As String, ByVal strPara2 As String)
'------------------------------------------------
'功能：填写自动路由或者消息转发的信息，消息转发方式现在还没有使用
'参数： strType--类型---“自动路由”或者“消息转发”
'       strData--数据内容，对于自动路由是目的地
'       strPara1--参数内容1，对于自动路由是压缩方式,
'       strPara2--参数内容2，对于自动路由是目录结构
'返回：无，直接填写控件
'------------------------------------------------
    Dim i As Integer, j As Integer
    Dim blnWritePara As Boolean
    '组成形式 "路由方式,目的地|---" 路由方式分为 路由/消息 ,路由即经DICOM传输,消息是UDP消息,将来可能为工作站开发工具以实现自动传输,无需查看时才提取
    
    If strData = "" Then Exit Sub
    
    '检查数据
    If strType = "自动路由" Then
        If UBound(Split(strData, "|")) = UBound(Split(strPara1, "|")) And UBound(Split(strData, "|")) = UBound(Split(strPara2, "|")) Then
            blnWritePara = True
        Else
            blnWritePara = False
        End If
    End If
    
    With vfgList
        For i = 0 To UBound(Split(strData, "|"))
            .TextMatrix(.Rows - 1, 0) = strType
            If strType = "自动路由" Then '自动路由保存的设备号,通过循环将Cbo中名称取出
                For j = 0 To UBound(Split(.ColComboList(1), "|"))
                    If InStr(Split(.ColComboList(1), "|")(j), Split(strData, "|")(i)) > 0 Then
                        .TextMatrix(.Rows - 1, 1) = Split(.ColComboList(1), "|")(j)
                        If blnWritePara = True Then '有参数，则按照参数来填写
                            .TextMatrix(.Rows - 1, 2) = IIf(Split(strPara1, "|")(i) = 1, AR不压缩, AR现有压缩)
                            .TextMatrix(.Rows - 1, 3) = IIf(Split(strPara2, "|")(i) = 1, AR序列级别, AR检查级别)
                        Else    '没有参数，则填写默认值
                            .TextMatrix(.Rows - 1, 2) = AR现有压缩
                            .TextMatrix(.Rows - 1, 3) = AR检查级别
                        End If
                    End If
                Next
            Else
                .TextMatrix(.Rows - 2, 1) = Split(strData, "|")(i)
            End If
            .Rows = .Rows + 1
        Next
        .TextMatrix(.Rows - 1, 0) = "自动路由"
    End With
End Sub

Public Sub SavePara()
    Dim strData As String
    Dim i As Integer, strData1 As String
    Dim arrSQL() As Variant
    Dim blnInTrans As Boolean       '是否在事务处理之中
    
    On Error GoTo ErrHandle
    If cboDevice.Text = "" Then
        MsgBoxD Me, "请选择存储设备！", vbInformation, gstrSysName: cboDevice.SetFocus: Exit Sub
    End If
    
    If cboEncode.Text = "" Then
        MsgBoxD Me, "请选择压缩方式！", vbInformation, gstrSysName: cboEncode.SetFocus: Exit Sub
    End If
    
    arrSQL = Array()
    
    If cboDevice.ListIndex <> -1 Then
        gstrSQL = "Zl_影像DICOM服务参数_SAVE(" & mlngSrvID & ",'存储设备','" & NeedNo(cboDevice.Text) & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
    End If
    
    If cboEncode.ListIndex <> -1 Then
        gstrSQL = "Zl_影像DICOM服务参数_SAVE(" & mlngSrvID & ",'压缩方式','" & NeedName(cboEncode.Text) & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
    End If
    
    gstrSQL = "Zl_影像DICOM服务参数_SAVE(" & mlngSrvID & ",'按检查UID匹配','" & chkMatchStudyUID.value & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL
    
    gstrSQL = "Zl_影像DICOM服务参数_SAVE(" & mlngSrvID & ",'按类型拆分序列','" & chkImageType.value & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL
    
    gstrSQL = "Zl_影像DICOM服务参数_SAVE(" & mlngSrvID & ",'存储过滤方式','" & NeedNo(cboMatchOther.Text) & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL
    
    strData = 0
    For i = 0 To optImgMatch.UBound
        If optImgMatch(i).value = True Then
            strData = i
            Exit For
        End If
    Next
    If strData = "" Then strData = "0"
    gstrSQL = "Zl_影像DICOM服务参数_SAVE(" & mlngSrvID & ",'匹配图像项目','" & strData & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL

    strData = 0
    For i = 0 To optMatch.UBound
        If optMatch(i).value = True Then
            strData = i
            Exit For
        End If
    Next
    If strData = "" Then strData = "0"
    gstrSQL = "Zl_影像DICOM服务参数_SAVE(" & mlngSrvID & ",'匹配数据库项目','" & strData & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL
    
    gstrSQL = "Zl_影像DICOM服务参数_Delete(" & mlngSrvID & ",'自动路由')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL

    gstrSQL = "Zl_影像DICOM服务参数_Delete(" & mlngSrvID & ",'消息转发')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL
    
    With vfgList
        strData = ""
        str自动路由压缩方式 = ""
        str自动路由目录结构 = ""
        For i = 1 To vfgList.Rows - 1
            If Trim(vfgList.TextMatrix(i, 1)) <> "" And vfgList.RowHidden(i) = False Then
                If vfgList.TextMatrix(i, 0) = "自动路由" Then
                    If InStr(strData, NeedNo(vfgList.TextMatrix(i, 1))) = 0 Then '重复的不增加
                        strData = strData & "|" & NeedNo(vfgList.TextMatrix(i, 1))
                        str自动路由压缩方式 = str自动路由压缩方式 & "|" & IIf(vfgList.TextMatrix(i, 2) = AR不压缩, 1, 0)
                        str自动路由目录结构 = str自动路由目录结构 & "|" & IIf(vfgList.TextMatrix(i, 3) = AR序列级别, 1, 0)
                    End If
                Else
                    If InStr(strData1, vfgList.TextMatrix(i, 1)) = 0 Then '重复的不增加
                        strData1 = strData1 & "|" & vfgList.TextMatrix(i, 1)
                    End If
                End If
            End If
        Next
    End With
    strData = Mid(strData, 2)
    gstrSQL = "Zl_影像DICOM服务参数_SAVE(" & mlngSrvID & ",'自动路由','" & strData & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL
    
    strData1 = Mid(strData1, 2)
    gstrSQL = "Zl_影像DICOM服务参数_SAVE(" & mlngSrvID & ",'消息转发','" & strData1 & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL
    
    str自动路由压缩方式 = Mid(str自动路由压缩方式, 2)
    gstrSQL = "Zl_影像DICOM服务参数_SAVE(" & mlngSrvID & ",'自动路由压缩方式','" & str自动路由压缩方式 & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL
    
    str自动路由目录结构 = Mid(str自动路由目录结构, 2)
    gstrSQL = "Zl_影像DICOM服务参数_SAVE(" & mlngSrvID & ",'自动路由目录结构','" & str自动路由目录结构 & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL
    
    gcnOracle.BeginTrans        '开始保存参数
    blnInTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "保存图像接收参数")
    Next i
    gcnOracle.CommitTrans
    blnInTrans = False
    
    RefreshPara
   Exit Sub
ErrHandle:
    If blnInTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Call InitvfgList
End Sub
Private Sub InitvfgList()
Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHandle
    With vfgList
        .Clear
        .FixedRows = 1
        .Rows = 2
        .Cols = 4
        .ColWidth(0) = 800
        .ColWidth(1) = 800
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        .TextMatrix(0, 0) = "转发类型"
        .TextMatrix(0, 1) = "转发目的地"
        .TextMatrix(0, 2) = "压缩方式"
        .TextMatrix(0, 3) = "目录结构"
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .TextMatrix(1, 0) = "自动路由"
        .TextMatrix(1, 2) = AR现有压缩
        .TextMatrix(1, 3) = AR检查级别
    End With
    gstrSQL = "select 设备号,设备名 from 影像设备目录 where 类型=1 and NVL(状态,0)=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取存储设备")
    cboDevice.Clear
    Dim strList As String
    Do Until rsTemp.EOF
        cboDevice.AddItem rsTemp!设备号 & "-" & rsTemp!设备名
        strList = strList & "|" & rsTemp!设备号 & "-" & rsTemp!设备名
        rsTemp.MoveNext
    Loop
    vfgList.ColComboList(1) = strList
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vfgList_Click()
    With vfgList
        If .Col = 0 Or .Col = 2 Or .Col = 3 Then
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub vfgList_DblClick()
    With vfgList
        If .Col = 0 Then
            If .TextMatrix(.Row, .Col) = "自动路由" Then
                .TextMatrix(.Row, .Col) = "消息转发"
            Else
                .TextMatrix(.Row, .Col) = "自动路由"
            End If
            .TextMatrix(.Row, 1) = ""
            If .TextMatrix(.Row, 0) = "自动路由" Then
                .TextMatrix(.Row, 2) = AR现有压缩
                .TextMatrix(.Row, 3) = AR检查级别
            Else
                .TextMatrix(.Row, 2) = ""
                .TextMatrix(.Row, 3) = ""
            End If
        End If
        
        If .Col = 2 Then
            If .TextMatrix(.Row, 2) = AR现有压缩 Then
                .TextMatrix(.Row, 2) = AR不压缩
            Else
                .TextMatrix(.Row, 2) = AR现有压缩
            End If
        End If
        
        If .Col = 3 Then
            If .TextMatrix(.Row, 3) = AR检查级别 Then
                .TextMatrix(.Row, 3) = AR序列级别
            Else
                .TextMatrix(.Row, 3) = AR检查级别
            End If
        End If
    End With
End Sub

Private Sub vfgList_KeyDown(KeyCode As Integer, Shift As Integer)
    '回车，新增一行
    If KeyCode = vbKeyReturn Then
        vfgList.Rows = vfgList.Rows + 1
        vfgList.TextMatrix(vfgList.Rows - 1, 0) = "自动路由"
        vfgList.TextMatrix(vfgList.Rows - 1, 2) = AR现有压缩
        vfgList.TextMatrix(vfgList.Rows - 1, 3) = AR检查级别
    End If
    'delete删除最后一行
    If KeyCode = vbKeyDelete And vfgList.Row >= 1 Then
        If MsgBoxD(Me, "是否删除本行？", vbYesNo) = vbYes Then
            vfgList.RowHidden(vfgList.Row) = True
        End If
    End If
End Sub

Private Sub vfgList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = 0 Then
        KeyAscii = 0
    ElseIf Col = 1 And vfgList.TextMatrix(vfgList.Row, 0) = "自动路由" Then
        KeyAscii = 0
    End If
End Sub
