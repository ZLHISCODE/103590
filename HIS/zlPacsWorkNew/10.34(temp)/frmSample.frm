VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSample 
   Caption         =   "Form1"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10050
   Icon            =   "frmSample.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   10050
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame framSample 
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.Frame fram核收 
         Caption         =   "1、标本核收"
         Height          =   3495
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   9615
         Begin VB.TextBox txtPathologyNo 
            Height          =   350
            Left            =   5040
            TabIndex        =   28
            Top             =   2400
            Width           =   1815
         End
         Begin VB.ComboBox cboCheckType 
            Height          =   300
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   2425
            Width           =   1575
         End
         Begin VB.TextBox txt拒收原因 
            Height          =   350
            Left            =   2520
            TabIndex        =   21
            Top             =   2880
            Width           =   6855
         End
         Begin VB.ComboBox cbo核收技师 
            Height          =   300
            Left            =   8160
            TabIndex        =   20
            Text            =   "Combo1"
            Top             =   2422
            Width           =   1260
         End
         Begin VB.CommandButton cmdRefuse 
            Caption         =   "拒收"
            Height          =   350
            Left            =   120
            TabIndex        =   19
            Top             =   2880
            Width           =   1100
         End
         Begin VB.CommandButton cmdCheckIn 
            Caption         =   "核收"
            Height          =   350
            Left            =   120
            TabIndex        =   18
            Top             =   2400
            Width           =   1100
         End
         Begin VSFlex8Ctl.VSFlexGrid vfgList 
            Height          =   1935
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   9375
            _cx             =   16536
            _cy             =   3413
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
            Cols            =   16
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
         Begin VB.Label Label11 
            Caption         =   "病理号"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   4320
            TabIndex        =   27
            Top             =   2445
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "检查类型"
            Height          =   255
            Left            =   1560
            TabIndex        =   25
            Top             =   2448
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "拒收原因"
            Height          =   255
            Left            =   1560
            TabIndex        =   23
            Top             =   2928
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "核收技师"
            Height          =   255
            Left            =   7200
            TabIndex        =   22
            Top             =   2445
            Width           =   855
         End
      End
      Begin VB.Frame fram取材 
         Caption         =   "2、巨检"
         Height          =   3855
         Left            =   120
         TabIndex        =   1
         Top             =   4200
         Width           =   9615
         Begin VB.CommandButton cmdSave 
            Caption         =   "保存"
            Height          =   350
            Left            =   120
            TabIndex        =   24
            Top             =   3360
            Width           =   1100
         End
         Begin VB.TextBox txt巨检诊断 
            Height          =   1695
            Left            =   600
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   240
            Width           =   8895
         End
         Begin VB.TextBox txt附言 
            Height          =   350
            Left            =   600
            TabIndex        =   7
            Top             =   2040
            Width           =   4575
         End
         Begin VB.TextBox txt备注 
            Height          =   350
            Left            =   600
            TabIndex        =   6
            Top             =   2475
            Width           =   8895
         End
         Begin VB.TextBox txt剩余标本位置 
            Height          =   350
            Left            =   6480
            TabIndex        =   5
            Top             =   2040
            Width           =   3015
         End
         Begin VB.ComboBox cbo取材技师 
            Height          =   300
            Left            =   1080
            TabIndex        =   4
            Text            =   "Combo2"
            Top             =   2932
            Width           =   1500
         End
         Begin VB.ComboBox cbo巨检医师 
            Height          =   300
            Left            =   4560
            TabIndex        =   3
            Text            =   "Combo3"
            Top             =   2910
            Width           =   1500
         End
         Begin VB.ComboBox cbo切片技师 
            Height          =   300
            Left            =   7800
            TabIndex        =   2
            Text            =   "Combo4"
            Top             =   2910
            Width           =   1500
         End
         Begin VB.Label Label4 
            Caption         =   "巨检"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "附言"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   2085
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "备注"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   2520
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "剩余标本位置"
            Height          =   255
            Left            =   5280
            TabIndex        =   12
            Top             =   2085
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "取材技师"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   2955
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "巨检医师"
            Height          =   255
            Left            =   3480
            TabIndex        =   10
            Top             =   2955
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "切片技师"
            Height          =   255
            Left            =   6840
            TabIndex        =   9
            Top             =   2955
            Width           =   735
         End
      End
      Begin VB.Label lblInfo 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   9495
      End
   End
End
Attribute VB_Name = "frmSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng科室ID As Long
Private mlng医嘱ID As Long
Private mlng发送号 As Long
Private blnInit As Boolean
Private blnChangedDept As Boolean
Private mblnMoved As Boolean
Private mstrPrivs As String

Private Enum ColList
        ColID = 0
        Col编号1
        Col标本部位和名称1
        Col数量1
        Col编号2
        Col标本部位和名称2
        Col数量2
        Col编号3
        Col标本部位和名称3
        Col数量3
End Enum

'状态改变事件lngState = 0 未操作；lngState=1 标本核收；lngState = 2 标本拒收
Public Event StateChanged(lngState As Long, str病理号 As String, str病理检查类别 As String)

'公共调用函数
Public Sub zlRefresh(lng科室ID As Long, lng医嘱ID As Long, lng发送号 As Long, strPrivs As String, ByVal blnReadOnly As Boolean, ByVal blnMoved As Boolean)
    If mlng科室ID <> lng科室ID Or lng科室ID = 0 Then
        blnChangedDept = True
        mlng科室ID = lng科室ID
    Else
        blnChangedDept = False
    End If
    
    mlng医嘱ID = lng医嘱ID
    mlng发送号 = lng发送号
    mblnMoved = blnMoved
    mstrPrivs = strPrivs
    
    Call InitBillSamples    '初始化标本部位名称
    If blnInit = True Then  '窗体装载，初始化基本数据
        Call FillCheckType      '初始化病理检查类型
    End If
    If blnChangedDept = True Then Call InitDoctors
    '填充界面元素
    Call FillInterface
    If lng医嘱ID = 0 Or lng发送号 = 0 Or blnReadOnly = True Or mblnMoved = True Then
        framSample.Enabled = False
    Else
        framSample.Enabled = True
    End If
End Sub

Private Sub FillCheckType()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    blnInit = False
    cboCheckType.Clear
    strSQL = "Select 名称 From 影像病理类别"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取影像病理类别")
    
    If rsTemp.EOF = True Then
        MsgBoxD Me, "查找不到病理检查类型的信息，请先在字典表“影像病理类别”中设置。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    While Not rsTemp.EOF
        cboCheckType.AddItem (Nvl(rsTemp!名称))
        rsTemp.MoveNext
    Wend
    If cboCheckType.ListCount > 0 And cboCheckType.ListIndex = -1 Then cboCheckType.ListIndex = 0
End Sub

Private Sub FillInterface()
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    '清空当前内容
    lblInfo.Caption = ""
    txt拒收原因.Text = ""
    txt巨检诊断.Text = ""
    txt附言.Text = ""
    txt剩余标本位置.Text = ""
    txt备注.Text = ""
    txtPathologyNo.Text = ""
    
    On Error GoTo errHandle
    
    '填写巨检等内容
    strSQL = "Select 医嘱ID,发送号,巨检所见,附言,剩余标本位置,备注,巨检医师,取材技师,切片技师,核收技师," & _
             "拒收原因,核收情况,病理号,病理检查类别 From 影像标本核收取材 Where 医嘱ID=[1]"
    If mblnMoved = True Then
        strSQL = Replace(strSQL, "影像标本核收取材", "H影像标本核收取材")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取病理标本核收取材", mlng医嘱ID)
    If Not rsTemp.EOF Then
        If Nvl(rsTemp!核收情况, 0) = 2 Then '拒收
            lblInfo.Caption = "标本被拒收"
            lblInfo.ForeColor = vbRed
            cbo核收技师.Text = Nvl(rsTemp!核收技师)
            txt拒收原因.Text = Nvl(rsTemp!拒收原因)
        Else
            lblInfo.Caption = "标本已核收"
            lblInfo.ForeColor = vbBlue
            txt巨检诊断.Text = Nvl(rsTemp!巨检所见)
            txt附言.Text = Nvl(rsTemp!附言)
            txt剩余标本位置.Text = Nvl(rsTemp!剩余标本位置)
            txt备注.Text = Nvl(rsTemp!备注)
            cbo取材技师.Text = Nvl(rsTemp!取材技师)
            cbo巨检医师.Text = Nvl(rsTemp!巨检医师)
            cbo切片技师.Text = Nvl(rsTemp!切片技师)
            Call SetCheckType(Nvl(rsTemp!病理检查类别))
            txtPathologyNo.Text = Nvl(rsTemp!病理号)
        End If
    End If
    
    If txtPathologyNo.Text = "" Then    '刷新提取最大病理号
        Call cboCheckType_Click
    End If
    '根据数据库填写表单内容
    Call FillBILL
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetCheckType(str病理检查类别 As String)
    Dim i As Integer
    
    For i = 0 To cboCheckType.ListCount - 1
        If cboCheckType.List(i) = str病理检查类别 Then Exit For
    Next i
    If i < cboCheckType.ListCount Then
        cboCheckType.ListIndex = i
    Else
        cboCheckType.ListIndex = -1
    End If
End Sub

Private Sub InitBillSamples()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    With vfgList
        .Clear
        .FixedRows = 1
        .Rows = 6
        .Cols = 10
        .ColWidth(ColID) = 0    'ID
        .ColWidth(Col编号1) = 500
        .ColWidth(Col标本部位和名称1) = 1500
        .ColWidth(Col数量1) = 1000
        .ColWidth(Col编号2) = 500
        .ColWidth(Col标本部位和名称2) = 1500
        .ColWidth(Col数量2) = 1000
        .ColWidth(Col编号3) = 500
        .ColWidth(Col标本部位和名称3) = 1500
        .ColWidth(Col数量3) = 1000
        
        .TextMatrix(0, ColID) = "ID"
        .TextMatrix(0, Col编号1) = "编号"
        .TextMatrix(0, Col标本部位和名称1) = "标本部位和名称"
        .TextMatrix(0, Col数量1) = "数量"
        .TextMatrix(0, Col编号2) = "编号"
        .TextMatrix(0, Col标本部位和名称2) = "标本部位和名称"
        .TextMatrix(0, Col数量2) = "数量"
        .TextMatrix(0, Col编号3) = "编号"
        .TextMatrix(0, Col标本部位和名称3) = "标本部位和名称"
        .TextMatrix(0, Col数量3) = "数量"
        
        .ColAlignment(ColID) = flexAlignCenterCenter
        .ColAlignment(Col编号1) = flexAlignCenterCenter
        .ColAlignment(Col标本部位和名称1) = flexAlignCenterCenter
        .ColAlignment(Col数量1) = flexAlignCenterCenter
        .ColAlignment(Col编号2) = flexAlignCenterCenter
        .ColAlignment(Col标本部位和名称2) = flexAlignCenterCenter
        .ColAlignment(Col数量2) = flexAlignCenterCenter
        .ColAlignment(Col编号3) = flexAlignCenterCenter
        .ColAlignment(Col标本部位和名称3) = flexAlignCenterCenter
        .ColAlignment(Col数量3) = flexAlignCenterCenter
        
        '输入标本部位和名称下拉列表的内容
        strSQL = "Select 名称 From 影像病理标本部位 Order By 名称"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取病理标本部位")
        strSQL = ""
        While Not rsTemp.EOF
            strSQL = strSQL & "|" & zlCommFun.SpellCode(Nvl(rsTemp!名称)) & "-" & Nvl(rsTemp!名称)
            
            rsTemp.MoveNext
        Wend
        strSQL = Mid(strSQL, 2)
        If strSQL = "" Then strSQL = " "
        .ColComboList(Col标本部位和名称1) = "|" & strSQL
        .ColComboList(Col标本部位和名称2) = "|" & strSQL
        .ColComboList(Col标本部位和名称3) = "|" & strSQL
        '填入固定的数字编号
        .TextMatrix(1, Col编号1) = "1"
        .TextMatrix(2, Col编号1) = "2"
        .TextMatrix(3, Col编号1) = "3"
        .TextMatrix(4, Col编号1) = "4"
        .TextMatrix(5, Col编号1) = "5"
        .TextMatrix(1, Col编号2) = "6"
        .TextMatrix(2, Col编号2) = "7"
        .TextMatrix(3, Col编号2) = "8"
        .TextMatrix(4, Col编号2) = "9"
        .TextMatrix(5, Col编号2) = "10"
        .TextMatrix(1, Col编号3) = "11"
        .TextMatrix(2, Col编号3) = "12"
        .TextMatrix(3, Col编号3) = "13"
        .TextMatrix(4, Col编号3) = "14"
        .TextMatrix(5, Col编号3) = "15"
    End With
End Sub

Private Sub FillBILL()
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim lng编号 As Long
    Dim i As Integer
    Dim int蜡块总数 As Integer
    
    If mlng医嘱ID = 0 And mlng发送号 = 0 Then Exit Sub
    On Error GoTo errHandle
    
    strSQL = "Select a.编号,a.医嘱ID,a.发送号,a.标本部位,a.块数 From 影像病理标本 a " & _
             "Where a.医嘱ID = [1] order by a.编号"
    
    If mblnMoved = True Then
        strSQL = Replace(strSQL, "影像病理标本", "H影像病理标本")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取病理标本", mlng医嘱ID)
    
    '清除原有内容
    For i = 1 To 5
        vfgList.TextMatrix(i, Col标本部位和名称1) = ""
        vfgList.TextMatrix(i, Col标本部位和名称2) = ""
        vfgList.TextMatrix(i, Col标本部位和名称3) = ""
        vfgList.TextMatrix(i, Col数量1) = ""
        vfgList.TextMatrix(i, Col数量2) = ""
        vfgList.TextMatrix(i, Col数量3) = ""
    Next i
    
    int蜡块总数 = 0
    While Not rsTemp.EOF
        lng编号 = rsTemp!编号
        With vfgList
            If lng编号 <= 5 Then
                .TextMatrix(lng编号, Col标本部位和名称1) = rsTemp!标本部位
                .TextMatrix(lng编号, Col数量1) = rsTemp!块数
            ElseIf lng编号 <= 10 Then
                .TextMatrix(lng编号 - 5, Col标本部位和名称2) = rsTemp!标本部位
                .TextMatrix(lng编号 - 5, Col数量2) = rsTemp!块数
            Else
                .TextMatrix(lng编号 - 10, Col标本部位和名称3) = rsTemp!标本部位
                .TextMatrix(lng编号 - 10, Col数量3) = rsTemp!块数
            End If
            int蜡块总数 = int蜡块总数 + Nvl(rsTemp!块数, 0)
        End With
        rsTemp.MoveNext
    Wend
    If lblInfo.Caption <> "" And int蜡块总数 <> 0 Then
        lblInfo.Caption = lblInfo.Caption & " 总数 " & int蜡块总数
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitDoctors()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    strSQL = "Select /*+RULE*/" & vbNewLine & _
                "Distinct b.id,b.姓名, Upper(b.简码) As 简码" & vbNewLine & _
                " From 部门人员 a, 人员表 b, 人员性质说明 c" & vbNewLine & _
                " Where a.部门id = [1] And a.人员id = b.Id And b.Id = c.人员id And c.人员性质 = '医生' And" & vbNewLine & _
                "      (b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.撤档时间 Is Null)" & vbNewLine & _
                " Order By 简码 Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取医生", mlng科室ID)
    cbo核收技师.Clear
    cbo巨检医师.Clear
    cbo切片技师.Clear
    cbo取材技师.Clear
    While Not rsTemp.EOF
        cbo核收技师.AddItem rsTemp!简码 & "-" & rsTemp!姓名
        If rsTemp!ID = UserInfo.ID Then cbo核收技师.ListIndex = cbo核收技师.NewIndex
        cbo巨检医师.AddItem rsTemp!简码 & "-" & rsTemp!姓名
        If rsTemp!ID = UserInfo.ID Then cbo巨检医师.ListIndex = cbo巨检医师.NewIndex
        cbo切片技师.AddItem rsTemp!简码 & "-" & rsTemp!姓名
        If rsTemp!ID = UserInfo.ID Then cbo切片技师.ListIndex = cbo切片技师.NewIndex
        cbo取材技师.AddItem rsTemp!简码 & "-" & rsTemp!姓名
        If rsTemp!ID = UserInfo.ID Then cbo取材技师.ListIndex = cbo取材技师.NewIndex
        rsTemp.MoveNext
    Wend
    If cbo核收技师.ListCount > 0 And cbo核收技师.ListIndex = -1 Then cbo核收技师.ListIndex = 0
    If cbo巨检医师.ListCount > 0 And cbo巨检医师.ListIndex = -1 Then cbo巨检医师.ListIndex = 0
    If cbo切片技师.ListCount > 0 And cbo切片技师.ListIndex = -1 Then cbo切片技师.ListIndex = 0
    If cbo取材技师.ListCount > 0 And cbo取材技师.ListIndex = -1 Then cbo取材技师.ListIndex = 0
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cboCheckType_Click()
    '提取最大病理号
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim lngBigNumber As Long
    
    strSQL = "Select 名称,最大号码,前导标记 From 影像病理类别 where 名称 = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取最大病理号", CStr(cboCheckType.Text))
    
    If Not rsTemp.EOF Then
        '判断是否存在前导标记和最大号码
        If IsNull(rsTemp!前导标记) Then
            MsgBoxD Me, "请先在字典表“影像病理类别”中设置病理号的前导标记。", vbInformation, gstrSysName
            Exit Sub
        End If
        lngBigNumber = Nvl(rsTemp!最大号码, 0)
        txtPathologyNo.Text = Nvl(rsTemp!前导标记) & lngBigNumber
    Else
        txtPathologyNo.Text = ""
    End If
End Sub

Private Sub cmdCheckIn_Click()
    Call CheckInSamples
End Sub

Private Sub CheckInSamples()
    '核收标本
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim str病理标本组 As String
    Dim i As Integer
    Dim j As Integer
    Dim iSampleCount As Integer
    Dim str病理号 As String
    Dim str病理检查类别 As String
    
    str病理号 = txtPathologyNo.Text
    str病理检查类别 = cboCheckType.Text
    If str病理号 = "" Or str病理检查类别 = "" Then
        MsgBoxD Me, "病理号或者病理检查类别输入不正确，请检查。 ", vbInformation, gstrSysName
        Exit Sub
    Else
        '检查病理号，是否有一个前导字符
        If Not IsNumeric(Mid(str病理号, 2)) Or IsNumeric(Left(str病理号, 1)) Then
            MsgBoxD Me, "病理号不符合“类别前导字符+数字编号”的规则，请检查。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    iSampleCount = 0
    For i = 1 To 5
        If vfgList.TextMatrix(i, Col标本部位和名称1) <> "" And vfgList.TextMatrix(i, Col数量1) <> "" _
            And Val(vfgList.TextMatrix(i, Col数量1)) <> 0 Then
            str病理标本组 = str病理标本组 & "<" & vfgList.TextMatrix(i, Col编号1) & "-" & _
                            vfgList.TextMatrix(i, Col标本部位和名称1) & "-" & _
                            Val(vfgList.TextMatrix(i, Col数量1)) & ">"
            iSampleCount = iSampleCount + Val(vfgList.TextMatrix(i, Col数量1))
        Else
            Exit For
        End If
    Next i
    
    For i = 1 To 5
        If vfgList.TextMatrix(i, Col标本部位和名称2) <> "" And vfgList.TextMatrix(i, Col数量2) <> "" _
            And Val(vfgList.TextMatrix(i, Col数量2)) <> 0 Then
            str病理标本组 = str病理标本组 & "<" & vfgList.TextMatrix(i, Col编号2) & "-" & _
                            vfgList.TextMatrix(i, Col标本部位和名称2) & "-" & _
                            Val(vfgList.TextMatrix(i, Col数量2)) & ">"
            iSampleCount = iSampleCount + Val(vfgList.TextMatrix(i, Col数量2))
        Else
            Exit For
        End If
    Next i
    
    For i = 1 To 5
        If vfgList.TextMatrix(i, Col标本部位和名称3) <> "" And vfgList.TextMatrix(i, Col数量3) <> "" _
            And Val(vfgList.TextMatrix(i, Col数量3)) <> 0 Then
            str病理标本组 = str病理标本组 & "<" & vfgList.TextMatrix(i, Col编号3) & "-" & _
                            vfgList.TextMatrix(i, Col标本部位和名称3) & "-" & _
                            Val(vfgList.TextMatrix(i, Col数量3)) & ">"
            iSampleCount = iSampleCount + Val(vfgList.TextMatrix(i, Col数量3))
        Else
            Exit For
        End If
    Next i

    On Error GoTo errHandle
    
    strSQL = "ZL_影像标本核收(" & mlng医嘱ID & "," & mlng发送号 & ",'" & NeedName(cbo核收技师.Text) & "',sysdate,'" _
                & str病理检查类别 & "','" & str病理号 & "'," & Mid(str病理号, 2) & ",'" & str病理标本组 & "')"
    zlDatabase.ExecuteProcedure strSQL, "影像病理标本核收"
    lblInfo.Caption = "标本已核收" & " 总数 " & iSampleCount
    lblInfo.ForeColor = vbBlue
    
    '检查病理号是否更改
    strSQL = "Select 病理号 From 影像标本核收取材 where 医嘱ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取影像病理号", mlng医嘱ID)
    If rsTemp!病理号 <> str病理号 Then
        MsgBoxD Me, "原病理号： " & str病理号 & " 已经被使用，自动将病理号修改为： " & rsTemp!病理号, vbInformation, gstrSysName
        txtPathologyNo.Text = rsTemp!病理号
    End If
    '触发事件
    RaiseEvent StateChanged(1, txtPathologyNo.Text, str病理检查类别)
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdRefuse_Click()
    '拒收标本
    Dim strSQL As String
    
    strSQL = "ZL_影像标本拒收(" & mlng医嘱ID & "," & mlng发送号 & ",'" & NeedName(cbo核收技师.Text) & "',sysdate,'" & txt拒收原因.Text & "')"
    zlDatabase.ExecuteProcedure strSQL, "影像病理标本拒收"
    lblInfo.Caption = "标本被拒收"
    lblInfo.ForeColor = vbRed
    RaiseEvent StateChanged(2, "", "")
End Sub

Private Sub CmdSave_Click()
    '保存巨检和取材
    Dim strSQL As String
    
    '先核收标本
    Call CheckInSamples
    
    strSQL = "ZL_影像病理巨检取材(" & mlng医嘱ID & "," & mlng发送号 & ",'" & txt巨检诊断.Text & "','" & _
            txt附言.Text & "','" & txt剩余标本位置.Text & "','" & txt备注.Text & "','" & NeedName(cbo巨检医师.Text) & _
            "',sysdate,'" & NeedName(cbo取材技师.Text) & "','" & NeedName(cbo切片技师.Text) & "')"
    zlDatabase.ExecuteProcedure strSQL, "影像病理巨检取材"
    
End Sub

Private Sub Form_Load()
    blnInit = True
    vfgList.SelectionMode = flexSelectionFree
    vfgList.Editable = flexEDKbdMouse
End Sub

Private Sub vfgList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '如果是标本部位，删除掉标本部位名称前面的拼音首字母
    Dim strTemp As String
    Dim strSQL As String
    
    On Error Resume Next
    
    If Col = Col标本部位和名称1 Or Col = Col标本部位和名称2 Or Col = Col标本部位和名称3 Then
        strTemp = vfgList.TextMatrix(Row, Col)
        If InStr(strTemp, "-") <> 0 Then
            strTemp = Mid(strTemp, InStr(strTemp, "-") + 1)
            vfgList.TextMatrix(Row, Col) = strTemp
        Else
            '判断用户是否有添加病理标本的权限，如果有，则检查标本是否存在，不存在则添加
            If InStr(mstrPrivs, "标本管理") > 0 And Trim(strTemp) <> "" Then
                strSQL = "ZL_影像病理标本部位_Insert('" & strTemp & "','" & zlCommFun.SpellCode(strTemp) & "' )"
                zlDatabase.ExecuteProcedure strSQL, "更新病理标本"
            End If
        End If
    End If
End Sub

Private Sub vfgList_DblClick()
    With vfgList
        If .Col = Col编号1 Or .Col = Col编号2 Or .Col = Col编号3 Then
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub vfgList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = Col标本部位和名称1 Or Col = Col标本部位和名称2 Or Col = Col标本部位和名称3 Then
'        If KeyAscii <> vbKeyReturn Then KeyAscii = 0
    ElseIf Col = Col编号1 Or Col = Col编号2 Or Col = Col编号3 Then
        KeyAscii = 0
    ElseIf Col = Col数量1 Or Col = Col数量2 Or Col = Col数量3 Then
        If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> vbKeyReturn Then
            KeyAscii = 0
        End If
    End If
    
    If KeyAscii = vbKeyReturn Then
        If Col = 2 Or Col = 5 Or Col = 8 Then
            vfgList.Select Row, Col + 1
            vfgList.EditCell
        ElseIf Col = 3 Or Col = 6 Or Col = 9 Then
            If Row < 5 Then
                vfgList.Select Row + 1, Col - 1
                vfgList.EditCell
            ElseIf Col <> 9 Then
                vfgList.Select 1, Col + 2
                vfgList.EditCell
            End If
        End If
    End If
End Sub
