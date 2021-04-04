VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frm存储库房 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "存储库房设置"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   Icon            =   "frm存储库房.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdStuff 
      Caption         =   "…"
      Height          =   285
      Left            =   8520
      TabIndex        =   23
      TabStop         =   0   'False
      Tag             =   "分类"
      ToolTipText     =   "按*打开选择器"
      Top             =   623
      Width           =   285
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   5280
      MaxLength       =   10
      TabIndex        =   20
      Tag             =   "规格编码"
      Top             =   990
      Width           =   1440
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "全选(&A)"
      Height          =   350
      Left            =   1200
      Picture         =   "frm存储库房.frx":000C
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5970
      Width           =   957
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "应用于本列(&L)"
      Height          =   350
      Left            =   3400
      Picture         =   "frm存储库房.frx":0156
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "当在服务科室选择了科室后，将服务科室列中选择的值应用前面勾选了的行中！"
      Top             =   5970
      Width           =   1395
   End
   Begin VB.CommandButton cmd参照 
      Caption         =   "参照最后设置(&Z)"
      Height          =   350
      Left            =   5760
      Picture         =   "frm存储库房.frx":02A0
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "存储库房和服务科室的设置参照最后一次进行设置!"
      Top             =   5970
      Width           =   1572
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   4635
      Left            =   -6135
      TabIndex        =   18
      Top             =   6000
      Visible         =   0   'False
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   8176
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Frame frame 
      Caption         =   "应用于同院区(&B)"
      Height          =   1050
      Left            =   135
      TabIndex        =   4
      Top             =   4800
      Width           =   9180
      Begin VB.OptionButton opt应用于 
         Caption         =   "应用于本品种所有规格(&2)"
         Height          =   255
         Index           =   4
         Left            =   2985
         TabIndex        =   6
         Top             =   285
         Width           =   2730
      End
      Begin VB.OptionButton opt应用于 
         Caption         =   "应用于此分类下的所有卫生材料(&4)"
         Height          =   255
         Index           =   2
         Left            =   210
         TabIndex        =   8
         Top             =   630
         Width           =   4320
      End
      Begin VB.OptionButton opt应用于 
         Caption         =   "应用于本级所有卫生材料(&3)"
         Height          =   255
         Index           =   1
         Left            =   5880
         TabIndex        =   7
         Top             =   285
         Width           =   3045
      End
      Begin VB.OptionButton opt应用于 
         Caption         =   "仅应用于本卫生材料(&1)"
         Height          =   255
         Index           =   0
         Left            =   195
         TabIndex        =   5
         Top             =   285
         Value           =   -1  'True
         Width           =   4065
      End
      Begin VB.OptionButton opt应用于 
         Caption         =   "应用于所有“卫生材料”(&5)"
         Height          =   255
         Index           =   3
         Left            =   5880
         TabIndex        =   9
         Top             =   630
         Width           =   3045
      End
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "恢复(&R)"
      Height          =   350
      Left            =   4800
      Picture         =   "frm存储库房.frx":03EA
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5970
      Width           =   957
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "全清(&C)"
      Height          =   350
      Left            =   2250
      Picture         =   "frm存储库房.frx":0534
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5970
      Width           =   957
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   7350
      TabIndex        =   10
      Top             =   5970
      Width           =   957
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   120
      Picture         =   "frm存储库房.frx":067E
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5970
      Width           =   885
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   8325
      TabIndex        =   11
      Top             =   5970
      Width           =   957
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2985
      Top             =   5610
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm存储库房.frx":07C8
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm存储库房.frx":0D62
            Key             =   "ItemStop"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm存储库房.frx":12FC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ZL9BillEdit.BillEdit mshBillEdit 
      Height          =   3180
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   5609
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.TextBox txtStuff 
      Height          =   300
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   2
      Top             =   615
      Width           =   6855
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   240
      Picture         =   "frm存储库房.frx":3006
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblFindComment 
      Caption         =   "输入编码，名称，简码，并按回车进行查找，连续查找按F3。"
      Height          =   375
      Left            =   6840
      TabIndex        =   22
      Top             =   980
      Width           =   2535
   End
   Begin VB.Label lblFind 
      AutoSize        =   -1  'True
      Caption         =   "查找库房"
      Height          =   180
      Left            =   4440
      TabIndex        =   21
      Top             =   1050
      Width           =   720
   End
   Begin VB.Label lbl分类 
      AutoSize        =   -1  'True
      Caption         =   "指定的卫材分类："
      Height          =   180
      Left            =   240
      TabIndex        =   19
      Top             =   1050
      Width           =   1440
   End
   Begin VB.Label lblMedi 
      AutoSize        =   -1  'True
      Caption         =   "指定卫生材料(&M)"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   675
      Width           =   1350
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    请选择卫生材料后，设置该卫生材料的存储库房的服务科室。"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   432
      TabIndex        =   0
      Top             =   240
      Width           =   5244
   End
End
Attribute VB_Name = "frm存储库房"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnFirst As Boolean
Private mbln编辑 As Boolean
Private mlng材料ID As Long                  '诊疗项目ID
Private Const mlngModule = 1711
Private mstrPreStuffName As String          '上次选择的文本数据
Private mlngLastFindRows As Long            '上次找到的行号
Private mstrFind As String
Private mrs科室 As ADODB.Recordset
Private mblnFind As Boolean             '是否是第一次查询
Private mstr记录值 As String            '记录所有的值
Private mblnSave As Boolean             '是否保存成功   true 保存成功或者已经保存 false 未保存成功或未保存
Private mblnKeyMethod As Boolean        '是否是通过编辑回车方式

'Private Sub FindRow(ByVal strInput As String, ByVal lngStartRows As Long)
'    Dim intTargetRow As Integer
'    Dim lngRows As Long
'    Dim blnFind As Boolean
'    Dim strMatch As String
'
'    If strInput = "" Then Exit Sub
'
'    '编码左匹配，名称和简码双向匹配
'    If IsNumeric(strInput) Then
'        intTargetRow = 5
'        strMatch = strInput
'    ElseIf zlStr.IsCharAlpha (strInput) Then
'        intTargetRow = 6
'        strMatch = "*" & strInput & "*"
'    Else
'        intTargetRow = 2
'        strMatch = "*" & strInput & "*"
'    End If
'
'    With mshBillEdit
'        If .Rows = 1 Then Exit Sub
'
'        If lngStartRows > .Rows - 1 Then
'            MsgBox "已查询到最后！", vbInformation, gstrSysName
'            lngRows = 1
'        Else
'            lngRows = lngStartRows
'        End If
'
'        .SetFocus
'
'        '从指定的开始位置开始查找
'        For lngRows = lngRows To .Rows - 1
'            If .TextMatrix(lngRows, intTargetRow) Like strMatch Then
'                .MsfObj.TopRow = lngRows
'                .Row = lngRows
'                .Col = 1
'                mlngLastFindRows = lngRows + 1
'                blnFind = True
'                Exit For
'            End If
'        Next
'
'        '没找到就从头再找一次
'        If Not blnFind And lngStartRows > 1 Then
'            For lngRows = 1 To .Rows - 1
'                If .TextMatrix(lngRows, intTargetRow) Like strInput & "*" Then
'                    .MsfObj.TopRow = lngRows
'                    .Row = lngRows
'                    .Col = 1
'                    mlngLastFindRows = lngRows + 1
'                    Exit For
'                End If
'            Next
'        End If
'    End With
'End Sub

Private Function SelectStuff(ByVal strSeach As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:选择指定的卫生材料
    '参数:strKey-多选择的条件
    '返回:选择成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/09/17
    '-----------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long, strWhere As String
    Dim objCtl As Object: Dim vRect As RECT
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    
    On Error GoTo ErrHand
    Set objCtl = txtStuff
    vRect = zlControl.GetControlRect(objCtl.hwnd)
    lngH = objCtl.Height
    strKey = GetMatchingSting(strSeach)
      
    strTittle = "卫行材料选择"
    If strSeach = "" Then

        gstrSQL = " " & _
        "   Select Id,上级id,编码,名称,'' As 计算单位,0 as 末级 ,'' 站点,'' 建档时间" & _
        "   From 诊疗分类目录  " & _
        "   Where 类型='7' " & _
        "   Start With 上级id Is Null Connect By Prior Id =上级id " & _
        "   Union All  " & _
        "   Select I.ID,B.分类ID As 上级ID, I.编码, I.名称 || LPad(' ', 3, ' ') || I.规格 || LPad(' ', 3, ' ') || I.产地 As 名称" & _
        "       ,I.计算单位,1 as  末级,I.站点,to_char(I.建档时间,'yyyy-mm-dd') as 建档时间 " & _
        "   From 收费项目目录 I, 材料特性 T,诊疗项目目录 B " & _
        "   Where I.ID = T.材料id And   T.诊疗id=b.Id     And I.类别 = '4' And (I.撤档时间 Is Null Or I.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))"
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 2, strTittle, True, "", "", False, False, False, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False)
    Else
        
        strWhere = " And (I.编码 Like [1] OR N.名称 Like [1] OR ( N.简码 LIKE Upper([1]) and N.码类=[2]))"
        If IsNumeric(strSeach) Then                         '如果是数字,则只取编码
            If Mid(gSystem_Para.Para_输入方式, 1, 1) = "1" Then strWhere = " And (I.编码 Like [1] And N.码类=[2])"
        ElseIf zlStr.IsCharAlpha(strSeach) Then          '输入全是字母时只匹配简码
            If Mid(gSystem_Para.Para_输入方式, 2, 1) = "1" Then strWhere = " And (N.简码 Like Upper([1]) And N.码类=[2]) "
        ElseIf zlStr.IsCharChinese(strSeach) Then
            strWhere = " And (N.名称 Like [1] And N.码类=[2]) "
        End If
    
        gstrSQL = "" & _
        "   Select distinct I.ID,I.编码, I.名称 ||LPAD(' ',3,' ')||I.规格||LPAD(' ',3,' ')||I.产地 as 名称,I.计算单位,I.站点,to_char(I.建档时间,'yyyy-mm-dd') as 建档时间" & _
        "   From 收费项目目录 I,材料特性 T,收费项目别名 N" & _
        "   Where I.ID=T.材料ID and I.ID=N.收费细目ID and I.类别='4'" & _
        "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
        "        " & strWhere
        
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strKey, gSystem_Para.int简码方式 + 1)
    End If
    If blnCancel = True Then
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    '加载部门
    If rsTemp Is Nothing Then
        ShowMsgBox "没有满足条件的卫生材料,请检查!"
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    With rsTemp
        mstrPreStuffName = "[" & !编码 & "]" & !名称:
        objCtl.Text = mstrPreStuffName
        objCtl.Tag = zlStr.Nvl(!Id)
        mlng材料ID = !Id
    End With
    Call ShowData
    SelectStuff = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function



Private Sub cmdApply_Click()
    Dim lngRow As Long, lngRowCurr As Long
    Dim strObject As String, strObjectID As String
    
    With mshBillEdit
        lngRowCurr = .Row
        strObject = .TextMatrix(.Row, 3)
        strObjectID = .TextMatrix(.Row, 4)
        For lngRow = 1 To .Rows - 1
            If lngRow <> lngRowCurr And .TextMatrix(lngRow, 1) = "√" Then
                .TextMatrix(lngRow, 3) = strObject
                .TextMatrix(lngRow, 4) = strObjectID
            End If
        Next
    End With
End Sub

Private Sub cmdChoose_Click()
    Dim lngRow As Long, lngRows As Long
    With mshBillEdit
        lngRows = .Rows - 1
        For lngRow = 1 To lngRows
            .TextMatrix(lngRow, 1) = "√"
'            .TextMatrix(lngRow, 3) = ""
'            .TextMatrix(lngRow, 4) = ""
        Next
    End With
End Sub

Private Sub cmdClear_Click()
    Dim lngRow As Long, lngRows As Long
    With mshBillEdit
        lngRows = .Rows - 1
        For lngRow = 1 To lngRows
            .TextMatrix(lngRow, 1) = ""
            .TextMatrix(lngRow, 3) = ""
            .TextMatrix(lngRow, 4) = ""
        Next
    End With
End Sub

Private Sub cmdClose_Click()
    Dim i As Integer
    Dim j As Integer
    Dim strTemp As String
    
    If mblnSave = False Then
        With mshBillEdit
            For i = 1 To .Rows - 1
                For j = 1 To .Cols - 1
                    strTemp = strTemp & .TextMatrix(i, j) & "|"
                Next
            Next
        End With
        strTemp = txtStuff.Text & "|" & opt应用于(0).Value & "|" & opt应用于(4).Value & "|" & opt应用于(1).Value & "|" & _
                    opt应用于(2).Value & "|" & opt应用于(3).Value & "|" & strTemp
                    
        If strTemp <> mstr记录值 Then
            If MsgBox("有数据被修改了，是否退出？", vbYesNo, gstrSysName) = vbYes Then
                Unload Me
            End If
        Else
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub cmdStuff_Click()
    If SelectStuff("") = False Then Exit Sub
End Sub

Private Sub cmdRestore_Click()
    Call ShowData
End Sub

Private Sub CmdSave_Click()
    Dim strPara As String
    Dim lngRow As Long, lngRows As Long
    Dim rsTemp As New ADODB.Recordset
    Dim arrInput As Variant
    Dim strTmp As String
    Dim intType As Integer
    Dim i As Integer
    
    arrInput = Array()
    
    On Error GoTo ErrHand
    
    If mlng材料ID = 0 Then
        MsgBox "请先选择卫生材料！", vbInformation, gstrSysName
        txtStuff.SetFocus
        Exit Sub
    End If
    
    If mshBillEdit.Active = False Then
        MsgBox "没有找到任何库房，请在部门管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If opt应用于(0).Value = False Then
        For i = 0 To opt应用于.UBound
            If opt应用于(i).Value = True Then
                If MsgBox("该卫材存储库房应用范围为“" & opt应用于(i).Caption & "”是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                Else
                    Exit For
                End If
            End If
        Next
    End If
    
    '产生输入串并保存
    lngRows = mshBillEdit.Rows - 1
    For lngRow = 1 To lngRows
        If mshBillEdit.TextMatrix(lngRow, 1) <> "" Then
            strTmp = "1," & mshBillEdit.RowData(lngRow) & "|" & mshBillEdit.TextMatrix(lngRow, 4)
        Else
            strTmp = "0," & mshBillEdit.RowData(lngRow) & "|" & ""
        End If
        
        '可能由于服务科室过多，造成传入的参数字符长度过长，所以最大传入4K的字符串
        If Len(IIf(strPara = "", "", strPara & "!!") & strTmp) > 4000 Then
            ReDim Preserve arrInput(UBound(arrInput) + 1)
            arrInput(UBound(arrInput)) = strPara
            
            strPara = strTmp
        Else
            strPara = IIf(strPara = "", "", strPara & "!!") & strTmp
        End If
    Next
    
    ReDim Preserve arrInput(UBound(arrInput) + 1)
    arrInput(UBound(arrInput)) = strPara
    
    '参数：
    '   诊疗ID_IN参数_IN(设置标志,库房ID|科室ID,科室ID!!...),应用_IN(1-本卫生材料;2-本级所有卫生材料;3-分类下的所有卫生材料;4-所有材料,5-本品种所有规格)
    '设置标志:0-不设置卫材在该库房的存储；1-要设置
    If opt应用于(0).Value = True Then
        intType = 1
    ElseIf opt应用于(1).Value = True Then
        intType = 2
    ElseIf opt应用于(2).Value = True Then
        intType = 3
    ElseIf opt应用于(4).Value = True Then
        intType = 5
    Else
        intType = 4
    End If
    
    For i = 0 To UBound(arrInput)
        If arrInput(i) <> "" Then
            gstrSQL = "zl_卫生材料存储库房_UPDATE(" & mlng材料ID & ",'" & arrInput(i) & "'," & intType & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
    Next
    
    MsgBox "该卫生材料的存储库房和服务科室保存成功！", vbInformation, gstrSysName
    mblnSave = True
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub



Private Sub cmd参照_Click()
    Dim str科室 As String
    Dim str科室id As String
    Dim blnSel As Boolean
    Dim intRow As Integer
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "" & _
        "   Select  开单科室id, 开单科室id, 执行科室id ,K.名称 " & _
        "   From 收费执行科室 A," & _
        "       (Select 收费细目id From 收费执行科室 Where Rowid = " & _
        "           (Select Max(a.Rowid) From 收费执行科室 a,材料特性 c where a.收费细目id=c.材料id)" & _
        "       ) B ,部门表 K" & _
        "   Where A.收费细目id = b.收费细目id and a.开单科室id=K.id(+) "
    
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    With mshBillEdit
        For intRow = 1 To .Rows - 1
            
            str科室 = "": str科室id = ""
            rsTemp.Filter = "执行科室ID=" & .RowData(intRow)
            
            blnSel = False
            Do While Not rsTemp.EOF
                blnSel = True
                str科室 = str科室 & "," & zlStr.Nvl(rsTemp!名称)
                str科室id = str科室id & "," & zlStr.Nvl(rsTemp!开单科室id, 0)
                rsTemp.MoveNext
            Loop
            If str科室 <> "" Then
                str科室 = Mid(str科室, 2)
                str科室id = Mid(str科室id, 2)
                If str科室id = "0" Then str科室id = ""
            End If
            mshBillEdit.TextMatrix(intRow, 0) = intRow
            If blnSel Then
                mshBillEdit.TextMatrix(intRow, 1) = "√"
            Else
                mshBillEdit.TextMatrix(intRow, 1) = ""
            End If
            mshBillEdit.TextMatrix(intRow, 3) = str科室
            mshBillEdit.TextMatrix(intRow, 4) = str科室id
        Next
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If lvwItems.Visible Then
            lvwItems.Visible = False
            mshBillEdit.SetFocus
            Exit Sub
        Else
            Unload Me
            Exit Sub
        End If
    End If
    
    If KeyCode = vbKeyF3 Then
        Call txtfind_KeyPress(vbKeyReturn)
    End If
        
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strTemp As String
    Dim j As Integer
    
    mlngLastFindRows = 1
    
    Call InitFace
    Call ShowData
    If mbln编辑 = False Then
        mshBillEdit.Active = False
        frame.Enabled = False
        cmdSave.Visible = False
        cmdRestore.Visible = False
        cmdClear.Visible = False
        cmdClose.Caption = "退出(&X)"
    End If
    Call CtlEnableSet
    
    With mshBillEdit
        For i = 1 To .Rows - 1
            For j = 1 To .Cols - 1
                strTemp = strTemp & .TextMatrix(i, j) & "|"
            Next
        Next
    End With
    mstr记录值 = ""
    mstr记录值 = txtStuff.Text & "|" & opt应用于(0).Value & "|" & opt应用于(4).Value & "|" & opt应用于(1).Value & "|" & _
                opt应用于(2).Value & "|" & opt应用于(3).Value & "|" & strTemp
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnFind = False
    mblnSave = False
End Sub

Private Sub mshBillEdit_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    mshBillEdit.TextMatrix(Row, 1) = ""
    mshBillEdit.TextMatrix(Row, 3) = ""
    mshBillEdit.TextMatrix(Row, 4) = ""
    Cancel = True
End Sub

Private Sub mshBillEdit_CommandClick()
    Dim str服务对象 As String
    Dim objItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    cmdSave.Enabled = True
    str服务对象 = ""
    gstrSQL = "select distinct 服务对象 from 部门性质说明 where 部门ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取服务对象", Val(mshBillEdit.RowData(mshBillEdit.Row)))
    
    Do While Not rsTemp.EOF
        str服务对象 = str服务对象 & "," & rsTemp!服务对象
        rsTemp.MoveNext
    Loop
    
    If str服务对象 <> "" Then
        str服务对象 = Mid(str服务对象, 2)
        If InStr(1, str服务对象, 3) <> 0 Then
            str服务对象 = "0,1,2,3"
        ElseIf InStr(1, str服务对象, 1) <> 0 Or InStr(1, str服务对象, 2) <> 0 Then
            str服务对象 = str服务对象 & ",3"
        End If
    Else
        str服务对象 = "0"
    End If
    
    '排除掉部分不是开单科室的部门性质
    gstrSQL = " Select distinct ID,编码,名称 From 部门表 A,部门性质说明 B, Table(Cast(f_Num2List([1]) As zlTools.t_NumList)) C " & _
              " Where A.ID=B.部门ID And B.服务对象=C.Column_Value and (a.撤档时间 is null or to_char(a.撤档时间,'yyyy-mm-dd')='3000-01-01')" & _
              " And B.工作性质 Not In ('虚拟库房', '配制中心', '中药库', '西药库', '成药库', '制剂室', '中药房', '西药房', '成药房', '卫材库', '发料部门') "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取服务科室", str服务对象)

    If rsTemp.RecordCount = 0 Then
        MsgBox "未设置临床，医技等科室！[部门管理]", vbInformation, gstrSysName
        mshBillEdit.TextMatrix(mshBillEdit.Row, 3) = ""
        mshBillEdit.TextMatrix(mshBillEdit.Row, 4) = ""
    End If
    
    Call AddColumnHeader(False)
    Me.lvwItems.ListItems.Clear
    Me.lvwItems.Checkboxes = True
    
    Do While Not rsTemp.EOF
        Set objItem = Me.lvwItems.ListItems.Add(, "_" & rsTemp!Id, rsTemp!名称, , 3)
        objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = rsTemp!编码
        If InStr(1, "," & mshBillEdit.TextMatrix(mshBillEdit.Row, 4) & ",", "," & rsTemp!Id & ",") > 0 Then
            objItem.Checked = True
        End If
        rsTemp.MoveNext
    Loop
    If lvwItems.ListItems.Count <> 0 Then lvwItems.ListItems(1).Selected = True
    With Me.lvwItems
        .Left = Me.mshBillEdit.Left + 2300
        .Top = Me.mshBillEdit.Top + Me.mshBillEdit.CellTop + Me.mshBillEdit.RowHeight(Me.mshBillEdit.Row)
        If .Top + .Height > Me.ScaleHeight Then
            .Top = Me.ScaleHeight - .Height
        End If
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshBillEdit_EnterCell(Row As Long, Col As Long)
    With mshBillEdit
        If .Col = 3 Then
            .ColData(3) = 1
        End If
    End With
End Sub


Private Sub mshBillEdit_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim str服务对象 As String
    Dim objItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    
    With mshBillEdit
'            mblnKeyMethod = False
        If KeyCode = vbKeyReturn And .Col = 3 Then
'            mblnKeyMethod = True
'            Call mshBillEdit_CommandClick
            If Trim(.Text) = "" Then
                If .Row <> .Rows - 1 Then
                    .Row = .Row + 1
                End If
                Exit Sub
            End If
            cmdSave.Enabled = True
            str服务对象 = ""
            gstrSQL = "select distinct 服务对象 from 部门性质说明 where 部门ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取服务对象", Val(mshBillEdit.RowData(mshBillEdit.Row)))
            
            Do While Not rsTemp.EOF
                str服务对象 = str服务对象 & "," & rsTemp!服务对象
                rsTemp.MoveNext
            Loop
            If str服务对象 <> "" Then
                str服务对象 = Mid(str服务对象, 2)
                If InStr(1, str服务对象, 3) <> 0 Then str服务对象 = "0,1,2,3"
            Else
                str服务对象 = "0"
            End If
            gstrSQL = " Select /*+ Rule*/ distinct ID,编码,名称 From 部门表 A,部门性质说明 B, Table(Cast(f_Num2List([1]) As zlTools.t_NumList)) C " & _
                      " Where A.ID=B.部门ID And B.服务对象=C.Column_Value"
            gstrSQL = gstrSQL & " and (编码 like [2] or 简码 like [2] or 名称 like [2])"
            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取服务科室", str服务对象, UCase(.Text) & "%")

            If rsTemp.RecordCount = 0 Then
                MsgBox "未设置该临床科室！[部门管理]", vbInformation, gstrSysName
                mshBillEdit.TextMatrix(mshBillEdit.Row, 3) = ""
                mshBillEdit.TextMatrix(mshBillEdit.Row, 4) = ""
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            If rsTemp.RecordCount = 1 Then
                .TextMatrix(mshBillEdit.Row, 1) = "√"
                .TextMatrix(.Row, 4) = rsTemp!Id
                .Text = IIf(IsNull(rsTemp!名称), "", rsTemp!名称)
                .TextMatrix(.Row, 3) = .Text
            Else
                Call AddColumnHeader(False)
                Me.lvwItems.ListItems.Clear
                Me.lvwItems.Checkboxes = True
                
                Do While Not rsTemp.EOF
                    Set objItem = Me.lvwItems.ListItems.Add(, "_" & rsTemp!Id, rsTemp!名称, , 3)
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = rsTemp!编码
                    If InStr(1, "," & mshBillEdit.TextMatrix(mshBillEdit.Row, 4) & ",", "," & rsTemp!Id & ",") > 0 Then
                        objItem.Checked = True
                    End If
                    rsTemp.MoveNext
                Loop
                lvwItems.ListItems(1).Checked = True
                With Me.lvwItems
                    .Left = Me.mshBillEdit.Left + 2300
                    .Top = Me.mshBillEdit.Top + Me.mshBillEdit.CellTop + Me.mshBillEdit.RowHeight(Me.mshBillEdit.Row)
                    If .Top + .Height > Me.ScaleHeight Then
                        .Top = Me.ScaleHeight - .Height
                    End If
                    .ZOrder 0: .Visible = True
                    Cancel = True
                    .SetFocus
                End With
            End If
        End If
    End With
End Sub

Private Sub opt应用于_Click(Index As Integer)
    Dim i As Integer
    
    For i = 1 To opt应用于.UBound
        If i = Index Then
            opt应用于(i).FontBold = True
        Else
            opt应用于(i).FontBold = False
        End If
    Next
End Sub

Private Sub opt应用于_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call OS.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtFind_Change()
    mblnFind = False
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Sub txtfind_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode <> vbKeyReturn Then Exit Sub
'    If Trim(txtFind.Text) = "" Then Exit Sub
'
'    mlngLastFindRows = 1
'    FindRow Trim(txtFind.Text), 1
'    txtFind.SetFocus
'    zlControl.TxtSelAll txtFind
End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    
    On Error GoTo ErrHandle
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then
        If Trim(txtFind.Text) = "" Then Exit Sub
        If mstrFind <> Trim(txtFind.Text) And mblnFind = False Then
            gstrSQL = "Select a.ID,a.编码, a.名称, a.简码" & _
                       " From 部门表 A, 部门性质说明 B" & _
                       " Where a.Id = b.部门id And b.工作性质 In ('卫材库', '发料部门', '制剂室', '虚拟库房') And (a.编码 Like [1] Or a.名称 Like [1] Or a.简码 like [1])"

            Set mrs科室 = zlDatabase.OpenSQLRecord(gstrSQL, "科室查询", UCase(txtFind.Text) & "%")
            If mrs科室.RecordCount > 0 Then
                mblnFind = True
                Call FindData(mrs科室)
            Else
                MsgBox "没有找到你想要的数据！", vbInformation, gstrSysName
                txtFind.SetFocus
                zlControl.TxtSelAll txtFind
            End If
        Else
            If Not mrs科室.EOF Then
                mrs科室.MoveNext
                If Not mrs科室.EOF Then
                    Call FindData(mrs科室)
                Else
                    MsgBox "已查询到最后！", vbInformation, gstrSysName
                End If
            Else
                mrs科室.MoveFirst
                Call FindData(mrs科室)
            End If
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub txtStuff_Change()
    txtStuff.Tag = ""
End Sub

Private Sub txtStuff_GotFocus()
    zlControl.TxtSelAll txtStuff
    OS.OpenIme False
End Sub

Private Sub txtStuff_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strKey As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txtStuff.Tag <> "" Then OS.PressKey vbKeyTab: Exit Sub
    If txtStuff.Text = "" Then OS.PressKey vbKeyTab: Exit Sub
    strKey = Trim(Me.txtStuff.Text)
    If SelectStuff(strKey) = False Then Exit Sub
End Sub
Private Sub txtStuff_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub
Private Sub txtStuff_LostFocus()
    Me.txtStuff.Text = mstrPreStuffName
End Sub

Private Sub ShowData()
    '提取数据并显示出来
    Dim str库房ID As String, str科室 As String, str科室id As String
    Dim intRow As Integer, intRows As Integer
    Dim blnSel As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long
    
    Dim lng分类id As Long, strWhere As String
    Dim str站点 As String
    Dim lng诊疗项目id As Long
    
    Call cmdClear_Click
    On Error GoTo ErrHandle
    
    gstrSQL = "Select 分类ID,id From 诊疗项目目录 where id =(Select Max(诊疗ID) ID from 材料特性 where 材料id=[1]) "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取分类信息", mlng材料ID)
    If rsTemp.EOF Then
        lng分类id = 0
    Else
        lng分类id = Val(zlStr.Nvl(rsTemp!分类id))
        lng诊疗项目id = rsTemp!Id
    End If
    
    gstrSQL = "Select A.编码,A.名称 From 诊疗分类目录 A where id =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取分类信息", lng分类id)
   
    
    If rsTemp.EOF Then
        lbl分类.Caption = "指定的卫材分类："
    Else
        lbl分类.Caption = "指定的卫材分类：" & "[" & rsTemp!编码 & "]" & rsTemp!名称
    End If
    
'    If mblnFirst Then
    '提取材料信息
    strWhere = "": str站点 = ""
    If mlng材料ID <> 0 Then
        gstrSQL = " Select A.ID, A.编码, A.名称 ||LPAD(' ',3,' ')||A.规格||LPAD(' ',3,' ')||A.产地 as 名称,A.站点" & _
                  " From 收费项目目录 A" & _
                  " Where A.ID=[1]"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取材料信息", mlng材料ID)
        If rsTemp.EOF = False Then
            txtStuff.Text = "[" & rsTemp!编码 & "]" & rsTemp!名称
            mstrPreStuffName = txtStuff.Text

            txtStuff.Tag = zlStr.Nvl(rsTemp!Id)
            str站点 = zlStr.Nvl(rsTemp!站点)
        End If
        If str站点 <> "" Then
            strWhere = " And (站点=[1] or 站点 is null)" '
            lbl分类.Caption = lbl分类.Caption & "    站点：" & str站点
        End If
    End If
        
    '根据材料的用途分类提取所允许存储的库房
    gstrSQL = "" & _
        " Select ID,编码,名称,简码 From 部门表 " & _
        " Where ID in (Select distinct 部门id from 部门性质说明 where 工作性质 In('卫材库','发料部门','制剂室','虚拟库房')) and (撤档时间 is null or to_char(撤档时间,'yyyy-mm-dd')='3000-01-01')  " & _
        strWhere
    gstrSQL = gstrSQL & " Order By 名称 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "根据卫生材料提取所允许存储的库房", str站点)
    
    With mshBillEdit
        .Rows = 2:
        If rsTemp.EOF Then
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
            Next
        End If
        Do While Not rsTemp.EOF
            .TextMatrix(.Rows - 1, 2) = rsTemp!名称
            .TextMatrix(.Rows - 1, 5) = rsTemp!编码
            .TextMatrix(.Rows - 1, 6) = rsTemp!简码
            .RowData(.Rows - 1) = rsTemp!Id
            .Rows = .Rows + 1
            str库房ID = str库房ID & "," & rsTemp!Id
            rsTemp.MoveNext
        Loop
        If str库房ID <> "" Then
            str库房ID = Mid(str库房ID, 2)
            .Rows = .Rows - 1
            .Active = True
        Else
            .Active = False
        End If
    End With
    'End If
    
    '取所有库房
    '    str库房ID = ""
    intRows = mshBillEdit.Rows - 1
    '    For intRow = 1 To intRows
    '        str库房ID = str库房ID & "," & mshBillEdit.RowData(intRow)
    '    Next
    '
    '    If str库房ID <> "" Then str库房ID = Mid(str库房ID, 2)
    
    
    '将相应数据组织后装入单据控件
    gstrSQL = "" & _
        "   Select A.收费细目ID,A.开单科室ID,A.执行科室ID,B.名称 " & _
        "   From 收费执行科室 A,部门表 B " & _
        "   Where A.开单科室ID=B.ID(+) And A.收费细目ID=[1] And A.执行科室ID in (select * from Table(Cast(f_Num2List([2]) As zlTools.t_NumList)))" & _
        "   Order by A.执行科室ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取已设置的诊疗执行科室数据", mlng材料ID, str库房ID)
       
    If rsTemp.RecordCount = 0 And mlng材料ID <> 0 Then
        '
        gstrSQL = "" & _
            " Select a.收费细目id,a.开单科室id, a.执行科室id, b.名称" & _
            " From 收费执行科室 A, 部门表 B, (Select 材料id From 材料特性 Where 诊疗id = [1] And Rownum < 2) C " & _
           " Where a.开单科室id = b.Id(+) And a.收费细目id = c.材料id And a.执行科室id in (select * from Table(Cast(f_Num2List([2]) As zlTools.t_NumList)))" & _
            " Order By a.执行科室id"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng诊疗项目id, str库房ID)
            
        If rsTemp.RecordCount = 0 And mlng材料ID <> 0 Then
            gstrSQL = "" & _
                "   Select A.收费细目ID,A.开单科室ID,A.执行科室ID,B.名称 " & _
                "   From 收费执行科室 A,部门表 B, " & _
                "       ( Select A.ID From 收费项目目录 A,材料特性 B,收费执行科室 C " & _
                "         Where A.ID=B.材料ID And A.类别='4' And A.ID=C.收费细目ID And Rownum<2 ) C" & _
                "   Where A.开单科室ID=B.ID(+) And A.收费细目ID=C.ID And A.执行科室ID in (select * from Table(Cast(f_Num2List([1]) As zlTools.t_NumList)))" & _
                "   Order by A.执行科室ID"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str库房ID)
            If rsTemp.RecordCount <> 0 Then
                MsgBox "当前规格卫材未设置存储库房，提取剂型相同的规格卫材的存储库房做为缺省数据！", vbInformation, gstrSysName
            End If
        Else
            MsgBox "当前规格卫材未设置存储库房，提取同品种下规格卫材的存储库房做为缺省数据！", vbInformation, gstrSysName
        End If
    End If
    With mshBillEdit
        For intRow = 1 To intRows
            str科室 = "": str科室id = ""
            rsTemp.Filter = "执行科室ID=" & .RowData(intRow)
            rsTemp.Sort = "开单科室ID"
            blnSel = False
            Do While Not rsTemp.EOF
                blnSel = True
                str科室 = str科室 & "," & zlStr.Nvl(rsTemp!名称)
                str科室id = str科室id & "," & Nvl(rsTemp!开单科室id, 0)
                rsTemp.MoveNext
            Loop
            If str科室 <> "" Then
                str科室 = Mid(str科室, 2)
                str科室id = Mid(str科室id, 2)
                If str科室id = "0" Then str科室id = ""
            End If
            .TextMatrix(intRow, 0) = intRow
            If blnSel Then .TextMatrix(intRow, 1) = "√"
            .TextMatrix(intRow, 3) = str科室
            .TextMatrix(intRow, 4) = str科室id
        Next
    End With
    mblnFirst = False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitFace()
    '初始化控件
    With mshBillEdit
        .Rows = 2
        .Cols = 7
        
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "选择"
        .TextMatrix(0, 2) = "存储库房"
        .TextMatrix(0, 3) = "服务科室"
        .TextMatrix(0, 4) = "服务科室ID"
        .TextMatrix(0, 5) = "库房编码"
        .TextMatrix(0, 6) = "库房简码"
        .TextMatrix(1, 0) = "1"
        .ColData(0) = 5
        .ColData(1) = -1
        .ColData(2) = 5
        .ColData(3) = 1
        .ColData(4) = 5
        .ColData(5) = 5
        .ColData(6) = 5
        .ColWidth(0) = 300
        .ColWidth(1) = 500
        .ColWidth(2) = 1500
        .ColWidth(3) = 5000
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        
        .PrimaryCol = 1
        .LocateCol = 1
        .AllowAddRow = False
        .Active = True
    End With
End Sub

Private Sub mshBillEdit_AfterAddRow(Row As Long)
    Dim lngCurRow As Long
    
    '修改行序号
    With mshBillEdit
        For lngCurRow = Row To .Rows - 1
            .TextMatrix(lngCurRow, 0) = lngCurRow
        Next
    End With
End Sub

Private Sub mshBillEdit_AfterDeleteRow()
    Dim lngCurRow As Long
    '修改行序号
    With mshBillEdit
        For lngCurRow = IIf(.Row <> 1, .Row - 1, .Row) To .Rows - 1
            .TextMatrix(lngCurRow, 0) = lngCurRow
        Next
    End With
End Sub
Private Sub AddColumnHeader(Optional ByVal bln卫材 As Boolean = True)
    If bln卫材 Then
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "名称", "名称", 3000
            .Add , "编码", "编码", 1000
            .Add , "计算单位", "计算单位", 800
        End With
        With Me.lvwItems
            .ColumnHeaders("编码").Position = 1
            .SortKey = .ColumnHeaders("编码").Index - 1
            .SortOrder = lvwAscending
        End With
        lvwItems.Tag = "1"
    Else
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "名称", "名称", 3000
            .Add , "编码", "编码", 1000
        End With
        With Me.lvwItems
            .ColumnHeaders("编码").Position = 1
            .SortKey = .ColumnHeaders("编码").Index - 1
            .SortOrder = lvwAscending
        End With
        lvwItems.Tag = "2"
    End If
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    Dim blnCancel As Boolean
    Dim lngRow As Long, lngRows As Long
    Dim str科室 As String, str科室id As String
     
    If lvwItems.Tag = "1" Then
'        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
'        With Me.lvwItems
'            If mlng材料ID <> Mid(.SelectedItem.Key, 2) Then
'                mlng材料ID = Mid(.SelectedItem.Key, 2)
'                Me.txtStuff.Tag = "[" & .SelectedItem.SubItems(.ColumnHeaders("编码").Index - 1) & "]" & .SelectedItem.Text
'                Me.txtStuff.Text = Me.txtStuff.Tag
'                Call ShowData
'            End If
'            Me.txtStuff.SetFocus
'            Call OS.PressKey(vbKeyTab)
'        End With
    Else
        '循环提取用户所选择的科室
        lngRows = lvwItems.ListItems.Count
        For lngRow = 1 To lngRows
            If lvwItems.ListItems(lngRow).Checked Then
                str科室 = str科室 & "," & lvwItems.ListItems(lngRow).Text
                str科室id = str科室id & "," & Mid(lvwItems.ListItems(lngRow).Key, 2)
            End If
        Next
        If str科室 <> "" Then
            str科室 = Mid(str科室, 2)
            str科室id = Mid(str科室id, 2)
        End If
        
        If str科室 <> "" Then mshBillEdit.TextMatrix(mshBillEdit.Row, 1) = "√"
        mshBillEdit.TextMatrix(mshBillEdit.Row, 3) = str科室
        mshBillEdit.TextMatrix(mshBillEdit.Row, 4) = str科室id
        mshBillEdit.SetFocus
        If mshBillEdit.Rows - 1 > mshBillEdit.Row Then mshBillEdit.Row = mshBillEdit.Row + 1
    End If
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        Call lvwItems_DblClick
    End Select
End Sub

Private Sub lvwItems_LostFocus()
    Me.lvwItems.Visible = False
End Sub

Public Sub ShowMe(ByVal frmParent As Object, ByVal lng材料ID As Long, ByVal bln编辑 As Boolean)
    On Error Resume Next
    mblnFirst = True
    mlng材料ID = lng材料ID
    mbln编辑 = bln编辑
    Me.Show 1, frmParent
End Sub
Private Sub CtlEnableSet()
    '功能:设置相关控件的Enable
    '---------------------------------------------------------------------------------------------------------------------
    Dim strReg As String
    err = 0: On Error GoTo ErrHand:
    '格式:3位字符构成,1,代表允许，0代表不允许,如111.其中第一位代表所有,第二位代表本级所有,第三位代表分类下所有
    strReg = zlDatabase.GetPara("允许应用于的范围", glngSys, mlngModule)
    If Len(strReg) < 3 Then
        '恢认全选中
        strReg = "111"
    End If
    opt应用于(3).Enabled = IIf(Val(Mid(strReg, 1, 1)) = 1, True, False)
    opt应用于(1).Enabled = IIf(Val(Mid(strReg, 2, 1)) = 1, True, False)
    opt应用于(2).Enabled = IIf(Val(Mid(strReg, 3, 1)) = 1, True, False)
    If opt应用于(1).Enabled = False And opt应用于(1).Value = True Then
         opt应用于(1).Value = False
    End If
    If opt应用于(2).Enabled = False And opt应用于(2).Value = True Then
         opt应用于(2).Value = False
    End If
    If opt应用于(3).Enabled = False And opt应用于(3).Value = True Then
         opt应用于(3).Value = False
    End If
    opt应用于(0).Value = True
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub FindData(ByVal rsTemp As ADODB.Recordset)
    '查询数据
    Dim i As Integer
    
    With mshBillEdit
        For i = 1 To .Rows - 1
            If .RowData(i) = rsTemp!Id Then
                .SetFocus
                .Row = i
                .Col = 1
                .MsfObj.TopRow = i
            End If
        Next
    End With
End Sub
