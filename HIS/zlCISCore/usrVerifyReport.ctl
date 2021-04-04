VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.UserControl usrVerifyReport 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000E&
   ClientHeight    =   3225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5700
   ScaleHeight     =   3225
   ScaleWidth      =   5700
   Begin VB.PictureBox picComment 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   1980
      TabIndex        =   2
      Top             =   2085
      Width           =   2010
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         Height          =   600
         Left            =   465
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   -15
         Width           =   2160
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "检验备注"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   45
         TabIndex        =   4
         Top             =   45
         Width           =   390
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf2 
      Height          =   1335
      Left            =   630
      TabIndex        =   0
      Top             =   540
      Width           =   2355
      _cx             =   4154
      _cy             =   2355
      Appearance      =   0
      BorderStyle     =   0
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
      BackColorFixed  =   -2147483643
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483645
      GridColorFixed  =   -2147483645
      TreeColor       =   -2147483639
      FloodColor      =   192
      SheetBorder     =   -2147483639
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   240
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
      Begin VB.Line lnX 
         Index           =   0
         Visible         =   0   'False
         X1              =   0
         X2              =   1785
         Y1              =   135
         Y2              =   135
      End
      Begin VB.Line lnY 
         Index           =   0
         Visible         =   0   'False
         X1              =   825
         X2              =   825
         Y1              =   0
         Y2              =   1215
      End
   End
   Begin zl9CISCore.VsfGrid vsf 
      Height          =   1530
      Left            =   1470
      TabIndex        =   1
      Top             =   0
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   2699
   End
   Begin VB.Shape shp 
      BorderColor     =   &H80000003&
      Height          =   435
      Left            =   0
      Top             =   150
      Width           =   585
   End
End
Attribute VB_Name = "usrVerifyReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strSQL As String
Private rsTmp As New ADODB.Recordset

Private mlng病历id As Long                      '外界传入
Private mlng医嘱id As Long                      '外界传入
Private mstr性别 As String

Private mblnMode As Boolean '为真是表示是用户进行的编辑，这时才赋值

Private mDispMode As Boolean
Private mReturnErrnumber As Long
Private mReturnErrDescription As String
Private mblnMoved As Boolean '数据是否已转出

Private mblnModify As Boolean
Private mblnCommon As Boolean

Private mstrSQL As String
Private mlngLoop As Long
Private mblnLoaded As Boolean

Private Enum mCol
    检验项目 = 1
    检验结果
    结果标志
    结果类型
    计算公式
    单位
    结果参考
    
    细菌 = 0
    计数
    抗生素
    结果
    类型
    培养描述
End Enum

Private Enum COLOR
    
    红色 = &HFF&
    兰色 = &HFF0000
    黑色 = 0
    非焦点 = &HFFEBD7
    焦点 = &HFFCC99
    橙色 = &H40C0&
    报警背景色 = &H40C0&
    报警前景色 = &H8000000E
    超标背景色 = &H80C0FF
    低标背景色 = &H80FFFF
    超标前景色 = &H80000012
    默认前景色 = &H80000008
End Enum

Private mobjParentObject As Object

Public Property Set ParentObject(vData As Object)
    Set mobjParentObject = vData
End Property

Public Property Get ParentObject() As Object
    Set ParentObject = mobjParentObject
End Property

Private Property Let Modified(vData As Boolean)
    
    On Error Resume Next
    
    If mobjParentObject Is Nothing Then Exit Property
    
    mobjParentObject.Modified = vData
    
End Property

Private Property Get Modified() As Boolean
    
    On Error Resume Next
    
    If mobjParentObject Is Nothing Then Exit Property
    
    Modified = mobjParentObject.Modified
    
End Property

'-------------------------------------------------------------------------------------------------------------------
Private Function AppendRows(ByVal objVsf As Object, ByRef objLineX As Variant, ByRef objLineY As Variant, Optional ByVal lngHideRows As Long = 0) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能:补充表格控件的空行
    '参数:objVsf 要补充行的表格控件对象
    '返回:若成功返回True,否则返回 False
    '--------------------------------------------------------------------------------------------------------
    Dim lngTop As Long
    Dim lngLoop As Long
    Dim lngIndex As Long
    Dim lngLastRow As Long
    
    On Error GoTo errHand
    
    If objVsf.Rows = 0 Then Exit Function
    
    For lngLoop = objVsf.Rows - 1 To 1 Step -1
        If objVsf.RowHidden(lngLoop) = False Then
            lngLastRow = lngLoop
            Exit For
        End If
    Next
    
    lngTop = objVsf.Cell(flexcpTop, lngLastRow, 0) + objVsf.RowHeight(lngLastRow)
    
    '1.隐藏所有的线
    For lngLoop = 1 To objLineX.UBound
        objLineX(lngLoop).Visible = False
    Next
    
    For lngLoop = 1 To objLineY.UBound
        objLineY(lngLoop).Visible = False
    Next
    
    '2.重新计算需要的纵线
    For lngLoop = 1 To objVsf.Cols - 1

        If objLineY.UBound < lngLoop Then Load objLineY(lngLoop)

        With objLineY(lngLoop)

            .ZOrder

            .X1 = objVsf.Cell(flexcpLeft, 0, lngLoop) - 15
            .X2 = .X1
            .Y1 = lngTop
            .Y2 = objVsf.Height

            .BorderColor = objVsf.GridColor

            .Visible = True
        End With

    Next

    '3.重新计算需要的横线
    lngIndex = 0
    Do While (lngTop + objVsf.RowHeight(0)) < objVsf.Height

        lngIndex = lngIndex + 1
        If objLineX.UBound < lngIndex Then Load objLineX(lngIndex)

        With objLineX(lngIndex)

            .ZOrder

            .X1 = 0
            .X2 = objVsf.Width
            .Y1 = lngTop + objVsf.RowHeight(0) + 15
            .Y2 = .Y1

            .BorderColor = objVsf.GridColor

            .Visible = True

            lngTop = .Y1
        End With

    Loop
        
    AppendRows = True
    
    Exit Function
    
errHand:
    
End Function
Private Function FilterKeyAscii(ByVal KeyAscii As Long, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Long
            
    FilterKeyAscii = KeyAscii
    
    If Chr(KeyAscii) = "'" Then
        FilterKeyAscii = 0
        Exit Function
    End If
    
    If KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or KeyAscii = vbKeyBack Then
        Exit Function
    End If
    
    Select Case bytMode
    Case 1      '纯数字
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 2      '正小数
        If InStr("0123456789.", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 99
        If InStr(KeyCustom, Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    End Select
    
End Function

Private Sub ApplyResultColor(vsf As Object, ByVal lngRow As Long, ByVal lngCol As Long, ByVal bytMode As Byte)
    '-----------------------------------------------------------------------------------------
    '功能:
    '-----------------------------------------------------------------------------------------
    Dim lngColor As Long, lngForeColor As Long
    
    Select Case bytMode
        Case 0, 1
            lngColor = &H80000005
            lngForeColor = COLOR.默认前景色
        Case 5 '异常低、高
            lngColor = COLOR.报警背景色
            lngForeColor = COLOR.报警前景色
        Case 2
            lngColor = COLOR.低标背景色
            lngForeColor = COLOR.超标前景色
        Case Else
            lngColor = COLOR.超标背景色
            lngForeColor = COLOR.超标前景色
    End Select
    
    vsf.Cell(flexcpBackColor, lngRow, lngCol, lngRow, lngCol) = lngColor
    vsf.Cell(flexcpForeColor, lngRow, lngCol, lngRow, lngCol) = lngForeColor
End Sub

Private Function CalcDefaultFlag(ByVal strValue As String, ByVal strReference As String, Optional ByVal bytMode As Byte = 1, _
    Optional ByVal strAlarmLow As String, Optional ByVal strAlarmHigh As String) As String
    
    '--------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '返回:
    '--------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    If Len(Trim(strValue)) = 0 Then CalcDefaultFlag = "": Exit Function
    
    CalcDefaultFlag = ""
    
    If InStr(strReference, vbCrLf) > 0 Then strReference = Mid(strReference, 1, InStr(strReference, vbCrLf) - 1)
    If Trim(strReference) = "" Then Exit Function
                
    If bytMode = 2 Or bytMode = 3 Then '定性、半定量
        If bytMode = 2 Or InStr(strReference, "～") = 0 Or Trim(strValue) Like "*阳*" Or Trim(strValue) Like "*+*" Or _
            Trim(strValue) Like "*±*" Or Trim(strValue) Like "*阴*" Or Trim(strValue) Like "*-*" Then
            '定性或无范围参考的半定量
            If (Len(Trim(strReference)) > 0 And (Trim(strReference) Like (Trim(strValue) & "*") Or Trim(strReference) Like ("*" & Trim(strValue)))) Or _
                (Not (Trim(strValue) Like "*阳*" Or Trim(strValue) Like "*+*" Or Trim(strValue) Like "*±*")) Then
                CalcDefaultFlag = ""
            Else
                CalcDefaultFlag = "异常"
            End If
            Exit Function
        Else
            '获取半定量值
            For i = 1 To Len(Trim(strValue))
                If InStr("01234567890.", Mid(strValue, i, 1)) > 0 Then Exit For
            Next
            If i > Len(Trim(strValue)) Then Exit Function
            strValue = Val(Mid(strValue, i))
        End If
    End If
    '高低判断
    If Len(Trim(strAlarmLow)) > 0 And Val(strAlarmLow) <> 0 Then
        If Val(strValue) < Val(strAlarmLow) Then
            CalcDefaultFlag = "↓↓"
            Exit Function
        End If
    End If
    If Len(Trim(strAlarmHigh)) > 0 And Val(strAlarmHigh) <> 0 Then
        If Val(strValue) > Val(strAlarmHigh) Then
            CalcDefaultFlag = "↑↑"
            Exit Function
        End If
    End If
    If InStr(strReference, "～") > 0 Then
        
        '如果小于参考低值
        If Val(strValue) < Val(Mid(strReference, 1, InStr(strReference, "～") - 1)) And _
            Len(Trim(Mid(strReference, 1, InStr(strReference, "～") - 1))) > 0 Then
            CalcDefaultFlag = "↓"
        End If
        
        '如果大于参考高值
        If Val(strValue) > Val(Mid(strReference, InStr(strReference, "～") + 1)) And _
            Len(Trim(Mid(strReference, InStr(strReference, "～") + 1))) > 0 Then
            CalcDefaultFlag = "↑"
        End If
            
    End If
End Function

Private Function CalcExpress(ByVal vsf As Object, ByVal strExPress As String) As Single
    
    '--------------------------------------------------------------------------------------------------------
    '功能:在表格中计算某一表达式的结果
    '参数:vsf           存放数据的表格
    '     strExpress    要计算的表达式
    '返回:计算结果值
    '--------------------------------------------------------------------------------------------------------
    
    Dim strTmpPress As String
    Dim rs As New ADODB.Recordset
    
    Dim lngTmpID As Long
    Dim lngLeftPos As Long
    Dim lngRightPos As Long
    Dim lngLoop As Long
    Dim sglValue As Single
    
    CalcExpress = 0
    
    strTmpPress = strExPress
    If strTmpPress <> "" Then
        
        lngLeftPos = InStr(strTmpPress, "[")
        lngRightPos = InStr(strTmpPress, "]")
        
        Do While lngLeftPos > 0
        
            lngTmpID = Val(Mid(strTmpPress, lngLeftPos + 1, lngRightPos - lngLeftPos - 1))
            
            '判断lngTmpID是否也是计算项目
            For lngLoop = 1 To vsf.Rows - 1
                If Val(vsf.RowData(lngLoop)) = lngTmpID Then
                    If Trim(vsf.TextMatrix(lngLoop, mCol.计算公式)) <> "" Then
                        '是计算项目,先计算出此结果
                        sglValue = CalcExpress(vsf, Trim(vsf.TextMatrix(lngLoop, mCol.计算公式)))
                    Else
                        '不是计算项目,直接取此结果
                        sglValue = Val(vsf.TextMatrix(lngLoop, mCol.检验结果))
                    End If
                    
                    Exit For
                    
                End If
            Next
            
            '在当前表格中没有此检验项目,认为结果为零
            If lngLoop = vsf.Rows Then sglValue = 0
                                        
            '以结果替代表达式中的计算因子
            strTmpPress = Mid(strTmpPress, 1, lngLeftPos - 1) & sglValue & Mid(strTmpPress, lngRightPos + 1)
            
            '查下一个计算因子的位置
            lngLeftPos = InStr(strTmpPress, "[")
            lngRightPos = InStr(strTmpPress, "]")
        Loop
                
        '计算表达式的结果
        On Error Resume Next
        Call OpenRecord(rs, "SELECT " & strTmpPress & " AS 结果 FROM DUAL", "检验结果")
        If rs.BOF = False Then CalcExpress = zlCommFun.NVL(rs("结果"), 0)
        On Error GoTo 0
        
    End If
    
End Function

Private Sub CalcPosition(ByRef x As Single, ByRef y As Single, ByVal objBill As Object)
    '----------------------------------------------------------------------
    '功能： 计算X,Y的实际坐标，并考虑屏幕超界的问题
    '参数： X---返回横坐标参数
    '       Y---返回纵坐标参数
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hWnd, objPoint)
    
    x = objPoint.x * 15 + objBill.CellLeft
    y = objPoint.y * 15 + objBill.CellTop + objBill.CellHeight
End Sub

Private Function ShowOpenList(Optional strText As String, Optional blnWhere As Boolean = False) As Byte
    '-----------------------------------------------------------------------------------------
    '功能:打开列表结构的诊疗检验标本数据
    '返回:出错返回2;成功返回1;取消返回0
    '-----------------------------------------------------------------------------------------
    Dim strLvw As String
    Dim sglX As Single
    Dim sglY As Single
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    
    strLvw = "编码,900,0,1;取值,1800,0,0;结果标志,900,0,0"

    ShowOpenList = 2
    
    strSQL = "SELECT ROWNUM AS ID,编码,取值,DECODE(结果标志,1,'',2,'↓',3,'↑',4,'异常',5,'↓↓',6,'↑↑','') AS 结果标志 FROM 检验项目取值 A WHERE 项目id=[1]"
    
    If blnWhere Then
        strSQL = strSQL & " AND (A.编码 Like [2] OR A.取值 Like [2])"
    End If
        
    Set rs = OpenSQLRecord(strSQL, "检验结果", CLng(Val(vsf.RowData(vsf.Row))), "%" & strText & "%")
    
    If rs.BOF Then
        
        ShowOpenList = 0
        
        Exit Function
    End If
    
    If rs.RecordCount = 1 And blnWhere Then GoTo Over
        
    Call CalcPosition(sglX, sglY, vsf)
    
    If frmSelectList.ShowSelect(Me, rs, strLvw, sglX + 60, sglY, 3600, 4500, "检验取值选择", "请从下表中选择一个取值") Then
        GoTo Over
    End If
    
    Exit Function
    
Over:
    vsf.EditText = zlCommFun.NVL(rs("取值").Value)
    vsf.Cell(flexcpData, vsf.Row, vsf.Col) = zlCommFun.NVL(rs("取值").Value)
    vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.NVL(rs("取值").Value)
    vsf.TextMatrix(vsf.Row, mCol.结果标志) = zlCommFun.NVL(rs("结果标志").Value)
    
    Call ApplyResultColor(vsf, vsf.Row, mCol.检验结果, _
            Decode(vsf.TextMatrix(vsf.Row, mCol.结果标志), "↑", 3, "↓", 2, "异常", 4, "↓↓", 5, "↑↑", 6, 1))
    
    ShowOpenList = 1
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

'公共方法、属性
Public Property Get DispMode() As Boolean
    '是否为显示模式
    DispMode = mDispMode
End Property

Public Property Let DispMode(ByVal New_DispMode As Boolean)
    mDispMode = New_DispMode
    ShowUsrControl mlng医嘱id, Not mDispMode
    PropertyChanged "DispMode"
    
    If mDispMode Then
        vsf.Body.Editable = flexEDNone
    End If
    
End Property

Public Property Get ID病人病历() As Long
    '返回病人病历ID
    
    ID病人病历 = mlng病历id
End Property

Public Property Let ID病人病历(ByVal New_ID病人病历 As Long)
    '设置病人病历ID,并检查该病历是不是存在
    
    mlng病历id = New_ID病人病历
    ShowUsrControl mlng医嘱id, Not mDispMode
    
End Property

Public Sub SetDiagItem(ByVal New_医嘱ID As Long, ByVal New_发送号)
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    mlng医嘱id = New_医嘱ID
    strSQL = "SELECT DECODE(相关id,NULL,ID,相关id) AS ID FROM 病人医嘱记录 WHERE ID=[1]"
    
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    Set rs = OpenSQLRecord(strSQL, "检验报告专用纸", mlng医嘱id)
    If rs.BOF = False Then mlng医嘱id = rs("ID").Value
        
        
    mblnModify = True
    mblnCommon = True       '是否为普通检验项目
    
    strSQL = "SELECT DISTINCT A.执行状态,D.项目类别 FROM 病人医嘱发送 A,病人医嘱记录 B,检验报告项目 C,检验项目 D WHERE B.诊疗项目id=C.诊疗项目id AND C.报告项目id=D.诊治项目id AND A.医嘱id=B.ID AND B.相关id=[1]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    Set rs = OpenSQLRecord(strSQL, "检验报告专用纸", mlng医嘱id)
    If rs.BOF = False Then
        mblnModify = (rs("执行状态").Value <> 1)
        mblnCommon = (rs("项目类别").Value <> 2)
    End If
                
    vsf.Visible = False
    vsf2.Visible = False
    
    If mblnCommon Then
        With vsf
            .Cols = 0
            .NewColumn "", 255, 4
            .NewColumn "项目", 2100, 1
            .NewColumn "结果", 900, 1, , 1
            .NewColumn "标志", 990, 1, " ", 1
            .NewColumn "类型", 0, 1
            .NewColumn "公式", 0, 1
            .NewColumn "单位", 0, 1
            .NewColumn "参考", 1500, 1
            
            .ExtendLastCol = True
            .Body.Appearance = flexFlat
            '.AppearanceFlat = True
            .Body.BorderStyle = flexBorderNone
            .Body.BackColorFixed = .Body.BackColor
            .FixedCols = 1
            
            .Cell(flexcpFontBold, 0, 0, 0, vsf.Cols - 1) = True
            .Visible = True
        End With
        
        If mblnModify = False Then
            vsf.EditMode(mCol.检验结果) = 0
            vsf.EditMode(mCol.结果标志) = 0
            vsf.ComboList(mCol.检验结果) = ""
            vsf.ComboList(mCol.结果标志) = ""
        End If
        
    Else
        With vsf2
            .Cols = 6
            .FixedRows = 2
            .FixedCols = 0
            
            .MergeRow(0) = True
            .MergeCol(0) = True
            .MergeCol(1) = True
            .MergeCol(5) = True
            
            .TextMatrix(0, 0) = "细菌"
            .TextMatrix(1, 0) = "细菌"
            
            .TextMatrix(0, 1) = "计数"
            .TextMatrix(1, 1) = "计数"
            
            .Cell(flexcpText, 0, 2, 0, 4) = "药敏试验"
            .TextMatrix(1, 2) = "抗生素"
            .TextMatrix(1, 3) = "结果"
            .TextMatrix(1, 4) = "类型"
            
            .TextMatrix(0, 5) = "培养描述"
            .TextMatrix(1, 5) = "培养描述"
            
            .ColWidth(0) = 1800
            .ColWidth(1) = 510
            .ColWidth(2) = 1800
            .ColWidth(3) = 1200
            .ColWidth(4) = 810
            .ColWidth(5) = 30
            
            .ColAlignment(0) = 1
            .ColAlignment(1) = 1
            .ColAlignment(2) = 1
            .ColAlignment(3) = 1
            .ColAlignment(4) = 1
            .ColAlignment(5) = 1
            
            .Cell(flexcpAlignment, 0, 0, 1, vsf2.Cols - 1) = 4
            
            .ExtendLastCol = True
            .Visible = True
            .Cell(flexcpFontBold, 0, 0, 1, vsf2.Cols - 1) = True
            Call AppendRows(vsf2, lnX, lnY)
            
            mblnModify = False
        End With
    End If
End Sub

Public Property Get Get医嘱id() As Long
        
    Get医嘱id = mlng医嘱id
        
End Property

Private Sub SetErr(lngErrNum As Long, strErr As String)
    '设置错误描述及错误号
    '如果lngErrNum=-1 表示 控件自己定义的错误
    mReturnErrnumber = lngErrNum
    mReturnErrDescription = strErr
End Sub

Public Property Get ReturnErrNumber() As Long
    '返回最后一次的错误号
    ReturnErrNumber = mReturnErrnumber
End Property

Public Property Get ReturnErrDescription() As String
    '返回最后一次错误描述字符串
    ReturnErrDescription = mReturnErrDescription
End Property

'-------------------------------------------------------------------------------------------------------------------
Private Function CheckStrValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0, Optional ByRef strError As String) As Boolean
'检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
        
    If InStr(strInput, "'") > 0 Or InStr(strInput, "|") > 0 Then
        strError = "所输入内容含有非法字符。"
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            strError = "所输入内容不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字母。"
            Exit Function
        End If
    End If
    
    CheckStrValid = True
End Function

'------------------------------------------------------------------------------------------------------------

Private Sub ShowUsrControl(lngKey As Long, Optional ByVal blnEditMode As Boolean = False)
    '------------------------------------------------------------------------------------------------------------
    '功能：外部调用显示手术概要的过程
    '------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHandle
        
    mDispMode = Not blnEditMode
    
    '按逻辑应先初始控件
    Call InitData
    
    If gcnOracle Is Nothing Then SetErr -1, "连接对象没有初始化": Exit Sub
    If gcnOracle.State <> adStateOpen Then SetErr -1, "连接对象没有连接": Exit Sub

    Call ReadData
    
    Exit Sub
    
ErrHandle:

    If Ambient.UserMode = False Or InDesign = False Then SetErr Err.Number, Err.Description:    Exit Sub
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub SetgcnOracle()
    '-------------------------------------------------------------------------------------------------
    '接口
    '-------------------------------------------------------------------------------------------------
    
    Call InitCommon(gcnOracle)
End Sub

Private Sub InitData()
    '初始化窗体
    
    Dim strTmp As String
    
    On Error GoTo ErrHandle
        
    If Not gcnOracle Is Nothing Then
        If Not gcnOracle.State <> adStateOpen Then
            If Ambient.UserMode = True Then

            End If
        End If
    
    End If
    
    
    
    Exit Sub
    
ErrHandle:

    If Ambient.UserMode = False Or InDesign = False Then
        SetErr Err.Number, Err.Description
        Exit Sub
    End If
    
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
End Sub

Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：读出数据库里的数据
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strTmp As String
    Dim strDec As String
    
    On Error GoTo ErrHandle
    
    mblnLoaded = True
    
    '连接检查
    If gcnOracle Is Nothing Then SetErr -1, "连接对象没有初始化": Exit Function
    If gcnOracle.State <> adStateOpen Then SetErr -1, "连接对象没有连接": Exit Function
                                        
    txt.Text = ""
    
    If mblnCommon Then
        '清除数据
        vsf.Rows = 2
        vsf.RowData(1) = 0
        vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
        
        '读取并装载数据
    
    '    检验项目 = 0
    '    检验结果
    '    结果标志
    '    结果参考
    '    结果类型
    '    计算公式
        
        mstrSQL = "SELECT 所见项id,所见内容 FROM 病人病历所见单 WHERE 病历id=[1]"
        
        If mlng病历id > 0 Then
            mstrSQL = "SELECT DISTINCT F.所见项id AS ID," & _
                               "G.中文名 AS 检验项目," & _
                               "F.所见内容 AS 检验结果," & _
                               "zlGetReference(F.所见项id,A.标本部位,DECODE(E.性别,'男',1,'女',2,0),E.出生日期) AS 结果参考," & _
                               "D.结果类型," & _
                               "D.单位," & _
                               "D.计算公式, " & _
                               "a.诊疗项目ID, " & _
                               "h.排列序号,h.检验备注 " & _
                        "FROM 病人医嘱记录 A," & _
                             "检验项目 D," & _
                             "病人信息 E, " & _
                             "检验报告项目 h, " & _
                             "(SELECT 所见项id,所见内容 FROM 病人病历所见单 WHERE NVL(行,0)=0 AND 病历id=[1]) F, " & _
                             "(SELECT Distinct x.医嘱id,y.检验备注 FROM 检验项目分布 x,检验标本记录 y WHERE x.标本id=y.ID And x.医嘱id=[2]) h, " & _
                             "诊治所见项目 G " & _
                        "Where A.相关ID = [2] And A.相关id=h.医嘱id(+) " & _
                              "AND E.病人ID=A.病人id " & _
                              "AND F.所见项id=D.诊治项目ID(+) " & _
                              "AND G.ID=D.诊治项目ID " & _
                              "AND a.诊疗项目ID = h.诊疗项目ID " & _
                              "AND D.诊治项目ID = h.报告项目ID " & _
                              " order by a.诊疗项目ID,h.排列序号 "
        Else
            mstrSQL = "SELECT DISTINCT C.报告项目ID AS ID," & _
                               "G.中文名 AS 检验项目," & _
                               "Decode(d.结果类型,3,Decode(F.所见内容,Null,'-',''),F.所见内容) AS 检验结果," & _
                               "zlGetReference(C.报告项目ID,A.标本部位,DECODE(E.性别,'男',1,'女',2,0),E.出生日期) AS 结果参考," & _
                               "D.结果类型," & _
                               "B.计算单位 AS 单位," & _
                               "D.计算公式,C.排列序号,h.检验备注 " & _
                        "FROM 病人医嘱记录 A," & _
                             "诊疗项目目录 B," & _
                             "检验报告项目 C," & _
                             "检验项目 D," & _
                             "病人信息 E, " & _
                             "(SELECT 所见项id,所见内容 FROM 病人病历所见单 WHERE NVL(行,0)=0 AND 病历id=[1]) F, " & _
                             "(SELECT Distinct x.医嘱id,y.检验备注 FROM 检验项目分布 x,检验标本记录 y WHERE x.标本id=y.ID And x.医嘱id=[2]) h, " & _
                             "诊治所见项目 G " & _
                        "Where A.相关ID = [2] And A.相关id=h.医嘱id(+) " & _
                              "AND E.病人ID=A.病人id " & _
                              "AND A.诊疗项目ID=B.ID " & _
                              "AND C.诊疗项目ID=B.ID " & _
                              "AND F.所见项id(+)=C.报告项目ID " & _
                              "AND D.诊治项目ID=C.报告项目ID " & _
                              "AND G.ID=C.报告项目ID Order By C.排列序号"
        End If
                    
        If mblnMoved Then
            mstrSQL = Replace(mstrSQL, "病人病历所见单", "H病人病历所见单")
            mstrSQL = Replace(mstrSQL, "病人医嘱记录", "H病人医嘱记录")
            mstrSQL = Replace(mstrSQL, "检验标本记录", "H检验标本记录")
            mstrSQL = Replace(mstrSQL, "检验项目分布", "H检验项目分布")
        End If
        Set rs = OpenSQLRecord(mstrSQL, "检验报告", mlng病历id, mlng医嘱id)
        If rs.BOF = False Then
            
            txt.Text = zlCommFun.NVL(rs("检验备注").Value)
            
            Do While Not rs.EOF
                
                If zlCommFun.NVL(rs("结果类型").Value) = 1 Then
                    strDec = "0.0000"
                Else
                    strDec = ""
                End If
                
                If Val(vsf.RowData(vsf.Rows - 1)) > 0 Then
                    vsf.Rows = vsf.Rows + 1
                End If
                
                vsf.RowData(vsf.Rows - 1) = zlCommFun.NVL(rs("ID").Value)
                vsf.TextMatrix(vsf.Rows - 1, mCol.检验项目) = zlCommFun.NVL(rs("检验项目").Value)
                
                vsf.TextMatrix(vsf.Rows - 1, mCol.结果参考) = zlCommFun.NVL(rs("结果参考").Value)
                
                strTmp = zlCommFun.NVL(rs("检验结果").Value)
                If strTmp <> "" Then
                    vsf.TextMatrix(vsf.Rows - 1, mCol.检验结果) = IIf(strDec = "", Split(strTmp, "'")(0), Format(Split(strTmp, "'")(0), strDec))
                    If UBound(Split(strTmp, "'")) > 0 Then vsf.TextMatrix(vsf.Rows - 1, mCol.结果标志) = Split(strTmp, "'")(1)
                    If UBound(Split(strTmp, "'")) > 1 Then vsf.TextMatrix(vsf.Rows - 1, mCol.结果参考) = Split(strTmp, "'")(2)
                End If
                    
                vsf.TextMatrix(vsf.Rows - 1, mCol.结果类型) = zlCommFun.NVL(rs("结果类型").Value)
                vsf.TextMatrix(vsf.Rows - 1, mCol.计算公式) = zlCommFun.NVL(rs("计算公式").Value)
                vsf.TextMatrix(vsf.Rows - 1, mCol.单位) = zlCommFun.NVL(rs("单位").Value)
                
                Call ApplyResultColor(vsf, vsf.Rows - 1, mCol.检验结果, _
                    Decode(vsf.TextMatrix(vsf.Rows - 1, mCol.结果标志), "↑", 3, "↓", 2, "异常", 4, "↓↓", 5, "↑↑", 6, 1))
                
                rs.MoveNext
            Loop
        End If
        
        For mlngLoop = 1 To vsf.Rows - 1
            Call ApplyResultColor(vsf, mlngLoop, mCol.检验结果, _
                Decode(vsf.TextMatrix(mlngLoop, mCol.结果标志), "↑", 3, "↓", 2, "异常", 4, "↑↑", 5, "↓↓", 6, 1))
        Next
        
        '自动设置高度
        If (vsf.Rows * (vsf.Body.RowHeight(0) + 15) + 30) > UserControl.Height Then
            UserControl.Height = vsf.Rows * (vsf.Body.RowHeight(0) + 15) + 30
        End If
    Else
        vsf2.Rows = 3
        vsf2.RowData(2) = 0
        vsf2.Cell(flexcpText, 2, 0, 2, vsf2.Cols - 1) = ""
        
        mstrSQL = _
            "SELECT A.ID," & _
                  "E.中文名 AS 细菌," & _
                  "A.检验结果 AS 计数," & _
                  "B.抗生素," & _
                  "B.结果," & _
                  "B.类型," & _
                  "B.颜色值," & _
                  "A.培养描述,h.检验备注 " & _
            "FROM 检验普通结果 A," & _
                 "检验项目 C," & _
                 "检验标本记录 D," & _
                 "检验细菌 E," & _
                 "(SELECT Distinct x.医嘱id,y.检验备注 FROM 检验项目分布 x,检验标本记录 y WHERE x.标本id=y.ID And x.医嘱id=[1]) h, " & _
                 "病人医嘱记录 F," & _
                 "(SELECT A.细菌结果ID," & _
                         "B.中文名 AS 抗生素," & _
                         "A.结果," & _
                         "DECODE(A.结果类型,'R','255','I','16711680','S','0','0') AS 颜色值," & _
                         "DECODE(A.结果类型,'R','耐药','I','中介','S','敏感','') AS 类型 " & _
                  "FROM 检验药敏结果 A," & _
                       "检验用抗生素 B " & _
                  "Where A.抗生素ID = B.ID " & _
                 ") B "
                 
        mstrSQL = mstrSQL & _
            "Where A.检验项目id = C.诊治项目ID(+) " & _
                "AND C.项目类别(+)=2 " & _
                "AND A.记录类型 =D.报告结果 " & _
                "AND D.ID=A.检验标本ID " & _
                "AND A.细菌id =E.ID(+) " & _
                "AND A.ID=B.细菌结果ID(+) " & _
                "AND (D.医嘱id=F.ID Or D.医嘱id=F.相关ID) " & _
                "AND F.ID=[1] And f.ID=h.医嘱id(+) ORDER BY A.ID "
                        
        mstrSQL = "SELECT ID,细菌,计数,抗生素,结果,类型,颜色值,培养描述,检验备注 FROM (" & mstrSQL & ") A "
                        
        If mblnMoved Then
            mstrSQL = Replace(mstrSQL, "检验普通结果", "H检验普通结果")
            mstrSQL = Replace(mstrSQL, "检验标本记录", "H检验标本记录")
            mstrSQL = Replace(mstrSQL, "病人医嘱记录", "H病人医嘱记录")
            mstrSQL = Replace(mstrSQL, "检验药敏结果", "H检验药敏结果")
            mstrSQL = Replace(mstrSQL, "检验项目分布", "H检验项目分布")
        End If
        Set rs = OpenSQLRecord(mstrSQL, "检验报告", mlng医嘱id)
        If rs.BOF = False Then
            
            txt.Text = zlCommFun.NVL(rs("检验备注").Value)
            
            Do While Not rs.EOF
                If Val(vsf2.RowData(vsf2.Rows - 1)) > 0 Then
                    vsf2.Rows = vsf2.Rows + 1
                End If
                
                vsf2.RowData(vsf2.Rows - 1) = zlCommFun.NVL(rs("ID").Value)
                
                vsf2.TextMatrix(vsf2.Rows - 1, mCol.细菌) = zlCommFun.NVL(rs("细菌").Value)
                vsf2.TextMatrix(vsf2.Rows - 1, mCol.计数) = zlCommFun.NVL(rs("计数").Value)
                vsf2.TextMatrix(vsf2.Rows - 1, mCol.抗生素) = zlCommFun.NVL(rs("抗生素").Value)
                vsf2.TextMatrix(vsf2.Rows - 1, mCol.结果) = zlCommFun.NVL(rs("结果").Value)
                vsf2.TextMatrix(vsf2.Rows - 1, mCol.类型) = zlCommFun.NVL(rs("类型").Value)
                vsf2.TextMatrix(vsf2.Rows - 1, mCol.培养描述) = zlCommFun.NVL(rs("培养描述").Value)
                
                vsf2.Cell(flexcpForeColor, vsf2.Rows - 1, mCol.结果) = zlCommFun.NVL(rs("颜色值").Value)
                
                rs.MoveNext
            Loop
        End If
        
        Dim lngSvrKey As Long
        Dim strSpace As String
        Dim lngLoop As Long
        
        For lngLoop = 2 To vsf2.Rows - 1
            If lngSvrKey <> Val(vsf2.RowData(lngLoop)) Then
                lngSvrKey = Val(vsf2.RowData(lngLoop))
                strSpace = IIf(strSpace = " ", "", " ")
            End If
            
            vsf2.TextMatrix(lngLoop, mCol.细菌) = vsf2.TextMatrix(lngLoop, mCol.细菌) & strSpace
            vsf2.TextMatrix(lngLoop, mCol.计数) = vsf2.TextMatrix(lngLoop, mCol.计数) & strSpace
            vsf2.TextMatrix(lngLoop, mCol.培养描述) = vsf2.TextMatrix(lngLoop, mCol.培养描述) & strSpace
            
        Next
        
        '自动设置高度
        If (vsf2.Rows * (vsf2.RowHeight(0) + 15) + 30) > UserControl.Height Then
            UserControl.Height = vsf2.Rows * (vsf2.RowHeight(0) + 15) + 30
        End If
        Call AppendRows(vsf2, lnX, lnY)
    End If
    
    mblnLoaded = False
    
    Exit Function
    
ErrHandle:
    
    If Ambient.UserMode = False Or InDesign = False Then SetErr Err.Number, Err.Description:    Exit Function
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
End Function

Public Sub ClearData()
    '------------------------------------------------------------------------------------------------------------------
    '功能:接口
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    vsf.Cell(flexcpText, 1, mCol.检验结果, vsf.Rows - 1, mCol.检验结果) = ""
    vsf.Cell(flexcpText, 1, mCol.结果标志, vsf.Rows - 1, mCol.结果标志) = ""
End Sub

Public Function SaveData(lng病人ID As Long, lng主页ID As Long, lng病历ID As Long, strReturnSQL As String, strError As String) As Boolean
    Dim lngLoop As Long
    Dim strTmp As String
    Dim strSQL() As String
    Dim strValue As String
    
    ReDim Preserve strSQL(0 To vsf.Rows)
    
    If mblnCommon = False Then
        strReturnSQL = ""
        SaveData = True
        Exit Function
    End If
    
    Call vsf_AfterEdit(vsf.Row, vsf.Col)
    
'    strSQL(0) = "ZL_病人病历内容_DELETE(" & lng病历ID & ")"
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            
            Select Case Val(Left(vsf.TextMatrix(lngLoop, mCol.结果标志), 1))
                Case 3
                    vsf.TextMatrix(lngLoop, mCol.结果标志) = "↑"
                Case 2
                    vsf.TextMatrix(lngLoop, mCol.结果标志) = "↓"
                Case 4
                    vsf.TextMatrix(lngLoop, mCol.结果标志) = "异常"
                Case 5
                    vsf.TextMatrix(lngLoop, mCol.结果标志) = "↓↓"
                Case 6
                    vsf.TextMatrix(lngLoop, mCol.结果标志) = "↑↑"
    '            Case Else
    '                vsf.TextMatrix(Row, Col) = ""
            End Select
            
            strValue = vsf.TextMatrix(lngLoop, mCol.检验结果) & "''" & vsf.TextMatrix(lngLoop, mCol.结果标志) & "''" & vsf.TextMatrix(lngLoop, mCol.结果参考)
            
            strSQL(lngLoop) = "ZL_病人病历所见单_SAVE(" & lng病历ID & "," & _
                                                        lngLoop & "," & _
                                                        "2,'" & _
                                                        vsf.TextMatrix(lngLoop, mCol.检验项目) & "'," & _
                                                        "NULL," & _
                                                        "NULL," & _
                                                        "NULL," & _
                                                        "NULL," & _
                                                        "NULL," & _
                                                        "NULL," & _
                                                        "NULL," & _
                                                        Val(vsf.RowData(lngLoop)) & "," & _
                                                        vsf.TextMatrix(lngLoop, mCol.结果类型) & ",'" & _
                                                        vsf.TextMatrix(lngLoop, mCol.单位) & "','" & _
                                                        strValue & "'" & _
                                                        ")"
        End If
    Next
        
    strTmp = ""
    For lngLoop = 0 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then
            If strTmp = "" Then
                strTmp = strSQL(lngLoop)
            Else
                strTmp = strTmp & Chr(9) & strSQL(lngLoop)
            End If
        End If
    Next
    
    '返回SQL语句
    strReturnSQL = strTmp
    
    SaveData = True
    
End Function

Private Sub UserControl_Initialize()

    '初始化控件属性
    
    On Error GoTo ErrHandle
    
                
    Exit Sub
    
ErrHandle:
    If Ambient.UserMode = False Or InDesign = False Then SetErr Err.Number, Err.Description:    Exit Sub
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function InDesign() As Boolean
    
    '功能：判断当前运行程序是否在VB的工程环境中
    
    On Error Resume Next
    
    Debug.Print 1 / 0
    If Err.Number <> 0 Then Err.Clear: InDesign = True
    
End Function

Private Sub UserControl_InitProperties()
    '初始病人病历为0
    mlng病历id = 0
    mDispMode = True
    mblnMoved = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mDispMode = PropBag.ReadProperty("DispMode", True)
    mblnMoved = PropBag.ReadProperty("DataMoved", False)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", BorderStyleSettings.flexBorderNone)
End Sub

Public Property Get BorderStyle() As BorderStyleSettings
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleSettings)
    UserControl.BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Let Locked(ByVal vData As Boolean)
    MsgBox "a"
End Property

Private Sub UserControl_Resize()
    On Error Resume Next
    
    With shp
        .Left = 0
        .Top = 0
        .Width = UserControl.Width
        .Height = UserControl.Height - picComment.Height
    End With
    
    With vsf
        .Left = 15
        .Top = 15
        .Width = UserControl.Width - .Left - 15
        .Height = UserControl.Height - .Top - 15
    End With
    
    With vsf2
        .Left = vsf.Left
        .Top = vsf.Top
        .Width = vsf.Width
        .Height = vsf.Height
    End With
    
    picComment.Move 0, shp.Top + shp.Height, shp.Width
    txt.Move txt.Left, -15, picComment.Width - txt.Left, picComment.Height + 15
    
    Call AppendRows(vsf2, lnX, lnY)
End Sub

Private Sub UserControl_Terminate()
    If rsTmp.State = adStateOpen Then rsTmp.Close
    Set rsTmp = Nothing
    
    On Error Resume Next
    
    Set mobjParentObject = Nothing
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("DispMode", mDispMode, True)
    Call PropBag.WriteProperty("DataMoved", mblnMoved, False)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, BorderStyleSettings.flexBorderNone)
End Sub

Private Sub UserControl_Show()
    Dim objCtl As Control
         
    '只在运行时显示
    
    On Error Resume Next
    
    If Ambient.UserMode = True And InDesign = False Then
        If mDispMode Then
            For Each objCtl In Controls
                If UCase(TypeName(objCtl)) <> UCase("ImageList") Then
                    objCtl.Enabled = False
                End If
            Next
        End If
    End If
    
    If mblnLoaded = False Then InitData
        
    
End Sub

Public Property Get Text() As String
    '为每一个控件加上文本转储属性
    Dim lngLoop As Long
    Dim strTmp As String
    Dim strSvrKey As Long
    
    '通过用户输入的内容得到转储文本
    strTmp = "检验报告：" & vbCrLf
    If mblnCommon Then
        For lngLoop = 0 To vsf.Rows - 1
            strTmp = strTmp & MidB(vsf.TextMatrix(lngLoop, mCol.检验项目) & Space(50), 1, 50)
            strTmp = strTmp & MidB(vsf.TextMatrix(lngLoop, mCol.检验结果) & Space(20), 1, 20)
            strTmp = strTmp & MidB(vsf.TextMatrix(lngLoop, mCol.结果标志) & Space(20), 1, 20)
            strTmp = strTmp & vsf.TextMatrix(lngLoop, mCol.结果参考) & vbCrLf
            If lngLoop = 0 Then strTmp = strTmp & "------------------------------------------------------------------------------------------" & vbCrLf
            If lngLoop = vsf.Rows - 1 Then strTmp = strTmp & "------------------------------------------------------------------------------------------" & vbCrLf
        Next
    Else
        strTmp = strTmp & MidB(vsf2.TextMatrix(1, mCol.细菌) & Space(40), 1, 40)
        strTmp = strTmp & MidB(vsf2.TextMatrix(1, mCol.计数) & Space(20), 1, 20)
        strTmp = strTmp & MidB(vsf2.TextMatrix(1, mCol.抗生素) & Space(40), 1, 40)
        strTmp = strTmp & MidB(vsf2.TextMatrix(1, mCol.结果) & Space(20), 1, 20)
        strTmp = strTmp & MidB(vsf2.TextMatrix(1, mCol.类型) & Space(20), 1, 20)
        strTmp = strTmp & vsf2.TextMatrix(1, mCol.培养描述) & vbCrLf
        strTmp = strTmp & "-----------------------------------------------------------------------------------------------------" & vbCrLf
        
        For lngLoop = 2 To vsf2.Rows - 1
            
            If strSvrKey <> vsf2.RowData(lngLoop) Then
                
                strTmp = strTmp & MidB(vsf2.TextMatrix(lngLoop, mCol.细菌) & Space(40), 1, 40)
                strTmp = strTmp & MidB(vsf2.TextMatrix(lngLoop, mCol.计数) & Space(20), 1, 20)
                
            End If
            strTmp = strTmp & MidB(vsf2.TextMatrix(lngLoop, mCol.抗生素) & Space(40), 1, 40)
            strTmp = strTmp & MidB(vsf2.TextMatrix(lngLoop, mCol.结果) & Space(20), 1, 20)
            strTmp = strTmp & MidB(vsf2.TextMatrix(lngLoop, mCol.类型) & Space(20), 1, 20)
            
            If strSvrKey <> vsf2.RowData(lngLoop) Then
                
                strSvrKey = vsf2.RowData(lngLoop)
                strTmp = strTmp & vsf2.TextMatrix(lngLoop, mCol.培养描述)
                
            End If
            strTmp = strTmp & vbCrLf
        Next
        
        strTmp = strTmp & "-----------------------------------------------------------------------------------------------------" & vbCrLf
    End If
    
    Text = strTmp
    
End Property

Private Sub UserControl_EnterFocus()
    On Error Resume Next
    
    UserControl.Parent.CallBack_GotFocus
    
End Sub


Private Sub vsf_AfterDeleteCell(ByVal Row As Long, ByVal Col As Long)
'    mblnChangeEdit = True
'    Call AdjustEnableState
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strReference As String
    Dim LngCount As Long
    
    If Col = mCol.结果标志 Then
        Select Case Val(Left(vsf.TextMatrix(Row, mCol.结果标志), 1))
            Case 3
                vsf.TextMatrix(Row, Col) = "↑"
            Case 2
                vsf.TextMatrix(Row, Col) = "↓"
            Case 4
                vsf.TextMatrix(Row, Col) = "异常"
            Case 5
                vsf.TextMatrix(Row, Col) = "↓↓"
            Case 6
                vsf.TextMatrix(Row, Col) = "↑↑"
'            Case Else
'                vsf.TextMatrix(Row, Col) = ""
        End Select
        Call ApplyResultColor(vsf, Row, mCol.检验结果, Decode(vsf.TextMatrix(Row, mCol.结果标志), "↑", 3, "↓", 2, "异常", 4, "↓↓", 5, "↑↑", 6, 1))
    End If
    
    If Col = mCol.检验结果 And Val(vsf.TextMatrix(Row, mCol.结果类型)) <> 2 Then
        
        '产生缺省的结果标志
        vsf.TextMatrix(Row, mCol.结果标志) = Format(CalcDefaultFlag(Trim(vsf.TextMatrix(Row, Col)), Trim(vsf.TextMatrix(Row, mCol.结果参考)), Val(vsf.TextMatrix(Row, mCol.结果类型))), "0.0000")
        
        '根据结果应用颜色标志
        Call ApplyResultColor(vsf, Row, mCol.检验结果, Decode(vsf.TextMatrix(Row, mCol.结果标志), "↑", 3, "↓", 2, "异常", 4, "↓↓", 5, "↑↑", 6, 1))
                                
                                
        '自动计算计算项目结果
        For mlngLoop = 1 To vsf.Rows - 1
            If Trim(vsf.TextMatrix(mlngLoop, mCol.计算公式)) <> "" Then
                
                vsf.TextMatrix(mlngLoop, Col) = Format(CalcExpress(vsf, Trim(vsf.TextMatrix(mlngLoop, mCol.计算公式))), "0.0000")
                
                '产生缺省的结果标志
                vsf.TextMatrix(mlngLoop, mCol.结果标志) = CalcDefaultFlag(Trim(vsf.TextMatrix(mlngLoop, Col)), Trim(vsf.TextMatrix(mlngLoop, mCol.结果参考)), Val(vsf.TextMatrix(mlngLoop, mCol.结果类型)))
        
                '根据结果应用颜色标志
                Call ApplyResultColor(vsf, Row, mCol.检验结果, Decode(vsf.TextMatrix(Row, mCol.结果标志), "↑", 3, "↓", 2, "异常", 4, "↓↓", 5, "↑↑", 6, 1))
            End If
        Next
        
    End If

'    mblnChangeEdit = True
'    Call AdjustEnableState
End Sub



Private Sub vsf_BeforeComboList(ByVal OldCol As Long, ByVal NewCol As Long, ComboList As String, Cancel As Boolean)
    On Error Resume Next
    
    If mblnModify = False Then Exit Sub
    
    '1-正常、2-偏低、3-偏高、4-阳性
    '1:数字,2:文字，3：阴阳型(+-)
    If NewCol = mCol.结果标志 Then
        Select Case Val(vsf.TextMatrix(vsf.Row, mCol.结果类型))
            Case 1  '数字
                ComboList = "1-正常|2-偏低|3-偏高"
'                ComboList = "|↓|↑"
            Case 2  '定性
                ComboList = "1-正常|4-异常"
'                ComboList = "1-正常|4-异常"
            Case 3  '半定量
                ComboList = "1-正常|2-偏低|3-偏高|4-异常"
        End Select
    ElseIf NewCol = mCol.检验结果 Then
        ComboList = "" '"|-|+|--|++|+-"
    End If
End Sub

Private Sub vsf_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Col = mCol.检验结果
    Cancel = True
End Sub

Private Sub vsf_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    
    '如果是文字型的
    On Error Resume Next
    
    If mblnModify = False Then Exit Sub
    
    If NewCol = mCol.检验结果 Then
        If Trim(vsf.TextMatrix(NewRow, mCol.计算公式)) <> "" Then
            vsf.EditMode(NewCol) = 0
        Else
            vsf.EditMode(NewCol) = 1
        End If
    
    
        Select Case Val(vsf.TextMatrix(NewRow, mCol.结果类型))
        Case 2
            vsf.ComboList(mCol.检验结果) = "..."
            vsf.VsfComboList = "..."
        Case 3
            vsf.ComboList(mCol.检验结果) = " "
            vsf.VsfComboList = " |-|+|--|++|+-"
        Case Else
            vsf.ComboList(mCol.检验结果) = ""
            vsf.VsfComboList = ""
        End Select
        
    End If
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Select Case ShowOpenList(, False)
    Case 0
        '没有匹配的项目
        'MsgBox "没有找到相匹配的结果！", vbInformation, gstrSysName
        
    Case 1
        '选取了一个项目
'        mblnChangeEdit = True
'        Call AdjustEnableState
    End Select
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    
    Dim strSvrText As String
    
    If mblnModify = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Then
        '对于2-文字型的情况
        If Val(vsf.TextMatrix(Row, mCol.结果类型)) <> 2 Then Exit Sub
        
        If InStr(vsf.EditText, "'") > 0 Then
            Cancel = True
            Exit Sub
        End If

        strSvrText = vsf.EditText
        Select Case ShowOpenList(vsf.EditText, True)
        Case 0
            '没有匹配的项目
            vsf.Cell(flexcpData, Row, Col) = strSvrText
        Case 1
            '选取了一个项目
'            mblnChangeEdit = True
'            Call AdjustEnableState
        Case 2
            '取消了本次选择
            Cancel = True

            vsf.Cell(flexcpData, Row, Col) = vsf.Cell(flexcpData, Row, Col)
            vsf.TextMatrix(Row, Col) = vsf.Cell(flexcpData, Row, Col)

        End Select
    Else
'        mblnChangeEdit = True
'        Call AdjustEnableState
    End If
End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    
    If mblnModify = False Then Exit Sub
    
    If KeyAscii = vbKeyReturn Then Exit Sub
    
    Select Case Val(vsf.TextMatrix(vsf.Row, mCol.结果类型))
    Case 1
        KeyAscii = FilterKeyAscii(KeyAscii, 2)
    Case Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End Select
    
End Sub

Private Sub vsf2_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsf2, lnX, lnY)
End Sub

Private Sub vsf2_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsf2, lnX, lnY)
End Sub
'数据是否转出
Public Property Get DataMoved() As Boolean
    DataMoved = mblnMoved
End Property

Public Property Let DataMoved(ByVal vNewValue As Boolean)
    mblnMoved = vNewValue
End Property


