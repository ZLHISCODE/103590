VERSION 5.00
Object = "{B0475000-7740-11D1-BDC3-0020AF9F8E6E}#6.1#0"; "TTF16.ocx"
Begin VB.Form frmDockView 
   BorderStyle     =   0  'None
   Caption         =   "预览窗体"
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox PicDy 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Index           =   0
      Left            =   1140
      ScaleHeight     =   885
      ScaleWidth      =   1170
      TabIndex        =   1
      Top             =   1470
      Visible         =   0   'False
      Width           =   1170
   End
   Begin TTF160Ctl.F1Book F1Main 
      Height          =   1305
      Left            =   180
      TabIndex        =   0
      Top             =   105
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   2302
      _0              =   $"frmDockView.frx":0000
      _1              =   $"frmDockView.frx":0409
      _2              =   $"frmDockView.frx":0812
      _3              =   $"frmDockView.frx":0C1B
      _4              =   $"frmDockView.frx":1024
      _count          =   5
      _ver            =   2
   End
End
Attribute VB_Name = "frmDockView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Doc As cTableEPR, mblnInit As Boolean
Private Sub F1Main_TopLeftChanged()
Dim i As Integer, strCellKey As String
    On Error GoTo errHand
    If mblnInit Then Exit Sub
    For i = 1 To PicDy.UBound
        If ChkControl(PicDy(i)) Then
            If PicDy(i).Picture.Handle <> 0 Then
                strCellKey = Split(PicDy(i).Tag, "|")(1)
                Call PaintPictureOnTable(strCellKey)
            End If
        End If
    Next
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    F1Main.Top = 0: F1Main.Left = 0: F1Main.Width = Me.ScaleWidth: F1Main.Height = Me.ScaleHeight
End Sub
Public Sub zlRefresh(Tmp As cTableEPR)
'功能：刷新界面
Dim l As Long, lCount As Long
    On Error GoTo errHand
    '清窗图片控件
    Set Doc = Tmp
    PicDy(0).Visible = False
    For l = 1 To PicDy.UBound
        If ChkControl(PicDy(l)) Then
            Unload PicDy(l)
        End If
    Next
    
    With F1Main '初始化表格
        .DeleteRange .MinRow, .MinCol, .MaxRow, .MaxCol, F1ShiftRows
        .MaxCol = 4: .MaxRow = 4
    End With
    
    If Doc.ReadFileStructure Then   '读取文件结构
        Doc.ReadFileContent Doc.mblnMove   '读取文件内容
    Else
        Exit Sub
    End If
    Call RefreshF1Main
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub RefreshF1Main()
Dim lngRow As Long, lngCol As Long, lngCell As Long, vCell As F1CellFormat, lngCount As Long, strShow As String
    mblnInit = True
    With F1Main
        .DeleteRange .MinRow, .MinCol, .MaxRow, .MaxCol, F1ShiftRows
        .ShowTabs = F1TabsOff
        .AllowMoveRange = False '移动选中区域
        .AllowFillRange = False '拖动范围赋值,无事件不可控制
        .AllowInCellEditing = False '单元格编辑
        .AllowEditHeaders = False '编辑列头
        .AllowDesigner = False  '允许设计
        .AllowDelete = False '提示是英文的，最好不要允许而自已通过KeyDown控制
        .ShowLockedCellsError = False '对锁定单元格进行编辑时的消息提示
        .ScrollToLastRC = False '允许滚动到最后一个单元格
        .ColWidthUnits = F1ColWidthUnitsTwips '列宽计算单位为堤
        .DefaultFontName = "宋体"
        .DefaultFontSize = 9
        .MaxCol = Doc.Cells.Cols
        .MaxRow = Doc.Cells.Rows

        '定行高列宽
        For lngRow = 1 To .MaxRow
            .RowHeight(lngRow) = Doc.Cells.Cell(lngRow, 1).Height
        Next
        For lngCol = 1 To .MaxCol
            .ColWidthTwips(lngCol) = Doc.Cells.Cell(1, lngCol).Width
        Next
        
        lngCount = Doc.Cells.Count
        For lngCell = 1 To lngCount
            lngRow = Doc.Cells(lngCell).Row: lngCol = Doc.Cells(lngCell).Col
            With Doc.Cells.Cell(lngRow, lngCol)
                '指定区域
                If .Merge And InStr(.MergeRange, ";") > 0 Then 'MergeRange数据格式 (左上方)行,列;(右下方)行,列
                    F1Main.SetSelection Split(Split(.MergeRange, ";")(0), ",")(0), Split(Split(.MergeRange, ";")(0), ",")(1), Split(Split(.MergeRange, ";")(1), ",")(0), Split(Split(.MergeRange, ";")(1), ",")(1)
                Else
                    F1Main.SetSelection lngRow, lngCol, lngRow, lngCol
                End If
                Set vCell = F1Main.CreateNewCellFormat
                If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then '只有合并单元格首个或非合并单元格才刷新
'                    vCell.ProtectionLocked = .保留对象  '是否锁定,保护区域,行控,列控,签名、保存时写入Database
                    vCell.MergeCells = .Merge
                    vCell.WordWrap = True
                    vCell.FontName = .FontName          '字体>宋体</字体>
                    vCell.FontSize = .FontSize          '<字号>9</字号>
                    vCell.FontBold = .FontBold          '<粗体>False</粗体>
                    vCell.FontItalic = .FontItalic        '<斜体>False</斜体>
                    vCell.FontUnderline = .FontUnderline     '<下划线>False</下划线>
                    vCell.FontStrikeout = .FontStrikeout    '<删除线>False</删除线>
                    vCell.FontColor = .FontColor         '<字体颜色>vbblack</字体颜色>
                    vCell.AlignHorizontal = .HAlignment       '<横向对齐>F1HAlignCenter</横向对齐>
                    vCell.AlignVertical = .VAlignment       '<纵向对齐>F1VAlignCenter</纵向对齐>

                    Select Case .对象类型
                        Case cprCTFixtext    '0-固定文本(不可编辑)
                            F1Main.TextRC(lngRow, lngCol) = .内容文本
                        Case cprCTText '1-文本型(可编辑多行文本)
                            F1Main.TextRC(lngRow, lngCol) = .内容文本
                        Case cprCTElement    '2-单要素
                            If Doc.ET = TabET_病历文件定义 Or Doc.ET = TabET_全文示范编辑 Then
                                If .ElementKey <> "" Then
                                    If Doc.Elements("K" & .ElementKey).输入形态 = 1 Then
                                        F1Main.TextRC(lngRow, lngCol) = Doc.Elements("K" & .ElementKey).内容文本
                                    Else
                                        F1Main.TextRC(lngRow, lngCol) = "[" & Doc.Elements("K" & .ElementKey).要素名称 & "]" & Doc.Elements("K" & .ElementKey).要素单位
                                    End If
                                End If
                            Else
                                strShow = ""
                                If .内容文本 = "" Then
                                    If Doc.Elements("K" & .ElementKey).替换域 = 1 Then '自动替换要素
                                        strShow = GetReplaceEleValue(Doc.Elements("K" & .ElementKey).要素名称, Doc.EPRPatiRecInfo.病人ID, Doc.EPRPatiRecInfo.主页ID, Doc.EPRPatiRecInfo.病人来源, Doc.EPRPatiRecInfo.医嘱id)
                                        If strShow = "" And Not Doc.Elements("K" & .ElementKey).自动转文本 Then '没取到值，是否自动转换成文本(空)
                                            strShow = "[" & Doc.Elements("K" & .ElementKey).要素名称 & "]" & Doc.Elements("K" & .ElementKey).要素单位
                                        Else
                                            Doc.Elements("K" & .ElementKey).内容文本 = strShow
                                            .内容文本 = strShow & Doc.Elements("K" & .ElementKey).要素单位
                                            strShow = .内容文本
                                        End If
                                    Else
                                        If Doc.Elements("K" & .ElementKey).输入形态 = 1 And Doc.Elements("K" & .ElementKey).要素类型 <> 2 Then '输入形态=展开
                                            .内容文本 = Doc.Elements("K" & .ElementKey).内容文本 & Doc.Elements("K" & .ElementKey).要素单位
                                            strShow = .内容文本
                                        Else
                                            strShow = "[" & Doc.Elements("K" & .ElementKey).要素名称 & "]" & Doc.Elements("K" & .ElementKey).要素单位
                                        End If
                                    End If
                                    F1Main.TextRC(lngRow, lngCol) = strShow
                                Else
                                    F1Main.TextRC(lngRow, lngCol) = .内容文本
                                End If
                            End If
                        Case cprCTTextElement '3-文本与多要素混合编辑
                            GetTextELement .Key     '跟据Text Element填写F1Main中的单元格及类的内容文本
                        Case cprCTReportPic, cprCTPicture    '5-报告图
                            If Doc.Pictures("K" & .PictureKey).OrigPic.Handle <> 0 Then
                                Call PaintPictureOnTable(.Key)
                            End If
                            F1Main.TextRC(lngRow, lngCol) = IIf(.对象类型 = cprCTPicture, "参考图", "报告图")
                        Case cprCTSign         '6-签名'签名在设计时仅为占位,无实际信息；没有签名时终止版=0；普通签名后审核时不显示，以便再次签名；行控/列控签名后审核时要显示
                            strShow = ""
                            If Doc.ET = TabET_单病历编辑 Or Doc.ET = TabET_单病历审核 Then
                                'mReadOnly 0-正常,1-签名后点修改,2-主界面打开查阅或查阅历次签名版本
                                If .终止版 <> 0 Then
                                    With Doc.Signs("K" & .SignKey)
                                        strShow = .前置文字 & .姓名 & IIf(.显示手签, "，手签：_____________", "")
                                        strShow = strShow & IIf(Trim(.显示时间) = "", "", "，" & Format(.签名时间, .显示时间))
                                    End With
                                Else
                                    strShow = "[签名位]"
                                End If
                            Else
                                strShow = "[签名位]"
                            End If
                            F1Main.TextRC(lngRow, lngCol) = strShow '前置文字 & 姓名 & 显示手签 & 显示时间<>""(format(签名时间,显示时间)
                        Case cprCTRowSign, cprCTColSign '7-行控签名 '8-列控签名
                            strShow = ""
                            If Doc.ET = TabET_单病历编辑 Or Doc.ET = TabET_单病历审核 Then
                                If .终止版 <> 0 Then
                                    With Doc.Signs("K" & .SignKey)
                                        strShow = .前置文字 & .姓名 & IIf(.显示手签, "，手签：_____________", "")
                                        strShow = strShow & IIf(Trim(.显示时间) = "", "", "，" & Format(.签名时间, .显示时间))
                                    End With
                                Else
                                    strShow = "[签名位]"
                                End If
                            Else
                                strShow = "[签名位]"
                            End If
                            F1Main.TextRC(lngRow, lngCol) = strShow '前置文字 & 姓名 & 显示手签 & 显示时间<>""(format(签名时间,显示时间)
                    End Select
                    F1Main.SetCellFormat vCell
                    Call F1Main.SetBorder(-1, .CellLineLeft, .CellLineRight, .CellLineTop, .CellLineBottom, 0, -1, .CellLineLeftColor, .CellLineRightColor, .CellLineTopColor, .CellLineBottomColor)
                End If
            End With
        Next
    End With
    mblnInit = False
End Sub
Private Sub GetTextELement(ByVal strCellKey As String)
'功能：跟据Text Element填写F1Main中的单元格及类的内容文本
Dim i As Long, lCount As Long, strTmp As String, ltCount As Long, leCount As Long, cleTmp As cTabElement
    With Doc.Cells(strCellKey)
        ltCount = UBound(Split(.TextKey, "|")): If ltCount < 0 Then ltCount = 0
        leCount = UBound(Split(.ElementKey, "|")): If leCount < 0 Then leCount = 0
        lCount = ltCount + leCount
        For i = 1 To lCount
            Set cleTmp = .clElement(Doc.Elements, i)
            If cleTmp Is Nothing Then '该次序为文本
                strTmp = strTmp & ToVarchar(.clText(Doc.Texts, i).内容文本, 4000)
            Else
                With Doc.Elements("K" & cleTmp.Key)
                    If .替换域 = 1 And (Doc.ET = TabET_单病历编辑 Or Doc.ET = TabET_单病历审核) Then
                        If Trim(.内容文本) = "" Then
                            If .自动转文本 Then
                                strTmp = strTmp & " " & .要素单位
                            Else
                                strTmp = strTmp & "[" & .要素名称 & "]" & .要素单位
                            End If
                        Else
                            strTmp = strTmp & .内容文本 & .要素单位
                        End If
                    Else
                        If .输入形态 = 0 Then
                            strTmp = strTmp & IIf(Trim(.内容文本) = "", "[" & .要素名称 & "]", .内容文本) & .要素单位
                        Else
                            strTmp = strTmp & .内容文本 & .要素单位
                        End If
                    End If
                End With
            End If
        Next
        .内容文本 = strTmp
        F1Main.TextRC(.Row, .Col) = strTmp
    End With
End Sub
Private Sub PaintPictureOnTable(ByVal strCellKey As String)
'功能:在指定单元格绘图
Dim objTmp As Object, vR As F1Rect, i As Integer, lHheight As Long, lHwidth As Long, lpLeft As Long, lpTop As Long '图片框,区域,固定列高度,固定行宽度,图片框XY坐标
Dim lsRow As Long, leRow As Long, lsCol As Long, leCol As Long '区域起止行列
Dim lsPosX As Long, lsPosY As Long, lpHeight As Long, lpWidth As Long '图片源剪切XY坐标,图片高宽

    If F1Main.ShowColHeading Then lHheight = F1Main.HdrHeight Else lHheight = 0 '固定行高度
    If F1Main.ShowRowHeading Then lHwidth = F1Main.HdrWidth Else lHwidth = 0    '固定列宽度

    With Doc.Cells(strCellKey)
        If .PictureKey = "" Then Exit Sub
        If Doc.Pictures("K" & .PictureKey).OrigPic.Handle = 0 Then Exit Sub
        
        '确定图片框所在区域
        If .Merge Then  'MergeRange数据格式 (左上方)行,列;(右下方)行,列
            lsRow = Split(Split(.MergeRange, ";")(0), ",")(0): leRow = Split(Split(.MergeRange, ";")(1), ",")(0)
            lsCol = Split(Split(.MergeRange, ";")(0), ",")(1): leCol = Split(Split(.MergeRange, ";")(1), ",")(1)
        Else
            lsRow = .Row: leRow = .Row: lsCol = .Col: leCol = .Col
        End If
        Set vR = F1Main.RangeToTwipsEx(lsRow, lsCol, leRow, leCol)
        '确定图片框大小及位置及裁剪坐标
        If vR.Right - lHwidth <= 0 Or vR.Bottom - lHheight <= 0 Then '不在可显示区域
            If ChkControl(PicDy(.Index)) Then
                PicDy(.Index).Visible = False
            End If
            Exit Sub
        ElseIf vR.Left >= 0 And vR.Top >= 0 Then '区域处在表格中间
            lpLeft = F1Main.Left + vR.Left: lpTop = F1Main.Top + vR.Top: lpWidth = vR.Width: lpHeight = vR.Height: lsPosX = 0: lsPosY = 0
        ElseIf vR.Left >= 0 And vR.Top < 0 Then '区域上方部份隐藏(滚动引起)
            lpLeft = F1Main.Left + vR.Left: lpTop = F1Main.Top + lHheight: lpWidth = vR.Width: lpHeight = vR.Height + vR.Top - lHheight: lsPosX = 0: lsPosY = vR.Height - lpHeight
        ElseIf vR.Left < 0 And vR.Top >= 0 Then '区域左方部份隐藏(滚动引起)
            lpLeft = F1Main.Left + lHwidth: lpTop = F1Main.Top + vR.Top: lpWidth = vR.Width + vR.Left - lHwidth: lpHeight = vR.Height: lsPosX = vR.Width - lpWidth: lsPosY = 0
        ElseIf vR.Left < 0 And vR.Top < 0 Then '区域上方左方都隐藏(滚动引起)
            lpLeft = F1Main.Left + lHwidth: lpTop = F1Main.Top + lHheight: lpWidth = vR.Width + vR.Left - lHwidth: lpHeight = vR.Height + vR.Top - lHheight: lsPosX = vR.Width - lpWidth: lsPosY = vR.Height - lpHeight
        Else                                    '不在可显示区域
            If ChkControl(PicDy(.Index)) Then
                PicDy(.Index).Visible = False
            End If
            Exit Sub
        End If
        
        '动态加载图片框数组
        If Not ChkControl(PicDy(.Index)) Then
            Load PicDy(.Index)
        End If
        Set objTmp = PicDy(.Index): objTmp.Cls
        objTmp.Tag = .MergeRange & "|" & strCellKey: objTmp.ToolTipText = IIf(.对象类型 = cprCTReportPic, "报告图", "参考图")
        objTmp.AutoRedraw = True: objTmp.BorderStyle = 0
        
        '先晳定图片大小并绘出标记
        LockWindowUpdate Me.hWnd
        objTmp.Width = vR.Width - Screen.TwipsPerPixelX * 2: objTmp.Height = vR.Height - Screen.TwipsPerPixelY * 2
        Set objTmp.Picture = Doc.Pictures("K" & .PictureKey).OrigPic
        objTmp.PaintPicture objTmp.Picture, 0, 0, objTmp.Width, objTmp.Height
        If .PicMarkKey <> "" Then '有标记图先绘标记
            For i = 1 To UBound(Split(.PicMarkKey, "|"))
                ShowPicMark objTmp, Doc.PicMarks("K" & Split(.PicMarkKey, "|")(i))
            Next
        End If
        Set objTmp.Picture = objTmp.Image
        '最后根据实际显示大小及坐标重绘
        objTmp.Move lpLeft + Screen.TwipsPerPixelX * 2, lpTop + Screen.TwipsPerPixelY * 2, lpWidth - Screen.TwipsPerPixelX * 2, lpHeight - Screen.TwipsPerPixelY * 2
        objTmp.PaintPicture objTmp.Picture, 0, 0, objTmp.Width, objTmp.Height, lsPosX, lsPosY
        objTmp.Visible = True: objTmp.ZOrder
        LockWindowUpdate 0
    End With
End Sub
