VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReportEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mDateType As Byte          '0表示出生日期，1表示发病日期，2表示诊断日期，3表示死亡日期,4表示填卡日期
Public mfrmReportEx As Object      '当前的报告卡窗体对象
Private marrSql() As Variant        '保存数据时候执行的SQL
Private mColCls As New Collection   '需要保存到数据库的数据
Private mColData As New Collection  '保存从数据库读取到的数据
Public blnHaveStatus As Boolean  '是否存在状态栏
Private blnFirstGot As Boolean  '第一次获得焦点

Private mlngPatiID As Long '病人id
Private mlngPageID As Long '主页ID（门诊传挂号ID）
Private mbytType As Byte   '编辑方式0-新增　1-修改，用于区别提取数据
Private mbytFrom As Byte   '病人来源1-门诊 2-住院
Private mlngDeptID As Long '当前科室ID
Private mlngFileID As Long   '文件ID,来源于电子病历记录.ID
Private mbytBabyNo As Long '婴儿ID
Private mbln身份证必填 As Boolean '身份证信息必填 参数：传染病报告身份证号码必填
Private mblnFirst As Boolean
Private mblnOneCard As Boolean '一病一卡 参数：传染病报告卡一病一卡

Private mstrChkType_2016 As String '数据格式是："[男][艾滋病][AIDS][...]......"
Private mstrChkType_2014 As String '数据格式是："[男][艾滋病][AIDS][...]......"
Private mstrVL2014 As String
Private mstrVL2016 As String
Private mcolLoadData As Collection


Public Event ClickPositives(blnSelected As Boolean)  '选择了阳性检测结果时触发

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Public Sub ClearMe()
    Dim objCtl As Control
    
    On Error GoTo errHand
    For Each objCtl In mfrmReportEx.Controls
        Call ClearInfo(objCtl)
    Next
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Public Sub ucCaseType2Change(Index As Integer)
    gbytAcute = 3
    If mfrmReportEx.ucCaseType2(Index).Checked = True Then
        gbytAcute = Index
    End If
End Sub

Public Sub ucCaseType1Change(Index As Integer)
    gbytDiseaseType = 5
    If mfrmReportEx.ucCaseType1(Index).Checked = True Then
        gbytDiseaseType = Index
    End If
    If Index = 4 Then
        RaiseEvent ClickPositives(mfrmReportEx.ucCaseType1(4).Checked)
    End If
End Sub

Public Sub txtIDCardKeyPress(Index As Integer, KeyAscii As Integer)
    If zlStr.IsNumOrChar(Chr(KeyAscii)) Then
        SendKeys "{TAB}"
    ElseIf Not (KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack) Then
        KeyAscii = 0
    End If
End Sub

Public Sub txtIDCardKeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyLeft Then
        SendKeys "+{TAB}"
    ElseIf KeyCode = vbKeyRight Then
        SendKeys "{TAB}"
    Else
        mfrmReportEx.txtIDCard(Index).SelStart = 0
        mfrmReportEx.txtIDCard(Index).SelLength = Len(mfrmReportEx.txtIDCard(Index).Text)
    End If
End Sub

Public Sub txtDiagnoseKeyPress(Index As Integer, KeyAscii As Integer)
    Call CheckVal(KeyAscii)
End Sub

Public Sub txtDeathKeyPress(Index As Integer, KeyAscii As Integer)
    Call CheckVal(KeyAscii)
End Sub

Public Sub txtBirthKeyPress(Index As Integer, KeyAscii As Integer)
    Call CheckVal(KeyAscii)
    mfrmReportEx.txtAge.Text = ""
End Sub

Public Sub txtAttackKeyPress(Index As Integer, KeyAscii As Integer)
    Call CheckVal(KeyAscii)
End Sub

Public Sub txtAgeKeyPress(KeyAscii As Integer)
    Call CheckVal(KeyAscii)
End Sub

Public Sub txtAgeChange()
    If mblnFirst = False Then
        mfrmReportEx.ucAge(1).Checked = IIf(Trim(mfrmReportEx.txtAge.Text) = "", False, True)
    End If
End Sub

Public Sub lblAttackClick(Index As Integer)
    mDateType = IIf(Index > 2, 4, 1)
    With mfrmReportEx
        Call ShowMView(.lblAttack(Index).Container.Left + 0.0658 * (.lblAttack(Index).Left), .lblAttack(Index).Container.Top + 0.07188 * (.lblAttack(Index).Top))
    End With
End Sub

Public Sub lblDeathClick(Index As Integer)
    mDateType = 3
    With mfrmReportEx
        Call ShowMView(.lblDeath(Index).Container.Left + 0.0658 * (.lblDeath(Index).Left), .lblDeath(Index).Container.Top + 0.07188 * (.lblDeath(Index).Top))
    End With
End Sub

Public Sub lblBirthClick(Index As Integer)
    mDateType = 0
    With mfrmReportEx
        Call ShowMView(.lblBirth(Index).Container.Left + 0.0658 * (.lblBirth(Index).Left), .lblBirth(Index).Container.Top + 0.07188 * (.lblBirth(Index).Top))
    End With
End Sub

Public Sub lblDiagnoseClick(Index As Integer)
    If Index <> 3 Then
        mDateType = 2
        With mfrmReportEx
            Call ShowMView(.lblDiagnose(Index).Container.Left + 0.0658 * (.lblDiagnose(Index).Left), .lblDiagnose(Index).Container.Top + 0.07188 * (.lblDiagnose(Index).Top))
        End With
    End If
End Sub

Public Sub lblAttackMouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set mfrmReportEx.lblAttack(Index).MouseIcon = mfrmReportEx.Image1.Picture
    mfrmReportEx.lblAttack(Index).MousePointer = vbCustom
End Sub

Public Sub lblDeathMouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set mfrmReportEx.lblDeath(Index).MouseIcon = mfrmReportEx.Image1.Picture
    mfrmReportEx.lblDeath(Index).MousePointer = vbCustom
End Sub

Public Sub lblBirthMouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set mfrmReportEx.lblBirth(Index).MouseIcon = mfrmReportEx.Image1.Picture
    mfrmReportEx.lblBirth(Index).MousePointer = vbCustom
End Sub

Public Sub lblDiagnoseMouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index <> 3 Then
        Set mfrmReportEx.lblDiagnose(Index).MouseIcon = mfrmReportEx.Image1.Picture
        mfrmReportEx.lblDiagnose(Index).MousePointer = vbCustom
    End If
End Sub

Public Sub MViewDateClick(ByVal DateClicked As Date)
    With mfrmReportEx
        .MView.Visible = False
        Select Case mDateType
            Case 0
                .txtBirth(0).Text = .MView.Year
                .txtBirth(1).Text = .MView.Month
                .txtBirth(2).Text = .MView.Day
            Case 1
                .txtAttack(0).Text = .MView.Year
                .txtAttack(1).Text = .MView.Month
                .txtAttack(2).Text = .MView.Day
            Case 2
                .txtDiagnose(0).Text = .MView.Year
                .txtDiagnose(1).Text = .MView.Month
                .txtDiagnose(2).Text = .MView.Day
            Case 3
                .txtDeath(0).Text = .MView.Year
                .txtDeath(1).Text = .MView.Month
                .txtDeath(2).Text = .MView.Day
            Case 4
                .txtEnter(0).Text = .MView.Year
                .txtEnter(1).Text = .MView.Month
                .txtEnter(2).Text = .MView.Day
        End Select
    End With
End Sub

Public Sub MViewLostFocus()
    mfrmReportEx.MView.Visible = False
End Sub

Public Sub ShowMView(X As Long, Y As Long)
'功能：显示日期选择控件
    With mfrmReportEx
        .MView.Left = X
        If mDateType = 0 Then
            .MView.Top = Y + 14.376
        ElseIf mDateType = 3 Then
            .MView.Top = Y - .MView.Height
            .MView.Left = X
        Else
            .MView.Top = Y - .MView.Height
        End If
        .MView.Visible = True
        Call .MView.SetFocus
    End With
End Sub


Public Sub FormLoad()
    blnFirstGot = True
    gbytDiseaseType = 5
    gbytAcute = 3
    With mfrmReportEx
        .picReport.ScaleHeight = .ScaleY(297, 6, 3)
        .picReport.ScaleWidth = .ScaleX(210, 6, 3)
        .picReport.Top = .ScaleTop + 200
        marrSql = Array()
        .Subclass.hWnd = .hWnd
        .Subclass.Messages(WM_MOUSEWHEEL) = True
        mbln身份证必填 = Val(zlDatabase.GetPara("传染病报告身份证号码必填", glngSys, 1277, 0)) = 1
	mblnOneCard = Val(zlDatabase.GetPara("传染病报告卡一病一卡", glngSys, 1277, 0)) = 1
        If mbln身份证必填 Then
            .lblReport(7).Caption = "有效证件号*："
        Else
            .lblReport(7).Caption = "有效证件号："
        End If
    End With
End Sub

Public Sub FormResize()
    On Error Resume Next
    With mfrmReportEx
        .picReport.Left = .ScaleLeft + (.ScaleWidth / 2) - (.picReport.Width / 2)
        If .ScaleWidth < .picReport.Width Then
            .hsbReport.Visible = True
        Else
            .hsbReport.Visible = False
        End If
        .vsbReport.Top = .ScaleTop
        .vsbReport.Left = .ScaleLeft + .ScaleWidth - .vsbReport.Width
        .vsbReport.Height = .ScaleHeight - IIf(.hsbReport.Visible = True, .hsbReport.Height, 0) - IIf(blnHaveStatus = True, 375, 0)
        .vsbReport.LargeChange = 100 / ((.picReport.Height + 800) / .ScaleHeight)
        .vsbReport.SmallChange = .vsbReport.LargeChange
        .hsbReport.Top = .vsbReport.Top + .vsbReport.Height
        .hsbReport.Left = .ScaleLeft
        .hsbReport.Width = .ScaleLeft + .ScaleWidth
        .hsbReport.LargeChange = 100 / (.picReport.Width / .ScaleWidth)
        .hsbReport.SmallChange = .hsbReport.LargeChange
        .picShadow.Move .picReport.Left + 50, .picReport.Top + 50, .picReport.Width, .picReport.Height
    End With
End Sub

Public Sub FormUnload()
    Set mColCls = Nothing
    Set mColData = Nothing
    
    Erase marrSql
    gstrKey = ""
    mfrmReportEx.Subclass.Messages(WM_MOUSEWHEEL) = False
End Sub

Public Sub hsbReportChange()
    With mfrmReportEx
        .picReport.Left = -((.picReport.Width - .Width) * (.hsbReport.Value / 100))
        .picShadow.Left = .picReport.Left + 50
    End With
End Sub

Public Sub picReportGotFocus()
    If blnFirstGot = True And mfrmReportEx.picReport.Enabled = True Then
        Call mfrmReportEx.txtNumber.SetFocus
        mfrmReportEx.txtNumber.SelStart = Len(mfrmReportEx.txtNumber.Text)
    End If
    blnFirstGot = False
End Sub

Public Sub SubclassWndProc(Msg As Long, wParam As Long, lParam As Long, result As Long)
    '自定义的消息处理函数
    Dim tP As POINTAPI
    Dim sngX As Single, sngY As Single   '鼠标坐标
    Dim intShift As Integer              '鼠标按键
    Dim bWay As Boolean                  '鼠标方向
    Dim bMouseFlag As Boolean            '鼠标事件激活标志
    Dim wzDelta, wKeys As Integer
    With mfrmReportEx
        Select Case Msg
            Case WM_MOUSEWHEEL   '滚动
                wzDelta = HIWORD(wParam)
                If wzDelta > 0 Then
                    .vsbReport.Value = IIf(.vsbReport.Value > 10, .vsbReport.Value - 10, 0)
                Else
                    .vsbReport.Value = IIf(.vsbReport.Value < 90, .vsbReport.Value + 10, 100)
                End If
        End Select
    End With
End Sub

Public Sub vsbReportChange()
    With mfrmReportEx
        .picReport.Top = 200 - ((.picReport.Height + 800 - .Height) * (.vsbReport.Value / 100))
        .picShadow.Top = .picReport.Top + 50
    End With
End Sub

Public Sub SetMyFocus()
    If mfrmReportEx.picReport.Enabled = True Then
        Call mfrmReportEx.picReport.SetFocus
    End If
End Sub

Public Sub CanWrite()
    Dim obj As Object
'功能：是界面可以编辑
    mfrmReportEx.picReport.Enabled = True
    If gblnLock = True Then
        mfrmReportEx.picPane(2).Enabled = (Not gblnLock)
        mfrmReportEx.picPane(2).BackColor = &H8000000F
        For Each obj In mfrmReportEx.Controls
            If TypeName(obj) = "Label" Or TypeName(obj) = "uCheckNorm" Then
                If obj.Container.Name = "picPane" Then
                    If obj.Container.Index = 2 Then
                        obj.BackColor = &H8000000F
                    End If
                End If
            End If
        Next
    End If
End Sub


Public Sub PrintReport(ByVal frmParent As Object, ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal lngFileId As Long, ByVal strPrintDeviceName As String)
'功能：打印报告
    Dim strSQL As String
    Dim strPos As String
    Dim strPosInfo() As String
    Dim strPosTmp() As String
    Dim i As Integer, j As Long, m As Long
    Dim objPage As Object

    On Error GoTo errHand

    Call zlRefresh(lngPatiID, lngPageID, lngFileId, False)
    
    For j = 0 To Val(mfrmReportEx.Tag)
        If j > 0 Then
            For m = 0 To mfrmReportEx.picPane.UBound
                If (mfrmReportEx.picPane(m).Tag) = j Then
                    Set objPage = mfrmReportEx.picPane(m)
                End If
            Next
            If objPage Is Nothing Then Exit For
        End If
    
        If Trim(strPrintDeviceName) <> "" Then
            For i = 0 To Printers.Count - 1
                If Trim(Printers(i).DeviceName) = Trim(strPrintDeviceName) Then
                    Set Printer = Printers(i)
                    Exit For
                End If
                If i = Printers.Count - 1 Then
                    MsgBox "没有找到相应的打印机，请核对打印机名称！", vbInformation + vbOKOnly, gstrSysName
                    Exit Sub
                End If
            Next
        End If
        Printer.PaperSize = vbPRPSA4 'A4纸
        Printer.ScaleMode = vbPixels
    
        glngOffsetX = -GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX) '可打印左边缘
        glngOffsetY = -GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY) '可打印上边缘
            
        strPos = getStrPos(j, objPage)
    
        Call PrintCtl(j)
        
        If strPos <> "" Then
            strPosInfo = Split(strPos, "|")
            For i = 0 To UBound(strPosInfo)
                strPosTmp = (Split(strPosInfo(i), ","))
                Printer.Line (glngOffsetX + PScaleX(Val(strPosTmp(0))), glngOffsetY + PScaleY(Val(strPosTmp(1))))-(glngOffsetX + PScaleX(Val(strPosTmp(2))), glngOffsetY + PScaleY(Val(strPosTmp(3)))), &H0&, B
            Next
        End If
    
        Printer.EndDoc
    Next

    strSQL = "Zl_电子病历打印_Insert(" & mlngFileID & ",20," & mlngPatiID & "," & mlngPageID & ",'" & UserInfo.姓名 & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "")

    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub


Public Function CheckContainer(obj As Object) As Boolean
'检查当前控件是否有父控件
    Dim strTmp As String
    On Error GoTo errHand
    strTmp = obj.Container.Name
    CheckContainer = True
    Exit Function
errHand:
    CheckContainer = False
    Err.Clear
End Function


Public Sub PrintCtl(lngPage As Long)
    '打印控件
    Dim objCtl As Control
    Dim lngTop As Long
    Dim dblTopZ As Double
    
    For Each objCtl In mfrmReportEx.Controls
        If CheckContainer(objCtl) = True Then
            If objCtl.Container.Name = "picPane" And Val(objCtl.Container.Tag) = lngPage Then
                '判断picPane里的控件Tag值是否为0,不为0则生成打印位置
                If objCtl.Tag <> "0" Then
                lngTop = Decode(objCtl.Container.Index, 0, 88, 1, gLngTwo, 2, gLngThree, 3, gLngFour, 52)
                dblTopZ = Decode(objCtl.Container.Index, 0, 0.05866, 1, 0.062, 2, 0.06315, 3, 0.06376, 0.05866)
                
                    Select Case TypeName(objCtl)
                            Case "Line"
                                If objCtl.Name = "LineNew" Then
                                    objCtl.Tag = "69," & (lngTop + dblTopZ * objCtl.Y1) & ",725"
                                Else
                                    objCtl.Tag = (70.23 + 0.0645 * objCtl.X1) & "," & (lngTop + dblTopZ * objCtl.Y1) & "," & (70.23 + 0.0645 * objCtl.X2)
                                End If
                            Case "TextBox"
                                objCtl.Tag = (70.23 + 0.0645 * objCtl.Left) & "," & (lngTop + dblTopZ * objCtl.Top)
                                If (objCtl.Width - mfrmReportEx.TextWidth(objCtl.Text)) / 2 > 0 And (objCtl.Width - mfrmReportEx.TextWidth(objCtl.Text)) / 2 < objCtl.Width / 2 And objCtl.Name <> "txtImportant" And objCtl.Name <> "txtRemarks" Then
                                    objCtl.Tag = (70.23 + 0.0645 * (objCtl.Left + ((objCtl.Width - mfrmReportEx.TextWidth(objCtl.Text)) / 2))) & "," & (lngTop + dblTopZ * objCtl.Top)
                                End If
                            Case Else
                                objCtl.Tag = (70.23 + 0.0645 * objCtl.Left) & "," & (lngTop + dblTopZ * objCtl.Top)
                    End Select
                    Call PrintInfo(objCtl)
                End If
            End If
        End If
    Next
End Sub


Public Function getStrPos(lngPage As Long, objPage As Object)
    Dim TOP1, Top2, Top3, Top4 As Long
    Dim lngX1 As Long
    Dim lngX2 As Long
    Dim strPos As String
    With mfrmReportEx
        lngX1 = 69: lngX2 = 725
    
        If lngPage = 0 Then
            '黑线框两个Y值
            TOP1 = 167.504111018364
            '两个Line控件的Y值
            Top3 = TOP1 + ((.picPane(1).ScaleHeight / 13.9107 / 0.07188) * 0.063) + 10
            Top4 = Top3 + ((.picPane(2).ScaleHeight / 13.9107 / 0.07188) * 0.063) + 10
            Top2 = Top4 + ((.picPane(3).ScaleHeight / 13.9107 / 0.07188) * 0.063) + 10
        
            gLngTwo = TOP1 + 2
            gLngThree = Top3 + 2
            gLngFour = Top4 + 2
        
            strPos = lngX1 & "," & TOP1 & "," & lngX2 & "," & TOP1 & "|" & _
                    lngX1 & "," & Top2 & "," & lngX2 & "," & Top2 & "|" & _
                    lngX1 & "," & TOP1 & "," & lngX1 & "," & Top2 & "|" & _
                    lngX2 & "," & TOP1 & "," & lngX2 & "," & Top2 & "|" & _
                    lngX1 & "," & Top3 & "," & lngX2 & "," & Top3 & "|" & _
                    lngX1 & "," & Top4 & "," & lngX2 & "," & Top4
        Else
            '打印多页
            If Not objPage Is Nothing Then
                If objPage.BorderStyle <> 0 Then
                    TOP1 = 50
                    Top2 = TOP1 + ((objPage.ScaleHeight / 13.9107 / 0.07188) * 0.063) + 10
                    strPos = lngX1 & "," & TOP1 & "," & lngX2 & "," & TOP1 & "|" & _
                            lngX1 & "," & Top2 & "," & lngX2 & "," & Top2 & "|" & _
                            lngX1 & "," & TOP1 & "," & lngX1 & "," & Top2 & "|" & _
                            lngX2 & "," & TOP1 & "," & lngX2 & "," & Top2
                End If
            End If
        End If
    End With
        getStrPos = strPos
End Function

Public Sub zlRefresh(ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal lngFileId As Long, ByVal blnMoved As Boolean)
    mlngPatiID = lngPatiID
    mlngPageID = lngPageID
    mlngFileID = lngFileId

    Call ClearMe
    Call InitReport(mbytType, mlngPatiID, mlngPageID, mbytFrom, 0, mlngDeptID, mlngFileID)
    If lngPatiID <> 0 Then
        Call LoadData(1, blnMoved)
    End If
End Sub



Public Sub InitReport(ByVal bytType As Byte, ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal bytFrom As Byte, ByVal bytBabyNo As Byte, ByVal lngDeptID As Long, ByVal lngFileId As Long)
    mbytType = bytType
    mlngPatiID = lngPatiID
    mlngPageID = lngPageID
    mbytFrom = bytFrom
    mlngDeptID = lngDeptID
    mlngFileID = lngFileId
    mbytBabyNo = bytBabyNo
End Sub

Public Function SaveData(ByVal blnFinish As Boolean) As Boolean
    Dim i As Integer
    Dim strSQL As String
    Dim blnBegin As Boolean
    Dim SLevel As SignLevel
    Dim lngFileId As Long       '文件ID 来源于病历文件列表
    Dim strFileName As String   '文件名称 来源于病历文件列表
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHand

    SaveData = False

    '新增需要提取新的文件ID
    If mbytType = 0 Then
        mlngFileID = zlDatabase.GetNextId("电子病历记录")
        mbytType = 1
    End If

    SLevel = GetUserSignLevel(UserInfo.ID, mlngPatiID, mlngPageID)

    strSQL = "select t.id,t.名称 from 病历文件列表 t where t.种类=5 and t.编号='000'"
    Set rsTemp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, "")
    lngFileId = Nvl(rsTemp!ID, 0)
    strFileName = Nvl(rsTemp!名称, "")
    strSQL = "Zl_传染病报告卡记录_Update(" & mlngFileID & "," & mbytFrom & "," & mlngPatiID & "," & _
              mlngPageID & "," & mlngDeptID & ",'" & UserInfo.姓名 & "'," & lngFileId & ",'" & strFileName & _
               "','" & UserInfo.姓名 & "'," & IIf(blnFinish, 1, 0) & "," & IIf(blnFinish, SLevel, "Null") & "," & mbytBabyNo & ")"

    Call MakeSaveSql(marrSql, mColCls, mlngFileID)

    gcnOracle.BeginTrans
    blnBegin = True
    Call zlDatabase.ExecuteProcedure(strSQL, "")
    For i = LBound(marrSql) To UBound(marrSql)
        If Not ((InStr(CStr(marrSql(i)), ",'填卡日期',") > 0 Or InStr(CStr(marrSql(i)), ",'填卡医生',") > 0) And gblnLock) Then
            Call zlDatabase.ExecuteProcedure(CStr(marrSql(i)), "")
        End If
        Call frmMain.HaveSavedSQL
    Next
    
    If mfrmReportEx.SaveData(mlngFileID, mlngPatiID, mlngPageID, mbytType, mbytFrom, mlngDeptID, mbytBabyNo) = False Then
        If blnBegin Then
            gcnOracle.RollbackTrans
        End If
        Exit Function
    End If
    
    gcnOracle.CommitTrans
    blnBegin = False
    SaveData = True
    If blnFinish = True Then
        mfrmReportEx.picReport.Enabled = False
    End If
    Exit Function
errHand:
    If blnBegin Then
        gcnOracle.RollbackTrans
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Function

Public Sub LoadData(ByVal bytType As Byte, Optional blnMoved As Boolean)
    Dim strSQL As String
    Dim strKey As String
    Dim strNo As String
    Dim strID As String
    Dim strTmp As String
    Dim strInfo() As String
    Dim objCls As clsReport
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer

    On Error GoTo errHand
    Set mColCls = New Collection
    mstrChkType_2014 = ""
    mstrChkType_2016 = ""
'   1-修改
    If bytType = 1 Then
        Set mColData = New Collection
        strSQL = "select t.id,t.对象序号,t.内容文本,t.要素名称 from 电子病历内容 t where t.文件id=[1]"
        If blnMoved = True Then
            strSQL = Replace(strSQL, "电子病历内容", "H电子病历内容")
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "电子病历内容", mlngFileID)

        For i = 0 To rsTemp.RecordCount - 1
            If rsTemp.EOF = False Then
                strID = Nvl(rsTemp!ID)
                strNo = Nvl(rsTemp!对象序号)
                strTmp = Nvl(rsTemp!内容文本)
                strKey = "K" & Trim(strNo)
                mColData.Add strTmp, strKey

                If InStr(GSTR_OBJNO_2014, "," & strNo & ",") > 0 Then
                    mstrChkType_2014 = mstrChkType_2014 & "[" & strNo & "," & Trim(strTmp) & "]"
                End If

                If InStr(GSTR_OBJNO_2016, "," & strNo & ",") > 0 Then
                    mstrChkType_2016 = mstrChkType_2016 & "[" & strNo & "," & Trim(strTmp) & "]"
                End If

                Set objCls = New clsReport
                objCls.ID = strID
                mColCls.Add objCls, strKey
                rsTemp.MoveNext
            End If
        Next
        If mColCls.Count = 38 Then
            Set objCls = New clsReport
            strKey = "K" & 39
            objCls.ID = 0
            mColCls.Add objCls, strKey
        ElseIf mColCls.Count = 44 Then
            Set objCls = New clsReport
            strKey = "K" & 45
            objCls.ID = 0
            mColCls.Add objCls, strKey
        End If
        
'   0-新增
    ElseIf bytType = 0 Then
        For i = 1 To 45
            Set objCls = New clsReport
            strKey = "K" & i
            objCls.ID = 0
            mColCls.Add objCls, strKey
        Next
        Set mColData = New Collection
        strTmp = "姓名|身份证号|性别|出生日期|年龄|工作单位|联系人电话|家庭电话|单位电话|婚姻状况|学历|单位名称|当前日期|家庭地址"
        strInfo = Split(strTmp, "|")

        For i = 0 To UBound(strInfo)
            If mbytBabyNo <> 0 And Trim(strInfo(i)) = "姓名" Then
                strSQL = "select Zl_Replace_Element_Value([1],[2],[3],[4],null,[5]) as 信息 from dual"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "数据读取", strInfo(i), mlngPatiID, mlngPageID, mbytFrom, mbytBabyNo)
                strTmp = Nvl(rsTemp!信息)
            ElseIf Trim(strInfo(i)) = "单位名称" Then
                strTmp = Nvl(zlRegInfo("单位名称"))
            Else
                strSQL = "select Zl_Replace_Element_Value([1],[2],[3],[4]) as 信息 from dual"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "数据读取", strInfo(i), mlngPatiID, mlngPageID, mbytFrom)
                strTmp = Nvl(rsTemp!信息)
            End If
            strNo = i
            mColData.Add strTmp, "K" & Trim(strNo)
        Next
        '家长姓名
        If mbytBabyNo <> 0 Then
            strSQL = "select Zl_Replace_Element_Value([1],[2],[3],[4],null,[5]) as 信息 from dual"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "数据读取", "家长姓名", mlngPatiID, mlngPageID, mbytFrom, mbytBabyNo)
            strTmp = Nvl(rsTemp!信息)
            mColData.Add strTmp, "KParent"
        Else
            mColData.Add "", "KParent"
        End If
        '发病日期
        If mbytFrom = 1 Then
            strSQL = "select t.登记时间 as 发病日期 from 病人挂号记录 t where t.id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "数据读取", mlngPageID)
        Else
            strSQL = "select t.入院日期 as 发病日期 from 病案主页 t where t.病人id=[1] and t.主页id=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "数据读取", mlngPatiID, mlngPageID)
        End If

        If rsTemp.RecordCount <> 0 Then
            mColData.Add Format(Nvl(rsTemp!发病日期), "yyyy-mm-dd"), "K14"
        Else
            mColData.Add "--", "K14"
        End If
        '诊断日期
        strSQL = "select decode(t.发病时间,null,t.记录日期,t.发病时间) as 诊断日期 from 病人诊断记录 t where t.病人id=[1] and t.主页id=[2] And NVL(T.编码序号,1) = 1"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "数据读取", mlngPatiID, mlngPageID)

        If rsTemp.RecordCount <> 0 Then
            mColData.Add Format(Nvl(rsTemp!诊断日期), "yyyy-mm-dd-hh"), "K15"
        Else
            mColData.Add "---", "K15"
        End If
        '死亡日期
        strSQL = " Select a.开始执行时间 as 死亡日期 " & _
                 " From 病人医嘱记录 A, 诊疗项目目录 B " & _
                 " Where a.诊疗项目id = b.Id And b.类别 = 'Z' And " & _
                 " b.操作类型 = '11'  And a.病人来源 = [1] And a.病人id=[2] and a.主页id=[3] "

        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "数据读取", mbytFrom, mlngPatiID, mlngPageID)
        If rsTemp.RecordCount <> 0 Then
            mColData.Add Format(Nvl(rsTemp!死亡日期), "yyyy-mm-dd"), "K17"
        Else
            mColData.Add "--", "K17"
        End If
        '病种
        strSQL = "Select a.Id, b.文件id, b.报告病种, a.病人id, a.主页id, a.医嘱id, a.诊断类型, a.疾病id, a.诊断id, a.诊断描述" & _
                 " From 病人诊断记录 A, 疾病报告前提 B " & _
                 " Where (a.疾病id = b.疾病id Or " & _
                 " a.诊断id = b.诊断id Or " & _
                 " b.诊断id = (Select c.诊断id From 疾病诊断对照 c Where c.疾病id =a.疾病id) or " & _
                 " b.疾病id = (select d.疾病id from 疾病诊断对照 d where d.诊断id=a.诊断id)) And " & _
                 " b.文件id =(select e.id from 病历文件列表 e where e.种类=5  and e.名称='中华人民共和国传染病报告卡' and e.保留=4 ) and " & _
                 " a.记录来源=3 And NVL(A.编码序号,1) = 1 and a.病人id=[1] and a.主页id=[2]"

        strTmp = ""
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "数据读取", mlngPatiID, mlngPageID)
        For i = 0 To rsTemp.RecordCount - 1
            If rsTemp.EOF = False Then
                strTmp = strTmp & Nvl(rsTemp!报告病种) & "|"
                rsTemp.MoveNext
            End If
        Next
        mColData.Add strTmp, "K16"
    End If
    '修改时候加载数据是44-45条(2016版本38-39条)如果少于44-45条(2016版本38-39条)说明病历文件不完整
    '新增时候加载数据是19条，如果少于19条说明信息来源破坏

    If bytType = 1 And (mColData.Count = 44 Or mColData.Count = 45) Then
        glngVersion = VL_2014                                       '修改时候加载数据是45条,2014版本
    ElseIf bytType = 1 And (mColData.Count = 38 Or mColData.Count = 39) Then
        glngVersion = VL_2016                                  '修改时候加载数据是39条,2016版本
    ElseIf mColData.Count = 19 And bytType = 0 Then
        glngVersion = VL_2016
    End If

    If glngVersion = VL_2014 Then
        Call SetData(mColData, bytType, mstrChkType_2014)
    ElseIf glngVersion = VL_2016 Then
        Call SetData(mColData, bytType, mstrChkType_2016)
    End If

    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub



Public Sub SetData(colData As Collection, bytType As Byte, ByVal strChkType As String)
    Dim strTmp As String
    Dim i As Integer
    Dim strInfo() As String
    Dim objCtl As Control
    Dim dteTmp As Date
    Dim strName As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim blnStructAdress As Boolean, blnShowTown As Boolean
    
    
    On Error GoTo errHand
    mblnFirst = True
    
    With mfrmReportEx
        If bytType = 1 Then  '修改
            .txtNumber.Text = CStr(colData("K1"))       '卡片编号
            .txtName.Text = CStr(colData("K3"))          '姓名
            .txtParentName.Text = CStr(colData("K4"))    '家长姓名
            
            '身份证号
            strTmp = CStr(colData("K5"))
            For i = 0 To .txtIDCard.Count - 1
                .txtIDCard(i).Text = Mid(strTmp, i + 1, 1)
            Next
            
            '性别，年龄单位，家人属于，人群分类，病例分类
            For Each objCtl In .Controls
                If TypeName(objCtl) = "uCheckNorm" Then
                    strTmp = Trim(objCtl.Caption)
                    strTmp = Replace(strTmp, "(", "")
                    strTmp = Replace(strTmp, ")", "")
                    strTmp = Replace(strTmp, "、", "")
                    
                    Select Case objCtl.Name
                        Case "ucCheckType" '报卡类别
                            If InStr(strChkType, strTmp) > 0 And Trim(strTmp) <> "" Then
                                objCtl.Checked = True
                            End If
                        Case "ucCheckJob" '人群分类
                            If InStr(strChkType, "14," & strTmp) > 0 And Trim(strTmp) <> "" Then
                              objCtl.Checked = True
                            End If
                        Case "ucCaseType1" '病例分类1
                            If InStr(strChkType, "15," & strTmp) > 0 And Trim(strTmp) <> "" Then
                                objCtl.Checked = True
                                If strTmp = "疑似病例" Then
                                    gbytDiseaseType = 0
                                ElseIf strTmp = "临床诊断病例" Then
                                    gbytDiseaseType = 1
                                ElseIf strTmp = "确诊病例" Then
                                    gbytDiseaseType = 2
                                ElseIf strTmp = "病原携带者" Then
                                    gbytDiseaseType = 3
                                End If
                            End If
                         Case "ucCaseType2"  '病例分类2
                            If InStr(strChkType, "16," & strTmp) > 0 And Trim(strTmp) <> "" Then
                                objCtl.Checked = True
                                If strTmp = "急性" Then
                                    gbytAcute = 0
                                ElseIf strTmp = "慢性" Then
                                    gbytAcute = 1
                                ElseIf strTmp = "未分型" Then
                                    gbytAcute = 2
                                End If
                            End If
                        Case "ucAIDS"           '艾滋病
                            If InStr(strChkType, "22," & strTmp) > 0 And Trim(strTmp) <> "" Then
                                objCtl.Checked = True
                            End If
                        Case "ucHepatitis"        '病毒性肝炎
                            If InStr(strChkType, "23," & strTmp) > 0 And Trim(strTmp) <> "" Then
                                objCtl.Checked = True
                            End If
                        Case "ucAnthrax"          '炭疽
                            If InStr(strChkType, "24," & strTmp) > 0 And Trim(strTmp) <> "" Then
                                objCtl.Checked = True
                            End If
                        Case "ucDysentery"        '痢疾
                            If InStr(strChkType, "25," & strTmp) > 0 And Trim(strTmp) <> "" Then
                                objCtl.Checked = True
                            End If
                        Case "ucPTB"              '肺结核
                            If InStr(strChkType, "26," & strTmp) > 0 And Trim(strTmp) <> "" Then
                                objCtl.Checked = True
                            End If
                        Case "ucTyphia"          '伤寒
                            If InStr(strChkType, "27," & strTmp) > 0 And Trim(strTmp) <> "" Then
                                objCtl.Checked = True
                            End If
                        Case "ucSyphilis"         '梅毒
                            If InStr(strChkType, "28," & strTmp) > 0 And Trim(strTmp) <> "" Then
                                objCtl.Checked = True
                            End If
                        Case "ucMalaria"          '疟疾
                            If InStr(strChkType, "29," & strTmp) > 0 And Trim(strTmp) <> "" Then
                                objCtl.Checked = True
                            End If
                        Case "ucInfectiousA"     '甲类传染病
                            If InStr(strChkType, "20," & strTmp) > 0 And Trim(strTmp) <> "" Then
                                objCtl.Checked = True
                            End If
                        Case Else
                            If InStr(GSTR_Controls, objCtl.Name) = 0 Then
                                If InStr(strChkType, objCtl.Name & ":" & strTmp) > 0 And Trim(strTmp) <> "" Then
                                    objCtl.Checked = True
                                End If
                            Else
                                If InStr(strChkType, strTmp) > 0 And Trim(strTmp) <> "" Then
                                    objCtl.Checked = True
                                End If
                            End If
                    End Select
                End If
            Next
            
            '出生日期
            strTmp = CStr(colData("K7"))
            strInfo = Split(strTmp, "-")
            For i = 0 To UBound(strInfo)
                .txtBirth(i).Text = IIf(Val(strInfo(i)) = 0, "", strInfo(i))
            Next
            
            .txtAge.Text = CStr(colData("K8"))           '年龄
            .txtAddress.Text = CStr(colData("K10"))      '地址
            .txtPhone.Text = CStr(colData("K11"))       '电话
            
            '住址
            strTmp = CStr(colData("K13"))
            strInfo = Split(strTmp, ";")
            For i = 0 To UBound(strInfo) - 1
                .txtAddInfo(i).Text = strInfo(i)
            Next
            
            '发病日期
            strInfo = Split(CStr(colData("K17")), "-")
            For i = 0 To UBound(strInfo)
                .txtAttack(i) = IIf(Val(strInfo(i)) = 0, "", strInfo(i))
            Next
            
            '诊断时间
            strInfo = Split(CStr(colData("K18")), " ")
            If UBound(strInfo) > 0 Then
                .txtDiagnose(3) = IIf(Val(strInfo(1)) = 0, "", strInfo(1))
            End If
            If UBound(strInfo) >= 0 Then
                strInfo = Split(strInfo(0), "-")
                For i = 0 To UBound(strInfo)
                    .txtDiagnose(i) = IIf(Val(strInfo(i)) = 0, "", strInfo(i))
                Next
            End If
            
            '死亡日期
            strInfo = Split(CStr(colData("K19")), "-")
            For i = 0 To UBound(strInfo)
                .txtDeath(i) = IIf(Val(strInfo(i)) = 0, "", strInfo(i))
            Next
            
            '其他法定管理以及重点监测传染病
            .txtImportant.Text = CStr(colData("K31"))
            If glngVersion = VL_2014 Then
                '订正病名
                .txtIName.Text = CStr(colData("K38"))
                '退卡原因
                .txtReason.Text = CStr(colData("K39"))
                '报告单位
                .txtUnit.Text = CStr(colData("K40"))
                '联系电话
                .txtDocNumber.Text = CStr(colData("K41"))
                '填卡医生
                .txtDoctor.Text = CStr(colData("K42"))
                '备注
                .txtRemarks.Text = CStr(colData("K44"))
                '填卡日期
                strTmp = CStr(colData("K43"))
                strInfo = Split(strTmp, "-")
                For i = 0 To UBound(strInfo)
                    .txtEnter(i) = strInfo(i)
                Next
                mstrVL2014 = CStr(colData("K32")) & "$" & CStr(colData("K33")) & "$" & CStr(colData("K34")) & "$" & CStr(colData("K35")) & "$" & CStr(colData("K36")) & "$" & CStr(colData("K37")) & "$"
            ElseIf glngVersion = VL_2016 Then
                '订正病名
                .txtIName.Text = CStr(colData("K32"))
                '退卡原因
                .txtReason.Text = CStr(colData("K33"))
                '报告单位
                .txtUnit.Text = CStr(colData("K34"))
                '联系电话
                .txtDocNumber.Text = CStr(colData("K35"))
                '填卡医生
                .txtDoctor.Text = CStr(colData("K36"))
                '备注
                .txtRemarks.Text = CStr(colData("K38"))
                '填卡日期
                strTmp = CStr(colData("K37"))
                strInfo = Split(strTmp, "-")
                For i = 0 To UBound(strInfo)
                    .txtEnter(i) = strInfo(i)
                Next
            End If
        Else   '新增
            .txtName.Text = CStr(colData("K0"))              '姓名
            .txtParentName.Text = CStr(colData("KParent"))   '家长姓名
            '身份证号
            strTmp = CStr(colData("K1"))
            For i = 0 To .txtIDCard.Count - 1
                .txtIDCard(i).Text = Mid(strTmp, i + 1, 1)
            Next
            
            .ucSex(IIf(CStr(colData("K2")) = "男", 0, 1)).Checked = True
            
            strTmp = Format(CStr(colData("K3")), "yyyy-mm-dd")
            If strTmp <> "年月日" Then
                strInfo = Split(strTmp, "-")
                For i = 0 To UBound(strInfo)
                    .txtBirth(i).Text = strInfo(i)
                Next
            End If
            
            strTmp = Trim(CStr(colData("K4")))
            i = InStr("大岁月日天", Right(strTmp, 1))
            If i > 1 Then
                i = IIf(i > 4, 4, i)
                .txtAge.Text = Val(CStr(colData("K4")))
                .ucAge(i - 2).Checked = True
            Else
                .txtAge.Text = Val(CStr(colData("K4")))
                .ucAge(0).Checked = True
                If Val(.txtAge.Text) = 0 Then
                    .txtAge.Text = ""
                    .ucAge(0).Checked = False
                End If
            End If
            
            .txtAddress.Text = CStr(colData("K5"))
            
            If CStr(colData("K6")) <> "" Then
                strTmp = CStr(colData("K6"))
            ElseIf CStr(colData("K7")) <> "" Then
                strTmp = CStr(colData("K7"))
            Else
                strTmp = CStr(colData("K8"))
            End If
            .txtPhone.Text = strTmp
            
            .lblReport(13).ToolTipText = CStr(colData("K13"))
            
            strTmp = Format(CStr(colData("K14")), "yyyy-mm-dd")
            If strTmp <> "年月日" Then
                strInfo = Split(strTmp, "-")
                For i = 0 To UBound(strInfo)
                    .txtAttack(i).Text = strInfo(i)
                Next
            End If
            
            strTmp = Format(CStr(colData("K15")), "yyyy-mm-dd-hh")
            If strTmp <> "年月日" Then
                strInfo = Split(strTmp, "-")
                For i = 0 To UBound(strInfo)
                    .txtDiagnose(i).Text = strInfo(i)
                Next
            End If
            
            strTmp = Format(CStr(colData("K17")), "yyyy-mm-dd")
            If strTmp <> "年月日" Then
                strInfo = Split(strTmp, "-")
                For i = 0 To UBound(strInfo)
                    .txtDeath(i).Text = strInfo(i)
                Next
            End If
            
            strTmp = CStr(colData("K16"))
            
            '甲类传染病
            For i = 0 To .ucInfectiousA.Count - 1
                strName = Trim(Replace(Replace(Replace(.ucInfectiousA(i).Caption, "、", ""), ")", ""), "(", ""))
                .ucInfectiousA(i).Checked = IIf(InStr(strTmp, strName) <> 0, True, False)
            Next
            
            '乙类传染病
            For i = 0 To .ucInfectiousB.Count - 1
                strName = Trim(.ucInfectiousB(i).Caption)
                strName = Mid(strName, 1, Len(strName) - 1)
                If InStr(strTmp, strName) <> 0 Then
                    If .ucInfectiousB(i).CheckedVisible Then .ucInfectiousB(i).Checked = True
                End If
            Next
            
            '丙类传染病
            For i = 0 To .ucInfectiousC.Count - 1
                strName = Trim(.ucInfectiousC(i).Caption)
                strName = Mid(strName, 1, Len(strName) - 1)
                If InStr(strTmp, strName) <> 0 Then
                    .ucInfectiousC(i).Checked = True
                End If
            Next
            
            '艾滋病
            For i = 0 To .ucAIDS.Count - 1
                strName = Trim(Replace(Replace(Replace(.ucAIDS(i).Caption, "、", ""), ")", ""), "(", ""))
                .ucAIDS(i).Checked = IIf(InStr(strTmp, strName) <> 0, True, False)
            Next
    
            
            '病毒性肝炎
            For i = 0 To .ucHepatitis.Count - 1
                strName = Trim(.ucHepatitis(i).Caption)
                strName = Mid(strName, 1, Len(strName) - IIf(i = 5, 2, 1))
                If InStr(strTmp, strName) <> 0 Then
                    .ucHepatitis(i).Checked = True
                End If
            Next
            
            '炭疽
            For i = 0 To .ucAnthrax.Count - 1
                strName = Trim(Replace(Replace(Replace(.ucAnthrax(i).Caption, "、", ""), ")", ""), "(", ""))
                .ucAnthrax(i).Checked = IIf(InStr(strTmp, strName) <> 0, True, False)
            Next
            
            '痢疾
            For i = 0 To .ucDysentery.Count - 1
                strName = Trim(Replace(Replace(Replace(.ucDysentery(i).Caption, "、", ""), ")", ""), "(", ""))
                .ucDysentery(i).Checked = IIf(InStr(strTmp, strName) <> 0, True, False)
            Next
     
            
            '肺结核
            For i = 0 To .ucPTB.Count - 1
                strName = Trim(.ucPTB(i).Caption)
                strName = Mid(strName, 1, Len(strName) - IIf(i = 2, 2, 1))
                If InStr(strTmp, strName) <> 0 Then
                    .ucPTB(i).Checked = True
                End If
            Next
            
            '伤寒
            For i = 0 To .ucTyphia.Count - 1
                strName = Trim(Replace(Replace(Replace(.ucTyphia(i).Caption, "、", ""), ")", ""), "(", ""))
                .ucTyphia(i).Checked = IIf(InStr(strTmp, strName) <> 0, True, False)
            Next
            
            '梅毒
            For i = 0 To .ucSyphilis.Count - 1
                strName = Trim(.ucSyphilis(i).Caption)
                strName = Mid(strName, 1, Len(strName) - IIf(i = 4, 2, 1))
                If InStr(strTmp, strName) <> 0 Then
                    .ucSyphilis(i).Checked = True
                End If
            Next
            
            '疟疾
            For i = 0 To .ucMalaria.Count - 1
                strName = Trim(Replace(Replace(Replace(.ucMalaria(i).Caption, "、", ""), ")", ""), "(", ""))
                .ucMalaria(i).Checked = IIf(InStr(strTmp, strName) <> 0, True, False)
            Next
                  
            .txtUnit.Text = CStr(colData("K11"))
            blnStructAdress = Val(zlDatabase.GetPara(251, glngSys)) <> 0 '病人地址结构化录入
            blnShowTown = Val(zlDatabase.GetPara(252, glngSys)) <> 0 '乡镇地址结构化录入
             i = 0
            If blnStructAdress Then
                strSQL = "Select 省,市,县,乡镇,其他 From 病人地址信息 Where 病人ID=[1] and (Nvl(主页ID,0)=[2] Or Nvl(主页ID,0)=0) and 地址类别=3 order by 主页ID"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "查询结构化地址", mlngPatiID, mlngPageID)
                If rsTmp.RecordCount > 0 Then
                    
                    If rsTmp!市 & "" = "市辖区" Then
                        .txtAddInfo(0).Text = ""
                        If rsTmp!省 & "" <> "" Then .txtAddInfo(1).Text = IIf(Right(rsTmp!省 & "", 1) = "市", Left(rsTmp!省 & "", Len(rsTmp!省 & "") - 1), rsTmp!省 & "")
                    Else
                        If rsTmp!省 & "" <> "" Then .txtAddInfo(0).Text = IIf(Right(rsTmp!省 & "", 1) = "省", Left(rsTmp!省 & "", Len(rsTmp!省 & "") - 1), rsTmp!省 & "")
                        If rsTmp!市 & "" <> "" Then .txtAddInfo(1).Text = IIf(Right(rsTmp!市 & "", 1) = "市", Left(rsTmp!市 & "", Len(rsTmp!市 & "") - 1), rsTmp!市 & "")
                    End If
                    If rsTmp!县 & "" <> "" Then
                        If InStr(rsTmp!县 & "", "镇") > 0 Then
                            .txtAddInfo(2).Text = ""
                            i = 1
                            .txtAddInfo(2 + i).Text = IIf(Right(rsTmp!县 & "", 1) = "县" Or Right(rsTmp!县 & "", 1) = "区", Left(rsTmp!县 & "", Len(rsTmp!县 & "") - 1), rsTmp!县 & "")
                        Else
                            .txtAddInfo(2).Text = IIf(Right(rsTmp!县 & "", 1) = "县" Or Right(rsTmp!县 & "", 1) = "区", Left(rsTmp!县 & "", Len(rsTmp!县 & "") - 1), rsTmp!县 & "")
                        End If
                    End If
                    If blnShowTown Then
                        If i <> 1 Then
                            If rsTmp!乡镇 & "" <> "" Then .txtAddInfo(3).Text = IIf(Right(rsTmp!乡镇 & "", 1) = "乡" Or Right(rsTmp!乡镇 & "", 1) = "镇", Left(rsTmp!乡镇 & "", Len(rsTmp!乡镇 & "") - 1), rsTmp!乡镇 & "")
                        End If
                        If rsTmp!其他 & "" <> "" Then .txtAddInfo(4).Text = rsTmp!其他 & ""
                    Else
                        If i = 1 Then
                            If rsTmp!其他 & "" <> "" Then .txtAddInfo(4).Text = rsTmp!其他 & ""
                        Else
                            If rsTmp!其他 & "" <> "" Then .txtAddInfo(3).Text = IIf(Right(rsTmp!其他 & "", 1) = "乡" Or Right(rsTmp!其他 & "", 1) = "镇", Left(rsTmp!其他 & "", Len(rsTmp!其他 & "") - 1), rsTmp!其他 & "")
                        End If
                    End If
                End If
            End If

        End If
        Call .LoadData(mlngFileID, mlngPatiID, mlngPageID, mbytType, mbytFrom, mlngDeptID, mbytBabyNo)
    End With
    
    mblnFirst = False
    Call SaveLoadData
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Public Function MakeSaveSql(arrSql() As Variant, colCls As Collection, ByVal strFileId As String) As Boolean
    Dim strObjNo As String
    Dim strContent As String
    Dim strReportInfo As String
    Dim i As Integer
    Dim strTmp As String
    Dim strTmpB As String
    Dim strObjNo_2014 As String
    Dim strObjNo_2016 As String
    Dim obj As Object


    On Error GoTo errHand
    With mfrmReportEx
        strObjNo_2014 = "1$2$3$4$5$6$7$8$9$10$11$12$13$14$15$16$17$18$19$20$21$22$23$24$25$26$27$28$29$30$31$32$33$34$35$36$37$38$39$40$41$42$43$44$45"
        strObjNo_2016 = "1$2$3$4$5$6$7$8$9$10$11$12$13$14$15$16$17$18$19$20$21$22$23$24$25$26$27$28$29$30$31$32$33$34$35$36$37$38$39"
        
        strTmp = IIf(.ucCheckType(0).Checked = True, 1, 2)
        strContent = .txtNumber.Text & "$"
        
        For i = 0 To .ucCheckType.Count - 1
            If .ucCheckType(i).Checked = True Then
                strTmp = .ucCheckType(i).Caption
                Exit For
            End If
        Next
        strContent = strContent & strTmp & "$"
        
        
        '姓名、患者父母姓名
        strContent = strContent & .txtName.Text & "$" & .txtParentName.Text & "$"
        
        '身份证号码
        strTmp = ""
        For i = 0 To .txtIDCard.Count - 1
            strTmp = Trim(strTmp) & Trim(.txtIDCard(i).Text)
        Next
        strContent = strContent & strTmp & "$"
        
        '性别
        strTmp = ""
        For i = 0 To .ucSex.Count - 1
            If .ucSex(i).Checked = True Then
                strTmp = .ucSex(i).Caption
                Exit For
            End If
        Next
        strContent = strContent & strTmp & "$"
        
        '出生日期
        strTmp = ""
        For i = 0 To .txtBirth.Count - 1
            strTmp = strTmp & "-" & IIf(Trim(.txtBirth(i).Text) = "", 0, Trim(.txtBirth(i).Text))
        Next
        strTmp = Mid(strTmp, 2)
        
        If Trim(strTmp) = "--" Then
            strTmp = ""
        End If
        strContent = strContent & strTmp & "$"
        
        '年龄
        strContent = strContent & Trim(.txtAge.Text) & "$"
        
        '年龄单位
        strTmp = ""
        For i = 0 To .ucAge.Count - 1
            If .ucAge(i).Checked = True Then
                strTmp = .ucAge(i).Caption
                Exit For
            End If
        Next
        strContent = strContent & strTmp & "$"
        
        '工作单位
        strContent = strContent & Trim(.txtAddress.Text) & "$"
        
        '联系电话
        strContent = strContent & Trim(.txtPhone.Text) & "$"
        
        '病人属于
        For i = 0 To .ucFrom.Count - 1
            If .ucFrom(i).Checked = True And .ucFrom(i).Visible = True Then
                strTmp = .ucFrom(i).Caption
                Exit For
            End If
            strTmp = ""
        Next
        strContent = strContent & strTmp & "$"
        
        '现居住
        strTmp = ""
        For i = 0 To .txtAddInfo.Count - 1
            strTmp = Trim(strTmp) & Trim(.txtAddInfo(i).Text) & ";"
        Next
        strContent = strContent & strTmp & "$"
        
        '患者职业
        For i = 0 To .ucCheckJob.Count - 1
            If .ucCheckJob(i).Checked = True And .ucCheckJob(i).Visible = True Then
                strTmp = .ucCheckJob(i).Caption
                Exit For
            End If
            strTmp = ""
        Next
        strContent = strContent & strTmp & "$"
        
        '病例分类1
        For i = 0 To .ucCaseType1.Count - 1
            If .ucCaseType1(i).Checked = True Then
                strTmp = .ucCaseType1(i).Caption
                Exit For
            End If
            strTmp = ""
        Next
        strContent = strContent & strTmp & "$"
        
        '病例分类2
        For i = 0 To .ucCaseType2.Count - 1
            If .ucCaseType2(i).Checked = True Then
                strTmp = .ucCaseType2(i).Caption
                Exit For
            End If
            strTmp = ""
        Next
        strContent = strContent & strTmp & "$"
    
        '发病日期
        strTmp = ""
        For i = 0 To .txtAttack.Count - 1
            strTmp = strTmp & "-" & IIf(Trim(.txtAttack(i).Text) = "", 0, Trim(.txtAttack(i).Text))
        Next
        
        strTmp = Mid(strTmp, 2)
        If Trim(strTmp) = "--" Then
            strTmp = ""
        End If
        strContent = strContent & strTmp & "$"
        
        '诊断日期
        strTmp = IIf(Trim(.txtDiagnose(0).Text) = "", 0, Trim(.txtDiagnose(0).Text)) & "-" & IIf(Trim(.txtDiagnose(1).Text) = "", 0, Trim(.txtDiagnose(1).Text)) & "-" & IIf(Trim(.txtDiagnose(2).Text) = "", 0, Trim(.txtDiagnose(2).Text)) & " " & IIf(Trim(.txtDiagnose(3).Text) = "", 0, Trim(.txtDiagnose(3).Text))
        strContent = strContent & strTmp & "$"
        
        '死亡日期
        strTmp = ""
        For i = 0 To .txtDeath.Count - 1
            strTmp = strTmp & "-" & IIf(Trim(.txtDeath(i).Text) = "", 0, Trim(.txtDeath(i).Text))
        Next
        
        strTmp = Mid(strTmp, 2)
        If Trim(strTmp) = "--" Then
            strTmp = ""
        End If
        strContent = strContent & strTmp & "$"
        
            '甲类传染病
        strTmp = ""
        For i = 0 To .ucInfectiousA.Count - 1
            If .ucInfectiousA(i).Checked = True Then
                strTmp = Trim(strTmp) & ";" & Trim(.ucInfectiousA(i).Caption)
            End If
        Next
        strTmp = Mid(strTmp, 2)
        strContent = strContent & strTmp & "$"
        
        '艾滋病
        strTmp = ""
        For i = 0 To .ucAIDS.Count - 1
            If .ucAIDS(i).Checked = True Then
                strTmp = .ucAIDS(i).Caption
                Exit For
            End If
        Next
        strTmpB = strTmpB & strTmp & "$"
        .ucInfectiousB(1).Checked = IIf(strTmp <> "", True, False)
        
        '病毒性肝炎
        strTmp = ""
        For i = 0 To .ucHepatitis.Count - 1
            If .ucHepatitis(i).Checked = True Then
                strTmp = .ucHepatitis(i).Caption
                Exit For
            End If
        Next
        strTmpB = strTmpB & strTmp & "$"
        .ucInfectiousB(2).Checked = IIf(strTmp <> "", True, False)
        
        '炭疽
        strTmp = ""
        For i = 0 To .ucAnthrax.Count - 1
            If .ucAnthrax(i).Checked = True Then
                strTmp = .ucAnthrax(i).Caption
                Exit For
            End If
        Next
        strTmpB = strTmpB & strTmp & "$"
        .ucInfectiousB(10).Checked = IIf(strTmp <> "", True, False)
        
        '痢疾
        strTmp = ""
        For i = 0 To .ucDysentery.Count - 1
            If .ucDysentery(i).Checked = True Then
                strTmp = .ucDysentery(i).Caption
                Exit For
            End If
        Next
        strTmpB = strTmpB & strTmp & "$"
        .ucInfectiousB(11).Checked = IIf(strTmp <> "", True, False)
        
        '肺结核
        strTmp = ""
        For i = 0 To .ucPTB.Count - 1
            If .ucPTB(i).Checked = True Then
                strTmp = .ucPTB(i).Caption
                Exit For
            End If
        Next
        strTmpB = strTmpB & strTmp & "$"
        .ucInfectiousB(12).Checked = IIf(strTmp <> "", True, False)
        
        '伤寒
        strTmp = ""
        For i = 0 To .ucTyphia.Count - 1
            If .ucTyphia(i).Checked = True Then
                strTmp = .ucTyphia(i).Caption
                Exit For
            End If
        Next
        strTmpB = strTmpB & strTmp & "$"
        .ucInfectiousB(13).Checked = IIf(strTmp <> "", True, False)
        
        '梅毒
        strTmp = ""
        For i = 0 To .ucSyphilis.Count - 1
            If .ucSyphilis(i).Checked = True Then
                strTmp = .ucSyphilis(i).Caption
                Exit For
            End If
        Next
        strTmpB = strTmpB & strTmp & "$"
        .ucInfectiousB(20).Checked = IIf(strTmp <> "", True, False)
        
        '疟疾
        strTmp = ""
        For i = 0 To .ucMalaria.Count - 1
            If .ucMalaria(i).Checked = True Then
                strTmp = .ucMalaria(i).Caption
                Exit For
            End If
        Next
        strTmpB = strTmpB & strTmp & "$"
        .ucInfectiousB(23).Checked = IIf(strTmp <> "", True, False)
        
        '乙类传染病
        strTmp = ""
        For i = 0 To .ucInfectiousB.UBound
            If .ucInfectiousB(i).Checked = True Then
                strTmp = strTmp & ";" & .ucInfectiousB(i).Caption
            End If
        Next
        If strTmp <> "" Then
            strTmp = Mid(strTmp, 2)
        End If
        strContent = strContent & strTmp & "$"
        
        '连接乙类单个疾病
        strContent = strContent & strTmpB
        
        '丙类传染病
        strTmp = ""
        For i = 0 To .ucInfectiousC.UBound
            If .ucInfectiousC(i).Checked = True Then
                strTmp = strTmp & ";" & .ucInfectiousC(i).Caption
            End If
        Next
        If strTmp <> "" Then
            strTmp = Mid(strTmp, 2)
        End If
        strContent = strContent & strTmp & "$"

        '其它传染病
        strTmp = ""
        strContent = strContent & Trim(.txtImportant.Text) & "$"
        If glngVersion = VL_2014 Then
            strContent = strContent & mstrVL2014
        End If
        '订正病名、退卡原因、报告单位、联系电话、填卡医生
        strContent = strContent & .txtIName.Text & "$" & .txtReason.Text & "$" & .txtUnit.Text & "$" & .txtDocNumber.Text & "$" & .txtDoctor.Text & "$"
    
        '填卡日期
        strTmp = .txtEnter(0).Text & "-" & .txtEnter(1).Text & "-" & .txtEnter(2).Text
        If Trim(strTmp) = "--" Then
            strTmp = ""
        End If
        strContent = strContent & strTmp & "$"
    
        '备注
        strContent = strContent & .txtRemarks.Text & "$"
        
        '附加传染病
        strTmp = ""
        For Each obj In .Controls
            If TypeName(obj) = "uCheckNorm" Then
                If InStr(GSTR_Controls, obj.Name) = 0 Then
                    If obj.Checked = True Then
                        strTmp = strTmp & ";" & obj.Name & ":" & obj.Caption
                    End If
                End If
            End If
        Next
        
        If strTmp <> "" Then
            strTmp = Mid(strTmp, 2)
        End If
        strContent = strContent & strTmp & "$"
        
        If glngVersion = VL_2014 Then
            strReportInfo = strObjNo_2014 & "|" & strContent
        Else
            strReportInfo = strObjNo_2016 & "|" & strContent
        End If
    End With
    
    MakeSaveSql = GetSaveSql(arrSql, colCls, strFileId, strReportInfo)
    Call SaveLoadData
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Function


Private Sub SaveLoadData()
'功能：保存控件显示信息
    Dim objCtl As Control
    Dim i As Integer
    i = 0
    Set mcolLoadData = New Collection
    For Each objCtl In mfrmReportEx.Controls
        Select Case TypeName(objCtl)
            Case "TextBox"
                Call mcolLoadData.Add(objCtl.Text, "K" & i)
            Case "uCheckNorm"
                Call mcolLoadData.Add(IIf(objCtl.Checked = True, 1, 0), "K" & i)
        End Select
        i = i + 1
    Next
End Sub

Public Function HaveChanged() As Boolean
'判断控件显示信息是否发生变化
    Dim objCtl As Control
    Dim i As Integer
    i = 0
    HaveChanged = False
    If mcolLoadData Is Nothing Then
        Set mcolLoadData = New Collection
    End If
    If mcolLoadData.Count <= 0 Then
        Exit Function
    End If
    For Each objCtl In mfrmReportEx.Controls
        Select Case TypeName(objCtl)
            Case "TextBox"
                If objCtl.Text <> mcolLoadData("K" & i) Then
                    HaveChanged = True
                    Exit Function
                End If
            Case "uCheckNorm"
                If IIf(objCtl.Checked = True, 1, 0) <> mcolLoadData("K" & i) Then
                    HaveChanged = True
                    Exit Function
                End If
        End Select
        i = i + 1
    Next
End Function


Public Sub SetCaption身份证()
    mbln身份证必填 = Val(zlDatabase.GetPara("传染病报告身份证号码必填", glngSys, 1277, 0)) = 1
    If mbln身份证必填 Then
        mfrmReportEx.lblReport(7).Caption = "有效证件号*："
    Else
        mfrmReportEx.lblReport(7).Caption = "有效证件号："
    End If
End Sub




Public Sub ClearEnterInfo()
    With mfrmReportEx
        .txtDoctor.Text = ""
        .txtEnter(0).Text = ""
        .txtEnter(1).Text = ""
        .txtEnter(2).Text = ""
    End With
End Sub

Public Sub SetEnterInfo()
    Dim strDate As String
    If mColData.Count < 44 Then
        strDate = Trim(CStr(mColData("K12")))
    Else
        strDate = Trim(CStr(mColData("K43")))
    End If
    If strDate = "" Or strDate = "--" Then
        strDate = zlDatabase.Currentdate
    End If
    Call SetInfo(UserInfo.姓名, strDate)
End Sub

Public Sub SetInfo(ByVal strDoctor As String, ByVal strDate As String)
    Dim strDateInfo() As String
    Dim strCurTime() As String
    strDateInfo = Split(Format(strDate, "yyyy-mm-dd"), "-")
    strCurTime = Split(Format(zlDatabase.Currentdate, "yyyy-mm-dd"), "-")
    With mfrmReportEx
        .txtDoctor.Text = strDoctor
        If UBound(strDateInfo) < 2 Then
            .txtEnter(0).Text = strCurTime(0)
            .txtEnter(1).Text = strCurTime(1)
            .txtEnter(2).Text = strCurTime(2)
        Else
            .txtEnter(0).Text = strDateInfo(0)
            .txtEnter(1).Text = strDateInfo(1)
            .txtEnter(2).Text = strDateInfo(2)
        End If
    End With
End Sub

Public Function CheckValidity() As Boolean
'功能：检查输入的合法性
    Dim strBirth As String      '生日，例如：1991-01-01
    Dim blnIsChild As Boolean   '判断是否为14岁以下的患者
    Dim intAge As Integer
    Dim blnDate As Boolean      '判断日期是否输入完整
    Dim i As Integer, j As Integer
    Dim strMsg As String
    Dim strTmp As String
    Dim strMsgInfo() As String
    On Error GoTo errHand

    With mfrmReportEx
        strMsg = ""
        strTmp = ""
        
        '检查姓名
        If Trim(.txtName.Text) = "" Then
            strMsg = strMsg & "<姓名>为必选项，请检查！$"
        End If
        '检查身份证号
        If mbln身份证必填 Then
            For i = 0 To .txtIDCard.Count - 1
                strTmp = strTmp & .txtIDCard(i).Text
            Next
            If Trim(strTmp) = "" Then
                strMsg = strMsg & "<有效证件号>为必选项，请检查！$"
            End If
        End If
        
        '检查性别，没有选择时必须选
        If .ucSex(0).Checked = False And .ucSex(1).Checked = False Then
            strMsg = strMsg & "<性别>为必选项，请检查！$"
        End If
        
        '检查出生日期 允许全空 不允许部份空或数值无效
        .txtBirth(0).Text = Trim(.txtBirth(0).Text): .txtBirth(1).Text = Trim(.txtBirth(1).Text): .txtBirth(2).Text = Trim(.txtBirth(2).Text)
        strBirth = .txtBirth(0).Text & "-" & .txtBirth(1).Text & "-" & .txtBirth(2).Text
        
        If .txtBirth(0).Text <> "" Or .txtBirth(1).Text <> "" Or .txtBirth(2).Text <> "" Then
            If Not IsDate(strBirth) Then
                strMsg = strMsg & "<出生日期>不完整或不是有效日期，请检查！$"
            Else
                If DateDiff("yyyy", strBirth, Now()) <= 14 And Trim(.txtParentName.Text) = "" Then
                    blnIsChild = True
                End If
            End If
        End If
        
        If Trim(.txtBirth(0).Text) = "" And Trim(.txtAge.Text) = "" Then
            strMsg = strMsg & "<出生日期>与<年龄>必需填写一项，请检查！$"
        End If
        
        '检查年龄，如果小于14，必须输入父母的名字,必须输入父母的联系电话
        If .txtBirth(0).Text = "" Then
            intAge = Val(.txtAge.Text) * IIf(.ucSex(0).Checked, 365, IIf(.ucSex(1).Checked, 30, 1))
            If intAge <= (14 * 365) And Trim(.txtParentName.Text) = "" Then
                blnIsChild = True
            End If
        End If
        If blnIsChild = True Then
            If Trim(.txtPhone.Text) = "" Then
                strMsg = strMsg & "14岁以下患者要求填写<联系电话>，请检查！$"
            Else
                strMsg = strMsg & "14岁以下患者要求填写<家长姓名>，请检查！$"
            End If
        End If
        
        '检查病人属于
        For i = 0 To .ucFrom.Count - 1
            If .ucFrom(i).Checked = True Then
                Exit For
            End If
            If i = 5 Then
                strMsg = strMsg & "<病人属于>为必选项，请检查！$"
            End If
        Next
        
        '检查地址
        For i = 0 To .txtAddInfo.Count - 1
            If Trim(.txtAddInfo(i).Text) <> "" Then
                Exit For
            End If
            If i = 5 Then
                strMsg = strMsg & "<现住址>为必选项，请检查！$"
            End If
        Next
        
        '检查职业，必须选择一项
        For i = 0 To .ucCheckJob.Count - 1
            If .ucCheckJob(i).Checked = True Then
                Exit For
            End If
            If i = .ucCheckJob.Count - 1 Then
                strMsg = strMsg & "<人群分类>为必选项，请检查！$"
            End If
        Next
        
        '检查病例分类
        For i = 0 To .ucCaseType1.Count - 1
            If .ucCaseType1(i).Checked = True Then
                Exit For
            End If
            If i = .ucCaseType1.Count - 1 Then
                For j = 0 To .ucCaseType2.Count - 1
                    If .ucCaseType2(j).Checked = False Then
                        Exit For
                    End If
                    If j = .ucCaseType2.Count - 1 Then
                        strMsg = strMsg & "<病例分类>为必选项，请检查！$"
                    End If
                Next
            End If
        Next
        
        '发病日期必须填写
        .txtAttack(0).Text = Trim(.txtAttack(0).Text): .txtAttack(1).Text = Trim(.txtAttack(1).Text): .txtAttack(2).Text = Trim(.txtAttack(2).Text)
        If Not IsDate(.txtAttack(0).Text & "-" & .txtAttack(1).Text & "-" & .txtAttack(2).Text) Then
            strMsg = strMsg & "<发病日期>不完整或不是有效日期，请检查！$"
        End If
        
        '诊断时间必须完整并精确到小时
        .txtDiagnose(0).Text = Trim(.txtDiagnose(0).Text): .txtDiagnose(1).Text = Trim(.txtDiagnose(1).Text):
        .txtDiagnose(2).Text = Trim(.txtDiagnose(2).Text): .txtDiagnose(3).Text = Trim(.txtDiagnose(3).Text)
        If (Not IsDate(.txtDiagnose(0).Text & "-" & .txtDiagnose(1).Text & "-" & .txtDiagnose(2).Text)) Or (Decode(.txtDiagnose(3).Text, "", "-1", .txtDiagnose(3).Text) < 0) Or (Decode(.txtDiagnose(3).Text, "", "-1", .txtDiagnose(3).Text) > 23) Then
            strMsg = strMsg & "<诊断时间>不完整，或不是有效日期，或未精确到小时，请检查！$"
        End If
        
        '死亡日期 允许全空 不允许部份空或数值无效
        .txtDeath(0).Text = Trim(.txtDeath(0).Text): .txtDeath(1).Text = Trim(.txtDeath(1).Text): .txtDeath(2).Text = Trim(.txtDeath(2).Text)
        If .txtDeath(0).Text <> "" Or .txtDeath(1).Text <> "" Or .txtDeath(2).Text <> "" Then
            If Not IsDate(.txtDeath(0).Text & "-" & .txtDeath(1).Text & "-" & .txtDeath(2).Text) Then
                strMsg = strMsg & "<死亡日期>不完整或不是有效日期，请检查！$"
            End If
        End If
        
        '1.检查病例分类为"病原携带者"时，病种是否是"霍乱、脊髓灰质炎、艾滋病"
        If gbytDiseaseType = 3 Then
            If .ucInfectiousA(1).Checked = False And (.ucAIDS(0).Checked = False And .ucAIDS(1).Checked = False) And .ucInfectiousB(3).Checked = False Then
                strMsg = strMsg & "需报告病原携带者的法定传染病病种包括<霍乱>、<脊髓灰质炎>、<艾滋病>，请检查！$"
            End If
        End If
        
        '2."梅毒"、"淋病"的病例分类只能为"确诊病例"和"疑似病例"
        If .ucInfectiousB(25).Checked Or .ucSyphilis(0).Checked Or .ucSyphilis(1).Checked Or .ucSyphilis(2).Checked Or .ucSyphilis(3).Checked Or .ucSyphilis(4).Checked Then
            If gbytDiseaseType <> 0 And gbytDiseaseType <> 2 Then
                strMsg = strMsg & "<梅毒、淋病>的病例分类只能为<确诊病例>和<疑似病例>！$"
            End If
        End If
        
        '3.乙肝、血吸虫病例须分急性或慢性填写
        If .ucHepatitis(1).Checked = True Or .ucInfectiousB(22).Checked = True Then
            If gbytAcute <> 0 And gbytAcute <> 1 Then
                strMsg = strMsg & "<乙肝>、<血吸虫病例>须分<急性>或<慢性>，请检查！$"
            End If
        End If

	'传染病一病一卡限制
        If mblnOneCard Then
            If Not CheckOneCard(strTmp) Then
                strMsg = strMsg & strTmp & "$"
            End If
        End If
        
        '调用frmReport自定义检查
         If .CheckData(mlngFileID, mlngPatiID, mlngPageID, mbytType, mbytFrom, mlngDeptID, mbytBabyNo, strTmp) = False Then
            strMsg = strMsg & strTmp
         End If
        
        strTmp = ""
        If Trim(strMsg) = "" Then
            CheckValidity = True
        Else
            strMsgInfo = Split(strMsg, "$")
            For i = 0 To UBound(strMsgInfo) - 1
                strTmp = strTmp & i + 1 & ". " & strMsgInfo(i) & vbCrLf
            Next
            Call ShowMsg(strTmp)
            CheckValidity = False
        End If

    End With
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Function

Public Function RelateFeedback(ByVal isRelated As Boolean) As Boolean
'功能：传染病报告卡，关联阳性结果反馈单，或者取消关联
'参数：isRelated  true-关联；false-取消关联
    Dim objDisease As Object

On Error GoTo errHand
    Set objDisease = CreateObject("zl9Disease.cDockDisease")
    If objDisease Is Nothing Then Exit Function
    Call objDisease.InitDockDisease(glngSys, gcnOracle)
    Call objDisease.RelateFeedback(Me, mlngFileID, mlngPatiID, mlngPageID, mbytFrom, isRelated)
    Set objDisease = Nothing
    RelateFeedback = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckOneCard(ByRef strMsg As String) As Boolean
'功能：判断是不是只选了一个病种
    Dim i As Long
    Dim strTmp As String
    Dim lngchk As Long '界面勾选的个数
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset '从当前病人已保存的诊断的疾病人去查可以用的病种
    
    On Error GoTo errH
        
    With mfrmReportEx
    
        '甲类传染病
        For i = 0 To .ucInfectiousA.Count - 1
            If .ucInfectiousA(i).Checked Then
                lngchk = lngchk + 1
            End If
        Next
        
        '乙类传染病
        For i = 0 To .ucInfectiousB.Count - 1
            If .ucInfectiousB(i).Checked Then
                lngchk = lngchk + 1
            End If
        Next
        
        '丙类传染病
        For i = 0 To .ucInfectiousC.Count - 1
            If .ucInfectiousC(i).Checked Then
                lngchk = lngchk + 1
            End If
        Next
        
        '艾滋病
        For i = 0 To .ucAIDS.Count - 1
            If .ucAIDS(i).Checked Then
                lngchk = lngchk + 1
            End If
        Next
    
        '病毒性肝炎
        For i = 0 To .ucHepatitis.Count - 1
            If .ucHepatitis(i).Checked Then
                lngchk = lngchk + 1
            End If
        Next
        
        '炭疽
        For i = 0 To .ucAnthrax.Count - 1
            If .ucAnthrax(i).Checked Then
                lngchk = lngchk + 1
            End If
        Next
        
        '痢疾
        For i = 0 To .ucDysentery.Count - 1
            If .ucDysentery(i).Checked Then
                lngchk = lngchk + 1
            End If
        Next
        
        '肺结核
        For i = 0 To .ucPTB.Count - 1
            If .ucPTB(i).Checked Then
                lngchk = lngchk + 1
            End If
        Next
        
        '伤寒
        For i = 0 To .ucTyphia.Count - 1
            If .ucTyphia(i).Checked Then
                lngchk = lngchk + 1
            End If
        Next
        
        '梅毒
        For i = 0 To .ucSyphilis.Count - 1
            If .ucSyphilis(i).Checked Then
                lngchk = lngchk + 1
            End If
        Next
        
        '疟疾
        For i = 0 To .ucMalaria.Count - 1
            If .ucMalaria(i).Checked Then
                lngchk = lngchk + 1
                Debug.Print .ucMalaria(i).Caption
            End If
        Next
    End With
    If lngchk = 0 Then
        strMsg = "请选择一个病种！"
    ElseIf lngchk > 1 Then
        strMsg = "只能选择一个病种！"
    End If
    CheckOneCard = lngchk = 1
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function