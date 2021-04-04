VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmHistoryStudy 
   BorderStyle     =   0  'None
   ClientHeight    =   4980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vsfStudy 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   8535
      _cx             =   15055
      _cy             =   7435
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16761024
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
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
      ExplorerBar     =   3
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
   Begin XtremeCommandBars.ImageManager imgList 
      Left            =   600
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmHistoryStudy.frx":0000
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   120
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmHistoryStudy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngModule As Long
Private mblnMoved As Boolean
Private mlngRow As Long
Private mdtBegin As Date
Private mdtEnd As Date
Private mblnCustom As Boolean
Private mlngPatId As Long
Private mlngCur科室ID As Long
Private mlngLinkID As Long
Private mstrCanUse科室IDs As String
Private mblnAllDepts As Boolean
Private mblnRelatingPatient As Boolean
Private mLngAdvice As Long
Private mblnDocPatient As Boolean        '是否住院病人
Private mlngBabyNum As Long              '婴儿序号
Private mblnImageEnable As Boolean
Private mblnReportEnable As Boolean
Private mTPListCfg As TPListCfg
Private mblnListCfgOk As Boolean

Private Enum FilterID
    ID_相关设置 = 1
    ID_本次相关 = 11
    ID_他科检查 = 12
    ID_嵌入查看 = 13
    ID_自动换行 = 14
    
    ID_时间范围 = 2
    ID_一月 = 21
    ID_两月 = 22
    ID_三月 = 23
    ID_半年 = 24
    ID_一年 = 25
    ID_两年 = 26
    ID_三年 = 27
    ID_不限 = 28
    ID_自定义 = 29
    
End Enum

Private Const M_STR_COLNAME = "序号;医嘱ID;检查号;年龄;类别;项目;部位;阴阳性;当前过程;检查时间;医嘱内容;随访描述"
'Private Const M_STR_COLNAME = "序号,300,1;医嘱ID,300,2;检查号,300,3;年龄,300,4;类别,300,5;项目,300,,6;部位,300,,7;阴阳性,300,8" _
'                                                  & ";当前过程,300,9;检查时间,300,10;医嘱内容,300,11;随访描述,300,12"

Private Const M_STR_CFG = "[历史]"

Private Type TPListCfg
    strSort As String
    strList As String
End Type

Public Event OnListLostFocus()
Public Event OnLoadCfg(ByRef strListCfg As String)
Public Event OnSaveCfg(ByVal strListCfg As String)
Public Event OnListMove()
Public Event OnListMouseClick(ByVal LngAdvice As Long, ByVal X As Long, ByVal Y As Long, ByVal blnClear As Boolean)
Public Event OnSelectStudy(ByVal LngAdvice As Long, ByVal strAdvice As String, ByVal blnEmbed As Boolean)
Public Event OnDoWork(ByVal LngAdvice As Long, ByVal strFuncName As String)
Public Event OnViewReport(ByVal LngAdvice As Long)
Public Event OnRefresh(ByVal lngCount As Long)

Property Let ListRow(value As Long)
    mlngRow = value
End Property

Public Function RefreshHistoryList(ByVal LngAdvice As Long, ByVal lngModule As Long, ByVal lngPatId As Long, ByVal blnDocPatient As Boolean, _
                            ByVal lngCur科室ID As Long, ByVal strCanUse科室IDs As String, ByVal lngLinkId As Long, _
                            ByVal blnAllDepts As Boolean, ByVal blnRelatingPatient As Boolean, Optional blnForce As Boolean = False, Optional lngNum As Long = 0) As Boolean
'blnDocPatient:是否住院病人
'lngNum:婴儿序号

    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    Dim objItem As ListItem
    Dim iCount As Integer
    Dim strTime As String
    Dim dtBegin As Date
    Dim dtEnd As Date
    Dim blnMoved As Boolean
    Dim objControl As CommandBarControl
    Dim blnNoTime As Boolean
    Dim strValue As String

    On Error GoTo errHandle
    
    
    
    If lngModule = G_LNG_PATHSTATION_MODULE Then
        vsfStudy.TextMatrix(0, vsfStudy.ColIndex("检查号")) = "病理号"
    End If
    
    If lngModule = G_LNG_VIDEOSTATION_MODULE Or lngModule = G_LNG_PATHSTATION_MODULE Then
        Set objControl = cbrMain.FindControl(, conMenu_Img_Look)
        objControl.Caption = "图像"
        objControl.ToolTipText = "图像"
        objControl.IconId = 10
        
        Set objControl = cbrMain.FindControl(, conMenu_Img_Contrast)
        objControl.IconId = 11
    End If
    
    If LngAdvice <= 0 Then Exit Function
    If mLngAdvice = LngAdvice And Not blnForce Then Exit Function

    mlngPatId = lngPatId
    mlngCur科室ID = lngCur科室ID
    mstrCanUse科室IDs = strCanUse科室IDs
    mlngLinkID = lngLinkId
    mblnAllDepts = blnAllDepts
    mblnRelatingPatient = blnRelatingPatient
    mlngModule = lngModule
    mLngAdvice = LngAdvice
    mblnDocPatient = blnDocPatient
    mlngBabyNum = lngNum
    
    Call SetVisible(ID_本次相关, mblnDocPatient, True)
    
    If blnAllDepts Then
        CheckCmd ID_他科检查, True
        SetVisible ID_他科检查, True, False
    Else
        SetVisible ID_他科检查, True, True
    End If
    
    mlngRow = 0
    vsfStudy.Rows = 1
    blnNoTime = False
    
    '获取时间范围
    Select Case GetTime
        Case "一月"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 30
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "两月"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 60
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "三月"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 90
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "半年"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 180
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "一年"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 365
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "两年"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 730
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "三年"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 1095
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "不限"
            blnNoTime = True
        Case "自定义"
            dtBegin = mdtBegin
            dtEnd = mdtEnd
    End Select
    
    '不限时间时，用dtBegin + 1（1899/12/31）去判断是否转存，不能用dtBegin或直接blnMoved = true
    If blnNoTime Then
        blnMoved = MovedByDate(dtBegin + 1)
    Else
        blnMoved = MovedByDate(dtBegin)
    End If
    
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        strSQL = "Select A.ID 医嘱ID,A.开嘱时间  开嘱时间,A.医嘱内容, B.执行过程, B.结果阳性,C.检查号, C.影像类别,C.随访描述,C.年龄,C.接收日期 检查时间,E.名称,E.标本部位 " & _
               " From 病人医嘱记录 A,病人医嘱发送 B,影像检查记录 C,病人信息 D,诊疗项目目录 E" & _
               " Where A.病人id = [1] And A.相关id Is Null And B.医嘱ID=A.ID and a.病人id = d.病人id  " & _
               " AND A.ID=C.医嘱ID(+) AND A.诊疗项目ID = E.ID AND b.执行过程 >= 2 "
    Else
        strSQL = "Select A.ID 医嘱ID,A.开嘱时间  开嘱时间,A.医嘱内容, B.执行过程, B.结果阳性,C.病理号 检查号,E.名称,E.标本部位,F.随访描述,F.影像类别,F.年龄,F.接收日期 检查时间 " & _
               " From 病人医嘱记录 A,病人医嘱发送 B,影像检查记录 F,病理检查信息 C,病人信息 D,诊疗项目目录 E" & _
               " Where A.病人id = [1] And A.相关id Is Null And B.医嘱ID=A.ID and a.病人id = d.病人id " & _
               " AND A.ID=C.医嘱ID(+) AND A.诊疗项目ID = E.ID and a.id=F.医嘱ID(+) AND b.执行过程 >= 2 "
    End If
    
    If Not blnNoTime Then
        strSQL = strSQL & " AND B.发送时间 between [6] and [7]"
    End If
    
    '本次检查
    If IsCheck(ID_本次相关) And mblnDocPatient Then
        strSQL = strSQL & " And (A.病人来源=2 And A.主页ID=D.主页ID)"
    End If
    
    '它科检查
    If blnAllDepts = False Then
        If Not IsCheck(ID_他科检查) Then
            strSQL = strSQL & " And A.执行科室id+0 =[2] "
        Else
            strSQL = strSQL & " And  (A.执行科室id+0 <>[2] and B.执行过程 >= 5 or A.执行科室id+0 =[2]) "
        End If
    Else
        strSQL = strSQL & " And (Instr( [3],',' || A.执行科室id || ',' ) >0)"
    End If
    
    '婴儿
    strSQL = strSQL & " And NVL(A.婴儿,0) = [8]"
    
    '启用关联病人，才查询关联ID
    If blnRelatingPatient And lngLinkId <> 0 Then
        If mlngModule <> G_LNG_PATHOLSYS_NUM Then
            strSQL = strSQL & " union select A.ID 医嘱ID,A.开嘱时间  开嘱时间,A.医嘱内容, B.执行过程, B.结果阳性,C.检查号, C.影像类别,C.随访描述,C.年龄,C.接收日期 检查时间,E.名称,E.标本部位 " & _
                " From 病人医嘱记录 A ,病人医嘱发送 B,影像检查记录 C,病人信息 D,诊疗项目目录 E" & _
                " Where B.医嘱ID=A.ID AND A.ID=C.医嘱ID(+) and a.病人id = d.病人id AND A.诊疗项目ID = E.ID AND b.执行过程 >= 2 AND A.id in (Select 医嘱ID from 影像检查记录 Where 关联ID =[4]) "
        Else
            strSQL = strSQL & " union select A.ID 医嘱ID,A.开嘱时间  开嘱时间,A.医嘱内容, B.执行过程, B.结果阳性,C.病理号 检查号,E.名称,E.标本部位,F.随访描述,F.影像类别,F.年龄,F.接收日期 检查时间" & _
                " From 病人医嘱记录 A,病人医嘱发送 B,影像检查记录 F,病理检查信息 C,病人信息 D,诊疗项目目录 E" & _
                " Where A.id in (Select 医嘱ID from 影像检查记录 Where 关联ID =[4]) And B.医嘱ID=A.ID and a.id=C.医嘱ID(+) and a.病人id = d.病人id AND A.诊疗项目ID = E.ID and a.id=F.医嘱ID(+) and b.执行过程 >= 2 "
        End If
        
        If Not blnNoTime Then
            strSQL = strSQL & " AND B.发送时间 between [6] and [7]"
        End If
        
        '本次检查
        If IsCheck(ID_本次相关) And mblnDocPatient Then
            strSQL = strSQL & " And (A.病人来源=2 And A.主页ID=D.主页ID)"
        End If
        
'        '它科检查
'        If chkOtherDeptReport.Value <> 1 Then
'            strSql = strSql & " And c.执行科室id+0 in(select  部门id  from 部门人员 where 人员id = [5] union all select to_Number([2]) from dual) "
'        End If
        '它科检查
        If blnAllDepts = False Then
            If Not IsCheck(ID_他科检查) Then
                strSQL = strSQL & " And A.执行科室id+0 =[2] "
            Else
                strSQL = strSQL & " And  (A.执行科室id+0 <>[2] and B.执行过程 >= 5 or A.执行科室id+0 =[2]) "
            End If
        Else
            strSQL = strSQL & " And (Instr( [3],',' || A.执行科室id || ',' ) >0)"
        End If
        
        strSQL = strSQL & " And NVL(A.婴儿,0) = [8]"
    End If
    
    If blnMoved Then
        strTemp = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strTemp = Replace(strTemp, "病人医嘱发送", "H病人医嘱发送")
        strTemp = Replace(strTemp, "影像检查记录", "H影像检查记录")
        strTemp = Replace(strTemp, "病人检查信息", "H病人检查信息")
        strSQL = strSQL & vbNewLine & " Union ALL " & vbNewLine & strTemp
    End If
    strSQL = "Select * From (" & vbNewLine & strSQL & vbNewLine & ") Order By 开嘱时间 Asc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "", lngPatId, _
            lngCur科室ID, "," & strCanUse科室IDs & ",", lngLinkId, UserInfo.ID, dtBegin, dtEnd, lngNum)
    
    If rsTemp.RecordCount > 0 Then

        rsTemp.Filter = "医嘱id <> " & LngAdvice
        
        With vsfStudy
            If mlngModule = G_LNG_PATHOLSYS_NUM Then .ColHidden(.ColIndex("类别")) = True
                Do While Not rsTemp.EOF
                    .Rows = .Rows + 1
                    iCount = iCount + 1
                    
                    .TextMatrix(iCount, .ColIndex("医嘱ID")) = Val(nvl(rsTemp!医嘱ID))
        
                    .TextMatrix(iCount, .ColIndex("序号")) = iCount
                    .TextMatrix(iCount, .ColIndex("检查号")) = nvl(rsTemp!检查号)
                    .TextMatrix(iCount, .ColIndex("年龄")) = nvl(rsTemp!年龄)
                    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
                        .TextMatrix(iCount, .ColIndex("类别")) = nvl(rsTemp!影像类别)
                    End If
                    .TextMatrix(iCount, .ColIndex("项目")) = nvl(rsTemp!名称)
                    .TextMatrix(iCount, .ColIndex("当前过程")) = Decode(Val(nvl(rsTemp!执行过程, 0)), -1, "已驳回", 0, "已登记", 1, "已登记", _
                                                        2, "已报到", 3, "已检查", 4, "已报告", 5, "已审核", "已完成")
                    .Cell(flexcpData, iCount, .ColIndex("当前过程")) = Val(nvl(rsTemp!执行过程, 0))
                    .TextMatrix(iCount, .ColIndex("阴阳性")) = IIf(Val(nvl(rsTemp!结果阳性)) = 1, "阳", "")
                    
                    
                    strTime = Format(rsTemp!检查时间, "yyyy-MM-dd hh:mm")
                    .TextMatrix(iCount, .ColIndex("随访描述")) = nvl(rsTemp!随访描述)
                    .TextMatrix(iCount, .ColIndex("检查时间")) = strTime
                    
                    
                    If UBound(Split(nvl(rsTemp!医嘱内容), ":")) > 0 Then
                        .TextMatrix(iCount, .ColIndex("医嘱内容")) = Split(nvl(rsTemp!医嘱内容), ":")(0)
                        .TextMatrix(iCount, .ColIndex("部位")) = Split(nvl(rsTemp!医嘱内容), ":")(1)
                    Else
                        .TextMatrix(iCount, .ColIndex("医嘱内容")) = nvl(rsTemp!医嘱内容)
                        .TextMatrix(iCount, .ColIndex("部位")) = ""
                    End If
                    
                    rsTemp.MoveNext
'                   If .Rows > 1 Then .Row = 1
                Loop
        End With
        
        If IsCheck(ID_自动换行) Then
            vsfStudy.WordWrap = True
            vsfStudy.AutoSize 1, vsfStudy.Cols - 1
        Else
            vsfStudy.WordWrap = False
            vsfStudy.AutoSize 1, vsfStudy.Cols - 1
        End If
    
    End If
    
    If Not mblnListCfgOk Then
        RaiseEvent OnLoadCfg(strValue)
        mblnListCfgOk = True
        
        If InStr(strValue, ";") > 0 Then
            mTPListCfg.strList = Split(strValue, ";")(1)
            mTPListCfg.strSort = Split(strValue, ";")(0)
            mTPListCfg.strSort = Replace(mTPListCfg.strSort, "[历史]", "")
            Call DoLoadListCfg(mTPListCfg.strList)
            Call DoLoadListSort(mTPListCfg.strSort)
        Else
            mTPListCfg.strList = ""
            mTPListCfg.strSort = ""
        End If
    Else
        If mTPListCfg.strList <> "" Then
            Call DoLoadListCfg(mTPListCfg.strList)
        End If
        
        If mTPListCfg.strSort <> "" Then
            Call DoLoadListSort(mTPListCfg.strSort)
        End If
    End If
    
    
        
    RaiseEvent OnRefresh(rsTemp.RecordCount)
    RefreshHistoryList = True
    Exit Function
errHandle:
    MsgBox err.Description, vbOKOnly, gstrSysName
    err.Clear
End Function


Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnResult As Boolean
    
    On Error GoTo errHandle
    
    Select Case Control.ID
        Case conMenu_Img_Look, conMenu_Img_Contrast, conMenu_PacsReport_Open
            RaiseEvent OnDoWork(Val(vsfStudy.TextMatrix(vsfStudy.Row, vsfStudy.ColIndex("医嘱ID"))), Control.Category)
        Case ID_自定义
            If mblnCustom Then Exit Sub
            CheckCmd Control.ID, Not Val(Control.Category) = 1
            mblnCustom = True
            blnResult = frmSetTime.ShowSetTime(mdtBegin, mdtEnd, Me)
            mblnCustom = False

            Call RefreshHistoryList(mLngAdvice, mlngModule, mlngPatId, mblnDocPatient, mlngCur科室ID, mstrCanUse科室IDs, mlngLinkID, mblnAllDepts, mblnRelatingPatient, True, mlngBabyNum)
        Case Else
            CheckCmd Control.ID, Not Val(Control.Category) = 1
            Call RefreshHistoryList(mLngAdvice, mlngModule, mlngPatId, mblnDocPatient, mlngCur科室ID, mstrCanUse科室IDs, mlngLinkID, mblnAllDepts, mblnRelatingPatient, True, mlngBabyNum)
    End Select
    Exit Sub
errHandle:
    MsgBox err.Description, vbOKOnly, gstrSysName
    err.Clear
End Sub

Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error Resume Next
    
    Select Case Control.ID
        Case conMenu_Img_Look, conMenu_Img_Contrast, conMenu_PacsReport_Open
            If vsfStudy.ColIndex("医嘱ID") = -1 Then Exit Sub
            Control.Enabled = Val(vsfStudy.TextMatrix(vsfStudy.Row, vsfStudy.ColIndex("医嘱ID"))) > 0
            Control.Visible = Not IsCheck(ID_嵌入查看)
            If Control.Visible And Control.Enabled Then
                If Control.ID = conMenu_PacsReport_Open Then
                    Control.Enabled = mblnReportEnable
                Else
                    Control.Enabled = mblnImageEnable
                End If
            End If
    End Select
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    
    mblnImageEnable = False
    mblnReportEnable = False
    
    mblnListCfgOk = False
    Call InitCommandBars
    Call GridInit(M_STR_COLNAME)
    Call SetFontSize(gbytFontSize)
    
    
    CheckCmd ID_他科检查, Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "他科检查", "0")) = 1
    CheckCmd ID_本次相关, Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "本次相关", "0")) = 1
    CheckCmd ID_嵌入查看, Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "嵌入查看", "0")) = 1
    CheckCmd ID_自动换行, Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "自动换行", "0")) = 1
    
    mdtBegin = CDate(Format(zlDatabase.Currentdate - 365, "yyyy-mm-dd 00:00:00"))
    mdtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
    Exit Sub
errHandle:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    
    cbrMain.RecalcLayout
    cbrMain.GetClientRect lngLeft, lngTop, lngRight, lngBottom
    
    vsfStudy.Left = 0
    vsfStudy.Top = lngTop
    vsfStudy.Width = Me.ScaleWidth
    vsfStudy.Height = Me.ScaleHeight - vsfStudy.Top
End Sub

Private Sub GridInit(strColName As String)
On Error GoTo errH
    '初始化配置列表
    Dim i As Integer
    Dim lngCount As Long
    Dim arrData() As String
    
    arrData = Split(strColName, ";")
    lngCount = UBound(arrData) + 1
    
    With vsfStudy
    
        .Cols = lngCount
        .FixedRows = 1
        .FixedCols = 0
        .RowHeightMin = 320
'        .Cell(flexcpAlignment, 0, 0, 0, lngCount - 1) = flexAlignCenterCenter

        '最后一列自动填充满列表
        .AllowUserResizing = flexResizeColumns
        .ExtendLastCol = True
        .AutoResize = True
        .ExplorerBar = 7 '用于列头拖动和排序
        .AutoSizeMode = flexAutoSizeRowHeight
    
        .WordWrap = True
        .AutoSizeMouse = True
        .SelectionMode = flexSelectionByRow
        .ScrollTrack = True
        
        For i = 0 To lngCount - 1
            .TextMatrix(0, i) = arrData(i)
            .ColKey(i) = arrData(i)
        Next
        
        .Rows = 1
        If .Rows > 1 Then .RowSel = 1
        
        .ColHidden(.ColIndex("医嘱ID")) = True '隐藏医嘱ID
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Public Sub SetFontSize(ByVal bytFontSize As Byte)
    Dim CtlFont As StdFont
    
    Set CtlFont = New StdFont
    CtlFont.Size = bytFontSize
    
    Call SetColWithd(bytFontSize)
    
    vsfStudy.FontSize = bytFontSize
    Set cbrMain.Options.Font = CtlFont
    
    If bytFontSize = 9 Then
        cbrMain.Options.SetIconSize True, 16, 16
    ElseIf bytFontSize = 12 Then
        cbrMain.Options.SetIconSize True, 20, 20
    ElseIf bytFontSize = 15 Then
        cbrMain.Options.SetIconSize True, 24, 24
    End If
    Call Form_Resize
End Sub

Private Sub SetColWithd(ByVal bytSize As Long)

    If mblnListCfgOk Then Exit Sub
    With vsfStudy
        Select Case bytSize
            Case 9
                .ColWidth(.ColIndex("序号")) = 500
                .ColWidth(.ColIndex("部位")) = 1000
                .ColWidth(.ColIndex("当前过程")) = 900
                .ColWidth(.ColIndex("检查号")) = 700
                .ColWidth(.ColIndex("类别")) = 700
                .ColWidth(.ColIndex("年龄")) = 700
                .ColWidth(.ColIndex("检查时间")) = 1600
                .ColWidth(.ColIndex("项目")) = 1200
                .ColWidth(.ColIndex("阴阳性")) = 800
            Case 12
                .ColWidth(.ColIndex("序号")) = 600
                .ColWidth(.ColIndex("部位")) = 1250
                .ColWidth(.ColIndex("当前过程")) = 1100
                .ColWidth(.ColIndex("检查号")) = 900
                .ColWidth(.ColIndex("类别")) = 900
                .ColWidth(.ColIndex("年龄")) = 900
                .ColWidth(.ColIndex("检查时间")) = 2200
                .ColWidth(.ColIndex("项目")) = 1450
                .ColWidth(.ColIndex("阴阳性")) = 1000
            Case 15
                .ColWidth(.ColIndex("序号")) = 700
                .ColWidth(.ColIndex("部位")) = 1500
                .ColWidth(.ColIndex("当前过程")) = 1300
                .ColWidth(.ColIndex("检查号")) = 1100
                .ColWidth(.ColIndex("类别")) = 1100
                .ColWidth(.ColIndex("年龄")) = 1100
                .ColWidth(.ColIndex("检查时间")) = 2800
                .ColWidth(.ColIndex("项目")) = 1700
                .ColWidth(.ColIndex("阴阳性")) = 1200
        End Select
    End With
End Sub

Private Sub vsfStudy_AfterMoveColumn(ByVal Col As Long, Position As Long)
    mTPListCfg.strList = GetListHeadString
    RaiseEvent OnSaveCfg(M_STR_CFG & mTPListCfg.strSort & ";" & mTPListCfg.strList)
End Sub

Private Sub vsfStudy_AfterSort(ByVal Col As Long, Order As Integer)
    Dim strName As String
    Dim i As Integer
    
    For i = 1 To vsfStudy.Rows - 1
        vsfStudy.TextMatrix(i, vsfStudy.ColIndex("序号")) = i
    Next
    
    strName = vsfStudy.TextMatrix(0, Col)
    mTPListCfg.strSort = strName & "," & Order
    
    RaiseEvent OnSaveCfg(M_STR_CFG & mTPListCfg.strSort & ";" & mTPListCfg.strList)
End Sub

Private Sub vsfStudy_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    mTPListCfg.strList = GetListHeadString
    RaiseEvent OnSaveCfg(M_STR_CFG & mTPListCfg.strSort & ";" & mTPListCfg.strList)
End Sub

Private Sub vsfStudy_Click()
    Call vsfStudy_RowColChange
End Sub

Private Sub vsfStudy_DblClick()
    Dim lngRow As Long
    Dim strAdvice As String
    Dim intCol As Integer
    
    On Error GoTo errHandle
    
    lngRow = vsfStudy.MouseRow
    intCol = vsfStudy.ColIndex("医嘱ID")

    If lngRow = 0 Or intCol = -1 Then Exit Sub
    
    If IsCheck(ID_嵌入查看) Then Exit Sub
    
    If lngRow <= 0 Then Exit Sub
    If Val(vsfStudy.TextMatrix(lngRow, intCol)) <= 0 Then Exit Sub
    
    RaiseEvent OnViewReport(Val(vsfStudy.TextMatrix(lngRow, intCol)))
    
    Exit Sub
errHandle:
    MsgBox err.Description, vbExclamation, "提示"
    err.Clear
End Sub

Private Sub vsfStudy_LostFocus()
    RaiseEvent OnListLostFocus
End Sub

Private Sub vsfStudy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OnListMove
End Sub

Private Sub vsfStudy_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errH
    Dim Popup As CommandBar
    Dim Control As CommandBarControl
    Dim intCol As Integer
    
    Dim lngID As Long
    Dim lngRow As Long
    
    
    If Button = 2 Then
        If IsCheck(ID_嵌入查看) Then Exit Sub
        Set Popup = cbrMain.Add("右键菜单", xtpBarPopup)
        With Popup.Controls
            Set Control = .Add(xtpControlButton, conMenu_Img_Look, "观片"): Control.IconId = 5: Control.Category = "观片"
            Set Control = .Add(xtpControlButton, conMenu_Img_Contrast, "对比"): Control.IconId = 6: Control.Category = "对比"
            Set Control = .Add(xtpControlButton, conMenu_PacsReport_Open, "查看报告"): Control.IconId = 7: Control.Category = "查看报告"
        End With
        
        Call Popup.ShowPopup
        
    ElseIf Button = 1 Then
        intCol = vsfStudy.ColIndex("医嘱ID")
    
        If intCol = -1 Then Exit Sub
    
        lngRow = vsfStudy.MouseRow
        If lngRow > 0 Then
            lngID = Val(vsfStudy.TextMatrix(lngRow, intCol))
        End If
        
        If vsfStudy.MouseRow > 0 Then
            RaiseEvent OnListMouseClick(lngID, X, Y, False)
        Else
            RaiseEvent OnListMouseClick(0, X, Y, True)
        End If
    End If
errH:
End Sub

Private Sub vsfStudy_RowColChange()
    Dim lngRow As Long
    Dim i As Long
    Dim intCol As Integer
    Dim strAdvice As String
    
    On Error GoTo errHandle
    
    lngRow = vsfStudy.Row
    intCol = vsfStudy.ColIndex("医嘱ID")
    
    If lngRow <= 0 Or intCol = -1 Then Exit Sub
    
    If mlngRow = lngRow Then Exit Sub
    If Val(vsfStudy.TextMatrix(lngRow, intCol)) <= 0 Then Exit Sub
    mlngRow = lngRow
    
    For i = 1 To vsfStudy.Rows - 1
        If Val(vsfStudy.TextMatrix(i, intCol)) > 0 Then
            strAdvice = strAdvice & IIf(Len(strAdvice) = 0, "", "|") & vsfStudy.TextMatrix(i, intCol)
        End If
    Next
    
    RaiseEvent OnSelectStudy(Val(vsfStudy.TextMatrix(lngRow, intCol)), strAdvice, IsCheck(ID_嵌入查看))
    
    Exit Sub
errHandle:
    MsgBox err.Description, vbExclamation, "提示"
    err.Clear
End Sub

Public Sub ClearData()
'清空数据
    vsfStudy.Rows = 1
    
    mLngAdvice = 0
End Sub

Private Sub InitCommandBars()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl, cbrPopControl As CommandBarControl
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    Set Me.cbrMain.Icons = imgList.Icons
    
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True  '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .UseSharedImageList = False
    End With
    
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.Visible = False
    
    '工具栏定义
    Set cbrToolBar = Me.cbrMain.Add("工具栏", xtpBarTop)
    cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = True
    cbrToolBar.Closeable = False
    cbrToolBar.ContextMenuPresent = False
    
    With cbrToolBar.Controls
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Look, "观片"): cbrControl.IconId = 5: cbrControl.Category = "观片": cbrControl.ToolTipText = "观片"
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Contrast, "对比"): cbrControl.IconId = 6: cbrControl.Category = "对比": cbrControl.ToolTipText = "对比"
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_Open, "报告"): cbrControl.IconId = 7: cbrControl.Category = "查看报告": cbrControl.ToolTipText = "查看报告"
    
        '时间.........................................................
        Set cbrControl = .Add(xtpControlButtonPopup, ID_时间范围, "日期"): cbrControl.IconId = 4: cbrControl.ToolTipText = "日期"
        cbrControl.BeginGroup = True

        Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, ID_一月, "一月"): objControl.IconId = 8
        Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, ID_两月, "两月"): objControl.IconId = 8
        Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, ID_三月, "三月"): objControl.IconId = 8
        Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, ID_半年, "半年"): objControl.IconId = 8
        Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, ID_一年, "一年"): objControl.IconId = 9: objControl.Category = 1
        Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, ID_两年, "两年"): objControl.IconId = 8
        Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, ID_三年, "三年"): objControl.IconId = 8
        Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, ID_不限, "不限"): objControl.IconId = 8
        Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, ID_自定义, "自定义"): objControl.IconId = 8
        
        For Each objControl In cbrControl.CommandBar.Controls
            objControl.CloseSubMenuOnClick = False
        Next
        
        Set cbrControl = .Add(xtpControlButtonPopup, ID_相关设置, "选项"): cbrControl.IconId = 3: cbrControl.ToolTipText = "选项"
        
        
        Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, ID_本次相关, "本次相关"): cbrPopControl.IconId = 1
        Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, ID_他科检查, "他科检查"): cbrPopControl.IconId = 1
        Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, ID_嵌入查看, "嵌入查看"): cbrPopControl.IconId = 1
        Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, ID_自动换行, "自动换行"): cbrPopControl.IconId = 1

        For Each cbrPopControl In cbrControl.CommandBar.Controls
            cbrPopControl.CloseSubMenuOnClick = False
        Next
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
End Sub

Private Function IsCheck(ByVal fltID As FilterID) As Boolean
    Dim objControl As CommandBarControl
    Dim obj As CommandBarControl
    
    If fltID > 10 And fltID < 20 Then
        Set objControl = cbrMain.FindControl(, ID_相关设置)
    ElseIf fltID > 20 And fltID < 30 Then
        Set objControl = cbrMain.FindControl(, ID_时间范围)
    End If
    Set obj = objControl.CommandBar.FindControl(, fltID)
    IsCheck = Val(obj.Category) = 1
End Function


Private Sub CheckCmd(ByVal fltID As FilterID, ByVal blnCheck As Boolean)
    Dim objControl As CommandBarControl
    Dim obj As CommandBarControl
    Dim i As Long
    
    Select Case fltID
        Case ID_本次相关, ID_嵌入查看, ID_他科检查, ID_自动换行
            Set objControl = cbrMain.FindControl(, ID_相关设置)
            Set obj = objControl.CommandBar.FindControl(, fltID)
            obj.IconId = IIf(blnCheck, 2, 1)
            obj.Category = IIf(blnCheck, 1, 0)
        Case ID_半年, ID_两年, ID_两月, ID_三年, ID_三月, ID_一年, ID_一月, ID_自定义, ID_不限
            If blnCheck Then
                Set objControl = cbrMain.FindControl(, ID_时间范围)
                For i = 21 To 29
                    
                    Set obj = objControl.CommandBar.FindControl(, i)
                    If fltID <> i Then
                        obj.IconId = 8
                        obj.Category = 0
                    Else
                        obj.IconId = 9
                        obj.Category = 1
                    End If
                Next
            End If
    End Select
End Sub

Private Sub SetVisible(ByVal fltID As FilterID, ByVal blnVisible As Boolean, ByVal blnEnabled As Boolean)
    Dim objControl As CommandBarControl
    Dim obj As CommandBarControl
    
    If fltID > 10 And fltID < 20 Then
        Set objControl = cbrMain.FindControl(, ID_相关设置)
    ElseIf fltID > 20 And fltID < 30 Then
        Set objControl = cbrMain.FindControl(, ID_时间范围)
    End If
    Set obj = objControl.CommandBar.FindControl(, fltID)
    obj.Visible = blnVisible
    obj.Enabled = blnEnabled
End Sub

Private Function GetTime() As String
    Dim objControl As CommandBarControl
    Dim obj As CommandBarControl
    Dim i As Long
    
    Set objControl = cbrMain.FindControl(, ID_时间范围)
    
    For i = 21 To 29
        Set obj = objControl.CommandBar.FindControl(, i)
        If Val(obj.Category) = 1 Then
            GetTime = obj.Caption
            Exit Function
        End If
    Next
    
    GetTime = "一年"
End Function

Private Function IsImageEnable(ByVal LngAdvice As Long) As Boolean
On Error GoTo errH
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "select 检查UID from 影像检查记录 where  医嘱id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询检查UID", LngAdvice)
    
    If rsTemp.RecordCount <= 0 Then Exit Function
    
    IsImageEnable = Len(nvl(rsTemp!检查UID)) > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function IsReportEnable(ByVal LngAdvice As Long) As Boolean
On Error GoTo errH
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "select count(1) 计数 from 病人医嘱报告 where  医嘱id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询报告", LngAdvice)
    
    If rsTemp.RecordCount <= 0 Then Exit Function
    
    IsReportEnable = Val(nvl(rsTemp!计数)) > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub Free()
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "他科检查", IIf(IsCheck(ID_他科检查), 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "本次相关", IIf(IsCheck(ID_本次相关), 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "嵌入查看", IIf(IsCheck(ID_嵌入查看), 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "自动换行", IIf(IsCheck(ID_自动换行), 1, 0)
End Sub

Private Sub vsfStudy_SelChange()
On Error GoTo errHandle
    Dim lngAdviceID As Long
    Dim intCol As Integer
    
    
    intCol = vsfStudy.ColIndex("医嘱ID")
    If intCol = -1 Then Exit Sub
    
    lngAdviceID = Val(vsfStudy.TextMatrix(vsfStudy.Row, intCol))
    
    mblnReportEnable = IsReportEnable(lngAdviceID)
    mblnImageEnable = IsImageEnable(lngAdviceID)
Exit Sub
errHandle:
    MsgBox err.Description, vbOKOnly, gstrSysName
End Sub

Private Function GetListHeadString() As String
'得到列名参数: 名称,宽度,是否显示  例如  "类别,1000,1|执行过程,2000,0|"
On Error GoTo errH
    Dim i As Integer
    Dim strTemp As String
    Dim strName As String
    Dim lngWidth As Long
    Dim blnShow As Boolean
    
    For i = 0 To vsfStudy.Cols - 1
        
        strName = vsfStudy.TextMatrix(0, i)
        lngWidth = vsfStudy.ColWidth(i)
        blnShow = Not vsfStudy.ColHidden(i)
        
        If Len(strTemp) > 0 Then
            strTemp = strTemp & "|"
        End If
        
        strTemp = strTemp & strName & "," & lngWidth & "," & blnShow
    Next

    GetListHeadString = strTemp
    
    Exit Function
errH:
    err.Raise -1, "历史检查", "[获取列头配置]" & vbCrLf & err.Description
    Resume
End Function

Private Sub DoLoadListCfg(ByVal strcfg As String)
'恢复列顺序和宽度
On Error GoTo errH
    Dim i As Integer, j As Integer
    Dim strName As String
    Dim lngW As Long
    Dim strCol() As String
    Dim intubound As Integer
    Dim blnHide As Boolean
    
    strCol = Split(strcfg, "|")
    intubound = UBound(strCol)
    
    With vsfStudy
        For i = 0 To intubound - 1
            strName = Split(strCol(i), ",")(0)
            lngW = Split(strCol(i), ",")(1)
            blnHide = Not Split(strCol(i), ",")(2)

            If strName <> .TextMatrix(0, i + 1) Then
                For j = 0 To .Cols - 1
                    If strName = .TextMatrix(0, j) Then
                        .ColPosition(j) = i
                        .ColWidth(i) = lngW
                        .ColHidden(i) = blnHide
                        Exit For
                    End If
                Next
            Else
                .ColWidth(i) = lngW
                .ColHidden(i) = blnHide
            End If

        Next
    End With
    
    Exit Sub
errH:
    err.Raise -1, "列表个性化设置", "[DoLoadListCfg]" & vbCrLf & err.Description
    Resume
End Sub

Private Sub DoLoadListSort(ByVal strcfg As String)
'恢复排序
On Error GoTo errH
    Dim strName As String
    Dim intWay As Integer
    Dim intPos As Integer
    Dim intCol As Integer
    Dim i As Integer
    
    intPos = InStr(strcfg, ",")
    If intPos = 0 Then Exit Sub
    
    strName = Split(strcfg, ",")(0)
    intWay = Val(Split(strcfg, ",")(1))
    
    With vsfStudy
        For i = 1 To .Cols - 1
            If strName = .TextMatrix(0, i) Then
                intCol = i
                Exit For
            End If
        Next
         
        .Col = intCol
        .Sort = intWay
        
        For i = 1 To .Rows - 1
            .TextMatrix(i, vsfStudy.ColIndex("序号")) = i
        Next
    End With
    
    Exit Sub
errH:
    err.Raise -1, "列表个性化设置", "[DoLoadListSort]" & vbCrLf & err.Description
    Resume
End Sub



