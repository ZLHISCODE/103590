VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPublicTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "西医诊断"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11640
   Icon            =   "frmPublicTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   11640
   Begin VSFlex8Ctl.VSFlexGrid vsTable 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      _cx             =   20558
      _cy             =   5530
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
      BackColorSel    =   16777215
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   9
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   325
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
End
Attribute VB_Name = "frmPublicTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mintType As Integer
Public mfrmParent As Form
Public mlng病人ID As Long, mlng主页ID As Long
Public mlngLeft As Long, mlngTop As Long, mlngHeight As Long
Public mrsTmp As New ADODB.Recordset

Public Function ShowMe(ByVal intType As Integer, ByVal lng病人ID As Long, ByVal lng主页ID As Long, frmParent As Form, ByVal x As Long, ByVal Y As Long, ByVal lngHeight As Long) As Boolean
'返回：ShowMe= 是确定还是取消
'参数:intType 1-西医诊断 2-中医诊断 3-手术记录
'     frmParent 父窗体
    mintType = intType
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mlngLeft = x
    mlngTop = Y
    mlngHeight = lngHeight
    Set mfrmParent = frmParent
    Set mrsTmp = LoadTableData(mintType)
    If mlng病人ID = 0 And mlng主页ID = 0 Then
        MsgBox "您还没有选择病人，不存在记录！", vbInformation, gstrSysName
        Exit Function
    Else
        If Not mrsTmp Is Nothing Then
            If mrsTmp.RecordCount < 1 Then
                If mintType = 1 Then
                    MsgBox "没有找到记录来源等于住院首页的医生西医诊断记录！", vbInformation, gstrSysName
                ElseIf mintType = 2 Then
                    MsgBox "没有找到记录来源等于住院首页的医生中医诊断记录！", vbInformation, gstrSysName
                ElseIf mintType = 3 Then
                    MsgBox "没有找到记录来源等于住院首页的医生手术记录记录！", vbInformation, gstrSysName
                End If
                Exit Function
            Else
                Show 0, frmParent
            End If
        Else
            MsgBox "该病人不存在记录来源等于住院首页" & IIf(mintType = 1 Or mintType = 1 = 2, "诊断记录", "手术记录"), , vbInformation, gstrSysName
            Exit Function
        End If
    End If
    ShowMe = True
End Function

Private Function InitTable(ByVal intType As Integer) As Boolean
    Dim strHead As String
    Dim strRow As String
On Error GoTo errH
    Select Case intType
        Case 1
            strHead = "诊断类型设置宽,1250,4;关联;诊断编码,900,4;诊断描述,3200,1;中医证候;发病时间;备注,1200,1;入院病情,850,1;出院情况,850,1;ICD附码,800,1;未治,350,4;疑诊,350,4;" & _
                                        ",270,4;,270,4;诊断ID;疾病ID;证候ID;医嘱IDs;诊断分类;固定附码;是否病人;疗效限制;分娩信息;附码ID;诊断来源;疾病编码;疾病类别;证候编码;记录日期;记录人员"
            strRow = DI_诊断类型 & ",门（急）诊诊断," & DI_诊断分类 & "," & DT_门诊诊断XY & ";" & _
                                    DI_诊断类型 & ",入院诊断," & DI_诊断分类 & "," & DT_入院诊断XY & ";" & _
                                    DI_诊断类型 & ",出院诊断," & DI_诊断分类 & "," & DT_出院诊断XY & ";" & _
                                    DI_诊断类型 & ",其他诊断," & DI_诊断分类 & "," & DT_出院诊断XY & ";" & _
                                    DI_诊断类型 & ",院内感染," & DI_诊断分类 & "," & DT_院内感染 & ";" & _
                                    DI_诊断类型 & ", 并 发 症 ," & DI_诊断分类 & "," & DT_并发症 & ";" & _
                                    DI_诊断类型 & ",病理诊断," & DI_诊断分类 & "," & DT_病理诊断 & ";" & _
                                    DI_诊断类型 & ",损伤中毒," & DI_诊断分类 & "," & DT_损伤中毒码
        Case 2
            strHead = "诊断类型设置宽,1250,4;关联;诊断编码,900,4;诊断描述,3000,1;中医证候,1500,1;发病时间;备注,1100,1;入院病情,850,1;出院情况,850,1;ICD附码;未治;疑诊,350,4;" & _
                                        ",270,4;,270,4;诊断ID;疾病ID;证候ID;医嘱IDs;诊断分类;固定附码;是否病人;疗效限制;分娩信息;附码ID;诊断来源;疾病编码;疾病类别;证候编码;记录日期;记录人员"
            strRow = DI_诊断类型 & ",门（急）诊诊断," & DI_诊断分类 & "," & DT_门诊诊断ZY & ";" & _
                                    DI_诊断类型 & ",入院诊断," & DI_诊断分类 & "," & DT_入院诊断ZY & ";" & _
                                    DI_诊断类型 & ",出院诊断," & DI_诊断分类 & "," & DT_出院诊断ZY & ";" & _
                                    DI_诊断类型 & ",其他诊断," & DI_诊断分类 & "," & DT_出院诊断ZY
        Case 3
            If gclsPros.MedPageSandard = ST_卫生部标准 Then
                strHead = ",300,4;" & IIf(gclsPros.UseOPSEndTime, "手术开始时间,1850,4;手术结束时间,1850,4", "手术及操作日期,1850,4;手术结束时间") & ";术前预防性抗菌用药时间;手术情况,875,1;准备天数;手术及操作编码,1500,1;手术及操作名称,2800,1;再次手术,850,4,11;术者,850,1;助产护士,850,1;第Ⅰ助手,850,1;第Ⅱ助手,850,1;" & _
                                "麻醉开始时间;麻醉方式,850,1;ASA分级,850,1;NNIS分级,850,1;手术级别,850,1;麻醉医师,850,1;切口愈合等级,1400,1;切口部位,850,1;重返手术室计划;重返手术室目的;切口感染;并发症;" & _
                                "术前0.5-2小时预防用抗菌药;清洁手术围术期预防用抗菌药天数;非预期的二次手术,1600,4,11;麻醉并发症;术中异物遗留;手术并发症;术后出血或血肿;手术伤口裂开;术后深静脉血栓;术后生理/代谢紊乱;术后呼吸衰竭;" & _
                                "术后肺栓塞;术后败血症;术后髋关节骨折;手术操作ID;诊疗项目ID;麻醉ID;麻醉类型;手麻来源"
            ElseIf gclsPros.MedPageSandard = ST_湖南省标准 Then
                strHead = ",300,4;" & IIf(gclsPros.UseOPSEndTime, "手术开始时间,1850,4;手术结束时间,1850,4", "手术及操作日期,1850,4;手术结束时间") & ";术前预防性抗菌用药时间;手术情况,875,1;准备天数;手术及操作编码,1500,1;手术及操作名称,2800,1;再次手术,850,4,11;术者,850,1;助产护士,850,1;第Ⅰ助手,850,1;第Ⅱ助手,850,1;" & _
                                "麻醉开始时间;麻醉方式,850,1;ASA分级,850,1;NNIS分级,850,1;手术级别,850,1;麻醉医师,850,1;切口愈合等级,1400,1;切口部位;重返手术室计划;重返手术室目的;切口感染;并发症;" & _
                                "术前0.5-2小时预防用抗菌药;清洁手术围术期预防用抗菌药天数;非预期的二次手术;麻醉并发症;术中异物遗留;手术并发症;术后出血或血肿;手术伤口裂开;术后深静脉血栓;术后生理/代谢紊乱;术后呼吸衰竭;" & _
                                "术后肺栓塞;术后败血症;术后髋关节骨折;手术操作ID;诊疗项目ID;麻醉ID;麻醉类型;手麻来源"
            ElseIf gclsPros.MedPageSandard = ST_四川省标准 Then
                strHead = ",300,4;" & "开始日期,1850,4;结束日期,1850,4;术前预防性抗菌用药时间,2150,4;手术情况,875,1;准备天数,850,7;手术编码,1500,1;手术名称,2800,1;再次手术,850,4,11;主刀医师,850,1;助产护士,850,1;第Ⅰ助手,850,1;第Ⅱ助手,850,1;" & _
                                "麻醉开始时间,1550,4;麻醉方式,850,1;ASA分级,850,1;NNIS分级,850,1;手术分级,850,1;麻醉医师,850,1;切口/愈合,1400,1;切口部位,850,1;重返手术室计划,1400,4,11;重返手术室目的,1400,1;切口感染,850,4,11;并发症,720,4,11;" & _
                                "术前0.5-2小时预防用抗菌药;清洁手术围术期预防用抗菌药天数;非预期的二次手术;麻醉并发症;术中异物遗留;手术并发症;术后出血或血肿;手术伤口裂开;术后深静脉血栓;术后生理/代谢紊乱;术后呼吸衰竭;" & _
                                "术后肺栓塞;术后败血症;术后髋关节骨折;手术操作ID;诊疗项目ID;麻醉ID;麻醉类型;手麻来源"
            ElseIf gclsPros.MedPageSandard = ST_云南省标准 Then
                strHead = ",300,4;" & "手术日期,1850,4;结束日期;术前预防性抗菌用药时间;手术情况,875,1;准备天数;手术编码,1500,1;手术名称,2800,1;再次手术,850,4,11;主刀医师,850,1;助产护士,850,1;第Ⅰ助手,850,1;第Ⅱ助手,850,1;" & _
                                "麻醉开始时间;麻醉方式,850,1;ASA分级,850,1;NNIS分级,850,1;手术分级,850,1;麻醉医师,850,1;切口/愈合,1400,1;切口部位;重返手术室计划;重返手术室目的;切口感染;并发症;" & _
                                "术前0.5-2小时预防用抗菌药,2400,4,11;清洁手术围术期预防用抗菌药天数,2850,7;非预期的二次手术,1600,4,11;麻醉并发症,1000,4,11;术中异物遗留,1200,4,11;手术并发症,1000,4,11;" & _
                                "术后出血或血肿,1450,4,11;手术伤口裂开,1200,4,11;术后深静脉血栓,1450,4,11;术后生理/代谢紊乱,1700,4,11;术后呼吸衰竭,1200,4,11;术后肺栓塞,1000,4,11;术后败血症,1000,4,11;" & _
                                "术后髋关节骨折,1450,4,11;手术操作ID;诊疗项目ID;麻醉ID;麻醉类型;手麻来源"
            End If
    End Select
    If Not setTableType(intType, strHead, strRow) Then Exit Function
    InitTable = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function setTableType(ByVal intType As Integer, Optional ByVal strHead As String, Optional ByVal strRow As String) As Boolean
    Dim vsTmp As VSFlexGrid
    Dim strTmp As String
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
On Error GoTo errH
    Select Case intType
        Case 1
            Set vsTmp = vsTable
            Call Grid.Init(vsTable, strHead, strRow, 1, 1)
            With vsTmp
                If gclsPros.FuncType <> f电子病案 Then
                    If Not .ColHidden(DI_入院病情) Then .ColData(DI_出院情况) = "有|临床未确定|情况不明|无"
                    If Not .ColHidden(DI_出院情况) Then
                        Set rsTmp = GetBaseCode("治疗结果")
                        If Not rsTmp.EOF Then
                            strTmp = Rec.ToComboList(rsTmp, "[0]-[1]|", "编码", "名称")
                            '用Chr(10)代替空白项是为了实现发送空格弹出下拉列表
                            .ColData(DI_出院情况) = Chr(10) & "|" & strTmp
                        Else
                            .ColData(DI_出院情况) = Chr(10) & "|1-治愈|2-好转|3-未愈|4-死亡|5-其他"
                        End If
                    End If
                End If
                If .Font.Size <> gclsPros.FontSize Then
                    .Font.Size = gclsPros.FontSize
                    Call Grid.AdjustCols(vsTmp, "," & DI_Del & "," & DI_增加 & ",")
                End If
                If .TextMatrix(0, DI_诊断类型) = "诊断类型设置宽" Then .TextMatrix(0, DI_诊断类型) = "诊断类型" '恢复列头
            End With
        Case 2
            Set vsTmp = vsTable
            Call Grid.Init(vsTable, strHead, strRow, 1, 1)
            With vsTmp
                If gclsPros.FuncType <> f电子病案 Then
                    If Not .ColHidden(DI_入院病情) Then .ColData(DI_出院情况) = "有|临床未确定|情况不明|无"
                    If Not .ColHidden(DI_出院情况) Then
                        If strTmp <> "" Then
                            '用Chr(10)代替空白项是为了实现发送空格弹出下拉列表
                            .ColData(DI_出院情况) = Chr(10) & "|" & strTmp
                        Else
                            .ColData(DI_出院情况) = Chr(10) & "|1-治愈|2-好转|3-未愈|4-死亡|5-其他"
                        End If
                    End If
                End If
                  If .Font.Size <> gclsPros.FontSize Then
                     .Font.Size = gclsPros.FontSize
                    Call Grid.AdjustCols(vsTmp, "," & DI_Del & "," & DI_增加 & ",")
                  End If
                If .TextMatrix(0, DI_诊断类型) = "诊断类型设置宽" Then .TextMatrix(0, DI_诊断类型) = "诊断类型" '恢复列头
            End With
        Case 3
            Set vsTmp = vsTable
            Call Grid.Init(vsTmp, strHead)
            With vsTmp
                .Font.Size = 9
                If gclsPros.FuncType <> f电子病案 Then
                    .ColComboList(PI_手术情况) = " |择期|急诊|限期"
                    .ColComboList(PI_ASA分级) = " |P1|P2|P3|P4|P5|P6"
                    .ColComboList(PI_NNIS分级) = " |NNIS0级|NNIS1级|NNIS2级|NNIS3级"
                    .ColComboList(PI_手术级别) = " |无|一级手术|二级手术|三级手术|四级手术"
                    '切口愈合
                    strSql = "Select Rownum As ID, To_Number(编码) As 编码, 编码 简码, 名称, 0 缺省 From 手术切口愈合 Order By 编码"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "手术切口愈合")
                    If Not rsTmp.EOF Then
                        strTmp = " |" & Rec.ToComboList(rsTmp, "[0]-[1]|", "编码", "名称")
                    Else
                        strTmp = " |0-0 / |1-Ⅰ/甲|2-Ⅰ/乙|3-Ⅰ/丙|4-Ⅰ/其他|5-Ⅱ/甲|6-Ⅱ/乙|7-Ⅱ/丙|8-Ⅱ/其他|9-Ⅲ/甲|10-Ⅲ/乙|11-Ⅲ/丙|12-Ⅲ/其他|13-IV/甲|14-IV/乙|15-IV/丙|16-IV/其他"
                    End If
                    .ColData(PI_切口愈合) = strTmp
                    '麻醉类型
                    Set rsTmp = GetBaseCode("诊疗麻醉类型")
                    If Not rsTmp.EOF Then
                        strTmp = " |" & Rec.ToComboList(rsTmp, "[0]-[1]|", "简码", "名称")
                    Else
                        strTmp = " |JM-局麻|QM-全麻|CY-持硬|QT-其他|JM-静脉|BC-臂丛|JC-颈丛"
                    End If
                    .ColData(PI_麻醉类型) = strTmp
                End If
                If gclsPros.FontSize <> 9 Then Call zlControl.VSFSetFontSize(vsTmp, gclsPros.FontSize)
            End With
    End Select
    setTableType = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function LoadTableData(ByVal intType As Integer) As ADODB.Recordset
    Dim strSql As String, strSQLTmp As String, strDiagType As String, strSQLJudge As String
    Dim rsTmp As New ADODB.Recordset
    Dim int记录来源 As Integer
On Error GoTo errH
    Select Case intType
        Case 1, 2
            If gclsPros.FuncType = f病案首页 Then
               int记录来源 = 3
               strSQLJudge = "Select 1 From 病人诊断记录 Where 病人id = [1] And 主页id =[2] And 记录来源 = [3] And Rownum < 2"
               Set rsTmp = zlDatabase.OpenSQLRecord(strSQLJudge, "首页来源诊断判断", mlng病人ID, mlng主页ID, int记录来源)
               If rsTmp.RecordCount > 0 Then
                   strDiagType = " And A.记录来源 =[3] "
               Else
                   strDiagType = ""
                   Set LoadTableData = rsTmp
                   Exit Function
               End If
            End If
            If intType = 1 Then
                strDiagType = strDiagType & " And A.诊断类型 IN(1,2,3,5,6,7,10,21) "
            Else
                strDiagType = strDiagType & " And A.诊断类型 IN(1,2,3,5,6,7,10,11,12,13,21) "
            End If
            strSql = "Select A.备注, A.Id, A.病人id, A.主页id, A.医嘱id, A.记录来源, A.诊断次序, Nvl(A.编码序号,1) 编码序号, A.诊断类型, A.入院病情, A.疾病id, A.诊断id, A.证候id,B.名称 疾病名称,C.名称 诊断名称,D.名称 证候名称," & vbNewLine & _
                "       A.诊断描述, A.出院情况, A.是否未治, A.是否疑诊, A.发病时间, B.编码 As 疾病编码,B.类别 As 疾病类别, B.附码, C.编码 As 诊断编码, D.编码 As 证候编码," & vbNewLine & _
                IIf(gclsPros.FuncType = f电子病案, " Null 医嘱id", " (Select F_List2str(Cast(Collect(C.医嘱id || '') As T_Strlist)) 医嘱id" & vbNewLine & _
                "         From 病人诊断医嘱 C,病人医嘱记录 F " & vbNewLine & _
                "         Where C.医嘱ID = F.ID and C.诊断id = A.Id and nvl(F.申请序号,0) = 0) As 医嘱id") & ",B.性别限制, B.疗效限制, B.分娩, B.附码, E.Id As 大类, E.是否病人,Null 附码ID,A.记录日期,A.记录人 " & vbNewLine & _
                "From 病人诊断记录 A, 疾病编码目录 B, 疾病诊断目录 C, 疾病编码目录 D,疾病编码分类 E" & vbNewLine & _
                "Where A.疾病id = B.Id(+) And A.诊断id = C.Id(+) And A.证候id = D.Id(+)  And  B.分类id = E.Id(+)" & strDiagType & "And A.取消时间 Is Null And A.诊断描述 Is Not Null And 病人id = [1] And 主页id =[2]" & vbNewLine & _
                "Order By A.诊断类型, A.记录来源 Desc, A.诊断次序, Nvl(A.编码序号,1), A.Id"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取首页诊断", mlng病人ID, mlng主页ID, int记录来源)
            Set LoadTableData = rsTmp
        Case 3
            strSql = "Select A.Id, A.病人id, A.主页id, A.手术情况, A.记录来源, A.手术日期, A.手术开始时间, A.手术结束时间, Nvl(B.编码, C.编码) As 手术编码, A.已行手术 手术名称," & vbNewLine & _
                "       Nvl(B.名称, C.名称) 手术原名, A.主刀医师, A.助产护士, A.第一助手, A.第二助手, A.麻醉医师, A.准备天数, A.抗菌用药时间, A.抗菌用药天数, A.麻醉开始时间, A.重返目的," & vbNewLine & _
                "       A.切口部位, A.麻醉类型, Decode(A.Asa分级, 'I级', 'P1', 'II级', 'P2', 'III级', 'P3', 'IV级', 'P4', 'V级', 'P5', A.Asa分级) Asa分级, A.Nnis分级, Decode(A.手术级别, 1, '一级手术', 2, '二级手术', 3, '三级手术', 4, '四级手术',9, '无', ' ') As 手术级别, A.切口," & vbNewLine & _
                "       A.愈合, A.再次手术, A.术前抗菌用药, A.非预期的二次手术, A.麻醉并发症, A.术中异物遗留, A.手术并发症, A.术后出血或血肿, A.手术伤口裂开, A.术后深静脉血栓, A.术后生理代谢紊乱," & vbNewLine & _
                "       A.术后呼吸衰竭, A.术后肺栓塞, A.术后败血症, A.术后髋关节骨折, A.重返计划, A.切口感染, A.并发症, A.手术操作id, A.诊疗项目id, A.麻醉方式 麻醉id, D.名称 麻醉方式, A.记录日期," & vbNewLine & _
                "       A.记录人, A.取消时间, A.取消人, Decode(B.手术类型, '甲', '四级手术', '乙', '三级手术', '丙', '二级手术', '丁', '一级手术', '四级', '四级手术', '三级', '三级手术', '二级', '二级手术', '一级', '一级手术', Null) 原手术级别 " & vbNewLine & _
                "From 病人手麻记录 A, 疾病编码目录 B, 诊疗项目目录 C, 诊疗项目目录 D" & vbNewLine & _
                "Where C.Id(+) = A.诊疗项目id And A.手术操作id = B.Id(+) And A.麻醉方式 = D.Id(+) And 病人id = [1] And 主页id = [2] And" & vbNewLine & _
                "      (记录来源 <> 1 Or" & vbNewLine & _
                "       (记录来源 = 1 And 取消时间 Is Null And" & vbNewLine & _
                "       记录日期 =" & vbNewLine & _
                "       (Select Max(记录日期) From 病人手麻记录 Where 病人id =[1] And 主页id = [2] And 取消时间 Is Null)))" & vbNewLine & _
                "Order By Nvl(A.手术次序,999),A.ID"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取病人手麻信息", mlng病人ID, mlng主页ID)
            rsTmp.Filter = "记录来源=3"
            Set LoadTableData = rsTmp
        End Select
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function LoadVsDiagData(ByRef vsTable As VSFlexGrid, ByVal rsInput As ADODB.Recordset, ByVal strDiagType As String)
    Dim strTmp As String
    Dim arrTmp As Variant
    Dim i As Long, j As Long, k As Long, LngRow As Long
    Dim bln分化程度 As Boolean
    Dim bln西医 As Boolean
    Dim lngPos As Long
    Dim strInfo As String, strMainInfo As String
    Dim arrWhole As Variant, arrMain As Variant
    Dim blnFreeDiag As Boolean
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim blnGet附码 As Boolean
    
On Error GoTo errH
    blnGet附码 = gclsPros.GetExtraCode
    arrTmp = Split(strDiagType, ",")
    bln西医 = mintType = 1
    With vsTable
        For i = LBound(arrTmp) To UBound(arrTmp)
            Call FilterDiagByType(rsInput, Val(arrTmp(i)), -1) '过滤诊断
            Do While Not rsInput.EOF
                If rsInput!编码序号 = 1 Then
                    '确定当前显示行
                    LngRow = .FindRow(arrTmp(i), , DI_诊断分类, , True)
                    For j = LngRow To .Rows - 1
                        If Val(.TextMatrix(j, DI_诊断分类)) = Val(arrTmp(i)) Then
                            LngRow = j
                            If .TextMatrix(j, DI_诊断描述) = "" Then Exit For
                        Else
                            Exit For
                        End If
                    Next
    
                    '新增行
                    If .TextMatrix(LngRow, DI_诊断描述) <> "" Then
                        LngRow = LngRow + 1: .AddItem "", LngRow
                        .TextMatrix(LngRow, DI_诊断分类) = arrTmp(i)
                                                    If .TextMatrix(LngRow, DI_诊断类型) <> "出院诊断" And Val(.TextMatrix(LngRow, DI_诊断分类)) = 3 Then
                            .Cell(flexcpData, LngRow, DI_诊断类型) = "其他诊断"
                        End If
                    End If
    
                    If gclsPros.FuncType = f诊断选择 Then
                        If InStr("," & gclsPros.DiagRowIDs & ",", "," & rsInput!ID & ",") > 0 Then
                            .TextMatrix(LngRow, DI_关联) = 1
                        End If
                    End If
    
                    strTmp = rsInput!诊断描述 & ""
                    '读取诊断编码，诊断描述为(编码)描述，或(编码)描述(证候) 类型的可以获取诊断描述
                    If strTmp Like "(?*)?*" Then
                        lngPos = InStr(1, strTmp, ")")
                        .TextMatrix(LngRow, DI_诊断编码) = Mid(strTmp, 2, lngPos - 2)
                        strTmp = Mid(strTmp, lngPos + 1)
                    End If
                    If .TextMatrix(LngRow, DI_诊断编码) = "" And Not (IsNull(rsInput!诊断ID) And IsNull(rsInput!疾病id)) Then
                        '由于疾病编码和诊断可以对应，如果两个都不为空的时候，先判断疾病编码，先取疾病编码
                        .TextMatrix(LngRow, DI_诊断编码) = IIf(Not IsNull(rsInput!疾病id), rsInput!疾病编码 & "", rsInput!诊断编码 & "")
                    End If
                    '获取中医证候，由于诊断描述可能会增加前后缀，前后缀包含括号，所以反向截取字符串
                    If strTmp Like "?*(?*)" And Not bln西医 Then
                        strTmp = StrReverse(strTmp)
                        lngPos = InStr(1, strTmp, "(")
                        .TextMatrix(LngRow, DI_中医证候) = StrReverse(Mid(strTmp, 2, lngPos - 2))
                        strTmp = StrReverse(Mid(strTmp, lngPos + 1))
                    End If
                    '取诊断描述
                    .TextMatrix(LngRow, DI_诊断描述) = strTmp
                    '诊断描述的备份数据
                    If Not (IsNull(rsInput!诊断ID) And IsNull(rsInput!疾病id)) Then
                        .Cell(flexcpData, LngRow, DI_诊断描述) = IIf(Not IsNull(rsInput!疾病id), rsInput!疾病名称 & "", rsInput!诊断名称 & "")
                    Else
                        .Cell(flexcpData, LngRow, DI_诊断描述) = .TextMatrix(LngRow, DI_诊断描述)
                    End If
                    If Val(rsInput!证候ID & "") <> 0 And .TextMatrix(LngRow, DI_中医证候) = "" Then
                        .TextMatrix(LngRow, DI_中医证候) = rsInput!证候名称 & ""
                    End If
                    .Cell(flexcpData, LngRow, DI_诊断编码) = .TextMatrix(LngRow, DI_诊断编码)
                    .Cell(flexcpData, LngRow, DI_中医证候) = .TextMatrix(LngRow, DI_中医证候)
                    If .TextMatrix(LngRow, DI_诊断描述) <> "" Then
                        .AutoSize DI_诊断编码, DI_诊断描述
                    End If
                    If .ColWidth(DI_诊断描述) < 3200 Then
                        .ColWidth(DI_诊断描述) = 3200
                    End If
                    '其他列数据加
                    .TextMatrix(LngRow, DI_发病时间) = Format(rsInput!发病时间 & "", "YYYY-MM-DD HH:mm")
                    .TextMatrix(LngRow, DI_备注) = rsInput!备注 & ""
                    .TextMatrix(LngRow, DI_出院情况) = rsInput!出院情况 & ""
                    .TextMatrix(LngRow, DI_入院病情) = rsInput!入院病情 & ""
                    If blnGet附码 Then
                        .TextMatrix(LngRow, DI_ICD附码) = rsInput!附码 & ""
                    End If
                    .TextMatrix(LngRow, DI_是否未治) = IIf(Val(rsInput!是否未治 & "") = 1, "√", "")
                    .TextMatrix(LngRow, DI_是否疑诊) = IIf(Val(rsInput!是否疑诊 & "") = 1, "？", "")
                    If gclsPros.FuncType <> f病案首页 Then
                        .TextMatrix(LngRow, DI_诊断ID) = rsInput!诊断ID & ""
                    End If
                    .TextMatrix(LngRow, DI_疾病ID) = rsInput!疾病id & ""
                    .TextMatrix(LngRow, DI_证候ID) = rsInput!证候ID & ""
                    .TextMatrix(LngRow, DI_医嘱IDs) = rsInput!医嘱ID & ""
                    If gclsPros.FuncType = f病案首页 Then
                        If (arrTmp(i) = DT_出院诊断XY Or arrTmp(i) = DT_出院诊断ZY Or arrTmp(i) = DT_院内感染 Or arrTmp(i) = DT_并发症) Then
    '                                .TextMatrix(LngRow, DI_固定附码) = IIf(IsNull(rsInput!附码), "", "1")
                            .TextMatrix(LngRow, DI_是否病人) = IIf(Val(rsInput!是否病人 & "") = 1, "1", "")
                        End If
                    End If
                    .TextMatrix(LngRow, DI_疗效限制) = rsInput!疗效限制 & ""
                    .TextMatrix(LngRow, DI_分娩信息) = IIf(IsNull(rsInput!分娩), "0", "1")
                    .TextMatrix(LngRow, DI_诊断来源) = Val(rsInput!记录来源 & "") '保存记录来源，以便保存时，保存为首页或病案来源
                    .TextMatrix(LngRow, DI_疾病编码) = rsInput!疾病编码 & ""
                    .TextMatrix(LngRow, DI_疾病类别) = rsInput!疾病类别 & ""
                    .TextMatrix(LngRow, DI_证候编码) = rsInput!证候编码 & ""
                    .TextMatrix(LngRow, DI_记录日期) = Format(rsInput!记录日期 & "", "YYYY-MM-DD HH:mm")
                    .TextMatrix(LngRow, DI_记录人员) = rsInput!记录人 & ""
                    .RowData(LngRow) = Val(rsInput!ID & "")
                Else
                    .TextMatrix(LngRow, DI_附码ID) = rsInput!疾病id & ""
                    .TextMatrix(LngRow, DI_ICD附码) = rsInput!疾病编码 & ""
                    .Cell(flexcpData, LngRow, DI_ICD附码) = .TextMatrix(LngRow, DI_ICD附码)
                End If
                rsInput.MoveNext
            Loop
        Next
    End With
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function LoadVsOPSData(ByRef vsOPSInput As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset)
'功能：加载病人手麻数据并缓存
'参数：vsOPSInput=需要加载病人手麻信息的表格
'      rsInput=病人手麻信息记录集
    Dim i As Long, LngRow As Long, j As Long
    Dim strInfo As String, strMainInfo As String
    Dim lngOrder As Long
    Dim strSql As String, rsTmp As ADODB.Recordset

    On Error GoTo errH
    With vsOPSInput
        '数据加载
        If rsInput Is Nothing Then .Rows = .FixedRows + 1: Exit Function
        .Rows = rsInput.RecordCount + 2 '固定行+新行
        For i = 1 To rsInput.RecordCount
            .TextMatrix(i, PI_手术日期) = Format(NVL(rsInput!手术开始时间, rsInput!手术日期) & "", "yyyy-MM-dd HH:mm")
            .TextMatrix(i, PI_结束日期) = Format(NVL(rsInput!手术结束时间, rsInput!手术日期) & "", "yyyy-MM-dd HH:mm")
            .TextMatrix(i, PI_手术编码) = rsInput!手术编码 & ""
            .TextMatrix(i, PI_手术名称) = rsInput!手术名称 & ""
            If (Not gclsPros.CNIndent And gclsPros.FuncType = f病案首页) Or .TextMatrix(i, PI_手术名称) = "" Then
                .TextMatrix(i, PI_手术名称) = rsInput!手术原名 & ""
                If .TextMatrix(i, PI_手术名称) = "" Then
                    .TextMatrix(i, PI_手术名称) = rsInput!手术名称 & ""
                End If
            End If
            If .TextMatrix(i, PI_手术名称) <> "" Then
                .AutoSize PI_手术编码, PI_手术名称
            End If
            .TextMatrix(i, PI_主刀医师) = rsInput!主刀医师 & ""
            .TextMatrix(i, PI_助产护士) = rsInput!助产护士 & ""
            .TextMatrix(i, PI_助手1) = rsInput!第一助手 & ""
            .TextMatrix(i, PI_助手2) = rsInput!第二助手 & ""
            .TextMatrix(i, PI_麻醉方式) = rsInput!麻醉方式 & ""
            .TextMatrix(i, PI_麻醉医师) = rsInput!麻醉医师 & ""
            If rsInput!切口 & rsInput!愈合 & "" <> "" Then
                .TextMatrix(i, PI_切口愈合) = rsInput!切口 & "/" & rsInput!愈合
            End If
            .TextMatrix(i, PI_手术操作ID) = rsInput!手术操作ID & ""
            .TextMatrix(i, PI_诊疗项目ID) = rsInput!诊疗项目id & ""
            .TextMatrix(i, PI_麻醉ID) = rsInput!麻醉ID & ""
            .TextMatrix(i, PI_麻醉类型) = rsInput!麻醉类型 & ""
            .TextMatrix(i, PI_手术情况) = rsInput!手术情况 & ""
            .TextMatrix(i, PI_ASA分级) = rsInput!asa分级 & ""
            .TextMatrix(i, PI_NNIS分级) = rsInput!NNIS分级 & ""
            .TextMatrix(i, PI_手术级别) = rsInput!手术级别 & ""
            .TextMatrix(i, PI_再次手术) = IIf(Val(rsInput!再次手术 & "") = 1, -1, 0)
            .TextMatrix(i, PI_准备天数) = IIf(Val(rsInput!准备天数 & "") = 0, "", Val(rsInput!准备天数 & ""))
            .TextMatrix(i, PI_抗菌用药时间) = Format(rsInput!抗菌用药时间 & "", "yyyy-MM-dd HH:mm")
            .TextMatrix(i, PI_麻醉开始时间) = Format(rsInput!麻醉开始时间 & "", "yyyy-MM-dd HH:mm")
            .TextMatrix(i, PI_切口部位) = rsInput!切口部位 & ""
            .TextMatrix(i, PI_重返手术室目的) = rsInput!重返目的 & ""
            .Cell(flexcpChecked, i, PI_重返手术室计划) = Val(rsInput!重返计划 & "")
            .Cell(flexcpChecked, i, PI_切口感染) = Val(rsInput!切口感染 & "")
            .Cell(flexcpChecked, i, PI_并发症) = Val(rsInput!并发症 & "")
            '10.34.10新增
            .TextMatrix(i, PI_抗菌药天数) = IIf(Val(rsInput!抗菌用药天数 & "") = 0, "", Val(rsInput!抗菌用药天数 & ""))
            .Cell(flexcpChecked, i, PI_预防用抗菌药) = Val(rsInput!术前抗菌用药 & "")
            .Cell(flexcpChecked, i, PI_非预期的二次手术) = Val(rsInput!非预期的二次手术 & "")
            .Cell(flexcpChecked, i, PI_麻醉并发症) = Val(rsInput!麻醉并发症 & "")
            .Cell(flexcpChecked, i, PI_术中异物遗留) = Val(rsInput!术中异物遗留 & "")
            .Cell(flexcpChecked, i, PI_手术并发症) = Val(rsInput!手术并发症 & "")
            .Cell(flexcpChecked, i, PI_术后出血或血肿) = Val(rsInput!术后出血或血肿 & "")
            .Cell(flexcpChecked, i, PI_手术伤口裂开) = Val(rsInput!手术伤口裂开 & "")
            .Cell(flexcpChecked, i, PI_术后深静脉血栓) = Val(rsInput!术后深静脉血栓 & "")
            .Cell(flexcpChecked, i, PI_术后生理代谢紊乱) = Val(rsInput!术后生理代谢紊乱 & "")
            .Cell(flexcpChecked, i, PI_术后呼吸衰竭) = Val(rsInput!术后呼吸衰竭 & "")
            .Cell(flexcpChecked, i, PI_术后肺栓塞) = Val(rsInput!术后肺栓塞 & "")
            .Cell(flexcpChecked, i, PI_术后败血症) = Val(rsInput!术后败血症 & "")
            .Cell(flexcpChecked, i, PI_术后髋关节骨折) = Val(rsInput!术后髋关节骨折 & "")
            .Cell(flexcpData, i, PI_手术名称) = rsInput!手术原名 & ""
            .TextMatrix(i, PI_手麻来源) = rsInput!记录来源 & ""
            .RowData(i) = Val(rsInput!ID & "")
            '记录用于编辑恢复
            For j = 0 To .Cols - 1
                If j = PI_手术名称 And .TextMatrix(i, PI_手术编码) <> "" Then
                    If .Cell(flexcpData, i, j) = "" Then
                        .Cell(flexcpData, i, j) = .TextMatrix(i, j)
                    End If
                Else
                    .Cell(flexcpData, i, j) = .TextMatrix(i, j)
                End If
            Next

            If Trim(.TextMatrix(i, PI_手术级别)) <> "" And rsInput!原手术级别 & "" <> "" Then
                .Cell(flexcpData, i, PI_手术级别) = 1
            End If
            rsInput.MoveNext
        Next
    End With
    Exit Function
errH:
    If ErrCenter() <> 1 Then
        Resume
    End If
End Function

Private Sub Form_Load()
    Dim lngScrH  As Long
    lngScrH = GetSystemMetrics(SM_CYFULLSCREEN) * 15 '屏幕可用高度
    If mlngTop + Me.Height > lngScrH Then
        Me.Top = mlngTop - Me.Height - 300
    Else
        Me.Top = mlngHeight + 1000
    End If
    Me.Left = mlngLeft
    If Not InitTable(mintType) Then Exit Sub
    Select Case mintType
        Case 1, 2
            mrsTmp.Filter = "记录来源=3"
            Call LoadVsDiagData(vsTable, mrsTmp, IIf(mintType = 1, "1,2,3,5,6,7,10", "11,12,13"))
        Case 3
            Call LoadVsOPSData(vsTable, mrsTmp)
    End Select
    If mintType = 1 Then
        Me.Caption = "医生西医诊断"
    ElseIf mintType = 2 Then
        Me.Caption = "医生中医诊断"
    ElseIf mintType = 3 Then
        Me.Caption = "医生手术记录"
    End If
    Exit Sub
End Sub



