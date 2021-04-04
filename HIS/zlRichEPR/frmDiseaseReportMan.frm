VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDiseaseReportMan 
   Caption         =   "疾病申报管理"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10230
   Icon            =   "frmDiseaseReportMan.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   10230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   3330
      Left            =   120
      TabIndex        =   1
      Top             =   630
      Width           =   6660
      _Version        =   589884
      _ExtentX        =   11747
      _ExtentY        =   5874
      _StockProps     =   0
      BorderStyle     =   2
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7065
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDiseaseReportMan.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15161
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgTemp 
      Height          =   900
      Left            =   240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5340
      Visible         =   0   'False
      Width           =   1080
      _cx             =   1905
      _cy             =   1587
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
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
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgInfo 
      Height          =   6300
      Left            =   7020
      TabIndex        =   3
      Top             =   630
      Width           =   3135
      _cx             =   5530
      _cy             =   11112
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483637
      GridColorFixed  =   -2147483638
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
      Rows            =   1
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   1545
      Top             =   5355
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseReportMan.frx":0E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseReportMan.frx":11B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseReportMan.frx":1550
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseReportMan.frx":18EA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   285
      Top             =   75
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmDiseaseReportMan.frx":1C84
      Left            =   960
      Top             =   210
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmDiseaseReportMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------
'常量
'-----------------------------------------------------
Private Enum mCol
    图标 = 0: ID: 状态: 报告: 科室: 就诊号: 姓名: 性别: 年龄: 填报时间: 填报人: 信息: 数据转出: 病人ID: 主页ID: 文件ID: 编辑方式
End Enum
Const conPane_Reports = 1
Const conPane_Preview = 2
Const conPane_AppInfo = 3

Private mobjDoc As cEPRDocument
Private mobjRichEMR As Object
Private mobjInfection As Object
'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mstrPrivs As String             '当前使用者权限串
Private mstrFiles As String             '本机管理的报告文件
Private mintDates As Integer            '默认查看最近记录的天数；为0时，说明进入程序执行参数设置，要求按日期范围查看
Private mstrDateFrom As String          '按范围查看开始日期，在mintDates=0时有效
Private mstrDateTo As String            '按范围查看截止日期，在mintDates=0时有效

Private mfrmPreview As frmDockEPRContent  '报告内容预览窗格
Private mstrCurId As String               '当前记录ID EMR库的ID是字符型
Private mstrContent As String             '新病历的XML内容
Private mintState As Integer            '当前记录状态
Private mblnCurMoved As Boolean         '当前记录转出状态 0-未转出 1-已转出
'-----------------------------------------------------

Private Function zlRefList(Optional strCurId As String, Optional strSender As String, Optional strPatient As String, Optional lngOutNo As Long, Optional lngInNo As Long) As Long
'功能：刷新装入符合条件的疾病报告，并定位到指定的记录上
'参数：strCurId 定位ID 接收、报送，撤报，刷新时传入
'       如果指定病人姓名、门诊号、住院号则不使用时间索引,报送人等四个参数仅会出0到1个
'       strSender  报送人 查找传入
'       strPatient 病人姓名 查找传入
'       lngOutNo   门诊号   查找传入
'       lngInNo    住院号   查找传入
Dim blnMoved As Boolean, strTemp As String, strFiles As String, i As Integer, strReturn As String
Dim rsTemp As New ADODB.Recordset, rsData As New ADODB.Recordset
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow

    Err = 0: On Error GoTo errHand
    
    If Trim(mstrFiles) = "" Then Exit Function
    Me.rptList.Records.DeleteAll
    
    For i = 0 To UBound(Split(mstrFiles, ","))
        If IsNumeric(Split(mstrFiles, ",")(i)) Then
            strFiles = strFiles & "," & Split(mstrFiles, ",")(i)
        End If
    Next
    If strFiles <> "" Then
        strFiles = Mid(strFiles, 2)
    End If
    
    If strFiles <> "" Then
        If mintDates <> 0 Then '如果指定病人姓名、门诊号、住院号则不使用时间索引
            gstrSQL = "And l.完成时间 >= trunc(Sysdate - [1])"
        Else
            blnMoved = MovedByDate(CDate(mstrDateFrom))
            gstrSQL = "And l.完成时间 Between To_Date([2],'yyyy-mm-dd') And To_Date([3],'yyyy-mm-dd')+1-1/24/60/60"
        End If
        
        If strPatient <> "" Or lngOutNo <> 0 Or lngInNo <> 0 Then
            gstrSQL = "And l.完成时间 is not null"
        End If
        
        gstrSQL = "Select l.Id,l.文件id,l.病人ID,l.主页ID, l.病历名称 As 报告, Decode(l.病人来源, 1, '门诊: ', 2, '住院: ', '') || d.名称 As 科室," & _
                "        Decode(l.病人来源, 1, p.门诊号, 2, p.住院号) As 就诊号, Nvl(l.姓名,p.姓名) as 姓名, Nvl(l.性别,p.性别) as 性别, Nvl(l.年龄,p.年龄) as 年龄, " & _
                "        To_Char(l.完成时间, 'yyyy-mm-dd hh24:mi') As 填报时间, l.保存人 As 填报人,l.编辑方式," & _
                "        Decode(l.状态, -1, Decode(Sign(l.保存时间 - l.收拒时间), 1, 0, -1), l.状态) As 状态," & _
                "        l.收拒人 || '|' || To_Char(l.收拒时间, 'yyyy-mm-dd hh24:mi') || '|' || l.收拒说明 || '|' || l.报送人 || '|' ||" & _
                "        To_Char(l.报送时间, 'yyyy-mm-dd hh24:mi') || '|' || l.报送单位 || '|' || l.报送备注 || '|' || l.登记人 || '|' ||" & _
                "        To_Char(l.登记时间, 'yyyy-mm-dd hh24:mi') || '|' || l.职业 || '|' || l.家庭地址 || '|' || l.家庭电话 || '|' ||" & _
                "        To_Char(l.发病日期, 'yyyy-mm-dd') || '|' || To_Char(l.确诊日期,'yyyy-mm-dd') || '|' ||" & _
                "        l.诊断描述1 || '|' || l.诊断描述2 || '|' || l.填报备注 As 信息,0 as 数据转出" & _
                " From (Select l.Id,l.文件id,l.病人ID,l.主页ID, l.病历名称, l.病人来源, l.科室id, l.完成时间, l.保存人, l.保存时间,l.编辑方式," & _
                "               Nvl(s.处理状态, 0) As 状态, s.收拒人, s.收拒时间, s.收拒说明, s.报送人, s.报送时间, s.报送单位, s.报送备注," & _
                "               s.登记人 , s.登记时间, s.姓名, s.性别, s.年龄, s.职业, s.家庭地址, s.家庭电话, s.发病日期, s.确诊日期, " & _
                "               s.诊断描述1, s.诊断描述2, s.填报备注" & _
                "        From 电子病历记录 l, 疾病申报记录 s" & _
                "        Where l.Id = s.文件id(+) And l.病历种类 = 5 And l.文件id In (" & strFiles & ") " & gstrSQL & _
                IIf(strSender = "", "", "And s.报送人=[4]") & _
                "       ) l,病人信息 p, 部门表 d" & IIf(lngInNo = 0, "", ",(Select Distinct 病人id,住院号 From 病案主页 Where 住院号 = [7]) A ") & _
                " Where l.病人id = p.病人id And l.科室id = d.Id" & IIf(strPatient = "", "", " And p.姓名=[5]") & _
                IIf(lngOutNo = 0, "", " And p.门诊号=[6]") & IIf(lngInNo = 0, "", " And a.住院号=[7] And a.病人ID=p.病人ID")
        If blnMoved Then
            strTemp = Replace(gstrSQL, "0 as 数据转出", "1 as 数据转出")
            strTemp = Replace(strTemp, "电子病历记录", "H电子病历记录")
            strTemp = Replace(strTemp, "疾病申报记录", "H疾病申报记录")
            gstrSQL = gstrSQL & " Union All " & strTemp
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mintDates, mstrDateFrom, mstrDateTo, strSender, strPatient, lngOutNo, lngInNo)
    
        Do While Not rsTemp.EOF
            Set rptRcd = Me.rptList.Records.Add()
            Set rptItem = rptRcd.AddItem(CStr(rsTemp!状态))
            Select Case rptItem.Value
            Case 0: rptItem.Icon = 0
            Case 1: rptItem.Icon = 1
            Case -1: rptItem.Icon = 2
            Case 2: rptItem.Icon = 3
            End Select
            rptRcd.AddItem CStr(rsTemp!ID)
            Select Case rsTemp!状态
            Case 0: rptRcd.AddItem CStr("a)新填写的疾病报告")
            Case 1: rptRcd.AddItem CStr("b)已接收的疾病报告")
            Case -1: rptRcd.AddItem CStr("c)已拒收的疾病报告")
            Case 2: rptRcd.AddItem CStr("d)已报送的疾病报告")
            Case Else: rptRcd.AddItem ""
            End Select
            rptRcd.AddItem CStr(rsTemp!报告)
            rptRcd.AddItem CStr(rsTemp!科室)
            rptRcd.AddItem CStr("" & rsTemp!就诊号)
            rptRcd.AddItem CStr(rsTemp!姓名)
            rptRcd.AddItem CStr("" & rsTemp!性别)
            rptRcd.AddItem CStr("" & rsTemp!年龄)
            rptRcd.AddItem CStr(NVL(rsTemp!填报时间))
            rptRcd.AddItem CStr(NVL(rsTemp!填报人))
            rptRcd.AddItem CStr(rsTemp!信息)
            rptRcd.AddItem CStr(rsTemp!数据转出)
            rptRcd.AddItem CStr(rsTemp!病人ID)
            rptRcd.AddItem CStr(rsTemp!主页ID)
            rptRcd.AddItem CStr(rsTemp!文件ID)
            rptRcd.AddItem CStr(NVL(rsTemp!编辑方式, 0))
            rsTemp.MoveNext
        Loop
    End If
    
    If Not gobjEmr Is Nothing Then
        strFiles = ""
        For i = 0 To UBound(Split(mstrFiles, ","))
            If Not IsNumeric(Split(mstrFiles, ",")(i)) Then
                strFiles = strFiles & ",Hextoraw('" & Split(mstrFiles, ",")(i) & "')"
            End If
        Next
        If strFiles <> "" Then
            strFiles = Mid(strFiles, 2)
        End If
        
        If strFiles <> "" Then
            If mintDates <> 0 Then
                gstrSQL = "l.complete_time >= trunc(Sysdate - :dates)"
            Else
                gstrSQL = "l.complete_time Between To_Date(:datef,'yyyy-mm-dd') And To_Date(:datet,'yyyy-mm-dd')+1-1/24/60/60"
            End If
            
            If strPatient <> "" Or lngOutNo <> 0 Or lngInNo <> 0 Then
                gstrSQL = "l.complete_time is not null"
            End If
            
            gstrSQL = "Select Rawtohex(m.Id) ID,Rawtohex(l.Antetype_Id) AntetypeId, m.Title 名称, Decode(o.Title, '门诊接诊', 1, 2) 病人来源,l.Completor 完成人, m.editor 编辑人, p.Code 病人id," & vbNewLine & _
                        "To_Char(m.Edit_Time, 'yyyy-mm-dd hh24:mi:ss') 保存时间, To_Char(l.Complete_Time, 'yyyy-mm-dd hh24:mi:ss') 完成时间, To_Char(n.Begin_Time, 'yyyy-mm-dd hh24:mi:ss') 事件时间" & vbNewLine & _
                    "From Bz_Doc_Tasks L, Bz_Doc_Log M, Bz_Act_Log N, Action_List O, Bz_Master_Codes P" & vbNewLine & _
                    "Where " & gstrSQL & " And l.Antetype_Id In (" & strFiles & ") And l.Real_Doc_Id = m.Id And m.Status >= 2 And" & vbNewLine & _
                    "      m.Actlog_Id = n.Id And n.Action_Id = o.Id And n.Master_Id = p.Master_Id And p.Kind = '病人ID'"
            If mintDates <> 0 Then
                strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, IIf(strPatient <> "" Or lngOutNo <> 0 Or lngInNo <> 0, "", mintDates & "^11^dates"), rsTemp)
            Else
                strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, IIf(strPatient <> "" Or lngOutNo <> 0 Or lngInNo <> 0, "", mstrDateFrom & "^16^datef|" & mstrDateTo & "^16^datet"), rsTemp)
            End If
            
            If strReturn = "" Then
            Do Until rsTemp.EOF
                gstrSQL = "Select '" & rsTemp!ID & "' As ID, p.病人id, '" & rsTemp!AntetypeId & "' As 文件id, q.主页id," & vbNewLine & _
                            "       '" & rsTemp!名称 & "' As 报告, Decode(" & rsTemp!病人来源 & ", 1, '门诊:', 2, '住院:') || a.名称 As 科室, Decode(" & rsTemp!病人来源 & ", 1, p.门诊号, 2, p.住院号) 就诊号," & vbNewLine & _
                            "       '" & Format(rsTemp!完成时间, "yyyy-mm-dd HH:MM") & "' As 填报时间, p.姓名, p.性别, p.年龄,0 As 数据转出, '" & rsTemp!编辑人 & "' As 填报人,3 as 编辑方式," & vbNewLine & _
                            "Nvl((Select Decode(Nvl(s.处理状态, 0), -1," & vbNewLine & _
                            "                 Decode(Sign(To_Date('2014-12-22 10:47:09', 'yyyy-mm-dd hh24:mi:ss') - s.收拒时间), 1, 0, -1)," & vbNewLine & _
                            "                 Nvl(s.处理状态, 0)) || '|' || s.收拒人 || '|' || To_Char(s.收拒时间, 'yyyy-mm-dd hh24:mi') || '|' || s.收拒说明 || '|' ||" & vbNewLine & _
                            "          s.报送人 || '|' || To_Char(s.报送时间, 'yyyy-mm-dd hh24:mi') || '|' || s.报送单位 || '|' || s.报送备注 || '|' || s.登记人 || '|' ||" & vbNewLine & _
                            "          To_Char(s.登记时间, 'yyyy-mm-dd hh24:mi') || '|' || s.职业 || '|' || s.家庭地址 || '|' || s.家庭电话 || '|' ||" & vbNewLine & _
                            "          To_Char(s.发病日期, 'yyyy-mm-dd') || '|' || To_Char(s.确诊日期, 'yyyy-mm-dd') || '|' || s.诊断描述1 || '|' || s.诊断描述2 || '|' ||" & vbNewLine & _
                            "          s.填报备注 As 信息" & vbNewLine & _
                            "  From 疾病申报记录 S" & vbNewLine & _
                            "  Where s.文档id = [1]" & IIf(strSender <> "", " And s.报送人=[4]", "") & "),'||||||||||||||||') 信息" & vbNewLine & _
                            "From 病人信息 P, 病案主页 Q, 部门表 A" & vbNewLine & _
                            "Where p.病人id = [2] And p.病人id = q.病人id And [3] Between q.入院日期 And" & vbNewLine & _
                            "      Nvl(q.出院日期, Sysdate) And q.出院科室ID = a.Id" & vbNewLine & _
                            IIf(strPatient <> "", " And P.姓名=[5]", "") & IIf(lngOutNo = 0, "", " And p.门诊号=[6]") & IIf(lngInNo = 0, "", " And Q.住院号=[7]")
                gstrSQL = gstrSQL & vbNewLine & " Union " & vbNewLine & _
                            "Select '" & rsTemp!ID & "' As ID, p.病人id, '" & rsTemp!AntetypeId & "' As 文件id, q.id 主页ID," & vbNewLine & _
                                        "       '" & rsTemp!名称 & "' As 报告, Decode(" & rsTemp!病人来源 & ", 1, '门诊:', 2, '住院:') || a.名称 As 科室, Decode(" & rsTemp!病人来源 & ", 1, p.门诊号, 2, p.住院号) 就诊号," & vbNewLine & _
                                        "       '" & Format(rsTemp!完成时间, "yyyy-mm-dd HH:MM") & "' As 填报时间, p.姓名, p.性别, p.年龄,0 As 数据转出, '" & rsTemp!编辑人 & "' As 填报人,3 as 编辑方式," & vbNewLine & _
                                        "Nvl((Select Decode(Nvl(s.处理状态, 0), -1," & vbNewLine & _
                                        "                 Decode(Sign(To_Date('2014-12-22 10:47:09', 'yyyy-mm-dd hh24:mi:ss') - s.收拒时间), 1, 0, -1)," & vbNewLine & _
                                        "                 Nvl(s.处理状态, 0)) || '|' || s.收拒人 || '|' || To_Char(s.收拒时间, 'yyyy-mm-dd hh24:mi') || '|' || s.收拒说明 || '|' ||" & vbNewLine & _
                                        "          s.报送人 || '|' || To_Char(s.报送时间, 'yyyy-mm-dd hh24:mi') || '|' || s.报送单位 || '|' || s.报送备注 || '|' || s.登记人 || '|' ||" & vbNewLine & _
                                        "          To_Char(s.登记时间, 'yyyy-mm-dd hh24:mi') || '|' || s.职业 || '|' || s.家庭地址 || '|' || s.家庭电话 || '|' ||" & vbNewLine & _
                                        "          To_Char(s.发病日期, 'yyyy-mm-dd') || '|' || To_Char(s.确诊日期, 'yyyy-mm-dd') || '|' || s.诊断描述1 || '|' || s.诊断描述2 || '|' ||" & vbNewLine & _
                                        "          s.填报备注 As 信息" & vbNewLine & _
                                        "  From 疾病申报记录 S" & vbNewLine & _
                                        "  Where s.文档id = [1]" & IIf(strSender <> "", " And s.报送人=[4]", "") & "),'||||||||||||||||') 信息" & vbNewLine & _
                                        "From 病人信息 P, 病人挂号记录 Q, 部门表 A" & vbNewLine & _
                                        "Where p.病人id = [2] And p.病人id = q.病人id And q.执行时间=[3]" & vbNewLine & _
                                        "      And q.执行部门ID = a.Id" & vbNewLine & _
                                        IIf(strPatient <> "", " And P.姓名=[5]", "") & IIf(lngOutNo = 0, "", " And p.门诊号=[6]")
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(rsTemp!ID), CLng(rsTemp!病人ID), CDate(rsTemp!事件时间), strSender, strPatient, lngOutNo, lngInNo)
                Do While Not rsData.EOF
                    Set rptRcd = Me.rptList.Records.Add()
                    Set rptItem = rptRcd.AddItem(Val(Split(rsData!信息, "|")(0)))
                    Select Case rptItem.Value
                    Case 0: rptItem.Icon = 0
                    Case 1: rptItem.Icon = 1
                    Case -1: rptItem.Icon = 2
                    Case 2: rptItem.Icon = 3
                    End Select
                    rptRcd.AddItem CStr(rsData!ID)
                    Select Case Val(Split(rsData!信息, "|")(0))
                    Case 0: rptRcd.AddItem CStr("a)新填写的疾病报告")
                    Case 1: rptRcd.AddItem CStr("b)已接收的疾病报告")
                    Case -1: rptRcd.AddItem CStr("c)已拒收的疾病报告")
                    Case 2: rptRcd.AddItem CStr("d)已报送的疾病报告")
                    Case Else: rptRcd.AddItem ""
                    End Select
                    rptRcd.AddItem CStr(rsData!报告)
                    rptRcd.AddItem CStr(rsData!科室)
                    rptRcd.AddItem CStr("" & rsData!就诊号)
                    rptRcd.AddItem CStr(rsData!姓名)
                    rptRcd.AddItem CStr("" & rsData!性别)
                    rptRcd.AddItem CStr("" & rsData!年龄)
                    rptRcd.AddItem CStr(rsData!填报时间)
                    rptRcd.AddItem CStr(rsData!填报人)
                    rptRcd.AddItem CStr(Mid(rsData!信息, InStr(rsData!信息, "|") + 1))
                    rptRcd.AddItem CStr(rsData!数据转出)
                    rptRcd.AddItem CStr(rsData!病人ID)
                    rptRcd.AddItem CStr(NVL(rsData!主页ID, 0))
                    rptRcd.AddItem CStr(rsData!文件ID)
                    rptRcd.AddItem CStr(NVL(rsData!编辑方式, 0))
                    rsData.MoveNext
                Loop
                rsTemp.MoveNext
            Loop
            End If
        End If
    End If
    
    Me.rptList.Populate
    
    If strCurId <> "" Then
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow = False Then
                If CStr(rptRow.Record(mCol.ID).Value) = strCurId Then
                    Set Me.rptList.FocusedRow = rptRow: Exit For
                End If
            End If
        Next
    End If
    If Me.rptList.Rows.Count > 0 Then
        If Me.rptList.FocusedRow Is Nothing Then Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
        If Me.rptList.FocusedRow.GroupRow Then
            strCurId = ""
        Else
            strCurId = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
        End If
    Else
        strCurId = ""
    End If
    zlRefList = Me.rptList.Records.Count
    Exit Function

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    zlRefList = Me.rptList.Records.Count
End Function

Public Sub zlRptPrint(ByVal bytMode As Byte)
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    If Me.rptList.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '复制数据表格
    If zlReportToVSFlexGrid(Me.vfgTemp, Me.rptList) = False Then Exit Sub
    
    '-------------------------------------------------
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vfgTemp
    objPrint.Title.Text = "病历文件清单"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

'-------------------------------------------------------
'功能：  报告预览及打印
'参数：  blnPreview  :是否是预览模式
'-------------------------------------------------------
Private Sub zlEPRPrint(blnPreview As Boolean)
Dim frmPrint As frmPrintPreview, ObjTabEprView As Object
Dim rsTemp As New ADODB.Recordset
    If mstrCurId = "" Then Exit Sub
    Err = 0: On Error GoTo errHand
    If IsNumeric(mstrCurId) Then
        gstrSQL = "Select l.病人来源, l.病人id, l.主页id,l.编辑方式, f.页面 From 电子病历记录 l, 病历文件列表 f Where l.文件id = f.Id And l.Id = [1]"
        If mblnCurMoved Then
            gstrSQL = Replace(gstrSQL, "电子病历记录", "H电子病历记录")
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(mstrCurId))
        With rsTemp
            If .RecordCount <= 0 Then MsgBox "该疾病报告可能已经被临床删除！", vbExclamation, gstrSysName: Exit Sub
            If rsTemp!编辑方式 = 0 Then
                Set frmPrint = New frmPrintPreview
                Select Case !病人来源
                Case 1
                    frmPrint.DoMultiDocPreview Me, cpr门诊病历, !病人ID, !主页ID, 5, !页面, CLng(mstrCurId), Not blnPreview, , , mblnCurMoved
                Case 2
                    frmPrint.DoMultiDocPreview Me, cpr住院病历, !病人ID, !主页ID, 5, !页面, CLng(mstrCurId), Not blnPreview, , , mblnCurMoved
                End Select
                Unload frmPrint
                Set frmPrint = Nothing
            ElseIf rsTemp!编辑方式 = 1 Then
                Set ObjTabEprView = DynamicCreate("zlTableEPR.cTableEPR", "打印表格病历", True)
                Call ObjTabEprView.InitTableEPR(gcnOracle, glngSys, gstrDbOwner)
                Call ObjTabEprView.InitOpenEPR(Me, cprEM_修改, cprET_单病历审核, CLng(mstrCurId), False, 0, !病人来源)
                ObjTabEprView.zlPrintDoc Me, blnPreview, ""
            ElseIf rsTemp!编辑方式 = 2 Then
                mobjInfection.PrintDoc Me, !病人ID, !主页ID, CLng(mstrCurId), ""
            End If
        End With
    Else
        If Not mobjRichEMR Is Nothing Then Call mobjRichEMR.zlPrintDoc(blnPreview)
    End If
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'-------------------------------------------------------
'以下为控件事件过程
'-------------------------------------------------------

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim cbrControl As CommandBarControl, strInfo As String, bytEdit As Byte
    If mblnCurMoved And (Control.ID = conMenu_File_Open Or Control.ID = conMenu_Edit_Reuse Or Control.ID = conMenu_Edit_Send Or Control.ID = conMenu_Edit_Untread) Then
        MsgBox "该病人的本次数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                        "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Sub
    End If

    If rptList.FocusedRow Is Nothing Then
        bytEdit = 0
    ElseIf rptList.FocusedRow.GroupRow = True Then
        bytEdit = 0
    Else
        bytEdit = Val(rptList.FocusedRow.Record.Item(mCol.编辑方式).Value)
    End If
            
    Select Case Control.ID
    Case conMenu_File_Open
        If bytEdit = 0 Then
            Dim f As New frmEPRView
            f.ShowMe Me, CLng(mstrCurId)
        ElseIf bytEdit = 3 Then
            '新病历编辑器
            Call mobjRichEMR.zlViewDoc(Me, "查阅", "")
        End If
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlEPRPrint(True)
    Case conMenu_File_Print: Call zlEPRPrint(False)
    Case conMenu_File_RowPrint: Call zlRptPrint(1)
    Case conMenu_File_Parameter
        Call frmDiseaseReportSet.ShowMe(Me, InStr(1, mstrPrivs, "范围设置") > 0, mstrFiles, mintDates, mstrDateFrom, mstrDateTo)
        Call zlRefList
    Case conMenu_File_Exit: Unload Me
    Case conMenu_Edit_Audit '审核病历
        '单病历审核模式
        If bytEdit = 0 Then
            Dim frmAudit As Form, bFindedAudit As Boolean
            For Each frmAudit In Forms
                If frmAudit.Name = "frmMain" Then
                    If frmAudit.Document.EPRPatiRecInfo.ID = CLng(mstrCurId) Then
                        frmAudit.Show
                        bFindedAudit = True
                    End If
                End If
            Next
            If bFindedAudit = False Then
                Set mobjDoc = New cEPRDocument
                mobjDoc.InitEPRDoc cprEM_修改, cprET_单病历审核, CLng(mstrCurId), cprPF_住院
                mobjDoc.ShowEPREditor Me
            End If
        ElseIf bytEdit = 3 Then
            '新病历编辑器
            Dim objAudit As Object
            Set objAudit = DynamicCreate("zlRichEMR.clsDockEMR", "新版病历", False)
            Call objAudit.Init(gobjEmr, gcnOracle, glngSys)
            Call objAudit.zlRefresh(rptList.FocusedRow.Record.Item(mCol.病人ID).Value, rptList.FocusedRow.Record.Item(mCol.主页ID).Value, glngDeptId, 0, IIf(InStr(rptList.FocusedRow.Record.Item(mCol.科室).Value, "门诊") > 0, 1, 2))
            Call objAudit.EditDoc(mstrCurId)
        End If
    Case conMenu_Edit_Reuse
        'strInfo=报告|科室|姓名|性别|年龄|就诊号|填报人|填报时间|病人ID|主页ID
        With rptList
            strInfo = .FocusedRow.Record.Item(mCol.报告).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.科室).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.姓名).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.性别).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.年龄).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.就诊号).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.填报人).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.填报时间).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.病人ID).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.主页ID).Value
        End With
        If Not IsNumeric(mstrCurId) Then
            '新病历 strInfo=strInfo & |甲类传染病|乙类传染病|丙类传染病|病例分类|病例分类2
            strInfo = strInfo & "|" & rptList.Tag
        End If
        If frmDiseaseReportIncept.ShowMe(Me, mstrCurId, strInfo) Then Call zlRefList(mstrCurId)
    Case conMenu_Edit_Send
        'strInfo=报告|科室|姓名|性别|年龄|就诊号|填报人|填报时间|病人ID|主页ID
        With rptList
            strInfo = .FocusedRow.Record.Item(mCol.报告).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.科室).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.姓名).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.性别).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.年龄).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.就诊号).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.填报人).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.填报时间).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.病人ID).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.主页ID).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.文件ID).Value
        End With
        If frmDiseaseReportSend.ShowMe(Me, mstrCurId, strInfo, Me.vfgInfo.TextMatrix(5, 1), Me.vfgInfo.TextMatrix(6, 1)) Then Call zlRefList(mstrCurId)
    Case conMenu_Edit_Untread
        Dim strMsg As String
        Select Case mintState
        Case 1:  strMsg = "真的取消该疾病报告的“接收处理”吗？"
        Case -1: strMsg = "真的取消该疾病报告的“拒绝处理”吗？"
        Case 2:  strMsg = "真的取消该疾病报告的“申报登记”吗？"
        Case Else: Exit Sub
        End Select
        If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        gstrSQL = "Zl_疾病申报记录_Untread('" & mstrCurId & "')"
        Err = 0: On Error GoTo errHand
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Call zlRefList(mstrCurId)
    Case conMenu_Edit_Compend '对应信息设置
        frmDiseaseReportRela.Show 1, Me
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.STYLE = IIf(cbrControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_Refresh
        Call zlRefList(mstrCurId)
    Case conMenu_View_Find
        Dim strSender As String, strPatient As String, lngOutNo As Long, lngInNo As Long
        If frmDiseaseReportFind.ShowMe(Me, mintDates, mstrDateFrom, mstrDateTo, strSender, strPatient, lngOutNo, lngInNo) Then
            Call zlRefList(mstrCurId, strSender, strPatient, lngOutNo, lngInNo)
        End If
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Forum '中联论坛
        Call zlWebForum(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    Case Else
        '执行发布到当前模块的报表
        Dim lng报告ID As Long
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            If rptList.SelectedRows.Count > 0 Then
                If Not rptList.SelectedRows(0).GroupRow Then
                    lng报告ID = Val(rptList.SelectedRows(0).Record(mCol.ID).Value)
                End If
            End If
            If lng报告ID <> 0 Then
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "报告ID=" & lng报告ID)
            Else
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
            End If
        End If
    End Select
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Open
         Control.Enabled = (mstrCurId <> "")
         If Control.Enabled Then Control.Enabled = InStr("0,3", Me.rptList.FocusedRow.Record.Item(mCol.编辑方式).Value) > 0 '目前仅支持老病历、智能病历
    Case conMenu_File_Preview
         Control.Enabled = (mstrCurId <> "")
         If Control.Enabled Then Control.Enabled = InStr("0,1,3", Me.rptList.FocusedRow.Record.Item(mCol.编辑方式).Value) > 0 '目前仅支持老病历、智能病历、表格病历
    Case conMenu_File_Print
         Control.Enabled = (mstrCurId <> "")
         If Control.Enabled Then Control.Enabled = InStr("0,1,3", Me.rptList.FocusedRow.Record.Item(mCol.编辑方式).Value) > 0 '目前仅支持老病历、智能病历、表格病历
    Case conMenu_File_RowPrint: Control.Enabled = (Me.rptList.Records.Count <> "")
    Case conMenu_Edit_Audit
        Control.Visible = (mstrCurId <> "")
        If Control.Visible Then Control.Visible = InStr("0,3", Me.rptList.FocusedRow.Record.Item(mCol.编辑方式).Value) > 0 '目前仅支持老病历、智能病历
        Control.Enabled = (InStr(1, mstrPrivs, "病历审阅") > 0)
        If Control.Enabled Then Control.Enabled = (mstrCurId <> "")
        If Control.Enabled Then Control.Enabled = (mintState = 0)
    Case conMenu_Edit_Reuse
        Control.Enabled = (InStr(1, mstrPrivs, "接收") > 0)
        If Control.Enabled Then Control.Enabled = (mstrCurId <> "" And mintState = 0)
    Case conMenu_Edit_Send
        Control.Enabled = (InStr(1, mstrPrivs, "报送") > 0)
        If Control.Enabled Then Control.Enabled = (mstrCurId <> "" And mintState = 1)
    Case conMenu_Edit_Untread
        Control.Enabled = (InStr(1, mstrPrivs, "回退") > 0)
        If Control.Enabled Then Control.Enabled = (mstrCurId <> "" And mintState <> 0)
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).STYLE = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    Case conMenu_View_Refresh:  Control.Enabled = (Trim(mstrFiles) <> "")
    Case conMenu_View_Find::  Control.Enabled = (Trim(mstrFiles) <> "")
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_Reports
        Item.Handle = Me.rptList.hWnd
    Case conPane_Preview
        If mfrmPreview Is Nothing Then Set mfrmPreview = New frmDockEPRContent
        Item.Handle = mfrmPreview.hWnd
    Case conPane_AppInfo
        Item.Handle = Me.vfgInfo.hWnd
    End Select
End Sub

Private Sub Form_Load()
Dim cbrMenuBar As CommandBarPopup
Dim cbrControl As CommandBarControl
Dim cbrToolBar As CommandBar
Dim rptCol As ReportColumn
Dim lngCount As Long
    '-----------------------------------------------------
    '权限限制串复制，避免同时进入其他模块而导致gstrPrivs变化，导致控制无效
    mstrPrivs = gstrPrivs
    mstrFiles = Trim(GetSetting("ZLSOFT", App.EXEName, "疾病申报文件范围", ""))
    mintDates = Val(GetSetting("ZLSOFT", App.EXEName, "疾病申报最近天数", 0))
    If mintDates = 0 Then mintDates = 7: Call SaveSetting("ZLSOFT", App.EXEName, "疾病申报最近天数", mintDates)
    
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '菜单定义
    Me.cbsThis.ActiveMenuBar.Title = "菜单"
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "打开(&O)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_RowPrint, "清单打印(&L)…"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&M)…"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "修订(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "接收(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Send, "报送(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "收回(&B)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Compend, "对应信息设置(&B)")
        cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Find, "查找(&F)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False  '固有
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With
    
    '快键绑定
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("O"), conMenu_File_Open
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("F"), conMenu_View_Find
        .Add 0, vbKeyF12, conMenu_File_Parameter
        .Add FCONTROL, Asc("A"), conMenu_Edit_Reuse
        .Add FCONTROL, Asc("T"), conMenu_Edit_Send
        .Add FCONTROL, Asc("U"), conMenu_Edit_Untread
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F6, conMenu_View_Jump
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '设置不常用菜单
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_RowPrint
        .AddHiddenCommand conMenu_View_Refresh
    End With
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "打开")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "修订")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "接收")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Send, "报送")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "回退")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.STYLE = xtpButtonIconAndCaption
    Next
    
    '读取发布到该模块的报表:因为是一次性读取,全局变量可用
    '---------------------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    
    '-----------------------------------------------------
    '设置词句显示停靠窗格
    If mfrmPreview Is Nothing Then Set mfrmPreview = New frmDockEPRContent
    If Not gobjEmr Is Nothing Then
        Set mobjRichEMR = DynamicCreate("zlRichEMR.clsDockContent", "新版病历", False)
        If Not mobjRichEMR Is Nothing Then Call mobjRichEMR.Init(gobjEmr, gcnOracle, glngSys, 0)
    End If

    Set mobjInfection = DynamicCreate("zlDisReportCard.clsDisReportCard", "传染病报告卡", True)
    If Not mobjInfection Is Nothing Then
        mobjInfection.Init gcnOracle, glngSys
    End If
    
    Dim panThis As Pane, panChild As Pane
    Set panThis = dkpMan.CreatePane(conPane_Reports, 400, 200, DockLeftOf, Nothing)
    panThis.Title = "报告列表": panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Set panChild = dkpMan.CreatePane(conPane_Preview, 400, 200, DockBottomOf, Nothing)
    panChild.Title = "报告内容": panChild.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Set panThis = dkpMan.CreatePane(conPane_AppInfo, 200, 400, DockRightOf, Nothing)
    panThis.Title = "附加信息": panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    
    '-----------------------------------------------------
    With Me.rptList
        Set rptCol = .Columns.Add(mCol.图标, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.状态, "状态", 90, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.报告, "报告", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.科室, "科室", 110, True): rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(mCol.就诊号, "门诊&住院号", 75, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.姓名, "姓名", 50, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.性别, "性别", 40, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.年龄, "年龄", 40, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.填报时间, "填报时间", 100, True): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.填报人, "填报人", 50, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.信息, "信息", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.数据转出, "数据转出", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.病人ID, "病人ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.主页ID, "主页ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.文件ID, "文件ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        .GroupsOrder.Add .Columns.Find(mCol.状态)
        .GroupsOrder(0).SortAscending = True
        
        .SetImageList Me.imgList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With
    
    '-----------------------------------------------------
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)
    '-----------------------------------------------------
    '数据装入
    If mstrFiles = "" Then
        Me.stbThis.Panels(2).Text = "未设置本工作站的疾病报告范围"
    Else
        lngCount = zlRefList()
        Me.stbThis.Panels(2).Text = "共有" & lngCount & "份疾病报告"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Unload mfrmPreview
    Set mfrmPreview = Nothing
    Set mobjDoc = Nothing
    Unload mobjRichEMR.zlGetForm
    Set mobjRichEMR.zlGetForm = Nothing
    Set mobjRichEMR = Nothing
    Set mobjInfection = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
Dim cbrPopupBar As CommandBar
Dim cbrPopupItem As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrControl As CommandBarControl
    
    If Button <> vbRightButton Then Exit Sub
    If Me.cbsThis.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = Me.cbsThis.Add("弹出菜单", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
    Next
    cbrPopupBar.ShowPopup
End Sub

Private Sub rptList_SelectionChanged()
    Dim strInfo As String, aryInfo() As String
    
    mstrContent = "": rptList.Tag = ""
    With Me.rptList
        If .FocusedRow Is Nothing Then
            mstrCurId = "": mintState = 0: strInfo = "": mblnCurMoved = False
        ElseIf .FocusedRow.GroupRow = True Then
            mstrCurId = "": mintState = 0: strInfo = "": mblnCurMoved = False
        Else
            mstrCurId = .FocusedRow.Record.Item(mCol.ID).Value
            mintState = .FocusedRow.Record.Item(mCol.图标).Value
            strInfo = .FocusedRow.Record.Item(mCol.信息).Value
            mblnCurMoved = (.FocusedRow.Record.Item(mCol.数据转出).Value = 1)
        End If
    End With
    
    aryInfo = Split(strInfo, "|")
    With Me.vfgInfo
        .Clear
        .ColWidth(0) = 900
        Select Case mintState
        Case 0
            If strInfo <> "" Then .Rows = 1: .TextMatrix(0, 0) = "………": .TextMatrix(0, 1) = "等待接收…"
        Case 1
            .Rows = 11
            .TextMatrix(0, 0) = "职业": .TextMatrix(0, 1) = aryInfo(9)
            .TextMatrix(1, 0) = "家庭地址": .TextMatrix(1, 1) = aryInfo(10)
            .TextMatrix(2, 0) = "家庭电话": .TextMatrix(2, 1) = aryInfo(11)
            .TextMatrix(3, 0) = "发病日期": .TextMatrix(3, 1) = aryInfo(12)
            .TextMatrix(4, 0) = "确诊日期": .TextMatrix(4, 1) = aryInfo(13)
            .TextMatrix(5, 0) = "诊断描述1": .TextMatrix(5, 1) = aryInfo(14)
            .TextMatrix(6, 0) = "诊断描述2": .TextMatrix(6, 1) = aryInfo(15)
            .TextMatrix(7, 0) = "填报备注": .TextMatrix(7, 1) = aryInfo(16)
            
            .TextMatrix(8, 0) = "接收人": .TextMatrix(8, 1) = aryInfo(0)
            .TextMatrix(9, 0) = "接收时间": .TextMatrix(9, 1) = aryInfo(1)
            .TextMatrix(10, 0) = "接收说明": .TextMatrix(10, 1) = aryInfo(2)
        Case -1
            .Rows = 3
            .TextMatrix(0, 0) = "拒收人": .TextMatrix(0, 1) = aryInfo(0)
            .TextMatrix(1, 0) = "拒收时间": .TextMatrix(1, 1) = aryInfo(1)
            .TextMatrix(2, 0) = "拒收原因": .TextMatrix(2, 1) = aryInfo(2)
        Case 2
            .Rows = 17
            .TextMatrix(0, 0) = "职业": .TextMatrix(0, 1) = aryInfo(9)
            .TextMatrix(1, 0) = "家庭地址": .TextMatrix(1, 1) = aryInfo(10)
            .TextMatrix(2, 0) = "家庭电话": .TextMatrix(2, 1) = aryInfo(11)
            .TextMatrix(3, 0) = "发病日期": .TextMatrix(3, 1) = aryInfo(12)
            .TextMatrix(4, 0) = "确诊日期": .TextMatrix(4, 1) = aryInfo(13)
            .TextMatrix(5, 0) = "诊断描述1": .TextMatrix(5, 1) = aryInfo(14)
            .TextMatrix(6, 0) = "诊断描述2": .TextMatrix(6, 1) = aryInfo(15)
            .TextMatrix(7, 0) = "填报备注": .TextMatrix(7, 1) = aryInfo(16)
            
            .TextMatrix(8, 0) = "接收人": .TextMatrix(8, 1) = aryInfo(0)
            .TextMatrix(9, 0) = "接收时间": .TextMatrix(9, 1) = aryInfo(1)
            .TextMatrix(10, 0) = "接收说明": .TextMatrix(10, 1) = aryInfo(2)
            .TextMatrix(11, 0) = "报送人": .TextMatrix(11, 1) = aryInfo(3)
            .TextMatrix(12, 0) = "报送时间": .TextMatrix(12, 1) = aryInfo(4)
            .TextMatrix(13, 0) = "报送单位": .TextMatrix(13, 1) = aryInfo(5)
            .TextMatrix(14, 0) = "报送备注": .TextMatrix(14, 1) = aryInfo(6)
            .TextMatrix(15, 0) = "登记人": .TextMatrix(15, 1) = aryInfo(7)
            .TextMatrix(16, 0) = "登记时间": .TextMatrix(16, 1) = aryInfo(8)
        End Select
    End With
    
    On Error Resume Next
    If IsNumeric(mstrCurId) Then
        dkpMan.FindPane(conPane_Preview).Handle = mfrmPreview.hWnd
        Call mfrmPreview.zlRefresh(CLng(mstrCurId), "", , mblnCurMoved, , NVL(rptList.FocusedRow.Record.Item(mCol.编辑方式).Value, 0))
    ElseIf mstrCurId <> "" Then
        dkpMan.FindPane(conPane_Preview).Handle = mobjRichEMR.zlGetForm.hWnd
        Call mobjRichEMR.zlShowDoc(mstrCurId, "")
        Call mobjRichEMR.zlGetForm.DocContent.SaveToXML(mstrContent, False)
        
        If mstrContent <> "" Then
            Dim xmldom As Object, xmlNode As Object, xmlEle As Object, sxpath As String, lngItem As Long, strCDATA As String, strItem As String
            Set xmldom = CreateObject("Msxml2.DOMDocument.6.0")
            Set xmlEle = CreateObject("Msxml2.DOMDocument.6.0")
            Call xmldom.loadXML(mstrContent)
            
            lngItem = 0: strItem = ""
            sxpath = "/zlxml/document/e_enum[contains(@title,""甲类传染病"")]"
            Call xmlEle.loadXML(xmldom.selectSingleNode(sxpath).xml)
            Set xmlNode = xmlEle.selectSingleNode("/e_enum/enumvalues/element")
            If Not xmlNode Is Nothing Then
                lngItem = Val(xmlNode.Text)
                If xmldom.selectSingleNode(sxpath).firstChild.nodeType = NODE_CDATA_SECTION Then
                    strCDATA = xmldom.selectSingleNode(sxpath).firstChild.nodeValue
                    strCDATA = Replace(strCDATA, "rangexml='", "")
                    strCDATA = Mid(strCDATA, 1, Len(strCDATA) - 1)
                    Call xmlEle.loadXML(strCDATA)
                    strItem = xmlEle.selectSingleNode("/root/item[" & lngItem & "]/meaning").Text
                End If
            End If
            strInfo = strItem
            
            lngItem = 0: strItem = ""
            sxpath = "/zlxml/document/e_enum[contains(@title,""乙类传染病"")]"
            Call xmlEle.loadXML(xmldom.selectSingleNode(sxpath).xml)
            Set xmlNode = xmlEle.selectSingleNode("/e_enum/enumvalues/element")
            If Not xmlNode Is Nothing Then
                lngItem = Val(xmlNode.Text)
                If xmldom.selectSingleNode(sxpath).firstChild.nodeType = NODE_CDATA_SECTION Then
                    strCDATA = xmldom.selectSingleNode(sxpath).firstChild.nodeValue
                    strCDATA = Replace(strCDATA, "rangexml='", "")
                    strCDATA = Mid(strCDATA, 1, Len(strCDATA) - 1)
                    Call xmlEle.loadXML(strCDATA)
                    strItem = xmlEle.selectSingleNode("/root/item[" & lngItem & "]/meaning").Text
                End If
            End If
            strInfo = strInfo & "|" & strItem
            
            lngItem = 0: strItem = ""
            sxpath = "/zlxml/document/e_enum[contains(@title,""丙类传染病"")]"
            Call xmlEle.loadXML(xmldom.selectSingleNode(sxpath).xml)
            Set xmlNode = xmlEle.selectSingleNode("/e_enum/enumvalues/element")
            If Not xmlNode Is Nothing Then
                lngItem = Val(xmlNode.Text)
                If xmldom.selectSingleNode(sxpath).firstChild.nodeType = NODE_CDATA_SECTION Then
                    strCDATA = xmldom.selectSingleNode(sxpath).firstChild.nodeValue
                    strCDATA = Replace(strCDATA, "rangexml='", "")
                    strCDATA = Mid(strCDATA, 1, Len(strCDATA) - 1)
                    Call xmlEle.loadXML(strCDATA)
                    strItem = xmlEle.selectSingleNode("/root/item[" & lngItem & "]/meaning").Text
                End If
            End If
            strInfo = strInfo & "|" & strItem
            
            lngItem = 0: strItem = ""
            sxpath = "/zlxml/document/e_enum[contains(@title,""病例分类"")][1]"
            Call xmlEle.loadXML(xmldom.selectSingleNode(sxpath).xml)
            Set xmlNode = xmlEle.selectSingleNode("/e_enum/enumvalues/element")
            If Not xmlNode Is Nothing Then
                lngItem = Val(xmlNode.Text)
                If xmldom.selectSingleNode(sxpath).firstChild.nodeType = NODE_CDATA_SECTION Then
                    strCDATA = xmldom.selectSingleNode(sxpath).firstChild.nodeValue
                    strCDATA = Replace(strCDATA, "rangexml='", "")
                    strCDATA = Mid(strCDATA, 1, Len(strCDATA) - 1)
                    Call xmlEle.loadXML(strCDATA)
                    strItem = xmlEle.selectSingleNode("/root/item[" & lngItem & "]/meaning").Text
                End If
            End If
            strInfo = strInfo & "|" & strItem
            
            lngItem = 0: strItem = ""
            sxpath = "/zlxml/document/e_enum[contains(@title,""病例分类"")][2]"
            Call xmlEle.loadXML(xmldom.selectSingleNode(sxpath).xml)
            Set xmlNode = xmlEle.selectSingleNode("/e_enum/enumvalues/element")
            If Not xmlNode Is Nothing Then
                lngItem = Val(xmlNode.Text)
                If xmldom.selectSingleNode(sxpath).firstChild.nodeType = NODE_CDATA_SECTION Then
                    strCDATA = xmldom.selectSingleNode(sxpath).firstChild.nodeValue
                    strCDATA = Replace(strCDATA, "rangexml='", "")
                    strCDATA = Mid(strCDATA, 1, Len(strCDATA) - 1)
                    Call xmlEle.loadXML(strCDATA)
                    strItem = xmlEle.selectSingleNode("/root/item[" & lngItem & "]/meaning").Text
                End If
            End If
            strInfo = strInfo & "|" & strItem
            
        End If
        rptList.Tag = strInfo
    End If
End Sub




