VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Begin VB.Form frmDockReport 
   BorderStyle     =   0  'None
   Caption         =   "诊疗报告管理"
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picRichEdit 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   2805
      Left            =   150
      ScaleHeight     =   2805
      ScaleWidth      =   5055
      TabIndex        =   2
      Top             =   435
      Width           =   5055
      Begin zlRichEditor.Editor edtThis 
         Height          =   1890
         Left            =   30
         TabIndex        =   3
         Top             =   15
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   3334
         WithViewButtonas=   0   'False
         ShowRuler       =   0   'False
      End
   End
   Begin VB.PictureBox picNote 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   135
      ScaleHeight     =   375
      ScaleWidth      =   6330
      TabIndex        =   0
      Top             =   45
      Width           =   6330
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "！"
         Height          =   180
         Left            =   135
         TabIndex        =   1
         Top             =   90
         Width           =   180
      End
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   15
      Top             =   765
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   135
      Top             =   4710
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmDockReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------
'常量
'-----------------------------------------------------
Const conPane_Note = 1
Const conPane_Content = 2
Const conPane_Table = 3
Const conPane_Annex = 4

'-----------------------------------------------------
'窗体事件
'-----------------------------------------------------
Public Event Activate()
Public Event AfterSaved(ByVal lngOrderId As Long, ByVal lngSaveType As Long)
Public Event AfterOpen(ByVal intEditType As EditTypeEnum)
Public Event AfterClosed(ByVal lngOrderId As Long)
Public Event AfterPrinted(ByVal lngOrderId As Long)
Public Event AfterDeleted(ByVal lngOrderId As Long)
'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mstrPrivs As String     '当前使用者对本程序(1258)的权限串
Private mblnSearch As Boolean   '当前使用者是否具备病历检索(1273)权

Private mlngOrderId     As Long         '医嘱id
Private mblnMoved       As Boolean      '是否转储
Private mblnCanPrint    As Boolean      '可否打印
Private mintPati来源    As Long
Private mlngPati病人ID  As Long
Private mlngPati主页ID  As Long
Private mlngPati婴儿    As Long
Private mlngEPR定义ID   As Long
Private mlngEPR报告ID   As Long
Private mstrEPR报告名称 As String
Private mstrEPR创建人   As String
Private mstrEPR保存人   As String
Private mstrEPR归档人   As String
Private mstrEPR完成时间 As String
Private mintEPR签名级别 As Integer
Private mintEPR签名版本 As Integer
Private mintEPR最后版本 As Integer
Private mlngEPR科室ID   As Long
Private mbyeEPR编辑方式 As Byte
Private mlngSingCount   As Long

Private mlngDeptId As Long          '当前操作科室id
Private mblnEdit As Boolean         '是否允许操作
Private mlngModule As Long

Private WithEvents mobjDoc As cEPRDocument
Attribute mobjDoc.VB_VarHelpID = -1
Private mObjTabEpr As cTableEPR
Private mObjTabEprView As cTableEPR
Private mfrmAnnex As frmDockAnnex    '病历附件窗体
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private WithEvents mfrmPrintPreview As frmPrintPreview
Attribute mfrmPrintPreview.VB_VarHelpID = -1

Private mstrPrinterDeviceName As String
Private mlngPrintCopies As Long

Dim mcbsThis As Object          'CommandBar控件


'------------------------------------------------------------
'以下为公共方法
'------------------------------------------------------------
Public Sub zlDefCommandBars(ByVal cbsThis As Object)
    '-----------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar
    Set mcbsThis = cbsThis
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    
    
    '编辑菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "报告(&E)", cbrMenuBar.Index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "书写(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "修订(&U)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Copy, "复制(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "查阅(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXML, "导出XML(&L)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Search, "报告检索(&S)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "验证签名(&V)")
    End With
    
    '工具栏定义
    '-----------------------------------------------------
    Set cbrToolBar = cbsThis(2)
    For Each cbrControl In cbrToolBar.Controls '先求出前面的最后一个Control
        If Val(Left(cbrControl.ID, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.ID, 1)) <> conMenu_ManagePopup Then
            Set cbrControl = cbrToolBar.Controls(cbrControl.Index - 1): Exit For
        End If
    Next
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "书写", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "修订", cbrControl.Index + 1)
    End With

    '命令的快键绑定
    '-----------------------------------------------------
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("O"), conMenu_File_Open
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("U"), conMenu_Edit_Audit
'        .Add FCONTROL, Asc("C"), conMenu_Edit_Copy
    End With
    
    '设置不常用命令
    '-----------------------------------------------------
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_ExportToXML
        .AddHiddenCommand conMenu_Tool_Search
    End With
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strInfo As String
    If mblnMoved And (Control.ID = conMenu_File_Open Or Control.ID = conMenu_File_ExportToXML Or _
                    Control.ID = conMenu_Edit_Modify Or Control.ID = conMenu_Edit_Delete Or Control.ID = conMenu_Edit_Audit) Then
        MsgBox "该病人的本次数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                    "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Select Case Control.ID
    Case conMenu_File_Open
        '病历阅读
        If mbyeEPR编辑方式 = 0 Then
            Dim fViewDoc As New frmEPRView
            fViewDoc.ShowMe Me, mlngEPR报告ID
        Else
            Call mObjTabEprView.InitOpenEPR(Me, cprEM_修改, cprET_单病历编辑, mlngEPR报告ID, True, 0, mintPati来源, mlngPati病人ID, mlngPati主页ID, mlngPati婴儿, mlngDeptId, mlngOrderId, mstrPrivs, mblnMoved, , Val(gstrESign))
        End If
    Case conMenu_File_PrintSet
        Call zlPrintSet
    Case conMenu_File_Preview
        If mblnCanPrint Then
            Call zlEPRPrint(True)
        Else
            MsgBox "当前报告未审核，不能打印，请检查！", vbInformation, gstrSysName
        End If
    Case conMenu_File_Print
        If mblnCanPrint Then
            Call zlEPRPrint(False)
        Else
            MsgBox "当前报告未审核，不能打印，请检查！", vbInformation, gstrSysName
        End If
    Case conMenu_File_BatPrint
        If mblnCanPrint Then
            Call zlEPRPrint(False, True)
        Else
            MsgBox "当前报告未审核，不能打印，请检查！", vbInformation, gstrSysName
        End If
    Case conMenu_File_ExportToXML:
        '导出到XML文件
        Dim strF As String
        dlgThis.Filename = "病历_" & mstrEPR报告名称 & "(" & mlngEPR报告ID & "," & mintEPR最后版本 & ").xml"
        dlgThis.Filter = "*.XML|*.xml|*.*|*.*"
        dlgThis.CancelError = True
        On Error Resume Next
        dlgThis.ShowSave
        strF = dlgThis.Filename
        If Err.Number <> 0 Then Err.Clear: Exit Sub
        On Error GoTo errHand
        If gobjFSO.FileExists(strF) Then
            DoEvents
            If MsgBox("该文件已经存在，是否覆盖？", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then Exit Sub
        End If
        
        If mbyeEPR编辑方式 = 0 Then
            Dim DocXML As New cEPRDocument
            '普通住院病历
            DocXML.InitAndOpenEPR mlngEPR报告ID, mintEPR最后版本, , True
            If DocXML.ExportToXMLFile(DocXML.frmEditor.Editor1, strF) Then
                DoEvents
                MsgBox "成功导出为XML文件！" & vbCrLf & "文件名:" & strF, vbOKOnly + vbInformation, gstrSysName
            End If
        Else    '表格式病历
            Call mObjTabEprView.InitOpenEPR(Me, cprEM_修改, cprET_单病历编辑, mlngEPR报告ID, False, 0, mintPati来源, mlngPati病人ID, mlngPati主页ID, mlngPati婴儿, mlngDeptId, mlngOrderId, mstrPrivs, mblnMoved)
            If mObjTabEprView.zlExportXML(strF) Then
                MsgBox "成功导出为XML文件！" & vbCrLf & "文件名:" & strF, vbOKOnly + vbInformation, gstrSysName
            End If
        End If
    Case conMenu_Edit_Modify
        If CheckCommitCheckup = False Then Exit Sub '出院病人病案提交审查后禁止修改
        Dim frmThis As Form, bFinded As Boolean
        If mbyeEPR编辑方式 = 1 Then '表格式病历
            If Not mObjTabEpr Is Nothing Then
                bFinded = mObjTabEpr.Showfrm(mlngEPR报告ID, mlngPati病人ID, mlngPati主页ID, mintPati来源, mlngDeptId)
            End If
            If bFinded = False Then
                Set mObjTabEpr = New cTableEPR
                mObjTabEpr.EPRFileInfo.lngModule = mlngModule
                mObjTabEpr.InitOpenEPR Me, IIf(mlngEPR报告ID = 0, cprEM_新增, cprEM_修改), cprET_单病历编辑, IIf(mlngEPR报告ID = 0, mlngEPR定义ID, mlngEPR报告ID), True, 0, mintPati来源, _
                    mlngPati病人ID, mlngPati主页ID, mlngPati婴儿, mlngDeptId, mlngOrderId, mstrPrivs, mblnMoved, , Val(gstrESign)
                RaiseEvent AfterOpen(cprET_单病历编辑)
            End If
            mlngSingCount = mObjTabEpr.Signs.Count
        Else
            For Each frmThis In Forms
                If frmThis.Name = "frmMain" Then
                    If Not frmThis.Document Is Nothing Then
                        If frmThis.Document.EPRPatiRecInfo.医嘱id = mlngOrderId And frmThis.ChildMode = False Then
                            frmThis.Show
                            bFinded = True
                        End If
                    Else
                        Unload frmThis
                    End If
                End If
            Next
            If bFinded = False Then
                strInfo = Clipboard.GetText '暂存
                Set mobjDoc = New cEPRDocument
                If mlngEPR报告ID = 0 Then
                    mobjDoc.InitEPRDoc cprEM_新增, cprET_单病历编辑, mlngEPR定义ID, _
                        mintPati来源, mlngPati病人ID, mlngPati主页ID, mlngPati婴儿, mlngDeptId, mlngOrderId, mblnMoved
                Else
                    mobjDoc.InitEPRDoc cprEM_修改, cprET_单病历编辑, mlngEPR报告ID, _
                        mintPati来源, mlngPati病人ID, mlngPati主页ID, mlngPati婴儿, mlngDeptId, mlngOrderId, mblnMoved
                End If
                
                mobjDoc.EPRFileInfo.lngModule = mlngModule
                
                mobjDoc.ShowEPREditor Me, mblnCanPrint
                If Trim(strInfo) <> "" Then '恢复粘贴板内容
                    DoEvents
                    Clipboard.SetText strInfo
                End If
                RaiseEvent AfterOpen(cprET_单病历编辑)
            End If
            mlngSingCount = mobjDoc.Signs.Count
        End If
        
    Case conMenu_Edit_Delete
        strInfo = "真的删除这份“" & mstrEPR报告名称 & "”吗？"
        If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        gstrSQL = "Zl_电子病历记录_Delete(" & mlngEPR报告ID & ")"
        Err = 0: On Error GoTo errHand
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Err = 0: On Error GoTo 0
        RaiseEvent AfterDeleted(mlngOrderId)
        Call Me.zlRefresh(mlngOrderId, mlngDeptId, mblnEdit, True, mblnMoved, mblnCanPrint, mlngModule)
    
    Case conMenu_Edit_Audit
        If CheckCommitCheckup = False Then Exit Sub '出院病人病案提交审查后禁止修改
        If mbyeEPR编辑方式 = 1 Then '表格式病历
            If Not mObjTabEpr Is Nothing Then
                bFinded = mObjTabEpr.Showfrm(mlngEPR报告ID, mlngPati病人ID, mlngPati主页ID, mintPati来源, mlngDeptId)
            End If
            If bFinded = False Then
                Set mObjTabEpr = New cTableEPR
                mObjTabEpr.EPRFileInfo.lngModule = mlngModule
                mObjTabEpr.InitOpenEPR Me, cprEM_修改, cprET_单病历审核, mlngEPR报告ID, True, 0, mintPati来源, _
                    mlngPati病人ID, mlngPati主页ID, mlngPati婴儿, mlngDeptId, mlngOrderId, mstrPrivs, mblnMoved, , Val(gstrESign)
                RaiseEvent AfterOpen(cprET_单病历审核)
            End If
            mlngSingCount = mObjTabEpr.Signs.Count
        Else
            Dim frmAudit As Form, bFindedAudit As Boolean
            For Each frmAudit In Forms
                If frmAudit.Name = "frmMain" Then
                    If Not frmAudit.Document Then
                        If frmAudit.Document.EPRPatiRecInfo.医嘱id = mlngOrderId And frmAudit.ChildMode = False Then
                            frmAudit.Show
                            bFindedAudit = True
                        End If
                    Else
                        Unload frmAudit
                    End If
                End If
            Next
            If bFindedAudit = False Then
                Set mobjDoc = New cEPRDocument
                mobjDoc.InitEPRDoc cprEM_修改, cprET_单病历审核, mlngEPR报告ID, _
                    mintPati来源, mlngPati病人ID, mlngPati主页ID, mlngPati婴儿, mlngDeptId, mlngOrderId
                    
                mobjDoc.EPRFileInfo.lngModule = mlngModule
                
                mobjDoc.ShowEPREditor Me
                RaiseEvent AfterOpen(cprET_单病历审核)
            End If
            mlngSingCount = mobjDoc.Signs.Count
        End If
        
    Case conMenu_Edit_Copy
        Call edtThis.Copy
    Case conMenu_Tool_Search: frmEPRSearchMan.ShowSearchReport Me, mlngDeptId
    Case conMenu_View_Refresh:  Call Me.zlRefresh(mlngOrderId, mlngDeptId, mblnEdit, True, mblnMoved, mblnCanPrint, mlngModule)
    Case conMenu_Help_Help
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Tool_SignVerify
        Call VerifySignature(Me, mlngEPR报告ID, mblnMoved)
    End Select
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error Resume Next
    If Me.Visible = False Then Exit Sub
    Select Case Control.ID
    Case conMenu_File_Open
        Control.Enabled = (Val(mlngEPR报告ID) <> 0)
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_ExportToXML
        Control.Enabled = (Val(mlngEPR报告ID) <> 0 And InStr(1, mstrPrivs, "报告打印") > 0)
    Case conMenu_Edit_Modify
        Control.Enabled = (mblnEdit And mlngOrderId > 0 And InStr(1, mstrPrivs, "报告书写") > 0)
        If Control.Enabled And mlngEPR报告ID > 0 Then
            If Control.Enabled Then Control.Enabled = (mlngDeptId = mlngEPR科室ID)   '本科病历才可以改
            If mstrEPR完成时间 = "" Then
                Control.Enabled = (InStr(1, mstrPrivs, "他人报告") > 0 Or mstrEPR创建人 = Trim(gstrUserName))
            ElseIf mstrEPR归档人 = "" And mintEPR最后版本 <= 1 And InStr(1, ",1,2,4,", mintEPR签名级别) > 0 Then
                Control.Enabled = (InStr(1, mstrPrivs, "他人报告") > 0 Or InStr(1, mstrEPR保存人, Trim(gstrUserName)) > 0)
            Else
                Control.Enabled = False
            End If
        End If
    Case conMenu_Edit_Delete
        Control.Enabled = (mblnEdit And mlngEPR报告ID <> 0) And (InStr(1, mstrPrivs, "报告书写") > 0 Or InStr(1, mstrPrivs, "强制删除") > 0)
        If Control.Enabled And InStr(1, mstrPrivs, "强制删除") > 0 Then Exit Sub                '具备强制删除权限，则不进行后续的判断
        If Control.Enabled Then Control.Enabled = (mlngDeptId = mlngEPR科室ID)     '本科病历才可以删
        If Control.Enabled Then Control.Enabled = (mstrEPR完成时间 = "")                '未完成病历可以删
        If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "他人报告") > 0 Or mstrEPR创建人 = Trim(gstrUserName))
    
    Case conMenu_Edit_Audit
        Control.Enabled = (mblnEdit And mlngPati病人ID > 0 And InStr(1, mstrPrivs, "报告修订") > 0)
        If Control.Enabled Then Control.Enabled = (mlngDeptId = mlngEPR科室ID)      '本科病历才可以审核
        If Control.Enabled Then Control.Enabled = (mstrEPR完成时间 <> "")           '完成病历才可以审
        If Control.Enabled Then Control.Enabled = (mstrEPR归档人 = "")              '未归档病历可以审
    Case conMenu_Tool_Search
        Control.Enabled = mblnSearch
    Case conMenu_Edit_Copy
        Control.Visible = mbyeEPR编辑方式 = 0
        If Control.Visible Then Control.Enabled = edtThis.Selection.EndPos <> edtThis.Selection.StartPos
    End Select

End Sub
Public Sub RefreshList()
    zlRefresh mlngOrderId, mlngDeptId, mblnEdit, True, mblnMoved, mblnCanPrint, mlngModule
End Sub
Public Function zlRefresh(ByVal lngOrderId As Long, ByVal lngDeptId As Long, ByVal blnEdit As Boolean, _
                            Optional ByVal blnForce As Boolean, Optional ByVal blnMoved As Boolean, Optional ByVal blnCanPrint As Boolean = True, Optional ByVal lngModule As Long) As Long
'异常返回0,否则返回1
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    
    strTemp = ""
    
    If mlngDeptId <> lngDeptId Or gstrESign = "" Then '提取是否本部门启用电子签名,科室变更或没取过时提取
        gstrESign = getPassESign(7, lngDeptId)
    End If
    
    mlngDeptId = lngDeptId: mblnEdit = blnEdit
    If mlngOrderId = lngOrderId And blnForce = False Then Exit Function
    mlngOrderId = lngOrderId: mblnMoved = blnMoved: mblnCanPrint = blnCanPrint: mlngModule = lngModule
    
    
    Err = 0: On Error GoTo errHand
    
    mintPati来源 = 0: mlngPati病人ID = 0: mlngPati主页ID = 0: mlngPati婴儿 = 0
    gstrSQL = "Select l.病人来源, l.病人id, l.挂号单, l.主页id, l.婴儿, a.病历文件id" & vbNewLine & _
            "From 病人医嘱记录 l, 病历单据应用 a" & vbNewLine & _
            "Where l.诊疗项目id = a.诊疗项目id(+) And a.应用场合(+) = Decode(l.病人来源, 2, 2, 4, 4, 1) And l.Id = [1]"
    If mblnMoved Then gstrSQL = Replace(gstrSQL, "病人医嘱记录", "H病人医嘱记录")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngOrderId)
    With rsTemp
        If .RecordCount > 0 Then
            mintPati来源 = Val("" & !病人来源)
            mlngPati病人ID = Val("" & !病人ID)
            If mintPati来源 <> 1 Then
                mlngPati主页ID = Val("" & !主页ID)
            Else
                strTemp = "" & !挂号单
            End If
            mlngPati婴儿 = Val("" & !婴儿)
            mlngEPR定义ID = Val("" & !病历文件id)
        End If
    End With
    
    If mlngEPR定义ID <> 0 Then
        gstrSQL = "Select 保留 From 病历文件列表 where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngEPR定义ID)
        mbyeEPR编辑方式 = IIf(NVL(rsTemp!保留, 0) = 2, 1, 0)
    Else
        mbyeEPR编辑方式 = 0
    End If
    
    If mintPati来源 = 1 Then
        gstrSQL = "Select ID From 病人挂号记录 Where NO = [1] and 记录性质=1  and 记录状态=1"
        If mblnMoved Then gstrSQL = Replace(gstrSQL, "病人挂号记录", "H病人挂号记录")
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTemp)
        If rsTemp.RecordCount > 0 Then
            mlngPati主页ID = rsTemp!ID
        End If
    End If

    mlngEPR报告ID = 0: mstrEPR报告名称 = "": mstrEPR创建人 = "": mstrEPR保存人 = "": mstrEPR归档人 = ""
    mlngEPR科室ID = 0: mstrEPR完成时间 = "": mintEPR签名级别 = 0: mintEPR签名版本 = 0: mintEPR最后版本 = 1
    Me.lblNote.Caption = "提示： 尚未书写报告！"
    gstrSQL = "Select l.Id, l.病历名称, l.创建人, l.保存人, l.完成时间, l.最后版本, l.签名级别, l.归档人, l.科室id, l.保存人," & vbNewLine & _
            "       Nvl(Max(c.开始版), 0) As 签名版本,l.编辑方式" & vbNewLine & _
            "From 电子病历记录 l, 电子病历内容 c, 病人医嘱报告 r" & vbNewLine & _
            "Where l.Id = c.文件id(+) And l.Id = r.病历id And l.病历种类 = 7 And c.对象类型(+) = 8 And r.医嘱id = [1]" & vbNewLine & _
            "Group By l.Id, l.病历名称, l.创建人, l.保存人, l.完成时间, l.最后版本, l.签名级别, l.归档人, l.科室id, l.保存人,编辑方式"
    If mblnMoved Then
        gstrSQL = Replace(gstrSQL, "电子病历记录", "H电子病历记录")
        gstrSQL = Replace(gstrSQL, "电子病历内容", "H电子病历内容")
        gstrSQL = Replace(gstrSQL, "病人医嘱报告", "H病人医嘱报告")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngOrderId)
    With rsTemp
        If Not .EOF Then
            mlngEPR报告ID = !ID
            mstrEPR报告名称 = "" & !病历名称
            mstrEPR创建人 = "" & !创建人
            mstrEPR保存人 = "" & !保存人
            mstrEPR完成时间 = Trim("" & Format(!完成时间, "yyyy年MM月dd日hh时mm分"))
            mintEPR最后版本 = Val("" & !最后版本)
            mintEPR签名级别 = Val("" & !签名级别)
            mstrEPR归档人 = "" & !归档人
            mintEPR签名版本 = Val("" & !签名版本)
            mlngEPR科室ID = Val("" & !科室ID)
            
            If mstrEPR完成时间 = "" Then
                Me.lblNote.Caption = "提示：当前报告正在由" & mstrEPR保存人 & "书写，尚未完成…"
            ElseIf mintEPR签名级别 = mintEPR最后版本 Then
                Me.lblNote.Caption = "提示：报告完成于" & mstrEPR完成时间 & "，" & !保存人 & IIf(mintEPR最后版本 = 1, "书写", "最后修订。")
            Else
                Me.lblNote.Caption = "提示：报告完成于" & mstrEPR完成时间 & "，" & !保存人 & IIf(mintEPR最后版本 = 1, "书写", "正在修订…")
            End If
            
            mbyeEPR编辑方式 = !编辑方式 '有报告时以报告实际编辑方式为准
        End If
        '调用显示文档
        If mbyeEPR编辑方式 = 1 And mlngEPR报告ID <> 0 Then '表格式病历并且已经写过报告
            With edtThis
                .Text = vbCrLf & Space(4) & "该文件为表格式病历，正在加载中..."
                .SelectAll
                .ForceEdit = True
                .Selection.Font.Name = "宋体": .Selection.Font.Size = 10.5
                .SelLength = 0
                .ForceEdit = False
            End With
            Call mObjTabEprView.InitOpenEPR(Me, cprEM_修改, cprET_单病历审核, mlngEPR报告ID, False, 0, mintPati来源, mlngPati病人ID, mlngPati主页ID, mlngPati婴儿, mlngDeptId, mlngOrderId, mstrPrivs, mblnMoved)
            Call mObjTabEprView.zlRefreshDockfrm
            
            dkpMan.FindPane(conPane_Content).Close
            dkpMan.ShowPane conPane_Table
            dkpMan.RedrawPanes
        Else
            Call zlRefDocment(mlngEPR报告ID)
            
            dkpMan.FindPane(conPane_Table).Close
            dkpMan.ShowPane conPane_Content
            dkpMan.RedrawPanes
        End If
        Call mfrmAnnex.zlRefresh(mlngEPR报告ID, IIf(mblnEdit, mstrPrivs, ""))
    End With
    If mlngEPR报告ID <> 0 Then zlRefresh = 1
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub zlRefDocment(ByVal lngEPRid As Long)
    '功能：刷新病历显示内容；
    '参数：lngEPRId-电子病历记录ID
    Dim mstrPrivs As String, blnPrivacy As Boolean, Elements As New cEPRElements
    Dim rs As New ADODB.Recordset, lngKey As Long
    
    Dim strTemp As String, strZipFile As String
    
    Me.edtThis.Freeze
    Me.edtThis.ReadOnly = False
    Me.edtThis.NewDoc
    strZipFile = zlBlobRead(5, lngEPRid, , mblnMoved)
    If gobjFSO.FileExists(strZipFile) Then
        strTemp = zlFileUnzip(strZipFile)
        If gobjFSO.FileExists(strTemp) Then
            '打开文件
            Me.edtThis.OpenDoc strTemp
            gobjFSO.DeleteFile strTemp, True
        End If
        gobjFSO.DeleteFile strZipFile, True
        Me.edtThis.SelStart = 0
    End If
    If lngEPRid > 0 Then
        '设置页面格式
        Dim mEPRFileInfo As New cEPRFileDefineInfo
        Err = 0: On Error GoTo errHand
        gstrSQL = "Select c.ID, a.格式 From   病历页面格式 a, 病历文件列表 b, 电子病历记录 c " & _
                " Where  c.文件id = b.id And a.种类 = b.种类 And a.编号 = b.页面 And c.ID = [1]"
        If mblnMoved Then gstrSQL = Replace(gstrSQL, "电子病历记录", "H电子病历记录")
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngEPRid)
        If Not rs.EOF Then
            mEPRFileInfo.格式 = zlCommFun.NVL(rs("格式").Value)
            mEPRFileInfo.SetFormat Me.edtThis, mEPRFileInfo.格式
            Me.edtThis.ResetWYSIWYG
        End If
        Set mEPRFileInfo = Nothing
    End If
    Me.edtThis.UnFreeze
    edtThis.RefreshTargetDC
    Me.edtThis.ReadOnly = True
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Public Sub ConfigPrint(ByVal strPrintDevice As String, ByVal lngCopies As Long)
'配置打印机
    mstrPrinterDeviceName = strPrintDevice
    mlngPrintCopies = lngCopies
End Sub


Private Sub zlEPRPrint(blnPreview As Boolean, Optional blnStilly As Boolean)
    '-------------------------------------------------
    '功能: 打印当前文档
    '参数:  blnPreview 预览
    '       blnStilly  强制静默打印
    '-------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim intOutMode As Integer, strBillNo As String, blnNoAsk As Boolean
    blnNoAsk = (zlDatabase.GetPara("NoAsk", glngSys, 1070, 0) = "1")
    If blnStilly Then blnNoAsk = True
    
    If Trim(mstrPrinterDeviceName) = "" Then
        mstrPrinterDeviceName = Printer.DeviceName
        mlngPrintCopies = Printer.Copies
    End If
    
    intOutMode = 0: strBillNo = ""
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select f.通用, f.编号 From 电子病历记录 l, 病历文件列表 f Where l.文件id = f.Id And l.Id = [1]"
    If mblnMoved Then gstrSQL = Replace(gstrSQL, "电子病历记录", "H电子病历记录")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngEPR报告ID)
    If rsTemp.RecordCount > 0 Then
        intOutMode = Val("" & rsTemp!通用)
        strBillNo = "ZLCISBILL" & Format(rsTemp!编号, "00000") & "-2"
    End If
    
    If intOutMode <> 2 Then
        If mbyeEPR编辑方式 = 0 Then 'RichEpr编辑
            '直接打印
            Set mfrmPrintPreview = New frmPrintPreview
            Call mfrmPrintPreview.DoMultiDocPreview(Me, cpr诊疗报告, , , cpr诊疗报告, , mlngEPR报告ID, Not blnPreview, False, blnNoAsk, mblnMoved, , mstrPrinterDeviceName, mlngPrintCopies)
            Unload mfrmPrintPreview: Set mfrmPrintPreview = Nothing
        ElseIf mlngEPR报告ID <> 0 Then '表格式病历
            mObjTabEprView.InitOpenEPR Me, cprEM_修改, cprET_单病历编辑, mlngEPR报告ID, False, 0, mintPati来源, mlngPati病人ID, mlngPati主页ID, mlngPati婴儿, mlngDeptId, mlngOrderId, mstrPrivs, mblnMoved
            mObjTabEprView.zlPrintDoc Me, blnPreview
        End If
    Else
        '自定义报表打印
        Dim strExseNo As String, intExseKind As Integer
        Dim objFile As New Scripting.FileSystemObject
        Dim strPicPath As String, strPicFile As String
        Dim cTable As cEPRTable, oPicture As StdPicture
        Dim aryPara(19) As String, intPCount As Integer
        Dim aryFlagPara(1) As String
        Dim intRows As Integer, intCols As Integer
        Dim dcmImages As New DicomImages, dcmResultImage As DicomImage
        Dim i As Integer
        
        gstrSQL = "Select 记录性质, No From 病人医嘱发送 Where 医嘱id = [1]"
        If mblnMoved Then gstrSQL = Replace(gstrSQL, "病人医嘱发送", "H病人医嘱发送")
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngOrderId)
        If rsTemp.RecordCount = 0 Then Exit Sub
        strExseNo = "" & rsTemp!NO
        intExseKind = Val("" & rsTemp!记录性质)
        If mobjReport Is Nothing Then Set mobjReport = New clsReport
        If Not blnNoAsk Then
            If mobjReport.ReportPrintSet(gcnOracle, glngSys, strBillNo, Me) = False Then Exit Sub
        End If
        
        '获取图像
        strPicPath = App.Path & "\TmpImage\"
        If objFile.FolderExists(strPicPath) = False Then objFile.CreateFolder strPicPath
        
        '获取报告图象(包括标记图)生成本地文件
        '一个报告表格中可能排列多个报告图
        intPCount = 0
        gstrSQL = "Select Id As 表格Id From 电子病历内容" & vbNewLine & _
        "       Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
        "       Order By 对象序号"
        If mblnMoved Then gstrSQL = Replace(gstrSQL, "电子病历内容", "H电子病历内容")
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngEPR报告ID)
        Do While Not rsTemp.EOF
            Set cTable = New cEPRTable
            If cTable.GetTableFromDB(cprET_单病历审核, mlngEPR报告ID, Val("" & rsTemp!表格Id), , IIf(mblnMoved, "H电子病历内容", "电子病历内容")) Then
                For i = 1 To cTable.Pictures.Count
                    strPicFile = strPicPath & "PACSPic" & i & ".JPG"
                    If objFile.FileExists(strPicFile) Then objFile.DeleteFile strPicFile, True
                    If cTable.Pictures(i).PictureType = EPRMarkedPicture Then
                        Set oPicture = cTable.Pictures(i).DrawFinalPic
                    Else
                        Set oPicture = cTable.Pictures(i).OrigPic
                    End If
                    SavePicture oPicture, strPicFile
                    If objFile.FileExists(strPicFile) Then
                        '保存标记图和图象的路径
                        If cTable.Pictures(i).PictureType = EPRMarkedPicture Then
                            aryFlagPara(0) = strPicFile
                        Else
                            aryPara(intPCount) = strPicFile
                            dcmImages.AddNew
                            dcmImages(dcmImages.Count).FileImport strPicFile, "BMP"
                            intPCount = intPCount + 1
                            If intPCount > UBound(aryPara) Then Exit Do
                        End If
                    End If
                Next
            End If
            rsTemp.MoveNext
        Loop
        
        '判断是否需要自动组合图象，自定义报表中只定义了一个图象框，则自动组合图象
        '重新查一次数据库
        gstrSQL = "Select b.名称,b.W,b.H From zlReports a, zlRptItems b" & vbNewLine & _
        "       Where a.Id = b.报表id And a.编号 = [1] And Nvl(b.下线, 0) = 1 And b.类型 = 11 And b.格式号 = 1 And b.名称 not like '标记%'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strBillNo)
        If rsTemp.RecordCount = 1 And intPCount >= 1 Then
            '组合图象
            ResizeRegion intPCount, rsTemp("W"), rsTemp("H"), intRows, intCols
            Set dcmResultImage = AssembleImage(dcmImages, intRows, intCols, rsTemp("H"), rsTemp("W"))
            dcmResultImage.FileExport Right(aryPara(0), Len(aryPara(0)) - InStr(aryPara(0), "=")), "JPEG"
        End If
        
        '获取自定义报表中的图象定义
        intPCount = 0
        gstrSQL = "Select b.名称 From zlReports a, zlRptItems b" & vbNewLine & _
        "       Where a.Id = b.报表id And a.编号 = [1] And Nvl(b.下线, 0) = 1 And b.类型 = 11 And b.格式号 = 1" & vbNewLine & _
        "       Order By b.名称" 'Trunc(b.y/567),Trunc(b.x/567)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strBillNo)
        Do While Not rsTemp.EOF
            If aryPara(intPCount) = "" Then Exit Do '报表中的图形比报告中多
            '分别装载标记图和报告图像
            If InStr(rsTemp!名称, "标记") <> 0 Then
                If aryFlagPara(0) <> "" Then aryFlagPara(0) = rsTemp!名称 & "=" & aryFlagPara(0)
            Else
                aryPara(intPCount) = rsTemp!名称 & "=" & aryPara(intPCount)
                intPCount = intPCount + 1
                If intPCount > UBound(aryPara) Then Exit Do
            End If
            rsTemp.MoveNext
        Loop
        For i = intPCount To UBound(aryPara) '报表中的图形比报告中少
            If aryPara(i) Like "*=*" Then aryPara(i) = ""
        Next
        
        '调用报表
       Call mobjReport.ReportOpen(gcnOracle, glngSys, strBillNo, Nothing, _
            "NO=" & strExseNo, "性质=" & intExseKind, "医嘱ID=" & mlngOrderId, aryFlagPara(0), _
            aryPara(0), aryPara(1), aryPara(2), aryPara(3), aryPara(4), aryPara(5), _
            aryPara(6), aryPara(7), aryPara(8), aryPara(9), aryPara(10), aryPara(11), _
            aryPara(12), aryPara(13), aryPara(14), aryPara(15), aryPara(16), aryPara(17), _
            aryPara(18), aryPara(19), IIf(blnPreview, 1, 2))
    End If
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Public Sub RefPacsPic()
'功能: 刷新正在编辑报告的PACS图片
    If mbyeEPR编辑方式 = 0 Then
        Dim frmThis As Form
        For Each frmThis In Forms
            If frmThis.Name = "frmMain" Then
                If Not frmThis.Document Then
                    With frmThis.Document
                        If .EPRPatiRecInfo.医嘱id = mlngOrderId Then
                            Call frmThis.RefPacsPic
                        End If
                    End With
                End If
                
                Exit Sub
            End If
        Next
    Else
        mObjTabEpr.zlRefreshPacsPic mlngOrderId
    End If
End Sub

'------------------------------------------------------------
'以下为窗体事件响应
'------------------------------------------------------------
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_Note
        Item.Handle = picNote.hWnd
    Case conPane_Content
        Item.Handle = picRichEdit.hWnd
    Case conPane_Table
        Item.Handle = mObjTabEprView.zlGetForm.hWnd
    Case conPane_Annex
        Item.Handle = mfrmAnnex.hWnd
    End Select
End Sub

Private Sub edtThis_KeyDown(ViewMode As zlRichEditor.ViewModeEnum, KeyCode As Integer, Shift As Integer)
    If Shift = 2 And KeyCode = vbKeyC Then
        Call edtThis.Copy
    End If
End Sub

Private Sub Form_Load()
Dim Pane1 As Pane, pane2 As Pane, pane3 As Pane, Pane4 As Pane
    Set mfrmAnnex = New frmDockAnnex
    Set mObjTabEprView = New cTableEPR
    mObjTabEprView.InitTableEPR gcnOracle, glngSys, gstrDbOwner
    mblnSearch = (InStr(1, GetPrivFunc(glngSys, 1273), "基本") > 0)
    mstrPrivs = GetPrivFunc(glngSys, 1258)
    
    Set Pane1 = dkpMan.CreatePane(conPane_Note, 200, 15, DockTopOf, Nothing)
    Pane1.Title = "提示": Pane1.MinTrackSize.Height = 360 / Screen.TwipsPerPixelY: Pane1.MaxTrackSize.Height = 360 / Screen.TwipsPerPixelY
    Pane1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set pane2 = dkpMan.CreatePane(conPane_Content, 1200, 200, DockBottomOf, Nothing)
    pane2.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set pane3 = dkpMan.CreatePane(conPane_Table, 1200, 200, DockBottomOf, Nothing)
    pane3.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    pane3.Close
    
    Set Pane4 = dkpMan.CreatePane(conPane_Annex, 200, 15, DockBottomOf, Nothing)
    Pane4.Title = "附件": Pane4.MinTrackSize.Height = 360 / Screen.TwipsPerPixelY: Pane4.MaxTrackSize.Height = 360 / Screen.TwipsPerPixelY
    Pane4.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    With dkpMan
        .VisualTheme = ThemeOffice2003
        .Options.HideClient = True
        .Options.UseSplitterTracker = True
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = False
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Unload mObjTabEprView.zlGetForm
    Unload mfrmAnnex
    Unload mfrmPrintPreview
    
    Set mfrmAnnex = Nothing
    Set mobjReport = Nothing
    Set mfrmPrintPreview = Nothing
    Set mobjDoc = Nothing
    Set mObjTabEpr = Nothing
    Set mObjTabEprView = Nothing
    Set mcbsThis = Nothing
End Sub

Private Sub edtThis_RequestRightMenu(ViewMode As zlRichEditor.ViewModeEnum, Shift As Integer, X As Single, Y As Single)
    If mcbsThis Is Nothing Then Exit Sub
Dim Popup As CommandBar
Dim cbrControl As CommandBarControl
    
    Set Popup = mcbsThis.Add("Popup", xtpBarPopup)
    With Popup.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "书写(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "修订(&U)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Copy, "复制(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "查阅(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXML, "导出XML(&L)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Search, "报告检索(&S)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "验证签名(&V)")
        Popup.ShowPopup
    End With
End Sub
Public Sub EditorClosed(lngOrderId As Long)
    RaiseEvent AfterClosed(lngOrderId)
End Sub

Private Sub mfrmPrintPreview_PrintEpr(ByVal lngRecordId As Long)
    Call Event_AfterPrinted(lngRecordId)
End Sub

Private Sub mobjDoc_AfterSaved(lngRecordId As Long)
    Dim rsTemp As New ADODB.Recordset, lng医嘱id As Long
    Dim lngSaveType As Long
        
    '如果是当前报告，则刷新当前显示内容
    If lngRecordId = mlngEPR报告ID Then
        Call Me.zlRefresh(mlngOrderId, mlngDeptId, mblnEdit, True, mblnMoved, mblnCanPrint, mlngModule)
    End If
    
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select 医嘱id,查阅状态 From 病人医嘱报告 Where 病历Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngRecordId)
    If rsTemp.RecordCount > 0 Then
        lng医嘱id = Val("" & rsTemp!医嘱id)
        '取消下面的注释，根据调用的模块填写该过程的权限脚本
'        If Val("" & rsTemp!查阅状态) = 1 Then
'            gstrSQL = "Zl_报告查阅记录_Cancel(" & lng医嘱ID & "," & lngRecordId & ",Null)"
'            Call zldatabase.ExecuteProcedure(gstrSQL, "更新查阅状态")
'        End If

        '区分是表格病历编辑器还是全文病历编辑器
        If mbyeEPR编辑方式 = 1 Then '表格式病历
            If mlngSingCount = mObjTabEpr.Signs.Count Then
                lngSaveType = 0 '普通保存
            Else
                If mObjTabEpr.ET <> TabET_单病历审核 Then
                    lngSaveType = 1 '诊断签名
                Else
                    lngSaveType = 2 '审核签名
                End If
            End If
        Else    '全文病历编辑器
            If mlngSingCount = mobjDoc.Signs.Count Then
                lngSaveType = 0 '普通保存
            ElseIf mlngSingCount < mobjDoc.Signs.Count Then
                If mobjDoc.Signs(mobjDoc.Signs.Count).签名级别 > cprSL_经治 Then
                    lngSaveType = 2 '审核签名
                Else
                    lngSaveType = 1 '诊断签名
                End If
            End If
            
            mlngSingCount = mobjDoc.Signs.Count
        End If
        
        RaiseEvent AfterSaved(lng医嘱id, lngSaveType)
    End If
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Public Sub Event_Saved(lngRecordId As Long)
    mobjDoc_AfterSaved lngRecordId
End Sub
Public Sub Event_AfterPrinted(lngRecordId As Long)
    Dim rsTemp As New ADODB.Recordset
    
    Err.Clear
    If mblnMoved Then Exit Sub '转储过的病人,不触发打印事件,目前的打印事件只是为影像检查记录标记标志
    On Error GoTo errHand
    gstrSQL = "Select 医嘱id From 病人医嘱报告 Where 病历Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngRecordId)
    If rsTemp.RecordCount > 0 Then
        RaiseEvent AfterPrinted(NVL(rsTemp!医嘱id, 0))
    End If
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mobjReport_AfterPrint(ByVal ReportNum As String)
    RaiseEvent AfterPrinted(mlngOrderId)
End Sub
Private Function CheckCommitCheckup() As Boolean
'功能：出院病人病案提交审查后,返回假，其余为真
Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrHandle
    '大医二院书写检查报告时，病案归档后不允许书写
    CheckCommitCheckup = False
    
    If mintPati来源 = 2 Then
        gstrSQL = "Select count(病人ID) 记录 From 病案主页 Where 病人id=[1] And 主页id =[2] And 出院日期 Is Not Null And Nvl(病案状态, 0) = 5"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询病案状态", mlngPati病人ID, mlngPati主页ID)
        If rsTemp!记录 >= 1 Then Exit Function '
    End If
    CheckCommitCheckup = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub picRichEdit_Resize()
    With edtThis
        .Top = 0: .Left = 0
        .Width = picRichEdit.ScaleWidth: .Height = picRichEdit.ScaleHeight
    End With
End Sub


