VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmMakeScripter 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   ScaleHeight     =   8925
   ScaleWidth      =   13080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picBack 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   6705
      Index           =   0
      Left            =   120
      ScaleHeight     =   6705
      ScaleWidth      =   10890
      TabIndex        =   0
      Top             =   855
      Width           =   10890
      Begin VB.PictureBox picBack 
         BorderStyle     =   0  'None
         Height          =   6465
         Index           =   1
         Left            =   15
         ScaleHeight     =   6465
         ScaleWidth      =   10830
         TabIndex        =   1
         Top             =   0
         Width           =   10830
         Begin VB.CommandButton cmdOpen 
            Height          =   300
            Index           =   2
            Left            =   6345
            Picture         =   "frmMakeScripter.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   525
            Width           =   315
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   2
            Left            =   960
            TabIndex        =   3
            Text            =   "E:\zlMipClientData"
            Top             =   540
            Width           =   5370
         End
         Begin RichTextLib.RichTextBox rtb 
            Height          =   4530
            Left            =   960
            TabIndex        =   5
            Top             =   885
            Width           =   5370
            _ExtentX        =   9472
            _ExtentY        =   7990
            _Version        =   393217
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            ScrollBars      =   3
            Appearance      =   0
            TextRTF         =   $"frmMakeScripter.frx":6852
         End
         Begin VB.Frame fra 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   105
            Left            =   60
            TabIndex        =   2
            Top             =   345
            Width           =   10635
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "输出结果"
            Height          =   180
            Index           =   1
            Left            =   105
            TabIndex        =   7
            Top             =   915
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "输出目录"
            Height          =   180
            Index           =   6
            Left            =   105
            TabIndex        =   6
            Top             =   570
            Width           =   720
         End
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmMakeScripter.frx":68EF
      Left            =   765
      Top             =   180
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager ImageManager1 
      Left            =   1800
      Top             =   300
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMakeScripter.frx":6903
   End
End
Attribute VB_Name = "frmMakeScripter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjFso As New FileSystemObject
Private mobjFile As TextStream

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim objPane As Object

    Set objPane = dkpMain.CreatePane(1, 100, 100, DockLeftOf, Nothing)
    objPane.Title = "XML格式(*.xsd)"
    objPane.Options = PaneNoCaption
    
    dkpMain.SetCommandBars cbsMain
    Call DockPannelInit(dkpMain)

End Sub

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim objFindKey As CommandBarControl
    Dim intPostion As Integer
    
    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
    
    Call CommandBarInit(cbsMain)
    cbsMain.VisualTheme = xtpThemeWhidbey
    Set cbsMain.Icons = ImageManager1.Icons
    
    cbsMain.Options.LargeIcons = False
    
    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份

    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    Set objControl = NewToolBar(objBar, xtpControlButton, 3, "开始输出", False, , xtpButtonIconAndCaption)
    objControl.IconId = 3
    
End Function


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
    Case 3
        rtb.Text = ""
        DoEvents
        If MakeMessageZLHISScript(txt(2).Text) Then
            
        End If
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
    Case 1
        Item.Handle = picBack(0).hWnd
    End Select
    
End Sub

Private Sub Form_Load()
    Call InitCommandBar
    Call InitDockPannel
End Sub

Private Sub picBack_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        picBack(1).Move 15, 15, picBack(Index).Width - 30, picBack(Index).Height - 30
    Case 1
        
        fra.Move fra.Left, fra.Top, picBack(Index).Width - fra.Left - 60
        txt(2).Move txt(2).Left, txt(2).Top, picBack(Index).Width - txt(2).Left - cmdOpen(2).Width - 120
        cmdOpen(2).Move txt(2).Left + txt(2).Width + 30
        
        rtb.Move rtb.Left, rtb.Top, picBack(Index).Width - rtb.Left - 60, picBack(Index).Height - rtb.Top - 60
        
    End Select
    
End Sub

Private Function MakeMessageZLHISScript(ByVal strFolder As String) As Boolean
    Dim strSQL As String
    Dim rsBusiness As ADODB.Recordset
    Dim strDataCode As String
    Dim strFile As String
    
    
    If mobjFso.FolderExists(strFolder) Then
        Call mobjFso.DeleteFolder(strFolder)
        DoEvents
    End If
    Call mobjFso.CreateFolder(strFolder)
        
    '1.输出公共部份
    strFile = strFolder & "\zlMipClientData.SQL"
    Call MakeBusinessDataScript("-", strFile)
    
    Call rtb.LoadFile(strFile)
    
    '2.输出业务部份
    strSQL = "Select a.data_code From zltools.zlmip_data_setup a Order By a.data_code"
    Set rsBusiness = New ADODB.Recordset
    rsBusiness.Open strSQL, gcnOracle
    If rsBusiness.BOF = False Then
        Do While Not rsBusiness.EOF
            strDataCode = rsBusiness("data_code").Value
'            strFile = "e:\zlMipClientData_" & strDataCode & ".SQL"

            If mobjFso.FolderExists(strFolder & "\" & strDataCode) Then
                Call mobjFso.DeleteFolder(strFolder & "\" & strDataCode)
                DoEvents
            End If
            Call mobjFso.CreateFolder(strFolder & "\" & strDataCode)
            
            strFile = strFolder & "\" & strDataCode & "\zlMipClientData.SQL"
            
            Call MakeBusinessDataScript(strDataCode, strFile)
            Call rtb.LoadFile(strFile)
            rsBusiness.MoveNext
        Loop
    End If
        
    MakeMessageZLHISScript = True
    
End Function

Private Function MakeBusinessDataScript(ByVal strDataCode As String, ByVal strFile As String) As Boolean
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strTemp As String
    Dim lngCount As Long
    Dim intGroup As Integer
        
    
    If strDataCode = "" Then strDataCode = "-"
    
    Set mobjFile = mobjFso.CreateTextFile(strFile, True)
    
        
    '------------------------------------------------------------------------------------------------------------------
    'zlmip_table
    mobjFile.WriteLine ""
    mobjFile.WriteLine "--zlmip_table"
    
    strSQL = "Select a.* From zltools.zlmip_table a Where Nvl(a.data_code,'-')='" & strDataCode & "' Order By a.tab_code"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open strSQL, gcnOracle
    If rsTemp.BOF = False Then
        intGroup = 0
        lngCount = 0
        Do While Not rsTemp.EOF
                                    
            intGroup = intGroup + 1
            If intGroup Mod 100 = 1 Then
                If intGroup > 1 Then mobjFile.WriteLine ""
                strTemp = "Insert Into zlmip_table(id,data_code,tab_type,tab_code,tab_title,tab_sqltext,tab_note)"
                mobjFile.WriteLine strTemp
                lngCount = 0
            End If
                        
            strTemp = "Select "
            strTemp = strTemp & "'" & rsTemp("id").Value & "'"
            strTemp = strTemp & ",'" & rsTemp("data_code").Value & "'"
            strTemp = strTemp & "," & rsTemp("tab_type").Value
            strTemp = strTemp & ",'" & rsTemp("tab_code").Value & "'"
            strTemp = strTemp & ",'" & rsTemp("tab_title").Value & "'"
            strTemp = strTemp & ",'" & Replace(rsTemp("tab_sqltext").Value, "'", "''") & "'"
            strTemp = strTemp & ",'" & rsTemp("tab_note").Value & "'"
            strTemp = strTemp & " From Dual"
            
            lngCount = lngCount + 1
            strTemp = IIf(lngCount = 1, "          ", "Union All ") & strTemp
            If lngCount = rsTemp.RecordCount Then strTemp = strTemp & ";"
                        
            mobjFile.WriteLine strTemp
            
            rsTemp.MoveNext
        Loop
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    'zlmip_tab_field
    mobjFile.WriteLine ""
    mobjFile.WriteLine "--zlmip_tab_field"
    
    strSQL = "Select b.* From zltools.zlmip_table a,zltools.zlmip_tab_field b Where Nvl(a.data_code,'-')='" & strDataCode & "' And a.id=b.tab_id Order By a.tab_code,b.fld_order"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open strSQL, gcnOracle
    If rsTemp.BOF = False Then
        intGroup = 0
        lngCount = 0
        Do While Not rsTemp.EOF
                                    
            intGroup = intGroup + 1
            If intGroup Mod 100 = 1 Then
                If intGroup > 1 Then mobjFile.WriteLine ""
                strTemp = "Insert Into zlmip_tab_field(tab_id,fld_order,fld_title,fld_type)"
                mobjFile.WriteLine strTemp
                lngCount = 0
            End If
                        
            strTemp = "Select "
            strTemp = strTemp & "'" & rsTemp("tab_id").Value & "'"
            strTemp = strTemp & "," & rsTemp("fld_order").Value
            strTemp = strTemp & ",'" & rsTemp("fld_title").Value & "'"
            strTemp = strTemp & "," & rsTemp("fld_type").Value
            strTemp = strTemp & " From Dual"
            
            lngCount = lngCount + 1
            strTemp = IIf(lngCount = 1, "          ", "Union All ") & strTemp
            If rsTemp.AbsolutePosition = rsTemp.RecordCount Or (intGroup Mod 100 = 0 And intGroup > 1) Then strTemp = strTemp & ";"
                        
            mobjFile.WriteLine strTemp
            
            rsTemp.MoveNext
        Loop
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    'zlmip_tab_parameter
    mobjFile.WriteLine ""
    mobjFile.WriteLine "--zlmip_tab_parameter"
    
    strSQL = "Select b.* From zltools.zlmip_table a,zltools.zlmip_tab_parameter b Where Nvl(a.data_code,'-')='" & strDataCode & "' And a.id=b.tab_id Order By a.tab_code,b.para_order"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open strSQL, gcnOracle
    If rsTemp.BOF = False Then
        intGroup = 0
        lngCount = 0
        Do While Not rsTemp.EOF
                                    
            intGroup = intGroup + 1
            If intGroup Mod 100 = 1 Then
                If intGroup > 1 Then mobjFile.WriteLine ""
                strTemp = "Insert Into zlmip_tab_parameter(tab_id,para_order,para_field,para_title,para_type,para_default,para_note)"
                mobjFile.WriteLine strTemp
                lngCount = 0
            End If
                        
            strTemp = "Select "
            strTemp = strTemp & "'" & rsTemp("tab_id").Value & "'"
            strTemp = strTemp & "," & rsTemp("para_order").Value
            strTemp = strTemp & ",'" & rsTemp("para_field").Value & "'"
            strTemp = strTemp & ",'" & rsTemp("para_title").Value & "'"
            strTemp = strTemp & "," & rsTemp("para_type").Value
            strTemp = strTemp & ",'" & rsTemp("para_default").Value & "'"
            strTemp = strTemp & ",'" & rsTemp("para_note").Value & "'"
            strTemp = strTemp & " From Dual"
            
            lngCount = lngCount + 1
            strTemp = IIf(lngCount = 1, "          ", "Union All ") & strTemp
            If lngCount = rsTemp.RecordCount Then strTemp = strTemp & ";"
                        
            mobjFile.WriteLine strTemp
            
            rsTemp.MoveNext
        Loop
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    'zlmip_tab_extend
    mobjFile.WriteLine ""
    mobjFile.WriteLine "--zlmip_tab_extend"
    
    strSQL = "Select a.* From zltools.zlmip_tab_extend a,zltools.zlmip_table b Where Nvl(b.data_code,'-')='" & strDataCode & "' And a.source_tab_id=b.id Order By a.source_tab_id"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open strSQL, gcnOracle
    If rsTemp.BOF = False Then
        intGroup = 0
        lngCount = 0
        Do While Not rsTemp.EOF
                                    
            intGroup = intGroup + 1
            If intGroup Mod 100 = 1 Then
                If intGroup > 1 Then mobjFile.WriteLine ""
                strTemp = "Insert Into zlmip_tab_extend(ID,SOURCE_TAB_ID,EXT_ORDER,EXT_TYPE,EXT_TITLE,TARGET_TAB_ID)"
                mobjFile.WriteLine strTemp
                lngCount = 0
            End If
                        
            strTemp = "Select "
            strTemp = strTemp & "'" & rsTemp("ID").Value & "'"
            strTemp = strTemp & ",'" & rsTemp("SOURCE_TAB_ID").Value & "'"
            strTemp = strTemp & "," & rsTemp("EXT_ORDER").Value
            strTemp = strTemp & "," & rsTemp("EXT_TYPE").Value
            strTemp = strTemp & ",'" & rsTemp("EXT_TITLE").Value & "'"
            strTemp = strTemp & ",'" & rsTemp("TARGET_TAB_ID").Value & "'"
            
            strTemp = strTemp & " From Dual"
            
            lngCount = lngCount + 1
            strTemp = IIf(lngCount = 1, "          ", "Union All ") & strTemp
            If lngCount = rsTemp.RecordCount Then strTemp = strTemp & ";"
                        
            mobjFile.WriteLine strTemp
            
            rsTemp.MoveNext
        Loop
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    'zlmip_tabext_condition
    mobjFile.WriteLine ""
    mobjFile.WriteLine "--zlmip_tabext_condition"
    
    strSQL = "Select b.* From zltools.zlmip_table x,zltools.zlmip_tab_extend a,zltools.zlmip_tabext_condition b Where Nvl(x.data_code,'-')='" & strDataCode & "' And x.id=a.source_tab_id And a.id=b.EXT_ID Order By a.source_tab_id,b.COND_ORDER"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open strSQL, gcnOracle
    If rsTemp.BOF = False Then
        intGroup = 0
        lngCount = 0
        Do While Not rsTemp.EOF
                                    
            intGroup = intGroup + 1
            If intGroup Mod 100 = 1 Then
                If intGroup > 1 Then mobjFile.WriteLine ""
                strTemp = "Insert Into zlmip_tabext_condition(EXT_ID,COND_ORDER,TARGET_FLD,SOURCE_FLD)"
                mobjFile.WriteLine strTemp
                
            End If
                        
            strTemp = "Select "
            strTemp = strTemp & "'" & rsTemp("EXT_ID").Value & "'"
            strTemp = strTemp & "," & rsTemp("COND_ORDER").Value
            strTemp = strTemp & ",'" & rsTemp("TARGET_FLD").Value & "'"
            strTemp = strTemp & ",'" & rsTemp("SOURCE_FLD").Value & "'"
            
            strTemp = strTemp & " From Dual"
            
            lngCount = lngCount + 1
            strTemp = IIf(lngCount = 1, "          ", "Union All ") & strTemp
            If lngCount = rsTemp.RecordCount Then strTemp = strTemp & ";"
                        
            mobjFile.WriteLine strTemp
            
            rsTemp.MoveNext
        Loop
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    'zlmip_item
    mobjFile.WriteLine ""
    mobjFile.WriteLine "--zlmip_item"
    
    strSQL = "Select * From zltools.zlmip_item a Where Nvl(a.data_code,'-')='" & strDataCode & "' And a.item_type=1 Order By a.item_code"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open strSQL, gcnOracle
    If rsTemp.BOF = False Then
        intGroup = 0
        lngCount = 0
        Do While Not rsTemp.EOF
                                    
            intGroup = intGroup + 1
            If intGroup Mod 100 = 1 Then
                If intGroup > 1 Then mobjFile.WriteLine ""
                strTemp = "Insert Into zlmip_item(id,data_code,item_code,item_title,item_request,item_type,trigger_type,again_policy,again_para,check_frequency,check_freq_internal,trigger_condition,trigger_frequency,tab_id,start_date,stop_date,item_note)"
                mobjFile.WriteLine strTemp
                lngCount = 0
            End If
            
            strTemp = "Select "
            strTemp = strTemp & "'" & rsTemp("id").Value & "'"
            strTemp = strTemp & ",'" & rsTemp("data_code").Value & "'"
            strTemp = strTemp & ",'" & rsTemp("item_code").Value & "'"
            strTemp = strTemp & ",'" & rsTemp("item_title").Value & "'"
            strTemp = strTemp & "," & rsTemp("item_request").Value
            strTemp = strTemp & "," & rsTemp("item_type").Value
            strTemp = strTemp & "," & rsTemp("trigger_type").Value
            strTemp = strTemp & "," & rsTemp("again_policy").Value
            strTemp = strTemp & ",Null"
            
            strTemp = strTemp & "," & NVL(rsTemp("check_frequency").Value, "Null")
            strTemp = strTemp & "," & NVL(rsTemp("check_freq_internal").Value, "Null")
            strTemp = strTemp & ",'" & Replace(NVL(rsTemp("trigger_condition").Value), "'", "''") & "'"
            strTemp = strTemp & "," & NVL(rsTemp("trigger_frequency").Value, "Null")
            
            If IsNull(rsTemp("tab_id").Value) Then
                strTemp = strTemp & ",Null"
            Else
                strTemp = strTemp & ",'" & rsTemp("tab_id").Value & "'"
            End If
            
            strTemp = strTemp & ",Sysdate"
            strTemp = strTemp & ",Null"
            strTemp = strTemp & ",'" & rsTemp("item_note").Value & "'"
            
            strTemp = strTemp & " From Dual"
            
            lngCount = lngCount + 1
            strTemp = IIf(lngCount = 1, "          ", "Union All ") & strTemp
                        
            If lngCount = rsTemp.RecordCount Then
                strTemp = strTemp & ";"
            End If
            
            mobjFile.WriteLine strTemp
            
            rsTemp.MoveNext
        Loop
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    'zlmip_tab_field
    mobjFile.WriteLine ""
    mobjFile.WriteLine "--zlmip_item_field"
    
    strSQL = "Select b.* From zltools.zlmip_item a,zltools.zlmip_item_field b Where Nvl(a.data_code,'-')='" & strDataCode & "' And a.item_type=1  And a.id=b.item_id Order By a.item_code,b.fld_order"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open strSQL, gcnOracle
    If rsTemp.BOF = False Then
        intGroup = 0
        lngCount = 0
        Do While Not rsTemp.EOF
                                    
            intGroup = intGroup + 1
            If intGroup Mod 100 = 1 Then
                If intGroup > 1 Then mobjFile.WriteLine ""
                strTemp = "Insert Into zlmip_item_field(item_id,fld_order,fld_title,fld_type)"
                mobjFile.WriteLine strTemp
                lngCount = 0
            End If
                        
            strTemp = "Select "
            strTemp = strTemp & "'" & rsTemp("item_id").Value & "'"
            strTemp = strTemp & "," & rsTemp("fld_order").Value
            strTemp = strTemp & ",'" & rsTemp("fld_title").Value & "'"
            strTemp = strTemp & "," & rsTemp("fld_type").Value
            strTemp = strTemp & " From Dual"
            
            lngCount = lngCount + 1
            strTemp = IIf(lngCount = 1, "          ", "Union All ") & strTemp
            If rsTemp.AbsolutePosition = rsTemp.RecordCount Or (intGroup Mod 100 = 0 And intGroup > 1) Then strTemp = strTemp & ";"
                        
            mobjFile.WriteLine strTemp
            
            rsTemp.MoveNext
        Loop
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    'zlmip_item_deliver
    mobjFile.WriteLine ""
    mobjFile.WriteLine "--zlmip_item_deliver"
    
    strSQL = "Select a.id,a.item_id,a.deliver_order,a.deliver_code,a.deliver_title,a.deliver_object.getStringVal() As deliver_object From zltools.zlmip_item_deliver a,zltools.zlmip_item b Where Nvl(b.data_code,'-')='" & strDataCode & "' And b.item_type=1 And a.item_id=b.id Order By b.item_code,a.deliver_order"
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open strSQL, gcnOracle
    If rsTemp.BOF = False Then
        intGroup = 0
        lngCount = 0
        Do While Not rsTemp.EOF
                                    
            intGroup = intGroup + 1
            If intGroup Mod 100 = 1 Then
                If intGroup > 1 Then mobjFile.WriteLine ""
                strTemp = "Insert Into zlmip_item_deliver(id,item_id,deliver_order,deliver_code,deliver_title,deliver_object)"
                mobjFile.WriteLine strTemp
                lngCount = 0
            End If
            
            lngCount = lngCount + 1
            If lngCount = 1 Then
                strTemp = "          Select '" & rsTemp("id").Value & "','" & rsTemp("item_id").Value & "'," & rsTemp("deliver_order").Value & ",'" & rsTemp("deliver_code").Value & "','" & rsTemp("deliver_title").Value & "',xmltype('" & rsTemp("deliver_object").Value & "') From Dual"
            Else
                strTemp = "Union All Select '" & rsTemp("id").Value & "','" & rsTemp("item_id").Value & "'," & rsTemp("deliver_order").Value & ",'" & rsTemp("deliver_code").Value & "','" & rsTemp("deliver_title").Value & "',xmltype('" & rsTemp("deliver_object").Value & "') From Dual"
            End If
            
            If lngCount = rsTemp.RecordCount Then
                strTemp = strTemp & ";"
            End If
            
            mobjFile.WriteLine strTemp
            
            rsTemp.MoveNext
        Loop
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    mobjFile.WriteLine ""
    mobjFile.WriteLine "--zlmip_item_config"
    
    strSQL = "Select a.* From zltools.zlmip_item_config a,zltools.zlmip_item b Where Nvl(b.data_code,'-')='" & strDataCode & "' And b.item_type=1 And a.item_id=b.id Order By b.item_code,a.node_order"
        
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open strSQL, gcnOracle
    If rsTemp.BOF = False Then
        intGroup = 0
        lngCount = 0
        Do While Not rsTemp.EOF
                                    
            intGroup = intGroup + 1
            If intGroup Mod 100 = 1 Then
                If intGroup > 1 Then mobjFile.WriteLine ""
                strTemp = "Insert Into zlmip_item_config(id,parent_id,item_id,node_order,node_type,node_title,data_type,min_occurs,max_occurs,config_occurs,config_occurs_key,config_express,config_express_key,config_note)"
                mobjFile.WriteLine strTemp
                lngCount = 0
            End If
                        
            strTemp = "Select "
            strTemp = strTemp & "'" & rsTemp("id").Value & "'"
            strTemp = strTemp & ",'" & rsTemp("parent_id").Value & "'"
            strTemp = strTemp & ",'" & rsTemp("item_id").Value & "'"
            strTemp = strTemp & "," & rsTemp("node_order").Value
            strTemp = strTemp & "," & rsTemp("node_type").Value
            strTemp = strTemp & ",'" & rsTemp("node_title").Value & "'"
            strTemp = strTemp & ",'" & rsTemp("data_type").Value & "'"
            strTemp = strTemp & ",'" & rsTemp("min_occurs").Value & "'"
            strTemp = strTemp & ",'" & rsTemp("max_occurs").Value & "'"
            strTemp = strTemp & ",'" & rsTemp("config_occurs").Value & "'"
            strTemp = strTemp & ",'" & rsTemp("config_occurs_key").Value & "'"
            strTemp = strTemp & ",'" & rsTemp("config_express").Value & "'"
            strTemp = strTemp & ",'" & rsTemp("config_express_key").Value & "'"
            strTemp = strTemp & ",'" & rsTemp("config_note").Value & "'"
            strTemp = strTemp & " From Dual"
            
            lngCount = lngCount + 1
            strTemp = IIf(lngCount = 1, "          ", "Union All ") & strTemp
            If lngCount = rsTemp.RecordCount Or (intGroup Mod 100 = 0 And intGroup > 1) Then strTemp = strTemp & ";"
                        
            mobjFile.WriteLine strTemp
            
            rsTemp.MoveNext
        Loop
    End If
    mobjFile.Close
    
    'zlmip_send_log
    '------------------------------------------------------------------------------------------------------------------
    
    'zlmip_receive_log
    '------------------------------------------------------------------------------------------------------------------
    
    MakeBusinessDataScript = True
    
End Function



