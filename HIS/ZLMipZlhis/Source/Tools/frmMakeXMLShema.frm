VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Begin VB.Form frmMakeXMLShema 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picBack 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2820
      Index           =   0
      Left            =   375
      ScaleHeight     =   2820
      ScaleWidth      =   9270
      TabIndex        =   0
      Top             =   1800
      Width           =   9270
      Begin VB.PictureBox picBack 
         BorderStyle     =   0  'None
         Height          =   2115
         Index           =   1
         Left            =   15
         ScaleHeight     =   2115
         ScaleWidth      =   8040
         TabIndex        =   1
         Top             =   15
         Width           =   8040
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   1
            Left            =   900
            TabIndex        =   6
            Text            =   "E:\ZLSOFT\Source\ZLSOFT10\zlMspZLHIS\Script\MessageStruct"
            Top             =   1725
            Width           =   6645
         End
         Begin VB.CommandButton cmdOpen 
            Height          =   300
            Index           =   0
            Left            =   7575
            Picture         =   "frmMakeXMLShema.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   1725
            Width           =   315
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            Height          =   1050
            Index           =   3
            Left            =   825
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Text            =   "frmMakeXMLShema.frx":6852
            Top             =   375
            Width           =   6645
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   0
            Left            =   825
            TabIndex        =   3
            Text            =   "E:\ZLSOFT\Source\ZLSOFT10\ZLMSPZLHIS\Docment\ZLHIS产品消息结构.xlsx"
            Top             =   45
            Width           =   6645
         End
         Begin VB.CommandButton cmdOpen 
            Height          =   300
            Index           =   1
            Left            =   7500
            Picture         =   "frmMakeXMLShema.frx":68F3
            Style           =   1  'Graphical
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   30
            Width           =   315
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "输出路径"
            Height          =   180
            Index           =   1
            Left            =   45
            TabIndex        =   9
            Top             =   1755
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "输出内容"
            Height          =   180
            Index           =   4
            Left            =   45
            TabIndex        =   8
            Top             =   390
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "消息文件"
            Height          =   180
            Index           =   0
            Left            =   45
            TabIndex        =   7
            Top             =   105
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
   Begin XtremeCommandBars.ImageManager ImageManager1 
      Bindings        =   "frmMakeXMLShema.frx":D145
      Left            =   1785
      Top             =   300
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMakeXMLShema.frx":D159
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   750
      Top             =   180
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMakeXMLShema"
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
    
    Set objControl = NewToolBar(objBar, xtpControlButton, 3, "开始生成", False, , xtpButtonIconAndCaption)
    objControl.IconId = 3
    
End Function

Private Sub MakeMessageStructScript()
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim strSource As String
    Dim strTargetFolder As String
    Dim intLoop As Integer
    Dim varTemp As Variant

    strSource = txt(0).Text
    strTargetFolder = txt(1).Text
    
    varTemp = Split(txt(3).Text, "|")
    
    For intLoop = 0 To UBound(varTemp)
        varTemp(intLoop) = Replace(varTemp(intLoop), vbCrLf, "")
        If MakeXSD(strSource, UCase(varTemp(intLoop)), strTargetFolder) = False Then
            MsgBox "从指定的Excel文件中生成消息结构(*.xsd)失败！", vbCritical
            Exit Sub
        End If
    Next
        
    MsgBox "从指定的Excel文件中生成消息结构(*.xsd)完成！", vbInformation
    
End Sub

Private Function MakeXSD(ByVal strExcelFile As String, ByVal strKind As String, ByVal strTargetFolder As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim intCount As Integer
    Dim objExcel As Object
    Dim objWorkbook As Object
    Dim objWorkSheet As Object
    Dim strName As String
    Dim strType As String
    Dim intIndentLevel As Integer
    Dim intPreIndentLevel As Integer
    Dim strMinOccurs As String
    Dim strMaxOccurs As String
    Dim strXSDPath As String
    Dim strXSDFile As String
    
    Dim objFile As TextStream
    Dim objScript As TextStream
    Dim objConfig As TextStream
    Dim strTemp As String
    
    On Error GoTo errHand
    
    
    Set objWorkbook = Nothing
    Set objWorkSheet = Nothing
    
    Set objExcel = CreateObject("Excel.Application")
    Set objWorkbook = CreateObject("Excel.Workbook")
    Set objWorkSheet = CreateObject("Excel.Worksheet")
    
    Set objWorkbook = objExcel.Workbooks.Open(strExcelFile)
    If objWorkbook Is Nothing Then Exit Function
    
    Set objWorkSheet = objWorkbook.Worksheets(strKind)
    If objWorkSheet Is Nothing Then Exit Function
    
    strXSDPath = strTargetFolder & "\" & strKind
    If mobjFso.FolderExists(strXSDPath) = True Then Call mobjFso.DeleteFolder(strXSDPath, True)
        
    Call mobjFso.CreateFolder(strXSDPath)
                        
    Set objConfig = mobjFso.CreateTextFile(strXSDPath & "\" & strKind & ".config", True)
    Call objConfig.WriteLine("<?xml version=""1.0"" encoding=""gbk""?>")
    Call objConfig.WriteLine("<config>")
    Call objConfig.WriteLine(GetIndent(1) & "<message>")
        
    For lngLoop = 2 To objWorkSheet.UsedRange.Rows.Count
        
        strName = Trim(objWorkSheet.Range("B" & lngLoop).Text)
        If strName <> "" Then
            If Left(strName, Len(strKind) + 1) = strKind & "_" Then
                
                
                If strXSDFile <> "" Then
                
                    For intCount = intPreIndentLevel To 1 Step -1
                        Call objFile.WriteLine(GetIndent(intCount * 3) & "</xs:sequence>")
                        Call objFile.WriteLine(GetIndent(intCount * 3 - 1) & "</xs:complexType>")
                        Call objFile.WriteLine(GetIndent(intCount * 3 - 2) & "</xs:element>")
                    Next

                    Call objFile.WriteLine("</xs:schema>")
                    objFile.Close
                    
                    strXSDFile = ""
                    intPreIndentLevel = 0
                End If
                                
                strXSDFile = strXSDPath & "\" & strName & ".xsd"
                
                Call objConfig.WriteLine(GetIndent(2) & "<code>" & strName & "</code><title>" & Trim(objWorkSheet.Range("C" & lngLoop).Text) & "</title><desc>" & Trim(objWorkSheet.Range("E" & lngLoop).Text) & "</desc>")
                                
                Set objFile = mobjFso.CreateTextFile(strXSDFile, True)
                Call objFile.WriteLine("<?xml version=""1.0"" encoding=""gbk""?>")
                Call objFile.WriteLine("<xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema""  elementFormDefault=""qualified"" attributeFormDefault=""unqualified"">")
                Call objFile.WriteLine(GetIndent(1) & "<xs:element name=""message"">")
                Call objFile.WriteLine(GetIndent(2) & "<xs:complexType>")
                Call objFile.WriteLine(GetIndent(3) & "<xs:sequence>")
                
                Call objFile.WriteLine(GetIndent(4) & "<xs:element name=""head""  minOccurs=""1"" maxOccurs=""1"">")
                Call objFile.WriteLine(GetIndent(5) & "<xs:complexType>")
                Call objFile.WriteLine(GetIndent(6) & "<xs:sequence>")
                Call objFile.WriteLine(GetIndent(7) & "<xs:element name=""" & LCase(strName) & """ type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>")
                                                
                Call objFile.WriteLine(GetIndent(7) & "<xs:element name=""msg_id"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>")
                Call objFile.WriteLine(GetIndent(7) & "<xs:element name=""msg_item_code"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>")
                Call objFile.WriteLine(GetIndent(7) & "<xs:element name=""send_station_name"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>")
                Call objFile.WriteLine(GetIndent(7) & "<xs:element name=""send_station_ip"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>")
                Call objFile.WriteLine(GetIndent(7) & "<xs:element name=""send_program"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>")
                Call objFile.WriteLine(GetIndent(7) & "<xs:element name=""send_instance"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>")
                Call objFile.WriteLine(GetIndent(7) & "<xs:element name=""send_system_code"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>")
                Call objFile.WriteLine(GetIndent(7) & "<xs:element name=""send_module_code"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>")
                Call objFile.WriteLine(GetIndent(7) & "<xs:element name=""send_mipuser"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>")
                Call objFile.WriteLine(GetIndent(7) & "<xs:element name=""send_hisuser"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>")
                Call objFile.WriteLine(GetIndent(7) & "<xs:element name=""send_time"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>")
                                
                Call objFile.WriteLine(GetIndent(6) & "</xs:sequence>")
                Call objFile.WriteLine(GetIndent(5) & "</xs:complexType>")
                Call objFile.WriteLine(GetIndent(4) & "</xs:element>")
                
            ElseIf strName <> "节点名称" Then
                '单元格的缩进
                intIndentLevel = objWorkSheet.Range("B" & lngLoop).IndentLevel + 1
                            
                Select Case Trim(objWorkSheet.Range("F" & lngLoop).Text)
                Case ""
                    strType = ""
                Case "N"
                    If Val(Trim(objWorkSheet.Range("G" & lngLoop).Text)) > 0 Then
                        strType = "xs:decimal"
                    Else
                        strType = "xs:integer"
                    End If
                Case "S"
                    strType = "xs:string"
                Case "D"
                    strType = "xs:date"
                Case "T"
                    strType = "xs:time"
                Case "DT"
                    strType = "xs:datetime"
                End Select
                
                Select Case Trim(objWorkSheet.Range("E" & lngLoop).Text)
                Case "0..1"
                    strMinOccurs = "0"
                    strMaxOccurs = "1"
                Case "0..*"
                    strMinOccurs = "0"
                    strMaxOccurs = "unbounded"
                Case "1"
                    strMinOccurs = "1"
                    strMaxOccurs = "1"
                Case "1..*"
                    strMinOccurs = "1"
                    strMaxOccurs = "unbounded"
                End Select
                
                If intPreIndentLevel > intIndentLevel Then
                    For intCount = intPreIndentLevel To intIndentLevel + 1 Step -1
                        Call objFile.WriteLine(GetIndent(intCount * 3) & "</xs:sequence>")
                        Call objFile.WriteLine(GetIndent(intCount * 3 - 1) & "</xs:complexType>")
                        Call objFile.WriteLine(GetIndent(intCount * 3 - 2) & "</xs:element>")
                    Next
                End If
                
                If strType = "" Then
                    Call objFile.WriteLine(GetIndent(intIndentLevel * 3 + 1) & "<xs:element name=""" & strName & """ minOccurs=""" & strMinOccurs & """ maxOccurs=""" & strMaxOccurs & """>")
                    Call objFile.WriteLine(GetIndent(intIndentLevel * 3 + 2) & "<xs:complexType>")
                    Call objFile.WriteLine(GetIndent(intIndentLevel * 3 + 3) & "<xs:sequence>")
                Else
                    Call objFile.WriteLine(GetIndent(intIndentLevel * 3 + 1) & "<xs:element name=""" & strName & """ type=""" & strType & """ minOccurs=""" & strMinOccurs & """ maxOccurs=""" & strMaxOccurs & """/>")
                End If
                intPreIndentLevel = intIndentLevel
            End If
        End If
    Next
    If strXSDFile <> "" Then
    
        If intPreIndentLevel > 0 Then
            
            For intCount = intPreIndentLevel To 1 Step -1
                Call objFile.WriteLine(GetIndent(intCount * 3) & "</xs:sequence>")
                Call objFile.WriteLine(GetIndent(intCount * 3 - 1) & "</xs:complexType>")
                Call objFile.WriteLine(GetIndent(intCount * 3 - 2) & "</xs:element>")
            Next

        End If
            
        Call objFile.WriteLine("</xs:schema>")
        objFile.Close
        
        strXSDFile = ""
    End If
    
    Call objConfig.WriteLine(GetIndent(1) & "</message>")
    Call objConfig.WriteLine("</config>")
    objConfig.Close
    objExcel.Quit
    
    MakeXSD = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox Err.Description
'    Resume

    If Not (objConfig Is Nothing) Then objConfig.Close
    If Not (objFile Is Nothing) Then objFile.Close
    objExcel.Quit

End Function

Private Function GetIndent(ByVal intIndentLevel As Integer) As String
    GetIndent = Space(intIndentLevel * 4)
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.id
    Case 3
        Call MakeMessageStructScript
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
    
    txt(3).Text = GetSetting("ZLSOFT", "生成消息结构", "输出内容", "ZLHIS_PUB|ZLHIS_QUEUE|ZLHIS_PATIENT|ZLHIS_CIS|ZLHIS_EMR|ZLHIS_LIS|ZLHIS_PACS|ZLHIS_PEIS|ZLHIS_OPER|ZLHIS_CHARGE|ZLHIS_REGIST|ZLHIS_TRANSFUSION|ZLHIS_BLOOD")
    txt(1).Text = GetSetting("ZLSOFT", "生成消息结构", "输出路径", "E:\ZLSOFT\Source\ZLSOFT10\zlMspZLHIS\Script\MessageStruct")
    txt(0).Text = GetSetting("ZLSOFT", "生成消息结构", "消息文件", "E:\ZLSOFT\Source\DataStructure\ZLHIS产品消息结构.xlsx")
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "ZLSOFT", "生成消息结构", "输出内容", Trim(txt(3).Text)
    SaveSetting "ZLSOFT", "生成消息结构", "输出路径", Trim(txt(1).Text)
    SaveSetting "ZLSOFT", "生成消息结构", "消息文件", Trim(txt(0).Text)
End Sub

Private Sub picBack_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        picBack(1).Move 15, 15, picBack(Index).Width - 30, picBack(Index).Height - 30
    Case 1
        txt(0).Move txt(0).Left, txt(0).Top, picBack(Index).Width - txt(0).Left - cmdOpen(1).Width - 60
        lbl(0).Top = txt(0).Top
        cmdOpen(1).Move txt(0).Left + txt(0).Width + 30
        txt(3).Move txt(0).Left, txt(0).Top + txt(0).Height + 45, txt(0).Width, picBack(Index).Height - (txt(0).Top + txt(0).Height + 45) - txt(1).Height - 90
        lbl(4).Top = txt(3).Top
        txt(1).Move txt(0).Left, txt(3).Top + txt(3).Height + 45, txt(3).Width
        lbl(1).Top = txt(1).Top
        cmdOpen(0).Move txt(1).Left + txt(1).Width + 30, txt(1).Top
    End Select
    
End Sub



