VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~1.OCX"
Object = "*\A..\zlRichEditor\zlRichEdit.vbp"
Begin VB.Form frmEPRFileContent 
   BackColor       =   &H80000003&
   BorderStyle     =   0  'None
   Caption         =   "病历文件提纲"
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picTab 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   2985
      Left            =   2175
      ScaleHeight     =   2985
      ScaleWidth      =   3420
      TabIndex        =   0
      Top             =   105
      Visible         =   0   'False
      Width           =   3420
      Begin VSFlex8Ctl.VSFlexGrid vfgThis 
         Height          =   4170
         Left            =   0
         TabIndex        =   1
         Top             =   810
         Width           =   7500
         _cx             =   13229
         _cy             =   7355
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
         BackColorFixed  =   15790320
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
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
         Rows            =   3
         Cols            =   6
         FixedRows       =   2
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
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
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "一般护理记录单"
         Height          =   180
         Left            =   2970
         TabIndex        =   3
         Top             =   0
         Width           =   1275
      End
      Begin VB.Label lblSubhead 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名:##"
         Height          =   180
         Left            =   180
         TabIndex        =   2
         Top             =   540
         Width           =   630
      End
   End
   Begin zlRichEditor.Editor edtThis 
      Height          =   2580
      Left            =   315
      TabIndex        =   4
      Top             =   225
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4551
      WithViewButtonas=   0   'False
      ShowRuler       =   0   'False
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   45
      Top             =   45
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmEPRFileContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum zlEnumCompendParentKind     '提纲父类型
    cprEmCPKFileDefine = 0              '文件定义内容
    cprEmCPKModelEssay = 1              '范文内容
End Enum

'-----------------------------------------------------
'窗体事件
'-----------------------------------------------------
Public Event DblClick()                                                 '返回双击操作事件

'-----------------------------------------------------
'临时变量
Dim rsTemp As New ADODB.Recordset
Dim lngCount As Long

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case ID_EDIT_COPY
        Control.Enabled = edtThis.Selection.EndPos <> edtThis.Selection.StartPos
    End Select
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case ID_EDIT_COPY
        Me.edtThis.Copy
    End Select
End Sub

Private Sub Form_Load()
    cbsThis.ActiveMenuBar.Visible = False
    cbsThis.KeyBindings.Add FCONTROL, Asc("C"), ID_EDIT_COPY
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With Me.edtThis
        .Left = Me.ScaleLeft + 90: .Width = Me.ScaleWidth - 2 * .Left
        .Top = Me.ScaleTop + 90: .Height = Me.ScaleHeight - 2 * .Top
    End With
    With Me.picTab
        .Left = Me.ScaleLeft + 90: .Width = Me.ScaleWidth - 2 * .Left
        .Top = Me.ScaleTop + 90: .Height = Me.ScaleHeight - 2 * .Top
    End With
End Sub

Private Sub picTab_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Err = 0: On Error Resume Next
    Me.lblTitle.Move Me.picTab.ScaleLeft, Me.picTab.ScaleTop + 120, Me.picTab.ScaleWidth
    Me.lblSubhead.Move Me.picTab.ScaleLeft + 210, Me.lblTitle.Top + Me.lblTitle.Height + 120
    Me.vfgThis.Move Me.picTab.ScaleLeft + 210, Me.lblSubhead.Top + Me.lblSubhead.Height + 45, Me.picTab.ScaleWidth - 210 * 2
    Me.vfgThis.Height = Me.picTab.ScaleHeight - Me.vfgThis.Top - 210
End Sub

Private Sub edtThis_DblClick(ViewMode As zlRichEditor.ViewModeEnum)
    RaiseEvent DblClick
End Sub

Private Sub edtThis_RequestRightMenu(ViewMode As zlRichEditor.ViewModeEnum, Shift As Integer, x As Single, y As Single)
    Dim Popup As CommandBar
    Dim Control As CommandBarControl
    
    Set Popup = cbsThis.Add("Popup", xtpBarPopup)
    With Popup.Controls
        Set Control = .Add(xtpControlButton, ID_EDIT_COPY, "复制(&C)")
        Popup.ShowPopup
    End With
End Sub

'-----------------------------------------------------
'窗体公共方法
'-----------------------------------------------------

Public Sub zlRefresh(ByVal lngParentId As Long, Optional bytParentKind As zlEnumCompendParentKind = cprEmCPKFileDefine)
    '功能：显示指定文件/范文的内容；
    Dim strTemp As String, strZipFile As String
    Dim rsTemp As New ADODB.Recordset
    
    Me.edtThis.Visible = True
    Me.picTab.Visible = False
    
    If lngParentId = 0 Then Me.edtThis.ReadOnly = False: Me.edtThis.NewDoc: Me.edtThis.ReadOnly = True: Exit Sub
    Me.edtThis.ReadOnly = False
    Me.edtThis.NewDoc
    Me.edtThis.Freeze
    If lngParentId > 0 Then
        '设置页面格式
        Dim mEPRFileInfo As New cEPRFileDefineInfo
        If bytParentKind = cprEmCPKFileDefine Then
            '病历文件格式
            gstrSQL = "Select b.ID, a.格式 From 病历页面格式 a, 病历文件列表 b " & _
                " Where  a.种类 = b.种类 And a.编号 = b.页面 And b.ID = [1]"
        Else
            '病历范文格式
            gstrSQL = "Select C.ID, a.格式 From 病历页面格式 a, 病历文件列表 b, 病历范文目录 c " & _
                " Where  c.文件ID = b.ID And b.种类 = a.种类 And b.页面 = a.编号 And c.ID = [1]"
        End If
        Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
        If Not rsTemp.EOF Then
            mEPRFileInfo.格式 = "" & rsTemp!格式
            mEPRFileInfo.SetFormat Me.edtThis, mEPRFileInfo.格式
        End If
        Set mEPRFileInfo = Nothing
    End If
    Me.edtThis.UnFreeze
    Me.edtThis.ResetWYSIWYG
    If bytParentKind = cprEmCPKFileDefine Then
        '病历文件格式
        Err = 0: On Error GoTo errHand
        gstrSQL = "Select l.种类,l.保留 From 病历文件列表 l Where l.Id = [1]"
        Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
        If Val("" & rsTemp!保留) < 0 And rsTemp!种类 <> 6 Then
            With Me.edtThis
                .Freeze
                .Text = vbCrLf & Space(4) & "该文件为特殊格式病历，不能浏览样式..."
                .SelectAll
                .ForceEdit = True
                .Selection.Font.Name = "宋体": .Selection.Font.Size = 10.5
                .SelLength = 0
                .ForceEdit = False
                .UnFreeze
            End With
        ElseIf rsTemp!种类 <> 3 Then
            strZipFile = zlBlobRead(1, lngParentId)
            If Len(strZipFile) > 0 Then
                If gobjFSO.FileExists(strZipFile) Then
                    strTemp = zlFileUnzip(strZipFile)
                    If gobjFSO.FileExists(strTemp) Then
                        Me.edtThis.OpenDoc strTemp
                        gobjFSO.DeleteFile strTemp, True
                    End If
                    gobjFSO.DeleteFile strZipFile, True
                End If
            End If
        Else
            Me.edtThis.Visible = False
            Me.picTab.Visible = True
            
            Dim lngCurColor As Long, strCurFont As String, objFont As StdFont
            Me.lblTitle.Caption = "": Me.lblSubhead.Caption = ""
            Me.vfgThis.Redraw = flexRDNone
            Me.vfgThis.Clear: Me.vfgThis.MergeCells = flexMergeFree
            Me.vfgThis.MergeRow(0) = True: Me.vfgThis.MergeRow(1) = True
            
            gstrSQL = "Select d.对象序号, d.内容文本, d.要素名称" & _
                " From 病历文件结构 d, 病历文件结构 p" & _
                " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表格样式'" & _
                " Order By d.对象序号"
            Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
            With rsTemp
                Do While Not .EOF
                    Select Case "" & !要素名称
                    Case "表头层数"
                        If Val("" & !内容文本) = 1 Then
                            Me.vfgThis.RowHidden(0) = True
                        Else
                            Me.vfgThis.RowHidden(0) = False
                        End If
                    Case "总列数":  Me.vfgThis.Cols = Val("" & !内容文本)
                    Case "最小行高": Me.vfgThis.RowHeightMin = Val("" & !内容文本)
                    Case "文本字体"
                        strCurFont = "" & !内容文本
                        Set objFont = New StdFont
                        With objFont
                            .Name = Split(strCurFont, ",")(0)
                            .Size = Val(Split(strCurFont, ",")(1))
                            .Bold = False: .Italic = False
                            If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                            If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                        End With
                        Set Me.vfgThis.Font = objFont
                        Set Me.lblSubhead.Font = Me.vfgThis.Font
                        
                    Case "文本颜色": Me.vfgThis.ForeColor = Val("" & !内容文本)
                    Case "表格颜色": Me.vfgThis.GridColor = Val("" & !内容文本): Me.vfgThis.GridColorFixed = Me.vfgThis.GridColor
                    
                    Case "标题文本": Me.lblTitle.Caption = "" & !内容文本
                    Case "标题字体"
                        strCurFont = "" & !内容文本
                        Set objFont = New StdFont
                        With objFont
                            .Name = Split(strCurFont, ",")(0)
                            .Size = Val(Split(strCurFont, ",")(1))
                            .Bold = False: .Italic = False
                            If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                            If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                        End With
                        Set Me.lblTitle.Font = objFont
                        Me.lblTitle.AutoSize = False
                    End Select
                    .MoveNext
                Loop
            End With
            '---------------------------------------------------
            gstrSQL = "Select d.对象序号, d.内容文本, d.要素名称, Nvl(d.是否换行, 0) As 是否换行" & _
                " From 病历文件结构 d, 病历文件结构 p" & _
                " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表上标签'" & _
                " Order By d.对象序号"
            Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
            With rsTemp
                Me.lblSubhead.Caption = ""
                Do While Not .EOF
                    Me.lblSubhead.Caption = Me.lblSubhead.Caption & " " & IIf(!是否换行 = 0, "", vbCrLf) & !内容文本 & "{" & !要素名称 & "}"
                    .MoveNext
                Loop
                Me.lblSubhead.Caption = Trim(Me.lblSubhead.Caption)
            End With
            '---------------------------------------------------
            gstrSQL = "Select d.对象序号, d.内容行次, d.内容文本" & _
                " From 病历文件结构 d, 病历文件结构 p" & _
                " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表头单元'" & _
                " Order By d.对象序号"
            Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
            With rsTemp
                Do While Not .EOF
                    Me.vfgThis.TextMatrix(!内容行次 - 1, !对象序号 - 1) = "" & !内容文本
                    Me.vfgThis.FixedAlignment(!对象序号 - 1) = flexAlignCenterCenter
                    .MoveNext
                Loop
            End With
            '---------------------------------------------------
            gstrSQL = "Select d.对象序号, d.对象属性, d.内容行次, d.内容文本, d.要素名称, d.要素单位" & _
                " From 病历文件结构 d, 病历文件结构 p" & _
                " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表列集合'" & _
                " Order By d.对象序号, d.内容行次"
            Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
            With rsTemp
                Do While Not .EOF
                    Me.vfgThis.ColWidth(!对象序号 - 1) = Val("" & !对象属性)
                    .MoveNext
                Loop
            End With
            Me.vfgThis.Redraw = flexRDDirect
                    
            '---------------------------------------------------
            Call picTab_Resize
        End If
    Else
        '病历范文内容
        strZipFile = zlBlobRead(3, lngParentId)
        If Len(strZipFile) > 0 Then
            If gobjFSO.FileExists(strZipFile) Then
                strTemp = zlFileUnzip(strZipFile)
                If gobjFSO.FileExists(strTemp) Then
                    Me.edtThis.OpenDoc strTemp
                    gobjFSO.DeleteFile strTemp, True
                End If
                gobjFSO.DeleteFile strZipFile, True
            End If
        End If
    End If
    edtThis.RefreshTargetDC
    Me.edtThis.ReadOnly = True
    Exit Sub
    
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
