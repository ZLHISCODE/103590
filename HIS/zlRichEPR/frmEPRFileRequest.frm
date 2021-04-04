VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmEPRFileRequest 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "病历应用要求"
   ClientHeight    =   3945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vgdRequest 
      Height          =   2445
      Left            =   315
      TabIndex        =   4
      Top             =   2070
      Visible         =   0   'False
      Width           =   5355
      _cx             =   9446
      _cy             =   4313
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
      BackColorSel    =   16764057
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmEPRFileRequest.frx":0000
      ScrollTrack     =   -1  'True
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
   Begin VB.Label lbl说明内容 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "说明：文件用于在××时候书写。"
      Height          =   180
      Left            =   45
      TabIndex        =   5
      Top             =   75
      Width           =   5475
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl要求内容 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "在病人入科后，一般每120小时书写一次，病重则每48小时书写一次，病危则每24小时书写一次。"
      Height          =   360
      Left            =   255
      TabIndex        =   3
      Top             =   1350
      Width           =   5475
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl要求标题 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2)时限要求:"
      Height          =   180
      Left            =   90
      TabIndex        =   2
      Top             =   1095
      Width           =   990
   End
   Begin VB.Label lbl通用内容 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "该病历适合所有科室。"
      Height          =   180
      Left            =   255
      TabIndex        =   1
      Top             =   720
      Width           =   5475
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl通用标题 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1)适用科室:"
      Height          =   180
      Left            =   90
      TabIndex        =   0
      Top             =   465
      Width           =   990
   End
End
Attribute VB_Name = "frmEPRFileRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum zlEnumDClick
    cprEmDClickApplyTo = 1         '双击适用科室
    cprEmDClickRequest = 2         '双击时限要求
End Enum

'-----------------------------------------------------
'窗体公共事件
'-----------------------------------------------------
Public Event DblClick(lngWhere As zlEnumDClick)    '返回双击事件

'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mintKind As Integer       '病历种类
Private mlngFileID As Long        '病历文件ID

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lbl通用内容.Font.Underline = False
    Me.lbl通用内容.ForeColor = Me.lbl通用标题.ForeColor
    Me.lbl要求内容.Font.Underline = False
    Me.lbl要求内容.ForeColor = Me.lbl要求标题.ForeColor
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    
    With Me.lbl说明内容
        .Left = 90: .Width = Me.ScaleWidth - .Left * 2: .Top = 195
    End With
    Me.lbl通用标题.Left = 90: Me.lbl通用标题.Top = Me.lbl说明内容.Top + Me.lbl说明内容.Height + 195
    With Me.lbl通用内容
        .Left = 255: .Width = Me.ScaleWidth - Me.lbl通用内容.Left - Me.lbl通用标题.Left
        .Top = Me.lbl通用标题.Top + Me.lbl通用标题.Height + 75
    End With
    With Me.lbl要求标题
        .Left = Me.lbl通用标题.Left
        .Top = Me.lbl通用内容.Top + Me.lbl通用内容.Height + 195
    End With
    With Me.lbl要求内容
        .Left = Me.lbl通用内容.Left: .Width = Me.ScaleWidth - Me.lbl要求内容.Left - Me.lbl通用标题.Left
        .Top = Me.lbl要求标题.Top + Me.lbl要求标题.Height + 75
    End With
    
    With Me.vgdRequest
        .Left = Me.lbl要求内容.Left: .Width = Me.lbl要求内容.Width
        .Top = Me.lbl要求内容.Top: .Height = Me.ScaleHeight - .Top - 195
        
        If mintKind = 5 Then
            .ColWidth(2) = .Width - .ColWidth(0) - .ColWidth(1) - .ColWidth(3) - 300
        Else
            .ColWidth(1) = .Width - .ColWidth(0) - .ColWidth(2) - 30
        End If
    End With
End Sub

Private Sub lbl通用内容_DblClick()
    RaiseEvent DblClick(cprEmDClickApplyTo)
End Sub

Private Sub lbl通用内容_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lbl通用内容.Font.Underline = True
    Me.lbl通用内容.ForeColor = RGB(0, 0, 128)
End Sub

Private Sub lbl要求内容_DblClick()
    RaiseEvent DblClick(cprEmDClickRequest)
End Sub

Private Sub lbl要求内容_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lbl要求内容.Font.Underline = True
    Me.lbl要求内容.ForeColor = RGB(0, 0, 128)
End Sub

Private Sub vgdRequest_DblClick()
    RaiseEvent DblClick(cprEmDClickRequest)
End Sub


'-----------------------------------------------------
'窗体公共方法
'-----------------------------------------------------
Public Sub zlRefresh(ByVal lngFileID As Long)
    '功能：刷新显示
Dim rsTemp As New ADODB.Recordset
Dim strTemp As String, lngCount As Long
    
    mlngFileID = lngFileID
    '--------------------------------------------
    Me.lbl通用内容 = "": Me.lbl要求内容.Caption = "": Me.lbl说明内容 = "": Me.vgdRequest.Visible = False
    If mlngFileID = 0 Then Call Form_Resize: Exit Sub
    
    '--------------------------------------------
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select 种类, 编号, 名称, 通用, 说明 From 病历文件列表 Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    With rsTemp
        If .RecordCount = 0 Then MsgBox "文件丢失(可能被其他用户删除)！", vbInformation, gstrSysName: Exit Sub
        mintKind = !种类
        Me.lbl通用内容.Tag = IIf(IsNull(!通用), 0, !通用)
        Me.lbl说明内容.Caption = "说明:" & !说明
    End With
    Select Case Val(Me.lbl通用内容.Tag)
    Case 0: Me.lbl通用内容.Caption = "该病历文件暂时还不允许使用。"
    Case 1: Me.lbl通用内容.Caption = "该病历文件适合所有科室使用。"
    Case Else
        gstrSQL = "Select d.编码, d.名称 From 部门表 d, 病历应用科室 s Where d.Id = s.科室id And 文件id =[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
        Me.lbl通用内容.Caption = ""
        With rsTemp
            Do While Not .EOF()
                Me.lbl通用内容.Caption = Me.lbl通用内容.Caption & "、[" & !编码 & "]" & !名称
                .MoveNext
            Loop
        End With
        If Me.lbl通用内容.Caption = "" Then
            Me.lbl通用内容.Caption = "尚未设置本病历文件的适用科室！"
        Else
            Me.lbl通用内容.Caption = "本病历文件的适用科室包括：" & Mid(Me.lbl通用内容, 2) & "。"
        End If
    End Select
    
    '--------------------------------------------
    Select Case mintKind
    Case 1      '门诊病历
        Me.lbl要求标题.Caption = "2)时限要求:"
        gstrSQL = "Select 事件, 必须 From 病历时限要求 Where 文件id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
        With rsTemp
            If .RecordCount = 0 Then
                Me.lbl要求内容.Caption = "尚未设置本病历文件的时限要求！"
            Else
                Me.lbl要求内容.Caption = "在病人" & !事件 & "时，" & IIf(!必须 = 0, "可以", "必须") & "书写本病历文件。"
            End If
        End With
    Case 2, 4       '住院病历、护理病历
        Me.lbl要求标题.Caption = "2)时限要求:"
        gstrSQL = "Select 事件, Nvl(必须,0) As 必须, Nvl(唯一,0) As 唯一," & _
                "       Nvl(书写时限,0) As 书写时限, Nvl(审阅时限,0) As 审阅时限, Nvl(诊断时限,0) As 诊断时限," & _
                "       Nvl(一般周期,0) As 一般周期, Nvl(病重周期,0) As 病重周期, Nvl(病危周期,0) As 病危周期" & _
                " From 病历时限要求 Where 文件id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
        With rsTemp
            If .RecordCount = 0 Then
                Me.lbl要求内容.Caption = "尚未设置本病历文件的时限要求！"
            Else
                Me.lbl要求内容.Caption = "在发生病人" & !事件 & IIf(!书写时限 < 0, "前，", "后，") & _
                    IIf(!必须 = 0, "可以", "必须") & IIf(!唯一 = 0, "循环记录", "书写一次") & "本病历文件。"
                strTemp = ""
                If !唯一 <> 0 Then
                    If !书写时限 > 0 Then strTemp = strTemp & "，" & !书写时限 & "小时完成病历"
                    If !审阅时限 > 0 Then strTemp = strTemp & "，" & !审阅时限 & "小时完成审阅"
                    If !诊断时限 > 0 Then strTemp = strTemp & "，" & !诊断时限 & "小时完成修正诊断"
                Else
                    If !一般周期 > 0 Then strTemp = strTemp & "；一般病人每" & !一般周期 & "小时记录一次"
                    If !病重周期 > 0 Then strTemp = strTemp & "；病重病人每" & !病重周期 & "小时记录一次"
                    If !病危周期 > 0 Then strTemp = strTemp & "；病危病人每" & !病危周期 & "小时记录一次"
                End If
                If strTemp <> "" Then Me.lbl要求内容.Caption = Me.lbl要求内容.Caption & vbCrLf & "要求" & Mid(strTemp, 2) & "。"
            End If
        End With
        
        gstrSQL = "Select l.编号, l.名称 From 病历文件列表 l, 病历替代关系 e Where l.Id = e.替代id And e.文件id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
        strTemp = ""
        With rsTemp
            Do While Not .EOF()
                strTemp = strTemp & "、[" & !编号 & "]" & !名称
                .MoveNext
            Loop
        End With
        If strTemp <> "" Then Me.lbl要求内容.Caption = Me.lbl要求内容.Caption & vbCrLf & "完成本病历，可不必书写同期的" & Mid(strTemp, 2) & "。"
    
    Case 3      '护理记录
        Me.lbl要求标题.Caption = "2)使用要求:"
        gstrSQL = "Select Decode(nvl(f.报表, 3), 0, '特级护理', 1, '一级护理', 2, '二级护理', 3, '三级护理') As 等级" & _
                " From 病历文件列表 l, 病历页面格式 f" & _
                " Where l.种类 = f.种类 And l.页面 = f.编号 And f.种类 = 3 And l.Id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
        If rsTemp.EOF Then
            Me.lbl要求内容.Caption = ""
        Else
            Me.lbl要求内容.Caption = "适用于“" & rsTemp.Fields(0).Value & "”及以上等级的病人。"
        End If
    
    Case 5      '疾病证明与报告
        Me.lbl要求标题.Caption = "2)发生以下诊断时书写该文件:"
        Me.vgdRequest.Visible = True
        gstrSQL = "Select '疾病' As 分类,编码, 名称, p.报告病种 From 疾病编码目录 i, 疾病报告前提 p Where i.Id = p.疾病id And p.文件id = [1]"
        gstrSQL = gstrSQL & " Union All Select  '诊断' As 分类,编码, 名称, p.报告病种 From 疾病诊断目录 i, 疾病报告前提 p Where i.Id = p.诊断id And p.文件id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
        Set Me.vgdRequest.DataSource = rsTemp
        With vgdRequest
            For lngCount = 0 To .Cols - 1
                .FixedAlignment(lngCount) = flexAlignCenterCenter
                .ColAlignment(lngCount) = flexAlignLeftCenter
            Next
            
            For lngCount = 1 To .Rows - 1
                If .TextMatrix(lngCount, 0) = "诊断" Then
                    .Cell(flexcpForeColor, lngCount, 0, .Rows - 1, .Cols - 1) = &HFF0000
                    Exit For
                End If
                
            Next
            
            .MergeCells = flexMergeFree
            .MergeCol(0) = True
            .ColWidth(0) = 510
            .ColWidth(1) = 1000: .ColWidth(3) = 1000
        End With
        
    Case 6      '知情文件
        Me.lbl要求标题.Caption = "2)进行以下诊疗措施前书写该文件:"
        Me.vgdRequest.Visible = True
        
        gstrSQL = "Select Distinct i.编码, i.名称, k.名称 As 类别" & _
                " From 诊疗项目类别 k, 诊疗项目目录 i, 病历单据应用 a" & _
                " Where k.编码 = i.类别 And i.Id = a.诊疗项目id And a.病历文件id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
        Set Me.vgdRequest.DataSource = rsTemp
        With Me.vgdRequest
            For lngCount = 0 To .Cols - 1
                .FixedAlignment(lngCount) = flexAlignCenterCenter
                .ColAlignment(lngCount) = flexAlignLeftCenter
            Next
            .ColWidth(0) = 1000: .ColWidth(2) = 700
        End With
    
    Case 7      '诊疗单据
    
    End Select
    
    Call Form_Resize
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

