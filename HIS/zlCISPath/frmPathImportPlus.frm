VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPathImportPlus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "临床路径选择"
   ClientHeight    =   9960
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12255
   Icon            =   "frmPathImportPlus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9960
   ScaleWidth      =   12255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8370
      Left            =   0
      ScaleHeight     =   8370
      ScaleWidth      =   12255
      TabIndex        =   5
      Top             =   800
      Width           =   12255
      Begin VSFlex8Ctl.VSFlexGrid vsPath 
         Height          =   4065
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   12015
         _cx             =   1973310153
         _cy             =   1973296130
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
         BackColorFixed  =   15597549
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16444122
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   32768
         GridColorFixed  =   32768
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   320
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPathImportPlus.frx":6852
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   0   'False
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
         BackColorFrozen =   14811105
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsDisease 
         Height          =   3465
         Left            =   120
         TabIndex        =   14
         Top             =   4800
         Width           =   12015
         _cx             =   1973310153
         _cy             =   1973295072
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
         BackColorFixed  =   15597549
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16444122
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   32768
         GridColorFixed  =   32768
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   320
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPathImportPlus.frx":68CB
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   0   'False
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
         BackColorFrozen =   14811105
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "请从下表中选择一个适用于该病人的诊断"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   4560
         Width           =   3240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   120
         X2              =   12840
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   120
         X2              =   12840
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "请从下表中选择一个适用于该病人的临床路径"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   3600
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   12255
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   9165
      Width           =   12255
      Begin VB.CommandButton cmdPathOut 
         Caption         =   "常规治疗"
         Height          =   350
         Left            =   9720
         TabIndex        =   15
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdPathIn 
         Caption         =   "入径治疗"
         Height          =   350
         Left            =   10920
         TabIndex        =   8
         Top             =   240
         Width           =   1100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   6
         X1              =   0
         X2              =   12720
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   7
         X1              =   0
         X2              =   12720
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   12255
      TabIndex        =   1
      Top             =   0
      Width           =   12255
      Begin VB.Frame fraSel 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFF0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   240
         TabIndex        =   10
         Top             =   277
         Width           =   3495
         Begin VB.OptionButton optSel 
            Appearance      =   0  'Flat
            BackColor       =   &H00EFF0E0&
            Caption         =   "按诊断编码输入"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton optSel 
            BackColor       =   &H00EFF0E0&
            Caption         =   "按疾病编码输入"
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   11
            Top             =   0
            Width           =   1575
         End
      End
      Begin VB.Frame fraDiag 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFF0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   3840
         TabIndex        =   2
         Top             =   230
         Width           =   5055
         Begin VB.CommandButton cmd 
            Caption         =   "…"
            Height          =   285
            Left            =   3660
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "选择诊断"
            Top             =   35
            Width           =   285
         End
         Begin VB.TextBox txtDiagnose 
            Height          =   330
            Left            =   480
            TabIndex        =   3
            ToolTipText     =   "录入诊断查找路径"
            Top             =   0
            Width           =   3495
         End
         Begin VB.Label lblDiagnose 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "诊断"
            Height          =   180
            Left            =   0
            TabIndex        =   4
            Top             =   75
            Width           =   360
         End
      End
      Begin MSComctlLib.ImageList imgSrc 
         Left            =   11520
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPathImportPlus.frx":692B
               Key             =   "chkRed"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPathImportPlus.frx":6EC5
               Key             =   "unchkRed"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPathImportPlus.frx":745F
               Key             =   "chkRedUnSquare"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPathImportPlus.frx":79F9
               Key             =   "unchkBlue"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPathImportPlus.frx":7F93
               Key             =   "chkBlue"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmPathImportPlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mPati As TYPE_Pati
Private mPP As TYPE_PATH_Pati

Private mfrmParent As Object

Private mstrSex    As String
Private mblnOK As Boolean
Private mbln中医科 As Boolean
Private mbln外挂 As Boolean

Private mintDiagInput As Integer        'mintDiagInput:1-由医生选择输入来源,2-按照诊断标准输入,3-按照疾病编码输入
Private mintDiagInputZY As Integer      '系统允许选择诊断输入方式（mintDiagInput=1）时住院中医诊断：0-根据诊断标准输入,1-根据疾病编码输入
Private mintDiagInputXY As Integer      '系统允许选择诊断输入方式（mintDiagInput=1）时住院西医诊断：0-根据诊断标准输入,1-根据疾病编码输入
Private mintDiag As Integer             '记录诊断编码输入方式

Private mrsPati As ADODB.Recordset
Private mrsPath As ADODB.Recordset      '缓存路径
Private mrsPathDept As ADODB.Recordset      '缓存路径
Private mrsDisease As ADODB.Recordset   '缓存病种

Private mcolPati As Collection

Private mblnICD11 As Boolean
Private mblnHave As Boolean    'T-存在诊断

Private Enum E_诊断
    E_IX_按诊断 = 0
    E_IX_按疾病 = 1
End Enum

Public Function ShowMe(frmParent As Object, t_pati As TYPE_Pati, ByRef t_pp As TYPE_PATH_Pati) As Boolean
'参数：
    mPati = t_pati
    mPP = t_pp
    Set mfrmParent = frmParent
    mbln中医科 = Sys.DeptHaveProperty(mPati.科室ID, "中医科")
    
    Set mrsPati = GetPatiInfo(mPati.病人ID, mPati.主页ID, mcolPati)
    If mrsPati.RecordCount = 0 Then
        MsgBox "读取病人当前住院信息失败。", vbInformation, gstrSysName
        Exit Function
    End If
    Set mrsPathDept = GetPathTable(0, 0, mPati.科室ID, -1)
    If mrsPathDept.RecordCount = 0 Then
        MsgBox "本科室没有符合当前病人的有效临床路径。", vbInformation, gstrSysName
        Exit Function
    End If
    mblnICD11 = IsICDElevent()
    mblnOK = False
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Sub cmdPathOut_Click()
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()
    Dim bln参数设置 As String
    '参数读取
    
    bln参数设置 = InStr(GetInsidePrivs(p住院医生站), "参数设置") > 0
    mintDiagInput = Val(zlDatabase.GetPara(55, glngSys, , 1))
    mintDiagInputXY = Val(zlDatabase.GetPara("西医诊断输入", glngSys, p住院医生站, 0, Array(optSel(E_IX_按疾病), optSel(E_IX_按诊断)), bln参数设置))
    If mbln中医科 Then
        mintDiagInputZY = Val(zlDatabase.GetPara("中医诊断输入", glngSys, p住院医生站, 0, Array(optSel(E_IX_按疾病), optSel(E_IX_按诊断)), bln参数设置))
    End If

    '加载路径, a.适用性别, a.适用年龄, a.最新版本,Nvl(a.病例分型,'无') as 病例分型
    Call Grid.Init(vsPath, "选择,600,4;编码,1200,1;名称,3495,1;说明,4995,1")
    Call Grid.Init(vsDisease, "选择,600,4;编码,1200,1;名称,3495,1;类别")
    With vsPath
        .RowHeightMin = 330
    End With
    With vsDisease
        .RowHeightMin = 330
    End With
    Set mrsPath = mrsPathDept
    Call LoadPath(mrsPath)
    If mbln外挂 Then mbln外挂 = False: Exit Sub
End Sub

Private Sub cmdPathIn_Click()
    '根据导入路径
    If Not SaveData() Then Exit Sub
    mblnOK = True
End Sub

Private Sub Form_Activate()
    Call txtDiagnose.SetFocus
End Sub

Private Sub ResizeDiagWay()
'功能:设置诊断录入方式
    If mblnICD11 Then
        fraSel.Visible = False
        fraDiag.Left = fraSel.Left
        mintDiag = 1
    Else
        'ICD-10
        If mintDiagInput = 1 Then
            fraSel.Visible = True
            If mbln中医科 Then
                mintDiag = mintDiagInputZY
            Else
                mintDiag = mintDiagInputXY
            End If
        Else
            fraSel.Visible = False
            fraDiag.Left = fraSel.Left
            mintDiag = mintDiagInput - 2
        End If
    End If
    optSel(mintDiag).Value = True
End Sub

Private Sub LoadPath(ByVal rsPath As ADODB.Recordset)
    Dim i As Long
    
    rsPath.Filter = ""
    vsDisease.Rows = 1: vsDisease.Rows = 2
    With vsPath
        .Rows = 1 '清空历史数据
        .Rows = rsPath.RecordCount + 1
        .AllowUserResizing = flexResizeColumns
        If .Rows = 1 Then .Rows = .Rows + 1
        For i = 1 To rsPath.RecordCount
            .RowData(i) = Val(rsPath!ID & "")
            .Cell(flexcpData, i, .ColIndex("编码")) = Val(rsPath!最新版本 & "")
            .Cell(flexcpPictureAlignment, i, .ColIndex("选择")) = flexAlignCenterCenter
            .TextMatrix(i, .ColIndex("编码")) = rsPath!编码 & ""
            .TextMatrix(i, .ColIndex("名称")) = rsPath!名称 & ""
            .TextMatrix(i, .ColIndex("说明")) = rsPath!说明 & ""
            rsPath.MoveNext
        Next
        If rsPath.RecordCount = 1 Then
            .Row = 1
            Set .Cell(flexcpPicture, .Row, .ColIndex("选择")) = imgSrc.ListImages("chkRedUnSquare").Picture
            Call LoadDisease(.RowData(.Row)) '仅一行时默认加载病种
        End If
        cmdPathIn.Enabled = rsPath.RecordCount > 0
    End With
End Sub

Private Sub LoadDisease(ByVal lngPathID As Long)
    Dim i As Long, lngSel As Long
    Dim rsTmp As ADODB.Recordset
    Dim blnRead As Boolean
    
    If mrsDisease Is Nothing Then
        Call InitRSDisease
        blnRead = True
    Else
        mrsDisease.Filter = "路径ID = " & lngPathID
        If mrsDisease.RecordCount = 0 Then blnRead = True
    End If
    If blnRead = True Then
        Set rsTmp = GetPathDisease(lngPathID)
        Do While Not rsTmp.EOF
            '路径ID,疾病ID,疾病码,疾病名,诊断ID,诊断码,诊断名
            mrsDisease.AddNew Array("路径ID", "疾病ID", "疾病码", "疾病名", "诊断ID", "诊断码", "诊断名", "类别"), _
            Array(lngPathID, Nvl(rsTmp!疾病id, 0), rsTmp!疾病码 & "", rsTmp!疾病名 & "", Nvl(rsTmp!诊断id, 0), rsTmp!诊断码 & "", rsTmp!诊断名 & "", rsTmp!类别 & "")
            rsTmp.MoveNext
        Loop
        mrsDisease.Filter = "路径ID = " & lngPathID
    End If
    
    With vsDisease
        .Rows = 1: '清空历史数据
        .Rows = mrsDisease.RecordCount + 1
        If .Rows = 1 Then .Rows = .Rows + 1
        For i = 1 To mrsDisease.RecordCount
            .RowData(i) = Val(mrsDisease!疾病id & "")
            .Cell(flexcpPictureAlignment, i, .ColIndex("选择")) = flexAlignCenterCenter
            .Cell(flexcpData, i, .ColIndex("编码")) = Val(mrsDisease!诊断id & "")
            .TextMatrix(i, .ColIndex("编码")) = IIf(Val(mrsDisease!疾病id & "") > 0, mrsDisease!疾病码 & "", mrsDisease!诊断码 & "")
            .TextMatrix(i, .ColIndex("名称")) = IIf(Val(mrsDisease!疾病id & "") > 0, mrsDisease!疾病名 & "", mrsDisease!诊断名 & "")
            .TextMatrix(i, .ColIndex("类别")) = mrsDisease!类别 & ""
            If mrsDisease!疾病id & "" = txtDiagnose.Tag Or mrsDisease!诊断id & "" = txtDiagnose.Tag Then
                .Row = i '缺省不定位数据 病种有且仅有一个缺省定位;病种和录入诊断一致缺省定位
            End If
            mrsDisease.MoveNext
        Next
        If mrsDisease.RecordCount > 0 And .Row < 1 Then
            .Row = 1
            Set .Cell(flexcpPicture, .Row, .ColIndex("选择")) = imgSrc.ListImages("chkRedUnSquare").Picture
        End If
        If .Row > 0 Then
            .ShowCell .Row, .ColIndex("名称")
        End If
    End With
End Sub

Private Sub InitRSDisease()
'功能:初始化记录
'     列名: 路径ID,疾病ID,疾病码,疾病名,诊断ID,诊断码,诊断名
    Set mrsDisease = New ADODB.Recordset
    mrsDisease.Fields.Append "路径ID", adBigInt
    mrsDisease.Fields.Append "疾病ID", adBigInt
    mrsDisease.Fields.Append "疾病码", adVarChar, 20, adFldIsNullable
    mrsDisease.Fields.Append "疾病名", adVarChar, 200, adFldIsNullable
    mrsDisease.Fields.Append "诊断ID", adBigInt
    mrsDisease.Fields.Append "诊断码", adVarChar, 20, adFldIsNullable
    mrsDisease.Fields.Append "诊断名", adVarChar, 200, adFldIsNullable
    mrsDisease.Fields.Append "类别", adVarChar, 1, adFldIsNullable

    mrsDisease.CursorLocation = adUseClient
    mrsDisease.LockType = adLockOptimistic
    mrsDisease.CursorType = adOpenStatic
    mrsDisease.Open
End Sub

Private Sub Form_Resize()
    Call ResizeDiagWay
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Not mblnOK Then
        Cancel = 1 '只能通过按钮取消
    Else
        Set mrsPath = Nothing
        Set mrsDisease = Nothing
        Set mrsPathDept = Nothing
        Set mrsPati = Nothing
        Set mcolPati = Nothing
    End If
End Sub

Private Sub optSel_Click(Index As Integer)
    If optSel(Index).Value Then mintDiag = Index
End Sub

Private Sub txtDiagnose_Change()
    If Trim(txtDiagnose.Text) = "" And txtDiagnose.Tag <> "" Then
        txtDiagnose.Tag = ""
        Set mrsPath = mrsPathDept
        Call LoadPath(mrsPath)
    End If
End Sub

Private Sub txtDiagnose_GotFocus()
    Call zlControl.TxtSelAll(txtDiagnose)
End Sub

Private Sub txtDiagnose_KeyPress(KeyAscii As Integer)
    Dim strInput As String
    Dim strSql As String
    Dim vRect As RECT
    Dim blnCancel As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim lngNum As Long
    Dim strType As String
    
    If KeyAscii = 13 Then
        strInput = UCase(Trim(txtDiagnose.Text))
        txtDiagnose.Tag = ""
        If strInput = "" Then
'            Set mrsPath = mrsPathDept
'            Call LoadPath(mrsPath)
            Exit Sub
        End If
        If mblnICD11 Then
            If mbln中医科 Then
                strType = ",1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,24,25,26,27,"
                strSql = "Select 序号 From 疾病编码分类 Where 章节 = [1] And 名称 = [2] And 编码 = [3]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, "26", "传统医学证候（TM1）", "L1-SE7")
                If Not rsTmp.EOF Then
                    lngNum = Val("" & rsTmp!序号)
                    strSql = IIf(lngNum <> 0, " And a.分类ID Not In (Select e.ID From 疾病编码分类 e where e.章节='26' And e.序号>=" & lngNum & ")", "")
                End If
            Else
                strType = ",1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,24,25,27,"
            End If

            If zlCommFun.IsCharChinese(strInput) Then
                strSql = strSql & " And A.名称 Like [2]" '输入汉字时,只匹配名称
            Else
                strSql = strSql & " And (A.编码 Like [1] Or A.名称 Like [2] Or " & IIf(gint简码 = 0, "A.简码", "A.五笔码") & " Like [2])"
            End If

            strSql = _
                " Select A.ID,A.ID as 项目ID,A.编码,A.附码,A.名称," & IIf(gint简码 = 0, "A.简码", "A.五笔码 as 简码") & ",A.说明" & _
                " From 疾病编码目录 A Where A.类别 ='E' And Instr([5],','||A.章节||',')>0 " & strSql & _
                IIf(mstrSex <> "", " And (A.性别限制=[3] Or A.性别限制 is NULL)", "") & _
                " And (A.撤档时间 is Null Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Order by A.编码"
        Else
            If mintDiag = E_IX_按诊断 Then
                '按诊断输入:西医部份，一个诊断可能属于多个分类
                If zlCommFun.IsCharChinese(strInput) Then
                    strSql = "B.名称 Like [2]" '输入汉字时,只匹配名称
                Else
                    strSql = "A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2]"
                End If
                strSql = _
                    " Select Distinct A.ID,A.ID as 项目ID,A.编码,A.名称,A.说明,A.编者" & _
                    " From 疾病诊断目录 A,疾病诊断别名 B" & _
                    " Where A.ID=B.诊断ID And A.类别=1" & _
                    " And B.码类=[4] And (" & strSql & ")" & _
                    " And (A.撤档时间 is Null Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Order by A.编码"
            Else
                'D-ICD-10疾病编码
                If zlCommFun.IsCharChinese(strInput) Then
                    strSql = "名称 Like [2]" '输入汉字时,只匹配名称
                Else
                    strSql = "编码 Like [1] Or 名称 Like [2] Or " & IIf(gint简码 = 0, "简码", "五笔码") & " Like [2]"
                End If
                strSql = _
                    " Select ID,ID as 项目ID,编码,附码,名称," & IIf(gint简码 = 0, "简码", "五笔码 as 简码") & ",说明" & _
                    " From 疾病编码目录 Where 类别 In('D','B') And (" & strSql & ")" & _
                    IIf(mstrSex <> "", " And (性别限制=[3] Or 性别限制 is NULL)", "") & _
                    " And (撤档时间 is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Order by 编码"
            End If
        End If
        vRect = zlControl.GetControlRect(txtDiagnose.Hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, IIf(optSel(1).Value, "诊断编码", "疾病编码"), _
            False, "", "", False, False, True, vRect.Left, vRect.Top, txtDiagnose.Height, blnCancel, False, True, _
            strInput & "%", gstrLike & strInput & "%", mstrSex, gint简码 + 1, strType)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "没有找到与输入匹配的内容。", vbInformation, gstrSysName
            End If
            Call zlControl.TxtSelAll(txtDiagnose)
            Exit Sub
        Else
            txtDiagnose.Text = "[" & rsTmp!编码 & "]" & Nvl(rsTmp!名称)
            txtDiagnose.Tag = Val(rsTmp!项目ID)
            Set mrsPath = GetPathTable(IIf(mintDiag = E_IX_按疾病, Val(rsTmp!项目ID), 0), IIf(mintDiag = E_IX_按诊断, Val(rsTmp!项目ID), 0), mPati.科室ID, 0)
            Call LoadPath(mrsPath)
        End If
    End If
End Sub

Private Sub cmd_Click()
    Dim rsTmp As ADODB.Recordset
    If mblnICD11 Then
        Set rsTmp = ShowILLSelect(Me, "E", mPati.科室ID, mstrSex, True, True, , , , 1, True, , 1)
    Else
        If mintDiag = E_IX_按诊断 Then
            '按诊断输入:西医部份，一个诊断可能属于多个分类
            Set rsTmp = ShowILLSelect(Me, "1", mPati.科室ID, mstrSex, False, False)
        Else
            'D-ICD-10疾病编码
            Set rsTmp = ShowILLSelect(Me, "D,B", mPati.科室ID, mstrSex, False, True)
        End If
    End If
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            txtDiagnose.Text = "[" & rsTmp!编码 & "]" & Nvl(rsTmp!名称)
            txtDiagnose.Tag = Val(rsTmp!项目ID)
            Set mrsPath = GetPathTable(IIf(mintDiag = E_IX_按疾病, Val(rsTmp!项目ID), 0), IIf(mintDiag = E_IX_按诊断, Val(rsTmp!项目ID), 0), mPati.科室ID, 0)
            Call LoadPath(mrsPath)
        End If
    End If
End Sub

Private Function SaveData() As Boolean
    Dim arrSQL As Variant
    Dim lng住院天数 As Long, lng标准住院日 As Long
    Dim rsTmp As ADODB.Recordset, rsCriterion As ADODB.Recordset
    Dim str未导入编码 As String, str未导入名称 As String
    Dim lngB As Long, lngE As Long, strUnit As String, strTmp As String, DatCur As Date, lngValue As Long

    Dim bln外挂判断 As Boolean
    Dim blnTrans As Boolean
    
    Dim dt入院时间 As Date
    Dim dtDate As Date
    Dim bytDiagSorce As Long
    Dim bytDiagType As Long
    Dim lngPatiDiagID As Long
    Dim strDiagInfo As String   '诊断描述
    
    Dim i As Long
    
    If vsPath.Row < 1 Then
         MsgBox "请选择一个适用于该病人的临床路径。", vbInformation + vbOKOnly, gstrSysName
         Exit Function
    End If
    If vsDisease.Row < 1 Then
        MsgBox "请选择一个适用于该病人的诊断。", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    Else
        If vsDisease.TextMatrix(vsDisease.Row, vsDisease.ColIndex("名称")) = "" Then
            MsgBox "该临床路径“" & vsPath.TextMatrix(vsPath.Row, vsPath.ColIndex("名称")) & "”没有有效的诊断信息。", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    If vsDisease.TextMatrix(vsDisease.Row, vsDisease.ColIndex("类别")) = "E" And Not mblnICD11 Then
        If mblnHave Then
            MsgBox "病人已经录入非ICD-11的诊断，请选择非ICD-11的诊断导入路径。", vbInformation + vbOKOnly, gstrSysName
        Else
            MsgBox "系统参数未开启ICD-11模式，请选择非ICD-11的诊断导入路径。", vbInformation + vbOKOnly, gstrSysName
        End If
        Exit Function
    ElseIf vsDisease.TextMatrix(vsDisease.Row, vsDisease.ColIndex("类别")) <> "E" And mblnICD11 Then
        If mblnHave Then
            MsgBox "病人已经录入ICD-11的诊断，请选择ICD-11的诊断导入路径。", vbInformation + vbOKOnly, gstrSysName
        Else
            MsgBox "系统参数开启ICD-11模式，请选择ICD-11的诊断导入路径。", vbInformation + vbOKOnly, gstrSysName
        End If
        Exit Function
    End If
    mrsPath.Filter = "ID = " & vsPath.RowData(vsPath.Row)
    mPP.路径ID = mrsPath!ID
    mPP.版本号 = mrsPath!最新版本
    
    With vsDisease
        If Val(.RowData(.Row)) > 0 Then  '疾病ID
            mrsDisease.Filter = "路径ID=" & mrsPath!ID & " And 疾病ID =" & .RowData(.Row)
        Else
            mrsDisease.Filter = "路径ID=" & mrsPath!ID & " And 诊断ID =" & Val(.Cell(flexcpData, .Row, .ColIndex("编码")))
        End If
    End With
    
    'mbytDiagSorce=诊断来源1-病历；2-入院登记；3-首页整理;4-病案
    bytDiagSorce = 3
    'mbytDiagType=诊断类型:1-西医门诊诊断;2-西医入院诊断;11-中医门诊诊断;12-中医入院诊断
    If InStr(",B,2,", "," & mrsDisease!类别 & ",") > 0 Then
        bytDiagType = 12
    Else
        bytDiagType = 2
    End If
    
    
    '填了病例分型，并且路径表也定义了要求
    If Not IsNull(mrsPati!病例分型) Then
        If mrsPath!病例分型 <> "无" And mrsPati!病例分型 <> mrsPath!病例分型 Then
            MsgBox "该路径要求的病例分型[" & mrsPath!病例分型 & "]不适合于该病人的病例分型[" & mrsPati!病例分型 & "]", vbInformation, gstrSysName
            str未导入名称 = "病例分型不适用"
            GoTo UnImport
        End If
    End If
    
    If Not IsNull(mrsPati!当前病况) Then
        If mrsPath!适用病情 <> "通用" And mrsPath!适用病情 <> mrsPati!当前病况 Then
            MsgBox "该路径[" & mrsPath!适用病情 & "]不适合于该病人病情[" & mrsPati!当前病况 & "]", vbInformation, gstrSysName
            str未导入名称 = "病情不适用"
            GoTo UnImport
        End If
    End If
    If Val(mrsPath!适用性别) <> 0 Then
        If Val(mrsPath!适用性别) <> IIf(mrsPati!性别 = "男", 1, IIf(mrsPati!性别 = "女", 2, 0)) Then
            MsgBox "该路径不适合于该病人性别[" & mrsPati!性别 & "]", vbInformation, gstrSysName
            str未导入名称 = "性别不适合"
            GoTo UnImport
        End If
    End If
    
    If Not IsNull(mrsPath!适用年龄) And Not IsNull(mrsPati!年龄) Then
        lngValue = 0
        lngB = Split(mrsPath!适用年龄, "-")(0)
        strTmp = Split(mrsPath!适用年龄, "-")(1)
        lngE = Mid(strTmp, 1, Len(strTmp) - 1)
        strUnit = Mid(strTmp, Len(strTmp))
    
        strTmp = mrsPati!年龄           '特殊：2岁3月等
        If strUnit = Mid(strTmp, Len(strTmp)) And IsNumeric(Mid(strTmp, 1, Len(strTmp) - 1)) Or IsNumeric(strTmp) Then
            '相同年龄单位才做比较
            lngValue = Val(strTmp)
        ElseIf mcolPati("_pati_birthdate") <> "" Then
            DatCur = zlDatabase.Currentdate
            lngValue = DateDiff(IIf(strUnit = "岁", "yyyy", IIf(strUnit = "月", "m", "d")), CDate(mcolPati("_pati_birthdate")), DatCur)
            If lngValue = 0 Then lngValue = 1
        End If
        If lngValue <> 0 Then
            If lngValue < lngB Or lngValue > lngE Then
                MsgBox "该路径不适合于该病人年龄[" & mrsPati!年龄 & "]", vbInformation, gstrSysName
    
                str未导入名称 = "年龄不适合"
                GoTo UnImport
            End If
        End If
    End If
    '住院日不能大于路径的标准住院日和确诊天数(如果没有设置了确诊天数，则不限制)
    dt入院时间 = GetPatiInDate(mPati, lng住院天数)
    dtDate = zlDatabase.Currentdate
    
    If InStr(mrsPath!标准住院日, "-") > 0 Then
        lng标准住院日 = Split(mrsPath!标准住院日, "-")(1)
    Else
        lng标准住院日 = Val(mrsPath!标准住院日)
    End If
    '住院天数超过确诊天数禁止导入路径;确诊天数未设置或为0时,则住院天数大于标准住院日时禁止导入路径
    
    If Not CheckPathSend(mPati.病人ID, mPati.主页ID) Then
        If mrsPath!确诊天数 <> 0 Then
            If dtDate > Format(DateAdd("d", Val(mrsPath!确诊天数), dt入院时间), "yyyy-MM-DD HH:mm:ss") Then
                MsgBox "该病人已入院" & lng住院天数 & "天，超过了规定的确诊天数(" & mrsPath!确诊天数 & "天)。", vbInformation, gstrSysName
                str未导入名称 = "超过确诊天数"
                GoTo UnImport
            End If
        Else
            If lng住院天数 > lng标准住院日 Then
                MsgBox "该病人已入院" & lng住院天数 & "天，超过了该路径的标准住院日(" & lng标准住院日 & "天)。", vbInformation, gstrSysName
                str未导入名称 = "超过标准住院日"
                GoTo UnImport
            End If
        End If
    End If
     
     
    Me.Hide
    bln外挂判断 = True
    '临床路径导入前调用外挂口
    If CreatePlugInOK(P临床路径应用) Then
        On Error Resume Next
        bln外挂判断 = gobjPlugIn.PathImportBefore(glngSys, P临床路径应用, mPati.病人ID, mPati.主页ID, mPP.路径ID, mPP.版本号, bytDiagType, bytDiagSorce, _
        Val(mrsDisease!疾病id & ""), Val(mrsDisease!诊断id & ""))
        '如果接口不存在，不影响原有逻辑
        If Not bln外挂判断 And Err.Number <> 0 Then bln外挂判断 = True
        Call zlPlugInErrH(Err, "PathImportBefore")
        Err.Clear: On Error GoTo 0
        If Not bln外挂判断 Then
            mbln外挂 = True
            mblnOK = True
            Unload Me
            Exit Function
        End If
    End If
    '
    arrSQL = Array()
    lngPatiDiagID = zlDatabase.GetNextId("病人诊断记录")
    strDiagInfo = IIf(Val(mrsDisease!疾病id) > 0, "(" & mrsDisease!疾病码 & ")" & mrsDisease!疾病名, "") & IIf(Val(mrsDisease!诊断id) > 0, "(" & mrsDisease!诊断码 & ")" & mrsDisease!诊断名, "")
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_INSERT(" & mPati.病人ID & "," & mPati.主页ID & ",3,NULL," & bytDiagType & "," & _
                        ZVal(mrsDisease!疾病id) & "," & ZVal(mrsDisease!诊断id) & ",NULL,'" & _
                        strDiagInfo & "','',0,0," & zlStr.To_Date(Format(dtDate, "yyyy-MM-DD HH:mm:ss"), "ymdhms") & ",'',1,'','',NULL,Null," & lngPatiDiagID & ",NULL,'',''" & _
                        IIf(mrsDisease!类别 & "" = "E", ",'E',1,'01'", ",NULL,NULL,NULL") & ",NULL,NULL,NULL,1)"
                        
 
    mblnOK = frmEvaluate.ShowMe(mfrmParent, 0, 1, mPati, mPP, mrsPath!名称, bytDiagType, bytDiagSorce, Val(mrsDisease!疾病id & ""), Val(mrsDisease!诊断id & ""), 0, , , , arrSQL)
     
    '临床路径导后前调用外挂口
    If CreatePlugInOK(P临床路径应用) Then
        On Error Resume Next
        Call gobjPlugIn.PathImportAfter(glngSys, P临床路径应用, mPati.病人ID, mPati.主页ID, mPP.路径ID, mPP.版本号)
        Call zlPlugInErrH(Err, "PathImportAfter")
        Err.Clear: On Error GoTo 0
    End If
    
    Unload Me
    Exit Function

UnImport:
    '首要诊断才保存未导入原因
    Set rsTmp = GetUnImportReson
    rsTmp.Filter = "名称='" & str未导入名称 & "'"
    If rsTmp.RecordCount = 0 Then
        str未导入编码 = ""
    Else
        str未导入编码 = rsTmp!编码
    End If
    
    Call SaveUnImport(mPati, mPP, str未导入编码, str未导入名称, bytDiagType, bytDiagSorce, Val(mrsDisease!疾病id & ""), Val(mrsDisease!诊断id & ""))
    mblnOK = True
    Unload Me
End Function


Private Sub vsDisease_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsDisease
        If Me.Visible And NewRow > 0 Then
            If OldRow > 0 Then
                 Set .Cell(flexcpPicture, OldRow, .ColIndex("选择")) = Nothing
            End If
            If .TextMatrix(NewRow, .ColIndex("名称")) <> "" Then
                Set .Cell(flexcpPicture, NewRow, .ColIndex("选择")) = imgSrc.ListImages("chkRedUnSquare").Picture
            End If
        End If
    End With
End Sub

Private Sub vsPath_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsPath
        If Me.Visible And NewRow > 0 Then
            If OldRow > 0 Then
                Set .Cell(flexcpPicture, OldRow, .ColIndex("选择")) = Nothing
            End If
            If .TextMatrix(NewRow, .ColIndex("名称")) <> "" Then
                Set .Cell(flexcpPicture, NewRow, .ColIndex("选择")) = imgSrc.ListImages("chkRedUnSquare").Picture
            End If
        End If
    End With
End Sub

Private Sub vsPath_Click()
    With vsPath
        If .RowData(.Row) <> "" Then
            Call LoadDisease(.RowData(.Row))
        End If
    End With
End Sub

Private Function IsICDElevent() As Boolean
' 新开导入路径
'    通过病人ID\主页ID查找病人诊断记录
'    记录为空
'        启用ICD-11, 走ICD-11
'        未启用ICD-11, 走ICD-10
'    记录不为空
'        存在ICD-10, 走ICD-10
'        存在ICD-11, 走ICD-11
    Dim rsTmp As ADODB.Recordset
    Dim blnResult As Boolean
    
    Set rsTmp = GetDiagType(mPati.病人ID, mPati.主页ID)
    If rsTmp.EOF Then
        blnResult = Mid(gstrICDEleven, 2, 1) = "1"
        mblnHave = False
    Else
        blnResult = (rsTmp!编码类别 & "" = "E")
        mblnHave = True
    End If
    IsICDElevent = blnResult
End Function

