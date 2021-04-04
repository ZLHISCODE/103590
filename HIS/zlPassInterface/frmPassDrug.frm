VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPassDrug 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   10320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11250
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmPassDrug.frx":0000
   ScaleHeight     =   10320
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer tmrShow 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8160
      Top             =   7560
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8F2EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7695
      Left            =   0
      ScaleHeight     =   7695
      ScaleWidth      =   1935
      TabIndex        =   5
      Top             =   1320
      Width           =   1935
      Begin VB.PictureBox picLeftCenter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6255
         Left            =   120
         ScaleHeight     =   6255
         ScaleWidth      =   2655
         TabIndex        =   8
         Top             =   360
         Width           =   2655
         Begin VB.PictureBox picNavigate 
            Appearance      =   0  'Flat
            BackColor       =   &H00F8F2EC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4455
            Left            =   120
            ScaleHeight     =   4455
            ScaleWidth      =   2535
            TabIndex        =   9
            Top             =   480
            Width           =   2535
            Begin VB.PictureBox picItem 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               FillStyle       =   0  'Solid
               ForeColor       =   &H80000008&
               Height          =   495
               Index           =   0
               Left            =   120
               ScaleHeight     =   495
               ScaleWidth      =   2295
               TabIndex        =   10
               Top             =   600
               Width           =   2295
               Begin VB.Label lblTitle 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "药品名称"
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   10.5
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   210
                  Index           =   0
                  Left            =   240
                  TabIndex        =   11
                  Top             =   120
                  Width           =   840
               End
            End
         End
      End
      Begin VB.PictureBox picDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00F8F2EC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1320
         Picture         =   "frmPassDrug.frx":08CA
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   7
         Top             =   7440
         Width           =   240
      End
      Begin VB.PictureBox picUp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00F8F2EC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1440
         Picture         =   "frmPassDrug.frx":711C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   6
         Top             =   120
         Width           =   240
      End
   End
   Begin VB.Timer tmrTime 
      Interval        =   50
      Left            =   7680
      Top             =   7560
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   2040
      ScaleHeight     =   5535
      ScaleWidth      =   9495
      TabIndex        =   1
      Top             =   1440
      Width           =   9495
      Begin VSFlex8Ctl.VSFlexGrid vsInfo 
         Height          =   1215
         Left            =   240
         TabIndex        =   12
         Top             =   120
         Width           =   5055
         _cx             =   8916
         _cy             =   2143
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         MouseIcon       =   "frmPassDrug.frx":D96E
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483641
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483643
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   2
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   0
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   500
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPassDrug.frx":E248
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
         AutoSizeMouse   =   0   'False
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
   Begin VB.PictureBox picTop 
      BackColor       =   &H00D48A00&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   12135
      TabIndex        =   0
      Top             =   0
      Width           =   12135
      Begin VB.PictureBox picFind 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3600
         ScaleHeight     =   330
         ScaleWidth      =   3195
         TabIndex        =   13
         Top             =   70
         Width           =   3200
         Begin VB.Frame fraText 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   15
            TabIndex        =   16
            Top             =   15
            Width           =   2295
            Begin VB.TextBox txtFind 
               Appearance      =   0  'Flat
               BackColor       =   &H00F8F2EC&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   300
               Left            =   0
               TabIndex        =   17
               Text            =   "输入药品名称\编码\简码"
               Top             =   0
               Width           =   2400
            End
         End
         Begin VB.Frame fraFind 
            BackColor       =   &H00D48A00&
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   2460
            TabIndex        =   14
            Top             =   15
            Width           =   690
            Begin VB.Label lblFind 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "查找"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   120
               TabIndex        =   15
               Top             =   45
               Width           =   420
            End
         End
      End
      Begin VB.PictureBox picClosed 
         Appearance      =   0  'Flat
         BackColor       =   &H00D48A00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   500
         Left            =   11640
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   3
         Top             =   0
         Width           =   500
         Begin VB.Label lblClose 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "×"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   300
            Left            =   90
            TabIndex        =   4
            Top             =   120
            Width           =   300
         End
      End
      Begin VB.Label lblTip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "提示:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   6960
         TabIndex        =   18
         Top             =   150
         Width           =   525
      End
      Begin VB.Label lblFrmName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "药品说明书"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1125
      End
   End
   Begin VB.Line linScope 
      BorderColor     =   &H80000010&
      Index           =   3
      X1              =   0
      X2              =   12720
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line linScope 
      Index           =   2
      X1              =   12600
      X2              =   12600
      Y1              =   0
      Y2              =   10560
   End
   Begin VB.Line linScope 
      Index           =   1
      X1              =   -240
      X2              =   13320
      Y1              =   10080
      Y2              =   10080
   End
   Begin VB.Line linScope 
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   8040
   End
End
Attribute VB_Name = "frmPassDrug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mMoveX As Long, mMoveY As Long  '记录窗体移动前，窗体左上角与鼠标指针位置间的纵横距离
Private mudtRect As RECT
Private mudtRectClose As RECT
Private mudtRectNavi    As RECT
Private mudtPoint As POINTAPI

Private mblnMoveStart As Boolean '判断移动是否开始
Private mblnMove As Boolean
    
Private mstrDrugJson    As String
Private mbytShowStyle   As Byte        '0-左下;1-右下;2-左上;3-右上
Private mbytOpen        As Byte        '是否加载
Private mbytMode        As Byte        '窗体加载模式
Private mfrmParent      As Object
Private mblnTip         As Boolean
Private mbytSelect      As Byte         '记录当前选中项目
Private mIntBeginTiem   As Long          '允许提示信息停留时间

Private mblnKeepTip          As Boolean

Private marrType As Variant

'药品名称:通用名称,商品名,汉语拼音,英文名称,化学名称
Private Const mstrJsonKey     As String = "药品名称,药物剂型,性状,适应症,药物规格,用法用量,不良反应,禁忌症,注意事项,孕妇用药,儿童用药,老年人用药,相互作用,药物过量,药理毒理,药代动力学,贮藏条件,批准文号,生产企业,修订日期"

Public Function ShowMe(frmParent As Object, ByVal strDrugJson As String, ByVal bytStyle As Byte, Optional ByVal blnTip As Boolean) As Boolean
'功能:显示审查结果
'参数:strDrugJson "[{\"通用名称\":\"异烟肼片\",\"商品名\":null}]
        
    Set mfrmParent = frmParent
    mbytMode = bytStyle
    mblnTip = blnTip
    If mbytSelect <> 0 Then
        picItem(mbytSelect).BackColor = picNavigate.BackColor
        mbytSelect = 0
    End If
    mstrDrugJson = strDrugJson
    vsInfo.Tag = ""
    If Not mblnTip Then
        If mbytMode = 0 Then
            Me.Show mbytMode, gobjFrm
            Call LoadJson
            If UCase(Me.ActiveControl.Name) = UCase("txtFind") Then vsInfo.SetFocus
            txtFind.Text = ""
            Call txtFind_LostFocus
            mbytOpen = 1
        Else
            Call Form_Load
            Call LoadJson
            Me.Show 1, gobjFrm
        End If
    Else
        Call Form_Load
        Call LoadJson
        Me.Show 0, gobjFrm
        Me.Top = Screen.Height
        Me.Left = Screen.Width - Me.Width - 120
        tmrShow.Enabled = True
        mIntBeginTiem = Timer()
    End If
End Function

Public Function IsOpen() As Boolean
'判断窗体是否加载
    IsOpen = mbytOpen = 1
End Function

Public Sub LoadJson(Optional ByVal strJson As String)
'功能:解析药品说明书
'参数:strCaption 从指定节点处开始加载说明书
    Dim strData As String
    Dim strValue   As String
    Dim arrTemp As Variant
    Dim i As Long, j As Long
    
    If strJson <> "" Then mstrDrugJson = strJson
    If mstrDrugJson = "" Then Exit Sub
    strData = mstrDrugJson
    strData = Left(strData, Len(strData) - 2)
    strData = Right(strData, Len(strData) - 2)
    strData = Replace(strData, "\""", """")
    If mblnTip Then
        For i = LBound(marrType) To UBound(marrType)
            If i = 0 Then
                strValue = JSONParse("通用名称", strData) & ""
                strValue = JSONReplace(strValue)
                lblFrmName.Caption = IIf(strValue <> "", "【" & strValue & "】", lblFrmName.Caption)
            End If
            strValue = JSONParse(marrType(i), strData) & ""
            lblTitle(i).Tag = "【" & marrType(i) & "】" & "[;]" & strValue
        Next
        Call picItem_Click(0) '缺省显示
    Else
        With vsInfo
            .Redraw = flexRDNone
            .Rows = 0: .Rows = 1
            .Cols = 2
            .ColWidth(0) = 120
            .ColWidth(1) = 3000
            .ColAlignment(0) = flexAlignLeftCenter
            .ColAlignment(1) = flexAlignLeftCenter
            .RowHeightMin = 200
            .RowHeightMax = 0
            .ColWidthMin = 0
            .ScrollBars = flexScrollBarVertical
            .AutoResize = True
            .AllowUserResizing = flexResizeRows
            .AutoSizeMode = flexAutoSizeRowHeight
            .MergeCells = flexMergeFree
            .WordWrap = True
            For i = LBound(marrType) To UBound(marrType)
                strValue = ""
                .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, 1) = "【" & marrType(i) & "】"
                .MergeRow(.Rows - 1) = True
                .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, 1) = True
                .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, 1) = RGB(17, 92, 84)
                lblTitle(i).Tag = .Rows - 1
                If marrType(i) = "药品名称" Then
                    arrTemp = Split("通用名称,商品名,汉语拼音,英文名称,化学名称", ",")
                    For j = LBound(arrTemp) To UBound(arrTemp)
                        .Rows = .Rows + 1
                        strValue = JSONParse(arrTemp(j), strData)
                        .TextMatrix(.Rows - 1, 1) = Space(4) & arrTemp(j) & ":" & Replace(Replace(strValue, vbCr, ""), vbLf, "")
                    Next
                Else
                    strValue = JSONParse(marrType(i), strData)
                    arrTemp = Split(strValue, vbCr & vbLf)
                    For j = LBound(arrTemp) To UBound(arrTemp)
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, 1) = Space(4) & arrTemp(j)
                    Next
                End If
                .Rows = .Rows + 1
            Next
            .Redraw = flexRDDirect
            .AutoSize 0, 1, , 45
             
        End With
    End If

End Sub

Private Sub Form_GotFocus()
    If mblnTip Then
        mblnKeepTip = True
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = Asc("`") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim strTemp As String
    Dim strPara As Variant
    
    If mblnTip Then
        Me.Width = 7000
        Me.Height = 5000
        lblFrmName.Caption = "提示信息"
        mblnKeepTip = False: mIntBeginTiem = 0
        lblTip.Visible = False
        picFind.Visible = False
    Else
        If mbytMode = 1 Then picFind.Visible = False
        Me.Width = 10000
        Me.Height = 10000
        lblFrmName.Caption = "药品说明书"
        lblTip.ForeColor = vbRed
        lblTip.Caption = ""
    End If
    Call ShowStyle(0)
    picMain.BackColor = vbWhite
    picTop.BackColor = conCOLOR_TITLE_BAR
    If mblnTip Then
        marrType = Split(conSTR_Key_Tip, ",")
        strPara = Mid(gstrParaTip, 2)
        If Len(strPara) > 0 Then
            If InStr(strPara, "1") > 0 Then
                strTemp = ""
                For i = 1 To Len(strPara)
                    If Mid(strPara, i, 1) = "1" Then strTemp = strTemp & "," & marrType(i - 1)
                Next
                If strTemp <> "" Then
                    strTemp = Mid(strTemp, 2)
                    marrType = Split(strTemp, ",")
                End If
            End If
        End If
    Else
        marrType = Split(mstrJsonKey, ",")
    End If
    For i = picItem.LBound To picItem.UBound
        If i <> 0 Then
            Unload lblTitle(i)
            Unload picItem(i)
        End If
    Next
    For i = LBound(marrType) To UBound(marrType)
        If i <> 0 Then
            Load picItem(i)
            Load lblTitle(i)
            picItem(i).Visible = True
            lblTitle(i).Visible = True
            lblTitle(i).ForeColor = vbBlack
            Set lblTitle(i).Container = picItem(i)
        End If
        lblTitle(i).Caption = marrType(i)
        picItem(i).BackColor = picNavigate.BackColor
    Next
    picLeft_Resize
    vsInfo.Tag = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbytOpen = 0
    If mblnTip Then
        mblnKeepTip = False: mIntBeginTiem = 0
    End If
End Sub

Private Sub fraFind_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraFind.BackColor = conCOLOR_BULE
End Sub

Private Sub fraFind_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      fraFind.BackColor = conCOLOR_TITLE_BAR
End Sub

Private Sub lblClose_Click()
    Call picClosed_Click
End Sub

Private Sub lblFind_Click()
    Call SerachDrug
End Sub

Private Sub lblFind_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraFind.BackColor = conCOLOR_BULE
End Sub

Private Sub lblFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = 99
End Sub

Private Sub lblFind_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraFind.BackColor = conCOLOR_TITLE_BAR
End Sub

Private Sub lblTitle_Click(Index As Integer)
    picItem_Click Index
End Sub

Private Sub picClosed_Click()
    Unload Me
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    picTop.Move 15, 15, Me.ScaleWidth - 30, 500
    If mblnTip Then
        picLeft.Move 15, picTop.Top + picTop.Height, 1500, Me.ScaleHeight - (picTop.Height + picTop.Top) - 30
        picMain.Move picLeft.Left + picLeft.Width, picLeft.Top, Me.ScaleWidth - picLeft.Width - 30, Me.ScaleHeight - picLeft.Top - 30
    Else
        picLeft.Move 15, picTop.Top + picTop.Height, 1500, Me.ScaleHeight - (picTop.Height + picTop.Top) - 30
        picMain.Move picLeft.Left + picLeft.Width, picLeft.Top, Me.ScaleWidth - picLeft.Width - 30, Me.ScaleHeight - picLeft.Top - 30
    End If
    'Left
    With linScope(0)
        .X1 = 0: .X2 = 0: .Y1 = 0: .Y2 = Me.ScaleHeight
        .BorderColor = conCOLOR_TITLE_BAR
        '&H00808080&
        '&H80000010& '按钮阴影
    End With
    'bottom
    With linScope(1)
        .X1 = 0: .X2 = Me.ScaleWidth: .Y1 = Me.ScaleHeight - 15: .Y2 = Me.ScaleHeight - 15
        .BorderColor = conCOLOR_TITLE_BAR
    End With
    'right
    With linScope(2)
        .X1 = Me.ScaleWidth - 15: .X2 = Me.ScaleWidth - 15: .Y1 = 0: .Y2 = Me.ScaleHeight - 15
        .BorderColor = conCOLOR_TITLE_BAR
    End With
    'Top
    With linScope(3)
        .X1 = 0: .X2 = Me.ScaleWidth - 15: .Y1 = 0: .Y2 = 0
        .BorderColor = conCOLOR_TITLE_BAR
    End With
End Sub

Private Sub picClosed_Resize()
    On Error Resume Next
    lblClose.Move picClosed.ScaleWidth / 2 + lblClose.Width / 2, (picClosed.ScaleHeight - lblClose.Height) / 2
End Sub

Private Sub picDown_Click()
    If picLeftCenter.Height < picNavigate.Height Then
        picNavigate.Top = picLeftCenter.Height - picNavigate.Height
    End If
End Sub

Private Sub picFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.MousePointer = 99 Then
        Me.MousePointer = 0
    End If
End Sub

Private Sub picFind_Resize()
    On Error Resume Next
    lblFind.Move 120, 45
    With fraFind
        .Width = lblFind.Width + 240
        .Height = picFind.ScaleHeight - 30
        .Left = picFind.ScaleWidth - .Width - 15
        .Top = 15
        .BackColor = conCOLOR_TITLE_BAR
    End With
    fraText.Move 15, 15, fraFind.Left - 15, picFind.ScaleHeight - 30
    txtFind.Move 0, 30, fraText.Width, fraText.Height  '下移居中显示
    fraText.BackColor = picNavigate.BackColor
    txtFind.BackColor = picNavigate.BackColor
    txtFind.ForeColor = &HC0C0C0
End Sub

Private Sub picItem_Click(Index As Integer)
    Dim i As Long
    Dim arrTmp As Variant
    Dim arrLine As Variant
    Dim strTemp As String
    
    With vsInfo
        picItem(mbytSelect).BackColor = picNavigate.BackColor
        mbytSelect = Index
        If mblnTip Then
            .Redraw = flexRDNone
            .Rows = 0: .Rows = 1
            .Cols = 2
            .ColWidth(0) = 120
            .ColWidth(1) = 3000
            .ColAlignment(0) = flexAlignLeftCenter
            .ColAlignment(1) = flexAlignLeftCenter
            .RowHeightMin = 200
            .ColWidthMin = 0
            .RowHeightMax = 0
            .AutoResize = True
            .AllowUserResizing = flexResizeRows
            .AutoSizeMode = flexAutoSizeRowHeight
            .MergeCells = flexMergeFree
            .WordWrap = True
            arrTmp = Split(lblTitle(Index).Tag, "[;]")
            
            .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, 1) = arrTmp(0)
            .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, 1) = True
            .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, 1) = RGB(17, 92, 84)
            .MergeRow(.Rows - 1) = True
            arrLine = Split(arrTmp(1), vbCr & vbLf)
            For i = LBound(arrLine) To UBound(arrLine)
                If Len(arrLine(i)) > 800 Then '由于提示界面高度限制,字符过长导致单行高度超过显示高度无法完全展示
                    strTemp = arrLine(i)
                    Do While strTemp <> ""
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, 1) = Mid(strTemp, 1, 200)
                        strTemp = Mid(strTemp, 201)
                    Loop
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = arrLine(i)
                End If
            Next
            .AutoSize 0, 1, , 45
            .Redraw = flexRDDirect
        Else
            .TopRow = CLng(lblTitle(Index).Tag)
        End If
    End With
End Sub

Private Sub picItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intPre As Integer
    If mblnTip Then Exit Sub
    intPre = Val(picNavigate.Tag)
    If intPre <> Index And picItem(intPre).BackColor = conCOLOR_BULELIGHT And mbytSelect <> intPre Then
        picItem(intPre).BackColor = picNavigate.BackColor
    End If
    picItem(Index).BackColor = conCOLOR_BULELIGHT
    picNavigate.Tag = Index
End Sub

Private Sub picLeft_Resize()
    On Error Resume Next
    '
    picLeftCenter.Move 0, 240, picLeft.ScaleWidth, picLeft.ScaleHeight - 480
    picLeftCenter.Move 0, 0, picLeft.ScaleWidth, picLeft.ScaleHeight - 0
    picUp.Move (picLeft.ScaleWidth - picUp.Width) / 2, 0
    picDown.Move (picLeft.ScaleWidth - picDown.Width) / 2, picLeft.ScaleHeight - 240
    picLeft.BackColor = picNavigate.BackColor
End Sub

Private Sub picLeftCenter_Resize()
    On Error Resume Next
    picNavigate.Move 0, 0, picLeftCenter.ScaleWidth, picLeftCenter.ScaleHeight
End Sub

Private Sub picNavigate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intPre As Integer
    If mblnTip Then Exit Sub
    intPre = Val(picNavigate.Tag)
    If picItem(intPre).BackColor = conCOLOR_BULELIGHT Then
        picItem(intPre).BackColor = picNavigate.BackColor
    End If
End Sub

Private Sub picNavigate_Resize()
    Dim lngH As Long
    Dim i As Long
    
    On Error Resume Next
    For i = picItem.LBound To picItem.UBound
        If i = 0 Then
            picItem(i).Move 120, 60, picLeft.ScaleWidth - 240, 360
            lngH = lngH + picItem(i).Height
        Else
            picItem(i).Move 120, picItem(i - 1).Top + picItem(i - 1).Height + 60, picLeft.ScaleWidth - 240, 360
            lngH = lngH + picItem(i).Height + 60
        End If
        lblTitle(i).Move 120, 75
    Next
    lngH = lngH + 60
    If lngH > picNavigate.Height Then picNavigate.Height = lngH
End Sub

Private Sub picTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mblnMove Then
        mMoveX = mudtPoint.X - mudtRect.Left
        mMoveY = mudtPoint.Y - mudtRect.Top
        mblnMoveStart = True
    End If
End Sub

Private Sub picTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngX As Long, lngY As Long
    
    If mblnMoveStart Then
        lngX = (mudtPoint.X - mMoveX)
        lngY = (mudtPoint.Y - mMoveY)
        Call ShowStyle(1, lngX * Screen.TwipsPerPixelX, lngY * Screen.TwipsPerPixelY)
    End If
    If Me.MousePointer = 99 Then
        Me.MousePointer = 0
    End If
End Sub

Private Sub picTop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call GetWindowRect(Me.hWnd, mudtRect)
    Call GetWindowRect(picClosed.hWnd, mudtRectClose)
    Call GetWindowRect(picNavigate.hWnd, mudtRectNavi)
    mblnMoveStart = False
End Sub

Private Sub picTop_Resize()
    On Error Resume Next
    lblFrmName.Move 120, picTop.ScaleHeight / 2 - lblFrmName.Height / 2
    picClosed.Move picTop.ScaleWidth - picTop.Height, 0, picTop.Height, picTop.Height
    If Not mblnTip And mbytMode = 0 Then
        picFind.Move picTop.ScaleWidth / 2 - 1600, picTop.Height / 2 - 165, 3200, 330
        lblTip.Move picFind.Left + picFind.Width + 120, picTop.Height / 2 - lblTip.Height / 2
        Call picFind_Resize
    End If
End Sub

Private Sub picMain_Resize()
    On Error Resume Next
    vsInfo.Move 120, 120, picMain.ScaleWidth - 180, picMain.ScaleHeight - 240
End Sub

Private Sub picUp_Click()
    picNavigate.Top = 0
End Sub

Private Sub tmrShow_Timer()
     
    If Me.Top > Screen.Height - Me.Height - 120 Then
        Me.Top = Me.Top - 60
    Else
        tmrShow.Enabled = False
    End If
     
End Sub

Private Sub tmrTime_Timer()
    Dim lngRet As Long
    If tmrTime.Tag = "" Then
        Call GetWindowRect(Me.hWnd, mudtRect)
        Call GetWindowRect(picClosed.hWnd, mudtRectClose)
        Call GetWindowRect(picNavigate.hWnd, mudtRectNavi)
        tmrTime.Tag = "1" '首次记录窗体位置
    End If
    lngRet = GetCursorPos(mudtPoint)
    '判断鼠标指针是否位于窗体拖动区
    If PtInRect(mudtRect, mudtPoint.X, mudtPoint.Y) Then
       mblnMove = True
    Else
       mblnMove = False
    End If
    If PtInRect(mudtRectClose, mudtPoint.X, mudtPoint.Y) Then
        picClosed.BackColor = "&H" & Hex(RGB(212, 64, 39))  '红色
    Else
        picClosed.BackColor = picTop.BackColor
    End If

    If PtInRect(mudtRectNavi, mudtPoint.X, mudtPoint.Y) = 0 And Val(picNavigate.Tag) >= 0 And Val(picNavigate.Tag) <> mbytSelect Then
        If picItem(Val(picNavigate.Tag)).BackColor = conCOLOR_BULELIGHT Then
            picItem(Val(picNavigate.Tag)).BackColor = picNavigate.BackColor
        End If
    End If
    '
    If picItem(mbytSelect).BackColor = picNavigate.BackColor Then picItem(mbytSelect).BackColor = conCOLOR_BULELIGHT

    If mblnTip Then
        If Timer - mIntBeginTiem > 4 And mblnKeepTip = False Then Unload Me
        If PtInRect(mudtRect, mudtPoint.X, mudtPoint.Y) Then
            mblnKeepTip = True
        End If
    Else
        If Val(lblTip.Tag) > 0 Then
            If Timer - Val(lblTip.Tag) > 5 Then
                lblTip.Caption = ""
                lblTip.Tag = 0
            End If
        End If
    End If
End Sub

Public Sub ShowStyle(ByVal bytFunc As Byte, Optional ByVal lngLeft As Long, Optional ByVal lngTop As Long)
'功能:根据主窗体位置,决定药品说明书显示方式
    Dim objPoint As RECT
    Dim lngSplit As Long
    lngSplit = 60
    If bytFunc = 0 Then
        '从主窗体进入
        If mblnTip Then
            Me.Left = Screen.Width - Me.Width
            Me.Top = Screen.Height
        Else
            If mbytMode = 0 Then
                Call GetWindowRect(mfrmParent.hWnd, objPoint)
                lngTop = objPoint.Top * Screen.TwipsPerPixelY
                lngLeft = objPoint.Left * Screen.TwipsPerPixelX
                If lngTop + mfrmParent.Height + Me.Height < Screen.Height And lngLeft + Me.Width < Screen.Width Then
                    mbytShowStyle = 0  '左下显示
                    Me.Top = lngTop + mfrmParent.Height + lngSplit
                    Me.Left = lngLeft
                ElseIf lngTop + mfrmParent.Height + Me.Height < Screen.Height And lngLeft + Me.Width > Screen.Width Then
                    mbytShowStyle = 1  '右下显示
                    Me.Top = lngTop + mfrmParent.Height + lngSplit
                    Me.Left = lngLeft - Me.Width + mfrmParent.Width
                ElseIf lngTop - Me.Height > 0 And lngLeft + Me.Width < Screen.Width Then
                    mbytShowStyle = 2  '左上显示
                    Me.Top = lngTop - Me.Height - lngSplit
                    Me.Left = lngLeft
                ElseIf lngTop - Me.Height > 0 And lngLeft + Me.Width > Screen.Width Then
                    mbytShowStyle = 3  '右上显示
                    Me.Top = lngTop - Me.Height - lngSplit
                    Me.Left = lngLeft - Me.Width + mfrmParent.Width
                ElseIf lngTop + Me.Height + mfrmParent.Height > Screen.Height Then
                    lngTop = Screen.Height - Me.Height - mfrmParent.Height
                    mfrmParent.Top = lngTop '悬浮窗体高度不够向上移动
                    If lngLeft + Me.Width < Screen.Width Then
                        mbytShowStyle = 0  '左下显示
                        Me.Top = lngTop + mfrmParent.Height + lngSplit
                        Me.Left = lngLeft
                    Else
                        mbytShowStyle = 1  '右下显示
                        Me.Top = lngTop + mfrmParent.Height + lngSplit
                        Me.Left = lngLeft - Me.Width + mfrmParent.Width
                    End If
                End If
            Else
                Me.Top = mfrmParent.Top + (mfrmParent.Height - Me.Height) / 2
                Me.Left = mfrmParent.Left + (mfrmParent.Width - Me.Width) / 2
            End If
        End If
    ElseIf bytFunc = 1 Then
        '拖动药品说明书、
        If mbytMode = 1 Or mblnTip Then
            If lngLeft < 0 Then lngLeft = 0
            If lngLeft + Me.Width > Screen.Width Then lngLeft = Screen.Width - Me.Width
            If lngTop < 0 Then lngTop = 0
            If lngTop + Me.Height > Screen.Height Then lngTop = Screen.Height - Me.Height
            Me.Left = lngLeft
            Me.Top = lngTop
        Else
            If lngLeft < 0 Then lngLeft = 0
            If lngLeft + Me.Width > Screen.Width Then lngLeft = Screen.Width - Me.Width
            If mbytShowStyle = 0 Or mbytShowStyle = 1 Then
                If lngTop < gobjFrm.Height + lngSplit Then lngTop = gobjFrm.Height + lngSplit
                If lngTop + Me.Height > Screen.Height Then lngTop = Screen.Height - Me.Height
            ElseIf mbytShowStyle = 2 Or mbytShowStyle = 3 Then
                If lngTop <= 0 Then lngTop = 0
                If lngTop + Me.Height + gobjFrm.Height + lngSplit >= Screen.Height Then lngTop = Screen.Height - Me.Height - gobjFrm.Height - lngSplit
            End If
            Me.Left = lngLeft
            Me.Top = lngTop
            If mbytShowStyle = 0 Then
                '左下显示
                gobjFrm.Top = lngTop - gobjFrm.Height - lngSplit
                gobjFrm.Left = lngLeft
            ElseIf mbytShowStyle = 1 Then
                '右下显示
                gobjFrm.Top = lngTop - gobjFrm.Height - lngSplit
                gobjFrm.Left = lngLeft + Me.Width - gobjFrm.Width
            ElseIf mbytShowStyle = 2 Then
                '左上显示
                gobjFrm.Top = Me.Top + Me.Height + lngSplit
                gobjFrm.Left = lngLeft
            ElseIf mbytShowStyle = 3 Then
                '右上显示
                gobjFrm.Top = Me.Top + Me.Height + lngSplit
                gobjFrm.Left = lngLeft + Me.Width - gobjFrm.Width
            End If
        End If
    ElseIf bytFunc = 2 Then
        '从主窗体拖拽
        If mbytShowStyle = 0 Then
            '靠左下显示
            If lngTop + gobjFrm.Height + Me.Height > Screen.Height Then
                lngTop = Screen.Height - Me.Height - gobjFrm.Height
            End If
            If lngLeft + Me.Width > Screen.Width Then
                lngLeft = Screen.Width - Me.Width
            End If
            Me.Top = lngTop + gobjFrm.Height + lngSplit
            Me.Left = lngLeft
        ElseIf mbytShowStyle = 1 Then
            '靠右下显示
            If lngTop + gobjFrm.Height + Me.Height > Screen.Height Then
                lngTop = Screen.Height - Me.Height - gobjFrm.Height
            End If
            If lngLeft < Me.Width Then
                lngLeft = Me.Width - gobjFrm.Width
            End If
            Me.Top = lngTop + gobjFrm.Height + lngSplit
            Me.Left = lngLeft - Me.Width + gobjFrm.Width
        ElseIf mbytShowStyle = 2 Then
            '左上显示
            If lngTop < Me.Height Then
                lngTop = Me.Height
            End If
            If lngLeft + Me.Width > Screen.Width Then
                lngLeft = Screen.Width - Me.Width
            End If
            Me.Top = lngTop - Me.Height - lngSplit
            Me.Left = lngLeft
        ElseIf mbytShowStyle = 3 Then
            '右上显示
            If lngTop < Me.Height Then
                lngTop = Me.Height
            End If
            If lngLeft < Me.Width Then
                lngLeft = Me.Width - gobjFrm.Width
            End If
            Me.Top = lngTop - Me.Height - lngSplit
            Me.Left = lngLeft - Me.Width + gobjFrm.Width
        End If
        gobjFrm.Left = lngLeft
        gobjFrm.Top = lngTop
        Call GetWindowRect(Me.hWnd, mudtRect)
        Call GetWindowRect(picClosed.hWnd, mudtRectClose)
        Call GetWindowRect(picNavigate.hWnd, mudtRectNavi)
    End If
End Sub

Private Sub txtFind_GotFocus()
    picFind.BackColor = vbMagenta       '&HF75000
    txtFind.BackColor = vbWhite
    fraText.BackColor = vbWhite
    txtFind.ForeColor = vbBlack
    If Trim(txtFind.Text) = "输入药品名称\编码\简码" Then txtFind.Text = ""
    zlControl.TxtSelAll txtFind
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If Val(lblTip.Tag) <> 0 Then
        lblTip.Tag = ""
        lblTip.Caption = ""
    End If
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call SerachDrug
    End If
End Sub


Private Sub txtFind_LostFocus()
    picFind.BackColor = vbWhite
    fraText.BackColor = picNavigate.BackColor
    txtFind.BackColor = picNavigate.BackColor
    txtFind.ForeColor = &HC0C0C0
    If txtFind.Text = "" Then txtFind.Text = "输入药品名称\编码\简码"
End Sub

Private Sub txtFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.MousePointer = 99 Then Me.MousePointer = 0
End Sub

Private Sub SerachDrug()
    Dim vPoint As POINTAPI
    Dim strSQL As String
    Dim strInput As String
    Dim strDrug As String
    Dim rsTemp As ADODB.Recordset
    Dim blnCancel As Boolean
    
     
    On Error GoTo errH
    strInput = Trim(txtFind.Text)
    If strInput = "" Then Exit Sub
    strSQL = " And (A.编码 Like [1] And C.码类=[3]" & _
            " Or C.名称 Like [2] And C.码类=[3] Or C.简码 Like [2] And C.码类 IN([3],3))"
    If IsNumeric(strInput) Then
        '1X.输入全是数字时只匹配编码'对于药品,则要匹配简码(码类为3的数字码)
        If Mid(gstrMatchMode, 1, 1) = "1" Then strSQL = " And (A.编码 Like [1] And C.码类=[3] Or C.简码 Like [2] And C.码类=3)"
    ElseIf zlCommFun.IsCharAlpha(strInput) Then
        'X1.输入全是字母时只匹配简码
        If Mid(gstrMatchMode, 2, 1) = "1" Then strSQL = " And C.简码 Like [2] And C.码类=[3]"
    ElseIf zlCommFun.IsCharChinese(strInput) Then
        '包含汉字,则只匹配名称
        strSQL = " And C.名称 Like [2] And C.码类=[3]"
    End If
    strSQL = "Select  ID,本位码,类别,编码,名称,规格,产地 " & vbNewLine & _
        "From (Select distinct B.药品ID As ID,B.本位码,Decode(a.类别, '5', '西成药', '6', '中成药', '中草药') As 类别,a.类别 AS 排序,a.编码, c.名称, a.规格, a.产地" & vbNewLine & _
        "       From 药品规格 B, 收费项目别名 C, 收费项目目录 A" & vbNewLine & _
        "       Where a.Id = b.药品id And a.Id = c.收费细目id And a.类别 In ('5', '6', '7') And c.性质 = 1 " & strSQL & vbNewLine & _
        "           And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And (a.站点 = '" & gstrNodeNo & "' Or a.站点 Is Null))" & vbNewLine & _
        "   Order By 排序"
    vPoint = zlControl.GetCoordPos(Me.hWnd, picFind.Left, picFind.Top)
    Set rsTemp = zlDatabase.ShowSQLSelect(gobjFrm, strSQL, 0, "药品检索", True, "", "", False, False, True, _
        vPoint.X + 15, vPoint.Y, picFind.Height + 15, blnCancel, False, True, UCase(strInput) & "%", mstrLike & UCase(strInput) & "%", mint简码 + 1, "ColSet:列宽设置|本位码,0")
    If rsTemp Is Nothing Then
        lblTip.Caption = "未找到匹配的药品数据。"
        lblTip.Tag = Timer
        txtFind.SetFocus
        Exit Sub
    Else
        txtFind.Text = rsTemp!名称 & ""
        strDrug = rsTemp!名称 & ""
        If rsTemp!本位码 & "" = "" Then
            lblTip.Caption = "没有找到对应的说明书。"
            lblTip.Tag = Timer
            Call txtFind_GotFocus
            Exit Sub
        End If
        If Not GetDrugInfo_ZL(rsTemp!本位码 & "", strDrug) Then Exit Sub
        gobjFrm.lblDrug = rsTemp!名称 & ""
        gobjFrm.mstrDrugCode = rsTemp!本位码 & ""
        Call ShowMe(gobjFrm, strDrug, 0)
        vsInfo.SetFocus
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


