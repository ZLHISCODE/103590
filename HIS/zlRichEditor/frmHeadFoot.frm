VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmHeadFoot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "页眉页脚"
   ClientHeight    =   6285
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   11625
   Icon            =   "frmHeadFoot.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   10485
      TabIndex        =   9
      Top             =   3525
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   10485
      TabIndex        =   10
      Top             =   3195
      Width           =   1100
   End
   Begin VB.CommandButton cmdPicture 
      Caption         =   "清除图片(&E)"
      Height          =   350
      Index           =   1
      Left            =   10320
      TabIndex        =   11
      Top             =   330
      Width           =   1275
   End
   Begin VB.CheckBox chkPic 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   3300
      Picture         =   "frmHeadFoot.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "插入图片(Alt+D)"
      Top             =   3525
      Width           =   345
   End
   Begin VB.CheckBox chkR 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2970
      Picture         =   "frmHeadFoot.frx":685E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "右对齐(Alt+R)"
      Top             =   3525
      Width           =   345
   End
   Begin VB.CheckBox chkM 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2640
      Picture         =   "frmHeadFoot.frx":D0B0
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "居中对齐(Alt+M)"
      Top             =   3525
      Width           =   345
   End
   Begin VB.CheckBox chkL 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2310
      Picture         =   "frmHeadFoot.frx":13902
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "左对齐(Alt+L)"
      Top             =   3525
      Width           =   345
   End
   Begin VB.CheckBox chkColor 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1980
      Picture         =   "frmHeadFoot.frx":1A154
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "指定选中文字的颜色(Alt+A)"
      Top             =   3525
      Width           =   345
   End
   Begin VB.CheckBox chkI 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1650
      Picture         =   "frmHeadFoot.frx":1A466
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "斜体(Alt+I)"
      Top             =   3525
      Width           =   345
   End
   Begin zlRichEditor.Editor edtThis 
      Height          =   225
      Left            =   420
      TabIndex        =   24
      Top             =   3345
      Visible         =   0   'False
      Width           =   210
      _extentx        =   370
      _extenty        =   397
      withviewbuttonas=   0
      showruler       =   0
   End
   Begin VB.ComboBox cboMode 
      Height          =   300
      ItemData        =   "frmHeadFoot.frx":20CB8
      Left            =   6495
      List            =   "frmHeadFoot.frx":20CBA
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3375
      Width           =   2715
   End
   Begin VB.CommandButton cmdFoot 
      Caption         =   "加入页脚(&F)"
      Height          =   350
      Left            =   9225
      TabIndex        =   15
      Top             =   3525
      Width           =   1275
   End
   Begin VB.PictureBox picPicture 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   15
      ScaleHeight     =   675
      ScaleWidth      =   10215
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   15
      Width           =   10215
      Begin VB.Image imgPicture 
         Appearance      =   0  'Flat
         Height          =   570
         Left            =   -30
         Top             =   0
         Width           =   4395
      End
      Begin VB.Label lblPicture 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "尺寸(像素):"
         Height          =   180
         Left            =   6135
         TabIndex        =   14
         Top             =   60
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdPicture 
      Caption         =   "选择图片(&P)"
      Height          =   350
      Index           =   0
      Left            =   10320
      TabIndex        =   12
      Top             =   0
      Width           =   1275
   End
   Begin VB.ComboBox cboFont 
      Height          =   300
      Left            =   990
      Style           =   2  'Dropdown List
      TabIndex        =   8
      ToolTipText     =   "字体(Ctrl+F)"
      Top             =   3225
      Width           =   1905
   End
   Begin VB.ComboBox cboFSize 
      Height          =   300
      Left            =   2895
      Style           =   2  'Dropdown List
      TabIndex        =   7
      ToolTipText     =   "字体(Ctrl+S)"
      Top             =   3220
      Width           =   750
   End
   Begin VB.CheckBox chkU 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1320
      Picture         =   "frmHeadFoot.frx":20CBC
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "下划线(Alt+U)"
      Top             =   3525
      Width           =   345
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin zlRichEditor.ColorPicker ColorPicker1 
      Height          =   2190
      Left            =   2130
      TabIndex        =   0
      Top             =   3870
      Visible         =   0   'False
      Width           =   2190
      _extentx        =   3863
      _extenty        =   3863
   End
   Begin zlRichEditor.Document DocHead 
      Height          =   2415
      Left            =   0
      TabIndex        =   18
      Top             =   720
      Width           =   11565
      _extentx        =   20399
      _extenty        =   4260
      margintop       =   850
      marginbottom    =   850
      marginleft      =   850
      marginright     =   850
   End
   Begin zlRichEditor.Document DocFoot 
      Height          =   2265
      Left            =   0
      TabIndex        =   19
      Top             =   3975
      Width           =   11565
      _extentx        =   20399
      _extenty        =   3995
   End
   Begin VB.CheckBox chkB 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   990
      Picture         =   "frmHeadFoot.frx":2750E
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "粗体(Alt+B)"
      Top             =   3525
      Width           =   345
   End
   Begin VB.CommandButton cmdHead 
      Caption         =   "加入页眉(&H)"
      Height          =   350
      Left            =   9225
      TabIndex        =   16
      Top             =   3195
      Width           =   1275
   End
   Begin VB.Label lblHead 
      AutoSize        =   -1  'True
      Caption         =   "<页眉>"
      Height          =   180
      Left            =   75
      TabIndex        =   23
      Top             =   3165
      Width           =   540
   End
   Begin VB.Label lblFoot 
      AutoSize        =   -1  'True
      Caption         =   "<页脚>"
      Height          =   180
      Left            =   75
      TabIndex        =   22
      Top             =   3720
      Width           =   540
   End
   Begin VB.Label lblGut 
      AutoSize        =   -1  'True
      Caption         =   "页眉页脚中可包含“{  }”形态的自动替换项目，实际输出时，将被转换为对应的实际内容。(&A)"
      Height          =   540
      Left            =   3795
      TabIndex        =   21
      Top             =   3255
      Width           =   2670
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmHeadFoot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mlngPaperWidth As Long, mlngPaperHeight As Long

Dim mblnOK As Boolean
Private Sub cboFont_Click()
    Call SetFontStyle("F_" & cboFont.List(cboFont.ListIndex))
End Sub
Private Sub cboFSize_Click()
    Call SetFontStyle("S_" & cboFSize.List(cboFSize.ListIndex))
End Sub



Private Sub chkB_Click()
    If chkB.Value = vbChecked Then
        Call SetFontStyle("B")
    Else
        Call SetFontStyle("UB")
    End If
End Sub

Private Sub chkColor_Click()
    If chkColor.Value = vbUnchecked Then Exit Sub
    
    ColorPicker1.Visible = True
    
End Sub

Private Sub chkI_Click()
    If chkI.Value = vbChecked Then
        Call SetFontStyle("I")
    Else
        Call SetFontStyle("UI")
    End If
End Sub

Private Sub chkL_Click()
    If chkL.Value = vbChecked Then
        Call SetFontStyle("L")
        chkM.Value = vbUnchecked
        chkR.Value = vbUnchecked
    End If
End Sub

Private Sub chkM_Click()
    If chkM.Value = vbChecked Then
        Call SetFontStyle("M")
        chkL.Value = vbUnchecked
        chkR.Value = vbUnchecked
    End If
End Sub

Private Sub chkPic_Click()
    If chkPic.Value = vbChecked Then
        With Me.dlgThis
            .DialogTitle = "插入图片"
            .FileName = ""
            .Filter = "图像|*.jpg;*.bmp;*.ico;*.gif"
            .CancelError = True
            Err = 0: On Error Resume Next
            .ShowOpen
            If Err.Number <> 0 Then chkPic.Value = vbUnchecked: Err.Clear: Exit Sub
        End With
         
        Dim picTmp As StdPicture, lngWidth As Long, lngHeight As Long
        Set picTmp = Nothing
        Set picTmp = LoadPicture(Me.dlgThis.FileName)
        If picTmp Is Nothing Then MsgBox "不是有效的图片文件！", vbExclamation, Me.Caption: chkPic.Value = vbUnchecked: Exit Sub
        'lngWidth = CLng(Me.ScaleX(DocHead.PaperWidth - DocHead.MarginLeft - DocHead.MarginRight, vbTwips, vbPixels))
        lngWidth = 200
        lngHeight = 50
        If Int(Me.ScaleX(picTmp.Width, vbHimetric, vbPixels)) > lngWidth Then
            MsgBox "图片宽度不能超过 " & lngWidth & " 像素，请检查。", vbInformation, Me.Caption
            chkPic.Value = vbUnchecked: Exit Sub
        End If
        If Int(Me.ScaleY(picTmp.Height, vbHimetric, vbPixels)) > lngHeight Then
            MsgBox "图片高度不能超过 " & lngHeight & " 像素，请检查。", vbInformation, Me.Caption
            chkPic.Value = vbUnchecked: Exit Sub
        End If
        
        edtThis.NewDoc
        DoEvents
        Clipboard.Clear
        Clipboard.SetData picTmp
        edtThis.PasteWithFormat
        edtThis.SelectAll
        edtThis.CopyWithFormat
        If Me.Tag = "Head" Then
            DocHead.PasteWithFormat
            DocHead.Range(DocHead.Selection.EndPos - 2, DocHead.Selection.EndPos).Text = ""
        Else
            DocFoot.PasteWithFormat
            DocFoot.Range(DocFoot.Selection.EndPos - 2, DocFoot.Selection.EndPos).Text = ""
        End If
        chkPic.Value = vbUnchecked
    End If
End Sub

Private Sub chkR_Click()
    If chkR.Value = vbChecked Then
        Call SetFontStyle("R")
        chkL.Value = vbUnchecked
        chkM.Value = vbUnchecked
    End If
End Sub

Private Sub chkU_Click()
    If chkU.Value = vbChecked Then
        Call SetFontStyle("U")
    Else
        Call SetFontStyle("UU")
    End If
End Sub
Private Sub SetFontStyle(ByVal strStyle As String)
Dim lngS As Long, lngE As Long
    If Not Me.Visible Then Exit Sub
    If Not (Me.Tag = "Head" Or Me.Tag = "Foot") Then Exit Sub
    Select Case strStyle
        Case "B"
            If Me.Tag = "Head" Then
                DocHead.Selection.Font.Bold = True
            Else
                DocFoot.Selection.Font.Bold = True
            End If
        Case "UB"
            If Me.Tag = "Head" Then
                DocHead.Selection.Font.Bold = False
            Else
                DocFoot.Selection.Font.Bold = False
            End If
        Case "U"
            If Me.Tag = "Head" Then
                DocHead.Selection.Font.Underline = cprHair
            Else
                DocFoot.Selection.Font.Underline = cprHair
            End If
        Case "UU"
            If Me.Tag = "Head" Then
                DocHead.Selection.Font.Underline = cprNone
            Else
                DocFoot.Selection.Font.Underline = cprNone
            End If
        Case "I"
            If Me.Tag = "Head" Then
                DocHead.Selection.Font.Italic = True
            Else
                DocFoot.Selection.Font.Italic = True
            End If
        Case "UI"
            If Me.Tag = "Head" Then
                DocHead.Selection.Font.Italic = False
            Else
                DocFoot.Selection.Font.Italic = False
            End If
        Case "L"
            If Me.Tag = "Head" Then
                DocHead.Selection.Para.Alignment = cprHALeft
            Else
                DocFoot.Selection.Para.Alignment = cprHALeft
            End If
        Case "R"
            If Me.Tag = "Head" Then
                DocHead.Selection.Para.Alignment = cprHARight
            Else
                DocFoot.Selection.Para.Alignment = cprHARight
            End If
        Case "M"
            If Me.Tag = "Head" Then
                DocHead.Selection.Para.Alignment = cprHACenter
            Else
                DocFoot.Selection.Para.Alignment = cprHACenter
            End If
        Case Else
            If Split(strStyle, "_")(0) = "F" Then '设置字体
                If Me.Tag = "Head" Then
                    DocHead.Selection.Font.Name = Split(strStyle, "_")(1)
                Else
                    DocFoot.Selection.Font.Name = Split(strStyle, "_")(1)
                End If
            Else '设置字号
                If Me.Tag = "Head" Then
                    DocHead.Selection.Font.SIZE = GetFontSizeNumber(Split(strStyle, "_")(1))
                Else
                    DocFoot.Selection.Font.SIZE = GetFontSizeNumber(Split(strStyle, "_")(1))
                End If
            End If
    End Select
End Sub
Private Sub cmdCancel_Click()
    mblnOK = False: Me.Hide
End Sub

Private Sub cmdFoot_Click()
Dim lngS As Long, lngE As Long
    With Me.DocFoot
        .ForceEdit = True
        lngE = .Selection.EndPos
        .Range(lngE, lngE).Text = Space(1) & Me.cboMode.Text & Space(1)
        lngE = lngE + Len(Me.cboMode.Text) + 2
        .Range(lngE, Len(Me.cboMode.Text) + 2).Selected
    End With
End Sub

Private Sub cmdHead_Click()
Dim lngE As Long
    With Me.DocHead
        .ForceEdit = True
        lngE = .Selection.EndPos
        .Range(lngE, lngE).Text = Space(1) & Me.cboMode.Text & Space(1)
        lngE = lngE + Len(Me.cboMode.Text) + 2
        .Range(lngE, lngE).Selected
    End With
End Sub

Private Sub cmdOK_Click()
    If Me.ScaleX(Me.imgPicture.Picture.Width, vbHimetric, vbTwips) > mlngPaperWidth Then
        MsgBox "页眉图片太宽！", vbExclamation, Me.Caption: Exit Sub
    End If
    If Me.ScaleX(Me.imgPicture.Picture.Height, vbHimetric, vbTwips) > mlngPaperHeight / 3 Then
        MsgBox "页眉图片太高！", vbExclamation, Me.Caption: Exit Sub
    End If
    mblnOK = True: Me.Hide
End Sub

Public Function ShowMe(Editor As Editor) As Boolean
'功能：显示本对话框
'参数：
'   Editor,编辑器对象
Dim i As Integer, sFont As String, lngWidth As Long
On Error Resume Next
    With Me.cboMode
        .AddItem "第{页码}页"
        .AddItem "第{页码}页，共{总页数}页"
        .AddItem "标题：{标题}"
        .AddItem "文件：{文件名}"
        .AddItem "文件：{路径}{文件名}"
        .AddItem "打印日期：{打印日期}"
        .AddItem "打印时间：{打印时间}"
        .AddItem "书写：{书写部门} {书写签名} {完成时间}"
        .AddItem "{医生签名} {主治签名} {主任签名}"
        .AddItem "{单位名称}{病历名称}"
        .AddItem "时间：{当前日期} {当前时间}"
        .AddItem "姓名：{姓名} 性别：{性别} 年龄：{年龄}"
        .AddItem "标识号：{标识号}"
        .AddItem "门诊号：{门诊号}"
        .AddItem "住院号：{住院号}"
        .AddItem "科室：{就诊科室}"
        .AddItem "科室：{入院科室}"
        .AddItem "病区：{入院病区}"
        .AddItem "科室：{当前科室}"
        .AddItem "床号：{当前床号}"
        .AddItem "住院日期：{入院日期}～{出院日期}"
        .AddItem "第{住院次数}住院"
        .AddItem "经治医师：{住院医师}"
        .AddItem "责任护士：{责任护士}"
        .ListIndex = 0
    End With
    SendMessage cboMode.Hwnd, CB_SETDROPPEDWIDTH, 300, 0
    SendMessage cboFont.Hwnd, CB_SETDROPPEDWIDTH, 250, 0
    
    For i = 0 To Screen.FontCount - 1
       sFont = Screen.Fonts(i)
       cboFont.AddItem sFont
       If sFont = "宋体" Then cboFont.ListIndex = i
    Next i
    With cboFSize
        .AddItem "初号"
        .AddItem "小初"
        .AddItem "一号"
        .AddItem "小一"
        .AddItem "二号"
        .AddItem "小二"
        .AddItem "三号"
        .AddItem "小三"
        .AddItem "四号"
        .AddItem "小四"
        .AddItem "五号"
        .AddItem "小五"
        .AddItem "六号"
        .AddItem "小六"
        .AddItem "七号"
        .AddItem "八号"
        .AddItem 5
        .AddItem 5.5
        .AddItem 6.5
        .AddItem 7.5
        .AddItem 8
        .AddItem 9
        .AddItem 10
        .AddItem 10.5
        .AddItem 11
        .AddItem 12
        .AddItem 14
        .AddItem 16
        .AddItem 18
        .AddItem 20
        .AddItem 22
        .AddItem 24
        .AddItem 26
        .AddItem 28
        .AddItem 36
        .AddItem 48
        .AddItem 72
        .ListIndex = 12
    End With
    
    With Editor
        DocHead.PaperWidth = .PaperWidth
        DocHead.MarginLeft = .MarginLeft
        DocHead.MarginRight = .MarginRight
        DocHead.ResetWYSIWYG
        DocHead.ForceEdit = True
        If .HeadFileText = "" Then '没有文件内容，将TXT文字读入Doc中，兼容原有文字型页眉页脚
            .HeadTextToFile
        End If
        DocHead.TextRTF = .HeadFileTextRTF
        
        DocFoot.PaperWidth = .PaperWidth
        DocFoot.MarginLeft = .MarginLeft
        DocFoot.MarginRight = .MarginRight
        DocFoot.ResetWYSIWYG
        DocFoot.ForceEdit = True
        If .FootFileText = "" Then
            .FootTextToFile '没有文件内容，将TXT文字读入Doc中，兼容原有文字型页眉页脚
        End If
        DocFoot.TextRTF = .FootFileTextRTF
        
        lngWidth = .PaperWidth - .MarginLeft - .MarginRight
        If lngWidth > Me.Width - 100 Then lngWidth = Me.Width - 100
    End With
    
    If Not (Editor.Picture Is Nothing) Then
        If Editor.Picture.Handle <> 0 Then
            Set Me.imgPicture.Picture = Editor.Picture
            Me.lblPicture.Caption = "尺寸(像素):" & Int(Me.ScaleX(Me.imgPicture.Picture.Width, vbHimetric, vbPixels)) & _
                                    "×" & Int(Me.ScaleY(Me.imgPicture.Picture.Height, vbHimetric, vbPixels))
        End If
    End If
    
    mlngPaperWidth = Editor.PaperWidth
    mlngPaperHeight = Editor.PaperHeight
    
    mblnOK = False
    Me.Show vbModal
    If mblnOK = False Then Unload Me: Exit Function '取消退出
    
    DocHead.ClearEndCrlfChar
    DocHead.SelectAll
    DocHead.CopyWithFormat
    Editor.DocHeadPasteWithFormat
    Editor.Head = DocHead.Text

    DocFoot.ClearEndCrlfChar
    DocFoot.SelectAll
    DocFoot.CopyWithFormat
    Editor.DocFootPasteWithFormat
    Editor.Foot = DocFoot.Text
    
    Set Editor.Picture = Me.imgPicture.Picture
    ShowMe = True: Unload Me
End Function

Private Sub cmdPicture_Click(Index As Integer)
    Dim picTemp As StdPicture
    If Index = 1 Then
        Set Me.imgPicture.Picture = Nothing
        Me.lblPicture.Caption = "尺寸(像素):"
    Else
        With Me.dlgThis
            .DialogTitle = "标志图选择"
            .FileName = ""
            .Filter = "图像|*.jpg;*.bmp;*.ico;*.gif"
            .CancelError = True
            Err = 0: On Error Resume Next
            .ShowOpen
            If Err.Number <> 0 Then Err.Clear: Exit Sub
        End With
        Set picTemp = Nothing
        Set picTemp = LoadPicture(Me.dlgThis.FileName)
        If picTemp Is Nothing Then MsgBox "不是有效的图片文件！", vbExclamation, Me.Caption: Exit Sub
        Set Me.imgPicture.Picture = picTemp
        Me.lblPicture.Caption = "尺寸(像素):" & Int(Me.ScaleX(Me.imgPicture.Picture.Width, vbHimetric, vbPixels)) & _
                                "×" & Int(Me.ScaleY(Me.imgPicture.Picture.Height, vbHimetric, vbPixels))
    End If
End Sub

Private Sub ColorPicker1_pCancel()
    chkColor.Value = vbUnchecked
    
End Sub

Private Sub ColorPicker1_pOK()
Dim lngColor As Long
    On Error Resume Next
    chkColor.Value = vbUnchecked
    lngColor = ColorPicker1.COLOR
    
    If Me.Tag = "Head" Then
        DocHead.Selection.Font.ForeColor = lngColor
        DocHead.Range(DocHead.Selection.EndPos, DocHead.Selection.EndPos).Selected
    Else
        DocFoot.Selection.Font.ForeColor = lngColor
        DocFoot.Range(DocFoot.Selection.EndPos, DocFoot.Selection.EndPos).Selected
    End If
    ColorPicker1.Visible = False
End Sub

Private Sub DocFoot_GotFocus()
    DocFoot.ForceEdit = True
    Me.Tag = "Foot"
End Sub

Private Sub DocFoot_LostFocus()
    DocFoot.ForceEdit = False
End Sub

Private Sub DocFoot_SelChange(ByVal lStart As Long, ByVal lEnd As Long)
    If Not Me.Visible Then Exit Sub
    If Me.ActiveControl.Name = "DocFoot" Then
        On Error Resume Next
        cboFont.Text = DocFoot.Selection.Font.Name
        cboFSize.Text = DocFoot.Selection.Font.SIZE
        chkB.Value = IIf(DocFoot.Selection.Font.Bold, vbChecked, vbUnchecked)
        chkU.Value = IIf(DocFoot.Selection.Font.Underline, vbChecked, vbUnchecked)
        chkI.Value = IIf(DocFoot.Selection.Font.Italic, vbChecked, vbUnchecked)
        chkL.Value = IIf(DocFoot.Selection.Para.Alignment = cprHALeft, vbChecked, vbUnchecked)
        chkM.Value = IIf(DocFoot.Selection.Para.Alignment = cprHACenter, vbChecked, vbUnchecked)
        chkR.Value = IIf(DocFoot.Selection.Para.Alignment = cprHARight, vbChecked, vbUnchecked)
        Err.Clear
    End If
End Sub

Private Sub DocHead_GotFocus()
    DocHead.ForceEdit = True
    Me.Tag = "Head"
End Sub

Private Sub DocHead_LostFocus()
    DocHead.ForceEdit = False
End Sub

Private Sub DocHead_SelChange(ByVal lStart As Long, ByVal lEnd As Long)
    If Not Me.Visible Then Exit Sub
    If Me.ActiveControl.Name = "DocHead" Then
        On Error Resume Next
        cboFont.Text = DocHead.Selection.Font.Name
        cboFSize.Text = DocHead.Selection.Font.SIZE
        chkB.Value = IIf(DocHead.Selection.Font.Bold, vbChecked, vbUnchecked)
        chkU.Value = IIf(DocHead.Selection.Font.Underline, vbChecked, vbUnchecked)
        chkI.Value = IIf(DocHead.Selection.Font.Italic, vbChecked, vbUnchecked)
        chkL.Value = IIf(DocHead.Selection.Para.Alignment = cprHALeft, vbChecked, vbUnchecked)
        chkM.Value = IIf(DocHead.Selection.Para.Alignment = cprHACenter, vbChecked, vbUnchecked)
        chkR.Value = IIf(DocHead.Selection.Para.Alignment = cprHARight, vbChecked, vbUnchecked)
        Err.Clear
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Or Shift = 4 Then
        Select Case KeyCode
            Case vbKeyB And Shift = 4
                chkB.Value = IIf(chkB.Value = vbChecked, vbUnchecked, vbChecked)
            Case vbKeyU And Shift = 4
                chkU.Value = IIf(chkU.Value = vbChecked, vbUnchecked, vbChecked)
            Case vbKeyI And Shift = 4
                chkI.Value = IIf(chkI.Value = vbChecked, vbUnchecked, vbChecked)
            Case vbKeyL And Shift = 4
                chkL.Value = vbChecked
            Case vbKeyM And Shift = 4
                chkM.Value = vbChecked
            Case vbKeyR And Shift = 4
                chkR.Value = vbChecked
            Case vbKeyA And Shift = 4
                chkColor.Value = vbChecked
            Case vbKeyD And Shift = 4
                chkPic.Value = vbChecked
            Case vbKeyF And Shift = 2
                SendMessage cboFont.Hwnd, CB_SHOWDROPDOWN, True, 0
            Case vbKeyS And Shift = 2
                SendMessage cboFSize.Hwnd, CB_SHOWDROPDOWN, True, 0
        End Select
    End If
End Sub

Private Sub imgPicture_DblClick()
    With Me.imgPicture
        .Stretch = Not .Stretch
        If .Stretch Then .Move 0, 0, Me.picPicture.ScaleWidth, Me.picPicture.ScaleHeight
    End With
End Sub
Private Function GetFontSizeNumber(ByVal strSize As String) As Single
    Dim sngNum As Single
    Select Case strSize
    Case "初号"
        sngNum = 42
    Case "小初"
        sngNum = 36
    Case "一号"
        sngNum = 26
    Case "小一"
        sngNum = 24
    Case "二号"
        sngNum = 22
    Case "小二"
        sngNum = 18
    Case "三号"
        sngNum = 16
    Case "小三"
        sngNum = 15
    Case "四号"
        sngNum = 14
    Case "小四"
        sngNum = 12
    Case "五号"
        sngNum = 10.5
    Case "小五"
        sngNum = 9
    Case "六号"
        sngNum = 7.5
    Case "小六"
        sngNum = 6.5
    Case "七号"
        sngNum = 5.5
    Case "八号"
        sngNum = 5
    Case Else
        sngNum = IIf(Val(strSize) <= 0, 10, Val(strSize))
    End Select
    GetFontSizeNumber = Format(sngNum, "0.0")
End Function
