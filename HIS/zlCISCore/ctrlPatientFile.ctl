VERSION 5.00
Object = "{B0475000-7740-11D1-BDC3-0020AF9F8E6E}#6.1#0"; "TTF16.ocx"
Begin VB.UserControl ctrlPatientFile 
   BackColor       =   &H80000015&
   ClientHeight    =   5760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4425
   KeyPreview      =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   4425
   Begin VB.TextBox txtSpecPaper 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   180
      Index           =   0
      Left            =   240
      MaxLength       =   2000
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   0
      ScaleHeight     =   5655
      ScaleWidth      =   3735
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3735
      Begin zl9CISCore.VisItem VisItem 
         Height          =   225
         Index           =   0
         Left            =   1560
         TabIndex        =   18
         Top             =   3840
         Visible         =   0   'False
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   397
         MousePointer    =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         AllowEdit       =   -1  'True
      End
      Begin VB.TextBox txtVisForm 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   180
         Index           =   0
         Left            =   3360
         MaxLength       =   2000
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   5280
         Visible         =   0   'False
         Width           =   90
      End
      Begin zl9CISCore.VisItem SpecItem 
         Height          =   225
         Index           =   0
         Left            =   2280
         TabIndex        =   13
         Top             =   4560
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2328
         _ExtentY        =   397
         MousePointer    =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AllowEdit       =   -1  'True
      End
      Begin VB.HScrollBar HSEdit 
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2520
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.VScrollBar VSEdit 
         Height          =   1215
         Index           =   0
         Left            =   3480
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   255
      End
      Begin zl9CISCore.ctrlVisForm VisForm 
         Height          =   735
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   4560
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   8811
         _ExtentY        =   1296
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
      Begin VB.PictureBox PicFlag 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   3000
         Index           =   0
         Left            =   360
         ScaleHeight     =   200
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   333
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   600
         Visible         =   0   'False
         Width           =   5000
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   420
            Index           =   0
            Left            =   345
            MultiLine       =   -1  'True
            TabIndex        =   19
            Top             =   2355
            Visible         =   0   'False
            Width           =   600
         End
      End
      Begin VB.PictureBox picSplit 
         BackColor       =   &H00E0E0E0&
         Height          =   15
         Index           =   0
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   7.5
         ScaleMode       =   0  'User
         ScaleWidth      =   2775
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.PictureBox picEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1455
         Index           =   0
         Left            =   840
         ScaleHeight     =   1455
         ScaleWidth      =   2535
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox txtBox1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   180
         Index           =   0
         Left            =   1080
         MaxLength       =   2000
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   3600
         Visible         =   0   'False
         Width           =   90
      End
      Begin TTF160Ctl.F1Book grdTable 
         Height          =   1335
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   2880
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   2355
         _0              =   $"ctrlPatientFile.ctx":0000
         _1              =   $"ctrlPatientFile.ctx":0409
         _2              =   $"ctrlPatientFile.ctx":0812
         _3              =   $"ctrlPatientFile.ctx":0C1B
         _4              =   $"ctrlPatientFile.ctx":1024
         _5              =   $"ctrlPatientFile.ctx":142D
         _6              =   $"ctrlPatientFile.ctx":1836
         _7              =   $"ctrlPatientFile.ctx":1C3F
         _8              =   $"ctrlPatientFile.ctx":2048
         _count          =   9
         _ver            =   2
      End
      Begin VB.Label lblVisForm 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Height          =   180
         Index           =   0
         Left            =   2280
         TabIndex        =   14
         Top             =   4560
         Visible         =   0   'False
         Width           =   90
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   11
         Top             =   4200
         Visible         =   0   'False
         Width           =   90
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   690
         TabIndex        =   4
         Top             =   720
         Visible         =   0   'False
         Width           =   165
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblFlag 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Δ"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   0
         Left            =   120
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   2040
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image picMargin 
         Appearance      =   0  'Flat
         Height          =   3135
         Left            =   0
         Picture         =   "ctrlPatientFile.ctx":221E
         Stretch         =   -1  'True
         Top             =   960
         Width           =   375
      End
   End
   Begin VB.VScrollBar VSMain 
      Height          =   3855
      Left            =   4080
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lblSpecPaper 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Height          =   180
      Index           =   0
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   90
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "ctrlPatientFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const LABEL_EXPAND = "Δ" '"━"
Private Const LABEL_COLLAPSE = "▼" ' "╋"
Private Const lnHeightDis = 0 'Label 和 TextBox的高度差
Private Const COLOR_COMBO = &H8000& '有下拉选择的文本颜色
Private MARGIN_PAPER   As Integer

Private PatientFileID As String '病历文件ID
'当PatientFileID为空时（新增病历）,需指定以下参数
Private PatientID As String '病人ID
Private CheckID As String '病案ID或挂号单ID
Private PatientType As Integer '0=门诊病人 1=住院病人
Private FileTypeID As String '病历模板文件ID
Private bSampleFile As Boolean '是否病历示范
Private AdviceID As Long '相关医嘱ID
Private blnMoved As Boolean '当前病人数据是否转出

Private SendAdviceID As Long, SendNO As Long '医嘱发送的医嘱ID和发送号

Private bOnLoadFile As Boolean

Private TitleWidth As Long, TitleHeight As Long '标题宽度
Private CtrlDistance As Integer '控件间隔
Private SplitDistance As Integer

'元素数组
'0-14列为病历文件组成字段
'0:专用纸部件
'15:元素是否作废
'16:元素高度
'17:元素控件索引
'18:元素控件类型
'19:元素宽度
'20:病人病历内容ID，为0表示增加的内容
'21:标记图元素ID
'22:病历元素编码
'23:元素是否修改
Private aElement() As Variant
Private FileHeight As Long '病历页面高度
Private bAllowEdit As Boolean
Private bModified As Boolean, bNotShowDiagItem As Boolean '是否显示检验选择栏目
Private bNotRunSelChange As Boolean

Private aPicFlag() As MapItems '标记图编辑返回值
Private SpecPaper() As VBControlExtender, WithEvents CurrSpecPaper As VBControlExtender
Attribute CurrSpecPaper.VB_VarHelpID = -1

'以下参数用于Rtf控件编辑
Private blnEvent_SelChange() As Boolean
Private blnCurrUnderLine() As Boolean
Private aTextItems() As String

Private blnMouseDown As Boolean

Public Event ElementGotFocus(ByVal ElementIndex As Integer, ByVal ElementType As Integer)
Public Event Resize()

Public Property Get FileID() As String
    FileID = PatientFileID
End Property

Public Sub Reload(Optional objProgressBar As ProgressBar, Optional blnReplaced As Boolean = False)
    '参数：blnReplaced 是否在所见单中强制替换具有替换域属性的所见项
    
    Dim tmpCtrl As VB.Control, CtrlIndex As Integer, CtrlHeight As Long, CtrlTop As Long
    Dim aFont() As String
    Dim i As Long, iNum As Long, Seq As Integer
    Dim rsTmp As New ADODB.Recordset, sTmpFile As String, FileObj As New Scripting.FileSystemObject
    Dim iTabLeft As Long, iTabTop As Long, iTabWidth As Long, iTabHeight As Long, iShown As Integer
    Dim iItemLen As Integer, sItemFormat As String, iItemType As Integer
    Dim bPicEnabled As Boolean
    
    Dim strTxtBox As String
    
    Dim lngInitProgValue As Long '初始进度值
    Dim TmpFont As StdFont, iTmpLines As Integer
    
    Dim sngLine_Indent As Single
    Dim strSQL As String, lngTmpID As Long
    
    bOnLoadFile = True
    blnMouseDown = False
    '卸载所有病历控件
    On Error Resume Next
    For Each tmpCtrl In UserControl.Controls
        If UCase(tmpCtrl.Name) Like "SPECPAPER*" Then
            UserControl.Controls.Remove tmpCtrl.Name
        Else
            Unload tmpCtrl
        End If
    Next
    '卸载PicEdit
    For Each tmpCtrl In UserControl.Controls
        Unload tmpCtrl
    Next
    Erase SpecPaper, aPicFlag
    ReDim SpecPaper(0): ReDim aPicFlag(0)
    
    FileHeight = 0
    
    CtrlIndex = 1
    CtrlTop = CtrlDistance
    Seq = 0: iNum = -1: iNum = UBound(aElement, 2)
    lngInitProgValue = objProgressBar.Value
    For i = 0 To iNum
        bPicEnabled = bAllowEdit
    
        Load HSEdit(CtrlIndex)
        Load VSEdit(CtrlIndex)
        If aElement(15, i) = 1 Then
            Load lblFlag(CtrlIndex)
            With lblFlag(CtrlIndex)
                .Left = 100
                .Top = CtrlTop
                .Caption = LABEL_EXPAND
                .ZOrder 0
                '文本段如果不显示标题，则不允许收折。
                .Visible = IIf(aElement(18, i) > 0 Or ((aElement(18, i) = 0 Or aElement(18, i) = -5) And aElement(7, i) <> 0), True, False)
            End With
            
            '加载病历元素
            Select Case aElement(18, i)
                Case 0, -5
                    strTxtBox = ""
                    If aElement(20, i) <> 0 Then '读取病历内容
                        strSQL = "Select * From 病人病历文本段 Where 病历ID=[1]"
                        If blnMoved Then
                            strSQL = Replace(strSQL, "病人病历文本段", "H病人病历文本段")
                        End If
                        lngTmpID = Val(aElement(3, i))
                        Set rsTmp = OpenSQLRecord(strSQL, "", lngTmpID)
                        If Not rsTmp.EOF Then
                            strTxtBox = rsTmp("内容")
                            If Not bSampleFile Then strTxtBox = ReplaceString(strTxtBox, _
                                PatientID, CheckID, PatientType)
                        End If
                    End If
                    Load lblText(lblText.Count)
                    With lblText(lblText.Count - 1)
                        Erase aFont
                        aFont = Split(aElement(11, i), ",")
                        .FontName = aFont(0)
                        .FontSize = aFont(1)
                        .FontBold = aFont(2)
                        .FontItalic = aFont(3)
                        .FontUnderline = aFont(4)
                        .FontStrikethru = aFont(5)
                        
                        .Caption = strTxtBox
                        .Width = UserControl.ScaleWidth - VSMain.Width - lblFlag(CtrlIndex).Left - lblFlag(CtrlIndex).Width - 15
                    
                        Set TmpFont = UserControl.Font
                        Set UserControl.Font = .Font
                        iTmpLines = CInt(.Height / UserControl.TextHeight(" "))
                        sngLine_Indent = UserControl.TextHeight(" ") * 1.35
'                        CtrlHeight = .Height * 1.4
                        Set UserControl.Font = TmpFont
                    End With
                    Load txtBox(txtBox.Count)
                    With txtBox(txtBox.Count - 1)
                        .Font.Name = aFont(0)
                        .Font.Size = aFont(1)
                        .Font.Bold = aFont(2)
                        .Font.Italic = aFont(3)
                        .Font.Underline = aFont(4)
                        .Font.Strikethrough = aFont(5)
                        
                        '必须显示才能求得行数
                        Set .Container = PicMain: .Visible = True
                        .Left = 0: .Top = 0: .Width = UserControl.ScaleWidth - VSMain.Width - lblFlag(CtrlIndex).Left - lblFlag(CtrlIndex).Width - 15
                        .Text = strTxtBox: .Refresh
'                        iTmpLines = .GetLineFromChar(Len(.Text))
                        .Visible = False
                        
'                        CtrlHeight = lblText(lblText.Count - 1).Height + sngLine_Indent * iTmpLines
                        CtrlHeight = sngLine_Indent * iTmpLines
                        aElement(16, i) = 10000
                        
                        .Enabled = True: .Locked = Not bAllowEdit: bPicEnabled = True  'bAllowEdit
                        If aElement(7, i) = 0 Then .ToolTipText = aElement(6, i)
                        .Visible = False
                        
                        .TabIndex = Seq: Seq = Seq + 1
                        '初始控件相关变量
                        ReDim Preserve blnCurrUnderLine(txtBox.Count - 1)
                        ReDim Preserve blnEvent_SelChange(txtBox.Count - 1)
                        ReDim Preserve aTextItems(txtBox.Count - 1)
                        blnCurrUnderLine(.Index) = False
                        blnEvent_SelChange(.Index) = False
                        aTextItems(.Index) = ""
                        Call FormatText(.Index, .Text)
                        
                        .Visible = True
                    End With
                    aElement(17, i) = txtBox.Count - 1
                Case 1
                    Load grdTable(grdTable.Count)
                    With grdTable(grdTable.Count - 1)
                        InitTable grdTable(grdTable.Count - 1)
                        
                        Erase aFont
                        aFont = Split(aElement(11, i), ",")
                        .DefaultFontName = aFont(0)
                        .DefaultFontSize = -1 * (aFont(1) * 1440 / 72) '将磅转为缇
                        
                        If aElement(20, i) <> 0 Then '读取病历内容
                            ReadTable_Patient grdTable(grdTable.Count - 1), aElement(3, i)
                        Else
                            ReadTable grdTable(grdTable.Count - 1), aElement(3, i)
                        End If
                        .SetSelection 1, 1, .MaxRow, .MaxCol
                        .WordWrap = True
                        .SetSelection 1, 1, 1, 1
                        
                        .EnableProtection = True
                        
                        .RangeToTwips 1, 1, .MaxRow, .MaxCol, iTabLeft, iTabTop, iTabWidth, iTabHeight, iShown
                        .Left = 0: .Top = 0
                        .Width = iTabWidth + 15
                        .Height = iTabHeight + 15
                        
                        CtrlHeight = .Height
                        aElement(16, i) = .Height
                        aElement(19, i) = .Width
                        
                        .Enabled = True 'bAllowEdit
                        
                        .TabIndex = Seq: Seq = Seq + 1
                        .Visible = True
                    End With
                    aElement(17, i) = grdTable.Count - 1
                Case 2
                    strTxtBox = ""
                    '读取内容
                    If aElement(20, i) <> 0 Then
                        strSQL = "Select * From 病人病历文本段 Where 病历ID=[1]"
                        If blnMoved Then
                            strSQL = Replace(strSQL, "病人病历文本段", "H病人病历文本段")
                        End If
                        lngTmpID = Val(aElement(3, i))
                        Set rsTmp = OpenSQLRecord(strSQL, "", lngTmpID)
                        If Not rsTmp.EOF Then
                            strTxtBox = rsTmp("内容")
                            If Not bSampleFile Then strTxtBox = ReplaceString(strTxtBox, _
                                PatientID, CheckID, PatientType)
                        End If
                    End If
                    Load lblVisForm(lblVisForm.Count)
                    With lblVisForm(lblVisForm.Count - 1)
                        Erase aFont
                        aFont = Split(aElement(11, i), ",")
                        .FontName = aFont(0)
                        .FontSize = aFont(1)
                        .FontBold = aFont(2)
                        .FontItalic = aFont(3)
                        .FontUnderline = aFont(4)
                        .FontStrikethru = aFont(5)
                        
                        .Width = UserControl.ScaleWidth - VSMain.Width - lblFlag(CtrlIndex).Left + lblFlag(CtrlIndex).Width - 15
                        .Caption = strTxtBox
                    End With
                    Load txtVisForm(txtVisForm.Count)
                    With txtVisForm(txtVisForm.Count - 1)
                        .FontName = aFont(0)
                        .FontSize = aFont(1)
                        .FontBold = aFont(2)
                        .FontItalic = aFont(3)
                        .FontUnderline = aFont(4)
                        .FontStrikethru = aFont(5)
                        
                        .Left = 0: .Top = 0
                        
                        .Enabled = True: .Locked = Not bAllowEdit ': bPicEnabled = True 'bAllowEdit
                        .Text = strTxtBox
                        
                        .TabIndex = Seq: Seq = Seq + 1
                    End With
                    
                    Load VisForm(VisForm.Count)
                    With VisForm(VisForm.Count - 1)
                        Erase aFont
                        aFont = Split(aElement(11, i), ",")
                        .Font.Name = aFont(0)
                        .Font.Size = aFont(1)
                        .Font.Bold = aFont(2)
                        .Font.Italic = aFont(3)
                        .Font.Underline = aFont(4)
                        .Font.Strikethrough = aFont(5)
                        
                        Set .ParentObject = Me
                        
                        If aElement(20, i) <> 0 Then '读取病历内容
                            .ReadForm aElement(3, i), False, PatientID, CheckID, PatientType, , blnReplaced, blnMoved
                        Else
                            .ReadForm aElement(3, i), , PatientID, CheckID, PatientType, , blnReplaced, blnMoved
                        End If
                        
                        .Left = 0: .Top = 0
                        
                        CtrlHeight = .Height
                        aElement(16, i) = .Height
                        aElement(19, i) = .Width
                        
                        .Enabled = True 'bAllowEdit
                        
                        .TabIndex = Seq: Seq = Seq + 1
                        .Visible = True
                    End With
                    aElement(17, i) = VisForm.Count - 1
                Case 3
                    ReDim Preserve aPicFlag(UBound(aPicFlag) + 1)
                    If aElement(20, i) <> 0 Then '读取MapItems
                        Set aPicFlag(UBound(aPicFlag)) = GetMapItems(CLng(aElement(3, i)), blnMoved)
                    Else
                        Set aPicFlag(UBound(aPicFlag)) = New MapItems
                    End If
                    
                    Load PicFlag(PicFlag.Count)
                    With PicFlag(PicFlag.Count - 1)
                        Set .Picture = ReadCaseMap(CLng(aElement(21, i)))
                        .Width = .ScaleX(.Picture.Width, , vbTwips): .Height = .ScaleY(.Picture.Height, , vbTwips)
                        .Width = IIf(.Width > 10000, 10000, .Width): .Height = .Height * .Width / .ScaleX(.Picture.Width, , vbTwips)
                        .Cls: Set .Picture = Nothing
                        
                        ShowFlagInOjbect PicFlag(PicFlag.Count - 1), CLng(aElement(21, i)), aPicFlag(PicFlag.Count - 1), blnMoved:=blnMoved
                        .Left = 0: .Top = 0
                        
                        CtrlHeight = .Height
                        aElement(16, i) = .Height
                        aElement(19, i) = .Width
                        
                        .Enabled = True ' bAllowEdit
                        If aElement(7, i) = 0 Then .ToolTipText = aElement(6, i)
                        
                        .TabIndex = Seq: Seq = Seq + 1
                        .Visible = True
                    End With
                    aElement(17, i) = PicFlag.Count - 1
                Case 4
                    strTxtBox = ""
                    '读取内容
                    If aElement(20, i) <> 0 Then
                        strSQL = "Select * From 病人病历文本段 Where 病历ID=[1]"
                        If blnMoved Then
                            strSQL = Replace(strSQL, "病人病历文本段", "H病人病历文本段")
                        End If
                        lngTmpID = Val(aElement(3, i))
                        Set rsTmp = OpenSQLRecord(strSQL, "", lngTmpID)
                        If Not rsTmp.EOF Then
                            strTxtBox = rsTmp("内容")
                            If Not bSampleFile Then strTxtBox = ReplaceString(strTxtBox, _
                                PatientID, CheckID, PatientType)
                        End If
                    End If
                    Load lblSpecPaper(lblSpecPaper.Count)
                    With lblSpecPaper(lblSpecPaper.Count - 1)
                        Erase aFont
                        aFont = Split(aElement(11, i), ",")
                        .FontName = aFont(0)
                        .FontSize = aFont(1)
                        .FontBold = aFont(2)
                        .FontItalic = aFont(3)
                        .FontUnderline = aFont(4)
                        .FontStrikethru = aFont(5)
                        
                        .Width = UserControl.ScaleWidth - VSMain.Width - lblFlag(CtrlIndex).Left + lblFlag(CtrlIndex).Width - 15
                        .Caption = strTxtBox
                    End With
                    Load txtSpecPaper(txtSpecPaper.Count)
                    With txtSpecPaper(txtSpecPaper.Count - 1)
                        .FontName = aFont(0)
                        .FontSize = aFont(1)
                        .FontBold = aFont(2)
                        .FontItalic = aFont(3)
                        .FontUnderline = aFont(4)
                        .FontStrikethru = aFont(5)
                        
                        .Left = 0: .Top = 0
                        
                        .Enabled = True: .Locked = Not bAllowEdit: bPicEnabled = True  'bAllowEdit
                        .Text = strTxtBox
                        
                        .TabIndex = Seq: Seq = Seq + 1
                    End With
                    
                    ReDim Preserve SpecPaper(UBound(SpecPaper) + 1)
                    Licenses.Add aElement(0, i)
                    Set SpecPaper(UBound(SpecPaper)) = UserControl.Controls.Add(aElement(0, i), "SpecPaper" & UBound(SpecPaper))
                    With SpecPaper(UBound(SpecPaper))
                        .SetgcnOracle gcnOracle
                        .DataMoved = blnMoved
                                                
                        Call .SetDiagItem(SendAdviceID, SendNO)
                        
                        Set .ParentObject = Me
                        
                        .ID病人病历 = aElement(20, i): .Get医嘱id = AdviceID
                        .病人id = PatientID
                        
                        If PatientType = 0 Then .挂号单 = CheckID
                        If aElement(0, i) Like "*SPECRESULT" And bNotShowDiagItem Then .ShowItem = False
                        .Left = 0: .Top = 0

                        CtrlHeight = .Height
                        aElement(16, i) = 10000 '.Height
                        aElement(19, i) = 10000 '.Width

                        .DispMode = Not bAllowEdit
                        .TabIndex = Seq: Seq = Seq + 1
                        .Visible = True
                    End With
                    aElement(17, i) = UBound(SpecPaper)
                Case -4
                    CtrlHeight = 0
                    aElement(16, i) = 0
                    aElement(17, i) = 0
                Case Else '特殊元素
                    strTxtBox = ""
                    If aElement(20, i) <> 0 Then '读取病历内容
                        strSQL = "Select * From 病人病历文本段 Where 病历ID=[1]"
                        If blnMoved Then
                            strSQL = Replace(strSQL, "病人病历文本段", "H病人病历文本段")
                        End If
                        lngTmpID = Val(aElement(3, i))
                        Set rsTmp = OpenSQLRecord(strSQL, "", lngTmpID)
                        If Not rsTmp.EOF Then strTxtBox = rsTmp("内容")
                    Else
                        strTxtBox = GetSpecValue(CStr(aElement(18, i)), PatientID, CheckID, PatientType)
                    End If
                    
                    Load SpecItem(SpecItem.Count)
                    With SpecItem(SpecItem.Count - 1)
                        Erase aFont
                        aFont = Split(aElement(8, i), ",")
                        .Font.Name = aFont(0)
                        .Font.Size = aFont(1)
                        .Font.Bold = aFont(2)
                        .Font.Italic = aFont(3)
                        .Font.Underline = aFont(4)
                        .Font.Strikethrough = aFont(5)
                        
                        Select Case aElement(18, i)
                            Case -2 '当前日期
                                iItemLen = 10: sItemFormat = "YYYY-MM-DD": iItemType = 2
                            Case -3 '当前时间
                                iItemLen = 19: sItemFormat = "YYYY-MM-DD HH:MM:SS": iItemType = 2
                            Case Else
                                iItemLen = 100
                                sItemFormat = "": iItemType = 1
                        End Select
                        .Init IIf(aElement(7, i) = 0, "", aElement(6, i)), "", 0, iItemType, iItemLen, 0, "", strTxtBox, 0, "", , , sItemFormat
                        
                        .Left = 0: .Top = 0: .Width = 0
                        
                        CtrlHeight = .Height
                        aElement(16, i) = .Height
                        
                        .Enabled = (aElement(18, i) <> -1) 'bAllowEdit
                        
                        .TabIndex = Seq: Seq = Seq + 1
                        .Visible = True
                    End With
                    aElement(17, i) = SpecItem.Count - 1
            End Select
            
            If aElement(18, i) <> -4 And (aElement(7, i) = 0 Or (aElement(18, i) < 0 And aElement(18, i) <> -5)) Then '不显示标题
                Load lblTitle(CtrlIndex)
                
                Load picSplit(CtrlIndex)
                With picSplit(CtrlIndex)
                    .Left = lblFlag(CtrlIndex).Left + lblFlag(CtrlIndex).Width
                    .Top = CtrlTop + CtrlHeight + SplitDistance
                    .Width = PicMain.ScaleWidth
                
                    .Visible = True
                End With
                
                Load picEdit(CtrlIndex)
                With picEdit(CtrlIndex)
                    .Left = lblFlag(CtrlIndex).Left + lblFlag(CtrlIndex).Width
                    .Top = CtrlTop
                    .Width = PicMain.ScaleWidth - .Left - 15
                    .Height = CtrlHeight
                    .ToolTipText = aElement(6, i)
                    .Enabled = bPicEnabled ' bAllowEdit
                    .Visible = True
                End With
            Else
                Load lblTitle(CtrlIndex)
                With lblTitle(CtrlIndex)
                    .Alignment = IIf(aElement(10, i) = 1, 0, IIf(aElement(10, i) = 2, 2, 1))
                    .Left = lblFlag(CtrlIndex).Left + lblFlag(CtrlIndex).Width
                    .Width = IIf(aElement(14, i) = 10, TitleWidth, PicMain.ScaleWidth - .Left - 15)
                    
                    .Caption = aElement(6, i)
                    
                    Erase aFont
                    aFont = Split(aElement(8, i), ",")
                    .FontName = aFont(0)
                    .FontSize = aFont(1)
                    .FontBold = aFont(2)
                    .FontItalic = aFont(3)
                    .FontUnderline = aFont(4)
                    .FontStrikethru = aFont(5)
                    
                    If aElement(14, i) = 10 Then
                        If .Height > CtrlHeight Then .Height = CtrlHeight
                        .Top = CtrlTop + (CtrlHeight - .Height) / 2
                    Else
                        If .Height < TitleHeight Then .Height = TitleHeight
                        .Top = CtrlTop
                    End If
                    .Visible = True
                End With
            
                Load picSplit(CtrlIndex)
                With picSplit(CtrlIndex)
                    .Left = IIf(aElement(14, i) = 10, lblTitle(CtrlIndex).Left + lblTitle(CtrlIndex).Width + CtrlDistance, lblFlag(CtrlIndex).Left + lblFlag(CtrlIndex).Width)
                    .Top = IIf(aElement(14, i) = 10, CtrlTop, lblTitle(CtrlIndex).Top + lblTitle(CtrlIndex).Height) + CtrlHeight + SplitDistance
                    .Width = PicMain.ScaleWidth
                
                    .Visible = True
                End With
                
                Load picEdit(CtrlIndex)
                With picEdit(CtrlIndex)
                    .Left = IIf(aElement(14, i) = 10, lblTitle(CtrlIndex).Left + lblTitle(CtrlIndex).Width + CtrlDistance, lblFlag(CtrlIndex).Left + lblFlag(CtrlIndex).Width)
                    .Top = IIf(aElement(14, i) = 10, CtrlTop, lblTitle(CtrlIndex).Top + lblTitle(CtrlIndex).Height)
                    .Width = PicMain.ScaleWidth - .Left - 15
                    .Height = CtrlHeight
                    .Enabled = bPicEnabled ' bAllowEdit
                    .Visible = True
                End With
                CtrlHeight = IIf(aElement(14, i) = 10, CtrlHeight, lblTitle(CtrlIndex).Height + CtrlHeight)
            End If
            
            FileHeight = FileHeight + CtrlHeight + SplitDistance + CtrlDistance
            CtrlTop = CtrlTop + CtrlHeight + SplitDistance + CtrlDistance
        Else
            Load lblFlag(CtrlIndex)
            
            '加载病历元素
            CtrlHeight = 1000
            
            Load lblTitle(CtrlIndex)
            Load picSplit(CtrlIndex)
            Load picEdit(CtrlIndex)
        End If
        
        objProgressBar.Value = lngInitProgValue + CLng((100 - lngInitProgValue) * (i + 1) / (iNum + 1))
        CtrlIndex = CtrlIndex + 1
    Next
    SetMainVscroll
    
    bOnLoadFile = False
End Sub

Private Sub Refresh(Optional objProgressBar As ProgressBar, Optional blnReplaced As Boolean = False)
    '参数：blnReplaced 是否在所见单中强制替换具有替换域属性的所见项
    
    Dim tmpCtrl As VB.Control, CtrlIndex As Integer, CtrlHeight As Long, CtrlTop As Long
    Dim aFont() As String
    Dim i As Long, iNum As Long, Seq As Integer
    Dim rsTmp As New ADODB.Recordset, sTmpFile As String, FileObj As New Scripting.FileSystemObject
    Dim iTabLeft As Long, iTabTop As Long, iTabWidth As Long, iTabHeight As Long, iShown As Integer
    Dim iItemLen As Integer, sItemFormat As String, iItemType As Integer
    Dim bPicEnabled As Boolean
     
    Dim strTxtBox As String
    
    Dim lngInitProgValue As Long '初始进度值
    Dim TmpFont As StdFont, iTmpLines As Integer
    
    Dim sngLine_Indent As Single
    Dim strSQL As String, lngTmpID As Long
    
    bOnLoadFile = True
    blnMouseDown = False
    On Error Resume Next
    
    FileHeight = 0
    
    CtrlIndex = 1
    CtrlTop = CtrlDistance
    Seq = 0: iNum = -1: iNum = UBound(aElement, 2)
    lngInitProgValue = objProgressBar.Value
    For i = 0 To iNum
        bPicEnabled = bAllowEdit
    
        Load HSEdit(CtrlIndex): HSEdit(CtrlIndex).Visible = False
        Load VSEdit(CtrlIndex): VSEdit(CtrlIndex).Visible = False
        If aElement(15, i) = 1 Then
            Load lblFlag(CtrlIndex)
            With lblFlag(CtrlIndex)
                .Left = 100
                .Top = CtrlTop
                .Caption = LABEL_EXPAND
                '文本段如果不显示标题，则不允许收折。
                .Visible = IIf(aElement(18, i) > 0 Or ((aElement(18, i) = 0 Or aElement(18, i) = -5) And aElement(7, i) <> 0), True, False)
            End With
            
            '加载病历元素
            Select Case aElement(18, i)
                Case 0, -5
                    If aElement(17, i) <= 0 Then
                        Load lblText(lblText.Count)
                        Load txtBox(txtBox.Count)
                        aElement(17, i) = txtBox.Count - 1
                        
                        strTxtBox = ""
                        If aElement(20, i) <> 0 Then '读取病历内容
                            strSQL = "Select * From 病人病历文本段 Where 病历ID=[1]"
                            If blnMoved Then
                                strSQL = Replace(strSQL, "病人病历文本段", "H病人病历文本段")
                            End If
                            lngTmpID = Val(aElement(3, i))
                            Set rsTmp = OpenSQLRecord(strSQL, "", lngTmpID)
                            If Not rsTmp.EOF Then strTxtBox = rsTmp("内容")
                        End If
                    Else
                        strTxtBox = txtBox(aElement(17, i)).Text
                    End If
                    If Not bSampleFile Then strTxtBox = ReplaceString(strTxtBox, _
                        PatientID, CheckID, PatientType)
                    With lblText(aElement(17, i))
                        Erase aFont
                        aFont = Split(aElement(11, i), ",")
                        .FontName = aFont(0)
                        .FontSize = aFont(1)
                        .FontBold = aFont(2)
                        .FontItalic = aFont(3)
                        .FontUnderline = aFont(4)
                        .FontStrikethru = aFont(5)
                        
                        .Caption = strTxtBox
                        .Width = UserControl.ScaleWidth - VSMain.Width - lblFlag(CtrlIndex).Left - lblFlag(CtrlIndex).Width - 15
                    
                        Set TmpFont = UserControl.Font
                        Set UserControl.Font = .Font
                        iTmpLines = CInt(.Height / UserControl.TextHeight(" "))
                        sngLine_Indent = UserControl.TextHeight(" ") * 1.35
'                        CtrlHeight = .Height * 1.4
                        Set UserControl.Font = TmpFont
                    End With
                    With txtBox(aElement(17, i))
                        .Font.Name = aFont(0)
                        .Font.Size = aFont(1)
                        .Font.Bold = aFont(2)
                        .Font.Italic = aFont(3)
                        .Font.Underline = aFont(4)
                        .Font.Strikethrough = aFont(5)
                        
                        '必须显示才能求得行数
                        .Visible = True
                        .Left = 0: .Top = 0
                        .Text = strTxtBox: .Refresh
'                        iTmpLines = .GetLineFromChar(Len(.Text))
                        .Visible = False
                        
'                        CtrlHeight = lblText(aElement(17, i)).Height + sngLine_Indent * iTmpLines
                        CtrlHeight = sngLine_Indent * iTmpLines
                        aElement(16, i) = 10000
                        
                        .Enabled = True: .Locked = Not bAllowEdit: bPicEnabled = True 'bAllowEdit
                        If aElement(7, i) = 0 Then .ToolTipText = aElement(6, i)
                        .Visible = False
                        
                        .TabIndex = Seq: Seq = Seq + 1
                        '初始控件相关变量
                        ReDim Preserve blnCurrUnderLine(txtBox.Count - 1)
                        ReDim Preserve blnEvent_SelChange(txtBox.Count - 1)
                        ReDim Preserve aTextItems(txtBox.Count - 1)
                        blnCurrUnderLine(.Index) = False
                        blnEvent_SelChange(.Index) = False
                        aTextItems(.Index) = ""
                        Call FormatText(.Index, .Text)
                        
                        .Visible = True
                    End With
                Case 1
                    If aElement(17, i) <= 0 Then
                        Load grdTable(grdTable.Count)
                        aElement(17, i) = grdTable.Count - 1
                    
                        With grdTable(aElement(17, i))
                            InitTable grdTable(grdTable.Count - 1)
                            
                            Erase aFont
                            aFont = Split(aElement(11, i), ",")
                            .DefaultFontName = aFont(0)
                            .DefaultFontSize = -1 * (aFont(1) * 1440 / 72) '将磅转为缇
                            
                            If aElement(20, i) <> 0 Then '读取病历内容
                                ReadTable_Patient grdTable(grdTable.Count - 1), aElement(3, i)
                            Else
                                ReadTable grdTable(grdTable.Count - 1), aElement(3, i)
                            End If
                            .SetSelection 1, 1, .MaxRow, .MaxCol
                            .WordWrap = True
                            .SetSelection 1, 1, 1, 1
                            
                            .EnableProtection = True
                            
                            .RangeToTwips 1, 1, .MaxRow, .MaxCol, iTabLeft, iTabTop, iTabWidth, iTabHeight, iShown
                            .Left = 0: .Top = 0
                            .Width = iTabWidth + 15
                            .Height = iTabHeight + 15
                            
                            CtrlHeight = .Height
                            aElement(16, i) = .Height
                            aElement(19, i) = .Width
                            
                            .Enabled = True ' bAllowEdit
                            
                            .TabIndex = Seq: Seq = Seq + 1
                            .Visible = True
                        End With
                    Else
                        With grdTable(aElement(17, i))
                            .RangeToTwips 1, 1, .MaxRow, .MaxCol, iTabLeft, iTabTop, iTabWidth, iTabHeight, iShown
                            .Left = 0: .Top = 0
                            .Width = iTabWidth + 15
                            .Height = iTabHeight + 15
                            
                            CtrlHeight = .Height
                            aElement(16, i) = .Height
                            aElement(19, i) = .Width
                            
                            .Enabled = True 'bAllowEdit
                            
                            .TabIndex = Seq: Seq = Seq + 1
                            .Visible = True
                        End With
                    End If
                Case 2
                    If aElement(17, i) <= 0 Then
                        Load lblVisForm(lblVisForm.Count)
                        Load txtVisForm(txtVisForm.Count)
                        
                        Load VisForm(VisForm.Count)
                        aElement(17, i) = VisForm.Count - 1
                        
                        strTxtBox = ""
                    
                        With lblVisForm(aElement(17, i))
                            Erase aFont
                            aFont = Split(aElement(11, i), ",")
                            .FontName = aFont(0)
                            .FontSize = aFont(1)
                            .FontBold = aFont(2)
                            .FontItalic = aFont(3)
                            .FontUnderline = aFont(4)
                            .FontStrikethru = aFont(5)
                            
                            .Width = UserControl.ScaleWidth - VSMain.Width - lblFlag(CtrlIndex).Left + lblFlag(CtrlIndex).Width - 15
                            .Caption = strTxtBox
                        End With
                        With txtVisForm(aElement(17, i))
                            .FontName = aFont(0)
                            .FontSize = aFont(1)
                            .FontBold = aFont(2)
                            .FontItalic = aFont(3)
                            .FontUnderline = aFont(4)
                            .FontStrikethru = aFont(5)
                            
                            .Left = 0: .Top = 0
                            
                            .Enabled = True: .Locked = Not bAllowEdit ': bPicEnabled = True 'bAllowEdit
                            .Text = strTxtBox
                            
                            .TabIndex = Seq: Seq = Seq + 1
                        End With
                        With VisForm(aElement(17, i))
                            Erase aFont
                            aFont = Split(aElement(11, i), ",")
                            .Font.Name = aFont(0)
                            .Font.Size = aFont(1)
                            .Font.Bold = aFont(2)
                            .Font.Italic = aFont(3)
                            .Font.Underline = aFont(4)
                            .Font.Strikethrough = aFont(5)
                            
                            Set .ParentObject = Me
                        
                            If aElement(20, i) <> 0 Then '读取病历内容
                                .ReadForm aElement(3, i), False, PatientID, CheckID, PatientType, , blnReplaced, blnMoved
                            Else
                                .ReadForm aElement(3, i), , PatientID, CheckID, PatientType, , blnReplaced, blnMoved
                            End If
                            
                            .Left = 0: .Top = 0
                            
                            CtrlHeight = .Height
                            aElement(16, i) = .Height
                            aElement(19, i) = .Width
                            
                            .Enabled = True 'bAllowEdit
                            
                            .TabIndex = Seq: Seq = Seq + 1
                            .Visible = True
                        End With
                    Else
                        strTxtBox = txtVisForm(aElement(17, i)).Text
                        If Not bSampleFile Then strTxtBox = ReplaceString(strTxtBox, _
                            PatientID, CheckID, PatientType)
                        
                        lblVisForm(aElement(17, i)).Caption = strTxtBox
                        
                        With txtVisForm(aElement(17, i))
                            If .Visible Then CtrlHeight = lblVisForm(aElement(17, i)).Height
                            
                            .Enabled = True: .Locked = Not bAllowEdit ': bPicEnabled = True 'bAllowEdit
                            
                            .Text = strTxtBox
                            .TabIndex = Seq: Seq = Seq + 1
                        End With
                        With VisForm(aElement(17, i))
                            .Left = 0: .Top = 0
                            
                            If .Visible Then CtrlHeight = .Height
                            
                            .Enabled = True 'bAllowEdit
                            
                            .TabIndex = Seq: Seq = Seq + 1
                        End With
                    End If
                Case 3
                    If aElement(17, i) <= 0 Then
                        ReDim Preserve aPicFlag(UBound(aPicFlag) + 1)
                        If aElement(20, i) <> 0 Then '读取MapItems
                            Set aPicFlag(UBound(aPicFlag)) = GetMapItems(CLng(aElement(3, i)), blnMoved)
                        Else
                            Set aPicFlag(UBound(aPicFlag)) = New MapItems
                        End If
                        
                        Load PicFlag(PicFlag.Count)
                        aElement(17, i) = PicFlag.Count - 1
                    
                        With PicFlag(aElement(17, i))
                            Set .Picture = ReadCaseMap(CLng(aElement(21, i)))
                            .Width = .ScaleX(.Picture.Width, , vbTwips): .Height = .ScaleY(.Picture.Height, , vbTwips)
                            .Width = IIf(.Width > 10000, 10000, .Width): .Height = .Height * .Width / .ScaleX(.Picture.Width, , vbTwips)
                            .Cls: Set .Picture = Nothing
                            
                            ShowFlagInOjbect PicFlag(PicFlag.Count - 1), CLng(aElement(21, i)), aPicFlag(PicFlag.Count - 1), blnMoved:=blnMoved
                            .Left = 0: .Top = 0
                            
                            CtrlHeight = .Height
                            aElement(16, i) = .Height
                            aElement(19, i) = .Width
                            
                            .Enabled = True ' bAllowEdit
                            If aElement(7, i) = 0 Then .ToolTipText = aElement(6, i)
                            
                            .TabIndex = Seq: Seq = Seq + 1
                            .Visible = True
                        End With
                    Else
                        With PicFlag(aElement(17, i))
                            CtrlHeight = .Height
                            aElement(16, i) = .Height
                            aElement(19, i) = .Width
                            
                            .Enabled = True ' bAllowEdit
                            
                            .TabIndex = Seq: Seq = Seq + 1
                            .Visible = True
                        End With
                    End If
                Case 4
                    If aElement(17, i) <= 0 Then
                        Load lblSpecPaper(lblSpecPaper.Count)
                        Load txtSpecPaper(txtSpecPaper.Count)
                        ReDim Preserve SpecPaper(UBound(SpecPaper) + 1)
                        Licenses.Add aElement(0, i)
                        Set SpecPaper(UBound(SpecPaper)) = UserControl.Controls.Add(aElement(0, i), "SpecPaper" & UBound(SpecPaper))
                        aElement(17, i) = UBound(SpecPaper)
                        
                        strTxtBox = ""
                    
                        With lblSpecPaper(aElement(17, i))
                            Erase aFont
                            aFont = Split(aElement(11, i), ",")
                            .FontName = aFont(0)
                            .FontSize = aFont(1)
                            .FontBold = aFont(2)
                            .FontItalic = aFont(3)
                            .FontUnderline = aFont(4)
                            .FontStrikethru = aFont(5)
                            
                            .Width = UserControl.ScaleWidth - VSMain.Width - lblFlag(CtrlIndex).Left + lblFlag(CtrlIndex).Width - 15
                            .Caption = strTxtBox
                        End With
                        With txtSpecPaper(aElement(17, i))
                            .FontName = aFont(0)
                            .FontSize = aFont(1)
                            .FontBold = aFont(2)
                            .FontItalic = aFont(3)
                            .FontUnderline = aFont(4)
                            .FontStrikethru = aFont(5)
                            
                            .Left = 0: .Top = 0
                            
                            .Enabled = True: .Locked = Not bAllowEdit: bPicEnabled = True  'bAllowEdit
                            .Text = strTxtBox
                            
                            .TabIndex = Seq: Seq = Seq + 1
                        End With
                    
                        With SpecPaper(aElement(17, i))
                            .SetgcnOracle gcnOracle
                            .DataMoved = blnMoved
                            
                            Call .SetDiagItem(SendAdviceID, SendNO)
                            
                            Set .ParentObject = Me
                        
                            .ID病人病历 = aElement(20, i): .Get医嘱id = AdviceID
                            .病人id = PatientID
                            
                            If PatientType = 0 Then .挂号单 = CheckID
                            If aElement(0, i) Like "*SPECRESULT" And bNotShowDiagItem Then .ShowItem = False
                            .Left = 0: .Top = 0:
    
                            CtrlHeight = .Height
                            aElement(16, i) = 10000 '.Height
                            aElement(19, i) = 10000 '.Width
    
                            .DispMode = Not bAllowEdit
                            .TabIndex = Seq: Seq = Seq + 1
                            .Visible = True
                        End With
                    Else
                        strTxtBox = txtSpecPaper(aElement(17, i)).Text
                        If Not bSampleFile Then strTxtBox = ReplaceString(strTxtBox, _
                            PatientID, CheckID, PatientType)
                        
                        lblSpecPaper(aElement(17, i)).Caption = strTxtBox
                        
                        With txtSpecPaper(aElement(17, i))
                            If .Visible Then CtrlHeight = lblSpecPaper(aElement(17, i)).Height
                            
                            .Enabled = True: .Locked = Not bAllowEdit: bPicEnabled = True  'bAllowEdit
                            
                            .Text = strTxtBox
                            .TabIndex = Seq: Seq = Seq + 1
                        End With
                        With SpecPaper(aElement(17, i))
                            If .Visible Then CtrlHeight = .Height
    
                            .DispMode = Not bAllowEdit
                            .TabIndex = Seq: Seq = Seq + 1
                        End With
                    End If
                Case -4
                    CtrlHeight = 0
                    aElement(16, i) = 0
                    aElement(17, i) = 0
                Case Else '特殊元素
                    If aElement(17, i) <= 0 Then
                        Load SpecItem(SpecItem.Count)
                        aElement(17, i) = SpecItem.Count - 1
                    
                        strTxtBox = ""
                        If aElement(20, i) <> 0 Then '读取病历内容
                            strSQL = "Select * From 病人病历文本段 Where 病历ID=[1]"
                            If blnMoved Then
                                strSQL = Replace(strSQL, "病人病历文本段", "H病人病历文本段")
                            End If
                            lngTmpID = Val(aElement(3, i))
                            Set rsTmp = OpenSQLRecord(strSQL, "", lngTmpID)
                            If Not rsTmp.EOF Then strTxtBox = rsTmp("内容")
                        Else
                            strTxtBox = GetSpecValue(CStr(aElement(18, i)), PatientID, CheckID, PatientType)
                        End If
                    Else
                        strTxtBox = SpecItem(aElement(17, i)).Value
                    End If
                    
                    With SpecItem(aElement(17, i))
                        Erase aFont
                        aFont = Split(aElement(8, i), ",")
                        .Font.Name = aFont(0)
                        .Font.Size = aFont(1)
                        .Font.Bold = aFont(2)
                        .Font.Italic = aFont(3)
                        .Font.Underline = aFont(4)
                        .Font.Strikethrough = aFont(5)
                        
                        Select Case aElement(18, i)
                            Case -2 '当前日期
                                iItemLen = 10: sItemFormat = "YYYY-MM-DD": iItemType = 2
                            Case -3 '当前时间
                                iItemLen = 19: sItemFormat = "YYYY-MM-DD HH:MM:SS": iItemType = 2
                            Case Else
                                iItemLen = 100
                                sItemFormat = "": iItemType = 1
                        End Select
                        .Init IIf(aElement(7, i) = 0, "", aElement(6, i)), "", 0, iItemType, iItemLen, 0, "", strTxtBox, 0, "", , , sItemFormat
                        
                        .Left = 0: .Top = 0: .Width = 0
                        
                        CtrlHeight = .Height
                        aElement(16, i) = .Height
                        
                        .Enabled = (aElement(18, i) <> -1) 'bAllowEdit
                        
                        .TabIndex = Seq: Seq = Seq + 1
                        .Visible = True
                    End With
            End Select
            
            If aElement(18, i) <> -4 And (aElement(7, i) = 0 Or (aElement(18, i) < 0 And aElement(18, i) <> -5)) Then '不显示标题
                Load lblTitle(CtrlIndex): lblTitle(CtrlIndex).Visible = False
        
                
                Load picSplit(CtrlIndex)
                With picSplit(CtrlIndex)
                    .Left = lblFlag(CtrlIndex).Left + lblFlag(CtrlIndex).Width
                    .Top = CtrlTop + CtrlHeight + SplitDistance
                    .Width = PicMain.ScaleWidth
                
                    .Visible = True
                End With
                
                Load picEdit(CtrlIndex)
                With picEdit(CtrlIndex)
                    .Left = lblFlag(CtrlIndex).Left + lblFlag(CtrlIndex).Width
                    .Top = CtrlTop
                    .Width = PicMain.ScaleWidth - .Left - 15: .Height = 0
                    .Height = CtrlHeight
                    .ToolTipText = aElement(6, i)
                    .Enabled = bPicEnabled ' bAllowEdit
                    .Visible = True
                End With
            Else
                Load lblTitle(CtrlIndex)
                With lblTitle(CtrlIndex)
                    .Alignment = IIf(aElement(10, i) = 1, 0, IIf(aElement(10, i) = 2, 2, 1))
                    .Left = lblFlag(CtrlIndex).Left + lblFlag(CtrlIndex).Width
                    .Width = IIf(aElement(14, i) = 10, TitleWidth, PicMain.ScaleWidth - .Left - 15)
                    
                    .Caption = aElement(6, i)
                    
                    Erase aFont
                    aFont = Split(aElement(8, i), ",")
                    .FontName = aFont(0)
                    .FontSize = aFont(1)
                    .FontBold = aFont(2)
                    .FontItalic = aFont(3)
                    .FontUnderline = aFont(4)
                    .FontStrikethru = aFont(5)
                    
                    If aElement(14, i) = 10 Then
                        If .Height > CtrlHeight Then .Height = CtrlHeight
                        .Top = CtrlTop + (CtrlHeight - .Height) / 2
                    Else
                        If .Height < TitleHeight Then .Height = TitleHeight
                        .Top = CtrlTop
                    End If
                    .Visible = True
                End With
            
                Load picSplit(CtrlIndex)
                With picSplit(CtrlIndex)
                    .Left = IIf(aElement(14, i) = 10, lblTitle(CtrlIndex).Left + lblTitle(CtrlIndex).Width + CtrlDistance, lblFlag(CtrlIndex).Left + lblFlag(CtrlIndex).Width)
                    .Top = IIf(aElement(14, i) = 10, CtrlTop, lblTitle(CtrlIndex).Top + lblTitle(CtrlIndex).Height) + CtrlHeight + SplitDistance
                    .Width = PicMain.ScaleWidth
                
                    .Visible = True
                End With
                
                Load picEdit(CtrlIndex)
                With picEdit(CtrlIndex)
                    .Left = IIf(aElement(14, i) = 10, lblTitle(CtrlIndex).Left + lblTitle(CtrlIndex).Width + CtrlDistance, lblFlag(CtrlIndex).Left + lblFlag(CtrlIndex).Width)
                    .Top = IIf(aElement(14, i) = 10, CtrlTop, lblTitle(CtrlIndex).Top + lblTitle(CtrlIndex).Height)
                    .Width = PicMain.ScaleWidth - .Left - 15: .Height = 0
                    .Height = CtrlHeight
                    .Enabled = bPicEnabled ' bAllowEdit
                    .Visible = True
                End With
                CtrlHeight = IIf(aElement(14, i) = 10, CtrlHeight, lblTitle(CtrlIndex).Height + CtrlHeight)
            End If
            
            FileHeight = FileHeight + CtrlHeight + SplitDistance + CtrlDistance
            CtrlTop = CtrlTop + CtrlHeight + SplitDistance + CtrlDistance
        Else
            Load lblFlag(CtrlIndex)
            lblFlag(CtrlIndex).Visible = False
            
            '卸载病历元素
            Select Case aElement(18, i)
                Case 0, -5
                    lblText(aElement(17, i)).Visible = False
                    txtBox(aElement(17, i)).Visible = False
                Case 1
                    grdTable(aElement(17, i)).Visible = False
                Case 2
                    VisForm(aElement(17, i)).Visible = False
                    txtVisForm(aElement(17, i)).Visible = False
                    lblVisForm(aElement(17, i)).Visible = False
                Case 3
                    PicFlag(aElement(17, i)).Cls: Set PicFlag(aElement(17, i)).Picture = Nothing
                    Set aPicFlag(aElement(17, i)) = Nothing
                Case 4
                    Set SpecPaper(aElement(17, i)) = Nothing
                    UserControl.Controls.Remove "SpecPaper" & aElement(17, i)
                    txtSpecPaper(aElement(17, i)).Visible = False
                    lblSpecPaper(aElement(17, i)).Visible = False
            End Select
            
            CtrlHeight = 1000
            
            Load lblTitle(CtrlIndex)
            lblTitle(CtrlIndex).Visible = False
            Load picSplit(CtrlIndex)
            picSplit(CtrlIndex).Visible = False
            Load picEdit(CtrlIndex)
            picEdit(CtrlIndex).Visible = False
        End If
        
        objProgressBar.Value = lngInitProgValue + CLng((100 - lngInitProgValue) * (i + 1) / (iNum + 1))
        CtrlIndex = CtrlIndex + 1
    Next
    
    SetMainVscroll
    
    bOnLoadFile = False
End Sub

Private Sub CurrSpecPaper_GotFocus()
    If bAllowEdit Then zlCommFun.OpenIme False
    ShowElement CurrSpecPaper

    On Error Resume Next
    RaiseEvent ElementGotFocus(CurrSpecPaper.Container.Index, 4)
End Sub

Private Sub grdTable_CancelEdit(Index As Integer)
    bNotRunSelChange = False
End Sub

Private Sub grdTable_DblClick(Index As Integer, ByVal nRow As Long, ByVal nCol As Long)
    grdTable(Index).StartEdit False, True, False
End Sub

Private Sub grdTable_EndEdit(Index As Integer, EditString As String, Cancel As Integer)
    Dim iOldHeight As Long
    Dim iTabLeft As Long, iTabTop As Long, iTabWidth As Long, iTabHeight As Long, iShown As Integer
    Dim iDecPos As Integer
    With grdTable(Index)
        bModified = True
    
        If IsNumeric(EditString) Then
            iDecPos = InStr(EditString, ".")
            If iDecPos > 0 And iDecPos < Len(EditString) Then
                .NumberFormat = "#." + String(Len(EditString) - iDecPos, "0")
            Else
                .NumberFormat = "General"
            End If
        Else
            .NumberFormat = "General"
        End If
        .TextRC(.Row, .Col) = EditString
        
        iOldHeight = .RowHeight(.Row)
        .SetRowHeightAuto .SelStartRow, 1, .SelEndRow, .MaxCol, True
        If .RowHeight(.Row) <> iOldHeight Then
            ExpandElement .Container.Index, .RowHeight(.Row) - iOldHeight
        
            .RangeToTwips 1, 1, .MaxRow, .MaxCol, iTabLeft, iTabTop, iTabWidth, iTabHeight, iShown
            .Height = iTabHeight + 15
            aElement(16, .Container.Index - 1) = .Height
        End If
        bNotRunSelChange = False
    End With
End Sub

Private Sub grdTable_GotFocus(Index As Integer)
    If bAllowEdit Then zlCommFun.OpenIme False
    ShowElement grdTable(Index)
    
    With grdTable(Index)
        .Row = IIf(.Row <= .FixedRows, .FixedRows + 1, .Row)
        .Col = IIf(.Col <= .FixedCols, .FixedCols + 1, .Col)
        
        .ShowActiveCell
        bNotRunSelChange = False
    End With

    On Error Resume Next
    RaiseEvent ElementGotFocus(grdTable(Index).Container.Index, 1)
End Sub

Private Sub grdTable_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim NextCtrl As Control
    
    If KeyCode = vbKeyTab Then
        Set NextCtrl = NextElement(grdTable(Index).Container.Index)
        On Error Resume Next
        If Not NextCtrl Is Nothing Then NextCtrl.SetFocus
    End If
End Sub

Private Sub grdTable_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim objCellFormat As TTF160Ctl.F1CellFormat
    On Error Resume Next
    With grdTable(Index)
        Set objCellFormat = .GetCellFormat
        If Len(objCellFormat.ValidationText) > 0 Then
            grdTable_SelChange Index
            KeyAscii = 0
        End If
    End With
End Sub

Private Sub grdTable_LostFocus(Index As Integer)
    bNotRunSelChange = True
End Sub

Private Sub grdTable_SelChange(Index As Integer)
    Dim objCellFormat As TTF160Ctl.F1CellFormat
    Dim aVisItemInfo() As String
    
    On Error Resume Next
    If bNotRunSelChange Then Exit Sub
    If UserControl.ActiveControl.Name <> "grdTable" Then Exit Sub
    With grdTable(Index)
        Set objCellFormat = .GetCellFormat
        If Len(objCellFormat.ValidationText) > 0 Then
            aVisItemInfo = Split(objCellFormat.ValidationText, ",")
            Me.VisItem(aVisItemInfo(1)).SetFocus
        End If
    End With
End Sub

Private Sub grdTable_StartEdit(Index As Integer, EditString As String, Cancel As Integer)
    Dim objCellFormat As TTF160Ctl.F1CellFormat
    On Error Resume Next
    bNotRunSelChange = True
    With grdTable(Index)
        Set objCellFormat = .GetCellFormat
        If Len(objCellFormat.ValidationText) > 0 Then
            Cancel = True
        End If
    End With
End Sub

Private Sub grdTable_TopLeftChanged(Index As Integer)
    If bNotRunSelChange Then Exit Sub
    
    bNotRunSelChange = True
    Proc_Table_TopLeftChanged grdTable(Index)
    bNotRunSelChange = False
End Sub

Private Sub HSEdit_Change(Index As Integer)
    Dim tmpCtrl As Control
    On Error Resume Next
    Select Case aElement(18, Index - 1)
        Case 1
            Set tmpCtrl = grdTable(aElement(17, Index - 1))
        Case 2
            Set tmpCtrl = VisForm(aElement(17, Index - 1))
        Case 3
            Set tmpCtrl = PicFlag(aElement(17, Index - 1))
        Case 4
            Set tmpCtrl = SpecPaper(aElement(17, Index - 1))
    End Select
    tmpCtrl.Left = -1 * HSEdit(Index).Value
End Sub

Private Sub lblFlag_Click(Index As Integer)
    ExpandElement Index
End Sub

Private Sub lblTitle_Click(Index As Integer)
    Select Case aElement(18, Index - 1)
        Case -4 '段落标题
            On Error Resume Next
            picEdit(Index).SetFocus
            RaiseEvent ElementGotFocus(Index, -4)
    End Select
End Sub

Private Sub lblTitle_DblClick(Index As Integer)
    If lblFlag(Index).Visible Then ExpandElement Index
End Sub

Private Sub picEdit_DblClick(Index As Integer)
    EditElement Index
End Sub

Private Sub picEdit_GotFocus(Index As Integer)
    Select Case aElement(18, Index - 1)
        Case 3 '标记图
            PicFlag_GotFocus CInt(aElement(17, Index - 1))
    End Select
End Sub

Private Sub picEdit_Resize(Index As Integer)
    Dim iNewHeight As Long
    Dim objCellFormat As TTF160Ctl.F1CellFormat
    Dim iCurrRow As Integer, iCurrCol As Integer
    Dim iRow As Integer, iCol As Integer, aVisItemInfo() As String
    Dim TmpFont As StdFont, iTmpLines As Integer
    Dim sngLine_Indent As Single
    On Error Resume Next
    If aElement(15, Index - 1) <> 1 Then Exit Sub
    
    Select Case aElement(18, Index - 1)
        Case 0, -5
            With lblText(aElement(17, Index - 1))
                Set .Container = picEdit(Index)
                .Width = picEdit(Index).Width
                
'                iNewHeight = .Height * 1.4
                Set TmpFont = UserControl.Font
                Set UserControl.Font = .Font
                iTmpLines = CInt(.Height / UserControl.TextHeight(" "))
                sngLine_Indent = UserControl.TextHeight(" ") * 1.35
                Set UserControl.Font = TmpFont
            End With
            With txtBox(aElement(17, Index - 1))
                Set .Container = picEdit(Index)
                .Width = picEdit(Index).Width
                .Height = picEdit(Index).Height
'                iTmpLines = .GetLineFromChar(Len(.Text))
'
''                iNewHeight = lblText(aElement(17, Index - 1)).Height + sngLine_Indent * iTmpLines
                iNewHeight = sngLine_Indent * iTmpLines
            End With
            
            If iNewHeight <> picEdit(Index).Height And Not bOnLoadFile Then ExpandElement Index, iNewHeight - picEdit(Index).Height
        Case 1
            With grdTable(aElement(17, Index - 1))
                Set .Container = picEdit(Index)
                
                '将表内所见项放入容器中
                iCurrRow = .Row: iCurrCol = .Col
                For iRow = 1 To .MaxRow
                    For iCol = 1 To .MaxCol
                        .SetActiveCell iRow, iCol

                        Set objCellFormat = .GetCellFormat
                        If Len(objCellFormat.ValidationText) > 0 And iRow = .SelStartRow And iCol = .SelStartCol Then
                            aVisItemInfo = Split(objCellFormat.ValidationText, ",")
                            Set VisItem(aVisItemInfo(1)).Container = picEdit(Index)
                        End If
                    Next iCol
                Next iRow
                .SetActiveCell iCurrRow, iCurrCol
                
                .Width = IIf(aElement(19, Index - 1) > picEdit(Index).Width, picEdit(Index).Width, aElement(19, Index - 1))
                .Height = IIf(aElement(16, Index - 1) > picEdit(Index).Height, picEdit(Index).Height, aElement(16, Index - 1))

                If .Width <= picEdit(Index).Width Then
                    Select Case aElement(13, Index - 1)
                        Case 1
                            .Left = 0
                        Case 2
                            .Left = (picEdit(Index).Width - .Width) / 2
                        Case 3
                            .Left = picEdit(Index).Width - .Width
                    End Select
                Else
                    If .Left > 0 Then .Left = 0
                End If
                If Not .Visible Then grdTable_TopLeftChanged CInt(aElement(17, Index - 1))
            End With
        Case 2
            With lblVisForm(aElement(17, Index - 1))
                Set .Container = picEdit(Index)
                .Width = picEdit(Index).Width
                
                iNewHeight = .Height
            End With
            With txtVisForm(aElement(17, Index - 1))
                Set .Container = picEdit(Index)
                .Width = picEdit(Index).Width
                .Height = picEdit(Index).Height
            End With
            If iNewHeight <> picEdit(Index).Height And Not bOnLoadFile And txtVisForm(aElement(17, Index - 1)).Visible Then
                ExpandElement Index, iNewHeight - picEdit(Index).Height
            End If
            
            With VisForm(aElement(17, Index - 1))
                Set .Container = picEdit(Index)
                
                If Not txtVisForm(aElement(17, Index - 1)).Visible Then
                    .Width = IIf(aElement(19, Index - 1) > picEdit(Index).Width, picEdit(Index).Width, aElement(19, Index - 1))
                    .Height = IIf(aElement(16, Index - 1) > picEdit(Index).Height, picEdit(Index).Height, aElement(16, Index - 1))
    
                    If .Width <= picEdit(Index).Width Then
                        Select Case aElement(13, Index - 1)
                            Case 1
                                .Left = 0
                            Case 2
                                .Left = (picEdit(Index).Width - .Width) / 2
                            Case 3
                                .Left = picEdit(Index).Width - .Width
                        End Select
                    Else
                        If .Left > 0 Then .Left = 0
                    End If
                End If
            End With
        Case 3
            Proc_PicEdit_Resize PicFlag(aElement(17, Index - 1)), Index
        Case 4
            With lblSpecPaper(aElement(17, Index - 1))
                Set .Container = picEdit(Index)
                .Width = picEdit(Index).Width
                
                iNewHeight = .Height
            End With
            With txtSpecPaper(aElement(17, Index - 1))
                Set .Container = picEdit(Index)
                .Width = picEdit(Index).Width
                .Height = picEdit(Index).Height
            End With
            If iNewHeight <> picEdit(Index).Height And Not bOnLoadFile And txtSpecPaper(aElement(17, Index - 1)).Visible Then
                ExpandElement Index, iNewHeight - picEdit(Index).Height
            End If
            
            With SpecPaper(aElement(17, Index - 1))
                Set .Container = picEdit(Index)
                
                If Not txtSpecPaper(aElement(17, Index - 1)).Visible Then
                    .Width = picEdit(Index).Width ' IIf(aElement(19, Index - 1) > picEdit(Index).Width, picEdit(Index).Width, aElement(19, Index - 1))
                    .Height = picEdit(Index).Height ' IIf(aElement(16, Index - 1) > picEdit(Index).Height, picEdit(Index).Height, aElement(16, Index - 1))
'                    If .Width > picEdit(Index).Width Then
'                        UserControl.Width = .Width + picEdit(Index).Left + 300 + picMain.Left + VSMain.Width
'                        Exit Sub
'                    End If
'                    If .Height <> picEdit(Index).Height Then
'                        picEdit(Index).Height = .Height
'                        Exit Sub
'                    End If
    
                    If .Width <= picEdit(Index).Width Then
                        Select Case aElement(13, Index - 1)
                            Case 1
                                .Left = 0
                            Case 2
                                .Left = (picEdit(Index).Width - .Width) / 2
                            Case 3
                                .Left = picEdit(Index).Width - .Width
                        End Select
                    Else
                        If .Left > 0 Then .Left = 0
                    End If
                End If
            End With
        Case Else '特殊元素
            With SpecItem(aElement(17, Index - 1))
                Set .Container = picEdit(Index)
                
                Set picEdit(Index).Font = .Font
                .Width = 0
                .Width = .Width + picEdit(Index).TextWidth(.Value) - picEdit(Index).TextWidth(" ")
                If .Width <= picEdit(Index).Width Then
                    Select Case aElement(10, Index - 1)
                        Case 1
                            .Left = 0
                        Case 2
                            .Left = (picEdit(Index).Width - .Width) / 2
                        Case 3
                            .Left = picEdit(Index).Width - .Width
                    End Select
                Else
                    If .Left > 0 Then .Left = 0
                End If
            End With
    End Select
End Sub

Public Sub EditElement(ByVal Index As Integer)
    Dim aMapFlags As Variant
    
    Select Case aElement(18, Index - 1)
        Case 3 '标记图
            Set aMapFlags = EditFlag(UserControl.Parent, CLng(aElement(21, Index - 1)), aPicFlag(aElement(17, Index - 1)))
            If Not aMapFlags Is Nothing Then
                bModified = True
                aElement(23, Index - 1) = 1
    
                Set aPicFlag(aElement(17, Index - 1)) = aMapFlags
                ShowFlagInOjbect PicFlag(aElement(17, Index - 1)), CLng(aElement(21, Index - 1)), aPicFlag(aElement(17, Index - 1)), blnMoved:=blnMoved
            End If
            
            On Error Resume Next
            picEdit(Index).SetFocus
    End Select
End Sub

Private Sub Proc_PicEdit_Resize(theControl As Control, ByVal Index As Integer)
    Dim iOrgWidth As Long, iOrgHeight As Long, iNewWidth As Long, iNewHeight As Long
    With theControl
        Set .Container = picEdit(Index)
            
        iNewHeight = IIf(lblFlag(Index).Caption = LABEL_EXPAND, picSplit(Index).Top - picEdit(Index).Top - SplitDistance, _
            picEdit(Index).Height + IIf(picEdit(Index).Width < aElement(19, Index - 1), HSEdit(Index).Height, 0))
        iNewWidth = PicMain.ScaleWidth - picEdit(Index).Left - 15
        iOrgWidth = iNewWidth
        iOrgHeight = iNewHeight
        
        If .Width > iNewWidth Then
            iNewHeight = iNewHeight - HSEdit(Index).Height
        End If
        If .Height > iNewHeight Then
            iNewWidth = iNewWidth - VSEdit(Index).Width
        End If
        If .Width > iNewWidth And iNewHeight = iOrgHeight Then
            iNewHeight = iNewHeight - HSEdit(Index).Height
        End If
        If .Height > iNewHeight And iNewWidth = iOrgWidth Then
            iNewWidth = iNewWidth - VSEdit(Index).Width
        End If
        
        picEdit(Index).Width = iNewWidth
        picEdit(Index).Height = iNewHeight
            
        If picEdit(Index).Height < iOrgHeight Then
            HSEdit(Index).Left = picEdit(Index).Left
            HSEdit(Index).Top = picEdit(Index).Top + picEdit(Index).Height
            HSEdit(Index).Width = picEdit(Index).Width
            
            SetHSEditScroll Index
            HSEdit(Index).Visible = (lblFlag(Index).Caption = LABEL_EXPAND)
        Else
            .Left = 0
            HSEdit(Index).Visible = False
        End If
        If picEdit(Index).Width < iOrgWidth Then
            VSEdit(Index).Left = picEdit(Index).Left + picEdit(Index).Width
            VSEdit(Index).Top = picEdit(Index).Top
            VSEdit(Index).Height = picEdit(Index).Height
            
            SetVSEditScroll Index
            VSEdit(Index).Visible = (lblFlag(Index).Caption = LABEL_EXPAND)
        Else
            .Top = 0
            VSEdit(Index).Visible = False
        End If
        
        If .Width <= picEdit(Index).Width Then
            Select Case aElement(13, Index - 1)
                Case 1
                    .Left = 0
                Case 2
                    .Left = (picEdit(Index).Width - .Width) / 2
                Case 3
                    .Left = picEdit(Index).Width - .Width
            End Select
        Else
            If .Left > 0 Then .Left = 0
        End If
    End With
End Sub

Private Sub PicFlag_GotFocus(Index As Integer)
    If bAllowEdit Then zlCommFun.OpenIme False
    ShowElement PicFlag(Index)

    On Error Resume Next
    RaiseEvent ElementGotFocus(PicFlag(Index).Container.Index, 3)
End Sub

Private Sub picMain_Resize()
    With picMargin
        .Left = 0: .Top = 0
        .Width = lblFlag(0).Width: .Height = PicMain.ScaleHeight
    End With
End Sub

Private Sub txtBox_Change(Index As Integer)
    Dim iTmpLines As Integer, TmpFont As StdFont, tmpStart As Long
    Dim blnTmp As Boolean
    Dim sngLine_Indent As Single

    On Error Resume Next
    bModified = True
    
    If bOnLoadFile Then Exit Sub
    If aElement(18, txtBox(Index).Container.Index - 1) < 0 And aElement(18, txtBox(Index).Container.Index - 1) <> -5 Then Exit Sub
    
    lblText(Index).Caption = txtBox(Index).Text
    
    Set TmpFont = UserControl.Font
    Set UserControl.Font = lblText(Index).Font
    iTmpLines = CInt(lblText(Index).Height / UserControl.TextHeight(" "))
    sngLine_Indent = UserControl.TextHeight(" ") * 1.35
    Set UserControl.Font = TmpFont
    
    If txtBox(Index).Visible Then
'        iTmpLines = txtBox(Index).GetLineFromChar(Len(txtBox(Index).Text))
        If sngLine_Indent * iTmpLines <> txtBox(Index).Height Then
            ExpandElement txtBox(Index).Container.Index, sngLine_Indent * iTmpLines - txtBox(Index).Height + lnHeightDis
        End If
'        If lblText(Index).Height * 1.4 <> txtBox(Index).Height Then
'            ExpandElement txtBox(Index).Container.Index, lblText(Index).Height * 1.4 - txtBox(Index).Height + lnHeightDis
'        End If
        With txtBox(Index)
            blnTmp = blnEvent_SelChange(Index)
            blnEvent_SelChange(Index) = True
            .SetFocus
            tmpStart = .SelStart
            .SelStart = 0
            .SelStart = tmpStart
            blnEvent_SelChange(Index) = blnTmp
        End With
    End If

    'RTF控件处理
    If Not txtBox(Index).Visible Then Exit Sub
    blnEvent_SelChange(Index) = False
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
    If bAllowEdit Then zlCommFun.OpenIme True
    If Not blnMouseDown Then ShowElement txtBox(Index)
    
    On Error Resume Next
    RaiseEvent ElementGotFocus(txtBox(Index).Container.Index, 0)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim NextCtrl As Control
    
    If (aElement(18, txtBox(Index).Container.Index - 1) < 0 And _
        aElement(18, txtBox(Index).Container.Index - 1) <> -5 And KeyCode = vbKeyReturn) Or _
        (KeyCode = vbKeyReturn And Shift = vbCtrlMask) Then
        txtBox(Index).Tag = "0" '不运行Key_Press事件
        Set NextCtrl = NextElement(txtBox(Index).Container.Index)
        On Error Resume Next
        If Not NextCtrl Is Nothing Then NextCtrl.SetFocus
    End If
End Sub

Private Sub txtBox_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim lngItemSeq As Long
    Dim tmpLeft As Long, tmpTop As Long, tmpPoint As POINTAPI
    
    If txtBox(Index).Tag = "0" Then txtBox(Index).Tag = "": KeyAscii = 0: Exit Sub

    If Not txtBox(Index).Visible Then Exit Sub
    blnEvent_SelChange(Index) = Not blnEvent_SelChange(Index)
'    If txtBox(Index).SelUnderline Then
'        If KeyAscii = 13 Then
'            KeyAscii = 0
'            '下一个输入项
'            NextItem Index
'            Exit Sub
'        End If
'        If KeyAscii = 32 And txtBox(Index).SelColor <> 0 Then
'            '空格调出选择
'            KeyAscii = 0
'            lngItemSeq = txtBox(Index).SelColor Xor COLOR_COMBO
'            tmpPoint.x = txtBox(Index).Left / Screen.TwipsPerPixelX: tmpPoint.y = txtBox(Index).Top / Screen.TwipsPerPixelY
'            Call ClientToScreen(txtBox(Index).hWnd, tmpPoint)
'            GetSelect Index, lngItemSeq - 1, tmpPoint.x * Screen.TwipsPerPixelX, tmpPoint.y * Screen.TwipsPerPixelY
'            blnEvent_SelChange(Index) = False
'            Exit Sub
'        End If
'    End If
End Sub

Private Sub txtBox_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    blnMouseDown = True
End Sub

Private Sub txtBox_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngItemSeq As Long
    Dim tmpLeft As Long, tmpTop As Long, tmpPoint As POINTAPI
    
    blnMouseDown = False

    If Not txtBox(Index).Visible Then Exit Sub
    If Button <> vbRightButton Then Exit Sub
'    With txtBox(Index)
'        If .SelUnderline And .SelColor <> 0 Then
'            lngItemSeq = .SelColor Xor COLOR_COMBO
'            tmpPoint.x = x / Screen.TwipsPerPixelX: tmpPoint.y = y / Screen.TwipsPerPixelY
'            Call ClientToScreen(.hWnd, tmpPoint)
'            GetSelect Index, lngItemSeq - 1, tmpPoint.x * Screen.TwipsPerPixelX, tmpPoint.y * Screen.TwipsPerPixelY
'        End If
'    End With
End Sub

Private Sub txtBox_SelChange(Index As Integer)
    If Not txtBox(Index).Visible Then Exit Sub
    If blnEvent_SelChange(Index) Then Exit Sub
'    With txtBox(Index)
'        If .SelUnderline And Not blnCurrUnderLine(Index) Then
'            SetSelect Index
'        Else
'            If Not .SelUnderline Then blnCurrUnderLine(Index) = False
'        End If
'    End With
End Sub

Private Sub txtSpecPaper_Change(Index As Integer)
    On Error Resume Next
    bModified = True
    
    If bOnLoadFile Then Exit Sub
    If aElement(18, txtSpecPaper(Index).Container.Index - 1) < 0 And aElement(18, txtSpecPaper(Index).Container.Index - 1) <> -5 Then Exit Sub
    
    lblSpecPaper(Index).Caption = txtSpecPaper(Index)
    If txtSpecPaper(Index).Visible Then
        If lblSpecPaper(Index).Height <> txtSpecPaper(Index).Height Then
            ExpandElement txtSpecPaper(Index).Container.Index, lblSpecPaper(Index).Height - txtSpecPaper(Index).Height + lnHeightDis
        End If
        txtSpecPaper(Index).SetFocus
    End If
End Sub

Private Sub txtSpecPaper_GotFocus(Index As Integer)
    If bAllowEdit Then zlCommFun.OpenIme True
    ShowElement txtSpecPaper(Index)
    
    On Error Resume Next
    RaiseEvent ElementGotFocus(txtSpecPaper(Index).Container.Index, 2)
End Sub

Private Sub txtSpecPaper_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim NextCtrl As Control
    
    If (aElement(18, txtSpecPaper(Index).Container.Index - 1) < 0 And _
        aElement(18, txtSpecPaper(Index).Container.Index - 1) <> -5 And KeyCode = vbKeyReturn) Or _
        (KeyCode = vbKeyReturn And Shift = vbCtrlMask) Then
        txtSpecPaper(Index).Tag = "0" '不运行Key_Press事件
        Set NextCtrl = NextElement(txtSpecPaper(Index).Container.Index)
        On Error Resume Next
        If Not NextCtrl Is Nothing Then NextCtrl.SetFocus
    End If
End Sub

Private Sub txtSpecPaper_KeyPress(Index As Integer, KeyAscii As Integer)
    If txtSpecPaper(Index).Tag = "0" Then txtSpecPaper(Index).Tag = "": KeyAscii = 0
End Sub

Private Sub txtVisForm_Change(Index As Integer)
    On Error Resume Next
    bModified = True
    
    If bOnLoadFile Then Exit Sub
    If aElement(18, txtVisForm(Index).Container.Index - 1) < 0 And aElement(18, txtVisForm(Index).Container.Index - 1) <> -5 Then Exit Sub
    
    lblVisForm(Index).Caption = txtVisForm(Index)
    If txtVisForm(Index).Visible Then
        If lblVisForm(Index).Height <> txtVisForm(Index).Height Then
            ExpandElement txtVisForm(Index).Container.Index, lblVisForm(Index).Height - txtVisForm(Index).Height + lnHeightDis
        End If
        txtVisForm(Index).SetFocus
    End If
End Sub

Private Sub txtVisForm_GotFocus(Index As Integer)
    If bAllowEdit Then zlCommFun.OpenIme True
    ShowElement txtVisForm(Index)
    
    On Error Resume Next
    RaiseEvent ElementGotFocus(txtVisForm(Index).Container.Index, 2)
End Sub

Private Sub txtVisForm_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim NextCtrl As Control
    
    If (aElement(18, txtVisForm(Index).Container.Index - 1) < 0 And _
        aElement(18, txtVisForm(Index).Container.Index - 1) <> -5 And KeyCode = vbKeyReturn) Or _
        (KeyCode = vbKeyReturn And Shift = vbCtrlMask) Then
        txtVisForm(Index).Tag = "0" '不运行Key_Press事件
        Set NextCtrl = NextElement(txtVisForm(Index).Container.Index)
        On Error Resume Next
        If Not NextCtrl Is Nothing Then NextCtrl.SetFocus
    End If
End Sub

Private Sub txtVisForm_KeyPress(Index As Integer, KeyAscii As Integer)
    If txtVisForm(Index).Tag = "0" Then txtVisForm(Index).Tag = "": KeyAscii = 0
End Sub

Private Sub UserControl_Initialize()
    '这几个属性可开放
    TitleWidth = 1080: TitleHeight = 300
    CtrlDistance = 60: SplitDistance = 15
    
    With PicMain
        .Left = MARGIN_PAPER: .Top = MARGIN_PAPER
    End With
End Sub

Private Sub ExpandElement(ByVal Index As Long, Optional ByVal iOffset As Single = 0)
    Dim bExpand As Boolean, iOldSplitTop As Long, iOldHeight As Long
    Dim i As Long
    Dim iOrgWidth As Long, iOrgHeight As Long
    
    On Error Resume Next
    
    If iOffset = 0 Then
        With lblFlag(Index)
            If .Caption = LABEL_EXPAND Then
                bExpand = False
                .Caption = LABEL_COLLAPSE
            Else
                bExpand = True
                .Caption = LABEL_EXPAND
            End If
        End With
        
        picEdit(Index).Visible = bExpand
            
        iOrgHeight = picSplit(Index).Top - picEdit(Index).Top - SplitDistance
        iOrgWidth = PicMain.ScaleWidth - picEdit(Index).Left - 15
        If bExpand And aElement(18, Index - 1) = 3 Then
            If picEdit(Index).Width < aElement(19, Index - 1) Then
                HSEdit(Index).Visible = True
            Else
                HSEdit(Index).Visible = False
            End If
            If picEdit(Index).Width < iOrgWidth Then
                VSEdit(Index).Visible = True
            Else
                VSEdit(Index).Visible = False
            End If
        Else
            HSEdit(Index).Visible = False
            VSEdit(Index).Visible = False
        End If
        
        iOldSplitTop = picSplit(Index).Top
        If aElement(7, Index - 1) = 0 Then
            picSplit(Index).Top = lblFlag(Index).Top + IIf(bExpand, picEdit(Index).Height + IIf(HSEdit(Index).Visible, HSEdit(Index).Height, 0), TitleHeight) + SplitDistance
        Else
            If aElement(14, Index - 1) = 10 Then
                picSplit(Index).Top = lblFlag(Index).Top + IIf(bExpand, picEdit(Index).Height + IIf(HSEdit(Index).Visible, HSEdit(Index).Height, 0), TitleHeight) + SplitDistance
            Else
                picSplit(Index).Top = picEdit(Index).Top + (picEdit(Index).Height + IIf(HSEdit(Index).Visible, HSEdit(Index).Height, 0)) * IIf(bExpand, 1, 0) + SplitDistance
            End If
        End If
        iOffset = picSplit(Index).Top - iOldSplitTop
        
        FileHeight = FileHeight + (picEdit(Index).Height + IIf(HSEdit(Index).Visible, HSEdit(Index).Height, 0)) * IIf(bExpand, 1, -1)
    Else
        iOldHeight = picEdit(Index).Height
        
        picEdit(Index).Height = picEdit(Index).Height + iOffset
        iOffset = picEdit(Index).Height - iOldHeight
        
        picSplit(Index).Top = picEdit(Index).Top + picEdit(Index).Height + IIf(HSEdit(Index).Visible, HSEdit(Index).Height, 0) + SplitDistance
        
        FileHeight = FileHeight + picEdit(Index).Height - iOldHeight
    End If
    
    If aElement(7, Index - 1) = 0 Then
    Else
        If aElement(14, Index - 1) = 10 Then
            lblTitle(Index).Caption = ""
            lblTitle(Index).Caption = aElement(6, Index - 1)
            
            If lblTitle(Index).Height > picSplit(Index).Top - lblFlag(Index).Top - SplitDistance Then lblTitle(Index).Height = picSplit(Index).Top - lblFlag(Index).Top - SplitDistance
            lblTitle(Index).Top = lblFlag(Index).Top + (picSplit(Index).Top - lblFlag(Index).Top - lblTitle(Index).Height) / 2
        End If
    End If
    
    For i = Index + 1 To lblTitle.Count - 1
        lblTitle(i).Top = lblTitle(i).Top + iOffset
        lblFlag(i).Top = lblFlag(i).Top + iOffset
        picSplit(i).Top = picSplit(i).Top + iOffset
        picEdit(i).Top = picEdit(i).Top + iOffset
        HSEdit(i).Top = HSEdit(i).Top + iOffset
        VSEdit(i).Top = VSEdit(i).Top + iOffset
    Next
    
    SetMainVscroll
End Sub

Private Sub picSplit_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    '文本及所见单文本不能变高度
    Select Case aElement(18, Index - 1)
        Case 0, -5
            Exit Sub
        Case 2
            If txtVisForm(aElement(17, Index - 1)).Visible Then Exit Sub
        Case 4
            If txtSpecPaper(aElement(17, Index - 1)).Visible Then Exit Sub
    End Select

    If lblFlag(Index).Caption = LABEL_COLLAPSE Then Exit Sub
    If (Not Button = 1) Or Abs(y) = 0 Then Exit Sub
    
    If picEdit(Index).Height <= TitleHeight And y < 0 Then Exit Sub
    If picEdit(Index).Height >= aElement(16, Index - 1) And y >= 0 Then Exit Sub
    
    picSplit(Index).Top = picSplit(Index).Top + y
    ExpandElement Index, y
End Sub

Private Sub UserControl_InitProperties()
    bAllowEdit = False
    MARGIN_PAPER = 60
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngNewRow As Long, lngNewCol As Long
    On Error Resume Next
    
    If Not ifEditKey(KeyCode, False) Then bModified = True: aElement(23, UserControl.ActiveControl.Container.Index) = 1
    If UCase(UserControl.ActiveControl.Name) <> "GRDTABLE" Then Exit Sub
    
    If KeyCode = vbKeyReturn Then
        With UserControl.ActiveControl
            If .Row = .MaxRow Then
                lngNewRow = .FixedRows + 1
                If .Col = .MaxCol Then
                    lngNewCol = .FixedCols + 1
                Else
                    lngNewCol = .Col + 1
                End If
            Else
                lngNewRow = .Row + 1: lngNewCol = .Col
            End If
            .SetActiveCell lngNewRow, lngNewCol
            .ShowActiveCell
        End With
        KeyCode = 0
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    bAllowEdit = PropBag.ReadProperty("AllowEdit", False)
    MARGIN_PAPER = PropBag.ReadProperty("Border_Width", 60)
End Sub

Private Sub UserControl_Resize()
    Dim i As Long, iNum As Long
    On Error Resume Next
    With PicMain
        .Left = MARGIN_PAPER
        .Width = UserControl.ScaleWidth - VSMain.Width - MARGIN_PAPER - .Left
    End With
    With VSMain
        .Left = UserControl.ScaleWidth - .Width: .Top = 0
        .Height = UserControl.ScaleHeight
    End With

    iNum = lblTitle.UBound
    For i = 1 To iNum
        lblTitle(i).Width = IIf(aElement(14, i - 1) = 10, TitleWidth, PicMain.ScaleWidth - lblTitle(i).Left - 15)
    Next
    iNum = picEdit.UBound
    For i = 1 To iNum
        picEdit(i).Width = PicMain.ScaleWidth - picEdit(i).Left - 15
    Next
    iNum = picSplit.UBound
    For i = 1 To iNum
        picSplit(i).Width = PicMain.ScaleWidth
    Next
    
    SetMainVscroll
    
    RaiseEvent Resize
End Sub

'判断是否为编辑键
Private Function ifEditKey(ByVal KeyAscii As Integer, Optional ByVal AllowSubtract As Boolean = True) As Boolean
    If KeyAscii = vbKeyBack Or (KeyAscii = vbKeyInsert And AllowSubtract) Or KeyAscii = vbKeyDelete Or _
      KeyAscii = vbKeyHome Or KeyAscii = vbKeyEnd Or KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or _
      KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Or KeyAscii = vbKeyShift Or KeyAscii = vbKeyControl Or KeyAscii = vbKeyMenu Or _
      KeyAscii = vbKeyCapital Or KeyAscii = vbKeyPageUp Or KeyAscii = vbKeyPageDown Or _
      KeyAscii = vbKeyUp Or KeyAscii = vbKeyDown Or KeyAscii = vbKeyNumlock Then
        ifEditKey = True
    Else
        ifEditKey = False
    End If
End Function

Private Sub SetMainVscroll()
    On Error Resume Next
    
    With PicMain
        .Height = IIf(FileHeight + TitleHeight + 2 * MARGIN_PAPER < UserControl.ScaleHeight, UserControl.ScaleHeight - 2 * MARGIN_PAPER, FileHeight + TitleHeight)
    End With
    With VSMain
        .Enabled = IIf(PicMain.Height + 2 * MARGIN_PAPER > UserControl.ScaleHeight, True, False)
        If .Enabled Then
            .Min = -1 * MARGIN_PAPER
            .Max = PicMain.Height + MARGIN_PAPER - UserControl.ScaleHeight
            
            '处理Max超界的情况
            .Tag = CInt((PicMain.Height + MARGIN_PAPER _
                - UserControl.ScaleHeight) / .Max) '滚动倍率
            If CInt(.Tag) * .Max < PicMain.Height + MARGIN_PAPER _
                - UserControl.ScaleHeight Then .Tag = CInt(.Tag) + 1
            .Min = .Min / CInt(.Tag)
            .Max = (PicMain.Height + MARGIN_PAPER - UserControl.ScaleHeight) / CInt(.Tag)
            
            .SmallChange = UserControl.ScaleHeight / (10 * CInt(.Tag))
            .LargeChange = UserControl.ScaleHeight / CInt(.Tag)
        Else
            If .Value > -1 * MARGIN_PAPER Then .Value = -1 * MARGIN_PAPER
        End If
    End With
End Sub

Public Sub Release()
    '卸载动态加载的控件
    Dim tmpCtrl As VB.Control
    
    If bAllowEdit Then zlCommFun.OpenIme False
    
    '卸载所有病历控件
    On Error Resume Next
    For Each tmpCtrl In UserControl.Controls
        If UCase(tmpCtrl.Name) Like "SPECPAPER*" Then
            UserControl.Controls.Remove tmpCtrl.Name
        Else
            Unload tmpCtrl
        End If
    Next
    '卸载PicEdit
    For Each tmpCtrl In UserControl.Controls
        Unload tmpCtrl
    Next
    Erase SpecPaper, aPicFlag
    ReDim SpecPaper(0): ReDim aPicFlag(0)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "AllowEdit", bAllowEdit, False
    PropBag.WriteProperty "Border_Width", MARGIN_PAPER, 60
End Sub

Private Sub VisForm_GotFocus(Index As Integer)
    If bAllowEdit Then zlCommFun.OpenIme False
    ShowElement VisForm(Index)

    On Error Resume Next
    RaiseEvent ElementGotFocus(VisForm(Index).Container.Index, 2)
End Sub

Private Sub VisForm_NextControl(Index As Integer)
    Dim NextCtrl As Control
    
    Set NextCtrl = NextElement(VisForm(Index).Container.Index)
    On Error Resume Next
    If Not NextCtrl Is Nothing Then NextCtrl.SetFocus
End Sub

Private Sub SpecItem_GotFocus(Index As Integer)
    ShowElement SpecItem(Index)
    
    On Error Resume Next
    RaiseEvent ElementGotFocus(SpecItem(Index).Container.Index, aElement(18, SpecItem(Index).Container.Index - 1))
End Sub

Private Sub SpecItem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim NextCtrl As Control
    
    If (aElement(18, SpecItem(Index).Container.Index - 1) < 0 And _
        aElement(18, SpecItem(Index).Container.Index - 1) <> -5 And KeyCode = vbKeyReturn) Or _
        (KeyCode = vbKeyReturn And Shift = vbCtrlMask) Then
        Set NextCtrl = NextElement(SpecItem(Index).Container.Index)
        On Error Resume Next
        If Not NextCtrl Is Nothing Then NextCtrl.SetFocus
    Else
        bModified = True
    End If
End Sub

Private Sub VisItem_GotFocus(Index As Integer)
    Dim aCellInfo() As String

    On Error Resume Next
    aCellInfo = Split(VisItem(Index).Tag, ",")
    
    grdTable(CInt(aCellInfo(2))).SetActiveCell aCellInfo(0), aCellInfo(1)
End Sub

Private Sub VisItem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim aCellInfo() As String
    
    On Error Resume Next
    If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then
        aCellInfo = Split(VisItem(Index).Tag, ",")
        grdTable(CInt(aCellInfo(2))).SetFocus
        zlCommFun.PressKey CByte(KeyCode)
    End If
End Sub

Private Sub VSEdit_Change(Index As Integer)
    Dim tmpCtrl As Control
    On Error Resume Next
    Select Case aElement(18, Index - 1)
        Case 1
            Set tmpCtrl = grdTable(aElement(17, Index - 1))
        Case 2
            Set tmpCtrl = VisForm(aElement(17, Index - 1))
        Case 3
            Set tmpCtrl = PicFlag(aElement(17, Index - 1))
        Case 4
            Set tmpCtrl = SpecPaper(aElement(17, Index - 1))
    End Select
    tmpCtrl.Top = -1 * VSEdit(Index).Value
End Sub

Private Sub VSMain_Change()
    On Error Resume Next
    PicMain.Top = CDbl(-1 * VSMain.Value) * CInt(VSMain.Tag)
End Sub

Private Sub SetHSEditScroll(ByVal Index As Integer)
    Dim tmpCtrl As Control
    On Error Resume Next
    Select Case aElement(18, Index - 1)
        Case 1
            Set tmpCtrl = grdTable(aElement(17, Index - 1))
        Case 2
            Set tmpCtrl = VisForm(aElement(17, Index - 1))
        Case 3
            Set tmpCtrl = PicFlag(aElement(17, Index - 1))
        Case 4
            Set tmpCtrl = SpecPaper(aElement(17, Index - 1))
    End Select
    With HSEdit(Index)
        .Min = 0
        .Max = tmpCtrl.Width - picEdit(Index).Width
        .SmallChange = picEdit(Index).Width / 10
        .LargeChange = picEdit(Index).Width
    End With
End Sub

Private Sub SetVSEditScroll(ByVal Index As Integer)
    Dim tmpCtrl As Control
    On Error Resume Next
    Select Case aElement(18, Index - 1)
        Case 1
            Set tmpCtrl = grdTable(aElement(17, Index - 1))
        Case 2
            Set tmpCtrl = VisForm(aElement(17, Index - 1))
        Case 3
            Set tmpCtrl = PicFlag(aElement(17, Index - 1))
        Case 4
            Set tmpCtrl = SpecPaper(aElement(17, Index - 1))
    End Select
    With VSEdit(Index)
        .Min = 0
        .Max = tmpCtrl.Height - picEdit(Index).Height
        .SmallChange = picEdit(Index).Height / 10
        .LargeChange = picEdit(Index).Height
    End With
End Sub

Public Property Get AllowEdit() As Boolean
Attribute AllowEdit.VB_Description = "是否允许编辑"
Attribute AllowEdit.VB_ProcData.VB_Invoke_Property = ";行为"
    AllowEdit = bAllowEdit
End Property

Public Property Let AllowEdit(ByVal vNewValue As Boolean)
    Dim tmpCtrl As Control
    
    bAllowEdit = vNewValue
    
    On Error Resume Next
    For Each tmpCtrl In UserControl.Controls
        If UCase(tmpCtrl.Name) = "PICEDIT" Then
            If tmpCtrl.Index > 0 Then tmpCtrl.Enabled = vNewValue
        End If
    Next
End Property

Public Property Get Border_Width() As Integer
Attribute Border_Width.VB_Description = "病历页面的边框宽度"
Attribute Border_Width.VB_ProcData.VB_Invoke_Property = ";外观"
    Border_Width = MARGIN_PAPER
End Property

Public Property Let Border_Width(ByVal vNewValue As Integer)
    MARGIN_PAPER = vNewValue

    With PicMain
        .Left = MARGIN_PAPER: .Top = MARGIN_PAPER
    End With
    UserControl_Resize
End Property

Private Function NextElement(ByVal Index As Integer) As Control
    Dim i As Long, iNum As Long
    On Error Resume Next
    iNum = -1
    iNum = UBound(aElement, 2)
    Set NextElement = Nothing
    
    For i = Index To iNum
        If aElement(15, i) = 1 And lblFlag(i + 1).Caption = LABEL_EXPAND And aElement(18, i) <> 3 And aElement(18, i) <> -4 Then
            Select Case aElement(18, i)
                Case 0, -5
                    Set NextElement = txtBox(aElement(17, i))
                Case 1
                    Set NextElement = grdTable(aElement(17, i))
                Case 2
                    Set NextElement = IIf(VisForm(aElement(17, i)).Visible, VisForm(aElement(17, i)), txtVisForm(aElement(17, i)))
                Case 3
                    Set NextElement = PicFlag(aElement(17, i))
                Case 4
                    Set NextElement = IIf(txtSpecPaper(aElement(17, i)).Visible, txtSpecPaper(aElement(17, i)), SpecPaper(aElement(17, i)))
                Case Else
                    Set NextElement = SpecItem(aElement(17, i))
            End Select
            Exit For
        End If
    Next
End Function

Private Function PrevElement(ByVal Index As Integer) As Control
    Dim i As Long
    On Error Resume Next
    Set PrevElement = Nothing
    
    For i = Index To 0 Step -1
        If aElement(15, i) = 1 And lblFlag(i + 1).Caption = LABEL_EXPAND And aElement(18, i) <> 3 And aElement(18, i) <> -4 Then
            Select Case aElement(18, i)
                Case 0, -5
                    Set PrevElement = txtBox(aElement(17, i))
                Case 1
                    Set PrevElement = grdTable(aElement(17, i))
                Case 2
                    Set PrevElement = IIf(VisForm(aElement(17, i)).Visible, VisForm(aElement(17, i)), txtVisForm(aElement(17, i)))
                Case 3
                    Set PrevElement = PicFlag(aElement(17, i))
                Case 4
                    Set PrevElement = IIf(txtSpecPaper(aElement(17, i)).Visible, txtSpecPaper(aElement(17, i)), SpecPaper(aElement(17, i)))
                Case Else
                    Set PrevElement = SpecItem(aElement(17, i))
            End Select
            Exit For
        End If
    Next
End Function
'保存病人病历
Public Function SaveFile() As String
    On Error GoTo DBError
    SaveFile = ""
    
    SaveFileData
    
    SaveFile = PatientFileID: bModified = False
    Exit Function
DBError:
    If Err.Number = vbObjectError + 1 Then
        If Len(Err.Description) > 0 Then MsgBox Err.Description, vbInformation, gstrSysName
    Else
        If ErrCenter() = 1 Then Resume
        SaveErrLog
    End If
End Function

Private Sub SaveFileData()
    Dim i As Long, iNum As Long
    Dim strSaveSql As String, tmpFileID As String, ItemID As String
    Dim rsTmp As New ADODB.Recordset
    Dim FileType As Integer, FileName As String
    Dim ErrorNumber As Long, ErrorMsg As String, strSQL As String, aSQLs() As String, iSQLSeq As Integer, iSQLNum As Integer
    Dim lngPageID As Long, lngPatientID As Long
    Dim bAddFile As Boolean
    Dim bNewVersion As Boolean '修订为新版本
    
    On Error Resume Next
    iNum = -1
    iNum = UBound(aElement, 2)
    
    bNewVersion = False
    
    gcnOracle.BeginTrans
    On Error GoTo DBError
    tmpFileID = PatientFileID
    If Not bSampleFile Then '新增病历文件
        bAddFile = False
        If Len(PatientFileID) = 0 Then
            bAddFile = True
            tmpFileID = zlDatabase.GetNextId("病人病历记录")
        Else
            strSQL = "Select * From 病人病历记录 Where ID=[1]"
            Set rsTmp = OpenSQLRecord(strSQL, "获取病历记录", PatientFileID)
            If rsTmp.EOF Then bAddFile = True
        End If
        
        If bAddFile Then
            '获取病历种类
            strSQL = "Select 种类,名称 From 病历文件目录 Where ID=[1]"
            Set rsTmp = OpenSQLRecord(strSQL, "获取病历种类", FileTypeID)
            If rsTmp.EOF Then
                '未找到则设为门诊或住院病历
                FileType = PatientType + 1
                FileName = ""
            Else
                FileType = rsTmp(0)
                FileName = rsTmp(1)
            End If
            
            If AdviceID = 0 Then
                strSaveSql = "ZL_病人病历_INSERT(" & tmpFileID & "," + PatientID + ",'" + IIf(PatientType = 0, "", CheckID) + "','" + IIf(PatientType = 0, CheckID, "") + "',0,'" & UserInfo.部门ID & "'," & FileType & "," + _
                    FileTypeID + ",'" + FileName + "','" + UserInfo.姓名 + "')"
            Else
                strSaveSql = "ZL_病人病历_INSERT(" & tmpFileID & "," + PatientID + ",'" + IIf(PatientType = 0, "", CheckID) + "','" + IIf(PatientType = 0, CheckID, "") + "',0,'" & UserInfo.部门ID & "'," & FileType & "," + _
                    FileTypeID + ",'" + FileName + "','" + UserInfo.姓名 + "'," & AdviceID & ")"
            End If
            gcnOracle.Execute strSaveSql, , adCmdStoredProc
        ElseIf NVL(rsTmp("审阅人")) <> UserInfo.姓名 Then '修改病历，处理是否修订存储
            strSaveSql = "ZL_病人病历修订_INSERT(" & PatientFileID & ",'" + UserInfo.姓名 + "')"
            gcnOracle.Execute strSaveSql, , adCmdStoredProc
            bNewVersion = True
        End If
    End If
    
    For i = 0 To iNum
        If aElement(15, i) = 1 Then
            If aElement(20, i) = 0 Or bNewVersion Then '新增病历内容
                ItemID = zlDatabase.GetNextId("病人病历内容")
                aElement(20, i) = ItemID
            Else '修改
                ItemID = aElement(20, i)
                
                strSaveSql = "ZL_病人病历内容_DELETE(" & ItemID & ")"
                gcnOracle.Execute strSaveSql, , adCmdStoredProc
            End If
            '处理修改后的签名
            If aElement(23, i) = 1 Then NewSign i
            
            Select Case aElement(18, i)
                Case 0, -5
                    strSaveSql = "ZL_病人病历内容_INSERT(" & ItemID & "," & IIf(bSampleFile, tmpFileID, "''") & "," & IIf(bSampleFile, "''", tmpFileID) & "," & i & _
                        "," & aElement(18, i) & ",'" & aElement(22, i) & "'," & aElement(5, i) & ",'" & aElement(6, i) & "','" & aElement(7, i) & "','" & aElement(8, i) & _
                        "'," & aElement(10, i) & ",0,'" & aElement(11, i) & _
                        "'," & aElement(13, i) & ",0," & aElement(14, i) & ")"
                    gcnOracle.Execute strSaveSql, , adCmdStoredProc
                    strSaveSql = "ZL_病人病历文本段_SAVE(" & ItemID & ",1,'" & Replace(txtBox(aElement(17, i)).Text, "'", "''") & "')"
                    gcnOracle.Execute strSaveSql, , adCmdStoredProc
                Case 1
                    strSaveSql = "ZL_病人病历内容_INSERT(" & ItemID & "," & IIf(bSampleFile, tmpFileID, "''") & "," & IIf(bSampleFile, "''", tmpFileID) & "," & i & _
                        ",1,'" & aElement(22, i) & "'," & aElement(5, i) & ",'" & aElement(6, i) & "','" & aElement(7, i) & "','" & aElement(8, i) & _
                        "'," & aElement(10, i) & ",0,'" & aElement(11, i) & _
                        "'," & aElement(13, i) & ",0," & aElement(14, i) & ")"
                    gcnOracle.Execute strSaveSql, , adCmdStoredProc
                    SaveTable_Patient ItemID, grdTable(aElement(17, i)), gcnOracle
                Case 2
                    strSaveSql = "ZL_病人病历内容_INSERT(" & ItemID & "," & IIf(bSampleFile, tmpFileID, "''") & "," & IIf(bSampleFile, "''", tmpFileID) & "," & i & _
                        ",2,'" & aElement(22, i) & "'," & aElement(5, i) & ",'" & aElement(6, i) & "','" & aElement(7, i) & "','" & aElement(8, i) & _
                        "'," & aElement(10, i) & ",0,'" & aElement(11, i) & _
                        "'," & aElement(13, i) & ",0," & aElement(14, i) & ")"
                    gcnOracle.Execute strSaveSql, , adCmdStoredProc
                    '保存所见单文本
                    If Len(Trim(txtVisForm(aElement(17, i)))) = 0 Then
                        txtVisForm(aElement(17, i)) = VisForm(aElement(17, i)).Text
                        lblVisForm(aElement(17, i)) = txtVisForm(aElement(17, i))
                    End If
                    If Len(Trim(txtVisForm(aElement(17, i)))) > 0 Then
                        strSaveSql = "ZL_病人病历文本段_SAVE(" & ItemID & ",1,'" & Replace(txtVisForm(aElement(17, i)), "'", "''") & "')"
                        gcnOracle.Execute strSaveSql, , adCmdStoredProc
                    End If
                    
                    VisForm(aElement(17, i)).SaveForm ItemID, gcnOracle, ErrorNumber, ErrorMsg
                    If ErrorNumber <> 0 Then
                        Err.Description = ErrorMsg
                        Err.Raise ErrorNumber, "病历编辑"
                    End If
                Case 3
                    strSaveSql = "ZL_病人病历内容_INSERT(" & ItemID & "," & IIf(bSampleFile, tmpFileID, "''") & "," & IIf(bSampleFile, "''", tmpFileID) & "," & i & _
                        ",3,'" & aElement(22, i) & "'," & aElement(5, i) & ",'" & aElement(6, i) & "','" & aElement(7, i) & "','" & aElement(8, i) & _
                        "'," & aElement(10, i) & ",0,'" & aElement(11, i) & _
                        "'," & aElement(13, i) & ",0," & aElement(14, i) & ")"
                    gcnOracle.Execute strSaveSql, , adCmdStoredProc
                    SaveFlag ItemID, aPicFlag(aElement(17, i)), gcnOracle
                Case 4
                    strSaveSql = "ZL_病人病历内容_INSERT(" & ItemID & "," & IIf(bSampleFile, tmpFileID, "''") & "," & IIf(bSampleFile, "''", tmpFileID) & "," & i & _
                        ",4,'" & aElement(22, i) & "'," & aElement(5, i) & ",'" & aElement(6, i) & "','" & aElement(7, i) & "','" & aElement(8, i) & _
                        "'," & aElement(10, i) & ",0,'" & aElement(11, i) & _
                        "'," & aElement(13, i) & ",0," & aElement(14, i) & ")"
                    gcnOracle.Execute strSaveSql, , adCmdStoredProc
                    '保存专用纸文本
                    If Len(Trim(txtSpecPaper(aElement(17, i)))) = 0 Or aElement(5, i) = 0 Then
                        txtSpecPaper(aElement(17, i)) = SpecPaper(aElement(17, i)).Text
                        lblSpecPaper(aElement(17, i)) = txtSpecPaper(aElement(17, i))
                    End If
                    If Len(Trim(txtSpecPaper(aElement(17, i)))) > 0 Then
                        strSaveSql = "ZL_病人病历文本段_SAVE(" & ItemID & ",1,'" & Replace(txtSpecPaper(aElement(17, i)), "'", "''") & "')"
                        gcnOracle.Execute strSaveSql, , adCmdStoredProc
                    End If
                    
                    If PatientType = 0 Then
                        lngPageID = 0 '门诊的有主页ID吗？
                    Else
                        If IsNumeric(CheckID) Then
                            lngPageID = CLng(CheckID)
                        Else
                            lngPageID = 0
                        End If
                    End If
                    If IsNumeric(PatientID) Then
                        lngPatientID = CLng(PatientID)
                    Else
                        lngPatientID = 0
                    End If
                    strSQL = ""
                    If Not SpecPaper(aElement(17, i)).SaveData(lngPatientID, lngPageID, CLng(ItemID), strSQL, ErrorMsg) Then
                        Err.Description = ErrorMsg
                        If Err.Number = 0 Then
                            Err.Raise vbObjectError + 1, "病历编辑"
                        Else
                            Err.Raise Err.Number, "病历编辑"
                        End If
                    Else
                        aSQLs = Split(strSQL, Chr(9))
                        iSQLNum = UBound(aSQLs, 1)
                        For iSQLSeq = 0 To iSQLNum
                            gcnOracle.Execute aSQLs(iSQLSeq), , adCmdStoredProc
                        Next
                    End If
                Case -4
                    strSaveSql = "ZL_病人病历内容_INSERT(" & ItemID & "," & IIf(bSampleFile, tmpFileID, "''") & "," & IIf(bSampleFile, "''", tmpFileID) & "," & i & _
                        ",-4,'" & aElement(22, i) & "'," & aElement(5, i) & ",'" & aElement(6, i) & "','" & aElement(7, i) & "','" & aElement(8, i) & _
                        "'," & aElement(10, i) & ",0,'" & aElement(11, i) & _
                        "'," & aElement(13, i) & ",0," & aElement(14, i) & ")"
                    gcnOracle.Execute strSaveSql, , adCmdStoredProc
                Case Else
                    strSaveSql = "ZL_病人病历内容_INSERT(" & ItemID & "," & IIf(bSampleFile, tmpFileID, "''") & "," & IIf(bSampleFile, "''", tmpFileID) & "," & i & _
                        "," & aElement(18, i) & ",'" & aElement(22, i) & "'," & aElement(5, i) & ",'" & aElement(6, i) & "','" & aElement(7, i) & "','" & aElement(8, i) & _
                        "'," & aElement(10, i) & ",0,'" & aElement(11, i) & _
                        "'," & aElement(13, i) & ",0," & aElement(14, i) & ")"
                    gcnOracle.Execute strSaveSql, , adCmdStoredProc
                    strSaveSql = "ZL_病人病历文本段_SAVE(" & ItemID & ",1,'" & Replace(SpecItem(aElement(17, i)).Value, "'", "''") & "')"
                    gcnOracle.Execute strSaveSql, , adCmdStoredProc
            End Select
        Else
            If aElement(20, i) <> 0 Then '从病历内容中删除该元素
                strSaveSql = "ZL_病人病历内容_DELETE(" & aElement(20, i) & ")"
                gcnOracle.Execute strSaveSql, , adCmdStoredProc
            End If
        End If
    Next
    
    gcnOracle.CommitTrans
    '清除元素修改标志
    For i = 0 To iNum
        aElement(23, i) = 0
    Next
    If Len(PatientFileID) = 0 Then PatientFileID = tmpFileID '病历变成修改状态
    Exit Sub
DBError:
    gcnOracle.RollbackTrans
    Err.Raise Err.Number, "病人病历保存"
End Sub

Public Sub ShowFile(ByVal FileID As String, Optional ByVal sPatientID As String = "", _
    Optional ByVal sPageID As String = "", Optional ByVal iPatientType As Integer = 0, _
    Optional ByVal sTemplateID As String = "", Optional ByVal bSample As Boolean = False, _
    Optional ByVal iFilter As Integer = 0, Optional objProgressBar As ProgressBar, _
    Optional lngAdviceID As Long = 0, Optional lngSendAdviceID As Long, Optional lngSendNO As Long, Optional DataMoved As Boolean = False)
'iFilter：过滤文件组成：0=不过滤、1=申请项目、2=报告项目
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    PatientFileID = FileID: bSampleFile = bSample: AdviceID = lngAdviceID
    SendAdviceID = lngSendAdviceID: SendNO = lngSendNO
    blnMoved = DataMoved
    If blnMoved Then bAllowEdit = False
    
    If Len(FileID) = 0 Then
        PatientID = IIf(bSample, "", sPatientID)
        CheckID = IIf(bSample, "", sPageID)
        PatientType = iPatientType
        FileTypeID = sTemplateID
    Else
        If bSample Then '示范
            strSQL = "Select nvl(a.文件ID,0),nvl(b.种类,1) From 病历示范目录 a,病历文件目录 b Where a.ID=[1] And a.文件ID=b.ID(+)"
            Set rsTmp = OpenSQLRecord(strSQL, "病历文件显示", PatientFileID)
            If Not rsTmp.EOF Then
                PatientType = IIf(rsTmp(1) = 1, 0, 1)
                FileTypeID = rsTmp(0)
            End If
        Else
            If FileID * 1 < 0 Then '低版本的病历
                strSQL = "Select 病人ID,主页ID,nvl(挂号单,'0'),nvl(文件ID,0) From 病人病历记录 a,病人病历修订记录 b Where a.ID=b.病历记录ID And b.ID=[1]"
                If blnMoved Then
                    strSQL = Replace(strSQL, "病人病历记录", "H病人病历记录")
                    strSQL = Replace(strSQL, "病人病历修订记录", "H病人病历修订记录")
                End If
                Set rsTmp = OpenSQLRecord(strSQL, "病历文件显示", -1 * PatientFileID)
            Else
                strSQL = "Select 病人ID,主页ID,nvl(挂号单,'0'),nvl(文件ID,0) From 病人病历记录 Where ID=[1]"
                If blnMoved Then
                    strSQL = Replace(strSQL, "病人病历记录", "H病人病历记录")
                End If
                Set rsTmp = OpenSQLRecord(strSQL, "病历文件显示", PatientFileID)
            End If
            If Not rsTmp.EOF Then
                PatientID = rsTmp(0)
                CheckID = IIf(IsNull(rsTmp(1)), rsTmp(2), rsTmp(1))
                PatientType = IIf(IsNull(rsTmp(1)), 0, 1)
                FileTypeID = rsTmp(3)
            End If
        End If
    End If

    On Error Resume Next
    objProgressBar.Value = 10
    
    aElement = Array()
    If Len(FileID) = 0 Then
        strSQL = "Select b.部件,a.排列序号,a.填写时机,a.病历元素ID,a.必输项目,nvl(a.文本转储,0),a.标题文本,a.标题显示,a.标题字体,0,a.标题位置,a.内容字体,0,a.内容位置,a.嵌入方式,1,0,0,b.类型,0,0,a.病历元素ID,b.编码,0" + _
            " From 病历文件组成 a,病历元素目录 b Where a.病历文件ID=[1] And a.病历元素ID=b.ID" + _
            IIf(iFilter = 0, "", " And 填写时机=[2]") + " Order By a.排列序号"
        Set rsTmp = OpenSQLRecord(strSQL, "病历文件显示", FileTypeID, iFilter)
        aElement = rsTmp.GetRows()
    Else
        If bSample Then
            strSQL = "Select b.部件,a.排列序号,0,a.ID,0,nvl(a.文本转储,0),a.标题文本,a.标题显示,a.标题字体,0,a.标题位置,a.内容字体,0,a.内容位置,a.嵌入方式,1,0,0,a.元素类型,0,a.ID,b.ID,b.编码,0" + _
                " From 病人病历内容 a,病历元素目录 b Where a.病历示范ID=[1] And a.元素编码=b.编码(+) Order By a.排列序号"
            If blnMoved Then
                strSQL = Replace(strSQL, "病人病历内容", "H病人病历内容")
            End If
        Else
            If FileID * 1 < 0 Then '低版本的病历
                strSQL = "Select b.部件,a.排列序号,0,a.ID,0,nvl(a.文本转储,0),a.标题文本,a.标题显示,a.标题字体,0,a.标题位置,a.内容字体,0,a.内容位置,a.嵌入方式,1,0,0,a.元素类型,0,a.ID,b.ID,b.编码,0" + _
                    " From 病人病历内容 a,病历元素目录 b Where a.病历修订ID=[2] And a.元素编码=b.编码(+) Order By a.排列序号"
                If blnMoved Then
                    strSQL = Replace(strSQL, "病人病历内容", "H病人病历内容")
                End If
            Else
                strSQL = "Select b.部件,a.排列序号,0,a.ID,0,nvl(a.文本转储,0),a.标题文本,a.标题显示,a.标题字体,0,a.标题位置,a.内容字体,0,a.内容位置,a.嵌入方式,1,0,0,a.元素类型,0,a.ID,b.ID,b.编码,0" + _
                    " From 病人病历内容 a,病历元素目录 b Where a.病历记录ID=[1] And a.元素编码=b.编码(+) Order By a.排列序号"
                If blnMoved Then
                    strSQL = Replace(strSQL, "病人病历内容", "H病人病历内容")
                End If
            End If
        End If
        Set rsTmp = OpenSQLRecord(strSQL, "病历文件显示", FileID, -1 * FileID)
        If rsTmp.EOF Then
            strSQL = "Select b.部件,a.排列序号,a.填写时机,a.病历元素ID,a.必输项目,nvl(a.文本转储,0),a.标题文本,a.标题显示,a.标题字体,0,a.标题位置,a.内容字体,0,a.内容位置,a.嵌入方式,1,0,0,b.类型,0,0,a.病历元素ID,b.编码,0" + _
                " From 病历文件组成 a,病历元素目录 b Where a.病历文件ID=[1] And a.病历元素ID=b.ID" + _
                IIf(iFilter = 0, "", " And 填写时机=[2]") + " Order By a.排列序号"
            Set rsTmp = OpenSQLRecord(strSQL, "病历文件显示", FileTypeID, iFilter)
            aElement = rsTmp.GetRows()
        Else
            aElement = rsTmp.GetRows()
        End If
    End If
    objProgressBar.Value = 20
    
    Reload objProgressBar
        
    bModified = False
End Sub
'使元素滚动到屏幕可见区域
Public Sub ShowElement(ctrlElement As Control)
    Dim TopMargin As Long, BottomMargin As Long
    On Error Resume Next
    TopMargin = ctrlElement.Top + picEdit(ctrlElement.Container.Index).Top
    BottomMargin = ctrlElement.Top + ctrlElement.Height + picEdit(ctrlElement.Container.Index).Top
    
    If TopMargin + PicMain.Top < 0 Or BottomMargin + PicMain.Top > UserControl.ScaleHeight Then VSMain.Value = IIf(lblFlag(ctrlElement.Container.Index).Top > VSMain.Max, VSMain.Max, lblFlag(ctrlElement.Container.Index).Top / CInt(VSMain.Tag))
End Sub
'专用纸获取焦点后的CallBack函数，相当于事件处理函数。
Public Sub CallBack_GotFocus()
    On Error Resume Next
    If UCase(UserControl.ActiveControl.Name) Like "SPECPAPER*" Then Set CurrSpecPaper = UserControl.ActiveControl
End Sub
'插入病历元素
Public Sub InsertElement(ByVal ElementID As Long, Optional ByVal Index As Integer = 0, Optional objProgressBar As ProgressBar)
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, j As Integer, iElementNum As Integer, iFldNum As Integer
    Dim CurrControl As Control
    Dim strTitle As String
    Dim strSQL As String
    Index = Index - 1

    On Error Resume Next
    strSQL = "Select 部件,0,0,ID,0,nvl(转文本,0),名称,1,'',0,0,'',0,0,2,1,0,0,类型,0,0,ID,编码,0" + _
        " From 病历元素目录 Where ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "病历文件显示", ElementID)
        
    If rsTmp.EOF Then Exit Sub
    '病历中不能有一个以上手术概要专用纸
    If UCase(rsTmp(0)) Like "*USROPERGENERAL" Then
        iElementNum = UBound(aElement, 2)
        For i = 0 To iElementNum
            If aElement(15, i) = 1 And UCase(aElement(0, i)) Like "*USROPERGENERAL" Then
                MsgBox "一份病历中只能书写一个手术概要！", vbInformation, gstrSysName: Exit Sub
            End If
        Next
    End If
    
    strTitle = InputBox("请输入病历项目的标题，如果标题为空则不显示：", "项目标题", NVL(rsTmp(6)))
    objProgressBar.Value = 10
    
    iElementNum = UBound(aElement, 2)
    iFldNum = UBound(aElement, 1)
    ReDim Preserve aElement(iFldNum, iElementNum + 1)
    
    For i = iElementNum To Index Step -1
        For j = 0 To iFldNum
            aElement(j, i + 1) = aElement(j, i)
        Next j
    Next i
    
    For j = 0 To iFldNum
        If InStr(",7,8,10,11,13,14,", "," & CStr(j) & ",") = 0 Then aElement(j, Index) = IIf(IsNull(rsTmp(j)), "", rsTmp(j))
    Next
    '处理标题
    If Len(Trim(strTitle)) = 0 Then
        aElement(6, Index) = NVL(rsTmp(6)): aElement(7, Index) = 0
    Else
        aElement(6, Index) = strTitle: aElement(7, Index) = 1
    End If
    
    objProgressBar.Value = 20
    
    Refresh objProgressBar
    
    DoEvents '有些控件要强制设焦点。
    Set CurrControl = NextElement(Index)
    If Not CurrControl Is Nothing Then CurrControl.SetFocus
End Sub
'删除病历元素
Public Sub DeleteElement(Optional ByVal Index As Integer = 0, Optional objProgressBar As ProgressBar)
    Dim CurrControl As Control
    Index = Index - 1

    On Error Resume Next
    If UBound(aElement, 2) = -1 Then Exit Sub
'    If aElement(18, Index) < 0 And aElement(18, Index) <> -5 Then MsgBox "项目【" + aElement(6, Index) + "】为特殊元素，不允许删除", vbExclamation, gstrSysName: Exit Sub
'    If aElement(4, Index) = 1 Then MsgBox "项目【" + aElement(6, Index) + "】必须输入，不允许删除", vbExclamation, gstrSysName: Exit Sub
    If MsgBox("是否将项目【" + aElement(6, Index) + "】从病历中删除", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    
    aElement(15, Index) = 0
    
    objProgressBar.Value = 0
    
    Refresh objProgressBar
    
    DoEvents '有些控件要强制设焦点。
    Set CurrControl = NextElement(Index)
    If Not CurrControl Is Nothing Then
        CurrControl.SetFocus
    Else
        Set CurrControl = PrevElement(Index)
        If Not CurrControl Is Nothing Then CurrControl.SetFocus
    End If
End Sub
'是否以文本显示
Public Property Get IsText(ByVal Index As Integer) As Boolean
    Index = Index - 1
    IsText = False
    
    Select Case aElement(18, Index)
        Case 2
            IsText = txtVisForm(aElement(17, Index)).Visible
        Case 4
            IsText = txtSpecPaper(aElement(17, Index)).Visible
    End Select
End Property
'是否修改
Public Property Get Modified() As Boolean
    Modified = bModified
End Property

'是否修改
Public Property Let Modified(vData As Boolean)
    bModified = vData
End Property

'返回元素ID
Public Property Get ElementID(ByVal Index As Integer) As Long
    Index = Index - 1
    On Error Resume Next
    ElementID = CLng(aElement(21, Index))
End Property
'所见单显示文本
Public Function ShowText(ByVal Index As Integer, ByVal bShow As Boolean) As Boolean
    Dim iOldHeight As Long
    Dim lblCtrl As Control, txtCtrl As Control, MainCtrl As Control
    Index = Index - 1
    
    If aElement(5, Index) = 0 Then MsgBox "项目【" + aElement(6, Index) + "】不允许转文本", vbExclamation, gstrSysName: Exit Function
    
    Select Case aElement(18, Index)
        Case 2
            Set lblCtrl = lblVisForm(aElement(17, Index))
            Set txtCtrl = txtVisForm(aElement(17, Index))
            Set MainCtrl = VisForm(aElement(17, Index))
        Case 4
            Set lblCtrl = lblSpecPaper(aElement(17, Index))
            Set txtCtrl = txtSpecPaper(aElement(17, Index))
            Set MainCtrl = SpecPaper(aElement(17, Index))
        Case Else
            Exit Function
    End Select
    
    iOldHeight = picEdit(Index + 1).Height
    If bShow Then
        '改变Label的高度
        lblCtrl.Caption = Chr(255): lblCtrl = txtCtrl
        txtCtrl.Visible = True
        MainCtrl.Visible = False
        
        If lblCtrl.Height <> iOldHeight Then ExpandElement Index + 1, lblCtrl.Height - iOldHeight
        txtCtrl.SetFocus
    Else
        txtCtrl.Visible = False
        MainCtrl.Visible = True
        
        If aElement(18, Index) = 2 Then
            If aElement(16, Index) <> iOldHeight Then ExpandElement Index + 1, aElement(16, Index) - iOldHeight
        Else
            If MainCtrl.Height <> iOldHeight Then ExpandElement Index + 1, MainCtrl.Height - iOldHeight
        End If
        MainCtrl.SetFocus
    End If
    ShowText = True
End Function
'所见单转储文本
Public Sub ChangeToText(ByVal Index As Integer)
    Dim txtCtrl As Control, MainCtrl As Control
    Index = Index - 1
    
    If aElement(5, Index) = 0 Then Exit Sub
    
    On Error Resume Next
    Select Case aElement(18, Index)
        Case 2
            Set txtCtrl = txtVisForm(aElement(17, Index))
            Set MainCtrl = VisForm(aElement(17, Index))
        Case 4
            Set txtCtrl = txtSpecPaper(aElement(17, Index))
            Set MainCtrl = SpecPaper(aElement(17, Index))
        Case Else
            Exit Sub
    End Select
    
    txtCtrl = MainCtrl.Text
End Sub
'加载病历示范
Public Sub LoadSample(ByVal SampleID As Long, Optional objProgressBar As ProgressBar, Optional ifSample As Boolean = True)
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, iNum As Long, iFldNum As Long, iRecNum As Long
    Dim strSQL As String
    
    On Error Resume Next
    If ifSample Then
        strSQL = "Select b.部件,a.排列序号,0,a.ID,0,nvl(a.文本转储,0),a.标题文本,a.标题显示,a.标题字体,0,a.标题位置,a.内容字体,0,a.内容位置,a.嵌入方式,1,0,0,a.元素类型,0,Decode(b.类型,Abs(类型),a.ID,-5,a.ID,0),b.ID,b.编码,0" + _
            " From 病人病历内容 a,病历元素目录 b Where a.元素编码=b.编码(+) And a.病历示范ID=[1] Order By a.排列序号"
    Else
        strSQL = "Select b.部件,a.排列序号,0,a.ID,0,nvl(a.文本转储,0),a.标题文本,a.标题显示,a.标题字体,0,a.标题位置,a.内容字体,0,a.内容位置,a.嵌入方式,1,0,0,a.元素类型,0,Decode(b.类型,Abs(类型),a.ID,-5,a.ID,0),b.ID,b.编码,0" + _
            " From 病人病历内容 a,病历元素目录 b Where a.元素编码=b.编码(+) And a.病历记录ID=[1] Order By a.排列序号"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "病历文件显示", SampleID)
    If rsTmp.EOF Then
        MsgBox "该示范病历无内容，不能加载", vbInformation, gstrSysName
    Else
        '删除以前的元素
        objProgressBar.Value = 10
        
        iNum = UBound(aElement, 2)
        For i = 0 To iNum
            aElement(15, i) = 0
            aElement(17, i) = -1
        Next
        
        iFldNum = UBound(aElement, 1): iRecNum = 0
        Do While Not rsTmp.EOF
            ReDim Preserve aElement(iFldNum, iNum + 1): iNum = iNum + 1
            For i = 0 To iFldNum
                aElement(i, iNum) = rsTmp(i)
            Next
        
            iRecNum = iRecNum + 1
            rsTmp.MoveNext
        Loop
        objProgressBar.Value = 20
        
        Reload objProgressBar, True
        
        '清除病人病历内容ID
        For i = iNum - iRecNum + 1 To iNum
            aElement(20, i) = 0
        Next
        
        bModified = True
    End If
End Sub
'加载元素示范
Public Sub LoadElementSample(ByVal ElementIndex As Integer, ByVal SampleID As Long)
    Dim rsTmp As New ADODB.Recordset, rsID As New ADODB.Recordset
    Dim i As Integer
    Dim strTxtBox As String
    Dim iTabLeft As Long, iTabTop As Long, iTabWidth As Long, iTabHeight As Long, iShown As Integer
    Dim CtrlHeight As Long, lngCurrPos As Long
    Dim strSQL As String, lngTmpID As Long
    
    On Error Resume Next
    strSQL = "Select ID From 病人病历内容 Where 病历示范ID=[1]"
    Set rsID = OpenSQLRecord(strSQL, "加载元素示范", SampleID)
    If rsID.EOF() Then Exit Sub
    '加载病历元素
    i = ElementIndex - 1
    Select Case aElement(18, i)
        Case 0, -5
            strTxtBox = ""
            strSQL = "Select * From 病人病历文本段 Where 病历ID=[1]"
            lngTmpID = rsID(0)
            Set rsTmp = OpenSQLRecord(strSQL, "", lngTmpID)
            If Not rsTmp.EOF Then
                strTxtBox = rsTmp("内容")
                If Not bSampleFile Then strTxtBox = ReplaceString(strTxtBox, _
                    PatientID, CheckID, PatientType)
            End If
            With txtBox(aElement(17, i))
                lngCurrPos = .SelStart
                .Text = Mid(.Text, 1, .SelStart) + strTxtBox + Mid(.Text, .SelStart + .SelLength + 1)
                .SelStart = lngCurrPos + Len(strTxtBox)
                CtrlHeight = .Height
            End With
        Case 1
            With grdTable(aElement(17, i))
                ReadTable_Patient grdTable(grdTable.Count - 1), rsID(0)
                
                .RangeToTwips 1, 1, .MaxRow, .MaxCol, iTabLeft, iTabTop, iTabWidth, iTabHeight, iShown
                .Width = iTabWidth + 15
                .Height = iTabHeight + 15
                
                CtrlHeight = .Height
                aElement(16, i) = .Height
                aElement(19, i) = .Width
            End With
        Case 2
            strTxtBox = ""
            '读取内容
            strSQL = "Select * From 病人病历文本段 Where 病历ID=[1]"
            lngTmpID = rsID(0)
            Set rsTmp = OpenSQLRecord(strSQL, "", lngTmpID)
            If Not rsTmp.EOF Then
                strTxtBox = rsTmp("内容")
                If Not bSampleFile Then strTxtBox = ReplaceString(strTxtBox, _
                    PatientID, CheckID, PatientType)
            End If
            With txtVisForm(aElement(17, i))
                .Text = strTxtBox
                If .Visible Then CtrlHeight = .Height
            End With
            
            With VisForm(aElement(17, i))
                .ReadForm rsID(0), False, PatientID, CheckID, PatientType, , True, blnMoved
                
                If .Visible Then CtrlHeight = .Height
                aElement(16, i) = .Height
                aElement(19, i) = .Width
            End With
        Case 3
            Set aPicFlag(aElement(17, i)) = GetMapItems(rsID(0), blnMoved)
            
            With PicFlag(aElement(17, i))
                ShowFlagInOjbect PicFlag(aElement(17, i)), CLng(aElement(21, i)), aPicFlag(aElement(17, i)), blnMoved:=blnMoved
                
                CtrlHeight = .Height
                aElement(16, i) = .Height
                aElement(19, i) = .Width
            End With
        Case 4
            strTxtBox = ""
            '读取内容
            strSQL = "Select * From 病人病历文本段 Where 病历ID=[1]"
            lngTmpID = rsID(0)
            Set rsTmp = OpenSQLRecord(strSQL, "", lngTmpID)
            If Not rsTmp.EOF Then
                strTxtBox = rsTmp("内容")
                If Not bSampleFile Then strTxtBox = ReplaceString(strTxtBox, _
                    PatientID, CheckID, PatientType)
            End If
            With txtSpecPaper(aElement(17, i))
                .Text = strTxtBox
                If .Visible Then CtrlHeight = .Height
            End With
            
            With SpecPaper(aElement(17, i))
                .SetgcnOracle gcnOracle
                .DataMoved = blnMoved
                
                Call .SetDiagItem(SendAdviceID, SendNO)
                .ID病人病历 = rsID(0): .Get医嘱id = AdviceID
                If PatientType = 0 Then .挂号单 = CheckID

                If .Visible Then CtrlHeight = .Height
            End With
    End Select
    If CtrlHeight <> picEdit(ElementIndex).Height Then ExpandElement ElementIndex, _
        CtrlHeight - picEdit(ElementIndex).Height
End Sub
'将焦点设在指定的元素上
Public Sub SetActiveElement(ByVal Index As Integer)
    Dim tmpCtrl As Control
    On Error Resume Next
    
    Set tmpCtrl = NextElement(Index - 1)
    tmpCtrl.SetFocus
End Sub
'返回附加表元素的所见项
Public Property Get VisItem() As Object
    Set VisItem = UserControl.VisItem
End Property

Public Property Get Controls() As Object
    Set Controls = UserControl.Controls
End Property

'复制病历文本
Public Sub CopyElement(ByVal ElementIndex As Integer, ByVal SampleID As Long, Optional Comp As String = "")
    Dim rsTmp As New ADODB.Recordset, rsID As New ADODB.Recordset
    Dim i As Integer
    Dim strTxtBox As String
    Dim iTabLeft As Long, iTabTop As Long, iTabWidth As Long, iTabHeight As Long, iShown As Integer
    Dim CtrlHeight As Long, lngCurrPos As Long
    Dim strSQL As String
    
    On Error Resume Next
'    zlDatabase.OpenRecordset rsID, "Select ID From 病人病历内容 Where ID=" & SampleID, "加载元素示范"
'    If rsID.EOF() Then Exit Sub
    '加载病历元素
    i = ElementIndex - 1
    Select Case aElement(18, i)
        Case 0, -5
            strTxtBox = ""
            Select Case Comp
                Case "ZL9CISCORE.USRINSPECRESULT"
                    strSQL = "select 标题 from 病人病历所见单 where 病历id=[1] and 控件号 in (-2,-1) order by 控件号"
                    Set rsTmp = OpenSQLRecord(strSQL, "zl9CISCore", SampleID)
                    If Not rsTmp.EOF Then strTxtBox = strTxtBox + ",检验项目：" & rsTmp(0): rsTmp.MoveNext
                    If Not rsTmp.EOF Then strTxtBox = strTxtBox + ",检验标本：" & rsTmp(0) & vbCrLf
                    rsTmp.Close
                    
                    strSQL = "select A.中文名||Decode(A.英文名,'','','('||A.英文名||')')||':'||B.所见内容||B.计量单位 " + _
                        "From 病人病历所见单 B,诊治所见项目 A " + _
                        "Where B.病历id=[1] And B.所见项ID=A.Id and B.所见内容 is not null"
                    Set rsTmp = OpenSQLRecord(strSQL, "zl9CISCore", SampleID)
                    If Len(strTxtBox) = 0 Then strTxtBox = " " & vbCrLf
                    Do While Not rsTmp.EOF
                        strTxtBox = strTxtBox & rsTmp(0) & vbCrLf
                        
                        rsTmp.MoveNext
                    Loop
                    If Len(strTxtBox) > 0 Then strTxtBox = Mid(strTxtBox, 2)
                Case Else
                    strSQL = "Select * From 病人病历文本段 Where 病历ID=[1]"
                    Set rsTmp = OpenSQLRecord(strSQL, "", SampleID)
                    If Not rsTmp.EOF Then
                        strTxtBox = rsTmp("内容")
                        If Not bSampleFile Then strTxtBox = ReplaceString(strTxtBox, _
                            PatientID, CheckID, PatientType)
                    End If
            End Select
            With txtBox(aElement(17, i))
                lngCurrPos = .SelStart
                .Text = Mid(.Text, 1, .SelStart) + strTxtBox + Mid(.Text, .SelStart + .SelLength + 1)
                .SelStart = lngCurrPos + Len(strTxtBox)
                CtrlHeight = .Height
            End With
    End Select
    If CtrlHeight <> picEdit(ElementIndex).Height Then ExpandElement ElementIndex, _
        CtrlHeight - picEdit(ElementIndex).Height
End Sub

Public Property Let ifShowDiagItem(vValue As Boolean)
    bNotShowDiagItem = Not vValue
End Property

Public Sub SetDiagItem(ByVal lngDiagItem As Long, ByVal strSample As String)
    Dim iNum As Long, i As Long
    
    On Error Resume Next
    iNum = -1
    iNum = UBound(aElement, 2)

    For i = 0 To iNum
        If aElement(15, i) = 1 And aElement(0, i) Like "*SPECRESULT" Then
            SpecPaper(aElement(17, i)).ID诊疗项目 = lngDiagItem
            SpecPaper(aElement(17, i)).Cur当前标本 = strSample
        End If
    Next
End Sub
'加载元素示范
Public Sub InsertTemplate(ByVal ElementIndex As Integer, ByVal strTemplate As String)
    Dim i As Integer
    Dim strTxtBox As String
    Dim iTabLeft As Long, iTabTop As Long, iTabWidth As Long, iTabHeight As Long, iShown As Integer
    Dim CtrlHeight As Long, lngCurrPos As Long
    
    On Error Resume Next
    i = ElementIndex - 1
    If aElement(18, i) <> 0 And aElement(18, i) <> -5 Then Exit Sub
    
    With txtBox(aElement(17, i))
        lngCurrPos = .SelStart
        .SelText = strTemplate
'        .Visible = False
        .SelStart = lngCurrPos + Len(strTemplate)
        Call FormatText(.Index, .Text)
        If .SelStart = 0 Then .SelStart = Len(.Text)
'        .Visible = True
        .SetFocus
        CtrlHeight = .Height
    End With
    If CtrlHeight <> picEdit(ElementIndex).Height Then ExpandElement ElementIndex, _
        CtrlHeight - picEdit(ElementIndex).Height
End Sub
'处理RTF控件的格式
Private Sub FormatText(ByVal Index As Integer, ByVal strText As String)
    Dim iPos1 As Long, iPos2 As Long, i As Long
    Dim strItems As String, iItemSeq As Integer
    Dim aTmpItems() As String, lngFirstItem As Long
    Dim aItemAttr() As Variant  '输入项属性：1列-类型、2列-开始位置、3列－长度
    
    On Error Resume Next
    iItemSeq = 0: lngFirstItem = -1
    With txtBox(Index)
        .Text = ""
        Do While Len(strText) > 0
            iPos1 = InStr(strText, "[")
            If iPos1 = 0 Then
                .Text = .Text & strText
                strText = ""
            Else
                iPos2 = InStr(iPos1, strText, "]")
                If iPos2 = 0 Then
                    .Text = .Text & strText
                    strText = ""
                Else
                    .Text = .Text & Mid(strText, 1, iPos1 - 1)
                    strItems = Mid(strText, iPos1 + 1, iPos2 - iPos1 - 1)
                    If Len(Trim(strItems)) = 0 Then '填空
                        .Text = .Text & IIf(Len(strItems) = 0, "    ", strItems)
                        '记录输入项格式信息
                        ReDim Preserve aItemAttr(2, iItemSeq)
                        aItemAttr(0, iItemSeq) = 0
                        aItemAttr(1, iItemSeq) = Len(.Text) - IIf(Len(strItems) = 0, 4, Len(strItems))
                        aItemAttr(2, iItemSeq) = IIf(Len(strItems) = 0, 4, Len(strItems))
                        
                        If lngFirstItem = -1 Then lngFirstItem = aItemAttr(1, iItemSeq)
                        iItemSeq = iItemSeq + 1
                    Else
                        If InStr(strItems, ";") = 0 Then  '没有多个选项，按原文处理
                            .Text = .Text & "[" & strItems & "]"
                        Else '处理选择
                            aTmpItems = Split(strItems, ";")
                            
                            .Text = .Text & aTmpItems(0)
                            '记录输入项格式信息
                            ReDim Preserve aItemAttr(2, iItemSeq)
                            aItemAttr(0, iItemSeq) = 1
                            aItemAttr(1, iItemSeq) = Len(.Text) - Len(aTmpItems(0))
                            aItemAttr(2, iItemSeq) = Len(aTmpItems(0))
                            
                            If lngFirstItem = -1 Then lngFirstItem = aItemAttr(1, iItemSeq)
                            iItemSeq = iItemSeq + 1
                            
                            aTextItems(Index) = IIf(Len(aTextItems(Index)) = 0, "", aTextItems(Index) & vbCrLf) & strItems
                        End If
                    End If
                    strText = Mid(strText, iPos2 + 1)
                End If
            End If
        Loop
        '处理输入项格式
'        If iItemSeq > 0 Then
'            iItemSeq = 1
'            For i = 0 To UBound(aItemAttr, 2)
'                .SelStart = aItemAttr(1, i)
'                .SelLength = aItemAttr(2, i)
'                .SelUnderline = True
'                Select Case aItemAttr(0, i)
'                    Case 0 '填空
'                        .SelColor = 0
'                    Case 1 '下拉选择
'                        .SelColor = COLOR_COMBO Xor iItemSeq
'                        iItemSeq = iItemSeq + 1
'                End Select
'            Next
'            .SelStart = lngFirstItem + 1
'            Call SetSelect(Index)
'        End If
    End With
End Sub

'Private Sub SetSelect(ByVal Index As Integer)
''选择整个输入文本
'    Dim i As Long, lngStart As Long, lngEnd As Long, lngTmpStart As Long
'
'    blnEvent_SelChange(Index) = True
'    On Error GoTo ProcError
'
'    With txtBox(Index)
'        lngTmpStart = .SelStart
'
'        lngStart = lngTmpStart
'        For i = lngTmpStart To 1 Step -1
'            .SelStart = i
'            If Not .SelUnderline Then
'                Exit For
'            Else
'                lngStart = i
'            End If
'        Next
'        lngEnd = lngTmpStart
'        For i = lngTmpStart To Len(.Text)
'            .SelStart = i
'            If Not .SelUnderline Then
'                Exit For
'            Else
'                lngEnd = i
'            End If
'        Next
'
'        .SelStart = lngStart - 1
'        .SelLength = lngEnd - lngStart + 1
'
'        blnCurrUnderLine(Index) = True
'    End With
'    blnEvent_SelChange(Index) = False
'    Exit Sub
'ProcError:
'    blnEvent_SelChange(Index) = False
'End Sub
'打开下拉选择
Private Sub GetSelect(ByVal Index As Integer, ByVal ItemSeq As Long, ByVal Left As Single, Top As Single)
    Dim aItems() As String, rsTmp As Recordset
    Dim strSQL As String, i As Integer
    Dim lngTmpStart As Long, lngTmpLength As Long
    Dim lngItemSeq As Long

    On Error Resume Next
    
    If Len(aTextItems(Index)) = 0 Then Exit Sub
    aItems = Split(aTextItems(Index), vbCrLf)
    If UBound(aItems) < ItemSeq Then Exit Sub
    
    aItems = Split(aItems(ItemSeq), ";")
    strSQL = ""
    For i = 0 To UBound(aItems)
        strSQL = strSQL & " Union All " & "Select " & i & " As ID,'" & _
            Replace(aItems(i), "'", "''") & "' As 选项 From Dual"
    Next
    If Len(strSQL) > 0 Then
        strSQL = Mid(strSQL, 12)
        Set rsTmp = zlDatabase.ShowSelect(UserControl.Parent, strSQL, 0, blnnonewin:=True, x:=Left, y:=Top)
        
        If rsTmp Is Nothing Then Exit Sub
        
        '屏蔽SelChange事件
        blnEvent_SelChange(Index) = True
    
        With txtBox(Index)
            lngTmpStart = .SelStart: lngTmpLength = .SelLength
        
            .SelText = rsTmp("选项")
            DoEvents
            '因Change事件改变了标志，需重新屏蔽SelChange事件
            blnEvent_SelChange(Index) = True
        
            '重置选择文本
            .SelStart = lngTmpStart + Len(rsTmp("选项")): .SelLength = 0
        End With
        
        blnEvent_SelChange(Index) = False
    End If
End Sub

'Private Sub NextItem(ByVal Index As Integer)
''到下一个输入项
'    Dim i As Long, lngStart As Long, lngLength As Long
'    Dim blnLastUnderLine As Boolean
'
'    blnEvent_SelChange(Index) = True
'    On Error GoTo ProcError
'
'    With txtBox(Index)
'        lngStart = .SelStart: lngLength = .SelLength
'
'        For i = lngStart + 1 To Len(.Text)
'            blnLastUnderLine = .SelUnderline
'            .SelStart = i
'            If .SelUnderline And Not blnLastUnderLine Then Exit For
'        Next
'        If i > Len(.Text) Then
'            .SelStart = lngStart: .SelLength = lngLength
'            blnEvent_SelChange(Index) = False
'            '如果没有后续项目，则跳到下一元素
'            Call txtBox_KeyDown(Index, vbKeyReturn, vbCtrlMask)
'        Else
'            blnEvent_SelChange(Index) = False
'            Call SetSelect(Index)
'        End If
'    End With
'    Exit Sub
'ProcError:
'    blnEvent_SelChange(Index) = False
'End Sub

'添加可附加的病历文件（如病程记录、护理记录等）
Public Sub AddRecord(ByVal lngAddFileID As Long, Optional ByVal Index As Integer = 0, Optional objProgressBar As ProgressBar)
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, j As Integer, iElementNum As Integer, iFldNum As Integer
    Dim CurrControl As Control
    Dim AddNum As Integer
    Dim strSQL As String
    Index = Index - 1

    On Error Resume Next
    strSQL = "Select a.部件,b.排列序号,0,a.ID,0,nvl(b.文本转储,0),b.标题文本,b.标题显示,b.标题字体,0,b.标题位置,b.内容字体,0,b.内容位置,b.嵌入方式,1,0,0,a.类型,0,0,ID,a.编码,0" + _
        " From 病历元素目录 a,病历文件组成 b Where a.ID=b.病历元素ID And b.病历文件ID=[1] Order By b.排列序号"
    Set rsTmp = OpenSQLRecord(strSQL, "病历文件显示", lngAddFileID)
        
    If rsTmp.EOF Then Exit Sub
    AddNum = rsTmp.RecordCount
    
    objProgressBar.Value = 10
    
    iElementNum = UBound(aElement, 2)
    iFldNum = UBound(aElement, 1)
    ReDim Preserve aElement(iFldNum, iElementNum + AddNum)
    
    For i = iElementNum To Index + 1 Step -1
        For j = 0 To iFldNum
            aElement(j, i + AddNum) = aElement(j, i)
        Next j
    Next i
    
    i = Index
    Do While Not rsTmp.EOF
        i = i + 1
        For j = 0 To iFldNum
            aElement(j, i) = IIf(IsNull(rsTmp(j)), "", rsTmp(j))
        Next
        rsTmp.MoveNext
    Loop
    
    objProgressBar.Value = 20
    
    Refresh objProgressBar
    
    DoEvents '有些控件要强制设焦点。
    Set CurrControl = NextElement(Index + 1)
    If Not CurrControl Is Nothing Then CurrControl.SetFocus
End Sub
'文本内容
Public Property Get CurrentText(ByVal Index As Integer) As String
    Index = Index - 1
    CurrentText = ""
    
    Select Case aElement(18, Index)
        Case 0, -5
            CurrentText = txtBox(aElement(17, Index)).Text
    End Select
End Property

'插入文本内容
Public Sub InsertString(ByVal ElementIndex As Integer, ByVal strTemplate As String)
    Dim i As Integer
    Dim strTxtBox As String
    Dim iTabLeft As Long, iTabTop As Long, iTabWidth As Long, iTabHeight As Long, iShown As Integer
    Dim CtrlHeight As Long, lngCurrPos As Long
    
    On Error Resume Next
    i = ElementIndex - 1
    If aElement(18, i) <> 0 And aElement(18, i) <> -5 Then Exit Sub
    
    With txtBox(aElement(17, i))
        lngCurrPos = .SelStart
        .SelText = strTemplate
        .Visible = False
        .SelStart = lngCurrPos + Len(strTemplate)
        If .SelStart = 0 Then .SelStart = Len(.Text)
        .Visible = True
        .SetFocus
        CtrlHeight = .Height
    End With
    If CtrlHeight <> picEdit(ElementIndex).Height Then ExpandElement ElementIndex, _
        CtrlHeight - picEdit(ElementIndex).Height
End Sub

'更新当前已修改元素最近的签名，格式为<新签名>/<签名>
Private Sub NewSign(ByVal iElementIndex As Integer)
    Dim i As Integer
    Dim iNum As Integer
    
    On Error Resume Next
    iNum = -1
    iNum = UBound(aElement, 2)
    On Error GoTo 0
    
    For i = iElementIndex To iNum
        If aElement(18, i) = -1 Then Exit For
    Next
    If i > iNum Then Exit Sub
    
    '判断目前最新签名是否为修改人
    If Not SpecItem(aElement(17, i)).Value & "/" Like UserInfo.姓名 & "/*" Then _
        SpecItem(aElement(17, i)).Value = UserInfo.姓名 & IIf(Len(Trim(SpecItem(aElement(17, i)).Value)) = 0, "", _
            "/" & SpecItem(aElement(17, i)).Value)
End Sub

Public Sub ClearContent()
    '清楚所有病历元素的内容
    Dim i As Long, iNum As Long
    Dim lngPageID As Long, lngPatientID As Long
    
    On Error Resume Next
    iNum = -1
    iNum = UBound(aElement, 2)
    
    For i = 0 To iNum
        If aElement(15, i) = 1 Then
            Select Case aElement(18, i)
                Case 0, -5
                    txtBox(aElement(17, i)).Text = ""
                Case 1 '表格
                    With grdTable(aElement(17, i))
                        .ClearRange 1, 1, .MaxRow, .MaxCol, F1ClearValues
                    End With
                    ReadTable grdTable(aElement(17, i)), aElement(21, i)
                Case 2 '所见单
                    txtVisForm(aElement(17, i)) = ""
                    lblVisForm(aElement(17, i)) = ""
                    VisForm(aElement(17, i)).ReadForm aElement(21, i), , PatientID, CheckID, PatientType, , , blnMoved
                Case 3 '标记图
                    Set aPicFlag(aElement(17, i)) = New MapItems
                    
                    ShowFlagInOjbect PicFlag(aElement(17, i)), CLng(aElement(21, i)), aPicFlag(aElement(17, i))
                Case 4 '专用纸文本
                    txtSpecPaper(aElement(17, i)) = ""
                    lblSpecPaper(aElement(17, i)) = ""
                    
                    Call SpecPaper(aElement(17, i)).ClearData
                Case -4
                Case Else '替换域
            End Select
            '置修改标志
            aElement(23, i) = 1
        End If
    Next
End Sub


