VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmInsSymbol 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "特殊符号"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   Icon            =   "frmInsSymbol.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picFree 
      BorderStyle     =   0  'None
      Height          =   2130
      Left            =   495
      ScaleHeight     =   2130
      ScaleWidth      =   6360
      TabIndex        =   43
      Top             =   1005
      Width           =   6360
      Begin VB.ComboBox cboGroup 
         Height          =   300
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   0
         Width           =   3615
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgFree 
         Height          =   1785
         Left            =   0
         TabIndex        =   45
         Top             =   345
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   3149
         _Version        =   393216
         Rows            =   1
         Cols            =   15
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   2
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   15
      End
      Begin VB.Label lblGroup 
         AutoSize        =   -1  'True
         Caption         =   "字符子集(&K)"
         Height          =   180
         Left            =   0
         TabIndex        =   46
         Top             =   60
         Width           =   990
      End
   End
   Begin VB.PictureBox picCard 
      BackColor       =   &H00FFFFFF&
      Height          =   2130
      Index           =   2
      Left            =   390
      ScaleHeight     =   2070
      ScaleWidth      =   6300
      TabIndex        =   0
      Top             =   915
      Width           =   6360
      Begin VB.TextBox txtYJ 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   3225
         TabIndex        =   4
         Top             =   750
         Width           =   2220
      End
      Begin VB.TextBox txtYJ 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1890
         TabIndex        =   3
         Top             =   960
         Width           =   1170
      End
      Begin VB.TextBox txtYJ 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1890
         TabIndex        =   2
         Top             =   555
         Width           =   1170
      End
      Begin VB.TextBox txtYJ 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   765
         TabIndex        =   1
         Top             =   750
         Width           =   915
      End
      Begin VB.Line Line1 
         X1              =   1815
         X2              =   3135
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "初潮年龄"
         Height          =   180
         Index           =   0
         Left            =   840
         TabIndex        =   8
         Top             =   525
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "每次行经日数"
         Height          =   180
         Index           =   1
         Left            =   1965
         TabIndex        =   7
         Top             =   315
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "经期相隔日数"
         Height          =   180
         Index           =   2
         Left            =   2010
         TabIndex        =   6
         Tag             =   "经期相隔日数"
         Top             =   1290
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "闭经年龄(或末次停经日期)"
         Height          =   180
         Index           =   3
         Left            =   3330
         TabIndex        =   5
         Top             =   510
         Width           =   2160
      End
   End
   Begin VB.PictureBox picCard 
      BackColor       =   &H80000005&
      Height          =   2130
      Index           =   1
      Left            =   300
      ScaleHeight     =   2070
      ScaleWidth      =   6300
      TabIndex        =   12
      Tag             =   "乳牙标注"
      Top             =   825
      Width           =   6360
      Begin VB.Frame fraLineRYH 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   30
         Left            =   435
         TabIndex        =   14
         Top             =   1515
         Width           =   4065
      End
      Begin VB.Frame fraLineRYV 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1635
         Left            =   2475
         TabIndex        =   13
         Top             =   225
         Width           =   30
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshRY 
         Height          =   675
         Left            =   435
         TabIndex        =   15
         Top             =   1185
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   1191
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   16
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         BackColorBkg    =   16777215
         GridColor       =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   16
      End
      Begin VB.Label lblRYLeft 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "右"
         Height          =   180
         Left            =   210
         TabIndex        =   24
         Top             =   1440
         Width           =   180
      End
      Begin VB.Label lblRYRight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "左"
         Height          =   180
         Left            =   4590
         TabIndex        =   23
         Top             =   1440
         Width           =   180
      End
      Begin VB.Label lblRYDn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "下颌"
         Height          =   180
         Left            =   2295
         TabIndex        =   22
         Top             =   1905
         Width           =   360
      End
      Begin VB.Label lblRYUp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上颌"
         Height          =   180
         Left            =   2295
         TabIndex        =   21
         Top             =   45
         Width           =   360
      End
      Begin VB.Label lblRY 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmInsSymbol.frx":000C
         Height          =   945
         Index           =   0
         Left            =   2670
         TabIndex        =   20
         Top             =   255
         Width           =   165
      End
      Begin VB.Label lblRY 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmInsSymbol.frx":001E
         Height          =   945
         Index           =   1
         Left            =   2985
         TabIndex        =   19
         Top             =   255
         Width           =   165
      End
      Begin VB.Label lblRY 
         BackStyle       =   0  'Transparent
         Caption         =   "    乳尖牙"
         Height          =   945
         Index           =   2
         Left            =   3330
         TabIndex        =   18
         Top             =   255
         Width           =   165
      End
      Begin VB.Label lblRY 
         BackStyle       =   0  'Transparent
         Caption         =   "第一乳磨牙"
         Height          =   945
         Index           =   3
         Left            =   3660
         TabIndex        =   17
         Top             =   255
         Width           =   165
      End
      Begin VB.Label lblRY 
         BackStyle       =   0  'Transparent
         Caption         =   "第二乳磨牙"
         Height          =   945
         Index           =   4
         Left            =   4005
         TabIndex        =   16
         Top             =   255
         Width           =   165
      End
   End
   Begin VB.PictureBox picCard 
      BackColor       =   &H80000005&
      Height          =   2130
      Index           =   0
      Left            =   195
      ScaleHeight     =   2070
      ScaleWidth      =   6300
      TabIndex        =   25
      Tag             =   $"frmInsSymbol.frx":0032
      Top             =   720
      Width           =   6360
      Begin VB.Frame fraLineHYV 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1635
         Left            =   3090
         TabIndex        =   27
         Top             =   210
         Width           =   30
      End
      Begin VB.Frame fraLineHYH 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   30
         Left            =   405
         TabIndex        =   26
         Top             =   1500
         Width           =   5505
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshHY 
         Height          =   675
         Left            =   405
         TabIndex        =   28
         Top             =   1170
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   1191
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   16
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         BackColorBkg    =   16777215
         GridColor       =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   16
      End
      Begin VB.Label lblHY 
         BackStyle       =   0  'Transparent
         Caption         =   "  第三磨牙"
         Height          =   930
         Index           =   7
         Left            =   5655
         TabIndex        =   40
         Top             =   255
         Width           =   165
      End
      Begin VB.Label lblHY 
         BackStyle       =   0  'Transparent
         Caption         =   "  第二磨牙"
         Height          =   930
         Index           =   6
         Left            =   5310
         TabIndex        =   39
         Top             =   255
         Width           =   165
      End
      Begin VB.Label lblHY 
         BackStyle       =   0  'Transparent
         Caption         =   "  第一磨牙"
         Height          =   930
         Index           =   5
         Left            =   4965
         TabIndex        =   38
         Top             =   255
         Width           =   165
      End
      Begin VB.Label lblHY 
         BackStyle       =   0  'Transparent
         Caption         =   "第二前磨牙"
         Height          =   930
         Index           =   4
         Left            =   4620
         TabIndex        =   37
         Top             =   255
         Width           =   165
      End
      Begin VB.Label lblHY 
         BackStyle       =   0  'Transparent
         Caption         =   "第一前磨牙"
         Height          =   930
         Index           =   3
         Left            =   4275
         TabIndex        =   36
         Top             =   255
         Width           =   165
      End
      Begin VB.Label lblHY 
         BackStyle       =   0  'Transparent
         Caption         =   "      尖牙"
         Height          =   930
         Index           =   2
         Left            =   3945
         TabIndex        =   35
         Top             =   255
         Width           =   165
      End
      Begin VB.Label lblHY 
         BackStyle       =   0  'Transparent
         Caption         =   "    侧切牙"
         Height          =   930
         Index           =   1
         Left            =   3600
         TabIndex        =   34
         Top             =   255
         Width           =   165
      End
      Begin VB.Label lblHY 
         BackStyle       =   0  'Transparent
         Caption         =   "    中切牙"
         Height          =   930
         Index           =   0
         Left            =   3255
         TabIndex        =   33
         Top             =   255
         Width           =   165
      End
      Begin VB.Label lblHYUp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上颌"
         Height          =   180
         Left            =   2910
         TabIndex        =   32
         Top             =   45
         Width           =   360
      End
      Begin VB.Label lblHYDn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "下颌"
         Height          =   180
         Left            =   2910
         TabIndex        =   31
         Top             =   1890
         Width           =   360
      End
      Begin VB.Label lblHYRight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "左"
         Height          =   180
         Left            =   5970
         TabIndex        =   30
         Top             =   1425
         Width           =   180
      End
      Begin VB.Label lblHYLeft 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "右"
         Height          =   180
         Left            =   195
         TabIndex        =   29
         Top             =   1425
         Width           =   180
      End
   End
   Begin VB.TextBox txtChar 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   75
      TabIndex        =   11
      Top             =   3375
      Width           =   7230
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4215
      TabIndex        =   10
      Top             =   4080
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5655
      TabIndex        =   9
      Top             =   4080
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgChar 
      Height          =   2130
      Left            =   255
      TabIndex        =   41
      Top             =   990
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   3757
      _Version        =   393216
      Rows            =   6
      Cols            =   15
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      ScrollBars      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   15
   End
   Begin MSComctlLib.TabStrip tabCard 
      Height          =   3180
      Left            =   60
      TabIndex        =   42
      Top             =   105
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   5609
      MultiRow        =   -1  'True
      TabFixedWidth   =   2646
      TabFixedHeight  =   616
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   9
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   $"frmInsSymbol.frx":003F
            Key             =   "恒牙标注"
            Object.Tag             =   "恒牙标注"
            Object.ToolTipText     =   "恒牙标注"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "乳牙标注(&Y)"
            Key             =   "乳牙标注"
            Object.Tag             =   "乳牙标注"
            Object.ToolTipText     =   "乳牙标注"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "标点符号(&P)"
            Key             =   "标点符号"
            Object.Tag             =   "标点符号"
            Object.ToolTipText     =   "标点符号"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "单位符号(&U)"
            Key             =   "单位符号"
            Object.Tag             =   "单位符号"
            Object.ToolTipText     =   "单位符号"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "数字序号(&N)"
            Key             =   "数字序号"
            Object.Tag             =   "数字序号"
            Object.ToolTipText     =   "数字序号"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "数学符号(&M)"
            Key             =   "数学符号"
            Object.Tag             =   "数学符号"
            Object.ToolTipText     =   "数学符号"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "特殊符号(&S)"
            Key             =   "特殊符号"
            Object.Tag             =   "特殊符号"
            Object.ToolTipText     =   "特殊符号"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "自由选择(&F)"
            Key             =   "自由选择"
            Object.Tag             =   "自由选择"
            Object.ToolTipText     =   "自由选择"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "月经史(&J)"
            Key             =   "月经史"
            Object.Tag             =   "月经史"
            Object.ToolTipText     =   "月经史"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmInsSymbol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'月经史分数表示
Private Const YJ分子 = ""
Private Const YJ分母 = ""
Private Const YJ分数1 = _
        "" & _
        "" & _
        "" & _
        "" & _
        "" & _
        "" & _
        "" & _
        "" & _
        ""
Private Const YJ分数2 = _
        "" & _
        "" & _
        "" & _
        "" & _
        "" & _
        "" & _
        "" & _
        "" & _
        "" & _
        ""
        
'乳牙标注字符
Private Const RY分数 = ""
Private Const RY小分子 = ""
Private Const RY小分母 = ""
Private Const RY大分子 = ""
Private Const RY大分母 = ""
Private Const RY左分子 = ""
Private Const RY左分母 = ""
Private Const RY右分子 = ""
Private Const RY右分母 = ""
'恒牙标注字符
Private Const HY分数 = ""
Private Const HY小分子 = ""
Private Const HY小分母 = ""
Private Const HY大分子 = ""
Private Const HY大分母 = ""
Private Const HY左分子 = ""
Private Const HY左分母 = ""
Private Const HY右分子 = ""
Private Const HY右分母 = ""

'Word特殊符号
Private Const CON标点符号 As String = "，、。．；：？！︰…‥′‵々～‖ˇˉ﹐﹑﹒﹔﹕﹖﹗｜–︱—︳︴﹏（）︵︶｛｝︷︸〔〕︹︺【】︻︼《》︽︾〈〉︿﹀「」﹁﹂『』﹃﹄﹙﹚﹛﹜﹝﹞‘’“”〝〞ˋˊ"
Private Const CON单位符号 As String = "°′″＄￥〒￠￡％＠℃℉﹩﹪‰﹫㏕㎜㎝㎞㏎㎡㎎㎏㏄°○¤"
Private Const CON数字序号 As String = "ⅰⅱⅲⅳⅴⅵⅶⅷⅸⅹⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ⒈⒉⒊⒋⒌⒍⒎⒏⒐⒑⒒⒓⒔⒕⒖⒗⒘⒙⒚⒛⑴⑵⑶⑷⑸⑹⑺⑻⑼⑽⑾⑿⒀⒁⒂⒃⒄⒅⒆⒇①②③④⑤⑥⑦⑧⑨⑩㈠㈡㈢㈣㈤㈥㈦㈧㈨㈩"
Private Const CON数学符号 As String = "≈≡≠＝≤≥＜＞≮≯∷±＋－×÷／∫∮∝∞∧∨∑∏∪∩∈∵∴⊥∥∠⌒⊙≌∽√≦≧≒≡﹢﹣﹤﹥﹦～∟⊿㏒㏑"
Private Const CON特殊符号 As String = "＃＠＆＊※§〃№〓○●△▲◎☆★◇◆□■▽▼㊣℅ˉ￣＿﹉﹊﹍﹎﹋﹌﹟﹠﹡♀♂⊕⊙↑↓←→↖↗↙↘∥∣／＼∕﹨"
Private Const CON医学符号 As String = ""

'牙齿标注颜色
Private Const M_FLAGCOLOR = &HC0E0FF

'内部变量
Dim blnEstopMedi As Boolean     '是否禁止医学符号
Dim blnOK As Boolean

Dim intRow As Integer, intCol As Integer
Dim i As Integer, j As Integer
Dim strTemp As String

Private Sub cboGroup_Click()
    Dim intStart As Integer
    If Me.cboGroup.Visible = False Then Exit Sub
    If Me.ActiveControl.Name <> Me.cboGroup.Name Then Exit Sub
    
    intStart = 0
    For i = 0 To Me.cboGroup.ListIndex - 1
        intStart = intStart + Me.cboGroup.ItemData(i)
    Next
    
    With Me.mfgFree
        .Row = intStart \ .Cols
        .Col = intStart Mod .Cols
        .TopRow = .Row
        .SetFocus
    End With
End Sub

Private Sub fraCard_DblClick(Index As Integer)
    Dim strTemp As String
    
    Select Case Index
        Case 0
            strTemp = MakeToothString(mshHY, 8)
            If strTemp <> "" Then
                txtChar.Text = strTemp
                If txtChar.SelLength = 0 Then txtChar.SelStart = Len(txtChar.Text)
            End If
        Case 1
            strTemp = MakeToothString(mshRY, 5)
            If strTemp <> "" Then
                txtChar.Text = strTemp
                If txtChar.SelLength = 0 Then txtChar.SelStart = Len(txtChar.Text)
            End If
    End Select
End Sub

Private Sub cmdCancel_Click()
    blnOK = False: Me.Hide
End Sub

Private Sub cmdOK_Click()
    blnOK = True: Me.Hide
End Sub

Private Sub Form_Activate()
    Call tabCard_Click
End Sub

Private Sub mfgChar_DblClick()
    With Me.mfgChar
        If Trim(.Text) = "" Then Exit Sub
        Me.txtChar.Text = Me.txtChar.Text + .Text
        Me.txtChar.SelStart = Len(Me.txtChar.Text)
    End With
End Sub

Private Sub mfgChar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then Call mfgChar_DblClick
End Sub

Private Sub mfgFree_DblClick()
    With Me.mfgFree
        If Trim(.Text) = "" Then Exit Sub
        Me.txtChar.Text = Me.txtChar.Text + .Text
        Me.txtChar.SelStart = Len(Me.txtChar.Text)
    End With
End Sub

Private Sub mfgFree_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then Call mfgFree_DblClick
End Sub

Private Sub mfgFree_RowColChange()
    Dim intPoint As Integer, intStart As Integer
    With Me.mfgFree
        intPoint = .Cols * .Row + .Col + 1
    End With
    intStart = 0
    For i = 0 To Me.cboGroup.ListCount - 1
        intStart = intStart + Me.cboGroup.ItemData(i)
        If intPoint <= intStart Then Me.cboGroup.ListIndex = i: Exit Sub
    Next
End Sub

Private Sub mshHY_Click()
    If mshHY.CellBackColor = vbWhite Then
        mshHY.CellBackColor = M_FLAGCOLOR
    Else
        mshHY.CellBackColor = vbWhite
    End If
    txtChar.Text = MakeToothString(mshHY, 8)
    If txtChar.SelLength = 0 Then txtChar.SelStart = Len(txtChar.Text)
End Sub

Private Sub mshHY_EnterCell()
    mshHY.CellFontBold = True
    mshHY.CellFontUnderline = True
    mshHY.CellForeColor = vbBlue
End Sub

Private Sub mshHY_GotFocus()
    mshHY_EnterCell
End Sub

Private Sub mshHY_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then mshHY_Click
End Sub

Private Sub mshHY_LeaveCell()
    mshHY.CellFontBold = False
    mshHY.CellFontUnderline = False
    mshHY.CellForeColor = mshHY.ForeColor
End Sub

Private Sub mshHY_LostFocus()
    mshHY_LeaveCell
End Sub

Private Sub mshRY_Click()
    If mshRY.CellBackColor = vbWhite Then
        mshRY.CellBackColor = M_FLAGCOLOR
    Else
        mshRY.CellBackColor = vbWhite
    End If
    txtChar.Text = MakeToothString(mshRY, 5)
    If txtChar.SelLength = 0 Then txtChar.SelStart = Len(txtChar.Text)
End Sub

Private Sub mshRY_EnterCell()
    mshRY.CellFontBold = True
    mshRY.CellFontUnderline = True
    mshRY.CellForeColor = vbBlue
End Sub

Private Sub mshRY_GotFocus()
    mshRY_EnterCell
End Sub

Private Sub mshRY_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then mshRY_Click
End Sub

Private Sub mshRY_LeaveCell()
    mshRY.CellFontBold = False
    mshRY.CellFontUnderline = False
    mshRY.CellForeColor = mshRY.ForeColor
End Sub

Private Sub mshRY_LostFocus()
    mshRY_LeaveCell
End Sub

Private Sub tabCard_Click()
    Select Case Me.tabCard.SelectedItem.Key
    Case "恒牙标注"
    
        If Me.picCard(0).Visible = False Then Me.txtChar.Text = ""
        Me.picCard(0).Visible = True
        Me.picCard(1).Visible = False
        Me.picCard(2).Visible = False
        Me.mfgChar.Visible = False
        Me.picFree.Visible = False
        
    Case "乳牙标注"
    
        If Me.picCard(1).Visible = False Then Me.txtChar.Text = ""
        Me.picCard(0).Visible = False
        Me.picCard(1).Visible = True
        Me.picCard(2).Visible = False
        Me.mfgChar.Visible = False
        Me.picFree.Visible = False
        
    Case "标点符号", "单位符号", "数字序号", "数学符号", "特殊符号"
    
        If Me.mfgChar.Visible = False Then Me.txtChar.Text = ""
        Me.picCard(0).Visible = False
        Me.picCard(1).Visible = False
        Me.picCard(2).Visible = False
        Me.mfgChar.Visible = True
        Me.picFree.Visible = False
        
        Select Case Me.tabCard.SelectedItem.Key
        Case "标点符号"
            strTemp = CON标点符号
        Case "单位符号"
            strTemp = CON单位符号
        Case "数字序号"
            strTemp = CON数字序号
        Case "数学符号"
            strTemp = CON数学符号
        Case "特殊符号"
            strTemp = CON特殊符号 + CON医学符号
        End Select
        
        With Me.mfgChar
            .Clear
            For i = 0 To Len(strTemp) - 1
                intRow = i \ .Cols: intCol = i Mod .Cols
                .TextMatrix(intRow, intCol) = Mid(strTemp, i + 1, 1)
            Next
            If .Visible Then .SetFocus
        End With
        
    Case "自由选择"
        If Me.picFree.Visible = False Then Me.txtChar.Text = ""
        Me.picCard(0).Visible = False
        Me.picCard(1).Visible = False
        Me.picCard(2).Visible = False
        Me.mfgChar.Visible = False
        Me.picFree.Visible = True
        Me.mfgFree.SetFocus
    Case "月经史"
        
        If Me.picCard(1).Visible = False Then Me.txtChar.Text = ""
        Me.picCard(0).Visible = False
        Me.picCard(1).Visible = False
        Me.picCard(2).Visible = True
        Me.mfgChar.Visible = False
        Me.picFree.Visible = False
        Call txtYJ_Change(3)
        
    End Select
End Sub

Private Sub txtChar_Change()
    Me.cmdOK.Enabled = Me.txtChar.Text <> ""
End Sub

Private Sub txtChar_KeyPress(KeyAscii As Integer)
    If InStr("'%?&", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Function MakeToothString(objMSH As MSHFlexGrid, bytCount As Byte) As String
    '功能：根据恒牙标注，产生表示恒牙标注的特殊字符串。
    '参数：objMSH=恒牙或乳牙标注表格
    '      bytCount=单侧牙齿数
    Dim byt分子 As Byte, byt分母 As Byte
    Dim A As String, b As String, C As String, D As String 'A=上左,B=上右,C=下左,D=下右
    Dim YC分数 As String
    Dim YC小分子 As String, YC小分母 As String
    Dim YC大分子 As String, YC大分母 As String
    Dim YC左分子 As String, YC左分母 As String
    Dim YC右分子 As String, YC右分母 As String
        
    strTemp = ""
    If objMSH.Name = "mshHY" Then
        YC分数 = HY分数
        YC小分子 = HY小分子: YC小分母 = HY小分母
        YC大分子 = HY大分子: YC大分母 = HY大分母
        YC左分子 = HY左分子: YC左分母 = HY左分母
        YC右分子 = HY右分子: YC右分母 = HY右分母
    Else
        YC分数 = RY分数
        YC小分子 = RY小分子: YC小分母 = RY小分母
        YC大分子 = RY大分子: YC大分母 = RY大分母
        YC左分子 = RY左分子: YC左分母 = RY左分母
        YC右分子 = RY右分子: YC右分母 = RY右分母
    End If
            
    '求ABCD四个方向的标注情况,以中心开始编齿号,如"37"
    objMSH.Redraw = False
    intRow = objMSH.Row: intCol = objMSH.Col
    
    objMSH.Row = 0
    For i = bytCount To 1 Step -1
        objMSH.Col = i - 1
        If objMSH.CellBackColor = M_FLAGCOLOR Then A = A & bytCount + 1 - i
    Next
    For i = bytCount + 1 To bytCount * 2
        objMSH.Col = i - 1
        If objMSH.CellBackColor = M_FLAGCOLOR Then b = b & i - bytCount
    Next
    
    objMSH.Row = 1
    For i = bytCount To 1 Step -1
        objMSH.Col = i - 1
        If objMSH.CellBackColor = M_FLAGCOLOR Then C = C & bytCount + 1 - i
    Next
    For i = bytCount + 1 To bytCount * 2
        objMSH.Col = i - 1
        If objMSH.CellBackColor = M_FLAGCOLOR Then D = D & i - bytCount
    Next
    
    objMSH.Row = intRow: objMSH.Col = intCol
    objMSH.Redraw = True
    
    '根据不同的给合情况，产生标注特殊字符串
    If A <> "" And b = "" And C = "" And D = "" Then
        '只有左上标注
        For i = Len(A) To 1 Step -1
            If i = 1 Then
                strTemp = strTemp & Mid(YC左分子, CByte(Mid(A, i, 1)), 1)
            Else
                strTemp = strTemp & Mid(YC大分子, CByte(Mid(A, i, 1)), 1)
            End If
        Next
    ElseIf A = "" And b <> "" And C = "" And D = "" Then
        '只有右上标注
        For i = 1 To Len(b)
            If i = 1 Then
                strTemp = strTemp & Mid(YC右分子, CByte(Mid(b, i, 1)), 1)
            Else
                strTemp = strTemp & Mid(YC大分子, CByte(Mid(b, i, 1)), 1)
            End If
        Next
    ElseIf A = "" And b = "" And C <> "" And D = "" Then
        '只有左下标注
        For i = Len(C) To 1 Step -1
            If i = 1 Then
                strTemp = strTemp & Mid(YC左分母, CByte(Mid(C, i, 1)), 1)
            Else
                strTemp = strTemp & Mid(YC大分母, CByte(Mid(C, i, 1)), 1)
            End If
        Next
    ElseIf A = "" And b = "" And C = "" And D <> "" Then
        '只有右下标注
        For i = 1 To Len(D)
            If i = 1 Then
                strTemp = strTemp & Mid(YC右分母, CByte(Mid(D, i, 1)), 1)
            Else
                strTemp = strTemp & Mid(YC大分母, CByte(Mid(D, i, 1)), 1)
            End If
        Next
    ElseIf A <> "" And b <> "" And C = "" And D = "" Then
        '只有上左右有标注
        For i = Len(A) To 1 Step -1
            strTemp = strTemp & Mid(YC大分子, CByte(Mid(A, i, 1)), 1)
        Next
        strTemp = strTemp & ""
        For i = 1 To Len(b)
            strTemp = strTemp & Mid(YC大分子, CByte(Mid(b, i, 1)), 1)
        Next
    ElseIf A = "" And b = "" And C <> "" And D <> "" Then
        '只有下左右有标注
        For i = Len(C) To 1 Step -1
            strTemp = strTemp & Mid(YC大分母, CByte(Mid(C, i, 1)), 1)
        Next
        strTemp = strTemp & ""
        For i = 1 To Len(D)
            strTemp = strTemp & Mid(YC大分母, CByte(Mid(D, i, 1)), 1)
        Next
    ElseIf A <> "" And b = "" And C = "" And D <> "" Then
        '只有左上右下有标注
        For i = Len(A) To 1 Step -1
            strTemp = strTemp & Mid(YC小分子, CByte(Mid(A, i, 1)), 1)
        Next
        strTemp = strTemp & ""
        For i = 1 To Len(D)
            strTemp = strTemp & Mid(YC小分母, CByte(Mid(D, i, 1)), 1)
        Next
    ElseIf A = "" And b <> "" And C <> "" And D = "" Then
        '只有右上左下有标注
        For i = Len(C) To 1 Step -1
            strTemp = strTemp & Mid(YC小分母, CByte(Mid(C, i, 1)), 1)
        Next
        strTemp = strTemp & ""
        For i = 1 To Len(b)
            strTemp = strTemp & Mid(YC小分子, CByte(Mid(b, i, 1)), 1)
        Next
    ElseIf Not (A = "" And b = "" And C = "" And D = "") Then
        '上下都有标注
        If A = "" And C = "" Then strTemp = ""
        
        '求左边分数串
        i = 1: j = 1 'i对应A,j对应C
        Do While i <= Len(A) Or j <= Len(C)
            byt分子 = 0: byt分母 = 0
            If i <= Len(A) Then byt分子 = Mid(A, i, 1)
            If j <= Len(C) Then byt分母 = Mid(C, j, 1)
            '根据分子分母求一个分数特殊符号
            If byt分子 <> 0 And byt分母 <> 0 Then
                strTemp = strTemp & Mid(YC分数, (byt分母 - 1) * bytCount + byt分子, 1)
            ElseIf byt分子 <> 0 And byt分母 = 0 Then
                strTemp = strTemp & Mid(YC小分子, byt分子, 1)
            ElseIf byt分子 = 0 And byt分母 <> 0 Then
                strTemp = strTemp & Mid(YC小分母, byt分母, 1)
            End If
            i = i + 1: j = j + 1
        Loop
        strTemp = StrReverse(strTemp)
        
        '连接符
        If (A <> "" Or C <> "") And (b <> "" Or D <> "") Then
            strTemp = strTemp & ""
        ElseIf b = "" And D = "" Then
            strTemp = strTemp & ""
        End If
        
        '求右边分数串
        i = 1: j = 1 'i对应B,j对应D
        Do While i <= Len(b) Or j <= Len(D)
            byt分子 = 0: byt分母 = 0
            If i <= Len(b) Then byt分子 = Mid(b, i, 1)
            If j <= Len(D) Then byt分母 = Mid(D, j, 1)
            '根据分子分母求一个分数特殊符号
            If byt分子 <> 0 And byt分母 <> 0 Then
                strTemp = strTemp & Mid(YC分数, (byt分母 - 1) * bytCount + byt分子, 1)
            ElseIf byt分子 <> 0 And byt分母 = 0 Then
                strTemp = strTemp & Mid(YC小分子, byt分子, 1)
            ElseIf byt分子 = 0 And byt分母 <> 0 Then
                strTemp = strTemp & Mid(YC小分母, byt分母, 1)
            End If
            i = i + 1: j = j + 1
        Loop
    End If
    MakeToothString = strTemp
End Function

Public Function ShowMe(Optional ByVal bytSex As Byte = 0) As String
    '功能：显示本对话框
    '参数：
    '   EstopMedi,是否禁止医学符号
    
    Dim intLoop As Integer
    
    '恒牙标注
    mshHY.Rows = 2: mshHY.Cols = 16
    mshHY.Height = mshHY.RowHeightMin * mshHY.Rows - 30
    mshHY.Width = mshHY.RowHeightMin * mshHY.Cols - 90
    mshHY.Left = (mshHY.Container.Width - mshHY.Width) / 2
    For i = 0 To mshHY.Cols - 1
        mshHY.ColWidth(i) = mshHY.RowHeight(0)
        mshHY.ColAlignment(i) = 4
        If i + 1 <= 8 Then
            mshHY.TextMatrix(0, i) = 8 - ((i + 1) Mod 9) + 1
            mshHY.TextMatrix(1, i) = 8 - ((i + 1) Mod 9) + 1
        Else
            mshHY.TextMatrix(0, i) = (i - 7) Mod 9
            mshHY.TextMatrix(1, i) = (i - 7) Mod 9
        End If
    Next
    fraLineHYH.Left = mshHY.Left
    fraLineHYH.Top = mshHY.Top + (mshHY.Height - fraLineHYH.Height) / 2
    fraLineHYH.Width = mshHY.Width
    fraLineHYV.Left = mshHY.Left + (mshHY.Width - fraLineHYV.Width) / 2
    
    For i = 0 To 7
        lblHY(i).Left = fraLineHYV.Left + (mshHY.ColWidth(0) - lblHY(i).Width) / 2 + mshHY.ColWidth(0) * i
    Next
    lblHYLeft.Top = fraLineHYH.Top - lblHYLeft.Height / 2
    lblHYLeft.Left = fraLineHYH.Left - lblHYLeft.Width - 60
    lblHYRight.Top = lblHYLeft.Top
    lblHYRight.Left = fraLineHYH.Left + fraLineHYH.Width + 60
    lblHYUp.Left = fraLineHYV.Left - lblHYUp.Width / 2
    lblHYUp.Top = fraLineHYV.Top - lblHYUp.Height - 30
    lblHYDn.Left = lblHYUp.Left
    lblHYDn.Top = mshHY.Top + mshHY.Height + 60
    mshHY.Row = 0: mshHY.Col = 8
    
    '乳牙标注
    mshRY.Rows = 2: mshRY.Cols = 10
    mshRY.Height = mshRY.RowHeightMin * mshRY.Rows - 30
    mshRY.Width = mshRY.RowHeightMin * mshRY.Cols - 90
    mshRY.Left = (mshRY.Container.Width - mshRY.Width) / 2
    
    mshRY.TextMatrix(0, 0) = "Ⅴ"
    mshRY.TextMatrix(0, 1) = "Ⅳ"
    mshRY.TextMatrix(0, 2) = "Ⅲ"
    mshRY.TextMatrix(0, 3) = "Ⅱ"
    mshRY.TextMatrix(0, 4) = "Ⅰ"
    For i = 0 To mshRY.Cols - 1
        mshRY.ColWidth(i) = mshRY.RowHeight(0)
        mshRY.ColAlignment(i) = 4
        
        If i >= 5 Then mshRY.TextMatrix(0, i) = mshRY.TextMatrix(0, mshRY.Cols - i - 1)
        mshRY.TextMatrix(1, i) = mshRY.TextMatrix(0, i)
    Next
    
    fraLineRYH.Left = mshRY.Left
    fraLineRYH.Top = mshRY.Top + (mshRY.Height - fraLineRYH.Height) / 2
    fraLineRYH.Width = mshRY.Width
    fraLineRYV.Left = mshRY.Left + (mshRY.Width - fraLineRYV.Width) / 2
    
    For i = 0 To 4
        lblRY(i).Left = fraLineRYV.Left + (mshRY.ColWidth(0) - lblRY(i).Width) / 2 + mshRY.ColWidth(0) * i
    Next
    lblRYLeft.Top = fraLineRYH.Top - lblRYLeft.Height / 2
    lblRYLeft.Left = fraLineRYH.Left - lblRYLeft.Width - 60
    lblRYRight.Top = lblRYLeft.Top
    lblRYRight.Left = fraLineRYH.Left + fraLineRYH.Width + 60
    lblRYUp.Left = fraLineRYV.Left - lblRYUp.Width / 2
    lblRYUp.Top = fraLineRYV.Top - lblRYUp.Height - 30
    lblRYDn.Left = lblRYUp.Left
    lblRYDn.Top = mshRY.Top + mshRY.Height + 60
    mshRY.Row = 0: mshRY.Col = 5
    
    'Word特殊符号网格
    With Me.mfgChar
        For i = 0 To .Rows - 1
            .RowHeight(i) = (.Height - 90) / .Rows
        Next
        For i = 0 To .Cols - 1
            .ColWidth(i) = (.Width - 150) / .Cols
            .ColAlignment(i) = 4
        Next
    End With
    
    '所有标准字符
    Dim aryFree(28, 1) As String
    aryFree(0, 0) = "基本拉丁语": aryFree(0, 1) = " !" & Chr(34) & "#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~"
    aryFree(1, 0) = "拉丁语-1和扩充": aryFree(1, 1) = "¤§¨°±·×àáèéêìíòó÷ùúüāēěīńňōūǎǐǒǔǖǘǚǜ"
    aryFree(2, 0) = "国际音标扩充": aryFree(2, 1) = "ɑɡ"
    aryFree(3, 0) = "进格修饰字符": aryFree(3, 1) = "ˇˉˊˋ˙"
    aryFree(4, 0) = "基本希腊语": aryFree(4, 1) = "ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩαβγδεζηθικλμνξοπρστυφχψω"
    aryFree(5, 0) = "西里尔文": aryFree(5, 1) = "ЁАБВГДЕЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯабвгдежзийклмнопрстуфхцчшщъыьэюяё"
    aryFree(6, 0) = "广义标点": aryFree(6, 1) = "‐–—―‖‘’“”‥…‰′″‵※"
    aryFree(7, 0) = "货币符号": aryFree(7, 1) = "�"
    aryFree(8, 0) = "类似字母的符号": aryFree(8, 1) = "℃℅℉№℡"
    aryFree(9, 0) = "数字形式": aryFree(9, 1) = "ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫⅰⅱⅲⅳⅴⅵⅶⅷⅸⅹ"
    aryFree(10, 0) = "箭头": aryFree(10, 1) = "←↑→↓↖↗↘↙"
    aryFree(11, 0) = "数学运算符": aryFree(11, 1) = "∈∏∑∕√∝∞∟∠∣∥∧∨∩∪∫∮∴∵∶∷∽≈≌≒≠≡≤≥≧≮≯⊕⊙⊥⊿"
    aryFree(12, 0) = "零杂技术用符号": aryFree(12, 1) = "⌒"
    aryFree(13, 0) = "带括号的字母数字": aryFree(13, 1) = "①②③④⑤⑥⑦⑧⑨⑩⑴⑵⑶⑷⑸⑹⑺⑻⑼⑽⑾⑿⒀⒁⒂⒃⒄⒅⒆⒇⒈⒉⒊⒋⒌⒍⒎⒏⒐⒑⒒⒓⒔⒕⒖⒗⒘⒙⒚⒛"
    aryFree(14, 0) = "制表符": aryFree(14, 1) = "─━│┃┄┅┆┇┈┉┊┋┌┍┎┏┐┑┒┓└┕┖┗┘┙┚┛├┝┞┟┠┡┢┣┤┥┦┧┨┩┪┫┬┭┮┯┰┱┲┳┴┵┶┷┸┹┺┻┼┽┾┿╀╁╂╃╄╅╆╇╈╉╊╋═║╒╓╔╕╖╗╘╙╚╛╜╝╞╟╠╡╢╣╤╥╦╧╨╩╪╫╬╭╮╯╰╱╲╳"
    aryFree(15, 0) = "方块元素": aryFree(15, 1) = "▁▂▃▄▅▆▇█▉▊▋▌▍▎▏▓▔▕"
    aryFree(16, 0) = "几何图形符": aryFree(16, 1) = "■□▲△▼▽◆◇○◎●◢◣◤◥"
    aryFree(17, 0) = "零杂丁贝符(示意符等)": aryFree(17, 1) = "★☆♀♂"
    aryFree(18, 0) = "CJK符号和标点": aryFree(18, 1) = "、。〃々〆〇〈〉《》「」『』【】〒〓〔〕〖〗〝〞〡〢〣〤〥〦〧〨〩"
    aryFree(19, 0) = "平假名": aryFree(19, 1) = "ぁあぃいぅうぇえぉおかがきぎくぐけげこごさざしじすずせぜそぞただちぢっつづてでとどなにぬねのはばぱひびぴふぶぷへべぺほぼぽまみむめもゃやゅゆょよらりるれろゎわゐゑをん゛゜ゝゞ"
    aryFree(20, 0) = "片假名": aryFree(20, 1) = "ァアィイゥウェエォオカガキギクグケゲコゴサザシジスズセゼソゾタダチヂッツヅテデトドナニヌネノハバパヒビピフブプベペホボポマミムメモャヤュユョヨラリルレロヮワヰヱヲンヴヵヶーヽヾ"
    aryFree(21, 0) = "注音": aryFree(21, 1) = "ㄅㄆㄇㄈㄉㄊㄋㄌㄍㄎㄏㄐㄑㄒㄓㄔㄕㄖㄗㄘㄙㄚㄛㄜㄝㄞㄟㄠㄡㄢㄣㄤㄥㄦㄧㄨㄩ"
    aryFree(22, 0) = "带括号的CJK字母和月份": aryFree(22, 1) = "㈠㈡㈢㈣㈤㈥㈦㈧㈨㈩㈱㊣"
    aryFree(23, 0) = "CJK兼容字符": aryFree(23, 1) = "㎎㎏㎜㎝㎞㎡㏄㏎㏑㏒㏕"
    aryFree(24, 0) = "CJK兼容形式": aryFree(24, 1) = "︰︱︳︴︵︶︷︸︹︺︻︼︽︾︿﹀﹁﹂﹃﹄﹉﹊﹋﹌﹍﹎﹏"
    aryFree(25, 0) = "小写变体": aryFree(25, 1) = "﹐﹑﹒﹔﹕﹖﹗﹙﹚﹛﹜﹝﹞﹟﹠﹡﹢﹣﹤﹥﹦﹨﹩﹪﹫"
    aryFree(26, 0) = "半行及全形字符": aryFree(26, 1) = "！" & Chr(-23646) & "＃＄％＆＇（）＊＋，－．／０１２３４５６７８９：；＜＝＞？＠ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ［＼］＾＿｀ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ｛｜｝～￠￡￢￣￤￥"
    aryFree(27, 0) = "分数字符": aryFree(27, 1) = ""

    With Me.mfgFree
        For i = 0 To .Cols - 1
            .ColWidth(i) = (.Width - 150 - 200) / .Cols
            .ColAlignment(i) = 4
        Next
        .RowHeight(0) = (.Height - 90) / 5
    End With
    
    intRow = 0: intCol = 0
    cboGroup.Clear
    For i = 0 To UBound(aryFree) - 1
        Me.cboGroup.AddItem aryFree(i, 0)
        Me.cboGroup.ItemData(Me.cboGroup.NewIndex) = Len(aryFree(i, 1))
        For j = 0 To Len(aryFree(i, 1)) - 1
            Me.mfgFree.TextMatrix(intRow, intCol) = Mid(aryFree(i, 1), j + 1, 1)
            intCol = intCol + 1
            If intCol = Me.mfgFree.Cols Then
                intRow = intRow + 1: intCol = 0
                If intRow >= Me.mfgFree.Rows - 1 Then
                    Me.mfgFree.Rows = Me.mfgFree.Rows + 1
                    Me.mfgFree.RowHeight(Me.mfgFree.Rows - 1) = Me.mfgFree.RowHeight(0)
                End If
            End If
        Next
    Next
    Me.cboGroup.ListIndex = 0
    If bytSex = 1 Then
        '男性时隐藏月经史
        For intLoop = 1 To Me.tabCard.Tabs.Count
            If Me.tabCard.Tabs(intLoop).Key = "月经史" Then
                Me.tabCard.Tabs.Remove "月经史"
                Exit For
            End If
        Next
    Else
        
        For intLoop = 1 To Me.tabCard.Tabs.Count
            If Me.tabCard.Tabs(intLoop).Key = "月经史" Then
                
                Exit For
            End If
        Next
        
        If intLoop > Me.tabCard.Tabs.Count Then
            Me.tabCard.Tabs.Add 9, "月经史", "月经史(&J)"
            Me.tabCard.Tabs("月经史").Tag = "月经史"
            Me.tabCard.Tabs("月经史").ToolTipText = "月经史"
        End If
    End If
    
    Call tabCard_Click
    
    Call txtYJ_Change(0)
    
    blnOK = False
    Me.Show vbModal
    If blnOK = False Then Unload Me: Exit Function
    ShowMe = Trim(Me.txtChar.Text): Unload Me
End Function

Private Sub txtYJ_Change(Index As Integer)
    If Visible Then
        txtChar.Text = MakeYJString
        If txtChar.SelLength = 0 Then txtChar.SelStart = Len(txtChar.Text)
    End If
End Sub

Private Sub txtYJ_DblClick(Index As Integer)
    txtYJ_Change Index
End Sub

Private Sub txtYJ_GotFocus(Index As Integer)
    If txtYJ(Index).Text = txtYJ(Index).ToolTipText Then
        'txtYJ(Index).Text = ""
    End If
    
    With txtYJ(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtYJ_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtYJ_LostFocus(Index As Integer)
    If Index = 3 Then
        If Not (IsNumeric(txtYJ(Index).Text) Or IsDate(txtYJ(Index).Text)) Then
            txtYJ(Index).Text = txtYJ(Index).ToolTipText
        End If
    Else
        If Not IsNumeric(txtYJ(Index).Text) Then
            txtYJ(Index).Text = txtYJ(Index).ToolTipText
        End If
    End If
End Sub

Private Function MakeYJString() As String
'功能：根据月经史填写的内容生成特殊字符标注串
    Dim str分子 As String, str分母 As String
    Dim strTmp As String
    
    If Not (IsNumeric(txtYJ(1).Text) And IsNumeric(txtYJ(2).Text)) Then Exit Function
    
    '求分数部分：数字向右对齐
    '------------------------
    str分子 = Right(Format(Int(txtYJ(1).Text), "00"), 2)
    str分母 = Right(Format(Int(txtYJ(2).Text), "00"), 2)
    
    '求10位的字符
    If Val(Left(str分母, 1)) <> 0 Or Val(Left(str分子, 1)) <> 0 Then
        If Val(Left(str分母, 1)) <> 0 And Val(Left(str分子, 1)) <> 0 Then
            strTmp = Mid(YJ分数1, (Val(Left(str分母, 1)) - 1) * 10 + Val(Left(str分子, 1)) + 1, 1)
        ElseIf Val(Left(str分子, 1)) = 0 Then
            strTmp = Mid(YJ分母, Val(Left(str分母, 1)) + 1, 1)
        ElseIf Val(Left(str分母, 1)) = 0 Then
            strTmp = Mid(YJ分子, Val(Left(str分子, 1)) + 1, 1)
        End If
    End If
        
    '求个位的字符
    strTmp = strTmp & Mid(YJ分数2, Val(Right(str分母, 1)) * 10 + Val(Right(str分子, 1)) + 1, 1)
        
    '组合其它字符
    If IsNumeric(txtYJ(0).Text) Then
        strTmp = txtYJ(0).Text & strTmp
    End If
    If IsNumeric(txtYJ(3).Text) Or IsDate(txtYJ(3).Text) Then
        strTmp = strTmp & txtYJ(3).Text
    End If
    
    MakeYJString = strTmp
End Function


