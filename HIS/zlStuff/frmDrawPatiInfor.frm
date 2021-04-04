VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmDrawPatiInfor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "卫材病人跟踪信息"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frmDrawPatiInfor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VSFlex8Ctl.VSFlexGrid vsfInfo 
      Height          =   2295
      Left            =   930
      TabIndex        =   46
      Top             =   1920
      Visible         =   0   'False
      Width           =   7335
      _cx             =   12938
      _cy             =   4048
      Appearance      =   1
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   18
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDrawPatiInfor.frx":000C
      ScrollTrack     =   0   'False
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
      WordWrap        =   0   'False
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
   Begin VB.CheckBox chk忽略科室 
      Caption         =   "忽略病人所在科室或病区"
      Height          =   180
      Left            =   4200
      TabIndex        =   45
      Top             =   360
      Width           =   3855
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   17
      Left            =   6465
      Locked          =   -1  'True
      TabIndex        =   34
      Tag             =   "住院号"
      Top             =   3405
      Width           =   1590
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   16
      Left            =   4260
      Locked          =   -1  'True
      TabIndex        =   32
      Tag             =   "门诊号"
      Top             =   3405
      Width           =   1050
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   15
      Left            =   930
      Locked          =   -1  'True
      TabIndex        =   30
      Tag             =   "当前床号"
      Top             =   3405
      Width           =   1395
   End
   Begin VB.Frame fra 
      Height          =   60
      Index           =   1
      Left            =   30
      TabIndex        =   44
      Top             =   4590
      Width           =   9000
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   14
      Left            =   930
      Locked          =   -1  'True
      TabIndex        =   38
      Tag             =   "当前病区"
      Top             =   4245
      Width           =   7140
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   13
      Left            =   930
      Locked          =   -1  'True
      TabIndex        =   36
      Tag             =   "当前科室"
      Top             =   3840
      Width           =   7140
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   12
      Left            =   930
      Locked          =   -1  'True
      TabIndex        =   26
      Tag             =   "区域"
      Top             =   2985
      Width           =   1395
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   11
      Left            =   6465
      Locked          =   -1  'True
      TabIndex        =   24
      Tag             =   "婚姻状况"
      Top             =   2550
      Width           =   1605
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   10
      Left            =   4245
      Locked          =   -1  'True
      TabIndex        =   22
      Tag             =   "身份"
      Top             =   2550
      Width           =   1035
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   9
      Left            =   930
      Locked          =   -1  'True
      TabIndex        =   20
      Tag             =   "学历"
      Top             =   2550
      Width           =   1395
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   8
      Left            =   4245
      Locked          =   -1  'True
      TabIndex        =   16
      Tag             =   "民族"
      Top             =   2115
      Width           =   1035
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   7
      Left            =   6465
      Locked          =   -1  'True
      TabIndex        =   18
      Tag             =   "身份证号"
      Top             =   2115
      Width           =   1605
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   6
      Left            =   4245
      Locked          =   -1  'True
      TabIndex        =   28
      Tag             =   "出生地点"
      Top             =   2985
      Width           =   3825
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   5
      Left            =   930
      Locked          =   -1  'True
      TabIndex        =   14
      Tag             =   "出生日期"
      Top             =   2115
      Width           =   1395
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   4
      Left            =   6465
      Locked          =   -1  'True
      TabIndex        =   12
      Tag             =   "年龄"
      Top             =   1680
      Width           =   1605
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   3
      Left            =   4245
      Locked          =   -1  'True
      TabIndex        =   10
      Tag             =   "性别"
      Top             =   1680
      Width           =   1035
   End
   Begin VB.CommandButton cmdPati 
      Caption         =   "…"
      Height          =   300
      Left            =   3090
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1680
      Width           =   270
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   1
      Left            =   930
      MaxLength       =   20
      TabIndex        =   3
      Tag             =   "材料条码"
      Top             =   1260
      Width           =   2430
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   0
      Left            =   930
      Locked          =   -1  'True
      TabIndex        =   1
      Tag             =   "材料信息"
      Top             =   810
      Width           =   7140
   End
   Begin VB.Frame fra 
      Height          =   60
      Index           =   0
      Left            =   -30
      TabIndex        =   43
      Top             =   645
      Width           =   9000
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   90
      Picture         =   "frmDrawPatiInfor.frx":0258
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   4785
      Width           =   1100
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   6975
      TabIndex        =   40
      Top             =   4785
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   -405
      Top             =   6090
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawPatiInfor.frx":03A2
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawPatiInfor.frx":093C
            Key             =   "ItemStop"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawPatiInfor.frx":0ED6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5880
      TabIndex        =   39
      Top             =   4785
      Width           =   1100
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   2
      Left            =   930
      TabIndex        =   7
      Tag             =   "姓名"
      Top             =   1680
      Width           =   2160
   End
   Begin MSMask.MaskEdBox MakTxtEdit 
      Height          =   300
      Left            =   6480
      TabIndex        =   5
      Top             =   1260
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   10
      Format          =   "yyyy-MM-DD"
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "住院号"
      Height          =   180
      Index           =   17
      Left            =   5880
      TabIndex        =   33
      Top             =   3465
      Width           =   540
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "门诊号"
      Height          =   180
      Index           =   16
      Left            =   3675
      TabIndex        =   31
      Top             =   3465
      Width           =   540
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "当前床号"
      Height          =   180
      Index           =   15
      Left            =   150
      TabIndex        =   29
      Top             =   3465
      Width           =   720
   End
   Begin VB.Label lblEdit 
      Caption         =   "当前病区"
      Height          =   210
      Index           =   14
      Left            =   120
      TabIndex        =   37
      Top             =   4290
      Width           =   765
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "当前科室"
      Height          =   180
      Index           =   13
      Left            =   120
      TabIndex        =   35
      Top             =   3900
      Width           =   720
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "区域"
      Height          =   180
      Index           =   12
      Left            =   480
      TabIndex        =   25
      Top             =   3045
      Width           =   360
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "婚姻状况"
      Height          =   180
      Index           =   11
      Left            =   5685
      TabIndex        =   23
      Top             =   2610
      Width           =   720
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "身份"
      Height          =   180
      Index           =   10
      Left            =   3780
      TabIndex        =   21
      Top             =   2610
      Width           =   360
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "学历"
      Height          =   180
      Index           =   9
      Left            =   480
      TabIndex        =   19
      Top             =   2610
      Width           =   360
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "民族"
      Height          =   180
      Index           =   8
      Left            =   3780
      TabIndex        =   15
      Top             =   2175
      Width           =   360
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "身份证号"
      Height          =   180
      Index           =   7
      Left            =   5685
      TabIndex        =   17
      Top             =   2175
      Width           =   720
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "出生地点"
      Height          =   180
      Index           =   6
      Left            =   3420
      TabIndex        =   27
      Top             =   3045
      Width           =   720
   End
   Begin VB.Label lblEditDate 
      AutoSize        =   -1  'True
      Caption         =   "使用时间"
      Height          =   180
      Left            =   5685
      TabIndex        =   4
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "出生日期"
      Height          =   180
      Index           =   5
      Left            =   150
      TabIndex        =   13
      Top             =   2175
      Width           =   720
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "年龄"
      Height          =   180
      Index           =   4
      Left            =   6045
      TabIndex        =   11
      Top             =   1740
      Width           =   360
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "性别"
      Height          =   180
      Index           =   3
      Left            =   3780
      TabIndex        =   9
      Top             =   1740
      Width           =   360
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "病人姓名"
      Height          =   180
      Index           =   2
      Left            =   150
      TabIndex        =   6
      Top             =   1740
      Width           =   720
   End
   Begin VB.Label lblEdit 
      Caption         =   "材料条码"
      Height          =   210
      Index           =   1
      Left            =   150
      TabIndex        =   2
      Top             =   1305
      Width           =   765
   End
   Begin VB.Label lblEdit 
      Caption         =   "材料信息"
      Height          =   210
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   855
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   330
      Picture         =   "frmDrawPatiInfor.frx":2BE0
      Top             =   105
      Width           =   480
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    请设置指定卫生材料的跟踪信息."
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   705
      TabIndex        =   41
      Top             =   315
      Width           =   2970
   End
End
Attribute VB_Name = "frmDrawPatiInfor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnFirst As Boolean
Private mbln编辑 As Boolean
Private mlng收发ID As Long
Private mlng材料ID As Long                  '材料ID
Private mlng病人id As Long                  '病人ID
Private mlng当前科室ID As Long
Private mstr使用时间 As String                  '使用时间:yyyy-mm-dd
Private mstr条码 As String                  '卫生材料的条码
Private mstr姓名 As String
Private mblnChange As Boolean
Private mblnOk As Boolean
Private mstr科室类型 As String

Private Sub cmdClose_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub


Private Function MulitSelectPati(ByVal strKey As String) As Boolean
    '----------------------------------------------------------------------------------
    '功能:选择领用部门下的病人信息
    '参数:strKey-选择的病人ID(-),住院号(+),姓名
    '返回:如果选择成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/08/20
    '----------------------------------------------------------------------------------
    
    Dim strSearchKey As String, strWhere As String
    Dim rsTemp As ADODB.Recordset, blnCancel As Boolean
    Dim lngH As Long
    Dim vRect  As RECT
    
    strWhere = ""
    strSearchKey = ""
    If strKey <> "" Then
        If Not IsNumeric(Mid(strKey, 2)) Then
            strWhere = " And  a.姓名 like [2]"
            strSearchKey = GetMatchingSting(strKey)
        Else
            Select Case Mid(strKey, 1, 1)
            Case "-"  '输入的病人ID
                strWhere = " And A.病人id=[1]"
                strSearchKey = Mid(strKey, 2)
            Case "+"  '输入的住院号
                strWhere = " And B.住院号=[1]  "
                strSearchKey = Mid(strKey, 2)
            Case Else   '其他模糊查找
                strWhere = " And  a.姓名 like [2]"
                strSearchKey = GetMatchingSting(strKey)
            End Select
        End If
    End If
    
    If strKey <> "" Then
        If Mid(strKey, 1, 1) = "-" Then
            '病人信息,可能存在门诊
            gstrSQL = "" & _
                "   Select Distinct decode(m.编码,NULL,NULL,'['||m.编码||']'||m.名称) 当前病区,decode(C.编码,NULL,NULL,'['||c.编码||']'||c.名称) as 当前科室,A.病人ID As ID,A.姓名,A.性别, A.年龄,to_char(A.出生日期,'yyyy-mm-dd') 出生日期,A.民族,A.身份证号,A.学历,A.身份,A.婚姻状况," & _
                "         A.区域,A.出生地点, a.当前床号 As 当前床号,A.门诊号,A.住院号" & _
                "   From  病人信息 A,部门表 C,部门表 M " & _
                "   Where  A.当前科室ID=C.id(+) and  a.当前病区ID=M.id(+) " & _
                "           " & strWhere
                If mlng当前科室ID <> 0 And chk忽略科室.Value = 0 Then
                    '问题:13415
                    If mstr科室类型 <> "护理" Then
                        gstrSQL = gstrSQL & " And A.当前科室ID=[3]"
                    Else
                        '临床
                        gstrSQL = gstrSQL & " And A.当前病区ID=[3]"
                    End If
                End If
        Else
            '病人信息
            gstrSQL = "" & _
                "   Select Distinct decode(m.编码,NULL,NULL,'['||m.编码||']'||m.名称) 当前病区,decode(C.编码,NULL,NULL,'['||c.编码||']'||c.名称) as 当前科室,A.病人ID As ID,A.姓名,A.性别, b.年龄,to_char(A.出生日期,'yyyy-mm-dd') 出生日期,A.民族,A.身份证号,A.学历,A.身份,A.婚姻状况," & _
                "         A.区域,A.出生地点, B.出院病床 As 当前床号,A.门诊号,A.住院号" & _
                "   From 病案主页 B, 病人信息 A,部门表 C,部门表 M " & _
                "   Where A.病人id = B.病人id And A.主页id=B.主页id and B.出院科室ID=C.id(+) and  a.当前病区ID=M.id(+) and B.入院日期 Is Not Null  " & _
                "           " & strWhere
                If mlng当前科室ID <> 0 And chk忽略科室.Value = 0 Then
                    '问题:13415
                    If mstr科室类型 <> "护理" Then
                        gstrSQL = gstrSQL & " And B.出院科室ID=[3]"
                    Else
                        '临床
                        gstrSQL = gstrSQL & " And B.当前病区ID=[3]"
                    End If
                End If
        End If
    Else
        '病人信息
        gstrSQL = "" & _
            "   Select Distinct Decode(M.编码,NULL,NULL,'['||m.编码||']'||m.名称) 当前病区,decode(C.编码,NULL,NULL,'['||c.编码||']'||c.名称 ) as 当前科室,A.病人ID As ID,A.姓名,A.性别, b.年龄,to_char(A.出生日期,'yyyy-mm-dd') 出生日期,A.民族,A.身份证号,A.学历,A.身份,A.婚姻状况," & _
            "         A.区域,A.出生地点, B.出院病床 As 当前床号,A.门诊号,A.住院号" & _
            "   From 病案主页 B, 病人信息 A,部门表 C,部门表 M " & _
            "   Where A.病人id = B.病人id And A.主页id=B.主页id and B.出院科室ID=C.id(+) and  a.当前病区ID=M.id(+) And B.入院日期 Is Not Null " & _
            "          " & strWhere
                If mlng当前科室ID <> 0 And chk忽略科室.Value = 0 Then
                    '问题:13415
                    If mstr科室类型 <> "护理" Then
                        gstrSQL = gstrSQL & " And B.出院科室ID=[3]"
                    Else
                        '临床
                        gstrSQL = gstrSQL & " And B.当前病区ID=[3]"
                    End If
                End If
    End If
    vRect = zlControl.GetControlRect(txtEdit(2).hwnd)
    lngH = txtEdit(2).Height
    
'    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "病人选择器", False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strSearchKey, CStr(UCase(strSearchKey)), mlng当前科室ID)
'    If blnCancel = True Then
'        If txtEdit(2).Enabled Then txtEdit(2).SetFocus
'        Exit Function
'    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "病人信息", strSearchKey, CStr(UCase(strSearchKey)), mlng当前科室ID)
    
    If rsTemp.RecordCount = 0 Then
        ShowMsgBox "没有满足条件的病人信息,请检查!"
        If txtEdit(2).Enabled Then txtEdit(2).SetFocus
        Exit Function
    Else
        vsfInfo.Visible = True
        vsfInfo.Rows = 1
        Set vsfInfo.DataSource = rsTemp
        With vsfInfo
            Do While rsTemp.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, .ColIndex("当前病区")) = rsTemp!当前病区
                .TextMatrix(.Rows - 1, .ColIndex("当前科室")) = rsTemp!当前科室
                .TextMatrix(.Rows - 1, .ColIndex("Id")) = rsTemp!Id
                .TextMatrix(.Rows - 1, .ColIndex("姓名")) = rsTemp!姓名
                .TextMatrix(.Rows - 1, .ColIndex("性别")) = rsTemp!性别
                .TextMatrix(.Rows - 1, .ColIndex("年龄")) = rsTemp!年龄
                .TextMatrix(.Rows - 1, .ColIndex("出生日期")) = rsTemp!出生日期
                .TextMatrix(.Rows - 1, .ColIndex("民族")) = rsTemp!民族
                .TextMatrix(.Rows - 1, .ColIndex("身份证号")) = rsTemp!身份证号
                .TextMatrix(.Rows - 1, .ColIndex("学历")) = rsTemp!学历
                .TextMatrix(.Rows - 1, .ColIndex("身份")) = rsTemp!身份
                .TextMatrix(.Rows - 1, .ColIndex("婚姻状况")) = rsTemp!婚姻状况
                .TextMatrix(.Rows - 1, .ColIndex("区域")) = rsTemp!区域
                .TextMatrix(.Rows - 1, .ColIndex("出生地点")) = rsTemp!出生地点
                .TextMatrix(.Rows - 1, .ColIndex("当前床号")) = rsTemp!当前床号
                .TextMatrix(.Rows - 1, .ColIndex("门诊号")) = rsTemp!门诊号
                .TextMatrix(.Rows - 1, .ColIndex("住院号")) = rsTemp!住院号
                rsTemp.MoveNext
            Loop
        End With
    End If
'    txtEdit(2).Text = zlStr.Nvl(rsTemp!姓名)
'    cmdPati.Tag = zlStr.Nvl(rsTemp!Id)
'    mlng病人id = Val(zlStr.Nvl(rsTemp!Id))
    Dim i As Integer
'    For i = 2 To txtEdit.UBound
'        txtEdit(i).Text = zlStr.Nvl(rsTemp.Fields(txtEdit(i).Tag))
'    Next
    
    MulitSelectPati = True
End Function
   
Private Function Init病人信息() As Boolean
    '------------------------------------------------------------------------------
    '功能:初始化病人信息
    '参数:
    '返回:初始成功,返回True,否则返回False
    '编制:刘兴宏
    '日期:2007/08/20
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim i1 As Long
    
    On Error GoTo ErrHandle
    gstrSQL = "Select id,编码,名称,规格,产地 From 收费项目目录 where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng材料ID)
    If rsTemp.EOF Then
        ShowMsgBox "不存在指定的卫生材料,请检查!"
        Exit Function
    End If
    
    
    txtEdit(0).Text = "[" & zlStr.Nvl(rsTemp!编码) & "]" & zlStr.Nvl(rsTemp!名称) & Space(5) & zlStr.Nvl(rsTemp!规格) & Space(5) & zlStr.Nvl(rsTemp!产地)
    txtEdit(1).Text = mstr条码
    If mstr使用时间 = "" And mbln编辑 = True Then
        MakTxtEdit.Text = Format(sys.Currentdate, "yyyy-mm-dd")
    ElseIf mstr使用时间 = "" Then
    Else
        MakTxtEdit.Text = mstr使用时间
    End If
        
    gstrSQL = "Select ID,编码,名称 From 部门表 where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng当前科室ID)
    If rsTemp.EOF Then
        ShowMsgBox "未选择领用部门,请检查!"
        Exit Function
    End If
    For i1 = 0 To txtEdit.UBound
        If txtEdit(i1).Tag = "当前科室" Then
            txtEdit(i1).Text = "[" & zlStr.Nvl(rsTemp!编码) & "]" & zlStr.Nvl(rsTemp!名称)
            Exit For
        End If
    Next
    '加载病人信息
    If mlng病人id <> 0 Then
        
        If mlng收发ID <> 0 Then
            gstrSQL = "" & _
                "   Select Distinct A.病人ID As ID,Q.姓名,Q.性别, Q.年龄,to_char(A.出生日期,'yyyy-mm-dd') 出生日期,A.民族,A.身份证号,A.学历,A.身份,A.婚姻状况," & _
                "         A.区域,A.出生地点, Q.床号 As 当前床号,A.门诊号,b.住院号,decode(C.编码,NULL,NULL,'['||c.编码||']'||c.名称) as 当前科室,decode(M.编码,NULL,NULL,'['||m.编码||']'||m.名称) 当前病区, " & _
                "        to_char(Q.使用时间,'yyyy-mm-dd') 使用时间,Q.条码" & _
                "   From  病人信息 A,材料领用信息 Q,病案主页 B,部门表 C,部门表 M" & _
                "   Where A.病人id = Q.病人id And A.主页id=Q.主页id And q.病人id = b.病人id And q.主页id = b.主页id and Q.当前科室ID=C.id(+) and  Q.当前病区ID=M.id(+) " & _
                "           And Q.收发ID= [2] "
        Else
            gstrSQL = "" & _
                "   Select Distinct A.病人ID As ID,A.姓名,A.性别, A.年龄,to_char(A.出生日期,'yyyy-mm-dd') 出生日期,A.民族,A.身份证号,A.学历,A.身份,A.婚姻状况," & _
                "         A.区域,A.出生地点, a.当前床号 As 当前床号,A.门诊号,A.住院号,decode(C.编码,NULL,NULL,'['||c.编码||']'||c.名称) as 当前科室,decode(m.编码,NULL,NULL,'['||m.编码||']'||m.名称) 当前病区" & _
                "   From  病人信息 A,部门表 C,部门表 M" & _
                "   Where A.当前科室ID=C.id(+) and  a.当前病区ID=M.id(+) " & _
                "           And A.病人ID= [1] "
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng病人id, mlng收发ID)
        If rsTemp.EOF Then
            GoTo Init:
            Exit Function
        End If
        txtEdit(2).Text = zlStr.Nvl(rsTemp!姓名)
        cmdPati.Tag = zlStr.Nvl(rsTemp!Id)
        Dim i As Integer
        For i = 2 To txtEdit.UBound
            txtEdit(i).Text = zlStr.Nvl(rsTemp.Fields(txtEdit(i).Tag))
        Next
        If mlng收发ID <> 0 Then
            txtEdit(1).Text = zlStr.Nvl(rsTemp!条码)
            If zlStr.Nvl(rsTemp!使用时间) <> "" Then
                MakTxtEdit.Text = zlStr.Nvl(rsTemp!使用时间)
            End If
        End If
    End If
Init:
    Init病人信息 = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ISValid() As Boolean
    '------------------------------------------------------------------------------------------
    '功能:分析输入的病种是否有效
    '参数:
    '返回值:有效返回True,否则为False
    '------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim strTemp As String
        
    If Val(cmdPati.Tag) = 0 Then
        ShowMsgBox "未选择相关的病人,请检查!"
        If txtEdit(2).Enabled Then txtEdit(2).SetFocus
        Exit Function
    End If
    If MakTxtEdit.Text <> "____-__-__" Then
        If IsDate(MakTxtEdit.Text) = False Then
            ShowMsgBox "不是日期格式,请重输!"
            If MakTxtEdit.Enabled Then MakTxtEdit.SetFocus
            Exit Function
        End If
    End If
    If zlCommFun.ActualLen(txtEdit(1).Text) > txtEdit(1).MaxLength Then
        ShowMsgBox "条码不能大于" & txtEdit(1).MaxLength & " 个字符或" & txtEdit(1).MaxLength / 2 & "个汉字!"
        If txtEdit(1).Enabled Then txtEdit(1).SetFocus
        Exit Function
    End If
    
    ISValid = True
End Function

Private Sub cmdPati_Click()
        If MulitSelectPati("") = False Then
            If txtEdit(2).Enabled Then txtEdit(2).SetFocus
            Exit Sub
        End If
        OS.PressKey vbKeyTab
End Sub

Private Sub CmdSave_Click()
    '功能:保证相关的信息
   
    If ISValid() = False Then Exit Sub
        
    If MakTxtEdit.Text = "____-__-__" Then '
        mstr使用时间 = ""
    Else
        mstr使用时间 = MakTxtEdit.Text
    End If
    mlng病人id = Val(cmdPati.Tag)
    mstr姓名 = Trim(txtEdit(2).Text)
    mstr条码 = txtEdit(1).Text
    mblnOk = True
    Unload Me
End Sub
Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    If Init病人信息() = False Then Unload Me: Exit Sub
    Call SetTxtCtrColor
    mblnChange = False
End Sub
Private Sub Form_Load()
    mblnFirst = True
End Sub
 
Public Function ShowEdit(ByVal frmMain As Object, ByVal lng收发ID As Long, ByVal lng当前科室ID As Long, ByVal str科室类型 As String, ByVal lng材料ID As Long, ByVal bln编辑 As Boolean, _
    ByRef lng病人id As Long, ByRef str条码 As String, ByRef str使用时间 As String, ByRef str姓名 As String) As Boolean
    '-------------------------------------------------------------------------------------------------
    '功能:显示编辑窗口,程序入口
    '入参:frmMain-父窗口
    '     lng收发ID-收发ID(为零时,表示新增单据,无收发ID,不为零表示,修改此条收发记录中的病人相关信息)
    '     lng当前科室ID-当前科室ID
    '     lng材料id-材料ID
    '     lng病人ID -病人ID
    '     str条码-条码
    '     str使用时间
    '     bln编辑=是否可以编辑
    '出参:
    '     lng病人ID -病人ID
    '     str条码-条码
    '     str使用时间
    '     str姓名
    '返回:按确定返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/08/20
    '-------------------------------------------------------------------------------------------------
    '问题:13415
    mstr科室类型 = str科室类型
    mblnFirst = True
    mlng收发ID = lng收发ID
    mlng病人id = lng病人id
    mlng材料ID = lng材料ID
    mlng当前科室ID = lng当前科室ID
    mstr条码 = str条码
    mstr使用时间 = str使用时间
    mbln编辑 = bln编辑
    Me.Show 1, frmMain
    lng病人id = mlng病人id
    str姓名 = mstr姓名
    str条码 = mstr条码
    str使用时间 = mstr使用时间
    ShowEdit = mblnOk
    
End Function

Private Sub CtlEnableSet()
    '---------------------------------------------------------------------------------------------------------------------
    '功能:设置相关控件的Enable
    '参数:
    '编制:刘兴宏
    '日期:2007/08/20
    '---------------------------------------------------------------------------------------------------------------------
    cmdSave.Enabled = Val(cmdPati.Tag) <> 0
    
End Sub
Private Sub SetTxtCtrColor()
    '----------------------------------------------------------------------------------------------------------------------
    '功能:设置不可编辑的文本框的北景色
    '编制:刘兴宏
    '日期:2007/08/20
    '----------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To txtEdit.UBound
        If txtEdit(i).Locked Or mbln编辑 = False Then
            txtEdit(i).Enabled = False
            txtEdit(i).BackColor = &H8000000F
        End If
    Next
    cmdSave.Visible = mbln编辑
    cmdPati.Enabled = mbln编辑
    MakTxtEdit.Enabled = mbln编辑
    If mbln编辑 = False Then
        MakTxtEdit.BackColor = &H8000000F
    End If
End Sub
 
Private Sub MakTxtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txtEDIT_Change(Index As Integer)
    If txtEdit(Index).Tag = "姓名" Then
        cmdPati.Tag = ""
    End If
    mblnChange = True
End Sub

Private Sub txtEDIT_GotFocus(Index As Integer)
     zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEDIT_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Select Case txtEdit(Index).Tag
    Case "姓名"
        If MulitSelectPati(txtEdit(Index).Text) = False Then Exit Sub
        If cmdSave.Enabled Then cmdSave.SetFocus
    Case Else
        OS.PressKey vbKeyTab
    End Select
End Sub

Private Sub txtEDIT_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = Asc("|") Then
        KeyAscii = 0
    End If
End Sub
Private Sub MakTxtEdit_Validate(Cancel As Boolean)
    Dim strFormat As Date
    
    If MakTxtEdit.Text = "____-__-__" Then Exit Sub
    
    err = 0:    On Error GoTo ErrHand:
    strFormat = CDate(MakTxtEdit.Text)
    Exit Sub
ErrHand:
    MsgBox "不是日期格式，请重输入！", vbInformation, gstrSysName
    Cancel = True
End Sub

Private Sub vsfInfo_DblClick()
    With vsfInfo
        mlng病人id = .TextMatrix(.Row, .ColIndex("id"))
        txtEdit(2).Text = .TextMatrix(.Row, .ColIndex("姓名"))
        txtEdit(3).Text = .TextMatrix(.Row, .ColIndex("性别"))
        txtEdit(4).Text = .TextMatrix(.Row, .ColIndex("年龄"))
        txtEdit(5).Text = .TextMatrix(.Row, .ColIndex("出生日期"))
        txtEdit(6).Text = .TextMatrix(.Row, .ColIndex("出生地点"))
        txtEdit(7).Text = .TextMatrix(.Row, .ColIndex("身份证号"))
        txtEdit(8).Text = .TextMatrix(.Row, .ColIndex("民族"))
        txtEdit(9).Text = .TextMatrix(.Row, .ColIndex("学历"))
        txtEdit(10).Text = .TextMatrix(.Row, .ColIndex("身份"))
        txtEdit(11).Text = .TextMatrix(.Row, .ColIndex("婚姻状况"))
        txtEdit(12).Text = .TextMatrix(.Row, .ColIndex("区域"))
        txtEdit(13).Text = .TextMatrix(.Row, .ColIndex("当前科室"))
        txtEdit(14).Text = .TextMatrix(.Row, .ColIndex("当前病区"))
        txtEdit(15).Text = .TextMatrix(.Row, .ColIndex("当前床号"))
        txtEdit(16).Text = .TextMatrix(.Row, .ColIndex("门诊号"))
        txtEdit(17).Text = .TextMatrix(.Row, .ColIndex("住院号"))
        cmdPati.Tag = .TextMatrix(.Row, .ColIndex("id"))
        vsfInfo.Visible = False
    End With
End Sub

Private Sub vsfInfo_LostFocus()
    vsfInfo.Visible = False
End Sub


