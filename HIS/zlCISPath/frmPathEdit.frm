VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPathEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "临床路径信息"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "frmPathEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   735
      Left            =   4800
      TabIndex        =   38
      Top             =   620
      Width           =   2160
      Begin VB.OptionButton opt首要路径 
         Caption         =   "首要路径"
         Height          =   195
         Left            =   600
         TabIndex        =   6
         Top             =   120
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.OptionButton opt合并路径 
         Caption         =   "合并路径"
         Height          =   195
         Left            =   600
         TabIndex        =   7
         Top             =   480
         Width           =   1020
      End
   End
   Begin VB.CheckBox chk诊断不同允许正常完成 
      Caption         =   "结束路径时，如果出院诊断不在适用病种范围内，允许选择正常完成。"
      Height          =   255
      Left            =   360
      TabIndex        =   37
      Top             =   6960
      Width           =   5895
   End
   Begin VB.TextBox txtConfirmDay 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   200
      IMEMode         =   2  'OFF
      Left            =   1245
      MaxLength       =   2
      TabIndex        =   35
      Text            =   "0"
      Top             =   6645
      Width           =   300
   End
   Begin VB.CheckBox chk连续 
      Caption         =   "连续增加"
      Height          =   195
      Left            =   360
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   7430
      Width           =   1020
   End
   Begin VB.ComboBox cbo适用性别 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2540
      Width           =   1740
   End
   Begin VB.Frame Frame4 
      Height          =   75
      Left            =   -105
      TabIndex        =   33
      Top             =   7215
      Width           =   7200
   End
   Begin VB.Frame Frame2 
      Height          =   75
      Left            =   225
      TabIndex        =   32
      Top             =   1980
      Width           =   6870
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   210
      TabIndex        =   31
      Top             =   540
      Width           =   6885
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5760
      TabIndex        =   40
      Top             =   7352
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4620
      TabIndex        =   39
      Top             =   7352
      Width           =   1100
   End
   Begin VB.OptionButton opt应用范围 
      Caption         =   "指定科室"
      Height          =   195
      Index           =   1
      Left            =   1560
      TabIndex        =   22
      Top             =   3360
      Value           =   -1  'True
      Width           =   1020
   End
   Begin VB.OptionButton opt应用范围 
      Caption         =   "全院通用"
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   21
      Top             =   3000
      Width           =   1020
   End
   Begin VSFlex8Ctl.VSFlexGrid vsDept 
      Height          =   840
      Left            =   3690
      TabIndex        =   23
      Top             =   2955
      Width           =   3195
      _cx             =   5636
      _cy             =   1482
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   0   'False
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPathEdit.frx":058A
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
      Editable        =   2
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
   Begin VB.TextBox txt说明 
      Height          =   510
      Left            =   960
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   1430
      Width           =   5920
   End
   Begin VB.ComboBox cbo年龄单位 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   5805
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2540
      Width           =   720
   End
   Begin VB.TextBox txt适用年龄 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   5415
      MaxLength       =   3
      TabIndex        =   18
      Top             =   2540
      Width           =   360
   End
   Begin VB.TextBox txt适用年龄 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   4875
      MaxLength       =   3
      TabIndex        =   17
      Top             =   2540
      Width           =   360
   End
   Begin VB.ComboBox cbo适用病情 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   4875
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2160
      Width           =   1650
   End
   Begin VB.ComboBox cbo病例分型 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2160
      Width           =   1740
   End
   Begin VB.ComboBox cbo分类 
      Height          =   300
      Left            =   960
      TabIndex        =   1
      Top             =   690
      Width           =   1695
   End
   Begin VB.TextBox txt名称 
      Height          =   300
      Left            =   960
      MaxLength       =   100
      TabIndex        =   5
      Top             =   1050
      Width           =   3855
   End
   Begin VB.TextBox txt编码 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3615
      MaxLength       =   5
      TabIndex        =   3
      Top             =   690
      Width           =   1200
   End
   Begin VSFlex8Ctl.VSFlexGrid vsDisease 
      Height          =   2385
      Index           =   0
      Left            =   300
      TabIndex        =   25
      Top             =   4200
      Width           =   3315
      _cx             =   5847
      _cy             =   4207
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPathEdit.frx":05BE
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
      Editable        =   2
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
   Begin VSFlex8Ctl.VSFlexGrid vsDisease 
      Height          =   2385
      Index           =   1
      Left            =   3690
      TabIndex        =   27
      Top             =   4200
      Width           =   3210
      _cx             =   5662
      _cy             =   4207
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPathEdit.frx":0623
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
      Editable        =   2
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
   Begin VB.Label Lbl跳转接收病种 
      AutoSize        =   -1  'True
      Caption         =   "跳转接收病种(&J)"
      Height          =   180
      Left            =   3720
      TabIndex        =   26
      Top             =   3960
      Width           =   1350
   End
   Begin VB.Label lblConfirm 
      Caption         =   "入院时间超过确诊天数后不允许导入临床路径。"
      Height          =   255
      Left            =   2040
      TabIndex        =   36
      Top             =   6705
      Width           =   3855
   End
   Begin VB.Label lblConfirmDay 
      Caption         =   "确诊天数：___ 天"
      Height          =   255
      Left            =   360
      TabIndex        =   34
      Top             =   6705
      Width           =   1455
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "临床路径的基本信息，适用对象，应用范围，对应病种等设置"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   720
      TabIndex        =   30
      Top             =   195
      Width           =   4860
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   150
      Picture         =   "frmPathEdit.frx":0688
      Top             =   50
      Width           =   480
   End
   Begin VB.Label lbl对应病种 
      AutoSize        =   -1  'True
      Caption         =   "导入适用病种(&I)"
      Height          =   180
      Left            =   285
      TabIndex        =   24
      Top             =   3960
      Width           =   1350
   End
   Begin VB.Label lbl应用范围 
      AutoSize        =   -1  'True
      Caption         =   "应用范围(&S)"
      Height          =   180
      Left            =   285
      TabIndex        =   20
      Top             =   2985
      Width           =   990
   End
   Begin VB.Label lbl说明 
      AutoSize        =   -1  'True
      Caption         =   "说明(&N)"
      Height          =   180
      Left            =   255
      TabIndex        =   8
      Top             =   1480
      Width           =   630
   End
   Begin VB.Label lbl年龄范围 
      AutoSize        =   -1  'True
      Caption         =   "-"
      Height          =   180
      Left            =   5280
      TabIndex        =   29
      Top             =   2595
      Width           =   90
   End
   Begin VB.Label lbl适用年龄 
      AutoSize        =   -1  'True
      Caption         =   "适用年龄(&Y)"
      Height          =   180
      Left            =   3735
      TabIndex        =   16
      Top             =   2595
      Width           =   990
   End
   Begin VB.Label lbl适用性别 
      AutoSize        =   -1  'True
      Caption         =   "适用性别(&X)"
      Height          =   180
      Left            =   285
      TabIndex        =   14
      Top             =   2600
      Width           =   990
   End
   Begin VB.Label lbl适用病情 
      AutoSize        =   -1  'True
      Caption         =   "适用病情(&B)"
      Height          =   180
      Left            =   3735
      TabIndex        =   12
      Top             =   2220
      Width           =   990
   End
   Begin VB.Label lbl病例分型 
      AutoSize        =   -1  'True
      Caption         =   "病例分型(&T)"
      Height          =   180
      Left            =   285
      TabIndex        =   10
      Top             =   2220
      Width           =   990
   End
   Begin VB.Label lbl分类 
      AutoSize        =   -1  'True
      Caption         =   "分类(&K)"
      Height          =   180
      Left            =   255
      TabIndex        =   0
      Top             =   750
      Width           =   630
   End
   Begin VB.Label lbl名称 
      AutoSize        =   -1  'True
      Caption         =   "名称(&N)"
      Height          =   180
      Left            =   255
      TabIndex        =   4
      Top             =   1110
      Width           =   630
   End
   Begin VB.Label lbl编码 
      AutoSize        =   -1  'True
      Caption         =   "编码(&C)"
      Height          =   180
      Left            =   2880
      TabIndex        =   2
      Top             =   750
      Width           =   630
   End
End
Attribute VB_Name = "frmPathEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event AfterSave(ByVal 分类 As String, ByVal 编码 As String)

Private mstrPrivs As String
Private mlng路径ID As Long
Private mstr分类 As String
Private mblnReturn As Boolean
Private mblnChange As Boolean
Private mblnOK As Boolean

Public Function ShowEdit(frmMain As Object, ByVal strPrivs As String, Optional ByVal lng路径ID As Long, Optional ByVal str分类 As String) As Boolean
'功能：新增或者修改临床路径
'参数：lng路径ID=修改时传入具体的ID值，新增时不传入
'      str分类=新增时，传入当前所选择的分类作为缺省，也可以不传入
    mstrPrivs = strPrivs
    mlng路径ID = lng路径ID
    mstr分类 = str分类
    
    Me.Show 1, frmMain
    ShowEdit = mblnOK
End Function

Private Sub cbo病例分型_Click()
    mblnChange = True
End Sub

Private Sub cbo分类_Change()
    mblnChange = True
End Sub

Private Sub cbo分类_Click()
    If mlng路径ID = 0 Then
        txt编码.Text = GetNextCode(cbo分类.Text)
    End If
    If vsDept.Enabled Then
        vsDept.Rows = 1
        vsDept.Rows = 2
        Call AddDept
    End If
    mblnChange = True
End Sub

Private Sub cbo分类_GotFocus()
    Call zlControl.TxtSelAll(cbo分类)
End Sub

Private Sub cbo分类_Validate(Cancel As Boolean)
    If mlng路径ID = 0 And cbo分类.ListIndex = -1 Then
        txt编码.Text = GetNextCode(cbo分类.Text)
    End If
End Sub

Private Sub cbo年龄单位_Click()
    mblnChange = True
End Sub

Private Sub cbo适用病情_Click()
    mblnChange = True
End Sub

Private Sub cbo适用性别_Click()
    mblnChange = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsTmp As ADODB.Recordset
    Dim str科室IDs As String
    Dim str病种IDs As String
    Dim strSql As String, i As Long
    Dim strTmp As String, intLimit As Integer
    Dim str跳转病种IDs As String

    '1)必须输入的项目
    If cbo分类.Text = "" Then
        MsgBox "必须指定临床路径的分类。", vbInformation, gstrSysName
        cbo分类.SetFocus: Exit Sub
    End If
    If txt编码.Text = "" Then
        MsgBox "必须指定临床路径的编码。", vbInformation, gstrSysName
        txt编码.SetFocus: Exit Sub
    End If
    If txt名称.Text = "" Then
        MsgBox "必须指定临床路径的名称。", vbInformation, gstrSysName
        txt名称.SetFocus: Exit Sub
    End If
    If txt适用年龄(0).Text <> "" And txt适用年龄(1).Text = "" Or _
       txt适用年龄(0).Text = "" And txt适用年龄(1).Text <> "" Then
        MsgBox "临床路径所适用的年龄应该是一个范围。", vbInformation, gstrSysName
        If txt适用年龄(0).Text = "" Then
            txt适用年龄(0).SetFocus
        Else
            txt适用年龄(1).SetFocus
        End If
        Exit Sub
    End If

    '2)输入长度检查
    If zlCommFun.ActualLen(cbo分类.Text) > 50 Then
        MsgBox "临床路径的分类信息最多只允许 25 个汉字或 50 个字符。", vbInformation, gstrSysName
        cbo分类.SetFocus: Exit Sub
    End If
    If zlCommFun.ActualLen(txt名称.Text) > txt名称.MaxLength Then
        MsgBox "临床路径的名称最多只允许 " & txt名称.MaxLength \ 2 & " 个汉字或 " & txt名称.MaxLength & " 个字符。", vbInformation, gstrSysName
        txt名称.SetFocus: Exit Sub
    End If
    If zlCommFun.ActualLen(txt说明.Text) > txt说明.MaxLength Then
        MsgBox "临床路径的说明信息最多只允许 " & txt说明.MaxLength \ 2 & " 个汉字或 " & txt说明.MaxLength & " 个字符。", vbInformation, gstrSysName
        txt说明.SetFocus: Exit Sub
    End If

    '3)其他检查
    If opt应用范围(1).Value Then
        With vsDept
            strTmp = ""
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    If InStr(strTmp & ",", "," & .RowData(i) & ",") > 0 Then
                        MsgBox "发现存在两行相同的科室。", vbInformation, gstrSysName
                        .Row = i: .Col = 0
                        .ShowCell .Row, .Col
                        .SetFocus: Exit Sub
                    Else
                        strTmp = strTmp & "," & .RowData(i)
                    End If
                End If
            Next
        End With
    End If
    With vsDisease(0)
        strTmp = ""
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, 2) <> "" Then
                If Val(.TextMatrix(i, 0)) <> 0 Then
                    strSql = "A" & .RowData(i)
                Else
                    strSql = "B" & .RowData(i)
                End If

                If InStr(strTmp & ",", "," & strSql & ",") > 0 Then
                    MsgBox "发现存在两行相同的病种。", vbInformation, gstrSysName
                    .Row = i: .Col = 2
                    .ShowCell .Row, .Col
                    .SetFocus: Exit Sub
                Else
                    strTmp = strTmp & "," & strSql
                End If
            End If
        Next
    End With
    If vsDisease(1).Enabled Then
        With vsDisease(1)
            strTmp = ""
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, 2) <> "" Then
                    If Val(.TextMatrix(i, 0)) <> 0 Then
                        strSql = "A" & .RowData(i)
                    Else
                        strSql = "B" & .RowData(i)
                    End If

                    If InStr(strTmp & ",", "," & strSql & ",") > 0 Then
                        MsgBox "发现存在两行相同的病种。", vbInformation, gstrSysName
                        .Row = i: .Col = 2
                        .ShowCell .Row, .Col
                        .SetFocus: Exit Sub
                    Else
                        strTmp = strTmp & "," & strSql
                    End If
                End If
            Next
        End With
    End If

    '4)保存数据
    If opt应用范围(1).Value Then
        For i = 1 To vsDept.Rows - 1
            If vsDept.RowData(i) <> 0 Then
                str科室IDs = str科室IDs & "," & vsDept.RowData(i)
            End If
        Next
        str科室IDs = Mid(str科室IDs, 2)
        If str科室IDs = "" Then
            MsgBox "必须指定临床路径的科室应用范围。", vbInformation, gstrSysName
            vsDept.SetFocus: Exit Sub
        End If
    End If

    For i = 1 To vsDisease(0).Rows - 1
        If vsDisease(0).RowData(i) <> 0 And Val(vsDisease(0).TextMatrix(i, 0)) <> 0 Then
            str病种IDs = str病种IDs & "," & vsDisease(0).RowData(i)
        End If
    Next
    str病种IDs = Mid(str病种IDs, 2) & ";"

    For i = 1 To vsDisease(0).Rows - 1
        If vsDisease(0).RowData(i) <> 0 And Val(vsDisease(0).TextMatrix(i, 1)) <> 0 Then
            str病种IDs = str病种IDs & vsDisease(0).RowData(i) & ","
        End If
    Next

    If vsDisease(1).Enabled Then
        For i = 1 To vsDisease(1).Rows - 1
            If vsDisease(1).RowData(i) <> 0 And Val(vsDisease(1).TextMatrix(i, 0)) <> 0 Then
                str跳转病种IDs = str跳转病种IDs & "," & vsDisease(1).RowData(i)
            End If
        Next

        str跳转病种IDs = Mid(str跳转病种IDs, 2) & ";"

        For i = 1 To vsDisease(1).Rows - 1
            If vsDisease(1).RowData(i) <> 0 And Val(vsDisease(1).TextMatrix(i, 1)) <> 0 Then
                str跳转病种IDs = str跳转病种IDs & vsDisease(1).RowData(i) & ","
            End If
        Next
    Else
        str跳转病种IDs = ";"
    End If

    If str病种IDs = ";" Then
        MsgBox "必须指定临床路径所对应的的病种。", vbInformation, gstrSysName
        vsDisease(0).SetFocus: Exit Sub
    End If
    If str跳转病种IDs = ";" Then
        str跳转病种IDs = ""
    End If
    '去掉最右边的逗号
    If Right(str病种IDs, 1) = "," Then str病种IDs = Left(str病种IDs, Len(str病种IDs) - 1)
    If Right(str跳转病种IDs, 1) = "," Then str跳转病种IDs = Left(str跳转病种IDs, Len(str跳转病种IDs) - 1)
    If mlng路径ID = 0 Then
        strSql = "Zl_临床路径目录_Insert('" & cbo分类.Text & "','" & txt编码.Text & "','" & txt名称.Text & "'," & _
                 "'" & txt说明.Text & "','" & zlCommFun.GetNeedName(cbo病例分型.Text) & "','" & IIf(cbo适用病情.ListIndex = 0, "", zlCommFun.GetNeedName(cbo适用病情.Text)) & "'," & _
                 cbo适用性别.ListIndex & ",'" & IIf(txt适用年龄(0).Text <> "", txt适用年龄(0).Text & "-" & txt适用年龄(1).Text & cbo年龄单位.Text, "") & "'," & _
                 IIf(opt应用范围(0).Value, 1, 2) & ",'" & str科室IDs & "','" & str病种IDs & "',Null," & IIf(txtConfirmDay.Enabled, ZVal(Val(txtConfirmDay.Text)), "Null") & ",'" & _
                 str跳转病种IDs & "'," & ZVal(chk诊断不同允许正常完成.Value) & "," & IIf(opt合并路径.Value, "1", "0") & ")"
    Else
        strSql = "Zl_临床路径目录_Update(" & mlng路径ID & ",'" & cbo分类.Text & "','" & txt编码.Text & "','" & txt名称.Text & "'," & _
                 "'" & txt说明.Text & "','" & zlCommFun.GetNeedName(cbo病例分型.Text) & "','" & IIf(cbo适用病情.ListIndex = 0, "", zlCommFun.GetNeedName(cbo适用病情.Text)) & "'," & _
                 cbo适用性别.ListIndex & ",'" & IIf(txt适用年龄(0).Text <> "", txt适用年龄(0).Text & "-" & txt适用年龄(1).Text & cbo年龄单位.Text, "") & "'," & _
                 IIf(opt应用范围(0).Value, 1, 2) & ",'" & str科室IDs & "','" & str病种IDs & "'," & IIf(txtConfirmDay.Enabled, ZVal(Val(txtConfirmDay.Text)), "Null") & ",'" & _
                 str跳转病种IDs & "'," & ZVal(chk诊断不同允许正常完成.Value) & "," & IIf(opt合并路径.Value, "1", "0") & ")"
    End If

    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    On Error GoTo 0

    '5)完成处理
    mblnOK = True
    RaiseEvent AfterSave(cbo分类.Text, txt编码.Text)

    '连续增加
    If mlng路径ID = 0 And chk连续.Value = 1 Then
        '授权数量检查
        If InStr(mstrPrivs, "不限制路径数量") > 0 Then
            intLimit = 0
        ElseIf InStr(mstrPrivs, "30个以下路径") > 0 Then
            intLimit = 30
        ElseIf InStr(mstrPrivs, "5个以下路径") > 0 Then
            intLimit = 5
        End If
        If intLimit > 0 Then
            On Error GoTo errH
            strSql = "Select Nvl(Count(*),0) as 数量 From 临床路径目录"
            Set rsTmp = New ADODB.Recordset
            Call zlDatabase.OpenRecordset(rsTmp, strSql, Me.Caption)
            If rsTmp!数量 < intLimit Then intLimit = 0
            On Error GoTo 0
        End If
        If intLimit = 0 Then
            txt编码.Text = GetNextCode(cbo分类.Text)
            txt名称.Text = "": txt说明.Text = ""
            txtConfirmDay.Text = "0"
            mblnChange = False: txt名称.SetFocus
            Exit Sub
        End If
    End If

    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TypeName(Me.ActiveControl) <> "VSFlexGrid" Then
            KeyAscii = 0
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim intIdx As Integer

    On Error GoTo errH

    mblnOK = False

    '字典信息读取
    '-------------------------------------------------------------------------------------
    '分类信息
    strSql = "Select Distinct 分类 From 临床路径目录 Where 分类 is Not NULL Order by 分类"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSql, Me.Caption)
    Do While Not rsTmp.EOF
        cbo分类.AddItem rsTmp!分类
        rsTmp.MoveNext
    Loop

    '病例分型
    cbo病例分型.AddItem ""
    cbo病例分型.ListIndex = 0
    strSql = "Select 编码,名称,简码 From 临床病例分型 Order by 编码"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSql, Me.Caption)
    Do While Not rsTmp.EOF
        cbo病例分型.AddItem rsTmp!编码 & "-" & rsTmp!名称
        rsTmp.MoveNext
    Loop

    '病情
    cbo适用病情.AddItem "0-不区分病情"
    cbo适用病情.ListIndex = 0
    strSql = "Select 编码,名称,简码 From 病情 Order by 编码"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSql, Me.Caption)
    Do While Not rsTmp.EOF
        cbo适用病情.AddItem rsTmp!编码 & "-" & rsTmp!名称
        rsTmp.MoveNext
    Loop

    '性别
    cbo适用性别.AddItem "0-不区分性别"
    cbo适用性别.AddItem "1-男性"
    cbo适用性别.AddItem "2-女性"
    cbo适用性别.ListIndex = 0

    '年龄单位
    cbo年龄单位.AddItem "岁"
    cbo年龄单位.AddItem "月"
    cbo年龄单位.AddItem "天"
    cbo年龄单位.ListIndex = 0

    '权限限制
    opt应用范围(1).Value = True    '缺省为指定科室
    If InStr(mstrPrivs, "全院路径") = 0 Then
        opt应用范围(0).Enabled = False
    End If

    '临床路径信息
    '-------------------------------------------------------------------------------------
    If mlng路径ID = 0 Then
        '新增临床路径
        vsDept.Enabled = opt应用范围(1).Value
        cbo分类.ListIndex = Cbo.FindIndex(cbo分类, mstr分类)    '隐含调用Call AddDept
    Else
        vsDept.Enabled = opt应用范围(1).Value
        chk连续.Visible = False

        '修改临床路径
        strSql = "Select 分类,编码,名称,说明,病例分型,适用病情,适用性别,适用年龄,通用,确诊天数,结束路径控制,性质 From 临床路径目录 Where ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng路径ID)

        cbo分类.Text = Nvl(rsTmp!分类)
        txt编码.Text = rsTmp!编码
        txt名称.Text = rsTmp!名称
        txt说明.Text = Nvl(rsTmp!说明)
        txtConfirmDay.Text = Val("" & rsTmp!确诊天数)
        chk诊断不同允许正常完成.Value = Val(rsTmp!结束路径控制 & "")

        If Not IsNull(rsTmp!病例分型) Then
            cbo病例分型.ListIndex = Cbo.FindIndex(cbo病例分型, CStr(rsTmp!病例分型))
        End If

        If Not IsNull(rsTmp!适用病情) Then
            cbo适用病情.ListIndex = Cbo.FindIndex(cbo适用病情, CStr(rsTmp!适用病情))
        End If

        cbo适用性别.ListIndex = Val(Nvl(rsTmp!适用性别, 0))

        If Not IsNull(rsTmp!适用年龄) Then
            txt适用年龄(0).Text = Split(rsTmp!适用年龄, "-")(0)
            txt适用年龄(1).Text = Val(Split(rsTmp!适用年龄, "-")(1))
            cbo年龄单位.ListIndex = Cbo.FindIndex(cbo年龄单位, CStr(Right(Split(rsTmp!适用年龄, "-")(1), 1)))
        End If

        If Val(rsTmp!性质 & "") = 1 Then
            opt合并路径.Value = True
        Else
            opt首要路径.Value = True
        End If

        '应用科室范围
        opt应用范围(0).Value = Val(Nvl(rsTmp!通用, 1)) = 1
        opt应用范围(1).Value = Val(Nvl(rsTmp!通用, 1)) = 2
        If Val(Nvl(rsTmp!通用, 1)) = 2 Then
            strSql = "Select B.ID,B.编码,B.名称 From 临床路径科室 A,部门表 B Where A.科室ID=B.ID And A.路径ID=[1] Order by B.编码"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng路径ID)
            If Not rsTmp.EOF Then
                vsDept.Rows = vsDept.FixedRows + rsTmp.RecordCount + 1    '加一空行
                For intIdx = 1 To rsTmp.RecordCount
                    vsDept.RowData(intIdx) = Val(rsTmp!ID)
                    vsDept.TextMatrix(intIdx, 0) = rsTmp!编码 & "-" & rsTmp!名称
                    vsDept.Cell(flexcpData, intIdx, 0) = vsDept.TextMatrix(intIdx, 0)

                    rsTmp.MoveNext
                Next
            End If
        End If
        vsDept.Row = 0: vsDept.Row = 1: vsDept.Col = 0

        '对应病种范围
        strSql = _
        " Select" & _
                 " A.疾病ID,B.编码 as 疾病编码,B.名称 as 疾病名称," & _
                 " A.诊断ID,C.编码 as 诊断编码,C.名称 as 诊断名称,Nvl(a.性质,0) as 性质" & _
                 " From 临床路径病种 A,疾病编码目录 B,疾病诊断目录 C" & _
                 " Where A.疾病ID=B.ID(+) And A.诊断ID=C.ID(+) And A.路径ID=[1] " & _
                 " Order by B.编码,C.编码"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng路径ID)
        If Not rsTmp.EOF Then
            rsTmp.Filter = "性质=0"
            vsDisease(0).Rows = vsDisease(0).FixedRows + rsTmp.RecordCount + 1    '加一空行
            For intIdx = 1 To rsTmp.RecordCount
                If Not IsNull(rsTmp!疾病id) Then
                    vsDisease(0).RowData(intIdx) = Val(rsTmp!疾病id & "")
                    vsDisease(0).TextMatrix(intIdx, 0) = -1
                    vsDisease(0).TextMatrix(intIdx, 1) = 0
                    vsDisease(0).TextMatrix(intIdx, 2) = "[" & rsTmp!疾病编码 & "]" & rsTmp!疾病名称
                    vsDisease(0).ColData(2) = vsDisease(1).ColData(2) & "," & rsTmp!疾病编码 & ","
                Else
                    vsDisease(0).RowData(intIdx) = Val(rsTmp!诊断id & "")
                    vsDisease(0).TextMatrix(intIdx, 1) = -1
                    vsDisease(0).TextMatrix(intIdx, 0) = 0
                    vsDisease(0).TextMatrix(intIdx, 2) = "[" & rsTmp!诊断编码 & "]" & rsTmp!诊断名称
                    vsDisease(0).ColData(2) = vsDisease(1).ColData(2) & "," & rsTmp!诊断编码 & ","
                End If
                vsDisease(0).Cell(flexcpData, intIdx, 2) = vsDisease(0).TextMatrix(intIdx, 0)

                rsTmp.MoveNext
            Next
            vsDisease(0).TextMatrix(vsDisease(0).Rows - 1, 0) = vsDisease(0).TextMatrix(vsDisease(0).Rows - 2, 0)
            vsDisease(0).TextMatrix(vsDisease(0).Rows - 1, 1) = vsDisease(0).TextMatrix(vsDisease(0).Rows - 2, 1)

            rsTmp.Filter = "性质=1"
            vsDisease(1).Rows = vsDisease(1).FixedRows + rsTmp.RecordCount + 1    '加一空行
            For intIdx = 1 To rsTmp.RecordCount
                If Not IsNull(rsTmp!疾病id) Then
                    vsDisease(1).RowData(intIdx) = Val(rsTmp!疾病id & "")
                    vsDisease(1).TextMatrix(intIdx, 0) = -1
                    vsDisease(1).TextMatrix(intIdx, 1) = 0
                    vsDisease(1).TextMatrix(intIdx, 2) = "[" & rsTmp!疾病编码 & "]" & rsTmp!疾病名称
                    vsDisease(1).ColData(2) = vsDisease(1).ColData(2) & "," & rsTmp!疾病编码 & ","
                Else
                    vsDisease(1).RowData(intIdx) = Val(rsTmp!诊断id & "")
                    vsDisease(1).TextMatrix(intIdx, 1) = -1
                    vsDisease(1).TextMatrix(intIdx, 0) = 0
                    vsDisease(1).TextMatrix(intIdx, 2) = "[" & rsTmp!诊断编码 & "]" & rsTmp!诊断名称
                    vsDisease(1).ColData(2) = vsDisease(1).ColData(2) & "," & rsTmp!诊断编码 & ","
                End If
                vsDisease(1).Cell(flexcpData, intIdx, 0) = vsDisease(1).TextMatrix(intIdx, 0)

                rsTmp.MoveNext
            Next
            If rsTmp.RecordCount = 0 Then
                vsDisease(1).TextMatrix(vsDisease(1).Rows - 1, 0) = -1
            Else
                vsDisease(1).TextMatrix(vsDisease(1).Rows - 1, 0) = vsDisease(1).TextMatrix(vsDisease(1).Rows - 2, 0)
                vsDisease(1).TextMatrix(vsDisease(1).Rows - 1, 1) = vsDisease(1).TextMatrix(vsDisease(1).Rows - 2, 1)
            End If
        End If
        vsDisease(0).Row = 0: vsDisease(0).Row = 1: vsDisease(0).Col = 2
        vsDisease(1).Row = 0: vsDisease(1).Row = 1: vsDisease(1).Col = 2
    End If
    vsDisease_AfterRowColChange 0, -1, -1, 1, 2
    vsDisease_AfterRowColChange 1, -1, -1, 1, 2
    mblnChange = False
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mblnOK And mlng路径ID <> 0 And mblnChange Then
        If MsgBox("该临床路径的信息已被更改，确实要放弃更改退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
    
    mstrPrivs = ""
    mlng路径ID = 0
    mstr分类 = ""
End Sub

Private Sub opt合并路径_Click()
    Call SetFirstPath(False)
End Sub

Private Sub opt首要路径_Click()
    Call SetFirstPath(True)
End Sub

Private Sub SetFirstPath(ByVal blnVisible As Boolean)
'功能：如果选择首要路径或合并路径，则设置界面的控件
'参数：blnVisible=true 设置为首要路径，否则为合并路径
    Lbl跳转接收病种.Enabled = blnVisible
    vsDisease(1).Enabled = blnVisible
    vsDisease(1).BackColor = IIf(blnVisible, vbWindowBackground, vbButtonFace)
    vsDisease(1).BackColorBkg = IIf(blnVisible, vbWindowBackground, vbButtonFace)
    lblConfirmDay.Enabled = blnVisible
    txtConfirmDay.Enabled = blnVisible
    txtConfirmDay.BackColor = IIf(blnVisible, &HC0E0FF, Me.BackColor)
    lblConfirm.Enabled = blnVisible
End Sub

Private Sub opt应用范围_Click(Index As Integer)
    vsDept.Enabled = opt应用范围(1).Value
    If Visible And vsDept.Enabled Then
        vsDept.SetFocus
    Else
        vsDept.Rows = 1
        vsDept.Rows = 2
    End If

    mblnChange = True
End Sub

Private Sub txtConfirmDay_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtConfirmDay_GotFocus()
    Call zlControl.TxtSelAll(txtConfirmDay)
End Sub

Private Sub txt编码_Change()
    mblnChange = True
End Sub

Private Sub txt编码_GotFocus()
    Call zlControl.TxtSelAll(txt编码)
End Sub

Private Sub txt编码_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt名称_Change()
    mblnChange = True
End Sub

Private Sub txt名称_GotFocus()
    Call zlControl.TxtSelAll(txt名称)
End Sub

Private Sub txt适用年龄_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub txt适用年龄_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txt适用年龄(Index))
End Sub

Private Sub txt适用年龄_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt说明_Change()
    mblnChange = True
End Sub

Private Sub txt说明_GotFocus()
    Call zlControl.TxtSelAll(txt说明)
End Sub

Private Sub vsDept_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call vsDept_AfterRowColChange(-1, -1, Row, Col)
End Sub

Private Sub vsDept_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsDept
        If NewCol <> 2 Then
            .ComboList = "..."
            .FocusRect = flexFocusSolid
        Else
            .FocusRect = flexFocusLight
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vsDept_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim vPoint As POINTAPI
    
    With vsDept
        If InStr(mstrPrivs, "全院路径") = 0 Then
            '当前人员所属临床科室
            strSql = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
                " From 部门表 A,部门人员 B,部门性质说明 C" & _
                " Where A.ID=B.部门ID And B.人员ID=[1]" & _
                " And A.ID=C.部门ID And C.服务对象 IN(2,3) And C.工作性质='临床'" & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                " Order by A.编码"
        Else
            '全院临床科室
            strSql = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
                " From 部门表 A,部门性质说明 C" & _
                " Where A.ID=C.部门ID And C.服务对象 IN(2,3) And C.工作性质='临床'" & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                " Order by A.编码"
        End If
        vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
        Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSql, 0, "临床科室", False, "", "", False, False, True, _
            vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, UserInfo.ID)
        
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "没有临床科室数据可以选择。", vbInformation, gstrSysName
            End If
        Else
            Call SetDeptInput(Row, rsTmp)
            Call DeptEnterNextCell(True)
        End If
    End With
End Sub

Private Sub vsDept_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    With vsDept
        If KeyCode = vbKeyF4 Then
            If .Col = 0 Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, 0) <> "" Then
                If MsgBox("确实要清除该行科室吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                    mblnChange = True
                End If
            End If
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsDept_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsDept_KeyPress(KeyAscii As Integer)
    With vsDept
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call DeptEnterNextCell
        ElseIf .Col = 0 Then
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsDept_CellButtonClick(.Row, .Col)
            Else
                .ComboList = "" '使按钮状态进入输入状态
            End If
        End If
    End With
End Sub

Private Sub vsDept_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
End Sub

Private Sub vsDept_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsDept.EditSelStart = 0
    vsDept.EditSelLength = zlCommFun.ActualLen(vsDept.EditText)
End Sub

Private Sub vsDept_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    
    With vsDept
        If Col = 0 Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then Call DeptEnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call DeptEnterNextCell
            Else
                strInput = UCase(.EditText)
                If InStr(mstrPrivs, "全院路径") = 0 Then
                    '当前人员所属临床科室
                    strSql = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
                        " From 部门表 A,部门人员 B,部门性质说明 C" & _
                        " Where A.ID=B.部门ID And B.人员ID=[3]" & _
                        " And A.ID=C.部门ID And C.服务对象 IN(2,3) And C.工作性质='临床'" & _
                        " And (A.编码 Like [1] Or A.名称 Like [2] Or A.简码 Like [2])" & _
                        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                        " Order by A.编码"
                Else
                    '全院临床科室
                    strSql = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
                        " From 部门表 A,部门性质说明 C" & _
                        " Where A.ID=C.部门ID And C.服务对象 IN(2,3) And C.工作性质='临床'" & _
                        " And (A.编码 Like [1] Or A.名称 Like [2] Or A.简码 Like [2])" & _
                        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                        " Order by A.编码"
                End If
                vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSql, 0, "临床科室", False, "", "", False, _
                    False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    strInput & "%", gstrLike & strInput & "%", UserInfo.ID)
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "没有找到匹配的临床科室。", vbInformation, gstrSysName
                    End If
                    Cancel = True
                Else
                    Call SetDeptInput(Row, rsTmp)
                    .EditText = .Text
                    If mblnReturn Then Call DeptEnterNextCell(True)
                End If
            End If
            mblnReturn = False
        End If
    End With
End Sub

Private Sub DeptEnterNextCell(Optional ByVal blnNewRow As Boolean)
    Dim i As Long, j As Long
    
    With vsDept
        If blnNewRow Then
            .Row = .Rows - 1: .Col = 0
            .ShowCell .Row, .Col
        Else
            If .Row + 1 <= .Rows - 1 Then
                .Row = .Row + 1
                .ShowCell .Row, .Col
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub

Private Sub SetDeptInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'功能：处理过敏药物的输入
    Dim i As Long
    Dim intCount As Integer

    With vsDept
        For i = 1 To rsInput.RecordCount
            If .FindRow(Val(rsInput!ID)) = -1 Then
                intCount = intCount + i
                If i > 1 Then
                    .AddItem "", lngRow + 1
                    lngRow = lngRow + 1
                End If

                .RowData(lngRow) = Val(rsInput!ID)
                .TextMatrix(lngRow, 0) = rsInput!编码 & "-" & rsInput!名称
                .Cell(flexcpData, lngRow, 0) = .TextMatrix(lngRow, 0)
            End If
            rsInput.MoveNext
        Next

        '始终保持一空行
        If lngRow = .Rows - 1 And intCount > 0 Then
            .AddItem "", lngRow + 1
        End If

        mblnChange = True
    End With
End Sub

Private Sub vsDisease_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Col = 0 Or Col = 1 Then
        With vsDisease(Index)
            If Val(.TextMatrix(Row, Col)) <> 0 Then
                .TextMatrix(Row, IIf(Col = 1, 0, 1)) = 0
                
                If .RowData(Row) <> 0 Then
                    .RowData(Row) = 0
                    .TextMatrix(Row, 2) = ""
                    .Cell(flexcpData, Row, 2) = ""
                    
                    mblnChange = True
                End If
            End If
        End With
    End If
    
    Call vsDisease_AfterRowColChange(Index, -1, -1, Row, Col)
End Sub

Private Sub vsDisease_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsDisease(Index)
        If NewCol <> 2 Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
            .ComboList = "..."
        End If
    End With
End Sub

Private Sub vsDisease_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Or Col = 1 Then
        If Val(vsDisease(Index).TextMatrix(Row, Col)) <> 0 Then Cancel = True
    End If
End Sub

Private Sub vsDisease_BeforeUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 2 Then Cancel = True
End Sub

Private Sub vsDisease_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset

    With vsDisease(Index)

        If Val(.TextMatrix(Row, 1)) <> 0 Then
            '按诊断输入:西医部份，一个诊断可能属于多个分类
            Set rsTmp = zlDatabase.ShowILLSelect(Me, "1", 0, , True, False, .ColData(2))
        Else
            'D-ICD-10疾病编码
            Set rsTmp = zlDatabase.ShowILLSelect(Me, "D,B", 0, Decode(cbo适用性别.ListIndex, 1, "男", 2, "女"), True, True, .ColData(2))
        End If
        If Not rsTmp Is Nothing Then
            Call SetDiseaseInput(Index, Row, rsTmp)
            Call DiseaseEnterNextCell(Index, True)
        End If
    End With
End Sub

Private Sub vsDisease_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim strTemp As String

    With vsDisease(Index)
        If KeyCode = vbKeyF4 Then
            If .Col = 2 Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, 2) <> "" Then
                If MsgBox("确实要清除该行内容吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    strTemp = .TextMatrix(.Row, 2)
                    strTemp = Mid(strTemp, 2, InStr(strTemp, "]") - 2)
                    .ColData(2) = Replace(.ColData(2), "," & strTemp & ",", "")
                    .RemoveItem .Row
                    mblnChange = True
                End If
            End If
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsDisease_KeyPress(Index, KeyCode)
        End If
    End With
End Sub

Private Sub vsDisease_KeyPress(Index As Integer, KeyAscii As Integer)
    With vsDisease(Index)
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call DiseaseEnterNextCell(Index)
        Else
            If .Col = 2 Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsDisease_CellButtonClick(Index, .Row, .Col)
                Else
                    .ComboList = "" '使按钮状态进入输入状态
                End If
            End If
        End If
    End With
End Sub

Private Sub vsDisease_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
End Sub

Private Sub vsDisease_SetupEditWindow(Index As Integer, ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsDisease(Index).EditSelStart = 0
    vsDisease(Index).EditSelLength = zlCommFun.ActualLen(vsDisease(Index).EditText)
End Sub

Private Sub vsDisease_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim str性别 As String, strInput As String
    Dim vPoint As POINTAPI, int诊断输入 As Integer
    
    With vsDisease(Index)
        If Col = 2 Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then Call DiseaseEnterNextCell(Index)
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call DiseaseEnterNextCell(Index)
            Else
                strInput = UCase(.EditText)
                If Val(.TextMatrix(Row, 1)) <> 0 Then
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
                    If cbo适用性别.ListIndex = 1 Then
                        str性别 = "男"
                    ElseIf cbo适用性别.ListIndex = 2 Then
                        str性别 = "女"
                    End If
                    'D-ICD-10疾病编码
                    If zlCommFun.IsCharChinese(strInput) Then
                        strSql = "名称 Like [2]" '输入汉字时,只匹配名称
                    Else
                        strSql = "编码 Like [1] Or 名称 Like [2] Or " & IIf(gint简码 = 0, "简码", "五笔码") & " Like [2]"
                    End If
                    strSql = _
                        " Select ID,ID as 项目ID,编码,附码,名称," & IIf(gint简码 = 0, "简码", "五笔码 as 简码") & ",说明" & _
                        " From 疾病编码目录 Where 类别 In('D','B') And (" & strSql & ")" & _
                        IIf(str性别 <> "", " And (性别限制=[3] Or 性别限制 is NULL)", "") & _
                        " And (撤档时间 is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " Order by 编码"
                End If
                
                vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSql, 0, IIf(Val(.TextMatrix(Row, 1)) <> 0, "诊断编码", "疾病编码"), _
                    False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    strInput & "%", gstrLike & strInput & "%", str性别, gint简码 + 1)
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "没有找到与输入匹配的内容。", vbInformation, gstrSysName
                    End If
                    Cancel = True
                Else
                    Call SetDiseaseInput(Index, Row, rsTmp): .EditText = .Text
                    If mblnReturn Then Call DiseaseEnterNextCell(Index, True)
                End If
            End If
            mblnReturn = False
        End If
    End With
End Sub

Private Sub DiseaseEnterNextCell(Index As Integer, Optional ByVal blnNewRow As Boolean)
    With vsDisease(Index)
        If blnNewRow Then
            .Row = .Rows - 1: .Col = 2
            .ShowCell .Row, .Col
        Else
            If .Col + 1 <= .Cols - 1 Then
                .Col = .Col + 1
            ElseIf .Row + 1 <= .Rows - 1 Then
                .Row = .Row + 1: .Col = 2
                .ShowCell .Row, .Col
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub

Private Sub SetDiseaseInput(Index As Integer, ByVal lngRow As Long, rsInput As ADODB.Recordset)
'功能：处理西医诊断项目的输入
    Dim i As Long
    Dim intCount As Integer

    With vsDisease(Index)
        For i = 1 To rsInput.RecordCount
            If .FindRow(Val(rsInput!项目ID)) = -1 Then
                intCount = intCount + 1    '可添加记录（不重复的记录）数
                If intCount > 1 Then
                    .AddItem "", lngRow + 1
                    lngRow = lngRow + 1
                    .TextMatrix(lngRow, 0) = .TextMatrix(lngRow - 1, 0)
                    .TextMatrix(lngRow, 1) = .TextMatrix(lngRow - 1, 1)
                End If
                .RowData(lngRow) = Val(rsInput!项目ID)
                .TextMatrix(lngRow, 2) = "[" & rsInput!编码 & "]" & Nvl(rsInput!名称)
                .Cell(flexcpData, lngRow, 2) = .TextMatrix(lngRow, 2)
                .ColData(2) = .ColData(2) & "," & rsInput!编码 & ","
            End If
            rsInput.MoveNext
        Next

        '始终保持一空行，intCount:当一条添加记录都没有时，禁止添加空行
        If lngRow = .Rows - 1 And intCount > 0 Then
            .AddItem "", lngRow + 1
            .TextMatrix(.Rows - 1, 0) = .TextMatrix(.Rows - 2, 0)
            .TextMatrix(.Rows - 1, 1) = .TextMatrix(.Rows - 2, 1)
        End If

        mblnChange = True
    End With
End Sub

Private Sub AddDept()
'功能:指定科室时，根据路径路径分类名称，自动添加临床科室部门

    Dim rsTmp       As ADODB.Recordset
    Dim strSql      As String
    Dim i           As Long

    On Error GoTo errH

    If InStr(mstrPrivs, "全院路径") = 0 Then
        '非全院路径权限
        '路径管理员属于多个临床科室的情况：先根据分类名称，从管理员所在临床科室中找到与分类名称相同的科室，否则不添加
        strSql = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
                 " From 部门表 A,部门人员 B,部门性质说明 C" & _
                 " Where A.ID=B.部门ID And B.人员ID=[1]" & _
                 " And A.ID=C.部门ID And C.服务对象 IN(2,3) And C.工作性质='临床'  " & _
                 " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                 " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                 " Order by A.编码"
    Else
        '全院路径权限
        '根据分类名称查找，找到就加载，找不到不自动加载，操作员手动加载
        strSql = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
                 " From 部门表 A,部门性质说明 C" & _
                 " Where A.ID=C.部门ID And C.服务对象 IN(2,3) And C.工作性质='临床'" & _
                 " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                 " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                 " Order by A.编码"
    End If

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)

    With vsDept
        rsTmp.Filter = "名称='" & cbo分类.List(cbo分类.ListIndex) & "'"
        For i = 1 To rsTmp.RecordCount
            If .FindRow(rsTmp!ID) = -1 Then    '已经添加过，禁止添加
                .TextMatrix(i, 0) = rsTmp!编码 & "-" & rsTmp!名称
                .RowData(i) = Val(rsTmp!ID)
                .Rows = .Rows + 1
            End If
            rsTmp.MoveNext
        Next
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
