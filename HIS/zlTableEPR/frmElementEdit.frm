VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmElementEdit 
   BackColor       =   &H00F6F6F6&
   BorderStyle     =   0  'None
   Caption         =   "数据编辑器"
   ClientHeight    =   3690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5490
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   -30
      ScaleHeight     =   285
      ScaleWidth      =   5415
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3330
      Width           =   5415
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "ESC 取消退出；回车:保存修改。"
         Height          =   240
         Left            =   90
         TabIndex        =   16
         Top             =   45
         Width           =   3570
      End
      Begin VB.Image imgResize 
         Height          =   270
         Left            =   5175
         MousePointer    =   8  'Size NW SE
         Picture         =   "frmElementEdit.frx":0000
         Top             =   0
         Width           =   225
      End
   End
   Begin VB.PictureBox pic替换项目 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   60
      ScaleHeight     =   705
      ScaleWidth      =   1740
      TabIndex        =   9
      Top             =   810
      Width           =   1740
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "名称:"
         Height          =   195
         Left            =   315
         TabIndex        =   14
         Top             =   360
         Width           =   510
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   45
         Picture         =   "frmElementEdit.frx":03A2
         Top             =   45
         Width           =   240
      End
      Begin VB.Label lbl示例 
         BackStyle       =   0  'Transparent
         Caption         =   "示例"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   810
         TabIndex        =   13
         Top             =   585
         Width           =   2940
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "示例:"
         Height          =   195
         Left            =   315
         TabIndex        =   12
         Top             =   585
         Width           =   510
      End
      Begin VB.Label lbl名称 
         BackStyle       =   0  'Transparent
         Caption         =   "名称"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   795
         TabIndex        =   11
         Top             =   375
         Width           =   2670
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "自动替换项目"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   315
         TabIndex        =   10
         Top             =   90
         Width           =   1635
      End
   End
   Begin VB.TextBox txt上下2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   945
      TabIndex        =   4
      Text            =   "99999"
      Top             =   2295
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picTitle 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F5BE9E&
      BorderStyle     =   0  'None
      Height          =   100
      Left            =   45
      MousePointer    =   5  'Size
      ScaleHeight     =   105
      ScaleWidth      =   5325
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   90
      Width           =   5325
      Begin VB.Image imgTitle 
         Height          =   45
         Left            =   1350
         MousePointer    =   5  'Size
         Picture         =   "frmElementEdit.frx":0617
         Top             =   30
         Width           =   2250
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfg复选 
      Height          =   555
      Left            =   3180
      TabIndex        =   0
      Top             =   270
      Visible         =   0   'False
      Width           =   780
      _cx             =   1376
      _cy             =   979
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   240
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmElementEdit.frx":0699
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
      Ellipsis        =   1
      ExplorerBar     =   7
      PicturesOver    =   -1  'True
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   1
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
      WallPaperAlignment=   1
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.TextBox txt上下1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Text            =   "99999"
      Top             =   2295
      Visible         =   0   'False
      Width           =   630
   End
   Begin MSComCtl2.UpDown ud上下 
      Height          =   300
      Left            =   1395
      TabIndex        =   5
      Top             =   2250
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      OrigLeft        =   1065
      OrigTop         =   2295
      OrigRight       =   1320
      OrigBottom      =   2595
      Enabled         =   -1  'True
   End
   Begin VSFlex8Ctl.VSFlexGrid vfg单选 
      Height          =   570
      Left            =   3225
      TabIndex        =   2
      Top             =   1005
      Visible         =   0   'False
      Width           =   690
      _cx             =   1217
      _cy             =   1005
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   16777215
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   240
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmElementEdit.frx":06D6
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
      Ellipsis        =   1
      ExplorerBar     =   7
      PicturesOver    =   -1  'True
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   1
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
      WallPaperAlignment=   1
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.TextBox txt文本 
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   3225
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1980
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Image imgOpt1 
      Height          =   195
      Left            =   2160
      Picture         =   "frmElementEdit.frx":0713
      Top             =   2655
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgOpt2 
      Height          =   195
      Left            =   2160
      Picture         =   "frmElementEdit.frx":0999
      Top             =   2925
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblDot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      Height          =   300
      Left            =   765
      TabIndex        =   8
      Top             =   2295
      Width           =   105
   End
   Begin VB.Shape shpBorder2 
      BorderColor     =   &H00E09060&
      Height          =   375
      Left            =   450
      Top             =   270
      Width           =   330
   End
   Begin VB.Label lbl单位 
      BackStyle       =   0  'Transparent
      Caption         =   "单位"
      Height          =   210
      Left            =   1665
      TabIndex        =   6
      Top             =   2340
      Width           =   555
   End
   Begin VB.Shape shpBorder1 
      BorderColor     =   &H00E09060&
      Height          =   375
      Left            =   45
      Top             =   270
      Width           =   330
   End
End
Attribute VB_Name = "frmElementEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public frmParent As Object      '父窗体
Public Element As cEPRElement   '诊治要素

Private lngX As Long, lngY As Long, mblnOk As Boolean

'################################################################################################################
'## 功能：  显示诊治要素编辑器
'##
'## 参数：  Ele         :所编辑的诊治要素
'##         (X,Y)       :显示位置（屏幕坐标）
'##         ofrmParent  :父窗体
'##         eEditType   :编辑模式
'################################################################################################################
Public Function ShowMe(ByRef Ele As cEPRElement, ByVal x As Long, ByVal y As Long, Optional ByRef ofrmParent As Object) As Boolean
Dim i As Long, j As Long, T As Variant, strTmp As String
    Set Element = New cEPRElement: mblnOk = False
    Call Ele.Clone(Element)
    Set Me.frmParent = ofrmParent
    With Me.Element
        Select Case .要素表示       '0-文本,1-上下,2-单选,3-复选  5-字典项目
        Case 0
            If Me.Element.替换域 = 1 Then
                lbl名称 = Me.Element.要素名称
                lbl示例 = GetReplaceEleValue(lbl名称, ofrmParent.Document.EPRPatiRecInfo.病人ID, ofrmParent.Document.EPRPatiRecInfo.主页ID, ofrmParent.Document.EPRPatiRecInfo.病人来源, ofrmParent.Document.EPRPatiRecInfo.医嘱id)
                If lbl示例 = "" Then
                    lbl示例.Visible = False
                    Label3.Visible = False
                    Me.Height = 1250
                Else
                    lbl示例.Visible = True
                    Label3.Visible = True
                    Me.Height = 1500
                End If
            Else
                txt文本.MaxLength = .要素长度
                txt文本 = .内容文本
                txt文本.Visible = True
                txt文本.SelStart = 0: txt文本.SelLength = Len(.内容文本)
            End If
        Case 1
            T = Split(.要素值域, ";")    '格式:  0;100000
            If UBound(T) < 1 Then
                ud上下.Min = 0
                ud上下.Max = 0
            Else
                ud上下.Min = IIf(CLng(T(0)) > CLng(T(1)), CLng(T(1)), CLng(T(0)))
                ud上下.Max = IIf(CLng(T(0)) > CLng(T(1)), CLng(T(0)), CLng(T(1)))
            End If
            txt上下1.Tag = "赋值..."
            i = InStr(1, .内容文本, ".")
            If i > 0 Then
                txt上下1 = Mid(.内容文本, 1, i - 1)
                txt上下1.Visible = True
                txt上下1.SelStart = 0: txt上下1.SelLength = Len(txt上下1)
                txt上下2 = Mid(.内容文本, i + 1)
            Else
                txt上下1 = .内容文本
                txt上下2 = ""
            End If
            txt上下1.Tag = ""
            txt上下1.MaxLength = .要素长度
            lbl单位 = .要素单位
            If Trim(.要素单位) <> "" Then
                lbl单位.Visible = True
            Else
                lbl单位.Visible = False
            End If
            If .要素小数 > 0 Then
                txt上下2.MaxLength = .要素小数
                txt上下2.Visible = True
                lblDot.Visible = True
            Else
                txt上下2.Visible = False
                lblDot.Visible = False
            End If
        Case 2
            T = Split(.要素值域, ";")
            vfg单选.Clear
            vfg单选.RowHeightMax = 240
            vfg单选.Cols = 3
            vfg单选.ColWidth(0) = 80
            vfg单选.ColWidth(1) = 200
            vfg单选.Rows = UBound(T) + 1
            For i = 0 To UBound(T)
                vfg单选.Cell(flexcpText, i, 2) = Trim(T(i))
                vfg单选.Cell(flexcpPicture, i, 1) = imgOpt1.Picture
            Next i
            
            If Element.输入形态 = 0 Then
                strTmp = Trim(.内容文本)
            Else
                '展开形式   '○●□■
                strTmp = ""
                i = InStr(1, .内容文本, "●")
                If i > 0 Then
                    j = InStr(i, .内容文本, "○")
                    If j > 0 Then
                        strTmp = Trim(Mid(.内容文本, i + 1, j - i - 1))
                    Else
                        strTmp = Trim(Mid(.内容文本, i + 1))
                    End If
                Else
                    strTmp = ""
                End If
            End If
            vfg单选.FocusRect = flexFocusNone
            vfg单选.Editable = flexEDKbdMouse
            vfg单选.Row = 0
            vfg单选.Col = 0
            Dim blnFinded As Boolean
            vfg单选.Row = 0
            For i = 0 To vfg单选.Rows - 1
                If strTmp = vfg单选.Cell(flexcpText, i, 2) And blnFinded = False Then
                    vfg单选.Cell(flexcpPicture, i, 1) = imgOpt2.Picture
                    vfg单选.Row = i
                    blnFinded = True
                Else
                    vfg单选.Cell(flexcpPicture, i, 1) = imgOpt1.Picture
                End If
            Next
        Case 3
            T = Split(.要素值域, ";")
            vfg复选.Clear
            vfg复选.RowHeightMax = 240
            vfg复选.Cols = 2
            vfg复选.Rows = UBound(T) + 1
            For i = 0 To UBound(T)
                vfg复选.Cell(flexcpText, i, 1) = T(i)
            Next i
            
            If Element.输入形态 = 0 Then
                strTmp = "、" & Trim(.内容文本) & "、"
            Else
                '展开形式
                strTmp = ""
                i = InStr(1, .内容文本, "■")
                Do While i > 0
                    j = InStr(i, .内容文本, " ")
                    If j > 0 Then
                        strTmp = strTmp & "、" & Mid(.内容文本, i + 1, j - i - 1)
                    Else
                        strTmp = strTmp & "、" & Mid(.内容文本, i + 1)
                    End If
                    i = InStr(i + 1, .内容文本, "■")
                Loop
                strTmp = strTmp & "、"
            End If
            vfg复选.Cell(flexcpChecked, 0, 0, vfg复选.Rows - 1, 0) = flexUnchecked
            vfg复选.Editable = flexEDKbdMouse
        
            vfg复选.ColWidth(0) = 240
            For i = 0 To vfg复选.Rows - 1
                If InStr(1, strTmp, "、" & vfg复选.Cell(flexcpText, i, 1) & "、") > 0 Then
                    vfg复选.Cell(flexcpChecked, i, 0) = flexChecked
                Else
                    vfg复选.Cell(flexcpChecked, i, 0) = flexUnchecked
                End If
            Next
            vfg复选.Row = 0

        End Select
    End With
    
    Me.Left = x
    Me.Top = y
    Me.Width = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\" & Me.Name, "MainWidth", 2500)
    If Me.Element.要素表示 <> 1 Then Me.Height = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\" & Me.Name, "MainHeight", 1545)
    Call Form_Resize
    Me.Show vbModal, frmParent
    ShowMe = mblnOk
End Function

Private Sub Form_Unload(Cancel As Integer)
    Form_Deactivate
End Sub

Private Sub txt上下1_GotFocus()
    zlCommFun.OpenIme
    txt上下1.SelStart = 0
    txt上下1.SelLength = Len(txt上下1)
    ud上下.BuddyControl = txt上下1
End Sub

Private Sub txt上下1_KeyPress(KeyAscii As Integer)
    If InStr("1234567890. " & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    If KeyAscii = vbKeySpace Or InStr(".", Chr(KeyAscii)) = 1 Then
        KeyAscii = 0
        If txt上下2.Visible And txt上下2.Enabled Then
            txt上下2.SelStart = 0
            txt上下2.SelLength = Len(txt上下2)
            txt上下2.SetFocus
        End If
    End If
End Sub

Private Sub txt上下2_Change()
    If txt上下1.Tag = "" Then
        If Me.Element.要素小数 > 0 Then
            Dim lngLen As Long, strR As String
            lngLen = Len(Trim(txt上下2))
            If lngLen > Me.Element.要素小数 Then
                strR = Trim(txt上下1.Text) & "." & Trim(txt上下2) & String(Me.Element.要素小数 - Len(Trim(txt上下2)), "0")
            Else
                strR = Trim(txt上下1.Text) & "." & Left(Trim(txt上下2), Me.Element.要素小数)
            End If
        Else
            strR = Trim(txt上下1.Text)
        End If
        Me.Element.内容文本 = IIf(Me.Element.要素小数 > 0, Format(strR, "0." & String(Me.Element.要素小数, "0")), strR)
    End If
End Sub

Private Sub txt上下2_GotFocus()
    zlCommFun.OpenIme
    txt上下2.SelStart = 0
    txt上下2.SelLength = Len(txt上下2)
    ud上下.BuddyControl = txt上下2
End Sub

Private Sub txt上下2_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt文本_GotFocus()
    If Me.Element.要素类型 = 0 Then
        zlCommFun.OpenIme
    End If
End Sub

Private Sub txt文本_KeyPress(KeyAscii As Integer)
    If Me.Element.要素类型 = 0 Then
        '数值型的控制：只能输入数字（小数点和负号，且小数点只能为1个，不能在开头；负号只能在开始处）
        'Asc(".") = vbKeyDelete = 46
        If Len(txt文本.Text) = 0 And KeyAscii = 46 Then KeyAscii = 0
        If InStr(1, txt文本.Text, ".") <> 0 And KeyAscii = 46 Then
            KeyAscii = 0
        ElseIf InStr(1, txt文本.Text, ".") = 0 And KeyAscii = 46 And txt文本.SelLength = Len(txt文本) And txt文本.SelStart = 0 Then
            KeyAscii = 0
        End If
        If txt文本.Text = "-" And KeyAscii = 46 Then KeyAscii = 0
        If KeyAscii = vbKeyBack Or KeyAscii = 46 Then Exit Sub
        If KeyAscii = Asc("-") Then
            If txt文本.SelStart <> 0 Then KeyAscii = 0
        Else
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub vfg单选_DblClick()
    Form_KeyPress vbKeyReturn
End Sub

Private Sub vfg单选_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim i As Long, j As Long, strValue As String
    strValue = ""
    Select Case KeyCode
    Case vbKeySpace
        For i = 0 To vfg单选.Rows - 1
            If i = vfg单选.Row Then
                If vfg单选.Cell(flexcpPicture, i, 1) = imgOpt2.Picture Then
                    vfg单选.Cell(flexcpPicture, i, 1) = imgOpt1.Picture
                Else
                    vfg单选.Cell(flexcpPicture, i, 1) = imgOpt2.Picture
                End If
            Else
                vfg单选.Cell(flexcpPicture, i, 1) = imgOpt1.Picture
            End If
        Next
        If Me.Element.输入形态 = 0 Then
            If vfg单选.Cell(flexcpPicture, vfg单选.Row, 1) = imgOpt2.Picture Then
                strValue = Trim(vfg单选.Cell(flexcpText, vfg单选.Row, 2))
            Else
                strValue = ""
            End If
        Else
            For i = 0 To vfg单选.Rows - 1
                If vfg单选.Cell(flexcpPicture, i, 1) = imgOpt2.Picture Then
                    strValue = strValue & IIf(j = 0, "●", "  ●") & Trim(vfg单选.Cell(flexcpText, i, 2))
                    j = j + 1
                Else
                    strValue = strValue & IIf(j = 0, "○", "  ○") & Trim(vfg单选.Cell(flexcpText, i, 2))
                    j = j + 1
                End If
            Next
        End If
        Element.内容文本 = strValue
        KeyCode = 0
    End Select
End Sub

Private Sub vfg单选_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long, j As Long, strValue As String
    strValue = ""
    
    LockWindowUpdate vfg单选.hWnd
    For i = 0 To vfg单选.Rows - 1
        If i = vfg单选.Row Then
            vfg单选.Cell(flexcpPicture, i, 1) = imgOpt2.Picture
        Else
            vfg单选.Cell(flexcpPicture, i, 1) = imgOpt1.Picture
        End If
    Next
    If Me.Element.输入形态 = 0 Then
        If vfg单选.Cell(flexcpPicture, vfg单选.Row, 1) = imgOpt2.Picture Then
            strValue = Trim(vfg单选.Cell(flexcpText, vfg单选.Row, 2))
        Else
            strValue = ""
        End If
    Else
        For i = 0 To vfg单选.Rows - 1
            If vfg单选.Cell(flexcpPicture, i, 1) = imgOpt2.Picture Then
                strValue = strValue & IIf(j = 0, "●", "  ●") & Trim(vfg单选.Cell(flexcpText, i, 2))
                j = j + 1
            Else
                strValue = strValue & IIf(j = 0, "○", "  ○") & Trim(vfg单选.Cell(flexcpText, i, 2))
                j = j + 1
            End If
        Next
    End If
    Element.内容文本 = strValue
    LockWindowUpdate 0
    UpdateWindow vfg单选.hWnd
End Sub

'#####################################################################################
'## 内部控件事件
'#####################################################################################

Private Sub vfg复选_AfterEdit(ByVal Row As Long, ByVal Col As Long) '○●□■
    Dim i As Long, j As Long, strValue As String
    strValue = ""
    For i = 0 To vfg复选.Rows - 1
        If Me.Element.输入形态 = 0 Then
            If vfg复选.Cell(flexcpChecked, i, 0) = flexChecked Then
                strValue = strValue & IIf(j = 0, "", "、") & Trim(vfg复选.Cell(flexcpText, i, 1))
                j = j + 1
            End If
        Else
            If vfg复选.Cell(flexcpChecked, i, 0) = flexChecked Then
                strValue = strValue & IIf(j = 0, "■", "  ■") & Trim(vfg复选.Cell(flexcpText, i, 1))
                j = j + 1
            Else
                strValue = strValue & IIf(j = 0, "□", "  □") & Trim(vfg复选.Cell(flexcpText, i, 1))
                j = j + 1
            End If
        End If
    Next
    Element.内容文本 = strValue
End Sub

Private Sub vfg复选_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    vfg复选.Col = 0
End Sub

Private Sub vfg单选_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    vfg单选.Col = 0
    Cancel = True
End Sub

Private Sub Form_Activate()
    SetCtlFocus
End Sub

Private Sub Form_Deactivate()
    On Error Resume Next
    SaveSetting "ZLSOFT", "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\" & Me.Name, "MainWidth", Me.Width
    SaveSetting "ZLSOFT", "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\" & Me.Name, "MainHeight", Me.Height
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Me.Element.要素类型 = 0 Then
            '数值型
            Dim T As Variant, dblMax As Double, dblMin As Double
            T = Split(Me.Element.要素值域, ";")    '格式:  0;100000
            If UBound(T) < 1 Then
                dblMin = 0#
                dblMax = 0#
            Else
                dblMin = IIf(CLng(T(0)) > CLng(T(1)), CLng(T(1)), CLng(T(0)))
                dblMax = IIf(CLng(T(0)) > CLng(T(1)), CLng(T(0)), CLng(T(1)))
            End If
            If Me.Element.要素表示 = 0 Then
                '文本表示
                If Trim(txt文本) = "" Then
                    Me.Element.内容文本 = ""
                ElseIf Me.Element.要素值域 <> ";" And Me.Element.要素值域 <> "0;0" And Me.Element.要素值域 <> "" Then
                    If Val(txt文本) > dblMax Then
                        txt文本 = dblMax
                    ElseIf Val(txt文本) < dblMin Then
                        txt文本 = dblMin
                    End If
                    Me.Element.内容文本 = IIf(Me.Element.要素小数 > 0, Format(txt文本, "0." & String(Me.Element.要素小数, "0")), txt文本)
                Else
                    Me.Element.内容文本 = IIf(Me.Element.要素小数 > 0, Format(txt文本, "0." & String(Me.Element.要素小数, "0")), txt文本)
                End If
            ElseIf Me.Element.要素表示 = 1 Then
                '上下表示
                If Trim(Me.Element.内容文本) <> "" And Me.Element.要素值域 <> ";" And Me.Element.要素值域 <> "0;0" Then
                    If Val(Me.Element.内容文本) > dblMax Then
                        Me.Element.内容文本 = dblMax
                    ElseIf Val(Me.Element.内容文本) < dblMin Then
                        Me.Element.内容文本 = dblMin
                    End If
                Else
                    Me.Element.内容文本 = IIf(Me.Element.要素小数 > 0, Format(Me.Element.内容文本, "0." & String(Me.Element.要素小数, "0")), Me.Element.内容文本)
                End If
            End If
        End If
        
        SaveSetting "ZLSOFT", "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\" & Me.Name, "MainWidth", Me.Width
        SaveSetting "ZLSOFT", "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\" & Me.Name, "MainHeight", Me.Height
        mblnOk = True
        Unload Me
    ElseIf KeyAscii = vbKeyEscape Then
        Form_Deactivate
    ElseIf KeyAscii = vbKeySpace Then
        If vfg单选.Visible Then vfg单选_KeyDown KeyAscii, 0
    ElseIf KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or KeyAscii = vbKeyUp Or KeyAscii = vbKeyDown Then
        If vfg单选.Visible Then vfg单选_KeyDown KeyAscii, 0
    End If
End Sub

Private Sub Form_Load()
    Me.KeyPreview = True
    Me.Width = 2000
End Sub

Private Sub Form_Paint()
    Cls
    Line (0, 0)-(ScaleWidth - Screen.TwipsPerPixelX, ScaleHeight - Screen.TwipsPerPixelY), &H996600, B
End Sub

Private Sub Form_Resize()
    Dim lX As Long, lY As Long
    lX = Screen.TwipsPerPixelX
    lY = Screen.TwipsPerPixelY
    
    txt上下1.Visible = False
    txt上下2.Visible = False
    lblDot.Visible = False
    lbl单位.Visible = False
    shpBorder2.Visible = False
    txt文本.Visible = False
    ud上下.Visible = False
    vfg单选.Visible = False
    vfg复选.Visible = False
    pic替换项目.Visible = False
    
    If Not Me.Element Is Nothing Then
        Select Case Me.Element.要素表示
        Case 0
            If Me.Element.替换域 = 1 Then
                pic替换项目.Move 80, picTitle.Height + 120, ScaleWidth - 160, ScaleHeight - 200 - picStatus.Height - picTitle.Height
                shpBorder1.Move pic替换项目.Left - lX, pic替换项目.Top - lY, pic替换项目.Width + lX * 2, pic替换项目.Height + lY * 2
                lbl示例.Height = Abs(pic替换项目.Height - lbl示例.Top)
                pic替换项目.Visible = True
                shpBorder1.Visible = True
            Else
                txt文本.Move 80, picTitle.Height + 120, ScaleWidth - 160, ScaleHeight - 200 - picStatus.Height - picTitle.Height
                shpBorder1.Move txt文本.Left - lX, txt文本.Top - lY, txt文本.Width + lX * 2, txt文本.Height + lY * 2
                txt文本.Visible = True
                shpBorder1.Visible = True
                If txt文本.Visible And txt文本.Enabled Then txt文本.SetFocus
            End If
        Case 1
            Dim lW1 As Long, lW2 As Long, lW3 As Long, lW4 As Long, lW5 As Long
            If Trim(Element.要素单位) <> "" Then
                lbl单位.Width = Me.TextWidth(lbl单位) + lX * 6
                lbl单位.Move Me.ScaleWidth - lbl单位.Width + lX * 3, picTitle.Height + 170
                lbl单位.Visible = True
                lW5 = lbl单位.Width
            Else
                lbl单位.Visible = False
                lW5 = 0
            End If
            lW4 = ud上下.Width + lX * 4
            ud上下.Move Me.ScaleWidth - lW4 - lW5 + lX * 3, picTitle.Height + 120
            ud上下.Visible = True
            If Element.要素小数 > 0 Then
                txt上下2.Width = Me.TextWidth(Space(Element.要素小数)) + lX * 4
                lW3 = txt上下2.Width + lX
                txt上下2.Move Me.ScaleWidth - lW5 - lW4 - lW3 + lX, picTitle.Height + 170
                shpBorder2.Move txt上下2.Left - lX, txt上下2.Top - lY - 50, txt上下2.Width + lX * 2, txt上下2.Height + 50 + lY * 2
                shpBorder2.Visible = True
                txt上下2.Visible = True
                lblDot.Width = Me.TextWidth(".") + lX * 2
                lW2 = lblDot.Width
                lblDot.Move txt上下2.Left - lW2 + lX * 2, picTitle.Height + 170
                lblDot.BackStyle = 0
                lblDot.Visible = True
            Else
                lW2 = 0
                lW3 = 0
                shpBorder2.Visible = False
                txt上下2.Visible = False
                lblDot.Visible = False
            End If
            lW1 = Me.TextWidth(txt上下1.Text) + lX * 2
            lW1 = IIf(lW1 < 400, 400, lW1)
            
            If Me.Width < lW1 + lW2 + lW3 + lW4 + lW5 Then Me.Width = lW1 + lW2 + lW3 + lW4 + lW5
            Me.Height = txt上下1.Height + lY * 3 + picStatus.Height + picTitle.Height + 180
            
            txt上下1.Move 80, picTitle.Height + 170, Me.ScaleWidth - lW5 - lW4 - lW3 - lW2 - lX * 4
            shpBorder1.Move txt上下1.Left - lX, txt上下1.Top - lY - 50, txt上下1.Width + lX * 2, txt上下1.Height + 50 + lY * 2
            txt上下1.Visible = True
            shpBorder1.Visible = True
            If txt上下1.Visible And txt上下1.Enabled Then txt上下1.SelStart = 0: txt上下1.SelLength = Len(txt上下1): txt上下1.SetFocus
        Case 2
            vfg单选.Move 80, picTitle.Height + 120, ScaleWidth - 160, ScaleHeight - 200 - picStatus.Height - picTitle.Height
            shpBorder1.Move vfg单选.Left - lX, vfg单选.Top - lY, vfg单选.Width + lX * 3, vfg单选.Height + lY * 2
            vfg单选.Visible = True
            shpBorder1.Visible = True
            If vfg单选.Visible And vfg单选.Enabled Then vfg单选.SetFocus
        Case 3
            vfg复选.Move 80, picTitle.Height + 120, ScaleWidth - 160, ScaleHeight - 200 - picStatus.Height - picTitle.Height
            shpBorder1.Move vfg复选.Left - lX, vfg复选.Top - lY, vfg复选.Width + lX * 3, vfg复选.Height + lY * 2
            vfg复选.Visible = True
            shpBorder1.Visible = True
            If vfg复选.Visible And vfg复选.Enabled Then vfg复选.SetFocus
        End Select
    End If
    
    picTitle.Move 60, 60, ScaleWidth - 120
    picStatus.Move lX, ScaleHeight - picStatus.Height - lY, ScaleWidth - lX * 2
    
    If Me.Top + Me.Height > Screen.Height - 800 Then Me.Top = Me.Top - Me.Height - 200
    If Me.Left + Me.Width > Screen.Width Then Me.Left = Me.Left - Me.Width
End Sub

Private Sub imgResize_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgResize.Tag = "Down"
    lngX = x
    lngY = y
End Sub

Private Sub imgResize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If imgResize.Tag = "Down" Then
        If Me.Width + x - lngX >= 1000 And Me.Width + x - lngX <= 12000 Then
            Me.Width = Me.Width + x - lngX
        End If
        If Me.Height + y - lngY >= 1000 And Me.Height + y - lngY <= 9000 Then
            Me.Height = Me.Height + y - lngY
        End If
        DoEvents
    End If
End Sub

Private Sub imgResize_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgResize.Tag = ""
    Call SetCtlFocus
End Sub

Private Sub imgTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgTitle.Tag = "Down"
    lngX = x
    lngY = y
End Sub

Private Sub imgTitle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If imgTitle.Tag = "Down" Then
        Me.Move Me.Left + x - lngX, Me.Top + y - lngY
    Else
        If x > 0 And x < picTitle.ScaleWidth And y > 0 And y < picTitle.ScaleHeight Then
            SetCapture picTitle.hWnd
            picTitle.Cls
            picTitle.BackColor = &HC2EEFF
            picTitle.Line (0, 0)-(picTitle.ScaleWidth - Screen.TwipsPerPixelX, picTitle.ScaleHeight - Screen.TwipsPerPixelY), &H800000, B
            lblInfo.Caption = "按下鼠标拖拽可以移动编辑器"
        Else
            ReleaseCapture
            picTitle.Cls
            picTitle.BackColor = &HF5BE9E
            lblInfo.Caption = "Esc:取消编辑。回车:保存修改。"
        End If
    End If
End Sub

Private Sub imgTitle_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgTitle.Tag = ""
    Call SetCtlFocus
End Sub

Private Sub picStatus_Resize()
    imgResize.Move picStatus.ScaleWidth - imgResize.Width, 0
End Sub

Private Sub picTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    picTitle.Tag = "Down"
    lngX = x
    lngY = y
End Sub

Private Sub picTitle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If picTitle.Tag = "Down" Then
        Me.Move Me.Left + x - lngX, Me.Top + y - lngY
    Else
        If x > 0 And x < picTitle.ScaleWidth And y > 0 And y < picTitle.ScaleHeight Then
            SetCapture picTitle.hWnd
            picTitle.Cls
            picTitle.BackColor = &HC2EEFF
            picTitle.Line (0, 0)-(picTitle.ScaleWidth - Screen.TwipsPerPixelX, picTitle.ScaleHeight - Screen.TwipsPerPixelY), &H800000, B
            lblInfo.Caption = "按下鼠标拖拽可以移动编辑器"
            If picTitle.Tag = "Down" Then
                Me.Move Me.Left + x - lngX, Me.Top + y - lngY
            End If
        Else
            ReleaseCapture
            picTitle.Cls
            picTitle.BackColor = &HF5BE9E
            lblInfo.Caption = "Esc:取消编辑。回车:保存修改。"
        End If
    End If
End Sub

Private Sub picTitle_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    picTitle.Tag = ""
    Call SetCtlFocus
End Sub

Private Sub picTitle_Resize()
    imgTitle.Move (picTitle.ScaleWidth - imgTitle.Width) / 2, 30
End Sub

Private Sub txt上下1_Change()
    If txt上下1.Tag = "" Then
        Me.Element.内容文本 = Trim(txt上下1.Text) & IIf(Me.Element.要素小数 > 0, "." & Format(Trim(txt上下2.Text), String(Me.Element.要素小数, "0")), "")
    End If
End Sub

Private Sub txt文本_Change()
    Me.Element.内容文本 = Trim(txt文本.Text)
End Sub

'#####################################################################################
'## 局部函数
'#####################################################################################

Private Sub SetCtlFocus()
    '设置控件焦点
    If txt上下1.Visible And txt上下1.Enabled Then
        txt上下1.SetFocus
    ElseIf txt上下2.Visible And txt上下2.Enabled Then
        txt上下2.SetFocus
    ElseIf txt文本.Visible And txt文本.Enabled Then
        txt文本.SetFocus
    ElseIf vfg单选.Visible And vfg单选.Enabled Then
        vfg单选.SetFocus
    ElseIf vfg复选.Visible And vfg复选.Enabled Then
        vfg复选.SetFocus
    End If
End Sub
