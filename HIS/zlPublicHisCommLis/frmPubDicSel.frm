VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmPubDicSel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选择"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5295
   Icon            =   "frmPubDicSel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   300
      Left            =   3750
      TabIndex        =   5
      Top             =   4980
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   300
      Left            =   2490
      TabIndex        =   4
      Top             =   4980
      Width           =   1100
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4845
      Left            =   570
      ScaleHeight     =   4815
      ScaleWidth      =   3945
      TabIndex        =   0
      Top             =   90
      Width           =   3975
      Begin VB.TextBox txtFilter 
         Height          =   285
         Left            =   690
         TabIndex        =   3
         Text            =   "txtFilter"
         Top             =   150
         Width           =   3045
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   3915
         Left            =   270
         TabIndex        =   1
         Top             =   660
         Width           =   3405
         _cx             =   6006
         _cy             =   6906
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
      Begin VB.Label lblFilter 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "过滤"
         Height          =   180
         Left            =   180
         TabIndex        =   2
         Top             =   180
         Width           =   360
      End
   End
   Begin VB.Label lblNotic 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "使用""↑""""↓""键选择项目"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   360
      TabIndex        =   6
      Top             =   5040
      Width           =   1980
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPubDicSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'创    建:蔡青松
'创建时间:2018/8/16
'模块功能:公共选择器
'---------------------------------------------------------------------------------------

Option Explicit

Private mstrRetur As String         '返回的内容
Private mstrHiddenID As String      '不需要显示的记录
Private mblnShowCheckBox As Boolean '是否显示复选框

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/8/16
'功    能:显示公共选择器
'入    参:
'           objfrm      调用来源
'           rsTmp       需要展示的数据来源
'           strFilter   需要过滤的内容
'           lngID       默认过滤ID，如果记录集中包含ID字段的话
'           intShowCol  需要展示多少列数据，从第0列开始依次往后数
'           strHiddenID 需要隐藏的行,多个ID使用","分割
'           blnShowCheckBox     是否显示复选框，若显示复选框，则表示可以多选

'出    参:
'返    回:  选择的内容，每列之间使用“;”分隔
'调整影响:
'---------------------------------------------------------------------------------------
Public Function ShowMe(objFrm As Object, ByVal rsTmp As ADODB.Recordset, Optional ByVal strFilter As String, _
                       Optional ByVal lngID As String, Optional ByVal intShowCol As Integer, _
                       Optional ByVal strHiddenID As String, Optional ByVal blnShowCheckBox As Boolean) As String

          Dim lngCount As Long
          Dim blnHaveChk As Boolean   '记录集中是否已经存在选择列
          Dim i As Integer

1         On Error GoTo showMe_Error

2         mstrRetur = ""
3         mstrHiddenID = strHiddenID
4         mblnShowCheckBox = blnShowCheckBox

5         If mblnShowCheckBox Then intShowCol = intShowCol + 1

          '绑定数据
6         If SetDataToVSF(Me.VSFList, rsTmp) = False Then
7             Unload Me
8             Exit Function
9         End If

10        With Me.VSFList
              '隐藏不需要显示的列
11            For i = 0 To .Cols - 1
                  '隐藏ID列
12                If UCase(.ColKey(i)) Like "*ID*" Then
13                    .ColHidden(i) = True
14                End If
                  '设置只显示多少列，从第0列开始依次往后数
15                If intShowCol <= i Then
16                    .ColHidden(i) = True
17                    lngCount = lngCount + 1
18                End If
19                If UCase(.ColKey(i)) Like "选择" Then
20                    blnHaveChk = True
21                End If
22            Next

23            Me.txtFilter.Text = strFilter

              '隐藏行
24            Call SetHiddenRow

              '显示选择列
25            If Not blnHaveChk Then
26                .Cols = .Cols + 1
27                .ColKey(.Cols - 1) = "选择"
28                .ColWidth(.ColIndex("选择")) = 800
29                .ColHidden(.ColIndex("选择")) = Not blnShowCheckBox
30                .Cell(flexcpAlignment, 0, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = flexAlignCenterCenter
31                .TextMatrix(0, .ColIndex("选择")) = "选择"
32                .ColDataType(.ColIndex("选择")) = flexDTBoolean
33                .ColPosition(.ColIndex("选择")) = 0
34            End If

              '得到显示行总数
35            lngCount = 0
36            For i = .Rows - 1 To .FixedRows Step -1
37                If .RowHidden(i) = False Then
38                    lngCount = lngCount + 1
39                    .Row = i
40                End If
41            Next

42            If lngCount = 1 Then
                  '只有一行数据时直接返回
43                If blnShowCheckBox Then .Cell(flexcpChecked, .Row, .ColIndex("选择")) = 1
44                Call cmdOK_Click
45                Unload Me
46            Else
47                Me.Show vbModal, objFrm
48            End If
49        End With
50        ShowMe = mstrRetur


51        Exit Function
showMe_Error:
52        Call WriteErrLog("zlPublicHisCommLis", "frmPubDicSel", "执行(ShowMe)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
53        Err.Clear
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/9/3
'功    能:设置隐藏的行
'入    参:
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Private Function SetHiddenRow() As Long
    Dim lngCount As Long
    Dim i As Integer
    
    With Me.VSFList
        If .ColIndex("ID") < 0 Then Exit Function
        For i = 1 To .Rows - 1
            If InStr("," & mstrHiddenID & ",", "," & .TextMatrix(i, .ColIndex("ID")) & ",") > 0 Then
                .RowHidden(i) = True
            End If
        Next
        
         '默认选中第一个没有隐藏的行
        For i = 1 To .Rows - 1
            If .RowHidden(i) = False Then
                If lngCount = 0 Then .Row = i
                .ShowCell .Row, 0
                lngCount = lngCount + 1
                If lngCount >= 1 Then Exit For
            End If
        Next
    End With
    SetHiddenRow = lngCount
End Function

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case CONFUN_UP      '上一行
            Call UpOrDown(Me.VSFList, vbKeyUp)
        Case CONFUN_DOWN    '下一行
            Call UpOrDown(Me.VSFList, vbKeyDown)
    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strSub As String

    mstrRetur = ""
    With Me.VSFList
        For lngRow = 1 To .Rows - 1
            If .Cell(flexcpChecked, lngRow, .ColIndex("选择")) = 1 And .RowHidden(lngRow) = False Then
                For lngCol = 1 To .Cols - 1
                    strSub = strSub & "<SP2>" & .TextMatrix(lngRow, lngCol)
                Next
                If Left(strSub, 5) = "<SP2>" Then strSub = Mid(strSub, 6)
                mstrRetur = mstrRetur & "<SP1>" & strSub
                strSub = ""
            End If
        Next
        If mstrRetur = "" Then
            If .Row <= 0 Or .RowHidden(.Row) Then Exit Sub
            For lngCol = 1 To .Cols - 1
                mstrRetur = mstrRetur & "<SP2>" & .TextMatrix(.Row, lngCol)
            Next
        End If
    End With
    If mstrRetur <> "" Then
        If Mid(mstrRetur, 1, 5) = "<SP1>" Then mstrRetur = Mid(mstrRetur, 6)
        If Mid(mstrRetur, 1, 5) = "<SP2>" Then mstrRetur = Mid(mstrRetur, 6)
    End If
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        '按回车
        Call cmdOK_Click
    End If
    If KeyAscii = vbKeyEscape Then
        '按ESC
        Unload Me
    End If
End Sub

Private Sub Form_Load()
     Dim Menus As New Collection
     
    '设置隐藏工具栏
    Menus.Add CONFUN_UP & ",上一个(&UP),False"
    Menus.Add CONFUN_DOWN & ",下一个(&DN),True"
    Call CbsButtonInit(cbrMain, Menus, True, xtpBarTop)
    Set Menus = Nothing
     '快键绑定
    With Me.cbrMain.KeyBindings
        .Add 0, vbKeyUp, CONFUN_UP
        .Add 0, vbKeyDown, CONFUN_DOWN
    End With
    
    Me.txtFilter.TabIndex = 0
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With Me.picMain
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.cmdOK.Top - 100
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrHiddenID = ""
End Sub

Private Sub PicMain_Resize()
    On Error Resume Next
    With Me.lblFilter
        .Left = 100
        .Top = 100
    End With
    With Me.txtFilter
        .Left = Me.lblFilter.Left + Me.lblFilter.Width + 100
        .Top = 50
        .Width = Me.picMain.Width - .Left - 100
    End With
    With Me.VSFList
        .Left = 0
        .Top = Me.txtFilter.Top + Me.txtFilter.Height + 100
        .Width = Me.picMain.Width
        .Height = Me.picMain.Height - .Top
    End With
End Sub

Private Sub txtFilter_Change()
          Dim strFilter As String
          Dim lngRow As Long
          Dim lngCol As Long
          Dim lngRowFind As Long

1         On Error GoTo txtFilter_Change_Error

2         strFilter = UCase(Trim(Me.txtFilter.Text))
3         With Me.VSFList
4             For lngRow = 1 To .Rows - 1
5                 .RowHidden(lngRow) = True
6                 If .ColIndex("ID") > -1 Then
7                     If InStr("," & mstrHiddenID & ",", "," & .TextMatrix(lngRow, .ColIndex("ID")) & ",") <= 0 Then
8                         For lngCol = 0 To .Cols - 1
9                             If Not .ColHidden(lngCol) Then
10                                If UCase(.TextMatrix(lngRow, lngCol)) Like strFilter & "*" Then
11                                    .RowHidden(lngRow) = False
12                                End If
13                            End If
14                        Next
15                    End If
16                Else
17                    For lngCol = 0 To .Cols - 1
18                        If Not .ColHidden(lngCol) Then
19                            If UCase(.TextMatrix(lngRow, lngCol)) Like strFilter & "*" Then
20                                .RowHidden(lngRow) = False
21                            End If
22                        End If
23                    Next
24                End If
25            Next
26        End With


27        Exit Sub
txtFilter_Change_Error:
28        Call WriteErrLog("zlPublicHisCommLis", "frmPubDicSel", "执行(txtFilter_Change)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
29        Err.Clear
End Sub

Private Sub txtFilter_GotFocus()
    Call selAllText(txtFilter)
End Sub

Private Sub vsfList_DblClick()
    Dim lngRow As Long
    Dim lngCol As Long
    
    With VSFList
        lngRow = .MouseRow
        lngCol = .MouseCol
        If lngRow < 1 Or lngCol < 0 Then Exit Sub
        If .ColIndex("选择") >= 0 Then
            .Cell(flexcpChecked, lngRow, .ColIndex("选择")) = True
        End If
    End With
    Call cmdOK_Click
End Sub

Private Sub vsfList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long
    Dim lngCol As Long
    
    With Me.VSFList
        lngRow = .MouseRow
        lngCol = .MouseCol
        If .ColIndex("选择") < 0 Then Exit Sub
        If lngRow < 1 Or lngCol <> .ColIndex("选择") Then Exit Sub
        If .Cell(flexcpChecked, lngRow, lngCol) = 1 Then
            .Cell(flexcpChecked, lngRow, lngCol) = 0
        Else
            .Cell(flexcpChecked, lngRow, lngCol) = 1
        End If
    End With
End Sub
