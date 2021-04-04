VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmEInvoice 
   BorderStyle     =   0  'None
   Caption         =   "电子票据管理"
   ClientHeight    =   6465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox chk住院 
      Caption         =   "住院预交"
      Height          =   255
      Left            =   2760
      TabIndex        =   15
      Top             =   6000
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chk门诊 
      Caption         =   "门诊预交"
      Height          =   255
      Left            =   1560
      TabIndex        =   14
      Top             =   6000
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmd客户端 
      Caption         =   "…"
      Height          =   280
      Left            =   4440
      TabIndex        =   12
      Top             =   705
      Width           =   280
   End
   Begin VB.TextBox txt客户端 
      Height          =   300
      Left            =   1080
      TabIndex        =   11
      Top             =   700
      Width           =   3345
   End
   Begin VB.CommandButton cmd配置 
      Caption         =   "电子票据配置"
      Height          =   300
      Left            =   7680
      TabIndex        =   10
      Top             =   6000
      Width           =   1335
   End
   Begin VB.OptionButton Option电子票据 
      Caption         =   "不启用电子票据"
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   4
      Top             =   240
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.OptionButton Option电子票据 
      Caption         =   "启用电子票据"
      Height          =   375
      Index           =   1
      Left            =   3000
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.OptionButton Option电子票据 
      Caption         =   "按客户端启用电子票据"
      Height          =   375
      Index           =   2
      Left            =   4680
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.CheckBox chk平台管理 
      Caption         =   "三方平台管理纸质票据"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5280
      TabIndex        =   1
      Top             =   6000
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   300
      Left            =   4800
      Picture         =   "frmEInvoice.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   700
      Width           =   300
   End
   Begin VSFlex8Ctl.VSFlexGrid VSF医保 
      Height          =   4620
      Left            =   5280
      TabIndex        =   5
      Top             =   1200
      Width           =   3780
      _cx             =   6667
      _cy             =   8149
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   2
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmEInvoice.frx":0A02
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
   Begin VSFlex8Ctl.VSFlexGrid VSF客户端 
      Height          =   4620
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   4860
      _cx             =   8572
      _cy             =   8149
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   2
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmEInvoice.frx":0AA3
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
   Begin VB.Label lbl预交 
      Caption         =   "启用预交类别"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   6030
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "挂号"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   300
      Width           =   615
   End
   Begin VB.Label lbl客户端 
      Alignment       =   1  'Right Justify
      Caption         =   "客户端"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lbl医保 
      Caption         =   "医保启用"
      Height          =   255
      Left            =   5280
      TabIndex        =   7
      Top             =   735
      Width           =   735
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000003&
      X1              =   6000
      X2              =   9720
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "frmEInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private mint场合 As Integer  '1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
Private mrs保险类别 As ADODB.Recordset
Private mstr挂号 As String, mstr收费 As String, mstr预交 As String, mstr结帐 As String
Private Type Para
    int类别  As Integer  '0-不区分；1-门诊；2-住院
    int启用电子票据  As Integer  '0-不启用电子票据；1-启用电子票据；2-分站点启用电子票据
    bln启用票据管理  As Boolean  'True-HIS管理票据；False-三方平台管理票据；
    str医保启用  As String  '字符串的格式为："0:"或"1:998"；0表示未启用，1表示启用，: 后边为险类(如998)，险类为空表示所有医保都启用
End Type
Private mPara As Para
Private Enum Page
    Pg_收费 = 1
    Pg_预交 = 2
    Pg_结帐 = 3
    Pg_挂号 = 4
    Pg_就诊卡 = 5
End Enum
Private mIndex As Integer

Private Sub chk门诊_Click()
    If chk住院.value = 0 Then
        If chk门诊.value = 0 Then
            MsgBox "必须设置至少一个预交类型!", vbInformation, gstrSysName
            chk门诊.value = 1
        End If
    End If
End Sub

Private Sub chk住院_Click()
    If chk门诊.value = 0 Then
        If chk住院.value = 0 Then
            MsgBox "必须设置至少一个预交类型!", vbInformation, gstrSysName
            chk住院.value = 1
        End If
    End If
End Sub

Private Sub cmdDelete_Click()
    If VSF客户端.Enabled = False Then Exit Sub
    If VSF客户端.Row <= 0 Then Exit Sub
    If mPara.int启用电子票据 = mIndex Then
        If CheckHaveData = False Then Exit Sub
    End If
    Call Delete客户端(VSF客户端.Row)
    zlcontrol.ControlSetFocus VSF客户端
End Sub

Private Sub cmd客户端_Click()
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    vRect = GetControlRect(txt客户端.hWnd)
    strSQL = "Select Rownum As id,Upper(工作站) as 工作站, Upper(用途) as 用途,Upper(部门) as 部门  From zlClients "
    Set rsTmp = zldatabase.ShowSQLSelect(Me, strSQL, 0, "读取客户端", 1, "", "请选择客户端", False, False, True, vRect.Left, vRect.Top, txt客户端.Height, blnCancel, False, False, "%" & Trim(txt客户端.Text) & "%", "bytSize=1")
    EnableWindow Me.hWnd, True '调用了ShowSQLSelect后会自动锁定窗口
    If rsTmp Is Nothing Then Exit Sub
    If rsTmp.State = 0 Then Exit Sub
    If rsTmp.EOF Then Exit Sub
    Call Add客户端(NVL(rsTmp!工作站), NVL(rsTmp!部门), NVL(rsTmp!用途))
End Sub

Private Sub cmd配置_Click()
    '调用电子发票设备参数配置接口
    Dim objEInvoice As Object
    
    If zlCreatEInvoice(objEInvoice, Me) = False Then Exit Sub
    If objEInvoice Is Nothing Then Exit Sub
    Call objEInvoice.zlEInvoiceSet(Me)
    Call objEInvoice.zlTerminate
    EnableWindow Me.hWnd, True '调用了该按钮后会自动锁定窗口
End Sub

Private Sub Form_GotFocus()
    If Option电子票据(0).value Then
        zlcontrol.ControlSetFocus Option电子票据(0)
    ElseIf Option电子票据(1).value Then
        zlcontrol.ControlSetFocus Option电子票据(1)
    Else
        zlcontrol.ControlSetFocus Option电子票据(2)
    End If
End Sub

Private Sub Form_Load()
    Call InitPara
    Call InitData
End Sub

Private Sub InitData()
    On Error GoTo ErrHand
    Call SetEnable
    If Get保险类别(mrs保险类别) = False Then Exit Sub
    Call Load客户端信息
    Call Load医保类别
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mint场合 = 0
End Sub

Private Function Get保险类别(ByRef rsTmp As ADODB.Recordset) As Boolean
    '功能：获取所有的保险类别
    Dim strSQL As String
    
    On Error GoTo ErrHand
    Set rsTmp = New ADODB.Recordset
    strSQL = "Select 序号,名称,说明,医院编码 From 保险类别 Where Nvl(是否禁止,0)=0 Order By 序号"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption)
    Get保险类别 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Load客户端信息()
    Dim i As Integer
    Dim strSQL  As String, rsData As ADODB.Recordset

     With VSF客户端
         If .Enabled Then
             .Clear 2
             strSQL = "Select b.工作站, b.部门, b.用途 From 电子票据站点控制 A, zlClients B Where a.站点 = b.工作站 And a.场合=[1] order by b.工作站 "
             Set rsData = zldatabase.OpenSQLRecord(strSQL, "读取电子票据站点控制", mint场合)
             If Not rsData.EOF Then
                 .Rows = rsData.RecordCount + 1
                 For i = 1 To rsData.RecordCount
                     .TextMatrix(i, .ColIndex("客户端名称")) = rsData!工作站
                     .TextMatrix(i, .ColIndex("部门")) = NVL(rsData!部门)
                     .TextMatrix(i, .ColIndex("用途")) = NVL(rsData!用途)
                     rsData.MoveNext
                 Next
             End If
         End If
     End With
End Sub

Private Sub Load医保类别()
    Dim i As Integer, j As Integer
    Dim str医保 As String, bln医保启用 As Boolean
    Dim varTmp As Variant
    
    If mrs保险类别 Is Nothing Then Exit Sub
    If mrs保险类别.RecordCount = 0 Then Exit Sub
    mrs保险类别.MoveFirst
    
    With VSF医保
        .Clear 2
        .Rows = mrs保险类别.RecordCount + 1
        varTmp = Split(mPara.str医保启用 & ":::", ":")
        bln医保启用 = varTmp(0) = 1: str医保 = varTmp(1)
        For j = 1 To mrs保险类别.RecordCount
            .TextMatrix(j, .ColIndex("保险类别")) = mrs保险类别!序号
            If bln医保启用 And InStr("," & str医保 & ",", "," & NVL(mrs保险类别!序号) & ",") > 0 Then
                .TextMatrix(j, .ColIndex("启用")) = "-1"
                .TextMatrix(j, .ColIndex("原启用")) = "1"
            End If
            .TextMatrix(j, .ColIndex("保险名称")) = mrs保险类别!名称
            mrs保险类别.MoveNext
        Next
    End With
End Sub

Private Function GetControlRect(ByVal lngHwnd As Long) As RECT
'功能：获取指定控件在屏幕中的位置(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function

Public Function Save电子票据控制() As Boolean
    Dim strSQL As String, i As Integer, j As Integer
    Dim str客户端 As String, blnTrans As Boolean
    Dim strTmp As String, intTmp As Integer
    Dim str险类 As String

    On Error GoTo ErrHand

    If Option电子票据(0).value Then
        strTmp = "0"
    ElseIf Option电子票据(1).value Then
        strTmp = "1"
    Else
        strTmp = "2"
    End If
    strTmp = strTmp & "|" & IIF(chk平台管理.value = 1, "1", "0")
    str险类 = ""
    With VSF医保
        For j = 1 To .Rows - 1
            If .TextMatrix(j, .ColIndex("启用")) = "-1" Then str险类 = str险类 & "," & .TextMatrix(j, .ColIndex("保险类别"))
        Next
    End With
    str险类 = Mid(str险类, 2)
    If str险类 = "" Then
        strTmp = strTmp & "|" & "0:"
    Else
        strTmp = strTmp & "|" & "1:" & str险类
    End If
    Select Case mint场合
        Case Pg_挂号
            zldatabase.SetPara "挂号电子票据控制", strTmp, glngSys
        Case Pg_收费
            zldatabase.SetPara "收费电子票据控制", strTmp, glngSys
        Case Pg_预交
            If chk门诊.value = 1 And chk住院.value = 1 Then
                intTmp = 0
            ElseIf chk门诊.value = 1 Then
                intTmp = 1
            Else
                intTmp = 2
            End If
            zldatabase.SetPara "预交电子票据控制", intTmp & "|" & strTmp, glngSys
        Case Pg_结帐
            zldatabase.SetPara "结帐电子票据控制", strTmp, glngSys
        Case Pg_就诊卡
            zldatabase.SetPara "就诊卡电子票据控制", strTmp, glngSys
        Case Else
    End Select

    str客户端 = ""
    With VSF客户端
        For j = 1 To .Rows - 1
            str客户端 = str客户端 & "," & .TextMatrix(j, .ColIndex("客户端名称"))
        Next
    End With
    str客户端 = Mid(str客户端, 2)
    strSQL = "Zl_电子票据站点控制_Update( " & mint场合 & "," & IIF(str客户端 = "", "NULL", "'" & str客户端 & "'") & ")"
    Call zldatabase.ExecuteProcedure(strSQL, Me.Caption)
    Save电子票据控制 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Check票据站点() As Boolean
    '功能:按客户端启用电子票据时,检查是否选择了客户端
    Dim str客户端 As String, i As Integer
    
    str客户端 = ""
    With VSF客户端
        For i = 1 To .Rows - 1
            str客户端 = str客户端 & "," & .TextMatrix(i, .ColIndex("客户端名称"))
        Next
    End With
     str客户端 = Mid(str客户端, 2)
     Check票据站点 = str客户端 <> ""
End Function

Private Sub Add客户端(ByVal str客户端 As String, ByVal str部门 As String, ByVal str用途 As String)
    Dim i As Integer
    If str客户端 = "" Then Exit Sub
    With VSF客户端
        For i = 1 To .Rows - 1
            If str客户端 = .TextMatrix(i, .ColIndex("客户端名称")) Then Exit Sub
        Next
        If .TextMatrix(.Rows - 1, .ColIndex("客户端名称")) <> "" Then .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, .ColIndex("客户端名称")) = str客户端
        .TextMatrix(.Rows - 1, .ColIndex("部门")) = str部门
        .TextMatrix(.Rows - 1, .ColIndex("用途")) = str用途
    End With
End Sub

Private Sub Delete客户端(ByVal intRow As Integer)
    Dim i As Integer
    If intRow = 0 Then Exit Sub
    With VSF客户端
        If intRow = 1 And .Rows = 2 Then
            .TextMatrix(intRow, .ColIndex("客户端名称")) = ""
            .TextMatrix(intRow, .ColIndex("部门")) = ""
            .TextMatrix(intRow, .ColIndex("用途")) = ""
        Else
            .RemoveItem intRow
        End If
    End With
End Sub

Private Sub InitPara()
    Dim strTmp As String, varTmp As Variant
    
    Select Case mint场合
        Case Pg_挂号
            strTmp = zldatabase.GetPara("挂号电子票据控制", glngSys, , "0|1|0:")
        Case Pg_收费
            strTmp = zldatabase.GetPara("收费电子票据控制", glngSys, , "0|1|0:")
        Case Pg_预交
            strTmp = zldatabase.GetPara("预交电子票据控制", glngSys, , "0|0|1|0:")
        Case Pg_结帐
            strTmp = zldatabase.GetPara("结帐电子票据控制", glngSys, , "0|1|0:")
        Case Pg_就诊卡
            strTmp = zldatabase.GetPara("就诊卡电子票据控制", glngSys, , "0|1|0:")
    End Select
    varTmp = Split(strTmp & "||||", "|")
    If mint场合 = Pg_预交 Then
        mPara.int类别 = varTmp(0)
        chk门诊.value = IIF(mPara.int类别 <> 2, 1, 0)
        chk住院.value = IIF(mPara.int类别 <> 1, 1, 0)
        mPara.int启用电子票据 = varTmp(1)
        mPara.bln启用票据管理 = varTmp(2) = 1
        mPara.str医保启用 = varTmp(3)
        Exit Sub
    End If
    mPara.int启用电子票据 = varTmp(0)
    mPara.bln启用票据管理 = varTmp(1) = 1
    mPara.str医保启用 = varTmp(2)
End Sub

Private Sub SetEnable(Optional ByVal intIndex As Integer = -1)
    Dim i As Integer, intTmp As Integer
    If intIndex = -1 Then
        intTmp = mPara.int启用电子票据
        Option电子票据(0).value = intTmp = 0
        Option电子票据(1).value = intTmp = 1
        Option电子票据(2).value = intTmp = 2
        cmd客户端.Enabled = intTmp = 2
        txt客户端.Enabled = intTmp = 2
        VSF客户端.Enabled = intTmp = 2
        VSF医保.Enabled = intTmp > 0
        chk平台管理.value = IIF(mPara.bln启用票据管理, 1, 0)
        cmdDelete.Enabled = intTmp = 2
    Else
        cmd客户端.Enabled = intIndex = 2
        txt客户端.Enabled = intIndex = 2
        VSF客户端.Enabled = intIndex = 2
        VSF医保.Enabled = intIndex > 0
        cmdDelete.Enabled = intIndex = 2
    End If
    
End Sub

Private Sub Clear客户端信息()
    '功能:清空客户端信息
    txt客户端.Text = ""
    VSF客户端.Rows = 2
    VSF客户端.Clear 2
End Sub

Private Sub Clear医保类别()
    '功能:清空医保类别
    VSF医保.Cell(flexcpText, 1, 0, VSF医保.Rows - 1, 0) = 0
End Sub

Public Sub InitMe(ByVal int场合 As Integer)
    '功能：初始化窗体
    mint场合 = int场合
    Select Case int场合
        Case Pg_挂号
            lbl.Caption = "挂号"
        Case Pg_收费
            lbl.Caption = "收费"
        Case Pg_预交
            lbl.Caption = "预交"
            lbl预交.Visible = True: chk门诊.Visible = True: chk住院.Visible = True
        Case Pg_结帐
            lbl.Caption = "结帐"
        Case Pg_就诊卡
            lbl.Caption = "就诊卡"
    End Select
End Sub

Private Function zlCreatEInvoice(ByRef objEInvoice As Object, ByVal frmMain As Object) As Boolean
    Dim strExtend As String
    err = 0: On Error Resume Next
    Set objEInvoice = CreateObject("zlPublicExpense.clsPubEInvoice")
    If err <> 0 Then
        MsgBox "不存在可用的电子票据接口部件(zlPublicExpense.clsPubEInvoice)，请与系统管理员联系,详细的错误信息为:" & vbCrLf & err.Description, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    zlCreatEInvoice = objEInvoice.zlInitialize(frmMain, 0, gcnOracle, glngSys, glngModul, True, strExtend)
End Function

Private Sub Option电子票据_Click(Index As Integer)
    If lbl.Tag = "1" Then Exit Sub
    If RemindUser(Index) = False Then
        lbl.Tag = "1"
        Option电子票据(mIndex).value = True
        lbl.Tag = "": Exit Sub
    End If
    Select Case Index
        Case 2
            SetEnable (Index)
            Call Load客户端信息
            Call Load医保类别
        Case 1
            Call SetEnable(Index)
            Call Clear客户端信息
            Call Load医保类别
        Case Else
            Call Clear客户端信息
            Call Clear医保类别
            Call SetEnable(Index)
    End Select
    mIndex = Index
End Sub

Private Sub txt客户端_KeyPress(KeyAscii As Integer)
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    If KeyAscii <> 13 Then Exit Sub
    vRect = GetControlRect(txt客户端.hWnd)
    strSQL = "Select Rownum As id,Upper(工作站) as 工作站, Upper(用途) as 用途,Upper(部门) as 部门  From zlClients " & _
                  "Where 工作站 Like Upper([1]) Or 用途 Like Upper([1]) Or 部门 Like Upper([1]) " & _
                  "   Or Upper(zlPinYinCode(工作站)) Like Upper([1]) Or Upper(zlPinYinCode(用途)) Like Upper([1]) Or Upper(zlPinYinCode(部门)) Like Upper([1]) Order By 工作站 "
    Set rsTmp = zldatabase.ShowSQLSelect(Me, strSQL, 0, "读取客户端", 1, "", "请选择客户端", False, False, True, vRect.Left, vRect.Top, txt客户端.Height, blnCancel, False, False, "%" & Trim(txt客户端.Text) & "%", "bytSize=1")
    EnableWindow Me.hWnd, True '调用了ShowSQLSelect后会自动锁定窗口
    If rsTmp Is Nothing Then Exit Sub
    If rsTmp.State = 0 Then Exit Sub
    If rsTmp.EOF Then Exit Sub
    Call Add客户端(NVL(rsTmp!工作站), NVL(rsTmp!部门), NVL(rsTmp!用途))
    
End Sub

Private Function RemindUser(ByVal intIndex As Integer) As Boolean
    '用户切换启用电子票据方式时提示用户
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strTmp As String, i As Integer
    Dim str场合 As String, int票种 As Integer  '1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
    
    If CheckHaveData = False Then Exit Function
    If intIndex > mIndex Then RemindUser = True: Exit Function
    '检查参数值是否改变
    With VSF客户端
        For i = 1 To .Rows - 1
            strTmp = strTmp & "," & .TextMatrix(i, .ColIndex("客户端名称"))
        Next
    End With
    strTmp = Mid(strTmp, 2)
    If strTmp <> "" Then
        If MsgBox("改变【电子票据启用方式】后，之前的修改将会清空。" & vbCrLf & "是否确认改变？", _
            vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            Else
                RemindUser = True: Exit Function
            End If
    End If
    
    If intIndex = 1 Then RemindUser = True: Exit Function
    
    With VSF医保
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("启用")) = "-1" Then strTmp = strTmp & "," & .TextMatrix(i, .ColIndex("保险类别"))
        Next
    End With
    strTmp = Mid(strTmp, 2)
    If strTmp <> "" Then
        If MsgBox("改变【电子票据启用方式】后，之前的修改将会清空。" & vbCrLf & "是否确认改变？", _
            vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    RemindUser = True
End Function

Private Function CheckHaveData() As Boolean
    '用户切换启用电子票据方式时检查是否已经产生数据
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim str场合 As String, int票种 As Integer  '1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
    If mIndex = 0 Then CheckHaveData = True: Exit Function
    
    '检查是否已经产生了数据
    strSQL = "Select 1 From 电子票据使用记录 Where 票种 = [1] And Rownum < 2 "
    Select Case mint场合
        Case 1 '挂号
            int票种 = 4
            str场合 = "挂号"
        Case 2 '收费
            int票种 = 1
            str场合 = "收费"
        Case 3 '预交
            int票种 = 2
            str场合 = "预交"
        Case 4 '结帐
            int票种 = 3
            str场合 = "结帐"
        Case 5 '就诊卡
            int票种 = 5
            str场合 = "就诊卡"
    End Select
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, int票种)
    If Not rsTmp.EOF Then
        If MsgBox(str场合 & "业务已产生电子票据使用记录，如果调整此参数，将会影响到" & str场合 & "业务的票据使用及打印，还有可能造成票据使用数据的混乱。" & vbCrLf & "是否确认调整参数？", _
            vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        Else
            CheckHaveData = True: Exit Function
        End If
    End If
    
    CheckHaveData = True
End Function

Private Sub VSF医保_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then
        With VSF医保
        If .TextMatrix(Row, Col) = "-1" And .TextMatrix(Row, .ColIndex("原启用")) = "1" Then
            If CheckHaveData = False Then
                Cancel = True
            End If
        End If
        End With
    End If
End Sub

Public Function Check电子票据Valid() As Boolean
    If Option电子票据(2).value Then
        If Check票据站点 = False Then
            MsgBox "按客户端启用电子票据时,必须设置至少一个客户端!", vbInformation, gstrSysName
            zlcontrol.ControlSetFocus cmd客户端
            Exit Function
        End If
    End If
    Check电子票据Valid = True
End Function
