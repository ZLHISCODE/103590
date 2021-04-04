VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSquareBrushCard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "消费卡刷卡"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10530
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSquareBrushCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   10530
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame2 
      Caption         =   "发卡信息"
      Height          =   2115
      Left            =   285
      TabIndex        =   23
      Top             =   420
      Width           =   3900
      Begin VB.TextBox txtEdit 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1155
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1155
         Width           =   2550
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1155
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   315
         Width           =   2550
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "当前余额"
         Height          =   240
         Index           =   10
         Left            =   90
         TabIndex        =   6
         Top             =   1680
         Width           =   960
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "密码(&W)"
         Height          =   240
         Index           =   2
         Left            =   210
         TabIndex        =   4
         Top             =   1245
         Width           =   840
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "卡号(&N)"
         Height          =   240
         Index           =   1
         Left            =   210
         TabIndex        =   0
         Top             =   390
         Width           =   840
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "卡类型"
         Height          =   240
         Index           =   0
         Left            =   330
         TabIndex        =   2
         Top             =   780
         Width           =   720
      End
      Begin VB.Label lblInfor 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Index           =   0
         Left            =   1155
         TabIndex        =   3
         Top             =   735
         Width           =   1005
      End
      Begin VB.Label lblInfor 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Index           =   1
         Left            =   1155
         TabIndex        =   7
         Top             =   1635
         Width           =   2550
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "本次刷卡消费"
      Height          =   2115
      Left            =   4290
      TabIndex        =   18
      Top             =   435
      Width           =   3990
      Begin VB.TextBox txtEdit 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1830
         TabIndex        =   9
         Top             =   1455
         Width           =   2025
      End
      Begin VB.Label lblInfor 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   1830
         TabIndex        =   22
         Top             =   360
         Width           =   2025
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "消费总额"
         Height          =   240
         Index           =   6
         Left            =   810
         TabIndex        =   21
         Top             =   435
         Width           =   960
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "本次最大刷卡额"
         Height          =   240
         Index           =   3
         Left            =   90
         TabIndex        =   20
         Top             =   975
         Width           =   1680
      End
      Begin VB.Label lblInfor 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   1830
         TabIndex        =   19
         Top             =   893
         Width           =   2025
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "本次消费(&X)"
         Height          =   240
         Index           =   4
         Left            =   450
         TabIndex        =   8
         Top             =   1530
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "删除当前行(&D)"
      Height          =   465
      Left            =   8535
      TabIndex        =   17
      Top             =   3285
      Width           =   1860
   End
   Begin VB.CommandButton cmd取消 
      Caption         =   "取消(&C)"
      Height          =   465
      Left            =   8655
      TabIndex        =   14
      Top             =   765
      Width           =   1185
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   465
      Left            =   8655
      TabIndex        =   13
      Top             =   255
      Width           =   1185
   End
   Begin VB.Frame fra 
      Height          =   3540
      Left            =   135
      TabIndex        =   15
      Top             =   75
      Width           =   8295
      Begin VB.TextBox txtEdit 
         Height          =   360
         Index           =   3
         Left            =   1095
         MaxLength       =   100
         TabIndex        =   11
         Top             =   2565
         Width           =   7035
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "继续刷下一张卡(&K)"
         Height          =   420
         Left            =   5835
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2970
         Width           =   2325
      End
      Begin VB.Label lbl失效额 
         Height          =   240
         Left            =   420
         TabIndex        =   24
         Top             =   3105
         Width           =   4455
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "备注(&S)"
         Height          =   240
         Index           =   5
         Left            =   180
         TabIndex        =   10
         Top             =   2640
         Width           =   840
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   3060
      Left            =   120
      TabIndex        =   16
      Top             =   3810
      Width           =   10305
      _cx             =   18177
      _cy             =   5397
      Appearance      =   1
      BorderStyle     =   1
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
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSquareBrushCard.frx":030A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   120
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
      ExplorerBar     =   7
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
End
Attribute VB_Name = "frmSquareBrushCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mstrPrivs As String, mintCallType As Integer, mintSucces As Integer, mblnChange As Boolean
Private mrsData As ADODB.Recordset, mrsFeeList As ADODB.Recordset
Private mlng接口编号 As Long
Private mblnCardNoSHowPW As Boolean '是否密文显示卡号
Private Type CardInfor
    lng消费卡ID As Long
    str卡号 As String
    dbl余额 As Double
    dbl最大消费额 As Double
    dbl失效面额 As Double '采取的原则是,先进先出的法则:先消费卡面额,再消费允值额:此金额为,到期后未消费的金额
    str限制类别 As String
    str接口名称 As String
    str结算方式 As String
End Type
Private mdbl最大消费额 As Double, mdbl已消费累计 As Double
Private mTyCurCardInfor As CardInfor
Private Enum mtxtIdx
    idx_txt卡号 = 0
    idx_txt密码 = 1
    idx_txt本次消费 = 2
    idx_txt备注 = 3
End Enum
Private Enum mlblIdx
    idx_lbl卡类型 = 0
    idx_lbl余额 = 1
    idx_lbl最大消费 = 2
    idx_lbl总费用 = 3
End Enum
Private mrsRequare As New ADODB.Recordset
Private WithEvents mobjBrushCard As clsBrushSequareCard
Attribute mobjBrushCard.VB_VarHelpID = -1
Private mobjKeyboard As Object
 

Private Function CheckDepended() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查数据的关联性
    '编制:刘兴洪
    '日期:2009-12-24 12:13:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    Dim rsTemp As New ADODB.Recordset
    Set rsTemp = zlGet消费卡接口
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    rsTemp.Find "编号=" & mlng接口编号, , , 1
    If rsTemp.EOF Then
        ShowMsgbox "接口未找到(编号为" & mlng接口编号 & "),请检查!"
        Exit Function
    End If
    With mTyCurCardInfor
        .str接口名称 = Nvl(rsTemp!名称)
        .str结算方式 = Nvl(rsTemp!结算方式)
        txtEdit(mtxtIdx.idx_txt卡号).MaxLength = Len(Nvl(rsTemp!前缀文本)) + Val(Nvl(rsTemp!卡号长度))
    End With
    CheckDepended = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function zlShowBrushCard(ByVal frmMain As Object, ByVal lng接口编号 As Long, ByVal intCallType As Integer, _
    ByVal rsFeeList As ADODB.Recordset, dbl最大消费额 As Double, rsRequare As ADODB.Recordset) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷卡接口
    '入参:frmMain-调用的主窗体
    '     lngModule-调用的模块号
    '     strPrivs-调用的权限串
    '     dbl最大消费额-本次刷卡的最大刷卡额
    '     rsFeeList-费用详细信息()
    '出参:rsRequare-返回结算信息
    '返回:调用成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-12-23 10:33:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    mdbl最大消费额 = dbl最大消费额: mlng接口编号 = lng接口编号: mintCallType = intCallType
    Set mrsFeeList = rsFeeList  '费用明细:
    Set mrsRequare = rsRequare
    
    If CheckDepended = False Then Exit Function
        
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    zlShowBrushCard = mintSucces > 0
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
'将上次刷卡的信息加载到网格
Private Function LoadPreBrushCardToVsGrid() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:将上次刷卡的信息加载到网格
    '返回:加成成功,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2009-12-24 11:42:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, dbl结算金额 As Double
    
    Err = 0: On Error GoTo ErrHand:
    lngRow = 1
    With vsGrid
        .Rows = 2
        .Clear 1
        .Cell(flexcpForeColor, 1, 0, 1, .Cols - 1) = &H80000008
        lngRow = 1
        If mrsRequare.RecordCount <> 0 Then mrsRequare.MoveFirst
        
        Do While Not mrsRequare.EOF
            If Val(Nvl(mrsRequare!接口编号)) = mlng接口编号 Then
                If lngRow > 1 Then
                    .Rows = .Rows + 1
                End If
                
                .TextMatrix(lngRow, .ColIndex("卡号")) = IIf(mblnCardNoSHowPW, "*******", Nvl(mrsRequare!卡号))
                .Cell(flexcpData, lngRow, .ColIndex("卡号")) = Nvl(mrsRequare!消费卡ID) & "-" & Nvl(mrsRequare!卡号)
                
                .TextMatrix(lngRow, .ColIndex("结算方式")) = Nvl(mrsRequare!结算方式)
                .Cell(flexcpData, lngRow, .ColIndex("结算方式")) = Nvl(mrsRequare!卡名称)
                .TextMatrix(lngRow, .ColIndex("卡余额")) = Format(Val(Nvl(mrsRequare!余额)), gVbFmtString.FM_金额)
                .Cell(flexcpData, lngRow, .ColIndex("卡余额")) = Val(Nvl(mrsRequare!余额))
                .TextMatrix(lngRow, .ColIndex("本次消费")) = Format(Val(Nvl(mrsRequare!结算金额)), gVbFmtString.FM_金额)
                dbl结算金额 = dbl结算金额 + Val(Nvl(mrsRequare!结算金额))
                .TextMatrix(lngRow, .ColIndex("备注")) = Nvl(mrsRequare!备注)
                lngRow = lngRow + 1
            End If
            mrsRequare.MoveNext
        Loop
        mdbl已消费累计 = dbl结算金额
         If .Rows - 1 >= 1 And dbl结算金额 <> 0 Then
            If .TextMatrix(1, .ColIndex("卡号")) <> "" Then
                .Rows = .Rows + 1
                lngRow = .Rows - 1
                .TextMatrix(lngRow, .ColIndex("卡号")) = "合计"
                .Cell(flexcpData, lngRow, .ColIndex("卡号")) = ""
                .TextMatrix(lngRow, .ColIndex("卡余额")) = ""
                .Cell(flexcpData, lngRow, .ColIndex("卡余额")) = ""
                .TextMatrix(lngRow, .ColIndex("本次消费")) = Format(dbl结算金额, gVbFmtString.FM_金额)
                .TextMatrix(lngRow, .ColIndex("备注")) = ""
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbBlue
            End If
        End If
    End With
    LoadPreBrushCardToVsGrid = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function


Private Sub cmdDel_Click()
    Call MoveVsGridRowData
End Sub

Private Sub cmdNext_Click()
    If zlInsertDataToGrid = False Then Exit Sub
    zlCtlSetFocus txtEdit(mtxtIdx.idx_txt卡号)
End Sub

Private Sub cmd取消_Click()
    mintSucces = 0: Unload Me
End Sub
Private Sub cmd确定_Click()
    Dim lngRow As Long, blnInputNotGrid As Boolean  '输入的信息,没有在网格当中体现,直接处理
    Dim dt交易时间 As Date, blnHaveData As Boolean
    blnInputNotGrid = False
    If Trim(txtEdit(mtxtIdx.idx_txt卡号).Text) <> "" Then
        '需要检查
        If CheckCardNotExists(Trim(txtEdit(mtxtIdx.idx_txt卡号).Text), False) Then
            '不存在,表示需要检查是否合法
            If CheckInput = False Then Exit Sub
            blnInputNotGrid = True
        End If
         
    Else
        '检查是否有数据
        blnHaveData = False
        With vsGrid
            For lngRow = 1 To .Rows - 1
                If Val(Split(.Cell(flexcpData, lngRow, .ColIndex("卡号")) & "-", "-")(0)) <> 0 Then
                    blnHaveData = True: Exit For
                End If
            Next
        End With
'        If blnHaveData = False Then
'            ShowMsgbox "不存在刷卡数据,请检查"
'            zlCtlSetFocus txtEdit(mtxtIdx.idx_txt卡号)
'            Exit Sub
'        End If
    End If
   
    '返回相关的结算信息
    '先删除
    With mrsRequare
        .Filter = 0
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Val(Nvl(mrsRequare!接口编号)) = mlng接口编号 Then
               .Delete adAffectCurrent
               .Update
               .MoveNext
               If .RecordCount <> 0 Then .MoveFirst
            Else
               .MoveNext
            End If
        Loop
    End With
    
    Dim varData As Variant
    dt交易时间 = zlDatabase.Currentdate
    With vsGrid
        
        For lngRow = 1 To .Rows - 1
            varData = Split(.Cell(flexcpData, lngRow, .ColIndex("卡号")) & "-", "-")
            If Val(varData(0)) <> 0 Then
                mrsRequare.AddNew
                mrsRequare!接口编号 = mlng接口编号
                mrsRequare!消费卡ID = Val(varData(0))
                mrsRequare!卡号 = Trim(varData(1))
                mrsRequare!结算方式 = mTyCurCardInfor.str结算方式
                mrsRequare!卡名称 = mTyCurCardInfor.str接口名称
                mrsRequare!余额 = Val(.Cell(flexcpData, lngRow, .ColIndex("卡余额")))
                mrsRequare!结算金额 = Val(.TextMatrix(lngRow, .ColIndex("本次消费")))
                mrsRequare!交易时间 = dt交易时间
                mrsRequare!备注 = Trim(.TextMatrix(lngRow, .ColIndex("备注")))
                mrsRequare!结算标志 = 0
                mrsRequare.Update
            End If
        Next
    End With
    If blnInputNotGrid Then
        mrsRequare.AddNew
        mrsRequare!接口编号 = mlng接口编号
        mrsRequare!消费卡ID = mTyCurCardInfor.lng消费卡ID
        mrsRequare!卡号 = mTyCurCardInfor.str卡号
        mrsRequare!结算方式 = mTyCurCardInfor.str结算方式
        mrsRequare!卡名称 = mTyCurCardInfor.str接口名称
        mrsRequare!余额 = mTyCurCardInfor.dbl余额
        mrsRequare!结算金额 = Val(txtEdit(mtxtIdx.idx_txt本次消费).Text)
        mrsRequare!交易时间 = dt交易时间
        mrsRequare!备注 = Trim(txtEdit(mtxtIdx.idx_txt备注).Text)
        mrsRequare!结算标志 = 0
        mrsRequare.Update
    End If
    mintSucces = mintSucces + 1
    mblnChange = False
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim dbl总费用 As Double
    '检查是否启用了相关的刷卡程序
     Call CreateObjectKeyboard
    Set mobjBrushCard = New clsBrushSequareCard
    Call mobjBrushCard.zlInitInterFacel(mlng接口编号)
    '获取总额
    mblnCardNoSHowPW = zlIsCardNoShowPW(mlng接口编号)
    If mblnCardNoSHowPW Then
        txtEdit(mtxtIdx.idx_txt卡号).PasswordChar = "*"
    Else
        txtEdit(mtxtIdx.idx_txt卡号).PasswordChar = ""
    End If
    
    Call LoadPreBrushCardToVsGrid
    lblInfor(mlblIdx.idx_lbl总费用).Caption = Format(grsStatic.dbl费用总额, gVbFmtString.FM_金额)
    Call vsGrid_LostFocus
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index <> mtxtIdx.idx_txt密码 Then txtEdit(Index).Tag = ""
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    Select Case Index
    Case mtxtIdx.idx_txt卡号
        gTy_TestBug.BytType = 2
        If Not mobjBrushCard Is Nothing Then Call mobjBrushCard.zlSetAutoBrush(Trim(txtEdit(Index).Text) = "")
        zlControl.TxtSelAll txtEdit(Index)
    Case mtxtIdx.idx_txt备注
        zlControl.TxtSelAll txtEdit(Index)
        zlCommFun.OpenIme True
    Case Else
        zlControl.TxtSelAll txtEdit(Index)
        zlCommFun.OpenIme False
        If Index = idx_txt密码 Then
            Call OpenPassKeyboard(txtEdit(Index))
        End If
    End Select
End Sub
Private Function CheckInputPassWord() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查密码输入是否正确
    '返回:正确,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-12-23 14:27:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Trim(txtEdit(mtxtIdx.idx_txt密码).Tag) <> "" And Trim(txtEdit(mtxtIdx.idx_txt密码).Text) = "" Then
        ShowMsgbox "密码未输入,请检查!"
        Exit Function
    End If
    
    If Trim(txtEdit(mtxtIdx.idx_txt密码).Tag) <> Trim(txtEdit(mtxtIdx.idx_txt密码).Text) Then
        ShowMsgbox "密码输入错误,请检查!"
        Exit Function
    End If
    CheckInputPassWord = True
End Function

Private Function CheckInputSquareMoney() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查输入的本次消费金额是否正确
    '返回:正确,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-12-23 14:27:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If zlDblIsValid(Trim(txtEdit(mtxtIdx.idx_txt本次消费).Text), 16, True, True, 0, "卡面额") = False Then
        Exit Function
    End If
    If Val(lblInfor(mlblIdx.idx_lbl最大消费).Caption) < Val(Trim(txtEdit(mtxtIdx.idx_txt本次消费).Text)) Then
        ShowMsgbox "本次最多能消费:" & Format(Val(lblInfor(mlblIdx.idx_lbl最大消费).Tag), gVbFmtString.FM_金额) & "元,请检查!"
        Exit Function
    End If
    If Val(lblInfor(mlblIdx.idx_lbl余额).Caption) < Val(Trim(txtEdit(mtxtIdx.idx_txt本次消费).Text)) Then
        ShowMsgbox "卡余额不足(" & Format(Val(lblInfor(mlblIdx.idx_lbl余额).Caption), gVbFmtString.FM_金额) & "元),请检查!"
        Exit Function
    End If
    
    CheckInputSquareMoney = True
End Function
Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim str编码 As String, str名称 As String, lngID As Long
    Dim strCardNo As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    Select Case Index
    Case mtxtIdx.idx_txt卡号
        If txtEdit(Index).Tag <> "" Then zlCommFun.PressKey vbKeyTab
         
        '考虑可能存在操作员乱刷卡的情况,因此暂不开放如下功能:
        If IsDesinMode = False Then Exit Sub
        If txtEdit(Index).Text = "" Then
            '直接调读卡
            If mobjBrushCard.zlReadCard(Me, strCardNo) = False Then
                Exit Sub
            End If
            txtEdit(Index).Text = strCardNo
            txtEdit(Index).Tag = strCardNo
        End If
        
        If zlBrusCard(Trim(txtEdit(Index))) = False Then
            zlCtlSetFocus txtEdit(Index)
        Else
            If txtEdit(mtxtIdx.idx_txt密码).Tag = "" Then
                If txtEdit(mtxtIdx.idx_txt备注).Enabled And txtEdit(mtxtIdx.idx_txt备注).Visible Then txtEdit(mtxtIdx.idx_txt备注).SetFocus
            Else
                zlCommFun.PressKey vbKeyTab
            End If
        End If
        
    Case mtxtIdx.idx_txt备注
        zlCommFun.PressKey vbKeyTab
        Exit Sub
    Case mtxtIdx.idx_txt密码
        If CheckInputPassWord = False Then
            zlControl.ControlSetFocus txtEdit(Index): Exit Sub
        End If
        zlCommFun.PressKey vbKeyTab
    Case mtxtIdx.idx_txt本次消费
        If CheckInputSquareMoney = False Then
            zlControl.ControlSetFocus txtEdit(Index): Exit Sub
        End If
        zlCommFun.PressKey vbKeyTab
    Case Else
        zlCommFun.PressKey vbKeyTab
    End Select
End Sub
Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim blnCard As Boolean
    
    Select Case Index
    Case mtxtIdx.idx_txt卡号
        If InStr(1, "'~～|`-'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        If IsDesinMode Then Exit Sub
        
        Call BrushCard(txtEdit(Index), KeyAscii)
    Case mtxtIdx.idx_txt备注
        blnCard = zlInputIsCard(txtEdit(Index), KeyAscii, glngSys, mblnCardNoSHowPW)
        If blnCard = True Then KeyAscii = 0
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m文本式
    Case mtxtIdx.idx_txt密码
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m文本式
    Case mtxtIdx.idx_txt本次消费
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m金额式
    Case Else
    End Select
End Sub
Private Sub BrushCard(ByVal objEdit As Object, KeyAscii As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷卡操作(目前只支持有卡进行刷卡)
    '编制:刘兴洪
    '日期:2010-02-09 14:07:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Static sngBegin As Single
    Dim sngNow As Single
    Dim blnCard As Boolean
    
    '是否刷卡完成
    blnCard = KeyAscii <> 8 And Len(objEdit.Text) = objEdit.MaxLength - 1 And objEdit.SelLength <> Len(objEdit.Text)
    
    If blnCard Then
        If KeyAscii <> 13 Then
            objEdit.Text = objEdit.Text & Chr(KeyAscii)
            objEdit.SelStart = Len(objEdit.Text)
        End If
        KeyAscii = 0
        '刷卡处理:
        If zlBrusCard(Trim(objEdit)) = False Then
            zlCtlSetFocus objEdit
        Else
            If txtEdit(mtxtIdx.idx_txt密码).Tag = "" Then
                If txtEdit(mtxtIdx.idx_txt备注).Enabled And txtEdit(mtxtIdx.idx_txt备注).Visible Then txtEdit(mtxtIdx.idx_txt备注).SetFocus
            Else
                zlCommFun.PressKey vbKeyTab
            End If
        End If
    Else
        If InStr(":：;；?？" & Chr(22), Chr(KeyAscii)) > 0 Then
            KeyAscii = 0 '去除特殊符号，并且不允许粘贴
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
        '安全刷卡检测
        If KeyAscii <> 0 And KeyAscii > 32 Then
            sngNow = Timer
            If objEdit.Text = "" Then
                sngBegin = sngNow
            ElseIf Format((sngNow - sngBegin) / (Len(objEdit.Text) + 1), "0.000") >= 0.04 Then '>0.007>=0.01
                objEdit.Text = Chr(KeyAscii)
                objEdit.SelStart = 1
                KeyAscii = 0
                sngBegin = sngNow
            End If
        End If
    End If
End Sub
Private Function CheckIsBreshCard(ByVal objEdit As Object, KeyAscii As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否刷卡操作
    '返回:是刷卡,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-10-25 09:52:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Static sngBegin As Single
    Dim sngNow As Single
    Dim blnCard As Boolean
    If InStr(":：;；?？" & Chr(22), Chr(KeyAscii)) > 0 Then
        KeyAscii = 0 '去除特殊符号，并且不允许粘贴
    End If
    '安全刷卡检测
    If KeyAscii <> 0 And KeyAscii > 32 Then
        sngNow = Timer
        If objEdit.Text = "" Then
            sngBegin = sngNow
        ElseIf Format((sngNow - sngBegin) / (Len(objEdit.Text) + 1), "0.000") > 0.04 Then '>0.007>=0.01
            '不是刷卡的
            blnCard = False
             sngBegin = sngNow
        Else
            blnCard = True
        End If
    End If
End Function

Private Sub txtEdit_LostFocus(Index As Integer)
    If Index = idx_txt密码 Then
        Call ClosePassKeyboard(txtEdit(Index))
    End If
End Sub

Private Sub txtEdit_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index <> mtxtIdx.idx_txt卡号 Then Exit Sub
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtEdit(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtEdit(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtEdit_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index <> mtxtIdx.idx_txt卡号 Then Exit Sub
    If Button = 2 Then
        Call SetWindowLong(txtEdit(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtEdit_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
    Case mtxtIdx.idx_txt卡号
    Case mtxtIdx.idx_txt备注
    Case mtxtIdx.idx_txt密码
        If CheckInputPassWord = False Then
        End If
    Case mtxtIdx.idx_txt本次消费
        If CheckInputSquareMoney = False Then
           'Cancel = 1
        End If
    Case Else
    End Select
End Sub

Private Function zlBrusCard(ByVal strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷卡操作
    '编制:刘兴洪
    '日期:2009-12-16 10:33:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    Dim rsTemp As ADODB.Recordset, i As Long, strTemp As String, blnFind As Boolean
    
    With mTyCurCardInfor
        .dbl失效面额 = 0
        .dbl余额 = 0
        .dbl最大消费额 = 0
        .str卡号 = ""
        .lng消费卡ID = 0
    End With
    
    gstrSQL = "" & _
    "   Select a.Id,a.卡类型,a.卡号,a.序号,a.可否充值,to_char(a.有效期,'yyyy-mm-dd hh24:mi:ss') as 有效期,  a.密码," & _
    "          to_char(a.回收时间,'yyyy-mm-dd hh24:mi:ss') as 回收时间 , " & _
    "          decode(a.当前状态,2,'回收',3,'退卡','回收') as 当前状态, " & _
    "          to_char(a.卡面金额," & gOraFmtString.FM_金额 & ") as 卡面金额 ," & _
    "          to_char(a.销售金额," & gOraFmtString.FM_金额 & ") as 销售金额 ," & _
    "          to_char(a.充值折扣率," & gOraFmtString.FM_折扣率 & ") as 充值折扣率 ," & _
    "          to_char(a.余额," & gOraFmtString.FM_金额 & ") as 余额 ," & _
    "          to_char(a.停用日期,'yyyy-mm-dd hh24:mi:ss') as 停用日期," & _
    "          a.限制类别 " & _
    "   From 消费卡目录 A  " & _
    "   Where A.卡号 = [1] and A.接口编号=[2] And 序号 = (Select Max(序号) From 消费卡目录 B Where 卡号 = A.卡号 and 接口编号=A.接口编号)  " & _
    "   Order by a.序号"
    Err = 0: On Error GoTo ErrHand:
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strCardNo, mlng接口编号)
    If rsTemp.EOF Then
       ShowMsgbox "未找到相关的消费卡记录,请检查!"
        Exit Function
    End If
    '检查:
    '是否回收
    If Nvl(rsTemp!回收时间, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox "卡号为" & strCardNo & "的消费卡已经被" & Nvl(rsTemp!当前状态) & ",不能再刷卡"
        Exit Function
    End If
    '是否停用
    If Nvl(rsTemp!停用日期, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox "卡号为" & strCardNo & "的消费卡已经被停止使用,不能再刷卡"
        Exit Function
    End If
    '是否停用
    If Nvl(rsTemp!停用日期, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox "卡号为" & strCardNo & "的消费卡已经被停止使用,不能再刷卡"
        Exit Function
    End If
    
    '检查效期
    mTyCurCardInfor.dbl余额 = Val(Nvl(rsTemp!余额))
    lbl失效额.Visible = False
    If Nvl(rsTemp!有效期, "3000-01-01 00:00:00") < Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") Then
       '到了有效期
       If Val(Nvl(rsTemp!可否充值)) = 1 Then
          '允许允值的,到期的,不能消费卡面金额,只能消费允值部分
          mTyCurCardInfor.dbl失效面额 = zlGet失效面额(Val(Nvl(rsTemp!ID)), mlng接口编号)
          mTyCurCardInfor.dbl余额 = IIf(mTyCurCardInfor.dbl余额 - mTyCurCardInfor.dbl失效面额 < 0, 0, mTyCurCardInfor.dbl余额 - mTyCurCardInfor.dbl失效面额)
          If mTyCurCardInfor.dbl失效面额 <> 0 Then
            lbl失效额.Caption = "当前卡号失效金额(卡面额)为：" & Format(mTyCurCardInfor.dbl失效面额, gVbFmtString.FM_金额) & "元"
            lbl失效额.Visible = True
            lbl失效额.ForeColor = vbRed
          End If
          
       Else
            '不允许允值的,不能再进行消费
            ShowMsgbox "卡号为" & strCardNo & "的消费卡已经失效,不能再刷卡"
            Exit Function
       End If
    End If
    If mTyCurCardInfor.dbl余额 <= 0 Then
        ShowMsgbox "卡号为" & strCardNo & "的消费卡已经没有余额,不能再刷卡"
        Exit Function
    End If
    If CheckCardNotExists(strCardNo, True) = False Then
    
        Exit Function
    End If
    
    With mTyCurCardInfor
        .lng消费卡ID = Val(Nvl(rsTemp!ID))
        .str卡号 = Nvl(rsTemp!卡号)
        .str限制类别 = Nvl(rsTemp!限制类别)
        .dbl最大消费额 = zl获取最大消费额(.str限制类别, mdbl最大消费额, mdbl已消费累计)
    End With
    txtEdit(mtxtIdx.idx_txt卡号).Text = Nvl(rsTemp!卡号)
    txtEdit(mtxtIdx.idx_txt卡号).Tag = Nvl(rsTemp!卡号)
    lblInfor(mlblIdx.idx_lbl余额).Caption = Format(Val(Nvl(rsTemp!余额)), gVbFmtString.FM_金额)
    lblInfor(mlblIdx.idx_lbl卡类型).Caption = Nvl(rsTemp!卡类型)
    txtEdit(mtxtIdx.idx_txt密码).Tag = Nvl(rsTemp!密码)
    lblInfor(mlblIdx.idx_lbl最大消费).Caption = Format(mTyCurCardInfor.dbl最大消费额, gVbFmtString.FM_金额)
    
    '缺省值:余额不足,缺省余额,否则为最大消费额
    If mTyCurCardInfor.dbl余额 < mTyCurCardInfor.dbl最大消费额 Then
        txtEdit(mtxtIdx.idx_txt本次消费).Text = Format(mTyCurCardInfor.dbl余额, gVbFmtString.FM_金额)
    Else
        txtEdit(mtxtIdx.idx_txt本次消费).Text = Format(mTyCurCardInfor.dbl最大消费额, gVbFmtString.FM_金额)
    End If
    zlBrusCard = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function
Private Function CheckCardNotExists(ByVal strCardNo As String, Optional blnMsgbox As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查卡片信息是否在刷卡中存在
    '     strCardNO-卡号
    '返回:不存在,返回True,否则返回False
    '编制:刘兴洪
    '日期:2009-12-23 17:07:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, varData As Variant
    Err = 0: On Error GoTo ErrHand:
    With vsGrid
        For lngRow = 1 To .Rows - 1
            varData = Split(.Cell(flexcpData, lngRow, .ColIndex("卡号")) & "-", "-")
            If Trim(varData(0)) <> "" Then
                If Trim(Trim(varData(1))) = strCardNo Then
                    If blnMsgbox Then
                        If mblnCardNoSHowPW Then
                            ShowMsgbox "当前卡号已经在第" & lngRow & "中存在了,不能再刷卡!"
                        Else
                            ShowMsgbox "卡号为" & strCardNo & "已经在第" & lngRow & "中存在了,不能再刷卡!"
                        End If
                        .Row = lngRow: .Col = .ColIndex("卡号")
                    End If
                    Exit Function
                End If
            End If
        Next
    End With
    CheckCardNotExists = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume

End Function
Private Function CheckInput() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查输入是否合法
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-12-23 17:03:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    If txtEdit(mtxtIdx.idx_txt卡号).Text <> Trim(txtEdit(mtxtIdx.idx_txt卡号).Tag) Or Trim(txtEdit(mtxtIdx.idx_txt卡号).Text) = "" Then
        ShowMsgbox "未刷卡或刷卡不正确,请检查!"
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt卡号)
        Exit Function
    End If
    If CheckInputPassWord = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt密码)
        Exit Function
    End If
    If CheckInputSquareMoney = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt本次消费)
        Exit Function
    End If
    '检查网格中是否存在相同的卡
    If CheckCardNotExists(Trim(txtEdit(mtxtIdx.idx_txt卡号))) = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt卡号)
        Exit Function
    End If
    CheckInput = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub ClearCtlData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除控件数据
    '编制:刘兴洪
    '日期:2009-12-24 11:11:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    txtEdit(mtxtIdx.idx_txt本次消费) = "0.00"
    txtEdit(mtxtIdx.idx_txt卡号) = ""
    txtEdit(mtxtIdx.idx_txt密码) = ""
    txtEdit(mtxtIdx.idx_txt密码).Tag = ""
    
    lblInfor(mlblIdx.idx_lbl卡类型).Caption = ""
    lblInfor(mlblIdx.idx_lbl余额).Caption = "0.00"
    lblInfor(mlblIdx.idx_lbl最大消费).Caption = "0.00"
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume

End Sub

Private Function zlInsertDataToGrid() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:将刷卡数据，放在网格行中
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-12-23 17:00:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, dbl本次消费总额 As Double
    Err = 0: On Error GoTo ErrHand:
    If CheckInput = False Then Exit Function
    
    With vsGrid
        If .Rows - 1 = 1 Then
            If Trim(.Cell(flexcpData, .Rows - 1, .ColIndex("卡号"))) = "" Then
                lngRow = .Rows - 1
            Else
                .Rows = .Rows + 1
                lngRow = .Rows - 1
            End If
        Else
            If Trim(.Cell(flexcpData, .Rows - 1, .ColIndex("卡号"))) = "" Then
                lngRow = .Rows - 1
            Else
                .Rows = .Rows + 1
                lngRow = .Rows - 1
            End If
        End If
        .TextMatrix(lngRow, .ColIndex("卡号")) = IIf(mblnCardNoSHowPW, "******", mTyCurCardInfor.str卡号)
        .Cell(flexcpData, lngRow, .ColIndex("卡号")) = mTyCurCardInfor.lng消费卡ID & "-" & mTyCurCardInfor.str卡号
        .TextMatrix(lngRow, .ColIndex("卡余额")) = Format(mTyCurCardInfor.dbl余额, gVbFmtString.FM_金额)
        .Cell(flexcpData, lngRow, .ColIndex("卡余额")) = mTyCurCardInfor.dbl余额
        
        .TextMatrix(lngRow, .ColIndex("结算方式")) = mTyCurCardInfor.str结算方式
        .TextMatrix(lngRow, .ColIndex("本次消费")) = Format(Val(txtEdit(mtxtIdx.idx_txt本次消费).Text), gVbFmtString.FM_金额)
        .TextMatrix(lngRow, .ColIndex("备注")) = Trim(txtEdit(mtxtIdx.idx_txt备注).Text)
        .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = lblEdit(0).ForeColor
        grsStatic.dbl已刷累计额 = grsStatic.dbl已刷累计额 + Val(txtEdit(mtxtIdx.idx_txt本次消费).Text)
        dbl本次消费总额 = 0
        For lngRow = 1 To .Rows - 1
            dbl本次消费总额 = dbl本次消费总额 + Val(.TextMatrix(lngRow, .ColIndex("本次消费")))
        Next
        
        If .Rows - 1 >= 1 And dbl本次消费总额 <> 0 Then
            .Rows = .Rows + 1
            lngRow = .Rows - 1
            .TextMatrix(lngRow, .ColIndex("卡号")) = "合计"
            .Cell(flexcpData, lngRow, .ColIndex("卡号")) = ""
            .TextMatrix(lngRow, .ColIndex("卡余额")) = ""
            .Cell(flexcpData, lngRow, .ColIndex("卡余额")) = ""
            .TextMatrix(lngRow, .ColIndex("本次消费")) = Format(dbl本次消费总额, gVbFmtString.FM_金额)
            .TextMatrix(lngRow, .ColIndex("备注")) = ""
            .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbBlue
        End If
        
        mdbl已消费累计 = dbl本次消费总额
        Call ClearCtlData
    End With
    
    Call SetDelRowCtrlEnabled
    zlInsertDataToGrid = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub SetDelRowCtrlEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置删除行的Eanbled属性
    '编制:刘兴洪
    '日期:2009-12-24 10:50:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsGrid
        If .Row < 0 Then cmdDel.Enabled = False: Exit Sub
        cmdDel.Enabled = Trim(.Cell(flexcpData, .Row, .ColIndex("卡号"))) <> ""
    End With
End Sub
Private Sub vsGrid_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsGrid, OldRow, NewRow, OldCol, NewCol, gSysColor.lngGridColorSel
    If OldRow = NewRow Then Exit Sub
    Call SetDelRowCtrlEnabled
End Sub
Private Sub MoveVsGridRowData(Optional lngRow As Long = -1)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:移出行数据
    '入参:lngRow-指定行(-1表示删除当前行)
    '编制:刘兴洪
    '日期:2009-12-24 10:52:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCurRow As Long, dbl本次消费 As Double, blnDeleCurRow As Long
    Dim lng合计Row As Long
    Err = 0: On Error GoTo ErrHand:
    With vsGrid
        If lngRow < 0 Then lngRow = .Row
        If lngRow < 0 Then Exit Sub
        blnDeleCurRow = lngRow = .Row
        lngCurRow = .Row
        
        dbl本次消费 = Val(.TextMatrix(lngRow, .ColIndex("本次消费")))
        grsStatic.dbl已刷累计额 = IIf(grsStatic.dbl已刷累计额 - dbl本次消费 < 0, 0, grsStatic.dbl已刷累计额 - dbl本次消费)
        
        If .Rows - 1 <= 1 Then
            .Cell(flexcpText, 1, 0, 1, .Cols - 1) = ""
            .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
            .Row = 1
        Else
            If .Cell(flexcpData, lngRow, .ColIndex("卡号")) = "" Then Exit Sub
            .RemoveItem lngRow
            If blnDeleCurRow Then
                If lngCurRow >= .Rows - 1 Then
                    .Row = .Rows - 1
                Else
                    .Row = lngCurRow + 1
                End If
            End If
        End If
        '重新计算合计数
        dbl本次消费 = 0
        For lngCurRow = 1 To .Rows - 1
            If .Cell(flexcpData, lngCurRow, .ColIndex("卡号")) <> "" Then
                dbl本次消费 = dbl本次消费 + Val(.TextMatrix(lngCurRow, .ColIndex("本次消费")))
            End If
            If .TextMatrix(lngCurRow, .ColIndex("卡号")) = "合计" Then lng合计Row = lngCurRow
        Next
        If dbl本次消费 = 0 And .Rows - 1 <= 2 Then
            .Clear 1: .Rows = 2: .Row = 1
            .Cell(flexcpForeColor, 1, 0, 1, .Cols - 1) = &H80000008
        Else
            '增加合计
            If lngCurRow < 1 Then
                .Rows = .Rows + 1
                lng合计Row = .Rows - 1
            End If
            .TextMatrix(lng合计Row, .ColIndex("卡号")) = "合计"
            .Cell(flexcpData, lng合计Row, .ColIndex("卡号")) = ""
            .TextMatrix(lng合计Row, .ColIndex("卡余额")) = ""
            .Cell(flexcpData, lng合计Row, .ColIndex("卡余额")) = ""
            .TextMatrix(lng合计Row, .ColIndex("本次消费")) = Format(dbl本次消费, gVbFmtString.FM_金额)
            .TextMatrix(lng合计Row, .ColIndex("备注")) = ""
            .Cell(flexcpForeColor, lng合计Row, 0, lng合计Row, .Cols - 1) = vbBlue
        End If
        mdbl已消费累计 = dbl本次消费
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub vsGrid_GotFocus()
  zl_VsGridGotFocus vsGrid, gSysColor.lngGridColorSel
End Sub

Private Sub vsGrid_LostFocus()
  zl_VsGridLOSTFOCUS vsGrid, gSysColor.lngGridColorLost
End Sub

Private Function CreateObjectKeyboard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建密码创建
    '返回:创建成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-24 23:59:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    Set mobjKeyboard = CreateObject("zl9Keyboard.clsKeyboard")
    If Err <> 0 Then Exit Function
    Err = 0
    CreateObjectKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function OpenPassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否者False
    '编制:刘兴洪
    '日期:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText) = False Then Exit Function
    OpenPassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function ClosePassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否者False
    '编制:刘兴洪
    '日期:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.ColsePassKeyoardInput(Me, ctlText) = False Then Exit Function
    ClosePassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

