VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSelect 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   ClientHeight    =   3570
   ClientLeft      =   2715
   ClientTop       =   3420
   ClientWidth     =   6855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin zl9NewQuery.ctlButton ctlClose 
      Height          =   540
      Left            =   5520
      TabIndex        =   4
      Top             =   1575
      Width           =   1200
      _extentx        =   2117
      _extenty        =   953
      caption         =   "取消"
      backcolor       =   16777215
      autosize        =   0   'False
      buttonheight    =   420
      textaligment    =   0
   End
   Begin zl9NewQuery.ctlButton ctlOK 
      Height          =   540
      Left            =   5520
      TabIndex        =   3
      Top             =   915
      Width           =   1200
      _extentx        =   2117
      _extenty        =   953
      caption         =   "确定"
      backcolor       =   16777215
      autosize        =   0   'False
      buttonheight    =   420
      textaligment    =   0
   End
   Begin zl9NewQuery.ctlButton UsrCmd 
      Height          =   540
      Index           =   0
      Left            =   4215
      TabIndex        =   1
      Top             =   2880
      Width           =   1245
      _extentx        =   2196
      _extenty        =   953
      caption         =   "上移"
      backcolor       =   16777215
      forecolor       =   8421504
      fontsize        =   10.5
      autosize        =   0   'False
      buttonheight    =   420
      textaligment    =   0
   End
   Begin zl9NewQuery.ctlButton UsrCmd 
      Height          =   540
      Index           =   1
      Left            =   2850
      TabIndex        =   2
      Top             =   2880
      Width           =   1260
      _extentx        =   2223
      _extenty        =   953
      caption         =   "下移"
      backcolor       =   16777215
      forecolor       =   8421504
      fontsize        =   10.5
      autosize        =   0   'False
      buttonheight    =   420
      textaligment    =   0
   End
   Begin MSComctlLib.ImageList ilsImage 
      Left            =   5895
      Top             =   2745
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelect.frx":0000
            Key             =   "down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelect.frx":039A
            Key             =   "up"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid msfAsk 
      Height          =   1815
      Left            =   615
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   930
      Width           =   4800
      _cx             =   8467
      _cy             =   3201
      Appearance      =   0
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
      BackColorFixed  =   15199202
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16633516
      ForeColorSel    =   16711680
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   16761024
      GridColorFixed  =   16761024
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   450
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
      WordWrap        =   -1  'True
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
   Begin VB.Image Image1 
      Height          =   480
      Left            =   75
      Picture         =   "frmSelect.frx":0734
      Top             =   165
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "请您选择要查询哪一次的情况,若是住院,选择住院次数."
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   645
      Left            =   630
      TabIndex        =   0
      Top             =   195
      Width           =   4845
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mvarFirst As Boolean
Private mvarRs As New ADODB.Recordset
Private mvarOK As Boolean

Private mvarCurPos1 As Long
Private mvarRows1 As Long

Private mvar病人id As Long
Private mvar主页id As Long

Private Sub cmdClose_Click()
    If Val(msfAsk.TextMatrix(msfAsk.Row, 0)) > 0 Then
        mvar病人id = Val(msfAsk.TextMatrix(msfAsk.Row, 3))
        mvar主页id = Val(msfAsk.TextMatrix(msfAsk.Row, 4))
    End If
    mvarOK = True
    Unload Me
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call InitInternal
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call InitInternal
End Sub

Private Sub ctlClose_CommandClick()
    Unload Me
End Sub

Private Sub ctlOK_CommandClick()
    If Val(msfAsk.TextMatrix(msfAsk.Row, 0)) > 0 Then
        mvar病人id = Val(msfAsk.TextMatrix(msfAsk.Row, 3))
        mvar主页id = Val(msfAsk.TextMatrix(msfAsk.Row, 4))
    End If
    mvarOK = True
    Unload Me
End Sub

Private Sub Form_Activate()
    If mvarFirst = False Then Exit Sub
    mvarFirst = False
    Dim i As Long
    Dim k As Long
    
    i = 1
    

        While Not mvarRs.EOF
            msfAsk.TextMatrix(i, 0) = mvarRs!No
            msfAsk.TextMatrix(i, 1) = IIf(IsNull(mvarRs!入院日期), "", Format(mvarRs!入院日期, "YYYY-MM-DD HH:MM"))
            msfAsk.TextMatrix(i, 2) = IIf(IsNull(mvarRs!出院日期), "", Format(mvarRs!出院日期, "YYYY-MM-DD HH:MM"))
            msfAsk.TextMatrix(i, 3) = IIf(IsNull(mvarRs!病人id), "", mvarRs!病人id)
            msfAsk.TextMatrix(i, 4) = IIf(IsNull(mvarRs!主页id), "", mvarRs!主页id)
            
            msfAsk.MergeRow(i) = False
            
            If msfAsk.TextMatrix(i, 1) = "门诊费用" Then
                msfAsk.MergeCells = flexMergeFree
                msfAsk.MergeRow(i) = True
            End If
            i = i + 1
            msfAsk.Rows = i + 1
            msfAsk.Row = 1
            mvarRs.MoveNext
        Wend

    
    If msfAsk.Rows > 2 Then msfAsk.Rows = msfAsk.Rows - 1
    mvarRows1 = msfAsk.Rows - 1
    msfAsk.Rows = msfAsk.Rows + 10
    
    Call EnablePageButton(msfAsk, mvarCurPos1, mvarRows1, UsrCmd(0), UsrCmd(1))
    
    
'    On Error Resume Next
'
'    pic.Cls
'    pic.Width = msfAsk.Width
'    pic.Height = msfAsk.Height
'    pic.PaintPicture Me.Image, 0, 0, pic.Width, pic.Height, msfAsk.Left, msfAsk.Top, msfAsk.Width, msfAsk.Height
'
'    Set msfAsk.WallPaper = pic.Image
    


End Sub

Private Sub Form_Load()
    mvarFirst = True
    mvarOK = False
    
    msfAsk.Rows = 2
    msfAsk.Cols = 0
            
    UsrCmd(0).Picture = ilsImage.ListImages("up")
    UsrCmd(1).Picture = ilsImage.ListImages("down")
    
    Call ClearSpecRowCol(msfAsk, 1, Array())
    Call AddColumn(msfAsk, "次数", 600, 1)
    Call AddColumn(msfAsk, "入院时间", 2100, 1)
    Call AddColumn(msfAsk, "出院时间", 2100, 1)
    Call AddColumn(msfAsk, "病人id", 0, 1)
    Call AddColumn(msfAsk, "主页id", 0, 1)
    msfAsk.Rows = 10
    
    mvar病人id = 0
    mvar主页id = 0
End Sub

Public Function ShowSelect(vRs As ADODB.Recordset, v病人id As Long, v主页id As Long) As Boolean
    If vRs.RecordCount = 1 Then
        v病人id = IIf(IsNull(vRs!病人id), 0, vRs!病人id)
        v主页id = IIf(IsNull(vRs!主页id), 0, vRs!主页id)
        ShowSelect = True
        Exit Function
    End If
    Set mvarRs = vRs
    frmSelect.Show 1
    If mvar病人id > 0 Then
        v病人id = mvar病人id
        v主页id = mvar主页id
    End If
    ShowSelect = mvarOK
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call InitInternal
End Sub

Private Sub Form_Paint()
    Call DrawColorToColor(Me, Me.BackColor, &HFFC0C0, , True)
End Sub

Private Sub msfAsk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call InitInternal
End Sub

Private Sub UsrCmd_CommandClick(Index As Integer)
    Call TurnToPage(msfAsk, IIf(Index = 0, -1, 1), mvarCurPos1)
    Call EnablePageButton(msfAsk, mvarCurPos1, mvarRows1, UsrCmd(0), UsrCmd(1))
End Sub

Private Sub UsrCmd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call InitInternal
End Sub
