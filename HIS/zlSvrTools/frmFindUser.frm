VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFindUser 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "用户过滤"
   ClientHeight    =   5595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4845
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ImageList img16 
      Left            =   1920
      Top             =   4800
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
            Picture         =   "frmFindUser.frx":0000
            Key             =   "unCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFindUser.frx":059A
            Key             =   "Check"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pctBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5595
      Left            =   0
      ScaleHeight     =   5565
      ScaleWidth      =   4815
      TabIndex        =   0
      Top             =   0
      Width           =   4845
      Begin VB.TextBox txtStop 
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   5160
         Width           =   90
      End
      Begin VB.CommandButton cmdYes 
         Caption         =   "确定(&S)"
         Height          =   300
         Left            =   2640
         TabIndex        =   7
         Top             =   5160
         Width           =   975
      End
      Begin VB.CommandButton cmdNo 
         Caption         =   "取消(&C)"
         Height          =   300
         Left            =   3720
         TabIndex        =   6
         Top             =   5160
         Width           =   975
      End
      Begin VB.Frame fraClose 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   -85
         Width           =   4825
         Begin VB.PictureBox PicClose 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   200
            Left            =   4560
            Picture         =   "frmFindUser.frx":0B34
            ScaleHeight     =   195
            ScaleWidth      =   210
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   120
            Width           =   215
         End
         Begin VB.Label LblHead 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "用户查找"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   4
            Top             =   120
            Width           =   720
         End
      End
      Begin VB.TextBox txtFind 
         ForeColor       =   &H80000010&
         Height          =   350
         Left            =   600
         TabIndex        =   2
         Text            =   "输入用户名、姓名或姓名简码后按回车进行定位"
         Top             =   395
         Width           =   4095
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfUser 
         Height          =   4215
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   4575
         _cx             =   8070
         _cy             =   7435
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
         BackColorBkg    =   16777215
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
         AllowUserResizing=   1
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
         ExplorerBar     =   1
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
         ComboSearch     =   0
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
      Begin VB.Label lblFind 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "查找"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmFindUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrResult As String    '返回的结果
Private Enum Color
    tipColor = &H80000010
    txtColor = &H80000012
    SelColor = &H8000000D
End Enum

Private Const conCol = "选择,250,1;用户名,1200,1;姓名,1200,1;部门,500,1"

Public Function ShowMe(ByVal frmMain As Object, ByVal rsUsers As ADODB.Recordset, strUsers As String, ByVal sngLeft As Single, ByVal sngTop As Single) As String
'功能:显示窗体,返回选取的用户名

    Left = sngLeft
    Top = sngTop
    
    mstrResult = strUsers
    Call LoadUserData(rsUsers, mstrResult)
    Me.Show 1
    ShowMe = mstrResult
End Function

Private Sub cmdNo_Click()
    Unload Me
End Sub

Private Sub cmdYes_Click()
    mstrResult = GetUsernames
    Unload Me
End Sub


Private Sub Form_Load()
    Call InitTable(vsfUser, conCol)
End Sub

Private Sub PicClose_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then RaisEffect PicClose, -2
End Sub

Private Sub PicClose_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then RaisEffect PicClose, 2
    
    If x > 0 And x < PicClose.Width And y > 0 And y < PicClose.Height Then Unload Me
End Sub

Private Sub LoadUserData(rsData As ADODB.Recordset, ByVal strCheck As String)
    '功能:获取账户数据,将包含的数据打勾
     Dim i As Integer
     
    On Error GoTo errh
    
    With vsfUser
        .Cell(flexcpPicture, 0, 0) = img16.ListImages("unCheck").Picture
        .Cell(flexcpText, 0, 0) = ""
        .Cell(flexcpPictureAlignment, 0, 0) = flexPicAlignCenterCenter
        If rsData.RecordCount = 0 Then
            .Rows = .FixedRows
            Exit Sub
        End If
        .Redraw = flexRDNone
        .Editable = flexEDKbdMouse
        rsData.MoveFirst
        .Rows = .FixedRows
        .Rows = rsData.RecordCount + .FixedRows
        
        .ColSort(-1) = flexSortCustom
        .ColSort(0) = flexSortNone
        .ColDataType(0) = flexDTBoolean
        
        i = .FixedRows
        Do While Not rsData.EOF
            If InstrEx(strCheck, rsData.Fields(0)) Then
                .TextMatrix(i, .ColIndex("选择")) = "-1"
            Else
                .TextMatrix(i, .ColIndex("选择")) = 0
            End If
            .TextMatrix(i, .ColIndex("用户名")) = rsData.Fields(0)
            .TextMatrix(i, .ColIndex("姓名")) = rsData.Fields(1) & ""
            .TextMatrix(i, .ColIndex("部门")) = rsData.Fields(2) & ""
            .RowData(i) = rsData.Fields(3) & ""
            i = i + 1: rsData.MoveNext
        Loop
        
        .Redraw = flexRDDirect
        vsfUser_AfterEdit 0, 0
        If .Rows > .FixedRows Then .Select .FixedRows, 0
    End With
    Exit Sub
errh:
    MsgBox err.Description
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then    '按下回车键
        Call GetRowPos(vsfUser, txtFind.Text, "用户名,姓名,部门")
    End If
End Sub

Private Sub txtFind_GotFocus()
    If txtFind.Text = "输入用户名、姓名或姓名简码后按回车进行定位" Then
        txtFind.Text = ""
        txtFind.ForeColor = txtColor
    End If
End Sub

Private Sub txtFind_LostFocus()
    If txtFind.Text = "" Then
        txtFind.Text = "输入用户名、姓名或姓名简码后按回车进行定位"
        txtFind.ForeColor = tipColor
    End If
End Sub

Private Sub vsfUser_AfterSort(ByVal Col As Long, Order As Integer)
    Dim i As Integer
    
    With vsfUser
        If Col = 0 Then
            If .Cell(flexcpPicture, 0, 0) = img16.ListImages("unCheck").Picture Then
                .Cell(flexcpPicture, 0, 0) = img16.ListImages("Check").Picture
                For i = .FixedRows To .Rows - .FixedRows
                    .TextMatrix(i, 0) = "-1"
                Next
            Else
                .Cell(flexcpPicture, 0, 0) = img16.ListImages("unCheck").Picture
                For i = .FixedRows To .Rows - .FixedRows
                    .TextMatrix(i, 0) = "0"
                Next
            End If
        End If
    End With
End Sub

Private Sub vsfUser_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Integer, blnAllSelectd As Boolean
    
    blnAllSelectd = True
    With vsfUser
        If .Redraw = flexRDNone Then Exit Sub
        
        For i = .FixedRows To .Rows - .FixedRows
            If .TextMatrix(i, 0) = "0" Then
                blnAllSelectd = False
                Exit For
            End If
        Next

        
        If blnAllSelectd Then
            .Cell(flexcpPicture, 0, 0) = img16.ListImages("Check").Picture
        Else
            .Cell(flexcpPicture, 0, 0) = img16.ListImages("unCheck").Picture
        End If
    End With
End Sub

Private Sub vsfUser_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub

Private Sub vsfUser_DblClick()
    Call vsfUser_KeyUp(32, 1)
End Sub

Private Sub vsfUser_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    
    If KeyCode = 13 Then
        cmdYes_Click
    End If
    
    With vsfUser
        If KeyCode = 32 And .Col <> 0 Then   '按下空格,进行勾选
            For i = .FixedRows To .Rows - .FixedRows
                If .IsSelected(i) Then
                    .TextMatrix(i, 0) = IIf(.TextMatrix(i, 0) = "-1", 0, -1)
                End If
            Next
        End If
    End With
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then cmdNo_Click
End Sub


Private Function GetUsernames() As String
    '功能:遍历表格,将表格中选中行的用户名返回.
    Dim strTmp As String, intRow As Integer
    
    With vsfUser
        For intRow = 0 To .Rows - 1
            If .TextMatrix(intRow, 0) = "-1" Then
                If strTmp = "" Then
                    strTmp = .TextMatrix(intRow, 1)
                Else
                    strTmp = strTmp & "," & .TextMatrix(intRow, 1)
                End If
            End If
        Next
    End With
    GetUsernames = strTmp
End Function
