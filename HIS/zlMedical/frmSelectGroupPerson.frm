VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSelectGroupPerson 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "单位人员选择"
   ClientHeight    =   5955
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8805
   Icon            =   "frmSelectGroupPerson.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd 
      Caption         =   "全清(&D)"
      Height          =   350
      Index           =   2
      Left            =   7605
      TabIndex        =   8
      Top             =   45
      Width           =   1100
   End
   Begin VB.CommandButton cmd 
      Caption         =   "全选(&A)"
      Height          =   350
      Index           =   1
      Left            =   6480
      TabIndex        =   7
      Top             =   45
      Width           =   1100
   End
   Begin VB.CommandButton cmdFilter 
      Caption         =   "过滤(&F)"
      Height          =   350
      Left            =   3795
      TabIndex        =   6
      Top             =   45
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7605
      TabIndex        =   11
      Top             =   5505
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6390
      TabIndex        =   10
      Top             =   5505
      Width           =   1100
   End
   Begin VB.TextBox txt 
      Height          =   300
      Index           =   1
      Left            =   3180
      TabIndex        =   5
      Text            =   "100"
      Top             =   60
      Width           =   510
   End
   Begin VB.TextBox txt 
      Height          =   300
      Index           =   0
      Left            =   2295
      TabIndex        =   3
      Text            =   "0"
      Top             =   60
      Width           =   510
   End
   Begin VB.ComboBox cbo 
      Height          =   300
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   60
      Width           =   1365
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   4890
      Left            =   0
      TabIndex        =   9
      Top             =   450
      Width           =   8700
      _cx             =   15346
      _cy             =   8625
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
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
      Begin VB.Line lnX 
         Index           =   0
         Visible         =   0   'False
         X1              =   -555
         X2              =   1230
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Line lnY 
         Index           =   0
         Visible         =   0   'False
         X1              =   270
         X2              =   270
         Y1              =   420
         Y2              =   1635
      End
   End
   Begin MSComctlLib.ImageList ils13 
      Left            =   3195
      Top             =   5475
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectGroupPerson.frx":000C
            Key             =   "状态"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectGroupPerson.frx":03A6
            Key             =   "个人"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectGroupPerson.frx":0940
            Key             =   "团体"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectGroupPerson.frx":59AA
            Key             =   "确认"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectGroupPerson.frx":5F44
            Key             =   "取消"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectGroupPerson.frx":64DE
            Key             =   "开始"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectGroupPerson.frx":6A78
            Key             =   "新开"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectGroupPerson.frx":7012
            Key             =   "完成"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectGroupPerson.frx":75AC
            Key             =   "up"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectGroupPerson.frx":776E
            Key             =   "down"
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   5550
      Width           =   90
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "到"
      Height          =   180
      Index           =   2
      Left            =   2910
      TabIndex        =   4
      Top             =   120
      Width           =   180
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "年龄"
      Height          =   180
      Index           =   1
      Left            =   1875
      TabIndex        =   2
      Top             =   120
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "性别"
      Height          =   180
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "frmSelectGroupPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Enum mCol
    选择 = 0
End Enum

Private mfrmMain As Object
Private mblnOK As Boolean
Private mlngKey As Long
Private mrsData As ADODB.Recordset
Private mrsSelData As ADODB.Recordset

Public Function ShowFilter(ByVal frmMain As Object, ByVal lngKey As Long, ByRef rsSelData As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '
    '
    '
    '******************************************************************************************************************
    Dim strVsf As String
    
    mblnOK = False
    
    If lngKey = 0 Then Exit Function

    gstrSQL = GetPublicSQL(SQL.单位人员选择)
    Set mrsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, "", 0, 1000)
    If mrsData.BOF Then
        ShowSimpleMsg "当前单位无对应人人员。"
        Exit Function
    End If
    
    mlngKey = lngKey
    Set mfrmMain = frmMain
    
    cbo.Clear
    cbo.AddItem ""
    cbo.AddItem "男"
    cbo.AddItem "女"
    
    strVsf = "选择,255,4,1,1,;姓名,810,1,1,1,;性别,750,1,1,1,;婚姻状况,990,1,1,1,;健康号,900,1,1,1,;门诊号,1200,1,1,1,;年龄,900,7,1,1,"
    Call CreateVsf(vsf, strVsf)
    vsf.Cols = vsf.Cols + 1
    vsf.ColWidth(vsf.Cols - 1) = 15
    vsf.ColDataType(mCol.选择) = flexDTBoolean
    Set vsf.Cell(flexcpPicture, 0, 0) = ils13.ListImages("状态").Picture
    vsf.Editable = True
        
    If mrsData.BOF = False Then
        Call LoadGrid(vsf, mrsData)
        lbl(3).Caption = "总人数：" & mrsData.RecordCount & " 人"
    Else
        lbl(3).Caption = ""
    End If
    
    Call AppendRows(vsf, lnX, lnY)
    
    Me.Show 1, mfrmMain
    
    Set rsSelData = mrsSelData
    ShowFilter = mblnOK

End Function

Private Function CreateRec(ByVal rsFrom As ADODB.Recordset) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim lngCol As Long
    
    For lngCol = 0 To rsFrom.Fields.Count - 1
        
        Select Case rsFrom.Fields(lngCol).Type
        Case 131, 139

            rs.Fields.Append rsFrom.Fields(lngCol).Name, adBigInt, , adFldIsNullable
            
        Case Else

            rs.Fields.Append rsFrom.Fields(lngCol).Name, adVarChar, rsFrom.Fields(lngCol).DefinedSize, adFldIsNullable
            
        End Select
        
    Next
    
    rs.CursorLocation = adUseClient
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenStatic
    rs.Open
    
    Set CreateRec = rs
End Function

Private Sub cbo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Select Case Index
    Case 1
        vsf.Cell(flexcpText, 1, mCol.选择, vsf.Rows - 1, mCol.选择) = 1
    Case 2
        vsf.Cell(flexcpText, 1, mCol.选择, vsf.Rows - 1, mCol.选择) = 0
    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFilter_Click()
    gstrSQL = GetPublicSQL(SQL.单位人员选择)
    Set mrsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey, cbo.Text, Val(txt(0).Text), Val(txt(1).Text))
    
    Call ResetVsf(vsf)
    If mrsData.BOF = False Then
        Call LoadGrid(vsf, mrsData)
        lbl(3).Caption = "总人数：" & mrsData.RecordCount & " 人"
    Else
        lbl(3).Caption = ""
    End If
    
    Call AppendRows(vsf, lnX, lnY)
    
End Sub

Private Sub cmdOK_Click()
    Dim lngLoop As Long
    Dim lngCol As Long
    
    Set mrsSelData = CreateRec(mrsData)

    For lngLoop = 1 To vsf.Rows - 1
        
        If Abs(Val(vsf.TextMatrix(lngLoop, 0))) = 1 Then
            mrsData.Filter = ""
            mrsData.Filter = "ID=" & Val(vsf.RowData(lngLoop))
            
            If mrsData.RecordCount > 0 Then
                mrsSelData.AddNew
                For lngCol = 0 To mrsData.Fields.Count - 1
                    mrsSelData.Fields(lngCol).Value = mrsData.Fields(lngCol).Value
                    mrsSelData("选择").Value = 1
                Next
            End If
        End If
    Next
        
    mblnOK = True
    Unload Me
End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txt(Index)
        
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        If FilterKeyAscii(KeyAscii, 99, "0123456789") = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngLoop As Long
    
    If Abs(Val(vsf.TextMatrix(Row, mCol.选择))) = 1 Then
        Exit Sub
    End If
        
    For lngLoop = 1 To vsf.Rows - 1
        If Abs(Val(vsf.TextMatrix(lngLoop, mCol.选择))) = 1 Then
            Exit Sub
        End If
    Next
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col = 0)
End Sub

Private Sub vsf_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeySpace And vsf.Col <> mCol.选择 Then
    
        If Abs(Val(vsf.TextMatrix(vsf.Row, mCol.选择))) = 1 Then
            vsf.TextMatrix(vsf.Row, mCol.选择) = 0
        Else
            vsf.TextMatrix(vsf.Row, mCol.选择) = 1
        End If

    End If
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> mCol.选择 Or Val(vsf.RowData(Row)) <= 0 Then
        Cancel = True
    End If
End Sub






