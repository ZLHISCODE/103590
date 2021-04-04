VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmClinicShortCut 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   6150
   ClientLeft      =   -30
   ClientTop       =   -315
   ClientWidth     =   4590
   ControlBox      =   0   'False
   Icon            =   "frmClinicShortCut.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmClinicShortCut"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picTitle 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   450
      ScaleHeight     =   240
      ScaleWidth      =   3825
      TabIndex        =   1
      Top             =   390
      Width           =   3825
      Begin VB.Label lblClose 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "×"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3585
         TabIndex        =   4
         Top             =   30
         Width           =   210
      End
      Begin VB.Label lblMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "▼"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   3315
         TabIndex        =   3
         Top             =   30
         Width           =   180
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "快捷"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   75
         TabIndex        =   2
         Top             =   30
         Width           =   390
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsList 
      Height          =   3885
      Left            =   480
      TabIndex        =   0
      Top             =   870
      Width           =   1815
      _cx             =   3201
      _cy             =   6853
      Appearance      =   0
      BorderStyle     =   0
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
      MousePointer    =   54
      BackColor       =   15659506
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   13811126
      ForeColorSel    =   0
      BackColorBkg    =   15659506
      BackColorAlternate=   15659506
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   15659506
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   15
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmClinicShortCut.frx":000C
      ScrollTrack     =   -1  'True
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
      Ellipsis        =   1
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
   Begin VB.Shape Bdr 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   5775
      Left            =   120
      Top             =   180
      Width           =   4260
   End
   Begin VB.Menu mnuType 
      Caption         =   "分类"
      Visible         =   0   'False
      Begin VB.Menu mnuTypeItem 
         Caption         =   "西药(&1) ALT+1"
         Index           =   0
      End
      Begin VB.Menu mnuTypeItem 
         Caption         =   "成药(&2) ALT+2"
         Index           =   1
      End
      Begin VB.Menu mnuTypeItem 
         Caption         =   "中药(&3) ALT+3"
         Index           =   2
      End
      Begin VB.Menu mnuTypeItem 
         Caption         =   "配方(&4) ALT+4"
         Index           =   3
      End
      Begin VB.Menu mnuTypeItem 
         Caption         =   "诊疗(&5) ALT+5"
         Index           =   4
      End
      Begin VB.Menu mnuTypeItem 
         Caption         =   "成套(&6) ALT+6"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmClinicShortCut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event ItemClick(ByVal 类型 As Integer, ByVal 分类ID As Long)

Private mrslist As ADODB.Recordset
Private mlngPreRow As Long
Private mintType As Integer
Private mfrmParent As Object
Private mblnShow As Boolean

Public Sub ShowMe(frmParent As Object, Optional ByVal blnBySave As Boolean)
'参数：blnBySave=是否根据上次面板显示与否进行显示
    Dim blnShow As Boolean
    
    Set mfrmParent = frmParent
    
    If blnBySave Then
        blnShow = CBool(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & mfrmParent.Name, "ShortCutState", False))
    Else
        blnShow = Not mblnShow
    End If
    
    If blnShow Then
        mblnShow = True
        Me.Show , frmParent
        Call ClearListSel
    Else
        If Not mrslist Is Nothing Then Me.Hide '加载了才隐藏
        mblnShow = False
    End If
    
    If mfrmParent.Visible Then
        mfrmParent.SetFocus
    End If
End Sub

Public Sub ShowShortCut(ByVal intType As Integer)
'功能：切换菜单项显示
'参数：intType=1-6,对应相应顺序的菜单
    If mblnShow Then
        If intType - 1 >= mnuTypeItem.LBound And intType - 1 <= mnuTypeItem.UBound Then
            Call mnuTypeItem_Click(intType - 1)
        End If
    End If
End Sub

Public Sub SaveShowState()
'功能：保存面板显示与否
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & mfrmParent.Name, "ShortCutState", mblnShow
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("`") Then
        KeyAscii = 0
        Me.Hide
        mblnShow = False
        mfrmParent.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim strPos As String
    
    Call FormSetCaption(Me, False, False)

    On Error GoTo errH
    
    strSQL = "Select ID,类型,编码,名称 From 诊疗分类目录 Where 类型 IN(1,2,3,4,5,6) And 上级ID Is Null Order by 类型,编码"
    Set mrslist = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(mrslist, strSQL, Me.Caption)
            
    Call mnuTypeItem_Click(Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & mfrmParent.Name, "ShortCut", 0)))
    
    strPos = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & mfrmParent.Name, "ShortCutPostion", "375,-100")
    Me.Top = mfrmParent.Top + Val(Split(strPos, ",")(0))
    Me.Left = mfrmParent.Left + mfrmParent.Width + Val(Split(strPos, ",")(1)) - Me.Width
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FillList(ByVal intType As Integer)
    Dim i As Long
    
    mlngPreRow = -1
    mintType = intType
    mrslist.Filter = "类型=" & intType
    
    With vsList
        .Redraw = flexRDNone
        
        .Rows = 0
        .Cols = 1
        .Rows = mrslist.RecordCount
        If .Rows = 0 Then .Rows = 1
        
        .Height = .RowHeightMin * .Rows
        
        For i = 1 To mrslist.RecordCount
            .TextMatrix(i - 1, 0) = " " & mrslist!名称
            .Cell(flexcpData, i - 1, 0) = CLng(mrslist!ID)
            mrslist.MoveNext
        Next
        
        .AutoSize 0
        If .ColWidth(0) < 900 Then .ColWidth(0) = 900
        If .ColWidth(0) > 1600 Then .ColWidth(0) = 1600
        .Width = .ColWidth(0)
        
        .Redraw = flexRDDirect
           
        Call SetFormSize
    End With
End Sub

Private Sub SetFormSize()
    Me.Width = vsList.Width + (Bdr.BorderWidth * 15 + 30) * 2
    Me.Height = vsList.Height + picTitle.Height + 15 + (Bdr.BorderWidth * 15 + 30) * 2
    
    Bdr.Left = 15
    Bdr.Top = 15
    Bdr.Width = Me.Width - 15
    Bdr.Height = Me.Height - 15
    
    picTitle.Left = Bdr.Left + Bdr.BorderWidth * 15 + 15
    picTitle.Top = Bdr.Top + Bdr.BorderWidth * 15 + 15
    picTitle.Width = Me.Width - picTitle.Left * 2
    
    vsList.Left = picTitle.Left
    vsList.Top = picTitle.Top + picTitle.Height + 30
    
    Call SetCloseButton(0, True)
    Call SetMenuButton(0, True)
End Sub

Private Sub SetCloseButton(ByVal intState As Integer, Optional ByVal blnSize As Boolean)
'参数：intState=0-正常,1-弹起,2-按下
    If intState = 0 Then
        lblClose.BackColor = picTitle.BackColor
        lblClose.ForeColor = vbWhite
        lblClose.BorderStyle = 0
    ElseIf intState = 1 Then
        lblClose.BackColor = vsList.BackColorSel
        lblClose.ForeColor = vbBlack
        lblClose.BorderStyle = 1
    ElseIf intState = 2 Then
        lblClose.BackColor = 11899525
        lblClose.ForeColor = vbWhite
        lblClose.BorderStyle = 1
    End If
    
    If blnSize Then
        lblClose.Width = 210
        lblClose.Height = 195
        lblClose.Left = picTitle.Width - lblClose.Width - 15
        lblClose.Top = (picTitle.Height - lblClose.Height) / 2
    End If
End Sub

Private Sub SetMenuButton(ByVal intState As Integer, Optional ByVal blnSize As Boolean)
'参数：intState=0-正常,1-弹起,2-按下
    If intState = 0 Then
        lblMenu.BackColor = picTitle.BackColor
        lblMenu.ForeColor = vbWhite
        lblMenu.BorderStyle = 0
    ElseIf intState = 1 Then
        lblMenu.BackColor = vsList.BackColorSel
        lblMenu.ForeColor = vbBlack
        lblMenu.BorderStyle = 1
    ElseIf intState = 2 Then
        lblMenu.BackColor = 11899525
        lblMenu.ForeColor = vbWhite
        lblMenu.BorderStyle = 1
    End If
    
    If blnSize Then
        lblMenu.Width = 210
        lblMenu.Height = 195
        lblMenu.Left = picTitle.Width - lblMenu.Width - lblClose.Width - 30
        lblMenu.Top = (picTitle.Height - lblMenu.Height) / 2
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Call MoveObj(Me.Hwnd)
    End If
    mfrmParent.SetFocus
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call SetCloseButton(0)
    Call SetMenuButton(0)
    Call ClearListSel
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim lngTop As Long, lngRight As Long
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & mfrmParent.Name, "ShortCut", mintType - 1
    
    '保存相对于主窗体右上角的位置
    If mfrmParent.WindowState = 0 Then
        lngTop = Me.Top - mfrmParent.Top
        lngRight = Me.Left + Me.Width - (mfrmParent.Left + mfrmParent.Width)
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & mfrmParent.Name, "ShortCutPostion", lngTop & "," & lngRight
    End If
    
    mblnShow = False
    Set mrslist = Nothing
End Sub

Private Sub lblClose_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Call SetCloseButton(2)
    End If
    mfrmParent.SetFocus
End Sub

Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call SetMenuButton(0)
    
    If x >= 0 And y >= 0 And x <= lblClose.Width And y <= lblClose.Height Then
        If Button = 1 Then
            Call SetCloseButton(2)
        Else
            Call SetCloseButton(1)
        End If
    Else
        Call SetCloseButton(1)
    End If
End Sub

Private Sub lblClose_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If x >= 0 And y >= 0 And x <= lblClose.Width And y <= lblClose.Height Then
        Me.Hide
        mblnShow = False
        Call SetCloseButton(0)
        mfrmParent.SetFocus
    End If
End Sub

Private Sub lblMenu_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Call SetMenuButton(2)
    End If
    mfrmParent.SetFocus
End Sub

Private Sub lblMenu_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call SetCloseButton(0)
    
    If x >= 0 And y >= 0 And x <= lblMenu.Width And y <= lblMenu.Height Then
        If Button = 1 Then
            Call SetMenuButton(2)
        Else
            Call SetMenuButton(1)
        End If
    Else
        Call SetMenuButton(1)
    End If
End Sub

Private Sub lblMenu_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If x >= 0 And y >= 0 And x <= lblMenu.Width And y <= lblMenu.Height Then
        Call SetCloseButton(0)
        PopupMenu mnuType, 8, picTitle.Left + lblMenu.Left + lblMenu.Width, picTitle.Top + lblMenu.Top + lblMenu.Height
    Else
        Call SetMenuButton(0)
    End If
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Call MoveObj(Me.Hwnd)
    End If
    mfrmParent.SetFocus
End Sub

Private Sub mnuTypeItem_Click(Index As Integer)
    Dim lngRight As Long, i As Integer
    
    For i = 0 To mnuTypeItem.UBound
        mnuTypeItem(i).Checked = i = Index
    Next
    lblTitle.Caption = mnuTypeItem(Index).Caption
    lblTitle.Caption = Split(lblTitle.Caption, "(")(0)
    
    lngRight = Me.Left + Me.Width
    LockWindowUpdate Me.Hwnd
    Call FillList(Index + 1)
    Me.Left = lngRight - Me.Width
    LockWindowUpdate 0
    
    If mfrmParent.Visible Then
        mfrmParent.SetFocus
    End If
End Sub

Private Sub picTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Call MoveObj(Me.Hwnd)
        mfrmParent.SetFocus
    End If
End Sub

Private Sub picTitle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call SetCloseButton(0)
    Call SetMenuButton(0)
End Sub

Private Sub vsList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mfrmParent.SetFocus
End Sub

Private Sub vsList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngRow As Long
    
    Call SetCloseButton(0)
    Call SetMenuButton(0)
    
    With vsList
        lngRow = .MouseRow
        If lngRow >= 0 And mlngPreRow <> lngRow Then
            Call ClearListSel
            .Cell(flexcpForeColor, lngRow, 0) = .ForeColorSel
            .Cell(flexcpBackColor, lngRow, 0) = .BackColorSel
            .CellBorderRange lngRow, 0, lngRow, 0, 1, 1, 1, 1, 1, 1, 1
            
            mlngPreRow = lngRow
            
            .ToolTipText = .Cell(flexcpText, lngRow, 0)
        End If
    End With
End Sub

Private Sub ClearListSel()
    With vsList
        If mlngPreRow >= 0 Then
            .Cell(flexcpForeColor, mlngPreRow, 0) = .ForeColor
            .Cell(flexcpBackColor, mlngPreRow, 0) = .BackColor
            .CellBorderRange mlngPreRow, 0, mlngPreRow, 0, 0, 0, 0, 0, 0, 0, 0
            mlngPreRow = -1
        End If
    End With
End Sub

Private Sub vsList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngRow As Long, lng分类ID As Long
    
    With vsList
        lngRow = .MouseRow
        If lngRow >= 0 Then
            lng分类ID = .Cell(flexcpData, lngRow, 0)
            If lng分类ID <> 0 Then
                Call ClearListSel
                RaiseEvent ItemClick(mintType, lng分类ID)
            End If
        End If
    End With
End Sub
