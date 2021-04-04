VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmEventMsgDeliveryApply 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "投递应用"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7080
   Icon            =   "frmEventMsgDeliveryApply.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4575
      TabIndex        =   10
      Top             =   7185
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(C)"
      Height          =   350
      Left            =   5895
      TabIndex        =   9
      Top             =   7185
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   4890
      Left            =   15
      TabIndex        =   4
      Top             =   2100
      Width           =   7005
      Begin VB.PictureBox picPane 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   2
         Left            =   885
         ScaleHeight     =   270
         ScaleWidth      =   1875
         TabIndex        =   13
         Top             =   150
         Width           =   1875
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   -15
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   -15
            Width           =   1905
         End
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "全清(&C)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   5790
         TabIndex        =   12
         Top             =   1560
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "全选(&A)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   5790
         TabIndex        =   11
         Top             =   1110
         Width           =   1100
      End
      Begin VB.PictureBox picPane 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   3690
         Index           =   1
         Left            =   315
         ScaleHeight     =   3690
         ScaleWidth      =   5385
         TabIndex        =   8
         Top             =   1125
         Width           =   5385
         Begin VSFlex8Ctl.VSFlexGrid vsf 
            Height          =   3660
            Index           =   1
            Left            =   15
            TabIndex        =   15
            Top             =   15
            Width           =   5355
            _cx             =   9446
            _cy             =   6456
            Appearance      =   0
            BorderStyle     =   0
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
            ForeColor       =   0
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16772055
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483638
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483638
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   270
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
         End
      End
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         Caption         =   "&3.指定业务事件的消息"
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   2
         Left            =   60
         TabIndex        =   7
         Top             =   765
         Width           =   2325
      End
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         Caption         =   "&2.所有业务事件的消息"
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   60
         TabIndex        =   6
         Top             =   435
         Width           =   4350
      End
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         Caption         =   "&1.所有                     类型的事件的消息"
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   60
         TabIndex        =   5
         Top             =   120
         Value           =   -1  'True
         Width           =   4755
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   3
      Left            =   30
      ScaleHeight     =   1515
      ScaleWidth      =   6990
      TabIndex        =   2
      Top             =   345
      Width           =   6990
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1485
         Index           =   0
         Left            =   15
         TabIndex        =   3
         Top             =   15
         Width           =   6960
         _cx             =   12277
         _cy             =   2619
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
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
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
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
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
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "指定消息"
      Height          =   180
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   1920
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "将下面已勾选中的投递目标服务应用到下面指定的消息中"
      Height          =   180
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   4500
   End
End
Attribute VB_Name = "frmEventMsgDeliveryApply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mfrmParent As Object
Private mbytMode As Byte
Private mblnDataChanged As Boolean
Private mblnReading As Boolean
Private mrsPara As ADODB.Recordset
Private mstrDataKey As String
Private mlngModualCode As Long
Private mblnContiune As Boolean
Private mclsVsf(1) As zlVSFlexGrid.clsVsf

Public Event AfterModifyData(ByVal DataKey As String)


Public Function ShowDialog(ByVal frmParent As Object, ByVal strDataKey As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Set mfrmParent = frmParent
    mstrDataKey = strDataKey
    
    Call InitGrid
    Call ReadData(mstrDataKey)
    Call opt_Click(0)
    
    Me.Show 1, mfrmParent
    
    ShowDialog = True
    
End Function

Private Function InitGrid() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    '初始网格控件
    Set mclsVsf(0) = New zlVSFlexGrid.clsVsf
    With mclsVsf(0)
        Call .Initialize(Me.Controls, vsf(0), True, True, GetImageList(16))
        Call .ClearColumn
        
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTBoolean, "", "[选择]", False)
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("名称", 1800, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("程序", 750, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("设备", 750, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("接口", 3000, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("注释", 15, flexAlignLeftCenter, flexDTString, , "", True)
                
        Call .InitializeEdit(True, False, False)
        Call .InitializeEditColumn(.ColIndex("选择"), True, vbVsfEditCheck)
        
        .AppendRows = True
        
    End With
    
    Set mclsVsf(1) = New zlVSFlexGrid.clsVsf
    With mclsVsf(1)
        Call .Initialize(Me.Controls, vsf(1), True, True, GetImageList(16))
        Call .ClearColumn
        
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTBoolean, "", "[选择]", False)
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("事件", 3000, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("消息", 900, flexAlignLeftCenter, flexDTString, , "", True)
                
        Call .InitializeEdit(True, False, False)
        Call .InitializeEditColumn(.ColIndex("选择"), True, vbVsfEditCheck)
        
        vsf(1).RowHidden(0) = True
        
        .AppendRows = True
    End With
    
    InitGrid = True
    
End Function

Private Function ReadData(ByVal strDataKey As String) As Boolean

    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim rsTmp As ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    
    mblnReading = True
        
    cbo.Clear
    Set rsTmp = gclsBusiness.ReadEventKind
    If rsTmp.BOF = False Then
        Do While Not rsTmp.EOF
            cbo.AddItem rsTmp("kind").Value
            rsTmp.MoveNext
        Loop
    End If
    If cbo.ListCount > 0 Then cbo.ListIndex = 0
    
    '------------------------------------------------------------------------------------------------------------------
    Set rsCondition = zlCommFun.CreateCondition
    Call zlCommFun.SetCondition(rsCondition, "事件消息id", strDataKey)
    Call mclsVsf(0).LoadDataSource(gclsBusiness.EventMsgServerRead("配置", rsCondition))
    
    '------------------------------------------------------------------------------------------------------------------
    Call mclsVsf(1).LoadDataSource(gclsBusiness.ReadEventMsg("所有"))
    
    mblnReading = False
    mblnDataChanged = False
    
    ReadData = True
    
End Function

Private Function ValidData() As Boolean
    '******************************************************************************************************************
    '功能：校验编辑数据的有效性
    '参数：
    '返回：
    '******************************************************************************************************************
        
    '
    
    ValidData = True
    
End Function

Private Function SaveData(ByRef strDataKey As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsPara As ADODB.Recordset
    Dim strTemp As String
    Dim lngLoop As Long
        
    On Error GoTo errHand
    
    Set rsPara = zlCommFun.CreateParameter
        
    '------------------------------------------------------------------------------------------------------------------
    strTemp = ""
    With vsf(0)
        For lngLoop = 1 To .Rows - 1
            If Abs(Val(.TextMatrix(lngLoop, .ColIndex("选择")))) = 1 Then
                strTemp = strTemp & ";" & Trim(.TextMatrix(lngLoop, .ColIndex("id")))
            End If
        Next
    End With
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    
    Call zlCommFun.SetParameter(rsPara, "TargetService", strTemp)
    If opt(0).Value = True Then
        Call zlCommFun.SetParameter(rsPara, "ConfigOption", 1)
        Call zlCommFun.SetParameter(rsPara, "ConfigString", cbo.Text)
    ElseIf opt(1).Value = True Then
        Call zlCommFun.SetParameter(rsPara, "ConfigOption", 2)
        Call zlCommFun.SetParameter(rsPara, "ConfigString", "")
    Else
        Call zlCommFun.SetParameter(rsPara, "ConfigOption", 3)
        strTemp = ""
        With vsf(1)
            For lngLoop = 1 To .Rows - 1
                If Abs(Val(.TextMatrix(lngLoop, .ColIndex("选择")))) = 1 Then
                    strTemp = strTemp & ";" & Trim(.TextMatrix(lngLoop, .ColIndex("id")))
                End If
            Next
        End With
        If strTemp <> "" Then strTemp = Mid(strTemp, 2)
        Call zlCommFun.SetParameter(rsPara, "ConfigString", strTemp)
    End If
    
    SaveData = gclsBusiness.EventMsgEdit("BatchTargetConfig", rsPara)
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClearAll_Click()
    With vsf(1)
        .Cell(flexcpText, 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = 0
    End With
End Sub

Private Sub cmdOK_Click()

    If ValidData = True Then
            
        If SaveData(mstrDataKey) = True Then
            
            RaiseEvent AfterModifyData(mstrDataKey)

            mblnDataChanged = False
            Unload Me
            
        End If
    End If

End Sub

Private Sub cmdSelectAll_Click()
    With vsf(1)
        .Cell(flexcpText, 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = 1
    End With
End Sub

Private Sub opt_Click(Index As Integer)
    Select Case Index
    Case 0
        vsf(1).Enabled = False
        cmdSelectAll.Enabled = False
        cmdClearAll.Enabled = False
        cbo.Enabled = True
        vsf(1).Cell(flexcpForeColor, 0, 0, vsf(1).Rows - 1, vsf(1).Cols - 1) = 12632256
    Case 1
        vsf(1).Enabled = False
        cmdSelectAll.Enabled = False
        cmdClearAll.Enabled = False
        cbo.Enabled = False
        vsf(1).Cell(flexcpForeColor, 0, 0, vsf(1).Rows - 1, vsf(1).Cols - 1) = 12632256
    Case 2
        vsf(1).Enabled = True
        cmdSelectAll.Enabled = True
        cmdClearAll.Enabled = True
        cbo.Enabled = False
        vsf(1).Cell(flexcpForeColor, 0, 0, vsf(1).Rows - 1, vsf(1).Cols - 1) = 0
        
    End Select
End Sub

Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    '编辑处理
    Call mclsVsf(Index).AfterEdit(Row, Col)
    mblnDataChanged = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf(Index).AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf(Index).BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_DblClick(Index As Integer)
        
    Call mclsVsf(Index).DbClick
    
    With vsf(Index)
        If Abs(Val(.TextMatrix(.Row, .ColIndex("选择")))) = 1 Then
            .TextMatrix(.Row, .ColIndex("选择")) = 0
        Else
            .TextMatrix(.Row, .ColIndex("选择")) = 1
        End If
    End With
    
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
        
    With vsf(Index)
        If KeyAscii = vbKeySpace Then
            Call vsf_DblClick(Index)
        End If
    End With
    
    Call mclsVsf(Index).KeyPress(KeyAscii)
End Sub

Private Sub vsf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call mclsVsf(Index).KeyDown(KeyCode, Shift)
End Sub

Private Sub vsf_SetupEditWindow(Index As Integer, ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call mclsVsf(Index).EditSelAll
End Sub

Private Sub vsf_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '编辑处理
    Call mclsVsf(Index).BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf(Index).ValidateEdit(Col, Cancel)
End Sub


