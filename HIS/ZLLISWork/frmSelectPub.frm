VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSelectPub 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选择"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5295
   Icon            =   "frmSelectPub.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5025
      Left            =   135
      ScaleHeight     =   5025
      ScaleWidth      =   4965
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   75
      Width           =   4965
      Begin VB.TextBox txtSel 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   165
         TabIndex        =   3
         Top             =   270
         Width           =   4515
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1830
         TabIndex        =   2
         Top             =   4530
         Width           =   1245
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3450
         TabIndex        =   1
         Top             =   4530
         Width           =   1245
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFItemSel 
         Height          =   2775
         Left            =   75
         TabIndex        =   4
         Top             =   930
         Width           =   4635
         _cx             =   8176
         _cy             =   4895
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
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
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
   End
End
Attribute VB_Name = "frmSelectPub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mRecordSource As ADODB.Recordset
Private mstrValue As String
Private mlngID As Long

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    GetRowValue
End Sub

Private Sub Form_Activate()
    Me.txtSel.SetFocus
End Sub

Private Sub Form_Resize()
    With Me.picItem
        .Top = 0
        .Left = 0
        .Width = Me.ScaleWidth - 0
        .Height = Me.ScaleHeight - 0
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.txtSel.Text = ""
End Sub

Private Sub picItem_Resize()
    With Me.txtSel
        .Top = 100
        .Left = 100
        .Width = Me.picItem.ScaleWidth - 200
    End With
    
    With Me.VSFItemSel
        .Top = Me.txtSel.Top + Me.txtSel.Height + 100
        .Left = 100
        .Width = Me.picItem.ScaleWidth - 200
        .Height = Me.picItem.ScaleHeight - .Top - Me.cmdOK.Height - 300
    End With
    
    With Me.cmdCancel
        .Top = Me.VSFItemSel.Top + Me.VSFItemSel.Height + 180
        .Left = Me.ScaleWidth - .Width - 300
    End With
    
    
    With Me.cmdOK
        .Top = cmdCancel.Top
        .Left = Me.cmdCancel.Left - .Width - 300
    End With
    
    Call PicDrowBorder(picItem)
    Call PicDrowSplit(picItem, Me.VSFItemSel)
End Sub

Public Function ShowMe(ByVal formParent As Object, ByVal RecordSource As Recordset, ByVal strFind As String, Optional lngID As Long) As String
    '功能   打开公共的选择器（单列)
    '参数   RecordSource    传入要查询的记录集
    '       strField        过滤字段
    '       strFind         过滤字段的查询条件
    Dim strFilter As String
    mstrValue = ""
    mlngID = lngID
    Set mRecordSource = RecordSource
    strFind = Trim(strFind)
    
    If strFind <> "" Then
        strFilter = GetFindString(RecordSource, strFind, lngID, "")
    Else
        strFilter = ""
        If lngID > 0 Then
            strFilter = "id=" & lngID
        End If
    End If
    
    mRecordSource.filter = strFilter
    
    If mRecordSource.RecordCount <> 1 Then
        If mRecordSource.RecordCount = 0 Then
            mRecordSource.filter = ""
            If lngID > 0 Then
                strFilter = "id=" & lngID
            End If
            mRecordSource.filter = strFilter
        End If
        Load frmSelectPub
        InitPublicDicVsf Me.VSFItemSel, mRecordSource, ""
        Me.txtSel.Text = strFind
        frmSelectPub.Show vbModal, formParent
        If mRecordSource.RecordCount > 0 Then
            Me.txtSel.Text = strFind
        End If
    Else
        InitPublicDicVsf Me.VSFItemSel, mRecordSource, ""
        mstrValue = GetVSFRowValue(VSFItemSel, VSFItemSel.Row, "")
    End If
    ShowMe = mstrValue
End Function

Private Sub txtSel_Change()
    If txtSel.Text <> "" Then
        mRecordSource.filter = GetFindString(mRecordSource, txtSel.Text, mlngID, "")
    Else
        If mlngID = 0 Then
            mRecordSource.filter = ""
        Else
            mRecordSource.filter = "id=" & mlngID
        End If
    End If
    InitPublicDicVsf Me.VSFItemSel, mRecordSource, ""
End Sub

Private Sub txtSel_GotFocus()
    txtSel.SelStart = 0
    txtSel.SelLength = Len(txtSel.Text)
End Sub

Private Sub txtSel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then
        '向上按键
        With Me.VSFItemSel
            If .Row > 1 Then
                .Row = .Row - 1
            End If
        End With
    End If
    If KeyCode = 40 Then
        With Me.VSFItemSel
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
            End If
        End With
    End If
End Sub

Private Sub txtSel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        GetRowValue
    End If
End Sub

Private Sub VSFItemSel_DblClick()
    GetRowValue
End Sub

Private Sub VSFItemSel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        GetRowValue
    End If
End Sub

Private Function GetRowValue()
    '功能           返回行结果
    mstrValue = GetVSFRowValue(VSFItemSel, VSFItemSel.Row, "")
    Unload Me
End Function

Private Function InitPublicDicVsf(vsfList As VSFlexGrid, RecordSource As Recordset, ByRef strErr As String) As Boolean
    
    On Error GoTo errH
    '初使化表格控件
    If Not vfgLoadFromRecord(vsfList, RecordSource, strErr) Then Exit Function
    
    With vsfList
        If .Cols - 1 >= 1 Then
            .ColWidth(1) = 1300: .ColHidden(1) = False
        End If
        If .Cols - 1 >= 2 Then
            .ColWidth(2) = 2500: .ColHidden(2) = False
        End If
        If .Cols - 1 >= 3 Then
            .ColWidth(3) = 300: .ColHidden(3) = False
        End If
        If .Cols - 1 < 1 Then
            '只有一列，则显示
            .ColWidth(0) = 1300: .ColHidden(0) = False
        End If
    End With
    Exit Function
errH:
    strErr = Err.Number & " " & Err.Description
End Function

Private Function GetFindString(RecordSource As Recordset, strFind As String, Optional lngID As Long, Optional ByRef strErr As String) As String
    '功能   从数据源中提取过滤字段并生成过滤字串
    '参数   RecordSource 数据源
    '       strFind 过滤字串
    Dim intLoop As Integer
    On Error GoTo errH
    For intLoop = 1 To RecordSource.Fields.Count - 1
        If RecordSource.Fields(intLoop).Type = 200 Then
            If strFind <> "" Then
                If lngID = 0 Then
                    GetFindString = GetFindString & "or " & RecordSource.Fields(intLoop).Name & " like '*" & strFind & "*' "
                Else
                    GetFindString = GetFindString & "or (" & RecordSource.Fields(intLoop).Name & " like '*" & strFind & "*' " & _
                                    " and id = " & lngID & " )"
                End If
            End If
        End If
    Next
    If GetFindString <> "" Then
        GetFindString = Mid(GetFindString, 3)
    End If
    Exit Function
errH:
    strErr = Err.Number & " " & Err.Description
End Function

Private Function GetVSFRowValue(vsfList As VSFlexGrid, intRow As Integer, ByRef strErr As String) As String
    '功能       到得当前行的值
    Dim intLoop As Integer
    On Error GoTo errH
    With vsfList
        For intLoop = 0 To vsfList.Cols - 1
            GetVSFRowValue = GetVSFRowValue & "," & .TextMatrix(intRow, intLoop)
        Next
        GetVSFRowValue = Mid(GetVSFRowValue, 2)
    End With
errH:
    strErr = Err.Number & " " & Err.Description
End Function

Public Sub PicDrowBorder(Picobj As PictureBox, Optional lngLineColour As Long = -1)
    '功能       画图片边框
    On Error Resume Next
    With Picobj
        .AutoRedraw = True
        .Cls
        .DrawWidth = 2
        
        If lngLineColour = -1 Then
            .ForeColor = &HE0E0E0
        Else
            .ForeColor = lngLineColour
        End If
        Picobj.Line (25, 25)-(.Width - 50, .Height - 50), , B
    End With
End Sub

Public Sub PicDrowSplit(Picobj As PictureBox, objSplit As Object, Optional lngHeightSplit As Long)
    '功能       画图片的分隔线
    On Error Resume Next
    With Picobj
        .AutoRedraw = True
        If lngHeightSplit = 0 Then
            Picobj.Line (25, objSplit.Top + objSplit.Height + 70)-(.Width - 50, objSplit.Top + objSplit.Height + 70), , B
        Else
            Picobj.Line (25, objSplit.Top + objSplit.Height + lngHeightSplit)-(.Width - 50, objSplit.Top + objSplit.Height + lngHeightSplit), , B
        End If
    End With
End Sub

