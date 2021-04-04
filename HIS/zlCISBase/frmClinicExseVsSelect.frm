VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmClinicExseVsSelect 
   BorderStyle     =   0  'None
   Caption         =   "VS选择器"
   ClientHeight    =   3585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fraSelect 
      Height          =   3675
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   6390
      Begin VB.CommandButton cmdCancle 
         Appearance      =   0  'Flat
         Height          =   320
         Left            =   5655
         Picture         =   "frmClinicExseVsSelect.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3285
         Width           =   350
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Height          =   320
         Left            =   5130
         Picture         =   "frmClinicExseVsSelect.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3285
         Width           =   350
      End
      Begin VB.OptionButton opt类型 
         Caption         =   "本项目部位"
         Height          =   270
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   3315
         Width           =   1275
      End
      Begin VB.OptionButton opt类型 
         Caption         =   "所有部位"
         Height          =   270
         Index           =   1
         Left            =   1515
         TabIndex        =   1
         Top             =   3315
         Width           =   1275
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgSelect 
         Height          =   3075
         Left            =   60
         TabIndex        =   3
         Top             =   165
         Width           =   6270
         _cx             =   11060
         _cy             =   5424
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
         BackColorFixed  =   15790320
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
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
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
         AutoResize      =   0   'False
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
End
Attribute VB_Name = "frmClinicExseVsSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnOK As Boolean
Private mrsInput As ADODB.Recordset
Private mbytType As Byte '1-选择部位 2-选择方法 3-选择收费项目
Private mstrReturn As String

Public Sub ShowSelect(ByVal bytType As Byte, ByVal rsInput As Recordset, ByRef strReturn As String)
    mblnOK = False: mbytType = bytType: Set mrsInput = rsInput
    Me.Show vbModal
    If mblnOK And mstrReturn <> strReturn Then strReturn = mstrReturn
End Sub


Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    mstrReturn = ""
    For i = 0 To vfgSelect.Cols - 1
        mstrReturn = mstrReturn & vfgSelect.TextMatrix(vfgSelect.Row, i) & "|"
    Next
    If mstrReturn <> "" Then mblnOK = True
    Unload Me
End Sub

Private Sub Form_Activate()
    '1-选择部位 2-选择方法 3-选择收费项目
    
    opt类型(0).Visible = False
    opt类型(1).Visible = False
    
    If mbytType = 1 Then
        opt类型(0).Visible = True
        opt类型(1).Visible = True
        opt类型(0).Value = True
        opt类型(0).Caption = "指定部位"
        opt类型(1).Caption = "所有部位"
    ElseIf mbytType = 2 Then
        opt类型(0).Visible = True
        opt类型(1).Visible = True
        opt类型(1).Value = True
        opt类型(0).Caption = "默认方法"
        opt类型(1).Caption = "所有方法"
    End If
    
    Call initVfgSelect(mbytType)
    vfgSelect.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call cmdOK_Click
    If KeyCode = vbKeyEscape Then Call cmdCancle_Click
End Sub

Private Sub initVfgSelect(ByVal bytType As Byte)
    Dim i As Integer
    Dim aryItem() As String, strItems As String, strTemp As String
    Dim aryChild() As String, lngChild As Long, lngCount As Long, strMode As String
    
    With vfgSelect
        '初始化表格
        .Cols = mrsInput.Fields.Count: .Rows = 1
        .FixedCols = 0: .FixedRows = 1
        .RowHeight(0) = 350
        .RowHeightMin = 300
        For i = 0 To mrsInput.Fields.Count - 1
            .TextMatrix(0, i) = mrsInput.Fields(i).Name
        Next
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColHidden(0) = True
        '合并行
        .MergeCol(1) = True
        .MergeCells = flexMergeRestrictColumns
        
        If opt类型(0).Visible Then
            If opt类型(0).Value = True Then
                mrsInput.Filter = " 分类='已选'"
            Else
                mrsInput.Filter = " 分类='可选'"
            End If
        Else
            mrsInput.Filter = ""
        End If
        
        '填数据
        If bytType = 2 And opt类型(1).Value Then
            '全部 方法的显示 要单独处理
            Do Until mrsInput.EOF
                strMode = "" & mrsInput.Fields(1)
                If InStr(1, strMode, vbTab) > 0 Then strMode = Mid(strMode, 1, InStr(1, strMode, vbTab) - 1) & ";" & Mid(strMode, InStr(1, strMode, vbTab))
                For lngCount = 1 To Len(strMode)
                    If Mid(strMode, lngCount, 1) = vbTab And lngCount <> 2 Then
                         If Mid(strTemp, Len(strTemp), 1) <> ";" Then strTemp = strTemp & ";"
                    End If
                    strTemp = strTemp & Mid(strMode, lngCount, 1)
                Next
                strMode = strTemp
                
                aryItem() = IIf(Mid(strMode, 1, 1) = ";", Split(Mid(strMode, 2), ";"), Split(strMode, ";"))
                For lngCount = 0 To UBound(aryItem)
                    strTemp = aryItem(lngCount)
                    If strTemp <> "" Then
                        If InStr(1, aryItem(lngCount), ",") > 0 Then strTemp = Left(aryItem(lngCount), InStr(1, aryItem(lngCount), ",") - 1)
                        .Rows = .Rows + 1
                        
                        If InStr(1, strTemp, vbTab) = 0 Then





                            .TextMatrix(.Rows - 1, 1) = Mid(strTemp, 2)
                        Else
                            .TextMatrix(.Rows - 1, 1) = Mid(strTemp, 3)
                        End If
                        
                        If InStr(1, aryItem(lngCount), ",") > 0 Then
                            strTemp = Mid(aryItem(lngCount), InStr(1, aryItem(lngCount), ",") + 1)
                            aryChild = Split(strTemp, ",")
                            For lngChild = 0 To UBound(aryChild)
                                strTemp = aryChild(lngChild)
                                .Rows = .Rows + 1
                                .TextMatrix(.Rows - 1, 1) = Mid(strTemp, 2)
                            Next
                        End If
                    End If
                Next

'                If UBound(Split(strMode, vbTab)) > 0 Then
'                    aryItem() = Split(Split(strMode, vbTab)(1), ";")
'                    For lngCount = 0 To UBound(aryItem)
'                        strTemp = aryItem(lngCount)
'                        .Rows = .Rows + 1
'                        .TextMatrix(.Rows - 1, 1) = Mid(strTemp, 2)
'                    Next
'                End If

                mrsInput.MoveNext
            Loop
        Else
            Set .DataSource = mrsInput
        End If
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        If .Rows > .FixedRows Then .Row = .FixedRows
    End With
End Sub

Private Sub opt类型_Click(Index As Integer)
    Call initVfgSelect(mbytType)
End Sub

Private Sub vfgSelect_DblClick()
    Call cmdOK_Click
End Sub
