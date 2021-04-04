VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEPRSearchPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "请选择需要打印列表"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13050
   Icon            =   "frmEPRSearchPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   13050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboPrinter 
      Height          =   300
      Left            =   5835
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   5680
      Width           =   3915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   10140
      TabIndex        =   3
      Top             =   5640
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   11790
      TabIndex        =   2
      Top             =   5655
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      Caption         =   "待输出清单(由主界面条件过滤)"
      Height          =   5415
      Left            =   30
      TabIndex        =   0
      Top             =   120
      Width           =   12975
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   5010
         Left            =   45
         TabIndex        =   1
         ToolTipText     =   "双击选中"
         Top             =   315
         Width           =   12870
         _cx             =   22701
         _cy             =   8837
         Appearance      =   2
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
         GridColor       =   12698049
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
         MergeCells      =   1
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
         Begin MSComctlLib.ImageList img16 
            Left            =   0
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEPRSearchPrint.frx":000C
                  Key             =   "Selected"
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.Label Label1 
      Caption         =   "打印机"
      Height          =   210
      Left            =   4995
      TabIndex        =   5
      Top             =   5730
      Width           =   780
   End
End
Attribute VB_Name = "frmEPRSearchPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrSelect As String, mstrPrint As String
Public Function ShowMe(ByVal frmObj As Object, ByVal parVsf As VSFlexGrid, ByRef strPrint As String) As String
Dim intCount As Integer
    
    On Error GoTo errHand
    With cboPrinter
        .Clear
        For intCount = 0 To Printers.Count - 1
            .AddItem Printers(intCount).DeviceName
            If Printers(intCount).DeviceName = Printer.DeviceName Then .ListIndex = intCount
        Next
    End With
    mstrSelect = ""
    Call FillVfg(parVsf)
    
    Me.Show 1, frmObj
    strPrint = mstrPrint
    ShowMe = mstrSelect
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub FillVfg(ByVal parVsf As VSFlexGrid)
Dim i As Integer, l As Integer
    On Error GoTo errHand
    With vsf
        .Clear
        .Rows = parVsf.Rows
        .Cols = parVsf.Cols + 1
        .ROWHEIGHT(0) = 350
        .TextMatrix(0, 0) = "选择": .ColWidth(0) = 400
        For i = 1 To parVsf.Cols
            .TextMatrix(0, i) = parVsf.TextMatrix(0, i - 1)
            .ColWidth(i) = parVsf.ColWidth(i - 1)
            
        Next
                
        For i = 1 To parVsf.Rows - 1
            .ROWHEIGHT(i) = 350
            .Cell(flexcpData, i, 0) = 0
            For l = 1 To parVsf.Cols
                .TextMatrix(i, l) = parVsf.TextMatrix(i, l - 1)
            Next
        Next
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
        If .Rows = 2 Then vsf_DblClick
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub cmdCancel_Click()
    mstrSelect = ""
    Unload Me
End Sub

Private Sub cmdOk_Click()
Dim i As Integer
Dim lngId As Long, EType As Byte
    mstrPrint = cboPrinter.Text
    On Error GoTo errHand
    With vsf
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, 0) = 1 Then
                lngId = Val(.TextMatrix(i, 1))
                EType = Val(.TextMatrix(i, 15))
                mstrSelect = mstrSelect & "|" & EType & "," & lngId
            End If
        Next
        If mstrSelect <> "" Then
            mstrSelect = Mid(mstrSelect, 2)
        End If
    End With
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mfrmPrint_PrintEpr(ByVal lngRecordId As Long)
    Debug.Print lngRecordId
End Sub

Private Sub vsf_DblClick()
Dim lngRow As Long
    With vsf
        lngRow = .Row
        If lngRow < 1 Then Exit Sub
        If .Cell(flexcpData, lngRow, 0) = 0 Then
            .Cell(flexcpData, lngRow, 0) = 1
            Set .Cell(flexcpPicture, lngRow, 0) = img16.ListImages("Selected").Picture
        Else
            .Cell(flexcpData, lngRow, 0) = 0
            Set .Cell(flexcpPicture, lngRow, 0) = Nothing
        End If
    End With
End Sub


