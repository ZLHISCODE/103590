VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSqlBindSelector 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "绑定变量选择器"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7320
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSqlBindSelector.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   6240
      TabIndex        =   3
      Top             =   3960
      Width           =   990
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "完成(&O)"
      Height          =   350
      Left            =   5160
      TabIndex        =   2
      Top             =   3960
      Width           =   990
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfBind 
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   7095
      _cx             =   12515
      _cy             =   5530
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
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
      GridColor       =   4210752
      GridColorFixed  =   -2147483640
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483640
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
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
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSqlBindSelector.frx":6852
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5400
   End
End
Attribute VB_Name = "frmSqlBindSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsBind As New ADODB.Recordset
Private mblnResult As Boolean

Private Enum GridColor
    ColorInactive = &HFCF7F4
    ColorActive = &HDCF8FF
End Enum
Public Function ShowSelector(ByRef rsBind As ADODB.Recordset, Optional ByVal ShowOriginal As Boolean = False) As Boolean
    'ShowOriginal = 当用于显示原句时,需要不同的提示信息
    
    If ShowOriginal Then
        lblTip.Caption = "请选择一组绑定变量,以便查看原句"
    End If
    
    Set mrsBind = rsBind

    Me.Show 1
    ShowSelector = mblnResult
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
        
    If vsfBind.Row = 0 Or vsfBind.Row = -1 Then
        MsgBox "请选择一组绑定变量", , "提示"
        Exit Sub
    End If
    
    mrsBind.Filter = "Hash_Value  = " & vsfBind.RowData(vsfBind.Row)
    
    mblnResult = True
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strCol As String, strTmp As String, strDate As String
    Dim i As Integer, intGroup As Integer, lngColor As Long
    
    strCol = "组号,1500,1;变量名,1500,1;位置,1500,1;变量值,1500,1"
    InitTable vsfBind, strCol
    
    
    With vsfBind
        .Redraw = flexRDNone
        .Rows = 1
        .Rows = mrsBind.RecordCount + 1
        i = 1
        Do While Not mrsBind.EOF
            If strTmp <> mrsBind!Hash_Value Then
                intGroup = intGroup + 1
                strTmp = mrsBind!Hash_Value
                lngColor = IIf(lngColor = GridColor.ColorActive, GridColor.ColorInactive, GridColor.ColorActive)
            End If
            
            .RowData(i) = strTmp
            .TextMatrix(i, .ColIndex("组号")) = intGroup
            .TextMatrix(i, .ColIndex("变量名")) = mrsBind!Name & ""
            .TextMatrix(i, .ColIndex("位置")) = mrsBind!Position & ""
            
            '时间类型查询出来的结果是 月份/日期/年 格式,需要特殊转换
            '10g为  : 04/18/19 00:00:00    11g为: 04/18/2019 00:00:00
            If mrsBind!DataType = 12 Or mrsBind!DataType = 180 Then '日期型参数
                If mrsBind!Value_String & "" <> "" Then

                    strDate = Split(mrsBind!Value_String, " ")(0)
                    If gstrBigVer = 10 Then
                        strDate = "20" & Split(strDate, "/")(2) & "/" & Split(strDate, "/")(0) & "/" & Split(strDate, "/")(1)
                        strDate = strDate & Right(mrsBind!Value_String, 9)
                    Else
                        strDate = Split(strDate, "/")(2) & "/" & Split(strDate, "/")(0) & "/" & Split(strDate, "/")(1)
                        strDate = strDate & Right(mrsBind!Value_String, 9)
                    End If
                End If
            Else
                strDate = mrsBind!Value_String & ""
            End If
            .TextMatrix(i, .ColIndex("变量值")) = strDate
            
            .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = lngColor
            i = i + 1
            mrsBind.MoveNext
        Loop
        
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(0) = True
        .AutoResize = True
        .AutoSize 0, .Cols - 1, False
        .Redraw = flexRDDirect
        .Select 1, 1
    End With
    
End Sub

