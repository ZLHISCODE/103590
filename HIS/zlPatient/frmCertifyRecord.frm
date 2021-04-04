VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmCertifyRecord 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "实名信息变更记录"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11625
   Icon            =   "frmCertifyRecord.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   11625
   StartUpPosition =   2  '屏幕中心
   Begin VSFlex8Ctl.VSFlexGrid vsfRecord 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      _cx             =   20558
      _cy             =   5530
      Appearance      =   0
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
      BackColorSel    =   16777215
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   9
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   325
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
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
Attribute VB_Name = "frmCertifyRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng实名id As Long
Dim mrsRescord As New ADODB.Recordset

Private Enum VSFRECORD_INFO
    COL_实名ID = 0
    COL_姓名
    COL_变更项目
    COL_原信息
    COL_新信息
    COL_变更人
    COL_变更时间
    COL_变更原因
End Enum
Private Sub InitVsfGridHeader()
'功能：初始化列表
    Dim strHeader As String
    
    strHeader = "实名ID;姓名,1000,1;变更项目,1500,1;原信息,3500,1;新信息,3500,1;变更人,1000,4;变更时间,1500,1;变更原因,3500,1"
    Call grid.Init(vsfRecord, strHeader)
End Sub

Private Sub Form_Load()
    Call InitVsfGridHeader
    LoadCertifyRecord mrsRescord
End Sub

Private Sub LoadCertifyRecord(ByVal rsTmp As ADODB.Recordset)
'功能：加载病人实名信息变更记录
    Dim i As Long, j As Long
    
    With vsfRecord
        j = .FixedRows
        If Not rsTmp.EOF Then
            For i = 0 To rsTmp.RecordCount - 1
                .AddItem "", j
                .TextMatrix(j, COL_实名ID) = "" & rsTmp!实名ID
                .TextMatrix(j, COL_姓名) = "" & rsTmp!姓名
                .TextMatrix(j, COL_变更项目) = "" & rsTmp!变更项目
                If Trim("" & rsTmp!变更项目) = "陪诊人出生日期" Or Trim("" & rsTmp!变更项目) = "出生日期" Then
                    .TextMatrix(j, COL_原信息) = Format("" & rsTmp!原信息, "YYYY-MM-DD HH:MM")
                    .TextMatrix(j, COL_新信息) = Format("" & rsTmp!新信息, "YYYY-MM-DD HH:MM")
                ElseIf Trim("" & rsTmp!变更项目) = "陪诊人身份证类型" Or Trim("" & rsTmp!变更项目) = "身份证类型" Then
                    .TextMatrix(j, COL_原信息) = decode(Val("" & rsTmp!原信息), 1, "居民身份证", 2, "港澳台居住证", 3, "外国人居留证", 0, "", -1, "", "" & rsTmp!原信息)
                    .TextMatrix(j, COL_新信息) = decode(Val("" & rsTmp!新信息), 1, "居民身份证", 2, "港澳台居住证", 3, "外国人居留证", 0, "", -1, "", "" & rsTmp!新信息)
                Else
                    .TextMatrix(j, COL_原信息) = "" & rsTmp!原信息
                    .TextMatrix(j, COL_新信息) = "" & rsTmp!新信息
                End If
                .TextMatrix(j, COL_变更时间) = Format("" & rsTmp!变更时间, "YYYY-MM-DD HH:MM")
                .TextMatrix(j, COL_变更人) = "" & rsTmp!变更人
                .TextMatrix(j, COL_变更原因) = "" & rsTmp!变更原因
                rsTmp.MoveNext
                j = j + 1
            Next
        End If
    End With
End Sub

Public Function ShowMe(frmParent As Object, ByVal lng实名ID As Long) As Boolean
    mlng实名id = lng实名ID
    If mlng实名id = 0 Then
        MsgBox "请选择一个病人！", vbInformation, gstrSysName
    Else
        Set mrsRescord = GetCertifyRecord(mlng实名id)
        If mrsRescord.EOF Then
            MsgBox "该病人没有实名信息变更记录！", vbInformation, gstrSysName
        Else
            If Not frmParent Is Nothing Then
                Me.Show , frmParent
            End If
        End If
    End If
End Function
