VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmDiseaseFileList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "病历文件选择"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6750
   Icon            =   "frmDiseaseFileList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6750
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   6750
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2910
      Width           =   6750
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   5520
         TabIndex        =   1
         Top             =   195
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   4200
         TabIndex        =   3
         Top             =   195
         Width           =   1100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   10000
         Y1              =   45
         Y2              =   45
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   10000
         Y1              =   30
         Y2              =   30
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsFileList 
      Height          =   2745
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _cx             =   11880
      _cy             =   4842
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
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   3
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
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   360
      RowHeightMax    =   360
      ColWidthMin     =   200
      ColWidthMax     =   5000
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmDiseaseFileList.frx":08CA
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
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
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmDiseaseFileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private mrsDiseaseFile As New ADODB.Recordset '需要填写的病历文件
Private mfrmParent     As Object '医生工作站界面
Private mlngFileID     As Long '选择的要书写的病历文的ID
Private mblnOk         As Boolean

'列的枚举
Private Enum Cols
    col编号 = 0
    col名称 = 1
    COL说明 = 2
End Enum

'病历文件选择窗体的显示
Public Function ShowMe(frmParent As Object, ByVal rsTmp As ADODB.Recordset, Optional ByRef lngFileID As Long) As Boolean
'参数：frmParent  父窗体
'      rsTmp      需要填写的病历文件记录集

    '保存传入变量的值
    Set mrsDiseaseFile = rsTmp
    Set mfrmParent = frmParent
    
    Me.Show 1, frmParent
    lngFileID = mlngFileID
    ShowMe = mblnOk
End Function

'取消时推出窗体
Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

'确定时显示病历文件编辑器
Private Sub cmdOK_Click()
    '如果没有选择文件，并确定进行询问
    If mlngFileID = 0 Then
        MsgBox "你还未选择病历文件，请选择病历文件", vbOKOnly, gstrSysName
    Else
        mblnOk = True
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    Dim i As Long
    '病历文件数据加载到VSGRID控件
    With vsFileList
        .Rows = .FixedRows
        mrsDiseaseFile.MoveFirst
        For i = 1 To mrsDiseaseFile.RecordCount
            '缺省不选择任何一行
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, col编号) = mrsDiseaseFile!编号
            .TextMatrix(.Rows - 1, col名称) = mrsDiseaseFile!名称
            .TextMatrix(.Rows - 1, COL说明) = "" & mrsDiseaseFile!说明
            .RowData(.Rows - 1) = mrsDiseaseFile!ID & ""
            mrsDiseaseFile.MoveNext
        Next
    End With
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    '确定退出时，并且文件已经选择时不阻止退出，否则阻止退出
    If mblnOk And mlngFileID = 0 Then
        Cancel = True
    Else
        Set mrsDiseaseFile = Nothing
        Set mfrmParent = Nothing
    End If
End Sub

'文件选定
Private Sub vsFileList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    mlngFileID = vsFileList.RowData(NewRow)
End Sub

'双击文件，打开文件编辑器
Private Sub vsFileList_DblClick()
    mlngFileID = vsFileList.RowData(vsFileList.Row)
    Call cmdOK_Click
End Sub


