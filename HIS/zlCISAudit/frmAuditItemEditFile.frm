VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAuditItemEditFile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "文件选择"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6900
   Icon            =   "frmAuditItemEditFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4410
      TabIndex        =   2
      Top             =   6177
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5625
      TabIndex        =   1
      Top             =   6177
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfFiles 
      Height          =   5625
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   6660
      _cx             =   11747
      _cy             =   9922
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmAuditItemEditFile.frx":000C
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
      Editable        =   1
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
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Line Line6 
      X1              =   -45
      X2              =   10950
      Y1              =   5955
      Y2              =   5955
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000005&
      X1              =   -30
      X2              =   10950
      Y1              =   5970
      Y2              =   5985
   End
End
Attribute VB_Name = "frmAuditItemEditFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const conField = "Select case when b.COLUMN_VALUE is null then '0' else '1' end 选择 ,a.id as 文件ID,a.编号 as 文件编码,a.名称 as 文件名称,a.说明 as 文件说明" & vbCrLf & _
                         "from 病历文件列表 A, Table (Cast(f_Str2List([1])  As zlTools.t_StrList)) B " & vbCrLf & _
                         "where a.id = b.COLUMN_VALUE(+) And a.种类 = [2]"
Private Const conEmrField = "Select /*+ Rule*/" & vbNewLine & _
                        " Decode(a.Id, Null, '0', '1') 选择, Rawtohex(b.Id) As 文件id, b.Code As 文件编码, b.Title As 文件名称, b.Note As 文件说明" & vbNewLine & _
                        "From (Select Hextoraw(Column_Value) As ID From Table(Zlcommunal.f_Str2list(:p0, ','))) A, Antetype_List B" & vbNewLine & _
                        "Where a.Id(+) = b.Id And b.Kind = :p1" & vbNewLine & _
                        "Order By 选择 Desc, 文件编码"
Private mintItemFileID          As String
Private mstrType                As String
Private mintSource              As Integer

Public Property Get intSource() As Integer
    intSource = mintSource
End Property

Public Property Let intSource(ByVal vNewValue As Integer)
    mintSource = vNewValue
End Property
Public Property Get intItemFileID() As String
    intItemFileID = mintItemFileID
End Property

Public Property Let intItemFileID(ByVal vNewValue As String)
    mintItemFileID = vNewValue
End Property

Public Property Get strType() As String
    strType = mstrType
End Property

Public Property Let strType(ByVal vNewValue As String)
    mstrType = vNewValue
End Property

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngLoop         As Long
    Dim strFiles        As String
    
    With vsfFiles
        If .Rows <= 1 Then
            MsgBox "没有可供选择的文件！", vbExclamation, ParamInfo.产品名称
            Exit Sub
        End If
        strFiles = ""
        For lngLoop = 1 To .Rows - 1
            If Abs(.TextMatrix(lngLoop, .ColIndex("选择"))) = "1" Then
                strFiles = strFiles & .TextMatrix(lngLoop, .ColIndex("文件ID")) & ","
            End If
        Next
        If Len(strFiles) <> 0 Then
            strFiles = Left(strFiles, Len(strFiles) - 1)
            If LenB(strFiles) > 2000 Then
                MsgBox "选择文件过多，请重新选择", vbCritical, ParamInfo.产品名称
                Exit Sub
            End If
        End If
    End With
    mintItemFileID = strFiles
    Unload Me
End Sub

Private Sub Form_Load()
Dim rsTemp As ADODB.Recordset, strReturn As String
    On Error GoTo errHand
    If intSource = 0 Then
        gstrSQL = conField
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mintItemFileID, mstrType)
    Else
        gstrSQL = conEmrField
        strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, mintItemFileID & "^" & DbType.T_String & "^p0|" & strType & "^" & DbType.T_String & "^p1", rsTemp)
        If strReturn <> "" Then
            MsgBox strReturn, vbCritical, ParamInfo.产品名称
            Exit Sub
        End If
    End If
    Set vsfFiles.DataSource = rsTemp
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfFiles_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsfFiles.ColKey(Col) <> "选择" Then Cancel = True
End Sub

