VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmEPRSignVerify 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   Icon            =   "frmEPRSignVerify.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   1
      Left            =   45
      TabIndex        =   4
      Top             =   3210
      Width           =   7050
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "退出(&E)"
      Height          =   350
      Left            =   5745
      TabIndex        =   3
      ToolTipText     =   "<ESC>退出窗口"
      Top             =   3300
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   7050
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgThis 
      Height          =   2085
      Left            =   120
      TabIndex        =   0
      Top             =   885
      Width           =   6930
      _cx             =   12224
      _cy             =   3678
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
      RowHeightMin    =   360
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
      Begin VB.CommandButton cmdVerify 
         Caption         =   "验证签名(&V)"
         Height          =   350
         Left            =   5640
         TabIndex        =   5
         Top             =   360
         Width           =   1200
      End
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "该病历可以进行数字验证的签名版次如下，请选择需要验证的版次进行验证。"
      Height          =   180
      Left            =   555
      TabIndex        =   1
      Top             =   180
      Width           =   6375
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   0
      Picture         =   "frmEPRSignVerify.frx":08CA
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmEPRSignVerify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlFiledId As Long
Public Sub ShowMe(ByVal frmParent As Object, ByVal lFileId As Long)
    Dim rsTemp As New ADODB.Recordset
    mlFiledId = lFileId
    gstrSQL = "Select c.Id, c.对象标记, c.开始版 As 版本," & vbNewLine & _
            "       Decode(l.病历种类, 4, Decode(c.要素表示, 3, '护士长', '护士'), Decode(c.要素表示, 3, '主任医师', 2, '主治医师', '经治医师')) ||" & vbNewLine & _
            "        Decode(c.开始版, 1, '签名', '修订') As 操作," & vbNewLine & _
            "       Decode(Nvl(Instr(c.内容文本, ';'), 0), 0, c.内容文本, Substr(c.内容文本, 1, Instr(c.内容文本, ';') - 1)) As 人员," & vbNewLine & _
            "       RTrim(Substr(c.对象属性, Instr(c.对象属性, ';', 1, 4) + 1, Instr(c.对象属性, ';', 1, 5) - Instr(c.对象属性, ';', 1, 4) - 1)) As 时间,'' as 验证,要素单位 签名ID" & vbNewLine & _
            "From 电子病历记录 L, 电子病历内容 C" & vbNewLine & _
            "Where l.Id = c.文件id And l.Id = [1] And c.对象类型 = 8 And" & vbNewLine & _
            "      Substr(c.对象属性, Instr(c.对象属性, ';', 1, 1) + 1, Instr(c.对象属性, ';', 1, 2) - Instr(c.对象属性, ';', 1, 1) - 1) = 2" & vbNewLine & _
            "Order By 对象标记"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取签名版次", mlFiledId)
    Set vfgThis.DataSource = rsTemp
    With vfgThis
        .ColWidth(0) = 0    'ID
        .ColWidth(1) = 0    '对象标记
        .ColWidth(2) = 600  '版本
        .ColWidth(3) = 1400 '操作
        .ColWidth(4) = 1400 '人员
        .ColWidth(5) = 1800 '时间
        .ColWidth(6) = 800 '验证按扭
        .ColWidth(7) = 0   '签名ID
    End With
    vfgThis.Col = 1
    Me.Show 1, frmParent
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdVerify_Click()
    '数字签名验证
    Dim strSource As String, lngSignID As String
    
    On Error GoTo errHand
    lngSignID = vfgThis.TextMatrix(vfgThis.Row, 7)
    strSource = GetSignSourceFromDB(mlFiledId, vfgThis.TextMatrix(vfgThis.Row, 1))
    If gobjESign Is Nothing Then
        Set gobjESign = CreateObject("zl9ESign.clsESign")
        Call gobjESign.Initialize(gcnOracle, glngSys)
    End If
    Call gobjESign.VerifySignature(strSource, lngSignID, 2)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vfgThis_RowColChange()
    If Val(vfgThis.Tag) <> vfgThis.Row Then
        vfgThis.Tag = vfgThis.Row
        cmdVerify.Move vfgThis.Left + 5100, vfgThis.Row * 350 + 30
    End If
End Sub

