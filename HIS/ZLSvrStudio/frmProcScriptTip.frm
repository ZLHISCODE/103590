VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmProcScriptTip 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "脚本缺失提醒"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7965
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProcScriptTip.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdContinue 
      Caption         =   "继续(&O)"
      Height          =   360
      Left            =   5880
      TabIndex        =   5
      Top             =   4440
      Width           =   990
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   5160
      Top             =   2160
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
            Picture         =   "frmProcScriptTip.frx":6852
            Key             =   "warning"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   360
      Left            =   6960
      TabIndex        =   4
      Top             =   4440
      Width           =   990
   End
   Begin VB.CommandButton cmdExp 
      Caption         =   "导出脚本列表(&L)"
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   4440
      Width           =   1575
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfScript 
      Height          =   3735
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   7935
      _cx             =   13996
      _cy             =   6588
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
      BackColorSel    =   -2147483626
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   350
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
   Begin VB.PictureBox pctTip 
      BackColor       =   &H00F8F0E9&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   7935
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7935
      Begin VB.Label lblTip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "为了检查结果的准确性，请执行等于或高于以下版本的SP安装包"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   5040
      End
   End
End
Attribute VB_Name = "frmProcScriptTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrTip As String
Private mblnReturn As Boolean

Private Sub InsertInfo(ByVal strInfo As String)
    With vsfScript
        .Rows = .Rows + 1
        
        .Cell(flexcpPicture, .Rows - 1, 0) = imgList.ListImages("warning").Picture
        .Cell(flexcpText, .Rows - 1, 0) = strInfo
    End With
End Sub

Public Function ShowMe(ByVal strTip As String) As Boolean
    mstrTip = strTip
    Me.Show 1
    
    ShowMe = mblnReturn
End Function


Private Sub cmdContinue_Click()
    mblnReturn = True
    Unload Me
End Sub

Private Sub cmdExp_Click()
    On Error GoTo errH:
    
    If Not IsInstallExcel() Then
        MsgBox "本机未安装Excel。", vbExclamation, Me.Caption
        Exit Sub
    End If
    

    vsfScript.SaveGrid App.Path & "\脚本缺失清单.xls", flexFileExcel, True
    MsgBox "保存成功，已经保存至" & VB.App.Path & "\脚本缺失清单.xls"

    Exit Sub
errH:
    MsgBox "脚本缺失清单导出失败." & vbNewLine & err.Description
End Sub

Private Sub cmdCancel_Click()
    mblnReturn = False
    Unload Me
End Sub

Private Sub Form_Load()
    Dim arrTmp() As String, i As Long
    
    arrTmp = Split(mstrTip, vbNewLine)
    For i = 0 To UBound(arrTmp)
        InsertInfo arrTmp(i)
    Next
End Sub

