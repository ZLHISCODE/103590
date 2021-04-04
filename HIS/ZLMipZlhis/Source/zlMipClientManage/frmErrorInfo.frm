VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form frmErrorInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "安装错误提示"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8145
   Icon            =   "frmErrorInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   8145
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdExport 
      Caption         =   "导出(&E)"
      Height          =   345
      Left            =   90
      TabIndex        =   7
      Top             =   4410
      Width           =   1100
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   9435
      TabIndex        =   5
      Top             =   0
      Width           =   9435
      Begin VB.Image Image1 
         Height          =   720
         Left            =   240
         Picture         =   "frmErrorInfo.frx":6852
         Top             =   45
         Width           =   720
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmErrorInfo.frx":C454
         Height          =   585
         Left            =   1185
         TabIndex        =   6
         Top             =   165
         Width           =   6720
      End
   End
   Begin VB.CommandButton cmdComm 
      Caption         =   "跳过(&C)"
      Height          =   345
      Left            =   6405
      TabIndex        =   3
      Top             =   4980
      Width           =   1100
   End
   Begin VB.CommandButton cmdReply 
      Caption         =   "重试(&R)"
      Height          =   345
      Left            =   5205
      TabIndex        =   2
      Top             =   4980
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -285
      TabIndex        =   1
      Top             =   4785
      Width           =   8385
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   3420
      Left            =   105
      TabIndex        =   4
      Top             =   930
      Width           =   7965
      _cx             =   2088842977
      _cy             =   2088834960
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
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483626
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
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   330
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
   Begin VB.Label Label1 
      Caption         =   "安装过程中遇到以下错误："
      Height          =   225
      Left            =   1275
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "frmErrorInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mclsVsf As New clsVsf
Private mblnResult As Boolean
Private mobjFso As New FileSystemObject

Public Function ShowError(ByVal objMain As Object, ByVal rsError As ADODB.Recordset) As Boolean
    On Error GoTo errHand
    
    If rsError.RecordCount = 0 Then Exit Function
    
    rsError.MoveFirst
    With mclsVsf
        Call .Initialize(Me.Controls, vsf, True, False)
        Call .ClearColumn
        Call .AppendColumn("序号", 600, flexAlignLeftCenter, flexDTString, "", , False)
        Call .AppendColumn("内容", 2400, flexAlignLeftCenter, flexDTString, "", , True)
        
        .ExtendLastCol = True
        .AppendRows = True
    End With
    
    Call mclsVsf.LoadGrid(rsError)
    Me.Show 1, objMain
    ShowError = mblnResult
    Exit Function
errHand:
    MsgBox Err.Description, vbInformation + vbOKOnly, "信息提示"
End Function

Private Sub cmdComm_Click()
    Dim var As Variant
    var = MsgBox("部分错误请手工在消息服务平台处理，是否先导出错误日志再跳过此步骤？", vbQuestion + vbYesNoCancel, "提示信息")
    If var = vbYes Then
        Call cmdExport_Click
    ElseIf var = vbCancel Then
        Exit Sub
    End If
    mblnResult = True
    Unload Me
End Sub

Private Sub cmdExport_Click()
    Dim objFile As TextStream
    Dim i As Integer
    Set objFile = mobjFso.OpenTextFile(App.Path & "\" & "ImportDataLog.log", ForWriting, True)
    With vsf
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                objFile.WriteLine .TextMatrix(i, .ColIndex("序号")) & "--" & .TextMatrix(i, .ColIndex("内容"))
                objFile.WriteLine "**************************************************************************************************************"
            Next
        End If
    End With
    objFile.Close
    MsgBox "导出成功，请查阅:" & App.Path & "\" & "ImportDataLog.log"
End Sub

Private Sub cmdReply_Click()
    mblnResult = False
    Unload Me
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow Then
        Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    End If
End Sub
