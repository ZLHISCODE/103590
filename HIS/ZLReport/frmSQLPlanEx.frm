VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSQLPlanEx 
   Caption         =   "查看执行计划"
   ClientHeight    =   8955
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15180
   Icon            =   "frmSQLPlanEx.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   15180
   StartUpPosition =   1  '所有者中心
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   13440
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RtbRemark 
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   10821
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      FileName        =   "D:\性能优化设计规范2014第六章.rtf"
      TextRTF         =   $"frmSQLPlanEx.frx":6852
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Height          =   720
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   1270
      ButtonWidth     =   1376
      ButtonHeight    =   1270
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "img灰色"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "复制"
            Key             =   "Copy"
            Description     =   "复制"
            Object.ToolTipText     =   "执行报表"
            Object.Tag             =   "执行"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "另存为···"
            Key             =   "Save"
            Description     =   "另存为"
            Object.ToolTipText     =   "另存为txt"
            Object.Tag             =   "另存为"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "刷新"
            Key             =   "Review"
            ImageKey        =   "View"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "退出"
            Key             =   "Quit"
            Description     =   "退出"
            Object.ToolTipText     =   "退出"
            Object.Tag             =   "退出"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img灰色 
      Left            =   11160
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSQLPlanEx.frx":11C06
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSQLPlanEx.frx":12300
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSQLPlanEx.frx":1251A
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSQLPlanEx.frx":12734
            Key             =   "View"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPlan 
      Height          =   7095
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   13020
      _cx             =   22966
      _cy             =   12515
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
      BackColorFixed  =   -2147483643
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   235
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSQLPlanEx.frx":1294E
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
      OutlineBar      =   4
      OutlineCol      =   1
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
Attribute VB_Name = "frmSQLPlanEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrSQLCheck As String
Private mblnPro As Boolean   '是否有性能问题
Private mintType As Integer  '0-查看执行计划，1-查看优化技巧
Private mstrDataName As String
Private mintConnect As Integer

Public Function ShowMe(frmParent As Object, ByVal intConnect As Integer, ByVal strSQLCheck As String, _
    Optional ByVal intType As Integer, Optional ByVal strDataName As String) As Boolean
    
    mstrSQLCheck = strSQLCheck
    mintType = intType
    mstrDataName = strDataName
    mintConnect = intConnect
    Me.Show 1, frmParent
    ShowMe = mblnPro
End Function

Private Sub Form_Activate()
    If Me.Visible And Val(Me.Tag) = Val("-1-异常") Then
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Long, StrPar As String, blnSuccess As Boolean
    
    If mintType = 0 Then
        mblnPro = CheckSQLPlan(mstrSQLCheck, vsPlan, mintConnect, blnSuccess)
        If blnSuccess = False Then
            Me.Tag = "-1"
        End If
        
        RtbRemark.Visible = False
        Me.Caption = "查看执行计划"
        tbrMain.Buttons("Review").Visible = True
    Else
        vsPlan.Visible = False
        Me.Caption = "查看优化技巧"
        tbrMain.Buttons("Review").Visible = False
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    vsPlan.Top = 0: vsPlan.Left = 0
    vsPlan.Width = Me.ScaleWidth - vsPlan.Left
    vsPlan.Height = Me.ScaleHeight - vsPlan.Top - 60
    
    RtbRemark.Top = 0: RtbRemark.Left = 0
    RtbRemark.Width = Me.ScaleWidth - vsPlan.Left
    RtbRemark.Height = Me.ScaleHeight - vsPlan.Top - 60
End Sub


Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim i As Long, strText As String
    Dim strFormat As String * 4
    Dim strSpace As String * 100
    
    Select Case Button.Key
    Case "Copy"
        If mintType = 1 Then
            Clipboard.Clear
            Call Clipboard.SetText(RtbRemark.Text)
        Else
            With vsPlan
                strSpace = " "
                For i = .FixedRows To .Rows - 1
                    strFormat = .TextMatrix(i, 0)
                    strText = strText & IIF(strText = "", "", vbCrLf) & strFormat & " " & Mid(strSpace, 100 - Val(.RowOutlineLevel(i))) & .TextMatrix(i, 1)
                Next
                If strText <> "" Then
                    Clipboard.Clear
                    Call Clipboard.SetText(strText)
                End If
            End With
        End If
    Case "Save"
        If mintType = 1 Then
            With CommonDialog
                .DialogTitle = "保存文件"
                .Filter = "RTF Files|*.rtf"
                .Flags = &H200000 + &H2000 + &H2 + &H800
                .InitDir = App.Path
                .FileName = "SQL查询优化"
                .CancelError = True
                On Error Resume Next
                .ShowSave
                If Err.Number = 0 Then
                    RtbRemark.SaveFile .FileName
                End If
            End With
        Else
            With CommonDialog
                .DialogTitle = "保存文件"
                .Filter = "RTF Files|*.RTF"
                .Flags = &H200000 + &H2000 + &H2 + &H800
                .InitDir = App.Path
                .FileName = mstrDataName
                .CancelError = True
                On Error Resume Next
                .ShowSave
            End With
            If Err.Number = 0 Then
                With vsPlan
                    strSpace = " "
                    strText = "--------------" & mstrDataName & "-执行计划" & "--------------"
                    strText = strText & vbCrLf & "行数  操作"
                    For i = .FixedRows To .Rows - 1
                        strFormat = .TextMatrix(i, 0)
                        strText = strText & vbCrLf & strFormat & " " & Mid(strSpace, 100 - Val(.RowOutlineLevel(i))) & .TextMatrix(i, 1)
                    Next
                    If strText <> "" Then
                        RtbRemark.Text = strText
                        RtbRemark.SelStart = 1: RtbRemark.SelLength = Len(RtbRemark.Text)
                        RtbRemark.SelFontSize = 5.5
                        RtbRemark.SaveFile CommonDialog.FileName
                    End If
                End With
            End If
        End If
    Case "Review"
        mblnPro = CheckSQLPlan(mstrSQLCheck, vsPlan)
    Case "Quit"
        Unload Me
    End Select
End Sub

Private Sub vsPlan_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    vsPlan.ForeColorSel = vsPlan.Cell(flexcpForeColor, NewRow, NewCol)
End Sub
