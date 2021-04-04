VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogClear 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "日志文件清理"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4545
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogClear.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdClear 
      Caption         =   "删除(&D)"
      Default         =   -1  'True
      Height          =   360
      Left            =   2280
      TabIndex        =   6
      Top             =   4200
      Width           =   990
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   360
      Left            =   3360
      TabIndex        =   5
      Top             =   4200
      Width           =   990
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfFile 
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   4215
      _cx             =   7435
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
      ForeColorSel    =   -2147483641
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483639
      GridColorFixed  =   -2147483639
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   3
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
      Editable        =   2
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
   Begin VB.CommandButton cmdPath 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   350
      Left            =   4035
      TabIndex        =   2
      Top             =   165
      Width           =   375
   End
   Begin VB.TextBox txtPath 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   310
      Left            =   900
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   195
      Width           =   3135
   End
   Begin MSComctlLib.ImageList imgIcon 
      Left            =   120
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogClear.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "文件列表"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   720
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "文件目录"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmLogClear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowClearLog()
    Me.Show 1
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Dim i As Long, strFile As String
    Dim blnSelected As Boolean, blnToday As Boolean
    
    On Error Resume Next

    ShowFlash "正在删除日志文件..."
    
    With vsfFile
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, 1) = flexChecked Then  '当天的日志文件,不需要删除
                blnSelected = True
                If Not .TextMatrix(i, .ColIndex("名称")) Like "*" & Replace(Date, "/", "") & "*" Then
                    gobjFile.DeleteFile txtPath.Text & "\" & .TextMatrix(i, .ColIndex("名称"))
                Else
                    blnToday = True
                End If
            End If
        Next
    End With
    
    ShowFlash ""
    
    If blnSelected = False Then
        MsgBox "请勾选文件后再进行清理！", , "提示"
        Exit Sub
    End If
        
    
    MsgBox "日志清理完成！" & IIf(blnToday, "当天生成的日志文件无法清理，请手动清理！", "")
    Unload Me
End Sub

Private Sub cmdPath_Click()
    Dim strTmp As String
    
    strTmp = OpenFolder(Me, "请选择日志路径", txtPath.Text)
    If strTmp = "" Then
        Exit Sub
    End If
    
    If Not gobjFile.FolderExists(strTmp) Then Exit Sub
    txtPath.Text = strTmp
    
    LoadFile strTmp
End Sub

Private Sub LoadFile(ByVal strPath As String)
    '功能: 根据传入的路径加载log日志
     Dim objFolder As Folder, objFile As File
     Dim i As Integer

    Set objFolder = gobjFile.GetFolder(strPath)
    
    With vsfFile
        .Rows = 1: i = 1
        .Rows = objFolder.Files.Count + 1
        
        For Each objFile In objFolder.Files
            .TextMatrix(i, 0) = i
            .TextMatrix(i, .ColIndex("名称")) = objFile.Name
            .Cell(flexcpPicture, i, .ColIndex("名称")) = imgIcon.ListImages(1).Picture
            i = i + 1
        Next
        
        If .Rows > 1 Then .Select 1, 1
        
    End With
End Sub

Private Sub InitVsf()
    Dim strCols As String
    
    strCols = ",300,1;,280,4;名称,2000,1"
    InitTable vsfFile, strCols
    
    With vsfFile
        .ColDataType(1) = flexDTBoolean
        .Cell(flexcpChecked, 0, 1) = flexUnchecked
    End With
End Sub

Private Sub Form_Load()
    Call InitVsf
    txtPath.Text = GetLogPath
    LoadFile txtPath.Text
End Sub

Private Sub vsfFile_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then
        Cancel = True
    Else
        Cancel = False
    End If
End Sub


Private Sub vsfFile_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    
    If Row = 0 And Col = 1 Then
    
        With vsfFile
            If .Redraw = flexRDNone Then Exit Sub
            If .Rows = 1 Then Exit Sub
            
            .Cell(flexcpChecked, 1, 1, .Rows - 1, 1) = .Cell(flexcpChecked, 0, 1)
        End With
    End If
End Sub


Private Sub vsfFile_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim i As Integer

    With vsfFile
        If KeyCode = 32 Then   '按下空格,进行勾选
            For i = .FixedRows To .Rows - .FixedRows
                If .IsSelected(i) Then
                    .TextMatrix(i, 1) = IIf(.TextMatrix(i, 1) = "-1", 0, -1)
                End If
            Next
        End If
    End With
End Sub
