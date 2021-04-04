VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmClientUpgradeFileUploadChoose 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "上传文件范围选择"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10395
   Icon            =   "frmClientUpgradeFileUploadChoose.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   10395
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   360
      Left            =   9240
      TabIndex        =   16
      Top             =   6120
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   360
      Left            =   8040
      TabIndex        =   15
      Top             =   6120
      Width           =   990
   End
   Begin VB.OptionButton optMode 
      Caption         =   "对指定文件升级(&2)"
      Height          =   180
      Index           =   1
      Left            =   3960
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin VB.OptionButton optMode 
      Caption         =   "对指定系统的文件升级(&1)"
      Height          =   180
      Index           =   0
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Value           =   -1  'True
      Width           =   2415
   End
   Begin VB.PictureBox picFiles 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      ScaleHeight     =   1545
      ScaleWidth      =   10065
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2280
      Width           =   10095
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除(&D)"
         Height          =   300
         Left            =   4560
         TabIndex        =   17
         Top             =   105
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "添加(&A)"
         Height          =   300
         Left            =   3600
         TabIndex        =   13
         Top             =   105
         Width           =   855
      End
      Begin VB.TextBox txtFile 
         Height          =   300
         Left            =   840
         TabIndex        =   12
         Top             =   120
         Width           =   2655
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfFiles 
         Height          =   855
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   1455
         _cx             =   2566
         _cy             =   1508
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
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
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmClientUpgradeFileUploadChoose.frx":0AE2
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
      Begin VB.Label lblFile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "文件(&F)"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   150
         Width           =   630
      End
   End
   Begin VB.PictureBox picSysFiles 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      ScaleHeight     =   1545
      ScaleWidth      =   10065
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   600
      Width           =   10095
      Begin VB.TextBox txtFind 
         Height          =   270
         Left            =   8280
         TabIndex        =   7
         Top             =   120
         Width           =   975
      End
      Begin VB.ComboBox cboSystem 
         Height          =   300
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfSysFiles 
         Height          =   855
         Left            =   2640
         TabIndex        =   9
         Top             =   480
         Width           =   1455
         _cx             =   2566
         _cy             =   1508
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
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
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmClientUpgradeFileUploadChoose.frx":0BB7
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
      Begin VSFlex8Ctl.VSFlexGrid vsfSys 
         Height          =   855
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1455
         _cx             =   2566
         _cy             =   1508
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
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
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmClientUpgradeFileUploadChoose.frx":0C8C
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
      Begin VB.Label lblSysFiles 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "文件列表"
         Height          =   180
         Left            =   2640
         TabIndex        =   18
         Top             =   150
         Width           =   720
      End
      Begin VB.Label lblFilter 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "文件定位(&I)"
         Height          =   180
         Left            =   7200
         TabIndex        =   6
         Top             =   150
         Width           =   990
      End
      Begin VB.Label lblSystem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "系统列表"
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   150
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   120
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientUpgradeFileUploadChoose.frx":0D61
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientUpgradeFileUploadChoose.frx":1229
            Key             =   "AllCheck"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblMode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "选择模式(&M)"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   255
      Width           =   990
   End
End
Attribute VB_Name = "frmClientUpgradeFileUploadChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MSTR_SYS As String = "选定,,3,300,B|编号,,0,0|系统,,3,100"
Private Const MSTR_SYS_READONLY As String = "编号|系统"
Private Const MSTR_SYSFILES As String = "选定,,3,300,B|序号,,3,600|所属系统,,0,0|文件,,3,2400|文件说明,,3,1000"
Private Const MSTR_SYSFILES_READONLY As String = "序号|所属系统|文件|文件说明"
Private Const MSTR_FILES As String = "序号,,3,600|文件,,3,1200"

Private WithEvents mobjSys As clsVSFlexGridEx
Attribute mobjSys.VB_VarHelpID = -1
Private WithEvents mobjFiles As clsVSFlexGridEx
Attribute mobjFiles.VB_VarHelpID = -1
Private WithEvents mobjSysFiles As clsVSFlexGridEx
Attribute mobjSysFiles.VB_VarHelpID = -1
Private mcolFiles As Collection

Public Function ShowMe(ByRef colFiles As Collection) As Boolean
    Show vbModal, frmMDIMain
    Set colFiles = mcolFiles
    ShowMe = val(Me.Tag) = 1
End Function

Private Sub Form_Load()
    Call InitVSF
    Call FillSysFiles
    Call FillSystem
    Call optMode_Click(0)
End Sub

Private Sub cmdAdd_Click()
    Dim i As Long
    Dim blnDo As Boolean
    
    If Trim(txtFile.Text) = "" Then Exit Sub
    
    '检查
    With vsfSysFiles
        blnDo = False
        For i = .FixedRows To .Rows - 1
            If UCase(Trim(.TextMatrix(i, .ColIndex("文件")))) = Trim(txtFile.Text) Then
                blnDo = True
                Exit For
            End If
        Next
        If blnDo = False Then
            MsgBox "文件（" & Trim(txtFile.Text) & "）不在升级文件清单中！" _
                    & vbCrLf & "请检查录入是否正确，或者先添加文件到升级文件清单中。" _
                , vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    
    With vsfFiles
        '检查
        For i = .FixedRows To .Rows - 1
            If UCase(Trim(.TextMatrix(i, .ColIndex("文件")))) = Trim(txtFile.Text) Then
                .Row = i
                Exit Sub
            End If
        Next
    
        '添加
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .TextMatrix(.Row, .ColIndex("序号")) = .Row
        .TextMatrix(.Row, .ColIndex("文件")) = Trim(txtFile.Text)
    End With
End Sub

Private Sub cmdCancel_Click()
    Me.Tag = ""
    Me.Hide
End Sub

Private Sub cmdDel_Click()
    Dim i As Long
    
    With vsfFiles
        i = .Row
        .RemoveItem i
        If i <= .Rows - 1 Then .Row = i
    End With
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    
    Set mcolFiles = New Collection
    
    On Error Resume Next
    If optMode(0).value Then
        With vsfSysFiles
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, .ColIndex("选定")) = flexChecked Then
                    mcolFiles.Add 0, Trim(.TextMatrix(i, .ColIndex("文件")))
                End If
            Next
        End With
    Else
        With vsfFiles
            If .Rows > 1 Then
                For i = .FixedRows To .Rows - 1
                    mcolFiles.Add 0, Trim(.TextMatrix(i, .ColIndex("文件")))
                Next
            End If
        End With
    End If
    On Error GoTo 0
    
    If mcolFiles.Count <= 0 Then
        MsgBox "请选定待上传的文件！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Me.Tag = "1"
    Me.Hide
End Sub

Private Sub InitVSF()
    Set mobjSys = New clsVSFlexGridEx
    With mobjSys
        .AppTemplate EM_Display, vsfSys, MSTR_SYS, MSTR_SYS_READONLY
        .Init True
        .Binding.Editable = flexEDKbdMouse
        .Binding.ScrollTrack = True
        .Binding.ExplorerBar = flexExNone
        .Binding.ExtendLastCol = True
        .Binding.AllowUserResizing = flexResizeNone
        .Binding.Cell(flexcpPicture, 0, .Binding.ColIndex("选定")) = img16.ListImages("UnCheck").Picture
    End With
    
    Set mobjSysFiles = New clsVSFlexGridEx
    With mobjSysFiles
        .AppTemplate EM_Display, vsfSysFiles, MSTR_SYSFILES, MSTR_SYSFILES_READONLY, True
        .Init True
        .Binding.Editable = flexEDKbdMouse
        .Binding.ScrollTrack = True
        .Binding.ExplorerBar = flexExNone
        .Binding.ExtendLastCol = True
        .Binding.Cell(flexcpPicture, 0, .Binding.ColIndex("选定")) = img16.ListImages("UnCheck").Picture
    End With
    
    Set mobjFiles = New clsVSFlexGridEx
    With mobjFiles
        .AppTemplate EM_Display, vsfFiles, MSTR_FILES, "", False
        .Init False
        .Binding.Editable = flexEDKbdMouse
        .Binding.ScrollTrack = True
        .Binding.ExplorerBar = flexExNone
        .Binding.ExtendLastCol = True
    End With
End Sub

Private Sub FillSysFiles()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo hErr
    
    strSQL = _
        "Select Nvl(a.文件名, b.名称) 文件 " & vbCr & _
        "  , Decode(a.所属系统, b.所属系统, a.所属系统, Nvl(a.所属系统, '') || ',' || Nvl(b.所属系统, '')) 所属系统 " & vbCr & _
        "  , Nvl(a.文件说明, b.文件说明) 文件说明 " & vbCr & _
        "From zlFilesUpgrade A Full Join Zlfiles B On a.文件名 = b.名称 " & vbCr & _
        "Order By 文件"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取待上传文件服务器的文件信息")
    mobjSysFiles.Recordset = rsTemp
    mobjSysFiles.Repaint RT_Rows
    rsTemp.Close
    
    With vsfSysFiles
        .Redraw = flexRDNone
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("序号")) = CLng(i)
            .RowData(i) = 1
        Next
        If .Rows > 1 Then
            .Row = 1
            i = .ColIndex("选定")
            .Cell(flexcpPicture, 0, i) = img16.ListImages("AllCheck").Picture
        End If
        .Redraw = flexRDDirect
    End With
    
    Exit Sub
    
hErr:
    MsgBox err.Number & "：" & err.Description, vbInformation, gstrSysName
End Sub

Private Sub FillSystem()
    Dim strSQL As String, strSys As String
    Dim rsTemp As ADODB.Recordset
    Dim arrSysNo As Variant, arrTemp As Variant
    Dim i As Long, lngCheck As Long
    Dim blnNoSysNo As Boolean

    On Error GoTo hErr

    '获取系统编号信息
    arrSysNo = Array()
    strSQL = _
        "Select Distinct 系统 " & vbCr & _
        "From (" & vbCr & _
        "  Select 所属系统 系统 From zlFilesUpgrade " & vbCr & _
        "  Union " & vbCr & _
        "  Select 所属系统 系统 From zlFiles " & vbCr & _
        ")"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "文件所属系统号")
    With rsTemp
        Do While .EOF = False
            If Trim("" & !系统) = "" Then
                '无系统...
                If blnNoSysNo = False Then blnNoSysNo = True
            ElseIf !系统 Like "*,*" Then
                arrTemp = Split(!系统, ",")
                For i = LBound(arrTemp) To UBound(arrTemp)
                    Call AddSytemNo(arrSysNo, val(arrTemp(i)))
                Next
            Else
                Call AddSytemNo(arrSysNo, val(!系统))
            End If

            .MoveNext
        Loop
        .Close
    End With

    '系统编号对应系统名称
    strSys = Join(arrSysNo, ",")
    strSQL = "Select * From(" & vbCr & _
        "Select /*+ cardinality(B, 10)*/ 系统, 编号 " & vbCr & _
        "From (" & vbCr & _
        "    Select 名称 系统, 编号 / 100 编号" & vbCr & _
        "    From zlSystems" & vbCr & _
        "    Union" & vbCr & _
        "    Select '管理工具', 0 From Dual" & vbCr & _
        ") A, Table(f_Str2List([1])) B " & vbCr & _
        "Where a.编号 = b.Column_Value " & vbCr & _
        IIf(blnNoSysNo, "Union Select '无', -1 From Dual ", "") & vbCr & _
        ") Order By 编号"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "文件所属系统信息", strSys)

    '加载数据
    mobjSys.Recordset = rsTemp
    mobjSys.Repaint RT_Rows
    rsTemp.Close
    
    With vsfSys
        .Redraw = flexRDNone
        If .Rows > 1 Then
            .Row = 1
            lngCheck = .ColIndex("选定")
            .Cell(flexcpPicture, 0, lngCheck) = img16.ListImages("AllCheck").Picture
            For i = .FixedRows To .Rows - 1
                .Cell(flexcpChecked, i, lngCheck) = flexChecked
                Call vsfSys_AfterEdit(i, lngCheck)
            Next
        End If
        .Redraw = flexRDDirect
    End With
    
    Exit Sub
    
hErr:
    MsgBox err.Number & "：" & err.Description, vbInformation, gstrSysName
End Sub

Private Sub AddSytemNo(ByRef arrVal As Variant, ByVal lngSysNo As Long)
    Dim i As Integer
    Dim blnDoAdd As Boolean
    
    If UBound(arrVal) < 0 Then
        blnDoAdd = True
    Else
        For i = LBound(arrVal) To UBound(arrVal)
            If arrVal(i) = lngSysNo Then
                Exit For
            End If
        Next
        blnDoAdd = i > UBound(arrVal)
    End If
    
    If blnDoAdd Then
        ReDim Preserve arrVal(UBound(arrVal) + 1)
        arrVal(UBound(arrVal)) = lngSysNo
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 120
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 120
    With picSysFiles
        .Width = cmdCancel.Width + cmdCancel.Left - 120
        .Height = cmdCancel.Top - .Top - 60
    End With
    picFiles.Move picSysFiles.Left, picSysFiles.Top, picSysFiles.Width, picSysFiles.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjFiles = Nothing
    Set mobjSysFiles = Nothing
End Sub

Private Sub optMode_Click(Index As Integer)
    picSysFiles.Visible = optMode(0).value
    picFiles.Visible = Not optMode(0).value
End Sub

Private Sub picFiles_Resize()
    On Error Resume Next
    cmdDel.Left = picFiles.ScaleWidth - cmdDel.Width - 120
    cmdAdd.Left = cmdDel.Left - cmdAdd.Width - 120
    txtFile.Width = cmdAdd.Left - txtFile.Left - 120
    vsfFiles.Move 120, _
        txtFile.Top + txtFile.Height + 60, _
        picFiles.ScaleWidth - vsfFiles.Left * 2, _
        picSysFiles.ScaleHeight - vsfFiles.Top - 120
End Sub

Private Sub picSysFiles_Resize()
    On Error Resume Next
    txtFind.Width = picSysFiles.ScaleWidth - txtFind.Left - 120
    vsfSys.Move 120, txtFind.Top + txtFind.Height + 60 _
        , lblSysFiles.Left - 60 - vsfSys.Left, picSysFiles.ScaleHeight - vsfSys.Top - 120
    vsfSysFiles.Move lblSysFiles.Left, _
        vsfSys.Top, picSysFiles.ScaleWidth - lblSysFiles.Left - 120, vsfSys.Height
End Sub

Private Sub txtFile_GotFocus()
    With txtFile
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtFile_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
        KeyAscii = KeyAscii - 32
    ElseIf KeyAscii = vbKeyReturn Then
        Call cmdAdd_Click
    End If
End Sub

Private Sub txtFind_Change()
    txtFind.Tag = ""
End Sub

Private Sub txtFind_GotFocus()
    With txtFind
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim i As Long, lngFile As Long, lngStart As Long
    
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyReturn Then
        With vsfSysFiles
            lngFile = .ColIndex("文件")
            lngStart = val(txtFind.Tag) + 1
            If lngStart < .FixedRows Then lngStart = .FixedRows
            For i = lngStart To .Rows - 1
                If .RowHidden(i) = False And InStr(UCase(.TextMatrix(i, lngFile)), UCase(Trim(txtFind.Text))) > 0 Then
                    If i - (.BottomRow - .TopRow) \ 2 > 0 Then
                        .TopRow = i - (.BottomRow - .TopRow) \ 2
                    Else
                        .TopRow = 1
                    End If
                    .Row = i
                    .Col = lngFile
                    txtFind.Tag = i
                    Exit For
                End If
            Next
            If i > .Rows - 1 Then
                If txtFind.Tag <> "" Then
                    txtFind.Tag = ""
                    If MsgBox("已查找到底部，需要从头开始查找吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        Call txtFind_KeyPress(vbKeyReturn)
                    End If
                Else
                    MsgBox "未找到配置文件！", vbInformation, gstrSysName
                End If
            End If
        End With
    End If
End Sub

Private Sub vsfFiles_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    If OldRowSel <> NewRowSel Then
        cmdDel.Enabled = vsfFiles.Rows > 1
    End If
End Sub

Private Function MatchingSystem(ByVal arrSys As Variant, ByVal strSysText As String) As Boolean
    Dim i As Long
    
    If strSysText = "" Then strSysText = "-1"
    For i = LBound(arrSys) To UBound(arrSys)
        If "," & Trim(strSysText) & "," Like "*," & val(arrSys(i)) & ",*" Then
            MatchingSystem = True
            Exit For
        End If
    Next
End Function

Private Sub vsfSys_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long, lngCheck As Long, lngSN As Long, lngFile As Long
    Dim blnAllChecked As Boolean
    Dim arrSys As Variant
    
    If Row <= 0 Then Exit Sub
    
    lngCheck = vsfSys.ColIndex("选定")
    
    With vsfSys
        blnAllChecked = True
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, lngCheck) <> flexChecked Then
                blnAllChecked = False
                Exit For
            End If
        Next
        If blnAllChecked = False Then
            .Cell(flexcpPicture, 0, lngCheck) = img16.ListImages("UnCheck").Picture
        Else
            .Cell(flexcpPicture, 0, lngCheck) = img16.ListImages("AllCheck").Picture
        End If
        
        arrSys = Array()
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, lngCheck) = flexChecked Then
                '收集勾选的系统编号
                ReDim Preserve arrSys(UBound(arrSys) + 1)
                arrSys(UBound(arrSys)) = .TextMatrix(i, .ColIndex("编号"))
            End If
        Next
    End With
        
    '更新文件列表
    With vsfSysFiles
        lngSN = 0
        lngCheck = .ColIndex("选定")
        lngFile = .ColIndex("文件")
        
        .Redraw = flexRDNone
        .ColSort(lngFile) = flexSortGenericAscending
        .Select .FixedRows, lngFile, .Rows - 1, lngFile
        .Sort = flexSortGenericAscending
        
        For i = .FixedRows To .Rows - 1
            .RowHidden(i) = Not MatchingSystem(arrSys, .TextMatrix(i, .ColIndex("所属系统")))
            .Cell(flexcpChecked, i, lngCheck) = IIf(.RowHidden(i), 0, 1)
            If .RowHidden(i) = False Then
                lngSN = lngSN + 1
                 .TextMatrix(i, .ColIndex("序号")) = CLng(lngSN)
                If lngSN = 1 Then
                    .Row = i
                    .TopRow = i
                End If
            End If
        Next
         .Redraw = flexRDDirect
    End With
End Sub

Private Sub vsfSys_Click()
    Dim lngCheck As Long, i As Long
    Dim blnChecked As Boolean
    
    '点击“选定”列头
    With vsfSys
        lngCheck = .ColIndex("选定")
        If Not (.MouseRow = 0 And .MouseCol = lngCheck) Then Exit Sub
        If .Rows < .FixedRows Then Exit Sub
        
        If .Cell(flexcpPicture, 0, lngCheck) = img16.ListImages("AllCheck").Picture Then
            .Cell(flexcpPicture, 0, lngCheck) = img16.ListImages("UnCheck").Picture
            blnChecked = False
        Else
            .Cell(flexcpPicture, 0, lngCheck) = img16.ListImages("AllCheck").Picture
            blnChecked = True
        End If
        
        '全选系统
        For i = .FixedRows To .Rows - 1
            .Cell(flexcpChecked, i, lngCheck) = IIf(blnChecked, flexChecked, flexNoCheckbox)
            Call vsfSys_AfterEdit(i, lngCheck)
        Next
        
        lngCheck = vsfSysFiles.ColIndex("选定")
        If blnChecked Then
            vsfSysFiles.Cell(flexcpPicture, 0, lngCheck) = img16.ListImages("AllCheck").Picture
        Else
            vsfSysFiles.Cell(flexcpPicture, 0, lngCheck) = img16.ListImages("UnCheck").Picture
        End If
    End With
End Sub

Private Sub vsfSysFiles_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long, lngCheck As Long
    Dim blnAllChecked As Boolean
    
    With vsfSysFiles
        lngCheck = .ColIndex("选定")
        blnAllChecked = True
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, lngCheck) <> flexChecked And .RowHidden(i) = False Then
                blnAllChecked = False
                Exit For
            End If
        Next
        If blnAllChecked = False Then
            .Cell(flexcpPicture, 0, .ColIndex("选定")) = img16.ListImages("UnCheck").Picture
        Else
            .Cell(flexcpPicture, 0, .ColIndex("选定")) = img16.ListImages("AllCheck").Picture
        End If
    End With
End Sub

Private Sub vsfSysFiles_BeforeSort(ByVal Col As Long, Order As Integer)
    Order = flexSortNone
End Sub

Private Sub vsfSysFiles_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col = vsfSysFiles.ColIndex("选定")
End Sub

Private Sub vsfSysFiles_Click()
    Dim lngCheck As Long, i As Long, j As Long
    Dim blnChecked As Boolean
    Dim arrSys As Variant, arrTemp As Variant
    
    With vsfSysFiles
        lngCheck = .ColIndex("选定")
        If .MouseRow = 0 And .MouseCol = lngCheck Then
            If .Rows < .FixedRows Then Exit Sub
            
            If .Cell(flexcpPicture, 0, lngCheck) = img16.ListImages("AllCheck").Picture Then
                .Cell(flexcpPicture, 0, lngCheck) = img16.ListImages("UnCheck").Picture
                blnChecked = False
            Else
                .Cell(flexcpPicture, 0, lngCheck) = img16.ListImages("AllCheck").Picture
                blnChecked = True
            End If
            
            arrSys = Array()
            For i = .FixedRows To .Rows - 1
                '更新
                .Cell(flexcpChecked, i, lngCheck) = blnChecked
                '收集文件对应的系统信息
                arrTemp = Split(.TextMatrix(i, .ColIndex("所属系统")), ",")
                For j = LBound(arrTemp) To UBound(arrTemp)
                    If Trim(arrTemp(j)) = "" Then
                        Call AddSytemNo(arrSys, -1)
                    Else
                        Call AddSytemNo(arrSys, val(arrTemp(j)))
                    End If
                Next
            Next
        End If
    End With
End Sub

Private Sub vsfSysFiles_KeyPress(KeyAscii As Integer)
    Dim blnVal As Boolean
    Dim i As Long, lngCheck As Long
    
    If KeyAscii = vbKeySpace Then
        With vsfSysFiles
            If .SelectedRows <= 0 Then Exit Sub
            
            lngCheck = .ColIndex("选定")
            blnVal = .Cell(flexcpChecked, .SelectedRow(0), lngCheck) = flexChecked
            For i = 0 To .SelectedRows - 1
                .Cell(flexcpChecked, .SelectedRow(i), lngCheck) = IIf(blnVal, flexNoCheckbox, flexChecked)
            Next
        End With
    End If
End Sub
