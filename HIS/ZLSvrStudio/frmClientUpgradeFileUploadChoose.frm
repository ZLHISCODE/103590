VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmClientUpgradeFileUploadChoose 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ϴ��ļ���Χѡ��"
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
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   360
      Left            =   9240
      TabIndex        =   16
      Top             =   6120
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   360
      Left            =   8040
      TabIndex        =   15
      Top             =   6120
      Width           =   990
   End
   Begin VB.OptionButton optMode 
      Caption         =   "��ָ���ļ�����(&2)"
      Height          =   180
      Index           =   1
      Left            =   3960
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin VB.OptionButton optMode 
      Caption         =   "��ָ��ϵͳ���ļ�����(&1)"
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
         Caption         =   "ɾ��(&D)"
         Height          =   300
         Left            =   4560
         TabIndex        =   17
         Top             =   105
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "���(&A)"
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
            Name            =   "����"
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
         Caption         =   "�ļ�(&F)"
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
            Name            =   "����"
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
            Name            =   "����"
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
         Caption         =   "�ļ��б�"
         Height          =   180
         Left            =   2640
         TabIndex        =   18
         Top             =   150
         Width           =   720
      End
      Begin VB.Label lblFilter 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ļ���λ(&I)"
         Height          =   180
         Left            =   7200
         TabIndex        =   6
         Top             =   150
         Width           =   990
      End
      Begin VB.Label lblSystem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ϵͳ�б�"
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
      Caption         =   "ѡ��ģʽ(&M)"
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

Private Const MSTR_SYS As String = "ѡ��,,3,300,B|���,,0,0|ϵͳ,,3,100"
Private Const MSTR_SYS_READONLY As String = "���|ϵͳ"
Private Const MSTR_SYSFILES As String = "ѡ��,,3,300,B|���,,3,600|����ϵͳ,,0,0|�ļ�,,3,2400|�ļ�˵��,,3,1000"
Private Const MSTR_SYSFILES_READONLY As String = "���|����ϵͳ|�ļ�|�ļ�˵��"
Private Const MSTR_FILES As String = "���,,3,600|�ļ�,,3,1200"

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
    
    '���
    With vsfSysFiles
        blnDo = False
        For i = .FixedRows To .Rows - 1
            If UCase(Trim(.TextMatrix(i, .ColIndex("�ļ�")))) = Trim(txtFile.Text) Then
                blnDo = True
                Exit For
            End If
        Next
        If blnDo = False Then
            MsgBox "�ļ���" & Trim(txtFile.Text) & "�����������ļ��嵥�У�" _
                    & vbCrLf & "����¼���Ƿ���ȷ������������ļ��������ļ��嵥�С�" _
                , vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    
    With vsfFiles
        '���
        For i = .FixedRows To .Rows - 1
            If UCase(Trim(.TextMatrix(i, .ColIndex("�ļ�")))) = Trim(txtFile.Text) Then
                .Row = i
                Exit Sub
            End If
        Next
    
        '���
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .TextMatrix(.Row, .ColIndex("���")) = .Row
        .TextMatrix(.Row, .ColIndex("�ļ�")) = Trim(txtFile.Text)
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
                If .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = flexChecked Then
                    mcolFiles.Add 0, Trim(.TextMatrix(i, .ColIndex("�ļ�")))
                End If
            Next
        End With
    Else
        With vsfFiles
            If .Rows > 1 Then
                For i = .FixedRows To .Rows - 1
                    mcolFiles.Add 0, Trim(.TextMatrix(i, .ColIndex("�ļ�")))
                Next
            End If
        End With
    End If
    On Error GoTo 0
    
    If mcolFiles.Count <= 0 Then
        MsgBox "��ѡ�����ϴ����ļ���", vbInformation, gstrSysName
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
        .Binding.Cell(flexcpPicture, 0, .Binding.ColIndex("ѡ��")) = img16.ListImages("UnCheck").Picture
    End With
    
    Set mobjSysFiles = New clsVSFlexGridEx
    With mobjSysFiles
        .AppTemplate EM_Display, vsfSysFiles, MSTR_SYSFILES, MSTR_SYSFILES_READONLY, True
        .Init True
        .Binding.Editable = flexEDKbdMouse
        .Binding.ScrollTrack = True
        .Binding.ExplorerBar = flexExNone
        .Binding.ExtendLastCol = True
        .Binding.Cell(flexcpPicture, 0, .Binding.ColIndex("ѡ��")) = img16.ListImages("UnCheck").Picture
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
        "Select Nvl(a.�ļ���, b.����) �ļ� " & vbCr & _
        "  , Decode(a.����ϵͳ, b.����ϵͳ, a.����ϵͳ, Nvl(a.����ϵͳ, '') || ',' || Nvl(b.����ϵͳ, '')) ����ϵͳ " & vbCr & _
        "  , Nvl(a.�ļ�˵��, b.�ļ�˵��) �ļ�˵�� " & vbCr & _
        "From zlFilesUpgrade A Full Join Zlfiles B On a.�ļ��� = b.���� " & vbCr & _
        "Order By �ļ�"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡ���ϴ��ļ����������ļ���Ϣ")
    mobjSysFiles.Recordset = rsTemp
    mobjSysFiles.Repaint RT_Rows
    rsTemp.Close
    
    With vsfSysFiles
        .Redraw = flexRDNone
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("���")) = CLng(i)
            .RowData(i) = 1
        Next
        If .Rows > 1 Then
            .Row = 1
            i = .ColIndex("ѡ��")
            .Cell(flexcpPicture, 0, i) = img16.ListImages("AllCheck").Picture
        End If
        .Redraw = flexRDDirect
    End With
    
    Exit Sub
    
hErr:
    MsgBox err.Number & "��" & err.Description, vbInformation, gstrSysName
End Sub

Private Sub FillSystem()
    Dim strSQL As String, strSys As String
    Dim rsTemp As ADODB.Recordset
    Dim arrSysNo As Variant, arrTemp As Variant
    Dim i As Long, lngCheck As Long
    Dim blnNoSysNo As Boolean

    On Error GoTo hErr

    '��ȡϵͳ�����Ϣ
    arrSysNo = Array()
    strSQL = _
        "Select Distinct ϵͳ " & vbCr & _
        "From (" & vbCr & _
        "  Select ����ϵͳ ϵͳ From zlFilesUpgrade " & vbCr & _
        "  Union " & vbCr & _
        "  Select ����ϵͳ ϵͳ From zlFiles " & vbCr & _
        ")"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "�ļ�����ϵͳ��")
    With rsTemp
        Do While .EOF = False
            If Trim("" & !ϵͳ) = "" Then
                '��ϵͳ...
                If blnNoSysNo = False Then blnNoSysNo = True
            ElseIf !ϵͳ Like "*,*" Then
                arrTemp = Split(!ϵͳ, ",")
                For i = LBound(arrTemp) To UBound(arrTemp)
                    Call AddSytemNo(arrSysNo, val(arrTemp(i)))
                Next
            Else
                Call AddSytemNo(arrSysNo, val(!ϵͳ))
            End If

            .MoveNext
        Loop
        .Close
    End With

    'ϵͳ��Ŷ�Ӧϵͳ����
    strSys = Join(arrSysNo, ",")
    strSQL = "Select * From(" & vbCr & _
        "Select /*+ cardinality(B, 10)*/ ϵͳ, ��� " & vbCr & _
        "From (" & vbCr & _
        "    Select ���� ϵͳ, ��� / 100 ���" & vbCr & _
        "    From zlSystems" & vbCr & _
        "    Union" & vbCr & _
        "    Select '������', 0 From Dual" & vbCr & _
        ") A, Table(f_Str2List([1])) B " & vbCr & _
        "Where a.��� = b.Column_Value " & vbCr & _
        IIf(blnNoSysNo, "Union Select '��', -1 From Dual ", "") & vbCr & _
        ") Order By ���"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "�ļ�����ϵͳ��Ϣ", strSys)

    '��������
    mobjSys.Recordset = rsTemp
    mobjSys.Repaint RT_Rows
    rsTemp.Close
    
    With vsfSys
        .Redraw = flexRDNone
        If .Rows > 1 Then
            .Row = 1
            lngCheck = .ColIndex("ѡ��")
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
    MsgBox err.Number & "��" & err.Description, vbInformation, gstrSysName
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
            lngFile = .ColIndex("�ļ�")
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
                    If MsgBox("�Ѳ��ҵ��ײ�����Ҫ��ͷ��ʼ������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        Call txtFind_KeyPress(vbKeyReturn)
                    End If
                Else
                    MsgBox "δ�ҵ������ļ���", vbInformation, gstrSysName
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
    
    lngCheck = vsfSys.ColIndex("ѡ��")
    
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
                '�ռ���ѡ��ϵͳ���
                ReDim Preserve arrSys(UBound(arrSys) + 1)
                arrSys(UBound(arrSys)) = .TextMatrix(i, .ColIndex("���"))
            End If
        Next
    End With
        
    '�����ļ��б�
    With vsfSysFiles
        lngSN = 0
        lngCheck = .ColIndex("ѡ��")
        lngFile = .ColIndex("�ļ�")
        
        .Redraw = flexRDNone
        .ColSort(lngFile) = flexSortGenericAscending
        .Select .FixedRows, lngFile, .Rows - 1, lngFile
        .Sort = flexSortGenericAscending
        
        For i = .FixedRows To .Rows - 1
            .RowHidden(i) = Not MatchingSystem(arrSys, .TextMatrix(i, .ColIndex("����ϵͳ")))
            .Cell(flexcpChecked, i, lngCheck) = IIf(.RowHidden(i), 0, 1)
            If .RowHidden(i) = False Then
                lngSN = lngSN + 1
                 .TextMatrix(i, .ColIndex("���")) = CLng(lngSN)
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
    
    '�����ѡ������ͷ
    With vsfSys
        lngCheck = .ColIndex("ѡ��")
        If Not (.MouseRow = 0 And .MouseCol = lngCheck) Then Exit Sub
        If .Rows < .FixedRows Then Exit Sub
        
        If .Cell(flexcpPicture, 0, lngCheck) = img16.ListImages("AllCheck").Picture Then
            .Cell(flexcpPicture, 0, lngCheck) = img16.ListImages("UnCheck").Picture
            blnChecked = False
        Else
            .Cell(flexcpPicture, 0, lngCheck) = img16.ListImages("AllCheck").Picture
            blnChecked = True
        End If
        
        'ȫѡϵͳ
        For i = .FixedRows To .Rows - 1
            .Cell(flexcpChecked, i, lngCheck) = IIf(blnChecked, flexChecked, flexNoCheckbox)
            Call vsfSys_AfterEdit(i, lngCheck)
        Next
        
        lngCheck = vsfSysFiles.ColIndex("ѡ��")
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
        lngCheck = .ColIndex("ѡ��")
        blnAllChecked = True
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, lngCheck) <> flexChecked And .RowHidden(i) = False Then
                blnAllChecked = False
                Exit For
            End If
        Next
        If blnAllChecked = False Then
            .Cell(flexcpPicture, 0, .ColIndex("ѡ��")) = img16.ListImages("UnCheck").Picture
        Else
            .Cell(flexcpPicture, 0, .ColIndex("ѡ��")) = img16.ListImages("AllCheck").Picture
        End If
    End With
End Sub

Private Sub vsfSysFiles_BeforeSort(ByVal Col As Long, Order As Integer)
    Order = flexSortNone
End Sub

Private Sub vsfSysFiles_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col = vsfSysFiles.ColIndex("ѡ��")
End Sub

Private Sub vsfSysFiles_Click()
    Dim lngCheck As Long, i As Long, j As Long
    Dim blnChecked As Boolean
    Dim arrSys As Variant, arrTemp As Variant
    
    With vsfSysFiles
        lngCheck = .ColIndex("ѡ��")
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
                '����
                .Cell(flexcpChecked, i, lngCheck) = blnChecked
                '�ռ��ļ���Ӧ��ϵͳ��Ϣ
                arrTemp = Split(.TextMatrix(i, .ColIndex("����ϵͳ")), ",")
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
            
            lngCheck = .ColIndex("ѡ��")
            blnVal = .Cell(flexcpChecked, .SelectedRow(0), lngCheck) = flexChecked
            For i = 0 To .SelectedRows - 1
                .Cell(flexcpChecked, .SelectedRow(i), lngCheck) = IIf(blnVal, flexNoCheckbox, flexChecked)
            Next
        End With
    End If
End Sub
