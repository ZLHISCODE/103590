VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmReused 
   BorderStyle     =   0  'None
   ClientHeight    =   7860
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   17970
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   17970
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdGO 
      Caption         =   "��λ��LOB"
      Height          =   350
      Left            =   16560
      TabIndex        =   19
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox chkFree 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "ֻ��ʾ�տ�"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   13560
      TabIndex        =   17
      Top             =   65
      Width           =   1275
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   17970
      TabIndex        =   6
      Top             =   7245
      Width           =   17970
      Begin VB.CommandButton cmdMore 
         Caption         =   "����(&4)"
         Height          =   350
         Left            =   15618
         TabIndex        =   18
         Top             =   120
         Width           =   1095
      End
      Begin VB.CheckBox chkOnline 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFF0E0&
         Caption         =   "������������"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10680
         TabIndex        =   16
         Top             =   180
         Value           =   1  'Checked
         Width           =   1400
      End
      Begin VB.TextBox txtParallel 
         Alignment       =   1  'Right Justify
         Height          =   280
         Left            =   10245
         TabIndex        =   15
         Text            =   "12"
         ToolTipText     =   "�����˲��жȺ�����(Move)����ʱ���׵����·���Ŀռ�λ�������ļ���β�����Ӷ������޷����������ļ���"
         Top             =   160
         Width           =   375
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "����(&S)"
         Height          =   350
         Left            =   2745
         TabIndex        =   13
         Top             =   117
         Width           =   1095
      End
      Begin VB.TextBox txtFind 
         Height          =   350
         Left            =   1500
         TabIndex        =   12
         Top             =   117
         Width           =   1200
      End
      Begin VB.CommandButton cmdShrink 
         Caption         =   "����(&2)"
         Height          =   350
         Left            =   13246
         TabIndex        =   7
         ToolTipText     =   "һ�����ڴ���ɾ�����ݺ󽵵͸�ˮ������ջؿռ�"
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "����(&1)"
         Height          =   350
         Left            =   12060
         TabIndex        =   8
         ToolTipText     =   "�����ƶ��������λ���Ա������ļ�"
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdResize 
         Caption         =   "����(&3)"
         Height          =   350
         Left            =   14432
         TabIndex        =   9
         ToolTipText     =   "������ǰ��ռ��е�ǰ�����ļ��Ĵ�С"
         Top             =   120
         Width           =   1100
      End
      Begin VB.Label lblParallel 
         BackStyle       =   0  'Transparent
         Caption         =   "�������ж�"
         Height          =   255
         Left            =   9315
         TabIndex        =   14
         Top             =   210
         Width           =   930
      End
      Begin VB.Label lblFind 
         BackColor       =   &H00EFF0E0&
         Caption         =   "�����������(&F)"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   195
         Width           =   1455
      End
      Begin VB.Label lblOptPrompt 
         AutoSize        =   -1  'True
         BackColor       =   &H00EFF0E0&
         ForeColor       =   &H00400000&
         Height          =   180
         Left            =   3945
         TabIndex        =   10
         Top             =   195
         Width           =   90
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfExtents 
      Height          =   6375
      Left            =   3600
      TabIndex        =   4
      Top             =   840
      Width           =   14175
      _cx             =   25003
      _cy             =   11245
      Appearance      =   1
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
      ForeColorSel    =   12582912
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
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
      AutoResize      =   0   'False
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
      OwnerDraw       =   1
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
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfTbs 
      Height          =   6735
      Left            =   60
      TabIndex        =   3
      Top             =   480
      Width           =   3435
      _cx             =   6059
      _cy             =   11880
      Appearance      =   1
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   380
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
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   1
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
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.ComboBox cboFiles 
      Height          =   300
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   65
      Width           =   8895
   End
   Begin VB.Label lblPrompt 
      Caption         =   "��ǰѡ��Extent����Ϣ"
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   3600
      TabIndex        =   5
      Top             =   480
      Width           =   12855
   End
   Begin VB.Label lblFiles 
      Caption         =   "��ռ��ļ�"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblTableSpaces 
      Caption         =   "��ռ��б�"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu mnuResize 
      Caption         =   "����ѡ��"
      NegotiatePosition=   2  'Middle
      Visible         =   0   'False
      Begin VB.Menu mnuResizeAll 
         Caption         =   "����ȫ�������ļ�"
      End
      Begin VB.Menu mnuResizeTemp 
         Caption         =   "������ʱ�����ļ�"
      End
      Begin VB.Menu mnuResizeUndo 
         Caption         =   "����Undo��ռ�"
      End
      Begin VB.Menu mnuAddFile 
         Caption         =   "��������ļ�"
      End
   End
End
Attribute VB_Name = "frmReused"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CONCOLS As Long = 50
Private Const CONBLOCKS As Long = 8
Private mrsExtents As ADODB.Recordset
Private mrsLobs As ADODB.Recordset
Private mcolCells As Collection
Private mlngRowPre As Long, mlngColPre As Long

Private Enum opt
    P1���� = 1
    P2����
    P3����
End Enum

Public Sub ShowMe()
    Me.Show
End Sub

Private Sub cboFiles_Click()
    
    '����ѭ����ʹ����doevents������������κοɲ����Ĺ���
    Call SetCommandEnable(0)
    
    On Error GoTo errH
    
    Call LoadExtents(vsfTbs.TextMatrix(vsfTbs.Row, vsfTbs.ColIndex("����")), Val(cboFiles.ItemData(cboFiles.ListIndex)))
    cboFiles.ToolTipText = cboFiles.List(cboFiles.ListIndex)
    Call SetCommandEnable(1)
    
    If Me.Visible And Me.Enabled Then
        vsfExtents.SetFocus
    End If
    vsfExtents.Select vsfExtents.Rows - 1, vsfExtents.Cols - 1
    vsfExtents.TopRow = vsfExtents.Row
    
    Exit Sub
errH:
    ErrCenter
End Sub

Private Sub SetCommandEnable(bytEnable As Byte)
'���ܣ��������ť�Ŀ�����
    cmdShrink.Enabled = bytEnable = 1
    cmdMove.Enabled = cmdShrink.Enabled
    cmdResize.Enabled = cmdShrink.Enabled
    cmdMore.Enabled = cmdShrink.Enabled
    chkFree.Enabled = cmdShrink.Enabled
    chkOnline.Enabled = cmdShrink.Enabled
    cmdFind.Enabled = cmdShrink.Enabled
    txtFind.Enabled = cmdShrink.Enabled
    If txtParallel.Locked = False Then txtParallel.Enabled = cmdShrink.Enabled
    
    If cmdGO.Visible Then cmdGO.Enabled = cmdShrink.Enabled
    
    vsfTbs.Enabled = cmdShrink.Enabled
    cboFiles.Enabled = cmdShrink.Enabled
End Sub

Private Sub chkFree_Click()
    If cboFiles.ListIndex >= 0 Then Call cboFiles_Click
End Sub

Private Function ResizeTBS(ByVal strTBS As String, Optional ByVal lngFile As Long) As Boolean
'���ܣ�������ռ�
'������strTBS-��ռ�����
'      blnPrompt-�����ļ���,������ʱ���ڲ���ʾ�������������ǰ��ռ�����������ļ�����С�ߴ�
    Dim strSql As String, dblMax As Double, dblFileSize As Double, dblLimit As Double, dblBlockSize As Double
    Dim i As Long, blnTry As Boolean
    Dim rsTmp As ADODB.Recordset
           
    On Error GoTo errH
    
    dblBlockSize = Val(vsfTbs.RowData(vsfTbs.Row))
    If dblBlockSize = 0 Then dblBlockSize = 8192
        
    If lngFile <> 0 Then
        dblLimit = CDbl(1024) * 1024 * 2
        
        strSql = "Select a.File_Id, a.Last_Block, b.Bytes" & vbNewLine & _
            "From (Select a.File_Id, Max(a.Block_Id + a.Blocks - 1) Last_Block" & vbNewLine & _
            "       From Dba_Extents A" & vbNewLine & _
            "       Where a.Tablespace_Name = [1] And File_Id = [2]" & vbNewLine & _
            "       Group By a.File_Id) A, Dba_Data_Files B" & vbNewLine & _
            "Where a.File_Id = b.File_Id"
        Set rsTmp = OpenSQLRecord(strSql, Me.Caption, strTBS, lngFile)
        
        If rsTmp.RecordCount = 0 Then
            MsgBox "��Dba_Extents��û���ҵ���ǰ��ռ估�����ļ��ļ�¼", vbInformation, "����"
            Exit Function
        End If
    
        dblMax = rsTmp!Last_Block * dblBlockSize
        dblFileSize = rsTmp!Bytes
        If dblFileSize - dblMax < dblLimit Then 'С��2M��������
            If MsgBox("�������Ŀռ�(" & Round((dblFileSize - dblMax) / 1024) & "KB)С��2M,�Ƿ�ȷʵҪ�������ļ���", vbYesNo + vbDefaultButton2, "����") = vbNo Then
                Exit Function
            End If
            dblMax = Round(dblMax / 1024 / 1024) + 1 'ȡ����1����λM
        Else
            dblMax = Round(dblMax / 1024 / 1024) + 1 'ȡ����1����λM
            If MsgBox("��ȷ��Ҫ����ǰ�ļ�������" & dblMax & "M��?", vbQuestion + vbOKCancel + vbDefaultButton1, Me.Caption) = vbCancel Then
                Exit Function
            End If
        End If
        
        If dblMax >= Round(rsTmp!Bytes / 1024 / 1024) Then
            MsgBox "�����ļ��Ѵﵽ���ߴ磬�޷����ģ�", vbInformation
        Else
            Err.Clear
            On Error Resume Next
retry1:     strSql = "Alter Database Datafile " & lngFile & " Resize " & CStr(dblMax) & "M"
            gcnOracle.Execute strSql
            
            If Err.Number <> 0 Then
                If MsgBox("���������ļ�ʧ�ܣ�������ɾ�������δ��ջ���վ����ģ��Ƿ���պ����ԣ�", vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
                    Err.Clear
                    strSql = "purge tablespace " & strTBS
                    gcnOracle.Execute strSql
                    GoTo retry1
                Else
                    GoTo errH
                End If
            End If
            ResizeTBS = True
        End If
        
    Else
        dblLimit = CDbl(1024) * 1024 * 10
        
        '�������ռ�С��10M����������������ѭ����Ƶ��ִ������
        strSql = "Select a.File_Id, a.Last_Block * " & dblBlockSize & " as MaxBytes, b.Bytes" & vbNewLine & _
                "From (Select a.File_Id, Max(a.Block_Id + a.Blocks - 1) Last_Block" & vbNewLine & _
                "       From Dba_Extents A" & vbNewLine & _
                "       Where a.Tablespace_Name = [1]" & vbNewLine & _
                "       Group By a.File_Id) A, Dba_Data_Files B" & vbNewLine & _
                "Where a.File_Id = b.File_Id And (b.Bytes - a.Last_Block * " & dblBlockSize & ") > " & dblLimit
    
        Set rsTmp = OpenSQLRecord(strSql, Me.Caption, strTBS)
        
        On Error Resume Next
        For i = 1 To rsTmp.RecordCount
            dblMax = Round(rsTmp!MaxBytes / 1024 / 1024) + 1 'ȡ����1����λM
            If dblMax < Round(rsTmp!Bytes / 1024 / 1024) Then
                lblOptPrompt.Caption = "����" & rsTmp!File_Id & "�������ļ���" & CStr(dblMax) & "M"
                                
                blnTry = False
retry2:         strSql = "Alter Database Datafile " & rsTmp!File_Id & " Resize " & CStr(dblMax) & "M"
                gcnOracle.Execute strSql
                If Err.Number <> 0 And blnTry = False Then
                    Err.Clear
                    strSql = "purge tablespace " & strTBS
                    gcnOracle.Execute strSql
                    blnTry = True
                    GoTo retry2
                Else
                    Err.Clear   '����һ�κ�����
                End If
                
                ResizeTBS = True
            End If
            
            rsTmp.MoveNext
        Next
    End If
    
    Exit Function
errH:
    Call ErrCenter(strSql)
    Call SetCommandEnable(1)
End Function


Private Sub cmdMore_Click()
    Me.PopupMenu mnuResize
End Sub

Private Sub cmdResize_Click()
'���ܣ�ִ�б�ռ�����
    If cboFiles.ListIndex < 0 Then
        MsgBox "��ѡ��һ�������ļ���", vbInformation, "����"
        If cboFiles.Enabled Then cboFiles.SetFocus
    Else
        Call SetCommandEnable(0)
        
        If ResizeTBS(vsfTbs.TextMatrix(vsfTbs.Row, vsfTbs.ColIndex("����")), Val(cboFiles.ItemData(cboFiles.ListIndex))) Then
            lblOptPrompt.Caption = "������ļ�����������ˢ�¡�"
            lblOptPrompt.Refresh
            
            Call RefreshData
            
            lblOptPrompt.Caption = "����ɲ�����"
        End If
        
        Call SetCommandEnable(1)
    End If
End Sub

Private Sub RefreshData()
'���ܣ�ˢ�µ�ǰ��ռ�ĵ�ǰ�����ļ��Ķε�������Ϣ

    Dim i As Long, strTBS As String
    Dim lngFile As Long
    
    strTBS = vsfTbs.TextMatrix(vsfTbs.Row, vsfTbs.ColIndex("����"))
    lngFile = cboFiles.ListIndex
    Call LoadTablespaces
    
    vsfTbs.Redraw = flexRDNone
    i = vsfTbs.FindRow(strTBS, , vsfTbs.ColIndex("����"))
    If i <> -1 Then vsfTbs.Row = i: vsfTbs.TopRow = i
    vsfTbs.Redraw = flexRDDirect
    
    Call LoadFiles(vsfTbs.TextMatrix(vsfTbs.Row, vsfTbs.ColIndex("����")))
    If lngFile <= cboFiles.ListCount Then
        cboFiles.ListIndex = lngFile
    Else
        cboFiles.ListIndex = 0
    End If
End Sub

Private Function CheckUnSuportObject(strSegment As String, strOpt As String) As Boolean
'���ܣ����ָ���ı��Ƿ����Move��Shrink��֧�ֵĶ���
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    strSql = "Select 1" & vbNewLine & _
            "From All_Tab_Columns" & vbNewLine & _
            "Where Table_Name = [2] And Owner = [1] And Data_Type In ('LONG','LONG RAW','UNDEFINED')"
    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
    If rsTmp.RecordCount > 0 Then
        lblOptPrompt.Caption = strSegment & "����LONG,LONG RAW�����ֶΣ����ܽ���" & strOpt & "����."
    Else
        CheckUnSuportObject = True
    End If
End Function

Private Function CheckIOT(strSegment As String) As Boolean
'���ܣ����ָ���������Ƿ�Ϊ������֯�����������֧���ؽ���
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    strSql = "Select 1 From All_Indexes Where Owner = [1] And Index_Name = [2] And Index_Type = 'IOT - TOP'"
    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
    CheckIOT = rsTmp.RecordCount > 0
End Function

Private Function GetIOTName(strSegment As String) As String
'���ܣ�����������֯�������������������֯����(��������ǰ׺)
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    strSql = "Select Table_Owner||'.'||Table_Name as Tab_Name From All_Indexes Where Owner = [1] And Index_Name = [2] And Index_Type = 'IOT - TOP'"
    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
    If rsTmp.RecordCount > 0 Then
        GetIOTName = rsTmp!Tab_Name
    End If
End Function


Private Function CheckLOBIndex(strSegment As String) As Boolean
'���ܣ����ָ���������Ƿ�ΪLOB����������֧���ؽ���
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    strSql = "Select 1 From All_Lobs Where Owner = [1] And Index_Name = [2]"
    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
    CheckLOBIndex = rsTmp.RecordCount > 0
End Function

Private Function GetLOBNameByIndex(strSegment As String) As String
'���ܣ����ָ���������Ƿ�ΪLOB����������֧���ؽ���
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    strSql = "Select Segment_Name From All_Lobs Where Owner = [1] And Index_Name = [2]"
    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
    If rsTmp.RecordCount > 0 Then
        GetLOBNameByIndex = Split(strSegment, ".")(0) & "." & rsTmp!Segment_Name
    End If
End Function

Private Sub ReBuildIndex(ByVal strOwner As String, ByVal strTable As String, ByVal strParallel As String)
'���ܣ��ؽ�ĳ�ű���ʧЧ������
'������strOwner=������,strTable=����
'      strParallel=" Parallel X",���ж�
    Dim rsTmp As ADODB.Recordset, rsIndex As ADODB.Recordset
    Dim strSql As String
    
    lblOptPrompt.Caption = "�����ؽ�[" & strOwner & "." & strTable & "]��ʧЧ������"
    On Error GoTo errH
    
    '�ؽ�ʧЧ������
    strSql = "Select Index_Name From DBA_Indexes Where Status='UNUSABLE' And Owner = [1] And Table_Name = [2]"
    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, strOwner, strTable)
    
    Do While Not rsTmp.EOF
        '����Ƿ�����������Ҫ��������
        strSql = "Select Partition_Name From Dba_Ind_Partitions Where Index_Owner = [1] And Index_Name = [2]"
        Set rsIndex = OpenSQLRecord(strSql, Me.Caption, strOwner, rsTmp!Index_Name)
        If rsIndex.RecordCount > 0 Then
            Do While Not rsIndex.EOF
                strSql = "Alter Index " & strOwner & "." & rsTmp!Index_Name & " Rebuild Partition " & rsIndex!Partition_Name & " Nologging" & strParallel
                gcnOracle.Execute strSql
                rsIndex.MoveNext
            Loop
        Else
            strSql = "Alter Index " & strOwner & "." & rsTmp!Index_Name & " Rebuild Nologging" & strParallel
            gcnOracle.Execute strSql
        End If
        
        If strParallel <> "" Then
            strSql = "Alter Index " & strOwner & "." & rsTmp!Index_Name & " NOParallel"
            gcnOracle.Execute strSql
        End If
        
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    Call ErrCenter(strSql)
End Sub

Private Sub cmdMove_Click()
'���ܣ�ִ�б�������Ŀ��пռ�����(Move)
    Dim strSql As String, strType As String, strPartition As String, strSegment As String, strSegmentPre As String, strSegmentAll As String
    Dim strTBSTemp As String, strTbsOriginal As String, strTbsLob As String, strColumn As String, strParallel As String, strTableName As String
    Dim rsTmp As ADODB.Recordset
    Dim rsTbs As ADODB.Recordset
    Dim r1 As Long, c1 As Long, r2 As Long, c2 As Long, c As Long, r As Long
    Dim arrTmp As Variant, blnRemove As Boolean
    Dim strPrompt As String, strOnline As String, strObjName As String
    
    '���ջؿռ�Ķ������Ƶ���ʱ�洢�ı�ռ䣬�����ƻ���
    Dim strRemoveIndex As String, strRemoveLob As String, strRemoveTable As String
    Dim strRemovePARTable As String, strRemovePARIndex As String, strRemovePARLOB As String
    Dim datBegin As Date, strTime As String
    
    If CheckExtent(P2����) = False Then Exit Sub
    On Error GoTo errH
    strType = Trim(lblPrompt.Tag)
    If strType <> "" And InStr(",TABLE,TABLE PARTITION,INDEX,INDEX PARTITION,LOBSEGMENT,LOBINDEX,LOB PARTITION,", "," & strType & ",") = 0 Then '��LOBINDEX������������LOBSEGMENT
        Call MsgBox("��֧�ֶԱ���������п��пռ��ջأ���֧�ֵ��������ͣ�" & strType, vbInformation, Me.Caption)
        Exit Sub
    End If
    Call SetCommandEnable(0)
    
reInput:    strTBSTemp = Trim(InputBox("    Ϊ�˽�����λ�������ļ�ĩβ�ĵ�ǰ���Ƶ�ǰ�棬��Ҫ���˶����Ƶ�һ����ʱ��ŵı�ռ䣬���������ļ�֮�����ƻ�����" & vbCrLf & _
                    "    ���ѡ��ȡ������ť����ֱ���ڵ�ǰ��ռ������������(�����󣬵�ǰ�ο�����Ȼλ�������ļ���ĩβ)��", "��ʱ��ŵı�ռ�", "SYSAUX"))
    If strTBSTemp <> "" Then
        strTBSTemp = UCase(strTBSTemp)
        
        strSql = "Select 1 From DBA_TABLESPACES Where TABLESPACE_NAME = [1]"
        Set rsTbs = OpenSQLRecord(strSql, "��ռ���", strTBSTemp)
        If rsTbs.RecordCount = 0 Then
            MsgBox "����ı�ռ䲻���ڣ�����������", vbExclamation, "��ʾ"
            GoTo reInput
        End If
    End If
    strTbsOriginal = vsfTbs.TextMatrix(vsfTbs.Row, vsfTbs.ColIndex("����"))
    If strTBSTemp = strTbsOriginal Then strTBSTemp = ""
    
    If txtParallel.Text <> "0" Then strParallel = " Parallel " & txtParallel.Text
    If chkOnline.Value = 1 Then strOnline = "Online"
    
    Me.Refresh  '�������ϵĲ�Ӱ
    datBegin = GetCurrentdate
    
    
    '����һ��ѡ����ж��е����
    With vsfExtents
        .GetSelection r1, c1, r2, c2
                
        For r = r2 To r1 Step -1
            For c = c2 To c1 Step -1
                strSegment = mcolCells("_" & r & "_" & c)     '��������
                If strSegment <> strSegmentPre Then
                    If InStr(strSegmentAll & ",", "," & strSegment & ",") = 0 Then
                    
                        mrsExtents.Filter = "Row=" & r & " And Col=" & c
                        If mrsExtents.RecordCount > 0 Then
                            DoEvents
                            strType = mrsExtents!Segment_Type
                            '1.��ͨ��
                            If strType = "TABLE" Then
                                'mdsys�û��´�������GridFile1044_TAB���ֺ���Сд��ĸ�ı�Σ�����dba_tables��ȴ�鲻����¼
                                If CheckUnSuportTable(Split(strSegment, ".")(0), Split(strSegment, ".")(1)) Then
                                
                                    If CheckUnSuportObject(strSegment, "����(Move)") Then
                                        lblOptPrompt.Caption = "���ڶ�[" & strSegment & "]��������"
                                        lblOptPrompt.Refresh
                                        If strTBSTemp = "" Then
                                            strSql = "Alter Table " & strSegment & " Move Nologging" & strParallel
                                            gcnOracle.Execute strSql
                                            Call ReBuildIndex(Split(strSegment, ".")(0), Split(strSegment, ".")(1), strParallel)
                                        Else
                                            strSql = "Alter Table " & strSegment & " Move TableSpace " & strTBSTemp & " Nologging" & strParallel
                                            gcnOracle.Execute strSql
                                            strRemoveTable = strRemoveTable & "," & strSegment & "||" & strTbsOriginal
                                        End If
                                        
                                        If strParallel <> "" Then
                                            strSql = "Alter Table " & strSegment & " NOParallel"
                                            gcnOracle.Execute strSql
                                        End If
                                    Else
                                        If InStr(strPrompt, ":" & strSegment) = 0 Then strPrompt = strPrompt & vbCrLf & "����Long��Long Raw�ֶεı�:" & strSegment
                                    End If
                                Else
                                    If InStr(strPrompt, ":" & strSegment) = 0 Then strPrompt = strPrompt & vbCrLf & "IOT���������������Զ����ֶεı�:" & strSegment
                                End If
                            '2.������(����LOB������)
                            ElseIf strType = "TABLE PARTITION" Then
                                If CheckUnSuportObject(strSegment, "����(Move)") Then
                                    
                                    strSql = "Select Partition_Name From Dba_Tab_Partitions Where Table_Owner = [1] And Table_Name = [2]"
                                    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
                                    Do While Not rsTmp.EOF
                                        strPartition = rsTmp!Partition_Name
                                        
                                        lblOptPrompt.Caption = "���ڶ�[" & strSegment & "(" & strPartition & ")]��������"
                                        lblOptPrompt.Refresh
                                        
                                        'δ�Ӽ�����������update indexes���ں������ReBuildIndex���ָ�����Ϊ����������
                                        If strTBSTemp = "" Then
                                            strSql = "Alter Table " & strSegment & " Move Partition " & strPartition & " Nologging" & strParallel
                                            gcnOracle.Execute strSql
                                            
                                            Call ReBuildIndex(Split(strSegment, ".")(0), Split(strSegment, ".")(1), strParallel)
                                        Else
                                            strSql = "Alter Table " & strSegment & " Move Partition " & strPartition & " TableSpace " & strTBSTemp & " Nologging" & strParallel
                                            gcnOracle.Execute strSql
                                            strRemovePARTable = strRemovePARTable & "," & strSegment & "||" & strPartition & "||" & strTbsOriginal
                                        End If
                                        rsTmp.MoveNext
                                    Loop
                                    
                                    If strParallel <> "" Then
                                        strSql = "Alter Table " & strSegment & " NOParallel"
                                        gcnOracle.Execute strSql
                                    End If
                                Else
                                    If InStr(strPrompt, ":" & strSegment) = 0 Then strPrompt = strPrompt & vbCrLf & "����Long��Long Raw�ֶεķ�����:" & strSegment
                                End If
                                
                            '3.LOB�Σ�����LOB����������LOB������
                            ElseIf strType = "LOBSEGMENT" Or strType = "LOBINDEX" Then
                                If strType = "LOBINDEX" Then
                                    strSql = "Select Owner ||'.'|| Segment_Name as Segment_Name From Dba_Lobs Where Owner = [1] And Index_Name = [2]"
                                    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
                                    If rsTmp.RecordCount > 0 Then
                                        If InStr(strSegmentAll & ",", "," & rsTmp!Segment_Name & ",") = 0 Then
                                            strSegment = rsTmp!Segment_Name
                                        Else
                                            GoTo NextCell '���LOB����������������
                                        End If
                                    End If
                                End If
                            
                                mrsLobs.Filter = "Owner='" & Split(strSegment, ".")(0) & "' And Segment_Name='" & Split(strSegment, ".")(1) & "'"   'Ϊ��ȡ����������
                                If mrsLobs.RecordCount > 0 Then
                                    'mdsys�û��´�������GridFile1044_TAB���ֺ���Сд��ĸ�ı�Σ�����dba_tables��ȴ�鲻����¼
                                    If CheckUnSuportTable(mrsLobs!Owner, mrsLobs!Table_name) Then
                                        strTableName = mrsLobs!Owner & "." & mrsLobs!Table_name
                                        strTbsLob = mrsLobs!Tablespace_Name
                                        strColumn = mrsLobs!Column_Name
                                        
                                        lblOptPrompt.Caption = "���ڶ�[" & strTableName & "(" & strColumn & ")]��������"
                                        lblOptPrompt.Refresh
                                        If strTBSTemp = "" Then
                                            strSql = "ALTER TABLE " & strTableName & " Move LOB (" & strColumn & ") Store as(Tablespace " & strTbsLob & ") Nologging" & strParallel
                                            gcnOracle.Execute strSql
                                            
                                        Else
                                            strSql = "ALTER TABLE " & strTableName & " Move LOB (" & strColumn & ") Store as(Tablespace " & strTBSTemp & ") Nologging" & strParallel
                                            gcnOracle.Execute strSql
                                            strRemoveLob = strRemoveLob & "," & strTableName & "||" & strColumn & "||" & strTbsLob
                                        End If
                                                                        
                                        'LOB����ִ�в��ᵼ�±�������degree���Ա����ã����Բ���ִ��noparallel
                                    Else
                                        If InStr(strPrompt, ":" & mrsLobs!Table_name) = 0 Then strPrompt = strPrompt & vbCrLf & "δ֧�ֵı�:" & mrsLobs!Table_name
                                    End If
                                Else
                                    lblOptPrompt.Caption = "����ͼDba_Lobs��δ�ҵ�LOB����" & strSegment & "��"
                                End If
                                
                            '4.��ͨ����
                            ElseIf strType = "INDEX" Then
                                If CheckIOT(strSegment) = False Then    'IOT����ֻ��ͨ��moveԭ���ؽ�
                                    lblOptPrompt.Caption = "���ڶ�[" & strSegment & "]�����ؽ�"
                                    lblOptPrompt.Refresh
                                    If strTBSTemp = "" Then
                                        strSql = "Alter Index " & strSegment & " Rebuild " & strOnline & " Nologging" & strParallel
                                        gcnOracle.Execute strSql
                                    Else
                                        strSql = "Alter Index " & strSegment & " Rebuild " & strOnline & " TableSpace " & strTBSTemp & " Nologging" & strParallel
                                        gcnOracle.Execute strSql
                                        strRemoveIndex = strRemoveIndex & "," & strSegment & "||" & strTbsOriginal
                                    End If
                                                                
                                    If strParallel <> "" Then
                                        strSql = "Alter Index " & strSegment & " NOParallel"
                                        gcnOracle.Execute strSql
                                    End If
                                    
                                Else 'IOT������֯��
                                    strObjName = GetIOTName(strSegment)
                                    
                                    lblOptPrompt.Caption = "���ڶ�[" & strObjName & "]��������"
                                    lblOptPrompt.Refresh
                                    
                                    If strTBSTemp = "" Then
                                        strSql = "Alter Table " & strObjName & " Move Nologging" & strParallel
                                        gcnOracle.Execute strSql
                                    Else
                                        strSql = "Alter Table " & strObjName & " Move TableSpace " & strTBSTemp & " Nologging" & strParallel
                                        gcnOracle.Execute strSql
                                        strRemoveTable = strRemoveTable & "," & strObjName & "||" & strTbsOriginal
                                    End If
                                    
                                    If strParallel <> "" Then
                                        strSql = "Alter Table " & strObjName & " NOParallel"
                                        gcnOracle.Execute strSql
                                    End If
                                End If
                                
                            '5.��������
                            ElseIf strType = "INDEX PARTITION" Then
                                If CheckLOBIndex(strSegment) Then
                                    'LOB����������LOB������һ��Move
                                    
                                ElseIf CheckIOT(strSegment) = False Then
                                    
                                    strSql = "Select Partition_Name From Dba_Ind_Partitions Where Index_Owner = [1] And Index_Name = [2]"
                                    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
                                    Do While Not rsTmp.EOF
                                        strPartition = rsTmp!Partition_Name
                                    
                                        lblOptPrompt.Caption = "���ڶ�[" & strSegment & "(" & strPartition & ")]�����ؽ�"
                                        lblOptPrompt.Refresh
                                        If strTBSTemp = "" Then
                                            strSql = "Alter Index " & strSegment & " Rebuild Partition " & strPartition & " Nologging" & strParallel & " " & strOnline
                                            gcnOracle.Execute strSql
                                        Else
                                            strSql = "Alter Index " & strSegment & " Rebuild Partition " & strPartition & " TableSpace " & strTBSTemp & " Nologging" & strParallel & " " & strOnline
                                            gcnOracle.Execute strSql
                                            strRemovePARIndex = strRemovePARIndex & "," & strSegment & "||" & strPartition & "||" & strTbsOriginal
                                        End If
                                        rsTmp.MoveNext
                                    Loop
                                                                
                                    If strParallel <> "" Then
                                        strSql = "Alter Index " & strSegment & " NOParallel"
                                        gcnOracle.Execute strSql
                                    End If
                                    
                                Else
                                    If InStr(strPrompt, ":" & strSegment) = 0 Then strPrompt = strPrompt & vbCrLf & "������֯��IOT���ķ�������:" & strSegment
                                End If
                                
                            '6.LOB������
                            ElseIf strType = "LOB PARTITION" Then
                                If CheckUnSuportObject(strSegment, "����(Move)") Then
                                                                        
                                    mrsLobs.Filter = "Owner='" & Split(strSegment, ".")(0) & "' And Segment_Name='" & Split(strSegment, ".")(1) & "'"   'Ϊ��ȡ��ռ���
                                    If mrsLobs.RecordCount > 0 Then
                                        strTableName = mrsLobs!Owner & "." & mrsLobs!Table_name
                                        strTbsLob = mrsLobs!Tablespace_Name
                                        strColumn = mrsLobs!Column_Name
                                        
                                        strSql = "Select Partition_Name From Dba_Lob_Partitions Where Table_Owner = [1] And Lob_Name = [2]"
                                        Set rsTmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
                                        Do While Not rsTmp.EOF
                                            strPartition = rsTmp!Partition_Name
                                            
                                            lblOptPrompt.Caption = "���ڶ�[" & strTableName & "(" & strPartition & ")]��������"
                                            lblOptPrompt.Refresh
                                            If strTBSTemp = "" Then
                                                strSql = "Alter Table " & strTableName & " Move Partition " & strPartition & " Lob(" & strColumn & ") Store as (Tablespace " & strTbsLob & ") Nologging" & strParallel
                                                gcnOracle.Execute strSql
                                            Else
                                                strSql = "Alter Table " & strTableName & " Move Partition " & strPartition & " Lob(" & strColumn & ") Store as (Tablespace " & strTBSTemp & ") Nologging" & strParallel
                                                gcnOracle.Execute strSql
                                                strRemovePARLOB = strRemovePARLOB & "," & strTableName & "||" & strPartition & "||" & strColumn & "||" & strTbsLob
                                            End If
                                            rsTmp.MoveNext
                                        Loop
                                        
                                        'LOB��������ִ�в��ᵼ�±�������degree���Ա����ã����Բ���ִ��noparallel
                                        If strTBSTemp = "" Then Call ReBuildIndex(Split(strTableName, ".")(0), Split(strTableName, ".")(1), strParallel)
                                        
                                    Else
                                        lblOptPrompt.Caption = "����ͼDba_Lobs��δ�ҵ�LOB����" & strSegment & "��"
                                    End If
                                Else
                                    If InStr(strPrompt, ":" & strSegment) = 0 Then strPrompt = strPrompt & vbCrLf & "����Long��Long Raw�ֶεķ�����:" & strSegment
                                End If
                            
                            ElseIf strType <> " " Then
                                lblOptPrompt.Caption = strSegment & ",��֧�ֵĶ������ͣ�" & strType
                            End If
                        End If
NextCell:               strSegmentAll = strSegmentAll & "," & strSegment
                    End If
                    strSegmentPre = strSegment
                End If
            Next
            lblOptPrompt.Caption = "�Ѵ������" & r & "��"
            lblOptPrompt.Refresh
        Next
    End With
    
    If strTBSTemp <> "" Then
        If strRemoveTable & strRemovePARTable & strRemoveLob & strRemoveIndex & strRemovePARIndex & strRemovePARLOB = "" Then
            lblOptPrompt.Caption = "��������" & strTBSTemp & "�Ŀռ�"
            Call ResizeTBS(strTBSTemp)
        End If
    End If
    
    
    '��û���ջؿռ�ģ��Ƶ���ʱ�洢�ı�ռ�Ķ����ƻ�ԭ��ռ䡣
reMove: blnRemove = True
    '1.��
   If strRemoveTable <> "" Then
        arrTmp = Split(Mid(strRemoveTable, 2), ",")
        For r = 0 To UBound(arrTmp)
            strSegment = Split(arrTmp(r), "||")(0)
            strTbsOriginal = Split(arrTmp(r), "||")(1)
            
            DoEvents
            If r = 0 Then
                lblOptPrompt.Caption = "��������" & strTbsOriginal & "�Ŀռ�"
                Call ResizeTBS(strTbsOriginal)
            End If
            
            lblOptPrompt.Caption = "���ڽ�[" & strSegment & "]�ƻ�ԭ��ռ�"
            lblOptPrompt.Refresh
            strSql = "Alter Table " & strSegment & " Move TableSpace " & strTbsOriginal & " Nologging" & strParallel
            gcnOracle.Execute strSql
                        
            If strParallel <> "" Then
                strSql = "Alter Table " & strSegment & " NOParallel"
                gcnOracle.Execute strSql
            End If
                        
            Call ReBuildIndex(Split(strSegment, ".")(0), Split(strSegment, ".")(1), strParallel)
        Next
        
        If strRemovePARTable & strRemoveLob & strRemoveIndex & strRemovePARIndex & strRemovePARLOB = "" Then
            lblOptPrompt.Caption = "��������" & strTBSTemp & "�Ŀռ�"
            lblOptPrompt.Refresh
            Call ResizeTBS(strTBSTemp)
            
            lblOptPrompt.Caption = "��������" & strTbsOriginal & "�Ŀռ�"
            lblOptPrompt.Refresh
            Call ResizeTBS(strTbsOriginal)
        End If
    End If
    
    '2.������
   If strRemovePARTable <> "" Then
        arrTmp = Split(Mid(strRemovePARTable, 2), ",")
        For r = 0 To UBound(arrTmp)
            strSegment = Split(arrTmp(r), "||")(0)
            strPartition = Split(arrTmp(r), "||")(1)
            strTbsOriginal = Split(arrTmp(r), "||")(2)
            
            DoEvents
            If r = 0 Then
                lblOptPrompt.Caption = "��������" & strTbsOriginal & "�Ŀռ�"
                Call ResizeTBS(strTbsOriginal)
            End If
            
            lblOptPrompt.Caption = "���ڽ�[" & strSegment & "(" & strPartition & ")]�ƻ�ԭ��ռ�"
            lblOptPrompt.Refresh
            strSql = "Alter Table " & strSegment & " Move Partition " & strPartition & " TableSpace " & strTbsOriginal & " Nologging" & strParallel
            gcnOracle.Execute strSql
            
            
            If strParallel <> "" Then
                strSql = "Alter Table " & strSegment & " NOParallel"
                gcnOracle.Execute strSql
                
            End If
                 
            '�ƻ����һ���������ؽ�����ʧЧ������
            If r = UBound(arrTmp) Then
                Call ReBuildIndex(Split(strSegment, ".")(0), Split(strSegment, ".")(1), strParallel)
            End If
        Next
        
        If strRemoveLob & strRemoveIndex & strRemovePARIndex & strRemovePARLOB = "" Then
            lblOptPrompt.Caption = "��������" & strTBSTemp & "�Ŀռ�"
            lblOptPrompt.Refresh
            Call ResizeTBS(strTBSTemp)
            
            lblOptPrompt.Caption = "��������" & strTbsOriginal & "�Ŀռ�"
            lblOptPrompt.Refresh
            Call ResizeTBS(strTbsOriginal)
        End If
    End If
    
    '3.LOB��
    If strRemoveLob <> "" Then
        arrTmp = Split(Mid(strRemoveLob, 2), ",")
        For r = 0 To UBound(arrTmp)
            strSegment = Split(arrTmp(r), "||")(0)
            strColumn = Split(arrTmp(r), "||")(1)
            strTbsLob = Split(arrTmp(r), "||")(2)
            
            DoEvents
            If r = 0 Then
                lblOptPrompt.Caption = "��������" & strTbsLob & "�Ŀռ�"
                Call ResizeTBS(strTbsLob)
            End If
                        
            lblOptPrompt.Caption = "���ڽ�[" & strSegment & "(" & strColumn & ")]�ƻ�ԭ��ռ�"
            lblOptPrompt.Refresh
            strSql = "ALTER TABLE " & strSegment & " Move LOB (" & strColumn & ") Store as(Tablespace " & strTbsLob & ") Nologging" & strParallel
            gcnOracle.Execute strSql
            
            
            'LOB����ִ�в��ᵼ�±�������degree���Ա����ã����Բ���ִ��noparallel
        Next
        
        If strRemoveIndex & strRemovePARIndex & strRemovePARLOB = "" Then
            lblOptPrompt.Caption = "��������" & strTBSTemp & "�Ŀռ�"
            lblOptPrompt.Refresh
            Call ResizeTBS(strTBSTemp)
            
            lblOptPrompt.Caption = "��������" & strTbsLob & "�Ŀռ�"
            lblOptPrompt.Refresh
            Call ResizeTBS(strTbsLob)
        End If
    End If
    
    '4.����
    If strRemoveIndex <> "" Then
        arrTmp = Split(Mid(strRemoveIndex, 2), ",")
        For r = 0 To UBound(arrTmp)
            strSegment = Split(arrTmp(r), "||")(0)
            strTbsOriginal = Split(arrTmp(r), "||")(1)
            
            DoEvents
            If r = 0 Then
                lblOptPrompt.Caption = "��������" & strTbsOriginal & "�Ŀռ�"
                Call ResizeTBS(strTbsOriginal)
            End If
                        
            lblOptPrompt.Caption = "���ڽ�[" & strSegment & "]�ƻ�ԭ��ռ�"
            lblOptPrompt.Refresh
            strSql = "Alter Index " & strSegment & " Rebuild " & strOnline & " TableSpace " & strTbsOriginal & " Nologging" & strParallel
            gcnOracle.Execute strSql
            
                               
            If strParallel <> "" Then
                strSql = "Alter Index " & strSegment & " NOParallel"
                gcnOracle.Execute strSql
            End If
        Next
        
        If strRemovePARIndex & strRemovePARLOB = "" Then
            lblOptPrompt.Caption = "��������" & strTBSTemp & "�Ŀռ�"
            lblOptPrompt.Refresh
            Call ResizeTBS(strTBSTemp)
            
            lblOptPrompt.Caption = "��������" & strTbsOriginal & "�Ŀռ�"
            lblOptPrompt.Refresh
            Call ResizeTBS(strTbsOriginal)
        End If
    End If
    
    '5.��������
    If strRemovePARIndex <> "" Then
        arrTmp = Split(Mid(strRemovePARIndex, 2), ",")
        For r = 0 To UBound(arrTmp)
            strSegment = Split(arrTmp(r), "||")(0)
            strPartition = Split(arrTmp(r), "||")(1)
            strTbsOriginal = Split(arrTmp(r), "||")(2)
            
            DoEvents
            If r = 0 Then
                lblOptPrompt.Caption = "��������" & strTbsOriginal & "�Ŀռ�"
                Call ResizeTBS(strTbsOriginal)
            End If
                        
            lblOptPrompt.Caption = "���ڽ�[" & strSegment & "(" & strPartition & ")]�ƻ�ԭ��ռ�"
            lblOptPrompt.Refresh
            strSql = "Alter Index " & strSegment & " Rebuild Partition " & strPartition & " TableSpace " & strTbsOriginal & " Nologging" & strParallel & " " & strOnline
            gcnOracle.Execute strSql
            
                               
            If strParallel <> "" Then
                strSql = "Alter Index " & strSegment & " NOParallel"
                gcnOracle.Execute strSql
                
            End If
        Next
        
        If strRemovePARLOB = "" Then
            lblOptPrompt.Caption = "��������" & strTBSTemp & "�Ŀռ�"
            lblOptPrompt.Refresh
            Call ResizeTBS(strTBSTemp)
            
            lblOptPrompt.Caption = "��������" & strTbsOriginal & "�Ŀռ�"
            lblOptPrompt.Refresh
            Call ResizeTBS(strTbsOriginal)
        End If
    End If
    
     '6.LOB����
    If strRemovePARLOB <> "" Then
        arrTmp = Split(Mid(strRemovePARLOB, 2), ",")
        For r = 0 To UBound(arrTmp)
            strSegment = Split(arrTmp(r), "||")(0)
            strPartition = Split(arrTmp(r), "||")(1)
            strColumn = Split(arrTmp(r), "||")(2)
            strTbsLob = Split(arrTmp(r), "||")(3)
            
            DoEvents
            If r = 0 Then
                lblOptPrompt.Caption = "��������" & strTbsLob & "�Ŀռ�"
                Call ResizeTBS(strTbsLob)
            End If
                        
            lblOptPrompt.Caption = "���ڽ�[" & strSegment & "(" & strPartition & ")]�ƻ�ԭ��ռ�"
            lblOptPrompt.Refresh
            strSql = "Alter Table " & strSegment & " Move Partition " & strPartition & " Lob(" & strColumn & ") Store as (Tablespace " & strTbsLob & ") Nologging" & strParallel
            gcnOracle.Execute strSql
            
            '�ƻ����һ���������ؽ�����ʧЧ������
            If r = UBound(arrTmp) Then
                Call ReBuildIndex(Split(strSegment, ".")(0), Split(strSegment, ".")(1), strParallel)
            End If
            
            'LOB��������ִ�в��ᵼ�±�������degree���Ա����ã����Բ���ִ��noparallel
        Next
                    
        lblOptPrompt.Caption = "��������" & strTBSTemp & "�Ŀռ�"
        lblOptPrompt.Refresh
        Call ResizeTBS(strTBSTemp)
        
        lblOptPrompt.Caption = "��������" & strTbsLob & "�Ŀռ�"
        lblOptPrompt.Refresh
        Call ResizeTBS(strTbsLob)
    End If
    
    
    'ˢ������
    Call RefreshData
    
    If strSegment <> "" Then
        mrsExtents.Filter = "Segment_Name='" & Split(strSegment, ".")(1) & "' And Owner='" & Split(strSegment, ".")(0) & "'"
        If mrsExtents.RecordCount > 0 Then
            vsfExtents.SetFocus
            vsfExtents.Select mrsExtents!Row, mrsExtents!Col
            vsfExtents.TopRow = vsfExtents.Row
        End If
    End If
    
    strTime = GetTimeString(datBegin, GetCurrentdate)
    
    If strPrompt <> "" Then
        strPrompt = Mid(strPrompt, 2, 1500)
        MsgBox "����ִ����ɣ����ι���ʱ��" & strTime & "��" & vbCrLf & _
            "δ��֧�����¶����������" & vbCrLf & strPrompt, vbInformation, gstrSysName
    Else
        MsgBox "����ִ����ɣ����ι���ʱ��" & strTime & "��", vbInformation, gstrSysName
    End If
    
    Call SetCommandEnable(1)
    Exit Sub
errH:
    Call ErrCenter(strSql)

    If 0 = 1 Then
        Resume
    End If
    If blnRemove = False Then GoTo reMove
    
    If txtParallel.Text <> "0" Then
        Call SetNOParallel(gcnOracle, 0)
        Call SetNOParallel(gcnOracle, 1)
    End If
    
    Call SetCommandEnable(1)
End Sub

Private Function CheckUnSuportTable(ByVal strOwner As String, ByVal strTable As String)
'���ܣ������Ƿ����('mdsys�û��´�������GridFile1044_TAB���ֺ���Сд��ĸ�ı�Σ�����dba_tables��ȴ�鲻����¼)
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    'iot_name��Ϊ�յģ���IOT�����������
    'mdsys��һ�ű�SDO_3DTXFMS_TABLE������SDO_NUMBER_ARRAY�������ͣ����²���Move
    'Data_Type_OwnerΪPublic����XMLTYPE
    strSql = "Select 1" & vbNewLine & _
            "From Dba_Tables A" & vbNewLine & _
            "Where Owner = [1] And Table_Name = [2] And Iot_Name Is Null And Not Exists" & vbNewLine & _
            " (Select 1" & vbNewLine & _
            "       From Dba_Tab_Cols B" & vbNewLine & _
            "       Where a.Owner = b.Owner And a.Table_Name = b.Table_Name And Nvl(b.Data_Type_Owner,'PUBLIC') <> 'PUBLIC' And b.Data_Type<> 'XMLTYPE')"

    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, strOwner, strTable)

    CheckUnSuportTable = rsTmp.RecordCount > 0
    
    Exit Function
errH:
    Call ErrCenter(strSql)
End Function

Private Sub cmdShrink_Click()
'���ܣ�ִ�б�������Ŀ��пռ��ջ�(Shrink Space)
    Dim strSql As String, strType As String, strSegment As String, strSegmentPre As String, strSegmentAll As String
    Dim rsTmp As ADODB.Recordset
    Dim blnRow_Movement As Boolean
    Dim r1 As Long, c1 As Long, r2 As Long, c2 As Long, r As Long, c As Long
    Dim strSegment_Type As String, strObjName As String
    
    If CheckExtent(P1����) = False Then Exit Sub
        
    On Error GoTo errH
    strType = Trim(lblPrompt.Tag)
    If strType <> "" And InStr(",TABLE,INDEX,LOBSEGMENT,", "," & strType & ",") = 0 Then
        Call MsgBox("��֧�ֶԱ���������п��пռ��ջأ���֧�ֵ��������ͣ�" & strType, vbInformation, Me.Caption)
        Exit Sub
    End If
    
    Call SetCommandEnable(0)
    vsfExtents.GetSelection r1, c1, r2, c2
    For r = r2 To r1 Step -1
        For c = c2 To c1 Step -1
            strSegment = mcolCells("_" & r & "_" & c)     '��������
            strSegment_Type = CStr(vsfExtents.Cell(flexcpData, r, c))
            
            If strSegment & "|" & strSegment_Type <> strSegmentPre Then
                If InStr(strSegmentAll & ",", "," & strSegment & "|" & strSegment_Type & ",") = 0 Then
                    mrsExtents.Filter = "Row=" & r & " And Col=" & c
                    If mrsExtents.RecordCount > 0 Then
                        DoEvents
            
                        strType = mrsExtents!Segment_Type
                        If strType = "TABLE" Then
                            If CheckUnSuportObject(strSegment, "�ջ�(Shrink Space)") Then
                                strSql = "Select Row_Movement From All_Tables Where Table_Name = [1] And Owner = [2]"
                                Set rsTmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(1), Split(strSegment, ".")(0))
                                If rsTmp.RecordCount = 0 Then
                                    Call MsgBox("����ͼAll_Tables��δ�ҵ�ָ���Ķ���" & strSegment, vbInformation, Me.Caption)
                                    Call SetCommandEnable(1)
                                    Exit Sub
                                End If
                                If rsTmp!Row_Movement = "DISABLED" Then 'enable row movement����������ñ�XXX�Ķ���(��洢���̡�������ͼ��)��Ϊ��Ч
                                    strSql = "Alter Table " & strSegment & " Enable Row Movement"
                                    gcnOracle.Execute strSql
                                    blnRow_Movement = True
                                End If
                                       
                                lblOptPrompt.Caption = "���ڶ�[" & strSegment & "]���пռ��ջ�"
                                                                
                                strSql = "Alter Table " & strSegment & " Shrink Space"
                                gcnOracle.Execute strSql
                                
                                If blnRow_Movement Then
                                    strSql = "Alter Table " & strSegment & " Disable Row Movement"
                                    gcnOracle.Execute strSql
                                End If
                            End If
                            
                        ElseIf strType = "LOBSEGMENT" Then
                            mrsLobs.Filter = "Owner='" & Split(strSegment, ".")(0) & "' And Segment_Name='" & Split(strSegment, ".")(1) & "'"
                            If mrsLobs.RecordCount > 0 Then
                                lblOptPrompt.Caption = "���ڶ�[" & mrsLobs!Table_name & "." & mrsLobs!Column_Name & "]���пռ����"
                                                                    
                                strSql = "ALTER TABLE " & mrsLobs!Owner & "." & mrsLobs!Table_name & " MODIFY LOB (" & mrsLobs!Column_Name & ") (SHRINK SPACE)"
                                gcnOracle.Execute strSql
                            Else
                                lblOptPrompt.Caption = "����ͼDba_Lobs��δ�ҵ�LOB����" & strSegment & "��"
                            End If
                            
                        ElseIf strType = "INDEX" Then
                            If Not CheckIOT(strSegment) Then
                                lblOptPrompt.Caption = "���ڶ�[" & strSegment & "]���пռ��ջ�"
                                                            
                                strSql = "Alter Index " & strSegment & " Shrink Space"
                                gcnOracle.Execute strSql
                            Else
                                strObjName = GetIOTName(strSegment)
                                strSql = "Alter Table " & strObjName & " Shrink Space"
                                gcnOracle.Execute strSql
                            End If
                        ElseIf strType <> " " Then
                            lblOptPrompt.Caption = strSegment & ",��֧�ֵĶ������ͣ�" & strType
                        End If
                    End If
                    strSegmentAll = strSegmentAll & "," & strSegment & "|" & strSegment_Type
                End If
                strSegmentPre = strSegment & "|" & strSegment_Type
            End If
        Next
        lblOptPrompt.Caption = "�Ѵ������" & r & "��"
        lblOptPrompt.Refresh
    Next
    
    'δ�ı������ļ���С������ˢ�±�ռ估�����ļ��б�
    Call LoadExtents(vsfTbs.TextMatrix(vsfTbs.Row, vsfTbs.ColIndex("����")), Val(cboFiles.ItemData(cboFiles.ListIndex)))
    
    If strSegment <> "" Then
        mrsExtents.Filter = "Segment_Name='" & Split(strSegment, ".")(1) & "' And Owner='" & Split(strSegment, ".")(0) & "'"
        If mrsExtents.RecordCount > 0 Then
            vsfExtents.SetFocus
            vsfExtents.Select mrsExtents!Row, mrsExtents!Col
            vsfExtents.TopRow = vsfExtents.Row
        End If
    End If
    
    Call SetCommandEnable(1)
    Exit Sub
errH:
    If MsgBox(Err.Description & vbCrLf & "���һ��ִ�е�SQL��" & vbCrLf & strSql & vbCrLf & "��������Ϊ��ǰ����ҵ���Ӱ�죬�Ƿ����ԣ�", vbRetryCancel, "����") = vbRetry Then
        Resume
    End If
    
    If blnRow_Movement Then
        strSql = "Alter Table " & strSegment & " Disable Row Movement"
        gcnOracle.Execute strSql
    End If
    lblOptPrompt.Caption = ""
    Call SetCommandEnable(1)
End Sub

Private Function CheckExtent(ByVal bytOpt As opt) As Boolean
    Dim strSegment As String, strPrompt As String
    Dim r1&, c1&, r2&, c2&, r&, c&
    
    On Error Resume Next
    If vsfExtents.Row = -1 Or vsfExtents.Col = -1 Then
        MsgBox "����ѡ��һ����Ԫ����ִ�б�����", vbInformation, Me.Caption
        Exit Function
    End If
    If mcolCells Is Nothing Then
        MsgBox "����ˢ�����ݲ�����һ���洢�����ݵĵ�Ԫ����ִ�б�����", vbInformation, Me.Caption
        Exit Function
    End If
    
    With vsfExtents
        .GetSelection r1, c1, r2, c2
        If r1 = r2 And c1 = c2 Then '��ѡ��һ����Ԫ��ʱ�ż��
            strSegment = mcolCells("_" & .Row & "_" & .Col)
            If strSegment = "" Or strSegment = "sys.free" Or cboFiles.ListIndex = -1 Then
                MsgBox "����ѡ��һ���洢�����ݵĵ�Ԫ����ִ�б�����", vbInformation, Me.Caption
                Exit Function
            End If
            mrsLobs.Filter = "Owner='" & Split(strSegment, ".")(0) & "' And Segment_Name='" & Split(strSegment, ".")(1) & "'"
            If mrsLobs.RecordCount > 0 Then
                strSegment = strSegment & "(" & mrsLobs!Table_name & "." & mrsLobs!Column_Name & ")"
            End If
        Else
            strSegment = mcolCells("_" & .Row & "_" & .Col) & "��"
        End If
    End With
        
    If bytOpt = P1���� Then
        strSegment = "����(Shrink)һ������ɾ���������ݺ󽵵͸�ˮ��ǣ��Ա�����ļ������������������̲�Ӱ��ҵ�����У���������ҵ������ڼ�ִ�У���ȷ��Ҫ��" & vbCrLf & vbTab & strSegment & vbCrLf & "���л��ղ�����"
        
    ElseIf bytOpt = P2���� Then
        strSegment = "����(Move Or Rebuild)һ�������ƶ��������λ�ã��������̻�����������Ҫ��ö�������Ŀ��пռ䣬�����ж�ҵ�����У���������ҵ������ڼ�ִ�У������ء�" & vbCrLf & _
                "Move��֮�����������ʧЧ�������������Զ��ؽ������ܺ�ʱ�ϳ�����ȷ��Ҫ��" & vbCrLf & vbTab & strSegment & vbCrLf & "��������������"
    End If
    If MsgBox(strSegment, vbOKCancel + vbDefaultButton1, Me.Caption) = vbCancel Then
        Exit Function
    End If
        
    CheckExtent = True
End Function


Private Sub Form_Load()
    Dim strCol As String, i As Long
    
    strCol = "��,300,1;״̬;����,1250,1;��С,500,1"
    Call InitTable(vsfTbs, strCol)
    vsfTbs.FixedCols = 1
    
    strCol = ""
    For i = 0 To CONCOLS
        If strCol = "" Then
            strCol = i & ",550,1"
        Else
            strCol = strCol & ";" & i & ",280,4"
        End If
    Next

    Call InitTable(vsfExtents, strCol)
    vsfExtents.FixedCols = 1
    vsfExtents.Rows = vsfExtents.FixedRows
    vsfExtents.TextMatrix(0, 0) = "��\��"
    
    
    Call LoadTablespaces
    
    vsfTbs.Editable = flexEDNone
    vsfExtents.Editable = flexEDNone
    
    Call LoadParallel
    
    'Me.Caption = Me.Caption & "(��������" & gstrServer & ")"
End Sub

Private Sub LoadParallel()
'���ܣ���ȡ����ʾ���ж�
    
    On Error GoTo errH
    If gintCpuCount = 0 Then
        txtParallel.Text = "0"
        txtParallel.Locked = True
        txtParallel.Enabled = False
        lblParallel.ToolTipText = "δ�ܶ�ȡ�����ݿ����cpu_count"
    Else
        txtParallel.Tag = gintCpuCount
        If gintCpuCount < 3 Then
            txtParallel.Text = "0"
            txtParallel.Enabled = False
            lblParallel.ToolTipText = "������Cpu��������3�������ܽ��в���ִ��"
        ElseIf gintCpuCount < 13 Then
            txtParallel.Text = gintCpuCount \ 2 'һ��ȡ��
        Else
            txtParallel.Text = "12"  '��ʹcpu�㹻�����Կ��������ڴ������ܣ����жȲ���Խ��Խ��
        End If
    End If

    Exit Sub
errH:
    Call ErrCenter
End Sub




Private Sub mnuAddFile_Click()
    '��������ļ�
    Dim strTblSpace As String, strQuery As String
    Dim strFileName As String, strFilePth As String
    Dim strSql As String
    
    On Error GoTo errH
    With vsfTbs
        If .Row = -1 Or .Row = 0 Then
            MsgBox "����ѡ��һ����ռ���ִ�в�����"
            Exit Sub
        End If
        
        strTblSpace = .TextMatrix(.Row, .ColIndex("����"))
    End With
    
    If strTblSpace = "" Then
        MsgBox "��ȡ��ռ�����ʧ�ܣ������²�����"
        Exit Sub
    Else
        Call SetCommandEnable(0)
        strFileName = GetDataFile(strTblSpace, strFilePth)
        strQuery = Trim(InputBox("Ϊ��ռ�" & strTblSpace & "��������ļ�" & vbCrLf & vbCrLf & _
                                                            "Ĭ����ӵ������ļ���СΪ100M�������������Ҫ�����ֶ�ִ������ָ�" & vbCrLf & vbCrLf & _
                                                            "ALTER TABLESPACE " & strTblSpace & " ADD DATAFILE " & vbCrLf & "'" & strFilePth & strFileName & "' SIZE 100M AUTOEXTEND ON" _
                                                        , "��ռ���������ļ�", strFileName))
        
        If strQuery = "" Then
            Call SetCommandEnable(1)
            Exit Sub
        Else
            lblOptPrompt.Caption = "����Ϊ��ռ�" & strTblSpace & "��������ļ�" & strFilePth & strFileName & "......"
            strSql = "ALTER TABLESPACE " & strTblSpace & " ADD DATAFILE '" & strFilePth & strQuery & "' SIZE 100M AUTOEXTEND ON"
            gcnOracle.Execute strSql
        End If
        lblOptPrompt.Caption = "��ռ�" & strTblSpace & "��������ļ�" & strFilePth & strQuery & "�ɹ���"
        Call SetCommandEnable(1)
    End If
    
    Exit Sub
errH:
    Call SetCommandEnable(1)
    lblOptPrompt.Caption = "��ռ�" & strTblSpace & "��������ļ�" & strFilePth & strQuery & "ʧ�ܡ�"
    If InStr(Err.Description, "ORA-01537") > 0 Then
        MsgBox "��ǰ��ռ��Ѿ�������Ϊ" & strQuery & "�������ļ������������롣"
        Exit Sub
    End If
    
    ErrCenter
End Sub

Private Sub mnuResizeAll_Click()
    Call ResizeAll
End Sub

Private Sub mnuResizeTemp_Click()
'������ʱ��ռ�
    Call ResizeTemp
End Sub

Private Sub mnuResizeUndo_Click()
'����Undo��ռ�
    Call frmResizeUndo.ShowMe(frmReused)
End Sub


Private Sub txtParallel_GotFocus()
    txtParallel.SelStart = 0
    txtParallel.SelLength = Len(txtParallel.Text)
End Sub

Private Sub txtParallel_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtParallel_Validate(Cancel As Boolean)
    If Val(txtParallel.Tag) <> 0 Then
        If Val(txtParallel.Text) > Val(txtParallel.Tag) Then
            MsgBox "���жȲ��ܳ���cpu����" & txtParallel.Tag, vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

Private Sub LoadTablespaces()
    Dim rsTmp As ADODB.Recordset, strSql  As String
    Dim i As Long, lngStart As Long
    
    strSql = "Select a.Status, a.Tablespace_Name, a.Block_Size, Round(Sum(b.Bytes) / 1024 / 1024, 2) Tsize , Max(Decode(b.autoextensible,'YES',0,1)) as autoextensible" & vbNewLine & _
            "From Dba_Tablespaces A, Dba_Data_Files B" & vbNewLine & _
            "Where a.Contents = 'PERMANENT' And a.Tablespace_Name = b.Tablespace_Name And b.Online_status in('ONLINE','SYSTEM')" & vbNewLine & _
            "Group By a.Tablespace_Name, a.Status, a.Block_Size" & vbNewLine & _
            "Order By 4 Desc"

    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSql, Me.Caption)
    
    With vsfTbs
        .Redraw = flexRDNone
        lngStart = .FixedRows
        .Rows = lngStart
        .Rows = lngStart + rsTmp.RecordCount
        For i = lngStart To rsTmp.RecordCount
            If rsTmp!autoextensible = 1 Then
                .Cell(flexcpBackColor, i, .ColIndex("��"), i, .ColIndex("��С")) = OFF_��ɫ
                .Cell(flexcpData, i, .ColIndex("��С")) = "NO"
            End If
            .TextMatrix(i, .ColIndex("��")) = i
            .TextMatrix(i, .ColIndex("״̬")) = rsTmp!Status
            .TextMatrix(i, .ColIndex("����")) = rsTmp!Tablespace_Name
            
            If Val("" & rsTmp!Tsize) > 1024 Then
                .TextMatrix(i, .ColIndex("��С")) = Round(rsTmp!Tsize / 1024, 2) & "G"
            Else
                .TextMatrix(i, .ColIndex("��С")) = rsTmp!Tsize & "M"
            End If
            
            .RowData(i) = Val(rsTmp!Block_Size)
            rsTmp.MoveNext
        Next
        .Redraw = flexRDDirect
    End With

    Exit Sub
errH:
    Call ErrCenter(strSql)
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    cmdGO.Left = Abs(Me.ScaleWidth - cmdGO.Width - 60)
    lblPrompt.Width = Abs(Me.ScaleWidth - cmdGO.Width)
    
    vsfExtents.Width = Abs(Me.ScaleWidth - vsfExtents.Left - 60)
    vsfExtents.ColWidth(-1) = (vsfExtents.Width - 550 - 120) / 51   '120Ϊ�������Ŀ��
    vsfExtents.ColWidth(vsfExtents.FixedRows - 1) = 550
    vsfExtents.RowHeight(-1) = vsfExtents.Width / 51
    
    vsfTbs.Height = Abs(Me.ScaleHeight - vsfTbs.Top - 60 - picBottom.Height)
    vsfExtents.Height = Abs(vsfTbs.Height - lblPrompt.Height)
    
    cmdMore.Left = Abs(Me.ScaleWidth - cmdMore.Width - 60)
    cmdResize.Left = Abs(cmdMore.Left - cmdResize.Width - 60)
    cmdShrink.Left = Abs(cmdResize.Left - cmdShrink.Width - 60)
    cmdMove.Left = Abs(cmdShrink.Left - cmdMove.Width - 60)
    
    chkOnline.Left = Abs(cmdMove.Left - chkOnline.Width - 60)
    txtParallel.Left = Abs(chkOnline.Left - txtParallel.Width - 60)
    lblParallel.Left = Abs(txtParallel.Left - lblParallel.Width)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsExtents = Nothing
    Set mcolCells = Nothing
    Set mrsLobs = Nothing
End Sub

Private Sub LoadLobs(ByVal strTBS As String)
'���ܣ���ȡ��ǰ��ռ��Lob����Ϣ
    Dim strSql As String
 
    strSql = "Select Table_Name, TableSpace_Name, Column_Name, Owner, Segment_Name, Index_Name From Dba_Lobs Where Tablespace_Name = [1]"
    On Error GoTo errH
    Set mrsLobs = OpenSQLRecord(strSql, Me.Caption, strTBS)

    Exit Sub
errH:
    Call ErrCenter(strSql)
End Sub

Public Sub LoadExtents(ByVal strTBS As String, ByVal lngFile As Long)
'���ܣ�����Extents����Ԫ��
    Dim rsTmp As ADODB.Recordset, strSql  As String, strSegment As String, strPreSegment As String, strFullSegment As String
    Dim i As Long, j As Long, n As Long, lngStart As Long, lngRows As Long
    Dim lngCells As Long, lngFixedCols As Long
    Dim blnFree As Boolean, blnSameCell As Boolean, strFirst As String
    
    lblOptPrompt.Caption = "���ڶ�ȡ���ݿ���Ϣ......"
    lblOptPrompt.Refresh
    
    If chkFree.Value = 1 Then
        strSql = "Select File_Id,Block_Id as Extent_ID, Block_Id as First_Block, Block_Id + Blocks - 1 as Last_Block,Blocks, 'free' as Segment_Name, 'sys.free' as Full_Segment_Name, ' ' as Segment_Type,' ' as Owner" & vbNewLine & _
            "From Dba_Free_Space A" & vbNewLine & _
            "Where Tablespace_Name = [1] And a.File_Id = [2]" & vbNewLine & _
            "Order By First_Block"
    Else
        strSql = "Select a.File_Id,a.Extent_ID, a.Block_Id First_Block, a.Block_Id + a.Blocks - 1 Last_Block,a.Blocks, a.Segment_Name, a.Owner || '.' || a.Segment_Name as Full_Segment_Name, b.Segment_Type, a.Owner" & vbNewLine & _
            "From Dba_Extents A, Dba_Segments B" & vbNewLine & _
            "Where a.Tablespace_Name = [1] And a.File_Id = [2] And a.Segment_Name = b.Segment_Name And a.Owner = b.Owner" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select File_Id,0, Block_Id, Block_Id + Blocks - 1,Blocks, 'free', 'sys.free' as Full_Segment_Name, ' ',' '" & vbNewLine & _
            "From Dba_Free_Space A" & vbNewLine & _
            "Where Tablespace_Name = [1] And a.File_Id = [2]" & vbNewLine & _
            "Order By First_Block"
    End If
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, strTBS, lngFile)
            
    Call InitmrsExtents
    
    If rsTmp.RecordCount = 0 Then
        lblOptPrompt.Caption = ""
        lblOptPrompt.Refresh
        vsfExtents.Rows = vsfExtents.FixedRows
        Exit Sub
    End If
    
    lblOptPrompt.Caption = "���ڼ������ݿ���Ϣ......"
    lblOptPrompt.Refresh
    lngFixedCols = vsfExtents.FixedCols
    lngStart = vsfExtents.FixedRows
    
    
    '�ȼ��������,������ʾ����
    j = lngFixedCols
    lngRows = lngStart
    Do While Not rsTmp.EOF
        lngCells = rsTmp!blocks \ CONBLOCKS 'ȡ��
        If rsTmp!blocks <> lngCells * CONBLOCKS Then lngCells = lngCells + 1
        
        For n = 1 To lngCells
            j = j + 1
            If j > CONCOLS Then '����
                lngRows = lngRows + 1
                j = lngFixedCols
            End If
        Next
        rsTmp.MoveNext
    Loop
    If rsTmp.RecordCount <> 0 Then
        rsTmp.MoveFirst
    End If
    
    vsfExtents.Redraw = flexRDNone  '���ⴥ���¼�vsfExtents_AfterRowColChange
    vsfExtents.Rows = lngStart
    
    vsfExtents.Redraw = flexRDDirect
    vsfExtents.ToolTipText = ""
    vsfExtents.Refresh
    lblPrompt.Caption = ""
    vsfExtents.Redraw = flexRDNone
    vsfExtents.Rows = lngStart + lngRows

    vsfExtents.Redraw = flexRDDirect
    
        
    With vsfExtents
        .Redraw = flexRDNone
                
        i = lngStart
        j = .FixedCols
        If i > 0 Then .TextMatrix(1, 0) = 1
        
        Do While Not rsTmp.EOF
            strSegment = rsTmp!Segment_Name
            blnFree = (strSegment = "free")
            strFullSegment = rsTmp!Full_Segment_Name
                                    
            strFirst = Mid$(strSegment, 1, 1)
            If strPreSegment <> strSegment & "|" & rsTmp!Segment_Type Then
                blnSameCell = Mid$(strPreSegment, 1, 1) = strFirst
            Else
                blnSameCell = False
            End If
            
            lngCells = rsTmp!blocks \ CONBLOCKS 'ȡ��
            If rsTmp!blocks <> lngCells * CONBLOCKS Then lngCells = lngCells + 1
           
            For n = 1 To lngCells
                If blnFree Then
                    .Cell(flexcpBackColor, i, j) = &HCCEDC7 '���пռ�
                    If n = 1 Then .TextMatrix(i, j) = "B"
                Else
                    .TextMatrix(i, j) = strFirst
                    .Cell(flexcpData, i, j) = CStr(rsTmp!Segment_Type)
                End If
                mcolCells.Add strFullSegment, "_" & i & "_" & j
               
                '��һ������ͬ��������ͬ���üӴ�������
                If blnSameCell Then .Cell(flexcpFontItalic, i, j) = True
                
                mrsExtents.AddNew Array("Row", "Col", "Segment_Name", "Extent_ID", "First_Block", "Blocks", "Last_Block", "Segment_Type", "Owner"), _
                            Array(i, j, strSegment, rsTmp!Extent_ID, rsTmp!First_Block, rsTmp!blocks, rsTmp!Last_Block, rsTmp!Segment_Type, rsTmp!Owner)
                               
                j = j + 1
                If j > CONCOLS Then '����
                    j = lngFixedCols
                    
                    i = i + 1
                   .TextMatrix(i, 0) = i   '�к�
                   
                   If i Mod 100 = 0 Then
                     DoEvents
                     lblOptPrompt.Caption = "���ڼ�����Ϣ(" & i & "/" & lngRows & ")"
                   End If
                End If
           Next
           strPreSegment = strSegment & "|" & rsTmp!Segment_Type
           rsTmp.MoveNext
        Loop
        
        'ʣ��Ŀյ�Ԫ����Ͽ�ֵ�Ա���Ӽ���ȡֵʱ����
        For n = j To CONCOLS
            mcolCells.Add "", "_" & i & "_" & n
        Next

        .Redraw = flexRDDirect
    End With
    lblOptPrompt.Caption = ""
    Exit Sub
errH:
    Call ErrCenter(strSql)
    If 0 = 1 Then
        Resume
    End If
End Sub

Private Sub txtFind_GotFocus()
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdFind_Click
    End If
End Sub

Private Sub cmdFind_Click()
    If Not mrsExtents Is Nothing Then
        If InStr(txtFind.Text, "*") > 0 Then
            mrsExtents.Filter = "Segment_Name Like '" & UCase(Trim(txtFind.Text)) & "'"
        Else
            mrsExtents.Filter = "Segment_Name='" & UCase(Trim(txtFind.Text)) & "'"
        End If
        If mrsExtents.RecordCount > 0 Then
            vsfExtents.SetFocus
            vsfExtents.Select mrsExtents!Row, mrsExtents!Col
            vsfExtents.TopRow = vsfExtents.Row
        Else
            lblOptPrompt.Caption = "û���ҵ�ƥ��ı��������"
            txtFind.SetFocus
            txtFind_GotFocus
        End If
    Else
        lblOptPrompt.Caption = "û���ҵ�ƥ��ı��������"
        txtFind.SetFocus
        txtFind_GotFocus
    End If
End Sub


Private Sub cmdGO_Click()
'���ܣ�����LOB������������� ��λ��LOB����
    Dim strObjName As String, strSegment As String, strSegment_Type As String
    Dim i As Long, j As Long
    
    If vsfExtents Is Nothing Then
        Exit Sub
    End If
    With vsfExtents
        If .Row < 0 Or .Col < 0 Then Exit Sub
        
        strSegment_Type = .Cell(flexcpData, .Row, .Col)
        strSegment = mcolCells("_" & .Row & "_" & .Col)
        
        If strSegment = "" Then Exit Sub
        
        If strSegment_Type = "LOBINDEX" Or strSegment_Type = "INDEX PARTITION" Then
            strObjName = GetLOBNameByIndex(strSegment)
        End If
        
        If strObjName <> "" Then
            For i = .FixedRows To .Rows - 1
                For j = .FixedCols To .Cols - 1
                    If strObjName = mcolCells("_" & i & "_" & j) Then
                        .Select i, j
                        .TopRow = i
                        .SetFocus
                        strObjName = ""
                        Exit Sub
                    End If
                Next
            Next
            If strObjName <> "" Then Call MsgBox("δ�ҵ�" & strObjName, vbInformation)
        End If
    End With
End Sub

Private Sub vsfExtents_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfExtents
        Dim strSegment As String, i As Long, lngBlockSize As Long, strSegment_Type As String
        
        If Me.Visible = False Or .Redraw = flexRDNone Or mcolCells Is Nothing Or vsfTbs.Enabled = False Then Exit Sub
        
        .Redraw = flexRDNone
        '��ȥ��֮ǰѡ�еĶεı���ɫ
        If OldRow > 0 And OldCol > 0 Then
            strSegment = mcolCells("_" & OldRow & "_" & OldCol)
            If strSegment <> "" Then
                strSegment_Type = .Cell(flexcpData, OldRow, OldCol)
                mrsExtents.Filter = "Segment_Name='" & Split(strSegment, ".")(1) & "' And Owner='" & Split(strSegment, ".")(0) & "' And Segment_Type='" & strSegment_Type & "'"
                For i = 1 To mrsExtents.RecordCount
                    If mrsExtents!Segment_Name = "free" Then
                        .Cell(flexcpBackColor, mrsExtents!Row, mrsExtents!Col) = &HCCEDC7 '���пռ�
                    Else
                        .Cell(flexcpBackColor, mrsExtents!Row, mrsExtents!Col) = &H80000005 '��ɫ
                    End If
                    .Cell(flexcpForeColor, mrsExtents!Row, mrsExtents!Col) = vbBlack
                    mrsExtents.MoveNext
                Next
            End If
        End If
                
        .Redraw = flexRDDirect
        
        
        '�����õ�ǰѡ�жεı���ɫ
        .Redraw = flexRDNone
        cmdGO.Visible = False
        
        strSegment = mcolCells("_" & NewRow & "_" & NewCol)
        lblPrompt.Tag = ""
        If strSegment <> "" Then
            strSegment_Type = .Cell(flexcpData, NewRow, NewCol)
            
            If strSegment_Type = "LOBINDEX" Then
                cmdGO.Visible = True
                cmdGO.Caption = "��λ��LOB"
            ElseIf strSegment_Type = "INDEX PARTITION" Then
                If CheckLOBIndex(strSegment) Then
                    cmdGO.Visible = True
                    cmdGO.Caption = "��λ��LOB"
                End If
            End If
            
            mrsExtents.Filter = "Row=" & NewRow & " And Col=" & NewCol
            If mrsExtents.RecordCount > 0 Then
                If mrsExtents!Segment_Type = "LOBSEGMENT" Then
                    mrsLobs.Filter = "Owner='" & mrsExtents!Owner & "' And Segment_Name='" & mrsExtents!Segment_Name & "'"
                    If mrsLobs.RecordCount > 0 Then .ToolTipText = mrsLobs!Table_name & "." & mrsLobs!Column_Name
                
                ElseIf mrsExtents!Segment_Type = "LOBINDEX" Then
                    mrsLobs.Filter = "Owner='" & mrsExtents!Owner & "' And Index_Name='" & mrsExtents!Segment_Name & "'"
                    If mrsLobs.RecordCount > 0 Then .ToolTipText = mrsLobs!Table_name & "." & mrsLobs!Column_Name & "(Index)"
                Else
                    .ToolTipText = strSegment & "(һ����Ԫ�����" & CONBLOCKS & "����)"
                End If
                
                lngBlockSize = Val(vsfTbs.RowData(vsfTbs.Row))
                If lngBlockSize = 0 Then lngBlockSize = 8192
                
                If strSegment = "sys.free" Then
                    lblPrompt.Caption = "�Ѹ�ʽ���Ŀ��пռ䣬" & mrsExtents!blocks & "�飺��" & Round(mrsExtents!First_Block * 8192 / 1024 / 1024, 2) & _
                                        "M��" & Round(mrsExtents!Last_Block * lngBlockSize / 1024 / 1024, 2) & "M"
                Else
                    lblPrompt.Caption = mrsExtents!Segment_Type & "��" & strSegment & "��Extent_ID��" & mrsExtents!Extent_ID & "(" & mrsExtents!blocks & "�飬��" & _
                                        Round(mrsExtents!First_Block * lngBlockSize / 1024 / 1024, 2) & "M��" & Round(mrsExtents!Last_Block * lngBlockSize / 1024 / 1024, 2) & "M)"
                    lblPrompt.Tag = mrsExtents!Segment_Type
                End If
            Else
                lblPrompt.Caption = "δѡ�����ݿ顣"
            End If
            
            mrsExtents.Filter = "Segment_Name='" & Split(strSegment, ".")(1) & "' And Owner='" & Split(strSegment, ".")(0) & "' And Segment_Type='" & strSegment_Type & "'"
            For i = 1 To mrsExtents.RecordCount
                .Cell(flexcpBackColor, mrsExtents!Row, mrsExtents!Col) = &H8000000D     '��ɫ
                .Cell(flexcpForeColor, mrsExtents!Row, mrsExtents!Col) = &H80000005
                mrsExtents.MoveNext
            Next
        Else
            lblPrompt.Caption = "δѡ�����ݿ顣"
        End If
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub vsfExtents_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngRow As Long
    Dim lngCol As Long
    
    If Me.Visible = False Or vsfExtents.Redraw = flexRDNone Or vsfTbs.Enabled = False Then Exit Sub
    
    lngRow = vsfExtents.MouseRow
    lngCol = vsfExtents.MouseCol
    If lngRow > 0 And lngCol > 0 And Not mcolCells Is Nothing Then
       If (lngRow <> mlngRowPre Or lngCol <> mlngColPre) And lngRow <> vsfExtents.Row And lngCol <> vsfExtents.Col Then
           vsfExtents.ToolTipText = mcolCells("_" & lngRow & "_" & lngCol) & "(һ����Ԫ�����" & CONBLOCKS & "����)"
           mlngRowPre = lngRow
           mlngColPre = lngCol
       End If
    Else
        vsfExtents.ToolTipText = ""
    End If
End Sub

Private Sub vsfTbs_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    If Me.Visible And NewRowSel <> OldRowSel And vsfTbs.Redraw <> flexRDNone Then
        vsfTbs.Refresh
        Call LoadFiles(vsfTbs.TextMatrix(NewRowSel, vsfTbs.ColIndex("����")))

        If cboFiles.ListCount < 2 Then
            cboFiles.ListIndex = 0
        Else
            vsfExtents.Redraw = flexRDNone '���ⴥ���¼�vsfExtents_AfterRowColChange
            vsfExtents.Rows = vsfExtents.FixedRows
            vsfExtents.Redraw = flexRDDirect
            vsfExtents.ToolTipText = ""
            vsfExtents.Refresh
        End If
        
        Call LoadLobs(vsfTbs.TextMatrix(NewRowSel, vsfTbs.ColIndex("����")))
        
        If vsfTbs.Cell(flexcpData, NewRowSel, vsfTbs.ColIndex("��С")) = "NO" Then
            lblOptPrompt.Caption = "��ѡ��ռ��д�������������ΪNO�������ļ���"
        End If
    End If
End Sub

Private Sub LoadFiles(strTBS As String)
    Dim rsTmp As ADODB.Recordset, strSql  As String
    Dim i As Long, lngStart As Long
    
    strSql = "Select a.File_Name, a.File_Id, Round(a.Bytes / 1024 / 1024) As Fsize, Round(Nvl(Sum(b.Bytes),0) / 1024 / 1024) As Free_Size , a.autoextensible " & vbNewLine & _
            "From Dba_Data_Files A, Dba_Free_Space B" & vbNewLine & _
            "Where a.Tablespace_Name = [1] And a.File_Id = b.File_Id(+) And a.Tablespace_Name = b.Tablespace_Name(+) And a.Online_status in('ONLINE','SYSTEM')" & vbNewLine & _
            "Group By a.File_Name, a.File_Id, a.Bytes,a.autoextensible" & vbNewLine & _
            "Order By a.File_Id"

    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, strTBS)
    
    cboFiles.Clear
    cboFiles.Tag = ""
    For i = 1 To rsTmp.RecordCount
        cboFiles.AddItem rsTmp!FILE_NAME & "(ռ��" & rsTmp!fsize & "M,����" & rsTmp!Free_Size & "M" & IIf(rsTmp!autoextensible & "" <> "YES", ",���Զ���չ", "") & ")"
        cboFiles.ItemData(cboFiles.NewIndex) = Val(rsTmp!File_Id)
        rsTmp.MoveNext
    Next
    
    Exit Sub
errH:
    Call ErrCenter(strSql)
End Sub


Private Sub InitmrsExtents()
    
    Set mcolCells = New Collection
    
    Set mrsExtents = New ADODB.Recordset
    mrsExtents.Fields.Append "Row", adBigInt
    mrsExtents.Fields.Append "Col", adBigInt
    mrsExtents.Fields.Append "Owner", adVarChar, 20
    mrsExtents.Fields.Append "Segment_Name", adVarChar, 100
    mrsExtents.Fields.Append "Segment_Type", adVarChar, 20
    
    mrsExtents.Fields.Append "Extent_ID", adBigInt
    mrsExtents.Fields.Append "Blocks", adBigInt
    mrsExtents.Fields.Append "First_Block", adBigInt
    mrsExtents.Fields.Append "Last_Block", adBigInt
    
    mrsExtents.CursorLocation = adUseClient
    mrsExtents.LockType = adLockOptimistic
    mrsExtents.CursorType = adOpenStatic
    mrsExtents.Open
End Sub


Private Sub ResizeAll()
'���ܣ��������������ļ�
    Dim strErr As String
    Dim rsTmp As ADODB.Recordset, rsSize As ADODB.Recordset
    Dim lngBlockSize As Long, lngSumSize As Long
    
    If MsgBox("��ȷ��Ҫ�������������ļ���" & vbCrLf & vbCrLf & "��������ҵ������ڼ�ִ�У������أ�", vbYesNo + vbQuestion + vbDefaultButton2, "ȷ������") = vbNo Then
        lblOptPrompt.Caption = "������ȡ����"
        Call SetCommandEnable(1)
        Exit Sub
    End If
    
    Call SetCommandEnable(0)
    '��ȡBlock_size��С
    gstrSQL = "select value from v$parameter where name = 'db_block_size'"
    Set rsTmp = OpenSQLRecord(gstrSQL, Me.Caption)
    lngBlockSize = Val("" & rsTmp!Value)
    
    '��¼ִ�в������
    lblOptPrompt.Caption = "���ڲ�ѯ�������������ļ���"
    gstrSQL = "Select File_Name,'alter database datafile ''' || Trim(File_Name) || ''' resize ' || Ceil((Nvl(Hwm, 1) * " & lngBlockSize & ") / 1024 / 1024+10) || 'm' Cmd" & vbNewLine & _
            "From Dba_Data_Files A, (Select File_Id, Max(Block_Id + Blocks ) Hwm From Dba_Extents Group By File_Id) B" & vbNewLine & _
            "Where a.File_Id = b.File_Id(+) And Exists(Select 1 From Dba_Tablespaces C Where a.Tablespace_Name = c.Tablespace_Name And c.Status = 'ONLINE' And Contents != 'UNDO')" & vbNewLine & _
            "      And Ceil(Blocks * " & lngBlockSize & " / 1024 / 1024) - Ceil((Nvl(Hwm, 1) * " & lngBlockSize & ") / 1024 / 1024) > 10"
    Set rsTmp = OpenSQLRecord(gstrSQL, Me.Caption)
    If rsTmp.RecordCount = 0 Then
        Call MsgBox("û��Ҫ���������ļ���", vbInformation, "���������ļ�")
        lblOptPrompt.Caption = ""
        Call SetCommandEnable(1)
        Exit Sub
    Else
    
        If MsgBox("����" & rsTmp.RecordCount & "���������������ļ�����ȷ��Ҫ������Щ�����ļ���", vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    lblOptPrompt.Caption = "��ʼ��������������"
    
    'ִ�в���
    '1.��¼����ǰ�Ĵ�С
    gstrSQL = "Select Trunc(Sum(Bytes) / 1024 / 1024) Mb_Size From Dba_Data_Files"
    Set rsSize = OpenSQLRecord(gstrSQL, Me.Caption)
    lngSumSize = rsSize!Mb_Size
    On Error Resume Next
    strErr = ""
    While Not rsTmp.EOF
        lblOptPrompt.Caption = "����������" & rsTmp!FILE_NAME
        lblOptPrompt.Refresh
        gstrSQL = rsTmp!cmd
        gcnOracle.Execute gstrSQL
        
        If Err.Number <> 0 Then
            strErr = strErr & vbCrLf & rsTmp!cmd & "������" & Err.Description
            Err.Clear
        End If
        
        rsTmp.MoveNext
    Wend
    
    '2.��¼��������ܴ�С
    gstrSQL = "Select Trunc(Sum(Bytes) / 1024 / 1024) Mb_Size From Dba_Data_Files"
    Set rsSize = OpenSQLRecord(gstrSQL, Me.Caption)
    lngSumSize = lngSumSize - rsSize!Mb_Size

    lblOptPrompt.Caption = ""
        
    If strErr <> "" Then
        MsgBox "������Ϣ��" & strErr, vbExclamation
    Else
        lblOptPrompt.Caption = "������ɣ���������" & lngSumSize & "M�Ŀռ䡣"
    End If
    
    Call RefreshData
    
    Call SetCommandEnable(1)
End Sub

Private Sub ResizeTemp()
    Dim strError As String, strVersion As String, strTbsInfo As String
    Dim rsTmp As ADODB.Recordset
    Dim strSize As String, lngMax As Long
    
    strVersion = getVersion
    If strVersion = "" Then
        Exit Sub
    End If
    
    Call SetCommandEnable(0)

    On Error GoTo errH
    gstrSQL = "Select Tablespace_Name, File_Name, Trunc(Bytes / 1024 / 1024) Siz" & vbNewLine & _
            "From Dba_Temp_Files" & vbNewLine & _
            "Where Bytes / 1024 / 1024 > 10" & vbNewLine & _
            "Order By Tablespace_Name, File_Name"
    Set rsTmp = OpenSQLRecord(gstrSQL, Me.Caption)
    
    If rsTmp.RecordCount <> 0 Then
        While Not rsTmp.EOF
            strTbsInfo = strTbsInfo & rsTmp!FILE_NAME & "," & rsTmp!Siz & "M" & vbCrLf
            If rsTmp!Siz > lngMax Then lngMax = rsTmp!Siz
            rsTmp.MoveNext
        Wend
        strTbsInfo = "��ǰ��ʱ��ռ䣺" & vbCrLf & vbCrLf & strTbsInfo
        
        '��ȡ���ú�Ĵ�С
input_line:
        strSize = Trim(InputBox(strTbsInfo & vbCrLf & vbCrLf & "������������������ļ���С(��λM)��С�ڵ���ָ��ֵ�Ĳ���������������ҵ������ڼ�ִ�С�", "������ʱ��ռ�"))
        If strSize = "" Then
            Call SetCommandEnable(1)
            Exit Sub
        Else
            strError = ""
            If Not IsNumeric(strSize) Then
                strError = "��������������"
            ElseIf Val(strSize) <= 0 Then
                strError = "��������������������"
            ElseIf Val(strSize) >= lngMax Then
                strError = "����������С��" & lngMax & "�����֡�"
            ElseIf InStr(strSize, ".") > 0 Then
                strError = "���������벻��С��������"
            End If
            
            If strError <> "" Then
                MsgBox strError, vbInformation, gstrSysName
                GoTo input_line
            End If
        End If
        
        On Error Resume Next
        strError = ""
        strTbsInfo = ""
        lblOptPrompt.Caption = ""
        rsTmp.MoveFirst
        rsTmp.Filter = "Siz>" & strSize
        While Not rsTmp.EOF
            lblOptPrompt.Caption = "����������ʱ��ռ� " & rsTmp!Tablespace_Name & "��"
            lblOptPrompt.Refresh
            If strVersion = 11 Then
                'һ����ռ��ж�������ļ���11GR1�ǰ���ռ���������
                'Ҳ���԰������ļ��������: alter tablespace temp shrink tempfile '/u01/app/oracle/oradata/anqing/temp01.dbf' keep 300M;
                If rsTmp!Tablespace_Name <> strTbsInfo Then
                    strTbsInfo = rsTmp!Tablespace_Name
                    gstrSQL = "alter tablespace " & strTbsInfo & "  shrink space keep " & Val(strSize) & "M"
                    gcnOracle.Execute gstrSQL
                End If
            Else
                gstrSQL = "alter database tempfile '" & rsTmp!FILE_NAME & "'  resize " & Val(strSize) & "M"
                gcnOracle.Execute gstrSQL
            End If
            
            If Err <> 0 Then
                strError = strError & vbCrLf & rsTmp!FILE_NAME & vbCrLf & Err.Description
                Err.Clear
            End If
            rsTmp.MoveNext
        Wend
        
        If strError <> "" Then
            MsgBox "������ռ���� " & vbCrLf & strError & vbCrLf & "������ָ�������ļ��Ĵ�С����������ϵͳ��ִ��������", vbInformation, gstrSysName
        Else
            lblOptPrompt.Caption = "��ʱ��ռ�������ϣ�"
        End If
    Else
        MsgBox "��ǰû�д���10M����ʱ�����ļ�������Ҫ������"
    End If
    
    Call SetCommandEnable(1)
    Exit Sub
errH:
    Call ErrCenter(gstrSQL)
    Call SetCommandEnable(1)
End Sub


Public Sub SetNOParallel(ByVal cnThis As ADODB.Connection, ByVal bytType As Byte)
'���ܣ�����ִ�к���Զ�Ϊ�����������ϲ��ж����ԣ������ȡ������Ӱ�����SQL��ִ�мƻ�(ȫ��ɨ��+���в�ѯ������)
'������bytType��0=������1=��

    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim cmdTmp As New ADODB.Command
        
    If bytType = 0 Then
        strSql = "Select Owner || '.' || Index_Name As Index_Name From DBA_Indexes Where Degree Not In ('1', '0')"
    Else
        strSql = "Select Owner || '.' || Table_Name As Table_Name From DBA_Tables Where Degree !=('         1')"
    End If
    Set rsTmp = New ADODB.Recordset
    rsTmp.Open strSql, cnThis, adOpenKeyset, adLockReadOnly
        
    Set cmdTmp.ActiveConnection = cnThis
    cmdTmp.CommandType = adCmdText
    
    '�����������֯���ᱨ��ORA-25176: �����������ʹ�ô洢˵��
    On Error Resume Next
    
    While Not rsTmp.EOF
        If bytType = 0 Then
            strSql = "alter index " & rsTmp!Index_Name & " noparallel"
        Else
            strSql = "alter table " & rsTmp!Table_name & " noparallel"
        End If
        cmdTmp.CommandText = strSql
        
        cmdTmp.Execute
        
        rsTmp.MoveNext
    Wend
End Sub


Private Function GetDataFile(ByVal strTblSpace As String, ByRef strLocation As String) As String
    '���ܣ������ռ����ƣ���ȡ�����������ļ�
    '����Ĭ�������ļ�����Ϊ��ǰ�����ļ���+1
    '����ֵ�� �µ������ļ����ƣ�strLocation-���޸�Ϊ�����ļ�����·��
    
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim strFileName As String, strFilePath As String
    Dim i As Integer, strTmp As String
    
    On Error GoTo errH
    strSql = "Select FILE_NAME From Dba_Data_Files Where tablespace_name = [1] Order By File_Name Desc"
    Set rsTmp = OpenSQLRecord(strSql, "GetDataFile", strTblSpace)
    
    If rsTmp.RecordCount = 0 Then Exit Function

    strFilePath = rsTmp!FILE_NAME
    
    '�жϷ������Ƿ�ΪWINDOWS����,WINDOWΪ \,LinuxΪ /
    strFileName = Mid(strFilePath, InStrRev(strFilePath, IIf(InStr(strFilePath, "\") > 0, "\", "/")) + 1)
    strFilePath = Left(strFilePath, InStrRev(strFilePath, IIf(InStr(strFilePath, "\") > 0, "\", "/")))

    If InStr(strFileName, ".DBF") > 0 Then
        strFileName = Left(strFileName, InStrRev(strFileName, ".DBF") - 1)
        '�ж��ļ�����β���Ƿ�Ϊ���֣��������־�����ֶ�Ϊ01������+1
        For i = Len(strFileName) To 1 Step -1
            If InStr("0123456789", Mid(strFileName, i, 1)) > 0 Then
                strTmp = Mid(strFileName, i, 1) & strTmp
            Else
                Exit For
            End If
        Next
        If strTmp <> "" Then
            strFileName = Left(strFileName, InStr(1, strFileName, strTmp) - 1) & Format(Val(strTmp) + 1, "00") & ".DBF"
        Else
            strFileName = strFileName & Format(Val(strTmp) + 1, "00") & ".DBF"
        End If
    End If

    strLocation = strFilePath
    GetDataFile = strFileName
    Exit Function
errH:
    GetDataFile = ""
    ErrCenter
End Function
