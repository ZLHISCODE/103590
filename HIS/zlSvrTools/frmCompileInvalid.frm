VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCompileInvalid 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "������Ч����"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmCompileInvalid.frx":0000
   ScaleHeight     =   4665
   ScaleWidth      =   6720
   WindowState     =   2  'Maximized
   Begin VSFlex8Ctl.VSFlexGrid vsError 
      Height          =   480
      Left            =   255
      TabIndex        =   1
      Top             =   3405
      Width           =   5490
      _cx             =   9684
      _cy             =   847
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmCompileInvalid.frx":04F9
      ScrollTrack     =   -1  'True
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
      ExplorerBar     =   5
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
   Begin MSComctlLib.ProgressBar pgbCompile 
      Height          =   165
      Left            =   270
      TabIndex        =   4
      Top             =   4365
      Visible         =   0   'False
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdCompile 
      Caption         =   "����(&P)"
      Height          =   350
      Left            =   3330
      TabIndex        =   3
      Top             =   3915
      Width           =   1100
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "���(&C)"
      Height          =   350
      Left            =   330
      TabIndex        =   2
      Top             =   3930
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsObj 
      Height          =   2250
      Left            =   255
      TabIndex        =   0
      Top             =   1155
      Width           =   5490
      _cx             =   9684
      _cy             =   3969
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   225
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmCompileInvalid.frx":0536
      ScrollTrack     =   -1  'True
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
      ExplorerBar     =   5
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
      Begin MSComctlLib.ImageList imgObj 
         Left            =   1725
         Top             =   1110
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   10
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCompileInvalid.frx":0618
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCompileInvalid.frx":0772
               Key             =   "����"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCompileInvalid.frx":08CC
               Key             =   "����"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCompileInvalid.frx":0A26
               Key             =   "������"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCompileInvalid.frx":0B80
               Key             =   "��ͼ"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCompileInvalid.frx":0CDA
               Key             =   "�ﻯ��ͼ"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCompileInvalid.frx":0E34
               Key             =   "��"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCompileInvalid.frx":0F8E
               Key             =   "����"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCompileInvalid.frx":10E8
               Key             =   "����"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCompileInvalid.frx":1242
               Key             =   "������"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label lbl˵�� 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���ʧЧ����϶࣬��Ӧ�ٶȻ��½�������ϵͳ����ʱʹ�á�"
      Height          =   180
      Left            =   1185
      TabIndex        =   6
      Top             =   615
      Width           =   5400
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   210
      Picture         =   "frmCompileInvalid.frx":139C
      Top             =   525
      Width           =   480
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������Ч����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   5
      Top             =   120
      Width           =   1440
   End
End
Attribute VB_Name = "frmCompileInvalid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum Col_OBJ
    colͼ�� = 0
    col���� = 1
    Col���� = 2
    colʱ�� = 3
    colOwner = 4
    colName = 5
    colType = 6
    colFind = 7
End Enum
Private mrsObjRef As ADODB.Recordset

Private Sub cmdCheck_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    
    '��ȡ���е���Ч����
    strSQL = _
        " Select Owner,Object_Name,Object_Type,Created,Last_DDL_Time From ALL_Objects " & _
        " Where Object_Type In('PROCEDURE','FUNCTION','VIEW','MATERIALIZED VIEW','TRIGGER','PACKAGE','PACKAGE BODY','TYPE','TYPE BODY')" & _
        " And Object_Name Not Like 'BIN$%' And Status = 'INVALID'" & _
        " Order by Object_Type,Decode(Owner,User,0,1),Owner,Object_Name"
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
        
    With vsObj
        .Redraw = flexRDNone
        .Rows = .FixedRows
        If rsTmp.EOF Then
            .Rows = .FixedRows + 1
        Else
            .Rows = rsTmp.RecordCount + 1
        End If
        For i = 1 To rsTmp.RecordCount
            '������
            .TextMatrix(i, colOwner) = rsTmp!Owner
            .TextMatrix(i, colName) = rsTmp!Object_Name
            .TextMatrix(i, colType) = rsTmp!Object_Type
            .TextMatrix(i, colFind) = rsTmp!Owner & "_" & rsTmp!Object_Name & "_" & rsTmp!Object_Type
            
            '����
            .TextMatrix(i, col����) = IIf(rsTmp!Owner <> UCase(gstrUserName), rsTmp!Owner & ".", "") & rsTmp!Object_Name
            Select Case rsTmp!Object_Type
            Case "PROCEDURE"
                .TextMatrix(i, Col����) = "����"
            Case "FUNCTION"
                .TextMatrix(i, Col����) = "����"
            Case "VIEW"
                .TextMatrix(i, Col����) = "��ͼ"
            Case "MATERIALIZED VIEW"
                .TextMatrix(i, Col����) = "�ﻯ��ͼ"
            Case "TRIGGER"
                .TextMatrix(i, Col����) = "������"
            Case "PACKAGE"
                .TextMatrix(i, Col����) = "��"
            Case "PACKAGE BODY"
                .TextMatrix(i, Col����) = "����"
            Case "TYPE"
                .TextMatrix(i, Col����) = "����"
            Case "TYPE BODY"
                .TextMatrix(i, Col����) = "������"
            End Select
            
            '������ʱ��
            .TextMatrix(i, colʱ��) = Format(rsTmp!Last_DDL_Time, "yyyy-MM-dd HH:mm:ss")
            
            'ͼ��
            Set .Cell(flexcpPicture, i, colͼ��) = imgObj.Overlay(.TextMatrix(i, Col����), 1)
            rsTmp.MoveNext
        Next
        .Cell(flexcpPictureAlignment, .FixedRows, colͼ��, .Rows - 1, colͼ��) = 4
        .Row = .FixedRows: .Col = col����
        .Redraw = flexRDDirect
    End With
    vsError.Rows = 0
    vsError.AddItem "" & vbTab & "���ҵ� " & rsTmp.RecordCount & " ����Ч����" & IIf(rsTmp.RecordCount > 0, ",��ʹ��[����]����������Щ��Ч����", "")
    
    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbExclamation, gstrSysName
End Sub

Private Sub cmdCompile_Click()
    Dim strSQL As String, i As Long
    Dim blnDo As Boolean, lngTopRow As Long
        
    If vsObj.TextMatrix(vsObj.FixedRows, colName) = "" Then Exit Sub
    
    Screen.MousePointer = 11
    pgbCompile.Min = 0
    pgbCompile.Max = vsObj.Rows - 1
    pgbCompile.Value = 0
    pgbCompile.Visible = True
    vsError.Rows = 0: vsError.Rows = 1
    Call Form_Resize
    Me.Refresh
    
    '��ȡ�������ù�ϵ��
    If mrsObjRef Is Nothing Then blnDo = True
    If Not blnDo Then
        '��ע����ɼ�¼���ر�
        If mrsObjRef.State = 0 Then blnDo = True
    End If
    If blnDo Then
        vsError.TextMatrix(0, 1) = "���ڳ�ʼ����������..."
        vsError.Refresh
        
        On Error GoTo errH
        strSQL = _
            " Select Owner,Name,Type,Referenced_Owner,Referenced_Name,Referenced_Type From ALL_Dependencies" & _
            " Where Type In('PROCEDURE','FUNCTION','VIEW','MATERIALIZED VIEW','TRIGGER','PACKAGE','PACKAGE BODY','TYPE','TYPE BODY')" & _
            " And Referenced_Type In('PROCEDURE','FUNCTION','VIEW','MATERIALIZED VIEW','TRIGGER','PACKAGE','PACKAGE BODY','TYPE','TYPE BODY')" & _
            " And Not(Name=Referenced_Name And Owner=Referenced_Owner And Type=Referenced_Type)" & _
            " And Name Not Like 'BIN$%' And Referenced_Name Not Like 'BIN$%'"
        Set mrsObjRef = New ADODB.Recordset
        mrsObjRef.CursorLocation = adUseClient
        mrsObjRef.Open strSQL, gcnOracle, adOpenKeyset
        On Error GoTo 0
    End If
    
    'ִ�б���
    With vsObj
        lngTopRow = .TopRow
        For i = .FixedRows To .Rows - 1
            Call .ShowCell(i, col����): .Refresh
            Call CompileObject(i)
            pgbCompile.Value = i
        Next
    End With
    vsObj.TopRow = lngTopRow
    Call vsObj_AfterRowColChange(-1, -1, vsObj.Row, vsObj.Col)
    
    pgbCompile.Visible = False
    Call Form_Resize

    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbExclamation, gstrSysName
End Sub

Private Sub CompileObject(ByVal lngRow As Long)
'���ܣ�����ָ���е���Ч����
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, lngFind As Long
    Dim arrObjRef As Variant, i As Long
    
    With vsObj
        vsError.TextMatrix(0, 1) = "���ڱ���" & .TextMatrix(lngRow, Col����) & " """ & .TextMatrix(lngRow, col����) & """ ..."
        vsError.Refresh
        'RowData:0-δ����,1-����ɹ�,2-����ʧ��
        If .RowData(lngRow) <> 0 Then Exit Sub
        
        '�������õĶ���(�ݹ�)
        arrObjRef = Array()
        mrsObjRef.Filter = "Owner='" & .TextMatrix(lngRow, colOwner) & "' And Name='" & .TextMatrix(lngRow, colName) & "' And Type='" & .TextMatrix(lngRow, colType) & "'"
        If Not mrsObjRef.EOF Then
            ReDim arrObjRef(mrsObjRef.RecordCount - 1) As String
        End If
        For i = 1 To mrsObjRef.RecordCount
            arrObjRef(i - 1) = CStr(mrsObjRef!Referenced_Owner & "_" & mrsObjRef!Referenced_Name & "_" & mrsObjRef!Referenced_Type)
            mrsObjRef.MoveNext
        Next
        For i = 0 To UBound(arrObjRef)
            lngFind = .FindRow(CStr(arrObjRef(i)), , colFind)
            If lngFind <> -1 Then
                Call CompileObject(lngFind)
            End If
        Next
                
        '���뵱ǰ�ж���
        strSQL = ""
        Select Case .TextMatrix(lngRow, colType)
        Case "PROCEDURE"
            strSQL = "ALTER PROCEDURE " & .TextMatrix(lngRow, colOwner) & "." & .TextMatrix(lngRow, colName) & " COMPILE"
        Case "FUNCTION"
            strSQL = "ALTER FUNCTION " & .TextMatrix(lngRow, colOwner) & "." & .TextMatrix(lngRow, colName) & " COMPILE"
        Case "VIEW"
            strSQL = "ALTER VIEW " & .TextMatrix(lngRow, colOwner) & "." & .TextMatrix(lngRow, colName) & " COMPILE"
        Case "MATERIALIZED VIEW"
            strSQL = "ALTER MATERIALIZED VIEW " & .TextMatrix(lngRow, colOwner) & "." & .TextMatrix(lngRow, colName) & " COMPILE"
        Case "TRIGGER"
            strSQL = "ALTER TRIGGER " & .TextMatrix(lngRow, colOwner) & "." & .TextMatrix(lngRow, colName) & " COMPILE"
        Case "PACKAGE"
            strSQL = "ALTER PACKAGE " & .TextMatrix(lngRow, colOwner) & "." & .TextMatrix(lngRow, colName) & " COMPILE"
        Case "PACKAGE BODY"
            strSQL = "ALTER PACKAGE " & .TextMatrix(lngRow, colOwner) & "." & .TextMatrix(lngRow, colName) & " COMPILE BODY"
        Case "TYPE"
            strSQL = "ALTER TYPE " & .TextMatrix(lngRow, colOwner) & "." & .TextMatrix(lngRow, colName) & " COMPILE"
        Case "TYPE BODY"
            strSQL = "ALTER TYPE " & .TextMatrix(lngRow, colOwner) & "." & .TextMatrix(lngRow, colName) & " COMPILE BODY"
        End Select
        If strSQL <> "" Then
            On Error Resume Next
            gcnOracle.Execute strSQL
            If Err.Number <> 0 Then
                .Cell(flexcpData, lngRow, colʱ��) = Err.Description
                .RowData(lngRow) = 2 '����δͨ��
            Else
                '������������Err.Number=0,Oracle.Errors.Count>0
                '1.[Microsoft][ODBC driver for Oracle]�����Ĺ��̻���������б������
                '2.û�и���Ĵ���
                strSQL = "Select Status,Last_DDL_Time From All_Objects Where Owner='" & .TextMatrix(lngRow, colOwner) & "' And Object_Name='" & .TextMatrix(lngRow, colName) & "' And Object_Type='" & .TextMatrix(lngRow, colType) & "'"
                Set rsTemp = New ADODB.Recordset
                rsTemp.CursorLocation = adUseClient
                rsTemp.Open strSQL, gcnOracle, adOpenKeyset
                If Nvl(rsTemp!Status) <> "VALID" Then
                    .RowData(lngRow) = 2 '����δͨ��
                Else
                    .RowData(lngRow) = 1 '����ͨ��
                    Set .Cell(flexcpPicture, lngRow, colͼ��) = imgObj.ListImages(.TextMatrix(lngRow, 2)).Picture
                End If
                .TextMatrix(lngRow, colʱ��) = Format(rsTemp!Last_DDL_Time, "yyyy-MM-dd HH:mm:ss")
            End If
            Err.Clear: On Error GoTo 0
            .Refresh
        End If
    End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    vsObj.ColWidth(colOwner) = 0
    vsObj.ColWidth(colType) = 0
    vsObj.ColWidth(colName) = 0
    vsObj.ColWidth(colFind) = 0
End Sub

Private Sub Form_Resize()
    Dim lngW As Long, i As Long
    
    On Error Resume Next
    
    vsObj.Width = Me.ScaleWidth - vsObj.Left * 2
    vsObj.Height = Me.ScaleHeight - vsObj.Top - IIf(pgbCompile.Visible, pgbCompile.Height, 0) - cmdCompile.Height - vsError.Height - 225
    
    vsError.Left = vsObj.Left
    vsError.Top = vsObj.Top + vsObj.Height + 15
    vsError.Width = vsObj.Width
    
    pgbCompile.Left = 0
    pgbCompile.Top = Me.ScaleHeight - pgbCompile.Height
    pgbCompile.Width = Me.ScaleWidth
    
    cmdCheck.Top = vsError.Top + vsError.Height + (IIf(pgbCompile.Visible, pgbCompile.Top, Me.ScaleHeight) - vsError.Top - vsError.Height - cmdCheck.Height) / 2
    cmdCompile.Top = cmdCheck.Top
    cmdCompile.Left = Me.ScaleWidth - cmdCompile.Width - cmdCheck.Left
    
    For i = 0 To vsObj.Cols - 1
        If i <> 1 And Not vsObj.ColHidden(i) Then
            lngW = lngW + vsObj.ColWidth(i) + 15
        End If
    Next
    lngW = lngW + 250
    If vsObj.Width - lngW < 1000 Then
        vsObj.ColWidth(1) = 1000
    Else
        vsObj.ColWidth(1) = vsObj.Width - lngW
    End If
    
    Me.Refresh
End Sub

Public Function RefreshData()
    'Nothing
End Function

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = True
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim bytR As Byte, i As Long
    Dim lngRow As Long, lngCol As Long
    
    If vsObj.TextMatrix(1, colName) = "" Then Exit Sub
    
    '��ͷ
    objOut.Title.Text = "��Ч�����嵥"
    objOut.Title.Font.name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    Set objRow = New zlTabAppRow
    objRow.Add "ʱ�䣺" & Format(CurrentDate(), "yyyy-MM-dd HH:mm:ss")
    objOut.UnderAppRows.Add objRow
    
    '����
    Set objOut.Body = vsObj
    
    '���
    vsObj.Redraw = False
    lngRow = vsObj.Row: lngCol = vsObj.Col
        
    If bytMode = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytMode
    End If
    
    vsObj.Row = lngRow: vsObj.Col = lngCol
    vsObj.Redraw = True
End Sub

Private Sub vsError_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'���ܣ���ʾ�����Ķ�����������Ϣ��ʾ
    Dim lngRow As Long
    
    With vsError
        lngRow = .MouseRow
        If lngRow >= 0 And lngRow <= .Rows - 1 Then
            .ToolTipText = .Cell(flexcpData, lngRow, 1)
        Else
            .ToolTipText = ""
        End If
    End With
End Sub

Private Sub vsObj_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'���ܣ���ʾ������������Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If NewRow = OldRow Or Not (NewRow >= vsObj.FixedRows And NewRow <= vsObj.Rows - 1) Then Exit Sub
    If vsObj.TextMatrix(NewRow, colName) = "" Then Exit Sub
            
    strSQL = "Select Owner,Name,Type,Sequence,Line,Text From ALL_Errors" & _
        " Where Owner='" & vsObj.TextMatrix(NewRow, colOwner) & "'" & _
        " And Name='" & vsObj.TextMatrix(NewRow, colName) & "'" & _
        " And Type='" & vsObj.TextMatrix(NewRow, colType) & "'" & _
        " Order By Sequence"
    On Error GoTo errH
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    vsError.Redraw = flexRDNone
    vsError.Rows = 0
    Do While Not rsTmp.EOF
        vsError.AddItem ""
        vsError.TextMatrix(vsError.Rows - 1, 0) = rsTmp!Line
        vsError.TextMatrix(vsError.Rows - 1, 1) = Replace(rsTmp!Text, vbLf, " ")
        vsError.Cell(flexcpData, vsError.Rows - 1, 1) = CStr(rsTmp!Text)
        rsTmp.MoveNext
    Loop
    If vsError.Rows > 0 Then vsError.Row = 0
    vsError.Redraw = flexRDDirect
    Exit Sub
errH:
    MsgBox Err.Description, vbExclamation, gstrSysName
End Sub

Private Sub vsObj_BeforeSort(ByVal Col As Long, Order As Integer)
    If Col = 0 Then Order = 0
End Sub

Private Sub vsObj_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'���ܣ���һ�ΰ�����ʾ����������
    Dim lngRow As Long
    
    With vsObj
        lngRow = .MouseRow
        If lngRow >= .FixedRows And lngRow <= .Rows - 1 Then
            If vsError.Rows = 1 Then
                If vsError.TextMatrix(0, 0) = "" Then
                    Call vsObj_AfterRowColChange(-1, -1, .Row, .Col)
                End If
            End If
        End If
    End With
End Sub

Private Sub vsObj_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'���ܣ���ʾ��������еĴ�����ʾ
    Dim lngRow As Long
    
    With vsObj
        lngRow = .MouseRow
        If lngRow >= .FixedRows And lngRow <= .Rows - 1 Then
            If .Cell(flexcpData, lngRow, colʱ��) <> "" Then
                .ToolTipText = .Cell(flexcpData, lngRow, colʱ��)
            Else
                .ToolTipText = ""
            End If
        End If
    End With
End Sub
