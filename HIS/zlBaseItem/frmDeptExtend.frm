VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDeptExtend 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "������չ��Ϣ"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      Height          =   1700
      Index           =   0
      Left            =   120
      ScaleHeight     =   1629.323
      ScaleMode       =   0  'User
      ScaleWidth      =   2235
      TabIndex        =   3
      Top             =   4185
      Visible         =   0   'False
      Width           =   2300
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7275
      TabIndex        =   2
      Top             =   5670
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8445
      TabIndex        =   1
      Top             =   5670
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   3765
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9630
      _cx             =   16986
      _cy             =   6641
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
      ForeColorSel    =   -2147483643
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   255
      RowHeightMax    =   3000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDeptExtend.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
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
      VirtualData     =   0   'False
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
   Begin MSComDlg.CommonDialog cdl��Ƭ 
      Left            =   7845
      Top             =   4740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image img 
      Height          =   1700
      Index           =   0
      Left            =   2505
      Stretch         =   -1  'True
      Top             =   4245
      Visible         =   0   'False
      Width           =   2300
   End
End
Attribute VB_Name = "frmDeptExtend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngId As Long
Private mblnPro As Boolean '�Ƿ�����չ��Ŀ
Private mintType As Integer '1-��Ա��0-����
Private mblnEdit As Boolean '�����Ƿ��޸�
Private mblnͼƬ As Boolean '�Ƿ���ͼƬ
Private mblnͼƬ���� As Boolean '�Ƿ������ͼƬ
Private mintIndex As Integer 'ͼƬ��������
Private mint�༭״̬ As Integer '1-�ɱ༭��0-���ɱ༭
Private mintCountPic As Integer 'ͼƬ����
Private mstrName As String   '��������

Public Sub ShowMe(ByVal fraPar As Form, ByVal strID As String, ByVal strName As String, Optional ByVal intType As Integer, Optional int�༭״̬ As Integer)
    mlngId = Val(strID)
    mintType = intType
    mint�༭״̬ = int�༭״̬
    mstrName = strName
    
    If mintType = 1 Then
        Me.Caption = "��Ա��չ��Ϣ-" & mstrName
    Else
        Me.Caption = "������չ��Ϣ-" & mstrName
    End If
    
    Call initVSf(mlngId, mintType)
    
    If mblnPro Then Me.Show vbModal, fraPar
End Sub

Public Sub ReadBlob(ByVal lngId As Long, ByVal strName As String, ByVal intIndex As Integer, ByVal intType As Integer)
    '��ȡͼƬ
    Dim strTempFile As String
    
    '��ʼ��ͼƬλ�óߴ�
    img(intIndex).Left = pic(intIndex).ScaleLeft
    img(intIndex).Top = pic(intIndex).ScaleTop
    img(intIndex).Width = pic(intIndex).ScaleWidth
    img(intIndex).Height = pic(intIndex).ScaleHeight
    
    If intType = 1 Then '��Ա
        strTempFile = sys.Readlob(100, 20, lngId & "," & strName)
    Else
        strTempFile = sys.Readlob(100, 19, lngId & "," & strName)
    End If
    
    img(intIndex).Tag = ""
    img(intIndex).Picture = Nothing
    pic(intIndex).Picture = Nothing
    pic(intIndex).AutoRedraw = True
    
    '����ͼƬ
    If strTempFile <> "" Then
        img(intIndex).Tag = strTempFile
        img(intIndex).Picture = LoadPicture(strTempFile)
        pic(intIndex).PaintPicture img(intIndex).Picture, 0, 0, pic(intIndex).Width, pic(intIndex).Height
        
    End If
End Sub

Private Function SaveBlob(ByVal lngId As Long, ByVal strName As String, ByVal intIndex As Integer) As Boolean
    '����ͼƬ
    Dim blnOk As Boolean
    
    On Error GoTo ErrHandle
    
    If img(intIndex).Tag = "" Then
        If mintType = 1 Then '��Ա
            gstrSQL = "Update ��Ա��չ��Ϣ Set ͼƬ=Null Where ��Աid=" & lngId & " And ��Ŀ='" & strName & "'"
        Else
            gstrSQL = "Update ������չ��Ϣ Set ͼƬ=Null Where ����id=" & lngId & " And ��Ŀ='" & strName & "'"
        End If
        gcnOracle.Execute gstrSQL
        blnOk = True
    Else
        If mintType = 1 Then '��Ա
            blnOk = sys.Savelob(100, 20, lngId & "," & strName, img(intIndex).Tag)
        Else
            blnOk = sys.Savelob(100, 19, lngId & "," & strName, img(intIndex).Tag)
        End If
    End If
    
    SaveBlob = blnOk
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub initVSf(ByVal lngId As Long, Optional ByVal intType As Integer)
    '��ʼ��vsf����������
    Dim rsTemp As ADODB.Recordset
    Dim rs��Ϣ As ADODB.Recordset
    Dim intIndex As Integer
    Dim intRow As Integer
    Dim i As Integer
    Dim blnͼƬ As Boolean
    
    On Error GoTo ErrHandle
    
    If intType = 1 Then '��Ա
        gstrSQL = "Select ����, ����, Nvl(�Ƿ�ͼƬ, 0) As �Ƿ�ͼƬ From ��Ա��չ��Ŀ Order By ����"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ա��չ����")
    Else
        gstrSQL = "Select ����, ����, Nvl(�Ƿ�ͼƬ, 0) As �Ƿ�ͼƬ From ������չ��Ŀ Order By ����"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "������չ����")
    End If
    
    With VSFList
        .Rows = 1
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .RowHeightMin = 255
        .RowHeightMax = 2000
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        For i = 1 To img.Count - 1
            Unload img(i)
        Next
        For i = 1 To pic.Count - 1
            Unload pic(i)
        Next
        
        If rsTemp.RecordCount > 0 Then
            .redraw = flexRDNone
            Do While Not rsTemp.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTemp!����
                .TextMatrix(.Rows - 1, .ColIndex("��Ŀ")) = rsTemp!����
                
                '����
                If intType = 1 Then '��Ա
                    gstrSQL = "Select ���� From ��Ա��չ��Ϣ Where ��Աid=[1] And ��Ŀ=[2]"
                    Set rs��Ϣ = zlDatabase.OpenSQLRecord(gstrSQL, "��Ա��չ����", lngId, .TextMatrix(.Rows - 1, .ColIndex("��Ŀ")))
                Else
                    gstrSQL = "Select ���� From ������չ��Ϣ Where ����id=[1] And ��Ŀ=[2]"
                    Set rs��Ϣ = zlDatabase.OpenSQLRecord(gstrSQL, "������չ����", lngId, .TextMatrix(.Rows - 1, .ColIndex("��Ŀ")))
                End If
                
                If Not rs��Ϣ.EOF Then
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = IIF(IsNull(rs��Ϣ!����), "", rs��Ϣ!����)
                End If
                
                'ͼƬ
                If rsTemp!�Ƿ�ͼƬ = 1 Then
                    blnͼƬ = True
                    If mint�༭״̬ = 0 Then
                        intIndex = 0
                    Else
                        If intIndex <> 0 Then
                            Load img(intIndex)
                            Load pic(intIndex)
                        End If
                    End If
                    
                    Call ReadBlob(lngId, .TextMatrix(.Rows - 1, .ColIndex("��Ŀ")), intIndex, intType)
                    
                    .Cell(flexcpPicture, .Rows - 1, .ColIndex("ͼƬ"), .Rows - 1, .ColIndex("ͼƬ")) = pic(intIndex).Image
                    .RowHeight(.Rows - 1) = 1700
                    .TextMatrix(.Rows - 1, .ColIndex("ѡ��ͼƬ")) = "��"
                    .TextMatrix(.Rows - 1, .ColIndex("���ͼƬ")) = "��"
                    intIndex = intIndex + 1
                Else
                    If .TextMatrix(.Rows - 1, .ColIndex("����")) = "" Then
                        .TextMatrix(.Rows - 1, .ColIndex("ͼƬ")) = " "
                        .TextMatrix(.Rows - 1, .ColIndex("����")) = " "
                    Else
                        .TextMatrix(.Rows - 1, .ColIndex("ͼƬ")) = .TextMatrix(.Rows - 1, .ColIndex("����"))
                    End If
                    .MergeRow(.Rows - 1) = True
                    .RowHeight(.Rows - 1) = 1000
                End If
                
                rsTemp.MoveNext
            Loop
            
            If Not blnͼƬ Then
                .ColHidden(.ColIndex("ѡ��ͼƬ")) = True
                .ColHidden(.ColIndex("���ͼƬ")) = True
            End If
            
            .redraw = flexRDDirect
            mblnPro = True
            Call FS.ShowTipInfo(VSFList.hwnd, "")
        Else
            If mint�༭״̬ = 0 Then Exit Sub
            If intType = 1 Then '��Ա
                MsgBox "δ������Ա��չ��Ŀ���뵽�����ֵ�->��Ա����->��Ա��չ��Ŀ�����ã�", vbInformation, gstrSysName
            Else
                MsgBox "δ���ò�����չ��Ŀ���뵽�����ֵ�->��������->������չ��Ŀ�����ã�", vbInformation, gstrSysName
            End If
            mblnPro = False
        End If
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If SaveData Then
        MsgBox "����ɹ���", vbInformation, gstrSysName
        mblnEdit = False
        mblnͼƬ���� = False
        Unload Me
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If mint�༭״̬ = 0 Then
        cmdCancel.Visible = False
        cmdOK.Visible = False
        VSFList.ColHidden(VSFList.ColIndex("ѡ��ͼƬ")) = True
        VSFList.ColHidden(VSFList.ColIndex("���ͼƬ")) = True
        VSFList.Move 10, 10, Me.ScaleWidth - 20, Me.ScaleHeight - 20
        VSFList.ColWidth(VSFList.ColIndex("����")) = VSFList.Width - VSFList.ColWidth(VSFList.ColIndex("��Ŀ")) - VSFList.ColWidth(VSFList.ColIndex("ͼƬ")) - 400
        Exit Sub
    End If
    
    cmdCancel.Move Me.Width - cmdCancel.Width - 300, Me.ScaleHeight - cmdCancel.Height - 50
    cmdOK.Move cmdCancel.Left - cmdOK.Width - 10, cmdCancel.Top
    VSFList.Move 10, 10, Me.ScaleWidth - 20, Me.ScaleHeight - cmdOK.Height - 100
    VSFList.ColWidth(VSFList.ColIndex("����")) = VSFList.Width - VSFList.ColWidth(VSFList.ColIndex("��Ŀ")) - VSFList.ColWidth(VSFList.ColIndex("ͼƬ")) - VSFList.ColWidth(VSFList.ColIndex("ѡ��ͼƬ")) - VSFList.ColWidth(VSFList.ColIndex("���ͼƬ")) - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnEdit Or mblnͼƬ���� Then
        If MsgBox("���޸����ݻ�δ���棬�Ƿ�ȷ���˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
        Else
            mblnEdit = False
            mblnͼƬ���� = False
        End If
    End If
End Sub

Private Sub vsfList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With VSFList
        If .TextMatrix(.Row, .ColIndex("ѡ��ͼƬ")) = "" Then
            If Trim(.TextMatrix(Row, .ColIndex("ͼƬ"))) = "" Then
                .TextMatrix(Row, .ColIndex("ͼƬ")) = " "
                .TextMatrix(Row, .ColIndex("����")) = " "
            Else
                .TextMatrix(Row, .ColIndex("����")) = .TextMatrix(Row, .ColIndex("ͼƬ"))
            End If
        End If
    End With
End Sub

Private Sub vsfList_ChangeEdit()
    mblnEdit = True
End Sub

Private Sub vsfList_EnterCell()
    If mint�༭״̬ = 0 Then Exit Sub
    With VSFList
        If .TextMatrix(.Row, .ColIndex("ѡ��ͼƬ")) = "" Then
            If .Col = .ColIndex("ͼƬ") Then
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
            End If
        Else
            If .Col = .ColIndex("����") Then
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
            End If
        End If
    End With
End Sub

Private Sub VSFList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack Then Exit Sub
    With VSFList
        If LenB(StrConv(.EditText + Chr(KeyAscii), vbFromUnicode)) > 1000 Then
            KeyAscii = 0
        End If
    End With
    
End Sub

Private Sub vsfList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    Dim dblHeight As Double
    Dim dblWidth As Double
    
    With VSFList
        If .Rows = 1 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("ѡ��ͼƬ")) = "" Then
            If .Col = .ColIndex("ͼƬ") Then
                Call FS.ShowTipInfo(VSFList.hwnd, Trim(.TextMatrix(.Row, .ColIndex("ͼƬ"))), True)
            Else
                Call FS.ShowTipInfo(VSFList.hwnd, "")
            End If
        Else
            If .Col = .ColIndex("����") Then
                Call FS.ShowTipInfo(VSFList.hwnd, Trim(.TextMatrix(.Row, .ColIndex("����"))), True)
            Else
                Call FS.ShowTipInfo(VSFList.hwnd, "")
            End If
            
            If .Col = .ColIndex("ѡ��ͼƬ") Or .Col = .ColIndex("���ͼƬ") Then
                For i = 0 To .Rows - 1
                    dblHeight = dblHeight + .RowHeight(i)
                Next
                
                For i = 0 To .Cols - 1
                    If .ColHidden(i) = False Then
                        dblWidth = dblWidth + .ColWidth(i)
                    End If
                Next
                
                If X < dblWidth And Y > .RowHeight(0) And Y < dblHeight Then
                    If .Col = .ColIndex("ѡ��ͼƬ") Then
                        Call SelectPic
                    Else
                        Call ClearPic
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub SelectPic()
    'ѡ��ͼƬ
    Dim intIndex As Integer
    
    With cdl��Ƭ
        .CancelError = True
        .Filter = "ͼƬ�ļ�(*.bmp,*.gif,*.jpg)|*.bmp;*.gif;*.jpg"
        
        On Error Resume Next
        .ShowOpen
        intIndex = GetPicStaition(VSFList.Row)
        If Err <> 0 Then
            'ûѡ���ļ�
            Err.Clear
        Else
            img(intIndex).Picture = LoadPicture(.FileName)
            pic(intIndex).PaintPicture img(intIndex).Picture, 0, 0, pic(intIndex).Width, pic(intIndex).Height
            VSFList.Cell(flexcpPicture, VSFList.Row, VSFList.ColIndex("ͼƬ"), VSFList.Row, VSFList.ColIndex("ͼƬ")) = pic(intIndex).Image
            
            If Err <> 0 Then
                MsgBox "ͼƬ�ļ���Ч�����ļ������ڣ�", vbInformation, gstrSysName
                Exit Sub
            End If
            img(intIndex).Tag = .FileName
            mblnͼƬ = True
            mblnͼƬ���� = True
        End If
    End With
End Sub

Private Function GetPicStaition(ByVal intCRow As Integer) As Integer
    '��ȡ����ͼƬ����
    Dim intRow As Integer
    
    With VSFList
        mintCountPic = -1
        For intRow = 1 To intCRow
            If .TextMatrix(intRow, .ColIndex("ѡ��ͼƬ")) = "��" Then
                mintCountPic = mintCountPic + 1
            End If
        Next
        GetPicStaition = mintCountPic
    End With
End Function

Private Sub ClearPic()
    '���ͼƬ
    Dim intIndex As Integer
    
    intIndex = GetPicStaition(VSFList.Row)
    
    If img(intIndex).Tag = "" Then Exit Sub
    
    mblnͼƬ = False
    mblnͼƬ���� = True
    img(intIndex).Tag = ""
    img(intIndex).Picture = Nothing
    pic(intIndex).Picture = Nothing
    VSFList.Cell(flexcpPicture, VSFList.Row, VSFList.ColIndex("ͼƬ"), VSFList.Row, VSFList.ColIndex("ͼƬ")) = pic(intIndex).Image
End Sub

Private Function SaveData() As Boolean
    '��������
    Dim blnTran As Boolean
    Dim intRow As Integer
    Dim arrSQL As Variant
    
    On Error GoTo ErrHandle
    
    SaveData = False
    arrSQL = Array()
    
    With VSFList
        For intRow = 1 To .Rows - 1
            If Check��Ŀ(.TextMatrix(intRow, .ColIndex("��Ŀ"))) Then
                If LenB(StrConv(.TextMatrix(intRow, .ColIndex("����")), vbFromUnicode)) > 1000 Then
                    MsgBox "��" & intRow & "�С�" & .TextMatrix(intRow, .ColIndex("��Ŀ")) & "���������ݴ���1000���ַ������������룡", vbInformation, gstrSysName
                    .Col = .Col
                    .Row = intRow
                    .SetFocus
                    Exit Function
                End If
                If mintType = 1 Then '��Ա
                    gstrSQL = "Zl_��Ա��չ��Ϣ_Delete(" & mlngId & ",'" & .TextMatrix(intRow, .ColIndex("��Ŀ")) & "')"

                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = gstrSQL
                    
                    gstrSQL = "Zl_��Ա��չ��Ϣ_Insert(" & mlngId & ",'" & .TextMatrix(intRow, .ColIndex("��Ŀ")) & "','" & .TextMatrix(intRow, .ColIndex("����")) & "')"

                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = gstrSQL
                Else
                    gstrSQL = "Zl_������չ��Ϣ_Delete(" & mlngId & ",'" & .TextMatrix(intRow, .ColIndex("��Ŀ")) & "')"

                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = gstrSQL
                    
                    gstrSQL = "Zl_������չ��Ϣ_Insert(" & mlngId & ",'" & .TextMatrix(intRow, .ColIndex("��Ŀ")) & "','" & .TextMatrix(intRow, .ColIndex("����")) & "')"

                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = gstrSQL
                End If
            End If
        Next
    
        gcnOracle.BeginTrans: blnTran = True
        For intRow = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(intRow)), "SaveData")
        Next
        
        For intRow = 1 To .Rows - 1
            If Check��Ŀ(.TextMatrix(intRow, .ColIndex("��Ŀ"))) Then
                If .TextMatrix(intRow, .ColIndex("ѡ��ͼƬ")) <> "" Then
                    Call SaveBlob(mlngId, .TextMatrix(intRow, .ColIndex("��Ŀ")), GetPicStaition(intRow))
                End If
            End If
        Next
        gcnOracle.CommitTrans: blnTran = False
    End With
    
    SaveData = True
    
    Exit Function
ErrHandle:
    If blnTran = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Check��Ŀ(ByVal strName As String) As Boolean
    Dim rsTemp As Recordset
    
    Check��Ŀ = False
    If mintType = 1 Then '��Ա
        gstrSQL = "Select 1 From ��Ա��չ��Ŀ Where ���� = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ա��չ����", strName)
    Else
        gstrSQL = "Select 1 From ������չ��Ŀ Where ���� = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "������չ����", strName)
    End If
    
    If Not rsTemp.EOF Then Check��Ŀ = True
    
End Function
