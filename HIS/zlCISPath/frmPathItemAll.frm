VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPathItemAll 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ȫ·����Ŀ����"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8760
   Icon            =   "frmPathItemAll.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   8760
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6720
      Width           =   8760
      Begin VB.CommandButton cmdDel 
         Caption         =   "ɾ��(&D)"
         Height          =   350
         Left            =   240
         TabIndex        =   13
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   7440
         TabIndex        =   12
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   6240
         TabIndex        =   11
         Top             =   120
         Width           =   1100
      End
      Begin VB.Timer tmrInfo 
         Interval        =   1000
         Left            =   3000
         Top             =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   7
         X1              =   0
         X2              =   12720
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   6
         X1              =   0
         X2              =   12720
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Frame fraInput 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   8295
      Begin VB.CommandButton cmdInput 
         Caption         =   "��"
         Height          =   285
         Left            =   4125
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   30
         Width           =   285
      End
      Begin VB.TextBox txtInput 
         Height          =   360
         Left            =   840
         TabIndex        =   8
         Top             =   0
         Width           =   3615
      End
      Begin VB.Label lblInput 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Ŀ����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   0
         TabIndex        =   9
         Top             =   90
         Width           =   720
      End
      Begin VB.Label lblTip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʾ��Ϣ:δ�ҵ�������Ŀ"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   4560
         TabIndex        =   7
         Top             =   75
         Width           =   3735
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   8760
      TabIndex        =   0
      Top             =   0
      Width           =   8760
      Begin VB.Image imgInfo 
         Height          =   720
         Index           =   1
         Left            =   120
         Picture         =   "frmPathItemAll.frx":6852
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2��¼����롢���롢���Ʋ���������Ŀ"
         Height          =   180
         Index           =   1
         Left            =   1680
         TabIndex        =   3
         Top             =   630
         Width           =   3150
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   12840
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1��������Ŀ�������ڵ�ǰ·���汾�������õ�������Ŀ"
         Height          =   180
         Index           =   0
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   4410
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "˵����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   1095
         TabIndex        =   1
         Top             =   165
         Width           =   585
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsItem 
      Height          =   5100
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   8715
      _cx             =   15372
      _cy             =   8996
      Appearance      =   2
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
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16444122
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   330
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathItemAll.frx":70DA
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
End
Attribute VB_Name = "frmPathItemAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrPrivs As String
Private mlngPathID As Long
Private mlngVerID As Long
Private mstrLike As String      '����ƥ�䷽ʽ
Private mint���� As Integer     '����ƥ�䷽ʽ��0-ƴ��,1-���
Private msngTime As Single
Private mstrDelIds As String     '��¼ɾ����������Ŀ
Private mblnView As Boolean     '������鿴���ɱ༭

Private Enum E_COL
    COL_���� = 0
    COL_����
End Enum
'-------------------------------------------------------------------------------------------------------
Public Sub ShowMe(frmParent As Object, ByVal strPrivs As String, ByVal lngPathID As Long, Optional ByVal lngVerID As Long, Optional ByVal blnView As Boolean)
'����:��ں���
'����:������
'   lngPathID-·��ID
'   lngVerID-�汾��
    mstrPrivs = strPrivs
    mlngPathID = lngPathID
    mlngVerID = lngVerID
    mstrDelIds = ""
    mblnView = blnView
    
    Me.Show 1, frmParent
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function DeleteItem() As Boolean
    Dim lngRow As Long
    
    With vsItem
        lngRow = .Row
        If lngRow <= 0 Then Exit Function
        If Val(.RowData(lngRow)) > 0 Then
            If .Cell(flexcpData, lngRow, COL_����) = 1 Then
                If MsgBox("ȷ��Ҫɾ����" & .TextMatrix(lngRow, COL_����) & "����", vbQuestion + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                    Exit Function
                End If
                mstrDelIds = mstrDelIds & "," & .RowData(lngRow)
            End If
            .RemoveItem lngRow
            If .Rows - 1 = 0 Then .AddItem ""
        End If
    End With
    DeleteItem = True
End Function

Private Sub cmdDel_Click()
    Call DeleteItem
End Sub

Private Sub cmdOK_Click()
    If Not mblnView Then
        Call SaveData
    End If
    
    Unload Me
End Sub

Private Sub cmdInput_Click()
    Dim strIds As String
    strIds = GetInputIDs()
    Call GetItem(1, "", strIds)
End Sub

Private Sub Form_Activate()
    If mblnView Then
        Call vsItem.SetFocus
    Else
        Call txtInput.SetFocus
    End If
End Sub

Private Sub Form_Load()
    mstrLike = IIf(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "") '����ƥ�䷽ʽ
    mint���� = Val(zlDatabase.GetPara("���뷽ʽ")) '����ƥ�䷽ʽ��0-ƴ��,1-���
    '��ʼ�����Ʊ��
    Call Grid.Init(vsItem, "����,1305,1;������Ŀ,5000,1")
    lblTip.Caption = "": tmrInfo.Enabled = False
    Call ReadItem
    If mblnView Then
        Me.Caption = "ȫ·����Ŀ"
        cmdCancel.Caption = "�˳�"
    End If
End Sub

Private Function ReadItem() As Boolean
    Dim strSql As String
    Dim i As Long
    Dim rsTmp As ADODB.Recordset
    
    strSql = "Select b.Id, b.����, b.����" & vbNewLine & _
            "From ·��ͨ��������Ŀ A, ������ĿĿ¼ B" & vbNewLine & _
            "Where a.������Ŀid = b.Id And a.·��id = [1] And a.�汾�� = [2]"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngPathID, mlngVerID)
    With vsItem
        .Enabled = flexEDKbdMouse
        .Rows = rsTmp.RecordCount + 1
        If .Rows = 1 Then .Rows = 2 'û������ʱ��ʾһ�п���
        For i = 1 To rsTmp.RecordCount
            .RowData(i) = Val(rsTmp!ID & "")
            .TextMatrix(i, COL_����) = rsTmp!���� & ""
            .Cell(flexcpData, i, COL_����) = 1  'ԭʼ����
            .TextMatrix(i, COL_����) = rsTmp!���� & ""
            rsTmp.MoveNext
        Next
    End With
    ReadItem = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveData() As Boolean
    Dim i As Long
    Dim strIds As String
    Dim blnTrans As Boolean
    Dim arrSQL As Variant
    
    On Error GoTo errH
    With vsItem
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, COL_����) = 2 And Val(.RowData(i)) > 0 Then
                strIds = strIds & "," & .RowData(i)
            End If
        Next
    End With
    strIds = Mid(strIds, 2)
    mstrDelIds = Mid(mstrDelIds, 2)
    
    arrSQL = Array()
    
    If mstrDelIds <> "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_·��ͨ��������Ŀ_Update(" & mlngPathID & "," & mlngVerID & ",'" & mstrDelIds & "',1)"
    End If
    
    If strIds <> "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_·��ͨ��������Ŀ_Update(" & mlngPathID & "," & mlngVerID & ",'" & strIds & "',0)"
    End If
    If UBound(arrSQL) >= 0 Then
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
        Next
        gcnOracle.CommitTrans: blnTrans = False
    End If
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub GetItem(ByVal bytFunc As Byte, ByVal strInput As String, ByVal strIds As String)
    Dim strSql As String
    Dim strFilter As String
    Dim rsTmp As ADODB.Recordset
    Dim varArr As Variant, varTemp As Variant
    Dim blnCancel As Boolean
    Dim i As Long, k As Long
    
    
    On Error GoTo errH
    With vsItem
        If bytFunc = 0 Then
            If strInput = "" Then Exit Sub
            
            strFilter = " And (D.���� Like [3] Or F.���� Like [4] Or F.���� Like [4]) And F.����=[5]"
            If IsNumeric(strInput) Then
                '1X.����ȫ������ʱֻƥ�����'
                strFilter = " And D.���� Like [3] And F.����=[5]"
            ElseIf zlCommFun.IsCharAlpha(strInput) Then
                'X1.����ȫ����ĸʱֻƥ�����
                strFilter = " And F.���� Like [4] And F.����=[5]"
            ElseIf zlCommFun.IsCharChinese(strInput) Then
                '��������,��ֻƥ������
                strFilter = " And F.���� Like [4] And F.����=[5]"
            End If
        End If
        varArr = Array("", "", "", "", "", "", "", "", "", "")
        If strIds <> "" Then
            varTemp = zlStr.Str2Array(strIds, ",", 4000)
            For i = LBound(varTemp) To UBound(varTemp)
                If i > UBound(varArr) Then Exit For
                varArr(i) = varTemp(i)
                strSql = strSql & " And Not Exists (Select 1 From Table(f_Str2list([" & (6 + i) & "])) where Column_Value = d.Id) "
            Next
        End If
        strSql = "Select Distinct d.Id, d.����, d.����, G.���� As ���" & vbNewLine & _
                "From �ٴ�·����Ŀ A, �ٴ�·��ҽ�� B, ·��ҽ������ C, ������ĿĿ¼ D,������Ŀ���� F,������Ŀ��� G" & vbNewLine & _
                "Where a.Id = b.·����Ŀid And b.ҽ������id = c.Id And c.������Ŀid = d.Id And D.ID=F.������ĿID And D.���=G.���� And a.·��id = [1] And a.�汾�� = [2] And" & vbNewLine & _
                "       ( Not (d.��� = 'E' And Instr(',2,3,4,6,8,9,', ',' || d.�������� || ',') > 0) Or (d.��� = 'E' And d.�������� = '8' And d.����Ӧ�� = 1)) And" & vbNewLine & _
                "      Not (Instr(',G,F,D,', ',' || d.��� || ',') > 0 And NVL(c.���id,0) <> 0)" & strFilter & strSql & vbNewLine & _
                "Order By G.����"
        If bytFunc = 1 Then
            strSql = "Select a.Id, a.����, a.����, a.��� From (" & strSql & ") A Where Rownum < 100"
        End If
        Set rsTmp = FS.ShowSQLSelectEx(Me, txtInput, strSql, 0, "������Ŀѡ��", False, "", "", False, False, True, blnCancel, _
            True, True, True, "bytSize=1#ColSet=�п�����|����,1200,0;����,4200,0;���,1000,0|������ʾ|����", mlngPathID, mlngVerID, UCase(strInput) & "%", mstrLike & UCase(strInput) & "%", mint���� + 1, _
                CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                lblTip.Caption = "û�п��õ�������Ŀ�����ȵ�·�������á�"
                msngTime = Timer
                tmrInfo.Enabled = True
            End If
            Call zlControl.TxtSelAll(txtInput)
            txtInput.SetFocus
            Exit Sub
        End If
        txtInput.Text = ""
        For i = 1 To rsTmp.RecordCount
            k = .FindRow(Val(rsTmp!ID & ""))
            If k = -1 Then
                If .RowData(.Rows - 1) > 0 Then .Rows = .Rows + 1
                .RowData(.Rows - 1) = Val(rsTmp!ID & "")
                .Cell(flexcpData, .Rows - 1, COL_����) = 2   '����
                .TextMatrix(.Rows - 1, COL_����) = rsTmp!���� & ""
                .TextMatrix(.Rows - 1, COL_����) = rsTmp!���� & ""
            End If
            rsTmp.MoveNext
        Next
        .Row = .Rows - 1
        .ShowCell .Row, COL_����
    End With

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    If mblnView Then
        fraInput.Visible = False: picInfo.Visible = False
        vsItem.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - picBottom.Height
        cmdOK.Visible = False
        cmdDel.Visible = False
    End If
End Sub

Private Sub tmrInfo_Timer()
    If Timer - msngTime > 5 Then
        lblTip.Caption = ""
        tmrInfo.Enabled = False
    End If
End Sub

Private Sub txtInput_GotFocus()
    Call zlControl.TxtSelAll(txtInput)
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call GetItem(0, txtInput.Text, GetInputIDs())
    End If
End Sub

Private Sub vsItem_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mblnView Then Cancel = False: Exit Sub
    If Col = COL_���� Then
        Cancel = True  '���벻����༭
    ElseIf Col = COL_���� Then
        Cancel = True
    End If
End Sub

Private Sub vsItem_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
    If mblnView Then Exit Sub
    With vsItem
        lngRow = .Row
        If lngRow <= 0 Then Exit Sub
        If KeyCode = vbKeyDelete Then
            Call DeleteItem
        End If
        If KeyCode <> vbKeyReturn Then Exit Sub
        If .Editable = flexEDNone Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        If .Col = COL_���� And .Row = .Rows - 1 Then
            .Row = 1: .Col = COL_����
        ElseIf .Col = COL_���� And .Row < .Rows - 1 Then
            .Row = .Row + 1: .Col = COL_����
        Else
            .Col = .Col + 1
        End If
        .ShowCell .Row, .Col
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetInputIDs
' Author    : YWJ
' Date      : 2019-05-07
' Purpose   : ��ȡ�Ѿ�¼���ID
'---------------------------------------------------------------------------------------
'
Private Function GetInputIDs() As String
    Dim i As Long
    Dim strTemp As String
    With vsItem
        For i = 1 To .Rows - 1
            strTemp = strTemp & "," & .RowData(i)
        Next
    End With
    GetInputIDs = Mid(strTemp, 2)
End Function
