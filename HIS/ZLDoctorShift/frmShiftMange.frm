VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmShiftMange 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ι���"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6735
   Icon            =   "frmShiftMange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   6735
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdTypeOK 
      Appearance      =   0  'Flat
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3960
      TabIndex        =   2
      Top             =   3840
      Width           =   1100
   End
   Begin VB.CommandButton cmdTypeCancel 
      Appearance      =   0  'Flat
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5355
      TabIndex        =   1
      Top             =   3840
      Width           =   1100
   End
   Begin VB.ComboBox cboDept 
      Height          =   300
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   5505
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfType 
      Height          =   2655
      Left            =   960
      TabIndex        =   3
      Top             =   600
      Width           =   5490
      _cx             =   9684
      _cy             =   4683
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmShiftMange.frx":6852
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   240
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShiftMange.frx":6934
            Key             =   "unCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShiftMange.frx":6ECE
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShiftMange.frx":7468
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShiftMange.frx":DCCA
            Key             =   "add"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShiftMange.frx":1452C
            Key             =   "Person"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShiftMange.frx":14F3E
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShiftMange.frx":1B7A0
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShiftMange.frx":1C1B2
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ʱ���ʽ����18:00          24:00����00:00��ʾ"
      Height          =   180
      Left            =   960
      TabIndex        =   6
      Top             =   3360
      Width           =   4050
   End
   Begin VB.Label lblDept 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "��    ��"
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   300
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "�����Ϣ"
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   720
   End
End
Attribute VB_Name = "frmShiftMange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrDeptId As String
Private mlngDeptID As Long
Private mblnOk As Boolean

Public Function ShowMe(ByVal strDeptID As String, ByVal lngDeptId As Long) As Boolean

    mstrDeptId = strDeptID
    mlngDeptID = lngDeptId
    Me.Show 1
    ShowMe = mblnOk
End Function

Private Sub cboDept_Click()
    mlngDeptID = cboDept.ItemData(cboDept.ListIndex)
    Call LoadData
End Sub

Private Sub cmdTypeCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdTypeOK_Click()
    Dim i As Long
    Dim arrSQL() As Variant, varTemp As Variant
    Dim strTemp As String
    Dim blnBegin As Boolean
    Dim lngDeptId As Long
    
    arrSQL = Array()
    If CheckTypeData = False Then Exit Sub
    lngDeptId = cboDept.ItemData(cboDept.ListIndex)
    With vsfType
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("�������")) <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                If .TextMatrix(i, .ColIndex("ԭ�������")) = "" Then
                    arrSQL(UBound(arrSQL)) = "Zl_ҽ��ֵ����_Edit(0," & lngDeptId & ",'" & .TextMatrix(i, .ColIndex("�������")) & "',null," & _
                        "to_date('" & Format(.TextMatrix(i, .ColIndex("��ʼʱ��")), "hh:mm") & "','hh24:mi'),to_date('" & Format(.TextMatrix(i, .ColIndex("����ʱ��")), "hh:mm") & "','hh24:mi'))"
                Else
                    arrSQL(UBound(arrSQL)) = "Zl_ҽ��ֵ����_Edit(1," & lngDeptId & ",'" & .TextMatrix(i, .ColIndex("�������")) & "','" & _
                        .TextMatrix(i, .ColIndex("ԭ�������")) & "',to_date('" & Format(.TextMatrix(i, .ColIndex("��ʼʱ��")), "hh:mm") & "','hh24:mi'),to_date('" & Format(.TextMatrix(i, .ColIndex("����ʱ��")), "hh:mm") & "','hh24:mi'))"
                End If
            End If
        Next
    End With
    gcnOracle.BeginTrans
    blnBegin = True
    On Error GoTo ErrHand
    If vsfType.Tag <> "" Then
        '�������ɾ���Ѿ��еİ�Σ�����ɾ��
        varTemp = Split(vsfType.Tag, "<�ָ���>")
        For i = 0 To UBound(varTemp)
            strTemp = "Zl_ҽ��ֵ����_Edit(2," & lngDeptId & ",'" & varTemp(i) & "')"
            If strTemp <> "" Then
                Call zlDatabase.ExecuteProcedure(strTemp, Me.Caption)
            End If
        Next
    End If
    For i = LBound(arrSQL) To UBound(arrSQL)
        strTemp = arrSQL(i)
        If strTemp <> "" Then
            Call zlDatabase.ExecuteProcedure(strTemp, Me.Caption)
        End If
    Next
    gcnOracle.CommitTrans
    mblnOk = True
    Unload Me
    Exit Sub
ErrHand:
    If blnBegin Then gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    
    mblnOk = False
    Set rsTemp = GetDeptName(mstrDeptId)
    Call zlControl.CboAddData(cboDept, rsTemp)
    For i = 0 To cboDept.ListCount - 1
        If cboDept.ItemData(i) = mlngDeptID Then
            cboDept.ListIndex = i
            Exit For
        Else
            If i = cboDept.ListCount - 1 Then
                cboDept.ListIndex = 0
            End If
        End If
    Next
    Call LoadData
End Sub

Private Sub LoadData()
'���ر��İ������
    Dim rsTemp As ADODB.Recordset

    Set rsTemp = GetShiftType(1, mlngDeptID)
    vsfType.Redraw = flexRDNone
    vsfType.Rows = rsTemp.RecordCount + 1
    Do While Not rsTemp.EOF
        vsfType.TextMatrix(rsTemp.AbsolutePosition, vsfType.ColIndex("ԭ�������")) = rsTemp!�������
        vsfType.TextMatrix(rsTemp.AbsolutePosition, vsfType.ColIndex("�������")) = rsTemp!�������
        vsfType.TextMatrix(rsTemp.AbsolutePosition, vsfType.ColIndex("��ʼʱ��")) = rsTemp!��ʼʱ�� & ""
        vsfType.TextMatrix(rsTemp.AbsolutePosition, vsfType.ColIndex("����ʱ��")) = rsTemp!����ʱ�� & ""
        rsTemp.MoveNext
    Loop
    If vsfType.Rows = 1 Then
        vsfType.Rows = 2
    End If
    If vsfType.Rows > 7 Then
        vsfType.ColWidth(vsfType.ColIndex("�������")) = 1575
    Else
        vsfType.ColWidth(vsfType.ColIndex("�������")) = 1830
    End If
    Call vsfType_AfterRowColChange(2, 1, 1, 1)
    vsfType.Redraw = flexRDDirect
End Sub

Private Sub vsfType_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Or NewRow < 1 Then Exit Sub
    With vsfType
        .Cell(flexcpPicture, NewRow, .ColIndex("ɾ����")) = imgList.ListImages("delete").Picture
        .Cell(flexcpPicture, NewRow, .ColIndex("������")) = imgList.ListImages("add").Picture
        If OldRow < .Rows Then
            .Cell(flexcpPicture, OldRow, .ColIndex("ɾ����")) = ""
            .Cell(flexcpPicture, OldRow, .ColIndex("������")) = ""
        End If
    End With
End Sub

Private Sub vsfType_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsfType.ColIndex("������") Or Col = vsfType.ColIndex("ɾ����") Then Cancel = True
End Sub

Private Sub vsfType_Click()
    Dim strName As String
    
    With vsfType
        strName = .TextMatrix(.Row, .ColIndex("�������"))
        If .Col = .ColIndex("������") Then
            If strName <> "" And .TextMatrix(.Row, .ColIndex("��ʼʱ��")) <> "" And .TextMatrix(.Row, .ColIndex("����ʱ��")) <> "" Then
                .AddItem "", .Row + 1
                .Row = .Row + 1
                .Col = .ColIndex("�������")
                Call .ShowCell(.Row, .Col)
            End If
        ElseIf .Col = .ColIndex("ɾ����") Then
            Call vsfType_KeyDown(vbKeyDelete, 0)
        End If
    End With
End Sub

Private Sub vsfType_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strName As String
    
    With vsfType
        strName = .TextMatrix(.Row, .ColIndex("�������"))
        If KeyCode = vbKeyDelete Then
            If strName = "" And .TextMatrix(.Row, .ColIndex("��ʼʱ��")) = "" And .TextMatrix(.Row, .ColIndex("����ʱ��")) = "" Then
            Else
                If MsgBox("��ȷ��ɾ��" & IIf(strName = "", "��" & .Row & "��", "����Ϊ��" & strName & "��") & "�İ����Ϣ��?", vbDefaultButton2 + vbYesNo, Me.Caption) = vbNo Then
                    Exit Sub
                End If
            End If
            If .TextMatrix(.Row, .ColIndex("ԭ�������")) <> "" Then
                vsfType.Tag = IIf(vsfType.Tag = "", "", vsfType.Tag & "<�ָ���>") & strName
            End If
            If .Rows <= 2 Then
                .TextMatrix(1, .ColIndex("�������")) = ""
                .TextMatrix(1, .ColIndex("��ʼʱ��")) = ""
                .TextMatrix(1, .ColIndex("����ʱ��")) = ""
                .ShowCell .Row, .ColIndex("�������")
            Else
                .RemoveItem .Row
            End If
            Call vsfType_AfterRowColChange(0, 0, vsfType.Row, 1)
        End If
    End With
End Sub

Private Function CheckTypeData() As Boolean
'������������ݵĺ�����
    Dim i As Long
    Dim strNames As String, strName As String, strTemp As String
    Dim strBegins As String, strEnds As String, strBegin As String, strEnd As String
    
    With vsfType
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("�������")) <> "" Then
                strBegins = strBegins & "," & Format(.TextMatrix(i, .ColIndex("��ʼʱ��")), "HH:MM")
                strEnds = strEnds & "," & Format(.TextMatrix(i, .ColIndex("����ʱ��")), "HH:MM")
            End If
        Next
        For i = 1 To .Rows - 1
            strTemp = .TextMatrix(i, .ColIndex("�������"))
            If zlstr.ActualLen(strTemp) > 10 Then
                MsgBox "ֵ�����Ʋ��ó���5�����֣����飡", vbExclamation, Me.Caption
                .ShowCell i, .ColIndex("�������")
                Exit Function
            End If
            If strTemp <> "" Then
                strBegin = Format(.TextMatrix(i, .ColIndex("��ʼʱ��")), "HH:MM")
                strEnd = Format(.TextMatrix(i, .ColIndex("����ʱ��")), "HH:MM")
                If Not CheckTime(strBegin, i, .Col) Then Exit Function
                If Not CheckTime(strEnd, i, .Col) Then Exit Function
                If InStr(strEnds & ",", "," & strBegin & ",") = 0 Then
                    MsgBox "��" & strTemp & "���Ŀ�ʼʱ��û�ж�Ӧ����ͬ����ʱ�䣬���飡", vbExclamation, Me.Caption
                    .ShowCell i, .ColIndex("��ʼʱ��")
                    Exit Function
                End If
                If InStr(strBegins & ",", "," & strEnd & ",") = 0 Then
                    MsgBox "��" & strTemp & "���Ľ���ʱ��û�ж�Ӧ����ͬ��ʼʱ�䣬���飡", vbExclamation, Me.Caption
                    .ShowCell i, .ColIndex("����ʱ��")
                    Exit Function
                End If
                If InStr(strNames & ",", "," & strTemp & ",") = 0 Then
                    strNames = strNames & "," & strTemp
                Else
                    strName = IIf(strName = "", "", strTemp & "��")
                End If
                If strName <> "" Then
                    MsgBox "����������ظ�������ƣ�" & strName, vbExclamation, Me.Caption
                    Exit Function
                End If
            End If
        Next
    End With
    CheckTypeData = True
End Function
Private Function CheckTime(ByVal strDate As String, ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
'���ʱ���ʽ�Ƿ���ȷ
    Dim varTemp As Variant
    
    If strDate = "" Then
        MsgBox "��" & strDate & "��ʱ�䲻��Ϊ�գ������룡", vbExclamation, Me.Caption
        vsfType.ShowCell lngRow, lngCol
        Exit Function
    End If
    If InStr(strDate, ":") = 0 Then
        MsgBox "��" & strDate & "��ʱ��û��ð�ţ����������룡", vbExclamation, Me.Caption
        vsfType.ShowCell lngRow, lngCol
        Exit Function
    End If
    varTemp = Split(strDate, ":")
    If varTemp(0) > 23 Then
        MsgBox "��" & strDate & "����Сʱ���ܳ���23�����������룡", vbExclamation, Me.Caption
        vsfType.ShowCell lngRow, lngCol
        Exit Function
    End If
    If Len(varTemp(0)) > 2 Then
        MsgBox "��" & strDate & "����Сʱ���Ȳ��ܴ���2λ�����������룡", vbExclamation, Me.Caption
        vsfType.ShowCell lngRow, lngCol
        Exit Function
    End If
    If varTemp(1) > 59 Then
        MsgBox "��" & strDate & "���ķ��Ӳ��ܳ���59�����������룡", vbExclamation, Me.Caption
        vsfType.ShowCell lngRow, lngCol
        Exit Function
    End If
    If Len(varTemp(1)) > 2 Then
        MsgBox "��" & strDate & "���ķ��ӳ��Ȳ��ܴ���2λ�����������룡", vbExclamation, Me.Caption
        vsfType.ShowCell lngRow, lngCol
        Exit Function
    End If
    CheckTime = True
End Function

Private Sub vsfType_KeyPress(KeyAscii As Integer)

    With vsfType
        If .Col < .ColIndex("����ʱ��") Then
            .Col = .Col + 1
            .ShowCell .Row, .Col
        Else
            zlCommFun.PressKey (vbKeyTab)
        End If
    End With
End Sub

Private Sub vsfType_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsfType
        If Col = .ColIndex("��ʼʱ��") Or Col = .ColIndex("����ʱ��") Then
            If KeyAscii = vbKeyBack Then Exit Sub
            If KeyAscii = vbKeyReturn Then Exit Sub
            If KeyAscii = Asc("��") Then KeyAscii = Asc(":")
            
            If KeyAscii = Asc(":") And InStr(1, .EditText, ":") > 0 Then KeyAscii = 0: Exit Sub
            If KeyAscii = Asc(":") Then Exit Sub
        
            If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
                KeyAscii = 0
                Exit Sub
            End If
        ElseIf Col = .ColIndex("�������") Then
            If KeyAscii = Asc("'") Then KeyAscii = 0
        End If
    End With
End Sub


