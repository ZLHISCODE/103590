VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form frmMediDosage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�䷽ԭ��"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8505
   Icon            =   "frmMediDosage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picSpilt 
      BackColor       =   &H80000005&
      Height          =   4455
      Left            =   3480
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4455
      ScaleWidth      =   15
      TabIndex        =   0
      Top             =   360
      Width           =   15
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfVariety 
      Height          =   3165
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   3015
      _cx             =   5318
      _cy             =   5583
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
      BackColorSel    =   16769992
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMediDosage.frx":6852
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
      ExplorerBar     =   1
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
   Begin VSFlex8Ctl.VSFlexGrid vsfSpec 
      Height          =   3165
      Left            =   4680
      TabIndex        =   4
      Top             =   600
      Width           =   3015
      _cx             =   5318
      _cy             =   5583
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
      BackColorSel    =   16769992
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMediDosage.frx":68C7
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
      ExplorerBar     =   1
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
   Begin VB.Label lblSpec 
      AutoSize        =   -1  'True
      Caption         =   "�����ǹ��"
      Height          =   180
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   900
   End
   Begin VB.Label lblVar 
      AutoSize        =   -1  'True
      Caption         =   "������Ʒ��"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "frmMediDosage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintDosageType As Integer '�䷽����
Private mstrName As String  'ѡȡ��ҩƷ
Private mstrFind As String  '������ѯ

'���
Private Enum menuSpec
    ID = 0
    ���� = 1
    ���� = 2
    ��� = 3
    ���㵥λ = 4
    ���� = 5
    �Ƿ��� = 6
    �������� = 7
    ������� = 8
    ҩ��ID = 9
    Cols = 10
End Enum

'Ʒ��
Private Enum menuVar
    ID = 0
    ���� = 1
    ���� = 2
    ���㵥λ = 3
    ������� = 4
    Cols = 5
End Enum

Public Sub ShowMe(ByVal intDosageType As Integer, ByVal frmPar As Form, ByVal strFind As String, ByRef strName As String)
    Select Case intDosageType
        Case 0
            mintDosageType = 3  '������̬
        Case 1
            mintDosageType = 0  'ɢװ
        Case 2
            mintDosageType = 1  '��Ƭ
        Case 3
            mintDosageType = 2  '����
    End Select
    
    mstrFind = strFind
    mstrName = ""
    
    Me.Show vbModal, frmPar
    strName = mstrName
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub GetVarInfo()
    '�õ�Ʒ����Ϣ
    Dim rsTemp As Recordset
    Dim intRow As Integer
    
    On Error GoTo ErrHand
    
    If mintDosageType <> 3 Then '������̬
        gstrSql = " b.��ҩ��̬= " & mintDosageType & " and "
    Else
        gstrSql = ""
    End If
    
    If Trim(mstrFind) = "" Then
        gstrSql = "Select a.Id, a.����, a.����,a.���㵥λ, Decode(a.�������, 1, '����', 2, 'סԺ', 3, '�����סԺ', 4, '���', '�������ڲ���') �������" & vbNewLine & _
            "From ������ĿĿ¼ A" & vbNewLine & _
            "Where Exists (Select 1 From ҩƷ��� B Where " & gstrSql & " a.Id = b.ҩ��id) And a.��� = '7' And Sysdate Between a.����ʱ�� And a.����ʱ��"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "getVarInfo")
    Else
        gstrSql = "Select Distinct a.Id, a.����, a.����, a.���㵥λ, Decode(a.�������, 1, '����', 2, 'סԺ', 3, '�����סԺ', 4, '���', '�������ڲ���') �������" & vbNewLine & _
            "From ������ĿĿ¼ A, ������Ŀ���� N" & vbNewLine & _
            "Where Exists (Select 1 From ҩƷ��� B Where " & gstrSql & " a.Id = b.ҩ��id) And a.Id = n.������Ŀid And a.��� = '7' And" & vbNewLine & _
            "      (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
            "      (a.���� Like [1] Or n.���� Like [2] Or n.���� Like [2])" & vbNewLine & _
            "Order By a.����"

        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "getVarInfo", mstrFind & "%", gstrMatch & mstrFind & "%")
    End If
    
    intRow = 1
    vsfVariety.Rows = rsTemp.RecordCount + 1
    Do While Not rsTemp.EOF
        With vsfVariety
            .TextMatrix(intRow, menuVar.ID) = rsTemp!ID
            .TextMatrix(intRow, menuVar.����) = rsTemp!����
            .TextMatrix(intRow, menuVar.����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            .TextMatrix(intRow, menuVar.���㵥λ) = IIf(IsNull(rsTemp!���㵥λ), "", rsTemp!���㵥λ)
            .TextMatrix(intRow, menuVar.�������) = rsTemp!�������
            intRow = intRow + 1
            rsTemp.MoveNext
        End With
    Loop
    
    vsfVariety.Cell(flexcpAlignment, 0, 0, 0, vsfVariety.Cols - 1) = flexAlignCenterCenter
    If vsfVariety.Rows > 1 Then
        vsfVariety.Cell(flexcpAlignment, 1, 0, vsfVariety.Rows - 1, vsfVariety.Cols - 1) = flexAlignLeftCenter
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub GetSpecInfo()
    '�õ������Ϣ
    Dim rsTemp As Recordset
    Dim intRow As Integer
    
    On Error GoTo ErrHand
    
    vsfSpec.Rows = 1
    
    If mintDosageType <> 3 Then '������̬
        gstrSql = " and b.��ҩ��̬=" & mintDosageType
    End If
    
    gstrSql = "Select a.Id, a.����, a.����, a.���,a.���㵥λ, a.����, Decode(a.�Ƿ���, 0, '����', 'ʱ��') As �Ƿ���, a.��������," & vbNewLine & _
        "       Decode(a.�������, 1, '����', 2, 'סԺ', 3, '�����סԺ', '�������ڲ���') �������, b.ҩ��id " & vbNewLine & _
        "From �շ���ĿĿ¼ A, ҩƷ��� B" & vbNewLine & _
        "Where a.Id = b.ҩƷid And  b.ҩ��id = [1]" & gstrSql & " and a.��� = '7' And Sysdate Between a.����ʱ�� And a.����ʱ��" & vbNewLine & _
        "Order By a.Id"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "getVarInfo", vsfVariety.TextMatrix(vsfVariety.Row, menuVar.ID))
    
    intRow = 1
    vsfSpec.Rows = rsTemp.RecordCount + 1
    Do While Not rsTemp.EOF
        With vsfSpec
            .TextMatrix(intRow, menuSpec.ID) = rsTemp!ID
            .TextMatrix(intRow, menuSpec.����) = rsTemp!����
            .TextMatrix(intRow, menuSpec.����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            .TextMatrix(intRow, menuSpec.���) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
            .TextMatrix(intRow, menuSpec.���㵥λ) = IIf(IsNull(rsTemp!���㵥λ), "", rsTemp!���㵥λ)
            .TextMatrix(intRow, menuSpec.����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            .TextMatrix(intRow, menuSpec.�Ƿ���) = rsTemp!�Ƿ���
            .TextMatrix(intRow, menuSpec.��������) = IIf(IsNull(rsTemp!��������), "", rsTemp!��������)
            .TextMatrix(intRow, menuSpec.�������) = rsTemp!�������
            .TextMatrix(intRow, menuSpec.ҩ��ID) = rsTemp!ҩ��ID
            intRow = intRow + 1
            
            rsTemp.MoveNext
        End With
    Loop
    vsfSpec.Cell(flexcpAlignment, 0, 0, 0, vsfSpec.Cols - 1) = flexAlignCenterCenter
    If vsfSpec.Rows > 1 Then
        vsfSpec.Cell(flexcpAlignment, 1, 0, vsfSpec.Rows - 1, vsfSpec.Cols - 1) = flexAlignLeftCenter
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Call initControl    '��ʼ���ؼ���С
    Call InitData   '��ʼ��vsf�ؼ���ͷ����ɫ
    Call GetVarInfo 'ΪƷ���б�ֵ
    Call SetCaption '���ñ���
End Sub

Private Sub SetCaption()
    '���ñ���
    Dim strCaption As String
    
    Select Case mintDosageType
        Case 0
            strCaption = "�䷽ԭ��(ɢװ)"
        Case 1
            strCaption = "�䷽ԭ��(��Ƭ)"
        Case 2
            strCaption = "�䷽ԭ��(����)"
        Case 3
            strCaption = "�䷽ԭ��(������̬)"
    End Select
    Me.Caption = strCaption
End Sub

Private Sub initControl()
    '��ʼ���ؼ�λ�ú�״̬
    Select Case mintDosageType
    Case 0 ' ɢװ
        lblVar.Left = 50
        vsfVariety.Top = lblVar.Height + lblVar.Top + 100
        vsfVariety.Width = Me.Width / 3 - picSpilt.Width
        vsfVariety.Height = Me.Height - lblVar.Height - lblVar.Top - 500
        vsfVariety.Left = 50
        picSpilt.Left = vsfVariety.Left + vsfVariety.Width
        picSpilt.Top = 0
        picSpilt.Height = Me.Height
        vsfSpec.Width = Me.Width - picSpilt.Width - picSpilt.Left - 100
        vsfSpec.Top = vsfVariety.Top
        vsfSpec.Height = Me.Height - lblVar.Height - lblVar.Top - 500
        vsfSpec.Left = picSpilt.Left + picSpilt.Width
        lblSpec.Left = vsfSpec.Left
    Case 1, 2, 3 '��Ƭ������,������̬
        vsfSpec.Visible = False
        lblSpec.Visible = False
        lblVar.Left = 50
        vsfVariety.Top = lblVar.Height + lblVar.Top + 100
        vsfVariety.Left = 50
        vsfVariety.Width = Me.ScaleWidth
        vsfVariety.Height = Me.Height - lblVar.Height - lblVar.Top - 500
        picSpilt.Visible = False
    End Select
End Sub

Private Sub picSpilt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = 1 Then
        If vsfVariety.Width + x < 200 Then Exit Sub
        If vsfSpec.Width + x < 200 Then Exit Sub
            picSpilt.Left = picSpilt.Left + x
            vsfVariety.Width = vsfVariety.Width + x
            vsfSpec.Width = vsfSpec.Width - x
            vsfSpec.Left = vsfSpec.Left + x
            lblSpec.Left = vsfSpec.Left
    End If
End Sub

Private Sub InitData()
    '��ʼ����������
    Dim intCol As Integer
    Dim intRow As Integer
    
    With vsfSpec
        .SelectionMode = flexSelectionByRow
        .AllowSelection = False '���ܶ�ѡ
        .ExplorerBar = flexExSortShowAndMove '������ƶ�
        .AllowUserResizing = flexResizeBoth  '���Ըı����п��
        .Cols = menuSpec.Cols
        .Rows = 1
        .TextMatrix(0, menuSpec.ID) = "ҩƷid"
        .TextMatrix(0, menuSpec.����) = "����"
        .TextMatrix(0, menuSpec.����) = "����"
        .TextMatrix(0, menuSpec.���) = "���"
        .TextMatrix(0, menuSpec.���㵥λ) = "���㵥λ"
        .TextMatrix(0, menuSpec.����) = "����"
        .TextMatrix(0, menuSpec.�Ƿ���) = "�Ƿ���"
        .TextMatrix(0, menuSpec.��������) = "��������"
        .TextMatrix(0, menuSpec.�������) = "�������"
        .TextMatrix(0, menuSpec.ҩ��ID) = "ҩ��id"
        
        .ColHidden(menuSpec.ID) = True
        .ColWidth(menuSpec.����) = 800
        .ColWidth(menuSpec.����) = 1500
        .ColWidth(menuSpec.���) = 1000
        .ColWidth(menuSpec.���㵥λ) = 1000
        .ColWidth(menuSpec.����) = 1200
        .ColWidth(menuSpec.�Ƿ���) = 850
        .ColWidth(menuSpec.��������) = 900
        .ColWidth(menuSpec.�������) = 1200
        .ColWidth(menuSpec.ҩ��ID) = 0
        
    End With
    vsfSpec.Cell(flexcpAlignment, 0, 0, 0, vsfSpec.Cols - 1) = flexAlignCenterCenter
    If vsfSpec.Rows > 1 Then
        vsfSpec.Cell(flexcpAlignment, 1, 0, vsfSpec.Rows - 1, vsfSpec.Cols - 1) = flexAlignLeftCenter
    End If
    vsfSpec.Cell(flexcpFontBold, 0, 0, 0, vsfSpec.Cols - 1) = 35
    
    With vsfVariety
        .SelectionMode = flexSelectionByRow
        .AllowSelection = False '���ܶ�ѡ
        .ExplorerBar = flexExSortShowAndMove '������ƶ�
        .AllowUserResizing = flexResizeBoth  '���Ըı����п��
        .Cols = menuVar.Cols
        .Rows = 1
        .TextMatrix(0, menuVar.ID) = "ID"
        .TextMatrix(0, menuVar.����) = "����"
        .TextMatrix(0, menuVar.����) = "����"
        .TextMatrix(0, menuVar.���㵥λ) = "���㵥λ"
        .TextMatrix(0, menuVar.�������) = "�������"
        
        .ColHidden(menuVar.ID) = True
        .ColWidth(menuVar.����) = 800
        .ColWidth(menuVar.����) = 1500
        .ColWidth(menuVar.���㵥λ) = 1000
        .ColWidth(menuVar.�������) = 1200
    End With
    vsfVariety.Cell(flexcpAlignment, 0, 0, 0, vsfVariety.Cols - 1) = flexAlignCenterCenter
    If vsfVariety.Rows > 1 Then
        vsfVariety.Cell(flexcpAlignment, 1, 0, vsfVariety.Rows - 1, vsfVariety.Cols - 1) = flexAlignLeftCenter
    End If
    vsfVariety.Cell(flexcpFontBold, 0, 0, 0, vsfVariety.Cols - 1) = 35
End Sub

Private Sub vsfSpec_DblClick()
    With vsfSpec
        If Val(.TextMatrix(.Row, menuSpec.ID)) <> 0 Then
            mstrName = .TextMatrix(.Row, menuSpec.ҩ��ID) & "," & .TextMatrix(.Row, menuSpec.ID) & "," & .TextMatrix(.Row, menuSpec.����) & "(" & .TextMatrix(.Row, menuSpec.���) & ")" & "," & .TextMatrix(.Row, menuSpec.���㵥λ)
            Unload Me
        End If
    End With
End Sub

Private Sub vsfSpec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call vsfSpec_DblClick
    End If
End Sub

Private Sub vsfVariety_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow And Val(vsfVariety.TextMatrix(vsfVariety.Row, menuVar.ID)) <> 0 And mintDosageType = 0 Then 'ɢװ�Ų�ѯ���
        Call GetSpecInfo
    End If
End Sub

Private Sub vsfVariety_DblClick()
    With vsfVariety
        If Val(.TextMatrix(.Row, menuVar.ID)) <> 0 Then
            mstrName = .TextMatrix(.Row, menuVar.ID) & ",0," & .TextMatrix(.Row, menuVar.����) & "," & .TextMatrix(.Row, menuVar.���㵥λ)
            Unload Me
        End If
    End With
End Sub

Private Sub vsfVariety_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call vsfVariety_DblClick
    End If
End Sub
