VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmRAParams 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7110
   Icon            =   "frmRAParams.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtAuto 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4920
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "20"
      Top             =   210
      Width           =   615
   End
   Begin VB.CheckBox chkAutoRefreshPatient 
      Caption         =   "���没���б�������ʱ�������Զ�ˢ�²����б� ���"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   360
      Left            =   4560
      TabIndex        =   6
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   360
      Left            =   5760
      TabIndex        =   7
      Top             =   6600
      Width           =   1095
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfFromDept 
      Height          =   5655
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   6855
      _cx             =   12091
      _cy             =   9975
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
   Begin VB.CheckBox chkAll 
      Caption         =   "ȫѡ(&A)"
      Height          =   180
      Left            =   6000
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�루10-999��"
      Height          =   180
      Left            =   5640
      TabIndex        =   2
      Top             =   240
      Width           =   1080
   End
   Begin VB.Label lblFromDept 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ñ���������Դ���ң�"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1980
   End
End
Attribute VB_Name = "frmRAParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MSTR_VSF As String = "����ID,,0,0|ѡ��,,3,500|����,,3,1000|����,,3,2000|��ѡ����,,3,3000"

Private mbytMode As Byte            '����չ�ֵ�ģʽ��0-���ﴦ����飻1-סԺҩ�����
Private mstrPCName As String
Private mblnEnter As Boolean

Public Sub ShowMe(ByVal frmOwner As Form, ByVal bytMode As Byte)
'���ܣ���ʾ����ӿڷ���
'������
'  frmOwner�������������
'  bytMode��ָ������չ�ֵ�ģʽ��0-���ﴦ����飻1-סԺҩ�����

    mbytMode = bytMode
    mstrPCName = UCase(OS.ComputerName)
    
    Show vbModal, frmOwner
    
End Sub

Private Sub chkAll_Click()
    Dim l As Long
    
    With vsfFromDept
        If .Rows <= 1 Then Exit Sub
        
        For l = 1 To .Rows - 1
            'ֻ������CheckBoxֵ����ͬ����
            If Val(.TextMatrix(l, .ColIndex("ѡ��"))) <> IIf(chkAll.Value = 1, -1, 0) Then
                .TextMatrix(l, .ColIndex("ѡ��")) = IIf(chkAll.Value = 1, -1, 0)
            End If
        Next
    End With
    
End Sub

Private Sub chkAutoRefreshPatient_Click()
    txtAuto.Enabled = chkAutoRefreshPatient.Value = 1
    If txtAuto.Enabled And txtAuto.Visible Then txtAuto.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String, strTmp As String
    Dim l As Long
    Dim intSecond As Integer
    
    '������
    If vsfFromDept.Rows <= 1 Then
        MsgBox "����Դ���ң��ܾ����棡", vbInformation, gstrSysName
        Exit Sub
    End If
    If mstrPCName = "" Then
        MsgBox "������δ��ȡ�����ܾ����棡", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(txtAuto.Text) <= 0 And chkAutoRefreshPatient.Value = 1 Then
        MsgBox "���������д���ԣ��ܾ����棡", vbInformation, gstrSysName
        If txtAuto.Visible And txtAuto.Enabled Then txtAuto.SetFocus
        Exit Sub
    End If
    
    '��֯����
    With vsfFromDept
        For l = 1 To .Rows - 1
            If Val(.TextMatrix(l, .ColIndex("ѡ��"))) = -1 Then
                If Val(.TextMatrix(l, .ColIndex("����ID"))) > 0 Then
                    strTmp = strTmp & zlStr.FormatString("[1],", Val(.TextMatrix(l, .ColIndex("����ID"))))
                End If
            End If
        Next
        If strTmp <> "" Then strTmp = Left(strTmp, Len(strTmp) - 1)
    End With
    If Len(strTmp) > 3999 Then
        MsgBox "����Դ���ҡ��������࣡", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�������
    If chkAutoRefreshPatient.Value = 1 Then
        intSecond = Val(txtAuto.Text)
    End If
    If mbytMode = Val("1-סԺҩ�����") Then
        Call zlDatabase.SetPara("�Զ�ˢ�²����б�", intSecond, glngSys, Val("1352-סԺҽ�����"))
    Else
        Call zlDatabase.SetPara("�Զ�ˢ�²����б�", intSecond, glngSys, Val("1351-���ﴦ�����"))
    End If
    
    On Error GoTo errHandle
    strSQL = zlStr.FormatString("Zl_����������_Save(1, [1], [2], Null, [3])", _
                    "'" & mstrPCName & "'", _
                    mbytMode, _
                    IIf(strTmp = "", "Null", "'" & strTmp & "'"))
    Call zlDatabase.ExecuteProcedure(strSQL, "���洦��������-��Դ����")

    Unload Me
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Load()
    Dim intSecond As Integer

    mblnEnter = False
    lblFromDept.Caption = zlStr.FormatString("���ñ�����[1]������Դ���ң�", mstrPCName)
    
    If mbytMode = Val("1-סԺҩ�����") Then
        intSecond = Val(zlDatabase.GetPara("�Զ�ˢ�²����б�", glngSys, Val("1352-סԺҩ�����")))
    Else
        intSecond = Val(zlDatabase.GetPara("�Զ�ˢ�²����б�", glngSys, Val("1351-���ﴦ�����")))
    End If
    txtAuto.Text = CStr(intSecond)
    chkAutoRefreshPatient.Value = IIf(intSecond > 0, 1, 0)
    
    Call InitVSF
    Call mdlDefine.SetVSFHead(vsfFromDept, MSTR_VSF)
    Call SetClinicDept
    
    With vsfFromDept
        .ColDataType(.ColIndex("ѡ��")) = flexDTBoolean
    End With
    
    Call SetPCName
    Call chkAutoRefreshPatient_Click
    
    zl9ComLib.RestoreWinState Me, App.ProductName
    mblnEnter = True
End Sub

Private Sub SetPCName()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim l As Long
    
    If vsfFromDept.Rows <= 1 Then Exit Sub
    
    On Error GoTo errHandle
    
    strSQL = "Select Upper(f_List2str(Cast(Collect(������) As t_Strlist), '��')) ������ " & vbNewLine & _
             "From ���������� " & vbNewLine & _
             "Where Nvl(�������, 0) = [1] And ',' || ��Դ���� || ',' Like [2] "
    
    With vsfFromDept
        For l = 1 To .Rows - 1
            'ȡ����id��ͬ�Ļ�����
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������", mbytMode, zlStr.FormatString("%,[1],%", .TextMatrix(l, .ColIndex("����ID"))))
            If rsTemp.EOF = False Then
                .TextMatrix(l, .ColIndex("��ѡ����")) = NVL(rsTemp!������)
                If .TextMatrix(l, .ColIndex("��ѡ����")) <> "" Then
                    If "��" & .TextMatrix(l, .ColIndex("��ѡ����")) & "��" Like "*��" & mstrPCName & "��*" Then
                        .TextMatrix(l, .ColIndex("ѡ��")) = -1
                    End If
                End If
            End If
        Next
    End With

    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub SetClinicDept()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    If mbytMode = 1 Then
        'סԺ
        strSQL = "Select a.����ID, b.����, b.���� " & vbNewLine & _
                 "From ��������˵�� A, ���ű� B " & vbNewLine & _
                 "Where a.����id = b.Id And a.�������� = '�ٴ�' And a.������� In ([1], 3) And (b.����ʱ�� Is Null Or To_Char(b.����ʱ��, 'yyyy') = '3000') "
    Else
        '����
        strSQL = "Select Count(1) Rec From ����������� Where Rownum < 2 And ��� <> 1 "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�����������-�ǿ��ҵ�����")
        If rsTemp!Rec > 0 Then
            strSQL = "Select a.����id, b.����, b.���� " & vbNewLine & _
                     "From ��������˵�� A, ���ű� B " & vbNewLine & _
                     "Where a.����id = b.Id And a.�������� = '�ٴ�' And a.������� In ([1], 3) " & vbNewLine & _
                     "    And (b.����ʱ�� Is Null Or To_Char(b.����ʱ��, 'yyyy') = '3000') "
        Else
            strSQL = "Select a.����id, b.����, b.���� " & vbNewLine & _
                     "From ��������˵�� A, ���ű� B, ����������� C " & vbNewLine & _
                     "Where a.����id = b.Id And a.����id = c.����id And a.�������� = '�ٴ�' And a.������� In ([1], 3) " & vbNewLine & _
                     "    And (b.����ʱ�� Is Null Or To_Char(b.����ʱ��, 'yyyy') = '3000') And c.��� = 1 And (c.����id Is Not Null Or c.����id > 0)"
        End If
        rsTemp.Close
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ" & IIf(mbytMode = 1, "סԺ", "����") & "ҵ����ٴ�����", IIf(mbytMode = 1, 2, 1))
    mdlDefine.FillVSFData vsfFromDept, rsTemp
    rsTemp.Close
    
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub InitVSF()
'���ܣ���ʼ�������VSFlexGrid�ؼ��ķ��

    With vsfFromDept
        .Appearance = flexFlat
        .AllowSelection = False
        .AllowUserResizing = flexResizeColumns
        .FixedCols = 0
        .FixedRows = 1
        .Editable = flexEDKbdMouse
        .SelectionMode = flexSelectionByRow
        .SheetBorder = .BackColor
        .BackColorBkg = .BackColor
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zl9ComLib.SaveWinState Me, App.ProductName
End Sub

Private Sub txtAuto_GotFocus()
    Call zlControl.TxtSelAll(txtAuto)
End Sub

Private Sub txtAuto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtAuto_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txtAuto_Validate(Cancel As Boolean)
    If Val(txtAuto.Text) < 0 Or Val(txtAuto.Text) > 0 And Val(txtAuto.Text) < 10 Then
        MsgBox "�Զ�ˢ�µ��������������д����ȷ��", vbInformation, gstrSysName
        Cancel = True
    End If
End Sub

Private Sub vsfFromDept_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col <> vsfFromDept.ColIndex("ѡ��")
End Sub

Private Sub vsfFromDept_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If mblnEnter = False Then Exit Sub

    With vsfFromDept
        If .ColIndex("��ѡ����") <= -1 Then Exit Sub
        
        If Col = .ColIndex("ѡ��") Then
            If Val(.TextMatrix(Row, Col)) = -1 Then
                '�ӻ�����
                If Trim(.TextMatrix(Row, .ColIndex("��ѡ����"))) = "" Then
                    .TextMatrix(Row, .ColIndex("��ѡ����")) = mstrPCName
                Else
                    .TextMatrix(Row, .ColIndex("��ѡ����")) = .TextMatrix(Row, .ColIndex("��ѡ����")) & "��" & mstrPCName
                End If
            Else
                '��������
                If .TextMatrix(Row, .ColIndex("��ѡ����")) Like "*��" & mstrPCName & "��*" _
                    Or .TextMatrix(Row, .ColIndex("��ѡ����")) Like mstrPCName & "��*" Then
                    .TextMatrix(Row, .ColIndex("��ѡ����")) = Replace(.TextMatrix(Row, .ColIndex("��ѡ����")), mstrPCName & "��", "")
                ElseIf .TextMatrix(Row, .ColIndex("��ѡ����")) Like "*��" & mstrPCName Then
                    .TextMatrix(Row, .ColIndex("��ѡ����")) = Replace(.TextMatrix(Row, .ColIndex("��ѡ����")), "��" & mstrPCName, "")
                ElseIf .TextMatrix(Row, .ColIndex("��ѡ����")) = mstrPCName Then
                    .TextMatrix(Row, .ColIndex("��ѡ����")) = ""
                End If
            End If
        End If
    End With
End Sub
