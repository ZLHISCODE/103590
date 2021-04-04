VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmImportFileCondition 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�������"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8970
   Icon            =   "frmImportFileCondition.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   8970
   StartUpPosition =   1  '����������
   Begin VSFlex8Ctl.VSFlexGrid vsfError 
      Height          =   4980
      Left            =   105
      TabIndex        =   5
      Top             =   495
      Width           =   8760
      _cx             =   15452
      _cy             =   8784
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
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   30
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
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
   Begin VB.OptionButton optPartImport 
      Caption         =   "��������"
      Height          =   240
      Left            =   1380
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.OptionButton optNoImport 
      Caption         =   "�����ֹ"
      Height          =   255
      Left            =   2805
      TabIndex        =   2
      Top             =   105
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����"
      Height          =   300
      Left            =   6870
      TabIndex        =   1
      Top             =   60
      Width           =   885
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�"
      Height          =   300
      Left            =   7950
      TabIndex        =   0
      Top             =   60
      Width           =   885
   End
   Begin VB.Label lblImportMethod 
      AutoSize        =   -1  'True
      Caption         =   "���뷽ʽ"
      Height          =   180
      Left            =   135
      TabIndex        =   4
      Top             =   135
      Width           =   720
   End
End
Attribute VB_Name = "frmImportFileCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MCONFIXECOLOR As Long = &H8000000F  '�����޸��б���ɫ
Private strPara             As String             '����ֵ
Private mlngModal           As Long               '��ǰģ���
Private mstrCheck           As String             '������

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim strTemp As String
    Dim intRow  As Integer
    
    With vsfError
        If optNoImport.Value = True Then
            strTemp = "1/"
        Else
            strTemp = "0/"
        End If
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 1) = "��ֹ" Then
                strTemp = strTemp & "1|"
            Else
                strTemp = strTemp & "0|"
            End If
        Next
    End With
    If strTemp <> "" Then
        strTemp = Mid(strTemp, 1, LenB(StrConv(strTemp, vbFromUnicode)) - 1)
    Else
        strTemp = "0/0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0"
    End If
    Call zlDatabase.SetPara("�����ļ���鷽ʽ", strTemp, glngSys, mlngModal)
    Unload Me
End Sub

Private Sub Form_Load()
    Call InitVsf
    Call LoadData
End Sub

Public Sub ShowMe(ByVal frmPar As Form, ByVal lngModal As Long)
    mlngModal = lngModal
    Me.Show vbModal, frmPar
End Sub

Private Sub optNoImport_Click()
    Dim intRow As Integer
    
    With vsfError
        If optNoImport.Value = True Then
            For intRow = 1 To .Rows - 1
                .TextMatrix(intRow, 1) = "��ֹ"
            Next
            .Cell(flexcpBackColor, 1, 0, .Rows - 1, 3) = MCONFIXECOLOR '�����޸�����ɫ
            .Cell(flexcpFontBold, 1, 1, .Rows - 1, 1) = True
            .Editable = flexEDNone
            .Row = 0
        End If
    End With
End Sub

Private Sub optPartImport_Click()
    Dim intRow As Integer
    
    If optPartImport.Value = True Then
        For intRow = 1 To vsfError.Rows - 1
            vsfError.TextMatrix(intRow, 1) = "��ʾ"
        Next
        vsfError.Row = 0
        vsfError.Cell(flexcpBackColor, 1, 1, vsfError.Rows - 1, 1) = &H80000005    '���޸�����ɫ
    End If
End Sub

Private Sub vsfError_CellChanged(ByVal Row As Long, ByVal Col As Long)
    With vsfError
        If Col = 1 Then
            If .TextMatrix(Row, Col) = "��ֹ" Then
                .Cell(flexcpFontBold, Row, 1, Row, 1) = True
            Else
                .Cell(flexcpFontBold, Row, 1, Row, 1) = False
            End If
        End If
    End With
End Sub

Private Sub vsfError_EnterCell()
    With vsfError
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = MCONFIXECOLOR Then
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub LoadData()
    '��������
    Dim strImportMethod As String
    Dim strPara         As String
    Dim intRow          As Integer
    Dim intCol          As Integer
    Dim arryPara        As Variant
    Dim arryTempPara    As Variant
    Dim strTemp         As String
    
    '�����ʽ(0-������ʾ1-�����ֹ/0-��ʾ1-��ֹ|0-��ʾ1-��ֹ|....)
    strPara = zlDatabase.GetPara("�����ļ���鷽ʽ", glngSys, mlngModal, "0/0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0")
    
    arryPara = Split(strPara, "|")
    With vsfError
        For intRow = 0 To UBound(arryPara)
            strTemp = arryPara(intRow)
            If intRow = 0 Then
                strImportMethod = Split(strTemp, "/")(0)
                If strImportMethod = "0" Then
                    optNoImport.Value = False
                    optPartImport.Value = True
                Else
                    optNoImport.Value = True
                    optPartImport.Value = False
                End If
                strTemp = Split(strTemp, "/")(1)
                If strTemp = "0" Then
                    .TextMatrix(intRow + 1, 1) = "��ʾ"
                Else
                    .TextMatrix(intRow + 1, 1) = "��ֹ"
                    .Cell(flexcpFontBold, intRow + 1, 1) = True
                End If
            End If
            If strTemp = "0" Then
                .TextMatrix(intRow + 1, 1) = "��ʾ"
            Else
                .TextMatrix(intRow + 1, 1) = "��ֹ"
                .Cell(flexcpFontBold, intRow + 1, 1) = True
            End If
        Next
    End With
End Sub

Private Sub InitVsf()
    '��ʼ��vsf�ؼ�
    With vsfError
        .Cols = 4
        .Rows = 28
        .FixedRows = 1
        .FixedCols = 1
        .RowHeight(-1) = 450
        .ColWidth(2) = 1500
        .Editable = flexEDNone
        .AllowSelection = False '��ѡ��һ��
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ExplorerBar = flexExNone
        .ExtendLastCol = True '���һ�������
        .ColComboList(1) = "��ֹ|��ʾ"
        .WordWrap = True
        .Cell(flexcpBackColor, 1, 0, .Rows - 1, 3) = MCONFIXECOLOR '�����޸�����ɫ
        .Cell(flexcpAlignment, 0, 0, 0, 3) = flexAlignCenterCenter '��ͷ���мӴ�
        .Cell(flexcpFontBold, 0, 0, 0, 3) = True
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, 3) = flexAlignLeftCenter
        .AutoResize = True
        .WordWrap = True '���ֻ���
        .AutoSizeMode = flexAutoSizeRowHeight '�Զ�����
        .MergeCells = flexMergeFree '��Ԫ��ϲ�
        .MergeCol(0) = True
        
        .TextMatrix(0, 0) = "�������"
        .TextMatrix(0, 1) = "��鷽ʽ"
        .Cell(flexcpFontBold, 0, 1, 0, 1) = True
        .TextMatrix(0, 2) = "������"
        .TextMatrix(0, 3) = "��ע"
        
        .TextMatrix(1, 0) = "����"
        .TextMatrix(1, 2) = "���"
        .TextMatrix(1, 3) = "���ֻ��������ҩ���г�ҩ���в�ҩ������Ϊ��"
        .TextMatrix(2, 0) = "����"
        .TextMatrix(2, 2) = "�ϼ�����"
        .TextMatrix(2, 3) = "�ϼ�������ձ������е����ݣ���ʽ����\�ָ���������"
        .TextMatrix(3, 0) = "����"
        .TextMatrix(3, 2) = "����"
        .TextMatrix(3, 3) = "���ܺ��зǷ��ַ�������Ϊ�գ����Ȳ��ܳ������ݿ��ֶγ���"
        .TextMatrix(4, 0) = "����"
        .TextMatrix(4, 2) = "����"
        .TextMatrix(4, 3) = "���ܺ��зǷ��ַ�������Ϊ�գ����Ȳ��ܳ������ݿ��ֶγ���"
        .TextMatrix(5, 0) = "����"
        .TextMatrix(5, 2) = "��������Ψһ���"
        .TextMatrix(5, 3) = "������������������л����ݿ�������������ͬ"
        .TextMatrix(6, 0) = "����"
        .TextMatrix(6, 2) = "����.���.�ϼ�����Ψһ���"
        .TextMatrix(6, 3) = "���ơ�����ϼ����಻����������л����ݿ�������������ͬ"
        .TextMatrix(7, 0) = "��ϸ"
        .TextMatrix(7, 2) = "���"
        .TextMatrix(7, 3) = "ֻ��������ҩ���г�ҩ���в�ҩ������Ϊ��"
        .TextMatrix(8, 0) = "��ϸ"
        .TextMatrix(8, 2) = "����"
        .TextMatrix(8, 3) = "���ձ������е����ݣ�����Ϊ�գ���ʽ����\�ָ���������"
        .TextMatrix(9, 0) = "��ϸ"
        .TextMatrix(9, 2) = "Ʒ�ֱ���"
        .TextMatrix(9, 3) = "���ܺ��зǷ��ַ�������Ϊ�գ����Ȳ��ܳ������ݿ��ֶγ���"
        .TextMatrix(10, 0) = "��ϸ"
        .TextMatrix(10, 2) = "Ʒ������"
        .TextMatrix(10, 3) = "���ܺ��зǷ��ַ�������Ϊ�գ����Ȳ��ܳ������ݿ��ֶγ���"
        .TextMatrix(11, 0) = "��ϸ"
        .TextMatrix(11, 2) = "������"
        .TextMatrix(11, 3) = "���ܺ��зǷ��ַ�������Ϊ�գ����Ȳ��ܳ������ݿ��ֶγ���"
        .TextMatrix(12, 0) = "��ϸ"
        .TextMatrix(12, 2) = "ҩƷ���"
        .TextMatrix(12, 3) = "���ܺ��зǷ��ַ�������Ϊ�գ����Ȳ��ܳ������ݿ��ֶγ���"
        .TextMatrix(13, 0) = "��ϸ"
        .TextMatrix(13, 2) = "������"
        .TextMatrix(13, 3) = "���ܺ��зǷ��ַ������Ȳ��ܳ������ݿ��ֶγ���"
        .TextMatrix(14, 0) = "��ϸ"
        .TextMatrix(14, 2) = "����"
        .TextMatrix(14, 3) = "���ձ������е����ݣ�����Ϊ�գ����ܺ��зǷ��ַ������Ȳ��ܳ������ݿ��ֶγ���"
        .TextMatrix(15, 0) = "��ϸ"
        .TextMatrix(15, 2) = "������λ���"
        .TextMatrix(15, 3) = "���ܺ��зǷ��ַ�������Ϊ�գ����Ȳ��ܳ������ݿ��ֶγ���"
        .TextMatrix(16, 0) = "��ϸ"
        .TextMatrix(16, 2) = "������λ������"
        .TextMatrix(16, 3) = "����Ϊ�գ���λ����ϵ���������Ҷ�>0����λ��ͬ����ϵ��������ͬ"
        .TextMatrix(17, 0) = "��ϸ"
        .TextMatrix(17, 2) = "��ۼ��"
        .TextMatrix(17, 3) = "Ϊ��Ĭ��Ϊ���ۣ����̡���ʾʱ�ۣ���������ֻ���ǡ��̡����"
        .TextMatrix(18, 0) = "��ϸ"
        .TextMatrix(18, 2) = "�۸���"
        .TextMatrix(18, 3) = "�۸��ֶ�ֻ���������ͣ����Ȳ��ܳ���������þ���"
        .TextMatrix(19, 0) = "��ϸ"
        .TextMatrix(19, 2) = "Ч��"
        .TextMatrix(19, 3) = "������������ֻ���ǲ�С��0������"
        .TextMatrix(20, 0) = "��ϸ"
        .TextMatrix(20, 2) = "������Ŀ"
        .TextMatrix(20, 3) = "���ܺ��зǷ��ַ�������Ϊ�գ�ֻ�������ݿ�����������Ŀ"
        .TextMatrix(21, 0) = "��ϸ"
        .TextMatrix(21, 2) = "����/סԺ����"
        .TextMatrix(21, 3) = "ֻ�����������õķ��㷽ʽ������Ϊ��"
        .TextMatrix(22, 0) = "��ϸ"
        .TextMatrix(22, 2) = "�������"
        .TextMatrix(22, 3) = "ֻ�����������ݿ����õķ������"
        .TextMatrix(23, 0) = "��ϸ"
        .TextMatrix(23, 2) = "��������"
        .TextMatrix(23, 3) = "ҩ�����ʱҩ�����ܷ�������������ֻ���ǡ��̡���գ�Ϊ�ձ�ʾ������"
        .TextMatrix(24, 0) = "��ϸ"
        .TextMatrix(24, 2) = "��Ӧ��"
        .TextMatrix(24, 3) = "ֻ�������ݿ������еĹ�Ӧ��"
        .TextMatrix(25, 0) = "��ϸ"
        .TextMatrix(25, 2) = "���ڸ�ʽ"
        .TextMatrix(25, 3) = "���ڸ�ʽ��2015-10-10����2015/10/10����2015.10.10"
        .TextMatrix(26, 0) = "��ϸ"
        .TextMatrix(26, 2) = "Ʒ��Ψһ�Լ��"
        .TextMatrix(26, 3) = "�жϵ�����Ŀ������������ݻ����ݿ�������Ʒ���Ƿ��ͻ"
        .TextMatrix(27, 0) = "��ϸ"
        .TextMatrix(27, 2) = "���Ψһ�Լ��"
        .TextMatrix(27, 3) = "�жϵ�����Ŀ������������ݻ����ݿ������й���Ƿ��ͻ"
    End With
    Call GetCheck
End Sub

Private Sub GetCheck()
    Dim intRow As Integer
    
    mstrCheck = ""
    With vsfError
        For intRow = 1 To .Rows - 1
            mstrCheck = mstrCheck & "|" & .TextMatrix(intRow, 2)
        Next
    End With
End Sub
