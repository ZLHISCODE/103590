VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPatiCureCardEInvoice 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid vsEInvoice 
      Height          =   1845
      Left            =   615
      TabIndex        =   0
      Top             =   495
      Width           =   1800
      _cx             =   3175
      _cy             =   3254
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
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
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
End
Attribute VB_Name = "frmPatiCureCardEInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long
Private mobjEInvoice As clsEInvoiceObj  '����Ʊ�ݲ���
Private mlngԭ����ID As Long
Private mlng����Ʊ��ID As Long, mbln�Ƿ񻻿� As Boolean
Private mbln�Ƿ����Ʊ�� As Boolean

Public Sub zlInitVar(ByRef objEInvoice As clsEInvoiceObj, ByVal strPrivs As String, ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ر���
    '���:objEinvoice-���ӷ�Ʊ����
    '     strPrivs-��ǰȨ�޴�
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-25 16:59:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnVisible As Boolean
    Set mobjEInvoice = objEInvoice: mstrPrivs = strPrivs: mlngModule = lngModule
End Sub

 
Public Property Get �Ƿ����Ʊ��() As Boolean
    �Ƿ����Ʊ�� = mbln�Ƿ����Ʊ��
End Property
Public Property Get ԭ����ID() As Long
    ԭ����ID = mlngԭ����ID
End Property
Public Property Get ����Ʊ��ID() As Long
    ����Ʊ��ID = mlng����Ʊ��ID
End Property
Public Property Get �Ƿ񻻿�() As Boolean
    �Ƿ񻻿� = mbln�Ƿ񻻿�
End Property
Public Sub zlReLoadData(ByVal strNo As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ص��ӷ�Ʊ��Ϣ
    '����:���˺�
    '����:2020-03-25 17:13:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strSQL As String
    Dim rsEInvoice As ADODB.Recordset
  
    On Error GoTo errHandle
     
    mlng����Ʊ��ID = 0: mbln�Ƿ񻻿� = False
    vsEInvoice.Clear 1: vsEInvoice.Rows = 2
     
    If strNo = "" Or mobjEInvoice Is Nothing Then Exit Sub
    
    mbln�Ƿ����Ʊ�� = mobjEInvoice.zlIsStartEinvoicFromNO(strNo, mlngԭ����ID)
    If mlngԭ����ID = 0 Then Exit Sub

    If Not mobjEInvoice.zlGetEInvoiceInforFromBalanceID(mlngԭ����ID, rsEInvoice, 5, 0) Then Exit Sub
    If rsEInvoice.EOF Then Exit Sub
    
    With vsEInvoice
        If rsEInvoice.RecordCount <> 0 Then rsEInvoice.MoveFirst
        i = 1
        Do While Not rsEInvoice.EOF
            .TextMatrix(i, .ColIndex("ID")) = Nvl(rsEInvoice!id)
            .TextMatrix(i, .ColIndex("��¼״̬")) = Nvl(rsEInvoice!��¼״̬)
            .TextMatrix(i, .ColIndex("����ID")) = Nvl(rsEInvoice!����ID)
            .TextMatrix(i, .ColIndex("��Ʊ����")) = Nvl(rsEInvoice!����)
            .TextMatrix(i, .ColIndex("��Ʊ����")) = Nvl(rsEInvoice!����)
            .TextMatrix(i, .ColIndex("Ʊ�ݽ��")) = Format(Nvl(rsEInvoice!Ʊ�ݽ��), "###0.00;-###0.00;;")
            .TextMatrix(i, .ColIndex("����ʱ��")) = Format(rsEInvoice!����ʱ��, "yyyy-mm-dd HH:MM:SS")
            .TextMatrix(i, .ColIndex("����ֽ�ʷ�Ʊ")) = IIf(Val(Nvl(rsEInvoice!�Ƿ񻻿�)) = 1, "�ѻ���", "δ����")
            .TextMatrix(i, .ColIndex("ֽ�ʷ�Ʊ��")) = Nvl(rsEInvoice!ֽ�ʷ�Ʊ��)
            .TextMatrix(i, .ColIndex("��ע")) = Nvl(rsEInvoice!��ע)
            .TextMatrix(i, .ColIndex("����Ա����")) = Nvl(rsEInvoice!����Ա����)
            If Val(Nvl(rsEInvoice!��¼״̬)) = 1 Then
                mlng����Ʊ��ID = Nvl(rsEInvoice!id): mbln�Ƿ񻻿� = Val(Nvl(rsEInvoice!�Ƿ񻻿�)) = 1
                 .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = Me.ForeColor
            Else
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = IIf(Val(Nvl(rsEInvoice!��¼״̬)) = 2, vbRed, vbBlue)
            End If
            i = i + 1: .Rows = .Rows + 1
            rsEInvoice.MoveNext
        Loop
        If .Rows > 2 Then .Rows = .Rows - 1
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Call InitEinvoiceGrid
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With vsEInvoice
        .Top = Me.ScaleTop
        .Left = Me.ScaleLeft
        .Height = Me.ScaleHeight
        .Width = Me.ScaleWidth
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    zl_vsGrid_Para_Save 1107, vsEInvoice, Me.Name, "����Ʊ����Ϣ�б�", False
    Set mobjEInvoice = Nothing
End Sub

Private Sub vsEInvoice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsEInvoice, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsEInvoice_GotFocus()
    zl_VsGridGotFocus vsEInvoice, &HFFC0C0
End Sub

Private Sub vsEInvoice_LostFocus()
    zl_VsGridLostFocus vsEInvoice, , vsEInvoice.Cell(flexcpForeColor, vsEInvoice.Row, vsEInvoice.Col)
End Sub
Private Sub InitEinvoiceGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�����ӷ�Ʊ����ؼ�
    '����:���˺�
    '����:2020-03-25 17:16:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With vsEInvoice
        .Redraw = flexRDNone
        .HighLight = flexHighlightWithFocus
        .Clear 1: .Rows = 2
        .Cols = 11
        .TextMatrix(0, i) = "ID": i = i + 1
        .TextMatrix(0, i) = "��¼״̬": i = i + 1
        .TextMatrix(0, i) = "����ID": i = i + 1
        .TextMatrix(0, i) = "��Ʊ����": i = i + 1
        .TextMatrix(0, i) = "��Ʊ����": i = i + 1
        .TextMatrix(0, i) = "Ʊ�ݽ��": i = i + 1
        .TextMatrix(0, i) = "����ʱ��": i = i + 1
        .TextMatrix(0, i) = "����ֽ�ʷ�Ʊ": i = i + 1
        .TextMatrix(0, i) = "ֽ�ʷ�Ʊ��": i = i + 1
        .TextMatrix(0, i) = "��ע": i = i + 1
        .TextMatrix(0, i) = "����Ա����": i = i + 1
        
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter: .ColAlignment(i) = flexAlignLeftCenter
            .ColKey(i) = .TextMatrix(0, i)
            .ColWidth(i) = 1000
            If .ColKey(i) Like "*ID" Then .ColHidden(i) = True: .ColWidth(i) = 0: .ColData(i) = "-1|1"
            Select Case .ColKey(i)
            Case "��¼״̬"
                .ColHidden(i) = True: .ColWidth(i) = 0: .ColData(i) = "-1|1"
            Case "��ע"
                .ColWidth(i) = 2000
            Case "����Ա����"
                 .ColWidth(i) = 1000
            Case "Ʊ�ݽ��"
                .ColAlignment(i) = flexAlignRightCenter
            End Select
        Next
         .Row = 1: .Col = 0: .ColSel = .Cols - 1
        .RowHeightMin = 350
        zl_vsGrid_Para_Restore 1107, vsEInvoice, Me.Name, "����Ʊ����Ϣ�б�", False
        If .Rows < 2 Then .Rows = 2
        .Redraw = flexRDBuffered
    End With
End Sub
