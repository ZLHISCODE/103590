VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm����ѡ���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ѡ����"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4920
   Icon            =   "frm����ѡ����.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VSFlex8Ctl.VSFlexGrid vfgColumn 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      _cx             =   6165
      _cy             =   6376
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
      BackColorSel    =   16635590
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   280
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
      TabBehavior     =   1
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
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   3630
      TabIndex        =   10
      Top             =   90
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3630
      TabIndex        =   9
      Top             =   510
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   3630
      TabIndex        =   8
      Top             =   1290
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtWidth 
      Height          =   300
      Left            =   1170
      TabIndex        =   1
      Top             =   3810
      Width           =   1155
   End
   Begin VB.ComboBox cmbAlign 
      Height          =   300
      Left            =   1170
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4230
      Width           =   1185
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "ȫѡ(&S)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   3630
      TabIndex        =   7
      Top             =   2040
      Width           =   1100
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "ȫ��(&L)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   3630
      TabIndex        =   6
      Top             =   2460
      Width           =   1100
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "����(&U)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   3630
      TabIndex        =   5
      Top             =   2880
      Width           =   1100
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "����(&D)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   3630
      TabIndex        =   4
      Top             =   3360
      Width           =   1100
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "�ָ�Ĭ������(&R)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   3030
      TabIndex        =   3
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label lblAlign 
      Caption         =   "���뷽ʽ(&A)"
      Height          =   180
      Left            =   60
      TabIndex        =   12
      Top             =   4290
      Width           =   990
   End
   Begin VB.Label lblWidth 
      Caption         =   "�п�(&W)"
      Height          =   180
      Left            =   420
      TabIndex        =   11
      Top             =   3870
      Width           =   630
   End
End
Attribute VB_Name = "frm����ѡ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mVsGrid As VSFlexGrid
Private mblnOK As Boolean

Public Function ShowColSet(ByVal frmMain As Form, ByVal strTittle As String, vsGrid As VSFlexGrid) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ýӿ�
    '����:
    '����:�����óɹ�,����true,���򷵻�False
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Err = 0: On Error Resume Next
    Set mVsGrid = vsGrid
    If strTittle <> "" Then Me.Caption = strTittle
    
    cmbAlign.AddItem "���϶���"
    cmbAlign.AddItem "���ж���"
    cmbAlign.AddItem "���¶���"
    cmbAlign.AddItem "���϶���"
    cmbAlign.AddItem "���ж���"
    cmbAlign.AddItem "���¶���"
    cmbAlign.AddItem "���϶���"
    cmbAlign.AddItem "���ж���"
    cmbAlign.AddItem "���¶���"
    
    Call LoadFulltoColSel
'    Call cmdRestore_Click
    With Me
        .Show vbModal, frmMain
    End With
    ShowColSet = mblnOK
End Function

Private Function LoadFulltoColSel() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����������
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:lesfeng
    '����:2009-08-25 16:46:43
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long, lngRow As Long, arrSplit As Variant
    Dim sngFrmHeight As Single, sngSelSumHeight As Single

    Call initVfgColumnTitle
    With mVsGrid
        lngRow = 1
        For i = 0 To .Cols - 1
            'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
            arrSplit = Split(.ColData(i) & "||", "||")
            
            If Trim(.ColKey(i)) <> "" And (Val(arrSplit(0)) = 1 Or Val(arrSplit(0)) = 0) Then
                vfgColumn.TextMatrix(lngRow, vfgColumn.ColIndex("����")) = .ColKey(i)
                vfgColumn.TextMatrix(lngRow, vfgColumn.ColIndex("ѡ��")) = IIf(.ColWidth(i) = 0 Or .ColHidden(i), False, True)
                vfgColumn.TextMatrix(lngRow, vfgColumn.ColIndex("����")) = .ColAlignment(i)
                vfgColumn.TextMatrix(lngRow, vfgColumn.ColIndex("�п�")) = .ColWidth(i)
                vfgColumn.TextMatrix(lngRow, vfgColumn.ColIndex("�̶�")) = arrSplit(0)
                If .ColWidth(i) = 0 Or .ColHidden(i) Then
                    vfgColumn.TextMatrix(lngRow, vfgColumn.ColIndex("ԭֵ")) = 0
                    vfgColumn.TextMatrix(lngRow, vfgColumn.ColIndex("�ı�")) = 0
                Else
                    vfgColumn.TextMatrix(lngRow, vfgColumn.ColIndex("ԭֵ")) = 1
                    vfgColumn.TextMatrix(lngRow, vfgColumn.ColIndex("�ı�")) = 1
                End If
                If Val(arrSplit(0)) = 1 Then
                    vfgColumn.TextMatrix(lngRow, vfgColumn.ColIndex("��ע")) = "��������"
                End If

                vfgColumn.RowData(lngRow) = Val(arrSplit(0))
                If Val(arrSplit(0)) = 1 Then
                    vfgColumn.Cell(flexcpForeColor, lngRow, 0, lngRow, vfgColumn.Cols - 1) = vbBlue
                End If
                vfgColumn.Rows = vfgColumn.Rows + 1
                lngRow = lngRow + 1
            End If
        Next
    End With
    
    If vfgColumn.Rows > 2 Then vfgColumn.Rows = vfgColumn.Rows - 1
    With vfgColumn
        .Visible = True
        .Editable = flexEDKbdMouse
        .ZOrder 0
        .ColDataType(.ColIndex("ѡ��")) = flexDTBoolean
'        '�������϶�
'        .ExplorerBar = flexExSortShowAndMove
        '��ѡ��
        .SelectionMode = flexSelectionByRow

        If .Rows > 1 Then
            .Row = 1
            .Select .Row, .ColIndex("ѡ��")
            Call vfgColumn_Click
            Call setenabled
        End If
        .SetFocus
    End With
    Call SetcmbAlign
End Function

Private Function SetVsGridCol(ByVal strColKey As String, ByVal blnShow As Boolean, ByVal blnBatch As Boolean, ByVal lngColWidth As Long, ByVal lngAlign As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:������ʾ��
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:lesfeng
    '����:2009-08-25 16:46:43
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long, lngRow As Long
        
    With mVsGrid
        .ColHidden(.ColIndex(strColKey)) = Not blnShow
        If .ColWidth(.ColIndex(strColKey)) = 0 Then .ColWidth(.ColIndex(strColKey)) = 1000
        If lngColWidth >= 0 Then .ColWidth(.ColIndex(strColKey)) = lngColWidth
        .ColAlignment(.ColIndex(strColKey)) = lngAlign
        '����29530 by lesfeng 2010-05-06
        If .Rows > 1 Then .Cell(flexcpAlignment, .FixedRows, .ColIndex(strColKey), .Rows - 1, .ColIndex(strColKey)) = lngAlign
    End With
End Function

Private Function SaveData() As Boolean
    Dim strOldValue As String
    Dim strNewValue As String
    Dim lngColWith As Long
    Dim lngAlign As Long
    Dim lngRow As Long
    Dim blnShow As Boolean
    
    With vfgColumn
        For lngRow = 1 To .Rows - 1
            strOldValue = .TextMatrix(lngRow, .ColIndex("ԭֵ"))
            strNewValue = .TextMatrix(lngRow, .ColIndex("�ı�"))
            If strOldValue <> strNewValue Then
                 blnShow = GetVsGridBoolColVal(vfgColumn, lngRow, .ColIndex("ѡ��"))
                 lngColWith = Val(.TextMatrix(lngRow, .ColIndex("�п�")))
                 lngAlign = .TextMatrix(lngRow, .ColIndex("����"))
                 Call SetVsGridCol(.TextMatrix(lngRow, .ColIndex("����")), blnShow, IIf(.Tag = "Head", False, True), lngColWith, lngAlign)
            End If
        Next
    End With
End Function

Private Sub cmbAlign_Click()
    Dim strAlign As String
    
    strAlign = cmbAlign.Text
    With vfgColumn
        If .Row > 0 Then
            Select Case strAlign
            Case "���϶���"
                .TextMatrix(.Row, .ColIndex("����")) = 0
            Case "���ж���"
                .TextMatrix(.Row, .ColIndex("����")) = 1
            Case "���¶���"
                .TextMatrix(.Row, .ColIndex("����")) = 2
            Case "���϶���"
                .TextMatrix(.Row, .ColIndex("����")) = 3
            Case "���ж���"
                .TextMatrix(.Row, .ColIndex("����")) = 4
            Case "���¶���"
                .TextMatrix(.Row, .ColIndex("����")) = 5
            Case "���϶���"
                .TextMatrix(.Row, .ColIndex("����")) = 6
            Case "���ж���"
                .TextMatrix(.Row, .ColIndex("����")) = 7
            Case "���¶���"
                .TextMatrix(.Row, .ColIndex("����")) = 8
            End Select
            .TextMatrix(.Row, .ColIndex("�ı�")) = 2
        End If
    End With
End Sub

Private Sub SetcmbAlign()
    Dim strAlign As String
    Dim lngColWith As Long
    
    With vfgColumn
        If .Row > 0 Then
            strAlign = .TextMatrix(.Row, .ColIndex("����"))
            lngColWith = Val(.TextMatrix(.Row, .ColIndex("�п�")))
        Else
            Exit Sub
        End If
    End With

    Select Case Val(strAlign)
    Case 0
        cmbAlign.Text = "���϶���"
    Case 1
        cmbAlign.Text = "���ж���"
    Case 2
        cmbAlign.Text = "���¶���"
    Case 3
        cmbAlign.Text = "���϶���"
    Case 4
        cmbAlign.Text = "���ж���"
    Case 5
        cmbAlign.Text = "���¶���"
    Case 6
        cmbAlign.Text = "���϶���"
    Case 7
        cmbAlign.Text = "���ж���"
    Case 8
        cmbAlign.Text = "���¶���"
    End Select
    txtWidth.Text = lngColWith
End Sub

Private Sub setenabled()
    With vfgColumn
        If .Row > 0 Then
            If .Row = 1 Then
                cmdUp.Enabled = False
            Else
                cmdUp.Enabled = True
            End If
            If .Row = .Rows - 1 Then
                cmdDown.Enabled = False
            Else
                cmdDown.Enabled = True
            End If
        Else
            cmdUp.Enabled = False
            cmdDown.Enabled = False
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    mblnOK = False
End Sub

Private Sub cmdClear_Click()
    Dim lngRow As Long
    
    With vfgColumn
        For lngRow = 1 To .Rows - 1
             If Val(.RowData(lngRow)) = 0 Then
                If .TextMatrix(lngRow, .ColIndex("ѡ��")) = "0" Then
                Else
                    .TextMatrix(lngRow, .ColIndex("�ı�")) = 0
                    .TextMatrix(lngRow, .ColIndex("ѡ��")) = False
                End If
            End If
        Next
    End With
End Sub

Private Sub CmdDown_Click()
    With vfgColumn
        If .Row = .Rows - 1 Then
        Else
            .Select .Row + 1, .ColIndex("ѡ��")
            Call vfgColumn_Click
        End If
        Call setenabled
    End With
End Sub

Private Sub cmdOK_Click()
    Call SaveData
    Unload Me
    mblnOK = True
End Sub

Private Sub cmdRestore_Click()
    Call LoadFulltoColSel
End Sub

Private Sub cmdSelect_Click()
    Dim lngRow As Long
    
    With vfgColumn
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, .ColIndex("ѡ��")) Then
            Else
                .TextMatrix(lngRow, .ColIndex("�ı�")) = 1
                .TextMatrix(lngRow, .ColIndex("ѡ��")) = True
            End If
        Next
    End With
End Sub

Private Sub CmdUP_Click()
    With vfgColumn
        If .Row = 1 Then
        Else
            .Select .Row - 1, .ColIndex("ѡ��")
            Call vfgColumn_Click
        End If
        Call setenabled
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnOK = False
End Sub

Private Sub txtWidth_Change()
    If Trim(txtWidth) <> "" Then Call IsValid
End Sub

Private Function IsValid() As Boolean
    Dim blnValid As Boolean
    
    blnValid = True
    If IsNumeric(txtWidth) = False Then
        blnValid = False
        MsgBox "������һ���Ϸ�����ֵ��", vbInformation, gstrSysName
    Else
        If Val(txtWidth.Text) > 10000 Or Val(txtWidth.Text) < 0 Then
            MsgBox "������һ��С��10000��������", vbInformation, gstrSysName
            blnValid = False
        End If
    End If
    IsValid = blnValid
End Function

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyTab And KeyAscii <> vbKeyReturn And _
        KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> Asc(".") Then KeyAscii = 0
End Sub

Private Sub txtWidth_Validate(Cancel As Boolean)
    Cancel = Not IsValid
    If Cancel = False Then
        With vfgColumn
            .TextMatrix(.Row, .ColIndex("�п�")) = txtWidth
            .TextMatrix(.Row, .ColIndex("�ı�")) = 2
        End With
    End If
End Sub

Private Sub vfgColumn_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '�޸ĺ�
    Dim strColKey As String, blnShow As Boolean
    With vfgColumn
        Select Case Col
        Case .ColIndex("ѡ��")
            blnShow = GetVsGridBoolColVal(vfgColumn, Row, .ColIndex("ѡ��"))
            If blnShow Then
                 .TextMatrix(Row, .ColIndex("�ı�")) = 1
            Else
                 .TextMatrix(Row, .ColIndex("�ı�")) = 0
            End If
        Case Else
        End Select
    End With
End Sub

Private Sub vfgColumn_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vfgColumn
        Select Case Col
        Case .ColIndex("ѡ��")
            'rowdata(i):1-�̶�,-1-����ѡ,0-��ѡ
            If Val(.RowData(Row)) = 1 Then
                Cancel = True
            End If
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub initVfgColumnTitle()
    Dim strHead As String
    strHead = "ѡ��,300,1,1;����,2000,1,1;��ע,1000,1,1;�п�,0,7,-1;����,0,7,-1;�̶�,0,7,-1;ԭֵ,0,7,-1;�ı�,0,7,-1"
    Call SetVsFlexGridChangeHead(strHead, vfgColumn, 0)
End Sub

Private Sub vfgColumn_Click()
    Call SetcmbAlign
End Sub
