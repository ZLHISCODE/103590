VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSquareCardParaSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7515
   Icon            =   "frmSquareCardParaSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   7515
   StartUpPosition =   1  '����������
   Begin VB.Frame fraTitle 
      Caption         =   "���ع������ѿ�"
      Height          =   1965
      Left            =   60
      TabIndex        =   6
      Top             =   120
      Width           =   7365
      Begin VSFlex8Ctl.VSFlexGrid vsBill 
         Height          =   1635
         Left            =   75
         TabIndex        =   7
         Top             =   270
         Width           =   7215
         _cx             =   12726
         _cy             =   2884
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
         GridColor       =   8421504
         GridColorFixed  =   8421504
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSquareCardParaSet.frx":030A
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
         ExplorerBar     =   2
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
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "�ɿ��ӡ����"
      Height          =   360
      Left            =   3750
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2185
      Width           =   1875
   End
   Begin VB.CheckBox chk������ֵ 
      Caption         =   "��ֵ���˳���ֵ����(&N)"
      Height          =   240
      Left            =   60
      TabIndex        =   4
      Top             =   2245
      Width           =   2400
   End
   Begin VB.Frame fraSplit 
      Height          =   120
      Index           =   1
      Left            =   -30
      TabIndex        =   3
      Top             =   2610
      Width           =   7875
   End
   Begin VB.CommandButton cmdDeviceSetup 
      Caption         =   "�豸����(&S)"
      Height          =   350
      Left            =   5790
      TabIndex        =   2
      Top             =   2190
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4800
      TabIndex        =   1
      Top             =   3000
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5985
      TabIndex        =   0
      Top             =   3000
      Width           =   1100
   End
End
Attribute VB_Name = "frmSquareCardParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mstrPrivs   As String, mblnFirst As Boolean, mblnChange As Boolean

Public Sub ShowParaSet(ByVal frmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '���:frmMain-������
    '     lngModule-ģ���
    '     strPrivs-Ȩ�޴�

    '����:���˺�
    '����:2009-11-19 15:29:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs: mblnFirst = True
    Me.Show 1, frmMain
End Sub

Private Sub LoadParaSet()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز�������
    '����:���˺�
    '����:2009-12-10 17:03:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���� As String, varData As Variant
    Dim blnIsHavePriv As Boolean
    blnIsHavePriv = InStr(1, mstrPrivs, ";��������;") > 0
    chk������ֵ.value = IIf(Val(zlDatabase.GetPara("������ֵ", glngSys, mlngModule, , Array(chk������ֵ), blnIsHavePriv)) = 1, 1, 0)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function SaveSet() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ݿⱣ���������
    '����:����ɹ�����True,���򷵻�False
    '����:���˺�
    '����:2009-12-10 16:59:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, i As Long
    Dim strValue As String
    
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    Err = 0: On Error GoTo Errhand:
    '���湲�����ѿ�
    strValue = ""
    With vsBill
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("ѡ��"))) <> 0 And Val(.RowData(i)) <> 0 Then
                strValue = strValue & "|" & Val(.RowData(i)) & "," & Trim(.Cell(flexcpData, i, .ColIndex("���ѿ����")))
            End If
        Next
    End With
    If strValue <> "" Then strValue = Mid(strValue, 2)
    zlDatabase.SetPara "�������ѿ�����", strValue, glngSys, mlngModule, blnHavePrivs
   
    Call zlDatabase.SetPara("������ֵ", IIf(chk������ֵ.value = 1, 1, 0), glngSys, mlngModule, blnHavePrivs)
    SaveSet = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    SaveErrLog
End Function

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, glngSys, mlngModule)
End Sub

Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ч�Լ��
    '����:���Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-06 18:39:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lngSelCount As Long, str��� As String
    
    On Error GoTo errHandle
    '���ÿ��ʹ����ʽֻ��һ��ѡ��
    With vsBill
        str��� = "-"
        For i = 1 To vsBill.Rows - 1
            If str��� <> Trim(.TextMatrix(i, .ColIndex("���ѿ����"))) Then
               str��� = Trim(.TextMatrix(i, .ColIndex("���ѿ����")))
               lngSelCount = 0
                For j = 1 To vsBill.Rows - 1
                    If Trim(.TextMatrix(i, .ColIndex("���ѿ����"))) = Trim(.TextMatrix(j, .ColIndex("���ѿ����"))) Then
                        If Val(.TextMatrix(j, .ColIndex("ѡ��"))) <> 0 Then
                            lngSelCount = lngSelCount + 1
                        End If
                    End If
                Next
                If lngSelCount > 1 Then
                    MsgBox "ע��:" & vbCrLf & "    ���ѿ����Ϊ��" & str��� & "����ֻ��ѡ��һ��Ʊ��,����!", vbInformation + vbOKOnly
                    Exit Function
                End If
            End If
        Next
    End With
    
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdOK_Click()
    If isValied = False Then Exit Sub
    If SaveSet = False Then Exit Sub
    Unload Me
End Sub

Private Sub cmdPrintSet_Click()
    Dim strBill As String
    strBill = "ZL1_BILL_1503"
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub
 
Private Sub Form_Load()
    Call InitShareInvoice
    Call LoadParaSet
End Sub

Private Sub InitShareInvoice()
    Dim rsTemp As ADODB.Recordset, lngRow As Long
    Dim strShareInvoice As String '����Ʊ������,��ʽ:����,����
    Dim varData As Variant, varTemp As Variant, VarType As Variant, varTemp1 As Variant
    Dim intType As Integer, intType1 As Integer   '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    Dim lngTemp As Long, i As Long, strSql As String
    Dim strPrintMode As String, blnHavePrivs As Boolean
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    '�ָ��п��
    
    On Error GoTo errHandle
    zl_vsGrid_Para_Restore mlngModule, vsBill, Me.Name, "�������ѿ��б�", False, False
    strShareInvoice = zlDatabase.GetPara("�������ѿ�����", glngSys, mlngModule, , , True, intType)
    '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    vsBill.Tag = ""
    Select Case intType
    Case 1, 3, 5, 15
        vsBill.ForeColor = vbBlue: vsBill.ForeColorFixed = vbBlue
        fraTitle.ForeColor = vbBlue: vsBill.Tag = 1
        If intType = 5 Then vsBill.Tag = ""
    Case Else
        vsBill.ForeColor = &H80000008: vsBill.ForeColorFixed = &H80000008
        fraTitle.ForeColor = &H80000008
    End Select
    With vsBill
        .Editable = flexEDKbdMouse
        If Val(.Tag) = 1 And InStr(1, mstrPrivs, ";��������;") = 0 Then .Editable = flexEDNone
    End With
    
    '��ʽ:����ID1,���ѿ����ID1|����IDn,���ѿ����IDn|...
    varData = Split(strShareInvoice, "|")

    '1.���ù���Ʊ��
    Set rsTemp = GetShareInvoiceGroupID(6)
    With vsBill
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(Nvl(rsTemp!id))
            .TextMatrix(lngRow, .ColIndex("���ѿ����")) = Nvl(rsTemp!ʹ�����)
            .Cell(flexcpData, lngRow, .ColIndex("���ѿ����")) = Val(Nvl(rsTemp!ʹ�����ID))
            .TextMatrix(lngRow, .ColIndex("������")) = Nvl(rsTemp!������)
            .TextMatrix(lngRow, .ColIndex("��������")) = Format(rsTemp!�Ǽ�ʱ��, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("���뷶Χ")) = rsTemp!��ʼ���� & "," & rsTemp!��ֹ����
            .TextMatrix(lngRow, .ColIndex("ʣ��")) = Format(Val(Nvl(rsTemp!ʣ������)), "##0;-##0;;")
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                lngTemp = Val(varTemp(0))
                If Val(.RowData(lngRow)) = lngTemp _
                    And Val(varTemp(1)) = Val(.Cell(flexcpData, lngRow, .ColIndex("���ѿ����"))) Then
                    .TextMatrix(lngRow, .ColIndex("ѡ��")) = -1: Exit For
                End If
            Next
            .MergeRow(lngRow) = True
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsBill, Me.Name, "�������ѿ��б�", False, False
End Sub

Private Sub vsBill_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsBill, Me.Name, "�������ѿ��б�", False, False
End Sub

Private Sub vsBill_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsBill, Me.Name, "�������ѿ��б�", False, False
End Sub

Private Sub vsBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    With vsBill
        Select Case Col
        Case .ColIndex("ѡ��")
            If Val(.TextMatrix(Row, Col)) <> 0 And Val(.RowData(Row)) <> 0 Then
                For i = 1 To .Rows - 1
                    If Val(.Cell(flexcpData, Row, .ColIndex("���ѿ����"))) = Val(.Cell(flexcpData, i, .ColIndex("���ѿ����"))) _
                        And i <> Row Then
                        .TextMatrix(i, Col) = 0
                    End If
                Next
            End If
        End Select
    End With
End Sub

Private Sub vsBill_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsBill
        If Val(.Tag) = 1 Then
            If InStr(1, mstrPrivs, ";��������;") = 0 Then Cancel = True: Exit Sub
        End If
        Select Case Col
        Case .ColIndex("ѡ��")
            If Val(.RowData(Row)) = 0 Then Cancel = True
        Case Else
            Cancel = True
        End Select
    End With
End Sub
