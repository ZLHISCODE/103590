VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmOutAdviceSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ҽ��ѡ��"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   Icon            =   "frmOutAdviceSetUp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6750
      Left            =   45
      ScaleHeight     =   6750
      ScaleWidth      =   5160
      TabIndex        =   2
      Top             =   -240
      Width           =   5160
      Begin VSFlex8Ctl.VSFlexGrid vsfDrugStore 
         Height          =   5700
         Left            =   0
         TabIndex        =   4
         Top             =   975
         Width           =   5055
         _cx             =   8916
         _cy             =   10054
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
         BackColorBkg    =   14737632
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmOutAdviceSetUp.frx":000C
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
      Begin VB.Label lblKYYF 
         Caption         =   $"frmOutAdviceSetUp.frx":00B9
         Height          =   615
         Left            =   0
         TabIndex        =   3
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3990
      TabIndex        =   1
      Top             =   6570
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2895
      TabIndex        =   0
      Top             =   6570
      Width           =   1100
   End
End
Attribute VB_Name = "frmOutAdviceSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const VsPubBackColor = &HFAEADA

Private Sub cmdOK_Click()
    Dim i As Long, bytType As Long
    Dim blnSetup As Boolean
    Dim arr����ҩ��(4) As String, arrȱʡҩ��(4) As String, arrTmp() As String
    Dim str��ҩ������ As String, str��ҩ������ As String, str��ҩ������ As String
     
    '----------------------------------------------------------------------------------------------------
    'ҩ��
    With vsfDrugStore
        For i = .FixedRows To .Rows - 1
            Select Case .TextMatrix(i, .ColIndex("���"))
            Case "��ҩ��"
                bytType = 0
                If .TextMatrix(i, .ColIndex("��ҩ����")) <> "�Զ�����" Then
                    str��ҩ������ = str��ҩ������ & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("��ҩ����"))
                End If
            Case "��ҩ��"
                bytType = 1
                If .TextMatrix(i, .ColIndex("��ҩ����")) <> "�Զ�����" Then
                    str��ҩ������ = str��ҩ������ & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("��ҩ����"))
                End If
            Case "��ҩ��"
                bytType = 2
                If .TextMatrix(i, .ColIndex("��ҩ����")) <> "�Զ�����" Then
                    str��ҩ������ = str��ҩ������ & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("��ҩ����"))
                End If
            Case "���ϲ���"
                bytType = 3
            End Select
            If .TextMatrix(i, .ColIndex("����")) <> 0 Then arr����ҩ��(bytType) = arr����ҩ��(bytType) & "," & .RowData(i)
            If .TextMatrix(i, .ColIndex("ȱʡ")) = "��" Then arrȱʡҩ��(bytType) = .RowData(i)
        Next
    End With
    
    blnSetup = InStr(GetInsidePrivs(p����ҽ���´�), ";ҽ��ѡ������;") > 0
    blnSetup = True
    arrTmp = Split("��ҩ��,��ҩ��,��ҩ��,���ϲ���", ",")
    For bytType = 0 To UBound(arrTmp)
        Call zlDatabase.SetPara("�������" & arrTmp(bytType), Mid(arr����ҩ��(bytType), 2), glngSys, p����ҽ���´�, blnSetup)
        Call zlDatabase.SetPara("����ȱʡ" & arrTmp(bytType), arrȱʡҩ��(bytType), glngSys, p����ҽ���´�, blnSetup)
    Next
    Call zlDatabase.SetPara("��ҩ������", Mid(str��ҩ������, 2), glngSys, p����ҽ���´�, blnSetup)
    Call zlDatabase.SetPara("��ҩ������", Mid(str��ҩ������, 2), glngSys, p����ҽ���´�, blnSetup)
    Call zlDatabase.SetPara("��ҩ������", Mid(str��ҩ������, 2), glngSys, p����ҽ���´�, blnSetup)
           
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim strPar As String
    Dim blnSetup As Boolean, arrTmp() As String
    Dim strDSIDs As String, strDefault As String, lngBackColor As Long, bytLockEdit As Byte
    Dim intType1 As Integer, intType2 As Integer, lngRow As Long
    Dim str���� As String, j As Integer
 
    On Error GoTo errH
    
    blnSetup = InStr(GetInsidePrivs(p����ҽ���´�), "ҽ��ѡ������") > 0
 
    'ҩ���뷢�ϲ���
    strSQL = _
        "Select Distinct A.ID,A.����,A.����,B.�������� " & _
        " From ���ű� A,��������˵�� B " & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " AND B.����ID=A.ID And B.������� IN(1,3) and B.�������� in('��ҩ��','��ҩ��','��ҩ��','���ϲ���')" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " Order by ��������,����"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    With vsfDrugStore
        .Rows = .FixedRows
        .Editable = flexEDKbdMouse
        .MergeCol(.ColIndex("���")) = True
        .MergeCells = flexMergeFixedOnly
            
        If Not rsTmp.EOF Then
            .Rows = .FixedRows + rsTmp.RecordCount
            lngRow = .FixedRows
            arrTmp = Split("��ҩ��,��ҩ��,��ҩ��,���ϲ���", ",")
            For i = 0 To UBound(arrTmp)
                rsTmp.Filter = "��������='" & arrTmp(i) & "'"
                strDefault = zlDatabase.GetPara("����ȱʡ" & arrTmp(i), glngSys, p����ҽ���´�, , , , intType1)
                strDSIDs = "," & zlDatabase.GetPara("�������" & arrTmp(i), glngSys, p����ҽ���´�, , , , intType2) & ","
                
                '���� ��ҩ����  �У�1252 ģ����3������û���õ� '��ҩ������','��ҩ������','��ҩ������'  ����������
                .ColHidden(.ColIndex("��ҩ����")) = True
                
                '��ҩ����
                str���� = zlDatabase.GetPara(arrTmp(i) & "����", glngSys, p����ҽ���´�, , , blnSetup)
                Do While Not rsTmp.EOF
                    .TextMatrix(lngRow, .ColIndex("���")) = arrTmp(i)
                    .TextMatrix(lngRow, .ColIndex("ҩ��")) = rsTmp!����
                    .RowData(lngRow) = Val(rsTmp!ID)
                    
                    If Val(rsTmp!ID) = Val(strDefault) Then
                        .TextMatrix(lngRow, .ColIndex("ȱʡ")) = "��"
                        .TextMatrix(lngRow, .ColIndex("����")) = -1   'true
                    Else
                        .TextMatrix(lngRow, .ColIndex("ȱʡ")) = ""
                        .TextMatrix(lngRow, .ColIndex("����")) = IIF(InStr(strDSIDs, "," & rsTmp!ID & ",") > 0, -1, 0)
                    End If
                    
                    'ȱʡ��Ԫ��
                    'intType-'���ز������ͣ�1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
                    bytLockEdit = 0
                    If InStr(1, ",1,3,15,", "," & intType1 & ",") > 0 Then
                        lngBackColor = IIF(blnSetup, VsPubBackColor, &H8000000F)      '��Ȩ�޿���
                        bytLockEdit = IIF(blnSetup, 0, 1)
                    ElseIf intType1 = 5 Then
                        lngBackColor = VsPubBackColor       '����ģ��,������Ȩ�޿���
                    Else
                        lngBackColor = &H80000005     '�����༭
                    End If
                    .Cell(flexcpBackColor, lngRow, .ColIndex("ȱʡ")) = lngBackColor
                    .Cell(flexcpData, lngRow, .ColIndex("ȱʡ")) = bytLockEdit
                     
                    '���õ�Ԫ��
                    bytLockEdit = 0
                    If InStr(1, ",1,3,15,", "," & intType2 & ",") > 0 Then
                        lngBackColor = IIF(blnSetup, VsPubBackColor, &H8000000F)      '��Ȩ�޿���
                        bytLockEdit = IIF(blnSetup, 0, 1)
                    ElseIf intType2 = 5 Then
                        lngBackColor = VsPubBackColor       '����ģ��,������Ȩ�޿���
                    Else
                        lngBackColor = &H80000005     '�����༭
                    End If
                    .Cell(flexcpBackColor, lngRow, .ColIndex("����")) = lngBackColor
                    .Cell(flexcpData, lngRow, .ColIndex("����")) = bytLockEdit
                    
                    '��ҩ����
                    For j = 0 To UBound(Split(str����, ","))
                        If Val(.RowData(lngRow)) = Val(Split(Split(str����, ",")(j), ":")(0)) Then
                            .TextMatrix(lngRow, .ColIndex("��ҩ����")) = Split(Split(str����, ",")(j), ":")(1)
                            Exit For
                        End If
                    Next
                    If .TextMatrix(lngRow, .ColIndex("��ҩ����")) = "" Then .TextMatrix(lngRow, .ColIndex("��ҩ����")) = "�Զ�����"
                    .Cell(flexcpBackColor, lngRow, .ColIndex("��ҩ����")) = lngBackColor
                    .Cell(flexcpData, lngRow, .ColIndex("��ҩ����")) = bytLockEdit
                    
                    lngRow = lngRow + 1
                    rsTmp.MoveNext
                Loop
                If lngRow < .Rows - 1 Then  '���ָ���
                    .Select lngRow, .FixedCols, lngRow, .Cols - 1
                    .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
                End If
            Next
        End If
    End With
         
    cmdCancel.Left = Me.Left + Me.Width - cmdCancel.Width - 200
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 200
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfDrugStore_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsfDrugStore.ColIndex("����") Then
        Call Set����ҩ��(Row, True)
    ElseIf Col = vsfDrugStore.ColIndex("����") Then
        Call Setȱʡҩ��
    End If
    If Col <> vsfDrugStore.ColIndex("��ҩ����") Then Cancel = True
End Sub

Private Sub vsfDrugStore_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfDrugStore
        Select Case Col
        Case .ColIndex("����")
            Cancel = Val(.Cell(flexcpData, Row, Col)) <> 0
        Case .ColIndex("ȱʡ")
            Cancel = Val(.Cell(flexcpData, Row, Col)) <> 0
        Case .ColIndex("��ҩ����")
            Cancel = Val(.Cell(flexcpData, Row, Col)) <> 0
        Case Else
            Cancel = True
            Exit Sub
        End Select
    End With
End Sub

Private Sub vsfDrugStore_DblClick()
    With vsfDrugStore
        If .MouseCol = .ColIndex("ȱʡ") Then
            Call Setȱʡҩ��
        ElseIf .MouseCol = .ColIndex("ҩ��") Then
            Call Set����ҩ��(.Row, True)
        ElseIf .MouseCol = .ColIndex("����") And .MouseRow = .FixedRows - 1 Then
            Dim i As Long
            For i = .FixedRows To .Rows - 1
                Call Set����ҩ��(i)
            Next
        End If
    End With
End Sub

Private Sub vsfDrugStore_EnterCell()
    Dim rsTmp As ADODB.Recordset, strList As String
    With vsfDrugStore
        If .Row > 0 Then
            If .Col = .ColIndex("��ҩ����") Then
                Set rsTmp = Read��ҩ����(.RowData(.Row))
                .ColComboList(.Col) = "�Զ�����|" & .BuildComboList(rsTmp, "����")
                .FocusRect = flexFocusSolid
            Else
                .FocusRect = flexFocusLight
            End If
        End If
    End With
End Sub

Private Sub vsfDrugStore_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        If vsfDrugStore.Col = vsfDrugStore.ColIndex("ȱʡ") Then
            Call Setȱʡҩ��
        End If
    End If
End Sub

Private Sub Setȱʡҩ��()
'���ܣ����õ�ǰ�е�ȱʡҩ����ͬʱ������ͬ���͵������е�ȱʡҩ��
    Dim i As Long
    
    With vsfDrugStore
        If Val("" & .Cell(flexcpData, .Row, .ColIndex("ȱʡ"))) = 0 Then  '�ò��������޸ĵ������
            If .TextMatrix(.Row, .ColIndex("ȱʡ")) = "��" Then
                .TextMatrix(.Row, .ColIndex("ȱʡ")) = ""
            Else
                '��û����Ȩ���޸Ŀ���ʱ�ҿ���Ϊ0��false)ʱ����������ȱʡ
                If Not (Val(.TextMatrix(.Row, .ColIndex("����"))) = 0 And Val("" & .Cell(flexcpData, .Row, .ColIndex("����"))) = 1) Then
                    'ͬ����������ȡ��ȱʡ
                    For i = .FixedRows To .Rows - 1
                        If .TextMatrix(.Row, .ColIndex("���")) = .TextMatrix(i, .ColIndex("���")) Then
                            If .TextMatrix(i, .ColIndex("ȱʡ")) = "��" Then .TextMatrix(i, .ColIndex("ȱʡ")) = ""
                        End If
                    Next
                    .TextMatrix(.Row, .ColIndex("����")) = -1    '�Զ�����Ϊ����
                    .TextMatrix(.Row, .ColIndex("ȱʡ")) = "��"
                Else
                    MsgBox "���õ�ǰҩ��Ϊȱʡʱ����ͬʱ����ǰҩ������Ϊ���ã�" & vbNewLine & "��û���޸Ŀ���ҩ����Ȩ�ޡ�", vbInformation, gstrSysName
                End If
            End If
        Else
            MsgBox "��û���޸�ȱʡҩ����Ȩ�ޡ�", vbInformation, gstrSysName
        End If
    End With
End Sub

Private Sub Set����ҩ��(ByVal lngRow As Long, Optional ByVal blnAsk As Boolean = False)
'���ܣ����õ�ǰ�еĿ���ҩ����ͬʱ����ǰ�е�ȱʡҩ��

    With vsfDrugStore
        If Val("" & .Cell(flexcpData, lngRow, .ColIndex("����"))) = 0 Then   '�ò��������޸ĵ������
            If Val(.TextMatrix(lngRow, .ColIndex("����"))) = -1 Then
                '��ǰ���ҹ�ѡ����
                If Not (Val("" & .Cell(flexcpData, lngRow, .ColIndex("ȱʡ"))) = 1 And .TextMatrix(lngRow, .ColIndex("ȱʡ")) = "��") Then
                    .TextMatrix(lngRow, .ColIndex("����")) = 0
                    .TextMatrix(lngRow, .ColIndex("ȱʡ")) = ""
                Else
                    If blnAsk Then
                        MsgBox "ȡ����ǰҩ������ʱ����ͬʱȡ����ǰҩ��ȱʡ��" & vbNewLine & "��û���޸�ȱʡҩ����Ȩ�ޡ�", vbInformation, gstrSysName
                    End If
                End If
            Else
                .TextMatrix(lngRow, .ColIndex("����")) = -1    '�Զ�����Ϊ����
            End If
        Else
            If blnAsk Then
                MsgBox "��û���޸Ŀ���ҩ����Ȩ�ޡ�", vbInformation, gstrSysName
            End If
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
