VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmIdentifySet 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��֤�ӿ�����"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8790
   Icon            =   "frmIdentifySet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   8790
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame frmIdentify 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "������֤�ӿ�"
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "ȡ��"
         Height          =   350
         Left            =   7320
         TabIndex        =   3
         Top             =   2760
         Width           =   1100
      End
      Begin VB.CommandButton cmdOk 
         Appearance      =   0  'Flat
         Caption         =   "ȷ��"
         Height          =   350
         Left            =   6120
         TabIndex        =   2
         Top             =   2760
         Width           =   1100
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfInterface 
         Height          =   2295
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   8295
         _cx             =   14631
         _cy             =   4048
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
         BackColorSel    =   4210752
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   9
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   325
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
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
   End
   Begin VB.Image imgDelete 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   0
      Picture         =   "frmIdentifySet.frx":6852
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAdd 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   495
      Picture         =   "frmIdentifySet.frx":7254
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmIdentifySet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmParent As Object
Private mrsSecdInfo As ADODB.Recordset
Private mrsIneterface As New ADODB.Recordset
Public Enum Cert_Interface
    COL_ID = 0
    COL_���
    COL_�ӿ���
    COL_������
    COL_˵��
    COL_�Ƿ�����
    COL_Add
    COL_Del
End Enum
Private Enum Change_State
    CS_ɾ���� = -1
    CS_δ�ı� = 0
    CS_������ = 1
    CS_�滻�� = 2
    CS_������ = 3
End Enum

Private Sub InitVsfGridHeader()
'���ܣ���ʼ���б�
    Dim strHeader As String
    strHeader = "ID;���;�ӿ���,2000,1;������,2000,1;˵��,2800,1;�Ƿ�����,900,4;,270,4;,270,4"
    Call grid.Init(vsfInterface, strHeader, , , 1)
    With vsfInterface
        .ColDataType(.ColIndex("�Ƿ�����")) = flexDTBoolean
        .TextMatrix(.FixedRows, COL_�Ƿ�����) = 0
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If SaveInterface Then
        Unload Me
    End If
End Sub

Private Function SaveInterface() As Boolean
'���ܣ�������֤�ӿ�������Ϣ
    Dim arrSQL() As Variant
    Dim blnTrans As Boolean
    Dim i As Long
    
    arrSQL = Array()
    On Error GoTo errH
    Call CachCertInterface(arrSQL)
    gcnOracle.BeginTrans: blnTrans = True
    For i = LBound(arrSQL) To UBound(arrSQL)
        Debug.Print CStr(arrSQL(i))
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    SaveInterface = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Load()
    Call InitVsfGridHeader
    Call InitBaseInfo
    Call LoadInterface
End Sub

Public Function ShowMe(frmParent As Object) As Boolean
    Set mfrmParent = frmParent
    If Not mfrmParent Is Nothing Then
        Me.Show , mfrmParent
    End If
End Function

Private Function LoadInterface() As Boolean
'���������ӿ���Ϣ
    On Error GoTo errH
    Set mrsIneterface = LoadCertInterface(1)
    If Not mrsIneterface.EOF Then
        Call LoadCachInterface(mrsIneterface)
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Sub LoadCachInterface(ByVal rsTmp As ADODB.Recordset)
'���ܣ���֤����Ϣ���ز�����

    Dim strTmp As String
    Dim i As Long, j As Long, k As Long, lngRow As Long
    Dim lngPos As Long
    Dim strInfo As String, strMainInfo As String
    Dim strsInfo As String, strsMainInfo As String
    Dim arrWhole As Variant, arrMain As Variant
    Dim lngTmp As Long
    Dim lngsTmp As Long
    Dim strType As String
    Dim rsImg As New ADODB.Recordset
    Dim strFile As String
    Dim objFile As New FileSystemObject
    
    On Error GoTo errH
    
     'ɾ��֮ǰ�Ļ���
    mrsSecdInfo.Filter = "�ؼ���='vsfInterface'"
    If Not mrsSecdInfo.EOF Then
        For i = 1 To mrsSecdInfo.RecordCount
            mrsSecdInfo.Delete
            mrsSecdInfo.Update
            mrsSecdInfo.MoveNext
        Next
    End If

    lngTmp = 1
    With vsfInterface
        .Rows = .FixedRows
        For i = 1 To rsTmp.RecordCount
            .Rows = .Rows + 1: lngRow = .Rows - 1
            .TextMatrix(lngRow, COL_ID) = "" & rsTmp!ID
            .TextMatrix(lngRow, COL_���) = "" & rsTmp!���
            .TextMatrix(lngRow, COL_�ӿ���) = "" & rsTmp!�ӿ���
            .TextMatrix(lngRow, COL_������) = "" & rsTmp!������
            .TextMatrix(lngRow, COL_˵��) = "" & rsTmp!˵��
            .TextMatrix(lngRow, COL_�Ƿ�����) = IIf(Val("" & rsTmp!�Ƿ�����) = 1, -1, 0)
            .Cell(flexcpPicture, lngRow, COL_Add, lngRow, COL_Add) = imgAdd
            .Cell(flexcpPictureAlignment, lngRow, COL_Add, lngRow, COL_Add) = 4
            
            .Cell(flexcpPicture, lngRow, COL_Del, lngRow, COL_Del) = imgDelete
            .Cell(flexcpPictureAlignment, lngRow, COL_Del, lngRow, COL_Del) = 4
            
            .RowData(lngRow) = Val(rsTmp!ID & "")
                
            strMainInfo = rsTmp!ID & "|" & rsTmp!��� & "|" & rsTmp!�ӿ��� & "|" & rsTmp!������ & "|" & rsTmp!˵�� & "|" & Val("" & rsTmp!�Ƿ�����)
            strInfo = strMainInfo
            mrsSecdInfo.AddNew Array("���", "ԭID", "�ؼ���", "��Ϣԭֵ", "����Ϣԭֵ"), Array(lngTmp, Val(rsTmp!ID & ""), "vsfInterface", strInfo, strMainInfo)
            lngTmp = lngTmp + 1
            rsTmp.MoveNext
        Next
        .Row = 1: .Col = COL_�ӿ���
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitBaseInfo()
    Set mrsSecdInfo = New ADODB.Recordset
    With mrsSecdInfo
        .Fields.Append "Sort", adInteger                                              '����¼��������
        .Fields.Append "���", adInteger                                              '��ʶ��Ϣ����������¼��
        .Fields.Append "�ؼ���", adVarChar, 100                                       'չʾ��Ϣ�Ŀؼ�����
        .Fields.Append "IndexEx", adInteger, , adFldIsNullable                        '�кŻ�ؼ�����Index
        .Fields.Append "ҳ��", adInteger                                              '��Ϣ���ڵ�ҳ��
        .Fields.Append "ԭID", adBigInt, , adFldIsNullable
        .Fields.Append "��Ϣԭֵ", adVarChar, 2000, adFldIsNullable      '��Ϣ�ڼ���ʱ��ֵ
        .Fields.Append "����Ϣԭֵ", adVarChar, 2000, adFldIsNullable    '��Ϣ����Ҫ���֣���ʶһ����Ϣ�Ƿ񱻳��׸ı䣬��Ϣ�ڼ���ʱ��ֵ
        .Fields.Append "��ID", adBigInt, , adFldIsNullable
        .Fields.Append "��Ϣ��ֵ", adVarChar, 2000, adFldIsNullable      '��Ϣ�ڼ��ʱ��ֵ
        .Fields.Append "����Ϣ��ֵ", adVarChar, 2000, adFldIsNullable    '��Ϣ�ڼ��ʱ��ֵ
        .Fields.Append "�ı�״̬", adInteger                             '��Ϣ�ı�̶�0-δ�ı䣬1-�μ���Ϣ�ı䣬2-����Ϣ�ı�,3-����,-1��ɾ��
        .Fields.Append "ID", adBigInt, , adFldIsNullable                 '��Ϣ�������ݿ��е�ID,һ������ؼ�ʹ��
        .Fields.Append "Tag", adVarChar, 2000                            '�洢��������
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
End Sub

Private Function CachCertInterface(ByRef arrSQL As Variant) As Boolean
'���ܣ���֤����Ϣ����
    Dim i As Long, j As Long, k As Long
    Dim lng״̬ As Long
    Dim strTmp As String
    Dim strInfo As String
    Dim strMainInfo As String
    Dim strDels As String
    Dim strAll As String
    Dim lngRow As Long
    Dim strVsName As String
    Dim arrWhole As Variant
    Dim arrOther As Variant
    Dim arrMain As Variant
    Dim DatCur As Date
    Dim lngID As Long
    Dim lngTmp As Long
    
    On Error GoTo errH
    With vsfInterface
        .Tag = ""
        lngTmp = 1
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, COL_�ӿ���) <> "" And .TextMatrix(i, COL_������) <> "" Then
                strMainInfo = Val(.TextMatrix(i, COL_ID)) & "|" & .TextMatrix(i, COL_���) & "|" & .TextMatrix(i, COL_�ӿ���) & "|" & .TextMatrix(i, COL_������) & "|" & .TextMatrix(i, COL_˵��) & "|" & IIf(Nvl(.TextMatrix(i, COL_�Ƿ�����), -1) = -1, 1, 0)
                strInfo = strMainInfo
                .RowData(i) = .TextMatrix(i, COL_ID)
                If InStr("," & strAll & ",", "," & strMainInfo & ",") > 0 Then
                    '��ͬ��ÿ��¼
                    .Tag = i
                    .Cell(flexcpBackColor, i, .FixedCols, i, COL_�ӿ���) = &HC0C0FF
                    Call .ShowCell(i, COL_�ӿ���)
                    Exit Function
                Else
                    strAll = strAll & "," & strMainInfo '�ռ����������ж��Ƿ����ظ���
                End If
            Else
                strMainInfo = ""
                strInfo = ""
            End If
               mrsSecdInfo.Filter = "�ؼ���='vsfInterface' and ���=" & lngTmp
               If mrsSecdInfo.EOF Then
                   mrsSecdInfo.AddNew
                   mrsSecdInfo!��� = lngTmp
                   mrsSecdInfo!�ؼ��� = "vsfInterface"
               End If
               mrsSecdInfo!��ID = Val(.RowData(i))
               mrsSecdInfo!��Ϣ��ֵ = IIf(strInfo = "", Null, strInfo)
               mrsSecdInfo!����Ϣ��ֵ = IIf(strMainInfo = "", Null, strMainInfo)
               mrsSecdInfo!IndexEx = i
               mrsSecdInfo.Update
               lngTmp = lngTmp + 1

               mrsSecdInfo.Filter = 0
        Next
        mrsSecdInfo.Filter = "�ؼ���='vsfInterface'"
        For i = 1 To mrsSecdInfo.RecordCount
            lng״̬ = CS_δ�ı�
            If mrsSecdInfo!��Ϣԭֵ & "" <> mrsSecdInfo!��Ϣ��ֵ & "" Then
                lng״̬ = CS_������
            End If
            If lng״̬ = CS_������ And IsNull(mrsSecdInfo!��Ϣԭֵ) Then
                lng״̬ = CS_������
            End If
            If lng״̬ = CS_������ And IsNull(mrsSecdInfo!��Ϣ��ֵ) Then
                lng״̬ = CS_ɾ����
            End If
            If lng״̬ = CS_������ And mrsSecdInfo!����Ϣԭֵ & "" <> mrsSecdInfo!����Ϣ��ֵ & "" Then
                lng״̬ = CS_�滻��
            End If
            mrsSecdInfo.Update "�ı�״̬", lng״̬
            mrsSecdInfo.MoveNext
        Next
        

        '����Ϣ�ı�����Ҫ����ɾ������
        mrsSecdInfo.Filter = "(�ı�״̬=" & CS_ɾ���� & " And �ؼ���='vsfInterface')" ' OR (�ı�״̬=" & CS_�滻�� & " And �ؼ���='vsfCert')"
        Do While Not mrsSecdInfo.EOF
            strDels = "" & mrsSecdInfo!ԭID
            If strDels <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_ʵ����֤�ӿ�_Delete(" & Val(strDels) & ")"
            End If
            mrsSecdInfo.MoveNext
        Loop

        '����Ϣ�ı��Լ���������Ҫ���ò������        '�μ���Ϣ�ı䣬���ø��¹���
        mrsSecdInfo.Filter = "�ؼ���='vsfInterface' And �ı�״̬>" & CS_δ�ı�
        Do While Not mrsSecdInfo.EOF
            lngRow = mrsSecdInfo!IndexEx
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            If mrsSecdInfo!�ı�״̬ = CS_������ Then
                arrSQL(UBound(arrSQL)) = "Zl_ʵ����֤�ӿ�_Insert(" & "'" & .TextMatrix(lngRow, COL_�ӿ���) & "','" & .TextMatrix(lngRow, COL_������) & "','" & .TextMatrix(lngRow, COL_˵��) & "'," & IIf(Nvl(.TextMatrix(lngRow, COL_�Ƿ�����), -1) = -1, 1, 0) & ")"
            Else
                arrSQL(UBound(arrSQL)) = "Zl_ʵ����֤�ӿ�_Update(" & Val(.TextMatrix(lngRow, COL_ID)) & ",'" & .TextMatrix(lngRow, COL_���) & "','" & .TextMatrix(lngRow, COL_�ӿ���) & "','" & _
                        .TextMatrix(lngRow, COL_������) & "','" & .TextMatrix(lngRow, COL_˵��) & "'," & IIf(Nvl(.TextMatrix(lngRow, COL_�Ƿ�����), -1) = -1, 1, 0) & ")"
            End If
            mrsSecdInfo.MoveNext
        Loop
    End With
    CachCertInterface = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set mfrmParent = Nothing
    Set mrsSecdInfo = Nothing
    Set mrsIneterface = Nothing
End Sub

Private Sub vsfInterface_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngNewCol As Long, lngNewRow As Long
    
    lngNewCol = NewCol
    lngNewRow = NewRow
    If lngNewCol = -1 Then Exit Sub
    With vsfInterface
        If lngNewCol = COL_Del Or lngNewCol = COL_Add Then
             .ComboList = "..."
             .FocusRect = flexFocusNone
             Set .CellButtonPicture = IIf(lngNewCol = COL_Del, imgDelete, imgAdd)
        Else
            .ComboList = ""
        End If
        If lngNewRow >= .FixedRows Then
            '��ʾͼƬ
            If lngNewCol <> COL_Add And .TextMatrix(lngNewRow, COL_�ӿ���) <> "" Then
                If .Rows - 1 <> lngNewRow Then
                    '��һ�����Ϊ������������
                    If .TextMatrix(lngNewRow + 1, COL_�ӿ���) = "" Then
                         Set .Cell(flexcpPicture, lngNewRow, COL_Add) = imgAdd
                    End If
                Else
                    Set .Cell(flexcpPicture, lngNewRow, COL_Add) = imgAdd
                End If
            End If
            '��ʾͼƬ
            If lngNewCol <> COL_Del Then Set .Cell(flexcpPicture, lngNewRow, COL_Del) = imgDelete
        End If
    End With
    
End Sub

Private Sub vsfInterface_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngCol As Long
    
    lngCol = Col
    If lngCol = COL_Add Or lngCol = COL_Del Then Cancel = True
End Sub

Private Sub vsfInterface_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim lngRow As Long, lngCol As Long, lngCount As Long
    Dim i As Long, j As Long
    Dim blnAdd As Boolean
    
    lngCol = Col
    lngRow = Row
    With vsfInterface
        Select Case lngCol
            Case COL_Add
                For i = .Rows - 1 To .FixedRows Step -1
                    If Trim(.TextMatrix(.Rows - 1, COL_�ӿ���)) <> "" And .RowHidden(.Rows - 1) = False Then
                        blnAdd = True
                        Exit For
                    ElseIf Trim(.TextMatrix(.Rows - 1, COL_�ӿ���)) = "" And .RowHidden(.Rows - 1) = False Then
                        Exit For
                    End If
                Next
                If blnAdd Then
                     lngRow = .Rows: .AddItem "", lngRow
                     .TextMatrix(lngRow, COL_�Ƿ�����) = 0
                     .Row = lngRow: .Col = COL_�ӿ���
                     .ShowCell .Row, COL_�ӿ���
                End If
                blnAdd = False
            Case COL_Del
                If Trim(.TextMatrix(lngRow, COL_�ӿ���)) <> "" Then
                    If MsgBox("ȷ��Ҫɾ���ӿ���Ϊ��" & .TextMatrix(lngRow, COL_�ӿ���) & "����֤����Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        If Val(.TextMatrix(lngRow, COL_ID)) <> 0 Then
                            If .Rows - 1 = .FixedRows Then
                                For i = COL_ID To COL_�Ƿ�����
                                    .TextMatrix(lngRow, i) = ""
                                    .Cell(flexcpData, lngRow, i, lngRow, i) = ""
                                Next
                            ElseIf .Rows - 1 > .FixedRows Then
                                .RemoveItem lngRow
                                .AddItem "", lngRow
                                .RowHidden(lngRow) = True
                            End If
                        Else
                             For i = .FixedRows To .Rows - 1
                                If .TextMatrix(i, COL_�ӿ���) <> "" Then
                                    lngCount = lngCount + 1
                                End If
                            Next
                            If lngCount = .FixedRows Then
                                For j = COL_ID To COL_�Ƿ�����
                                    .TextMatrix(lngRow, j) = ""
                                    .Cell(flexcpData, lngRow, j, lngRow, j) = ""
                                Next
                            End If
                        End If
                    Else
                        .Row = lngRow: .Col = COL_�ӿ���
                        .ShowCell .Row, .Col
                    End If
                Else
                    If .Rows - 1 = .FixedRows Or lngRow = .FixedRows Then
                        Exit Sub
                    Else
                        For i = .FixedRows To .Rows - 1
                            If .TextMatrix(i, COL_�ӿ���) <> "" Then
                                lngCount = lngCount + 1
                            End If
                        Next
                        If lngCount <> 0 Then
                            .RemoveItem lngRow
                        End If
                    End If
                End If
        End Select
    End With
End Sub

Private Sub vsfInterface_Click()
    Dim lngRow As Long, lngCol As Long
    
    With vsfInterface
        lngRow = .Row
        lngCol = .Col
        If (lngCol = COL_Add Or lngCol = COL_Del) And lngRow >= .FixedRows Then
            If lngCol = COL_Add Then
                If .TextMatrix(lngRow, COL_�ӿ���) = "" Then Exit Sub
            End If
            .Select lngRow, lngCol
            Call vsfInterface_CellButtonClick(lngRow, lngCol)
        End If
    End With
End Sub

Private Sub vsfInterface_KeyPress(KeyAscii As Integer)
    Dim lngRow As Long
    Dim lngCol As Long
      
    lngRow = vsfInterface.Row
    lngCol = vsfInterface.Col
    With vsfInterface
        If lngCol = COL_������ Then
            If zlCommFun.ActualLen(.TextMatrix(lngRow, COL_������)) >= 100 And UCase(Chr(KeyAscii)) <> Chr(8) And UCase(Chr(KeyAscii)) <> Chr(13) Then
                KeyAscii = 0
            End If
        ElseIf lngCol = COL_�ӿ��� Then
            If zlCommFun.ActualLen(.TextMatrix(lngRow, COL_�ӿ���)) >= 50 And UCase(Chr(KeyAscii)) <> Chr(8) And UCase(Chr(KeyAscii)) <> Chr(13) Then
                KeyAscii = 0
            End If
        ElseIf lngCol = COL_˵�� Then
            If zlCommFun.ActualLen(.TextMatrix(lngRow, COL_˵��)) >= 100 And UCase(Chr(KeyAscii)) <> Chr(8) And UCase(Chr(KeyAscii)) <> Chr(13) Then
                KeyAscii = 0
            End If
        End If
    End With
End Sub


Private Sub vsfInterface_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngRow As Long
    Dim lngCol As Long
      
    lngRow = Row
    lngCol = Col
    With vsfInterface
        If lngCol = COL_������ Then
            .TextMatrix(lngRow, COL_������) = .EditText
            If zlCommFun.ActualLen(.TextMatrix(lngRow, COL_������)) >= 100 Then
                MsgBox "���������ַ��������ܴ���100���ַ�����50�����֣�", vbInformation, gstrSysName
                Cancel = True
            End If
        ElseIf lngCol = COL_�ӿ��� Then
            .TextMatrix(lngRow, COL_�ӿ���) = .EditText
            If zlCommFun.ActualLen(.TextMatrix(lngRow, COL_�ӿ���)) >= 50 Then
                MsgBox "�ӿ������ַ��������ܴ���50���ַ�����25�����֣�", vbInformation, gstrSysName
                Cancel = True
            End If
        ElseIf lngCol = COL_˵�� Then
            .TextMatrix(lngRow, COL_˵��) = .EditText
            If zlCommFun.ActualLen(.TextMatrix(lngRow, COL_˵��)) >= 100 Then
                MsgBox "˵�����ַ��������ܴ���100���ַ�����50�����֣�", vbInformation, gstrSysName
                Cancel = True
            End If
        End If
    End With
End Sub


