VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm�ӳ������� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ӳ�������"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   Icon            =   "frm�ӳ�������.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6285
      TabIndex        =   1
      Top             =   5805
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7555
      TabIndex        =   2
      Top             =   5805
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   8295
   End
   Begin TabDlg.SSTab sstCustom 
      Height          =   4815
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "��ҩ/��ҩ"
      TabPicture(0)   =   "frm�ӳ�������.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vsfList"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "��ҩ"
      TabPicture(1)   =   "frm�ӳ�������.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "vsfListHerbal"
      Tab(1).ControlCount=   1
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   4335
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   8295
         _cx             =   14631
         _cy             =   7646
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
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm�ӳ�������.frx":0342
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
      Begin VSFlex8Ctl.VSFlexGrid vsfListHerbal 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   6
         Top             =   360
         Width           =   8295
         _cx             =   14631
         _cy             =   7646
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
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm�ӳ�������.frx":0442
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
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   165
      Picture         =   "frm�ӳ�������.frx":0542
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblInfor 
      Caption         =   $"frm�ӳ�������.frx":0BC3
      Height          =   480
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   7665
   End
End
Attribute VB_Name = "frm�ӳ�������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrValue As String
Private mintҩƷ���� As Integer '0-��ҩ/��ҩ,1-��ҩ

Private Enum mCol
    ��� = 0
    ��ͼ� = 1
    ��߼� = 2
    �ӳ���
    ��۶�
    ˵��
    intCount
End Enum

Public Sub ShowMe(ByVal objFra As frmMediLists)
    Me.Show vbModal, objFra
End Sub

Private Sub cmdCancel_Click()
    Dim intRow As Integer
    Dim intCol As Integer
    Dim strValue As String
    
    With vsfList
        For intRow = 1 To .Rows - 1
            For intCol = 1 To .Cols - 1
                strValue = strValue + .TextMatrix(intRow, intCol) & "|"
            Next
        Next
    End With
      
    With vsfListHerbal
        For intRow = 1 To .Rows - 1
            For intCol = 1 To .Cols - 1
                strValue = strValue + .TextMatrix(intRow, intCol) & "|"
            Next
        Next
    End With
    
    If strValue <> mstrValue Then
        If MsgBox("���ݱ��޸��ˣ��Ƿ��˳���", vbYesNo, gstrSysName) = vbYes Then
            Unload Me
        End If
    Else
        Unload Me
    End If
    
End Sub

Private Sub InitVsf()
    With vsfList
        .Rows = 2
        .Editable = flexEDKbdMouse
        .ExtendLastCol = True
        .Cols = mCol.intCount
        .TextMatrix(1, mCol.���) = 1
        .Col = 1
        .Row = 1
        .AllowUserResizing = flexResizeColumns
    End With
    
    With vsfListHerbal
        .Rows = 2
        .Editable = flexEDKbdMouse
        .ExtendLastCol = True
        .Cols = mCol.intCount
        .TextMatrix(1, mCol.���) = 1
        .Col = 1
        .Row = 1
        .AllowUserResizing = flexResizeColumns
    End With
End Sub

Private Sub FillVSF()
    Dim i As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    gstrSql = "select ���, ��ͼ�, ��߼�, �ӳ���, ��۶�, ˵��,���� from ҩƷ�ӳɷ���  order by ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "��ѯҩƷ�ӳɷ���")
    
    rsTemp.Filter = "����=0"
    With vsfList
        If rsTemp.RecordCount <> 0 Then
            .Rows = rsTemp.RecordCount + 1
            For i = 1 To .Rows - 1
                .TextMatrix(i, mCol.���) = i
                .TextMatrix(i, mCol.��ͼ�) = formatex(IIf(IsNull(rsTemp!��ͼ�), "", rsTemp!��ͼ�), 2)
                .TextMatrix(i, mCol.��߼�) = formatex(IIf(IsNull(rsTemp!��߼�), "", rsTemp!��߼�), 2)
                .TextMatrix(i, mCol.�ӳ���) = formatex(IIf(IsNull(rsTemp!�ӳ���), "", rsTemp!�ӳ���), 2)
                .TextMatrix(i, mCol.��۶�) = formatex(IIf(IsNull(rsTemp!��۶�), "", rsTemp!��۶�), 2)
                .TextMatrix(i, mCol.˵��) = IIf(IsNull(rsTemp!˵��), "", rsTemp!˵��)
                rsTemp.MoveNext
            Next
        End If
        
        For intRow = 1 To .Rows - 1
            For intCol = 1 To .Cols - 1
                mstrValue = mstrValue + .TextMatrix(intRow, intCol) & "|"
            Next
        Next
    End With
    
    rsTemp.Filter = "����=1"
    With vsfListHerbal
        If rsTemp.RecordCount <> 0 Then
            .Rows = rsTemp.RecordCount + 1
            For i = 1 To .Rows - 1
                .TextMatrix(i, mCol.���) = i
                .TextMatrix(i, mCol.��ͼ�) = formatex(IIf(IsNull(rsTemp!��ͼ�), "", rsTemp!��ͼ�), 2)
                .TextMatrix(i, mCol.��߼�) = formatex(IIf(IsNull(rsTemp!��߼�), "", rsTemp!��߼�), 2)
                .TextMatrix(i, mCol.�ӳ���) = formatex(IIf(IsNull(rsTemp!�ӳ���), "", rsTemp!�ӳ���), 2)
                .TextMatrix(i, mCol.��۶�) = formatex(IIf(IsNull(rsTemp!��۶�), "", rsTemp!��۶�), 2)
                .TextMatrix(i, mCol.˵��) = IIf(IsNull(rsTemp!˵��), "", rsTemp!˵��)
                rsTemp.MoveNext
            Next
        End If
        
        For intRow = 1 To .Rows - 1
            For intCol = 1 To .Cols - 1
                mstrValue = mstrValue + .TextMatrix(intRow, intCol) & "|"
            Next
        Next
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    Dim str˵�� As String
    Dim blnTrans As Boolean
    Dim arrSql() As Variant     '��¼�洢���̵�����
    
    On Error GoTo ErrHandle
    arrSql = Array()
    If Validate = False Then Exit Sub
    
    gstrSql = " ZL_ҩƷ�ӳɷ���_DELETE"
    ReDim Preserve arrSql(UBound(arrSql) + 1)
    arrSql(UBound(arrSql)) = gstrSql
    
    If vsfList.TextMatrix(1, 1) <> "" And vsfList.TextMatrix(1, 2) <> "" Then
        With vsfList
            For i = 1 To .Rows - 1
                str˵�� = IIf(.TextMatrix(i, mCol.˵��) = "", "Null", "'" & .TextMatrix(i, mCol.˵��) & "'")
                gstrSql = "Zl_ҩƷ�ӳɷ���_Insert(" & .TextMatrix(i, mCol.���) & "," & .TextMatrix(i, mCol.��ͼ�) & "," & .TextMatrix(i, mCol.��߼�) & _
                             "," & .TextMatrix(i, mCol.�ӳ���) & "," & IIf(.TextMatrix(i, mCol.��۶�) = "", "Null", .TextMatrix(i, mCol.��۶�)) & "," & str˵�� & ",0)"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSql
            Next
        End With
    End If
    
    If vsfListHerbal.TextMatrix(1, 1) <> "" And vsfListHerbal.TextMatrix(1, 2) <> "" Then
        With vsfListHerbal
            For i = 1 To .Rows - 1
                str˵�� = IIf(.TextMatrix(i, mCol.˵��) = "", "Null", "'" & .TextMatrix(i, mCol.˵��) & "'")
                gstrSql = "Zl_ҩƷ�ӳɷ���_Insert(" & vsfList.Rows - 1 + i & "," & .TextMatrix(i, mCol.��ͼ�) & "," & .TextMatrix(i, mCol.��߼�) & _
                             "," & .TextMatrix(i, mCol.�ӳ���) & "," & IIf(.TextMatrix(i, mCol.��۶�) = "", "Null", .TextMatrix(i, mCol.��۶�)) & "," & str˵�� & ",1)"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSql
            Next
        End With
    End If

    gcnOracle.BeginTrans: blnTrans = True          '��������
    For i = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False     '�ύ����
    
    MsgBox "����ɹ���", vbInformation, gstrSysName
    Unload Me
    
    Exit Sub
ErrHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Validate() As Boolean
    '����ʱ���ݼ���ڽ���һ�����ݼ��
    Dim intRow As Integer
    Dim intCol As Integer
    Dim j As Integer
    
     If vsfList.TextMatrix(1, 1) <> "" Or vsfList.TextMatrix(1, 2) <> "" Then
        With vsfList
            For intRow = 1 To .Rows - 1
                If Trim(.TextMatrix(intRow, mCol.��ͼ�)) = "" Then
                    MsgBox "��ҩ/��ҩ�����е�" & intRow & "����ͼ۲���Ϊ�գ�", vbInformation, gstrSysName
                    .SetFocus
                    .Row = intRow
                    .Col = mCol.��ͼ�
                    Validate = False
                    Exit Function
                ElseIf Trim(.TextMatrix(intRow, mCol.��߼�)) = "" Then
                    MsgBox "��ҩ/��ҩ�����е�" & intRow & "����߼۲���Ϊ�գ�", vbInformation, gstrSysName
                    .SetFocus
                    .Row = intRow
                    .Col = mCol.��߼�
                    Validate = False
                    Exit Function
                ElseIf Trim(.TextMatrix(intRow, mCol.�ӳ���)) = "" Then
    '                MsgBox "��ҩ/��ҩ�����е�" & intRow & "�мӳ��ʲ���Ϊ�գ�", vbInformation, gstrSysName
    '                .SetFocus
    '                .Row = intRow
    '                .Col = mCol.�ӳ���
    '                Validate = False
    '                Exit Function
                    .TextMatrix(intRow, mCol.�ӳ���) = 0
                End If
                If Val(.TextMatrix(intRow, mCol.��ͼ�)) > Val(.TextMatrix(intRow, mCol.��߼�)) Then
                    MsgBox "��ҩ/��ҩ�����е�" & intRow & "����ͼ۱���С�ڵ�����߼ۣ�", vbInformation, gstrSysName
                    .Col = mCol.��߼�
                    .Row = intRow
                    .SetFocus
                    Validate = False
                    Exit Function
                End If
                If Val(Trim(.TextMatrix(intRow, mCol.��ͼ�))) > 9999999999# Then
                    MsgBox "��ҩ/��ҩ�����е�" & intRow & "����ͼ۹���", vbInformation, gstrSysName
                    .Col = mCol.��ͼ�
                    .Row = intRow
                    .SetFocus
                    Validate = False
                    Exit Function
                End If
                If Val(Trim(.TextMatrix(intRow, mCol.��߼�))) > 9999999999# Then
                    MsgBox "��ҩ/��ҩ�����е�" & intRow & "����߼۹���", vbInformation, gstrSysName
                    .Col = mCol.��߼�
                    .Row = intRow
                    .SetFocus
                    Validate = False
                    Exit Function
                End If
                If Val(Trim(.TextMatrix(intRow, mCol.�ӳ���))) > 100 Then
                    MsgBox "��ҩ/��ҩ�����е�" & intRow & "�мӳ��ʹ���", vbInformation, gstrSysName
                    .Col = mCol.�ӳ���
                    .Row = intRow
                    .SetFocus
                    Validate = False
                    Exit Function
                End If
                If Val(Trim(.TextMatrix(intRow, mCol.��۶�))) > 9999999999# Then
                    MsgBox "��ҩ/��ҩ�����е�" & intRow & "�в�۶����", vbInformation, gstrSysName
                    .Col = mCol.��۶�
                    .Row = intRow
                    .SetFocus
                    Validate = False
                    Exit Function
                End If
                
                For j = intRow To .Rows - 1
                    If Val(.TextMatrix(intRow, mCol.��ͼ�)) + Val(.TextMatrix(intRow, mCol.��߼�)) > Val(.TextMatrix(j, mCol.��ͼ�)) + Val(.TextMatrix(j, mCol.��߼�)) Then
                        MsgBox "��ҩ/��ҩ�����е�" & intRow & "�����ݣ������˵�" & j & "�����ݣ����������룡", vbInformation, gstrSysName
                        .Row = j
                        .Col = mCol.��ͼ�
                        .SetFocus
                        Validate = False
                        Exit Function
                    ElseIf intRow < .Rows - 1 Then
                        If Val(.TextMatrix(intRow, mCol.��߼�)) > Val(.TextMatrix(intRow + 1, mCol.��ͼ�)) Then
                            MsgBox "��ҩ/��ҩ�����е�" & intRow & "����߼۴����˵�" & intRow + 1 & "����ͼۣ����������룡", vbInformation, gstrSysName
                            .Row = intRow + 1
                            .Col = mCol.��ͼ�
                            .SetFocus
                            Validate = False
                            Exit Function
                        End If
                    End If
                Next
            Next
        End With
    End If
    
    If vsfListHerbal.TextMatrix(1, 1) <> "" Or vsfListHerbal.TextMatrix(1, 2) <> "" Then
        With vsfListHerbal
            For intRow = 1 To .Rows - 1
                If Trim(.TextMatrix(intRow, mCol.��ͼ�)) = "" Then
                    MsgBox "��ҩ�����е�" & intRow & "����ͼ۲���Ϊ�գ�", vbInformation, gstrSysName
                    .SetFocus
                    .Row = intRow
                    .Col = mCol.��ͼ�
                    Validate = False
                    Exit Function
                ElseIf Trim(.TextMatrix(intRow, mCol.��߼�)) = "" Then
                    MsgBox "��ҩ�����е�" & intRow & "����߼۲���Ϊ�գ�", vbInformation, gstrSysName
                    .SetFocus
                    .Row = intRow
                    .Col = mCol.��߼�
                    Validate = False
                    Exit Function
                ElseIf Trim(.TextMatrix(intRow, mCol.�ӳ���)) = "" Then
    '                MsgBox "��ҩ�����е�" & intRow & "�мӳ��ʲ���Ϊ�գ�", vbInformation, gstrSysName
    '                .SetFocus
    '                .Row = intRow
    '                .Col = mCol.�ӳ���
    '                Validate = False
    '                Exit Function
                    .TextMatrix(intRow, mCol.�ӳ���) = 0
                End If
                If Val(.TextMatrix(intRow, mCol.��ͼ�)) > Val(.TextMatrix(intRow, mCol.��߼�)) Then
                    MsgBox "��ҩ�����е�" & intRow & "����ͼ۱���С�ڵ�����߼ۣ�", vbInformation, gstrSysName
                    .Col = mCol.��߼�
                    .Row = intRow
                    .SetFocus
                    Validate = False
                    Exit Function
                End If
                If Val(Trim(.TextMatrix(intRow, mCol.��ͼ�))) > 9999999999# Then
                    MsgBox "��ҩ�����е�" & intRow & "����ͼ۹���", vbInformation, gstrSysName
                    .Col = mCol.��ͼ�
                    .Row = intRow
                    .SetFocus
                    Validate = False
                    Exit Function
                End If
                If Val(Trim(.TextMatrix(intRow, mCol.��߼�))) > 9999999999# Then
                    MsgBox "��ҩ�����е�" & intRow & "����߼۹���", vbInformation, gstrSysName
                    .Col = mCol.��߼�
                    .Row = intRow
                    .SetFocus
                    Validate = False
                    Exit Function
                End If
                If Val(Trim(.TextMatrix(intRow, mCol.�ӳ���))) > 100 Then
                    MsgBox "��ҩ�����е�" & intRow & "�мӳ��ʹ���", vbInformation, gstrSysName
                    .Col = mCol.�ӳ���
                    .Row = intRow
                    .SetFocus
                    Validate = False
                    Exit Function
                End If
                If Val(Trim(.TextMatrix(intRow, mCol.��۶�))) > 9999999999# Then
                    MsgBox "��ҩ�����е�" & intRow & "�в�۶����", vbInformation, gstrSysName
                    .Col = mCol.��۶�
                    .Row = intRow
                    .SetFocus
                    Validate = False
                    Exit Function
                End If
                
                For j = intRow To .Rows - 1
                    If Val(.TextMatrix(intRow, mCol.��ͼ�)) + Val(.TextMatrix(intRow, mCol.��߼�)) > Val(.TextMatrix(j, mCol.��ͼ�)) + Val(.TextMatrix(j, mCol.��߼�)) Then
                        MsgBox "��ҩ�����е�" & intRow & "�����ݣ������˵�" & j & "�����ݣ����������룡", vbInformation, gstrSysName
                        .Row = j
                        .Col = mCol.��ͼ�
                        .SetFocus
                        Validate = False
                        Exit Function
                    ElseIf intRow < .Rows - 1 Then
                        If Val(.TextMatrix(intRow, mCol.��߼�)) > Val(.TextMatrix(intRow + 1, mCol.��ͼ�)) Then
                            MsgBox "��ҩ�����е�" & intRow & "����߼۴����˵�" & intRow + 1 & "����ͼۣ����������룡", vbInformation, gstrSysName
                            .Row = intRow + 1
                            .Col = mCol.��ͼ�
                            .SetFocus
                            Validate = False
                            Exit Function
                        End If
                    End If
                Next
            Next
        End With
    End If
    Validate = True
End Function

Private Sub Form_Activate()
    vsfList.SetFocus
End Sub

Private Sub Form_Load()
    Call InitVsf
    Call FillVSF
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrValue = ""
    mintҩƷ���� = 0
End Sub

Private Sub sstCustom_Click(PreviousTab As Integer)
    If sstCustom.Caption = "��ҩ/��ҩ" Then
        mintҩƷ���� = 0
    Else
        mintҩƷ���� = 1
    End If
End Sub

Private Sub vsfList_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    With vsfList
        If OldRow <= .Rows - 1 Then
            If Trim(.TextMatrix(OldRow, OldCol)) = "" And OldCol <> mCol.��۶� And OldCol <> mCol.˵�� Then
                If OldCol = mCol.�ӳ��� Then
                    .TextMatrix(OldRow, mCol.�ӳ���) = 0
                    Exit Sub
                End If
                MsgBox "�۸���Ϊ�գ�", vbInformation, gstrSysName
                Cancel = True
                Exit Sub
            End If
            If OldCol = mCol.��ͼ� Or OldCol = mCol.��߼� Or OldCol = mCol.�ӳ��� Then
                If Val(Trim(.TextMatrix(OldRow, OldCol))) > 9999999999# Then
                    MsgBox "�۸����", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
            End If
            
            If Val(Trim(.TextMatrix(OldRow, OldCol))) > 100 And OldCol = mCol.�ӳ��� Then
                MsgBox "�ӳ��ʹ���", vbInformation, gstrSysName
                Cancel = True
                Exit Sub
            End If
            
            If OldCol <> mCol.˵�� Then
                .TextMatrix(OldRow, OldCol) = formatex(.TextMatrix(OldRow, OldCol), 2)
            End If
        End If
    End With
End Sub

Private Sub vsfListHerbal_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    With vsfListHerbal
        If OldRow <= .Rows - 1 Then
            If Trim(.TextMatrix(OldRow, OldCol)) = "" And OldCol <> mCol.��۶� And OldCol <> mCol.˵�� Then
                If OldCol = mCol.�ӳ��� Then
                    .TextMatrix(OldRow, mCol.�ӳ���) = 0
                    Exit Sub
                End If
                MsgBox "�۸���Ϊ�գ�", vbInformation, gstrSysName
                Cancel = True
                Exit Sub
            End If
            If OldCol = mCol.��ͼ� Or OldCol = mCol.��߼� Or OldCol = mCol.�ӳ��� Then
                If Val(Trim(.TextMatrix(OldRow, OldCol))) > 9999999999# Then
                    MsgBox "�۸����", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
            End If
            
            If Val(Trim(.TextMatrix(OldRow, OldCol))) > 100 And OldCol = mCol.�ӳ��� Then
                MsgBox "�ӳ��ʹ���", vbInformation, gstrSysName
                Cancel = True
                Exit Sub
            End If
            
            If OldCol <> mCol.˵�� Then
                .TextMatrix(OldRow, OldCol) = formatex(.TextMatrix(OldRow, OldCol), 2)
            End If
        End If
    End With
End Sub

Private Sub vsfList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If KeyCode = vbKeyDelete Then
        With vsfList
            If .Row = 1 And .Rows = 2 Then
                For i = 1 To .Cols - 1
                    .TextMatrix(1, i) = ""
                Next
            Else
                .RemoveItem .Row
                For i = 1 To .Rows - 1
                    .TextMatrix(i, mCol.���) = i
                Next
            End If
            Exit Sub
        End With
    End If
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsfList
        
        If .Col <> mCol.intCount - 1 Then
            If CheckData(.Row, .Col, .TextMatrix(.Row, .Col)) = True Then
                .TextMatrix(.Row, .Col) = formatex(.TextMatrix(.Row, .Col), 2)
                .Col = .Col + 1
            Else
                KeyCode = 0
                .SetFocus
                .EditSelStart = 0
                .EditSelLength = Len(.TextMatrix(.Row, .Col))
            End If
        Else
            If Trim(.TextMatrix(.Rows - 1, mCol.��ͼ�)) = "" And Trim(.TextMatrix(.Rows - 1, mCol.��߼�)) = "" And Trim(.TextMatrix(.Rows - 1, mCol.�ӳ���)) = "" Then
                .Row = .Rows - 1
                .TextMatrix(.Row, mCol.��ͼ�) = .TextMatrix(.Row - 1, mCol.��߼�)
                .Col = mCol.��߼�
            Else
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .TextMatrix(.Row, mCol.���) = .Row
                .TextMatrix(.Row, mCol.��ͼ�) = .TextMatrix(.Row - 1, mCol.��߼�)
                .Col = mCol.��߼�
            End If
        End If
    End With
End Sub
Private Sub vsfListHerbal_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If KeyCode = vbKeyDelete Then
        With vsfListHerbal
            If .Row = 1 And .Rows = 2 Then
                For i = 1 To .Cols - 1
                    .TextMatrix(1, i) = ""
                Next
            Else
                .RemoveItem .Row
                For i = 1 To .Rows - 1
                    .TextMatrix(i, mCol.���) = i
                Next
            End If
            Exit Sub
        End With
    End If
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsfListHerbal
        
        If .Col <> mCol.intCount - 1 Then
            If CheckData(.Row, .Col, .TextMatrix(.Row, .Col)) = True Then
                .TextMatrix(.Row, .Col) = formatex(.TextMatrix(.Row, .Col), 2)
                .Col = .Col + 1
            Else
                KeyCode = 0
                .SetFocus
                .EditSelStart = 0
                .EditSelLength = Len(.TextMatrix(.Row, .Col))
            End If
        Else
            If Trim(.TextMatrix(.Rows - 1, mCol.��ͼ�)) = "" And Trim(.TextMatrix(.Rows - 1, mCol.��߼�)) = "" And Trim(.TextMatrix(.Rows - 1, mCol.�ӳ���)) = "" Then
                .Row = .Rows - 1
                .TextMatrix(.Row, mCol.��ͼ�) = .TextMatrix(.Row - 1, mCol.��߼�)
                .Col = mCol.��߼�
            Else
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .TextMatrix(.Row, mCol.���) = .Row
                .TextMatrix(.Row, mCol.��ͼ�) = .TextMatrix(.Row - 1, mCol.��߼�)
                .Col = mCol.��߼�
            End If
        End If
    End With
End Sub
Private Sub vsfList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     If ((Not InStr("0123456789.", Chr(KeyAscii)) > 0) Or (Chr(KeyAscii) = "." And InStr(vsfList.EditText, ".") > 0)) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn And Col <> mCol.˵�� Then
              KeyAscii = 0
      End If
    
'    If Not IsNumeric(Chr(KeyAscii)) Then
'        If KeyAscii <> vbKeyReturn And Chr(KeyAscii) = "." And InStr(vsfList.TextMatrix(Row, Col), ".") > 0 And KeyAscii <> vbKeyBack And Col <> mCol.˵�� Then
'            KeyAscii = 0
'        End If
'    End If
End Sub
Private Sub vsfListHerbal_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     If ((Not InStr("0123456789.", Chr(KeyAscii)) > 0) Or (Chr(KeyAscii) = "." And InStr(vsfListHerbal.EditText, ".") > 0)) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn And Col <> mCol.˵�� Then
              KeyAscii = 0
      End If
    
'    If Not IsNumeric(Chr(KeyAscii)) Then
'        If KeyAscii <> vbKeyReturn And Chr(KeyAscii) <> "." And KeyAscii <> vbKeyBack And Col <> mCol.˵�� Then
'            KeyAscii = 0
'        End If
'    End If
End Sub
Private Function CheckData(ByVal intRow As Integer, ByVal intCol As Integer, ByVal strText As String) As Boolean
    '���ܣ�������ƶ���ǰ�е�ʱ�򣬱���������Ƿ���ȷ
    '����ֵ�� true �����֣�false ��������
    Dim strTemp As String
    
    If mintҩƷ���� = 0 Then
        With vsfList
            strTemp = Trim(strText)
            If intCol <> mCol.˵�� And intCol <> mCol.��۶� Then
                If strTemp <> "" Then
                    If IsNumeric(strTemp) = False Then
                        MsgBox "��������ȷ���������ͣ�", vbInformation, gstrSysName
                        CheckData = False
                        Exit Function
                    Else
                        If intCol = mCol.��߼� And Val(strTemp) < Val(.TextMatrix(intRow, mCol.��ͼ�)) Then
                            MsgBox "��߼۲���С����ͼۣ�", vbInformation, gstrSysName
                            CheckData = False
                        ElseIf intCol = mCol.��ͼ� And intRow <> 1 And Val(strTemp) < Val(.TextMatrix(intRow - 1, mCol.��߼�)) Then
                            MsgBox "��ͼ�Ҫ���ڵ�����һ�е���߼ۣ�", vbInformation, gstrSysName
                            CheckData = False
                        Else
                            CheckData = True
                        End If
                        
                        If intCol <> mCol.�ӳ��� Then
                            If Val(strTemp) > 99999999999# Then
                                MsgBox "�ü۸����", vbInformation, gstrSysName
                                CheckData = False
                            End If
                        Else
                            If Val(strTemp) > 100 Then
                                MsgBox "��������ȷ�ļӳ��ʣ�", vbInformation, gstrSysName
                                CheckData = False
                            End If
                        End If
                        Exit Function
                    End If
                Else
                    If intCol = mCol.�ӳ��� Then
                        .TextMatrix(intRow, mCol.�ӳ���) = 0
                        CheckData = True
                        Exit Function
                    End If
                    
                    MsgBox "��Ԫ�����ݲ���Ϊ�գ�", vbInformation, gstrSysName
                    CheckData = False
                    Exit Function
                End If
            Else
                CheckData = True
            End If
        End With
    Else
        With vsfListHerbal
            strTemp = Trim(strText)
            If intCol <> mCol.˵�� And intCol <> mCol.��۶� Then
                If strTemp <> "" Then
                    If IsNumeric(strTemp) = False Then
                        MsgBox "��������ȷ���������ͣ�", vbInformation, gstrSysName
                        CheckData = False
                        Exit Function
                    Else
                        If intCol = mCol.��߼� And Val(strTemp) < Val(.TextMatrix(intRow, mCol.��ͼ�)) Then
                            MsgBox "��߼۲���С����ͼۣ�", vbInformation, gstrSysName
                            CheckData = False
                        ElseIf intCol = mCol.��ͼ� And intRow <> 1 And Val(strTemp) < Val(.TextMatrix(intRow - 1, mCol.��߼�)) Then
                            MsgBox "��ͼ�Ҫ���ڵ�����һ�е���߼ۣ�", vbInformation, gstrSysName
                            CheckData = False
                        Else
                            CheckData = True
                        End If
                        
                        If intCol <> mCol.�ӳ��� Then
                            If Val(strTemp) > 99999999999# Then
                                MsgBox "�ü۸����", vbInformation, gstrSysName
                                CheckData = False
                            End If
                        Else
                            If Val(strTemp) > 100 Then
                                MsgBox "��������ȷ�ļӳ��ʣ�", vbInformation, gstrSysName
                                CheckData = False
                            End If
                        End If
                        Exit Function
                    End If
                Else
                    If intCol = mCol.�ӳ��� Then
                        .TextMatrix(intRow, mCol.�ӳ���) = 0
                        CheckData = True
                        Exit Function
                    End If
                    
                    MsgBox "��Ԫ�����ݲ���Ϊ�գ�", vbInformation, gstrSysName
                    CheckData = False
                    Exit Function
                End If
            Else
                CheckData = True
            End If
        End With
    End If
End Function

Private Sub vsfList_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfList
        If Col = mCol.˵�� Then
            If LenB(StrConv(.EditText, vbFromUnicode)) > 50 Then
                Cancel = True
                .EditText = ""
                .TextMatrix(Row, Col) = .EditText
                MsgBox "˵���д���50���ַ������������룡", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End With
End Sub

Private Sub vsfListHerbal_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfListHerbal
        If Col = mCol.˵�� Then
            If LenB(StrConv(.EditText, vbFromUnicode)) > 50 Then
                Cancel = True
                .EditText = ""
                .TextMatrix(Row, Col) = .EditText
                MsgBox "˵���д���50���ַ������������룡", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End With
End Sub


