VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPathOutDefinition 
   Caption         =   "���˳����ǼǱ���"
   ClientHeight    =   8970
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12015
   Icon            =   "frmPathOutDefinition.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8970
   ScaleWidth      =   12015
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   12015
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   8295
      Width           =   12015
      Begin VB.CommandButton cmdDelete 
         Caption         =   "ɾ��(&D)"
         Height          =   350
         Left            =   6120
         TabIndex        =   10
         ToolTipText     =   "��""Delete""�����ɾ����ѡ���б��п�ɾ���б����ݣ�ѡ����������ɾ���������ݡ�"
         Top             =   200
         Width           =   1095
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   9480
         TabIndex        =   9
         Top             =   200
         Width           =   1100
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "����(&U)"
         Height          =   350
         Index           =   0
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   200
         Width           =   1100
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "����(&L)"
         Height          =   350
         Index           =   1
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   200
         Width           =   1100
      End
      Begin VB.CheckBox chk������д 
         BackColor       =   &H00F0F4E4&
         Caption         =   "�������˳�·��ʱ������д����"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   188
         Width           =   3015
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   10680
         TabIndex        =   4
         Top             =   200
         Width           =   1100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   12000
         Y1              =   30
         Y2              =   30
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   10000
         Y1              =   45
         Y2              =   45
      End
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   12015
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   240
         Picture         =   "frmPathOutDefinition.frx":6852
         Top             =   45
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   12000
         Y1              =   795
         Y2              =   795
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   10000
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmPathOutDefinition.frx":70DA
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   8175
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsItem 
      Height          =   7410
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   11895
      _cx             =   20981
      _cy             =   13070
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
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   8
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   280
      RowHeightMax    =   500
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPathOutDefinition.frx":7193
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
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
      WordWrap        =   -1  'True
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
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.PictureBox picAddRow 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00808080&
         Height          =   200
         Left            =   11480
         Picture         =   "frmPathOutDefinition.frx":73B9
         ScaleHeight     =   195
         ScaleWidth      =   240
         TabIndex        =   6
         Top             =   280
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmPathOutDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum Item_Type
    T0���� = 0
    T1�ַ� = 1
    T2���� = 2
    T3������ = 3
    T4��ѡ�� = 4
    T5��ѡ�� = 5
    T6���� = 6
End Enum

Private Enum CNAME
    col_��Ŀ��� = 0    '˳���
    col_��Ŀ���� = 1
    Col_ҳ�� = 2
    col_���� = 3
    col_���� = 4
    col_ͨ�� = 5
    col_�б� = 6
    col_���� = 7
    '������
    col_״̬ = 8    '0-������1-ԭʼ��2-�޸�
    COL_�к� = 9    '��ID
    col_��ѡ��� = 10
End Enum
Private Const color_Unmodify = &H8000000F
Private Const mstrComboList = "0-����|1-�ַ�|2-����|3-����|4-��ѡ��|5-��ѡ��|6-����"
Private mstrDelItem As String 'ɾ���˵���Ŀ��Ŵ�
Private mstrCaption As String   '������
Private mlng·��ID As Long
Private mintType As Integer  '0-סԺ��1-����

Public Function ShowMe(frmMain As Object, ByVal lng·��ID As Long, ByVal strCaption As String, Optional ByVal intType As Integer) As Boolean
    mlng·��ID = lng·��ID
    mstrCaption = strCaption
    mintType = intType
    Me.Show 1, frmMain
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Call vsItem_KeyDown(vbKeyDelete, 0)
End Sub

Private Sub cmdMove_Click(Index As Integer)
    With vsItem
        If Index = 0 And .Row > .FixedRows Then
            .RowPosition(.Row) = .Row - 1
            .Row = .Row - 1
        ElseIf Index = 1 And .Row < .Rows - 1 Then
            .RowPosition(.Row) = .Row + 1
            .Row = .Row + 1
        End If
        Call FuncNoASC
    End With
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, j As Long, arrtmp As Variant
    Dim blnOneNullRow As Boolean
    Dim lngMaxPage As Long
    Dim strPage As String
    Dim strMsg As String
    Dim MaxPage As Long
    
    With vsItem
        '���ɾ���������У���ֻ��һ�п���
        blnOneNullRow = (.Rows = .FixedRows + 1 And .TextMatrix(.FixedRows, col_��Ŀ����) = "")
        
        If Not blnOneNullRow Then
            '�������
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, col_��Ŀ����) = "" And i >= 1 Then
                    '��ѡ���в���Ϊ��
                    If Val(.TextMatrix(i - 1, col_����)) <> T5��ѡ�� Or Val(.TextMatrix(i, col_����)) <> T5��ѡ�� Then
                        Exit For
                    End If
                End If
                If Val(.TextMatrix(i, col_����)) = T5��ѡ�� And .TextMatrix(i, col_�б�) = "" Then
                    .Select i, col_�б�
                    MsgBox "��" & i & "�е��б�ֵΪ�գ���������Ŀ�б�ֵ��", vbInformation, gstrSysName
                    .SetFocus
                    Exit Sub
                End If
                'ҳ��
                strPage = strPage & "," & IIf(Val(.TextMatrix(i, Col_ҳ��)) = 0, 1, Val(.TextMatrix(i, Col_ҳ��)))
                If IIf(Val(.TextMatrix(i, Col_ҳ��)) = 0, 1, Val(.TextMatrix(i, Col_ҳ��))) > MaxPage Then
                    MaxPage = IIf(Val(.TextMatrix(i, Col_ҳ��)) = 0, 1, Val(.TextMatrix(i, Col_ҳ��)))
                End If
            Next
            strPage = Mid(strPage, 2)
            If i <= .Rows - 1 Then
                MsgBox "��" & i & "����Ŀ����Ϊ�գ���������Ŀ���ơ�", vbInformation, gstrSysName
                .Select i, col_��Ŀ����
                .SetFocus
                Exit Sub
            End If
            
            '���ѡ���б�
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, col_����)) = T4��ѡ�� Then
                    If Trim(.TextMatrix(i, col_�б�)) = "" Then
                        Exit For
                    ElseIf InStr(.TextMatrix(i, col_�б�), ",") = 0 And Mid(.TextMatrix(i, col_�б�), 1, 1) <> "[" Then
                        Exit For
                    Else
                        If ZLCommFun.ActualLen(.TextMatrix(i, col_�б�)) >= 100 Then
                            Exit For
                        End If
                        arrtmp = Split(.TextMatrix(i, col_�б�), ",")
                        For j = 0 To UBound(arrtmp)
                            If Trim(arrtmp(j)) = "" Then
                                Exit For
                            End If
                        Next
                        If j <= UBound(arrtmp) Then
                            Exit For
                        End If
                    End If
                End If
            Next
            If i <= .Rows - 1 Then
                MsgBox "��" & i & "��ѡ���б��ʽ������Ҫ��ÿ��ѡ���Ϊ�գ����ѡ�����Զ��ŷָ���", vbInformation, gstrSysName
                .Select i, col_�б�
                .SetFocus
                Exit Sub
            End If
            
        ElseIf chk������д.Value = 1 Then
            MsgBox "û�ж�����д��Ŀʱ����������Ϊ������д��", vbInformation, gstrSysName
            .Select .FixedRows, col_�б�
            .SetFocus
            Exit Sub
        End If
        '���ҳ���Ƿ�����
        If Not (.Rows = 2 And .TextMatrix(1, col_��Ŀ����) = "") Then
            If Not CheckPageNum(strPage, MaxPage, strMsg) Then
                MsgBox strMsg, vbInformation, gstrSysName
                .SetFocus
                Exit Sub
            End If
        End If
        
    End With
    
    
    If SaveData(blnOneNullRow) Then
        Unload Me
    End If
End Sub

Private Function CheckPageNum(ByVal strPage As String, ByVal MaxPage As Long, strMsg As String) As Boolean
'���ܣ��ж�ҳ���Ƿ�����
    Dim strSql As String, rsTmp As Recordset
    
    strSql = "Select Rownum From Dual Connect By Rownum < [1]+1 Minus Select Column_Value From Table(f_Num2list([2]))"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, MaxPage, strPage)
    If rsTmp.RecordCount > 0 Then
        strMsg = "ȱ�ٵ�" & rsTmp!Rownum & "ҳ�����ݣ�����ҳ���Ƿ�������"
    Else
       CheckPageNum = True
    End If
End Function

Private Function SaveData(blnOneNullRow As Boolean) As Boolean
'���ܣ���������
    Dim colSQL As New Collection, blnTrans As Boolean
    Dim strSql As String, strTmp As String
    Dim lng���� As Long, i As Long, arrtmp As Variant
    Dim lngMaxNO As Long, lngPage As Long, j As Long
    Dim blnGrant As Boolean
    
    If Not blnOneNullRow Then
        With vsItem
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, col_״̬)) = 0 Then
                    If lngMaxNO = 0 Then lngMaxNO = GetMaxItemNO
                    lngMaxNO = lngMaxNO + 1
                    
                    strSql = "0,2," & lngMaxNO
                ElseIf Val(.TextMatrix(i, col_״̬)) = 2 Then
                    strSql = "1,2," & .TextMatrix(i, COL_�к�)
                Else
                    strSql = ""
                End If
                If strSql <> "" Then
                    lng���� = Val("" & .TextMatrix(i, col_����))
                    If lng���� = T0���� Or lng���� = T1�ַ� Or lng���� = T2���� Then
                        strTmp = lng���� & "|" & IIf(.Cell(flexcpChecked, i, col_����) = 1, 1, 0)
                    ElseIf lng���� = T4��ѡ�� Then
                        strTmp = lng���� & "|" & .TextMatrix(i, col_�б�)
                        If Mid(.TextMatrix(i, col_�б�), 1, 1) = "[" And Mid(.TextMatrix(i, col_�б�), Len(.TextMatrix(i, col_�б�))) = "]" Then
                            blnGrant = True
                        End If
                    ElseIf lng���� = T5��ѡ�� Then
                        strTmp = lng���� & "|" & IIf(.Cell(flexcpChecked, i, col_����) = 1, 1, 0) & "|" & .TextMatrix(i, col_�б�)
                    Else
                        strTmp = lng����
                    End If
                    lngPage = Val(.TextMatrix(i, Col_ҳ��))
                    If lngPage = 0 Then
                        For j = i - 1 To 1 Step -1
                            If Val(.TextMatrix(j, Col_ҳ��)) <> 0 Then lngPage = Val(.TextMatrix(j, Col_ҳ��)): Exit For
                        Next
                    End If
                    strSql = "Zl_·��������_Update(" & strSql & "," & _
                            ZVal(Val(.TextMatrix(i, col_��Ŀ���))) & ",'" & .TextMatrix(i, col_��Ŀ����) & "','" & strTmp & "',Null," & ZVal(lngPage) & "," _
                            & IIf(.Cell(flexcpChecked, i, col_ͨ��) = 1, 1, 0) & "," & ZVal(mlng·��ID) & "," & ZVal(.TextMatrix(i, col_��ѡ���)) & "," & glngSys & ")"
                    colSQL.Add strSql, "C" & colSQL.count + 1
                End If
            Next
        End With
    End If
    
    If mstrDelItem <> "" Then
        arrtmp = Split(mstrDelItem, ",")
        For i = 0 To UBound(arrtmp)
            strSql = "Zl_·��������_Update(2,2," & arrtmp(i) & ",Null,Null,Null,Null,NULL,NULL,NULL,NULL," & glngSys & ")"
            colSQL.Add strSql, "C" & colSQL.count + 1
        Next
    End If
    
    On Error GoTo errH
    If colSQL.count > 0 Then
        gcnOracle.BeginTrans: blnTrans = True
            For i = 1 To colSQL.count
                Call zlDatabase.ExecuteProcedure(IIf(mintType = 1, Replace(colSQL("C" & i), "·��������", "����·��������"), colSQL("C" & i)), Me.Caption)
            Next
        gcnOracle.CommitTrans: blnTrans = False
    End If
    
    '�����������
    Call zlDatabase.SetPara("������д�����ǼǱ�", chk������д.Value, glngSys, IIf(mintType = 1, P����·��Ӧ��, P�ٴ�·��Ӧ��))
    
    SaveData = True
    If blnGrant Then
        MsgBox "��ѡ�����ֵ����Ϊ��ѡ�������Դ���뵽�������ж�""�ٴ�·��Ӧ��""ģ�����������Ȩ��", vbInformation, Me.Caption
    End If
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("|'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0 '����������ָ�����������
    End If
End Sub

Private Sub LoadData()
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long, arrtmp As Variant, arrType As Variant
 
    arrType = Split(mstrComboList, "|")
    
    strSql = "Select a.�к�, a.ҳ��, Nvl(b.���, a.��Ŀ���) As ��Ŀ���, a.��Ŀ�ı�1, a.��Ŀ�ı�2, a.·��id, a.��ѡ���" & vbNewLine & _
                "From ·������ṹ A, ·��������� B" & vbNewLine & _
                "Where a.����id = b.����id(+) And a.�к� = b.�к�(+) And a.����id = 2 And" & vbNewLine & _
                "      (Nvl(a.·��id, b.·��id) = [1] And (Exists (Select 1 From ·��������� Where ����id = 2 And ·��id = [1]) Or Not Exists (Select 1 From ·������ṹ Where ����id = 2 And a.·��id Is Null)))" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select a.�к�, a.ҳ��, a.��Ŀ���, a.��Ŀ�ı�1, a.��Ŀ�ı�2, a.·��id, a.��ѡ���" & vbNewLine & _
                "From ·������ṹ A" & vbNewLine & _
                "Where a.����id = 2 And a.·��id Is Null And Not Exists (Select 1 From ·��������� Where ����id = 2 And ·��id = [1])" & vbNewLine & _
                "Order By ��Ŀ���, ��ѡ���"


    On Error GoTo errH
    If mintType = 1 Then strSql = Replace(strSql, "·������ṹ", "����·������ṹ"): strSql = Replace(strSql, "·���������", "����·���������")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng·��ID)
    With vsItem
        .Rows = .FixedRows
        If rsTmp.RecordCount = 0 Then
            Call AddNewRow
        Else
            .Rows = .FixedRows + rsTmp.RecordCount
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, col_��Ŀ���) = rsTmp!��Ŀ��� & ""
                .TextMatrix(i, Col_ҳ��) = rsTmp!ҳ�� & ""
                .TextMatrix(i, col_��Ŀ����) = rsTmp!��Ŀ�ı�1 & ""
                .Cell(flexcpChecked, i, col_ͨ��) = IIf(rsTmp!·��ID & "" = "", 1, 0)
                
                arrtmp = Split(rsTmp!��Ŀ�ı�2, "|")    '=����|�Ƿ�����ѡ��1,ѡ��2,...
                .TextMatrix(i, col_����) = arrType(Val(arrtmp(0)))
                
                If Val(arrtmp(0)) = T3������ Or Val(arrtmp(0)) = T4��ѡ�� Then
                    .Cell(flexcpChecked, i, col_����) = 1
                    .Cell(flexcpBackColor, i, col_����) = color_Unmodify
                ElseIf Val(arrtmp(0)) = T6���� Then
                    .Cell(flexcpChecked, i, col_����) = 0
                    .Cell(flexcpBackColor, i, col_����) = color_Unmodify
                ElseIf Val(arrtmp(0)) = T5��ѡ�� Then
                    If Val(arrtmp(1)) = 1 Then
                        .Cell(flexcpChecked, i, col_����) = 1
                    End If
                    .TextMatrix(i, col_�б�) = arrtmp(2)
                ElseIf UBound(arrtmp) > 0 Then  '���֣��ַ�������
                    If Val(arrtmp(1)) = 1 Then
                        .Cell(flexcpChecked, i, col_����) = 1
                    End If
                End If
                
                If Val(arrtmp(0)) = T4��ѡ�� Then
                    If UBound(arrtmp) > 0 Then
                        .TextMatrix(i, col_�б�) = arrtmp(1)
                    End If
                End If
                
                .TextMatrix(i, col_״̬) = 1
                .TextMatrix(i, COL_�к�) = rsTmp!�к�
                
                rsTmp.MoveNext
            Next
        End If
    End With
    Call FuncNoASC
    vsItem.Row = 1
    chk������д.Value = zlDatabase.GetPara("������д�����ǼǱ�", glngSys, IIf(mintType = 1, P����·��Ӧ��, P�ٴ�·��Ӧ��), 0)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)

    mstrDelItem = ""
    Me.Caption = mstrCaption
    vsItem.ColComboList(col_����) = mstrComboList
    vsItem.ColComboList(Col_ҳ��) = "1|2|3|4|5|6|7|8|9|10"
    vsItem.ColDataType(col_����) = flexDTBoolean
    vsItem.Editable = flexEDKbdMouse
    picAddRow.Visible = False
    Call LoadData
End Sub

Private Sub Form_Resize()
    Dim lngWidth As Long, i As Long
    On Error Resume Next
    Line1(0).X2 = Me.Width
    Line1(1).X2 = Me.Width
    Line1(2).X2 = Me.Width
    Line1(3).X2 = Me.Width
    vsItem.Width = Me.Width - 320
    vsItem.Height = Me.Height - vsItem.Top - picBottom.Height - 590
    picAddRow.Left = vsItem.Left + vsItem.Width - vsItem.ColWidth(col_����) - 30
    For i = 0 To vsItem.Cols - 1
        If Not vsItem.ColHidden(i) And i <> col_�б� And i <> col_���� Then lngWidth = lngWidth + vsItem.ColWidth(i)
    Next
    vsItem.ColWidth(col_�б�) = vsItem.Width - lngWidth - 470
    picBottom.Top = Me.Height - vsItem.Height - vsItem.Top
    picBottom.Width = Me.Width
    cmdOK.Left = Me.Width - cmdOK.Width - 1800
    cmdCancel.Left = Me.Width - cmdCancel.Width - 500
    If vsItem.Width < 7888 Then
        picAddRow.Visible = False
    Else
        picAddRow.Visible = True
    End If
    If Me.Width < 9900 Then Me.Width = 9900
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub picAddRow_Click()
    Dim i As Long
    
    If vsItem.Row = vsItem.Rows - 1 Then
        Call AddNewRow
    Else
        Call AddNewRow(vsItem.Row)
    End If
    
End Sub

Private Sub vsItem_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngType As Long, i As Long
    '���ݵ�ǰ���������ù�����
    With vsItem
        If Col = col_���� Then
            lngType = Val(.TextMatrix(Row, col_����))
            
            If lngType = T3������ Or lngType = T4��ѡ�� Then
                .Cell(flexcpChecked, Row, col_����) = 1
                .Cell(flexcpBackColor, Row, col_����) = color_Unmodify
            ElseIf lngType = T6���� Then
                .Cell(flexcpChecked, Row, col_����) = 0
                .Cell(flexcpBackColor, Row, col_����) = color_Unmodify
            Else
                .Cell(flexcpBackColor, Row, col_����) = vbWhite
            End If
            
            '����б�ֵ
            If lngType <> T4��ѡ�� And lngType <> T5��ѡ�� Then
                If .TextMatrix(Row, col_�б�) <> "" Then
                    .TextMatrix(Row, col_�б�) = ""
                End If
            End If
            '��ѡ�����Ŀ���
            Call FuncNoASC
        ElseIf Col = col_��Ŀ���� Then
            .TextMatrix(Row, col_��Ŀ����) = .TextMatrix(Row, col_��Ŀ����)
            '��ѡ�����Ŀ���
            Call FuncNoASC
        ElseIf Col = Col_ҳ�� Or Col = col_���� Or Col = col_ͨ�� Then
            For i = Row + 1 To .Rows - 1
                If (.TextMatrix(i, col_��Ŀ����) = "" Or .TextMatrix(i, col_��Ŀ����) = .TextMatrix(Row, col_��Ŀ����)) And Val(.TextMatrix(i, col_����)) = T5��ѡ�� And Val(.TextMatrix(Row, col_����)) = T5��ѡ�� Then
                    .TextMatrix(i, Col) = .TextMatrix(Row, Col)
                    If Val(.TextMatrix(i, col_״̬)) = 1 Then 'Ϊ�˼򻯴���ֻҪ�޸��˶����Ϊ�Ѹģ�����ʵ���Ƿ�ı�
                        .TextMatrix(i, col_״̬) = 2
                    End If
                Else
                    Exit For
                End If
            Next
            If (Col = col_���� Or Col = col_ͨ��) And .Cell(flexcpData, Row, col_��Ŀ���) <> "" Then
                For i = Row - 1 To 1 Step -1
                    If (.TextMatrix(i, col_��Ŀ����) = "" Or .TextMatrix(i, col_��Ŀ����) = .TextMatrix(Row, col_��Ŀ����)) And Val(.TextMatrix(i, col_����)) = T5��ѡ�� And Val(.TextMatrix(Row, col_����)) = T5��ѡ�� Then
                        .TextMatrix(i, Col) = .TextMatrix(Row, Col)
                        If Val(.TextMatrix(i, col_״̬)) = 1 Then 'Ϊ�˼򻯴���ֻҪ�޸��˶����Ϊ�Ѹģ�����ʵ���Ƿ�ı�
                            .TextMatrix(i, col_״̬) = 2
                        End If
                    Else
                        If Val(.TextMatrix(i, col_����)) = T5��ѡ�� Then
                            .TextMatrix(i, Col) = .TextMatrix(Row, Col)
                            If Val(.TextMatrix(i, col_״̬)) = 1 Then 'Ϊ�˼򻯴���ֻҪ�޸��˶����Ϊ�Ѹģ�����ʵ���Ƿ�ı�
                                .TextMatrix(i, col_״̬) = 2
                            End If
                        End If
                        Exit For
                    End If
                Next
            End If
        ElseIf Col = col_�б� Then
            If Val(.TextMatrix(Row, col_����)) = T4��ѡ�� Then .ColComboList(col_�б�) = "..."
        End If
        
        If Val(.TextMatrix(Row, col_״̬)) = 1 Then 'Ϊ�˼򻯴���ֻҪ�޸��˶����Ϊ�Ѹģ�����ʵ���Ƿ�ı�
            .TextMatrix(Row, col_״̬) = 2
        End If
    End With
End Sub

Private Sub vsItem_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow And NewRow > vsItem.FixedRows - 1 Then
        If Me.Visible Then
            If picAddRow.Visible = False Then picAddRow.Visible = True
        End If
        picAddRow.Top = vsItem.Cell(flexcpTop, NewRow, col_����) + 30
        picAddRow.Left = vsItem.Left + vsItem.Cell(flexcpLeft, NewRow, col_����) + 50
    End If
    If NewCol = col_�б� And Val(vsItem.TextMatrix(NewRow, col_����)) = T4��ѡ�� Then
        vsItem.ColComboList(col_�б�) = "..."
    Else
        vsItem.ColComboList(col_�б�) = ""
    End If
End Sub

Private Sub vsItem_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = col_���� Then
        If vsItem.Cell(flexcpBackColor, Row, col_����) = color_Unmodify Then
            Cancel = True
        End If
    ElseIf Col = col_�б� Then
        If Val(vsItem.TextMatrix(Row, col_����)) <> T4��ѡ�� And Val(vsItem.TextMatrix(Row, col_����)) <> T5��ѡ�� Then
            Cancel = True
        End If
    ElseIf Col = Col_ҳ�� Then
        If vsItem.Cell(flexcpData, Row, col_��Ŀ���) <> "" Then
            Cancel = True
        End If
    ElseIf Col = col_���� Then
        Cancel = True
    End If
End Sub

Private Sub vsItem_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As Recordset, strSql As String
    Dim vPoint As POINTAPI, blnCancel As Boolean
    
    With vsItem
        If Col = col_�б� And Val(.TextMatrix(Row, col_����)) = T4��ѡ�� Then
            '����Ƿ���ڳ�����¼
            If Val(.TextMatrix(Row, col_״̬)) <> 0 Then
                If CheckItemNO(Val(.TextMatrix(Row, COL_�к�))) Then
                    If MsgBox("�ڡ����˳�����¼�����Ѵ��ڵ�ǰ�е�������ݣ��޸ĺ���ܻᶪʧ���ݣ���ȷ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Sub
                    End If
                End If
            End If
            '����ж����򲻲����ֵ��
            strSql = "Select Rownum As ID, ϵͳ, ���� From zlBaseCode"
            vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "�����ֵ��", _
                False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True)
            '�ж��Ƿ�������
            If Not rsTmp Is Nothing Then
                .TextMatrix(Row, Col) = "[" & rsTmp!���� & "]"
            End If
        End If
    End With
End Sub

Private Sub vsItem_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, lngRowNO As Long
    'ɾ����
    If KeyCode = vbKeyDelete Then
        With vsItem
            If .Col = col_�б� Then
                .TextMatrix(.Row, .Col) = ""
            Else
                If .Row = .FixedRows - 1 Then .Row = .FixedRows
                If .TextMatrix(.Row, col_��Ŀ����) = "" And .Rows = 2 Then
                    MsgBox "û�п�ɾ������Ŀ�ˡ�", vbInformation, gstrSysName
                    Exit Sub
                End If
                If MsgBox("��ȷ��Ҫɾ����ǰ����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    lngRowNO = Val(.TextMatrix(.Row, COL_�к�))

                    If CheckItemNO(lngRowNO) Then
                        If MsgBox("��" & IIf(mintType = 1, "���������������¼��", "�����˳�����¼��") & "���Ѵ��ڵ�ǰ�е�������ݣ�ɾ�����ƻ�����֮��Ĺ�������ȷ��Ҫɾ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Exit Sub
                        End If
                    End If
                    
                    If lngRowNO <> 0 Then
                        If mstrDelItem = "" Then
                            mstrDelItem = lngRowNO
                        Else
                            mstrDelItem = mstrDelItem & "," & lngRowNO
                        End If
                    End If
                    For i = .Row + 1 To .Rows - 1
                        .TextMatrix(i, col_��Ŀ���) = Val(.TextMatrix(i, col_��Ŀ���)) - 1
                    Next
                    .RemoveItem .Row
                    If .Rows = .FixedRows Then Call AddNewRow
                End If
            End If
        End With
    ElseIf KeyCode > 127 Then
        '���ֱ�����뺺�ֵ�����
        Call vsItem_KeyPress(KeyCode)
    End If
End Sub

Private Sub vsItem_KeyPress(KeyAscii As Integer)
    '������(���һ�а��س�)
    Dim i As Long
    With vsItem
        If KeyAscii = vbKeyReturn Then
            If .Row = .Rows - 1 And .Col = col_�б� Then
                Call AddNewRow
                .Select .Rows - 1, col_��Ŀ����
            ElseIf .Col = col_�б� Then
                KeyAscii = 0
                .Select .Row + 1, col_��Ŀ����
            Else
                KeyAscii = 0
                .Col = .Col + 1
            End If
        Else
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsItem_CellButtonClick(.Row, .Col)
            Else
                .ColComboList(col_�б�) = "" 'ʹ��ť״̬��������״̬
            End If
            
        End If
    End With
End Sub

Private Function FuncNoASC() As Long
'���ܣ�Ϊ����������,����һ��ı����ҳ���޸�Ϊ��ͬ
    Dim i As Long, lngNo As Long, lng��� As Long
    
    With vsItem
        lngNo = 0
        .TextMatrix(1, col_��Ŀ���) = "1"
        If Val(.TextMatrix(1, col_״̬)) = 1 Then
            .TextMatrix(1, col_״̬) = 2
        End If
        For i = 2 To .Rows - 1
            If Val(.TextMatrix(i, col_����)) = T5��ѡ�� Then
                If Val(.TextMatrix(i - 1, col_����)) = T5��ѡ�� And (.TextMatrix(i, col_��Ŀ����) = .TextMatrix(i - 1, col_��Ŀ����) Or .TextMatrix(i, col_��Ŀ����) = "") Then
                    .TextMatrix(i, col_��Ŀ���) = .TextMatrix(i - 1, col_��Ŀ���)
                    .Cell(flexcpData, i, col_��Ŀ���) = "1"
                    .TextMatrix(i, Col_ҳ��) = .TextMatrix(i - 1, Col_ҳ��)
                    .Cell(flexcpChecked, i, col_����) = .Cell(flexcpChecked, i - 1, col_����)
                    .TextMatrix(i, col_��ѡ���) = Val(.TextMatrix(i - 1, col_��ѡ���)) + 1
                    lngNo = lngNo + 1
                Else
                    .TextMatrix(i, col_��Ŀ���) = i - lngNo
                    .Cell(flexcpData, i, col_��Ŀ���) = ""
                    lng��� = 1
                    If Val(.TextMatrix(i, col_����)) = T5��ѡ�� Then .TextMatrix(i, col_��ѡ���) = lng���
                End If
            Else
                .TextMatrix(i, col_��Ŀ���) = i - lngNo
                .Cell(flexcpData, i, col_��Ŀ���) = ""
                lng��� = 1
                If Val(.TextMatrix(i, col_����)) = T5��ѡ�� Then .TextMatrix(i, col_��ѡ���) = lng���
            End If
            If Val(.TextMatrix(i, col_״̬)) = 1 Then 'Ϊ�˼򻯴���ֻҪ�޸��˶����Ϊ�Ѹģ�����ʵ���Ƿ�ı�
                .TextMatrix(i, col_״̬) = 2
            End If
        Next
    End With
End Function


Private Sub AddNewRow(Optional ByVal lngRow As Long)
'���ܣ�����һ�հ���
'������lngRow-0�����һ������������Ϊ����
    With vsItem
        If lngRow = 0 Then
            .Rows = .Rows + 1
            lngRow = .Rows - 1
        Else
            vsItem.AddItem "", lngRow
        End If
        If lngRow <> 1 Then
            .TextMatrix(lngRow, col_����) = Split(mstrComboList, "|")(Val(.TextMatrix(lngRow - 1, col_����)))
            .TextMatrix(lngRow, Col_ҳ��) = .TextMatrix(lngRow - 1, Col_ҳ��)
            .Cell(flexcpChecked, lngRow, col_����) = .Cell(flexcpChecked, lngRow - 1, col_����)
            .Cell(flexcpChecked, lngRow, col_ͨ��) = .Cell(flexcpChecked, lngRow - 1, col_ͨ��)
        Else
            .TextMatrix(lngRow, col_����) = Split(mstrComboList, "|")(0)
        End If
        Call vsItem_AfterEdit(lngRow, col_����)
        .TextMatrix(lngRow, col_״̬) = 0
        Call FuncNoASC
    End With
End Sub

Private Function GetMaxItemNO() As Long
'���ܣ���ȡ��Ŀ�б�ĵ�ǰ����к�
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select Nvl(Max(�к�),0) as ����к� From ·������ṹ Where ����ID = 2"
    On Error GoTo errH
    If mintType = 1 Then strSql = Replace(strSql, "·������ṹ", "����·������ṹ"): strSql = Replace(strSql, "·���������", "����·���������")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        GetMaxItemNO = rsTmp!����к�
    Else
        GetMaxItemNO = 0
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckItemNO(ByVal lngRowNO As Long) As Boolean
'���ܣ���鵱ǰ���Ƿ��Ѵ��ڲ����������
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    If mintType = 1 Then
        strSql = "Select 1 From �������������¼ Where �к� = [1] And Rownum=1"
    Else
        strSql = "Select 1 From ���˳�����¼ Where �к� = [1] And Rownum=1"
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngRowNO)
    CheckItemNO = rsTmp.RecordCount > 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsItem_LostFocus()
    vsItem.ForeColorSel = vbBlack
    vsItem.BackColorSel = vbWhite
End Sub

Private Sub vsItem_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As Recordset, strSql As String
    Dim vPoint As POINTAPI, blnCancel As Boolean
    
    With vsItem
        If Col = col_�б� And Val(.TextMatrix(Row, col_����)) = T4��ѡ�� Or Col = col_���� And .EditText <> .TextMatrix(Row, col_����) Or Col = col_ͨ�� And Val(.TextMatrix(Row, col_ͨ��)) <> 0 Then
            '����Ƿ���ڳ�����¼
            If Val(.TextMatrix(Row, col_״̬)) <> 0 Then
                If CheckItemNO(Val(.TextMatrix(Row, COL_�к�))) Then
                    If MsgBox("��" & IIf(mintType = 1, "���������������¼��", "�����˳�����¼��") & "���Ѵ��ڵ�ǰ�е�������ݣ��޸ĺ���ܻᶪʧ���ݣ���ȷ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Cancel = True
                        If Col = col_���� Then .EditText = .TextMatrix(Row, col_����)
                        Exit Sub
                    End If
                End If
            End If
        End If
        If Col = col_�б� And Val(.TextMatrix(Row, col_����)) = T4��ѡ�� Then
            '����ж����򲻲����ֵ��
            If .EditText = "" Then Exit Sub
            If InStr(.EditText, ",") = 0 Then
                strSql = "Select Rownum As ID, ϵͳ, ���� From zlBaseCode Where ���� Like [1]"
                vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
                On Error GoTo errH
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "�����ֵ��", _
                    False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    gstrLike & .EditText & "%")
                If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                    Cancel = True
                Else
                    '�ж��Ƿ�������
                    If Not rsTmp Is Nothing Then
                        .EditText = "[" & rsTmp!���� & "]"
                    End If
                End If
            End If
            If Mid(.EditText, 1, 1) = "[" And Mid(.EditText, Len(.EditText)) = "]" Then
                strSql = "Select Count(1) as ���� From zlBaseCode Where ����=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Mid(.EditText, 2, Len(.EditText) - 2))
                If Val(rsTmp!���� & "") = 0 Then
                    MsgBox "û������ҵ�����ֵ��" & Mid(.EditText, 2, Len(.EditText) - 2)
                    Cancel = True
                    Exit Sub
                End If
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
