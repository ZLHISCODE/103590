VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmPathOutLog 
   Caption         =   "���˳����Ǽ�"
   ClientHeight    =   8970
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12015
   Icon            =   "frmPathOutLog.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8970
   ScaleWidth      =   12015
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   0
      ScaleHeight     =   7455
      ScaleWidth      =   12000
      TabIndex        =   6
      Top             =   840
      Width           =   12000
      Begin VSFlex8Ctl.VSFlexGrid vsItem 
         Height          =   7410
         Left            =   0
         TabIndex        =   7
         Top             =   0
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
         Rows            =   7
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   280
         RowHeightMax    =   500
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPathOutLog.frx":6852
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   2
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
         OwnerDraw       =   1
         Editable        =   2
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
      End
   End
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
      Begin VB.CommandButton cmdPrintToEXCEL 
         Caption         =   "�����EXCEL"
         Height          =   350
         Left            =   7800
         TabIndex        =   8
         Top             =   200
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   10800
         TabIndex        =   4
         Top             =   200
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   9600
         TabIndex        =   3
         Top             =   200
         Width           =   1100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   11880
         Y1              =   45
         Y2              =   45
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   12000
         Y1              =   30
         Y2              =   30
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
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "����д����Ҫ��Ǽǵ���Ϣ��ǳ��ɫ�����ĵ�Ԫ��Ϊ����������밴YYYY-MM-DD��ʽ¼�롣�ύ�������󽫲��������޸ġ�"
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   7455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   10000
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   12000
         Y1              =   795
         Y2              =   795
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   195
         Picture         =   "frmPathOutLog.frx":6990
         Top             =   45
         Width           =   720
      End
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   5055
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   8895
      _Version        =   589884
      _ExtentX        =   15690
      _ExtentY        =   8916
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmPathOutLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngFun As Long '0-������1-�鿴��2-�޸�
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mlng·��ID As Long
Private mlng����·��ID As Long
Private mcolSQL As New Collection

Private mblnOK As Boolean
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
    col_��Ŀֵ = 2
    col_��ע = 3
    
    '������
    col_���� = 4
    col_���� = 5
    col_״̬ = 6    '1-ԭʼ��2-�޸�
    col_�к� = 7    '��ID
    Col_ҳ�� = 8
End Enum
Private Const Color_MustAddBack = &HE1FFE1

Public Function ShowMe(frmMain As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lngFun As Long, ByRef colSQL As Collection, _
                    ByVal lng·��ID As Long, Optional ByVal lng����·��Id As Long) As Boolean
'������ lngFun=0-������1-�鿴��2-�޸�
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mlngFun = lngFun
    mlng·��ID = lng·��ID
    mlng����·��ID = lng����·��Id
    
    If mlngFun = 1 Then
        If CheckPatiPathOutLog(lng����ID, lng��ҳID) = False Then
            MsgBox "δ�ǼǸò��˵ĳ�����Ϣ��", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    Set mcolSQL = Nothing
    
    Me.Show 1, frmMain
    
    Set colSQL = mcolSQL
    
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, strMsg As String, intType As Integer
    Dim strValue As String
    Dim blnIsCheck As Boolean   '�ж϶�ѡ��������
    Dim j As Long
    
    With vsItem
        For i = .FixedRows To .Rows - 1
            intType = Val(.TextMatrix(i, col_����))
            strValue = Trim(.TextMatrix(i, col_��Ŀֵ))
            If intType = T0���� Then
                If strValue <> "" And Not IsNumeric(strValue) Then
                    strMsg = "Ҫ����������ݱ����������͡�"
                    Exit For
                End If
            ElseIf intType = T2���� Then
                If strValue <> "" And Not IsDate(strValue) Then
                    strMsg = "Ҫ����������ݱ����������͡�"
                    Exit For
                End If
            ElseIf intType = T1�ַ� Then
                If strValue <> "" Then
                    If zlCommFun.ActualLen(strValue) >= 100 Then
                        strMsg = "��Ŀֵ�����������100���ַ���50�����֡�"
                        Exit For
                    End If
                End If
            End If
            If intType = T0���� Or intType = T1�ַ� Or intType = T2���� Or intType = T4��ѡ�� Then
                If Val(.TextMatrix(i, col_����)) = 1 Then
                    If strValue = "" Then
                        strMsg = "Ҫ�������д���ݡ�"
                        Exit For
                    End If
                End If
            End If
            
            If intType = T5��ѡ�� And Val(.TextMatrix(i, col_����)) = 1 Then
                For j = i To .Rows - 1
                    If .TextMatrix(j, col_��Ŀ���) <> .TextMatrix(i, col_��Ŀ���) Then Exit For
                    If .Cell(flexcpChecked, j, col_��Ŀֵ) = 1 Then blnIsCheck = True
                Next
                If Not blnIsCheck Then
                    strMsg = "Ҫ�������д���ݡ�"
                    Exit For
                End If
                blnIsCheck = False
                i = j - 1
            End If
            
            If .TextMatrix(i, col_��ע) <> "" Then
                If zlCommFun.ActualLen(.TextMatrix(i, col_��ע)) >= 1000 Then
                    strMsg = "��ע���������������1000���ַ���500�����֡�"
                    Exit For
                End If
            End If
        Next
        If i <= .Rows - 1 Then
            tbcSub.Item(Val(.TextMatrix(i, Col_ҳ��)) - 1).Selected = True
            MsgBox "��" & .TextMatrix(i, col_��Ŀ���) & "����Ŀ��" & strMsg, vbInformation, gstrSysName
            .Select i, col_��Ŀֵ
            .SetFocus
            Exit Sub
        End If
        
        If SaveData = False Then
            Exit Sub
        End If
    End With
    
    mblnOK = True
    Unload Me
End Sub

Private Function SaveData() As Boolean
    Dim colSQL As New Collection, blnTrans As Boolean
    Dim strSql As String, i As Long, intType As Integer
    Dim strDate As String, str����ֵ As String, str�ַ�ֵ As String, str����ֵ As String, strValue As String

    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-DD HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"

    With vsItem
        For i = .FixedRows To .Rows - 1
            str����ֵ = "Null"
            str�ַ�ֵ = "Null"
            str����ֵ = "Null"
            strSql = ""

            strValue = Trim(.TextMatrix(i, col_��Ŀֵ))
            intType = Val(.TextMatrix(i, col_����))
            Select Case intType
            Case T0����
                str����ֵ = strValue
            Case T3������, T5��ѡ��
                str����ֵ = IIf(.Cell(flexcpChecked, i, col_��Ŀֵ) = 1, 1, 0)
            Case T1�ַ�, T4��ѡ��
                str�ַ�ֵ = "'" & strValue & "'"
            Case T2����
                If strValue <> "" Then
                    str����ֵ = "To_Date('" & Format(strValue, "yyyy-MM-DD") & "','yyyy-mm-dd')"
                End If
            End Select

            If mlngFun = 0 Then
                '����(δ��д���в�����)
                If intType = T3������ Or strValue <> "" Or Trim(.TextMatrix(i, col_��ע)) <> "" Then
                    strSql = "0"
                End If
            Else
                '�޸�
                If Val(.TextMatrix(i, col_״̬)) = 2 Then
                    strSql = "1"
                End If
            End If

            If strSql <> "" Then
                strSql = "Zl_���˳�����¼_Update(" & strSql & "," & mlng����ID & "," & mlng��ҳID & "," & .TextMatrix(i, col_�к�) & _
                         "," & str����ֵ & "," & str�ַ�ֵ & "," & str����ֵ & ",'" & Trim(.TextMatrix(i, col_��ע)) & "','" & UserInfo.���� & "'," & strDate & "," & mlng����·��ID & ")"
                colSQL.Add strSql, "C" & colSQL.count + 1
            End If
        Next
    End With

    Set mcolSQL = colSQL

    SaveData = True
End Function

Private Sub cmdPrintToEXCEL_Click()
'����:�����EXCEL
    Dim objReport As ReportControl
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    Dim rsTmp As Recordset, strSql As String
    
    On Error GoTo errH
    
    strSql = "Select NVL(B.����,A.����) ���� From ������Ϣ A,������ҳ B Where A.����ID=B.����ID And A.����ID=[1] And B.��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mlng��ҳID)
    vsItem.ColWidth(col_��ע) = 3950
    Set objPrint.Body = Me.vsItem
    vsItem.AutoSize vsItem.FixedCols, vsItem.Cols - 1, , 45 '�߶�����Ӧ
    objPrint.Title.Text = "�����ǼǱ�"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("���ˣ�" & rsTmp!����)
    strSql = "Select ���� From �ٴ�·��Ŀ¼ Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng·��ID)
    Call objAppRow.Add("·����" & rsTmp!����)
    Call objPrint.UnderAppRows.Add(objAppRow)
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("��ӡ��:" & UserInfo.����)
    Call objAppRow.Add("��ӡʱ��:" & Format(Now, "yyyy-MM-dd HH:mm"))
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    zlPrintOrView1Grd objPrint, 3
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("|'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0 '����������ָ�����������
    End If
End Sub

Private Sub Form_Load()
    Dim lngPage As Long, i As Long
    
    Call RestoreWinState(Me, App.ProductName, mlngFun)
    
    mblnOK = False
    For i = 0 To vsItem.Cols - 1
        If vsItem.ColHidden(i) Then vsItem.ColWidth(i) = 0
    Next
    If mlngFun = 1 Then
        vsItem.Editable = flexEDNone
        cmdOK.Visible = False
        cmdCancel.Caption = "�˳�(&X)"
    End If
    
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
            .ColorSet.HeaderFaceLight = &HF0F4E4
            .ColorSet.HeaderFaceDark = &HF0F4E4
        End With
        '���ֻ��һҳ������ѡ�ͷ
        lngPage = CheckPage
        If lngPage > 0 Then
            For i = 0 To lngPage - 1
                .InsertItem(i, "��" & i + 1 & "ҳ", picItem.Hwnd, 0).Tag = i + 1
            Next
            .Item(0).Selected = True
            Call tbcSub_SelectedChanged(.Item(0))
        End If
        If lngPage <= 1 Then
            .PaintManager.HeaderMargin.Top = -20   '����ʾTab��������
        End If
    End With
    
    '�󶨺��ڼ��أ�����������ʾ����������
    Call LoadData
End Sub

Private Function CheckPage() As Long
'���أ�ҳ��
    Dim strSql As String
    Dim rsTmp As Recordset
    
    strSql = "Select Count(Distinct NVL(ҳ��,1)) as ҳ�� From ·������ṹ Where ����id = 2 And ·��id Is Null Or ·��id = [1]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng·��ID)
    CheckPage = Val(rsTmp!ҳ�� & "")
    
End Function

Private Sub LoadData()
    Dim i As Long, arrtmp As Variant, intType As Integer
    Dim strSql As String
    Dim rsTmp As Recordset
    
    On Error GoTo errH
    If mlngFun = 1 Or mlngFun = 2 Then
        strSql = "Select a.�к�, NVL(a.ҳ��,1)  as ҳ��, a.��Ŀ���, a.��Ŀ�ı�1, a.��Ŀ�ı�2, a.·��id, a.��ѡ���, b.����ֵ, b.�ַ�ֵ, b.����ֵ, b.��ע" & vbNewLine & _
                "From (Select a.�к�, a.ҳ��, Nvl(b.���, a.��Ŀ���) As ��Ŀ���, a.��Ŀ�ı�1, a.��Ŀ�ı�2, a.·��id, a.��ѡ���" & vbNewLine & _
                "       From ·������ṹ A, ·��������� B" & vbNewLine & _
                "       Where a.����id = b.����id(+) And a.�к� = b.�к�(+) And a.����id = 2 And" & vbNewLine & _
                "             (Nvl(a.·��id, b.·��id) = [3] And (Exists (Select 1 From ·��������� Where ����ID = 2  And ·��id = [3]) Or Not Exists (Select 1 From ·������ṹ Where ����id = 2 And a.·��id Is Null)))" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select a.�к�, NVL(a.ҳ��,1)  as ҳ��, a.��Ŀ���, a.��Ŀ�ı�1, a.��Ŀ�ı�2, a.·��id, a.��ѡ���" & vbNewLine & _
                "From ·������ṹ A" & vbNewLine & _
                "Where a.����id = 2 And a.·��id Is Null And Not Exists (Select 1 From ·��������� Where ����id = 2 And ·��id = [3])) A, ���˳�����¼ B" & vbNewLine & _
                "Where a.�к� = b.�к�(+) And b.����id(+) = [1] And b.��ҳid(+) = [2] And B.·����¼ID(+)=[4]" & vbNewLine & _
                "Order By ��Ŀ���, ��ѡ���"
 
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mlng��ҳID, mlng·��ID, mlng����·��ID)
    Else
        strSql = "Select a.�к�, NVL(a.ҳ��,1) as ҳ��, Nvl(b.���, a.��Ŀ���) As ��Ŀ���, a.��Ŀ�ı�1, a.��Ŀ�ı�2, a.·��id, a.��ѡ���" & vbNewLine & _
                "From ·������ṹ A, ·��������� B" & vbNewLine & _
                "Where a.����id = b.����id(+) And a.�к� = b.�к�(+) And a.����id = 2 And" & vbNewLine & _
                "      (Nvl(a.·��id, b.·��id) = [1] And (Exists (Select 1 From ·��������� Where ����id = 2 And ·��id = [1]) Or Not Exists (Select 1 From ·������ṹ Where ����id = 2 And a.·��id Is Null)))" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select a.�к�, NVL(a.ҳ��,1)  as ҳ��, a.��Ŀ���, a.��Ŀ�ı�1, a.��Ŀ�ı�2, a.·��id, a.��ѡ���" & vbNewLine & _
                "From ·������ṹ A" & vbNewLine & _
                "Where a.����id = 2 And a.·��id Is Null And Not Exists (Select 1 From ·��������� Where ����id = 2 And ·��id = [1])" & vbNewLine & _
                "Order By ��Ŀ���, ��ѡ���"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng·��ID)
    End If
    With vsItem
        .Rows = .FixedRows
        .Rows = .FixedRows + rsTmp.RecordCount
        For i = 1 To rsTmp.RecordCount
            .TextMatrix(i, col_��Ŀ���) = rsTmp!��Ŀ��� & ""
            .TextMatrix(i, col_��Ŀ����) = rsTmp!��Ŀ�ı�1 & ""
            .TextMatrix(i, col_״̬) = "1"
            .TextMatrix(i, col_�к�) = rsTmp!�к� & ""
            .TextMatrix(i, Col_ҳ��) = rsTmp!ҳ�� & ""
            
            arrtmp = Split(rsTmp!��Ŀ�ı�2, "|")    '=����|�Ƿ�����ѡ��1,ѡ��2,...
            intType = arrtmp(0)
            .TextMatrix(i, col_����) = intType
            
            If intType = T3������ Or intType = T4��ѡ�� Then
                .TextMatrix(i, col_����) = 1
                .Cell(flexcpPictureAlignment, i, col_��Ŀֵ) = flexPicAlignCenterCenter
                .Cell(flexcpAlignment, i, col_��Ŀֵ) = flexAlignCenterCenter
                .Cell(flexcpBackColor, i, col_��Ŀֵ) = Color_MustAddBack
            ElseIf UBound(arrtmp) > 0 Then  '���֣��ַ�������
                .TextMatrix(i, col_����) = arrtmp(1)
                If Val(arrtmp(1)) = 1 Then
                    .Cell(flexcpBackColor, i, col_��Ŀֵ) = Color_MustAddBack
                End If
                If intType = T5��ѡ�� Then
                    .TextMatrix(i, col_��ע) = arrtmp(2)
                    .Cell(flexcpPictureAlignment, i, col_��Ŀֵ) = flexPicAlignCenterCenter
                    .Cell(flexcpAlignment, i, col_��Ŀֵ) = flexAlignCenterCenter
                End If
            Else
                .TextMatrix(i, col_����) = 0
            End If
            
            If intType = T4��ѡ�� Then
                If UBound(arrtmp) > 0 Then
                    .RowData(i) = CStr(Replace(arrtmp(1), ",", "|"))   '��ΪColComboList��ֵ
                    .TextMatrix(i, col_��Ŀֵ) = Split(arrtmp(1), ",")(0)   '��һ����Ϊȱʡֵ
                    If .TextMatrix(i, col_��Ŀֵ) <> "" Then
                        If Mid(.TextMatrix(i, col_��Ŀֵ), 1, 1) = "[" And Mid(.TextMatrix(i, col_��Ŀֵ), Len(.TextMatrix(i, col_��Ŀֵ))) = "]" Then
                            .Cell(flexcpData, i, col_��Ŀֵ) = Mid(.TextMatrix(i, col_��Ŀֵ), 2, InStr(.TextMatrix(i, col_��Ŀֵ), "]") - 2)
                            .TextMatrix(i, col_��Ŀֵ) = ""
                            .RowData(i) = "" '������Ϊ��T4��ѡ�ʱ��������Ϊ������ʽ���߰󶨷�ʽ���ж�����
                        End If
                    End If
                End If
            ElseIf intType = T3������ Or intType = T5��ѡ�� Then
                .Cell(flexcpChecked, i, col_��Ŀֵ) = 2
            ElseIf intType = T6���� Then
                .TextMatrix(i, col_��Ŀֵ) = .TextMatrix(i, col_��Ŀ����)
                .TextMatrix(i, col_��ע) = .TextMatrix(i, col_��Ŀ����)
                .MergeRow(i) = True
            End If
            
            
            If (mlngFun = 1 Or mlngFun = 2) And intType <> T6���� Then
                Select Case intType
                Case T0����
                    .TextMatrix(i, col_��Ŀֵ) = "" & rsTmp!����ֵ
                Case T3������, T5��ѡ��
                    .Cell(flexcpChecked, i, col_��Ŀֵ) = IIf(Val("" & rsTmp!����ֵ) = 1, 1, 2)
                Case T1�ַ�, T4��ѡ��
                    .TextMatrix(i, col_��Ŀֵ) = "" & rsTmp!�ַ�ֵ
                Case T2����
                    If Not IsNull(rsTmp!����ֵ) Then
                        .TextMatrix(i, col_��Ŀֵ) = Format(rsTmp!����ֵ & "", "yyyy-MM-DD")
                    End If
                End Select
                
                If intType <> T5��ѡ�� Then .TextMatrix(i, col_��ע) = "" & rsTmp!��ע
                
                '����ԭֵ�������ж��Ƿ��޸�
                If mlngFun = 2 Then
                    .Cell(flexcpData, i, col_��ע) = "" & .TextMatrix(i, col_��ע)
                    If .Cell(flexcpData, i, col_��Ŀֵ) = "" Then .Cell(flexcpData, i, col_��Ŀֵ) = "" & .TextMatrix(i, col_��Ŀֵ)
                End If
            End If
            
            rsTmp.MoveNext
        Next
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Line1(0).X2 = Me.Width
    Line1(1).X2 = Me.Width
    Line1(2).X2 = Me.Width
    Line1(3).X2 = Me.Width
    tbcSub.Top = picInfo.Height
    tbcSub.Left = 20
    picBottom.Top = Me.Height - tbcSub.Height - tbcSub.Top
    tbcSub.Width = Me.Width - 270
    tbcSub.Height = Me.Height - tbcSub.Top - picBottom.Height - 590
    picBottom.Width = Me.Width
    cmdOK.Left = Me.Width - cmdOK.Width - 1800
    cmdCancel.Left = Me.Width - cmdCancel.Width - 500
    cmdPrintToEXCEL.Left = Me.Width - cmdPrintToEXCEL.Width - 3000
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, mlngFun)
End Sub


Private Sub picItem_Resize()
    On Error Resume Next
    vsItem.Move 0, 0, picItem.Width, picItem.Height
End Sub

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim i As Long
    
    For i = 1 To vsItem.Rows - 1
        If Val(vsItem.TextMatrix(i, Col_ҳ��)) = Val(Item.Tag & "") Then
            vsItem.RowHidden(i) = False
        Else
            vsItem.RowHidden(i) = True
        End If
    Next
End Sub

Private Sub vsItem_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = col_��ע Or Col = col_��Ŀֵ Then
        With vsItem
            If CStr(.Cell(flexcpData, Row, Col)) <> CStr(.TextMatrix(Row, Col)) Or .TextMatrix(Row, col_����) = T3������ Or .TextMatrix(Row, col_����) = T5��ѡ�� Then
                .TextMatrix(Row, col_״̬) = 2
            Else
                .TextMatrix(Row, col_״̬) = 1
            End If
            If .TextMatrix(Row, col_����) = T4��ѡ�� And Col = col_��Ŀֵ And .RowData(Row) = "" Then
                .ColComboList(col_��Ŀֵ) = "..."
            End If
        End With
    End If
End Sub

Private Sub vsItem_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow And NewRow >= vsItem.FixedRows Then
        If vsItem.TextMatrix(NewRow, col_����) = T4��ѡ�� Then
            If vsItem.RowData(NewRow) = "" Then   '������Դ��ʽ
                vsItem.ColComboList(col_��Ŀֵ) = "..."
            Else '������ʽ
                vsItem.ColComboList(col_��Ŀֵ) = vsItem.RowData(NewRow)
            End If
        Else
            vsItem.ColComboList(col_��Ŀֵ) = ""
        End If
    End If
End Sub

Private Sub vsItem_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (Col = col_��Ŀֵ Or Col = col_��ע) Then
        Cancel = True
    Else
        If vsItem.TextMatrix(Row, col_����) = "6" Then
            Cancel = True
        End If
        If Col = col_��ע And vsItem.TextMatrix(Row, col_����) = "5" Then
            Cancel = True
        End If
            
    End If
End Sub

Private Sub vsItem_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strTable As String, strSql As String
    Dim rsTmp As Recordset
    Dim vPoint As POINTAPI, blnCancel As Boolean

    With vsItem
        If Col = col_��Ŀֵ Then
            strTable = .Cell(flexcpData, Row, Col)
            If strTable <> "" Then
                strSql = "Select Rownum as ID,���� From " & strTable
                vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
                On Error GoTo errH
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, strTable, True, "", "", True, True, True, _
                                                     vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True)
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "û���ҵ� " & strTable & " �����ݡ�", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = "-"
                    End If
                    Exit Sub
                Else
                    .TextMatrix(Row, Col) = rsTmp!���� & ""
                    If CStr(.Cell(flexcpData, Row, Col)) <> CStr(.TextMatrix(Row, Col)) Then
                        .TextMatrix(Row, col_״̬) = 2  'ֱ��˫��ѡ��ʱ���ᴥ��vsItem_AfterEdit
                    End If
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

Private Sub vsItem_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    '˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsItem
        If Val(.TextMatrix(Row, col_����)) = T6���� And Col > 0 And Col < col_��ע Then
            vRect.Left = Right - 2
            vRect.Right = Right
            vRect.Top = Top
            vRect.Bottom = Bottom - 1
        Else
            lngLeft = col_��Ŀ���: lngRight = col_��Ŀ����
            If Not Between(Col, lngLeft, lngRight) Then Exit Sub
            If Not RowInһ����Ŀ(Row, lngBegin, lngEnd) Then Exit Sub
    
            vRect.Left = Left '������߱����
            vRect.Right = Right - 1 '�����ұ߱����
            If Row = lngBegin Then
                vRect.Top = Bottom - 1 '���б�����������
                vRect.Bottom = Bottom
            Else
                If Row = lngEnd Then
                    vRect.Top = Top
                    vRect.Bottom = Bottom - 2 '���б����±���(���������õ��±��ߴ�Ϊ2)
                Else
                    vRect.Top = Top
                    vRect.Bottom = Bottom
                End If
            End If
        End If
      

        If Between(Row, .Row, .RowSel) Then
            'SetBkColor hDC, SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, OS.SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Function RowInһ����Ŀ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
    Dim i As Long, blnTmp As Boolean
    
    With vsItem
        If .TextMatrix(lngRow, col_��Ŀ���) = "" Then Exit Function
        If lngRow = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, col_��Ŀ���)) = Val(.TextMatrix(lngRow, col_��Ŀ���)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, col_��Ŀ���)) = Val(.TextMatrix(lngRow, col_��Ŀ���)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col_��Ŀ���)) = Val(.TextMatrix(lngRow, col_��Ŀ���)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col_��Ŀ���)) = Val(.TextMatrix(lngRow, col_��Ŀ���)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInһ����Ŀ = blnTmp
    End With
End Function

Private Sub vsItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode > 127 Then
        '���ֱ�����뺺�ֵ�����
        Call vsItem_KeyPress(KeyCode)
    End If
End Sub

Private Sub vsItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With vsItem
            If .Row = .Rows - 1 And .Col = col_��ע Then
                Call zlCommFun.PressKey(vbKeyTab)
            ElseIf .Col = col_��ע Then
                KeyAscii = 0
                .Select .Row + 1, col_��Ŀֵ
            Else
                KeyAscii = 0
                .Col = .Col + 1
            End If
        End With
    Else
        If KeyAscii = Asc("*") Then
            KeyAscii = 0
            Call vsItem_CellButtonClick(vsItem.Row, vsItem.Col)
        Else
            vsItem.ColComboList(col_��Ŀֵ) = "" 'ʹ��ť״̬��������״̬
        End If
    End If
End Sub

Private Sub vsItem_LostFocus()
    vsItem.ForeColorSel = vbBlack
    vsItem.BackColorSel = vbWhite
End Sub

Private Sub vsItem_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'���ܣ�������Ҫ������ͼ���������ݵ���Ч��
    Dim intType As Integer, blnValidate As Boolean, strMsg As String
    Dim strValue As String
    Dim strTable As String, strSql As String
    Dim rsTmp As Recordset
    Dim vPoint As POINTAPI, blnCancel As Boolean

    With vsItem
        If Col = col_��Ŀֵ Then
            blnValidate = True
            intType = Val(vsItem.TextMatrix(Row, col_����))
            strValue = vsItem.EditText

            If strValue <> "" Then
                Select Case intType
                Case T0����
                    blnValidate = IsNumeric(strValue)
                    strMsg = "Ҫ����������ݱ����������͡�"
                Case T2����
                    blnValidate = IsDate(strValue)
                    strMsg = "Ҫ����������ݱ����������͡�"
                Case T4��ѡ��
                    If .RowData(Row) = "" Then
                        strTable = .Cell(flexcpData, Row, Col)
                        If strTable <> "" Then
                            strSql = "Select Rownum as ID,���� From " & strTable & " Where ���� Like [1] Or Upper(zlspellcode(����)) like [2]"
                            vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
                            On Error GoTo errH
                        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, strTable, True, "", "", True, True, True, _
                                vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, gstrLike & strValue & "%", UCase(strValue) & "%")
                            If rsTmp Is Nothing Then
                                If Not blnCancel Then
                                    strMsg = "û�в��ҵ�ָ�������ݡ�"
                                    blnValidate = False
                                Else
                                    Cancel = True
                                End If
                            Else
                                .EditText = rsTmp!���� & ""
                                .TextMatrix(Row, Col) = rsTmp!���� & ""
                            End If
                        End If
                    End If
                End Select
            End If
            If blnValidate = False Then
                MsgBox strMsg, vbInformation, gstrSysName
                Cancel = True
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



