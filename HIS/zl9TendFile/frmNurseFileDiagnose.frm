VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmNurseFileDiagnose 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����������"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9015
   Icon            =   "frmNurseFileDiagnose.frx":0000
   LinkTopic       =   "frmNurseFileDiagnose"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picStb 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      FillColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   90
      ScaleHeight     =   360
      ScaleWidth      =   8700
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5280
      Width           =   8700
      Begin VB.Label lblStb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   -15
         Width           =   75
      End
   End
   Begin VB.PictureBox picDiagnosis 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   5625
      Left            =   15
      ScaleHeight     =   5595
      ScaleWidth      =   8895
      TabIndex        =   0
      Top             =   30
      Width           =   8925
      Begin VB.ComboBox cboDate 
         Height          =   300
         Left            =   1515
         TabIndex        =   1
         Text            =   "cboDate"
         Top             =   2505
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.TextBox txtTime 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   300
         Left            =   2760
         TabIndex        =   2
         Top             =   2250
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CommandButton cmdOK 
         Height          =   345
         Left            =   5910
         Picture         =   "frmNurseFileDiagnose.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "ȷ��"
         Top             =   4875
         Width           =   1155
      End
      Begin VB.CommandButton cmdCancle 
         Height          =   360
         Left            =   7500
         Picture         =   "frmNurseFileDiagnose.frx":6DDC
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "ȡ��"
         Top             =   4860
         Width           =   1155
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDiag 
         Height          =   4410
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   8655
         _cx             =   15266
         _cy             =   7779
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
         BackColorFixed  =   14811105
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
         FixedCols       =   1
         RowHeightMin    =   305
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmNurseFileDiagnose.frx":7366
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
End
Attribute VB_Name = "frmNurseFileDiagnose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mlng�ļ�ID As Long
Private mstrBeginDate  As String
Private mstrEndDate As String

Private Enum TYPE_diag
    Col_diagName = 0
    Col_choose = 1
    Col_diag���� = 2
    Col_diagʱ�� = 3
    Col_diag������� = 4
    Col_diag���� = 5
    Col_diagID = 6
    Col_diag���id = 7
    Col_diag������� = 8
    Col_diag��ʾ = 9
End Enum

Private Function zlRefreshData(ByVal strTime As String)
    Dim aryPeriod() As String
    Dim blnEnd As Boolean
    Dim i As Integer
    Dim strBeginDate As String, strEndDate As String
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim strDiagnose As String
    
    On Error GoTo ErrHand
    aryPeriod = Split(strTime, "��")
    If UBound(aryPeriod) > 0 Then
        strBeginDate = Format(aryPeriod(0), "YYYY-MM-DD")
        strEndDate = Format(aryPeriod(1), "YYYY-MM-DD")
    Else
        blnEnd = True
    End If
    
    blnEnd = False
     
    Do While Not blnEnd
        If DateDiff("d", strBeginDate, strEndDate) >= 0 Then
            cboDate.AddItem Format(strBeginDate, "YYYY-MM-DD")
        Else
            blnEnd = True
        End If
        strBeginDate = DateAdd("d", 1, strBeginDate)
    Loop

        
        
        strSQL = "Select a.����id, a.��ҳid, a.�������, a.�������, Decode(�Ƿ�����, 1, �Ƿ�����, 0) �Ƿ�����" & vbNewLine & _
            " From ������ϼ�¼ A, ��������Ŀ¼ B, �������Ŀ¼ C, ��������Ŀ¼ D, ����������� E, " & vbNewLine & _
            " (Select Max(��¼��Դ) ��¼��Դ From ������ϼ�¼ Where ����id = [1] And ��ҳid = [2]) G " & vbNewLine & _
            " Where a.����id = b.Id(+) And a.���id = c.Id(+) And a.֤��id = d.Id(+) And b.����id = e.Id(+) And a.��¼��Դ =g.��¼��Դ And" & vbNewLine & _
            "      a.ȡ��ʱ�� Is Null And a.������� Is Not Null And ����id = [1] And ��ҳid = [2]" & vbNewLine & _
            " Union " & vbNewLine & _
            " Select ����id, ��ҳid,�������, �������,�Ƿ�����" & vbNewLine & _
            " From ���˻������" & vbNewLine & _
            " Where  ����id =[1]  And ��ҳid = [2] And �ļ�id = [3]"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��ǰ���е����", mlng����ID, mlng��ҳID, mlng�ļ�ID)
    
    Call showVsfData(rsTemp)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function showVsfData(ByVal rsTemp As ADODB.Recordset)
    Dim lngRow  As Long
    Dim strSQL As String
    Dim strTime As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    vsfDiag.Rows = rsTemp.RecordCount + 1
    vsfDiag.MergeCol(0) = True
    If Not rsTemp.RecordCount > 0 Then Exit Function
    With rsTemp
    lngRow = 1
    .Filter = "�������= 1 or �������= 11 "
    Do While Not .EOF
        vsfDiag.TextMatrix(lngRow, Col_diagName) = "�������"
        vsfDiag.TextMatrix(lngRow, Col_diag�������) = Val(NVL(!�������))
        vsfDiag.TextMatrix(lngRow, Col_diag�������) = NVL(!�������)
        vsfDiag.TextMatrix(lngRow, Col_diag����) = IIf(Val(NVL(!�Ƿ�����)) = 0, "", "��")
        lngRow = lngRow + 1
        .MoveNext
    Loop
    
    .Filter = "�������= 2 or �������= 12 "
    Do While Not .EOF
        vsfDiag.TextMatrix(lngRow, Col_diagName) = "��Ժ���"
        vsfDiag.TextMatrix(lngRow, Col_diag�������) = Val(NVL(!�������))
        vsfDiag.TextMatrix(lngRow, Col_diag�������) = NVL(!�������)
        vsfDiag.TextMatrix(lngRow, Col_diag����) = IIf(Val(NVL(!�Ƿ�����)) = 0, "", "��")
        lngRow = lngRow + 1
        .MoveNext
    Loop
    
    .Filter = "�������= 3 or �������= 13 "
    Do While Not .EOF
        vsfDiag.TextMatrix(lngRow, Col_diagName) = "��Ժ���"
        vsfDiag.TextMatrix(lngRow, Col_diag�������) = Val(NVL(!�������))
        vsfDiag.TextMatrix(lngRow, Col_diag�������) = NVL(!�������)
        vsfDiag.TextMatrix(lngRow, Col_diag����) = IIf(Val(NVL(!�Ƿ�����)) = 0, "", "��")
        lngRow = lngRow + 1
        .MoveNext
    Loop
    
    .Filter = "������� <> 1 and ������� <> 2 and ������� <> 3 and ������� <>  11 and ������� <>  12 and ������� <> 13 "
    Do While Not .EOF
        vsfDiag.TextMatrix(lngRow, Col_diagName) = "�������"
        vsfDiag.TextMatrix(lngRow, Col_diag�������) = Val(NVL(!�������))
        vsfDiag.TextMatrix(lngRow, Col_diag�������) = NVL(!�������)
        vsfDiag.TextMatrix(lngRow, Col_diag����) = IIf(Val(NVL(!�Ƿ�����)) = 0, "", "��")
        lngRow = lngRow + 1
        .MoveNext
    Loop

    End With
    
    strSQL = "select Id,�������,�������,���ʱ��,�Ƿ����� from ���˻������ where ����id= [1] And ��ҳid=[2] And �ļ�id=[3] And ���ʱ�� between [4] and [5] Order By ���ʱ�� desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ȡ���õ����", mlng����ID, mlng��ҳID, mlng�ļ�ID, CDate(mstrBeginDate), CDate(mstrEndDate))
    
    strTime = Format(mstrBeginDate, "YYYY-MM-DD") & " " & Format(zlDatabase.Currentdate, "hh:mm")
    For lngRow = 1 To vsfDiag.Rows - 1
        vsfDiag.TextMatrix(lngRow, Col_diag����) = Format(mstrBeginDate, "YYYY-MM-DD")
        If Not (strTime > mstrBeginDate And strTime < mstrEndDate) Then
            strTime = mstrBeginDate
        End If
        vsfDiag.TextMatrix(lngRow, Col_diagʱ��) = Format(strTime, "hh:mm")
        vsfDiag.TextMatrix(lngRow, Col_diag��ʾ) = 0
    Next
                
    Do While Not rsTmp.EOF
        For lngRow = 1 To vsfDiag.Rows - 1
            If NVL(rsTmp!�������) = vsfDiag.TextMatrix(lngRow, Col_diag�������) And Val(vsfDiag.TextMatrix(lngRow, Col_diag�������)) = Val(rsTmp!�������) _
            And IIf(vsfDiag.TextMatrix(lngRow, Col_diag����) = "", 0, 1) = Val(rsTmp!�Ƿ�����) Then
                vsfDiag.TextMatrix(lngRow, Col_choose) = 1
                vsfDiag.TextMatrix(lngRow, Col_diagID) = Val(rsTmp!ID)
                vsfDiag.TextMatrix(lngRow, Col_diag����) = Format(rsTmp!���ʱ��, "YYYY-MM-DD")
                vsfDiag.TextMatrix(lngRow, Col_diagʱ��) = Format(rsTmp!���ʱ��, "hh:mm")
                vsfDiag.TextMatrix(lngRow, Col_diag��ʾ) = 1
            End If
        Next
        rsTmp.MoveNext
    Loop
    
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ShowEdtor(ByVal frmParent As Object, ByVal PatiId As Long, ByVal lngPageId As Long, ByVal FileID As Long, ByVal strTime As String)
    Dim strSQL As String
    
    mlng����ID = PatiId
    mlng��ҳID = lngPageId
    mlng�ļ�ID = FileID
    mstrBeginDate = Split(strTime, "��")(0)
    mstrEndDate = Split(strTime, "��")(1)
    Call zlRefreshData(strTime)
    Me.Show 1, frmParent
End Function

Private Sub cboDate_Click()
    cboDate.Visible = False
End Sub

Private Sub cboDate_DblClick()
    Call cboDate_Click
End Sub

Private Sub cboDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cboDate.Visible = False
    End If
End Sub

Private Sub cboDate_LostFocus()
    cboDate.Visible = False
End Sub

Private Sub cboDate_Validate(Cancel As Boolean)
    Dim strText As String
    strText = Format(cboDate.Text, "YYYY-MM-DD")
    If Trim(strText) = "" Then
        lblStb.Caption = "���ڲ���Ϊ�գ�"
        lblStb.ForeColor = 255
        Cancel = True
        Exit Sub
    End If
    If Not IsDate(strText) Then
        lblStb.Caption = "¼������ݲ��ǺϷ������ڣ���1��12�գ�2011-01-12"
        lblStb.ForeColor = 255
        Cancel = True
        Exit Sub
    End If
    If Not vsfDiag.TextMatrix(vsfDiag.ROW, Col_diag����) = strText Then
        vsfDiag.TextMatrix(vsfDiag.ROW, Col_diag����) = strText
        vsfDiag.TextMatrix(vsfDiag.ROW, Col_diag��ʾ) = 3
    End If
    cboDate.Text = Format(cboDate.Text, "yyyy-MM-dd")
End Sub

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
'------------------------------------------------
'����:����������Ϣ
'------------------------------------------------
    Dim blnTran As Boolean
    Dim lngID As Long
    Dim strSQL As String
    Dim ArrSQL() As String
    Dim i As Integer, lngItemCode As Long
    Dim lngRow As Long
    Dim str������� As String
    Dim int������� As Integer
    Dim strTime As String
    
    On Error GoTo ErrHand
    
    For lngRow = 1 To vsfDiag.Rows - 1
        strTime = Format(vsfDiag.TextMatrix(vsfDiag.ROW, Col_diag����) & " " & vsfDiag.TextMatrix(vsfDiag.ROW, Col_diagʱ��), "YYYY-MM-DD hh:mm:ss")
        If Not (strTime >= mstrBeginDate And strTime <= mstrEndDate) Then
            If Format(strTime, "YYYY-MM-DD hh:mm") = Format(mstrBeginDate, "YYYY-MM-DD hh:mm") And Format(strTime, "YYYY-MM-DD hh:mm") <= Format(mstrEndDate, "YYYY-MM-DD hh:mm") Then
                strTime = mstrBeginDate
            Else
                MsgBox "��" & lngRow & "��¼�������Ҫ�ڱ�ҳ����ʱ�䷶Χ��:" & mstrBeginDate & " ��" & mstrEndDate & "��", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    Next
    
    Screen.MousePointer = 11
    
    ReDim Preserve ArrSQL(1 To 1)
    
    For lngRow = 1 To vsfDiag.Rows - 1
        If Not Val(vsfDiag.TextMatrix(lngRow, Col_choose)) = Val(vsfDiag.TextMatrix(lngRow, Col_diag��ʾ)) Then
            lngID = Val(vsfDiag.TextMatrix(lngRow, Col_diagID))
            str������� = NVL(vsfDiag.TextMatrix(lngRow, Col_diag�������))
            int������� = Val(vsfDiag.TextMatrix(lngRow, Col_diag�������))
            strTime = Format(NVL(vsfDiag.TextMatrix(lngRow, Col_diag����)) & " " & NVL(vsfDiag.TextMatrix(lngRow, Col_diagʱ��)), "YYYY-MM-DD HH:mm:ss")
            If strTime < mstrBeginDate Then strTime = mstrBeginDate
            strTime = " To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss')"
            strSQL = "Zl_���˻������_Update("
            strSQL = strSQL & lngID & ","
            strSQL = strSQL & mlng����ID & ","
            strSQL = strSQL & mlng��ҳID & ","
            strSQL = strSQL & mlng�ļ�ID & ","
            strSQL = strSQL & IIf(Val(vsfDiag.TextMatrix(lngRow, Col_choose)) = 0, "''", "'" & str������� & "'") & ","
            strSQL = strSQL & int������� & ","
            strSQL = strSQL & strTime & ","
            strSQL = strSQL & IIf(NVL(vsfDiag.TextMatrix(lngRow, Col_diag����)) = "", 0, 1) & ")"
            ArrSQL(ReDimArray(ArrSQL)) = strSQL
            
        End If
    Next
    
    '------------------------------------------------------------------------------------------------------------------
    'ѭ��ִ��SQL��������
    'Debug.Print "----���濪ʼ:" & Now
    gcnOracle.BeginTrans
    blnTran = True
    For i = 1 To UBound(ArrSQL)
        If ArrSQL(i) <> "" Then Call zlDatabase.ExecuteProcedure(CStr(ArrSQL(i)), "������������"): ' Debug.Print CStr(arrSQL(i))
    Next
    gcnOracle.CommitTrans
    blnTran = False

    Screen.MousePointer = 0
    Unload Me
    Exit Sub
ErrHand:
    If blnTran = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
    With picStb
        .Top = picDiagnosis.Height + picDiagnosis.Top - picStb.Height
        .Left = picDiagnosis.Left + 20
        .Height = TextHeight("����") + 50
        .Width = picDiagnosis.Width - 70
    End With

    With lblStb
        .Font.Size = 9
        .Height = TextHeight("����")
        .Top = (picStb.Height - .Height) \ 2
        .Left = 10
    End With
End Sub

Private Sub txtTime_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strText As String
    Dim strInfo As String
    
    strText = txtTime.Text
    If KeyCode = vbKeyReturn Then
        strText = CheckTime(strText, strInfo)
        If strInfo = "" Then
            If Not vsfDiag.TextMatrix(vsfDiag.ROW, Col_diagʱ��) = strText Then
                vsfDiag.TextMatrix(vsfDiag.ROW, Col_diagʱ��) = strText
                vsfDiag.TextMatrix(vsfDiag.ROW, Col_diag��ʾ) = 3
            End If
            txtTime.Visible = False
        Else
            lblStb.Caption = strInfo
            lblStb.ForeColor = 255
        End If
    End If
    
End Sub

Private Function CheckTime(ByVal strText As String, ByRef strInfo As String) As String
    Dim arrTime() As String
    Dim strTime As String
    
    If Trim(strText) = "" Then
        strInfo = "ʱ�䲻��Ϊ�գ�"
        Exit Function
    End If
    If InStr(1, Trim(strText), ":") = 0 Then
        Select Case Len(strText)
        Case 3, 4
            strText = String(4 - Len(strText), "0") & strText
            strText = Mid(strText, 1, 2) & ":" & Mid(strText, 3)
        Case Is < 3
            strText = String(2 - Len(strText), "0") & strText
            strText = Format(Now, "HH") & ":" & strText
        End Select
    End If
    arrTime = Split(Trim(strText), ":")
    
    If UBound(arrTime) <> 1 Then
        strInfo = "¼���ʱ���ʽ�Ƿ���[Сʱ:����]"
        Exit Function
    Else
        If Len(Trim(arrTime(0))) < 2 Then arrTime(0) = String(2 - Len(Trim(arrTime(0))), "0") & Trim(arrTime(0))
        If Len(Trim(arrTime(1))) < 2 Then arrTime(1) = String(2 - Len(Trim(arrTime(1))), "0") & Trim(arrTime(1))
        strText = arrTime(0) & ":" & arrTime(1)
    End If
    
    '�Ϸ��Լ��
    If IsNumeric(arrTime(0)) = False Or IsNumeric(arrTime(1)) = False Or Len(Trim(arrTime(0))) > 2 Or Len(Trim(arrTime(1))) > 2 Then
        strInfo = "¼���ʱ���ʽ�Ƿ���[Сʱ:����]"
        Exit Function
    End If
    If Mid(strText, 3, 1) <> ":" Then
        strInfo = "¼���ʱ���ʽ�Ƿ���[Сʱ:����]"
        Exit Function
    End If
    If Val(arrTime(0)) < 0 Or Val(arrTime(0)) > 23 Then
        strInfo = "¼���ʱ���ʽ�Ƿ���[СʱӦ��0��23֮��]"
        Exit Function
    End If
    If Val(arrTime(1)) < 0 Or Val(arrTime(1)) > 59 Then
        strInfo = "¼���ʱ���ʽ�Ƿ���[����Ӧ��0��59֮��]"
        Exit Function
    End If
    strTime = Format(vsfDiag.TextMatrix(vsfDiag.ROW, Col_diag����) & " " & strText, "YYYY-MM-DD hh:mm")
    If Not (strTime >= Format(mstrBeginDate, "YYYY-MM-DD hh:mm") And strTime <= Format(mstrEndDate, "YYYY-MM-DD hh:mm")) Then
        strInfo = "¼���ʱ��Ҫ�ڱ�ҳ����ʱ�䷶Χ��:" & mstrBeginDate & " ��" & mstrEndDate
        Exit Function
    End If
    CheckTime = strText
    
        
End Function

Private Sub txtTime_LostFocus()
    txtTime.Visible = False
End Sub

Private Sub vsfDiag_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    lblStb.Caption = ""
End Sub

Private Sub vsfDiag_Click()
    If vsfDiag.COL = Col_choose And vsfDiag.ROW > 0 Then
        If Val(vsfDiag.TextMatrix(vsfDiag.ROW, Col_choose)) = 1 Then
            vsfDiag.TextMatrix(vsfDiag.ROW, Col_choose) = 0
        Else
            vsfDiag.TextMatrix(vsfDiag.ROW, Col_choose) = 1
        End If
    End If
End Sub

Private Sub vsfDiag_DblClick()
    Dim lngRow As Long, lngCol As Long
    lngRow = vsfDiag.ROW
    lngCol = vsfDiag.COL
    If lngCol = Col_diag���� Or lngCol = Col_diagʱ�� Then
        If vsfDiag.COL = Col_diag���� Then
            cboDate.Visible = True
            cboDate.Text = vsfDiag.TextMatrix(lngRow, lngCol)
            cboDate.Move vsfDiag.CellLeft + vsfDiag.Left + 20, vsfDiag.CellTop + vsfDiag.Top + 20, vsfDiag.CellWidth
            cboDate.SetFocus
        End If
        If vsfDiag.COL = Col_diagʱ�� Then
            txtTime.Visible = True
            txtTime.Text = vsfDiag.TextMatrix(lngRow, lngCol)
            txtTime.Move vsfDiag.CellLeft + vsfDiag.Left + 20, vsfDiag.CellTop + vsfDiag.Top + 20, vsfDiag.CellWidth
            txtTime.SetFocus
        End If
    End If
    
End Sub
