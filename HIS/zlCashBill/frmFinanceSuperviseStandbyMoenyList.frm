VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmFinanceSuperviseStandbyMoenyList 
   BorderStyle     =   0  'None
   Caption         =   "���ý��б�"
   ClientHeight    =   8565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11625
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11625
      TabIndex        =   0
      Top             =   0
      Width           =   11625
      Begin VB.CheckBox chkCancel 
         Caption         =   "�����ռ�¼(&C)"
         Height          =   210
         Left            =   6240
         TabIndex        =   8
         Top             =   180
         Width           =   2130
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "���¹�������(&R)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   8655
         TabIndex        =   3
         Top             =   105
         Width           =   1605
      End
      Begin VB.ComboBox cboPerson 
         Height          =   330
         Left            =   1020
         TabIndex        =   2
         Text            =   "cboPerson"
         Top             =   120
         Width           =   2040
      End
      Begin VB.TextBox txtNO 
         Height          =   345
         Left            =   3600
         TabIndex        =   1
         Top             =   113
         Width           =   2415
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2355
         TabIndex        =   6
         Top             =   150
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label lblPerson 
         AutoSize        =   -1  'True
         Caption         =   "�շ�Ա"
         Height          =   210
         Left            =   315
         TabIndex        =   5
         Top             =   180
         Width           =   630
      End
      Begin VB.Label lblNo 
         AutoSize        =   -1  'True
         Caption         =   "NO"
         Height          =   210
         Left            =   3330
         TabIndex        =   4
         Top             =   165
         Width           =   210
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsList 
      Height          =   1800
      Left            =   360
      TabIndex        =   7
      Top             =   1110
      Width           =   8625
      _cx             =   15214
      _cy             =   3175
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
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
      BackColorSel    =   12632256
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFinanceSuperviseStandbyMoenyList.frx":0000
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
Attribute VB_Name = "frmFinanceSuperviseStandbyMoenyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsPerson As ADODB.Recordset
Private mlngModule As Long, mstrPrivs As String
Private mblnDrop As Boolean

Public Sub zlInitVar(ByVal lngModule As Long, ByVal strPrivs As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ر���
    '���:lngModule-ģ���
    '       strPrivs-Ȩ�޴�
    '����:���˺�
    '����:2013-09-09 14:41:46
    '˵��:���ش����,��������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
End Sub
Private Sub InitGrid()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������ؼ�
    '����:���˺�
    '����:2013-09-11 17:34:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strHead As String, varData As Variant
    strHead = "ID,���ݺ�,���,��ע,������,�ջ���,�ջ�ʱ��,�Ǽ���,�Ǽ�ʱ��"
    varData = Split(strHead, ",")
    With vsList
        .Clear
        .Rows = 2: .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = varData(i)
            .ColKey(i) = varData(i)
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*ʱ��" Or .ColKey(i) = "���ݺ�" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*���" Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsList, Me.Name, "���ý���Ϣ�б�", False
    End With
End Sub
Private Function LoadData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ʷ�տ�����
    '����:���ݼ��سɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-11 17:08:50
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strPerson As String, strWhere As String, i As Long, blnDel As Boolean
    
    
    On Error GoTo errHandle
     strPerson = zlStr.NeedName(cboPerson.Text)
    If txtNO.Text <> "" Then
        strWhere = strWhere & " And A.NO=[1]"
    Else
        strWhere = strWhere & " And A.�տ�Ա=[2]"
    End If
    If chkCancel.Value <> 1 Then
        strWhere = strWhere & " And A.�ջ�ʱ�� is null "
    End If
    strSQL = "" & _
    "   Select A.ID, A.NO As ���ݺ�, LTrim(To_Char(A.���, '99999999990.00')) As ���, A.��ע, " & _
    "        A.�տ�Ա As ������,  " & _
    "        A.�ջ���, to_char(A.�ջ�ʱ��,'yyyy-mm-dd hh24:mi:ss') as �ջ�ʱ��, " & _
    "        A.�Ǽ���,  to_char(A.�Ǽ�ʱ��,'yyyy-mm-dd hh24:mi:ss') as �Ǽ�ʱ�� " & _
    "   From ��Ա�ݴ��¼ A" & _
    "   Where MOD(A.��¼����,10) = 1 " & strWhere & _
    "   Order by �Ǽ�ʱ�� Desc,NO Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(Trim(txtNO.Text)), strPerson)
    
    With vsList
        .Clear 1: .Rows = 2
        .FixedRows = 1
        If rsTemp.RecordCount <> 0 Then
            Set .DataSource = rsTemp
        End If
        For i = 0 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*ID" Then .ColHidden(i) = True
            If .ColKey(i) Like "*ʱ��" Or .ColKey(i) = "���ݺ�" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*���" Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        For i = 1 To .Rows - 1
            blnDel = Trim(.TextMatrix(i, .ColIndex("�ջ�ʱ��"))) <> ""
            If blnDel Then
                '���ϼ�¼���ú�ɫ����
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
            End If
        Next
        .Row = 1
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsList, Me.Name, "���ý���Ϣ�б�", False
        If .Enabled And .Visible Then .SetFocus
    End With
    LoadData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function

Private Sub cboPerson_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer, iCount As Integer
    Dim strText As String, strResult As String, strFilter As String
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮
    Dim rsTemp As ADODB.Recordset
    If KeyAscii <> 13 Then Exit Sub
    
    If cboPerson.Locked Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    
    strText = UCase(cboPerson.Text)
    If cboPerson.ListIndex <> -1 Then
        '�����б�ʱ,�����ı�������������
        If strText <> UCase(cboPerson.List(cboPerson.ListIndex)) Then Call zlcontrol.CboSetIndex(cboPerson.hWnd, -1)
    End If
    If strText = "" Then cboPerson.ListIndex = -1: Exit Sub
    '69061,������,2013-12-30,����ˢ���б�ķ�ʽ����
    If cboPerson.ListIndex >= 0 Then
        Call cmdRefresh_Click
        Exit Sub
    End If
    intIdx = -1
    '�ȸ��Ƽ�¼��
    Set rsTemp = zlDatabase.zlCopyDataStructure(mrsPerson)
    strCompents = Replace(gstrLike, "%", "*") & strText & "*"
    If IsNumeric(strText) Then
        intInputType = 0 '0-�������ȫ����
    ElseIf zlCommFun.IsCharAlpha(strText) Then
        intInputType = 1 '1-�������ȫ��ĸ
    Else
        intInputType = 2 '2-����
    End If
    mrsPerson.Filter = 0: iCount = 0
    With mrsPerson
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not mrsPerson.EOF
            Select Case intInputType
            Case 0  '�������ȫ����
                '������������,��Ҫ���:
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012���ֿ�,������������01����01���,��ֱ�Ӷ�λ��01,�򲻶�λ��1��.
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                If Nvl(!���) = strText Then strResult = Nvl(!����): iCount = 0: Exit Do
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                If Val(Nvl(!���)) = Val(strText) Then
                    If iCount = 0 Then strResult = Nvl(!����)
                    iCount = iCount + 1
                End If
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                 If Val(mrsPerson!���) Like strText & "*" Then
                    If CheckPersonExists(Nvl(!����)) Then Call zlDatabase.zlInsertCurrRowData(mrsPerson, rsTemp)
                 End If
                 
            Case 1  '�������ȫ��ĸ
                '����:
                ' 1.����ļ������,��ֱ�Ӷ�λ
                ' 2.���ݲ�����ƥ����ͬ����
                
                '1.����ļ������,��ֱ�Ӷ�λ
                If Trim(Nvl(!����)) = strText Then
                    If iCount = 0 Then strResult = Nvl(!����)   '���ܴ��ڶ����ͬ����
                    iCount = iCount + 1
                End If
                '2.���ݲ�����ƥ����ͬ����
                If Trim(Nvl(!����)) Like strCompents Then
                    If CheckPersonExists(Nvl(!����)) Then Call zlDatabase.zlInsertCurrRowData(mrsPerson, rsTemp)
                End If
            Case Else  ' 2-����
                '����:���ܴ��ں��ֵ����,����������N001���������ZYK01�������
                '1.����\�������,ֱ�Ӷ�λ
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                
                '1.����\�������,ֱ�Ӷ�λ
                If Trim(!���) = strText Or Trim(!����) = strText Or Trim(!����) = strText Then
                    If iCount = 0 Then strResult = Nvl(!����)   '���ܴ��ڶ����ͬ�Ķ��
                    iCount = iCount + 1
                End If
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                If Trim(!���) Like strText & "*" Or Trim(Nvl(!����)) Like strCompents Or Trim(Nvl(!����)) Like strCompents Then
                    If CheckPersonExists(Nvl(!����)) Then Call zlDatabase.zlInsertCurrRowData(mrsPerson, rsTemp)
                End If
            End Select
            mrsPerson.MoveNext
        Loop
    End With
    
    If iCount > 1 Then strResult = ""
    If strResult = "" And rsTemp.RecordCount = 1 Then strResult = Nvl(rsTemp!����)
    'ֱ�Ӷ�λ
    If strResult <> "" Then
        rsTemp.Close: Set rsTemp = Nothing
        If CheckPersonExists(strResult, True) Then Call cmdRefresh_Click
        Exit Sub
    End If
     If rsTemp.RecordCount = 0 Then
        'δ�ҵ�
        rsTemp.Close: Set rsTemp = Nothing
        KeyAscii = 0: zlcontrol.TxtSelAll cboPerson: Exit Sub
     End If
     
    '�Ȱ�ĳ�ַ�ʽ��������
    Select Case intInputType
    Case 0 '����ȫ����
        rsTemp.Sort = "���"
    Case 1 '����ȫƴ��
        rsTemp.Sort = "����"
    Case Else
        '����ѡ������
        rsTemp.Sort = "���"
    End Select
    '����ѡ����
    Dim rsReturn As ADODB.Recordset
    If zlDatabase.zlShowListSelect(Me, glngSys, mlngModule, cboPerson, rsTemp, True, "", "", rsReturn) Then
        If cboPerson.Enabled Then cboPerson.SetFocus
        If Not rsReturn Is Nothing Then
            If rsReturn.RecordCount <> 0 Then
                '���ж�λ
                If CheckPersonExists(Nvl(rsReturn!����), True) Then
                    'zlCommFun.PressKey vbKeyTab
                End If
            End If
        End If
    End If
    rsTemp.Close: Set rsTemp = Nothing
End Sub

Private Sub cboPerson_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub cboPerson_Validate(Cancel As Boolean)
    If cboPerson.Text <> "" Then
        If cbo.FindIndex(cboPerson, zlStr.NeedName(cboPerson.Text), True) = -1 Then cboPerson.ListIndex = -1: cboPerson.Text = ""
    End If
    If cboPerson.Text = "" Then Call cboPerson_KeyPress(vbKeyReturn)
    '�����ݣ���������
    If cboPerson.ListIndex = -1 And cboPerson.ListCount <> 0 Then Cancel = True
End Sub
Private Sub cboPerson_GotFocus()
    Call zlCommFun.OpenIme(True)
    Call zlcontrol.TxtSelAll(cboPerson)
End Sub
 

Private Sub cmdRefresh_Click()
    txtNO.Text = ""
    Call LoadData
End Sub
Private Sub Form_Load()
    Call InitGrid
    Call LoadPerson
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With vsList
        .Left = ScaleLeft + 50
        .Top = picTop.Top + picTop.Height + 50
        .Height = ScaleHeight - .Top + 50
        .Width = ScaleWidth - .Left * 2
    End With
End Sub
Private Function LoadPerson() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����շ�Ա
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-23 11:59:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    gstrSQL = "" & _
    "   Select distinct A.ID,A.���,A.����,A.����  " & _
    "   From ��Ա�� A,��Ա����˵�� B " & _
    "   Where A.id = B.��ԱID  " & _
    "               And B.��Ա���� In ('����Һ�Ա','�����շ�Ա','Ԥ���տ�Ա','סԺ����Ա','��Ժ�Ǽ�Ա','�����Ǽ���')  " & _
    "               And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
    "               And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
    "   Order By ���"
    Set mrsPerson = zlDatabase.OpenSQLRecord(gstrSQL, "��鵱ǰ����Ա�Ƿ�Ϊ��Ӧ������Ա")
    With cboPerson
        Do While Not mrsPerson.EOF
            .AddItem Nvl(mrsPerson!���) & "-" & Nvl(mrsPerson!����)
            .ItemData(.NewIndex) = Val(Nvl(mrsPerson!ID))
            If .ListIndex < 0 Then .ListIndex = .NewIndex
            mrsPerson.MoveNext
        Loop
        If .ListCount <> 0 And .ListIndex < 0 Then .ListIndex = 0
    End With
    LoadPerson = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function isCheckPersonExists(ByVal str���� As String, _
    Optional blnLocateItem As Boolean = False) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ƿ����շ�Ա�����б���
    '���:str����-����
    '     blnLocateItem:�Ƿ�ֱ�Ӷ�λ
    '����:���ڷ���true,���򷵻�False
    '����:���˺�
    '����:2013-09-23 14:34:47
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To cboPerson.ListCount - 1
        If zlStr.NeedName(cboPerson.List(i)) = str���� Then
            If blnLocateItem Then cboPerson.ListIndex = i
            isCheckPersonExists = True
            Exit Function
        End If
    Next
End Function
Private Sub txtNO_GotFocus()
    zlcontrol.TxtSelAll txtNO
    zlCommFun.OpenIme False

End Sub
Private Sub txtNO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Or Trim(txtNO.Text) = "" Then Exit Sub
    txtNO.Text = GetFullNO(Trim(txtNO.Text), 141)
    Call LoadData
End Sub

Public Sub zlPrint(ByVal bytMode As Byte)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����б���Ϣ
    '���:bytMode=1-��ӡ,2-Ԥ��,3-�����Excel
    '����:���˺�
    '����:2013-09-13 10:23:30
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As New zlPrint1Grd, objRow As New zlTabAppRow
    
    Err = 0: On Error GoTo ErrHand:
    
    '����տ���Ϣ
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstr��λ���� & "���ý𷢷����"
    Set objRow = New zlTabAppRow
    If txtNO.Text <> "" Then
        objRow.Add "���ݺţ�" & txtNO.Text
    Else
        objRow.Add "�շ�Ա��" & cboPerson.Text
    End If
    If chkCancel.Value = 1 Then
        objRow.Add "�����ϵı��ý�"
    End If
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    Set objPrint.Body = vsList
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub
Public Sub RePrintBill()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ش��տ��վ�
    '����:���˺�
    '����:2013-09-13 16:00:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As String
    If Not (zlStr.IsHavePrivs(mstrPrivs, "���ý����õ���ӡ") And zlStr.IsHavePrivs(mstrPrivs, "�ش��ý����õ�")) Then Exit Sub
    With vsList
        If .Row < 1 Then Exit Sub
        strNO = Trim(.TextMatrix(.Row, .ColIndex("���ݺ�")))
    End With
    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1500_1", Me, "NO=" & strNO, 2)
End Sub
Public Sub zlRefresh()
    '���½�������ˢ��
    Call cmdRefresh_Click
End Sub
Public Sub CallCustomRpt(ByVal frmMain As Object, ByVal lngSys As Long, ByVal strRptCode As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Զ��屨��
    '���:lngSys-ϵͳ��
    '        strRptCode-������
    '����:���˺�
    '����:2013-09-17 10:18:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngDeptID As Long
    Dim lngID As Long, dtStartDate As Date, dtEndDate As Date, blnDel As Boolean
    Dim strNO As String
    With vsList
        If .Row < 1 Then Exit Sub
        strNO = Trim(.TextMatrix(.Row, .ColIndex("���ݺ�")))
        lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        blnDel = Trim(.TextMatrix(.Row, .ColIndex("�ջ�ʱ��"))) <> ""
    End With
    Call ReportOpen(gcnOracle, lngSys, strRptCode, frmMain, _
        "NO=" & strNO, _
        "ID=" & lngID, _
        "���ϱ�־=" & IIf(blnDel, 1, 0))
End Sub

Public Function zlPayOnWorkMoney(ByVal frmMain As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ϸڱ��ý�
    '����:���ųɹ�����true,���򷵻�False
    '����:������
    '����:2013-12-4
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str������ As String
    Dim frmNew As New frmFinanceSuperviseStandbyMoneyEdit
    Dim blnReturn As Boolean
    On Error GoTo errHandle
    If cboPerson.ListIndex > 0 Then
        If cboPerson.ItemData(cboPerson.ListIndex) <> 0 Then
            str������ = zlStr.NeedName(cboPerson.Text)
        End If
    End If
    blnReturn = frmNew.EditCard(frmMain, EM_ED_�ϸ�, mlngModule, mstrPrivs, str������, 0)
    If Not frmNew Is Nothing Then Unload frmNew
    Set frmNew = Nothing
    If blnReturn Then zlRefresh
    zlPayOnWorkMoney = blnReturn
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlPayStandbyMoney(ByVal frmMain As Object) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ű��ý�
    '����:���ųɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-10-12 16:45:53
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str������ As String
    Dim frmNew As New frmFinanceSuperviseStandbyMoneyEdit
    Dim blnReturn As Boolean
    On Error GoTo errHandle
    If cboPerson.ListIndex > 0 Then
        If cboPerson.ItemData(cboPerson.ListIndex) <> 0 Then
            str������ = zlStr.NeedName(cboPerson.Text)
        End If
    End If
    blnReturn = frmNew.EditCard(frmMain, EM_ED_����, mlngModule, mstrPrivs, str������, 0)
    If Not frmNew Is Nothing Then Unload frmNew
    Set frmNew = Nothing
    If blnReturn Then zlRefresh
    zlPayStandbyMoney = blnReturn
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckPersonExists(ByVal str���� As String, Optional blnLocateItem As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ƿ������շ�Ա�����б���.
    '���:str����-����
    '     blnLocateItem:�Ƿ�ֱ�Ӷ�λ
    '����:
    '����:���ڷ���true,���򷵻�False
    '����:���˺�
    '����:2009-07-20 17:53:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To cboPerson.ListCount - 1
        If zlStr.NeedName(cboPerson.List(i)) = str���� Then
            If blnLocateItem Then cboPerson.ListIndex = i
            CheckPersonExists = True
            Exit Function
        End If
    Next
End Function

Private Sub vsList_DblClick()
    Dim lngID As Long
    Dim frmNew As frmFinanceSuperviseStandbyMoneyEdit
    With vsList
        If .Row < 0 Or .Row > .Rows - 1 Then Exit Sub
        lngID = .TextMatrix(.Row, .ColIndex("ID"))
        If lngID = 0 Then Exit Sub
    End With
    On Error GoTo errHandle
    Set frmNew = New frmFinanceSuperviseStandbyMoneyEdit
    Call frmNew.EditCard(Me, EM_ED_�鿴, mlngModule, mstrPrivs, "", lngID)
    If Not frmNew Is Nothing Then Unload frmNew
    Set frmNew = Nothing
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsList_GotFocus()
    Call zl_VsGridGotFocus(vsList)
End Sub
Private Sub vsList_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsList, GRD_LOSTFOCUS_COLORSEL)
    vsList.Tag = "0"
End Sub
Private Sub vsList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Name, "���ý���Ϣ�б�", False, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub
Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsList, OldRow, NewRow, OldCol, NewCol)
End Sub
Private Sub vsList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Name, "���ý���Ϣ�б�", False, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub
Public Function CancelStandbyMoney() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ϱ��ý�
    '����:���ϳɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-10-12 18:00:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    Dim strSQL As String, lngID As Long, strNO As String
    Dim strTime As String
    On Error GoTo errHandle
    With vsList
        If .Row < 0 Or .Row > .Rows - 1 Then
            Exit Function
        End If
        lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        strNO = Trim(.TextMatrix(.Row, .ColIndex("���ݺ�")))
        If lngID = 0 Then Exit Function
    End With
    If MsgBox("���Ƿ����Ҫ�ջص���Ϊ" & strNO & "�ı��ý���?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    strTime = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    ' Zl_��Ա�ݴ��¼_Cancel
    strSQL = "Zl_��Ա�ݴ��¼_Cancel("
    '  Id_In       In ��Ա�ݴ��¼.Id%Type,
    strSQL = strSQL & "" & lngID & ","
    '  �ջ���_In   In ��Ա�ݴ��¼.�ջ���%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  �ջ�ʱ��_In In ��Ա�ݴ��¼.�ջ�ʱ��%Type
    strSQL = strSQL & "to_date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'))"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    With vsList
        If chkCancel.Value = 1 Then
            .TextMatrix(.Row, .ColIndex("�ջ���")) = UserInfo.����
            .TextMatrix(.Row, .ColIndex("�ջ�ʱ��")) = strTime
            .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = vbRed
        Else
            lngRow = .Row
            If (.Row < .Rows - 1 Or .Row >= 1) And .Rows - 1 > 1 Then
                If .Row = .Rows - 1 Then
                    .Row = lngRow - 1
                Else
                    .Row = lngRow + 1
                End If
                If .RowIsVisible(.Row) = False Then Call .ShowCell(.Row, .Col)
                .RemoveItem lngRow
            ElseIf .Row = 1 Then
                .Cell(flexcpText, .Row, 0, .Row, .Cols - 1) = ""
            Else
                Call zlRefresh
            End If
        End If
    End With
    If vsList.Enabled And vsList.Visible Then vsList.SetFocus
    Call vsList_GotFocus
    CancelStandbyMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Property Get IsAllowCancel() As Boolean
    '�����Ƿ����
    Dim lngID As Long
    With vsList
        If .Row < 0 Or .Row > .Rows - 1 Then IsAllowCancel = False: Exit Property
        lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        If lngID = 0 Then IsAllowCancel = False: Exit Property
        IsAllowCancel = Trim(.TextMatrix(.Row, .ColIndex("�ջ�ʱ��"))) = ""
    End With
End Property
 
