VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPathImportOut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�����ٴ�·��ѡ��"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6690
   Icon            =   "frmPathImportOut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   6690
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   6690
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2145
      Width           =   6690
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   5520
         TabIndex        =   5
         Top             =   195
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   4200
         TabIndex        =   4
         Top             =   195
         Width           =   1100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   10000
         Y1              =   45
         Y2              =   45
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   10000
         Y1              =   30
         Y2              =   30
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPath 
      Height          =   1185
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6495
      _cx             =   11456
      _cy             =   2090
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
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   320
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathImportOut.frx":169B2
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   0   'False
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
   End
   Begin VB.Label lblPait 
      Caption         =   "��ǰ���ˣ���С��,��,46��"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblPath 
      Caption         =   "����±���ѡ��һ�������ڸò��˵��ٴ�·��"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6495
   End
End
Attribute VB_Name = "frmPathImportOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mPati               As TYPE_Pati
Private mfrmParent          As Object

Private mlngDiagnosisType   As Long             '�������:1-��ҽ�������;11-��ҽ�������
Private mlngDiagnosisSorce  As Long             '�����Դ 1-������3-��ҳ����
Private mlng����ID          As Long
Private mlng���ID          As Long

Private mblnOK              As Boolean

Private mrsDiag             As ADODB.Recordset
Private mrsPath             As ADODB.Recordset
Private mrsPati             As ADODB.Recordset

Public Function ShowMe(frmParent As Object, t_pati As TYPE_Pati, Optional blnImport As Boolean) As Boolean
'����: blnImport:true-�����ť����;false-�Զ�����
'���ܣ�����һ������������д�����֮���Զ�����Ĺ��ܣ�ֻ������Ҫ��ϣ��������Ҫ��ϣ��������·����ť��ʱ���ټ����Ҫ���
    Dim str��Ҫ������� As String
    Dim str��Ҫ������� As String
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim bln��ҽ As Boolean
    Dim i As Long
    Dim blnPath As Boolean
    Dim blnSecondPath As Boolean
    Dim blnPathSend As Boolean
    
    mPati = t_pati
    Set mfrmParent = frmParent
    
    mblnOK = False
    
    mlngDiagnosisType = 0
    mlngDiagnosisSorce = 0
    '���ò����Ƿ����ɹ���Ŀ
    blnPathSend = CheckOutPathSend(mPati.�Һ�ID)

    If blnPathSend Then
        MsgBox "�ò����Ѿ��������ٴ�·�����������ٴε��롣", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���ò����Ƿ��ڱ�������δ��ɵ��ٴ�·�����Ƿ���Լ���·��
    If Not blnPathSend Then
        strSql = " Select 1 From ��������·�� Where ����ID=[1] and ����ID+0 = [2] and ״̬ = 1 "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "CheckOutPathSend", mPati.����ID, mPati.����ID)
        If rsTmp.RecordCount > 0 Then
            MsgBox "�ò����ڱ����Ҵ���δ��ɵ��ٴ�·�������ܹ���������·����", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '����·��
    'ȡ��������ϣ�û�еĻ���ʾ��
    '�жϵ�һ��ϵ��Ƿ��к��ʵ��ٴ�·�����еĻ����룬û�еĻ�����ʾ��
    '�ж����������Ƿ��з��ϵ��ٴ�·�����еĻ�ѯ�ʵ��룬û�еĻ���ʾ��
    Set rsTmp = Get���ﲡ��ID(mPati.����ID, mPati.�Һ�ID, 0, mPati.����ID, bln��ҽ)
    Set mrsDiag = rsTmp
    If rsTmp.RecordCount = 0 Then
        MsgBox "�ò���û����д�κ���ϣ�������д����ִ�е��롣", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Not blnPathSend Then
'        rsTmp.Filter = "��ϴ��� = 1"
        For i = 1 To rsTmp.RecordCount                      '֮������ѭ��������Ϊ�п�����ҽһ������ҽһ��
            mlng����ID = Val("" & rsTmp!����id)
            mlng���ID = Val("" & rsTmp!���id)
            str��Ҫ������� = "" & rsTmp!�������
            mlngDiagnosisType = Val("" & rsTmp!�������)
            mlngDiagnosisSorce = Val("" & rsTmp!��¼��Դ)
            Set mrsPath = GetOutPathTable(mlng����ID, mlng���ID, mPati.����ID)
            If mrsPath.RecordCount > 0 Then
                blnPath = True
                Exit For
            End If
        Next
        
'        If Not blnPath And blnImport Then
'            rsTmp.Filter = "��ϴ��� <> 1"
'            For i = 1 To rsTmp.RecordCount
'                mlng����ID = Val("" & rsTmp!����id)
'                mlng���ID = Val("" & rsTmp!���id)
'                str��Ҫ������� = "" & rsTmp!�������
'                mlngDiagnosisType = Val("" & rsTmp!�������)
'                mlngDiagnosisSorce = Val("" & rsTmp!��¼��Դ)
'                Set mrsPath = GetOutPathTable(mlng����ID, mlng���ID, mPati.����ID)
'                If mrsPath.RecordCount > 0 Then
'                    blnSecondPath = True
'                    Exit For
'                End If
'            Next
'        End If
        
        If Not blnPath And Not blnSecondPath Then
            MsgBox "��ǰ����û���ʺ��ڸò�����Ҫ���[" & str��Ҫ������� & "]�������ٴ�·����", vbInformation, gstrSysName
            Exit Function
        ElseIf (Not blnPath And blnSecondPath And blnImport) Then
            If MsgBox("��ǰ����û���ʺ��ڸò�����Ҫ���[" & str��Ҫ������� & "]�������ٴ�·����" & vbCrLf & _
                            "�������ʺ��ڸò��˴�Ҫ���[" & str��Ҫ������� & "]�������ٴ�·����" & vbCrLf & _
                            "�Ƿ����Ҫ��ϵ��ٴ�·����", vbInformation + vbYesNo, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
        
        Set mrsPati = GetPatiInfoOut(mPati.����ID, mPati.�Һ�ID)
        If mrsPati.RecordCount = 0 Then
            MsgBox "��ȡ���˵�ǰ������Ϣʧ�ܡ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    On Error Resume Next
    
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim t_pp As TYPE_PATH_Pati, arrtmp As Variant
    Dim lng�������� As Long, lng��׼����ʱ�� As Long
    Dim rsTmp As ADODB.Recordset, strδ������� As String, strδ�������� As String
    Dim lngB As Long, lngE As Long, strUnit As String, strTmp As String, DatCur As Date, lngValue As Long
    Dim i As Long, strFilter As String
    
    If vsPath.Row <= 0 Then
        MsgBox "��ѡ��һ�������ڸò��˵������ٴ�·����", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    arrtmp = Split(vsPath.RowData(vsPath.Row), ":")
    t_pp.·��ID = arrtmp(0)
    t_pp.�汾�� = arrtmp(1)
    mrsPath.Filter = "ID=" & t_pp.·��ID

    Set rsTmp = GetUnImportReson

    If Val(mrsPath!�����Ա�) <> 0 Then
        If Val(mrsPath!�����Ա�) <> IIf(mrsPati!�Ա� = "��", 1, IIf(mrsPati!�Ա� = "Ů", 2, 0)) Then
            MsgBox "��·�����ʺ��ڸò����Ա�[" & mrsPati!�Ա� & "]", vbInformation, gstrSysName
            strδ�������� = "�Ա��ʺ�"
            GoTo UnImport
        End If
    End If

    If Not IsNull(mrsPath!��������) And Not IsNull(mrsPati!����) Then
        lngValue = 0
        lngB = Split(mrsPath!��������, "-")(0)
        strTmp = Split(mrsPath!��������, "-")(1)
        lngE = Mid(strTmp, 1, Len(strTmp) - 1)
        strUnit = Mid(strTmp, Len(strTmp))

        strTmp = mrsPati!����           '���⣺2��3�µ�
        If strUnit = Mid(strTmp, Len(strTmp)) And IsNumeric(Mid(strTmp, 1, Len(strTmp) - 1)) Or IsNumeric(strTmp) Then
            '��ͬ���䵥λ�����Ƚ�
            lngValue = Val(strTmp)
        ElseIf Not IsNull(mrsPati!��������) Then
            DatCur = zlDatabase.Currentdate
            lngValue = DateDiff(IIf(strUnit = "��", "yyyy", IIf(strUnit = "��", "m", "d")), CDate(mrsPati!��������), DatCur)
            If lngValue = 0 Then lngValue = 1
        End If
        If lngValue <> 0 Then
            If lngValue < lngB Or lngValue > lngE Then
                MsgBox "��·�����ʺ��ڸò�������[" & mrsPati!���� & "]", vbInformation, gstrSysName
                strδ�������� = "���䲻�ʺ�"
                GoTo UnImport
            End If
        End If
    End If

    Me.Hide
    mblnOK = frmEvaluateOut.ShowMe(mfrmParent, 0, 1, mPati, t_pp, mrsPath!����, mlngDiagnosisType, mlngDiagnosisSorce, mlng����ID, mlng���ID)
    
    If cmdOK.Tag <> "Unload" Then Unload Me
    Exit Sub
UnImport:
    '��Ҫ��ϲű���δ����ԭ��
    rsTmp.Filter = "����='" & strδ�������� & "'"
    If rsTmp.RecordCount = 0 Then
        strδ������� = ""
    Else
        strδ������� = rsTmp!����
    End If

    Call SaveUnImport(mPati, t_pp, strδ�������, strδ��������)
    mblnOK = True
    If cmdOK.Tag <> "Unload" Then
        Unload Me
    End If
End Sub

Private Sub SaveUnImport(mPati As TYPE_Pati, mPP As TYPE_PATH_Pati, strVariationCode As String, strVariationTitle As String)
'���ܣ�����δ����ԭ��
'������strVariationCode=δ����ԭ�����,strVariationTitle=δ����ԭ������
    Dim strSql As String, strID As String, DateInPath As Date
    Dim str���ϵ��� As String
    
    On Error GoTo errH
    
    strID = zlDatabase.GetNextId("��������·��")
    If CheckOutPathSend(mPati.�Һ�ID) Then
        DateInPath = zlDatabase.Currentdate
    Else
        DateInPath = GetPatiInDateOut(mPati)
    End If
    str���ϵ��� = "0"
    
    strSql = "Zl_��������·������_Insert(" & mPati.����ID & "," & mPati.�Һ�ID & "," & mPati.����ID & "," & _
            mPP.·��ID & "," & mPP.�汾�� & "," & strID & ",'" & UserInfo.���� & "','" & strVariationTitle & "'," & _
            str���ϵ��� & ",To_Date('" & Format(DateInPath, "yyyy-MM-DD HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),'" & _
            strVariationCode & "'," & mlngDiagnosisType & "," & mlngDiagnosisSorce & "," & IIf(mlng����ID = 0, "NULL", mlng����ID) & "," & IIf(mlng���ID = 0, "NULL", mlng���ID) & ",Null,1)"
    
    Call zlDatabase.ExecuteProcedure(strSql, "δ����ԭ��")
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetUnImportReson() As ADODB.Recordset
'���ܣ���ȡ�̶��ļ���δ����ԭ��
    Dim strSql As String
 
    strSql = "Select ����, ���� From ������쳣��ԭ�� Where ���� = 0 And ĩ�� = 1"
    On Error GoTo errH
    Set GetUnImportReson = zlDatabase.OpenSQLRecord(strSql, "δ����ԭ��")

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Activate()
    If cmdOK.Tag = "Unload" Then
        cmdOK.Tag = ""
        Unload Me
    Else
        If vsPath.Rows = vsPath.FixedRows + 1 Then
            vsPath.Row = vsPath.Rows - 1
            vsPath.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long

    lblPait.Caption = "��ǰ���ˣ�" & mrsPati!���� & "," & mrsPati!�Ա� & "," & mrsPati!����

    If mrsPath.RecordCount > 0 Then mrsPath.MoveFirst

    With vsPath
        .Rows = .FixedRows
        For i = 1 To mrsPath.RecordCount
            'ȱʡ��ѡ���κ�һ��
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = mrsPath!����
            .TextMatrix(.Rows - 1, 2) = mrsPath!����
            .TextMatrix(.Rows - 1, 3) = "" & mrsPath!˵��
            .RowData(.Rows - 1) = mrsPath!ID & ":" & mrsPath!���°汾
            mrsPath.MoveNext
        Next
        
        cmdCancel.Visible = False
        cmdOK.Left = cmdCancel.Left
        
        'ֻ��һ��·����
        If .Rows = .FixedRows + 1 Then
            .Row = .Rows - 1
            cmdOK.Tag = "Unload"
            cmdOK_Click
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '������Ҫ���ʱ������ȡ����ť��ֻ�ܵ�ȷ��
    If mblnOK = False Then
        Cancel = 1
    Else
        Set mrsPath = Nothing
        Set mrsPati = Nothing
        Set mfrmParent = Nothing
        Set mrsDiag = Nothing
        mlng����ID = 0
        mlng���ID = 0
    End If
End Sub

Private Sub vsPath_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub

Private Sub vsPath_DblClick()
    Call cmdOK_Click
End Sub
