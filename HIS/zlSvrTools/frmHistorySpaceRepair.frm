VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmHistorySpaceRepair 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ʷ��ṹ�޸�"
   ClientHeight    =   6780
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   10710
   Icon            =   "frmHistorySpaceRepair.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdRepair 
      Caption         =   "�޸�(&R)"
      Height          =   350
      Left            =   8280
      TabIndex        =   5
      Top             =   6000
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�(&X)"
      Height          =   350
      Left            =   9480
      TabIndex        =   4
      Top             =   6000
      Width           =   1100
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      TabIndex        =   0
      Top             =   6408
      Width           =   10704
      _ExtentX        =   18891
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmHistorySpaceRepair.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15319
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "16:31"
            Key             =   "STANUM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraCheck 
      Height          =   5985
      Left            =   0
      TabIndex        =   1
      Top             =   -60
      Width           =   10680
      Begin VSFlex8Ctl.VSFlexGrid vsCheckResult 
         Height          =   5100
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   10395
         _cx             =   18336
         _cy             =   8996
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
         BackColorSel    =   16764057
         ForeColorSel    =   0
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483628
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   100
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmHistorySpaceRepair.frx":0E1C
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
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
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Frame fraTop 
         Height          =   120
         Left            =   15
         TabIndex        =   2
         Top             =   570
         Width           =   10680
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "����ת����Ķ��壬������߿�����ʷ��Ľṹһ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   3
         Top             =   225
         Width           =   5400
      End
   End
   Begin ComctlLib.ImageList ist 
      Left            =   120
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmHistorySpaceRepair.frx":0F57
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmHistorySpaceRepair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngSys  As Long                   '��ǰϵͳ��ϵͳ���
Private mstrVersion   As String            '��ǰϵͳ�İ汾��
Private mstrBakOwnerName As String         '��ǰϵͳ���߷�ֻ������ʷ��ռ��������
Private mstrOwnerName As String

Private mcnBakDB As New ADODB.Connection '��ʷ��ռ������߽���������

Private marrBakAddSQL() As Variant '�Ժ󱸿����Ա��¼ִ�е�SQL���������߿����Ա�ȵ�¼ִ�е�SQL��ִ��
Private marrOnlineAddSQL() As Variant '���߿����Ա�ȵ�¼ִ�е�SQL

Private mblnSucced As Boolean '�޸��Ƿ�ɹ�
Private mblnUpdate As Boolean '�Ƿ��������������
Private mblnFirstAct As Boolean '�����Ƿ��״μ���
Private mblnAllRepair As Boolean '�Ƿ��޸��ɹ�
Private mblnCurDB As Boolean '�Ƿ��ǵ�ǰϵͳ�ĵ�ǰ��ʷ��
Private mrsErrInfo As ADODB.Recordset

Private mrsSQL As ADODB.Recordset
Private mlngIndex As Long '��¼SQL˳��
Private mstrBakDB           As String   '��ʷ��ռ�
Private mstrBakIndexDB      As String   '��ʷ��������ռ�
Private mstrBakLobDB        As String   '��ʷ��LOB��ռ�
Private mstrDBLink  As String   '����@����
Private mstrServer As String

Private Enum RepCols
    RC_DifInfo = 0
    RC_DifType = 1
    RC_TabName = 2
    RC_ObjName = 3
    RC_ColName = 4
    RC_ObjType = 5
    RC_ObjLen = 6
    RC_ObjScale = 7
    RC_AutoRep = 8
    RC_RepSQL = 9
    RC_RepMethod = 10
    
End Enum

Private Enum DifType
    DT_HLackTab = 0 '��ʷ��ȱʧ
    DT_HMoreCol = 1 '��ʷ���һ��
    DT_HLessCol = 2 '��ʷ����һ��
    DT_HDataTypeDif = 3 '�������Ͳ�ͬ
    DT_HRepLenDif = 4 '���޸��г��Ȼ򾫶Ȳ���
    DT_HNotRepLenDif = 5 '�����޸����г��Ȼ򾫶Ȳ���
    DT_HLobTablespace = 6 'LOB�ֶα�ı�ռ����
    DT_HIndUsable = 7 '��ʷ��ʧЧ������
    DT_HIndDel = 8 '��ʷ����������
    DT_HIndAdd = 9 '��ʷ��ȱ�ٵ�����
    DT_HIndColDif = 10 '��ʷ�������в���
    DT_HConDisable = 11 '��ʷ����õ�Լ��
    DT_URefConDel = 12 '�ӱ�����δת��
    DT_HConDel = 13 '��ʷ������Լ��
    DT_HConAdd = 14 '��ʷ��ȱ�ٵ�Լ��
    DT_HConColDIf = 15 '��ʷ��Լ���в���
    DT_HIndexTablesapce = 16 '������ռ����
End Enum

Public Function ShowRepair(ByVal frmMain As Form, ByVal lngϵͳ As Long, ByVal blnUpdate As Boolean, Optional ByVal strBakUser As String, Optional ByVal strBakDB As String, Optional ByVal blnCurDB As Boolean = True, Optional ByRef rsRepairSQL As ADODB.Recordset, Optional cnDBBAK As ADODB.Connection, Optional strDbLink As String) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------
    '����:�޸���ʷ��ռ�����ݽṹ
    '����:cnOracle-ϵͳ����
    '     strOwner-�������û���
    '     lngϵͳ-ϵͳ��
    '     blnUpdate -����ʱ�Ľṹ�޸�
    '     strDBLink=DBLInk����
    '����:��װ�ɹ�,����true,���򷵻�False
    '----------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    mlngSys = lngϵͳ

    'ϵͳȨ�޿���
    strSQL = "select ������,�汾��,���� from zlSystems where ���=" & mlngSys
    Call OpenRecordset(rsTemp, strSQL, "��ȡ������")
    
    If Not rsTemp.EOF Then
        mstrOwnerName = Nvl(rsTemp!������)
        mstrVersion = Nvl(rsTemp!�汾��)
    Else
        If Not blnUpdate Then MsgBox "ϵͳ������,���ܱ����˲�ж,���ܼ���!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    
    If gstrUserName <> mstrOwnerName Then
        If Not blnUpdate Then MsgBox "�㲻�ǵ�ǰӦ�ó����������,���ܼ���!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    
    If mstrVersion <> "" Then
        If Val(Split(mstrVersion, ".")(0)) < 10 Then
                If Not blnUpdate Then MsgBox "��֧��9���µİ汾,���ܼ���!", vbInformation + vbDefaultButton1, gstrSysName
                Exit Function
        End If
    End If
    mblnCurDB = False
    mstrServer = gstrServer
    If Not blnUpdate Then
        '��ʷ��ռ������ߣ��Լ���ռ�����
        strSQL = "Select ����,������,DB���� From Zltools.Zlbakspaces Where ϵͳ = " & mlngSys & "  And ��ǰ = 1 And ֻ�� = 0"
        Call OpenRecordset(rsTemp, strSQL, "��ȡ��ʷ��ռ�������")
        If rsTemp.EOF Then
            MsgBox "��ǰû�п��õ���ʷ���ݿռ������ʷ���ݿռ�Ŀǰ��״̬Ϊֻ��,���ܼ���!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        Else
            mstrBakOwnerName = Nvl(rsTemp!������)
            mblnCurDB = True
            mstrDBLink = Nvl(rsTemp!DB����)
            mstrBakDB = Nvl(rsTemp!����)
        End If
    Else
        mstrBakOwnerName = strBakUser
        mblnCurDB = blnCurDB
        mstrDBLink = strDbLink
        mstrBakDB = strBakDB
    End If
    mstrBakDB = UCase(mstrBakDB)
    If mstrDBLink <> "" Then
        strSQL = "Select Owner, Db_Link, Username, Host" & vbNewLine & _
                    "From All_Db_Links" & vbNewLine & _
                    "Where Owner =[1] And Username =[2] And Db_Link||'.' Like [3]"
        Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡDBLink������", gstrUserName, UCase(mstrBakOwnerName), UCase(mstrDBLink) & ".%")
        If Not rsTemp.EOF Then mstrServer = rsTemp!Host & ""
    End If
    mstrDBLink = IIf(mstrDBLink = "", "", "@") & mstrDBLink
    '��ȡ������ռ���LOB��ռ�
    strSQL = "Select a.Name From V$tablespace" & mstrDBLink & " a Where a.Name Like '" & mstrBakDB & "_%'"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡ��ʷ���ռ�")
    rsTemp.Filter = "Name='" & mstrBakDB & "_IDX'"
    If Not rsTemp.EOF Then
        mstrBakIndexDB = rsTemp!name
    Else
        mstrBakIndexDB = mstrBakDB
    End If
    rsTemp.Filter = "Name='" & mstrBakDB & "_LOB'"
    If Not rsTemp.EOF Then
        mstrBakLobDB = rsTemp!name
    Else
        mstrBakLobDB = mstrBakDB
    End If
    mblnFirstAct = True
    mblnAllRepair = True
    mblnUpdate = blnUpdate
    If blnUpdate Then
        Set mcnBakDB = cnDBBAK
        Set mrsSQL = rsRepairSQL
        If mrsSQL Is Nothing Then
            Set mrsSQL = GetIniRec
        End If
        Call LoadCheckData
    End If
    On Error Resume Next
    If Not blnUpdate Then
        Me.Show 1
        On Error GoTo 0
        ShowRepair = mblnSucced
    Else
        Set rsRepairSQL = mrsSQL
        Exit Function
    End If
    On Error GoTo 0
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdRepair_Click()
    Dim i As Long
    Dim comTmp As New ADODB.Command
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim lngCount As Long

    '��ʷ��ռ���������֤
    If Not frmUserCheckLogin.ShowLogin(UCT_CurZLBAK, mcnBakDB, mstrBakOwnerName, mstrServer, mlngSys) Then Exit Sub
    

    lngCount = 5
    
    mblnAllRepair = False
    
    Call SetFaceCtlEnable
    If mcnBakDB Is Nothing Then Exit Sub
    
    On Error Resume Next
    SetPromptText ("��1/" & lngCount & ")���ڽ����߿�����ر���Ȩ����ʷ���û�")
    comTmp.CommandType = adCmdText
    Set comTmp.ActiveConnection = gcnOracle
    For i = LBound(marrOnlineAddSQL) To UBound(marrOnlineAddSQL)
        comTmp.CommandText = marrOnlineAddSQL(i)
        comTmp.Execute
        If err <> 0 Then
            Call AddErrIntoRs(0, err.Description, , , marrOnlineAddSQL(i))
            err.Clear
        End If
    Next
    
    SetPromptText ("��2/" & lngCount & ")��ʼ�޸���ʷ�����ݽṹ")
    With vsCheckResult
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, RC_RepSQL) <> "" Then
                Select Case Val(.TextMatrix(i, RC_DifType))
                    Case DT_URefConDel
                        Set comTmp.ActiveConnection = gcnOracle
                    Case DT_HConAdd
                        '��������Ƿ���ڣ����ڣ���ɾ��
                        strSQL = "Select /*+rule*/" & vbNewLine & _
                                    " 1" & vbNewLine & _
                                    "From User_Indexes A" & vbNewLine & _
                                    "Where A.Index_Name = '" & .TextMatrix(i, RC_ObjName) & "'"
                         Call OpenRecordset(rsTmp, strSQL, "����ת����ع�����Ч���", , , mcnBakDB)
                         Set comTmp.ActiveConnection = mcnBakDB
                         If Not rsTmp.EOF Then
                            comTmp.CommandText = " Drop Index  " & .TextMatrix(i, RC_ObjName)
                            comTmp.Execute
                            If err <> 0 Then
                                Call AddErrIntoRs(IIf(Val(.TextMatrix(i, RC_DifType)) = DT_URefConDel, 0, 1), err.Description, .TextMatrix(i, RC_TabName), .TextMatrix(i, RC_ObjName), " Drop Index  " & .TextMatrix(i, RC_ObjName))
                                err.Clear
                            End If
                         End If
                    Case Else
                        Set comTmp.ActiveConnection = mcnBakDB
                End Select
                
                comTmp.CommandText = .TextMatrix(i, RC_RepSQL)
                comTmp.Execute
                If err <> 0 Then
                    Call AddErrIntoRs(IIf(Val(.TextMatrix(i, RC_DifType)) = DT_URefConDel, 0, 1), err.Description, .TextMatrix(i, RC_TabName), .TextMatrix(i, RC_ObjName), .TextMatrix(i, RC_RepSQL))
                    err.Clear
                End If
            End If
        Next
    End With
    SetPromptText ("��3/" & lngCount & ")��ʷ����Ȩ���߿�������")
    Set comTmp.ActiveConnection = mcnBakDB
    For i = LBound(marrBakAddSQL) To UBound(marrBakAddSQL)
        comTmp.CommandText = marrBakAddSQL(i)
        comTmp.Execute
        If err <> 0 Then
            Call AddErrIntoRs(0, err.Description, , , marrBakAddSQL(i))
            err.Clear
        End If
    Next
    
    SetPromptText ("��4/" & lngCount & ")��ʼ���´�����ʷ��H��ͼ����Ȩ")
    If mstrDBLink = "" Then
        Set comTmp.ActiveConnection = gcnOracle
        Call GrantBakToUser(mcnBakDB, mstrOwnerName)
    End If
    If mblnCurDB Then
        Call CreateAppView(mstrOwnerName, mstrBakOwnerName, mlngSys, mstrDBLink)
    End If
    SetPromptText ("��5/" & lngCount & ")����ת����ع�����Ч������ر���")
    Set comTmp.ActiveConnection = gcnOracle
    strSQL = "Select 'Alter Procedure Zl" & mlngSys \ 100 & "_Datamove_Tag compile' As Sql" & vbNewLine & _
            "From User_Objects A" & vbNewLine & _
            "Where a.Object_Name = Upper('Zl" & mlngSys \ 100 & "_Datamove_Tag') And a.Object_Type = 'PROCEDURE' And a.Status = 'INVALID'" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select 'Alter Procedure ' || a.Object_Name || ' compile' As Sql" & vbNewLine & _
            "From User_Objects A" & vbNewLine & _
            "Where a.Object_Name In (Upper('Zl" & mlngSys \ 100 & "_Datamoveout1'), Upper('Zl_Retu_Clinic'), Upper('Zl_Retu_Exes')) And" & vbNewLine & _
            "      a.Object_Type = 'PROCEDURE' And a.Status = 'INVALID'"
    Call OpenRecordset(rsTmp, strSQL, "����ת����ع�����Ч���")
    While Not rsTmp.EOF
        comTmp.CommandText = rsTmp!SQL & ""
        comTmp.Execute
        If err <> 0 Then
            Call AddErrIntoRs(0, err.Description, , , strSQL)
            err.Clear
        End If
        rsTmp.MoveNext
    Wend
    
    SetPromptText ("�޸����")
    Call LoadErrInfo(mrsErrInfo)
    mblnAllRepair = True
    Call SetFaceCtlEnable
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If mblnFirstAct Then
        mblnFirstAct = False
        Me.Refresh
        If Not LoadCheckData Then
            SetPromptText ("������,��ʷ��δ���ֽṹ����")
        Else
            SetPromptText ("������,���޸�")
        End If
        If mrsErrInfo Is Nothing Then Set mrsErrInfo = New ADODB.Recordset
        With mrsErrInfo
            .Fields.Append "���ݿ�", adInteger
            .Fields.Append "��������", adInteger
            .Fields.Append "������Ϣ", adVarChar, 100
            .Fields.Append "����", adVarChar, 50
            .Fields.Append "������", adVarChar, 50
            .Fields.Append "����SQL", adVarChar, 200
            .Open
        End With
    End If
    
End Sub

Private Sub Form_Load()
    On Error Resume Next
    marrBakAddSQL = Array()
    marrOnlineAddSQL = Array()
    Call ApplyOEM(stbThis)
End Sub

Private Function LoadCheckData() As Boolean

    '----------------------------------------------------------------------------------------------------------------------------------
    '����:������ʷ��ռ����ݽṹ�����
    '����:���ɹ�,����true,���򷵻�False
    '----------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim strLackTable    As String
    Dim strLobTable     As String
    Dim lngTotal        As Long
    Dim lngCur          As Long
    
    If Not mblnUpdate Then
        vsCheckResult.Redraw = False
    End If
    On Error GoTo errH:
    lngTotal = 15
    '��һ���� ��ʷ��ȱʧ
    lngCur = lngCur + 1
    SetPromptText ("��" & lngCur & "/" & lngTotal & ")���ڼ����ʷ��")
    strSQL = "Select '��ʷ��ȱʧ' ������Ϣ, " & DT_HLackTab & " ��������, t.����, Null ������, Null ����, Null ��������, Null ����, Null ����, '��' �Զ��޸�, '�����˱�' �޸�˵��" & vbNewLine & _
            "From Zltools.Zlbaktables t, (Select Table_Name From All_Tables" & mstrDBLink & " a Where a.Owner = '" & UCase(mstrBakOwnerName) & "') b" & vbNewLine & _
            "Where t.ϵͳ = " & mlngSys & " And b.Table_Name(+) = t.���� And b.Table_Name Is Null"
    
    Call OpenRecordset(rsTmp, strSQL, "��ʷ��ȱʧ��")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp, True)
    End If
    If rsTmp.RecordCount <> 0 Then rsTmp.MoveFirst
    While Not rsTmp.EOF
        strLackTable = strLackTable & ",'" & UCase(rsTmp!���� & "") & "'"
        rsTmp.MoveNext
    Wend
    If Len(strLackTable) <> 0 Then strLackTable = Mid(strLackTable, 2)
    '(����) �м�飺 1-��ʷ��ȱ���У������޸� 2-��ʷ���ж����У���Ӹ���) �������ų��ˡ����ȱʧ��
    lngCur = lngCur + 1
    SetPromptText ("��" & lngCur & "/" & lngTotal & "/" & lngTotal & ")���ڼ����ʷ�����")
    '(1)�ֲ�����ʷ��ռ���ڵı�
    strSQL = IIf(Len(strLackTable) = 0, "", " And T.���� Not In (" & strLackTable & ")")
    '(2)������ʷ��ռ䣬ȱ�ٻ�������
    strSQL = "Select  Decode(a.Column_Name, Null, '��ʷ���ж�����', '��ʷ��ȱ����') ������Ϣ, Decode(B.Column_Name, Null, " & DT_HLessCol & ", " & DT_HMoreCol & ") ��������," & vbNewLine & _
            "       Nvl(b.Table_Name, a.Table_Name) ����, Nvl(b.Column_Name, a.Column_Name) ������, Nvl(b.Column_Name, a.Column_Name) ����," & vbNewLine & _
            "       Nvl(b.Data_Type, a.Data_Type) ��������," & vbNewLine & _
            "       Decode(Nvl(b.Data_Type, a.Data_Type),  'XMLTYPE',Null,'DATE',Null,'Long Raw',Null,'BLOB',Null,'CLOB',Null," & vbNewLine & _
            "       Nvl(Nvl(b.Data_Precision, a.Data_Precision),Nvl(b.Data_Length, a.Data_Length))) ����, Nvl(b.Data_Scale, a.Data_Scale) ����," & vbNewLine & _
            "       Decode(a.Column_Name, Null, '��', '��') �Զ��޸�," & vbNewLine & _
            "       Decode(a.Column_Name, Null, '�����޸�', '��������') �޸�˵��" & vbNewLine & _
            "From (Select c.Table_Name,c.Column_Name,c.Data_Type,c.Data_Precision,c.Data_Scale,c.Data_Length From User_Tab_Columns C, Zltools.Zlbaktables T Where c.Table_Name = t.���� And t.ϵͳ = " & mlngSys & strSQL & ") A" & vbNewLine & _
            "Full Join (Select d.Table_Name,d.Column_Name,d.Data_Type,d.Data_Precision,d.Data_Scale,d.Data_Length " & vbNewLine & _
            "           From All_Tab_Columns" & mstrDBLink & " D, Zltools.Zlbaktables T" & vbNewLine & _
            "           Where d.Table_Name = t.���� And t.ϵͳ = " & mlngSys & strSQL & "  And d.Owner = '" & UCase(mstrBakOwnerName) & "') B" & vbNewLine & _
            "On a.Table_Name = b.Table_Name And a.Column_Name = b.Column_Name" & vbNewLine & _
            "Where a.Column_Name Is Null Or b.Column_Name Is Null"

    Call OpenRecordset(rsTmp, strSQL, "��ʷ���м��")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp)
    End If
    
    '(�����������Ͳ���
    lngCur = lngCur + 1
    SetPromptText ("��" & lngCur & "/" & lngTotal & ")���ڼ����ʷ���������")
    strSQL = "Select '�����Ͳ���' ������Ϣ, " & DT_HDataTypeDif & " ��������, b.Table_Name ����, b.Column_Name ������, b.Column_Name ����, b.Data_Type ��������," & vbNewLine & _
                    "       Decode(b.Data_Type,'XMLTYPE',Null,'DATE',Null,'Long Raw',Null,'BLOB',Null,'CLOB',Null,Nvl(b.Data_Precision, b.Data_Length)) ����," & vbNewLine & _
                    "        '��' �Զ��޸�, '�޸�˵������' �޸�˵��" & vbNewLine & _
                    "From User_Tab_Columns a, All_Tab_Columns" & mstrDBLink & " b, (Select t.���� From Zltools.Zlbaktables t Where t.ϵͳ = " & mlngSys & ") c" & vbNewLine & _
                    "Where a.Table_Name = b.Table_Name And b.Table_Name = c.���� And b.Owner = '" & UCase(mstrBakOwnerName) & "' And a.Column_Name = b.Column_Name And" & vbNewLine & _
                    "      a.Data_Type <> b.Data_Type "
    Call OpenRecordset(rsTmp, strSQL, "�����Ͳ�ͬ����")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp)
    End If
    '(�ġ�) �еľ����Լ����Ȳ���(4-���޸���5-�����޸��� �������͵���Ϣ���������жϵó�
    lngCur = lngCur + 1
    SetPromptText ("��" & lngCur & "/" & lngTotal & ")���ڼ����ʷ����о���")
    strSQL = "Select Null ������Ϣ, -1 ��������, ����, ������, ������ ����, ��������, a.�󱸳���, a.�󱸾���, a.���߳���, a.���߾���, Null �Զ��޸�, Null �޸�˵��" & vbNewLine & _
            "From (Select a.Table_Name ����, a.Column_Name ������, a.Data_Type ��������," & vbNewLine & _
            "              Decode(a.Data_Type, 'XMLTYPE',Null,'DATE',Null,'Long Raw',Null, 'BLOB',Null,'CLOB',Null, Nvl(a.Data_Precision,a.Data_Length)) ���߳���, a.Data_Scale ���߾���," & vbNewLine & _
            "              Decode(b.Data_Type, 'XMLTYPE',Null,'DATE',Null,'Long Raw',Null, 'BLOB',Null,'CLOB',Null,Nvl(b.Data_Precision,b.Data_Length)) �󱸳���, b.Data_Scale �󱸾���" & vbNewLine & _
            "       From User_Tab_Columns A, All_Tab_Columns" & mstrDBLink & " B" & vbNewLine & _
            "       Where a.Table_Name = b.Table_Name And b.Owner = '" & UCase(mstrBakOwnerName) & "' And a.Column_Name = b.Column_Name And Exists" & vbNewLine & _
            "        (Select 1 From Zltools.Zlbaktables T Where t.ϵͳ = " & mlngSys & " And t.���� = a.Table_Name) And a.Data_Type = b.Data_Type) A" & vbNewLine & _
            "Where ���߳��� <> �󱸳��� Or ���߾��� <> �󱸾���"


    Call OpenRecordset(rsTmp, strSQL, "�еľ��Ȳ���")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp)
    End If
    
    '(S)��ʷ��LOB��ռ��飬�����������Լ��֮ǰ��飬��Ϊ�������ܻᵼ������Լ��ʧЧ
    lngCur = lngCur + 1
    SetPromptText ("��" & lngCur & "/" & lngTotal & ")���ڼ����ʷ��LOB��ռ�")
    strSQL = "Select '��ʷ����ռ����' ������Ϣ, " & DT_HLobTablespace & " ��������, a.Table_Name ����, Null ������, Null ����, Null ��������, Null ����, Null ����, '��' �Զ��޸�," & vbNewLine & _
                    "       '�����ֹ����ñ��ƶ�����ռ�" & mstrBakLobDB & "' �޸�˵��" & vbNewLine & _
                    "From All_Tables" & mstrDBLink & " a" & vbNewLine & _
                    "Where a.Owner = '" & UCase(mstrBakOwnerName) & "'" & vbNewLine & _
                    "And a.Table_Name In (Select Distinct c.Table_Name" & vbNewLine & _
                    "                    From User_Tab_Cols c, Zltools.Zlbaktables t" & vbNewLine & _
                    "                    Where c.Table_Name = t.����" & vbNewLine & _
                    "                    And t.ϵͳ = " & mlngSys & vbNewLine & _
                    "                    And c.Data_Type In ('BLOB', 'CLOB', 'BFILE', 'XMLTYPE'))" & vbNewLine & _
                    "And a.Tablespace_Name Not in( '" & mstrBakLobDB & "'" & IIf(mstrBakLobDB = mstrBakDB, ",'" & mstrBakLobDB & "_LOB')", ")")
    Call OpenRecordset(rsTmp, strSQL, "��ʷ��Լ���в���")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp)
    End If
    If rsTmp.RecordCount <> 0 Then rsTmp.MoveFirst
    While Not rsTmp.EOF
        strLobTable = strLobTable & ",'" & UCase(rsTmp!���� & "") & "'"
        rsTmp.MoveNext
    Wend
    strLobTable = Mid(strLobTable, 2)
    '(�����޸�����
    '��1������/Ψһ�����ǿ�������������Ч�Լ��
    lngCur = lngCur + 1
    SetPromptText ("��" & lngCur & "/" & lngTotal & ")���ڼ����ʷ����������Ч��")
    strSQL = "Select '��ʷ����ʧЧ����' ������Ϣ, " & DT_HIndUsable & "  ��������, a.Table_Name ����, a.Index_Name ������, a.Colstr ����, Null ��������, Null ����, Null ����, '��' �Զ��޸�," & vbNewLine & _
                    "       '�����ؽ�' �޸�˵��" & vbNewLine & _
                    "From (Select d.Table_Name, d.Index_Name," & vbNewLine & _
                    "              f_List2str(Cast(Collect(d.Column_Name Order By d.Column_Position) As t_Strlist)) Colstr" & vbNewLine & _
                    "       From User_Ind_Columns d, Zltools.Zlbaktables t" & vbNewLine & _
                    "       Where d.Table_Name = t.���� And t.ϵͳ = " & mlngSys & " And Instr(d.Index_Name, '_PK') = 0 And Instr(d.Index_Name, '_UQ') = 0 And Instr(d.Index_Name, '_IX_��ת��') = 0" & vbNewLine & _
                    "       Group By d.Table_Name, d.Index_Name) a," & vbNewLine & _
                    "     (Select e.Table_Name, f_List2str(Cast(Collect(e.Column_Name Order By e.Position) As t_Strlist)) Colstr" & vbNewLine & _
                    "       From User_Cons_Columns e, User_Constraints f, Zltools.Zlbaktables t, User_Constraints c," & vbNewLine & _
                    "            (Select ���� From Zltools.Zlbaktables" & vbNewLine & _
                    "              Union All Select '������ҳ' From Dual" & vbNewLine & _
                    "              Union All Select '������Ϣ' From Dual) g" & vbNewLine & _
                    "       Where e.Table_Name = t.���� And t.ϵͳ = " & mlngSys & " And e.Constraint_Name = f.Constraint_Name And f.Constraint_Type = 'R' And" & vbNewLine & _
                    "             c.Constraint_Name = f.r_Constraint_Name And g.����(+) = c.Table_Name And g.���� Is Null" & vbNewLine & _
                    "       Group By e.Table_Name, e.Constraint_Name) b," & vbNewLine & _
                    "     (Select Index_Name" & vbNewLine & _
                    "       From All_Indexes" & mstrDBLink & " k, Zltools.Zlbaktables t" & vbNewLine & _
                    "       Where k.Table_Owner = '" & UCase(mstrBakOwnerName) & "' And k.Status = 'UNUSABLE' And k.Index_Type <> 'LOB' And k.Table_Name = t.���� And" & vbNewLine & _
                    "             t.ϵͳ = " & mlngSys & ") h" & vbNewLine & _
                    "Where a.Table_Name = b.Table_Name(+) And a.Colstr = b.Colstr(+) And b.Colstr Is Null And a.Index_Name = h.Index_Name"

    Call OpenRecordset(rsTmp, strSQL, "��ʷ����ʧЧ����")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp)
    End If
    
    '��2����ʷ��ռ�����ϵ�������ɾ������������õı����ʷ��ռ���ڵı�����ɾ��)
    lngCur = lngCur + 1
    SetPromptText ("��" & lngCur & "/" & lngTotal & ")���ڼ����ʷ����������")
    
    strSQL = "Select '��ʷ����������' ������Ϣ, " & DT_HIndDel & " ��������, a.Table_Name ����, a.Index_Name ������, a.Colstr ����, Null ��������, Null ����, Null ����, '��' �Զ��޸�," & vbNewLine & _
                "       'ɾ������' �޸�˵��" & vbNewLine & _
                "From (Select d.Table_Name, d.Index_Name," & vbNewLine & _
                "              f_List2str(Cast(Collect(d.Column_Name Order By d.Column_Position) As t_Strlist)) Colstr" & vbNewLine & _
                "       From User_Ind_Columns d, Zltools.Zlbaktables t" & vbNewLine & _
                "       Where d.Table_Name = t.���� And t.ϵͳ = " & mlngSys & " And Instr(d.Index_Name, '_PK') = 0 And Instr(d.Index_Name, '_UQ') = 0" & vbNewLine & _
                "       Group By d.Table_Name, d.Index_Name) a," & vbNewLine & _
                "     (Select e.Table_Name, f_List2str(Cast(Collect(e.Column_Name Order By e.Position) As t_Strlist)) Colstr" & vbNewLine & _
                "       From User_Cons_Columns e, User_Constraints f, Zltools.Zlbaktables t, User_Constraints c," & vbNewLine & _
                "            (Select ���� From Zltools.Zlbaktables" & vbNewLine & _
                "              Union All Select '������ҳ' From Dual" & vbNewLine & _
                "              Union All Select '������Ϣ' From Dual) g" & vbNewLine & _
                "       Where e.Table_Name = t.���� And t.ϵͳ = " & mlngSys & " And e.Constraint_Name = f.Constraint_Name And f.Constraint_Type = 'R' And" & vbNewLine & _
                "             c.Constraint_Name = f.r_Constraint_Name And g.����(+) = c.Table_Name And g.���� Is Null" & vbNewLine & _
                "       Group By e.Table_Name, e.Constraint_Name) b," & vbNewLine & _
                "     (Select Index_Name" & vbNewLine & _
                "       From All_Indexes" & mstrDBLink & " k, Zltools.Zlbaktables t" & vbNewLine & _
                "       Where k.Table_Owner = '" & UCase(mstrBakOwnerName) & "'  And k.Table_Name = t.���� And t.ϵͳ = " & mlngSys & ") h" & vbNewLine & _
                "Where a.Table_Name = b.Table_Name And a.Colstr = b.Colstr And h.Index_Name = a.Index_Name" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select '��ʷ����������' ������Ϣ, " & DT_HIndDel & " ��������, k.Table_Name ����, k.Index_Name ������, '��ת��' ����, Null ��������, Null ����, Null ����, '��' �Զ��޸�," & vbNewLine & _
                "       'ɾ������' �޸�˵��" & vbNewLine & _
                "From All_Indexes" & mstrDBLink & " k, Zltools.Zlbaktables t" & vbNewLine & _
                "Where k.Table_Owner = '" & UCase(mstrBakOwnerName) & "'  And k.Table_Name = t.���� And t.ϵͳ = " & mlngSys & " And k.Index_Name Like '%_��ת��'"
    Call OpenRecordset(rsTmp, strSQL, "��ʷ����������")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp)
    End If
    '(3)��ʷ��ռ�����ϵ���������ӣ���������õı�����ʷ��ռ���ڵı��������)
    lngCur = lngCur + 1
    SetPromptText ("��" & lngCur & "/" & lngTotal & ")���ڼ����ʷ��ȱ�ٵ�����")
    strSQL = "Select  '��ʷ��ȱ�ٵ�����' ������Ϣ, " & DT_HIndAdd & " ��������, a.Table_Name ����, a.Index_Name ������, a.Colstr ����, Null ��������, Null ����, Null ����, '��' �Զ��޸�," & vbNewLine & _
                    "       '��������' �޸�˵��" & vbNewLine & _
                    "From (Select d.Table_Name, d.Index_Name,f_List2str(Cast(Collect(d.Column_Name Order By d.Column_Position) As t_Strlist)) Colstr" & vbNewLine & _
                    "       From User_Ind_Columns D, Zltools.Zlbaktables T" & vbNewLine & _
                    "       Where d.Table_Name = t.���� And t.ϵͳ = " & mlngSys & " And Instr(d.Index_Name, '_PK') = 0 And Instr(d.Index_Name, '_UQ') = 0 And Instr(d.Index_Name, '_IX_��ת��')=0" & vbNewLine & _
                    "       Group By d.Table_Name, d.Index_Name) A," & vbNewLine & _
                    "     (Select e.Table_Name, f_List2str(Cast(Collect(e.Column_Name Order By e.Position) As t_Strlist)) Colstr" & vbNewLine & _
                    "       From User_Cons_Columns e, User_Constraints f, Zltools.Zlbaktables t, User_Constraints c," & vbNewLine & _
                    "            (Select ���� From Zltools.Zlbaktables" & vbNewLine & _
                    "              Union All Select '������ҳ' From Dual" & vbNewLine & _
                    "              Union All Select '������Ϣ' From Dual) g" & vbNewLine & _
                    "       Where e.Table_Name = t.���� And t.ϵͳ = " & mlngSys & " And e.Constraint_Name = f.Constraint_Name And f.Constraint_Type = 'R' And" & vbNewLine & _
                    "             c.Constraint_Name = f.r_Constraint_Name And g.����(+) = c.Table_Name And g.���� Is Null" & vbNewLine & _
                    "       Group By e.Table_Name, e.Constraint_Name) b," & vbNewLine & _
                    "     (Select Index_Name" & vbNewLine & _
                    "       From All_Indexes" & mstrDBLink & " k, Zltools.Zlbaktables t" & vbNewLine & _
                    "       Where k.Table_Owner = '" & UCase(mstrBakOwnerName) & "' And k.Table_Name = t.���� And t.ϵͳ = " & mlngSys & ") h" & vbNewLine & _
                    "Where a.Table_Name = b.Table_Name(+) And a.Colstr = b.Colstr(+) And b.Colstr Is Null and a.Index_Name=h.index_name(+) and h.index_name is null"
    Call OpenRecordset(rsTmp, strSQL, "��ʷ��ȱ�ٵ�����")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp)
    End If
    '(4)��ʷ��ռ������з����仯
    lngCur = lngCur + 1
    SetPromptText ("��" & lngCur & "/" & lngTotal & ")���ڼ����ʷ�������еĲ���")
    strSQL = "Select '�����в���' ������Ϣ, " & DT_HIndColDif & " ��������, u.Table_Name ����, u.Index_Name ������, u.Colstr ����, Null ��������, Null ����, Null ����, '��' �Զ��޸�," & vbNewLine & _
                    "       '�ؽ�����' �޸�˵��" & vbNewLine & _
                    "From (Select a.Table_Name, a.Index_Name, f_List2str(Cast(Collect(a.Column_Name Order By a.Column_Position) As t_Strlist)) Colstr" & vbNewLine & _
                    "       From User_Ind_Columns A, Zltools.Zlbaktables T" & vbNewLine & _
                    "       Where a.Table_Name = t.���� And t.ϵͳ = " & mlngSys & " And Instr(a.Index_Name, '_PK') = 0 And Instr(a.Index_Name, '_UQ') = 0 And Instr(a.Index_Name, '_IX_��ת��')=0" & vbNewLine & _
                    "       Group By a.Table_Name, a.Index_Name) U," & vbNewLine & _
                    "     (Select a.Table_Name, a.Index_Name, f_List2str(Cast(Collect(a.Column_Name Order By a.Column_Position) As t_Strlist)) Colstr" & vbNewLine & _
                    "       From All_Ind_Columns" & mstrDBLink & " A, Zltools.Zlbaktables T" & vbNewLine & _
                    "       Where a.Table_Name = t.���� And t.ϵͳ = " & mlngSys & " And a.Table_Owner ='" & UCase(mstrBakOwnerName) & "'" & vbNewLine & _
                    "       Group By a.Table_Name, a.Index_Name) H," & vbNewLine & _
                    "     (Select e.Table_Name, f_List2str(Cast(Collect(e.Column_Name Order By e.Position) As t_Strlist)) Colstr" & vbNewLine & _
                    "       From User_Cons_Columns e, User_Constraints f, Zltools.Zlbaktables t, User_Constraints c," & vbNewLine & _
                    "            (Select ���� From Zltools.Zlbaktables" & vbNewLine & _
                    "              Union All Select '������ҳ' From Dual" & vbNewLine & _
                    "              Union All Select '������Ϣ' From Dual) g" & vbNewLine & _
                    "       Where e.Table_Name = t.���� And t.ϵͳ = " & mlngSys & " And e.Constraint_Name = f.Constraint_Name And f.Constraint_Type = 'R' And" & vbNewLine & _
                    "             c.Constraint_Name = f.r_Constraint_Name And g.����(+) = c.Table_Name And g.���� Is Null" & vbNewLine & _
                    "       Group By e.Table_Name, e.Constraint_Name) b" & vbNewLine & _
                    "Where u.Table_Name = h.Table_Name And u.Index_Name = h.Index_Name And u.Colstr <> h.Colstr And" & vbNewLine & _
                    "      u.Table_Name = b.Table_Name(+) And u.Colstr = b.Colstr(+) And b.Table_Name Is Null"
    Call OpenRecordset(rsTmp, strSQL, "�����в���")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp)
    End If
    '���塢���޸�Լ��
    '��1����/Ψһ��Լ������Ч�Լ��
    lngCur = lngCur + 1
    SetPromptText ("��" & lngCur & "/" & lngTotal & ")���ڼ����ʷ���н��õ�Լ��")
    strSQL = "Select '��ʷ���н��õ�Լ��' ������Ϣ, " & DT_HConDisable & " ��������, a.Table_Name ����, a.Constraint_Name ������," & vbNewLine & _
            "       f_List2str(Cast(Collect(b.Column_Name Order By b.Position) As t_Strlist)) ����, a.Constraint_Type ��������, Null ����, Null ����, '��' �Զ��޸�," & vbNewLine & _
            "       '����Լ��' �޸�˵��" & vbNewLine & _
            "From All_Constraints" & mstrDBLink & " A, All_Cons_Columns" & mstrDBLink & " B, Zltools.Zlbaktables T" & vbNewLine & _
            "Where a.Owner = '" & UCase(mstrBakOwnerName) & "'  And a.Owner = b.Owner And a.Constraint_Type In ('P', 'U') And a.Table_Name = t.���� And t.ϵͳ = " & mlngSys & " And" & vbNewLine & _
            "      a.Status = 'DISABLED' And Exists" & vbNewLine & _
            " (Select 1" & vbNewLine & _
            "       From User_Constraints C" & vbNewLine & _
            "       Where c.Constraint_Name = a.Constraint_Name And c.Constraint_Type = a.Constraint_Type) And a.Owner = b.Owner And" & vbNewLine & _
            "      a.Constraint_Name = b.Constraint_Name" & vbNewLine & _
            "Group By a.Table_Name, a.Constraint_Name, a.Constraint_Type"

    Call OpenRecordset(rsTmp, strSQL, "Լ����Ч��")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp)
    End If
    '(2)���߿���ɾ�������Լ��
    lngCur = lngCur + 1
    SetPromptText ("��" & lngCur & "/" & lngTotal & ")���ڼ�����߿�ת��������Լ��")
    strSQL = "Select  Distinct '�ӱ�����δת��' ������Ϣ, " & DT_URefConDel & " ��������, d.Table_Name ����, d.Constraint_Name ������," & vbNewLine & _
            "                f_List2str(Cast(Collect(e.Column_Name Order By e.Position) As t_Strlist)) ����, Null ��������, Null ����, Null ����, '��' �Զ��޸�," & vbNewLine & _
            "                '���޸�' �޸�˵��" & vbNewLine & _
            "From (Select Table_Name, Constraint_Name, Owner" & vbNewLine & _
            "       From (Select a.Owner, a.r_Constraint_Name, a.Constraint_Name, a.Table_Name, b.Table_Name r_Table_Name" & vbNewLine & _
            "              From User_Constraints A, User_Constraints B" & vbNewLine & _
            "              Where a.r_Constraint_Name = b.Constraint_Name(+)) C" & vbNewLine & _
            "       Start With c.r_Table_Name In (Select t.���� From Zltools.Zlbaktables T Where t.ϵͳ = " & mlngSys & ")" & vbNewLine & _
            "       Connect By Nocycle Prior c.Table_Name = c.r_Table_Name) D, User_Cons_Columns E" & vbNewLine & _
            "Where Not Exists (Select 1 From Zltools.Zlbaktables T Where t.���� = d.Table_Name) And" & vbNewLine & _
            "      e.Constraint_Name = d.Constraint_Name And e.Table_Name = d.Table_Name And e.Owner = d.Owner" & vbNewLine & _
            "Group By d.Constraint_Name, d.Table_Name"
            
    Call OpenRecordset(rsTmp, strSQL, "�ӱ�����δת��")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp)
    End If
    
    '(3)��ʷ��ռ���ɾ����Լ�������ڷ�������Ψһ����Լ������ɾ��,��ʷ��ռ��е�����Ψһ��û�ж�Ӧ�������ݿ��Ҳ��ɾ����
    lngCur = lngCur + 1
    SetPromptText ("��" & lngCur & "/" & lngTotal & ")���ڼ����ʷ�����Լ��")
    strSQL = "Select '��ʷ������Լ��' ������Ϣ, " & DT_HConDel & " ��������, a.Table_Name ����, a.Constraint_Name ������," & vbNewLine & _
            "       f_List2str(Cast(Collect(b.Column_Name Order By b.Position) As t_Strlist)) ����, Null ��������, Null ����, Null ����, '��' �Զ��޸�, '�ֹ�ɾ��Լ��' �޸�˵��" & vbNewLine & _
            "From All_Constraints" & mstrDBLink & " A, All_Cons_Columns" & mstrDBLink & " B" & vbNewLine & _
            "Where a.Owner =  '" & UCase(mstrBakOwnerName) & "' And a.Owner = b.Owner And a.Constraint_Name = b.Constraint_Name And Exists" & vbNewLine & _
            " (Select 1 From Zltools.Zlbaktables T Where t.ϵͳ =  " & mlngSys & " And t.���� = a.Table_Name) And" & vbNewLine & _
            "      (a.Constraint_Type Not In ('P', 'U') Or a.Constraint_Type In ('P', 'U') And Not Exists" & vbNewLine & _
            "       (Select 1 From User_Constraints C Where c.Constraint_Name = a.Constraint_Name))" & vbNewLine & _
            "Group By a.Table_Name, a.Constraint_Name"

    Call OpenRecordset(rsTmp, strSQL, "��ʷ������Լ��")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp)
    End If
    '��4����ʷ��ռ�ȱ�ٵ�������Ψһ��Լ��
    lngCur = lngCur + 1
    SetPromptText ("��" & lngCur & "/" & lngTotal & ")���ڼ����ʷ��ȱ�ٵ�Լ��")
    strSQL = "Select  '��ʷ��ȱ�ٵ�Լ��' ������Ϣ, " & DT_HConAdd & " ��������, Table_Name ����, Constraint_Name ������," & vbNewLine & _
            "       f_List2str(Cast(Collect(Column_Name Order By a.Position) As t_Strlist)) ����, Constraint_Type ��������, Null ����, Null ����, '��' �Զ��޸�," & vbNewLine & _
            "       '����Լ��' �޸�˵��" & vbNewLine & _
            "From (Select a.Table_Name, a.Constraint_Name, a.Column_Name, Nvl(a.Position, 1) Position, b.Constraint_Type" & vbNewLine & _
            "       From User_Cons_Columns A, User_Constraints B, Zltools.Zlbaktables T" & vbNewLine & _
            "       Where a.Constraint_Name = b.Constraint_Name And b.Table_Name = t.���� And t.ϵͳ = " & mlngSys & " And" & vbNewLine & _
            "             b.Constraint_Type In ('P', 'U') And Not Exists" & vbNewLine & _
            "        (Select 1" & vbNewLine & _
            "              From All_Constraints" & mstrDBLink & " C" & vbNewLine & _
            "              Where c.Owner = '" & UCase(mstrBakOwnerName) & "' And c.Constraint_Type In ('P', 'U') And c.Table_Name = t.���� And" & vbNewLine & _
            "                    c.Constraint_Name = b.Constraint_Name)) A" & vbNewLine & _
            "Group By a.Table_Name, a.Constraint_Name, Constraint_Type"
            
    Call OpenRecordset(rsTmp, strSQL, "��ʷ��ȱ�ٵ�Լ��")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp)
    End If
    '(5)��ʷ��ռ�Լ���б䶯
    lngCur = lngCur + 1
    SetPromptText ("��" & lngCur & "/" & lngTotal & ")���ڼ����ʷ��Լ���еĲ���")
    strSQL = "Select  'Լ���в���' ������Ϣ, " & DT_HConColDIf & " ��������, u.Table_Name ����, u.Constraint_Name ������, u.Colstr ����," & vbNewLine & _
            "       (Select c.Constraint_Type" & vbNewLine & _
            "         From User_Constraints C" & vbNewLine & _
            "         Where c.Constraint_Name = u.Constraint_Name And u.Table_Name = c.Table_Name) ��������, Null ����, Null ����, '��' �Զ��޸�," & vbNewLine & _
            "       '�ؽ�Լ��' �޸�˵��" & vbNewLine & _
            "From (Select a.Table_Name, a.Constraint_Name, f_List2str(Cast(Collect(a.Column_Name Order By a.Position) As t_Strlist)) Colstr" & vbNewLine & _
            "       From User_Cons_Columns A, Zltools.Zlbaktables T" & vbNewLine & _
            "       Where a.Table_Name = t.���� And t.ϵͳ = " & mlngSys & vbNewLine & _
            "       Group By a.Table_Name, a.Constraint_Name) U," & vbNewLine & _
            "     (Select a.Table_Name, a.Constraint_Name, f_List2str(Cast(Collect(a.Column_Name Order By a.Position) As t_Strlist)) Colstr" & vbNewLine & _
            "       From All_Cons_Columns" & mstrDBLink & " A, Zltools.Zlbaktables T" & vbNewLine & _
            "       Where a.Table_Name = t.���� And t.ϵͳ = " & mlngSys & " And a.Owner =  '" & UCase(mstrBakOwnerName) & "'" & vbNewLine & _
            "       Group By a.Table_Name, a.Constraint_Name) H" & vbNewLine & _
            "Where u.Table_Name = h.Table_Name And u.Constraint_Name = h.Constraint_Name And u.Colstr <> h.Colstr"

    '(5)��ʷ��������ռ����
    lngCur = lngCur + 1
    SetPromptText ("��" & lngCur & "/" & lngTotal & ")���ڼ����ʷ��Լ���еĲ���")
    strSQL = "Select '��ʷ��������ռ����' ������Ϣ, " & DT_HIndexTablesapce & " ��������, a.Table_Name ����, a.Index_Name ������, Null ����, Null ��������, Null ����, Null ����, '��' �Զ��޸�," & vbNewLine & _
            "       '�ƶ���������ռ�" & mstrBakIndexDB & "' �޸�˵��" & vbNewLine & _
            "From (Select d.Table_Name, d.Index_Name," & vbNewLine & _
            "              f_List2str(Cast(Collect(d.Column_Name Order By d.Column_Position) As t_Strlist)) Colstr" & vbNewLine & _
            "       From User_Ind_Columns d, Zltools.Zlbaktables t" & vbNewLine & _
            "       Where d.Table_Name = t.����" & vbNewLine & _
            "       And t.ϵͳ = " & mlngSys & "" & vbNewLine & _
            "       And Instr(d.Index_Name, '_IX_��ת��') = 0" & vbNewLine & _
            "       Group By d.Table_Name, d.Index_Name) a," & vbNewLine & _
            "     (Select e.Table_Name, f_List2str(Cast(Collect(e.Column_Name Order By e.Position) As t_Strlist)) Colstr" & vbNewLine & _
            "       From User_Cons_Columns e, User_Constraints f, Zltools.Zlbaktables t, User_Constraints c," & vbNewLine & _
            "            (Select ���� From Zltools.Zlbaktables Union All" & vbNewLine & _
            "              Select '������ҳ' From Dual Union All" & vbNewLine & _
            "              Select '������Ϣ' From Dual) g" & vbNewLine & _
            "       Where e.Table_Name = t.���� And t.ϵͳ = " & mlngSys & " And e.Constraint_Name = f.Constraint_Name" & vbNewLine & _
            "       And f.Constraint_Type = 'R' And c.Constraint_Name = f.r_Constraint_Name And g.����(+) = c.Table_Name And g.���� Is Null" & vbNewLine & _
            "       Group By e.Table_Name, e.Constraint_Name) b," & vbNewLine & _
            "     (Select Index_Name From All_Indexes" & mstrDBLink & " k, Zltools.Zlbaktables t" & vbNewLine & _
            "       Where k.Table_Owner = '" & UCase(mstrBakOwnerName) & "' And k.Table_Name = t.���� And t.ϵͳ = " & mlngSys & vbNewLine & _
            "       And k.Status = 'VALID'  And k.Tablespace_Name Not in( '" & mstrBakIndexDB & "'" & IIf(mstrBakIndexDB = mstrBakDB, ",'" & mstrBakDB & "_IDX')", ")") & ") h" & vbNewLine & _
            "Where a.Table_Name = b.Table_Name(+) And a.Colstr = b.Colstr(+)" & vbNewLine & _
            "And (b.Colstr Is Null Or Instr(a.Index_Name, '_PK') > 0 Or Instr(a.Index_Name, '_UQ') > 0)" & vbNewLine & _
            "And a.Index_Name = h.Index_Name"
            
    'strLobTable��Ӧ�ĵ���SQL
    'k.Status = 'VALID'  And  (k.Tablespace_Name Not in( '" & mstrBakIndexDB & "'" & IIf(mstrBakIndexDB = mstrBakDB, ",'" & mstrBakDB & "_IDX')", ")") & " OR K.Table_Name In(" & strLobTable & "))) h" & vbNewLine & _

    Call OpenRecordset(rsTmp, strSQL, "��ʷ��Լ���в���")
    If mblnUpdate Then
        Call GetFixSQL(rsTmp, mrsSQL)
    Else
        Call LoadDataByRecord(rsTmp)
    End If
    
    If Not mblnUpdate Then
        LoadCheckData = vsCheckResult.Rows <> vsCheckResult.FixedRows
        
        With vsCheckResult
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, RC_DifInfo) = "H��ͼ�ؽ�"
            .TextMatrix(.Rows - 1, RC_AutoRep) = "��"
            .TextMatrix(.Rows - 1, RC_RepMethod) = "���´���H��ͼ"
             .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, RC_DifInfo) = "��ع������±���"
            .TextMatrix(.Rows - 1, RC_AutoRep) = "��"
            .TextMatrix(.Rows - 1, RC_RepMethod) = "���±�������ת����صĴ洢����"
        End With
        
        vsCheckResult.Redraw = True
    End If
    Exit Function
errH:
    If 1 = 0 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, Me.Caption
End Function


Private Sub LoadDataByRecord(ByVal rsTmp As ADODB.Recordset, Optional ByVal blnClear As Boolean)
    '-------------------------------------------------------------------------------------------------------------
    '����:����¼�����ص���ʷ��ռ�������
    '������rsTmp �������¼��
    '      blnClear ��ձ������
    '-------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    Dim strTmp As String
    Dim strSQL As String
    
    On Error GoTo errH:
    With vsCheckResult
    
        If blnClear Then
            .Rows = .FixedRows
        End If
        
        If rsTmp.RecordCount = 0 Then
            Exit Sub
        End If
        
        While Not rsTmp.EOF
            .Rows = .Rows + 1
            lngRow = .Rows - 1
            
            If Val(rsTmp!��������) = -1 Then
                .TextMatrix(lngRow, RC_ObjType) = rsTmp!�������� & ""
                .TextMatrix(lngRow, RC_ObjLen) = rsTmp!�󱸳��� & ""
                .TextMatrix(lngRow, RC_ObjScale) = rsTmp!�󱸾��� & ""
                .Cell(flexcpData, lngRow, RC_ObjLen) = rsTmp!���߳��� & ""
                .Cell(flexcpData, lngRow, RC_ObjScale) = rsTmp!���߾��� & ""
                .TextMatrix(lngRow, RC_DifInfo) = "�о��Ȳ���"
                If Val(rsTmp!�󱸾��� & "") <= Val(rsTmp!���߾��� & "") And (Val(rsTmp!�󱸳��� & "") - Val(rsTmp!�󱸾��� & "")) <= (Val(rsTmp!���߳��� & "") - Val(rsTmp!���߾��� & "")) Then
                    .TextMatrix(lngRow, RC_DifType) = 4
                    .TextMatrix(lngRow, RC_AutoRep) = "��"
                    .TextMatrix(lngRow, RC_RepMethod) = "���󾫶�"
                    
                Else
                    .TextMatrix(lngRow, RC_DifType) = 5
                    .TextMatrix(lngRow, RC_AutoRep) = "��"
                    .TextMatrix(lngRow, RC_RepMethod) = "��ʷ��ռ���о��ȴ������߿�,���ܻ�Ӱ���ѡ���ع���"
                End If
            Else
                .TextMatrix(lngRow, RC_ObjType) = rsTmp!�������� & ""
                .TextMatrix(lngRow, RC_ObjLen) = rsTmp!���� & ""
                .TextMatrix(lngRow, RC_ObjScale) = rsTmp!���� & ""
                .TextMatrix(lngRow, RC_DifType) = rsTmp!�������� & ""
                .TextMatrix(lngRow, RC_DifInfo) = rsTmp!������Ϣ & ""
                .TextMatrix(lngRow, RC_AutoRep) = rsTmp!�Զ��޸� & ""
                .TextMatrix(lngRow, RC_RepMethod) = rsTmp!�޸�˵�� & ""
            End If
            .TextMatrix(lngRow, RC_TabName) = rsTmp!���� & ""
            .TextMatrix(lngRow, RC_ObjName) = rsTmp!������ & ""
            .TextMatrix(lngRow, RC_ColName) = rsTmp!���� & ""
            
            If .TextMatrix(lngRow, RC_DifType) <> "" Then
                Select Case Val(.TextMatrix(lngRow, RC_DifType))
                    Case DT_HLackTab
                        If mstrDBLink = "" Then
                            ReDim Preserve marrOnlineAddSQL(UBound(marrOnlineAddSQL) + 1)
                            marrOnlineAddSQL(UBound(marrOnlineAddSQL)) = " Grant Select On " & .TextMatrix(lngRow, RC_TabName) & " To " & mstrBakOwnerName  '�Ժ󱸿����Ա�������߿���Ӧ���SelectȨ��
                            If Not ExistsSynonym(.TextMatrix(lngRow, RC_TabName)) Then  'Ϊ��������ͬ���
                                ReDim Preserve marrOnlineAddSQL(UBound(marrOnlineAddSQL) + 1)
                                marrOnlineAddSQL(UBound(marrOnlineAddSQL)) = " Create Public Synonym " & .TextMatrix(lngRow, RC_TabName) & " For " & .TextMatrix(lngRow, RC_TabName)
                            End If
                        End If
                        strSQL = CreateTable(gcnOracle, mstrOwnerName, mstrBakDB, mstrBakOwnerName, .TextMatrix(lngRow, RC_TabName), mstrBakLobDB)
                        If strSQL <> "" Then
                            .TextMatrix(lngRow, RC_RepSQL) = strSQL
                        End If
                        If mstrDBLink = "" Then
                            '�����߿����Ա����󱸿���Ӧ�������Ȩ��
                            ReDim Preserve marrBakAddSQL(UBound(marrBakAddSQL) + 1)
                            marrBakAddSQL(UBound(marrBakAddSQL)) = " Grant All On " & .TextMatrix(lngRow, RC_TabName) & " To " & mstrOwnerName & " with Grant option"
                        End If
                    Case DT_HMoreCol '��ʷ��ռ��һ��
                    
                    Case DT_HLessCol '��ʷ��ռ���һ��
                        If Val(.TextMatrix(lngRow, RC_ObjLen)) = 0 Then
                            strTmp = .TextMatrix(lngRow, RC_ObjType)
                        Else
                            If Val(.TextMatrix(lngRow, RC_ObjScale)) = 0 Then
                                strTmp = .TextMatrix(lngRow, RC_ObjType) & "(" & .TextMatrix(lngRow, RC_ObjLen) & ")"
                            Else
                                strTmp = .TextMatrix(lngRow, RC_ObjType) & "(" & .TextMatrix(lngRow, RC_ObjLen) & "," & Val(.TextMatrix(lngRow, RC_ObjScale)) & ")"
                            End If
                        End If
                        .TextMatrix(lngRow, RC_RepSQL) = "Alter Table " & .TextMatrix(lngRow, RC_TabName) & " Add " & .TextMatrix(lngRow, RC_ColName) & " " & strTmp
                    Case DT_HDataTypeDif '�����Ͳ���
                        
                    Case DT_HRepLenDif '���޸��о��Ȳ���
                        If Val(.Cell(flexcpData, lngRow, RC_ObjLen)) = 0 Then
                            strTmp = .TextMatrix(lngRow, RC_ObjType)
                        Else
                            If Val(.Cell(flexcpData, lngRow, RC_ObjScale)) = 0 Then
                                strTmp = .TextMatrix(lngRow, RC_ObjType) & "(" & .Cell(flexcpData, lngRow, RC_ObjLen) & ")"
                            Else
                                strTmp = .TextMatrix(lngRow, RC_ObjType) & "(" & .Cell(flexcpData, lngRow, RC_ObjLen) & "," & Val(.Cell(flexcpData, lngRow, RC_ObjScale)) & ")"
                            End If
                        End If
                        .TextMatrix(lngRow, RC_RepSQL) = "Alter Table " & .TextMatrix(lngRow, RC_TabName) & " Modify " & .TextMatrix(lngRow, RC_ColName) & " " & strTmp
                    Case DT_HNotRepLenDif '�����޸��о��Ȳ���
                    
                    Case DT_HIndUsable, DT_HIndexTablesapce '��ʷ��ʧЧ����
                        .TextMatrix(lngRow, RC_RepSQL) = "Alter Index " & .TextMatrix(lngRow, RC_ObjName) & " Rebuild Tablespace " & mstrBakIndexDB
                    Case DT_HIndDel '��ʷ����������
                        .TextMatrix(lngRow, RC_RepSQL) = "Drop Index " & .TextMatrix(lngRow, RC_ObjName)
                    Case DT_HIndAdd '��ʷ��ȱ�ٵ�����
                        .TextMatrix(lngRow, RC_RepSQL) = "Create Index " & .TextMatrix(lngRow, RC_ObjName) & " On " & .TextMatrix(lngRow, RC_TabName) & "(" & .TextMatrix(lngRow, RC_ColName) & ")  Tablespace " & mstrBakIndexDB
                    Case DT_HIndColDif  '�����в���
                        .TextMatrix(lngRow, RC_RepSQL) = "Drop Index " & .TextMatrix(lngRow, RC_ObjName)
                        ReDim Preserve marrBakAddSQL(UBound(marrBakAddSQL) + 1)
                        marrBakAddSQL(UBound(marrBakAddSQL)) = "Create Index " & .TextMatrix(lngRow, RC_ObjName) & " On " & .TextMatrix(lngRow, RC_TabName) & "(" & .TextMatrix(lngRow, RC_ColName) & ")  Tablespace " & mstrBakIndexDB
                    Case DT_HConDisable '��ʷ���н���Լ��
                        .TextMatrix(lngRow, RC_RepSQL) = "Alter Table " & .TextMatrix(lngRow, RC_TabName) & " Enable Constraint " & .TextMatrix(lngRow, RC_ObjName)
                    Case DT_HConDel '��ʷ�����Լ��
                        If .TextMatrix(lngRow, RC_ObjName) Like "*_PK" Or .TextMatrix(lngRow, RC_ObjName) Like "*_UQ_*" Then
                            .TextMatrix(lngRow, RC_RepSQL) = "Alter Table " & .TextMatrix(lngRow, RC_TabName) & " Drop Constraint " & .TextMatrix(lngRow, RC_ObjName) & " Cascade Drop Index"
                        Else
                            .TextMatrix(lngRow, RC_RepSQL) = "Alter Table " & .TextMatrix(lngRow, RC_TabName) & " Drop Constraint " & .TextMatrix(lngRow, RC_ObjName)
                        End If
                    Case DT_URefConDel '�ӱ�����δת��
                        '���޸����û��ֹ��޸�
                    Case DT_HConAdd '��ʷ��ȱ�ٵ�Լ��
                        If .TextMatrix(lngRow, RC_ObjType) = "P" Then
                            .TextMatrix(lngRow, RC_RepSQL) = "Alter table " & .TextMatrix(lngRow, RC_TabName) & " Add Constraint " & .TextMatrix(lngRow, RC_ObjName) & " Primary Key(" & .TextMatrix(lngRow, RC_ColName) & ")  Using Index  Tablespace " & mstrBakIndexDB
                        ElseIf .TextMatrix(lngRow, RC_ObjType) = "U" Then
                            .TextMatrix(lngRow, RC_RepSQL) = "Alter table " & .TextMatrix(lngRow, RC_TabName) & " Add Constraint " & .TextMatrix(lngRow, RC_ObjName) & " Unique(" & .TextMatrix(lngRow, RC_ColName) & ")  Using Index  Tablespace " & mstrBakIndexDB
                        End If
                    Case DT_HConColDIf 'Լ���в���
                        If .TextMatrix(lngRow, RC_ObjType) = "P" Or .TextMatrix(lngRow, RC_ObjType) = "U" Then
                            .TextMatrix(lngRow, RC_RepSQL) = "Alter Table " & .TextMatrix(lngRow, RC_TabName) & " Drop Constraint " & .TextMatrix(lngRow, RC_ObjName) & " Cascade Drop Index"
                        Else
                            .TextMatrix(lngRow, RC_RepSQL) = "Alter Table " & .TextMatrix(lngRow, RC_TabName) & " Drop Constraint " & .TextMatrix(lngRow, RC_ObjName)
                        End If
                        ReDim Preserve marrBakAddSQL(UBound(marrBakAddSQL) + 1)
                        If .TextMatrix(lngRow, RC_ObjType) = "P" Then
                            marrBakAddSQL(UBound(marrBakAddSQL)) = "Alter table " & .TextMatrix(lngRow, RC_TabName) & " Add Constraint " & .TextMatrix(lngRow, RC_ObjName) & " Primary Key(" & .TextMatrix(lngRow, RC_ColName) & ") Using Index  Tablespace " & mstrBakIndexDB
                        ElseIf .TextMatrix(lngRow, RC_ObjType) = "U" Then
                            marrBakAddSQL(UBound(marrBakAddSQL)) = "Alter table " & .TextMatrix(lngRow, RC_TabName) & " Add Constraint " & .TextMatrix(lngRow, RC_ObjName) & " Unique(" & .TextMatrix(lngRow, RC_ColName) & ") Using Index  Tablespace " & mstrBakIndexDB
                        End If
                    Case DT_HLobTablespace '�����ֹ�����,��ȡ�����Σ�ע�����������ռ���SQL
    '                    .TextMatrix(lngRow, RC_RepSQL) ="Alter Table " & .TextMatrix(lngRow, RC_TabName)�� & " Move Tablespace " & mstrBakLobDB
                End Select
            End If
            If .TextMatrix(lngRow, RC_AutoRep) = "��" Then .Cell(flexcpBackColor, lngRow, .FixedCols, lngRow, .Cols - 1) = &H8000000F
            rsTmp.MoveNext
        Wend
    End With
    Exit Sub
errH:
    If 1 = 0 Then
        Resume
    End If
End Sub

Private Function ExistsSynonym(ByVal strTableName As String) As Boolean
'����:��ѯ��ǰ���Ƿ���ڹ���ͬ���
'       strTableName Ҫ���ı�
'���أ�true-���ڹ���ͬ��ʣ�false-�����ڹ���ͬ���

    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    strSQL = "Select 1 From All_Synonyms A Where a.Table_Owner = User And a.Synonym_Name ='" & strTableName & "' And a.Owner = 'PUBLIC'"
    Call OpenRecordset(rsTmp, strSQL, Me.Caption)
    ExistsSynonym = Not rsTmp.EOF
    
End Function

Private Sub LoadErrInfo(ByRef rsTmp As ADODB.Recordset)
    Dim lngRow As Long, i As Long

    With vsCheckResult
        .Rows = .FixedRows
        If rsTmp.RecordCount <> 0 Then
            rsTmp.MoveFirst
            For i = 0 To .Cols - 1
                 .ColHidden(i) = True
            Next
            .ColHidden(RC_TabName) = False
            .ColHidden(RC_ObjName) = False
            .ColHidden(RC_DifType) = False
            .ColWidth(RC_DifType) = 3000
            .ColHidden(RC_DifInfo) = False
            .ColHidden(RC_RepSQL) = False
            .ColWidth(RC_RepSQL) = 2700
            .TextMatrix(0, RC_DifInfo) = "���ݿռ�"
            .TextMatrix(0, RC_RepSQL) = "����SQL"
            .TextMatrix(0, RC_DifType) = "������Ϣ"
            .TextMatrix(0, RC_RepSQL) = "����SQL"
            While Not rsTmp.EOF
                .Rows = .Rows + 1
                lngRow = .Rows - 1
                .TextMatrix(lngRow, RC_TabName) = rsTmp!���� & ""
                .Cell(flexcpData, lngRow, RC_TabName) = .TextMatrix(lngRow, RC_TabName)
                .TextMatrix(lngRow, RC_ObjName) = rsTmp!������ & ""
                .Cell(flexcpData, lngRow, RC_ObjName) = .TextMatrix(lngRow, RC_ObjName)
                .TextMatrix(lngRow, RC_DifInfo) = IIf(Val(rsTmp!���ݿ�) = 0, "���߿�", "��ʷ��ռ�")
                .Cell(flexcpData, lngRow, RC_DifInfo) = .TextMatrix(lngRow, RC_DifInfo)
                .TextMatrix(lngRow, RC_DifType) = rsTmp!������Ϣ & ""
                .Cell(flexcpData, lngRow, RC_DifType) = .TextMatrix(lngRow, RC_DifType)
                .TextMatrix(lngRow, RC_RepSQL) = rsTmp!����SQL & ""
                .Cell(flexcpData, lngRow, RC_RepSQL) = .TextMatrix(lngRow, RC_RepSQL)
                rsTmp.MoveNext
            Wend
        End If
    End With
End Sub

Private Sub SetPromptText(ByVal strText As String)
    If Not mblnUpdate Then
        stbThis.Panels(2).Text = strText
        stbThis.Panels(2).ToolTipText = strText
    End If
End Sub

Private Sub SetFaceCtlEnable()
    cmdExit.Enabled = mblnAllRepair
    cmdRepair.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mblnAllRepair Then
        Cancel = True
    Else
        Set mrsErrInfo = Nothing
        Set mcnBakDB = Nothing
    End If
End Sub

Private Sub vsCheckResult_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    vsCheckResult.TextMatrix(Row, Col) = vsCheckResult.Cell(flexcpData, Row, Col)
End Sub

Private Sub vsCheckResult_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= vsCheckResult.FixedRows And NewCol >= vsCheckResult.FixedCols Then
        If NewRow <> OldRow Then
            vsCheckResult.ForeColorSel = &H0&
        End If
    End If
End Sub

Private Sub AddErrIntoRs(ByVal intDBType As Integer, Optional ByVal strErrInfo As String, Optional ByVal strTabName As String, Optional ByVal strObjName As String, Optional ByVal strSQL As String)
'���ܣ����ش�����Ϣ�ڼ�¼��
    With mrsErrInfo
        .AddNew
        !���ݿ� = intDBType
        If InStr(strErrInfo, "ORA-") > 0 Then
            !������Ϣ = Mid(strErrInfo, InStr(strErrInfo, "ORA-"))
        Else
            !������Ϣ = strErrInfo
        End If
        
        !���� = strTabName
        !������ = strObjName
        !����SQL = strSQL
        .Update
    End With
End Sub

Private Sub vsCheckResult_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = cmdRepair.Enabled Or Row < vsCheckResult.FixedRows
    If Not Cancel Then Cancel = Col <> RC_RepSQL
End Sub

Private Function GetIniRec() As ADODB.Recordset
'���ܣ���ȡ��ʼSQL��¼���������������ʱ���践����ʷ��������SQL��¼�����˴���SQL��¼����ʼ��
    Dim rsReturn As New ADODB.Recordset
    
    With rsReturn
        .Fields.Append "BAKDBName", adVarChar, 100 '��ʷ������
        .Fields.Append "BAKUser", adVarChar, 100
        .Fields.Append "SQL", adVarChar, 500 '���ݿ��޸�SQL
        .Fields.Append "ExecOrder", adInteger '����ȷ��SQLִ�е�ǰ��˳����Щ�޸���ҪһЩȨ�ޣ���ЩȨ��SQL��Ҫ��ǰִ�У��޸���ɺ�������ҪһЩ����������Щ����Ҫ�Ժ�ִ�С�
        .Fields.Append "FixType", adInteger '�޸����ͣ���������SQLִ��˳��
        .Fields.Append "ExecDB", adInteger 'ִ��SQL�����ݿ⣬0-��ʷ�⣬1-���߿�
        .Fields.Append "ExecIndex", adInteger 'SQL����˳��
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    Set GetIniRec = rsReturn
End Function

Private Function GetFixSQL(ByVal rsInput As ADODB.Recordset, ByRef rsSQL As ADODB.Recordset)
'���ܣ����ݲ�ѯ���Ĳ�����Ϣ��ȡ����SQL
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim strTmp As String
    
    With rsInput
        While Not .EOF
            Select Case Val(!��������)
                Case DT_HLackTab
                    '��ǰִ��SQL,ExecOrder=-1
                    If mstrDBLink = "" Then
                        Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, " Grant Select On " & !���� & " To " & mstrBakOwnerName, -1, DT_HLackTab, 1, mlngIndex)
                        mlngIndex = mlngIndex + 1
                    End If
                    '����ִ��SQL,ExecOrder=0
                    strSQL = CreateTable(gcnOracle, mstrOwnerName, mstrBakDB, mstrBakOwnerName, !����, mstrBakLobDB)
                    If strSQL <> "" Then
                        Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, strSQL, 0, DT_HLackTab, 0, mlngIndex)
                        mlngIndex = mlngIndex + 1
                    End If
                    If mstrDBLink = "" Then
                        '�Ӻ�ִ��SQL,ExecOrder=1
                        Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, " Grant All On " & !���� & " To " & mstrOwnerName & " with Grant option", 1, DT_HLackTab, 0, mlngIndex)
                        mlngIndex = mlngIndex + 1
                    End If
                Case DT_HMoreCol
                Case DT_HLessCol
                    If Val(!���� & "") = 0 Then
                        strTmp = !�������� & ""
                    Else
                        If Val(!���� & "") = 0 Then
                            strTmp = !�������� & "(" & !���� & ")"
                        Else
                            strTmp = !�������� & "(" & !���� & "," & !���� & ")"
                        End If
                    End If
                    Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Alter Table " & !���� & " Add " & !���� & " " & strTmp, 0, DT_HLessCol, 0, mlngIndex)
                    mlngIndex = mlngIndex + 1
                Case DT_HDataTypeDif
                
                Case -1
                    'Case DT_HRepLenDif
                    If Val(!�󱸾��� & "") <= Val(!���߾��� & "") And (Val(!�󱸳��� & "") - Val(!�󱸾��� & "")) <= (Val(!���߳��� & "") - Val(!���߾��� & "")) Then
                        If Val(!���߳��� & "") = 0 Then
                            strTmp = !�������� & ""
                        Else
                            If Val(!���߾��� & "") = 0 Then
                                strTmp = !�������� & "(" & !���߳��� & ")"
                            Else
                                strTmp = !�������� & "(" & !���߳��� & "," & !���߾��� & ")"
                            End If
                        End If
                        Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Alter Table " & !���� & " Modify " & !���� & " " & strTmp, 0, DT_HRepLenDif, 0, mlngIndex)
                        mlngIndex = mlngIndex + 1
                    Else
                    'Case DT_HNotRepLenDif
                    
                    End If
                Case DT_HIndUsable, DT_HIndexTablesapce
                    Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Alter Index " & !������ & " Rebuild TableSpace " & mstrBakIndexDB, 0, DT_HIndUsable, 0, mlngIndex)
                    mlngIndex = mlngIndex + 1
                Case DT_HIndDel
                    Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Drop Index " & !������, 0, DT_HIndDel, 0, mlngIndex)
                    mlngIndex = mlngIndex + 1
                Case DT_HIndAdd
                    Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Create Index " & !������ & " On " & !���� & "(" & !���� & ") TableSpace " & mstrBakIndexDB, 0, DT_HIndAdd, 0, mlngIndex)
                    mlngIndex = mlngIndex + 1
                Case DT_HIndColDif
                    Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Drop Index " & !������, 0, DT_HIndColDif, 0, mlngIndex)
                    mlngIndex = mlngIndex + 1
                    Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Create Index " & !������ & " On " & !���� & "(" & !���� & ")  TableSpace " & mstrBakIndexDB, 0, DT_HIndColDif, 0, mlngIndex)
                    mlngIndex = mlngIndex + 1
                Case DT_HConDisable
                    Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Alter Table " & !���� & " Enable Constraint " & !������, 0, DT_HConDisable, 0, mlngIndex)
                    mlngIndex = mlngIndex + 1
                Case DT_HConDel
                    If !������ Like "*_PK" Or !������ Like "*_UQ_*" Then
                        Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Alter Table " & !���� & " Drop Constraint " & !������ & " Cascade Drop Index", 0, DT_HConDel, 0, mlngIndex)
                    Else
                        Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Alter Table " & !���� & " Drop Constraint " & !������, 0, DT_HConDel, 0, mlngIndex)
                    End If
                    mlngIndex = mlngIndex + 1
                Case DT_URefConDel
                Case DT_HConAdd '����ȱ�����Լ��
                    '��������Ƿ���ڣ����ڣ���ɾ��
                    strSQL = "Select /*+rule*/" & vbNewLine & _
                                " 1" & vbNewLine & _
                                "From User_Indexes A" & vbNewLine & _
                                "Where A.Index_Name = '" & !������ & "'"
                    Call OpenRecordset(rsTmp, strSQL, "����ת����ع�����Ч���", , , mcnBakDB)
                    If Not rsTmp.EOF Then
                        Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, " Drop Index  " & !������, -1, DT_HConAdd, 0, mlngIndex)
                    End If
                    If !�������� & "" = "P" Then
                        Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Alter table " & !���� & " Add Constraint " & !������ & " Primary Key(" & !���� & ") Using Index  TableSpace " & mstrBakIndexDB, 0, DT_HConAdd, 0, mlngIndex)
                    ElseIf !�������� & "" = "U" Then
                        Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Alter table " & !���� & " Add Constraint " & !������ & " Unique(" & !���� & ")  Using Index  TableSpace " & mstrBakIndexDB, 0, DT_HConAdd, 0, mlngIndex)
                    End If
                    mlngIndex = mlngIndex + 1
                Case DT_HConColDIf
                    If !�������� & "" = "P" Or !�������� & "" = "U" Then
                        Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Alter Table " & !���� & " Drop Constraint " & !������ & " Cascade Drop Index", 0, DT_HConAdd, 0, mlngIndex)
                    Else
                        Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Alter Table " & !���� & " Drop Constraint " & !������, 0, DT_HConAdd, 0, mlngIndex)
                    End If
                    mlngIndex = mlngIndex + 1
                    If !�������� & "" = "P" Then
                        Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Alter table " & !���� & " Add Constraint " & !������ & " Primary Key(" & !���� & ")  Using Index  TableSpace " & mstrBakIndexDB, 0, DT_HConAdd, 0, mlngIndex)
                    ElseIf !�������� & "" = "U" Then
                        Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Alter table " & !���� & " Add Constraint " & !������ & " Unique(" & !���� & ")  Using Index  TableSpace " & mstrBakIndexDB, 0, DT_HConAdd, 0, mlngIndex)
                    End If
                    mlngIndex = mlngIndex + 1
                Case DT_HLobTablespace '�����ֹ�����,��ȡ�����Σ�ע�����������ռ���SQL
'                    Call ADDSQLToRec(rsSQL, mstrBakDB, mstrBakOwnerName, "Alter Table " & !���� & " Move Tablespace " & mstrBakLobDB, 0, DT_HConAdd, 0, mlngIndex)
            End Select
            .MoveNext
        Wend
    End With
End Function

Private Sub ADDSQLToRec(ByRef rsSQL As ADODB.Recordset, ByVal strBakDB As String, strBakUser As String, ByVal strSQL As String, ByVal intExecOrder As Integer, ByVal intFixType As Integer, ByVal intExecDB As Integer, ByVal lngExecIndex As Long)
    With rsSQL
        .AddNew
        !BAKDBName = strBakDB
        !BAKUser = strBakUser
        !SQL = strSQL
        !ExecOrder = intExecOrder
        !FixType = intFixType
        !ExecDB = intExecDB
        !ExecIndex = lngExecIndex
        .Update
    End With
End Sub

