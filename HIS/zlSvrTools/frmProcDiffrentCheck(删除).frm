VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProcDiffrentCheck 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���̲�����"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10590
   Icon            =   "frmProcDiffrentCheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4545
      Index           =   0
      Left            =   240
      ScaleHeight     =   4545
      ScaleWidth      =   10080
      TabIndex        =   3
      Top             =   960
      Width           =   10080
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1755
         Index           =   0
         Left            =   210
         TabIndex        =   4
         Top             =   405
         Width           =   1935
         _cx             =   3413
         _cy             =   3096
         Appearance      =   1
         BorderStyle     =   0
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483626
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   330
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
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
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   9210
      TabIndex        =   2
      Top             =   5625
      Width           =   1100
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   195
      Picture         =   "frmProcDiffrentCheck.frx":6852
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   1
      Top             =   75
      Width           =   720
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "��ʼ(&O)"
      Height          =   350
      Left            =   7995
      TabIndex        =   0
      Top             =   5625
      Width           =   1100
   End
   Begin MSComctlLib.ProgressBar pbr 
      Height          =   105
      Left            =   180
      TabIndex        =   5
      Top             =   6045
      Visible         =   0   'False
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   6705
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   180
      Left            =   180
      TabIndex        =   8
      Top             =   5805
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��ȡ���½ű������뵱ǰ�Զ����̶�Ӧ�ı�׼���̽��бȽϵó�����"
      Height          =   180
      Left            =   1215
      TabIndex        =   7
      Top             =   630
      Width           =   5400
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "���̲�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1185
      TabIndex        =   6
      Top             =   150
      Width           =   1980
   End
End
Attribute VB_Name = "frmProcDiffrentCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private mobjMain As Object
Private mclsVsf As clsVsf
Private mblnReading As Boolean

Public Function ShowMe(ByVal objMain As Object)
    On Error GoTo errHand
    mblnOk = False
    Set mobjMain = objMain
    Me.Show 1, mobjMain
    
    ShowMe = mblnOk
    
    Exit Function
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
End Function

Private Function ExecuteCommand(ByVal strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim blnAllowModify As Boolean
    Dim strSQL As String
    Dim objItem As Object
    Dim intRow As Integer
    Dim intFlag As Integer
    Dim strUpPath As String
    Dim strFlag As String
    
    On Error GoTo errHand
    mblnReading = True
    Select Case strCommand
    '--------------------------------------------------------------------------------------------------------------
    Case "��ʼ�ؼ�"
        Set mclsVsf = New clsVsf
        With mclsVsf
            Call .Initialize(Me.Controls, vsf(0), True, True)
            Call .ClearColumn
            Call .AppendColumn("ѡ��", 500, flexAlignLeftCenter, flexDTBoolean, , "", False)
            Call .AppendColumn("�汾��", 0, flexAlignLeftCenter, flexDTString, , "", False)
            Call .AppendColumn("ϵͳ����", 2000, flexAlignLeftCenter, flexDTString, , "", False)
            Call .AppendColumn("��װ�ű�", 2800, flexAlignLeftCenter, flexDTString, , "", True)
            Call .AppendColumn("�����ű�", 0, flexAlignLeftCenter, flexDTString, , "", True)
            
            Call .InitializeEdit(True, False, False)
            Call .InitializeEditColumn(.ColIndex("ѡ��"), True, vbVsfEditCheck)
            Call .InitializeEditColumn(.ColIndex("��װ�ű�"), True, vbVsfEditCommand)
            Call .InitializeEditColumn(.ColIndex("�����ű�"), True, vbVsfEditCommand)

'            .AppendRows = True
        End With
    '--------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"
        
        With vsf(0)
            strSQL = "Select A.���,A.�汾��,A.���� as ϵͳ����,B.�ļ��� From zlSystems A,zlSysFiles B Where A.��� = B.ϵͳ And B.����=1"
            Set rs = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "")
            If rs.BOF = False Then
                For intRow = 0 To rs.RecordCount - 1
                    intFlag = intFlag + 1
                    If .Rows < intFlag + 1 Then .Rows = intFlag + 1
                    .TextMatrix(intRow + 1, .ColIndex("ϵͳ����")) = rs("ϵͳ����").value
                    .TextMatrix(intRow + 1, .ColIndex("��װ�ű�")) = rs("�ļ���").value
                    
                    strFlag = rs("�汾��").value
                    .TextMatrix(intRow + 1, .ColIndex("�汾��")) = strFlag
                    strFlag = Split(strFlag, ".")(0) & "." & Split(strFlag, ".")(1) & ".0"
                    'ȱʡ�����ű�
                    strUpPath = Split(rs("�ļ���").value, "Ӧ�ýű�")(0) & "�����ű�\" & strFlag & "\zlUpgrade.ini"
                    If gobjFile.FileExists(strUpPath) = True Then
                        .TextMatrix(intRow + 1, .ColIndex("�����ű�")) = strUpPath
                    End If
                    
                    .RowData(intRow + 1) = rs("���").value
                    rs.MoveNext
                Next
            End If
        End With
    End Select
    ExecuteCommand = True
    GoTo errEnd
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
    Exit Function
errEnd:
    mblnReading = False
    Exit Function
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHand
    
    Dim strTemp As String
    Dim str�ϴα�׼����·�� As String
    Dim str���±�׼����·�� As String
    Dim str�Աȱ���·�� As String
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim lngLoop As Long
    Dim i As Integer
    Dim strProcName As String
    Dim strIniPath As String
    Dim rsInit As ADODB.Recordset
    Dim intSysNumLast As Integer
    Dim strFlag As String
    Dim strCommand As String
    Dim lngTemp As Long
    Dim lngProcess As Long
    Dim rsSQL As ADODB.Recordset
    Dim objFolder As Folder
    Dim objFile As File
    Dim objFSO As TextStream
    Dim lngMaxLength As Long
    Dim str As String
    Dim strArr() As String
    Dim strIni1 As String
    Dim strIniSys As String
    Dim strIniApp As String
    Dim lngSys As Long
    
    Call gclsBase.SQLRecord(rsSQL)
    
    cmdOK.Enabled = False
    
    lblTitle = "���ڳ�ʼ��.."
    lblTitle.Visible = True
    
    str�ϴα�׼����·�� = App.Path & "\Tmp1"
    str���±�׼����·�� = App.Path & "\NewProcedure"
    str�Աȱ���·�� = App.Path & "\Reports"
        
        
    With vsf(0)
        strSQL = "Select ���,����,�汾�� From zlSystems a"
        Set rsData = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "")
        If rsData.BOF = True Then
            MsgBox "��ǰ���ݿ�û�а�װ�κ�ϵͳ��", vbInformation + vbOKOnly, "�������"
            GoTo errEnd
        End If
        For i = 1 To .Rows - 1
            If IIf(Abs(Val(.TextMatrix(i, .ColIndex("ѡ��")))) = 1, True, False) = True Then
                rsData.Filter = ""
                rsData.Filter = "���=" & .RowData(i)

                If .TextMatrix(i, vsf(0).ColIndex("��װ�ű�")) = "" Then
                    MsgBox "��ѡ��" & .TextMatrix(i, .ColIndex("ϵͳ����")) & "��װ�ű�"
                    GoTo errEnd
                End If
                Set rsInit = ReadINIToRec(vsf(0).TextMatrix(i, vsf(0).ColIndex("��װ�ű�")))
                rsInit.Filter = "��Ŀ='�汾��'"
                strIniApp = rsInit("����").value

                rsData.Filter = ""
                rsData.Filter = "���=" & .RowData(i)
                strIniSys = Trim(rsData("�汾��").value)
                
                If strIniSys <> strIniApp Then
                    MsgBox .TextMatrix(i, .ColIndex("ϵͳ����")) & "���ݿ�ϵͳ�汾�������ļ��汾��ƥ�䡣", vbInformation + vbOKOnly, "�������"
                    GoTo errEnd
                End If
            End If
        Next
    End With
    
    '����������ʱ�ļ���
    If gobjFile.FolderExists(str�ϴα�׼����·��) Then Call gobjFile.DeleteFolder(str�ϴα�׼����·��)
    If gobjFile.FolderExists(str���±�׼����·��) Then Call gobjFile.DeleteFolder(str���±�׼����·��)
    DoEvents
    
    Call gobjFile.CreateFolder(str�ϴα�׼����·��)
    Call gobjFile.CreateFolder(str���±�׼����·��)
    lblTitle.Visible = True
    
    
    '------------------------------------------------------------------------------------------------------------------
    '��ȡ���°�װ�ű��������ű��а����ı䶯���̣����ŵ���ʱ�ļ���1��
    
    For i = 1 To vsf(0).Rows - 1
        If Abs(Val(vsf(0).TextMatrix(i, vsf(0).ColIndex("ѡ��")))) = 1 Then
            
            '��ȡ��װ�ű��������ű��Ĺ��������ɵ����ű��ļ�
            '��ȡ��װ�ű�
            
            If Not gobjFile.FileExists(vsf(0).TextMatrix(i, vsf(0).ColIndex("��װ�ű�"))) Then
                MsgBox "�޷��򿪽ű��ļ�" & vsf(0).TextMatrix(i, vsf(0).ColIndex("��װ�ű�")) & ",ִ���жϡ�", vbExclamation, gstrSysName
                Exit Sub
            Else
                strIniPath = Mid(vsf(0).TextMatrix(i, vsf(0).ColIndex("��װ�ű�")), 1, Len(vsf(0).TextMatrix(i, vsf(0).ColIndex("��װ�ű�"))) - 11)
                strIniPath = strIniPath & "zlProgram.sql"
            End If
            
            lblTitle.Caption = "������ȡ��" & vsf(0).TextMatrix(i, vsf(0).ColIndex("ϵͳ����")) & "����װ�ű�.."
            Call CheckProcedure(strIniPath, str���±�׼����·��)
            pbr.value = 0
            pbr.Visible = False
            
            '��ȡ�����ű�
            strIniSys = vsf(0).TextMatrix(i, vsf(0).ColIndex("�汾��"))
            If Split(strIniSys, ".")(2) = 0 Then
                GoTo errNext
            ElseIf Not gobjFile.FolderExists(Split(vsf(0).TextMatrix(i, vsf(0).ColIndex("��װ�ű�")), "Ӧ�ýű�")(0) & "�����ű�\" & Split(strIniSys, ".")(0) & "." & Split(strIniSys, ".")(1) & ".0") Then
                MsgBox "�޷���⵽�����ű��ļ���,ִ���жϡ�", vbExclamation, gstrSysName
                GoTo errEnd
            Else
                strIniPath = Split(vsf(0).TextMatrix(i, vsf(0).ColIndex("��װ�ű�")), "Ӧ�ýű�")(0) & "�����ű�\" & Split(strIniSys, ".")(0) & "." & Split(strIniSys, ".")(1) & ".0" & "\"
            End If
        
'            If Not gobjFile.FileExists(vsf(0).TextMatrix(i, vsf(0).ColIndex("�����ű�"))) And vsf(0).TextMatrix(i, vsf(0).ColIndex("�����ű�")) <> "" Then
'                MsgBox "�޷��򿪽ű��ļ�" & vsf(0).TextMatrix(i, vsf(0).ColIndex("�����ű�")) & ",ִ���жϡ�", vbExclamation, gstrSysName
'                Exit Sub
'            ElseIf Trim(vsf(0).TextMatrix(i, vsf(0).ColIndex("�����ű�"))) = "" Then
'                GoTo errNext
'            Else
'                strIniPath = Mid(vsf(0).TextMatrix(i, vsf(0).ColIndex("�����ű�")), 1, Len(vsf(0).TextMatrix(i, vsf(0).ColIndex("�����ű�"))) - 13)
'            End If
            
'            Set rsInit = ReadINIToRec(vsf(0).TextMatrix(i, vsf(0).ColIndex("�����ű�")))
'            If Not CheckINIValid(rsInit, "ϵͳ��|Ŀ��汾") Then
'                MsgBox "��Ǩ�����ļ���ʽ����ȷ��", vbExclamation, "�������"
'                Exit Sub
'            End If
            
            
'            lblTitle.Caption = "������ȡ��" & vsf(0).TextMatrix(i, vsf(0).ColIndex("ϵͳ����")) & "�������ű�.."
'            rsInit.Filter = "��Ŀ='Ŀ��汾'"
             intSysNumLast = Split(strIniSys, ".")(2)
            For lngLoop = 10 To intSysNumLast Step 10
                strFlag = Split(strIniSys, ".")(0) & "." & Split(strIniSys, ".")(1) & "." & CStr(lngLoop)
                Call CheckProcedure(strIniPath & "ZL1_" & strFlag & ".sql", str���±�׼����·��)
            Next
        End If

errNext:

    Next
    
    '------------------------------------------------------------------------------------------------------------------
    '��ȡ�䶯���̶�Ӧ���ϴα�׼����
    lblTitle = "����׼���ϴεı�׼����.."
    lblTitle.Visible = True
    strSQL = "Select ID,����,������ From zlprocedure Where ���� In (1,2)"
    Set rsData = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "")
    If rsData.BOF = False Then
        pbr.value = 0
        pbr.Visible = True
        pbr.Max = rsData.RecordCount
        For lngLoop = 0 To rsData.RecordCount - 1
            strProcName = Nvl(rsData("����").value)
            
            If strProcName = "NEXTNO" Then
                strProcName = "NEXTNO"
            End If
            
            If gobjFile.FileExists(str���±�׼����·�� & "\" & strProcName & ".sql") Then
                strSQL = "Select A.ID,A.����,Upper(B.����) As ���� From zlProcedure A,zlProcedureText B Where A.ID = B.����ID And B.���� = 4 And A.ID=[1] Order By B.���"
                Set rs = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "", Val(Nvl(rsData("ID").value)))
                If rs.BOF = False Then
                    strTemp = ""
                    Do While Not rs.EOF
                        strTemp = strTemp & UCase(Nvl(rs("����").value))
                        rs.MoveNext
                    Loop
'                    rs.MoveFirst
                    Set objFSO = gobjFile.CreateTextFile(str�ϴα�׼����·�� & "\" & CStr(strProcName) & ".sql")
                    Call objFSO.Write(strTemp)
                    Call objFSO.Close
                End If
            End If
            rsData.MoveNext
            pbr.value = pbr.value + 1
        Next
    Else
        lblTitle.Visible = False
        MsgBox "��ǰ������û�б�׼���̺Ϳհ׹��̣�", vbInformation + vbOKOnly, "�������"
        Exit Sub
    End If
    
    '----------------------------���õ��������߽��жԱ������ļ����еĽű������ɱ���---------------------------------------------
    If gobjFile.FolderExists(str�Աȱ���·��) Then
        Call gobjFile.DeleteFolder(str�Աȱ���·��)
    End If
    Call gobjFile.CreateFolder(str�Աȱ���·��)
    '�����ݿ��еĹ�����ű����бȶԣ�����html����
    lblTitle.Caption = "���ڱȶ�.."
    If Not CompareFolder(str�ϴα�׼����·��, str���±�׼����·��, str�Աȱ���·��) Then
        Exit Sub
    End If
    '--------------------------���в���Ĺ����Զ������ĵ���״̬�޸�Ϊ"������"---------------------------------------------------
    Set objFolder = gobjFile.GetFolder(str�Աȱ���·��)
    lblTitle.Caption = "���ڵ�������״̬.."
    '�����д��ڵļ�Ϊ��Ҫ�����Ĺ���
    rsData.MoveFirst
    For i = 0 To rsData.RecordCount - 1
        If gobjFile.FileExists(str���±�׼����·�� & "\" & Nvl(rsData("����").value) & ".sql") Then
            If gobjFile.FileExists(str�Աȱ���·�� & "\" & Nvl(rsData("����").value) & ".sql.htm") Then
            '��׼����������ǰ���б仯
                strProcName = Nvl(rsData("����").value)
                Set rs = gclsBase.GetProInfo(strProcName)
                If rs.BOF = False Then
                    strSQL = "Zl_Zlprocedure_Update(" & rs("ID").value & "," & rs("����").value & ",'" & strProcName & "'," & ProcState.������ & ",'','" & Nvl(rsData("������").value) & "')"
                    Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
                End If
                strTemp = ""
                Set objFSO = gobjFile.OpenTextFile(str���±�׼����·�� & "\" & strProcName & ".sql")
                Do While Not objFSO.AtEndOfStream
                    If objFSO.Line = 1 Then
                        strTemp = strTemp & Replace(objFSO.ReadLine, "'", "''")
                    Else
                        strTemp = strTemp & vbCrLf & Replace(objFSO.ReadLine, "'", "''")
                    End If
                    DoEvents
                Loop
                Call objFSO.Close
                lngMaxLength = 3900
                If LenB(StrConv(strTemp, vbFromUnicode)) > lngMaxLength Then
                    strFlag = ""
                    str = ""
                    For lngLoop = 1 To Len(strTemp)
                        str = str & Mid(strTemp, lngLoop, 1)
                        If (LenB(StrConv(str, vbFromUnicode)) > lngMaxLength - 1 Or lngLoop = Len(strTemp)) And Mid(strTemp, lngLoop, 1) <> "'" Then
                            strFlag = strFlag & gstrSplite & str
                            str = ""
                        End If
                    Next
                    strFlag = Mid(strFlag, Len(gstrSplite) + 1)
                    strTemp = strFlag
                End If
                strArr = Split(strTemp, gstrSplite)
'                strSQL = "Zl_Zlproceduretext_Move(" & NVL(rsData("ID").value) & ",3,1,4,2)"
                For lngLoop = 0 To UBound(strArr)
'                    strSQL = "Zl_Zlproceduretext_Update(" & NVL(rsData("ID").value) & ",3," & (lngLoop + 1) & ",'" & strArr(lngLoop) & "')"
'                    Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
                    strSQL = "Zl_Zlproceduretext_Update(" & Nvl(rsData("ID").value) & ",4," & (lngLoop + 1) & ",'" & TrimNull(strArr(lngLoop)) & "')"
                    Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
                Next
            Else
                '��׼����������ǰ���ޱ仯
                strProcName = Nvl(rsData("����").value)
                strSQL = "Select A.ID,A.����,A.����,Upper(B.����) As ���� From zlProcedure A,zlProcedureText B Where A.ID = B.����ID And B.���� = 3 And A.����=[1]"
                Set rs = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "", strProcName)
                If rs.BOF = False Then
                    strSQL = "Zl_Zlprocedure_Update(" & rs("ID").value & "," & rs("����").value & ",'" & strProcName & "',3,'','" & Nvl(rsData("������").value) & "')"
                    Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
                    strTemp = ""
                    For lngLoop = 0 To rs.RecordCount - 1
                        strTemp = Replace(Nvl(rs("����").value), "'", "''")
                        
                        strSQL = "Zl_Zlproceduretext_Update(" & Nvl(rsData("ID").value) & ",3," & (lngLoop + 1) & ",'" & strTemp & "')"
                        Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
                        rs.MoveNext
                    Next
                End If
                Set objFSO = gobjFile.OpenTextFile(str���±�׼����·�� & "\" & strProcName & ".sql")
                strTemp = ""
                Do While Not objFSO.AtEndOfStream
                    If objFSO.Line = 1 Then
                        strTemp = strTemp & Replace(objFSO.ReadLine, "'", "''")
                    Else
                        strTemp = strTemp & vbCrLf & Replace(objFSO.ReadLine, "'", "''")
                    End If
                    DoEvents
                Loop
                
                Call objFSO.Close
                lngMaxLength = 3900
                If LenB(StrConv(strTemp, vbFromUnicode)) > lngMaxLength Then
                    strFlag = ""
                    str = ""
                    For lngLoop = 1 To Len(strTemp)
                        str = str & Mid(strTemp, lngLoop, 1)
                        If (LenB(StrConv(str, vbFromUnicode)) > lngMaxLength - 1 Or lngLoop = Len(strTemp)) And Mid(strTemp, lngLoop, 1) <> "'" Then
                            strFlag = strFlag & gstrSplite & str
                            str = ""
                        End If
                    Next
                    strFlag = Mid(strFlag, Len(gstrSplite) + 1)
                    strTemp = strFlag
                End If
                strArr = Split(strTemp, gstrSplite)
                
                For lngLoop = 0 To UBound(strArr)
                    strSQL = "Zl_Zlproceduretext_Update(" & Nvl(rsData("ID").value) & ",4," & (lngLoop + 1) & ",'" & TrimNull(strArr(lngLoop)) & "')"
                    Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
                Next
                
                
            End If
        End If
        rsData.MoveNext
    Next
    
    On Error Resume Next
    
    
    objFSO.Close
    Set objFSO = Nothing
    On Error GoTo errHand
    Call SQLRecordExecute(rsSQL, "")
    If gobjFile.FolderExists(str���±�׼����·��) Then
        Call gobjFile.DeleteFolder(str���±�׼����·��)
    End If
    If gobjFile.FolderExists(str�ϴα�׼����·��) Then
        Call gobjFile.DeleteFolder(str�ϴα�׼����·��)
    End If
    If gobjFile.FolderExists(str�Աȱ���·��) Then
        Call gobjFile.DeleteFolder(str�Աȱ���·��)
    End If
    
    lblTitle.Visible = False
    
    MsgBox "�������Ѿ���ɣ�", vbInformation, Me.Caption
    cmdOK.Enabled = True
    mblnOk = True
    Exit Sub
errEnd:
    mblnOk = True
    cmdOK.Enabled = True
    Exit Sub
errHand:
    MsgBox "������ʧ�ܣ�" & vbCrLf & err.Description, vbCritical, Me.Caption
    cmdOK.Enabled = True
End Sub

Public Function ReadINIToRec(ByVal strFile As String) As ADODB.Recordset
'���ܣ���ָ��INI�����ļ������ݶ�ȡ����¼����
'���أ�Nothing�����"��Ŀ,����"�ļ�¼��,����ͬһ��Ŀ�����ж�������
    Dim rsTmp As New ADODB.Recordset
    Dim objINI As Scripting.TextStream
    
    Dim strItem As String, strText As String
    Dim strLine As String
            
    rsTmp.Fields.Append "��Ŀ", adVarChar, 100
    rsTmp.Fields.Append "����", adVarChar, 4000, adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set objINI = gobjFile.OpenTextFile(strFile, ForReading)
    Do While Not objINI.AtEndOfStream
        strLine = Replace(objINI.ReadLine, vbTab, " ")
        If Left(Trim(strLine), 1) = "[" And InStr(strLine, "]") > InStr(strLine, "[") Then
            
            If strItem <> "" And strText = "" Then
                rsTmp.AddNew
                rsTmp!��Ŀ = strItem
                rsTmp!���� = Null
                rsTmp.Update
            End If
            
            strItem = Trim(Mid(strLine, InStr(strLine, "[") + 1, InStr(strLine, "]") - InStr(strLine, "[") - 1))
            strText = Trim(Mid(strLine, InStr(strLine, "]") + 1))
            If strItem <> "" And strText <> "" Then
                rsTmp.AddNew
                rsTmp!��Ŀ = strItem
                rsTmp!���� = strText
                rsTmp.Update
            End If
        ElseIf Trim(strLine) <> "" And strItem <> "" Then
            strText = Trim(strLine)
            rsTmp.AddNew
            rsTmp!��Ŀ = strItem
            rsTmp!���� = strText
            rsTmp.Update
        End If
    Loop
    
    If strItem <> "" And strText = "" Then
        rsTmp.AddNew
        rsTmp!��Ŀ = strItem
        rsTmp!���� = Null
        rsTmp.Update
    End If
    
    objINI.Close
    
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    
    Set ReadINIToRec = rsTmp
End Function

Private Function CheckINIValid(rsINI As ADODB.Recordset, ByVal strItem As String) As Boolean
'���ܣ�����Ӧ�������ļ���ʽ�Ƿ���ȷ
'������rsINI=��������ļ����ݵļ�¼��������"��Ŀ,����"�ֶ�
'      strItem=�����ļ��б���Ҫ�������ݵ���Ŀ��,��"��Ŀ1|��Ŀ2|..."
    Dim arrItem As Variant, i As Long
    
    arrItem = Split(strItem, "|")
    For i = 0 To UBound(arrItem)
        rsINI.Filter = "��Ŀ='" & arrItem(i) & "'"
        If rsINI.EOF Then Exit Function
        If IsNull(rsINI!����) Then Exit Function
    Next
    CheckINIValid = True
End Function

Private Function CheckProcedure(ByVal strFile As String, Optional strFilePath As String)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Dim lngLine As Long
    Dim strLine As String
    Dim strTemp As String
    Dim strFMT As String
    Dim blnSQL As Boolean
    Dim blnBlock As Boolean
    Dim strFlag As String
    Dim strFileProName As String
    Dim lngFileLines As Long
    Dim objFileTemp As TextStream
    Dim objFile As TextStream
    Dim blnFlag As Boolean
    Dim objPercent As New clsPercent
    Dim lngMsg As Long
    
    On Error GoTo errHand
    
    pbr.value = 0
    pbr.Visible = True

    Set objFile = gobjFile.OpenTextFile(strFile, ForReading)
    If objFile.AtEndOfStream Then
        objFile.Close
        Exit Function
    End If
        
    Do While Not objFile.AtEndOfStream
        objFile.ReadLine
    Loop
    lngFileLines = objFile.Line
    
    Call objPercent.InitPercent(pbr, lngFileLines)
    
    objFile.Close
    
    Dim blnSpaceProc As Boolean
    
    Set objFile = gobjFile.OpenTextFile(strFile, ForReading)
    Do While Not objFile.AtEndOfStream
        lngLine = objFile.Line '��ǰ�к�:δ��ȡ��֮ǰ,��ָ��δ�Ƶ���һ��
        strLine = objFile.ReadLine
        strFMT = UCase(TrimComment(TrimEx(strLine)))
        If strFMT Like "PROMPT *" Then GoTo NextLine
        
        
        If blnBlock Then
            If strFMT = "/" Then
                blnSQL = True
                blnBlock = False
                Do While Right(strTemp, 1) = Chr(10) Or Right(strTemp, 1) = Chr(13)
                   strTemp = Left(strTemp, Len(strTemp) - 1)
                Loop
                
                
                objFileTemp.Write "CREATE OR REPLACE " & strTemp
                DoEvents
                objFileTemp.Close
                strTemp = ""
                
                If blnSpaceProc = True Then
                    blnSpaceProc = False
                    
                    Set objFileTemp = gobjFile.OpenTextFile(strFilePath & "\" & strFileProName & ".sql")
                    strTemp = objFileTemp.ReadAll
                    objFileTemp.Close
                    strTemp = GetBlankProcedure(strTemp)
                    
                    DoEvents
                    Set objFileTemp = gobjFile.CreateTextFile(strFilePath & "\" & strFileProName & ".sql", True)
                    objFileTemp.Write strTemp
                    objFileTemp.Close
                    strTemp = ""
                End If
                
            Else
                strTemp = strTemp & vbCrLf & strLine
            End If
        ElseIf strFMT Like "CREATE OR REPLACE PROCEDURE *" Or strFMT Like "CREATE PROCEDURE *" _
            Or strFMT Like "CREATE OR REPLACE FUNCTION *" Or strFMT Like "CREATE FUNCTION *" _
            Or strFMT Like "CREATE OR REPLACE TRIGGER *" Or strFMT Like "CREATE TRIGGER *" _
            Or strFMT Like "CREATE OR REPLACE TYPE *" Or strFMT Like "CREATE TYPE *" _
            Or strFMT Like "CREATE OR REPLACE PACKAGE *" Or strFMT Like "CREATE PACKAGE *" Then
            
            blnBlock = True
            
            '�����������̽ű��ļ�
            strFlag = Replace(strFMT, "CREATE OR REPLACE ", "")
            strFlag = Replace(strFlag, "CREATE ", "")
            
            If InStr(strFlag, "(") > 0 Then strFlag = Left(strFlag, InStr(strFlag, "(") - 1)
            If InStr(strFlag, ".") > 0 Then strFlag = Split(strFlag, ".")(1)
            strFileProName = Split(strFlag, " ")(1)
            If gobjFile.FileExists(strFilePath & "\" & strFileProName & ".sql") Then
                Call gobjFile.DeleteFile(strFilePath & "\" & strFileProName & ".sql")
            End If
            
            '����Ƿ�Ϊ�հ׹���
            blnSpaceProc = False
            If IsSpaceProcedure("ZLHIS", strFileProName) = True Then
                blnSpaceProc = True
            End If
            
            Set objFileTemp = gobjFile.CreateTextFile(strFilePath & "\" & strFileProName & ".sql", True)
             
            strFlag = Replace(strFMT, "CREATE OR REPLACE ", "")
            strFlag = Replace(strFlag, "CREATE ", "")
            strTemp = strTemp & UCase(strFlag)
        End If
        
        Call objPercent.LoopPercent

NextLine:
    Loop
    objFile.Close
    pbr.Visible = False
    pbr.value = 0
'    MsgBox blnFlag
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
End Function

Public Function TrimEx(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
'���ܣ�ȥ��TAB�ַ������߿ո񣬻س������ֻ�ɵ��ո�ָ���
'˵������Ҫ��RunSQLFile���Ӻ���
    If blnCrlf Then
        strText = Replace(strText, vbCrLf, " ")
        strText = Replace(strText, vbCr, " ")
        strText = Replace(strText, vbLf, " ")
    End If
    strText = Trim(Replace(strText, vbTab, " "))
    
    Do While InStr(strText, "  ") > 0
        strText = Replace(strText, "  ", " ")
    Loop
    TrimEx = strText
End Function

Public Function TrimComment(ByVal strSQL As String) As String
'���ܣ�ȥ��д�ڵ���strSQL�������"--"ע��
'˵������Ҫ��RunSQLFile���Ӻ���
    Dim blnStr As Boolean
    Dim i As Long, K As Long
    
    If Left(strSQL, 2) <> "--" And InStr(strSQL, "--") > 0 Then
        For i = 1 To Len(strSQL)
            If Mid(strSQL, i, 1) = "'" Then blnStr = Not blnStr
            If Mid(strSQL, i, 2) = "--" And Not blnStr Then
                K = i: Exit For
            End If
        Next
        If K > 0 Then strSQL = RTrim(Left(strSQL, K - 1))
    End If
    TrimComment = strSQL
End Function

Private Sub Form_Load()
    Call ExecuteCommand("��ʼ�ؼ�")
    Call ExecuteCommand("��ʼ����")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    vsf(0).Move 15, 15, picPane(0).ScaleWidth - 30, picPane(0).ScaleHeight - 30
'    mclsVsf.AppendRows = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (mclsVsf Is Nothing) Then
        Set mclsVsf = Nothing
    End If
End Sub

Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If mblnReading = True Then Exit Sub
    Call mclsVsf.AfterEdit(Row, Col)
End Sub

Private Sub vsf_AfterMoveColumn(Index As Integer, ByVal Col As Long, Position As Long)
    Call mclsVsf.AfterMoveColumn(Col, Position)
'    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnReading = True Then Exit Sub
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
'    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
'    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mblnReading = True Then Exit Sub
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim rsData As New ADODB.Recordset
    Dim bytRet As Byte
    
    With vsf(0)
        Select Case Col
        '--------------------------------------------------------------------------------------------------------------
        Case .ColIndex("��װ�ű�")
            With dlg
                .DialogTitle = "ѡ��Ӧ�ð�װ�����ļ�"
                .Filter = "(Ӧ�ð�װ�����ļ�)|zlSetup.ini"
                .ShowOpen
                If .FileName = "" Then
                    Exit Sub
                Else
                    vsf(0).TextMatrix(vsf(0).Row, vsf(0).Col) = .FileName
                End If
            End With
        Case .ColIndex("�����ű�")
            With dlg
                .DialogTitle = "ѡ��Ӧ����Ǩ�����ļ�"
                .Filter = "Ӧ����Ǩ�����ļ�(zlUpgrade.ini)|zlUpgrade.ini"
                .Flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
                .ShowOpen
                On Error GoTo 0
                Me.Refresh
                If .FileName = "" Then
                    Exit Sub
                Else
                    vsf(0).TextMatrix(vsf(0).Row, vsf(0).Col) = .FileName
                End If
            End With
        End Select
        
        Call mclsVsf.SetFocus(, , True)
    End With
End Sub




