VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInputTools 
   BackColor       =   &H80000005&
   Caption         =   "����������ɹ���"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmInputTools.frx":0000
   ScaleHeight     =   5715
   ScaleWidth      =   8865
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPolyphone 
      Caption         =   "��������(&P)"
      Height          =   350
      Left            =   2100
      TabIndex        =   17
      Top             =   5265
      Width           =   1145
   End
   Begin VB.CommandButton cmdBasic 
      Caption         =   "�����ִ�(&B)"
      Height          =   350
      Left            =   885
      TabIndex        =   16
      Top             =   5265
      Width           =   1145
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "ȫ��(&C)"
      Height          =   350
      Left            =   7305
      TabIndex        =   12
      Top             =   1035
      Width           =   1100
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "ȫѡ(&A)"
      Height          =   350
      Left            =   6150
      TabIndex        =   11
      Top             =   1035
      Width           =   1100
   End
   Begin VB.Frame fra2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "�����ļ�"
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   900
      TabIndex        =   2
      Top             =   4050
      Width           =   7035
      Begin VB.CommandButton cmdMakeFile 
         Caption         =   "�������(&G)"
         Height          =   350
         Left            =   5610
         TabIndex        =   15
         Top             =   660
         Width           =   1145
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "��"
         Height          =   255
         Left            =   5295
         TabIndex        =   9
         Top             =   705
         Width           =   255
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "�����"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   3270
         TabIndex        =   6
         Top             =   345
         Value           =   1  'Checked
         Width           =   1155
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ƴ����"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   2160
         TabIndex        =   5
         Top             =   345
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   2
         Left            =   855
         MaxLength       =   30
         TabIndex        =   4
         Top             =   315
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   690
         Width           =   4725
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���뷨��"
         Height          =   180
         Index           =   4
         Left            =   90
         TabIndex        =   3
         Top             =   375
         Width           =   720
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ļ�����"
         Height          =   180
         Index           =   2
         Left            =   90
         TabIndex        =   7
         Top             =   765
         Width           =   720
      End
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   6390
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cboSystem 
      Height          =   300
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   675
      Width           =   4185
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2505
      Left            =   900
      TabIndex        =   13
      Top             =   1410
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   4419
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "���ݱ�"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ѡ����"
      Height          =   180
      Index           =   1
      Left            =   900
      TabIndex        =   14
      Top             =   1155
      Width           =   720
   End
   Begin VB.Image imgICO 
      Height          =   480
      Left            =   150
      Picture         =   "frmInputTools.frx":803A
      Top             =   600
      Width           =   480
   End
   Begin VB.Label lblSys 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ӧ��ϵͳ"
      Height          =   180
      Left            =   900
      TabIndex        =   0
      Top             =   735
      Width           =   720
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����������ɹ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   105
      Width           =   1920
   End
End
Attribute VB_Name = "frmInputTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnStartUp As Boolean
Private mstr������ As String
Private mlngSys As Long
Private mintColumn As Integer
Private mlngLoop As Long

Private mcolBasicWord_PY As colBasicWord
Private mcolBasicWord_WB As colBasicWord
Private mcolBasicWord_POLY As colBasicWord

Private Const mstrNoHave As String = "��"
                                        
Private Const mstrNoSingle As String = "���£ãģţƣǣȣɣʣˣ̣ͣΣϣУѣңӣԣգ֣ףأ٣ڣ�������������������" & _
                                        "���࣡�����������������������ݣۣ���������������������������"

Private Function GetTimerString(ByVal sglSecend As Single) As String
    '
    
    GetTimerString = (sglSecend \ 3600) & "Сʱ"
    GetTimerString = GetTimerString & (sglSecend Mod 3600) \ 60 & "��"
    GetTimerString = GetTimerString & ((sglSecend Mod 3600) Mod 60) & "��"
    
End Function

Private Function GetValidWord(ByVal strWord As String) As String
    
    '��ȡ��Ч�����뷨��/���飬��ɾ����Ч���ַ�
    
    Dim lngLoop As Long
    Dim strTmp As String
    Dim lngLength As Long
    
    
    
    lngLength = Len(strWord)
    
    If InStr(mstrNoSingle, strWord) > 0 And Len(strWord) = 1 Then
        Exit Function
    End If
    
    For lngLoop = 1 To lngLength
        
        strTmp = Mid(strWord, lngLoop, 1)
        
        If Asc(strTmp) < 0 Then
            'Ϊ˫�ֽ��ַ�
            
            '����Ƿ��зǷ��ַ�����
            If InStr(mstrNoHave, UCase(strTmp)) = 0 Then
            
                GetValidWord = GetValidWord & strTmp
                
            End If
            
        End If
    Next
    
    GetValidWord = Trim(GetValidWord)
    
End Function

Private Function GetValidCode(ByRef strWord As String, ByRef strPy As String, ByRef strWb As String, Optional ByVal strTable As String) As Boolean
    Dim lngLoop As Long
    Dim rs As New ADODB.Recordset
    Dim strTmp As String
    Dim blnWb As Boolean
    Dim blnPy As Boolean
    Dim blnWord As Boolean
    Dim blnFirstPy As Boolean
    Dim blnFirstWb As Boolean
    Dim strCode As String
    
    Dim clsItem As clsBasicWord
    
    On Error GoTo errHand
    
    GetValidCode = True
    
    strTmp = strWord
    strWord = ""
    strPy = ""
    strWb = ""

    If Trim(strTmp) = "" Then Exit Function
        
    '����Ƿ��Ѿ��ڻ������д���
    If rs.State = adStateOpen Then rs.Close
    rs.Open "SELECT 1 FROM zlWordBasic where �ִ�='" & strTmp & "' AND ���뷨=1 and rownum<2", gcnOracle
    If rs.RecordCount > 0 Then blnPy = True
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open "SELECT 1 FROM zlWordBasic where �ִ�='" & strTmp & "' AND ���뷨=2 and rownum<2", gcnOracle
    If rs.RecordCount > 0 Then blnWb = True
            
    If blnPy And blnWb Then Exit Function
            
    For lngLoop = 1 To Len(strTmp)
        
        blnWord = False
                
        '-------------���Ϸ�ʽ
        If blnPy = False Then
        
            strCode = ""
            On Error Resume Next
            Err = 0
            Set clsItem = mcolBasicWord_POLY("K" & strTable & Mid(strTmp, lngLoop, 1))
            If Err = 0 Then strCode = clsItem.Codes
            On Error GoTo errHand
            
            If strCode = "" Then
                On Error Resume Next
                Err = 0
                Set clsItem = mcolBasicWord_PY("K" & Asc(Mid(strTmp, lngLoop, 1)))
                If Err = 0 Then strCode = clsItem.Codes
                On Error GoTo errHand
            End If
            
            If strCode <> "" Then
                strPy = strPy & strCode
                blnWord = True
            End If
        End If
        
        If blnWb = False Then
            
            strCode = ""
            
            On Error Resume Next
            Err = 0
            Set clsItem = mcolBasicWord_WB("K" & Asc(Mid(strTmp, lngLoop, 1)))
            If Err = 0 Then strCode = clsItem.Codes
            On Error GoTo errHand
            
            If strCode <> "" Then
                strWb = strWb & strCode
                blnWord = True
            End If
        End If
        
        '-------------���Ϸ�ʽ
                
        If blnWord Then strWord = strWord & Mid(strTmp, lngLoop, 1)
    Next
    
    '����볤����12λ�����Զ��ص�β��
    If Len(strPy) > 12 Then strPy = Left(strPy, 12)
    If Len(strWb) > 12 Then strWb = Left(strWb, 12)
        
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
    Exit Function
    
errHand:
    GetValidCode = False
End Function
    
Private Function CollectWords() As Boolean
    Dim strPy As String
    Dim strWb As String
    Dim strTmp As String
    Dim lngLoop As Long
    Dim rs As New ADODB.Recordset
    Dim lngCount As Long
    Dim strSQL As String
    
    '1.���ԭ�����ռ�����
    frmInputStatus.Describle = "�������ԭ�����ռ�����..."
    DoEvents

    On Error GoTo errHand
    
    gcnOracle.BeginTrans
    
    gcnOracle.Execute "ZL_zlWordCodes_DELETE", , adCmdStoredProc
    
    '2.��ȡ����ƴ���ֵ�������
    frmInputStatus.Describle = "����׼������ƴ����..."
    frmInputStatus.Value = 0
    lngLoop = 0
    DoEvents
    Set mcolBasicWord_PY = New colBasicWord
    rs.Open "select �ִ�,substr(������,1,1) AS ������ from zlwordbasic where ���뷨=1 and �Ƿ���=1 group by �ִ�,substr(������,1,1)", gcnOracle
    If rs.BOF = False Then
        Do While Not rs.EOF
            mcolBasicWord_PY.Add "K" & Asc(rs("�ִ�").Value), rs("������").Value
            rs.MoveNext
            
            lngLoop = lngLoop + 1
            frmInputStatus.Value = CInt(100 * lngLoop / rs.RecordCount)
            DoEvents
        Loop
    End If
    
    '3.��ȡ��������ֵ�������
    frmInputStatus.Describle = "����׼�����������..."
    frmInputStatus.Value = 0
    lngLoop = 0
    DoEvents
    Set mcolBasicWord_WB = New colBasicWord
    If rs.State = adStateOpen Then rs.Close
    rs.Open "select �ִ�,substr(������,1,1) AS ������ from zlwordbasic where ���뷨=2 and �Ƿ���=1 group by �ִ�,substr(������,1,1)", gcnOracle
    If rs.BOF = False Then
        Do While Not rs.EOF
            mcolBasicWord_WB.Add "K" & Asc(rs("�ִ�").Value), rs("������").Value
            rs.MoveNext
            
            lngLoop = lngLoop + 1
            frmInputStatus.Value = CInt(100 * lngLoop / rs.RecordCount)
            DoEvents
        Loop
    End If
    
    '4.��ȡ�����ֵ�������
    frmInputStatus.Describle = "����׼�����������÷���..."
    frmInputStatus.Value = 0
    lngLoop = 0
    DoEvents
    Set mcolBasicWord_POLY = New colBasicWord
    If rs.State = adStateOpen Then rs.Close
    rs.Open "SELECT ����,�ִ�,���� FROM zlWordPolyphone WHERE ϵͳ=" & mlngSys, gcnOracle
    If rs.BOF = False Then
        Do While Not rs.EOF
            mcolBasicWord_POLY.Add "K" & rs("����").Value & rs("�ִ�").Value, rs("����").Value
            rs.MoveNext
            
            lngLoop = lngLoop + 1
            frmInputStatus.Value = CInt(100 * lngLoop / rs.RecordCount)
            DoEvents
        Loop
    End If
    
    '5.��ʼ�ռ�����
    lngCount = 0
    For mlngLoop = 1 To lvw.ListItems.Count
        If lvw.ListItems(mlngLoop).Checked Then
            
            frmInputStatus.Describle = "�����ռ�""" & lvw.ListItems(mlngLoop).Text & """�е�""" & lvw.ListItems(mlngLoop).SubItems(1) & """..."
            frmInputStatus.Value = 0
            DoEvents
            
            If frmInputStatus.State Then
                'ȡ��
                gcnOracle.RollbackTrans
                Exit Function
            End If
            
            gstrSQL = "SELECT DISTINCT " & lvw.ListItems(mlngLoop).SubItems(1) & " AS ���� FROM " & lvw.ListItems(mlngLoop).Text
            
            If rs.State = adStateOpen Then rs.Close
            
            rs.Open gstrSQL, gcnOracle
            
            If rs.BOF = False Then
                lngLoop = 0
                Do While Not rs.EOF

                    'strTmp = GetValidWord(rsTmp("����").Value)
                    strTmp = rs("����").Value
                    
                    If strTmp <> "" Then
                        If GetValidCode(strTmp, strPy, strWb, lvw.ListItems(mlngLoop).Text) = False Then GoTo errHand
                        If Trim(strTmp) <> "" Then

                            If Trim(strPy) <> "" Then
                                
'                                If lngCount = 0 Then strSQL = "INSERT INTO zlWordCodes(�ִ�,������,���뷨) "
'                                strSQL = strSQL & IIf(lngCount = 0, "", " UNION ALL ") & " SELECT '" & strTmp & "','" & strPy & "',1 FROM DUAL "
'                                lngCount = lngCount + 1
                                gcnOracle.Execute "ZL_zlWordCodes_INSERT('" & strTmp & "','" & strPy & "',1)", , adCmdStoredProc
                            End If

                            If Trim(strWb) <> "" Then
                                                            
'                                If lngCount = 0 Then strSQL = "INSERT INTO zlWordCodes(�ִ�,������,���뷨) "
'                                strSQL = strSQL & IIf(lngCount = 0, "", " UNION ALL ") & " SELECT '" & strTmp & "','" & strWb & "',2 FROM DUAL "
'                                lngCount = lngCount + 1
                                
                                gcnOracle.Execute "ZL_zlWordCodes_INSERT('" & strTmp & "','" & strWb & "',2)", , adCmdStoredProc
                            End If
                            
'                            If lngCount >= 50 Then
'                                gcnOracle.Execute strSQL
'                                strSQL = ""
'                                lngCount = 0
'                            End If
                        End If
                        
                        If frmInputStatus.State Then
                            'ȡ��
                            gcnOracle.RollbackTrans
                            Exit Function
                        End If
                    End If

                    rs.MoveNext
                    
                    lngLoop = lngLoop + 1
                    frmInputStatus.Value = Int(100 * lngLoop / rs.RecordCount)
                    DoEvents
                Loop
            End If
            
        End If

    Next
    
'    If lngCount > 0 Then gcnOracle.Execute strSQL
    
    gcnOracle.CommitTrans
    
    CollectWords = True
    
    '6.��������
    Set mcolBasicWord_PY = Nothing
    Set mcolBasicWord_WB = Nothing
    Set mcolBasicWord_POLY = Nothing
    
    Exit Function
    
errHand:
    gcnOracle.RollbackTrans
    MsgBox "�ռ��ʻ�ʧ�ܣ��������ռ���", vbInformation, gstrSysName
End Function

Private Function MakeInputFile(ByVal strFile As String, ByVal strName As String, ByVal bytPy As Byte, ByVal bytWb As Byte) As Boolean
    Dim fso As FileSystemObject
    Dim tsFile As TextStream
    Dim rsTmp As New ADODB.Recordset
    Dim lngLoop As Long
    
    On Error GoTo errHand
    
    frmInputStatus.Describle = "���������������ļ�......"
    frmInputStatus.Value = 0
    
    If bytPy = 1 And bytWb = 1 Then
        gstrSQL = "SELECT �ִ�,������ FROM (SELECT �ִ�,������ FROM zlWordBasic " & _
                    " UNION " & _
                    "SELECT �ִ�,������ FROM zlWordCodes) GROUP BY ������,�ִ�"
    ElseIf bytPy = 1 And bytWb = 0 Then
        gstrSQL = "SELECT ������,�ִ� FROM (SELECT �ִ�,������ FROM zlWordBasic WHERE ���뷨=1 " & _
                    " UNION " & _
                    "SELECT �ִ�,������ FROM zlWordCodes  WHERE ���뷨=1)  GROUP BY ������,�ִ�"
    Else
        gstrSQL = "SELECT ������,�ִ� FROM (SELECT �ִ�,������ FROM zlWordBasic  WHERE ���뷨=2 " & _
                    " UNION " & _
                    "SELECT �ִ�,������ FROM zlWordCodes  WHERE ���뷨=2)  GROUP BY ������,�ִ�"
    End If
    
    rsTmp.Open gstrSQL, gcnOracle
    
    If rsTmp.BOF = False Then
        Set fso = New FileSystemObject
        If fso.FileExists(strFile) Then
            fso.DeleteFile strFile, True
        End If
        
        Set tsFile = fso.OpenTextFile(strFile, ForWriting, True, TristateTrue)
        
        tsFile.WriteLine "[Description]"
        tsFile.WriteLine "name=" & strName
        tsFile.WriteLine "MaxCodes=12"
        tsFile.WriteLine "MaxElement=1"
        tsFile.WriteLine "UsedCodes=abcdefghijklmnopqrstuvwxyz"
        tsFile.WriteLine "WildChar=?"
        tsFile.WriteLine "NumRules=0"
        tsFile.WriteLine "[Rule]"
        tsFile.WriteLine "[Text]"
        
        lngLoop = 0
        Do While Not rsTmp.EOF
            tsFile.WriteLine rsTmp("�ִ�").Value & rsTmp("������").Value
            rsTmp.MoveNext
            
            lngLoop = lngLoop + 1
            frmInputStatus.Value = Int(100 * lngLoop / rsTmp.RecordCount)
            DoEvents
            
            If frmInputStatus.State Then
                'ȡ��
                tsFile.Close
                Exit Function
            End If
        Loop
        tsFile.Close
    End If
    
    MakeInputFile = True
            
    Exit Function
    
errHand:
    Select Case Err.Number
    Case 76
        MsgBox "ָ�����ļ�·��δ�ҵ���", vbInformation, gstrSysName
    Case Else
        MsgBox "�����������ļ�ʧ�ܣ����飡", vbInformation, gstrSysName
    End Select
    
End Function

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�

End Sub

Private Sub cboSystem_Click()
    Dim rs As New ADODB.Recordset
    Dim objItem As ListItem
    
    If mblnStartUp Then Exit Sub
    If cboSystem.ListIndex = -1 Then Exit Sub
    
    mlngSys = cboSystem.ItemData(cboSystem.ListIndex)
    lvw.ListItems.Clear
        
    '��ȡ�ϴ�����ʱ��ѡ��
    gstrSQL = "SELECT C.*,D.ϵͳ " & _
                "FROM ( " & _
                    "SELECT * FROM (SELECT A.TABLE_NAME, A.COLUMN_NAME " & _
                      "FROM all_col_comments A, All_Tab_Columns B,all_objects E " & _
                     "WHERE A.OWNER = '" & mstr������ & "' " & _
                            "AND B.OWNER = '" & mstr������ & "' " & _
                            "AND INSTR(',����,������,', ',' || A.COLUMN_NAME || ',') > 0 " & _
                            "AND A.TABLE_NAME = B.TABLE_NAME " & _
                            "AND B.DATA_TYPE = 'VARCHAR2' " & _
                            "AND E.OBJECT_NAME=A.TABLE_NAME " & _
                            "AND E.object_type='TABLE' And Instr(E.OBJECT_NAME,'BIN$')<=0" & _
                    "UNION " & _
                    "SELECT table_name,'����' AS COLUMN_NAME from user_tables where table_name='��Ա��') " & _
                     "GROUP BY TABLE_NAME, COLUMN_NAME) C," & _
                    "(SELECT * FROM zlWordTable WHERE ϵͳ=" & mlngSys & ") D " & _
                 "WHERE C.TABLE_NAME=D.����(+)"
                 
    rs.Open gstrSQL, gcnOracle
    If rs.BOF = False Then
        Do While Not rs.EOF
            Set objItem = lvw.ListItems.Add(, , rs("TABLE_NAME").Value)
            objItem.SubItems(1) = rs("COLUMN_NAME").Value
            If IsNull(rs("ϵͳ").Value) = False Then
                objItem.Checked = True
            End If
            rs.MoveNext
        Loop
    End If
    
End Sub

Private Sub cboSystem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cmdBasic_Click()
    frmInputBasic.Show 1, frmMDIMain
End Sub

Private Sub cmdClear_Click()
    Dim lngLoop As Long
    
    For lngLoop = 1 To lvw.ListItems.Count
        lvw.ListItems(lngLoop).Checked = False
    Next
End Sub

Private Sub cmdMakeFile_Click()
    Dim fso As New FileSystemObject
    Dim tsFile As TextStream
    Dim sglStart As Single
    
    If MsgBox("�����Ҫ�����������ļ���", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
        
    If Trim(txt(2).Text) = "" Then
        MsgBox "��ȷ�����뷨�����ƣ�", vbInformation, gstrSysName
        txt(2).SetFocus
        Exit Sub
    End If
    
    If Trim(txt(1).Text) = "" Then
        MsgBox "��ȷ���������ŵ��ļ���·����", vbInformation, gstrSysName
        txt(1).SetFocus
        Exit Sub
    End If
    
    If chk(0).Value = 0 And chk(1).Value = 0 Then
        MsgBox "��ȷ��������ı��뷽ʽ��", vbInformation, gstrSysName
        chk(0).SetFocus
        Exit Sub
    End If
    
    On Error Resume Next
    
    Set fso = New FileSystemObject
    If fso.FileExists(txt(1).Text) Then
        If MsgBox(txt(1).Text & "�ļ��Ѿ����ڣ��Ƿ񸲸ǣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txt(1).SetFocus
            Exit Sub
        End If
    End If
    
    
    Err = 0
    Set fso = New FileSystemObject
    Set tsFile = fso.OpenTextFile(txt(1).Text, ForWriting, True, TristateTrue)
    If Err.Number = 76 Then
        MsgBox "ָ�����ļ�·��δ�ҵ���", vbInformation, gstrSysName
        tsFile.Close
        txt(1).SetFocus
        Exit Sub
    End If
    Err = 0
    tsFile.Close
    On Error GoTo 0
    
    '1.�ռ�����
    
    frmInputStatus.Show , frmMDIMain
    frmInputStatus.Describle = ""
    frmInputStatus.Value = 0
    DoEvents
    
    sglStart = Timer
    
    If CollectWords = False Then
        Unload frmInputStatus
        Exit Sub
    End If
    
    
    '2.�������뷨��װ�ļ�
    
    If MakeInputFile(txt(1).Text, txt(2).Text, chk(0).Value, chk(1).Value) = False Then
        Unload frmInputStatus
        Exit Sub
    End If
    
    Unload frmInputStatus
    
    MsgBox "�Ѿ�����������ļ������ɣ�          " & vbCrLf & _
            "��ʱ" & GetTimerString(Timer - sglStart), vbInformation, gstrSysName
    
    '����ѡ��ı����
    
    On Error GoTo errHand
    
    gcnOracle.BeginTrans
    
    gcnOracle.Execute "ZL_zlWordTable_DELETE(" & mlngSys & ")", , adCmdStoredProc
    
    For mlngLoop = 1 To lvw.ListItems.Count
        If lvw.ListItems(mlngLoop).Checked Then
            gstrSQL = "ZL_zlWordTable_INSERT(" & mlngSys & ",'" & lvw.ListItems(mlngLoop).Text & "','" & lvw.ListItems(mlngLoop).SubItems(1) & "')"
            gcnOracle.Execute gstrSQL, , adCmdStoredProc
        End If
    Next
    gcnOracle.CommitTrans
    
    Exit Sub
    
errHand:
    gcnOracle.RollbackTrans
End Sub

Private Sub cmdOpen_Click()
    
    dlg.Filter = "���뷨�ı��ļ�(*.txt)|*.txt"
    dlg.FileName = txt(1).Text
    dlg.ShowSave
    If dlg.FileName <> "" Then txt(1).Text = dlg.FileName
    
End Sub


Private Sub cmdPolyphone_Click()
    Call frmInputPolyphone.ShowEdit(frmMDIMain, mlngSys)
End Sub

Private Sub cmdSelect_Click()
    Dim lngLoop As Long
    
    For lngLoop = 1 To lvw.ListItems.Count
        lvw.ListItems(lngLoop).Checked = True
    Next
End Sub

Private Sub Form_Activate()
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errHandle
    If mblnStartUp = False Then Exit Sub

    '��ʾ����ʾ��ϵͳ
    mstr������ = UCase(gstrUserName)
    mlngSys = 0
    chk(0).Value = 0
    chk(1).Value = 0
        
    txt(1).Text = GetSetting("ZLSOFT", "ʵ�ù���\�����빤��\" & gstrUserName, "�ļ�����", App.Path & "\winpy.txt")
    txt(2).Text = GetSetting("ZLSOFT", "ʵ�ù���\�����빤��\" & gstrUserName, "���뷨��", gstrSysName)
    Select Case GetSetting("ZLSOFT", "ʵ�ù���\�����빤��\" & gstrUserName, "���뷽ʽ", 3)
    Case 1
        chk(0).Value = 1
    Case 2
        chk(1).Value = 1
    Case 3
        chk(0).Value = 1
        chk(1).Value = 1
    Case Else
        chk(0).Value = 0
        chk(1).Value = 0
    End Select
    
    Set rsTmp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", mstr������)
    
    If rsTmp.BOF = False Then
        Do While Not rsTmp.EOF
            cboSystem.AddItem rsTmp("����") & " v" & rsTmp("�汾��") & "��" & rsTmp("���") & "��"
            cboSystem.ItemData(cboSystem.NewIndex) = rsTmp("���")
            rsTmp.MoveNext
        Loop
        
        cboSystem.ListIndex = 0
        mlngSys = cboSystem.ItemData(cboSystem.ListIndex)
    Else
        cboSystem.Enabled = False
        lvw.Enabled = False
        cmdSelect.Enabled = False
        cmdClear.Enabled = False
        cmdBasic.Enabled = False
        cmdPolyphone.Enabled = False
        fra2.Enabled = False
    End If
    
    mblnStartUp = False
    DoEvents
    
    Call cboSystem_Click
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub Form_Load()
    mblnStartUp = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    With lvw
        .Width = Me.ScaleWidth - .Left - 60
        .Height = Me.ScaleHeight - fra2.Height - cmdBasic.Height - .Top - 150
    End With
    
    With fra2
        .Top = lvw.Top + lvw.Height + 60
        .Width = lvw.Width
    End With
    
    
    cmdBasic.Top = fra2.Top + fra2.Height + 60
    cmdPolyphone.Top = cmdBasic.Top
    
    With cmdClear
        .Left = lvw.Left + lvw.Width - .Width - 15
        cmdSelect.Left = .Left - .Width - 60
    End With
    
    
    With cmdMakeFile
        .Left = fra2.Width - .Width - 60
    End With
    
    With txt(1)
        .Width = fra2.Width - .Left - cmdMakeFile.Width - 120
    End With
    
    With cmdOpen
        .Left = txt(1).Left + txt(1).Width - .Width - 30
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call SaveSetting("ZLSOFT", "ʵ�ù���\�����빤��\" & gstrUserName, "�ļ�����", txt(1).Text)
    Call SaveSetting("ZLSOFT", "ʵ�ù���\�����빤��\" & gstrUserName, "���뷨��", txt(2).Text)
    
    If chk(0).Value = 1 And chk(1).Value = 1 Then
        Call SaveSetting("ZLSOFT", "ʵ�ù���\�����빤��\" & gstrUserName, "���뷽ʽ", 3)
    ElseIf chk(0).Value = 1 Then
        Call SaveSetting("ZLSOFT", "ʵ�ù���\�����빤��\" & gstrUserName, "���뷽ʽ", 1)
    ElseIf chk(1).Value = 1 Then
        Call SaveSetting("ZLSOFT", "ʵ�ù���\�����빤��\" & gstrUserName, "���뷽ʽ", 2)
    Else
        Call SaveSetting("ZLSOFT", "ʵ�ù���\�����빤��\" & gstrUserName, "���뷽ʽ", 0)
    End If
    
End Sub

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then
        lvw.SortOrder = IIf(lvw.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvw.SortKey = mintColumn
        lvw.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txt_GotFocus(Index As Integer)
    txt(Index).SelStart = 0
    txt(Index).SelLength = Len(txt(Index).Text)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub
