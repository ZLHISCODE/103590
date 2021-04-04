VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmChildQuestionFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3825
   Icon            =   "frmChildQuestionFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cbo 
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   0
      Left            =   1230
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   195
      Width           =   2430
   End
   Begin VB.TextBox txt������ 
      Height          =   300
      Left            =   1230
      TabIndex        =   7
      Top             =   2115
      Width           =   2070
   End
   Begin VB.CommandButton cmd������ 
      Height          =   300
      Left            =   3345
      Picture         =   "frmChildQuestionFilter.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2100
      Width           =   300
   End
   Begin VB.CommandButton cmdRef 
      Caption         =   "ˢ��(&R)"
      Height          =   350
      Left            =   255
      TabIndex        =   9
      Top             =   2685
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.ComboBox cbo������ 
      Height          =   300
      Left            =   1230
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1620
      Width           =   2430
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2475
      TabIndex        =   11
      Top             =   2685
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1365
      TabIndex        =   10
      Top             =   2685
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtp 
      Height          =   300
      Index           =   0
      Left            =   1230
      TabIndex        =   1
      Top             =   615
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   288227331
      CurrentDate     =   38083
   End
   Begin MSComCtl2.DTPicker dtp 
      Height          =   300
      Index           =   1
      Left            =   1230
      TabIndex        =   3
      Top             =   1050
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   288161795
      CurrentDate     =   38083
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "������(&3)"
      Height          =   180
      Index           =   2
      Left            =   330
      TabIndex        =   6
      Top             =   2160
      Width           =   810
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "������(&2)"
      Height          =   180
      Index           =   1
      Left            =   150
      TabIndex        =   4
      Top             =   1680
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "���ʱ��(&1)"
      Height          =   180
      Index           =   0
      Left            =   165
      TabIndex        =   0
      Top             =   225
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   180
      Index           =   9
      Left            =   810
      TabIndex        =   2
      Top             =   1080
      Width           =   180
   End
End
Attribute VB_Name = "frmChildQuestionFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################

Private mblnDataChanged As Boolean
Private mblnOK As Boolean
Private mstr��鿪ʼʱ�� As String
Private mstr������ʱ�� As String
Private mstr����ѡ�� As String
Private mstr������    As String
Private mlngCurNum As Long

Private mblnDataExecute As Boolean


'######################################################################################################################
Private Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
End Property

Private Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Public Function ShowPara(ByVal frmMain As Object, ByRef str��鿪ʼʱ�� As String, ByRef str������ʱ�� As String, ByRef str����ѡ�� As String, ByRef lngCurNum As Long, ByRef str������ As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mblnOK = False
    mblnDataExecute = True
    With cbo(0)
        .Clear
        .AddItem "��  ��"
        .AddItem "�Զ���"
        .AddItem "��  ��"
        .AddItem "��  ��"
        .AddItem "��  ��"
        .AddItem "��  ��"
        .AddItem "��  ��"
        .AddItem "������"
        .AddItem "��  ��"
        .AddItem "ǰ����"
        .AddItem "ǰһ��"
        .AddItem "ǰ����"
        .AddItem "ǰһ��"
        .AddItem "ǰ����"
        .AddItem "ǰ����"
        .AddItem "ǰ����"
        .AddItem "ǰһ��"
        .AddItem "ǰ����"
        .Text = "ǰһ��"
    End With
    mblnDataExecute = False
    
    If str����ѡ�� <> "" Then
        cbo(0).Text = str����ѡ��
    End If
    
    If cbo(0).Text = "�Զ���" Then
        If str��鿪ʼʱ�� = "" Then
            dtp(0).Value = Format(Now, "YYYY-MM-DD 00:00:00")
        Else
            dtp(0).Value = CDate(str��鿪ʼʱ��)
        End If
        
        If str������ʱ�� = "" Then
            dtp(1).Value = Format(Now, "YYYY-MM-DD 23:59:59")
        Else
            dtp(1).Value = CDate(str������ʱ��)
        End If
   
        Call Init������(str��鿪ʼʱ��, str������ʱ��)
    
        If lngCurNum = 0 Then
            cbo������.ListIndex = 0
        Else
    '        cbo������.Text = lngCurNum
            cbo������.ListIndex = 0
        End If
    End If
    
    If str������ = "" Then
        txt������.Text = ""
    Else
        txt������.Text = str������
    End If
    
    Call SetCob(lngCurNum)
    
    
    Me.Show 1, frmMain
    
    If mblnOK Then
        str��鿪ʼʱ�� = mstr��鿪ʼʱ��
        str������ʱ�� = mstr������ʱ��
        str����ѡ�� = mstr����ѡ��
        lngCurNum = mlngCurNum
        str������ = mstr������
        ShowPara = mblnOK
    End If
    
End Function

Private Sub cbo_Click(Index As Integer)
    
    If mblnDataExecute Then Exit Sub
    
    Select Case Index
    Case 0
        Select Case cbo(Index).Text
        Case "��  ��"
            dtp(0).Enabled = False
            dtp(1).Enabled = False
            dtp(0).Value = Format("2000-01-01 00:00:00", dtp(0).CustomFormat)
            dtp(1).Value = Format("3000-01-01 23:59:59", dtp(1).CustomFormat)
        Case "�Զ���"
            dtp(0).Enabled = True
            dtp(1).Enabled = True
        Case Else
            If dtp(0).Enabled = False Then
                dtp(0).Enabled = True
                dtp(1).Enabled = True
            End If
            dtp(0).Value = Format(GetBasePeriod(cbo(0).Text, 1), dtp(0).CustomFormat)
            dtp(1).Value = Format(GetBasePeriod(cbo(0).Text, 2), dtp(1).CustomFormat)
        End Select
        
         Call Init������(dtp(0).Value, dtp(1).Value)
         DataChanged = True
    End Select
    
    Dim strTempNum As String
    strTempNum = CLng(cbo������.ItemData(cbo������.ListIndex))
    Call Init������(dtp(0).Value, dtp(1).Value)
    If strTempNum = "" Then Exit Sub
'    cbo������.Text = strTempNum
    
    
    
End Sub

Private Sub cbo������_Change()
    DataChanged = True
End Sub

Private Sub cbo������_Click()
'    Dim strTempNum As String
'    strTempNum = CLng(cbo������.ItemData(cbo������.ListIndex))
'    Call Init������(dtp(0).Value, dtp(1).Value)
'    If strTempNum = "" Then Exit Sub
''    cbo������.Text = strTempNum
End Sub

Private Sub cbo������_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 And KeyAscii <> 46 And KeyAscii < Asc(0) Or KeyAscii > Asc(9) Then KeyAscii = 0
End Sub

Private Sub cbo������_Validate(Cancel As Boolean)
     DataChanged = True
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    If DataChanged Then
        On Error Resume Next
        mstr��鿪ʼʱ�� = CStr(dtp(0).Value)
        mstr������ʱ�� = CStr(dtp(1).Value)
        mstr����ѡ�� = cbo(0).Text
        mstr������ = CStr(txt������.Text)
        mlngCurNum = CLng(cbo������.ItemData(cbo������.ListIndex))
        
        mblnOK = True
        DataChanged = False
    End If
    Unload Me
End Sub

Private Sub cmdRef_Click()
    Call Init������(dtp(0).Value, dtp(1).Value)
End Sub



Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub dtp_Change(Index As Integer)
    Call zlControl.CboLocate(cbo(0), "�Զ���")
    DataChanged = True
End Sub

'''Private Sub Init������(ByVal str��鿪ʼʱ�� As String, ByVal str������ʱ�� As String)
'''    On Error GoTo errH
'''        Dim rs As ADODB.Recordset
'''        cbo������.Clear
'''        gstrSQL = "select distinct(��������) as ���� from ����������¼ A where A.����ʱ�� Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400 and A.�������� is not null order by A.��������"
'''        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Format(str��鿪ʼʱ��, "yyyy-mm-dd"), Format(str������ʱ��, "yyyy-mm-dd"))
'''        If rs.RecordCount > 0 Then
'''            If rs.BOF = False Then
'''                Call AddComboData(cbo������, rs, "����", "����", , False)
'''            End If
'''        Else
'''            cbo������.AddItem "1"
'''        End If
'''    Exit Sub
'''errH:
'''    Err.Clear
'''    Exit Sub
'''End Sub

'��ȡ������������Ϣ
Private Sub Init������(ByVal str��鿪ʼʱ�� As String, ByVal str������ʱ�� As String)
    On Error GoTo errH
        Dim rs As ADODB.Recordset
        Dim lngCount As Long '��¼����
        cbo������.Clear
        lngCount = 0
        gstrSQL = "select distinct(��������),Sum(A.��ֵ) as �ܿ۷���,Min(A.����ʱ��) as ���練��ʱ�� from ����������¼ A where A.����ʱ�� Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400 group by A.�������� order by A.�������� ASC"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Format(str��鿪ʼʱ��, "yyyy-mm-dd"), Format(str������ʱ��, "yyyy-mm-dd"))
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            cbo������.AddItem "����"
            Do Until rs.EOF
                If NVL(rs!��������, 0) = 0 Then
                    cbo������.AddItem "��" & NVL(rs!��������, 0) & "��-" & Format(NVL(rs!���練��ʱ��, Now()), "YYYY-MM-DD") & "(" & NVL(rs!�ܿ۷���, 0) & ")"
                    cbo������.ItemData(cbo������.NewIndex) = NVL(rs!��������, 0)
                End If
                rs.MoveNext
            Loop
            
            rs.MoveFirst
            Do Until rs.EOF
                    If lngCount >= 10 Then Exit Do
'                        Call AddComboData(cbo������, rs, "���練��ʱ��", "����", , False)
                        If NVL(rs!��������, 0) <> 0 Then
                            cbo������.AddItem "��" & NVL(rs!��������, 0) & "��-" & Format(NVL(rs!���練��ʱ��, Now()), "YYYY-MM-DD") & "(" & NVL(rs!�ܿ۷���, 0) & ")"
                            cbo������.ItemData(cbo������.NewIndex) = NVL(rs!��������, 0)
                        End If
                    lngCount = lngCount + 1
                    rs.MoveNext
            Loop
            cbo������.ListIndex = 0
        Else
            cbo������.AddItem "����"
            cbo������.ItemData(cbo������.NewIndex) = 0
            cbo������.ListIndex = 0
        End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err.Clear
    Exit Sub
End Sub

Private Sub dtp_Validate(Index As Integer, Cancel As Boolean)
    
    Dim strTempNum As String
    If cbo(0).Text = "�Զ���" Then
        strTempNum = CLng(cbo������.ItemData(cbo������.ListIndex))
        Call Init������(dtp(0).Value, dtp(1).Value)
        If strTempNum = "" Then Exit Sub
    End If
'    cbo������.Text = strTempNum
End Sub

Private Sub txt������_Change()
    DataChanged = True
End Sub

Private Sub cmd������_Click()
    On Error GoTo errH
    SelectDoctor
    Exit Sub
errH:
    Err.Clear
    Exit Sub
End Sub

Private Sub txt������_KeyPress(KeyAscii As Integer)
    If Trim(txt������.Text) = "" Then Exit Sub
    If InStr(1, "��'|[](){}*%", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        SelectDoctor txt������.Text
    End If
End Sub

'ѡ��ҽ��
Private Sub SelectDoctor(Optional strShortName As String = "")
    Dim rsTmp           As ADODB.Recordset
    Dim rsResult        As ADODB.Recordset
    Dim bytRet          As Byte
On Error GoTo errH
    gstrSQL = ""
    If strShortName <> "" Then
        gstrSQL = gstrSQL & vbCrLf & "select distinct (A.������) as ����,B.ID as id,B.��� From ����������¼ A,��Ա�� B"
        gstrSQL = gstrSQL & vbCrLf & "Where A.����ʱ�� Between To_Date([1], 'yyyy-mm-dd') And"
        gstrSQL = gstrSQL & vbCrLf & "To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400"
        gstrSQL = gstrSQL & vbCrLf & "And (B.���� like '%'||[3]||'%')"
        gstrSQL = gstrSQL & vbCrLf & "And A.������ = B.����"
        gstrSQL = gstrSQL & vbCrLf & "order by A.������"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Format(dtp(0).Value, "yyyy-mm-dd"), Format(dtp(1).Value, "yyyy-mm-dd"), UCase(strShortName))
        bytRet = ShowPubSelect(Me, txt������, 2, "���,1200,0,;����,1200,0,", Me.Name & "\������ѡ��", "����±���ѡ��һ������������", rsTmp, rsResult, 5000, 4500, True)
    Else
        Dim strTemp As String
        gstrSQL = gstrSQL & vbCrLf & "select distinct (A.������) as ����,B.ID as id,B.��� From ����������¼ A,��Ա�� B"
        gstrSQL = gstrSQL & vbCrLf & "Where A.����ʱ�� Between To_Date([1], 'yyyy-mm-dd') And"
        gstrSQL = gstrSQL & vbCrLf & "To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400"
        gstrSQL = gstrSQL & vbCrLf & "And A.������ = B.����"
        gstrSQL = gstrSQL & vbCrLf & "order by A.������"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Format(dtp(0).Value, "yyyy-mm-dd"), Format(dtp(1).Value, "yyyy-mm-dd"))
        bytRet = ShowPubSelect(Me, txt������, 2, "���,1200,0,;����,1200,0,", Me.Name & "\������ѡ��", "����±���ѡ��һ������������", rsTmp, rsResult, 5000, 4500, True)
        
    End If
     
    If rsResult Is Nothing Then
'        txt������.Text = ""
    ElseIf rsResult.EOF Or rsResult.BOF Then
        txt������.Text = ""
    Else
        rsResult.MoveFirst
        Do Until rsResult.EOF
            If Len(txt������.Text) = 0 Then
                txt������.Text = rsResult("����").Value
            Else
                If InStrRev(txt������.Text, rsResult("����").Value, -1) = 0 Then
                    txt������.Text = txt������.Text & "," & rsResult("����").Value
                End If
            End If
            rsResult.MoveNext
        Loop
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err.Clear
    Exit Sub
End Sub

Private Function GetBasePeriod(ByVal strMode As String, Optional ByVal bytFlag As Byte = 1) As String
    '******************************************************************************************************************
    '����:��ȡ����ʱ��
    '����:
    '����:
    '******************************************************************************************************************
    Dim intDay As Integer
    Dim varValue As Variant
    
    If Left(strMode, 3) = "�Զ���" Then
        '�Զ���:3,4
        varValue = Split(Mid(strMode, 5), ",")
        
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", Val(varValue(0)), zlDatabase.Currentdate), "yyyy-MM-dd") & " 00:00:00"
        Else
            If UBound(varValue) < 1 Then
                GetBasePeriod = Format(zlDatabase.Currentdate, "yyyy-MM-dd") & " 23:59:59"
            Else
                GetBasePeriod = Format(DateAdd("d", Val(varValue(1)), zlDatabase.Currentdate), "yyyy-MM-dd") & " 23:59:59"
            End If
        End If
            
        Exit Function
    End If
    
    Select Case strMode
    Case "��  ʱ"      '��ʱ
        GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"       '����,bytFlag=1,���ܿ�ʼʱ��,=2,���ܽ���ʱ��
        intDay = Weekday(CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD")))
        
        If intDay = 1 Then
            intDay = 7
        Else
            intDay = intDay - 1
        End If
        
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", 0 - intDay + 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", 7 - intDay, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM") & "-01 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM") & "-01"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"      '������
        Select Case Format(zlDatabase.Currentdate, "MM")
        Case "01", "02", "03"
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-03-31 23:59:59"
            End If
        Case "04", "05", "06"
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-04-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Case "07", "08", "09"
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-09-30 23:59:59"
            End If
        Case "10", "11", "12"
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-10-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End Select
    Case "������"      '������
        If Val(Format(zlDatabase.Currentdate, "MM")) < 7 Then
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Else
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End If
    Case "��  ��"   'ȫ��
        If bytFlag = 1 Then
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
        End If
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", -1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -3, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰһ��"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -7, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -15, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰһ��"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -30, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -60, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -90, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    
    Case "ǰ����"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -180, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    Case "ǰһ��"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -365, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    Case "ǰ����"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -365 * 2, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    End Select
    
End Function

Private Sub SetCob(ByVal lngCurNum As Long)
    Dim i As Integer
    For i = 0 To cbo������.ListCount - 1
        If cbo������.ItemData(i) = lngCurNum Then
            cbo������.ListIndex = i
            Exit Sub
        End If
    Next
End Sub
