VERSION 5.00
Begin VB.Form frmSetͭɽ�� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ����������"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5235
      TabIndex        =   5
      Top             =   2730
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4110
      TabIndex        =   4
      Top             =   2730
      Width           =   1100
   End
   Begin VB.Frame fraIC 
      Caption         =   "IC������"
      Height          =   2250
      Left            =   75
      TabIndex        =   0
      Top             =   150
      Width           =   6240
      Begin VB.CommandButton cmd�Ա��� 
         Caption         =   "ͬ���Ա���"
         Height          =   390
         Left            =   4650
         TabIndex        =   17
         Top             =   1740
         Width           =   1320
      End
      Begin VB.TextBox txtPass 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   1335
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   1095
         Width           =   1425
      End
      Begin VB.TextBox txtUser 
         Height          =   270
         Left            =   1335
         TabIndex        =   13
         Top             =   720
         Width           =   1425
      End
      Begin VB.CommandButton cmdδ���� 
         Caption         =   "��δ������Ŀ"
         Height          =   390
         Left            =   4650
         TabIndex        =   12
         Top             =   1200
         Width           =   1320
      End
      Begin VB.CommandButton cmd���ƿ� 
         Caption         =   "���ƿ����"
         Height          =   390
         Left            =   4650
         TabIndex        =   11
         Top             =   772
         Width           =   1320
      End
      Begin VB.CommandButton cmdҩƷ�� 
         Caption         =   "ҩƷ�����"
         Height          =   390
         Left            =   4650
         TabIndex        =   10
         Top             =   315
         Width           =   1320
      End
      Begin VB.CommandButton cmdҽԺ���� 
         Caption         =   "ҽԺ�������"
         Height          =   390
         Left            =   3120
         TabIndex        =   9
         Top             =   315
         Width           =   1320
      End
      Begin VB.CommandButton cmdסԺ���� 
         Caption         =   "סԺ���ָ���"
         Height          =   360
         Left            =   3120
         TabIndex        =   8
         Top             =   772
         Width           =   1350
      End
      Begin VB.CommandButton cmd���ﲡ�ָ��� 
         Caption         =   "���ﲡ�ָ���"
         Height          =   390
         Left            =   3120
         TabIndex        =   6
         Top             =   1200
         Width           =   1320
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   1335
         MaxLength       =   40
         TabIndex        =   1
         Text            =   "1"
         Top             =   315
         Width           =   360
      End
      Begin VB.Label Label4 
         Caption         =   "��ע�⣺ͬ���Ա��빦��ֻ�������Ա�����շ�ϸĿID��ͬ�ĵ�λ��"
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   570
         TabIndex        =   18
         Top             =   1725
         Width           =   3765
      End
      Begin VB.Label Label3 
         Caption         =   "����Ա����"
         Height          =   165
         Left            =   270
         TabIndex        =   16
         Top             =   1170
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "ҽ������Ա"
         Height          =   165
         Left            =   270
         TabIndex        =   15
         Top             =   780
         Width           =   1005
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�Ŵ���"
         Height          =   180
         Index           =   4
         Left            =   1740
         TabIndex        =   3
         Top             =   375
         Width           =   540
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��ǰ����(&D)"
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   2
         Top             =   375
         Width           =   990
      End
   End
   Begin VB.Label Label1 
      Caption         =   "���ָ���ʱ��ϳ��������ĵȺ�"
      Height          =   315
      Left            =   90
      TabIndex        =   7
      Top             =   2730
      Width           =   3630
   End
End
Attribute VB_Name = "frmSetͭɽ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mblnChange As Boolean
Private mlngIcdev As Long
Private st%


Dim strUser As String, strServer As String, strPass As String
Dim strFileName As String
Dim objStream As TextStream, lngReturn As Long
Dim objFileSystem As New FileSystemObject, lngID As Long
Dim strLine As String, rsBzgx As New ADODB.Recordset
Dim lngRount As Long
Dim lng�ɸ��� As Long

Private Const P_FILENAME = 167782162

 
Private Sub cmd���ﲡ�ָ���_Click()

On Error GoTo errHand
    strUser = "tsxyb"
    strServer = GetSetting("ZLSOFT", "ע����Ϣ\��½��Ϣ", "SERVER", "")
    strPass = "tsxyb"
    lngRount = 0
    If OraDataOpen(gcnͭɽ��, strServer, strUser, strPass, False) = False Then
        MsgBox "�޷����ӵ��м�⣬���鱣�ղ����Ƿ�������ȷ��", vbInformation, gstrSysName
        Exit Sub
    End If
    Label1.Caption = "��ʼ�������ﲡ��!"
    '������ļ�,����ʾ,�Ƿ񸲸�,���������
    DoEvents
    strFileName = App.Path & "\MZBZ.TXT"
    If Not objFileSystem.FileExists(strFileName) Then
        '�������ﲡ��
        Call ҽ����ʼ��_ͭɽ��
        If tsx_createparams(1024, 1024) = -1 Then
            MsgBox "�����ڴ�ռ�ʧ��!" & tsx_getlasterr(), vbInformation, gstrSysName
            Exit Sub
        End If
        
        lngReturn = tsx_setstringparam(P_FILENAME, 0, strFileName)
        If tsx_jkcall("D_MZBZB") = -1 Then
            MsgBox tsx_getlasterr()
            Exit Sub
        End If
        Label1.Caption = "���ﲡ���������!"
        lngReturn = tsx_destroyparams()
    Else
        '��ʾ��
        If MsgBox("�������ﲡ���ļ�,�Ƿ���������?", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            '����
            Call ҽ����ʼ��_ͭɽ��
            If tsx_createparams(1024, 1024) = -1 Then
                MsgBox "�����ڴ�ռ�ʧ��!" & tsx_getlasterr(), vbInformation, gstrSysName
                Exit Sub
            End If
            
            lngReturn = tsx_setstringparam(P_FILENAME, 0, strFileName)
            If tsx_jkcall("D_MZBZB") = -1 Then
                MsgBox tsx_getlasterr()
                Exit Sub
            End If
            Label1.Caption = "���ﲡ���������!"
            lngReturn = tsx_destroyparams()
        End If
    End If
    DoEvents

    strFileName = App.Path & "\MZBZ.TXT"
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForReading)
    
    Do Until objStream.AtEndOfLine
        strLine = objStream.ReadLine
        gstrSQL = "Select * from Mzbz where ���ֱ���='" & _
                    Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 1, 2), vbUnicode)) & "'"
        Call OpenRecordset_OtherBase(rsBzgx, "���ղ���", , gcnͭɽ��)
        
        If rsBzgx.EOF Then
            gstrSQL = "Select ���ղ���_ID.nextval as ID from dual"
            Set rsBzgx = zlDatabase.OpenSQLRecord(gstrSQL, "���ղ���")
            lngID = rsBzgx!ID
            gstrSQL = "Insert into MZBZ(ID,���ֱ���,��������,ƴ����) values(" & lngID & ",'" & _
                                        Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 1, 2), vbUnicode)) & "','" & _
                                        Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 3, 20), vbUnicode)) & "','" & _
                                        Mid(Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 24, 10), vbUnicode)), 2) & "')"
        End If
        gcnͭɽ��.Execute gstrSQL
        DoEvents
        lngRount = lngRount + 1
        Label1.Caption = "������" & lngRount & "����¼"
    Loop
    Label1.Caption = "������ﲡ�ָ���,���ι�����" & lngRount & "����¼"
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Sub

Private Sub cmdδ����_Click()

    Dim rsYpzlk As New ADODB.Recordset, rsSfxm As New ADODB.Recordset
    Dim lngRount As Long, strSqlTemp As String
On Error GoTo errHand

    cmdδ����.Enabled = False
    lngRount = 0
    
    strUser = "tsxyb"
    strServer = GetSetting("ZLSOFT", "ע����Ϣ\��½��Ϣ", "SERVER", "")
    strPass = "tsxyb"
    lngRount = 0
    
    If OraDataOpen(gcnͭɽ��, strServer, strUser, strPass, False) = False Then
        MsgBox "�޷����ӵ��м�⣬���鱣�ղ����Ƿ�������ȷ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strSqlTemp = "Delete ypzlk where substr(֧�����,1,1)<>'1' and �շ�ϸĿID is not null"
    gcnͭɽ��.Execute strSqlTemp
    
    strSqlTemp = "update ypzlk set ֧�����=substr(֧�����,2) where substr(֧�����,1,1)='1'"
    gcnͭɽ��.Execute strSqlTemp
    
    gstrSQL = "Select a.*," & _
              "Decode(Nvl(����ʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')), To_Date('3000-01-01', 'YYYY-MM-DD'), 0, 1) As ͣ�� " & _
              " From �շ�ϸĿ a"
    Set rsSfxm = zlDatabase.OpenSQLRecord(gstrSQL, "�շ�ϸĿ")
    strFileName = App.Path & "\δ������Ŀ.TXT"
    Call objFileSystem.CreateTextFile(strFileName)
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
    cmdδ����.Tag = 0
    
    Do Until rsSfxm.EOF
        gstrSQL = "Select * from ypzlk where �շ�ϸĿID=" & rsSfxm!ID
        Call OpenRecordset_OtherBase(rsYpzlk, "ypzlk", , gcnͭɽ��)
        If rsYpzlk.EOF Then
            'д�ļ�
            If rsSfxm!ͣ�� = 0 Then
                objStream.WriteLine rsSfxm!��� & "  " & rsSfxm!ID & "  " & rsSfxm!���� & "  " & rsSfxm!����
                lngRount = lngRount + 1
            End If
            gstrSQL = "ZL_����֧����Ŀ_Delete(" & rsSfxm!ID & "," & TYPE_ͭɽ�� & ")"
        Else
            gstrSQL = "ZL_����֧����Ŀ_Modify(" & rsSfxm!ID & "," & TYPE_ͭɽ�� & ",NUll,'" & _
                           rsYpzlk!�Ա��� & "','" & rsSfxm!���� & "','" & rsYpzlk!֧����� & "',1)"
                           
        End If
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����֧����Ŀ")
        rsSfxm.MoveNext
        DoEvents
        cmdδ����.Tag = cmdδ����.Tag + 1
        Label1.Caption = "�Ѽ��" & cmdδ����.Tag & "����¼"
    Loop
    
    objStream.WriteLine "��" & lngRount & "����¼δ����"
    objStream.Close
    Set objStream = Nothing
    
    If lng�ɸ��� = 3 Then
        cmdδ����.Enabled = False
        cmdҩƷ��.Enabled = True
        cmd���ƿ�.Enabled = True
        cmd�Ա���.Enabled = True
        lng�ɸ��� = 0
    End If
    
    Shell "notepad.exe " & strFileName, vbMaximizedFocus

    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdҩƷ��_Click()
    Dim str��� As String
On Error GoTo errHand


    strUser = "tsxyb"
    strServer = GetSetting("ZLSOFT", "ע����Ϣ\��½��Ϣ", "SERVER", "")
    strPass = "tsxyb"
    lngRount = 0
    
    
    If OraDataOpen(gcnͭɽ��, strServer, strUser, strPass, False) = False Then
        MsgBox "�޷����ӵ��м�⣬���鱣�ղ����Ƿ�������ȷ��", vbInformation, gstrSysName
        Exit Sub
    End If
    Label1.Caption = "��ʼ����ҩƷ��!"
    '������ļ�,����ʾ,�Ƿ񸲸�,���������
    DoEvents
    strFileName = App.Path & "\YPK.TXT"
    If Not objFileSystem.FileExists(strFileName) Then
        MsgBox "���ҽ��ǰ̨�����е���ҩƷ��,��������ΪYPK.TXT�ŵ�" & App.Path & "Ŀ¼�¡���ִ�д˹��ܣ�", vbInformation, gstrSysName
        Exit Sub
    Else
        If MsgBox("��ȷ�ϴ�ҽ��ǰ̨�������ļ�YPK.TXT�ŵ���" & App.Path & "Ŀ¼��", vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    lng�ɸ��� = lng�ɸ��� + 1
    cmdҩƷ��.Enabled = False
    
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForReading)
    
    Do Until objStream.AtEndOfLine
        strLine = objStream.ReadLine
        gstrSQL = "Select * from YPZLK where �Ա���='" & _
                    Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 1, 10), vbUnicode)) & "'"
        Call OpenRecordset_OtherBase(rsBzgx, "ҽԺ����", , gcnͭɽ��)
        
        If rsBzgx.EOF Then
            Select Case Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 21, 1), vbUnicode))
                Case 1
                    str��� = "1����"
                Case 2
                    str��� = "1����"
                Case Else
                    str��� = "1�Է�"
            End Select
            gstrSQL = "Insert into YPZLK(�Ա���,ҽ������,֧�����,������־) values('" & _
                                        Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 1, 10), vbUnicode)) & "','" & _
                                        Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 11, 10), vbUnicode)) & "','" & _
                                        str��� & "','" & _
                                        Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 22, 1), vbUnicode)) & "')"
        Else
            Select Case Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 21, 1), vbUnicode))
                Case 1
                    str��� = "1����"
                Case 2
                    str��� = "1����"
                Case Else
                    str��� = "1�Է�"
            End Select
            gstrSQL = "Update YPZLK Set ҽ������='" & Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 11, 10), vbUnicode)) & "'," & _
                                       "֧�����='" & str��� & "'," & _
                                       "������־='" & Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 22, 1), vbUnicode)) & "' " & _
                                       " Where �Ա���='" & Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 1, 10), vbUnicode)) & "'"
            
        End If
        gcnͭɽ��.Execute gstrSQL
        DoEvents
        lngRount = lngRount + 1
        Label1.Caption = "������" & lngRount & "����¼"
        
    Loop
    Label1.Caption = "���ҩƷ�����,���ι�����" & lngRount & "����¼"
    objStream.Close
    Set objStream = Nothing
    

    If lng�ɸ��� = 3 Then
        cmdδ����.Enabled = True
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdҽԺ����_Click()
On Error GoTo errHand
    strUser = "tsxyb"
    strServer = GetSetting("ZLSOFT", "ע����Ϣ\��½��Ϣ", "SERVER", "")
    strPass = "tsxyb"
    lngRount = 0
    If OraDataOpen(gcnͭɽ��, strServer, strUser, strPass, False) = False Then
        MsgBox "�޷����ӵ��м�⣬���鱣�ղ����Ƿ�������ȷ��", vbInformation, gstrSysName
        Exit Sub
    End If
    Label1.Caption = "��ʼ����ҽԺ����!"
    '������ļ�,����ʾ,�Ƿ񸲸�,���������
    DoEvents

    strFileName = App.Path & "\YYDA.TXT"
    If Not objFileSystem.FileExists(strFileName) Then
        '����סԺ����
        Call ҽ����ʼ��_ͭɽ��
        If tsx_createparams(1024, 1024) = -1 Then
            Label1.Caption = "�����ڴ�ռ�ʧ��!" & tsx_getlasterr()
            Exit Sub
        End If
        
        lngReturn = tsx_setstringparam(P_FILENAME, 0, strFileName)
        If tsx_jkcall("D_YYDA") = -1 Then
            MsgBox tsx_getlasterr()
            Exit Sub
        End If
        Label1.Caption = "ҽԺ�����������!��ʼд�뱾�����ݿ��У�"
        lngReturn = tsx_destroyparams()
    Else
        '��ʾ��
        If MsgBox("����ҽԺ�����ļ�,�Ƿ���������?", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            '����
            Call ҽ����ʼ��_ͭɽ��
            If tsx_createparams(1024, 1024) = -1 Then
                MsgBox "�����ڴ�ռ�ʧ��!" & tsx_getlasterr(), vbInformation, gstrSysName
                Exit Sub
            End If
            
            lngReturn = tsx_setstringparam(P_FILENAME, 0, strFileName)
            If tsx_jkcall("D_YYDA") = -1 Then
                MsgBox tsx_getlasterr()
                Exit Sub
            End If
            Label1.Caption = "ҽԺ�����������!��ʼд�뱾�����ݿ��У�"
            lngReturn = tsx_destroyparams()
        End If
    End If
    DoEvents
    strFileName = App.Path & "\YYDA.TXT"
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForReading)
    
    Do Until objStream.AtEndOfLine
        strLine = objStream.ReadLine
        gstrSQL = "Select * from YYDA where ҽԺ����='" & _
                    Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 1, 2), vbUnicode)) & "'"
        Call OpenRecordset_OtherBase(rsBzgx, "ҽԺ����", , gcnͭɽ��)
        
        If rsBzgx.EOF Then
            gstrSQL = "Insert into YYDA(ҽԺ����,ҽԺ����) values('" & _
                                        Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 1, 2), vbUnicode)) & "','" & _
                                        Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 3, 25), vbUnicode)) & "')"
        End If
        gcnͭɽ��.Execute gstrSQL
        DoEvents
        lngRount = lngRount + 1
        Label1.Caption = "������" & lngRount & "����¼"
        
    Loop
    Label1.Caption = "���ҽԺ�������,���ι�����" & lngRount & "����¼"
    objStream.Close
    Set objStream = Nothing
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Sub

Private Sub cmd���ƿ�_Click()
    Dim str��� As String
    
On Error GoTo errHand

    strUser = "tsxyb"
    strServer = GetSetting("ZLSOFT", "ע����Ϣ\��½��Ϣ", "SERVER", "")
    strPass = "tsxyb"
    lngRount = 0
    
    If OraDataOpen(gcnͭɽ��, strServer, strUser, strPass, False) = False Then
        MsgBox "�޷����ӵ��м�⣬���鱣�ղ����Ƿ�������ȷ��", vbInformation, gstrSysName
        Exit Sub
    End If
    Label1.Caption = "��ʼ����ҩƷ��!"
    '������ļ�,����ʾ,�Ƿ񸲸�,���������
    DoEvents
    strFileName = App.Path & "\ZLK.TXT"
    If Not objFileSystem.FileExists(strFileName) Then
        MsgBox "���ҽ��ǰ̨�����е���ҩƷ��,��������ΪZLK.TXT�ŵ�" & App.Path & "Ŀ¼�¡���ִ�д˹��ܣ�", vbInformation, gstrSysName
        Exit Sub
    Else
        If MsgBox("��ȷ�ϴ�ҽ��ǰ̨�������ļ�ZLK.TXT�ŵ���" & App.Path & "Ŀ¼��", vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Sub
        End If
        
    End If
    
    lng�ɸ��� = lng�ɸ��� + 1
    cmd���ƿ�.Enabled = False
    
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForReading)
    
    Do Until objStream.AtEndOfLine
        strLine = objStream.ReadLine
        gstrSQL = "Select * from YPZLK where �Ա���='" & _
                    Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 1, 10), vbUnicode)) & "'"
        Call OpenRecordset_OtherBase(rsBzgx, "ҽԺ����", , gcnͭɽ��)
        
        If rsBzgx.EOF Then
            Select Case Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 21, 1), vbUnicode))
                Case 1
                    str��� = "1����"
                Case 2
                    str��� = "1����"
                Case Else
                    str��� = "1�Է�"
            End Select
            gstrSQL = "Insert into YPZLK(�Ա���,ҽ������,֧�����,������־) values('" & _
                                        Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 1, 10), vbUnicode)) & "','" & _
                                        Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 11, 10), vbUnicode)) & "','" & _
                                        str��� & "','" & _
                                        Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 22, 1), vbUnicode)) & "')"
        Else
            Select Case Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 21, 1), vbUnicode))
                Case 1
                    str��� = "1����"
                Case 2
                    str��� = "1����"
                Case Else
                    str��� = "1�Է�"
            End Select
            gstrSQL = "Update YPZLK Set ҽ������='" & Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 11, 10), vbUnicode)) & "'," & _
                                       "֧�����='" & str��� & "'," & _
                                       "������־='" & Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 22, 1), vbUnicode)) & "' " & _
                                       " Where �Ա���='" & Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 1, 10), vbUnicode)) & "'"
        End If
        gcnͭɽ��.Execute gstrSQL
        DoEvents
        lngRount = lngRount + 1
        Label1.Caption = "������" & lngRount & "����¼"
        
    Loop
    Label1.Caption = "������ƿ����,���ι�����" & lngRount & "����¼"
    objStream.Close
    Set objStream = Nothing
    If lng�ɸ��� = 3 Then
        cmdδ����.Enabled = True
    End If

    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdסԺ����_Click()
On Error GoTo errHand

    strUser = "tsxyb"
    strServer = GetSetting("ZLSOFT", "ע����Ϣ\��½��Ϣ", "SERVER", "")
    strPass = "tsxyb"
    lngRount = 0
    If OraDataOpen(gcnͭɽ��, strServer, strUser, strPass, False) = False Then
        MsgBox "�޷����ӵ��м�⣬���鱣�ղ����Ƿ�������ȷ��", vbInformation, gstrSysName
        Exit Sub
    End If
    Label1.Caption = "��ʼ����סԺ����!"
    '������ļ�,����ʾ,�Ƿ񸲸�,���������
    DoEvents

    strFileName = App.Path & "\ICD10.TXT"
    If Not objFileSystem.FileExists(strFileName) Then
        '����סԺ����
        Call ҽ����ʼ��_ͭɽ��
        If tsx_createparams(1024, 1024) = -1 Then
            Label1.Caption = "�����ڴ�ռ�ʧ��!" & tsx_getlasterr()
            Exit Sub
        End If
        
        lngReturn = tsx_setstringparam(P_FILENAME, 0, strFileName)
        If tsx_jkcall("D_ICD10") = -1 Then
            MsgBox tsx_getlasterr()
            Exit Sub
        End If
        Label1.Caption = "סԺ�����������!��ʼд��HIS���ݿ��У�"
        lngReturn = tsx_destroyparams()
    Else
        '��ʾ��
        If MsgBox("����סԺ�����ļ�,�Ƿ���������?", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            '����
            Call ҽ����ʼ��_ͭɽ��
            If tsx_createparams(1024, 1024) = -1 Then
                MsgBox "�����ڴ�ռ�ʧ��!" & tsx_getlasterr(), vbInformation, gstrSysName
                Exit Sub
            End If
            
            lngReturn = tsx_setstringparam(P_FILENAME, 0, strFileName)
            If tsx_jkcall("D_ICD10") = -1 Then
                MsgBox tsx_getlasterr()
                Exit Sub
            End If
            Label1.Caption = "סԺ�����������!��ʼд��HIS���ݿ��У�"
            lngReturn = tsx_destroyparams()
        End If
    End If
    DoEvents
    strFileName = App.Path & "\ICD10.TXT"
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForReading)
    
    Do Until objStream.AtEndOfLine
        strLine = objStream.ReadLine
        gstrSQL = "Select * from ICD10 where ���ֱ���='" & _
                    Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 1, 10), vbUnicode)) & "'"
        Call OpenRecordset_OtherBase(rsBzgx, "���ղ���", , gcnͭɽ��)
        
        If rsBzgx.EOF Then
            gstrSQL = "Select ���ղ���_ID.nextval as ID from dual"
            Set rsBzgx = zlDatabase.OpenSQLRecord(gstrSQL, "���ղ���")
            lngID = rsBzgx!ID
            gstrSQL = "Insert into ICD10(ID,���ֱ���,��������,ƴ����) values(" & lngID & ",'" & _
                                        Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 1, 10), vbUnicode)) & "','" & _
                                        Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 11, 20), vbUnicode)) & "','" & _
                                        Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 31, 10), vbUnicode)) & "')"
        End If
        gcnͭɽ��.Execute gstrSQL
        DoEvents
        lngRount = lngRount + 1
        Label1.Caption = "������" & lngRount & "����¼"
        
    Loop
    Label1.Caption = "���סԺ���ָ���,���ι�����" & lngRount & "����¼"
    objStream.Close
    Set objStream = Nothing
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Sub

Private Sub cmd�Ա���_Click()
    
On Error GoTo errHand
    strUser = "tsxyb"
    strServer = GetSetting("ZLSOFT", "ע����Ϣ\��½��Ϣ", "SERVER", "")
    strPass = "tsxyb"
    lngRount = 0
    lng�ɸ��� = lng�ɸ��� + 1
    cmd�Ա���.Enabled = False
    
    If OraDataOpen(gcnͭɽ��, strServer, strUser, strPass, False) = False Then
        MsgBox "�޷����ӵ��м�⣬����ִ�г�ʼ���ű���", vbInformation, gstrSysName
        'Exit Function
    Else
        gstrSQL = "Update ypzlk set �շ�ϸĿID=�Ա��� Where �շ�ϸĿID is null"
        gcnͭɽ��.Execute gstrSQL
    End If
    If lng�ɸ��� = 3 Then
        cmdδ����.Enabled = True
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    lng�ɸ��� = 0
    cmdδ����.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If IsValid = False Then Exit Sub
    If SaveData = False Then Exit Sub
    
    mblnOK = True
    mblnChange = False
    Unload Me
End Sub

Private Function IsValid() As Boolean
    Dim lngCount As Long
    Dim strTitle As String
    
    If Not IsNumeric(txtEdit(4).Text) Then
        MsgBox "�뽫���ں�����������Ϣ", vbInformation, gstrSysName
        Exit Function
    End If
    
    On Error Resume Next
'    mlngIcdev = init_com(txtEdit(4).Text - 1) 'Init COM2
'    If mlngIcdev <> 0 Then
'        If MsgBox("���ڳ�ʼ��ʧ�ܣ����鴮�ڡ��Ƿ�������棿", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
'            txtEdit(4).SetFocus
'            Exit Function
'        End If
'    End If
'    st = close_com()
    IsValid = True
End Function

Public Function ��������() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim str����ֵ As String
    
    mblnOK = False
    On Error Resume Next
    txtEdit(4).Text = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "��ǰʹ�õĴ���") + 1
    
    strUser = "tsxyb"
    strServer = GetSetting("ZLSOFT", "ע����Ϣ\��½��Ϣ", "SERVER", "")
    strPass = "tsxyb"
    lngRount = 0
    If OraDataOpen(gcnͭɽ��, strServer, strUser, strPass, False) = False Then
        MsgBox "�޷����ӵ��м�⣬����ִ�г�ʼ���ű���", vbInformation, gstrSysName
        'Exit Function
    End If
        
    gstrSQL = "Select * from czry where P_GLY=1"
    Call OpenRecordset_OtherBase(rsTemp, "czry", , gcnͭɽ��)
    If rsTemp.EOF = False Then
        txtUser.Text = rsTemp!P_RYH
        txtPass.Text = rsTemp!P_MM
    End If
    
    mblnChange = False
    frmSetͭɽ��.Show vbModal, frmҽ�����
    �������� = mblnOK
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SaveData() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lst As ListItem
    
    '����ǰʹ�õĴ���д��ע���֮��
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "��ǰʹ�õĴ���", CStr(txtEdit(4).Text - 1)
    If Len(Trim(txtUser.Text)) > 0 Then
        
        gstrSQL = "Delete czry where P_GLY=1"
        gcnͭɽ��.Execute gstrSQL
        gstrSQL = "Insert into czry(P_RYH,P_XM,P_MM,P_GLY) values('" & Trim(txtUser.Text) & _
                "','����Ա','" & Trim(txtPass.Text) & "',1)"
        gcnͭɽ��.Execute gstrSQL
    
    End If
    
    
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
    If Index = 4 Then
        If CheckIsInclude(UCase(Chr(KeyAscii)), "������") = True Then KeyAscii = 0
    End If
End Sub

Private Function CheckIsInclude(strSource As String, strTarge As String) As Boolean
    '���strSource�е�ÿһ���ַ��Ƿ���strTarge��
    Dim i As Long
    CheckIsInclude = False
    
    Select Case strTarge
    Case "����"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "ʱ��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+-_)(*&^%$#@!`~"
    Case "����ʱ��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+_)(*&^%$#@!`~"
    Case "����"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "С��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "������"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "��С��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "�ɴ�ӡ�ַ�"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/."":;|\=+-_)(*&^%$#@!`~0123456789"
    End Select
    For i = 1 To Len(strSource)
        If InStr(strTarge, Mid(strSource, i, 1)) <= 0 Then Exit Function
    Next
    CheckIsInclude = True
End Function

Private Sub txtEdit_Validate(Index As Integer, Cancel As Boolean)
    If Index = 4 Then
        If Not IsNumeric(txtEdit(4).Text) Then
            MsgBox "�뽫���ں�����������Ϣ", vbInformation, gstrSysName
        End If
    End If
End Sub




