VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFilmFontSize 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ƭ��������"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8070
   Icon            =   "frmFilmFontSize.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CmdAdd 
      Caption         =   "����(&A)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   120
      TabIndex        =   13
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton CmdModify 
      Caption         =   "�޸�(&M)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1470
      TabIndex        =   12
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "ɾ��(&D)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2790
      TabIndex        =   11
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�(&Q)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   6840
      TabIndex        =   7
      Top             =   4680
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   2820
      Width           =   7875
      Begin VB.CheckBox chkFontTransparent 
         Caption         =   "����͸��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   16
         Top             =   233
         Width           =   1575
      End
      Begin VB.CheckBox chkFontShadow 
         Caption         =   "������Ӱ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   15
         Top             =   593
         Width           =   2055
      End
      Begin VB.CheckBox chkFontInverse 
         Caption         =   "���巴ɫ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   14
         Top             =   998
         Width           =   1455
      End
      Begin VB.TextBox txtPostureFontSize 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2070
         TabIndex        =   9
         Top             =   968
         Width           =   1665
      End
      Begin VB.CheckBox chkPostureAutoZoom 
         Caption         =   "��λ��ע��ͼ������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   8
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CheckBox ChkAutoZoom 
         Caption         =   "��Ϣ��ͼ������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   6
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox TxtFontSize 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2070
         TabIndex        =   5
         Top             =   578
         Width           =   1665
      End
      Begin VB.TextBox txtImageType 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2070
         TabIndex        =   3
         Top             =   218
         Width           =   1665
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "��λ��ע�����С:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   10
         Top             =   1020
         Width           =   1785
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "��Ϣ�����С:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   4
         Top             =   630
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ӱ������:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   2
         Top             =   270
         Width           =   945
      End
   End
   Begin MSComctlLib.ListView LivMain 
      Height          =   2715
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   4789
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmFilmFontSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CmdAdd_Click()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    '���������Ч��
    If Len(Trim(Me.txtImageType)) < 1 Then
        MsgBox "������Ӱ�����", vbInformation, gstrSysName
        Me.txtImageType.SetFocus
        Exit Sub
    End If
    If Len(Trim(Me.TxtFontSize)) < 1 Then
        MsgBox "�������ӡ����Ϣ�����С", vbInformation, gstrSysName
        Me.TxtFontSize.SetFocus
        Exit Sub
    End If
    If Len(Trim(Me.txtPostureFontSize)) < 1 Then
        MsgBox "�������ӡ����λ��ע�����С", vbInformation, gstrSysName
        Me.txtPostureFontSize.SetFocus
        Exit Sub
    End If
    On Error GoTo errh
    
    '��ѯȷ��Ӱ������Ƿ��Ѿ�����
    If blLocalRun = True Then
        strSQL = "select count(*) as �ܼ� from DICOM��Ƭ��ӡ���� where Ӱ����� = """ & Me.txtImageType & """"
        Set rsTmp = cnAccess.Execute(strSQL)
    Else
        strSQL = "select count(*) as �ܼ� from Ӱ��Ƭ��ӡ���� where Ӱ����� = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(UCase(Me.txtImageType)))
    End If
    If rsTmp("�ܼ�") > 0 Then
        MsgBox "�����ӵ�Ӱ������Ѵ��ڣ����������룡", vbInformation, gstrSysName
        Me.txtImageType.SetFocus
        Exit Sub
    End If
    
    '�����µ���������
    If blLocalRun = True Then
        strSQL = "insert into DICOM��Ƭ��ӡ���� (Ӱ�����,�����С,�Ƿ���ͼ������,��λ��ע�����С,��λ��ע��ͼ������) values (""" & _
                 UCase(Me.txtImageType) & """,""" & Me.TxtFontSize & """," & IIf(Me.ChkAutoZoom.Value = 1, True, False) & _
                 ",""" & Me.txtPostureFontSize & """," & IIf(Me.chkPostureAutoZoom.Value = 1, True, False) & ")"
        cnAccess.Execute strSQL
    Else
        strSQL = "ZL_Ӱ��Ƭ��ӡ����_INSERT('" & UCase(Me.txtImageType) & "'," & Me.TxtFontSize & _
                 "," & IIf(Me.ChkAutoZoom.Value = 1, 1, 0) & ",'" & Me.txtPostureFontSize & "'," & _
                 IIf(Me.chkPostureAutoZoom.Value = 1, 1, 0) & "," & IIf(Me.chkFontInverse.Value = 1, 1, 0) & _
                 "," & IIf(Me.chkFontShadow.Value = 1, 1, 0) & "," & IIf(Me.chkFontTransparent.Value = 1, 1, 0) & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    
    '���¼����Ϣ������Ĭ��ֵ
    Me.txtImageType.Text = ""
    Me.TxtFontSize.Text = ""
    Me.ChkAutoZoom.Value = 0
    Me.txtPostureFontSize.Text = ""
    Me.chkPostureAutoZoom.Value = 1
    Me.chkFontInverse.Value = 0
    Me.chkFontShadow = 0
    Me.chkFontTransparent.Value = 1
    
    '�б�������ʾ
    LoadDate
    Me.LivMain.ListItems(Me.LivMain.ListItems.Count).Selected = True
    
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
End Sub

Private Sub CmdDelete_Click()
    Dim strSQL As String
    Dim i As Integer
    If Me.LivMain.ListItems.Count < 1 Then Exit Sub
    If Len(Trim(Me.LivMain.ListItems(Me.LivMain.SelectedItem.Index).Text)) < 1 Then
        MsgBox "��ѡ��һ��Ҫɾ����Ӱ�����", vbInformation, gstrSysName
        Me.txtImageType.SetFocus
        Exit Sub
    End If
    
    On Error GoTo errh
    i = Me.LivMain.SelectedItem.Index
    
    If blLocalRun = True Then
        strSQL = "delete from DICOM��Ƭ��ӡ���� where Ӱ����� = '" & Me.LivMain.ListItems(Me.LivMain.SelectedItem.Index).Text & "'"
        cnAccess.Execute strSQL
    Else
        strSQL = "ZL_Ӱ��Ƭ��ӡ����_DELETE('" & Me.LivMain.ListItems(Me.LivMain.SelectedItem.Index).Text & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    
    Me.txtImageType = ""
    Me.TxtFontSize = ""
    Me.ChkAutoZoom.Value = 0
    Me.txtPostureFontSize.Text = ""
    Me.chkPostureAutoZoom.Value = 1
    Me.chkFontInverse.Value = 0
    Me.chkFontShadow = 0
    Me.chkFontTransparent.Value = 1
    
    LoadDate
    If i > 1 Then
        Me.LivMain.ListItems(i - 1).Selected = True
    End If
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdModify_Click()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    '���������Ч��
    If Me.LivMain.ListItems.Count < 1 Then Exit Sub
    If Len(Trim(Me.txtImageType)) < 1 Then
        MsgBox "��ѡ��һ��Ӱ�����", vbInformation, gstrSysName
        Me.txtImageType.SetFocus
        Exit Sub
    End If
    If Len(Trim(Me.TxtFontSize)) < 1 Then
        MsgBox "�������ӡ����Ϣ�����С", vbInformation, gstrSysName
        Me.TxtFontSize.SetFocus
        Exit Sub
    End If
    If Len(Trim(Me.txtPostureFontSize)) < 1 Then
        MsgBox "�������ӡ����λ��ע�����С", vbInformation, gstrSysName
        Me.txtPostureFontSize.SetFocus
        Exit Sub
    End If
    On Error GoTo errh
    
    '�޸Ľ�Ƭ������Ϣ
    If blLocalRun = True Then
        strSQL = "update DICOM��Ƭ��ӡ���� set Ӱ����� = '" & UCase(Me.txtImageType) & "',�����С ='" & TxtFontSize & _
                 "',�Ƿ���ͼ������ = " & IIf(Me.ChkAutoZoom.Value = 1, True, False) & ",��λ��ע�����С = '" & txtPostureFontSize & _
                 "',��λ��ע��ͼ������ = " & IIf(Me.chkPostureAutoZoom.Value = 1, True, False) & " where Ӱ����� = '" & _
                 Me.LivMain.ListItems(Me.LivMain.SelectedItem.Index).Text & "'"
        cnAccess.Execute strSQL
    Else
        strSQL = "ZL_Ӱ��Ƭ��ӡ����_UPDATE('" & Me.LivMain.ListItems(Me.LivMain.SelectedItem.Index).Text & _
                 "','" & UCase(Me.txtImageType) & "'," & Me.TxtFontSize & "," & IIf(Me.ChkAutoZoom.Value = 1, 1, 0) & _
                 ",'" & Me.txtPostureFontSize & "'," & IIf(Me.chkPostureAutoZoom.Value = 1, 1, 0) & "," & _
                  IIf(Me.chkFontInverse.Value = 1, 1, 0) & "," & IIf(Me.chkFontShadow.Value = 1, 1, 0) & "," & _
                  IIf(Me.chkFontTransparent.Value = 1, 1, 0) & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    
    i = Me.LivMain.SelectedItem.Index
    LoadDate
    Me.LivMain.ListItems(i).Selected = True
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
End Sub

Private Sub Form_Load()
    InitLivHead
    LoadDate
End Sub
Sub InitLivHead()
    Dim chColHeader As ColumnHeader
    '��ʹ���б�ͷ
    With Me.LivMain
        Set chColHeader = .ColumnHeaders.Add(, "A", "Ӱ�����")
        chColHeader.width = 900
        Set chColHeader = .ColumnHeaders.Add(, "B", "��Ϣ����")
        chColHeader.width = 900
        Set chColHeader = .ColumnHeaders.Add(, "C", "��Ϣ��ͼ����")
        chColHeader.width = 1200
        Set chColHeader = .ColumnHeaders.Add(, "D", "��λ����")
        chColHeader.width = 900
        Set chColHeader = .ColumnHeaders.Add(, "E", "��λ��ͼ����")
        chColHeader.width = 1200
        Set chColHeader = .ColumnHeaders.Add(, "F", "����͸��")
        chColHeader.width = 900
        Set chColHeader = .ColumnHeaders.Add(, "G", "������Ӱ")
        chColHeader.width = 900
        Set chColHeader = .ColumnHeaders.Add(, "H", "���巴ɫ")
        chColHeader.width = 900
    End With
End Sub
Sub LoadDate()
    '��������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim objItem As ListItem
    
    Me.LivMain.ListItems.Clear
    
    If blLocalRun = True Then
        strSQL = "select Ӱ����� , �����С, �Ƿ���ͼ������,��λ��ע�����С,��λ��ע��ͼ������ from DICOM��Ƭ��ӡ����"
        Set rsTmp = cnAccess.Execute(strSQL)
    Else
        strSQL = "select Ӱ����� , �����С, �Ƿ���ͼ������,��λ��ע�����С,��λ��ע��ͼ������,���巴ɫ,������Ӱ,���屳��͸�� from Ӱ��Ƭ��ӡ����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    End If
    Do Until rsTmp.EOF
        With Me.LivMain
            Set objItem = .ListItems.Add(, "A" & rsTmp("Ӱ�����"), rsTmp("Ӱ�����"))
            objItem.SubItems(1) = Nvl(rsTmp("�����С"))
            objItem.SubItems(2) = IIf(Nvl(rsTmp("�Ƿ���ͼ������"), 0), "��", "")
            objItem.SubItems(3) = Nvl(rsTmp("��λ��ע�����С"))
            objItem.SubItems(4) = IIf(Nvl(rsTmp("��λ��ע��ͼ������"), 0), "��", "")
            objItem.SubItems(5) = IIf(Nvl(rsTmp("���屳��͸��"), 1), "��", "")
            objItem.SubItems(6) = IIf(Nvl(rsTmp("������Ӱ"), 0), "��", "")
            objItem.SubItems(7) = IIf(Nvl(rsTmp("���巴ɫ"), 0), "��", "")
        End With
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    If LivMain.ListItems.Count >= 1 Then
        Call LivMain_ItemClick(LivMain.ListItems(1))
    End If
End Sub

Private Sub LivMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Me.txtImageType = Item.Text
    Me.TxtFontSize = Item.SubItems(1)
    Me.ChkAutoZoom.Value = IIf(Item.SubItems(2) <> "", 1, 0)
    Me.txtPostureFontSize = Item.SubItems(3)
    Me.chkPostureAutoZoom.Value = IIf(Item.SubItems(4) <> "", 1, 0)
    Me.chkFontTransparent.Value = IIf(Item.SubItems(5) <> "", 1, 0)
    Me.chkFontShadow.Value = IIf(Item.SubItems(6) <> "", 1, 0)
    Me.chkFontInverse.Value = IIf(Item.SubItems(7) <> "", 1, 0)
End Sub

Private Sub TxtFontSize_GotFocus()
    Me.TxtFontSize.SelStart = 0
    Me.TxtFontSize.SelLength = Len(Me.TxtFontSize)
End Sub

Private Sub TxtFontSize_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtImageType_GotFocus()
    Me.txtImageType.SelStart = 0
    Me.txtImageType.SelLength = Len(Me.txtImageType)
End Sub

Private Sub txtPostureFontSize_GotFocus()
    Me.txtPostureFontSize.SelStart = 0
    Me.txtPostureFontSize.SelLength = Len(Me.txtPostureFontSize)
End Sub

Private Sub txtPostureFontSize_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
