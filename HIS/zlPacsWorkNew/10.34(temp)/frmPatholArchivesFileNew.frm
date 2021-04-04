VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatholArchivesFileNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7215
   Icon            =   "frmPatholArchivesFileNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picShow 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   6975
      TabIndex        =   28
      Top             =   4200
      Width           =   6975
      Begin VB.TextBox txtShow 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   120
         Width           =   6735
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         DrawMode        =   1  'Blackness
         FillColor       =   &H000000FF&
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   6975
      End
   End
   Begin VB.CheckBox chkContinue 
      Caption         =   "ȷ�������ִ�е�ǰ����"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3720
      Width           =   2415
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "ȷ ��&S)"
      Height          =   400
      Left            =   4560
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ ��(&C)"
      Height          =   400
      Left            =   5880
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.TextBox txtPlace 
         Height          =   300
         Left            =   4440
         TabIndex        =   30
         Top             =   2160
         Width           =   2295
      End
      Begin VB.ComboBox cbxDrawer 
         Height          =   300
         Left            =   960
         TabIndex        =   27
         Top             =   2160
         Width           =   2295
      End
      Begin VB.ComboBox cbxBox 
         Height          =   300
         Left            =   4440
         TabIndex        =   26
         Top             =   1680
         Width           =   2295
      End
      Begin VB.ComboBox cbxRoom 
         Height          =   300
         Left            =   960
         TabIndex        =   25
         Top             =   1680
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker dtpArchivesCreate 
         Height          =   300
         Left            =   4440
         TabIndex        =   21
         Top             =   3120
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   154796035
         CurrentDate     =   40865
      End
      Begin VB.TextBox txtCreateUser 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   960
         TabIndex        =   19
         Top             =   3120
         Width           =   2295
      End
      Begin VB.TextBox txtArchivesDescription 
         Height          =   300
         Left            =   960
         TabIndex        =   17
         Top             =   2640
         Width           =   5775
      End
      Begin MSComCtl2.DTPicker dtpArchivesEnd 
         Height          =   300
         Left            =   4440
         TabIndex        =   14
         Top             =   1200
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   154796035
         CurrentDate     =   40864
      End
      Begin MSComCtl2.DTPicker dtpArchivesStart 
         Height          =   300
         Left            =   960
         TabIndex        =   12
         Top             =   1200
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   154796035
         CurrentDate     =   40864
      End
      Begin VB.ComboBox cbxArchivesClass 
         Height          =   300
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtArchivesStudyArea 
         Height          =   300
         Left            =   960
         TabIndex        =   8
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtArchivesCode 
         Height          =   300
         Left            =   4440
         TabIndex        =   6
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtArchivesName 
         Height          =   300
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label14 
         Caption         =   "��ϸ��ַ��"
         Height          =   255
         Left            =   3600
         TabIndex        =   33
         Top             =   2205
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "�������룺"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   2200
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "������ţ�"
         Height          =   255
         Left            =   3600
         TabIndex        =   31
         Top             =   1740
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   6760
         TabIndex        =   24
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label10 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3280
         TabIndex        =   22
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "�������ڣ�"
         Height          =   255
         Left            =   3600
         TabIndex        =   20
         Top             =   3180
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "�� �� �ˣ�"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   3180
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "����˵����"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2670
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "�������䣺"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1740
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "�������ڣ�"
         Height          =   255
         Left            =   3600
         TabIndex        =   13
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "��ʼ���ڣ�"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "�������ࣺ"
         Height          =   255
         Left            =   3600
         TabIndex        =   9
         Top             =   780
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "��鷶Χ��"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   780
         Width           =   975
      End
      Begin VB.Label labArchivesCode 
         Caption         =   "������ţ�"
         Height          =   255
         Left            =   3600
         TabIndex        =   5
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "�������ƣ�"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmPatholArchivesFileNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mufgParentGrid As ucFlexGrid

Private mblnIsSucceed As Boolean
Private mblnIsUpdate As Boolean

Private mrsArchivesClass As ADODB.Recordset




Public Function ShowAddArchivesFileWindow(ufgParentGrid As ucFlexGrid, owner As Form) As Boolean
'��ʾ�����ļ���������
    Dim curDate As Date
    
    ShowAddArchivesFileWindow = False
    
    Set mufgParentGrid = ufgParentGrid
    
    Me.Caption = "��������"
    mblnIsUpdate = False
    mblnIsSucceed = False
    
    curDate = zlDatabase.Currentdate
    
    dtpArchivesStart.value = curDate
    dtpArchivesEnd.value = curDate
    dtpArchivesCreate.value = curDate
    txtCreateUser.Text = UserInfo.����
    
    
    Call CloseProcessHint
    
    chkContinue.value = False
    chkContinue.Visible = True
    
    Call Me.Show(1, owner)
    
    ShowAddArchivesFileWindow = mblnIsSucceed

End Function



Public Function ShowUpdateArchivesFileWindow(ufgParentGrid As ucFlexGrid, owner As Form) As Boolean
'��ʾ������´���
    ShowUpdateArchivesFileWindow = False
    
    Set mufgParentGrid = ufgParentGrid
        
    Me.Caption = "���µ���"
    mblnIsUpdate = True
    mblnIsSucceed = False
        
    Call CloseProcessHint
    
    Call ConfigUpdateFace
    
    chkContinue.value = False
    chkContinue.Visible = False

    
    Call Me.Show(1, owner)
    
    ShowUpdateArchivesFileWindow = mblnIsSucceed
End Function



Public Sub ConfigUpdateFace()
On Error Resume Next
    Dim strPlace As String
    
    With mufgParentGrid
        txtArchivesName.Text = .Text(.SelectionRow, gstrPatholCol_��������)
        txtArchivesCode.Text = .Text(.SelectionRow, gstrPatholCol_�������)
        txtArchivesStudyArea.Text = .Text(.SelectionRow, gstrPatholCol_��鷶Χ)
        dtpArchivesStart.value = .Text(.SelectionRow, gstrPatholCol_��ʼ����)
        dtpArchivesEnd.value = .Text(.SelectionRow, gstrPatholCol_��������)
        dtpArchivesCreate.value = .Text(.SelectionRow, gstrPatholCol_��������)
        txtCreateUser.Text = .Text(.SelectionRow, gstrPatholCol_������)
        txtArchivesDescription.Text = .Text(.SelectionRow, gstrPatholCol_����˵��)
        
        '��ȡ��������
        cbxArchivesClass.Text = .Text(.SelectionRow, gstrPatholCol_��������)


        '��ȡ���λ��
        cbxRoom.Text = .Text(.SelectionRow, gstrPatholCol_��������)
        cbxBox.Text = .Text(.SelectionRow, gstrPatholCol_�������)
        cbxDrawer.Text = .Text(.SelectionRow, gstrPatholCol_��������)
        txtPlace.Text = .Text(.SelectionRow, gstrPatholCol_��ϸ��ַ)
    End With
    
err.Clear
    
End Sub



Private Sub ShowProcessHint(ByVal strHint As String)
'��ʾ������Ϣ
    txtShow.Text = strHint
End Sub


Private Sub CloseProcessHint()
'�رմ�����ʾ
    txtShow.Text = ""
End Sub





Private Sub cmdCancel_Click()
On Error GoTo errHandle
    Call Unload(Me)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Function CheckArchivesFileDataIsValid() As String
    CheckArchivesFileDataIsValid = ""
    
    '��鵵�������Ƿ�Ϊ��
    If Trim(txtArchivesName.Text) = "" Then
        CheckArchivesFileDataIsValid = "�������Ʋ���Ϊ�ա�"
        
        Call txtArchivesName.SetFocus
        Exit Function
    End If
    
    
    '��鵵�������Ƿ�Ϊ��
    If Trim(cbxArchivesClass.Text) = "" Then
        CheckArchivesFileDataIsValid = "�������಻��Ϊ�ա�"
        
        Call cbxArchivesClass.SetFocus
        Exit Function
    End If
    
    
    
    '��鵵�������Ƿ��ظ�
    Dim i As Integer
    For i = 1 To mufgParentGrid.GridRows - 1
        If Not mufgParentGrid.RowHidden(i) Then
            If Not mblnIsUpdate Then
                If mufgParentGrid.Text(i, gstrPatholCol_��������) = txtArchivesName.Text Then
                    CheckArchivesFileDataIsValid = "���������ظ���"
    
                    Call txtArchivesName.SetFocus
                    Exit Function
                End If
            Else
                If Not mufgParentGrid.SelectionRow = i Then
                    If mufgParentGrid.Text(i, gstrPatholCol_��������) = txtArchivesName.Text Then
                        CheckArchivesFileDataIsValid = "���������ظ���"
    
                        Call txtArchivesName.SetFocus
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function



Private Sub NewArchivesInf()
'�����ݿ�������������¼
'���ص���ID
    Dim strSql As String
    Dim rsReture As ADODB.Recordset
    Dim lngNewRecordIndex As Long
    Dim lngNewArchivesId As Long
    
    

    strSql = "select Zl_������_�����ļ�����([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13]) as ����ֵ from dual"
                                
    Set rsReture = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                txtArchivesName.Text, _
                                txtArchivesCode.Text, _
                                txtArchivesStudyArea.Text, _
                                Val(cbxArchivesClass.ItemData(cbxArchivesClass.ListIndex)), _
                                CDate(dtpArchivesStart.value), _
                                CDate(dtpArchivesEnd.value), _
                                CDate(dtpArchivesCreate.value), _
                                UserInfo.����, _
                                cbxRoom.Text, _
                                cbxBox.Text, _
                                cbxDrawer.Text, _
                                txtPlace.Text, _
                                txtArchivesDescription.Text)
                                
    If rsReture.RecordCount <= 0 Then
        Call err.Raise(0, "NewArchivesFile", "δ�ɹ���ȡ������ĵ���ID,���β���ʧ�ܡ�")
        Exit Sub
    End If
    
    
    With mufgParentGrid
        lngNewRecordIndex = .NewRow
        
        .Text(lngNewRecordIndex, gstrPatholCol_ID) = Nvl(rsReture!����ֵ)
        .Text(lngNewRecordIndex, gstrPatholCol_��������) = txtArchivesName.Text
        .Text(lngNewRecordIndex, gstrPatholCol_�������) = txtArchivesCode.Text
        .Text(lngNewRecordIndex, gstrPatholCol_��鷶Χ) = txtArchivesStudyArea.Text
        .Text(lngNewRecordIndex, gstrPatholCol_��ʼ����) = Format(dtpArchivesStart.value, gstrDateFormat)
        .Text(lngNewRecordIndex, gstrPatholCol_��������) = Format(dtpArchivesEnd.value, gstrDateFormat)
        .Text(lngNewRecordIndex, gstrPatholCol_��������) = Format(dtpArchivesCreate.value, gstrDateFormat)
        .Text(lngNewRecordIndex, gstrPatholCol_����˵��) = txtArchivesDescription.Text
        .Text(lngNewRecordIndex, gstrPatholCol_������) = txtCreateUser.Text
        .Text(lngNewRecordIndex, gstrPatholCol_����״̬) = "δ�鵵"
        .Text(lngNewRecordIndex, gstrPatholCol_��������) = cbxRoom.Text
        .Text(lngNewRecordIndex, gstrPatholCol_�������) = cbxBox.Text
        .Text(lngNewRecordIndex, gstrPatholCol_��������) = cbxDrawer.Text
        .Text(lngNewRecordIndex, gstrPatholCol_��ϸ��ַ) = txtPlace.Text
        
        .Text(lngNewRecordIndex, gstrPatholCol_��������) = cbxArchivesClass.Text
        
        mrsArchivesClass.Filter = "��������='" & cbxArchivesClass.Text & "'"
        If mrsArchivesClass.RecordCount > 0 Then
            .Text(lngNewRecordIndex, gstrPatholCol_��������) = Val(Nvl(mrsArchivesClass!��������))
            .Text(lngNewRecordIndex, gstrPatholCol_��������) = Nvl(mrsArchivesClass!��������)
        End If
        
        Call .LocateRow(lngNewRecordIndex)
        
    End With
End Sub




Private Sub UpdateArchivesInf()
'�������ݿ��еĵ�����Ϣ
    Dim strSql As String
    Dim lngCurArchivesId As Long
    Dim lngUpdateRecordIndex As Long
    


    lngCurArchivesId = mufgParentGrid.KeyValue(mufgParentGrid.SelectionRow)
    
    strSql = "Zl_������_�����ļ�����(" & lngCurArchivesId & ",'" & txtArchivesName.Text & "','" & txtArchivesCode.Text & "','" & txtArchivesStudyArea.Text & "'," & _
                                cbxArchivesClass.ItemData(cbxArchivesClass.ListIndex) & "," & To_Date(dtpArchivesStart.value) & "," & _
                                To_Date(dtpArchivesEnd.value) & ",'" & cbxRoom.Text & "','" & cbxBox.Text & "','" & cbxDrawer.Text & "','" & txtPlace.Text & "','" & txtArchivesDescription.Text & "')"
                                
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    
    lngUpdateRecordIndex = mufgParentGrid.SelectionRow
    
    With mufgParentGrid
        .Text(lngUpdateRecordIndex, gstrPatholCol_��������) = txtArchivesName.Text
        .Text(lngUpdateRecordIndex, gstrPatholCol_�������) = txtArchivesCode.Text
        .Text(lngUpdateRecordIndex, gstrPatholCol_��鷶Χ) = txtArchivesStudyArea.Text
        .Text(lngUpdateRecordIndex, gstrPatholCol_��ʼ����) = Format(dtpArchivesStart.value, gstrDateFormat)
        .Text(lngUpdateRecordIndex, gstrPatholCol_��������) = Format(dtpArchivesEnd.value, gstrDateFormat)
        .Text(lngUpdateRecordIndex, gstrPatholCol_����˵��) = txtArchivesDescription.Text
        .Text(lngUpdateRecordIndex, gstrPatholCol_��������) = cbxRoom.Text
        .Text(lngUpdateRecordIndex, gstrPatholCol_�������) = cbxBox.Text
        .Text(lngUpdateRecordIndex, gstrPatholCol_��������) = cbxDrawer.Text
        .Text(lngUpdateRecordIndex, gstrPatholCol_��ϸ��ַ) = txtPlace.Text
        
        .Text(lngUpdateRecordIndex, gstrPatholCol_��������) = cbxArchivesClass.Text
        
        mrsArchivesClass.Filter = "��������='" & cbxArchivesClass.Text & "'"
        If mrsArchivesClass.RecordCount > 0 Then
            .Text(lngUpdateRecordIndex, gstrPatholCol_��������) = Val(Nvl(mrsArchivesClass!��������))
            .Text(lngUpdateRecordIndex, gstrPatholCol_��������) = Nvl(mrsArchivesClass!��������)
        End If
        
    End With
End Sub




Private Sub cmdSure_Click()
On Error GoTo errHandle
    Dim strErr As String
    Dim strNewArchivesId As String

    '����Ƿ�¼����Ч����
    strErr = CheckArchivesFileDataIsValid()
    If Trim(strErr) <> "" Then
        Call ShowProcessHint(strErr)
        Exit Sub
    End If
    
    
    If Not mblnIsUpdate Then
        '��������
        Call NewArchivesInf
        
        Call mufgParentGrid.LocateRow(mufgParentGrid.GridRows - 1)
    Else
        '���µ���
        Call UpdateArchivesInf
    End If
    
    mblnIsSucceed = True
    
    If Not CBool(chkContinue.value) Then
        Call Unload(Me)
    End If
    
    Call CloseProcessHint
Exit Sub
errHandle:
    Call ShowProcessHint(err.Description)
    err.Clear
End Sub

Private Sub Form_Initialize()
    mblnIsSucceed = False
    mblnIsUpdate = False
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    Call RestoreWinState(Me, App.ProductName)
    Call LoadArchivesClassData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub LoadArchivesClassData()
'���ص�����������
    Dim strSql As String
    
    strSql = "select ID,��������,��������,�������� from ����������"
    
    Set mrsArchivesClass = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    Call cbxArchivesClass.Clear
    If mrsArchivesClass.RecordCount <= 0 Then Exit Sub
    
    While Not mrsArchivesClass.EOF
        Call cbxArchivesClass.AddItem(Nvl(mrsArchivesClass!��������))
        
        cbxArchivesClass.ItemData(cbxArchivesClass.ListCount - 1) = Nvl(mrsArchivesClass!ID)
        
        mrsArchivesClass.MoveNext
    Wend
    
    If cbxArchivesClass.ListCount > 0 Then cbxArchivesClass.ListIndex = 0
End Sub






Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
err.Clear
End Sub

