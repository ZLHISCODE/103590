VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmLabSampleRegisterRefuse 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�걾����"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboRefuse 
      Height          =   300
      Left            =   1230
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2280
      Width           =   6345
   End
   Begin VB.PictureBox PicRecord 
      Height          =   1935
      Left            =   60
      ScaleHeight     =   1875
      ScaleWidth      =   7455
      TabIndex        =   6
      Top             =   300
      Width           =   7515
      Begin XtremeReportControl.ReportControl rptAlist 
         Height          =   1335
         Left            =   360
         TabIndex        =   7
         Top             =   60
         Width           =   4665
         _Version        =   589884
         _ExtentX        =   8229
         _ExtentY        =   2355
         _StockProps     =   0
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
   End
   Begin VB.CheckBox chkHide 
      Caption         =   "����δѡ�е�ҽ��"
      Height          =   225
      Left            =   5850
      TabIndex        =   5
      Top             =   68
      Value           =   1  'Checked
      Width           =   1875
   End
   Begin VB.CommandButton cmdRefuse 
      Caption         =   "����(&F)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4440
      TabIndex        =   4
      Top             =   4140
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6060
      TabIndex        =   3
      Top             =   4140
      Width           =   1100
   End
   Begin VB.TextBox txt���� 
      Height          =   1335
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2610
      Width           =   7545
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   1500
      Top             =   4020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampleRegisterRefuse.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampleRegisterRefuse.frx":006C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampleRegisterRefuse.frx":0606
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampleRegisterRefuse.frx":0BA0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "��д��������"
      Height          =   180
      Left            =   90
      TabIndex        =   1
      Top             =   2340
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ѡ����ձ걾"
      Height          =   180
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   1080
   End
End
Attribute VB_Name = "frmLabSampleRegisterRefuse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum mAcol                                  'ҽ���б�
    ID
    ѡ��
    ͼ��
    ��ִ��
    �ɼ���ʽ
    ҽ������
    ����
    ִ�п���
    ����ҽ��
    ����ʱ��
    ������
    ����ʱ��
    �걾
    ����ʱ��
    �Թ���ɫ
    �ϲ�ҽ��
    �Թܱ���
    ������
    ��Ѫ��
    �Թ�����
    ����
    ������Դ
    �������
    Ӥ��
    ����
    ���ID
    ҽ��id
    ����ID
    ����
    �Ա�
    ����
    ��ʶ��
    ����
    ���˿���
    ����ʱ��
    ������ĿID
    ִ��״̬
End Enum
Dim mRecords As ReportRecords
Public Sub ShowMe(Objfrm As Object, Recordset As ReportRecords)
    Set mRecords = Recordset
    Me.Show vbModal, Objfrm
End Sub

Private Sub cboRefuse_Click()
    Me.txt����.Text = Mid(Me.cboRefuse.Text, InStr(Me.cboRefuse.Text, "-") + 1)
End Sub

Private Sub chkHide_Click()
    Dim intLoop As Integer
    With Me.rptAlist
        If Me.chkHide.Value = 1 Then
            For intLoop = 0 To .Records.Count - 1
                If .Records(intLoop).Item(mAcol.ѡ��).Checked = True Then
                    .Records(intLoop).Visible = True
                Else
                    .Records(intLoop).Visible = False
                End If
            Next
        Else
            If .Records.Count > 0 Then
                .Records(intLoop).Visible = True
            End If
        End If
        .Populate
    End With
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRefuse_Click()
    Dim blnSelect As Boolean
    Dim intLoop As Integer
    
    
    With Me.rptAlist
        If .Rows.Count > 0 Then
            For intLoop = 0 To .Rows.Count - 1
                If .Rows(intLoop).Record(mAcol.ѡ��).Checked = True Then
                    blnSelect = True
                    Exit For
                End If
            Next
        End If
    End With
    
    'û��ѡ�����ҽ��
    If blnSelect = False Then
        MsgBox "��ѡ��һ��ҽ�����ܽ��о���!", vbInformation, Me.Caption
        Exit Sub
    End If
    
    '������д��������
    If Trim(Me.txt����) = "" Then
        MsgBox "����д�������ɣ�", vbInformation, Me.Caption
        Me.txt����.SetFocus
        Exit Sub
    End If
    
    '��ʼ����
    On Error GoTo errH
    
    gcnOracle.BeginTrans
    With Me.rptAlist
        For intLoop = 0 To .Rows.Count - 1
            If .Rows(intLoop).Record(mAcol.ѡ��).Checked = True Then
                gstrSql = "Zl_����걾��¼_�걾����(" & _
                            .Rows(intLoop).Record(mAcol.ҽ��id).Value & ",'" & _
                            Me.txt����.Text & "','" & UserInfo.���� & "')"
                zlDatabase.ExecuteProcedure gstrSql, Me.Caption
            End If
        Next
    End With
    gcnOracle.CommitTrans
    Unload Me
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Sub

Private Sub Form_Load()
    Dim intLoop As Integer
    Dim lngLoop As Long
    Dim Record As ReportRecord
    Dim rsTmp As New ADODB.Recordset
    
    With Me.rptAlist
        .Top = 0
        .Left = 0
        .Width = Me.PicRecord.ScaleWidth
        .Height = Me.PicRecord.ScaleHeight
    End With
    
    rptAlist.SetImageList ImgList
    With rptAlist.PaintManager
        .ColumnStyle = xtpColumnShaded
        .GridLineColor = RGB(225, 225, 225)
        .NoGroupByText = "�϶��б��⵽����,�����з���..."
        .NoItemsText = "û�п���ʾ����Ŀ..."
        .VerticalGridStyle = xtpGridSolid
        .HideSelection = True
    End With
    With Me.rptAlist.Columns
        Set Column = .Add(mAcol.ID, "ID", 0, False): Column.Visible = False
        Set Column = .Add(mAcol.ѡ��, "Check", 18, False): Column.Icon = 0
        Set Column = .Add(mAcol.ͼ��, "", 18, False): Column.Icon = 3
        Set Column = .Add(mAcol.�ɼ���ʽ, "�ɼ���ʽ", 75, True)
        Set Column = .Add(mAcol.�걾, "�걾", 55, True)
        Set Column = .Add(mAcol.ҽ������, "ҽ������", 75, True)
        Set Column = .Add(mAcol.����, "����", 75, True)
        Set Column = .Add(mAcol.ִ�п���, "ִ�п���", 75, True)
        Set Column = .Add(mAcol.����ҽ��, "����ҽ��", 75, True)
        Set Column = .Add(mAcol.����ʱ��, "����ʱ��", 75, True)
        Set Column = .Add(mAcol.������, "������", 65, True)
        Set Column = .Add(mAcol.����ʱ��, "����ʱ��", 75, True)
        Set Column = .Add(mAcol.����ʱ��, "����ʱ��", 75, True)
        Set Column = .Add(mAcol.����ʱ��, "����ʱ��", 75, True)
        Set Column = .Add(mAcol.�Թ���ɫ, "��ɫ����", 18, True): Column.Visible = False
        Set Column = .Add(mAcol.�Թܱ���, "�Թܱ���", 18, True): Column.Visible = False
        Set Column = .Add(mAcol.������, "������", 60, True)
        Set Column = .Add(mAcol.��Ѫ��, "��Ѫ��", 60, True): Column.Visible = False
        Set Column = .Add(mAcol.�Թ�����, "�Թ�����", 60, True): Column.Visible = False
        Set Column = .Add(mAcol.����, "����", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.������Դ, "������Դ", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.�������, "�������", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.Ӥ��, "Ӥ��", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.����, "����", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.���ID, "���ID", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.����ID, "����ID", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.����, "����", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.�Ա�, "�Ա�", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.����, "����", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.��ʶ��, "��ʶ��", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.����, "����", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.���˿���, "���˿���", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.������ĿID, "������ĿId", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.ҽ��id, "ҽ��ID", 50, True): Column.Visible = False
    End With
    
    For lngLoop = 0 To mRecords.Count - 1
        Set Record = Me.rptAlist.Records.Add
        For intLoop = 0 To Me.rptAlist.Columns.Count + 1
            Record.AddItem ""
        Next
                
        Record(mAcol.ID).Value = mRecords(lngLoop).Item(mAcol.ID).Value
        Record(mAcol.ѡ��).HasCheckbox = True
        Record(mAcol.ѡ��).Checked = mRecords(lngLoop).Item(mAcol.ѡ��).Checked
        Record(mAcol.ͼ��).BackColor = mRecords(lngLoop).Item(mAcol.ͼ��).BackColor
        Record(mAcol.�ɼ���ʽ).Value = mRecords(lngLoop).Item(mAcol.�ɼ���ʽ).Value
        Record(mAcol.ҽ������).Value = mRecords(lngLoop).Item(mAcol.ҽ������).Value
        Record(mAcol.����).Value = mRecords(lngLoop).Item(mAcol.����).Value
        Record(mAcol.ִ�п���).Value = mRecords(lngLoop).Item(mAcol.ִ�п���).Value
        Record(mAcol.����ҽ��).Value = mRecords(lngLoop).Item(mAcol.����ҽ��).Value
        Record(mAcol.����ʱ��).Value = mRecords(lngLoop).Item(mAcol.����ʱ��).Value
        Record(mAcol.������).Value = mRecords(lngLoop).Item(mAcol.������).Value
        Record(mAcol.����ʱ��).Value = mRecords(lngLoop).Item(mAcol.����ʱ��).Value
        Record(mAcol.�Թ���ɫ).Value = mRecords(lngLoop).Item(mAcol.�Թ���ɫ).Value
        Record(mAcol.�Թܱ���).Value = mRecords(lngLoop).Item(mAcol.�Թܱ���).Value
        Record(mAcol.�걾).Value = mRecords(lngLoop).Item(mAcol.�걾).Value
        Record(mAcol.����ʱ��).Value = mRecords(lngLoop).Item(mAcol.����ʱ��).Value
        Record(mAcol.������).Value = mRecords(lngLoop).Item(mAcol.������).Value
        Record(mAcol.��Ѫ��).Value = mRecords(lngLoop).Item(mAcol.��Ѫ��).Value
        Record(mAcol.�Թ�����).Value = mRecords(lngLoop).Item(mAcol.�Թ�����).Value
        Record(mAcol.����).Value = mRecords(lngLoop).Item(mAcol.����).Value
        Record(mAcol.������Դ).Value = mRecords(lngLoop).Item(mAcol.������Դ).Value
        Record(mAcol.Ӥ��).Value = mRecords(lngLoop).Item(mAcol.Ӥ��).Value
        Record(mAcol.����).Value = mRecords(lngLoop).Item(mAcol.����).Value
        Record(mAcol.���ID).Value = mRecords(lngLoop).Item(mAcol.���ID).Value
        
        Record(mAcol.����ID).Value = mRecords(lngLoop).Item(mAcol.����ID).Value
        Record(mAcol.����).Value = mRecords(lngLoop).Item(mAcol.����).Value
        Record(mAcol.�Ա�).Value = mRecords(lngLoop).Item(mAcol.�Ա�).Value
        Record(mAcol.����).Value = mRecords(lngLoop).Item(mAcol.����).Value
        Record(mAcol.��ʶ��).Value = mRecords(lngLoop).Item(mAcol.��ʶ��).Value
        Record(mAcol.����).Value = mRecords(lngLoop).Item(mAcol.����).Value
        Record(mAcol.���˿���).Value = mRecords(lngLoop).Item(mAcol.���˿���).Value
        Record(mAcol.����ʱ��).Value = mRecords(lngLoop).Item(mAcol.����ʱ��).Value
        Record(mAcol.������ĿID).Value = mRecords(lngLoop).Item(mAcol.������ĿID).Value
        Record(mAcol.ҽ��id).Value = mRecords(lngLoop).Item(mAcol.ҽ��id).Value
        
        For intLoop = 0 To Me.rptAlist.Columns.Count + 1
            Record(intLoop).ForeColor = mRecords(lngLoop).Item(mAcol.�Թ���ɫ).Value
        Next
        
    Next
    Call chkHide_Click
    Me.rptAlist.Populate
    
    On Error GoTo errH:
    
    gstrSql = "select ����,���� from �����������"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Do While Not rsTmp.EOF
        With Me.cboRefuse
            .AddItem Nvl(rsTmp("����")) & "-" & Nvl(rsTmp("����"))
        End With
        rsTmp.MoveNext
    Loop
        
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Picture1_Resize()
    
End Sub

Private Sub rptAlist_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call chkHide_Click
End Sub

