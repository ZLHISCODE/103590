VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPathItemBatReplace 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "��Ŀ��������"
   ClientHeight    =   9855
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   13725
   FillColor       =   &H00404040&
   ForeColor       =   &H8000000C&
   Icon            =   "frmPathItemBatReplace.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10521.35
   ScaleMode       =   0  'User
   ScaleWidth      =   13973.42
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraSplit 
      BorderStyle     =   0  'None
      Height          =   42
      Left            =   4200
      TabIndex        =   40
      Top             =   5280
      Width           =   4695
   End
   Begin VB.PictureBox picAdvice 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2145
      Left            =   4680
      ScaleHeight     =   2145
      ScaleWidth      =   6615
      TabIndex        =   37
      Top             =   5520
      Width           =   6615
      Begin VB.CommandButton cmdEdit 
         Caption         =   "�滻��Ŀ�༭"
         Height          =   420
         Left            =   0
         TabIndex        =   39
         Top             =   120
         Width           =   1500
      End
      Begin zlCISPath.UCAdviceList ucAdvice 
         Height          =   1575
         Left            =   0
         TabIndex        =   38
         Top             =   600
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2778
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   13725
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   9180
      Width           =   13725
      Begin VB.CommandButton cmdQuit 
         BackColor       =   &H8000000E&
         Caption         =   "�˳�(&Q)"
         Height          =   420
         Left            =   12000
         TabIndex        =   9
         Top             =   120
         Width           =   1500
      End
      Begin VB.CommandButton cmdBatExe 
         BackColor       =   &H80000014&
         Caption         =   "�����滻(&B)"
         Height          =   420
         Left            =   10440
         TabIndex        =   8
         Top             =   120
         Width           =   1500
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   20400
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   20280
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.PictureBox picTop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   4680
      ScaleHeight     =   975
      ScaleWidth      =   9735
      TabIndex        =   19
      Top             =   960
      Width           =   9735
      Begin VB.PictureBox picItem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   6855
         TabIndex        =   20
         Top             =   0
         Width           =   6855
         Begin VB.CommandButton cmd�÷� 
            Height          =   240
            Left            =   4440
            Picture         =   "frmPathItemBatReplace.frx":6852
            Style           =   1  'Graphical
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   165
            Width           =   270
         End
         Begin VB.CommandButton cmdƵ�� 
            Height          =   240
            Left            =   4440
            Picture         =   "frmPathItemBatReplace.frx":6948
            Style           =   1  'Graphical
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   570
            Width           =   270
         End
         Begin VB.TextBox txt���� 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   525
            MaxLength       =   10
            TabIndex        =   0
            Top             =   135
            Width           =   1290
         End
         Begin VB.TextBox txt���� 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   525
            MaxLength       =   10
            TabIndex        =   2
            Top             =   540
            Width           =   1290
         End
         Begin VB.CheckBox chkPra 
            BackColor       =   &H80000005&
            Caption         =   "ֻ�滻Ƶ����ͬ��"
            Height          =   255
            Index           =   2
            Left            =   5040
            TabIndex        =   6
            Top             =   600
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox chkPra 
            BackColor       =   &H80000005&
            Caption         =   "ֻ�滻�÷���ͬ��"
            Height          =   255
            Index           =   1
            Left            =   5040
            TabIndex        =   5
            Top             =   360
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox chkPra 
            BackColor       =   &H80000005&
            Caption         =   "ֻ�滻������ͬ��"
            Height          =   255
            Index           =   0
            Left            =   5040
            TabIndex        =   4
            Top             =   120
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.TextBox txtƵ�� 
            Height          =   300
            Left            =   2925
            TabIndex        =   3
            Top             =   540
            Width           =   1815
         End
         Begin VB.TextBox txt�÷� 
            Height          =   300
            Left            =   2925
            TabIndex        =   1
            Top             =   135
            Width           =   1815
         End
         Begin VB.Label lbl������λ 
            BackColor       =   &H00C0C0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "g"
            Height          =   180
            Left            =   1935
            TabIndex        =   26
            Top             =   195
            Width           =   405
         End
         Begin VB.Label lbl���� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Left            =   120
            TabIndex        =   25
            Top             =   195
            Width           =   360
         End
         Begin VB.Label lbl���� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Left            =   120
            TabIndex        =   24
            Top             =   600
            Width           =   360
         End
         Begin VB.Label lbl������λ 
            BackColor       =   &H00C0C0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Height          =   180
            Left            =   1920
            TabIndex        =   23
            Top             =   600
            Width           =   450
         End
         Begin VB.Label lbl�÷� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�÷�"
            Height          =   180
            Left            =   2520
            TabIndex        =   22
            Top             =   195
            Width           =   360
         End
         Begin VB.Label lblƵ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ƶ��"
            Height          =   180
            Left            =   2520
            TabIndex        =   21
            Top             =   600
            Width           =   360
         End
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H80000005&
         Caption         =   "����·��(&F)"
         Height          =   420
         Left            =   6960
         TabIndex        =   7
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.PictureBox picSplit 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   7335
      Left            =   3840
      MousePointer    =   9  'Size W E
      ScaleHeight     =   7335
      ScaleWidth      =   45
      TabIndex        =   18
      Top             =   1200
      Width           =   45
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   13725
      TabIndex        =   16
      Top             =   0
      Width           =   13725
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "6����������滻������滻��"
         Height          =   255
         Index           =   8
         Left            =   8040
         TabIndex        =   36
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "3�����Ұ���ѡ���������Ŀ��·����"
         Height          =   255
         Index           =   7
         Left            =   8040
         TabIndex        =   35
         Top             =   120
         Width           =   3135
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   195
         Picture         =   "frmPathItemBatReplace.frx":6A3E
         Top             =   45
         Width           =   720
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "5������������Ŀ��Ӧ���滻��Ŀ��"
         Height          =   255
         Index           =   6
         Left            =   4560
         TabIndex        =   31
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "˵��:"
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   30
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "2������������Ŀ�滻�Ĺ���"
         Height          =   255
         Index           =   3
         Left            =   4560
         TabIndex        =   28
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "4����ѡ��Ҫ�滻��·������"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   27
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "1��ѡ����Ҫ�滻��������Ŀ��"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   17
         Top             =   120
         Width           =   3375
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   10000
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   20280
         Y1              =   840
         Y2              =   840
      End
   End
   Begin VB.PictureBox picMain 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   4680
      ScaleHeight     =   2775
      ScaleWidth      =   7335
      TabIndex        =   12
      Top             =   2280
      Width           =   7335
      Begin XtremeReportControl.ReportControl rptPath 
         Height          =   2055
         Left            =   0
         TabIndex        =   13
         Top             =   240
         Width           =   7215
         _Version        =   589884
         _ExtentX        =   12726
         _ExtentY        =   3625
         _StockProps     =   0
         BorderStyle     =   2
      End
      Begin VB.Label lblNote 
         BackColor       =   &H80000005&
         Caption         =   "·���б�"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   0
      ScaleHeight     =   6135
      ScaleWidth      =   4575
      TabIndex        =   10
      Top             =   840
      Width           =   4575
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   4215
         Left            =   120
         TabIndex        =   11
         Top             =   2280
         Width           =   3255
         _Version        =   589884
         _ExtentX        =   5741
         _ExtentY        =   7435
         _StockProps     =   0
         BorderStyle     =   2
         ShowItemsInGroups=   -1  'True
         AutoColumnSizing=   0   'False
      End
      Begin VB.Frame fraFind 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   900
         Left            =   120
         TabIndex        =   41
         Top             =   135
         Width           =   4455
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   480
            TabIndex        =   44
            ToolTipText     =   "������һ��(F3)"
            Top             =   480
            Width           =   3255
         End
         Begin VB.OptionButton optType 
            BackColor       =   &H80000005&
            Caption         =   "·������"
            Height          =   300
            Index           =   1
            Left            =   2760
            TabIndex        =   43
            Top             =   60
            Width           =   1215
         End
         Begin VB.OptionButton optType 
            BackColor       =   &H80000005&
            Caption         =   "ֱ�Ӳ���"
            Height          =   300
            Index           =   0
            Left            =   0
            TabIndex        =   42
            Top             =   60
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.Label lblFind 
            BackColor       =   &H80000005&
            Caption         =   "����"
            Height          =   255
            Left            =   0
            TabIndex        =   45
            Top             =   510
            Width           =   495
         End
      End
      Begin VB.Label lblStopNote 
         BackColor       =   &H80000005&
         Caption         =   "·�������õ�����ͣ����Ŀ"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   2295
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   2880
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathItemBatReplace.frx":7193
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathItemBatReplace.frx":772D
            Key             =   "Check"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblNote 
      BackStyle       =   0  'Transparent
      Caption         =   "1��ѡ��ͣ����Ŀ�����ù����������������������·��"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmPathItemBatReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrSelItem As String    '��¼ѡ����Ŀ ��ʽ:������ĿID_�շ�ϸĿID
Private mstrPrivs As String      '�ٴ�·��ģ��Ȩ��
Private mrsAdvice As ADODB.Recordset   '�滻��Ŀҽ����¼��
Private mblnChange As Boolean   '��ʶ���˲���ֵ���غ�����ޱ䶯��� T-�䶯��F-δ�䶯��

'---------------------------------
Private Enum CHK_INDEX
    chk_���� = 0
    chk_�÷� = 1
    chk_Ƶ�� = 2
End Enum

Private Enum COL_LIST
    COL_���� = 0
    COL_����
    COL_����
    COL_��Ʒ��
    COL_����
    COL_ҩƷ����
    
    '������
    COL_������ĿID
    COL_�շ�ϸĿID
    COL_�������
    COL_��������
    COL_ִ�з���
    COL_���㷽ʽ
    COL_�걾��λ
    COL_��鷽��
    COL_����ID
    COL_���ID
    COL_����
End Enum

Private Enum COL_PATH
    Path_ID = 0
    Path_ѡ��
    Path_����
    Path_����
    Path_����
    Path_�汾
    Path_˵��
End Enum

Private Enum CONST_COLOR
    Color_Enabled = &H80000005
    Color_UNEnabled = &H8000000F
End Enum
'-------------------------------------------------------------------------------------------------------
Public Sub ShowMe(frmParent As Object, ByVal strPrivs As String)
'����:��ں���
'����:������
'
    mstrPrivs = strPrivs
    
    Me.Show 1, frmParent
End Sub

Private Sub chkPra_Click(Index As Integer)
    Dim blnCheck As Boolean
    
    blnCheck = chkPra(Index) = vbChecked
    If Index = chk_���� Then
        SetEditable IIf(blnCheck, 1, -1), IIf(blnCheck, 1, -1)
    ElseIf Index = chk_�÷� Then
        SetEditable , , IIf(blnCheck, 1, -1)
    ElseIf Index = chk_Ƶ�� Then
        SetEditable , , , IIf(blnCheck, 1, -1)
    End If
    If rptPath.Records.count > 0 Then
        Call ClearPath
    End If
End Sub

Private Sub cmdBatExe_Click()
'����:�����滻����
    '��������ǰ�ļ�����
    Dim i As Long
    Dim strPath As String
    Dim strTmp As String
    
    If mrsAdvice.RecordCount = 0 Then
        MsgBox "���������滻��Ŀ,��ִ�С������滻�����ܡ�", vbInformation + vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    With rptPath
        For i = 1 To .Rows.count - 1
            If Not .Rows(i).GroupRow Then
                If .Rows(i).Record(Path_ѡ��).Checked Then
                    strPath = strPath & ":" & .Rows(i).Record(Path_ID).Value & "," & .Rows(i).Record(Path_�汾).Value
                    If .Rows(i).Record(Path_�汾).Value = .Rows(i).Record.Tag Then  '��¼�����°汾��·��ID�Ͱ汾��
                        strTmp = strTmp & "," & .Rows(i).Record(Path_ID).Value & "_" & .Rows(i).Record.Tag
                    End If
                End If
            End If
        Next
        If strPath <> "" Then strPath = Mid(strPath, 2)
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
    End With
    
    If strPath = "" Then
        MsgBox "����ѡ���滻��·��,��ִ�С������滻�����ܡ�", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
        Exit Sub
    End If
    '���ݱ���
    If SaveData(strPath, strTmp) Then
        'ˢ�½���
        Call RefreshData
    End If
End Sub

Private Sub cmdEdit_Click()
'����:�滻��Ŀ�༭
    Dim rsScheme As ADODB.Recordset
    Dim colAdviceID As New Collection
    Dim lng��� As Long
    Dim lngҽ��ID As Long
    
    Call InitSchemeRecordset(rsScheme)
    
    Do While Not mrsAdvice.EOF
        rsScheme.AddNew
        rsScheme!��� = mrsAdvice!ID
        rsScheme!������ = mrsAdvice!���id
        rsScheme!��Ч = mrsAdvice!��Ч
        rsScheme!������ĿID = mrsAdvice!������ĿID
        rsScheme!�շ�ϸĿID = mrsAdvice!�շ�ϸĿID
        rsScheme!ҽ������ = mrsAdvice!ҽ������
        rsScheme!�������� = mrsAdvice!��������
        rsScheme!�ܸ����� = mrsAdvice!�ܸ�����
        rsScheme!ҽ������ = mrsAdvice!ҽ������
        rsScheme!ִ��Ƶ�� = mrsAdvice!ִ��Ƶ��
        rsScheme!Ƶ�ʴ��� = mrsAdvice!Ƶ�ʴ���
        rsScheme!Ƶ�ʼ�� = mrsAdvice!Ƶ�ʼ��
        rsScheme!�����λ = mrsAdvice!�����λ
        rsScheme!ʱ�䷽�� = mrsAdvice!ʱ�䷽��
        rsScheme!ִ�п���ID = mrsAdvice!ִ�п���ID
        rsScheme!ִ������ = mrsAdvice!ִ������
        rsScheme!�걾��λ = mrsAdvice!�걾��λ
        rsScheme!��鷽�� = mrsAdvice!��鷽��
        rsScheme!�䷽ID = mrsAdvice!�䷽ID
        rsScheme!�����ĿID = mrsAdvice!�����ĿID
        rsScheme!ִ�б�� = mrsAdvice!ִ�б��
        
        rsScheme.Update
        mrsAdvice.MoveNext
    Loop
    
    Set rsScheme = gobjKernel.ShowSchemeEdit(Me, 2, rsScheme, False, False, "", 2, rptList.SelectedRows(0).Record(COL_�������).Value & "", _
                    rptList.SelectedRows(0).Record(COL_��������).Value & "", rptList.SelectedRows(0).Record(COL_ִ�з���).Value & "")
    
    
    '��ɾ����ǰ��ҽ��ID
    If mrsAdvice.RecordCount > 0 And Not rsScheme Is Nothing Then
        Call InitAdviceRecordset '���³�ʼ��
    End If

    If Not rsScheme Is Nothing Then
         '�Ȳ����µ�ҽ��ID
        Do While Not rsScheme.EOF
            lngҽ��ID = zlDatabase.GetNextId("·��ҽ������")
            colAdviceID.Add lngҽ��ID, "_" & rsScheme!���
            rsScheme.MoveNext
        Loop
        rsScheme.MoveFirst: lng��� = 1
        Do While Not rsScheme.EOF
            mrsAdvice.AddNew
            mrsAdvice!ID = colAdviceID("_" & rsScheme!���)
            If Not IsNull(rsScheme!������) Then
                mrsAdvice!���id = colAdviceID("_" & rsScheme!������)
            End If
            mrsAdvice!��� = lng���
            mrsAdvice!��Ч = rsScheme!��Ч
            mrsAdvice!������ĿID = rsScheme!������ĿID
            mrsAdvice!�շ�ϸĿID = rsScheme!�շ�ϸĿID
            If IsNull(rsScheme!������ĿID) Then
                mrsAdvice!ҽ������ = rsScheme!ҽ������ '����¼��ҽ���ű���
            End If
            mrsAdvice!�������� = rsScheme!��������
            mrsAdvice!�ܸ����� = rsScheme!�ܸ�����
            mrsAdvice!ҽ������ = rsScheme!ҽ������
            mrsAdvice!ִ��Ƶ�� = rsScheme!ִ��Ƶ��
            mrsAdvice!Ƶ�ʴ��� = rsScheme!Ƶ�ʴ���
            mrsAdvice!Ƶ�ʼ�� = rsScheme!Ƶ�ʼ��
            mrsAdvice!�����λ = rsScheme!�����λ
            mrsAdvice!ʱ�䷽�� = rsScheme!ʱ�䷽��
            mrsAdvice!ִ�п���ID = rsScheme!ִ�п���ID
            mrsAdvice!ִ������ = rsScheme!ִ������
            mrsAdvice!�걾��λ = rsScheme!�걾��λ
            mrsAdvice!��鷽�� = rsScheme!��鷽��
            mrsAdvice!�Ƿ�ȱʡ = rsScheme!�Ƿ�ȱʡ
            mrsAdvice!�Ƿ�ѡ = rsScheme!�Ƿ�ѡ
            mrsAdvice!�䷽ID = rsScheme!�䷽ID
            mrsAdvice!�����ĿID = rsScheme!�����ĿID
            mrsAdvice!ִ�б�� = rsScheme!ִ�б��
            
            mrsAdvice.Update
            
            lng��� = lng��� + 1
            rsScheme.MoveNext
        Loop
        If mrsAdvice.RecordCount > 1 Then mrsAdvice.MoveFirst
    End If
    
    Call ShowAdvice
    cmdBatExe.Enabled = True
End Sub

Private Sub cmdFind_Click()
    Dim strSql As String
    Dim strTmp As String
    Dim rsTmp As ADODB.Recordset
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim objCol As ReportColumn
    Dim str��� As String
    
    Dim i As Long
    '�������
    If rptList.Records.count = 0 Then Exit Sub
    If rptList.SelectedRows.count < 1 Then Exit Sub
    
    rptPath.Records.DeleteAll: rptPath.Populate: cmdEdit.Enabled = False
    
    On Error GoTo errH
    strSql = "Select Distinct d.Id, d.����, d.����, d.����,d.˵��,H.�汾��,D.���°汾 " & vbNewLine & _
            "From ·��ҽ������ A, �ٴ�·��ҽ�� B, �ٴ�·����Ŀ C, �ٴ�·���汾 H,�ٴ�·��Ŀ¼ D" & vbNewLine & _
            "Where a.Id = b.ҽ������id And b.·����Ŀid = c.Id And c.·��id = H.·��Id And c.�汾�� = H.�汾�� And H.ͣ���� is null And H.·��Id=D.ID"
    With rptList.SelectedRows(0)
    
        If Val(.Record(COL_�շ�ϸĿID).Value) = 0 Then
            strSql = strSql & " And a.������ĿID =[1]"
        Else
            strSql = strSql & " And a.�շ�ϸĿID =[1]"
        End If
        str��� = .Record(COL_�������).Value
        If InStr(",D,C,", "," & str��� & ",") > 0 Then
            strSql = strSql & " And Instr([7], ',' || NVl(a.���ID,a.Id)|| ',') > 0 "
        End If
        
        If chkPra(chk_����).Value = vbChecked Then
            If txt����.Text <> "" Then
                strSql = strSql & " And a.�������� =[2] "
            End If
            If txt����.Text <> "" Then
                strSql = strSql & " And a.�ܸ����� = [3] "
            End If
        End If
        
        If chkPra(chk_�÷�).Value = vbChecked Then
            If txt�÷�.Text <> "" Then
                strSql = strSql & " and exists (select 1 from ·��ҽ������  E where e.id=a.���id and e.������Ŀid = [4]) "
            End If
        End If
        
        If chkPra(chk_Ƶ��).Value = vbChecked Then
            If txtƵ��.Text <> "" Then
                strSql = strSql & " And a.ִ��Ƶ�� =[5] "
            End If
        End If
            
        If InStr(mstrPrivs, "ȫԺ·��") = 0 Then
            'û��Ȩ��ʱ��ֻ�ܶ�ֻӦ���ڱ��Ƶ�·�����д���
            strSql = strSql & _
                     " And D.ͨ�� = 2 And Exists" & vbNewLine & _
                     "      (Select 1 From ������Ա E,�ٴ�·������ F  " & vbNewLine & _
                     "       Where E.��Աid = [6] And F.����id = E.����id And F.·��id = D.ID)"
        End If
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, IIf(Val(.Record(COL_�շ�ϸĿID).Value) = 0, Val(.Record(COL_������ĿID).Value), Val(.Record(COL_�շ�ϸĿID).Value)), _
                    Val(txt����.Text), Val(txt����.Text), Val(txt�÷�.Tag), txtƵ��.Text, UserInfo.ID, "," & .Record(COL_����ID).Value & ",")
        If rsTmp.RecordCount = 0 Then Exit Sub
        
        With rptPath
            For i = 1 To rsTmp.RecordCount
                Set objRecord = .Records.Add
                objRecord.AddItem rsTmp!ID & ""
                
                Set objItem = objRecord.AddItem("")
                objItem.HasCheckbox = True
                If .Columns(Path_ѡ��).Icon = img16.ListImages("UnCheck").Index - 1 Then
                    objItem.Checked = True
                Else
                    objItem.Checked = False
                End If
                objRecord.AddItem rsTmp!���� & ""
                objRecord.AddItem rsTmp!���� & ""
                objRecord.AddItem rsTmp!���� & ""
                objRecord.AddItem rsTmp!�汾�� & ""
                objRecord.AddItem rsTmp!˵�� & ""
                objRecord.Tag = rsTmp!���°汾 & ""  '���ݸ�ֵ�жϱ䶯��¼�Ƿ����
                rsTmp.MoveNext
            Next
            .Populate
        End With
        
        cmdEdit.Enabled = rptPath.Records.count > 0
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdƵ��_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim str��Χ As String, intƵ�� As Integer, vRect As RECT
    Dim lng������ĿID As Long
       
    If rptList.SelectedRows.count = 0 Then Exit Sub  '���������
    With rptList.SelectedRows(0)
         strSql = "Select ִ��Ƶ�� From ������ĿĿ¼ Where ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(.Record(COL_������ĿID).Value))
        If Not rsTmp.EOF Then intƵ�� = NVL(rsTmp!ִ��Ƶ��, 0)
        
        If txt����.Text <> "" Then '����
            If .Record(COL_�������).Value <> "7" And intƵ�� = 0 Then
                str��Χ = "1,-1" '��������Ϊһ����
            Else
                str��Χ = GetƵ�ʷ�Χ(intƵ��)
            End If
        Else
            str��Χ = GetƵ�ʷ�Χ(intƵ��)
            intƵ�� = Decode(str��Χ, "1", 0, "2", 0, "-1", 1, "-2", 2, "-3", 1, "-5", 1)
        End If
        
        '��ѡ��Ƶ�ʵĳ���Ƶ��
        lng������ĿID = Val(.Record(COL_������ĿID).Value)
        strSql = ""
        If InStr("," & str��Χ & ",", ",1,") > 0 Then
            strSql = " And (Exists(Select 1 From �����÷����� Where ��ĿID=[2] And �÷�ID is NULL And Ƶ��=A.���� And A.���÷�Χ=1)" & _
                " Or (Select Count(*) From �����÷����� Where ��ĿID=[2] And �÷�ID is NULL And Ƶ�� Is Not NULL)<=1)"
        End If
        strSql = _
            " Select Rownum as ID,A.����,A.����,A.����," & _
            " A.Ӣ������,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,A.���÷�Χ as ��ΧID" & _
            " From ����Ƶ����Ŀ A" & _
            " Where (Instr([1],','||A.���÷�Χ||',')>0  Or a.���÷�Χ=[3])" & strSql & _
            " Order by A.���÷�Χ,A.����"
        vRect = zlControl.GetControlRect(txtƵ��.Hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "����Ƶ��", False, "", "", False, False, True, _
            vRect.Left, vRect.Top, txtƵ��.Height, blnCancel, False, True, "," & str��Χ & ",", lng������ĿID, IIf(txt����.Text <> "", -5, -3))
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "û�п��õ�����Ƶ����Ŀ�����ȵ�ҽ��Ƶ�ʹ��������á�", vbInformation, gstrSysName
            End If
            Call zlControl.TxtSelAll(txtƵ��)
            txtƵ��.SetFocus: Exit Sub
        End If
        txtƵ��.Text = rsTmp!���� & ""
        Call zlControl.TxtSelAll(txtƵ��)
        txtƵ��.SetFocus
  
    End With
End Sub

Private Sub cmd�÷�_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim int���� As Integer, vRect As RECT
    
    If rptList.SelectedRows.count = 0 Then Exit Sub  '���������
    
    With rptList.SelectedRows(0)
        If InStr(",5,6,", .Record(COL_�������).Value) > 0 Then
            int���� = 2 '��ҩ;��
        ElseIf .Record(COL_�������).Value = "C" Then
            int���� = 6 '�ɼ�����
        ElseIf .Record(COL_�������).Value = "K" Then
            int���� = 8 '��Ѫ;��
        Else
            int���� = 4 '��ҩ�÷�
        End If
        If int���� = 2 Then 'ֻȡ��Ч��Χ�ĸ�ҩ;��(�����û��һ��ʱ����ѡ)
            strSql = " And (A.ID IN(Select �÷�ID From �����÷����� Where ��ĿID=[2] And ����>0)" & _
                " Or (Select Count(A.�÷�ID) From �����÷����� A,������ĿĿ¼ B" & _
                    " Where A.�÷�ID=B.ID And B.������� IN([3],3) And A.��ĿID=[2] And A.����>0)<=1)"
        End If
        strSql = "Select Distinct A.ID,A.����,A.����,C.���� as ����" & _
            " From ������Ŀ���� B,������ĿĿ¼ A,���Ʒ���Ŀ¼ C" & _
            " Where A.ID=B.������ĿID And A.����ID=C.ID(+)" & _
            " And A.���='E' And A.��������=[1] And A.������� IN([3],3)" & strSql & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
            " Order by A.����"
        vRect = zlControl.GetControlRect(txt�÷�.Hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, lbl�÷�.Caption, False, "", "", False, False, True, _
            vRect.Left, vRect.Top, txt�÷�.Height, blnCancel, False, True, CStr(int����), .Record(COL_������ĿID).Value, 2)
        If rsTmp Is Nothing Then
            txt�÷�.SetFocus: Exit Sub
        End If

        txt�÷�.SetFocus
        txt�÷�.Text = rsTmp!���� & ""
        txt�÷�.Tag = rsTmp!ID & ""
        Call zlControl.TxtSelAll(txt�÷�)
    End With
End Sub

Private Sub Form_Load()
    Call InitRPTListColumn
    Call InitRPTPathColumn
    '����ͣ����Ŀ
    Call RefreshData
    cmdEdit.Enabled = False
    '�滻��Ŀ��ʼ
    Call InitAdviceTable
    optType(0).Value = True
    Call optType_Click(0)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    '�ָ���
    With picSplit
        .Left = Me.ScaleWidth \ 3
        .Top = picInfo.Height
        .Width = 45
        .Height = Me.ScaleHeight - picInfo.Height
        
    End With
    '�����
    With picLeft
        .Left = 0
        .Top = picInfo.Height
        .Width = picSplit.Left
        .Height = Me.ScaleHeight - picInfo.Height - picBottom.Height
    End With
    '����
    With picTop
        .Left = picSplit.Left + picSplit.Width
        .Top = picInfo.Height
        .Width = Me.ScaleWidth - .Left - 45
    End With
    
    'ҽ����ʾ�� �߶ȹ̶�
    With picAdvice
        .Left = picSplit.Left + picSplit.Width
        .Top = Me.ScaleHeight - picBottom.Height - .Height + 30
        .Width = Me.ScaleWidth - .Left - 45 '
    End With
    
    fraSplit.Move picSplit.Left + picSplit.Width, picAdvice.Top + 45, picAdvice.Width, 45
    '�м�
    With picMain
        .Left = picSplit.Left + picSplit.Width
        .Top = picInfo.Height + picTop.Height
        .Width = Me.ScaleWidth - .Left - 45
        .Height = Me.ScaleHeight - picInfo.Height - picTop.Height - picBottom.Height - picAdvice.Height - fraSplit.Height
    End With
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'����:ж�ش���

    mstrSelItem = ""
    If Not mrsAdvice Is Nothing Then
        Set mrsAdvice = Nothing
    End If
    
End Sub

Private Sub optType_Click(Index As Integer)
    If Index = 1 Then
        lblStopNote.Caption = "·�������õ�����ͣ����Ŀ"
        txtFind.ToolTipText = "������һ��(F3)"
    Else
        lblStopNote.Caption = "·���������õ�������Ŀ"
        txtFind.ToolTipText = "�����롢���Ʋ���������Ŀ"
    End If
    If Me.Visible Then
        txtFind.Text = ""
        txtFind.SetFocus
        Call RefreshData
    End If
End Sub

Private Sub picAdvice_Resize()
    On Error Resume Next
    ucAdvice.Width = picAdvice.ScaleWidth - 90
End Sub

Private Sub picBottom_Resize()
    On Error Resume Next
    cmdBatExe.Left = picBottom.ScaleWidth - (cmdBatExe.Width + cmdQuit.Width + 400)
    cmdQuit.Left = picBottom.ScaleWidth - (cmdQuit.Width + 300)
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    Line1(3).X1 = 0
    Line1(3).X2 = picInfo.ScaleWidth
End Sub

Private Sub picLeft_Resize()
    On Error Resume Next
    fraFind.Move 120, 60, 4455, 900
    lblStopNote.Move 120, picTop.Height - 45
    With rptList
        .Left = 120
        .Top = picTop.Height + lblNote(0).Height - 45
        .Width = picLeft.ScaleWidth - .Left * 2
        .Height = picLeft.ScaleHeight - .Top + 15
    End With
End Sub

Private Sub picMain_Resize()
    On Error Resume Next
   
    '·���嵥
    lblNote(0).Move 0, 0, 1000, 255
    With rptPath
        .Left = 0
        .Top = lblNote(0).Top + lblNote(0).Height
        .Width = picMain.Width - 180
        .Height = picMain.Height - .Top - 100
    End With
End Sub

Private Sub InitRPTListColumn()
    Dim objCol As ReportColumn, lngIdx As Long, i As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    
    With rptList
        '����˳�������(�������Ϊ����)�ı��,Ҫ��Find(�к�)������,���Կ���Record(�к�)����������
        Set objCol = .Columns.Add(COL_����, "����", 60, True)
    
        '���࣬���ƣ����룬��Ʒ�������أ�ҩƷ����
        Set objCol = .Columns.Add(COL_����, "����", 80, True): objCol.Visible = True
        Set objCol = .Columns.Add(COL_����, "����", 200, False): objCol.Visible = True
        
        Set objCol = .Columns.Add(COL_��Ʒ��, "��Ʒ��", 100, True)
        Set objCol = .Columns.Add(COL_����, "����", 75, True)
        Set objCol = .Columns.Add(COL_ҩƷ����, "ҩƷ����", 75, True)
        '������
        Set objCol = .Columns.Add(COL_������ĿID, "������ĿID", 1, True): objCol.Visible = False
        Set objCol = .Columns.Add(COL_�շ�ϸĿID, "�շ�ϸĿID", 1, True): objCol.Visible = False
        Set objCol = .Columns.Add(COL_�������, "�������", 1, True): objCol.Visible = False
        Set objCol = .Columns.Add(COL_��������, "��������", 1, True): objCol.Visible = False
        Set objCol = .Columns.Add(COL_ִ�з���, "ִ�з���", 1, True): objCol.Visible = False
        Set objCol = .Columns.Add(COL_���㷽ʽ, "���㷽ʽ", 1, True): objCol.Visible = False
        Set objCol = .Columns.Add(COL_�걾��λ, "�걾��λ", 0, True): objCol.Visible = False
        Set objCol = .Columns.Add(COL_��鷽��, "��鷽��", 0, True): objCol.Visible = False
        Set objCol = .Columns.Add(COL_����ID, "����ID", 0, True): objCol.Visible = False
        Set objCol = .Columns.Add(COL_���ID, "���ID", 0, True): objCol.Visible = False
        Set objCol = .Columns.Add(COL_����, "����", 0, True): objCol.Visible = False
        
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = objCol.Index = COL_����
        Next
        
        rptList.Populate
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '�з�����ʱ�������߱��ϻ�����һ������
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ��������Ŀ..."
        End With
        
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '������SelectionChanged�¼�
        .ShowItemsInGroups = False 'True ʱ������ʱ,�Զ���������ӵ�������

        
        .GroupsOrder.Add .Columns(COL_����)
        .GroupsOrder(0).SortAscending = True '����֮��,��������в���ʾ,�����е������ǲ����
        
        '����֮�����ʧȥ��¼���е�˳��,���ǿ�м���������
        .SortOrder.Add .Columns(COL_����)
        .SortOrder(0).SortAscending = True
    End With
End Sub


Private Sub InitRPTPathColumn()
    Dim objCol As ReportColumn, lngIdx As Long, i As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    
    With rptPath
        '����˳�������(�������Ϊ����)�ı��,Ҫ��Find(�к�)������,���Կ���Record(�к�)����������
        Set objCol = .Columns.Add(Path_ID, "ID", 0, True)
        objCol.Visible = False
        Set objCol = .Columns.Add(Path_ѡ��, "ѡ��", 60, True)
        objCol.Sortable = False
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentLeft
        objCol.Editable = True
        objCol.Icon = img16.ListImages("UnCheck").Index - 1
        Set objCol = .Columns.Add(Path_����, "����", 100, True)
        objCol.Visible = True
        Set objCol = .Columns.Add(Path_����, "����", 100, True)
        objCol.Visible = True
        Set objCol = .Columns.Add(Path_����, "����", 200, True)
        objCol.Visible = True
        Set objCol = .Columns.Add(Path_�汾, "�汾", 45, True)
        objCol.Visible = True
        Set objCol = .Columns.Add(Path_˵��, "˵��", 200, True)
        objCol.Visible = True
        
        For Each objCol In .Columns
            If objCol.Index <> Path_ѡ�� Then
                objCol.Editable = False
            End If
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '�з�����ʱ�������߱��ϻ�����һ������
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ���ٴ�·��..."
        End With
        
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '������SelectionChanged�¼�
        .ShowItemsInGroups = False
         .SetImageList Me.img16
        
        .GroupsOrder.Add .Columns(Path_����)
        .GroupsOrder(0).SortAscending = True '����֮��,��������в���ʾ,�����е������ǲ����
     
        '����֮�����ʧȥ��¼���е�˳��,���ǿ�м���������
        .SortOrder.Add .Columns(Path_����)
        .SortOrder(0).SortAscending = True
    End With
End Sub

Private Sub LoadStopedItem()
'---------------------------------------
'����:����δͣ��·�������õ�����ͣ����Ŀ��ҩƷ���Ϊ�����Ŀ
'����:
'˵��:
'---------------------------------------
    Dim strSql As String, str��� As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim strIDs As String
    Dim objRecord As ReportRecord
    Dim lng��ID As Long
    Dim str��Id As String
    Dim lngBegin As Long
    Dim strtxt As String
    
    Dim i As Long
    Dim j As Long
    If optType(1).Value Then
        strSql = "Select a.������Ŀid, a.�շ�ϸĿid, f.���, Nvl(g.����, f.����) As ����," & vbNewLine & _
                "       Nvl(g.����, f.����) || Decode(Nvl(g.���, '0'), '0', '', ' ' || g.���) As ����, f.��������, f.ִ�з���, f.���㷽ʽ, f.���㵥λ, g.����," & vbNewLine & _
                "       k.���� As ��Ʒ��, h.ҩƷ����, -null As �걾��λ, -null As ��鷽��, -null As ID, -null As ���id" & vbNewLine & _
                "From ������ĿĿ¼ F, �շ���ĿĿ¼ G, ҩƷ���� H, �շ���Ŀ���� K," & vbNewLine & _
                "     (" & vbNewLine & _
                "     Select Distinct a.������Ŀid, a.�շ�ϸĿid" & vbNewLine & _
                "       From ·��ҽ������ A, ������ĿĿ¼ B, �շ���ĿĿ¼ C, �ٴ�·��ҽ�� E, �ٴ�·����Ŀ F, �ٴ�·���汾 D" & vbNewLine & _
                "       Where a.������Ŀid = b.Id And a.�շ�ϸĿid = c.Id(+) And a.Id = e.ҽ������id And e.·����Ŀid = f.Id And f.·��id = d.·��id And" & vbNewLine & _
                "             f.�汾�� = d.�汾�� And d.ͣ��ʱ�� Is Null And" & vbNewLine & _
                "             ((Nvl(b.����ʱ��, Sysdate) <> To_Date('3000-01-01', 'yyyy-mm-dd') And b.����ʱ�� Is Not Null) Or" & vbNewLine & _
                "             (Nvl(c.����ʱ��, Sysdate) <> To_Date('3000-01-01', 'yyyy-mm-dd') And c.����ʱ�� Is Not Null) Or Exists" & vbNewLine & _
                "              (Select 1" & vbNewLine & _
                "               From ҩƷ���" & vbNewLine & _
                "               Where ҩƷid = a.�շ�ϸĿid And (Nvl(����, 0) = 0 Or Ч�� Is Null Or Ч�� > Trunc(Sysdate)) And ���� = 1" & vbNewLine & _
                "               Group By ҩƷid" & vbNewLine & _
                "               Having Nvl(Sum(��������), 0) <= 0)) And b.��� In ('5', '6', '7')" & vbNewLine & _
                "       ) A" & vbNewLine & _
                "Where a.������Ŀid = f.Id And a.������Ŀid = h.ҩ��id(+) And a.�շ�ϸĿid = g.Id(+) And a.�շ�ϸĿid = k.�շ�ϸĿid(+) And k.����(+) = 3 And" & vbNewLine & _
                "      k.����(+) = 1"
        strSql = strSql & " Union All "
        strSql = strSql & "Select a.������Ŀid, -null As �շ�ϸĿid, a.���, a.����, a.����, a.��������, a.ִ�з���, a.���㷽ʽ, '' As ���㵥λ, '' As ����, '' As ��Ʒ��, '' As ҩƷ����," & vbNewLine & _
                "       a.�걾��λ, a.��鷽��, a.Id, a.���id" & vbNewLine & _
                "From (Select h.·����Ŀid, h.������Ŀid, h.����, h.���, b.����, h.��������, h.ִ�з���, h.���㷽ʽ, a.Id, a.���id, a.���, a.�걾��λ, a.��鷽��" & vbNewLine & _
                "       From (Select Distinct Nvl(a.���id, a.Id) As ��id, b.·����Ŀid, a.������Ŀid, g.���, g.����, g.���� As ����, g.��������, g.ִ�з���, g.���㷽ʽ" & vbNewLine & _
                "              From ·��ҽ������ A, �ٴ�·��ҽ�� B, �ٴ�·����Ŀ C, �ٴ�·���汾 D, ������ĿĿ¼ G" & vbNewLine & _
                "              Where a.Id = b.ҽ������id And b.·����Ŀid = c.Id And c.·��id = d.·��id And c.�汾�� = d.�汾�� And d.ͣ���� Is Null And" & vbNewLine & _
                "                    a.������Ŀid = g.Id And Nvl(g.����ʱ��, Sysdate) <> To_Date('3000-01-01', 'yyyy-mm-dd') And g.����ʱ�� Is Not Null And" & vbNewLine & _
                "                    g.��� In ('D', 'C')) H, ·��ҽ������ A, ������ĿĿ¼ B" & vbNewLine & _
                "       Where (h.��id = a.Id Or h.��id = a.���id) And a.������Ŀid = b.Id" & vbNewLine & _
                "       Order By h.·����Ŀid, a.���) A"
        strSql = strSql & " Union All "
        strSql = strSql & "Select Distinct a.������Ŀid, -null As �շ�ϸĿid, a.���, a.����, a.����, a.��������, a.ִ�з���,a.���㷽ʽ, '' As ���㵥λ, '' As ����, '' As ��Ʒ��, '' As ҩƷ����," & vbNewLine & _
                "                a.�걾��λ, a.��鷽��, a.Id, a.���id  " & vbNewLine & _
                "From (Select a.������Ŀid, g.���, g.����, g.���� As ����, g.��������, g.ִ�з���,g.���㷽ʽ, a.�걾��λ, a.��鷽��, -null As ID, -null As ���id" & vbNewLine & _
                "       From ·��ҽ������ A, �ٴ�·��ҽ�� B, �ٴ�·����Ŀ C, �ٴ�·���汾 D,������ĿĿ¼ G" & vbNewLine & _
                "       Where a.Id = b.ҽ������id And b.·����Ŀid = c.Id And c.·��id = d.·��id And c.�汾�� = d.�汾�� And d.ͣ���� Is Null And a.������Ŀid = g.Id And" & vbNewLine & _
                "             Nvl(g.����ʱ��, Sysdate) <> To_Date('3000-01-01', 'yyyy-mm-dd') And g.����ʱ�� Is Not Null And" & vbNewLine & _
                "             g.��� Not In ('D', 'C', '5', '6', '7')) A"
    Else
        strtxt = Trim(txtFind.Text)
        If strtxt = "" Then
            rptList.Records.DeleteAll
            rptList.Populate
            Exit Sub
        End If
        If zlCommFun.IsCharChinese(strtxt) Then
            strtxt = " And g.���� like [1]"
        Else
            strtxt = " And g.���� like [1]"
        End If
        strSql = "Select a.������Ŀid, a.�շ�ϸĿid, f.���, f.����, Nvl(g.����, f.����) || Decode(Nvl(g.���, '0'), '0', '', ' ' || g.���) As ����, f.��������," & vbNewLine & _
            "       f.ִ�з���, f.���㷽ʽ, f.���㵥λ, g.����, k.���� As ��Ʒ��, h.ҩƷ����, -null As �걾��λ, -null As ��鷽��, -null As ID, -null As ���id" & vbNewLine & _
            "From ������ĿĿ¼ F, �շ���ĿĿ¼ G, ҩƷ���� H, �շ���Ŀ���� K," & vbNewLine & _
            "     (Select Distinct a.������Ŀid, a.�շ�ϸĿid" & vbNewLine & _
            "       From ·��ҽ������ A, �ٴ�·��ҽ�� E, �ٴ�·����Ŀ F, �ٴ�·���汾 D, ������ĿĿ¼ G" & vbNewLine & _
            "       Where a.������Ŀid = G.Id And a.Id = e.ҽ������id And e.·����Ŀid = f.Id And f.·��id = d.·��id And f.�汾�� = d.�汾�� And" & vbNewLine & _
            "             d.ͣ��ʱ�� Is Null And g.��� In ('5', '6', '7')  " & strtxt & ") A" & vbNewLine & _
            "Where a.������Ŀid = f.Id And a.������Ŀid = h.ҩ��id(+) And a.�շ�ϸĿid = g.Id(+) And a.�շ�ϸĿid = k.�շ�ϸĿid(+) And k.����(+) = 3 And" & vbNewLine & _
            "      k.����(+) = 1"
        strSql = strSql & " Union All "
        strSql = strSql & "Select a.������Ŀid, -null As �շ�ϸĿid, a.���, a.����, a.����, a.��������, a.ִ�з���, a.���㷽ʽ, '' As ���㵥λ, '' As ����, '' As ��Ʒ��, '' As ҩƷ����," & vbNewLine & _
            "       a.�걾��λ, a.��鷽��, a.Id, a.���id" & vbNewLine & _
            "From (Select h.·����Ŀid, h.������Ŀid, h.����, h.���, b.����, h.��������, h.ִ�з���, h.���㷽ʽ, a.Id, a.���id, a.���, a.�걾��λ, a.��鷽��" & vbNewLine & _
            "       From (Select Distinct Nvl(a.���id, a.Id) As ��id, b.·����Ŀid, a.������Ŀid, g.���, g.����, g.���� As ����, g.��������, g.ִ�з���, g.���㷽ʽ" & vbNewLine & _
            "              From ·��ҽ������ A, �ٴ�·��ҽ�� B, �ٴ�·����Ŀ C, �ٴ�·���汾 D, ������ĿĿ¼ G" & vbNewLine & _
            "              Where a.Id = b.ҽ������id And b.·����Ŀid = c.Id And c.·��id = d.·��id And c.�汾�� = d.�汾�� And d.ͣ���� Is Null And" & vbNewLine & _
            "                    a.������Ŀid = g.Id And g.��� In ('D', 'C') " & strtxt & ") H, ·��ҽ������ A, ������ĿĿ¼ B" & vbNewLine & _
            "       Where (h.��id = a.Id Or h.��id = a.���id) And a.������Ŀid = b.Id" & vbNewLine & _
            "       Order By h.·����Ŀid, a.���) A"
            strSql = strSql & " Union All "
        strSql = strSql & "Select Distinct a.������Ŀid, -null As �շ�ϸĿid, a.���, a.����, a.����, a.��������, a.ִ�з���, a.���㷽ʽ, '' As ���㵥λ, '' As ����, '' As ��Ʒ��," & vbNewLine & _
                "                '' As ҩƷ����, a.�걾��λ, a.��鷽��, a.Id, a.���id" & vbNewLine & _
                "From (Select a.������Ŀid, g.���, g.����, g.���� As ����, g.��������, g.ִ�з���, g.���㷽ʽ, a.�걾��λ, a.��鷽��, -null As ID, -null As ���id" & vbNewLine & _
                "       From ·��ҽ������ A, �ٴ�·��ҽ�� B, �ٴ�·����Ŀ C, �ٴ�·���汾 D, ������ĿĿ¼ G" & vbNewLine & _
                "       Where a.Id = b.ҽ������id And b.·����Ŀid = c.Id And c.·��id = d.·��id And c.�汾�� = d.�汾�� And d.ͣ���� Is Null And" & vbNewLine & _
                "             a.������Ŀid = g.Id And g.��� Not In ('D', 'C', '5', '6', '7') " & strtxt & ") A"
        
    End If

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UCase(Trim(txtFind.Text)) & "%")
    txtFind.Text = ""
    With rptList
        For i = 1 To rsTmp.RecordCount
            If InStr(",5,6,7,", rsTmp!��� & "") > 1 Then
                str��� = "01-ҩƷ"
            ElseIf rsTmp!��� & "" = "D" Then
                str��� = "02-���"
            ElseIf rsTmp!��� & "" = "C" Then
                str��� = "03-����"
            Else
                str��� = "04-����"
            End If
            Set objRecord = .Records.Add
            objRecord.AddItem str���
            objRecord.AddItem rsTmp!���� & ""
            objRecord.AddItem rsTmp!���� & ""
            objRecord.AddItem rsTmp!��Ʒ�� & ""
            objRecord.AddItem rsTmp!���� & ""
            objRecord.AddItem rsTmp!ҩƷ���� & ""
            objRecord.AddItem rsTmp!������ĿID & ""
            objRecord.AddItem rsTmp!�շ�ϸĿID & ""
            objRecord.AddItem rsTmp!��� & ""
            objRecord.AddItem rsTmp!�������� & ""
            objRecord.AddItem rsTmp!ִ�з��� & ""
            objRecord.AddItem rsTmp!���㷽ʽ & ""
            objRecord.AddItem rsTmp!�걾��λ & ""
            objRecord.AddItem rsTmp!��鷽�� & ""
            objRecord.AddItem rsTmp!ID & ""
            objRecord.AddItem rsTmp!���id & ""
            objRecord.AddItem zlCommFun.SpellCode(NVL(rsTmp!����) & "��0")
            rsTmp.MoveNext
        Next

        '������
        For i = 0 To .Records.count - 1
            str��Id = IIf(.Records.Record(i).Item(COL_���ID).Value <> "", .Records.Record(i).Item(COL_���ID).Value, .Records.Record(i).Item(COL_����ID).Value)
            If .Records.Record(i).Item(COL_�������).Value = "D" Then
                If .Records.Record(i).Item(COL_����ID).Value <> str��Id Then    '��ID
                    .Records.Record(i).Visible = False
                Else
                    .Records.Record(i).Visible = True
                    .Records.Record(i).Item(COL_����).Value = AdviceMakeText(i, "D", strTmp)
                    .Records.Record(i).Tag = strTmp
                    strIDs = strIDs & "," & .Records.Record(i).Item(COL_������ĿID).Value & "_" & i
                End If
            ElseIf .Records.Record(i).Item(COL_�������).Value = "C" Then
                If .Records.Record(i).Item(COL_����ID).Value <> str��Id Then    '��ID
                    If lng��ID <> Val(str��Id) Then
                        lng��ID = str��Id
                        lngBegin = i    '��¼�����е�����
                    End If
                    .Records.Record(i).Visible = False
                Else
                    .Records.Record(i).Visible = True
                    .Records.Record(i).Item(COL_����).Value = AdviceMakeText(i, "C", strTmp, lngBegin)
                    .Records.Record(i).Tag = strTmp
                    strIDs = strIDs & "," & .Records.Record(i).Item(COL_������ĿID).Value & "_" & i
                End If
            End If
        Next
        strIDs = strIDs & ","
        
        For i = 0 To .Records.count - 1
            If .Records.Record(i).Visible And InStr(",C,D,", "," & .Records.Record(i).Item(COL_�������).Value & ",") > 0 Then
                strTmp = .Records.Record(i).Item(COL_������ĿID).Value
                
                If InStr(Mid(strIDs, 1, InStr(strIDs, "," & strTmp & "_" & i)), "," & strTmp & "_") > 0 Then  '�ҵ�ǰ��������ĿID�뵱ǰ������Ŀ��ͬ��
                    For j = i - 1 To 0 Step -1
                        If .Records.Record(j).Visible And .Records.Record(i).Item(COL_�������).Value = .Records.Record(j).Item(COL_�������).Value Then
                            If CompareStr(.Records.Record(i).Tag, .Records.Record(j).Tag) Then  '����������ͬ
                                .Records.Record(j).Item(COL_����ID).Value = .Records.Record(j).Item(COL_����ID).Value & "," & .Records.Record(i).Item(COL_����ID).Value
                                .Records.Record(i).Visible = False
                            End If
                        End If
                    Next
                End If
            End If
        Next
        .Populate
    End With

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub picSplit_KeyPress(KeyAscii As Integer)
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'����:�϶�ͣ����Ŀ�б�
   If Button = 1 Then
        On Error Resume Next
        If picSplit.Left + X < Me.ScaleWidth / 10 Or picSplit.Left + X > Me.ScaleWidth / 10 * 9 Then Exit Sub
        picSplit.Left = picSplit.Left + X
        picLeft.Width = picLeft.Width + X
        picTop.Left = picTop.Left + X: picTop.Width = picTop.Width - X
        picMain.Left = picMain.Left + X: picMain.Width = picMain.Width - X
        picAdvice.Left = picAdvice.Left + X: picAdvice.Width = picAdvice.Width - X
        fraSplit.Left = fraSplit.Left + X: fraSplit.Width = fraSplit.Width - X
    End If
End Sub

Private Sub SetFilterInfo()
'����:���ù�����Ϣ
'
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim lng�շ�ϸĿID As Long
    Dim lng������ĿID As Long
    Dim str��� As String
    Dim blnTooLong As Boolean
    Dim str����ID As String
    Dim strTmp As String
    Dim i As Long, j As Long
    
    On Error GoTo errH
    '���
    Call ClearParaValue
    
    lng�շ�ϸĿID = Val(rptList.SelectedRows(0).Record(COL_�շ�ϸĿID).Value)
    lng������ĿID = Val(rptList.SelectedRows(0).Record(COL_������ĿID).Value)
    str��� = rptList.SelectedRows(0).Record(COL_�������).Value
    str����ID = "," & rptList.SelectedRows(0).Record(COL_����ID).Value & ","
    '�շ�ϸĿIDΪ�ղ�ȡ������ĿID
    
    strSql = "Select d.��������, d.�ܸ�����, d.ִ��Ƶ��, f.���� As �÷�,f.���㷽ʽ, f.Id as ������ĿID  " & _
            " From �ٴ�·���汾 A, �ٴ�·����Ŀ B, �ٴ�·��ҽ�� C, ·��ҽ������ D, ·��ҽ������ E, ������ĿĿ¼ F " & _
            " Where a.·��ID = b.·��id And a.�汾�� = b.�汾�� and A.ͣ���� is null And b.Id = c.·����Ŀid And c.ҽ������id = d.Id And " & _
            IIf("K" = rptList.SelectedRows(0).Record(COL_�������).Value, "d.id = e.���Id(+)", "d.���id = e.Id(+)") & " And e.������Ŀid = f.Id(+) "
    
    If str��� = "C" Or str��� = "D" Then
        If Len(str����ID) > 4000 Then
            blnTooLong = True
        End If
        strSql = strSql & " And Instr([2],','||NVl(d.���Id,d.ID)||',' ) >0 "
    End If
    
    If lng�շ�ϸĿID = 0 Then
        strSql = strSql & " And d.������Ŀid = [1] And Rownum < 2 "
    Else
        strSql = strSql & " And d.�շ�ϸĿID = [1] And Rownum < 2 "
    End If
    
    If blnTooLong Then
        j = 1
        Do While j < Len(str����ID)
            strTmp = Mid(str����ID, j, 4000)
            i = InStrRev(strTmp, ",")
            strTmp = Mid(strTmp, 1, i)
            j = j + i - 1
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, IIf(lng�շ�ϸĿID = 0, lng������ĿID, lng�շ�ϸĿID), strTmp)
            If rsTmp.RecordCount > 0 Then
                Exit Do
            End If
        Loop
        If rsTmp.RecordCount = 0 Then Exit Sub
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, IIf(lng�շ�ϸĿID = 0, lng������ĿID, lng�շ�ϸĿID), str����ID)
        If rsTmp.RecordCount = 0 Then Exit Sub
    End If
    
    With rptList.SelectedRows(0)
        If InStr(",5,6,C,", "," & .Record(COL_�������).Value & ",") > 0 Then
            txt����.Text = FormatEx(NVL(rsTmp!��������), 4)
            txt����.Text = FormatEx(NVL(rsTmp!�ܸ�����), 4)
            txt�÷�.Text = rsTmp!�÷� & "": txt�÷�.Tag = rsTmp!������ĿID & ""
            txtƵ��.Text = rsTmp!ִ��Ƶ�� & ""
        ElseIf .Record(COL_�������).Value = "D" Then
            txt����.Text = rsTmp!�ܸ����� & ""
            txtƵ��.Text = rsTmp!ִ��Ƶ�� & ""
        ElseIf InStr(",1,2,", "," & .Record(COL_���㷽ʽ).Value & ",") > 0 Then '1-������2-��ʱ
            txt����.Text = FormatEx(NVL(rsTmp!��������), 4)
            txt����.Text = FormatEx(NVL(rsTmp!�ܸ�����), 4)
            txtƵ��.Text = rsTmp!ִ��Ƶ�� & ""
        End If
    End With
    If lng�շ�ϸĿID = 0 Then
        strSql = "Select a.��� As ���id, a.���㷽ʽ, a.ִ��Ƶ��, a.���㵥λ, NULL as סԺ��λ From ������ĿĿ¼ A Where a.Id = [1]"
    Else
        strSql = "Select a.��� As ���id, a.���㷽ʽ, a.ִ��Ƶ��, a.���㵥λ, b.סԺ��λ" & vbNewLine & _
                "From ������ĿĿ¼ A, ҩƷ��� B" & vbNewLine & _
                "Where a.Id = [1] And b.ҩƷid = [2]"

    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng������ĿID, lng�շ�ϸĿID)
    If rsTmp.RecordCount = 0 Then Exit Sub
    '������λ
    If txt����.Text = "" Then '����
        If InStr(",5,6,", rsTmp!���ID) > 0 Or InStr(",1,2,", NVL(rsTmp!���㷽ʽ, 0)) > 0 Then
            lbl������λ.Caption = NVL(rsTmp!���㵥λ)   'ҩƷΪ������λ
        End If
    Else
        If InStr(",5,6,", rsTmp!���ID) > 0 Or (NVL(rsTmp!ִ��Ƶ��, 0) = 0 And InStr(",1,2,", NVL(rsTmp!���㷽ʽ, 0)) > 0) Then
            lbl������λ.Caption = NVL(rsTmp!���㵥λ)   'ҩƷΪ������λ
        End If
    End If

    '������λ
    If txt����.Text <> "" Then '����
        If InStr(",5,6,", rsTmp!���ID) > 0 Then
            '�С�����ҩ������������λ����סԺ��λ
            lbl������λ.Caption = rsTmp!סԺ��λ & ""
        ElseIf rsTmp!���ID = "4" Then
            lbl������λ.Caption = rsTmp!סԺ��λ & ""  'ɢװ��λ
        Else
            '��������Ҫ��������
            '���Ϊһ���Ի�ƴ�����ȱʡ����Ϊ1
            If NVL(rsTmp!ִ��Ƶ��, 0) = 1 Or NVL(rsTmp!���㷽ʽ, 0) = 3 Then
               lbl������λ.Caption = 1
            End If
            lbl������λ.Caption = NVL(rsTmp!���㵥λ)
        End If
    End If
        
    mblnChange = False
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub rptList_SelectionChanged()
    Dim objItem As ReportRecordItem
    Dim strTmp As String
    
    cmdFind.Enabled = False
    If rptList.SelectedRows.count = 0 Then Exit Sub  '���������
    If rptList.SelectedRows(0).GroupRow Then
        SetEditable -1, -1, -1, -1, -1, -1, -1
        Call ClearParaValue
        Call ClearPath
        Exit Sub
    End If
    With rptList.SelectedRows(0)
        '
        strTmp = "�÷�" 'ȱʡ����Ϊ�÷�
        If InStr(",5,6,7,", "," & .Record(COL_�������).Value & ",") > 0 Then
            SetEditable 1, 1, 1, 1, 1, 1, 1
        ElseIf .Record(COL_�������).Value = "D" Then
            SetEditable -1, -1, -1, 1, -1, -1, 1
        ElseIf .Record(COL_�������).Value = "C" Then
            strTmp = "�ɼ���ʽ"
            SetEditable -1, -1, 1, 1, -1, 1, 1
        ElseIf InStr(",1,2,", "," & .Record(COL_���㷽ʽ).Value & ",") > 0 Then   '1-������2-��ʱ
            SetEditable 1, 1, -1, 1, 1, -1, 1
        Else
            SetEditable -1, -1, -1, -1, -1, -1, -1
        End If
        lbl�÷�.Caption = strTmp
        If .Record(COL_�������).Value = "C" Then
            chkPra(chk_�÷�).Caption = "ֻ�滻�ɼ���ͬ��"
        Else
            chkPra(chk_�÷�).Caption = "ֻ�滻�÷���ͬ��"
        End If
        
        If mblnChange Or mstrSelItem <> .Record(COL_������ĿID).Value & "_" & .Record(COL_�շ�ϸĿID).Value & "_" & .Record.Index Then
            'ѡ����Ŀ�л�ʱ,��Ҫ�������
            Call ClearPath
            mstrSelItem = .Record(COL_������ĿID).Value & "_" & .Record(COL_�շ�ϸĿID).Value & "_" & .Record.Index
            Call SetFilterInfo
        End If
    End With
    cmdFind.Enabled = True
End Sub

Private Sub rptPath_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objColumn As ReportColumn
    Dim i As Long
    
    '��������ͷ��ͼƬ����ѡ��ȫ��
    If Button = 1 Then
        If rptPath.HitTest(X, Y).ht = xtpHitTestHeader Then
            Set objColumn = rptPath.HitTest(X, Y).Column
            If Not objColumn Is Nothing Then
                If objColumn.Index = Path_ѡ�� Then
                    If rptPath.Columns(Path_ѡ��).Icon = img16.ListImages("Check").Index - 1 Then
                        rptPath.Columns(Path_ѡ��).Icon = img16.ListImages("UnCheck").Index - 1
                        For i = 0 To rptPath.Records.count - 1
                            rptPath.Records(i)(Path_ѡ��).Checked = False
                        Next
                    Else
                        rptPath.Columns(Path_ѡ��).Icon = img16.ListImages("Check").Index - 1
                        For i = 0 To rptPath.Records.count - 1
                            rptPath.Records(i)(Path_ѡ��).Checked = True
                        Next
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub rptPath_SortOrderChanged()
    Dim objCol As ReportColumn
    '����ʱ��ǿ���Ȱ���������
    '������������Ч����������һ������
    If rptPath.SortOrder.count = 1 Then
        If rptPath.SortOrder(0).Index <> Path_���� Then
            Set objCol = rptPath.SortOrder(0)
            rptPath.SortOrder.DeleteAll
            rptPath.SortOrder.Add rptPath.Columns(Path_����)
            rptPath.SortOrder.Add objCol
        End If
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And txtFind.Text <> "" Then
        If optType(1).Value Then
            Call FindRPTList(True)
        Else
            Call RefreshData
        End If
        txtFind.SetFocus '��λ�����ҿ�
    End If
End Sub

Private Sub txtFind_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 And txtFind.Text <> "" Then
        Call FindRPTList(True)
        txtFind.SetFocus '��λ�����ҿ�
    End If
End Sub

Private Sub txt����_Change()
    mblnChange = True
    If rptPath.Records.count > 0 Then
        Call ClearPath
    End If
End Sub

Private Sub txtƵ��_Change()
    mblnChange = True
    If rptPath.Records.count > 0 Then
        Call ClearPath
    End If
End Sub

Private Sub txtƵ��_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim str��Χ As String, intƵ�� As Integer, vRect As RECT
    Dim lng������ĿID As Long
    
    If rptList.SelectedRows.count = 0 Then Exit Sub  '���������
    With rptList.SelectedRows(0)
        If KeyAscii = 13 Then
            KeyAscii = 0
            If txtƵ��.Text = "" Then
                If cmdƵ��.Enabled And cmdƵ��.Visible Then cmdƵ��_Click
            Else
                intƵ�� = Get��ĿƵ��
                If txt����.Text <> "" Then '����
                    If .Record(COL_�������).Value <> "7" And intƵ�� = 0 Then
                        str��Χ = "1,-1" '��������Ϊһ����
                    Else
                        str��Χ = GetƵ�ʷ�Χ(intƵ��)
                    End If
                Else
                    str��Χ = GetƵ�ʷ�Χ(intƵ��)
                    intƵ�� = intƵ�� = Decode(str��Χ, "1", 0, "2", 0, "-1", 1, "-2", 2, "-3", 1, "-5", 1)
                End If
                
                '��ѡ��Ƶ�ʵĳ���Ƶ��
                lng������ĿID = Val(.Record(COL_������ĿID).Value)
                strSql = ""
                If InStr("," & str��Χ & ",", ",1,") > 0 Then
                    strSql = " And (Exists(Select 1 From �����÷����� Where ��ĿID=[4] And �÷�ID is NULL And Ƶ��=A.���� And A.���÷�Χ=1)" & _
                        " Or (Select Count(*) From �����÷����� Where ��ĿID=[4] And �÷�ID is NULL And Ƶ�� Is Not NULL)<=1)"
                End If
                strSql = _
                    " Select Rownum as ID,A.����,A.����,A.����," & _
                    " A.Ӣ������,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,A.���÷�Χ as ��ΧID" & _
                    " From ����Ƶ����Ŀ A" & _
                    " Where (Instr([3],','||A.���÷�Χ||',')>0   Or a.���÷�Χ=[5])" & strSql & _
                    " And (A.���� Like [1] Or Upper(A.����) Like [2]" & _
                    " Or Upper(A.����) Like [2] Or Upper(A.Ӣ������) Like [2])" & _
                    " Order by A.���÷�Χ,A.����"
                vRect = zlControl.GetControlRect(txtƵ��.Hwnd)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "����Ƶ��", False, "", "", False, False, True, _
                    vRect.Left, vRect.Top, txtƵ��.Height, blnCancel, False, True, UCase(txtƵ��.Text) & "%", _
                    gstrLike & UCase(txtƵ��.Text) & "%", "," & str��Χ & ",", lng������ĿID, IIf(txt����.Text <> "", -5, -3))
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "δ�ҵ�ƥ�������Ƶ����Ŀ��", vbInformation, gstrSysName
                    End If
                    Call zlControl.TxtSelAll(txtƵ��)
                    txtƵ��.SetFocus: Exit Sub
                End If
                txtƵ��.Text = rsTmp!���� & ""
                Call zlControl.TxtSelAll(txtƵ��)
                txtƵ��.SetFocus
            End If
        End If
    End With
End Sub

Private Sub txt�÷�_Change()
    mblnChange = True
    If rptPath.Records.count > 0 Then
        Call ClearPath
    End If
End Sub

Private Sub txt�÷�_KeyPress(KeyAscii As Integer)
    Dim int���� As Integer
    Dim strSql As String
    Dim strLike As String
    Dim vRect As RECT
    Dim blnCancel As Boolean
    Dim rsTmp As ADODB.Recordset
    
    If rptList.SelectedRows.count = 0 Then Exit Sub  '���������
    If rptList.SelectedRows(0).GroupRow Then Exit Sub
    If KeyAscii = 13 Then
        KeyAscii = 0
        With rptList.SelectedRows(0)
            If txt�÷�.Text = "" Then
                If cmd�÷�.Enabled And cmd�÷�.Visible Then cmd�÷�_Click
            Else
                If InStr(",5,6,", .Record(COL_�������).Value) > 0 Then
                    int���� = 2 '��ҩ;��
                ElseIf .Record(COL_�������).Value = "C" Then
                    int���� = 6 '�ɼ�����
                ElseIf .Record(COL_�������).Value = "K" Then
                    int���� = 8 '��Ѫ;��
                Else
                    int���� = 4 '��ҩ�÷�
                End If
                If int���� = 2 Then 'ֻȡ��Ч��Χ�ĸ�ҩ;��(�����û��һ��ʱ����ѡ)
                    strSql = " And (A.ID IN(Select �÷�ID From �����÷����� Where ��ĿID=[4] And ����>0)" & _
                        " Or (Select Count(A.�÷�ID) From �����÷����� A,������ĿĿ¼ B" & _
                            " Where A.�÷�ID=B.ID And B.������� IN([6],3) And A.��ĿID=[4] And A.����>0)<=1)"
                End If
                
                '�Ż�
                strLike = gstrLike
                If Len(txt�÷�.Text) < 2 Then strLike = ""
                
                strSql = "Select Distinct A.ID,A.����,A.����" & _
                    " From ������ĿĿ¼ A,������Ŀ���� B" & _
                    " Where A.ID=B.������ĿID" & _
                    " And A.���='E' And A.��������=[3] And A.������� IN([6],3)" & strSql & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                    " And (A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2])" & _
                    Decode(gint����, 0, " And B.���� IN([5],3)", 1, " And B.���� IN([5],3)", "") & _
                    " Order by A.����"
                vRect = zlControl.GetControlRect(txt�÷�.Hwnd)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, lbl�÷�.Caption, False, "", "", False, False, True, _
                    vRect.Left, vRect.Top, txt�÷�.Height, blnCancel, False, True, UCase(txt�÷�.Text) & "%", _
                    strLike & UCase(txt�÷�.Text) & "%", CStr(int����), Val(.Record(COL_������ĿID).Value), gint���� + 1, 2)
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "δ�ҵ�ƥ���" & lbl�÷�.Caption & "��", vbInformation, gstrSysName
                    End If
                    Call zlControl.TxtSelAll(txt�÷�)
                    txt�÷�.SetFocus: Exit Sub
                End If
                txt�÷�.SetFocus
                txt�÷�.Text = rsTmp!���� & ""
                txt�÷�.Tag = rsTmp!ID & ""
                Call zlControl.TxtSelAll(txt�÷�)
            End If
        End With
    End If
End Sub

Private Sub txt����_Change()
    mblnChange = True
    If rptPath.Records.count > 0 Then
        Call ClearPath
    End If
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Function GetƵ�ʷ�Χ(ByVal lngƵ������ As Long) As Integer
    Dim lngFind As Long
    
    With rptList.SelectedRows(0)
        If .Record(COL_�������).Value = "7" Then
            GetƵ�ʷ�Χ = 2 '��ҽ
        Else
            If lngƵ������ = 0 Then
                GetƵ�ʷ�Χ = 1 '��ѡƵ�ʵ���Ŀʹ����ҽƵ����Ŀ
            ElseIf lngƵ������ = 1 Then
                GetƵ�ʷ�Χ = -1 'һ����
            ElseIf lngƵ������ = 2 Then
                GetƵ�ʷ�Χ = -2 '������
            End If
        End If
    End With
End Function

Private Function Get��ĿƵ��() As Integer
'���ܣ���ȡָ����Ŀ��ԭʼִ��Ƶ������
'������lngRow=��ǰ�ɼ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select ִ��Ƶ�� From ������ĿĿ¼ Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(rptList.SelectedRows(0).Record(COL_������ĿID).Value))
    If Not rsTmp.EOF Then Get��ĿƵ�� = NVL(rsTmp!ִ��Ƶ��, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetEditable(Optional int���� As Integer, Optional int���� As Integer, _
    Optional int�÷� As Integer, Optional intƵ�� As Integer, Optional intPra���� As Integer, _
    Optional intPra�÷� As Integer, Optional intPraƵ�� As Integer)
'���ܣ�����ָ���༭��Ŀ���״̬
'������0-���ֲ���,-1-��ֹ,1-����
    
    If int���� = 1 Then
        txt����.Enabled = True
        txt����.BackColor = Color_Enabled
        lbl������λ.Visible = True
    ElseIf int���� = -1 Then
        txt����.Enabled = False
        txt����.BackColor = Color_UNEnabled
        lbl������λ.Visible = False
    End If
    
    If int���� = 1 Then
        txt����.Enabled = True
        txt����.BackColor = Color_Enabled
    ElseIf int���� = -1 Then
        txt����.Enabled = False
        txt����.BackColor = Color_UNEnabled
    End If
    
    If intƵ�� = 1 Then
        txtƵ��.Enabled = True
        txtƵ��.BackColor = Color_Enabled
        cmdƵ��.Enabled = True
    ElseIf intƵ�� = -1 Then
        txtƵ��.Enabled = False
        cmdƵ��.Enabled = False
        txtƵ��.BackColor = Color_UNEnabled
    End If
    
    If int�÷� = 1 Then
        cmd�÷�.Enabled = True
        txt�÷�.Enabled = True
        txt�÷�.BackColor = Color_Enabled
    ElseIf int�÷� = -1 Then
        cmd�÷�.Enabled = False
        txt�÷�.Enabled = False
        txt�÷�.BackColor = Color_UNEnabled
    End If
    
    If intPra���� = 1 Then
        chkPra(chk_����).Enabled = True
        chkPra(chk_����).Value = Checked
    ElseIf intPra���� = -1 Then
        chkPra(chk_����).Enabled = False
        chkPra(chk_����).Value = Unchecked
    End If
    
    If intPra�÷� = 1 Then
        chkPra(chk_�÷�).Enabled = True
        chkPra(chk_�÷�).Value = Checked
    ElseIf intPra�÷� = -1 Then
        chkPra(chk_�÷�).Enabled = False
        chkPra(chk_�÷�).Value = Unchecked
    End If
    
    If intPraƵ�� = 1 Then
        chkPra(chk_Ƶ��).Enabled = True
        chkPra(chk_Ƶ��).Value = Checked
    ElseIf intPraƵ�� = -1 Then
        chkPra(chk_Ƶ��).Enabled = False
        chkPra(chk_Ƶ��).Value = Unchecked
    End If
    
End Sub

Private Function ShowAdvice() As Boolean
'���ܣ���ʾ·����Ŀ��Ӧ��ҽ������(�ٴ�·����Ŀ�༭)
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim i As Long, j As Long
    
    strSql = ""
    '���ɶ�̬SQL
    With mrsAdvice
        .Filter = ""
        Do While Not .EOF
            strSql = strSql & " Union ALL Select "
            For i = 0 To .Fields.count - 1
                If Not IsNull(.Fields(i).Value) Then
                    If Rec.IsType(.Fields(i).Type, adVarChar) Then
                        strSql = strSql & "'" & Replace(Replace(.Fields(i).Value, "[", "("), "]", ")") & "'"
                    Else
                        strSql = strSql & .Fields(i).Value 'û��������
                    End If
                Else
                    If Rec.IsType(.Fields(i).Type, adBigInt) Or Rec.IsType(.Fields(i).Type, adSmallInt) Or Rec.IsType(.Fields(i).Type, adSingle) Then
                        strSql = strSql & "-Null"
                    Else
                        strSql = strSql & "Null"
                    End If
                End If
                strSql = strSql & " As " & .Fields(i).Name & ","
            Next
            strSql = Left(strSql, Len(strSql) - 1) & " From Dual"
            .MoveNext
        Loop
        .Filter = ""
        strSql = Mid(strSql, 12)
    End With
    
    If strSql = "" Then
        Call ucAdvice.ShowAdvice(4, "", 0, 0, 0)
    Else
        Call ucAdvice.ShowAdvice(4, strSql, 0, 0, 0)
    End If
    ShowAdvice = True
End Function

Private Sub InitAdviceRecordset()
    If Not mrsAdvice Is Nothing Then
        If mrsAdvice.State = 1 Then mrsAdvice.Close
    End If
    Set mrsAdvice = New ADODB.Recordset
    
    mrsAdvice.Fields.Append "ID", adBigInt
    mrsAdvice.Fields.Append "�Ƿ�ȱʡ", adSmallInt
    mrsAdvice.Fields.Append "�Ƿ�ѡ", adSmallInt
    mrsAdvice.Fields.Append "���ID", adBigInt, , adFldIsNullable
    mrsAdvice.Fields.Append "���", adBigInt
    mrsAdvice.Fields.Append "��Ч", adSmallInt
    mrsAdvice.Fields.Append "������ĿID", adBigInt, , adFldIsNullable
    mrsAdvice.Fields.Append "�շ�ϸĿID", adBigInt, , adFldIsNullable
    mrsAdvice.Fields.Append "ҽ������", adVarChar, 1000, adFldIsNullable
    mrsAdvice.Fields.Append "��������", adSingle, , adFldIsNullable
    mrsAdvice.Fields.Append "�ܸ�����", adSingle, , adFldIsNullable
    mrsAdvice.Fields.Append "�걾��λ", adVarChar, 100, adFldIsNullable
    mrsAdvice.Fields.Append "��鷽��", adVarChar, 100, adFldIsNullable
    mrsAdvice.Fields.Append "ҽ������", adVarChar, 1000, adFldIsNullable
    mrsAdvice.Fields.Append "ִ��Ƶ��", adVarChar, 100, adFldIsNullable
    mrsAdvice.Fields.Append "Ƶ�ʴ���", adSmallInt, , adFldIsNullable
    mrsAdvice.Fields.Append "Ƶ�ʼ��", adSmallInt, , adFldIsNullable
    mrsAdvice.Fields.Append "�����λ", adVarChar, 10, adFldIsNullable
    mrsAdvice.Fields.Append "ִ������", adSmallInt
    mrsAdvice.Fields.Append "ִ�п���ID", adBigInt, , adFldIsNullable
    mrsAdvice.Fields.Append "ʱ�䷽��", adVarChar, 100, adFldIsNullable
    mrsAdvice.Fields.Append "�䷽ID", adBigInt, , adFldIsNullable
    mrsAdvice.Fields.Append "�����ĿID", adBigInt, , adFldIsNullable
    mrsAdvice.Fields.Append "ִ�б��", adSingle, , adFldIsNullable
    
    mrsAdvice.CursorLocation = adUseClient
    mrsAdvice.LockType = adLockOptimistic
    mrsAdvice.CursorType = adOpenStatic
    mrsAdvice.Open
End Sub

Private Sub InitSchemeRecordset(rsScheme As ADODB.Recordset)
    Set rsScheme = New ADODB.Recordset
    rsScheme.Fields.Append "�Ƿ�ѡ", adSmallInt
    rsScheme.Fields.Append "�Ƿ�ȱʡ", adSmallInt
    rsScheme.Fields.Append "���", adBigInt
    rsScheme.Fields.Append "������", adBigInt, , adFldIsNullable
    rsScheme.Fields.Append "��Ч", adSmallInt
    rsScheme.Fields.Append "������ĿID", adBigInt, , adFldIsNullable
    rsScheme.Fields.Append "�շ�ϸĿID", adBigInt, , adFldIsNullable
    rsScheme.Fields.Append "ҽ������", adVarChar, 1000, adFldIsNullable
    rsScheme.Fields.Append "����", adSingle, , adFldIsNullable
    rsScheme.Fields.Append "��������", adSingle, , adFldIsNullable
    rsScheme.Fields.Append "�ܸ�����", adSingle, , adFldIsNullable
    rsScheme.Fields.Append "ҽ������", adVarChar, 1000, adFldIsNullable
    rsScheme.Fields.Append "ִ��Ƶ��", adVarChar, 100, adFldIsNullable
    rsScheme.Fields.Append "Ƶ�ʴ���", adSmallInt, , adFldIsNullable
    rsScheme.Fields.Append "Ƶ�ʼ��", adSmallInt, , adFldIsNullable
    rsScheme.Fields.Append "�����λ", adVarChar, 10, adFldIsNullable
    rsScheme.Fields.Append "ʱ�䷽��", adVarChar, 100, adFldIsNullable
    rsScheme.Fields.Append "ִ�п���ID", adBigInt, , adFldIsNullable
    rsScheme.Fields.Append "ִ������", adSmallInt
    rsScheme.Fields.Append "�걾��λ", adVarChar, 100, adFldIsNullable
    rsScheme.Fields.Append "��鷽��", adVarChar, 100, adFldIsNullable
    rsScheme.Fields.Append "�䷽ID", adBigInt, , adFldIsNullable
    rsScheme.Fields.Append "�����ĿID", adBigInt, , adFldIsNullable
    rsScheme.Fields.Append "ִ�б��", adSingle, , adFldIsNullable
    
    rsScheme.CursorLocation = adUseClient
    rsScheme.LockType = adLockOptimistic
    rsScheme.CursorType = adOpenStatic
    rsScheme.Open
End Sub

Private Function SaveData(ByVal strPath As String, ByVal strPathTag As String) As Boolean
'����:��������滻
'����:
'      strPath    ��:·��ID1,�汾��1:·��ID2,�汾��2:....
'      strPathTag ����: ·��Id_���°汾1,·��Id_���°汾2:...
    Dim str��IDs As String, lng��� As Long
    Dim lng������ĿID As Long, lng�շ�ϸĿID As Long
    Dim str��� As String
    Dim strSql As String
    Dim i As Long, j As Long
    Dim strTmp As String
    Dim arrSQL As Variant
    Dim blnTran As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strItemIDs As String  '��¼·����ĿID
    Dim strItemAdvices As String   '����:·����ĿID1,��ҽ��ID1:·����ĿID2,��ҽ��ID2...
    Dim colItem As Collection
    Dim colAdvice As Collection
    
    On Error GoTo errH
    
    With rptList.SelectedRows(0)
        lng������ĿID = Val(.Record(COL_������ĿID).Value)
        lng�շ�ϸĿID = Val(.Record(COL_�շ�ϸĿID).Value)
        str��� = .Record(COL_�������).Value
        '�ҵ��滻����Ŀ
        If InStr(",5,6,", str���) > 0 Then
            strSql = "Select *" & vbNewLine & _
            "From (Select Distinct /*+cardinality(A,10)*/ a.C1 as ·��ID,a.C2 as �汾��,c.·����ĿID,e.��Ч,e.Id As ����id, E.���id, E.������Ŀid, E.�շ�ϸĿid, E.��� " & vbNewLine & _
            "From Table(f_Num2list2([1], ':', ',')) A, �ٴ�·����Ŀ B, �ٴ�·��ҽ�� C, ·��ҽ������ D, ·��ҽ������ E" & vbNewLine & _
            "Where a.C1 = b.·��id And a.C2 = b.�汾�� And b.Id = c.·����Ŀid And c.ҽ������id = d.Id " & vbNewLine & _
            IIf(lng�շ�ϸĿID = 0, " And d.������ĿID =[2]", " And d.�շ�ϸĿID =[2]") & _
            " And (d.���id = e.���id Or d.���id = e.Id)"
                
            If chkPra(chk_����).Value = vbChecked Then
                If txt����.Text <> "" Then
                    strSql = strSql & " And d.�������� =[3] "
                End If
                If txt����.Text <> "" Then
                    strSql = strSql & " And d.�ܸ����� = [4] "
                End If
            End If
            
            If chkPra(chk_�÷�).Value = vbChecked Then
                If txt�÷�.Text <> "" Then
                    strSql = strSql & " and exists (select 1 from ·��ҽ������  H where H.id=d.���id and H.������Ŀid = [5]) "
                End If
            End If
            
            If chkPra(chk_Ƶ��).Value = vbChecked Then
                If txtƵ��.Text <> "" Then
                    strSql = strSql & " And d.ִ��Ƶ�� =[6] "
                End If
            End If
            strSql = strSql & ") ��order By ·����Ŀid, ���"
        ElseIf InStr(",D,C,", str���) > 0 Then
            '��ȡ����ҽ��ID
            str��IDs = IIf(.Record(COL_���ID).Value = "", .Record(COL_����ID).Value, .Record(COL_���ID).Value)
            strSql = "Select /* +Rule */" & vbNewLine & _
                " a.C1 as  ·��ID,a.C2 as �汾��,c.·����Ŀid, c.ҽ������id as ����ID" & vbNewLine & _
                "From Table(f_Num2list2([1], ':', ',')) A, �ٴ�·����Ŀ B, �ٴ�·��ҽ�� C" & vbNewLine & _
                "Where a.C1 = b.·��id And a.C2 = b.�汾�� And b.Id = c.·����Ŀid And Instr([7], ','||c.ҽ������id||',') > 0"

        Else '������� ����ҽ��,����ҽ���������滻��
            strSql = "Select /*+ RULE*/" & vbNewLine & _
                    "a.C1 as  ·��ID,a.C2 as �汾��,c.·����ĿID,d.Id As ����id" & vbNewLine & _
                    "From Table(f_Num2list2([1], ':', ',')) A, �ٴ�·����Ŀ B, �ٴ�·��ҽ�� C, ·��ҽ������ D" & vbNewLine & _
                    "Where a.C1 = b.·��id And a.C2 = b.�汾�� And b.Id = c.·����Ŀid And c.ҽ������id = d.Id " & _
                    IIf(lng�շ�ϸĿID = 0, " And d.������ĿID =[2]", " And d.�շ�ϸĿID =[2]")
            If chkPra(chk_����).Value = vbChecked And InStr(",1,2,", "," & rptList.SelectedRows(0).Record(COL_���㷽ʽ).Value & ",") > 0 Then   '������ʱ
               If txt����.Text <> "" Then
                   strSql = strSql & " And d.�������� =[3] "
               End If
               If txt����.Text <> "" Then
                   strSql = strSql & " And d.�ܸ����� = [4] "
               End If
            End If
            
            If chkPra(chk_Ƶ��).Value = vbChecked Then
                If txtƵ��.Text <> "" Then
                    strSql = strSql & " And d.ִ��Ƶ�� =[6] "
                End If
            End If
        End If
  
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strPath, IIf(lng�շ�ϸĿID = 0, lng������ĿID, lng�շ�ϸĿID), _
            Val(txt����.Text), Val(txt����.Text), Val(txt�÷�.Tag), txtƵ��.Text, "," & str��IDs & ",")
        If rsTmp.RecordCount = 0 Then
            MsgBox "û���ҵ�����������ͣ����Ŀ,�滻ʧ��!", vbInformation + vbOKOnly, "�����滻"
            Exit Function
        End If
        'ֻ�������°汾�����һ����˰汾�������䶯��¼
        Set colItem = New Collection: Set colAdvice = New Collection
        strItemIDs = "": strItemAdvices = ""
        For i = 1 To rsTmp.RecordCount
            If Not InStr(strItemIDs & ",", "," & rsTmp!·����ĿID & ",") > 0 And InStr(strPathTag, rsTmp!·��ID & "_" & rsTmp!�汾��) > 0 Then  '����·��ҽ���䶯����ĿID
                If Len(strItemIDs & "," & rsTmp!·����ĿID) > 4000 Then
                    colItem.Add Mid(strItemIDs, 2)
                    strItemIDs = "," & rsTmp!·����ĿID
                Else
                    strItemIDs = strItemIDs & "," & rsTmp!·����ĿID
                End If
            End If
            If str��� = "D" Or str��� = "C" Then
                If Len(strItemAdvices & ":" & rsTmp!·����ĿID & "," & rsTmp!����ID) > 4000 Then
                    colAdvice.Add Mid(strItemAdvices, 2)
                    strItemAdvices = ":" & rsTmp!·����ĿID & "," & rsTmp!����ID
                Else
                    strItemAdvices = strItemAdvices & ":" & rsTmp!·����ĿID & "," & rsTmp!����ID
                End If
            End If
            rsTmp.MoveNext
        Next
        If strItemIDs <> "" Then
            colItem.Add Mid(strItemIDs, 2)
        End If
        If strItemAdvices <> "" Then
            colAdvice.Add Mid(strItemAdvices, 2)
        End If
    End With
    arrSQL = Array()
    With mrsAdvice
        For i = 1 To colItem.count
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_·��ҽ���䶯_Insert('" & colItem(i) & "'," & "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')" & ",'" & UserInfo.���� & "')"
        Next
        rsTmp.MoveFirst
        If InStr(",5,6,", str���) > 0 Then
            For i = 1 To rsTmp.RecordCount
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                If rsTmp!���id & "" = "" Then
                    .Filter = "���Id=NULL"
                Else
                    .Filter = "���Id<>NULL"
                End If
                If (InStr(",5,6,", str���) > 0 And lng������ĿID = Val(rsTmp!������ĿID & "")) Then
                    arrSQL(UBound(arrSQL)) = "zl_·��ҽ������_Update(1," & rsTmp!����ID & "," & ZVal(NVL(!������ĿID)) & "," & ZVal(NVL(!�շ�ϸĿID)) & ",'" & !ҽ������ & "'," & _
                    ZVal(NVL(!��������)) & "," & IIf(rsTmp!��Ч = 1, ZVal(NVL(!�ܸ�����)), "NULL") & ",'" & !�걾��λ & "','" & !��鷽�� & "','" & !ҽ������ & "','" & !ִ��Ƶ�� & "'," & _
                    ZVal(NVL(!Ƶ�ʴ���)) & "," & ZVal(NVL(!Ƶ�ʼ��)) & ",'" & !�����λ & "','" & !ʱ�䷽�� & "')"
                Else
                    If rsTmp!���id & "" = "" Then
                    '�÷�
                        arrSQL(UBound(arrSQL)) = "zl_·��ҽ������_Update( 1," & rsTmp!����ID & "," & ZVal(NVL(!������ĿID)) & ", NULL,NULL,NULL,NULL,NULL,NULL,NULL,'" & !ִ��Ƶ�� & "'," & _
                                     ZVal(NVL(!Ƶ�ʴ���)) & "," & ZVal(NVL(!Ƶ�ʼ��)) & ",'" & !�����λ & "','" & !ʱ�䷽�� & "')"
                    Else
                    'һ����ҩ,�滻һ����Ŀ
                        arrSQL(UBound(arrSQL)) = "zl_·��ҽ������_Update(1," & rsTmp!����ID & ",Null, NULL,NULL,NULL,NULL,NULL,NULL,NULL,'" & !ִ��Ƶ�� & "'," & _
                                     ZVal(NVL(!Ƶ�ʴ���)) & "," & ZVal(NVL(!Ƶ�ʼ��)) & ",'" & !�����λ & "','" & !ʱ�䷽�� & "')"
                    End If
                End If
            
                rsTmp.MoveNext
            Next
        ElseIf InStr(",D,C,", str���) > 0 Then
            '��������ҽ���Ȳ���[·��ҽ������]
            strTmp = "": .MoveFirst
            For i = 1 To .RecordCount
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                lng��� = lng��� + 1
                arrSQL(UBound(arrSQL)) = "Zl_·��ҽ������_Insert(" & _
                        !ID & "," & ZVal(NVL(!���id, 0)) & "," & !��� & "," & !��Ч & "," & _
                        ZVal(NVL(!������ĿID, 0)) & ",'" & NVL(!ҽ������) & "'," & ZVal(NVL(!��������, 0)) & "," & _
                        ZVal(NVL(!�ܸ�����, 0)) & "," & ZVal(NVL(!�շ�ϸĿID, 0)) & ",'" & NVL(!�걾��λ) & "'," & _
                        "'" & NVL(!��鷽��) & "','" & NVL(!ִ��Ƶ��) & "'," & ZVal(NVL(!Ƶ�ʴ���, 0)) & "," & _
                        ZVal(NVL(!Ƶ�ʼ��, 0)) & ",'" & NVL(!�����λ) & "','" & NVL(!ҽ������) & "'," & _
                        NVL(!ִ������, 0) & "," & ZVal(NVL(!ִ�п���ID, 0)) & ",'" & NVL(!ʱ�䷽��) & "',Null,Null)"
                strTmp = strTmp & "," & !ID
                .MoveNext
            Next
            strTmp = Mid(strTmp, 2)
            '���ԭҽ����ɾ��,��ҽ��Id��·����ĿID�Ĺ�������·����Ŀ������ҽ���������
            For i = 1 To colAdvice.count
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_·��ҽ������_Update(2,Null,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,Null,Null,Null,null,null,'" & strTmp & "','" & colAdvice(i) & "','" & str��� & "')"
            Next
        Else
            '���������滻
            '��ѪK,����L,H����,����F,Ƥ�Ե�ֻ�滻������ĿId;
            '������ʱ
            '�滻��ҩ�䷽��ĳһҩ:ֻ�滻������ĿID,�շ�ϸĿID,������������Ƶ�ʲ��ı�
            For i = 1 To rsTmp.RecordCount
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                If str��� = "7" Then
                    arrSQL(UBound(arrSQL)) = "zl_·��ҽ������_Update(1," & rsTmp!����ID & "," & ZVal(NVL(!������ĿID)) & "," & ZVal(NVL(!�շ�ϸĿID)) & ",'" & !ҽ������ & "'," & _
                                ZVal(NVL(!��������)) & ")"
                ElseIf InStr(",1,2,", "," & rptList.SelectedRows(0).Record(COL_���㷽ʽ).Value & ",") > 0 Then '������ʱ
                    arrSQL(UBound(arrSQL)) = "zl_·��ҽ������_Update(1," & rsTmp!����ID & "," & ZVal(NVL(!������ĿID)) & ",NULL,NULL," & ZVal(NVL(!��������)) & "," & ZVal(NVL(!�ܸ�����)) & ")"
                Else
                    arrSQL(UBound(arrSQL)) = "zl_·��ҽ������_Update(1," & rsTmp!����ID & "," & ZVal(NVL(!������ĿID)) & ")"
                End If
                rsTmp.MoveNext
            Next
        End If
    End With

    '�ύ����
    gcnOracle.BeginTrans: blnTran = True
    For i = 0 To UBound(arrSQL)
        If CStr(arrSQL(i)) <> "" Then
            zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
        End If
    Next
    gcnOracle.CommitTrans: blnTran = False

    '�滻�ɹ�����Ҫˢ�½�������
    SaveData = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function AdviceMakeText(ByVal lngRow As Long, ByVal str��� As String, ByRef strTag As String, Optional ByVal lngBegin As Long) As String
'����:������飬�����ʾ������
    Dim str��λ As String
    Dim str���� As String
    Dim str��λLast As String
    Dim strReturn As String
    Dim str���� As String, str�걾  As String
    Dim i As Long
    
    str��λ = "": str���� = "": strTag = ""
    With rptList.Records
        If str��� = "D" Then
            For i = lngRow + 1 To .count - 1
                If Val(.Record(i).Item(COL_���ID).Value) = Val(.Record(lngRow).Item(COL_����ID).Value) Then
                    If .Record(i).Item(COL_�걾��λ).Value <> "" Then
                        If .Record(i).Item(COL_�걾��λ).Value <> str��λLast And str��λLast <> "" Then
                            str��λ = str��λ & "," & str��λLast & IIf(str���� <> "", "(" & Mid(str����, 2) & ")", "")
                            str���� = ""
                        End If
                        
                        If .Record(i).Item(COL_��鷽��).Value <> "" Then
                            str���� = str���� & "," & .Record(i).Item(COL_��鷽��).Value
                        End If
                        
                        str��λLast = .Record(i).Item(COL_�걾��λ).Value
                        
                        '��鷽��,�걾��λ
                        strTag = strTag & "," & .Record(i).Item(COL_�걾��λ).Value & "_" & .Record(i).Item(COL_��鷽��).Value
                        
                    End If
                Else
                    Exit For
                End If
            Next
            If str��λLast <> "" Then
                str��λ = str��λ & "," & str��λLast & IIf(str���� <> "", "(" & Mid(str����, 2) & ")", "")
            End If
            str��λ = Mid(str��λ, 2) '��������Ŀ�Ĳ�λ
            strReturn = .Record(lngRow).Item(COL_����).Value & ":" & str��λ
        ElseIf str��� = "C" Then
            str���� = "": str�걾 = ""
            
            For i = lngBegin To lngRow - 1    '��������δ��ɼ���ʽ
         
                str���� = .Record(i).Item(COL_����).Value & "," & str����
                str�걾 = .Record(i).Item(COL_�걾��λ).Value
                '��¼����걾��λ
                strTag = strTag & "," & .Record(i).Item(COL_����).Value & "_" & .Record(i).Item(COL_�걾��λ).Value
          
            Next
            str���� = Left(str����, Len(str����) - 1)
            strReturn = str���� & "(" & str�걾 & ")"
        End If
    End With
    strTag = Mid(strTag, 2)
    If strTag = "" Then
        strTag = strReturn
    End If
    AdviceMakeText = strReturn
End Function

Private Sub ClearPath()
'����:���·��������
    rptPath.Records.DeleteAll
    rptPath.Populate
    Call InitAdviceTable
End Sub

Private Sub InitAdviceTable()
'����:���ҽ������
    cmdEdit.Enabled = False
    cmdBatExe.Enabled = False
    Call InitAdviceRecordset
    Call ShowAdvice
End Sub

Private Sub RefreshData()
'����:ˢ������
    rptList.Records.DeleteAll
    Call ClearParaValue
    Call ClearPath
    Call LoadStopedItem
    Call SetEditable(-1, -1, -1, -1, -1, -1, -1)
End Sub

Private Sub FindRPTList(Optional ByVal blnNext As Boolean)
'������blnNext=�Ƿ������һ��
    Static blnReStart As Boolean
    Dim blnHave As Boolean, i As Long
    Call zlControl.TxtSelAll(txtFind)

    '��ʼ������
    If rptList.SelectedRows.count > 0 Then blnHave = True
    If Not blnNext Or blnReStart Or Not blnHave Then
        i = 0    'ReportControl����������0��ʼ
    Else
        i = rptList.SelectedRows(0).Index + 1
    End If

    
    For i = i To rptList.Rows.count - 1
        With rptList.Rows(i)
            If Not .GroupRow Then
                If IsNumeric(txtFind.Text) Then
                    '1X.����ȫ������ʱֻƥ�����
                    If .Record(COL_����).Value Like "*" & UCase(Trim(txtFind.Text)) & "*" Then
                        Exit For
                    End If
                ElseIf zlCommFun.IsCharAlpha(txtFind.Text) Then
                    'X1.����ȫ����ĸʱֻƥ�����
                    If .Record(COL_����).Value Like "*" & UCase(Trim(txtFind.Text)) & "*" Then
                        Exit For
                    End If
                ElseIf zlCommFun.IsCharChinese(txtFind.Text) Then
                    '��������,��ֻƥ������
                    If .Record(COL_����).Value Like "*" & UCase(Trim(txtFind.Text)) & "*" Then
                        Exit For
                    End If
                End If
            End If
        End With
    Next
    
    
    If i <= rptList.Rows.count - 1 Then
        blnReStart = False
        '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
        Set rptList.FocusedRow = rptList.Rows(i)
        If rptList.Visible Then rptList.SetFocus
    Else
        blnReStart = True
        MsgBox IIf(blnNext, "������", "") & "�Ҳ�������������������Ŀ��", vbInformation, gstrSysName
    End If
End Sub

Private Function CompareStr(ByVal str1 As String, ByVal str2 As String, Optional ByVal strDelimiter As String = ",") As Boolean
'����:�Ƚ������Զ��ŷָ����ַ����Ƿ���ȫ���,�����ַ����е�˳��
'����:
'     str1-�ַ���1���ֽ��֮����ַ����ܳ����ظ��ģ�
'     str2-�ַ���2
'     strDelimiter-�ָ���
'����ֵ:True-�൱,false-���뵱
'˵��: str1="1,2,3";str2="1,3,2" ,����ֵ -true
    Dim arrOne As Variant
    Dim arrTwo As Variant
    Dim i As Long
    
    arrOne = Split(str1, strDelimiter)
    arrTwo = Split(str2, strDelimiter)
    
    str2 = strDelimiter & str2 & strDelimiter
    If UBound(arrOne) <> UBound(arrTwo) Then Exit Function
    
    For i = LBound(arrOne) To UBound(arrOne)
        If InStr(str2, strDelimiter & arrOne(i) & strDelimiter) = 0 Then
            Exit Function
        End If
    Next
    CompareStr = True
End Function

Private Sub ClearParaValue()
'����:��չ��˲���ֵ
    
    txt����.Text = ""
    txt����.Text = ""
    txt�÷�.Text = ""
    txt�÷�.Tag = ""
    txtƵ��.Text = ""
    lbl������λ.Caption = ""
    lbl������λ.Caption = ""
End Sub
