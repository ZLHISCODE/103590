VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLabItemBase 
   BorderStyle     =   0  'None
   Caption         =   "��Ŀ������Ϣ"
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   Enabled         =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CheckBox chk��ο� 
      Caption         =   "��ο���Ŀ"
      Height          =   210
      Left            =   6735
      TabIndex        =   56
      Top             =   525
      Width           =   1200
   End
   Begin VB.TextBox txt���鷽�� 
      Height          =   300
      Left            =   1215
      MaxLength       =   40
      ScrollBars      =   2  'Vertical
      TabIndex        =   50
      Top             =   3825
      Width           =   6630
   End
   Begin VB.TextBox txt����ƴ�� 
      Height          =   300
      Left            =   1215
      MaxLength       =   12
      TabIndex        =   54
      Top             =   1221
      Width           =   1335
   End
   Begin VB.TextBox txt��Ŀ���� 
      Height          =   300
      Left            =   1215
      MaxLength       =   60
      TabIndex        =   53
      Top             =   849
      Width           =   3975
   End
   Begin VB.TextBox txt��Ŀ���� 
      Height          =   300
      Left            =   1215
      MaxLength       =   13
      TabIndex        =   52
      Top             =   477
      Width           =   1335
   End
   Begin VB.TextBox txtӢ����д 
      Height          =   300
      Left            =   1215
      MaxLength       =   10
      TabIndex        =   51
      Top             =   1965
      Width           =   1335
   End
   Begin VB.TextBox txtȡֵ���� 
      Height          =   300
      Left            =   1215
      MaxLength       =   200
      ScrollBars      =   2  'Vertical
      TabIndex        =   49
      ToolTipText     =   "(�붨���Ͷ�����Ŀ������ȡֵ���У������ѡȡֵ�ǲ��á�;���ָ�)"
      Top             =   3453
      Width           =   6630
   End
   Begin VB.TextBox txt���㹫ʽ 
      Height          =   300
      Left            =   1215
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   48
      Top             =   3081
      Width           =   6330
   End
   Begin VB.TextBox txt���Ʒ��� 
      Height          =   300
      Left            =   1215
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   47
      Top             =   105
      Width           =   3660
   End
   Begin VB.ComboBox cbo�걾���� 
      Height          =   300
      ItemData        =   "frmLabItemBase.frx":0000
      Left            =   1215
      List            =   "frmLabItemBase.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   46
      Top             =   2337
      Width           =   1365
   End
   Begin VB.TextBox txtAlias 
      Height          =   300
      Left            =   1215
      MaxLength       =   60
      TabIndex        =   45
      Top             =   1593
      Width           =   3975
   End
   Begin VB.ComboBox cbo�Թ� 
      Height          =   300
      Left            =   1215
      Style           =   2  'Dropdown List
      TabIndex        =   44
      Top             =   2709
      Width           =   1365
   End
   Begin VB.TextBox txt������� 
      Height          =   300
      Left            =   6315
      MaxLength       =   14
      TabIndex        =   41
      Top             =   2337
      Width           =   1410
   End
   Begin VB.Frame fraø�� 
      Caption         =   "ø����Ŀ��ʽ"
      Height          =   1365
      Left            =   90
      TabIndex        =   33
      Top             =   4215
      Width           =   7770
      Begin VB.TextBox txtCutOff��ʽ 
         Height          =   300
         Left            =   1110
         MaxLength       =   200
         ScrollBars      =   2  'Vertical
         TabIndex        =   36
         Top             =   975
         Width           =   6555
      End
      Begin VB.TextBox txt�����Թ�ʽ 
         Height          =   300
         Left            =   1110
         MaxLength       =   200
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Top             =   600
         Width           =   6555
      End
      Begin VB.TextBox txt���Թ�ʽ 
         Height          =   300
         Left            =   1110
         MaxLength       =   200
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   240
         Width           =   6555
      End
      Begin VB.Label lblCutOff��ʽ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CutOff��ʽ"
         Height          =   180
         Left            =   165
         TabIndex        =   39
         Top             =   1020
         Width           =   900
      End
      Begin VB.Label lbl�����Թ�ʽ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����Թ�ʽ"
         Height          =   180
         Left            =   165
         TabIndex        =   38
         Top             =   690
         Width           =   900
      End
      Begin VB.Label lbl���Թ�ʽ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���Թ�ʽ"
         Height          =   180
         Left            =   165
         TabIndex        =   37
         Top             =   315
         Width           =   900
      End
   End
   Begin VB.CheckBox chkPrivacy 
      Caption         =   "��˽��Ŀ"
      Height          =   210
      Left            =   6735
      TabIndex        =   32
      Top             =   150
      Width           =   1065
   End
   Begin VB.CommandButton cmdFormula 
      Caption         =   "��"
      Height          =   285
      Left            =   7530
      TabIndex        =   25
      Top             =   3075
      Width           =   300
   End
   Begin VB.OptionButton OptApplyType 
      Caption         =   "Ӧ���ڱ����Թܱ���"
      Height          =   285
      Left            =   5835
      TabIndex        =   31
      Top             =   2762
      Width           =   1995
   End
   Begin VB.OptionButton OptApplyOnly 
      Caption         =   "Ӧ���ڱ����Թܱ���"
      Height          =   285
      Left            =   3780
      TabIndex        =   30
      Top             =   2762
      Value           =   -1  'True
      Width           =   1995
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   2835
      Left            =   -3540
      TabIndex        =   28
      TabStop         =   0   'False
      Tag             =   "1000"
      Top             =   240
      Visible         =   0   'False
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   5001
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgList"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.TextBox txtĬ�Ͻ�� 
      Height          =   300
      Left            =   6330
      MaxLength       =   40
      TabIndex        =   23
      Top             =   1965
      Width           =   1380
   End
   Begin VB.ComboBox cbo�����Χ 
      Height          =   300
      ItemData        =   "frmLabItemBase.frx":0004
      Left            =   6330
      List            =   "frmLabItemBase.frx":0006
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   1593
      Width           =   1380
   End
   Begin VB.CommandButton cmd���Ʒ��� 
      Caption         =   "&P"
      Height          =   285
      Left            =   4875
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   113
      Width           =   285
   End
   Begin VB.ComboBox cbo��Ŀ��� 
      Height          =   300
      ItemData        =   "frmLabItemBase.frx":0008
      Left            =   6330
      List            =   "frmLabItemBase.frx":000A
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   849
      Width           =   1380
   End
   Begin VB.ComboBox cbo������� 
      Height          =   300
      ItemData        =   "frmLabItemBase.frx":000C
      Left            =   6330
      List            =   "frmLabItemBase.frx":000E
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   1221
      Width           =   1380
   End
   Begin VB.ComboBox cbo�������� 
      Height          =   300
      Left            =   3975
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   480
      Width           =   1200
   End
   Begin VB.CheckBox chk�����Ŀ 
      Caption         =   "��ϼ�����Ŀ"
      Height          =   210
      Left            =   5310
      TabIndex        =   16
      Top             =   525
      Width           =   1440
   End
   Begin VB.ComboBox cbo�����Ա� 
      Height          =   300
      Left            =   3975
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2337
      Width           =   1200
   End
   Begin VB.TextBox txt������� 
      Height          =   300
      Left            =   3255
      MaxLength       =   12
      TabIndex        =   7
      Top             =   1221
      Width           =   1320
   End
   Begin VB.CheckBox chk����Ӧ�� 
      Caption         =   "������Ӧ��"
      Height          =   210
      Left            =   5295
      TabIndex        =   15
      Top             =   150
      Value           =   1  'Checked
      Width           =   1530
   End
   Begin VB.TextBox txt���㵥λ 
      Height          =   300
      Left            =   3975
      MaxLength       =   10
      TabIndex        =   11
      Top             =   1965
      Width           =   1200
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   7680
      Top             =   1650
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabItemBase.frx":0010
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabItemBase.frx":006E
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabItemBase.frx":00CC
            Key             =   "Dept"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ʵ�鷽��(&Y)"
      Height          =   180
      Left            =   165
      TabIndex        =   55
      Top             =   3870
      Width           =   990
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3375
      TabIndex        =   43
      Top             =   2709
      Width           =   300
   End
   Begin VB.Label Label2 
      Caption         =   "�Թ���ɫ"
      Height          =   225
      Left            =   2595
      TabIndex        =   42
      Top             =   2762
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������(&L)"
      Height          =   180
      Left            =   5295
      TabIndex        =   40
      Top             =   2385
      Width           =   990
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��    ��(&B)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   8
      Top             =   1649
      Width           =   990
   End
   Begin VB.Label lbl�Թܱ��� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�Թܱ���(&C)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   29
      Top             =   2762
      Width           =   990
   End
   Begin VB.Label lbl�걾���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�걾����(&M)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   12
      Top             =   2391
      Width           =   990
   End
   Begin VB.Label lblĬ�Ͻ�� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ĭ�Ͻ��(&R)"
      Height          =   180
      Left            =   5310
      TabIndex        =   22
      Top             =   2025
      Width           =   990
   End
   Begin VB.Label lbl�����Χ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����Χ(W)"
      Height          =   180
      Left            =   5310
      TabIndex        =   27
      Top             =   1650
      Width           =   990
   End
   Begin VB.Label lbl���Ʒ��� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���Ʒ���(&F)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   0
      Top             =   165
      Width           =   990
   End
   Begin VB.Label lbl��Ŀ��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ŀ����(&K)"
      Height          =   180
      Left            =   5310
      TabIndex        =   17
      Top             =   900
      Width           =   990
   End
   Begin VB.Label lbl������� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������(T)"
      Height          =   180
      Left            =   5310
      TabIndex        =   19
      Top             =   1275
      Width           =   990
   End
   Begin VB.Label lbl���㹫ʽ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���㹫ʽ(F)"
      Height          =   180
      Left            =   165
      TabIndex        =   24
      Top             =   3133
      Width           =   990
   End
   Begin VB.Label lblȡֵ���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ȡֵ����(&P)"
      Height          =   180
      Left            =   165
      TabIndex        =   26
      Top             =   3510
      Width           =   990
   End
   Begin VB.Label lblӢ����д 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ӣ����д(&E)"
      Height          =   180
      Left            =   165
      TabIndex        =   9
      Top             =   2020
      Width           =   990
   End
   Begin VB.Label lbl�������� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������(&T)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2985
      TabIndex        =   3
      Top             =   540
      Width           =   990
   End
   Begin VB.Label lbl��Ŀ���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��Ŀ����(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   2
      Top             =   536
      Width           =   990
   End
   Begin VB.Label lbl��Ŀ���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   5
      Top             =   907
      Width           =   990
   End
   Begin VB.Label lbl�����Ա� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�����Ա�(&X)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2940
      TabIndex        =   13
      Top             =   2391
      Width           =   990
   End
   Begin VB.Label lbl���Ƽ��� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���Ƽ���(&S)               (ƴ��)                 (���)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   6
      Top             =   1278
      Width           =   4950
   End
   Begin VB.Label lbl���㵥λ 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���㵥λ(&U)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2940
      TabIndex        =   10
      Top             =   2020
      Width           =   990
   End
End
Attribute VB_Name = "frmLabItemBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngItemID As Long          '��ǰ��ʾ����Ŀid

Dim objNode As Node
Dim strTemp As String, aryTemp() As String
Dim lngCount As Long

'--------------------------------------------
'����Ϊ���幫������
'--------------------------------------------
Public Function zlRefresh(lngItemId As Long) As Boolean
    '���ܣ�������Ŀidˢ�µ�ǰ��ʾ����
    Dim rsTemp As New ADODB.Recordset, rsGS As New ADODB.Recordset
    Dim strTmp As String, strItem As String, lngLength As Long
    mlngItemID = lngItemId
    
    '�����ǰ��Ŀ����ʾ
    Me.cbo��������.ListIndex = -1
    Me.txt��Ŀ����.Text = "": Me.txt��Ŀ����.Text = ""
    Me.txt����ƴ��.Text = "": Me.txt�������.Text = ""
    Me.txtӢ����д.Text = "": Me.txt���㵥λ.Text = ""
    Me.cbo�걾����.ListIndex = -1: Me.cbo�����Ա�.ListIndex = -1
    Me.cbo��Ŀ���.ListIndex = -1: Me.cbo�������.ListIndex = -1: Me.txtĬ�Ͻ��.Text = ""
    Me.txt���㹫ʽ.Text = "": Me.txtȡֵ����.Text = "": Me.txtAlias.Text = ""
    Me.OptApplyOnly.Value = True
    Me.chkPrivacy.Value = 0
    Me.txt���Թ�ʽ.Text = ""
    Me.txt�����Թ�ʽ.Text = ""
    Me.txtCutOff��ʽ.Text = ""
    Me.txt���鷽��.Text = ""
    Me.chk��ο�.Value = 0
    
    If lngItemId = 0 Then zlRefresh = True: Exit Function
    
    '��ȡָ����Ŀ����Ϣ
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select ����id, ��������, ����, ����, ���㵥λ, �걾��λ, �����Ա�, ����Ӧ��, �����Ŀ,�Թܱ��� From ������ĿĿ¼ Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
    With rsTemp
        If .RecordCount > 0 Then
            If Val("" & !����id) > 0 Then
                Me.tvwClass.Nodes("_" & !����id).Selected = True
                Me.txt���Ʒ���.Text = Me.tvwClass.SelectedItem.Text
                Me.txt���Ʒ���.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
            End If
            For lngCount = 0 To Me.cbo��������.ListCount - 1
                If Mid(Me.cbo��������.List(lngCount), 4) = "" & !�������� Then Me.cbo��������.ListIndex = lngCount: Exit For
            Next
            Me.txt��Ŀ����.Text = "" & !����
            Me.txt��Ŀ����.Text = "" & !����
            Me.txt���㵥λ.Text = "" & !���㵥λ

            If "" & !�Թܱ��� = "" Or "" & !�Թܱ��� = "NULL" Then
                Me.cbo�Թ�.ListIndex = 0
            Else
                For lngCount = 0 To Me.cbo�Թ�.ListCount - 1
                    If Split(Me.cbo�Թ�.List(lngCount), "-")(0) = "" & !�Թܱ��� Then Me.cbo�Թ�.ListIndex = lngCount: Exit For
                Next
            End If
            
            For lngCount = 0 To Me.cbo�걾����.ListCount - 1
                If Mid(Me.cbo�걾����.List(lngCount), 4) = "" & !�걾��λ Then Me.cbo�걾����.ListIndex = lngCount: Exit For
            Next
            For lngCount = 0 To Me.cbo�����Ա�.ListCount - 1
                If Left(Me.cbo�����Ա�.List(lngCount), 1) = "" & !�����Ա� Then Me.cbo�����Ա�.ListIndex = lngCount: Exit For
            Next
            Me.chk����Ӧ��.Value = IIf(Val("" & !����Ӧ��) = 1, 1, 0)
            Me.chk�����Ŀ.Value = IIf(Val("" & !�����Ŀ) = 1, 1, 0)
            If Me.chk�����Ŀ.Value = 1 Then
                Me.chk����Ӧ��.Enabled = False
            Else
                Me.chk����Ӧ��.Enabled = True
            End If
        End If
    End With
        
    gstrSql = "Select ����,����,����,���� From ������Ŀ���� Where ������ĿID=[1] And ����=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
    With rsTemp
        Do While Not .EOF
            If !���� = 1 And !���� = 1 Then Me.txt����ƴ��.Text = !����
            If !���� = 1 And !���� = 2 Then Me.txt�������.Text = !����
            .MoveNext
        Loop
    End With
            
    gstrSql = "select ����,����,����,���� from ������Ŀ���� where ������ĿID=" & lngItemId
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
    Do While Not rsTemp.EOF
        If rsTemp!���� = 9 And rsTemp!���� = 1 Then Me.txtAlias.Text = rsTemp!����
        If rsTemp!���� = 9 And rsTemp!���� = 2 Then Me.txtAlias.Text = rsTemp!����
        rsTemp.MoveNext
    Loop
            
    '��ѯ����������Ŀ��Ӧ�ļ���ָ��
    If Me.chk�����Ŀ.Value = 0 Then
        gstrSql = "Select A.��д, A.��Ŀ���, A.�������, A.�����Χ, A.Ĭ��ֵ, A.���㹫ʽ, A.ȡֵ����,A.��˽��Ŀ, " & vbNewLine & _
                  " A.���Թ�ʽ, A.�����Թ�ʽ, A.CutOff��ʽ,A.�������, A.���鷽��, A.��ο� " & vbNewLine & _
                "From ������Ŀ A, ���鱨����Ŀ C" & vbNewLine & _
                "Where A.������Ŀid = C.������Ŀid And C.������Ŀid = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
        With rsTemp
            If .RecordCount > 0 Then
                Me.txtӢ����д.Text = "" & !��д
                For lngCount = 0 To Me.cbo��Ŀ���.ListCount - 1
                    If Left(Me.cbo��Ŀ���.List(lngCount), 1) = "" & !��Ŀ��� Then Me.cbo��Ŀ���.ListIndex = lngCount: Exit For
                Next
                For lngCount = 0 To Me.cbo�������.ListCount - 1
                    If Left(Me.cbo�������.List(lngCount), 1) = "" & !������� Then Me.cbo�������.ListIndex = lngCount: Exit For
                Next
                Me.txtĬ�Ͻ��.Text = "" & !Ĭ��ֵ
                Me.txt���㹫ʽ.Text = "" & !���㹫ʽ
                Me.chkPrivacy.Value = Nvl(!��˽��Ŀ, 0)
                Me.txt�������.Text = "" & !�������
                
                If Me.txt���㹫ʽ.Text <> "" Then
                    Do While Me.txt���㹫ʽ.Text Like "*[[]*[]]*"
                        strTmp = strTmp & Mid(Me.txt���㹫ʽ.Text, 1, InStr(Me.txt���㹫ʽ.Text, "[") - 1)
                        lngLength = InStr(Me.txt���㹫ʽ.Text, "]") - InStr(Me.txt���㹫ʽ.Text, "[") - 1
                        strItem = Mid(Me.txt���㹫ʽ.Text, InStr(Me.txt���㹫ʽ.Text, "[") + 1, lngLength)
                        gstrSql = "Select ������ĿID,��д From ������Ŀ Where ������ĿID=[1] "
                        Set rsGS = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(strItem))
                        Do Until rsGS.EOF
                            If Trim("" & rsGS.Fields("��д")) <> "" Then
                                strTmp = strTmp & "[" & Trim("" & rsGS.Fields("��д")) & "]"
                            Else
                                strTmp = strTmp & "[" & Val(strItem) & "]"
                            End If
                            rsGS.MoveNext
                        Loop
                        Me.txt���㹫ʽ.Text = Mid(Me.txt���㹫ʽ.Text, InStr(Me.txt���㹫ʽ.Text, "]") + 1)
                    Loop
                    strTmp = strTmp & Mid(Me.txt���㹫ʽ.Text, InStr(Me.txt���㹫ʽ.Text, "]") + 1)
                    Me.txt���㹫ʽ.Text = strTmp
                End If
                
                Me.txtȡֵ����.Text = "" & !ȡֵ����
                
                Me.txt���Թ�ʽ.Text = "" & !���Թ�ʽ
                Me.txt�����Թ�ʽ.Text = "" & !�����Թ�ʽ
                Me.txtCutOff��ʽ.Text = "" & !CutOff��ʽ
                Me.txt���鷽��.Text = "" & !���鷽��
                
                Me.chk��ο�.Value = Val("" & !��ο�)
            End If
        End With
    End If
    
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart(blnAdd As Boolean, lngItemId As Long) As Boolean
    '���ܣ���ʼ��Ŀ�༭
    '������ blnAdd-�Ƿ����ӣ�����Ϊ�޸�
    '       lngItemId-���ӵĲ�����Ŀ������ָ���༭����Ŀ
    Dim rsTemp As New ADODB.Recordset
    If Me.tvwClass.Nodes.Count = 0 Then
        MsgBox "�������ֵ��г�ʼ�������Ʒ���Ŀ¼����", vbInformation, gstrSysName
        zlEditStart = False: Exit Function
    End If
    If Me.cbo��������.ListCount = 0 Then
        MsgBox "�������ֵ��г�ʼ�����������͡���", vbInformation, gstrSysName
        zlEditStart = False: Exit Function
    End If
    
    If blnAdd Then
        Err = 0: On Error GoTo ErrHand
        If Val(zlDatabase.GetPara(61, glngSys)) = 0 Then '������Ŀ�������ģʽ
            gstrSql = "Select Nvl(Max(����),'0000000') As ���� From ������ĿĿ¼ Where ��� >= 'A'"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "zlEditStart")
            
            Me.txt��Ŀ����.Text = IncStr(rsTemp!����)
        Else
            strTemp = Mid(Me.txt���Ʒ���.Text, 2, InStr(1, Me.txt���Ʒ���.Text, "]") - 2)
            gstrSql = "Select Nvl(Max(����),'0000000') As ����" & _
                    " From ������ĿĿ¼" & _
                    " Where ��� >= 'A' And ���� Like '" & strTemp & "%'"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "zlEditStart")
            
            Err = 0: On Error Resume Next
            If rsTemp!���� = "0000000" Then
                Me.txt��Ŀ����.Text = strTemp & IncStr(rsTemp!����)
            Else
                Me.txt��Ŀ����.Text = strTemp & IncStr(Right(rsTemp!����, Len(rsTemp!����) - Len(strTemp)))
            End If
        End If
        
        '���������Ĭ��ֵ
        Me.txt��Ŀ����.Text = "": Me.cbo�Թ�.ListIndex = 0
        Me.txt����ƴ��.Text = "": Me.txt�������.Text = ""
        Me.txtӢ����д.Text = "": Me.txtĬ�Ͻ��.Text = ""
        Me.txt���㹫ʽ.Text = "": Me.txtȡֵ����.Text = ""
        Me.txtAlias.Text = ""
        If lngItemId = 0 Then
            '˵����ǰû�пɼ̳���Ϣ����Ҫ���ò���Ĭ��ֵ
            Me.cbo��������.ListIndex = 0
            Me.cbo�걾����.ListIndex = 0: Me.cbo�����Ա�.ListIndex = 0
            Me.cbo��Ŀ���.ListIndex = 0: Me.cbo�������.ListIndex = 0
        End If
    End If

    mlngItemID = lngItemId
    Me.Enabled = True: Me.Tag = IIf(blnAdd, "����", "�޸�")
    Me.BackColor = RGB(250, 250, 250): Me.chk����Ӧ��.BackColor = Me.BackColor: Me.chk�����Ŀ.BackColor = Me.BackColor
    Me.chkPrivacy.BackColor = Me.BackColor: Me.chk��ο�.BackColor = Me.BackColor
    Me.OptApplyOnly.BackColor = Me.BackColor: Me.OptApplyType.BackColor = Me.BackColor
    Me.txt���Ʒ���.SetFocus
    zlEditStart = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditStart = False: Exit Function
End Function

Public Function IncStr(ByVal strVal As String, Optional intUpDown As Integer, Optional ByRef strErr As String) As String
'���ܣ���һ���ַ����Զ���1��
'˵����ÿһλ��λʱ,���������,��ʮ���ƴ���,����26���ƴ���
'������strVal=Ҫ��1���ַ���
'      intUpDown = 0 ��1 =1 ��1
    Dim strValuse As String
    Dim intAdd As Integer
    Dim intUp As Integer
    Dim strValue As String
    Dim strValueOne As String
    Dim strHead As String
    
    Dim i  As Integer
    
    On Error GoTo errH
    
    strVal = UCase(strVal)

    For i = Len(strVal) To 1 Step -1
        strValueOne = Mid(strVal, i, 1)
        If Asc(strValueOne) >= Asc("0") And Asc(strValueOne) <= Asc("9") Then
        Else
            '��������
            strHead = Mid$(strVal, 1, i)
            strVal = Mid$(strVal, i + 1)
            Exit For
        End If
    Next
    
    strVal = UCase(strVal)
    
    If intUpDown = 0 Then
        '��1
        For i = Len(strVal) To 1 Step -1
            If i = Len(strVal) Then
                intAdd = 1
            Else
                intAdd = 0
            End If
            strValueOne = Mid(strVal, i, 1)
    
            If IsNumeric(strValueOne) Then
                If Val(strValueOne) + intAdd + intUp < 10 Then
                    strValue = Chr(Asc(strValueOne) + intAdd + intUp) & strValue
                    intUp = 0
                Else
                    strValue = "0" & strValue
                    intUp = 1
                End If
            Else
                If Asc(strValueOne) + intAdd + intUp <= Asc("Z") Then
                    strValue = Chr(Asc(strValueOne) + intAdd + intUp) & strValue
                    intUp = 0
                Else
                    strValue = "A" & strValue
                    intUp = 1
                End If
            End If
        Next
        
        If intUp = 1 Then
            If IsNumeric(strValueOne) Then
                strValue = "1" & strValue
            Else
                strValue = "A" & strValue
            End If
        End If
        IncStr = IIf(strHead <> "", strHead & strValue, strValue)
    Else
        For i = Len(strVal) To 1 Step -1
            If i = Len(strVal) Then
                intAdd = -1
            Else
                intAdd = 0
            End If
            strValueOne = Mid(strVal, i, 1)
    
            If IsNumeric(strValueOne) Then
                If Val(strValueOne) + intAdd + intUp >= 0 Then
                    strValue = Chr(Asc(strValueOne) + intAdd + intUp) & strValue
                    intUp = 0
                Else
                    strValue = "9" & strValue
                    intUp = -1
                End If
            Else
                If Asc(strValueOne) + intAdd + intUp >= Asc("A") Then
                    strValue = Chr(Asc(strValueOne) + intAdd + intUp) & strValue
                    intUp = 0
                Else
                    If intAdd = 0 Then
                        strValue = "Z" & strValue
                    End If
                    intUp = -1
                End If
            End If
        Next
        
        If intUp = 1 Then
            strValue = -1
        End If
        
        If Mid(strValue, 1, 1) = "0" Or Mid(strValue, 1, 1) = "A" Then
            strValue = Mid(strValue, 2)
            If strValue = "" Then strValue = 1
        End If
        IncStr = IIf(strHead <> "", strHead & strValue, strValue)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub zlEditCancel()
    '���ܣ��������ڽ��еı༭
    Me.Enabled = False: Me.Tag = ""
    Me.BackColor = Me.cmd���Ʒ���.BackColor: Me.chk����Ӧ��.BackColor = Me.BackColor: Me.chk�����Ŀ.BackColor = Me.BackColor:
    Me.chkPrivacy.BackColor = Me.BackColor: Me.chk��ο�.BackColor = Me.BackColor
    Me.OptApplyOnly.BackColor = Me.cmd���Ʒ���.BackColor: Me.OptApplyType.BackColor = Me.cmd���Ʒ���.BackColor
    Call Me.zlRefresh(mlngItemID)
End Sub

Public Function zlEditSave() As Long
    '���ܣ��������ڽ��еı༭,���������ڱ༭��Ŀid,����ʧ�ܷ���0
    Dim lngNewId As Long
    Dim rsGS As New ADODB.Recordset, strItem As String, strTmp As String, lngLength As Long
    Dim str�Թ� As String
    
    'һ�����Լ��
    Err = 0: On Error GoTo ErrHand
    If Trim(Me.txt��Ŀ����.Text) = "" Then
        MsgBox "��������Ŀ���룡", vbInformation, gstrSysName
        Me.txt��Ŀ����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt��Ŀ����.Text), vbFromUnicode)) > Me.txt��Ŀ����.MaxLength Then
        MsgBox "��Ŀ����ĳ��ȳ��������" & Me.txt��Ŀ����.MaxLength & " ���ַ�����", vbInformation, gstrSysName
        Me.txt��Ŀ����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Trim(Me.txt��Ŀ����.Text) = "" Then
        MsgBox "��������Ŀ���ƣ�", vbInformation, gstrSysName
        Me.txt��Ŀ����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt��Ŀ����.Text), vbFromUnicode)) > Me.txt��Ŀ����.MaxLength Then
        MsgBox "��Ŀ���Ƴ��������" & Me.txt��Ŀ����.MaxLength & "���ַ���ȳ����֣���", vbInformation, gstrSysName
        Me.txt��Ŀ����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt���㵥λ.Text), vbFromUnicode)) > Me.txt���㵥λ.MaxLength Then
        MsgBox "���㵥λ���������" & Me.txt���㵥λ.MaxLength & "���ַ���ȳ����֣���", vbInformation, gstrSysName
        Me.txt���㵥λ.SetFocus: zlEditSave = 0: Exit Function
    End If
    '10804 ��������ʱ�������������Ƿ�ɾ��
    If Not zlExistItem("���Ƽ�������", "����", Mid(Me.cbo��������.Text, InStr(1, Me.cbo��������, "-") + 1), "�������ͣ�" & Mid(Me.cbo��������.Text, InStr(1, Me.cbo��������, "-") + 1)) Then
        Me.cbo��������.SetFocus: zlEditSave = 0: Exit Function
    End If
'    If Trim(Me.txtӢ����д.Text) = "" And Me.cbo��Ŀ���.ListIndex = 0 Then
'        MsgBox "������Ŀ������Ӣ����д��", vbInformation, gstrSysName
'        Me.txtӢ����д.SetFocus: zlEditSave = 0: Exit Function
'    End If
    If Trim(Me.txt����ƴ��) = "" Then
        Me.txt����ƴ��.Text = zlStr.GetCodeByORCL(Me.txt��Ŀ����.Text, False, Me.txt����ƴ��.MaxLength)
    End If
    
    If Trim(Me.txt�������) = "" Then
        Me.txt�������.Text = zlStr.GetCodeByORCL(Me.txt��Ŀ����.Text, True, Me.txt�������.MaxLength)
    End If
    If LenB(StrConv(Trim(Me.txtӢ����д.Text), vbFromUnicode)) > Me.txtӢ����д.MaxLength Then
        MsgBox "Ӣ����д���������" & Me.txtӢ����д.MaxLength & "���ַ���ȳ����֣���", vbInformation, gstrSysName
        Me.txtӢ����д.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    '������� Ϊ �Ƕ����ģ�����Ƿ�����������Ŀ����
    If Mid(Me.cbo�������.Text, 1, 1) <> "1" And Me.Tag <> "����" Then
        gstrSql = "Select ������Ŀid, ��д, B.������, B.����" & vbNewLine & _
                "From ����������Ŀ B, ������Ŀ A" & vbNewLine & _
                "Where A.������Ŀid = B.ID And" & vbNewLine & _
                "      ���㹫ʽ Like (Select '%' || Chr(91) || A.������Ŀid || Chr(93) || '%' From ���鱨����Ŀ A ,������ĿĿ¼ B Where A.������Ŀid=B.ID and B.�����Ŀ=0 and A.������Ŀid = [1])"
        Set rsGS = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngItemID)
        strTmp = "����Ŀ��������Ŀ���ã����ܸ��Ľ�����ͣ�"
        Do Until rsGS.EOF
            strItem = strItem & "(" & rsGS.Fields("����") & ")" & rsGS.Fields("������") & vbNewLine
            rsGS.MoveNext
        Loop
        If strItem <> "" Then
            MsgBox strTmp & vbNewLine & strItem, vbInformation, Me.Caption
            Exit Function
        End If
    End If
    
    strItem = "": strTmp = ""
    If Me.txt���㹫ʽ.Text <> "" And Me.cbo��Ŀ���.ListIndex = 2 Then
         
        Do While Me.txt���㹫ʽ.Text Like "*[[]*[]]*"
            strTmp = strTmp & Mid(Me.txt���㹫ʽ.Text, 1, InStr(Me.txt���㹫ʽ.Text, "[") - 1)
            lngLength = InStr(Me.txt���㹫ʽ.Text, "]") - InStr(Me.txt���㹫ʽ.Text, "[") - 1
            strItem = Mid(Me.txt���㹫ʽ.Text, InStr(Me.txt���㹫ʽ.Text, "[") + 1, lngLength)
            gstrSql = "Select ������ĿID,��д From ������Ŀ Where (������ĿID=[1] or ��д=[2]) "
            Set rsGS = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(strItem), strItem)
            Do Until rsGS.EOF
                strTmp = strTmp & "[" & Val("" & rsGS.Fields("������ĿID")) & "]"
                rsGS.MoveNext
            Loop
            Me.txt���㹫ʽ.Text = Mid(Me.txt���㹫ʽ.Text, InStr(Me.txt���㹫ʽ.Text, "]") + 1)
        Loop
        strTmp = strTmp & Mid(Me.txt���㹫ʽ.Text, InStr(Me.txt���㹫ʽ.Text, "]") + 1)
        Me.txt���㹫ʽ.Text = strTmp
    Else
        Me.txt���㹫ʽ.Text = ""
    End If
    
    If LenB(StrConv(Trim(Me.txt���㹫ʽ.Text), vbFromUnicode)) > Me.txt���㹫ʽ.MaxLength Then
        MsgBox "���㹫ʽ���������" & Me.txt���㹫ʽ.MaxLength & "���ַ���ȳ����֣���", vbInformation, gstrSysName
        Me.txt���㹫ʽ.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    If LenB(StrConv(Trim(Me.txtĬ�Ͻ��.Text), vbFromUnicode)) > Me.txtĬ�Ͻ��.MaxLength Then
        MsgBox "Ĭ�Ͻ�����������" & Me.txtĬ�Ͻ��.MaxLength & "���ַ���ȳ����֣���", vbInformation, gstrSysName
        Me.txtĬ�Ͻ��.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txtȡֵ����.Text), vbFromUnicode)) > Me.txtȡֵ����.MaxLength Then
        MsgBox "ȡֵ���г��������" & Me.txtȡֵ����.MaxLength & "���ַ���ȳ����֣���", vbInformation, gstrSysName
        Me.txtȡֵ����.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    If LenB(StrConv(Trim(Me.txt���Թ�ʽ.Text), vbFromUnicode)) > Me.txt���Թ�ʽ.MaxLength Then
        MsgBox "���Թ�ʽ���������" & Me.txt���Թ�ʽ.MaxLength & "���ַ���ȳ����֣���", vbInformation, gstrSysName
        Me.txt���Թ�ʽ.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt�����Թ�ʽ.Text), vbFromUnicode)) > Me.txt�����Թ�ʽ.MaxLength Then
        MsgBox "�����Թ�ʽ���������" & Me.txt�����Թ�ʽ.MaxLength & "���ַ���ȳ����֣���", vbInformation, gstrSysName
        Me.txt�����Թ�ʽ.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txtCutOff��ʽ.Text), vbFromUnicode)) > Me.txtCutOff��ʽ.MaxLength Then
        MsgBox "CutOff��ʽ���������" & Me.txtCutOff��ʽ.MaxLength & "���ַ���ȳ����֣���", vbInformation, gstrSysName
        Me.txtCutOff��ʽ.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Me.txt���Թ�ʽ.Text <> "" And txt���Թ�ʽ.Enabled Then
        strTmp = UCase(Me.txt���Թ�ʽ.Text)
        strTmp = Replace(strTmp, "OD", "100"): strTmp = Replace(strTmp, "NC", "1"): strTmp = Replace(strTmp, "BC", "1")
        strTmp = Replace(strTmp, "QC", "1"): strTmp = Replace(strTmp, "PC", "1")
        If Check_Expression(0, strTmp) = False Then
            MsgBox "���Թ�ʽ�������飡", vbInformation, Me.Caption
            Me.txt���Թ�ʽ.SetFocus: zlEditSave = 0: Exit Function
        End If
        Me.txt���Թ�ʽ.Text = UCase(Me.txt���Թ�ʽ.Text)
    End If
    If Me.txt�����Թ�ʽ.Text <> "" And txt�����Թ�ʽ.Enabled Then
        strTmp = UCase(Me.txt�����Թ�ʽ.Text)
        strTmp = Replace(strTmp, "OD", "100"): strTmp = Replace(strTmp, "NC", "1"): strTmp = Replace(strTmp, "BC", "1")
        strTmp = Replace(strTmp, "QC", "1"): strTmp = Replace(strTmp, "PC", "1")
        If Check_Expression(0, strTmp) = False Then
            MsgBox "�����Թ�ʽ�������飡", vbInformation, Me.Caption
            Me.txt�����Թ�ʽ.SetFocus: zlEditSave = 0: Exit Function
        End If
        Me.txt�����Թ�ʽ.Text = UCase(Me.txt�����Թ�ʽ.Text)
    End If
    If Me.txtCutOff��ʽ.Text <> "" And txtCutOff��ʽ.Enabled Then
        strTmp = UCase(Me.txtCutOff��ʽ.Text)
        strTmp = Replace(strTmp, "OD", "100"): strTmp = Replace(strTmp, "NC", "1"): strTmp = Replace(strTmp, "BC", "1")
        strTmp = Replace(strTmp, "QC", "1"): strTmp = Replace(strTmp, "PC", "1")
        If Check_Expression(1, strTmp) = False Then
            MsgBox "CutOff��ʽ�������飡", vbInformation, Me.Caption
            Me.txtCutOff��ʽ.SetFocus: zlEditSave = 0: Exit Function
        End If
        Me.txtCutOff��ʽ.Text = UCase(Me.txtCutOff��ʽ.Text)
    End If
    '���ݱ��������֯
    If Me.Tag = "����" Then
        lngNewId = zlDatabase.GetNextId("������ĿĿ¼")
        If zlClinicCodeRepeat(Trim(Me.txt��Ŀ����.Text)) = True Then zlEditSave = 0: Exit Function
    Else
        If zlClinicCodeRepeat(Trim(Me.txt��Ŀ����.Text), mlngItemID) = True Then zlEditSave = 0: Exit Function
        '�����Ŀ�Ƿ���
        If zlExistItem("������ĿĿ¼", "ID", mlngItemID, Trim(Me.txt��Ŀ����.Text)) = False Then zlEditSave = 0: Exit Function
       
    End If

    gstrSql = Me.txt���Ʒ���.Tag & ",'" & Mid(Me.cbo��������.Text, InStr(1, Me.cbo��������.Text, "-") + 1) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txt��Ŀ����.Text) & "','" & Trim(Me.txt��Ŀ����.Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txt����ƴ��.Text) & "','" & Trim(Me.txt�������.Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txtAlias.Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txtӢ����д.Text) & "','" & Trim(Me.txt���㵥λ.Text) & "'"
    If InStr(Me.cbo�걾����.Text, "-") > 0 Then
        gstrSql = gstrSql & ",'" & Mid(Me.cbo�걾����.Text, InStr(Me.cbo�걾����.Text, "-") + 1) & "'," & Me.cbo�����Ա�.ListIndex
    Else
        gstrSql = gstrSql & ",'" & Mid(Me.cbo�걾����.Text, 4) & "'," & Me.cbo�����Ա�.ListIndex
    End If
    gstrSql = gstrSql & "," & Me.chk����Ӧ��.Value & "," & Me.chk�����Ŀ.Value
    
    '�����������
    gstrSql = gstrSql & "," & IIf(Trim(Me.txt�������.Text) = "", "Null", Val(Me.txt�������))
    '-- 2008-12-24 ���� ���鷽��
    gstrSql = gstrSql & ",'" & Trim(Me.txt���鷽��) & "'"
    
    If Me.chk�����Ŀ.Value = 0 Then
        If Me.cbo��Ŀ���.ListIndex = -1 Then
            MsgBox "�������Ŀ����˵����Ŀ���ʣ�", vbInformation, gstrSysName
            Me.cbo��Ŀ���.SetFocus: zlEditSave = 0: Exit Function
        End If
        gstrSql = gstrSql & "," & Left(Me.cbo��Ŀ���.Text, 1)
        
        If Me.cbo�������.ListIndex = -1 Then
            MsgBox "�������Ŀ����˵��������ͣ�", vbInformation, gstrSysName
            Me.cbo�������.SetFocus: zlEditSave = 0: Exit Function
        End If
        gstrSql = gstrSql & "," & Left(Me.cbo�������.Text, 1)
        
        gstrSql = gstrSql & ",'" & Me.cbo�����Χ.Text & "','" & Trim(Me.txtĬ�Ͻ��.Text) & "'"
        gstrSql = gstrSql & ",'" & Trim(Me.txt���㹫ʽ.Text) & "','" & Trim(Me.txtȡֵ����.Text) & "'"
        gstrSql = gstrSql & "," & Me.chkPrivacy.Value
        gstrSql = gstrSql & "," & Me.chk��ο�.Value
        
        '-begin 20080318 ����ø����Ŀ������
        If txt���Թ�ʽ.Enabled Then gstrSql = gstrSql & ",'" & Trim(Me.txt���Թ�ʽ) & "'"
        If txt�����Թ�ʽ.Enabled Then gstrSql = gstrSql & ",'" & Trim(Me.txt�����Թ�ʽ) & "'"
        If txtCutOff��ʽ.Enabled Then gstrSql = gstrSql & ",'" & Trim(Me.txtCutOff��ʽ) & "'"
        '-- End 20080318 ����ø����Ŀ������
        
        
    End If
    
    If Me.Tag = "����" Then
        gstrSql = "Zl_������Ŀ_Edit(1," & lngNewId & "," & gstrSql & ")"
    Else
        gstrSql = "Zl_������Ŀ_Edit(2," & mlngItemID & "," & gstrSql & ")"
    End If
    
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
'    If Trim(Me.Txt�Թܱ���) <> "" Then
        str�Թ� = ""
        If Me.cbo�Թ�.ListIndex > 0 Then
            str�Թ� = Me.cbo�Թ�.List(Me.cbo�Թ�.ListIndex)
            str�Թ� = Split(str�Թ�, "-")(0)
        End If
        gstrSql = "Zl_������ĿĿ¼_Batch_Update('" & IIf(Trim(str�Թ�) = "", "NULL", str�Թ�) & "'," & IIf(Me.Tag = "����", lngNewId, mlngItemID) & _
        ",'" & IIf(Me.OptApplyOnly.Value = True, "", Mid(Me.cbo��������.Text, InStr(1, Me.cbo��������, "-") + 1)) & "')"
        zlDatabase.ExecuteProcedure gstrSql, gstrSysName
'    End If
    
    If Me.Tag = "����" Then mlngItemID = lngNewId
    Me.Enabled = False: Me.Tag = ""
    Me.BackColor = Me.cmd���Ʒ���.BackColor: Me.chk����Ӧ��.BackColor = Me.BackColor: Me.chk�����Ŀ.BackColor = Me.BackColor
    Me.OptApplyOnly.BackColor = Me.cmd���Ʒ���.BackColor: Me.OptApplyType.BackColor = Me.cmd���Ʒ���.BackColor
    zlEditSave = mlngItemID: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

Private Function Check_Expression(ByVal intType As Integer, strExpression As String) As Boolean
    '��֤ø����ʾ�Ƿ���ȷ
    'inttype =0 �߼����ʽ��1���Ϸ��ļ�����ʽ
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    If intType = 0 Then
        strSql = "Select 1 From Dual Where " & strExpression
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        Check_Expression = True
    ElseIf intType = 1 Then
        strSql = "Select " & strExpression & " From Dual "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        Check_Expression = True
    End If
    Exit Function
errHandle:
    Check_Expression = False
End Function
'--------------------------------------------
'����Ϊ����ؼ���Ӧ�¼�
'--------------------------------------------

Private Sub cbo�걾����_Click()
'   ���ݱ걾�����Ա�
    If Me.cbo�걾����.ListIndex >= 0 Then
        If Me.cbo�걾����.ItemData(Me.cbo�걾����.ListIndex) = 1 Then
            Me.cbo�����Ա�.ListIndex = 1
            Me.cbo�����Ա�.Enabled = False
        ElseIf cbo�걾����.ItemData(Me.cbo�걾����.ListIndex) = 2 Then
            Me.cbo�����Ա�.ListIndex = 2
            Me.cbo�����Ա�.Enabled = False
        Else
            Me.cbo�����Ա�.Enabled = True
        End If
    End If
End Sub

Private Sub cbo�걾����_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo�걾����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo��������_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo��������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo�����Χ_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo�����Χ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo�������_Click()
    Select Case Left(Me.cbo�������, 1)
    Case 1: Me.txtȡֵ����.Text = "": Me.txtȡֵ����.Enabled = False
    Case 2
        Me.txtȡֵ����.Enabled = True
        If Me.txtȡֵ����.Text = "-;��;+;++;+++;++++" Then Me.txtȡֵ���� = ""
    Case 3
        Me.txtȡֵ����.Enabled = True
        If Trim(Me.txtȡֵ����.Text) = "" Then Me.txtȡֵ����.Text = "-;��;+;++;+++;++++"
    End Select
End Sub

Private Sub cbo�������_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo�������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo�Թ�_Click()
    Dim lngColor As Long
    
    If cbo�Թ�.ListIndex > 0 Then
        
        lngColor = cbo�Թ�.ItemData(cbo�Թ�.ListIndex)
        If lngColor < 0 Then lngColor = 0
        On Error Resume Next
        lblColor.BackColor = lngColor
    Else
        lblColor.BackColor = Label2.BackColor
    End If
End Sub

Private Sub cbo�Թ�_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo�Թ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo�����Ա�_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo�����Ա�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo��Ŀ���_Click()
    '΢������Ŀ�������������Ŀ
    If Me.cbo��Ŀ���.ListIndex = 1 Then
        Me.chk�����Ŀ.Value = 0: Me.chk�����Ŀ.Enabled = False
    Else
        Me.chk�����Ŀ.Enabled = True
    End If
    '�Ǽ�����Ŀ�������ü��㹫ʽ
    If Me.cbo��Ŀ���.ListIndex = 2 Then
        Me.txt���㹫ʽ.Text = Me.txt���㹫ʽ.Tag: Me.txt���㹫ʽ.Enabled = True
        Me.cmdFormula.Enabled = True
    Else
        Me.txt���㹫ʽ.Tag = Me.txt���㹫ʽ.Text: Me.txt���㹫ʽ.Enabled = False
        Me.cmdFormula.Enabled = False
    End If
    
    '��ø����Ŀ��������ø�깫ʽ
    If Me.cbo��Ŀ���.ListIndex = 3 Then
        Me.txt���Թ�ʽ.Text = Me.txt���Թ�ʽ.Tag: Me.txt���Թ�ʽ.Enabled = True
        Me.txt�����Թ�ʽ.Text = Me.txt�����Թ�ʽ.Tag: Me.txt�����Թ�ʽ.Enabled = True
        Me.txtCutOff��ʽ.Text = Me.txtCutOff��ʽ.Tag: Me.txtCutOff��ʽ.Enabled = True
    Else
        Me.txt���Թ�ʽ.Tag = Me.txt���Թ�ʽ.Text: Me.txt���Թ�ʽ.Enabled = False
        Me.txt�����Թ�ʽ.Tag = Me.txt�����Թ�ʽ.Text: Me.txt�����Թ�ʽ.Enabled = False
        Me.txtCutOff��ʽ.Tag = Me.txtCutOff��ʽ.Text: Me.txtCutOff��ʽ.Enabled = False
    End If
End Sub

Private Sub cbo��Ŀ���_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo��Ŀ���_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk����Ӧ��_Click()
    If Me.chk�����Ŀ.Value = 0 Then
        Me.chkPrivacy.Visible = True
        Me.chk��ο�.Visible = True
    Else
        Me.chkPrivacy.Value = 0
        Me.chkPrivacy.Visible = False
        Me.chk��ο�.Value = 0
        Me.chk��ο�.Visible = False
    End If
End Sub

Private Sub chk����Ӧ��_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chk����Ӧ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk�����Ŀ_Click()
    If Me.chk�����Ŀ.Value = 0 Then
        Me.txtӢ����д.Text = Me.txtӢ����д.Tag: Me.txtӢ����д.Enabled = True
        Me.cbo��Ŀ���.ListIndex = Val(Me.cbo��Ŀ���.Tag): Me.cbo��Ŀ���.Enabled = True
        Me.cbo�������.ListIndex = Val(Me.cbo�������.Tag): Me.cbo�������.Enabled = True
        Me.cbo�����Χ.ListIndex = Val(Me.cbo�����Χ.Tag): Me.cbo�����Χ.Enabled = True
        Me.txtĬ�Ͻ��.Text = Me.txtĬ�Ͻ��.Tag: Me.txtĬ�Ͻ��.Enabled = True
        
        Me.txt���㹫ʽ.Enabled = (Me.cbo��Ŀ���.ListIndex = 2)
        If Me.txt���㹫ʽ.Enabled Then Me.txt���㹫ʽ.Text = Me.txt���㹫ʽ.Tag
        
        Me.txtȡֵ����.Tag = Me.txtȡֵ����.Text: Me.txtȡֵ����.Enabled = True
        chk����Ӧ��.Enabled = True
        Me.chkPrivacy.Visible = True
        
        Me.txt���Թ�ʽ.Enabled = (Me.cbo��Ŀ���.ListIndex = 3)
        Me.txt�����Թ�ʽ.Enabled = (Me.cbo��Ŀ���.ListIndex = 3)
        Me.txtCutOff��ʽ.Enabled = (Me.cbo��Ŀ���.ListIndex = 3)
        If Me.txt���Թ�ʽ.Enabled Then Me.txt���Թ�ʽ.Text = Me.txt���Թ�ʽ.Tag
        If Me.txt�����Թ�ʽ.Enabled Then Me.txt�����Թ�ʽ.Text = Me.txt�����Թ�ʽ.Tag
        If Me.txtCutOff��ʽ.Enabled Then Me.txtCutOff��ʽ.Text = Me.txtCutOff��ʽ.Tag
        
        Me.chk��ο�.Visible = True
    Else
        Me.txtӢ����д.Tag = Me.txtӢ����д.Text: Me.txtӢ����д.Text = "": Me.txtӢ����д.Enabled = False
        Me.cbo��Ŀ���.Tag = Me.cbo��Ŀ���.ListIndex: Me.cbo��Ŀ���.ListIndex = -1: Me.cbo��Ŀ���.Enabled = False
        Me.cbo�������.Tag = Me.cbo�������.ListIndex: Me.cbo�������.ListIndex = -1: Me.cbo�������.Enabled = False
        Me.cbo�����Χ.Tag = Me.cbo�����Χ.ListIndex: Me.cbo�����Χ.ListIndex = -1: Me.cbo�����Χ.Enabled = False
        Me.txtĬ�Ͻ��.Tag = Me.txtĬ�Ͻ��.Text: Me.txtĬ�Ͻ��.Text = "": Me.txtĬ�Ͻ��.Enabled = False
        Me.txt���㹫ʽ.Tag = Me.txt���㹫ʽ.Text: Me.txt���㹫ʽ.Text = "": Me.txt���㹫ʽ.Enabled = False
        Me.txtȡֵ����.Tag = Me.txtȡֵ����.Text: Me.txtȡֵ����.Text = "": Me.txtȡֵ����.Enabled = False
        
        Me.txt���Թ�ʽ.Tag = Me.txt���Թ�ʽ.Text: Me.txt���Թ�ʽ.Text = "": Me.txt���Թ�ʽ.Enabled = False
        Me.txt�����Թ�ʽ.Tag = Me.txt�����Թ�ʽ.Text: Me.txt�����Թ�ʽ.Text = "": Me.txt�����Թ�ʽ.Enabled = False
        Me.txtCutOff��ʽ.Tag = Me.txtCutOff��ʽ.Text: Me.txtCutOff��ʽ.Text = "": Me.txtCutOff��ʽ.Enabled = False
        
        chk����Ӧ��.Enabled = False: chk����Ӧ��.Value = 1: Me.chkPrivacy.Visible = False: Me.chkPrivacy.Value = 0: Me.chk��ο�.Visible = False:  Me.chk��ο�.Value = 0
    End If
End Sub

Private Sub chk�����Ŀ_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chk�����Ŀ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdFormula_Click()
    txt���㹫ʽ = FrmLabItemFormula.DefFormula(mlngItemID, txt���㹫ʽ, Me)
End Sub

Private Sub cmd���Ʒ���_Click()
    With Me.tvwClass
        .Left = Me.txt���Ʒ���.Left
        .Top = Me.txt���Ʒ���.Top + Me.txt���Ʒ���.Height
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    If Me.tvwClass.Visible Then
        On Error Resume Next
        Me.tvwClass.Visible = False: Me.txt���Ʒ���.SetFocus: Exit Sub
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
    
Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    Err = 0: On Error GoTo ErrHand
    
    '�ֶγ�������
    gstrSql = "Select A.����, A.����, A.���㵥λ, B.���� From ������ĿĿ¼ A, ������Ŀ���� B " & _
            " Where A.ID = B.������Ŀid And A.ID = 0 And B.���� = 1"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
    With rsTemp
        Me.txt��Ŀ����.MaxLength = .Fields("����").DefinedSize
        Me.txt��Ŀ����.MaxLength = .Fields("����").DefinedSize
        Me.txt���㵥λ.MaxLength = .Fields("���㵥λ").DefinedSize
        Me.txt����ƴ��.MaxLength = .Fields("����").DefinedSize
        Me.txt�������.MaxLength = .Fields("����").DefinedSize
    End With
    
    gstrSql = "Select A.��д, A.Ĭ��ֵ, A.���㹫ʽ, A.ȡֵ����, A.���Թ�ʽ, A.�����Թ�ʽ, A.CutOff��ʽ From ������Ŀ A Where A.������ĿID = 0"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
    With rsTemp
        Me.txtӢ����д.MaxLength = .Fields("��д").DefinedSize
        Me.txtӢ����д.MaxLength = .Fields("Ĭ��ֵ").DefinedSize
        Me.txt���㹫ʽ.MaxLength = .Fields("���㹫ʽ").DefinedSize
        Me.txtȡֵ����.MaxLength = .Fields("ȡֵ����").DefinedSize
        Me.txt���Թ�ʽ.MaxLength = .Fields("���Թ�ʽ").DefinedSize
        Me.txt�����Թ�ʽ.MaxLength = .Fields("�����Թ�ʽ").DefinedSize
        Me.txtCutOff��ʽ.MaxLength = .Fields("CutOff��ʽ").DefinedSize
    End With
    
    '���Ʒ���װ��
    gstrSql = "select ID,�ϼ�ID,����,����,����" & _
            " From ���Ʒ���Ŀ¼" & _
            " Where ���� = 5" & _
            " start with �ϼ�ID is null" & _
            " connect by prior ID=�ϼ�ID"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
'        If .BOF Or .EOF Then MsgBox "�����Ƚ������Ʒ�����Ŀ!", vbExclamation, gstrSysName: Exit Sub
    Me.tvwClass.Nodes.Clear
    With rsTemp
        Do While Not .EOF
            If IsNull(!�ϼ�ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, "[" & !���� & "]" & !����, "close")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !�ϼ�ID, tvwChild, "_" & !ID, "[" & !���� & "]" & !����, "close")
            End If
            objNode.Sorted = True
            objNode.Tag = IIf(IsNull(!����), "", !����)
            objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
        If Me.tvwClass.Nodes.Count > 0 Then
            Me.tvwClass.Nodes(1).Selected = True
            Me.txt���Ʒ���.Text = Me.tvwClass.SelectedItem.Text
            Me.txt���Ʒ���.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
        End If
    End With
    
    '�����������
    gstrSql = "Select ����,���� From ���Ƽ�������"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
    Me.cbo��������.Clear
    
    With rsTemp
        Do While Not .EOF
            Me.cbo��������.AddItem !���� & "-" & !����
            .MoveNext
        Loop
        If Me.cbo��������.ListCount > 0 Then Me.cbo��������.ListIndex = 0
    
        '�Ա�����Ҫ���ڼ���걾����װ�룬�����걾����Ҫ�õ��������
        aryTemp = Split("0-���Ա�����;1-����;2-Ů��", ";")
        For lngCount = LBound(aryTemp) To UBound(aryTemp)
            Me.cbo�����Ա�.AddItem aryTemp(lngCount)
        Next
        Me.cbo�����Ա�.ListIndex = 0
    End With
        '�����������
        
        gstrSql = "Select ����,����,�����Ա� From ���Ƽ���걾"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
        Me.cbo�걾����.Clear
    With rsTemp
        Do While Not .EOF
            Me.cbo�걾����.AddItem !���� & "-" & !����
            If InStr(Trim("" & !�����Ա�), "��") > 0 Then
                Me.cbo�걾����.ItemData(Me.cbo�걾����.NewIndex) = 1
            ElseIf InStr(Trim("" & !�����Ա�), "Ů") > 0 Then
                Me.cbo�걾����.ItemData(Me.cbo�걾����.NewIndex) = 2
            Else
                Me.cbo�걾����.ItemData(Me.cbo�걾����.NewIndex) = 0
            End If
            
            .MoveNext
        Loop
        If Me.cbo�걾����.ListCount > 0 Then Me.cbo�걾����.ListIndex = 0
    End With
    
    '��������Χ
    gstrSql = "Select Distinct ����  From ����������"
'    If .State = adStateOpen Then .Close
'    Call SQLTest(App.ProductName, Me.Caption, gstrSql)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'    Call SQLTest
    Me.cbo�����Χ.Clear
    Me.cbo�����Χ.AddItem ""
    With rsTemp
        Do While Not .EOF
            Me.cbo�����Χ.AddItem "" & !����
            .MoveNext
        Loop
        If Me.cbo�����Χ.ListCount > 0 Then Me.cbo�����Χ.ListIndex = 0
    
    End With
    '�����̶�����װ��
    aryTemp = Split("1-��ͨ����;2-΢����;3-������Ŀ;4-ø����Ŀ", ";")
    For lngCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo��Ŀ���.AddItem aryTemp(lngCount)
    Next
    Me.cbo��Ŀ���.ListIndex = 0
    
    aryTemp = Split("1-����;2-����;3-�붨��", ";")
    For lngCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo�������.AddItem aryTemp(lngCount)
    Next
    Me.cbo�������.ListIndex = 0
    
    '�Թ�
    gstrSql = "Select ����,����,��ɫ From ��Ѫ������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Me.cbo�Թ�.Clear
    Me.cbo�Թ�.AddItem "<δ����>"
    Me.cbo�Թ�.ItemData(cbo�Թ�.NewIndex) = Me.BackColor
    Do Until rsTemp.EOF
        Me.cbo�Թ�.AddItem rsTemp!���� & "-" & rsTemp!����
        Me.cbo�Թ�.ItemData(cbo�Թ�.NewIndex) = rsTemp!��ɫ
        rsTemp.MoveNext
    Loop
    If Me.cbo�Թ�.ListCount > 0 Then Me.cbo�Թ�.ListIndex = 0
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Sub tvwClass_DblClick()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Me.txt���Ʒ���.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
    Me.txt���Ʒ���.Text = Me.tvwClass.SelectedItem.Text
    Me.txt���Ʒ���.SetFocus
End Sub

Private Sub tvwClass_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
        If Me.tvwClass.SelectedItem.Children > 0 Then Exit Sub
        Call tvwClass_DblClick
    Case vbKeySpace
        Call tvwClass_DblClick
    Case vbKeyEscape
        Call tvwClass_LostFocus
    End Select
End Sub

Private Sub tvwClass_LostFocus()
    If Me.cmd���Ʒ��� Is ActiveControl Then Exit Sub
    Me.tvwClass.Visible = False
End Sub

Private Sub txtAlias_GotFocus()
    Me.txtAlias.SelStart = 0
    Me.txtAlias.SelLength = Len(Me.txtAlias.Text)
End Sub

Private Sub txtAlias_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtCutOff��ʽ_GotFocus()
    Me.txtCutOff��ʽ.SelStart = 0: Me.txtCutOff��ʽ.SelLength = 1000
End Sub

Private Sub txtCutOff��ʽ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt���㵥λ_GotFocus()
    Me.txt���㵥λ.SelStart = 0: Me.txt���㵥λ.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt���㵥λ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
'    If InStr(" ~!@#$%^&*_|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt���㹫ʽ_GotFocus()
    Me.txt���㹫ʽ.SelStart = 0: Me.txt���㹫ʽ.SelLength = 1000
End Sub

Private Sub txt���㹫ʽ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt���鷽��_GotFocus()
    Me.txt���鷽��.SelStart = 0: Me.txt���鷽��.SelLength = Me.txt���鷽��.MaxLength
End Sub

Private Sub txt���鷽��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt����ƴ��_GotFocus()
    Me.txt����ƴ��.SelStart = 0: Me.txt����ƴ��.SelLength = 1000
End Sub

Private Sub txt����ƴ��_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt�������_GotFocus()
    Me.txt�������.SelStart = 0: Me.txt�������.SelLength = 1000
End Sub

Private Sub txt�������_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtĬ�Ͻ��_GotFocus()
    Me.txtĬ�Ͻ��.SelStart = 0: Me.txtĬ�Ͻ��.SelLength = 1000
End Sub

Private Sub txtĬ�Ͻ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt�������_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txtȡֵ����_GotFocus()
    Me.txtȡֵ����.SelStart = 0: Me.txtȡֵ����.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtȡֵ����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt�����Թ�ʽ_GotFocus()
    Me.txt�����Թ�ʽ.SelStart = 0: Me.txt�����Թ�ʽ.SelLength = 1000
End Sub

Private Sub txt�����Թ�ʽ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt��Ŀ����_GotFocus()
    Me.txt��Ŀ����.SelStart = 0: Me.txt��Ŀ����.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt��Ŀ����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt��Ŀ����_GotFocus()
    Me.txt��Ŀ����.SelStart = 0: Me.txt��Ŀ����.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��Ŀ����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txt��Ŀ����.Text = MoveSpecialChar(txt��Ŀ����.Text)
        Me.txt����ƴ��.Text = zlStr.GetCodeByORCL(Me.txt��Ŀ����.Text, False, Me.txt����ƴ��.MaxLength)
        Me.txt�������.Text = zlStr.GetCodeByORCL(Me.txt��Ŀ����.Text, True, Me.txt�������.MaxLength)
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End If
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt��Ŀ����_LostFocus()
    Me.txt����ƴ��.Text = zlStr.GetCodeByORCL(Me.txt��Ŀ����.Text, False, Me.txt����ƴ��.MaxLength)
    Me.txt�������.Text = zlStr.GetCodeByORCL(Me.txt��Ŀ����.Text, True, Me.txt�������.MaxLength)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt���Թ�ʽ_GotFocus()
    Me.txt���Թ�ʽ.SelStart = 0: Me.txt���Թ�ʽ.SelLength = 1000
End Sub

Private Sub txt���Թ�ʽ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtӢ����д_GotFocus()
    Me.txtӢ����д.SelStart = 0: Me.txtӢ����д.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtӢ����д_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt���Ʒ���_Change()
    Me.txt���Ʒ���.SelStart = 0: Me.txt���Ʒ���.SelLength = 1000
End Sub

Private Sub txt���Ʒ���_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End Select
    KeyAscii = 0
End Sub


