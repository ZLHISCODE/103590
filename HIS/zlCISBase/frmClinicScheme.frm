VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClinicScheme 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���׷���"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   Icon            =   "frmClinicScheme.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkAll 
      Caption         =   "���ñ�����ʱȫѡ"
      Height          =   270
      Left            =   4890
      TabIndex        =   47
      ToolTipText     =   "��ѡʱҽ���´���ñ�����ʱĬ��ȫѡ������Ŀ������ѡ�κ���Ŀ��"
      Top             =   2310
      Width           =   1770
   End
   Begin VB.TextBox txtFind 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   300
      Left            =   5160
      TabIndex        =   26
      Top             =   3225
      Width           =   975
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Enabled         =   0   'False
      Height          =   300
      Left            =   6120
      TabIndex        =   27
      Top             =   3225
      Width           =   855
   End
   Begin VB.TextBox txt����ʱ�� 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   300
      Left            =   4560
      MaxLength       =   13
      TabIndex        =   45
      Top             =   5280
      Width           =   2370
   End
   Begin VB.TextBox txt������ 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   300
      Left            =   1080
      MaxLength       =   13
      TabIndex        =   43
      Top             =   5280
      Width           =   1890
   End
   Begin VB.Frame fraline 
      Height          =   45
      Index           =   3
      Left            =   0
      TabIndex        =   41
      Top             =   5640
      Width           =   7335
   End
   Begin VB.ComboBox cmbStationNo 
      Height          =   300
      Left            =   825
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   2250
      Visible         =   0   'False
      Width           =   2220
   End
   Begin MSComctlLib.ListView lvw���� 
      Height          =   1380
      Left            =   1125
      TabIndex        =   29
      Top             =   3615
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   2434
      View            =   2
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483633
      Appearance      =   1
      Enabled         =   0   'False
      NumItems        =   0
   End
   Begin VB.CheckBox chk��Χ 
      Caption         =   "סԺʹ��(&I)"
      Height          =   195
      Index           =   1
      Left            =   5805
      TabIndex        =   22
      Top             =   2880
      Value           =   1  'Checked
      Width           =   1290
   End
   Begin VB.CheckBox chk��Χ 
      Caption         =   "����ʹ��(&C)"
      Height          =   195
      Index           =   0
      Left            =   4470
      TabIndex        =   21
      Top             =   2880
      Value           =   1  'Checked
      Width           =   1290
   End
   Begin VB.Frame fraline 
      Height          =   30
      Index           =   2
      Left            =   0
      TabIndex        =   38
      Top             =   2640
      Width           =   8490
   End
   Begin VB.OptionButton opt��Χ 
      Caption         =   "����(&1)"
      Height          =   180
      Index           =   0
      Left            =   1110
      TabIndex        =   18
      Top             =   2880
      Width           =   930
   End
   Begin VB.OptionButton opt��Χ 
      Caption         =   "ȫԺ(&3)"
      Height          =   180
      Index           =   2
      Left            =   3135
      TabIndex        =   20
      Top             =   2880
      Value           =   -1  'True
      Width           =   930
   End
   Begin VB.OptionButton opt��Χ 
      Caption         =   "����(&2)"
      Height          =   180
      Index           =   1
      Left            =   2115
      TabIndex        =   19
      Top             =   2880
      Width           =   930
   End
   Begin VB.ComboBox cbo��Ա 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   300
      Left            =   1125
      TabIndex        =   24
      Top             =   3225
      Width           =   3030
   End
   Begin VB.CommandButton cmdScheme 
      Caption         =   "��������(&E)��"
      Height          =   350
      Left            =   1590
      TabIndex        =   30
      Top             =   5775
      Width           =   1590
   End
   Begin VB.TextBox txt˵�� 
      Height          =   300
      Left            =   825
      MaxLength       =   30
      TabIndex        =   16
      Top             =   1875
      Width           =   5835
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Index           =   1
      Left            =   825
      MaxLength       =   40
      TabIndex        =   11
      Top             =   1485
      Width           =   2250
   End
   Begin VB.TextBox txtƴ�� 
      Height          =   300
      Index           =   1
      Left            =   4080
      MaxLength       =   12
      TabIndex        =   13
      Top             =   1485
      Width           =   960
   End
   Begin VB.TextBox txt��� 
      Height          =   300
      Index           =   1
      Left            =   5700
      MaxLength       =   12
      TabIndex        =   14
      Top             =   1485
      Width           =   960
   End
   Begin VB.Frame fraline 
      Height          =   30
      Index           =   1
      Left            =   -45
      TabIndex        =   37
      Top             =   5160
      Width           =   8490
   End
   Begin VB.TextBox txt��� 
      Height          =   300
      Index           =   0
      Left            =   5700
      MaxLength       =   12
      TabIndex        =   9
      Top             =   1110
      Width           =   960
   End
   Begin VB.TextBox txtƴ�� 
      Height          =   300
      Index           =   0
      Left            =   4080
      MaxLength       =   12
      TabIndex        =   8
      Top             =   1110
      Width           =   960
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Index           =   0
      Left            =   825
      MaxLength       =   40
      TabIndex        =   6
      Top             =   1110
      Width           =   2250
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   825
      MaxLength       =   13
      TabIndex        =   1
      Top             =   735
      Width           =   2250
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   3570
      Left            =   180
      TabIndex        =   35
      TabStop         =   0   'False
      Tag             =   "1000"
      Top             =   6345
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   6297
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
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5865
      TabIndex        =   32
      Top             =   5775
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   420
      Picture         =   "frmClinicScheme.frx":058A
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   5775
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4800
      TabIndex        =   31
      Top             =   5775
      Width           =   1100
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   4080
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   3
      Top             =   735
      Width           =   2580
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "&P"
      Height          =   285
      Left            =   6675
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   735
      Width           =   285
   End
   Begin VB.Frame fraline 
      Height          =   60
      Index           =   0
      Left            =   -30
      TabIndex        =   36
      Top             =   540
      Width           =   8490
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   3780
      Top             =   6375
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
            Picture         =   "frmClinicScheme.frx":06D4
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicScheme.frx":0C6E
            Key             =   "expend"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ҿ��ң�"
      Height          =   180
      Left            =   4320
      TabIndex        =   25
      Top             =   3285
      Width           =   900
   End
   Begin VB.Label lbl����ʱ�� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����ʱ��(&T)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3480
      TabIndex        =   46
      Top             =   5340
      Width           =   990
   End
   Begin VB.Label lbl������ 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������(&P)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   44
      Top             =   5340
      Width           =   810
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   15
      Left            =   480
      TabIndex        =   42
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label lblStationNo 
      AutoSize        =   -1  'True
      Caption         =   "Ժ��(&C)"
      Height          =   180
      Left            =   135
      TabIndex        =   40
      Top             =   2310
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "ʹ�ÿ��ң�"
      Height          =   180
      Left            =   210
      TabIndex        =   28
      Top             =   3690
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ʹ�÷�Χ��"
      Height          =   180
      Left            =   210
      TabIndex        =   17
      Top             =   2880
      Width           =   900
   End
   Begin VB.Label lbl��Ա 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ʹ����Ա��"
      Height          =   180
      Left            =   210
      TabIndex        =   23
      Top             =   3285
      Width           =   900
   End
   Begin VB.Label lbl˵�� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "˵��(&Z)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   135
      TabIndex        =   15
      Top             =   1920
      Width           =   630
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   390
      Picture         =   "frmClinicScheme.frx":1208
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&A)"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   135
      TabIndex        =   10
      Top             =   1545
      Width           =   630
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&M)           (ƴ��)            (���)"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   3420
      TabIndex        =   12
      Top             =   1545
      Width           =   3780
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    ���ݳ��õĵ���ҽ���������ʵ�ɸѡ���γɳ��׵�ҽ���������Է���ҽ�����ٵ��´ﲡ��ҽ����"
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   1155
      TabIndex        =   34
      Top             =   105
      Width           =   5925
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&S)           (ƴ��)            (���)"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   3420
      TabIndex        =   7
      Top             =   1170
      Width           =   3780
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   5
      Top             =   1170
      Width           =   630
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   135
      TabIndex        =   0
      Top             =   795
      Width           =   630
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&F)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3420
      TabIndex        =   2
      Top             =   795
      Width           =   630
   End
End
Attribute VB_Name = "frmClinicScheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'˵����
'   1���ϼ�����ͨ��������ShowMe�������������塢Ȩ�ޡ��༭��Ŀ�ķ���ID��ID,�༭״̬����Ϣ���ݽ��뱾����
'   2���༭״̬����Me.tag��ţ��ֱ�Ϊ"����"��"�޸�"��"����"�����ϼ�����ͨ��ShowMe����
'---------------------------------------------------
Private mint��Χ As Integer '1-����,2-סԺ,3-�����סԺ
Private mstrPrivs As String
Private lngClassId As Long       '���༭�ķ���ID���ϼ�����ͨ��ShowMe���ݽ���
Private lngItemId As Long        '���༭����ĿID���޸ġ�����ʱ���ϼ�����ͨ��ShowMe���ݽ���,����ʱΪ0��
Private mblnNoCheck As Boolean
Private mblnFirst As Boolean
Private mstrLike As String
Private mblnChange As Boolean
Private mlngFind As Long

Private rsTemp As New ADODB.Recordset
Private mrsScheme As ADODB.Recordset
Private mblnOK As Boolean

Public Function ShowMe(ByVal frmParent As Object, ByVal strPrivs As String, ByVal byt״̬ As Byte, _
    ByVal lng����id As Long, Optional ByVal lng��Ŀid As Long, Optional ByVal int��Χ As Integer = 3) As Boolean
    '---------------------------------------------------
    '���ܣ��ϼ�������ñ�����ģ����ݲ���������ʾ����
    '---------------------------------------------------
    Dim objNode As Node
    
    mint��Χ = int��Χ
    mstrPrivs = strPrivs
    
    Me.Tag = Switch(byt״̬ = 0, "����", byt״̬ = 1, "�޸�", byt״̬ = 2, "����")
    Me.Caption = "���׷���" & Me.Tag
    lngClassId = lng����id: lngItemId = lng��Ŀid
    
    '��д��Ҫѡ�������
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select ID,�ϼ�ID,����,����,����" & _
                " From ���Ʒ���Ŀ¼ Where ����=6 And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Start With �ϼ�ID is Null Connect by Prior ID=�ϼ�ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "�����Ƚ����䷽���Ʒ�����Ŀ֮�������䷽", vbExclamation, gstrSysName
            Unload Me: Exit Function
        End If
        
        Me.tvwClass.Nodes.Clear
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
        Me.tvwClass.Nodes("_" & lng����id).Selected = True
        Me.txt����.Text = Me.tvwClass.SelectedItem.Text
        Me.txt����.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
    End With
    
    '��ʾ����
    Me.Show 1, frmParent
    ShowMe = mblnOK
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub GetDefineSize()
    '���ܣ��õ����ݿ�ı��ֶεĳ���
    On Error GoTo errHandle
    Dim rsTmp As New ADODB.Recordset
    gstrSql = "Select A.����,A.�걾��λ,B.����,B.���� From ������ĿĿ¼ A, ������Ŀ���� B Where A.ID=B.������ĿID and A.ID=0"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSql, Me.Caption)
    
    txt����.MaxLength = rsTmp.Fields("����").DefinedSize
    txt����(0).MaxLength = rsTmp.Fields("����").DefinedSize
    txt����(1).MaxLength = rsTmp.Fields("����").DefinedSize
    txtƴ��(0).MaxLength = rsTmp.Fields("����").DefinedSize
    txtƴ��(1).MaxLength = rsTmp.Fields("����").DefinedSize
    txt���(0).MaxLength = rsTmp.Fields("����").DefinedSize
    txt���(1).MaxLength = rsTmp.Fields("����").DefinedSize
    txt˵��.MaxLength = rsTmp.Fields("�걾��λ").DefinedSize

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo��Ա_Click()
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Long, strSql As String
    Dim intIdx As Integer, i As Long
    Dim vRect As RECT, blnCancel As Boolean
    
    If cbo��Ա.ListIndex <> -1 Then
        If cbo��Ա.ItemData(cbo��Ա.ListIndex) = 0 Then
            strSql = "Select ID,���,����,����,�Ա� From ��Ա�� Where ����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null Order by ���"
            vRect = zlControl.GetControlRect(cbo��Ա.hWnd)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSql, 0, "��Ա", , , , , , True, vRect.Left, vRect.Top, cbo��Ա.Height, blnCancel, , True)
            If Not rsTmp Is Nothing Then
                intIdx = Cbo.FindIndex(cbo��Ա, rsTmp!ID)
                If intIdx <> -1 Then
                    cbo��Ա.ListIndex = intIdx
                Else
                    cbo��Ա.AddItem rsTmp!��� & "-" & rsTmp!����, 0
                    cbo��Ա.ItemData(cbo��Ա.NewIndex) = rsTmp!ID
                    cbo��Ա.ListIndex = cbo��Ա.NewIndex
                End If
                mblnChange = True
            Else
                If Not blnCancel Then
                    MsgBox "û����Ա���ݣ����ȵ���Ա���������á�", vbInformation, gstrSysName
                End If
                Call zlControl.CboSetIndex(cbo��Ա.hWnd, Val(cbo��Ա.Tag))
            End If
        Else
            cbo��Ա.Tag = cbo��Ա.ListIndex
        End If
    Else
        cbo��Ա.Tag = cbo��Ա.ListIndex
    End If
End Sub

Private Sub cbo��Ա_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo��Ա.ListIndex = -1 Then
            Call cbo��Ա_Validate(blnCancel)
        End If
        If Not blnCancel Then Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo��Ա_Validate(Cancel As Boolean)
'���ܣ��������������,�Զ�ƥ��ִ�п���
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, intIdx As Long, i As Long
    Dim vRect As RECT, blnCancel As Boolean
    
    If cbo��Ա.ListIndex <> -1 Then Exit Sub '��ѡ��
    If cbo��Ա.Text = "" Then Exit Sub '������
    
    On Error GoTo errH
    
    strSql = "Select ID,���,����,����,�Ա� From ��Ա��" & _
        " Where Upper(���) Like '" & UCase(cbo��Ա.Text) & "%'" & _
        " Or Upper(����) Like '" & mstrLike & UCase(cbo��Ա.Text) & "%'" & _
        " Or Upper(����) Like '" & mstrLike & UCase(cbo��Ա.Text) & "%'" & _
        " And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) " & _
        " Order by ���"
    vRect = zlControl.GetControlRect(cbo��Ա.hWnd)
    Set rsTmp = zlDatabase.ShowSelect(Me, strSql, 0, "��Ա", , , , , , True, vRect.Left, vRect.Top, cbo��Ա.Height, blnCancel, , True)
    If Not rsTmp Is Nothing Then
        intIdx = Cbo.FindIndex(cbo��Ա, rsTmp!ID)
        If intIdx <> -1 Then
            cbo��Ա.ListIndex = intIdx
        Else
            cbo��Ա.AddItem rsTmp!��� & "-" & rsTmp!����, 0
            cbo��Ա.ItemData(cbo��Ա.NewIndex) = rsTmp!ID
            cbo��Ա.ListIndex = cbo��Ա.NewIndex
        End If
        mblnChange = True
    Else
        If Not blnCancel Then
            MsgBox "δ�ҵ���Ӧ����Ա��", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub chkAll_Click()
    If mblnNoCheck Then Exit Sub
    mblnChange = True
End Sub

Private Sub chk��Χ_Click(Index As Integer)
    
    If mblnNoCheck Then Exit Sub
    
    If Index = 1 And chk��Χ((Index + 1) Mod 2).Value = 1 And chk��Χ(Index).Value = 0 Then
        If Not mrsScheme Is Nothing Then
            mrsScheme.Filter = "��Ч=0"
            If mrsScheme.RecordCount > 0 Then
                MsgBox "�˳��׷����д��ڳ�������������Ϊ��ʹ�������", vbInformation, gstrSysName
                mblnNoCheck = True
                chk��Χ(Index).Value = 1
                mblnNoCheck = False
                mrsScheme.Filter = "": Exit Sub
            End If
            mrsScheme.Filter = ""
        End If
        
    ElseIf chk��Χ((Index + 1) Mod 2).Value = 0 And chk��Χ(Index).Value = 0 Then
        mblnNoCheck = True
        chk��Χ(Index).Value = 1
        mblnNoCheck = False
        Exit Sub
    End If
    
    If InStr(mstrPrivs, "ȫԺ���׷���") > 0 Then
        Call LoadDeptList(True)
    ElseIf InStr(mstrPrivs, "���Ƴ��׷���") > 0 Then
        Call LoadDeptList(False)
    Else
    End If
    
    mblnChange = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me: Exit Sub
End Sub

Private Sub cmdFind_Click()
    Dim strFind As String
    Dim i As Long
    Dim blnIsFind As Boolean
    
    strFind = UCase(Trim(txtFind.Text))
    If strFind = "" Then Exit Sub
    For i = mlngFind To Lvw����.ListItems.Count
        If zlCommFun.SpellCode(Mid(Lvw����.ListItems(i).Text, InStr(Lvw����.ListItems(i).Text, "-") + 1)) Like UCase(IIf(gstrMatch <> "", "*", "") & strFind & "*") Or _
                UCase(Lvw����.ListItems(i).Text) Like UCase(IIf(gstrMatch <> "", "*", "") & strFind & "*") Then
            Lvw����.ListItems(i).Selected = True
            Lvw����.ListItems(i).EnsureVisible
            Lvw����.SetFocus
            blnIsFind = True
            mlngFind = i + 1
            Exit For
        End If
    Next
    If blnIsFind = False Then
        If mlngFind = 1 Then
            MsgBox "û���ҵ������ҵĿ��ҡ�", vbInformation, Me.Caption
        Else
            MsgBox "�Ѿ������һ�������ˡ�", vbInformation, Me.Caption
            mlngFind = 1
        End If
    End If
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Function Get�������() As Integer
    If chk��Χ(0).Value = 1 And chk��Χ(1).Value = 1 Then
        Get������� = 3
    ElseIf chk��Χ(0).Value = 1 Then
        Get������� = 1
    ElseIf chk��Χ(1).Value = 1 Then
        Get������� = 2
    End If
End Function

Private Sub cmdOK_Click()
    Dim arrSql() As Variant
    Dim strTmp As String, i As Long
    Dim str����IDs As String, lng��ԱID As Long
    Dim strSql As String
    Dim strվ�� As String
    Dim str���� As String
    
    '���¼�����ƣ���ȥ�������ַ�
    strTmp = MoveSpecialChar(txt����(0).Text)
    If txt����(0).Text <> strTmp Then
        txt����(0).Text = strTmp
        Me.txtƴ��(0).Text = zlStr.GetCodeByORCL(Me.txt����(0).Text, False)
        Me.txt���(0).Text = zlStr.GetCodeByORCL(Me.txt����(0).Text, True)
    End If
    strTmp = MoveSpecialChar(txt����(1).Text)
    If txt����(1).Text <> strTmp Then
        txt����(1).Text = strTmp
        Me.txtƴ��(1).Text = zlStr.GetCodeByORCL(Me.txt����(1).Text, False)
        Me.txt���(1).Text = zlStr.GetCodeByORCL(Me.txt����(1).Text, True)
    End If
    
    'һ�����Լ��
    If Trim(Me.txt����.Text) = "" Then MsgBox "��������룡", vbInformation, gstrSysName: Me.txt����.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > Me.txt����.MaxLength Then MsgBox "����ĳ��������" & Me.txt����.MaxLength & "���ַ�����", vbInformation, gstrSysName: Me.txt����.SetFocus: Exit Sub
    If Trim(Me.txt����(0).Text) = "" Then MsgBox "���������ƣ�", vbInformation, gstrSysName: Me.txt����(0).SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt����(0).Text), vbFromUnicode)) > Me.txt����(0).MaxLength Then
        MsgBox "���Ƴ�����" & Me.txt����(0).MaxLength & "���ַ���" & Me.txt����(0).MaxLength / 2 & "�����֣���", vbInformation, gstrSysName: Me.txt����(0).SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt����(1).Text), vbFromUnicode)) > Me.txt����(1).MaxLength Then
        MsgBox "����������" & Me.txt����(1).MaxLength & "���ַ���" & Me.txt����(1).MaxLength / 2 & "�����֣���", vbInformation, gstrSysName: Me.txt����(1).SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txtƴ��(0).Text), vbFromUnicode)) > Me.txtƴ��(0).MaxLength Then
        MsgBox "����ƴ�����볬����" & Me.txtƴ��(0).MaxLength & "���ַ�����", vbInformation, gstrSysName: Me.txtƴ��(0).SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txtƴ��(1).Text), vbFromUnicode)) > Me.txtƴ��(1).MaxLength Then
        MsgBox "����ƴ�����볬����" & Me.txtƴ��(1).MaxLength & "���ַ�����", vbInformation, gstrSysName: Me.txtƴ��(1).SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt���(0).Text), vbFromUnicode)) > Me.txt���(0).MaxLength Then
        MsgBox "������ʼ��볬����" & Me.txt���(0).MaxLength & "���ַ�����", vbInformation, gstrSysName: Me.txt���(0).SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt���(1).Text), vbFromUnicode)) > Me.txt���(1).MaxLength Then
        MsgBox "������ʼ��볬����" & Me.txt���(1).MaxLength & "���ַ�����", vbInformation, gstrSysName: Me.txt���(1).SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt˵��.Text), vbFromUnicode)) > Me.txt˵��.MaxLength Then
        MsgBox "˵��������" & Me.txt˵��.MaxLength & "���ַ���" & Me.txt˵��.MaxLength / 2 & "�����֣���", vbInformation, gstrSysName: Me.txt˵��.SetFocus: Exit Sub
    End If
    
    '������Ŀʱ����֤�������ظ����룬������ظ��Զ���ԭ��������ϼ�1��ֱ�����ظ�
    str���� = Trim(txt����.Text)
    If Me.Tag = "����" Then
        Do While True
            gstrSql = "select a.���� from ������ĿĿ¼ a,������Ŀ��� b where a.����=[1] and a.���=b.����"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "�����Ƿ��ظ�", str����)
            If rsTemp.RecordCount <> 0 Then
                str���� = zlCommFun.IncStr(str����)
            Else
                Exit Do
            End If
        Loop
    End If
    
    'ʹ�÷�Χ���
    If opt��Χ(0).Value Then
        If cbo��Ա.ListIndex = -1 Then
            MsgBox "��ָ�����׷�����ʹ����Ա��", vbInformation, gstrSysName
            cbo��Ա.SetFocus: Exit Sub
        End If
        lng��ԱID = cbo��Ա.ItemData(cbo��Ա.ListIndex)
    ElseIf opt��Χ(1).Value Then
        For i = 1 To Lvw����.ListItems.Count
            If Lvw����.ListItems(i).Checked Then
                str����IDs = str����IDs & "," & Mid(Lvw����.ListItems(i).Key, 2)
            End If
        Next
        If str����IDs = "" Then
            MsgBox "��ָ�����׷�����ʹ�ÿ��ҡ�", vbInformation, gstrSysName
            Lvw����.SetFocus: Exit Sub
        End If
        str����IDs = Mid(str����IDs, 2)
    End If
    
    '���ݼ��
    If mrsScheme Is Nothing Then
        MsgBox "���׷�����û�����ݣ�����¼����׷������ݣ�", vbInformation, gstrSysName
        cmdScheme.SetFocus: Exit Sub
    ElseIf mrsScheme.RecordCount = 0 Then
        MsgBox "���׷�����û�����ݣ�����¼����׷������ݣ�", vbInformation, gstrSysName
        cmdScheme.SetFocus: Exit Sub
    End If
    
    If cmbStationNo.Text = "" Then
        strվ�� = "Null"
    Else
        strվ�� = Mid(cmbStationNo.Text, 1, InStr(1, cmbStationNo.Text, "-") - 1)
    End If
    
    '���ݱ���
    arrSql = Array()
    If Me.Tag = "����" Then
        lngItemId = zlDatabase.GetNextId("������ĿĿ¼")
    Else
        If zlClinicCodeRepeat(str����, lngItemId) = True Then Exit Sub
    End If
    
    ReDim Preserve arrSql(UBound(arrSql) + 1)
    arrSql(UBound(arrSql)) = "ZL_���׷�����Ŀ_Update(" & _
        lngItemId & "," & Val(Me.txt����.Tag) & ",'" & str���� & "'," & _
        "'" & Trim(Me.txt����(0).Text) & "','" & Trim(Me.txtƴ��(0).Text) & "','" & Trim(Me.txt���(0).Text) & "'," & _
        "'" & Trim(Me.txt����(1).Text) & "','" & Trim(Me.txtƴ��(1).Text) & "','" & Trim(Me.txt���(1).Text) & "'," & _
        "'" & Trim(Me.txt˵��.Text) & "'," & IIf(opt��Χ(0).Value, lng��ԱID, "Null") & "," & _
        IIf(opt��Χ(1).Value, "'" & str����IDs & "'", "Null") & "," & Get������� & "," & IIf(cmbStationNo.Visible = False Or cmbStationNo.Text = "", "Null", strվ��) & ",'" & UserInfo.���� & "'," & chkAll.Value & ")"
    
    If mrsScheme.RecordCount > 0 Then mrsScheme.MoveFirst
    Do While Not mrsScheme.EOF
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = "ZL_���׷�������_Insert(" & _
            lngItemId & "," & mrsScheme!��� & "," & ZVal(Nvl(mrsScheme!������, 0)) & "," & _
            mrsScheme!��Ч & "," & ZVal(Nvl(mrsScheme!������Ŀid, 0)) & "," & _
            IIf(IsNull(mrsScheme!������Ŀid), "'" & Nvl(mrsScheme!ҽ������) & "',", "NULL,") & _
            ZVal(Nvl(mrsScheme!����, 0)) & "," & ZVal(Nvl(mrsScheme!��������, 0)) & "," & ZVal(Nvl(mrsScheme!�ܸ�����, 0)) & "," & _
            ZVal(Nvl(mrsScheme!�շ�ϸĿID, 0)) & ",'" & Nvl(mrsScheme!�걾��λ) & "'," & _
            "'" & Nvl(mrsScheme!ִ��Ƶ��) & "'," & ZVal(Nvl(mrsScheme!Ƶ�ʴ���, 0)) & "," & _
            ZVal(Nvl(mrsScheme!Ƶ�ʼ��, 0)) & ",'" & Nvl(mrsScheme!�����λ) & "'," & _
            "'" & Nvl(mrsScheme!ҽ������) & "'," & Nvl(mrsScheme!ִ������, 0) & "," & _
            ZVal(Nvl(mrsScheme!ִ�п���ID, 0)) & ",'" & Nvl(mrsScheme!ʱ�䷽��) & "'," & _
            "'" & Nvl(mrsScheme!��鷽��) & "'," & ZVal(Val(mrsScheme!�䷽ID & "")) & "," & _
            ZVal(Val(mrsScheme!�����ĿID & "")) & "," & Val(mrsScheme!ִ�б�� & "") & ")"
        mrsScheme.MoveNext
    Loop

    On Error GoTo ErrHand
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    If Me.Tag = "����" Then
        If Val(zlDatabase.GetPara("������Ŀ��������", glngSys, 1054, 0)) = 1 Then
            lngItemId = 0: mblnFirst = True
            Call Form_Activate
            Me.txt����.SetFocus
            Exit Sub
        End If
    End If
    mblnChange = False
    mblnOK = True
    Unload Me
    Exit Sub
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdScheme_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strʹ�ÿ��� As String, i As Long
    
    If Lvw����.Enabled Then
        For i = 1 To Lvw����.ListItems.Count
            If Lvw����.ListItems(i).Checked Then strʹ�ÿ��� = strʹ�ÿ��� & "," & Mid(Lvw����.ListItems(i).Key, 2)
        Next
    End If
    
'    '���Դ���
'    Dim mobjCISKernel As New clsCISKernel
'    Call mobjCISKernel.InitCISKernel(gcnOracle, Me, glngSys, mstrPrivs)
'    Set rsTmp = mobjCISKernel.ShowSchemeEdit(Me, Get�������, mrsScheme, Me.Tag = "����", , Mid(strʹ�ÿ���, 2))

    Call gobjKernel.InitCISKernel(gcnOracle, Me, glngSys, mstrPrivs)
    Set rsTmp = gobjKernel.ShowSchemeEdit(Me, Get�������, mrsScheme, Me.Tag = "����", , Mid(strʹ�ÿ���, 2))
    If Not rsTmp Is Nothing Then
        Set mrsScheme = rsTmp
        mblnChange = True
    End If
End Sub

Private Sub cmd����_Click()
    With Me.tvwClass
        .Left = Me.txt����.Left
        .Top = Me.txt����.Top + Me.txt����.Height
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
End Sub

Private Sub Form_Activate()
    Dim strTemp As String
    Dim bln�޸� As Boolean
    
    
    If Not mblnFirst Then Exit Sub
    mblnFirst = False
    bln�޸� = True
    
    '����ʱ�����ý��治�ɱ༭(OptionButton��Enabledʱֵ��仯)
    '-------------------------------------------------
    If Me.Tag = "����" Then
        Me.cmdOk.Visible = False
        Me.cmdCancel.Caption = "�ر�(&C)"
        Me.txt����.Enabled = False: Me.cmd����.Enabled = False
        Me.txt����.Enabled = False
        Me.txt����(0).Enabled = False: Me.txtƴ��(0).Enabled = False: Me.txt���(0).Enabled = False
        Me.txt����(1).Enabled = False: Me.txtƴ��(1).Enabled = False: Me.txt���(1).Enabled = False
        Me.txt˵��.Enabled = False
        
        opt��Χ(0).Enabled = False: opt��Χ(1).Enabled = False: opt��Χ(2).Enabled = False
        chk��Χ(0).Enabled = False: chk��Χ(1).Enabled = False
        cbo��Ա.Enabled = False: cbo��Ա.BackColor = vbButtonFace
        Lvw����.Enabled = False: Lvw����.BackColor = vbButtonFace
    End If
    
    Err = 0: On Error GoTo ErrHand
    
    '��ȡִ����Ŀ����Ϣ
    '-------------------------------------------------
    If Me.Tag = "����" Then
        
        lngItemId = 0
        Set mrsScheme = Nothing '������������

        If Val(zlDatabase.GetPara(61, glngSys)) = 0 Then '������Ŀ�������ģʽ
            gstrSql = "Select Nvl(Max(����),'0000000') as ���� From ������ĿĿ¼"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
            Me.txt����.Text = Right(String(10, "0") & Val(rsTemp!����) + 1, Len(rsTemp!����))
        Else
            strTemp = Mid(Me.txt����.Text, 2, InStr(1, Me.txt����.Text, "]") - 2)
            gstrSql = "Select Nvl(Max(����),'0000000') as ����" & _
                    " From ������ĿĿ¼" & _
                    " Where ���� like [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, "9" & strTemp & "%")
            
            Err = 0: On Error Resume Next
            Me.txt����.Text = "9" & strTemp & Right(String(10, "0") & Val(rsTemp!����) + 1, Len(rsTemp!����) - 1 - Len(strTemp))
        End If

        Me.txt����(0).Text = "": Me.txt����(1).Text = ""
        Me.txtƴ��(0).Text = "": Me.txtƴ��(1).Text = ""
        Me.txt���(0).Text = "": Me.txt���(1).Text = ""
        Me.txt˵��.Text = ""
        Me.txt������.Text = UserInfo.����
        Me.txt����ʱ��.Text = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    Else
        '��ʾ������Ϣ
        gstrSql = "Select A.����,A.����,A.�걾��λ as ˵��,A.�������,A.��ԱID,B.���,B.����,A.վ��,A.������,A.����ʱ��,a.ִ�з��� From ������ĿĿ¼ A,��Ա�� B Where A.��ԱID=B.ID(+) And A.ID=[1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
        With rsTemp
            Me.txt����.MaxLength = .Fields("����").DefinedSize
            If .RecordCount > 0 Then
                Me.txt����.Text = !����
                Me.txt����(0).Text = !����
                Me.txt˵��.Text = Nvl(!˵��)
                Me.txt������.Text = Nvl(!������)
                Me.txt����ʱ��.Text = IIf(Nvl(!����ʱ��) = "", "", Format(!����ʱ��, "yyyy-mm-dd"))
                SetStationNo IIf(IsNull(!վ��), "", !վ��)
                mblnNoCheck = True
                If Nvl(!�������, 0) = 3 Then
                    chk��Χ(0).Value = 1
                    chk��Χ(1).Value = 1
                ElseIf Nvl(!�������, 0) = 1 Then
                    chk��Χ(0).Value = 1
                    chk��Χ(1).Value = 0
                ElseIf Nvl(!�������, 0) = 2 Then
                    chk��Χ(0).Value = 0
                    chk��Χ(1).Value = 1
                End If
                If Nvl(!ִ�з���, 0) = 1 Then
                    chkAll.Value = 1
                Else
                    chkAll.Value = 0
                End If
                mblnNoCheck = False
                
                If Nvl(!��ԱID, 0) <> 0 Then
                    Me.cbo��Ա.AddItem Nvl(!���) & "-" & Nvl(!����), 0
                    Me.cbo��Ա.ItemData(Me.cbo��Ա.NewIndex) = Nvl(!��ԱID, 0)
                    Me.cbo��Ա.ListIndex = Me.cbo��Ա.NewIndex
                Else
                    Me.cbo��Ա.Text = ""
                    Me.cbo��Ա.ListIndex = -1
                End If
            End If
        End With
        
        '��ʾ����
        gstrSql = "Select ����,����,����,���� From ������Ŀ���� Where ������ĿID=[1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
        With rsTemp
            Do While Not .EOF
                If !���� = 1 And !���� = 1 Then Me.txtƴ��(0).Text = !����
                If !���� = 1 And !���� = 2 Then Me.txt���(0).Text = !����
                If !���� = 9 Then Me.txt����(1).Text = !����
                If !���� = 9 And !���� = 1 Then Me.txtƴ��(1).Text = !����
                If !���� = 9 And !���� = 2 Then Me.txt���(1).Text = !����
                .MoveNext
            Loop
        End With
        
        '��ȡ��������
        Call LoadScheme(lngItemId)
    
        'ȷ����Ŀʹ�÷�Χ
        If cbo��Ա.Text <> "" Then
            opt��Χ(0).Value = True
        Else
            gstrSql = "Select B.ID,B.���� From �������ÿ��� A,���ű� B Where A.����ID=B.ID And A.��ĿID=[1] Order by B.����"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
            If Not rsTemp.EOF Then opt��Χ(1).Value = True
            'ԭ��������Ҫ��������,Ҳ��Ϊ���汣��ѡ��Ļ���
            Do While Not rsTemp.EOF
                Lvw����.ListItems.Add(, "_" & rsTemp!ID, rsTemp!����).Checked = True
                rsTemp.MoveNext
            Loop
        End If
    End If
    
    '����Ȩ�����ÿؼ�������
    '-------------------------------------------------
    If Me.Tag <> "����" Then
        '����Ȩ������ʹ�÷�Χ
        If InStr(mstrPrivs, "ȫԺ���׷���") > 0 Then
            '��ȫԺ���׷���Ȩ��ʱ��������
            Call LoadDeptList(True)
        ElseIf InStr(mstrPrivs, "���Ƴ��׷���") > 0 Then
            'ֻ�б��Ƴ��׷���Ȩ��ʱ�����ڱ����ڻ����ѵ�
            opt��Χ(2).Enabled = False
            If opt��Χ(2).Value Then opt��Χ(1).Value = True
            Call LoadDeptList(False)
        Else
            '��û����ֻ�ܿ����ѵ�
            opt��Χ(1).Enabled = False
            opt��Χ(2).Enabled = False
            If opt��Χ(1).Value Or opt��Χ(2).Value Then opt��Χ(0).Value = True
        End If
        If InStr(mstrPrivs, "ȫԺ���׷���") = 0 Then
            cbo��Ա.Locked = True
            If cbo��Ա.Text = "" Then '�������������޸�ʱѡ��Ϊ����ʹ�ã���ǰ��һ��ѡ���˱���ʹ��
                Me.cbo��Ա.AddItem UserInfo.��� & "-" & UserInfo.����, 0
                Me.cbo��Ա.ItemData(Me.cbo��Ա.NewIndex) = UserInfo.ID
                Me.cbo��Ա.ListIndex = Me.cbo��Ա.NewIndex
            End If
        End If
        
        mblnNoCheck = True
        If mint��Χ = 1 Then
            '�̶�������ʹ��
            chk��Χ(0).Value = 1: chk��Χ(0).Visible = False
            chk��Χ(1).Value = 0: chk��Χ(1).Visible = False
        ElseIf mint��Χ = 2 Then
            '�̶���סԺʹ��
            chk��Χ(0).Value = 0: chk��Χ(0).Visible = False
            chk��Χ(1).Value = 1: chk��Χ(1).Visible = False
        Else
            '��������ֵ,����ѡ��������ʹ��
        End If
        mblnNoCheck = False
    Else
        On Error Resume Next
        cmdCancel.SetFocus
    End If
    
    If Me.Tag = "�޸�" Then
        If opt��Χ(0).Value = True And InStr(mstrPrivs, "�޸ĸ��˳��׷���") < 1 Then
            bln�޸� = False
        ElseIf opt��Χ(1).Value = True And InStr(mstrPrivs, "�޸Ŀ��ҳ��׷���") < 1 Then
            bln�޸� = False
        ElseIf opt��Χ(2).Value = True And InStr(mstrPrivs, "�޸�ȫԺ���׷���") < 1 Then
            bln�޸� = False
        End If
    End If
    
    
     
    '����ʱ�����ý��治�ɱ༭(OptionButton��Enabledʱֵ��仯)
    '-------------------------------------------------
    If Me.Tag = "����" Or bln�޸� = False Then
        Me.cmdOk.Visible = False
        Me.cmdCancel.Caption = "�ر�(&C)"
        Me.txt����.Enabled = False: Me.cmd����.Enabled = False
        Me.txt����.Enabled = False
        Me.txt����(0).Enabled = False: Me.txtƴ��(0).Enabled = False: Me.txt���(0).Enabled = False
        Me.txt����(1).Enabled = False: Me.txtƴ��(1).Enabled = False: Me.txt���(1).Enabled = False
        Me.txt˵��.Enabled = False
        
        opt��Χ(0).Enabled = False: opt��Χ(1).Enabled = False: opt��Χ(2).Enabled = False
        chk��Χ(0).Enabled = False: chk��Χ(1).Enabled = False
        cbo��Ա.Enabled = False: cbo��Ա.BackColor = vbButtonFace
        Lvw����.Enabled = False: Lvw����.BackColor = vbButtonFace
        cmbStationNo.Enabled = False
    End If
    
    mblnChange = False
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadDeptList(ByVal BlnAll As Boolean)
'���ܣ�����Ȩ�޶�ȡ����ʹ�õĿ����б�
'������blnAll=�Ƿ��ȡ���еĿ��ң�����ֻ��ȡ���ѵĿ���
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, strTmp As String
    Dim objItem As ListItem, i As Long
    
    On Error GoTo errH
    
    strTmp = IIf(chk��Χ(0).Value = 1, ",1", "") & IIf(chk��Χ(1).Value = 1, ",2", "") & ",3,"
    If BlnAll Then
        '����ָ����ȫԺ����
        strSql = "Select Distinct A.ID,A.����,A.���� From ���ű� A,��������˵�� B" & _
            " Where A.ID=B.����ID And Instr([1],B.�������)>0" & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is Null)" & _
            " And B.�������� IN('�ٴ�','����','���','����','����','����','Ӫ��')" & _
            " Order by A.����"
    Else
        'ֻ��ָ�����ѵĿ���
        strSql = "Select Distinct A.ID,A.����,A.���� From ���ű� A,��������˵�� B,������Ա C" & _
            " Where A.ID=B.����ID And Instr([1],B.�������)>0 And A.ID=C.����ID And C.��ԱID=[2]" & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is Null)" & _
            " And B.�������� IN('�ٴ�','����','���','����','����','����','Ӫ��')" & _
            " Order by A.����"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strTmp, UserInfo.ID)
    
    strTmp = ""
    For i = 1 To Lvw����.ListItems.Count
        If Lvw����.ListItems(i).Checked Then
            strTmp = strTmp & "," & Mid(Lvw����.ListItems(i).Key, 2)
        End If
    Next
    If strTmp <> "" Then strTmp = strTmp & ","
    Lvw����.ListItems.Clear
    
    i = 0
    Do While Not rsTmp.EOF
        Set objItem = Lvw����.ListItems.Add(, "_" & rsTmp!ID, rsTmp!����)
        If InStr(strTmp, "," & rsTmp!ID & ",") > 0 Then '����ԭ�ȵ�ѡ��
            objItem.Checked = True
            objItem.ForeColor = vbBlue
            If i = 0 Then
                objItem.Selected = True
                objItem.EnsureVisible
            End If
            i = i + 1
        End If
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If Me.tvwClass.Visible Then
            Me.tvwClass.Visible = False: Me.txt����.SetFocus
        Else
            Call cmdCancel_Click
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnFirst = True
    mblnOK = False
    mstrLike = IIf(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
    
    Me.cbo��Ա.AddItem "[ѡ����Ա...]"
    cbo��Ա.Tag = cbo��Ա.ListIndex
    mlngFind = 1
    
    Call GetDefineSize
    Call IniStationNo
End Sub

Private Sub SetStationNo(ByVal strNo As String)
    Dim n As Integer
    
'    If gstrNodeNo = "-" Then Exit Sub
    
    If strNo = "" Then
        cmbStationNo.ListIndex = 0
    Else
        For n = 1 To cmbStationNo.ListCount - 1
            If Mid(cmbStationNo.List(n), 1, InStr(1, cmbStationNo.List(n), "-") - 1) = strNo Then
                cmbStationNo.ListIndex = n
            End If
        Next
    End If
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnChange Then
        If MsgBox("���Ѿ����������˸��ģ�ȷʵҪ���������˳���", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsScheme = Nothing
End Sub

Private Sub IniStationNo()
    Dim dblHeight As Double
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
    
'    lblStationNo.Visible = False
'    cmbStationNo.Visible = False
'
'    If gstrNodeNo <> "-" Then
        lblStationNo.Visible = True
        cmbStationNo.Visible = True
        
        strSql = "select ���,���� from zlnodelist"
        Set rsRecord = zlDatabase.OpenSQLRecord(strSql, "վ���ѯ")
        With cmbStationNo
            .AddItem ""
            Do While Not rsRecord.EOF
                .AddItem rsRecord!��� & "-" & rsRecord!����
                rsRecord.MoveNext
            Loop
        End With
        
'        With cmbStationNo
'            .Clear
'            .AddItem ""
'            .AddItem "0"
'            .AddItem "1"
'            .AddItem "2"
'            .AddItem "3"
'            .AddItem "4"
'            .AddItem "5"
'            .AddItem "6"
'            .AddItem "7"
'            .AddItem "8"
'            .AddItem "9"
'
'            .ListIndex = 0
'        End With
'    Else
'        dblHeight = cmbStationNo.Height
'
'        fraLine(1).Top = fraLine(1).Top - dblHeight
'        fraLine(2).Top = fraLine(2).Top - dblHeight
'        Label1.Top = Label1.Top - dblHeight
'        opt��Χ(0).Top = opt��Χ(0).Top - dblHeight
'        opt��Χ(1).Top = opt��Χ(1).Top - dblHeight
'        opt��Χ(2).Top = opt��Χ(2).Top - dblHeight
'        chk��Χ(0).Top = chk��Χ(0).Top - dblHeight
'        chk��Χ(1).Top = chk��Χ(1).Top - dblHeight
'        lbl��Ա.Top = lbl��Ա.Top - dblHeight
'        cbo��Ա.Top = cbo��Ա.Top - dblHeight
'        lbl����.Top = lbl����.Top - dblHeight
'        lvw����.Top = lvw����.Top - dblHeight
'        cmdHelp.Top = cmdHelp.Top - dblHeight
'        cmdScheme.Top = cmdScheme.Top - dblHeight
'        cmdOK.Top = cmdOK.Top - dblHeight
'        cmdCancel.Top = cmdCancel.Top - dblHeight
'        Me.Height = Me.Height - dblHeight
'    End If
End Sub


Private Sub lvw����_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        Item.ForeColor = vbBlue
    Else
        Item.ForeColor = Lvw����.ForeColor
    End If
    mlngFind = Item.Index + 1
    mblnChange = True
End Sub

Private Sub lvw����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then Call cmdFind_Click
End Sub

Private Sub opt��Χ_Click(Index As Integer)
    If Me.Tag = "����" Then Exit Sub
    
    cbo��Ա.Enabled = Index = 0
    Lvw����.Enabled = Index = 1
    txtFind.Enabled = Lvw����.Enabled
    cmdFind.Enabled = Lvw����.Enabled
    
    If cbo��Ա.Enabled Then
        cbo��Ա.BackColor = vbWindowBackground
    Else
        cbo��Ա.BackColor = vbButtonFace
    End If
    If Lvw����.Enabled Then
        Lvw����.BackColor = vbWindowBackground
        txtFind.BackColor = vbWindowBackground
    Else
        Lvw����.BackColor = vbButtonFace
        txtFind.BackColor = vbButtonFace
    End If
    
    If cbo��Ա.Enabled Then
        If Trim(cbo��Ա.Text) = "" And cbo��Ա.ListCount = 1 Then
            If cbo��Ա.List(0) = "[ѡ����Ա...]" Then
                cbo��Ա.AddItem UserInfo.��� & "-" & UserInfo.����, 0
                cbo��Ա.ItemData(cbo��Ա.NewIndex) = UserInfo.ID
                cbo��Ա.ListIndex = cbo��Ա.NewIndex
                cbo��Ա.Tag = cbo��Ա.ListIndex
            End If
        End If
    End If
    
    mblnChange = True
End Sub

Private Sub tvwClass_DblClick()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Me.txt����.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
    Me.txt����.Text = Me.tvwClass.SelectedItem.Text
    Me.txt����.SetFocus
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
    If Me.cmd���� Is ActiveControl Then Exit Sub
    Me.tvwClass.Visible = False
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txtFind.Text <> "" Then Call cmdFind_Click
End Sub

Private Sub txt����_Change()
    mblnChange = True
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 100
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
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

Private Sub txt����_Change()
    mblnChange = True
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 100
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt����_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub txt����_GotFocus(Index As Integer)
    Me.txt����(Index).SelStart = 0: Me.txt����(Index).SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txt����(Index).Text = MoveSpecialChar(txt����(Index).Text)
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End If
'    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
             
End Sub

Private Sub txt����_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Me.txtƴ��(Index).Text = zlStr.GetCodeByORCL(Me.txt����(Index).Text, False, txtƴ��(Index).MaxLength)
    Me.txt���(Index).Text = zlStr.GetCodeByORCL(Me.txt����(Index).Text, True, txt���(Index).MaxLength)
End Sub

Private Sub txt����_LostFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtƴ��_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub txtƴ��_GotFocus(Index As Integer)
    Me.txtƴ��(Index).SelStart = 0: Me.txtƴ��(Index).SelLength = 100
End Sub

Private Sub txtƴ��_KeyPress(Index As Integer, KeyAscii As Integer)
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

Private Sub cbo��Ա_GotFocus()
    Call zlControl.TxtSelAll(cbo��Ա)
End Sub

Private Sub txt˵��_Change()
    mblnChange = True
End Sub

Private Sub txt˵��_GotFocus()
    Me.txt˵��.SelStart = 0: Me.txt˵��.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt˵��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt˵��_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt���_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub txt���_GotFocus(Index As Integer)
    Me.txt���(Index).SelStart = 0: Me.txt���(Index).SelLength = 100
End Sub

Private Sub txt���_KeyPress(Index As Integer, KeyAscii As Integer)
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

Private Function LoadScheme(ByVal lng����ID As Long) As Boolean
'���ܣ���ȡ����ʾ���ݿ��еĳ��׷�������
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select ���,������,��Ч,������ĿID,�շ�ϸĿID,ҽ������,����,��������,�ܸ�����," & _
        " ҽ������,ִ��Ƶ��,Ƶ�ʴ���,Ƶ�ʼ��,�����λ,ʱ�䷽��,ִ�п���ID,�걾��λ,��鷽��,ִ������,ִ�б��,�䷽ID,�����ĿID" & _
        " From ������Ŀ��� Where �������ID=[1] Order by ���"
    Set mrsScheme = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID)
    LoadScheme = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
