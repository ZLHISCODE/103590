VERSION 5.00
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.12#0"; "zlIDKind.ocx"
Begin VB.UserControl usrCardPeople 
   ClientHeight    =   6405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3825
   ScaleHeight     =   6405
   ScaleWidth      =   3825
   Begin VB.PictureBox picFont 
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   2535
      ScaleHeight     =   150
      ScaleWidth      =   210
      TabIndex        =   21
      Top             =   135
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox pic10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   120
      ScaleHeight     =   270
      ScaleWidth      =   3465
      TabIndex        =   15
      Top             =   5640
      Visible         =   0   'False
      Width           =   3495
      Begin VB.Frame frm10 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   1800
         TabIndex        =   20
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox TXT10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   1800
         TabIndex        =   19
         Text            =   "7"
         ToolTipText     =   "��ʾ��ǰҳ������������ָ��ҳ���������س���ת��ָ��ҳ"
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lbl10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "�ܵ�ҳ��"
         Top             =   0
         Width           =   650
      End
      Begin VB.Label lbl12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "��һҳ"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2520
         TabIndex        =   17
         ToolTipText     =   "���������һҳ"
         Top             =   30
         Width           =   705
      End
      Begin VB.Label lbl11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "��һҳ"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   960
         TabIndex        =   16
         ToolTipText     =   "���������һҳ"
         Top             =   0
         Width           =   705
      End
   End
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4425
      Left            =   120
      ScaleHeight     =   4395
      ScaleWidth      =   3465
      TabIndex        =   2
      Top             =   1110
      Width           =   3495
      Begin VB.PictureBox Pic2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1000
         Index           =   0
         Left            =   240
         Picture         =   "usrCardPeople.ctx":0000
         ScaleHeight     =   1030.769
         ScaleMode       =   0  'User
         ScaleWidth      =   4095
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   4095
         Begin VB.PictureBox pic4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   255
            TabIndex        =   5
            Top             =   480
            Width           =   255
            Begin VB.CheckBox chk1 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   6
               Top             =   0
               Width           =   255
            End
         End
         Begin VB.Shape shpRight 
            BorderColor     =   &H8000000D&
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00FF0000&
            FillStyle       =   0  'Solid
            Height          =   495
            Index           =   0
            Left            =   3090
            Top             =   225
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Shape shpLeft 
            BorderColor     =   &H8000000D&
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00FF0000&
            FillStyle       =   0  'Solid
            Height          =   495
            Index           =   0
            Left            =   2625
            Top             =   210
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Shape shpBottom 
            BorderColor     =   &H8000000D&
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00FF0000&
            FillStyle       =   0  'Solid
            Height          =   44
            Index           =   0
            Left            =   2640
            Top             =   660
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Shape shpTop 
            BorderColor     =   &H00FF8080&
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00FF0000&
            FillStyle       =   0  'Solid
            Height          =   44
            Index           =   0
            Left            =   2625
            Top             =   210
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lbl1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   7.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   150
            Index           =   0
            Left            =   840
            TabIndex        =   14
            Top             =   405
            Width           =   150
         End
         Begin VB.Label lbl2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   255
            TabIndex        =   13
            Top             =   120
            Width           =   345
         End
         Begin VB.Label lbl3 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "��"
            Height          =   180
            Index           =   0
            Left            =   270
            TabIndex        =   12
            Top             =   600
            Width           =   195
         End
         Begin VB.Label lbl4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "��"
            Height          =   180
            Index           =   0
            Left            =   720
            TabIndex        =   11
            Top             =   120
            Width           =   180
         End
         Begin VB.Label lbl5 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "��"
            Height          =   180
            Index           =   0
            Left            =   1290
            TabIndex        =   10
            Top             =   120
            Width           =   195
         End
         Begin VB.Label lbl6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "��"
            Height          =   175
            Index           =   0
            Left            =   1995
            TabIndex        =   9
            Top             =   105
            Width           =   180
         End
         Begin VB.Label lbl7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "��"
            Height          =   180
            Index           =   0
            Left            =   720
            TabIndex        =   8
            Top             =   600
            Width           =   180
         End
         Begin VB.Label lbl8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "��"
            Height          =   180
            Index           =   0
            Left            =   1875
            TabIndex        =   7
            Top             =   600
            Width           =   180
         End
         Begin VB.Image ImgCard 
            Height          =   255
            Index           =   0
            Left            =   120
            Stretch         =   -1  'True
            Top             =   0
            Width           =   135
         End
      End
      Begin VB.VScrollBar VS1 
         Height          =   840
         Left            =   100
         TabIndex        =   3
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.PictureBox Pic3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      ScaleHeight     =   585
      ScaleWidth      =   3465
      TabIndex        =   0
      Top             =   360
      Width           =   3495
      Begin VB.CheckBox chkFilter 
         Height          =   375
         Left            =   3000
         Picture         =   "usrCardPeople.ctx":752A
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   375
      End
      Begin zlIDKind.PatiIdentify pi1 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"usrCardPeople.ctx":84A4
         BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindAppearance=   0
         ShowSortName    =   -1  'True
         DefaultCardType =   "���￨"
         IDKindWidth     =   555
         BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AllowAutoCommCard=   -1  'True
         NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
      End
   End
End
Attribute VB_Name = "usrCardPeople"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngModule As Long
Private mlngCount As Long '��ʾ���˵ĸ��������ڲ�ѯ���˻���˲��˺�ı�
Private mblnFilter As Boolean '�ж������Ƿ񾭹��˹��ˣ���������˹��˾���ʾ���˵����ݣ����û�й�������ʾ����ĳ��֣������־������false����Ϊÿ�δ���˶�����������
Private mstrFilterName As String '����������
Private mArr���� '�洢�û��Զ���ĸ����ؼ�Ӧ�ô�����ݵı��⣬Ҳ�൱��һ�ֹ����޶�����������Ϊ����
Private mRsBR As ADODB.Recordset '��ŵ�ǰ������Ҫ��ʾ������
Private mRsAll As ADODB.Recordset '��Ŵ�������е�����
Private mstrReturn As String '���ѡ���ķ���ֵ������ѡ���1~8��label��ֵ��ͨ��"|"�ָ���ɵ��ַ���
Private mrsReturn As ADODB.Recordset
Private mlngSelTab As Long
Private m_CanCheck As Boolean
Private m_def_CanCheck As Boolean
Private mstrLocalID As String
Private mlngLocalIDNum As Long
Private mblnInit As Boolean
Private mstrPIText As String '�����һ��PI1.text�е�����
Private mblnFineseSearch As Boolean
Private mblnNewSearch As Boolean '��ʾ���¿�ʼ��ѯһ����ѯ
Private mstrCardNo As String
Private mlngPatiID As Long
Private mstrFindKey As String
Private mImgList As Object
Private mdblVSϵ�� As Double '���(���п�Ƭ�ؼ��ܺ͵ĸ߶�/10000�����ϵ��,�����п�Ƭ�ؼ��ܺ͵ĸ߶�>10000�������ʹ�á�
'*\CardChanged�¼��������ͬ��ѡ�ʱ����Ӧ���¼��������ڻ�ȡ�ؼ��ķ���ֵmstrReturn�����������ܡ�
Public Event CardChanged() 'ÿ�α��ѡ�ʱ���¼��������ڻ�ȡѡ��ѡ��е�����
Public Event AfterPatiFind(ByVal strIDKindstr As String, ByVal strValue As String, ByVal blnNext As Boolean, blnfind As Boolean)  '���ҵ�IDKindStr���濨Ƭ�ϣ��򷵻��¼��е���������
Public Event CodeFilter(ByVal strCode As String)
Private mbln��ʼ�� As Boolean '��Ƭ�Ƿ��Ѿ�����
Private mlngҳ�� As Long 'ͨ���������ݵĸ�������ȷ��ҳ��
Private mColRs As New Collection '�����¼���е���������50����Ҫʹ�ü��Ͻ���¼���е����ݷֿ���
Private marrFilter '��Ź��˺������
Private mstrOldPiText As String '��žɵĲ�ѯ��������
Private mblnFilterClick As Boolean '�Ƿ������˰�ť
'Public Event GetChecked() '��ť����¼�
'API����ɫ����ת��Ϊ��ɫ
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Sub ShowPeople(Optional ByVal rsBR As ADODB.Recordset, Optional blnSelFree As Boolean = False, Optional blnFilterRefresh As Boolean = False)
    '���ܣ����øÿؼ��ķ������ܹ�δ�ؼ��ṩ��ʼ�Ĺ���������
    '������rsBRҪ��ʾ������Դ������Դ��Ҫ����ID������ֵ�л᷵��ID�ţ�����id��Ϊ�˷����û���ѯ��
    '       blnSelFreeΪfalse��ʾ�ؼ��Զ�������ǰ�����ݶ�λ��true��ؼ������Զ���λ����
    Dim lngi As Long
    Dim lngselnum As Long
    Dim rsbrcopy As ADODB.Recordset
    ReDim marrFilter(0 To 0)
    If mblnInit = False Then Exit Sub 'δ��ʼ��������
    
    mdblVSϵ�� = 1 'ϵ��Ϊ1����ʾ����,���ڹ�������λ
    mstrReturn = ""
    '��ʼ��ҳ���ͳ�ʼ�������ݡ�
    If rsBR Is Nothing Then
        pic2(0).Visible = False
        pi1.Enabled = False
        Exit Sub
    End If
    mblnFilter = False
    Set rsbrcopy = rsBR
'    Set mRsAll = rsBR
    If Not blnFilterRefresh Then Call CopyRecord(rsBR, mRsAll)
    mlngCount = rsbrcopy.RecordCount '�����¼�����ݵĸ���
    mlngҳ�� = Fix(mlngCount / 50) + IIf(mlngCount Mod 50 = 0, 0, 1) '���Ի��ֵ�ҳ��
    
    '�������ȷ��Ҫ��ʾ�Ĳ�������
    If mlngCount <= 50 Then
        Set mRsBR = rsbrcopy
        pic10.Visible = False
    Else
        splitRsToCol rsbrcopy
        Set mRsBR = mColRs("1ҳ")
        mlngCount = mRsBR.RecordCount
        pic10.Visible = True
    End If
    
    lbl10.Caption = "��" & mlngҳ�� & "ҳ"
    TXT10.Text = 1
    
    UserControl_Resize '���resize��Ϊ�˸���ҳ��ȷ���Ƿ���ʾ�������ҳ��Ϣ��
    
    '��ʼ��ҳ��ؼ�����״̬
    Call ExecuteCommand("����ؼ�") '��ʼ���û��ؼ�ʱ��ѡ����
    Call ExecuteCommand("��ʼ�ؼ�")
    
    '********************************************************************
    '���������ҳ����Ҫ�����л��Ļ�������Ĳ������Լ�¼�л�ǰ���ݵ�״̬��
    '����ʵ���ؼ�ֻ��Ҫ�ṩѡ�к���(setCardFocus)��������ҳ����в�����������
    '������Ĳ���������Ϊ�˼�����ǰ�ĳ����������ݱ�����
    '********************************************************************
    mlngLocalIDNum = -1 'id����λ��

    For lngi = 0 To UBound(mArr����) 'idֻ����һ��
        If UCase(mArr����(lngi, 0)) = "ID" Then
            mlngLocalIDNum = (lngi + 1) * 2 - 1
            Exit For
        End If
        If UCase(mArr����(lngi, 1)) = "ID" Then
            mlngLocalIDNum = (lngi + 1) * 2
            Exit For
        End If
    Next
    
    lngselnum = -1
    If mstrLocalID <> "" And mRsBR.RecordCount > 0 And mlngLocalIDNum >= 0 Then '��������ID��mrsbr�����������о�ID��ȡѡ���
        For lngi = 0 To mRsBR.RecordCount - 1
            If mRsBR.Fields("ID").Value = mstrLocalID Then
                lngselnum = lngi
                Exit For
            End If
            mRsBR.MoveNext
        Next
        mRsBR.MoveFirst
    End If

    If mstrLocalID = "" Or lngselnum = -1 Or mlngLocalIDNum < 0 Or mlngCount = 0 Then
        RaiseEvent CardChanged
    Else
        If blnSelFree = False Then
            Call SelectPeopleCard(lngselnum)
        End If
    End If
End Sub

Private Sub splitRsToCol(rs As ADODB.Recordset)
    '���ܣ�����¼���е����ݰ���50��һ���ԭ����飬���ŵ�������
    Dim lngi As Long
    Dim lngj As Long
    Dim lngk As Long
    Dim lngCount As Long
    Dim lngPage As Long
    Dim rsCopy As ADODB.Recordset
    Dim ArrRs()
    ReDim ArrRs(0 To 0)

    If rs Is Nothing Then Exit Sub
    Call CopyRecord(rs, rsCopy)

    lngCount = rsCopy.RecordCount
    Set mColRs = Nothing
    rsCopy.PageSize = 50 '50������һ�黮�ּ�¼��
    lngPage = rsCopy.PageCount '�洢����ҳ��
    
    For lngi = 1 To lngPage '��̬������¼������
        ReDim Preserve ArrRs(0 To UBound(ArrRs) + 1)
        Set ArrRs(UBound(ArrRs)) = New ADODB.Recordset
        Call RsTitelCopy(rsCopy, ArrRs(UBound(ArrRs)))
    Next
    
    rsCopy.MoveFirst
    
    For lngi = 1 To lngPage '��ϼ���
        rsCopy.AbsolutePage = lngi
        For lngj = 1 To rsCopy.PageSize
            If rsCopy.EOF Then Exit For
            
            ArrRs(lngi).AddNew
                
            For lngk = 0 To rsCopy.Fields.Count - 1
                ArrRs(lngi).Fields(lngk).Value = rsCopy.Fields(lngk).Value
            Next
            
            ArrRs(lngi).Update
            
            rsCopy.MoveNext
        Next
        ArrRs(lngi).MoveFirst
        mColRs.Add ArrRs(lngi), lngi & "ҳ"
    Next
    
End Sub

Private Function GetValue(lngnum As Long, Index As Integer) As String
    '���ܣ�����ָ��indexѡ��ϵ�ָ���ؼ��ϵ�����
    Dim lngi As Long
    Select Case lngnum
        Case 1
            GetValue = lbl1(Index).Caption & ""
        Case 2
            GetValue = lbl1(Index).Tag & ""
        Case 3
            GetValue = lbl2(Index).Caption & ""
        Case 4
            GetValue = lbl2(Index).Tag & ""
        Case 5
            GetValue = lbl3(Index).Caption & ""
        Case 6
            GetValue = lbl3(Index).Tag & ""
        Case 7
            GetValue = lbl4(Index).Caption & ""
        Case 8
            GetValue = lbl4(Index).Tag & ""
        Case 9
            GetValue = lbl5(Index).Caption & ""
        Case 10
            GetValue = lbl5(Index).Tag & ""
        Case 11
            GetValue = lbl6(Index).Caption & ""
        Case 12
            GetValue = lbl6(Index).Tag & ""
        Case 13
            GetValue = lbl7(Index).Caption & ""
        Case 14
            GetValue = lbl7(Index).Tag & ""
        Case 15
            GetValue = lbl8(Index).Caption & ""
        Case 16
            GetValue = lbl8(Index).Tag & ""
    End Select
End Function

'�����������������
Public Property Get objPicBack() As PictureBox
    Set objPicBack = pic1
End Property

Public Property Get FScrollBar() As VScrollBar
    Set FScrollBar = VS1
End Property

Public Property Let FindStart(newFindStart As Boolean)
    '���ܣ��û��л�ҳ���Ҫ���³�ʼ����ѯ���������ʹ��һ�����ؼ������в�ѯ����ͨ�õģ�
    '      ���������ҳ���ѯ�������һ��ҳ�棬mblnFineseSearch����ı䣬Ҳ����˵�����Ĭ���Ѳ�ѯ��ϣ�
    '      ��ʱ���޷����ж��˴εĲ�ѯ��ע�����ڲ�ѯ���ֱ仯��FindStart������Ч��ֻ��Ϊ�˼�����ǰ�ĳ������ﱣ��
    mblnNewSearch = newFindStart
    pi1.Text = ""
End Property

Public Property Let locked(blnlocked As Boolean)
    pic1.Enabled = Not blnlocked
    pic3.Enabled = Not blnlocked
    pic10.Enabled = Not blnlocked
End Property

Public Property Get strReturn() As String
    strReturn = mstrReturn
End Property

Public Property Get rsReturn() As ADODB.Recordset
    Set rsReturn = mrsReturn
End Property
Public Property Get CanCheck() As Boolean
    CanCheck = m_CanCheck
    Call pic1_Resize
    Call UserControl_Resize
    Call pic1_Resize
End Property
Public Property Let CanCheck(newCanCheck As Boolean)
    Dim lngi As Long
    m_CanCheck = newCanCheck
    For lngi = 0 To chk1.Count - 1
        chk1(lngi).Value = 0
    Next
    Call pic1_Resize
    Call UserControl_Resize
    Call pic1_Resize
End Property

Public Sub FilterRefreshByCode(rs As Recordset)
    Dim strFilter As String
    
    If rs.State = adStateClosed Then Exit Sub
    If rs.RecordCount = 0 Then Exit Sub
    rs.MoveFirst
    Do While Not rs.EOF
        strFilter = " or id = " & rs!id & strFilter
        rs.MoveNext
    Loop
    mRsAll.Filter = Mid(strFilter, 5)
End Sub

Private Sub FilterRefresh()
    Dim rs As New Recordset
    Dim strPatiID As String
'-------------------------------
    On Error GoTo errH
    
    If chkFilter.Value <> 1 Or pi1.Text = "" Then
        mRsAll.Filter = ""
    Else
        If mstrFilterName = "����" Then
            mRsAll.Filter = "���� like '" & pi1.Text & "%'"
        ElseIf mstrFilterName <> "����" And mstrFilterName <> "Ѫ�����" Then     'And mRsAll.Fields("ID") > 0
            If mlngPatiID <> 0 Then mRsAll.Filter = "����ID = " & mlngPatiID
        ElseIf mstrFilterName = "Ѫ�����" Then
            RaiseEvent CodeFilter(pi1.Text)
        End If
    End If
    mstrOldPiText = mstrFilterName & "-" & pi1.Text
    mblnFilterClick = False
    Call ShowPeople(mRsAll, False, True)
    Exit Sub
errH:
    If Err.Number = 0 Then
        Resume Next
    End If
End Sub

Private Sub chkFilter_Click()
    mblnFilterClick = True
    If chkFilter.Value = 1 And mstrFilterName <> "����" And mstrFilterName <> "Ѫ�����" Then
        pi1.SetFocus
        Call gobjCommFun.PressKey(vbKeyReturn)
    Else
        Call FilterRefresh
    End If
End Sub

Private Sub lbl1_DblClick(Index As Integer)
    chk1(Index).Value = IIf(chk1(Index).Value = 1, 0, 1)
End Sub

Private Sub lbl1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    Set picFont.Font = lbl1(Index).Font
    If lbl1(Index).Width < picFont.TextWidth(lbl1(Index).Caption) Then
        strInfo = lbl1(Index).Caption
    End If
    Call gobjCommFun.ShowTipInfo(pic2(Index).hWnd, strInfo, True)
End Sub

Private Sub lbl11_Click()
    Dim lngTXT As Long

    lngTXT = Val(TXT10.Text) - 1
    If lngTXT > 0 Then
        TXT10.Text = lngTXT
        SetPage lngTXT
    End If
End Sub

Private Sub lbl12_Click()
    Dim lngTXT As Long

    lngTXT = Val(TXT10.Text) + 1
    If lngTXT <= mlngҳ�� Then
        TXT10.Text = lngTXT
        SetPage lngTXT
    End If
End Sub

Private Sub lbl2_DblClick(Index As Integer)
    chk1(Index).Value = IIf(chk1(Index).Value = 1, 0, 1)
End Sub

Private Sub lbl3_DblClick(Index As Integer)
    chk1(Index).Value = IIf(chk1(Index).Value = 1, 0, 1)
End Sub

Private Sub lbl3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    Set picFont.Font = lbl3(Index).Font
    If lbl3(Index).Width < picFont.TextWidth(lbl3(Index).Caption) Then
        strInfo = lbl3(Index).Caption
    End If
    Call gobjCommFun.ShowTipInfo(pic2(Index).hWnd, strInfo, True)
End Sub

Private Sub lbl4_DblClick(Index As Integer)
    chk1(Index).Value = IIf(chk1(Index).Value = 1, 0, 1)
End Sub

Private Sub lbl4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    Set picFont.Font = lbl4(Index).Font
    If lbl4(Index).Width < picFont.TextWidth(lbl4(Index).Caption) Then
        strInfo = lbl4(Index).Caption
    End If
    Call gobjCommFun.ShowTipInfo(pic2(Index).hWnd, strInfo, True)
End Sub

Private Sub lbl5_DblClick(Index As Integer)
    chk1(Index).Value = IIf(chk1(Index).Value = 1, 0, 1)
End Sub

Private Sub lbl5_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    Set picFont.Font = lbl5(Index).Font
    If lbl5(Index).Width < picFont.TextWidth(lbl5(Index).Caption) Then
        strInfo = lbl5(Index).Caption
    End If
    Call gobjCommFun.ShowTipInfo(pic2(Index).hWnd, strInfo, True)
End Sub

Private Sub lbl6_DblClick(Index As Integer)
    chk1(Index).Value = IIf(chk1(Index).Value = 1, 0, 1)
    pic4(Index).Tag = chk1(Index).Value
End Sub

Private Sub lbl6_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    Set picFont.Font = lbl6(Index).Font
    If lbl6(Index).Width < picFont.TextWidth(lbl6(Index).Caption) Then
        strInfo = lbl6(Index).Caption
    End If
    Call gobjCommFun.ShowTipInfo(pic2(Index).hWnd, strInfo, True)
End Sub

Private Sub lbl7_DblClick(Index As Integer)
    chk1(Index).Value = IIf(chk1(Index).Value = 1, 0, 1)
End Sub

Private Sub lbl7_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    Set picFont.Font = lbl7(Index).Font
    If lbl7(Index).Width < picFont.TextWidth(lbl7(Index).Caption) Then
        strInfo = lbl7(Index).Caption
    End If
    Call gobjCommFun.ShowTipInfo(pic2(Index).hWnd, strInfo, True)
End Sub

Private Sub lbl8_DblClick(Index As Integer)
    chk1(Index).Value = IIf(chk1(Index).Value = 1, 0, 1)
End Sub

Private Sub lbl8_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    Set picFont.Font = lbl8(Index).Font
    If lbl8(Index).Width < picFont.TextWidth(lbl8(Index).Caption) Then
        strInfo = lbl8(Index).Caption
    End If
    Call gobjCommFun.ShowTipInfo(pic2(Index).hWnd, strInfo, True)
End Sub

Private Sub SetPage(lngPage As Long)
    '���ݴ���������ص�lngPageҳ
    Dim strcbo As String
    
    TXT10.Text = lngPage
    strcbo = lngPage & "ҳ"
    Set mRsBR = mColRs(strcbo)
    mlngCount = mRsBR.RecordCount
    Call ExecuteCommand("����ؼ�")
    Call ExecuteCommand("��ʼ�ؼ�")
End Sub

Private Sub pi1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0 '"'"������������Ч
End Sub

Private Sub pic10_Resize()
    On Error GoTo Errorhand
    lbl10.Move pic10.ScaleLeft, pic10.ScaleTop + 30
    lbl11.Move (pic10.Width - lbl10.Width - TXT10.Width - lbl12.Width - lbl11.Width) / 2 + lbl10.Left + lbl10.Width, lbl10.Top
    TXT10.Move lbl11.Left + lbl11.Width, pic10.ScaleTop
    frm10.Move TXT10.Left, TXT10.Top + TXT10.Height, TXT10.Width
    lbl12.Move TXT10.Left + TXT10.Width, lbl10.Top
Errorhand:
End Sub

Private Sub pic2_DblClick(Index As Integer)
    chk1(Index).Value = IIf(chk1(Index).Value = 1, 0, 1)
End Sub

Private Sub PI1_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    '���ܣ���ȡ���š���������
    If objHisPati Is Nothing Then
        mlngPatiID = 0
    Else
        mlngPatiID = objHisPati.����id
    End If
    mstrCardNo = ""
    If blnCard = True And Not objCardData Is Nothing Then mstrCardNo = objCardData.���� '����п��ҿ������ݲ�Ϊ���򽫿��Ÿ�ֵ��mstrCardNo
    If mblnFilterClick Or mstrOldPiText <> mstrFilterName & "-" & pi1.Text And chkFilter.Value = 1 Then    '�ϴ��뱾����ͬʱ����Ҫ���ݹ�������ˢ�¼�¼��
        Call FilterRefresh
    End If
    Call FindPatiCard(False, True)
End Sub

Public Sub FindPatiByVbKey(Optional ByVal blnNext As Boolean)
    'blnNext=false ��ʾ��λ���ң�True��ʾ��ʼ���һ������һ��
    If blnNext = True Then '������һ��
        If pi1.Text = "" Then
            If pi1.Enabled And pi1.Visible Then pi1.SetFocus
        Else
            Call FindPatiCard(True)
        End If
    Else  '���ң���λ����ǰ�ؼ�
        If UserControl.ActiveControl Is pi1 Then
            pi1.SetFocus '��ʱ��Ҫ��λһ��
            If pi1.Text <> "" Then
                Call FindPatiCard
            End If
        Else
            pi1.SetFocus
        End If
    End If
End Sub

Private Sub FindPatiCard(Optional ByVal blnNext As Boolean, Optional ByVal blnPati As Boolean = True)
    '���Ҳ���,ͨ����ݼ�ֱ��
    Dim blnfind As Boolean
    If pi1.Enabled = False Then Exit Sub
    If pi1.Text = "" Then mlngPatiID = 0: Exit Sub
    
    If mlngPatiID <> 0 Then blnfind = findIdPeoPle(CStr(mlngPatiID), blnPati, blnNext) 'ͨ��id��ѯ����
    If pi1.Text <> "" And blnfind = False Then blnfind = FindPati(pi1.Text, blnNext) '��������
    If blnfind = False Then
        RaiseEvent AfterPatiFind(mstrFilterName, pi1.Text, blnNext, blnfind)
        If blnfind = False Then
            MsgBox "û���ҵ��������������ݣ�", vbInformation, gstrSysName
        End If
    End If
End Sub

Private Function FindPati(ByVal strValue As String, Optional blnNext As Boolean = True) As Boolean
    '���ܣ����ҷ������������ݲ���λ
    Dim lngi As Long
    Dim lngj As Long
    Dim blnisend As Boolean
    Dim blnTitle As Boolean
    Dim lngFind As Long '�Ӽ�¼�����ҵ������ݵ�λ��
    Dim lngselnum As Long 'lngfind���������ܹ���λ����Ƭ��λ��
    Dim lng����ҳ As Long '��ѯ���Ŀ�Ƭ���ڵ�ҳ��
    Dim rsData As ADODB.Recordset
    Dim bln������ As Boolean, strCardName As String
    
    If strValue = "" Then Exit Function
    strCardName = mstrFilterName
    '�жϲ�ѯ������mArr�������Ƿ����
    For lngi = 0 To UBound(mArr����)
        If mArr����(lngi, 0) = strCardName Or mArr����(lngi, 1) = strCardName Then
            blnTitle = True
            Exit For
        End If
        If mArr����(lngi, 0) = "������" Or mArr����(lngi, 1) = "������" Then
            bln������ = True
        End If
    Next
    If blnTitle = False Then
        If (strCardName = "סԺ��" Or strCardName = "�����") And bln������ = True Then
            strCardName = "������"
        Else
            Exit Function '�����ѯ�����ڹ����в�������ֱ���˳���
        End If
    End If
    
    If Not mRsAll.EOF Then
        mRsAll.MoveFirst
    Else
        Exit Function
    End If
    
    CopyRecord mRsAll, rsData
    
    If rsData Is Nothing Then Exit Function '���mrsbr����û������ʱ��ֱ����������
    
   If blnNext = False Then
        lngselnum = -1

    Else
        If pic10.Visible = True Then
            lngselnum = (TXT10.Text - 1) * 50 + mlngSelTab + 1
        Else
            lngselnum = mlngSelTab + 1
        End If
    End If
    
     '�����ѯ����û�иı䣬����Ҫ���¹������ݣ�ֱ��ʹ���ϴι��˵õ�������
    If mstrOldPiText = strCardName & "-" & strValue And UBound(marrFilter) > 0 Then
        setPosition marrFilter, lngselnum
        FindPati = True
        Exit Function
    End If

    ReDim marrFilter(0 To 0)
    '����
    If strCardName = "����" Then
        rsData.Filter = "���� like '" & strValue & "%'"
    Else
        rsData.Filter = strCardName & IIf(IsNumeric(strValue) = True, "=" & Val(strValue), "='" & strValue & "'")
    End If
    If rsData.RecordCount = 0 Then Exit Function '�޷����˵����ݣ���ֱ���˳�����
    '��λ�����ҵ�������
    For lngi = 0 To rsData.RecordCount - 1
        ReDim Preserve marrFilter(UBound(marrFilter) + 1)
        marrFilter(UBound(marrFilter)) = rsData.Bookmark
        rsData.MoveNext
    Next
    setPosition marrFilter, lngselnum, blnNext
    '�����¼֮ǰ���ҵ���Ϣ�������жϲ�ѯ�����Ƿ�ı�
    mstrOldPiText = strCardName & "-" & strValue
    FindPati = True
End Function

Private Sub setPosition(arr As Variant, lngnum As Long, Optional blnNext As Boolean = True)
    '���ҷ������������ݣ�����λ
    Dim lngi As Long
    Dim lngselnum As Long
    Dim blnisend As Boolean
    Dim lng����ҳ As Long '��ѯ���Ŀ�Ƭ���ڵ�ҳ��
    If UBound(arr) = 0 Then Exit Sub
    
    blnisend = False
    For lngi = 1 To UBound(arr)
        If lngnum < arr(lngi) Then  '�����ǰѡ������֮����ƥ��������λ��ƥ������
            lngselnum = arr(lngi)
            blnisend = True
            Exit For
        End If
    Next
    
    If blnisend = False Then
        If blnNext = True Then
            MsgBox "�������Ҳ������������Ĳ��ˡ�", vbInformation, gstrSysName
            Exit Sub
        Else
            lngselnum = arr(1) '�����ǰѡ������֮����ƥ��������λ����һ��ƥ�����ݵ�λ��
        End If
    End If
    
    If pic10.Visible = True Then '����ж�ҳ���ݣ���Ҫ����ѯ������������λ����һ���Ĵ�����Ҫ�ı䵱ǰҳ
        lng����ҳ = Fix((lngselnum - 1) / 50) + IIf(mRsAll.RecordCount Mod 50 = 0, 0, 1)
        lngselnum = (lngselnum - 1) Mod 50
        If lng����ҳ <> Val(TXT10.Text) Then
            SetPage lng����ҳ
        End If
    Else
        lngselnum = lngselnum - 1
    End If
    
    Call SelectPeopleCard(lngselnum)
End Sub

Public Function findIdPeoPle(strKey As String, Optional ByVal blnPatiID As Boolean = True, Optional blnNext As Boolean = True) As Boolean
'���ܣ���ѯ�����ļ򻯰棬����֪������id������£�ֱ��ͨ������id��ѯ����,FindPati������ͨ������id���в�ѯ��
'lngID������ID��һ���ǿؼ�ͨ����ѯ�ؼ��ڲ�����ʹ�ã���key(��������������������ڶ�λ����)
    Dim rsData As ADODB.Recordset
    Dim arrFilter
    Dim lngselnum As Long
    Dim lngi As Long
    
    CopyRecord mRsAll, rsData
    
    If rsData Is Nothing Then Exit Function '���mrsbr����û������ʱ��ֱ����������
    
    If blnNext = False Then
        lngselnum = -1
    Else
        If pic10.Visible = True Then
            lngselnum = (TXT10.Text - 1) * 50 + mlngSelTab + 1
        Else
            lngselnum = mlngSelTab + 1
        End If
    End If
    
    If blnPatiID = True Then
        rsData.Filter = "����ID=" & Val(strKey)
    Else
        If IsNumeric(strKey) Then
            rsData.Filter = "ID=" & Val(strKey)
        Else
            rsData.Filter = "ID='" & strKey & "'"
        End If
    End If

    If rsData.RecordCount = 0 Then findIdPeoPle = False: Exit Function '�޷����˵����ݣ���ֱ���˳�����
    ReDim arrFilter(0 To 0)
    '��λ�����ҵ�������
    For lngi = 0 To rsData.RecordCount - 1
        ReDim Preserve arrFilter(UBound(arrFilter) + 1)
        arrFilter(UBound(arrFilter)) = rsData.Bookmark
        rsData.MoveNext
    Next
    
    setPosition arrFilter, lngselnum, blnNext
    
    findIdPeoPle = True
End Function

Private Function GetLblValue(Title As String, Index As Integer) As String
    '���ܣ���ȡPic2����Ӧ�ؼ���ֵ,�ڱȽ�ʱ����ʹ��ucase����ĸת��Ϊ��д���Է������������ʧ��
    '������title-����Titleֵ��ѯ���ĸ��ؼ���index-����title��ѯ���Ŀؼ���index
    Dim lngi As Long
    
    If mblnInit = False Then Exit Function
    
    If UBound(mArr����) <> 0 Then
        For lngi = 0 To UBound(mArr����)
            If UCase(mArr����(lngi, 0)) = UCase(Title) Then
                GetLblValue = GetValue((lngi + 1) * 2 - 1, Index)
                Exit For
            End If
            If UCase(mArr����(lngi, 1)) = UCase(Title) Then
                GetLblValue = GetValue((lngi + 1) * 2, Index)
                Exit For
            End If
        Next
    End If
End Function

Public Sub SetPIFocus()
    '������۽���PI1��
    If pi1.Enabled And pi1.Visible Then pi1.SetFocus
End Sub

Private Sub pic2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call gobjCommFun.ShowTipInfo(pic2(Index).hWnd, "")
End Sub

Private Sub TXT10_KeyPress(KeyAscii As Integer)
    Dim strKey As String
    Dim strTXT As String
    strKey = Chr(KeyAscii)
    If Not IsNumeric(strKey) And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then KeyAscii = 0: Exit Sub '�����֡��س����˸����˳�
    If KeyAscii = vbKeyReturn Then
        strTXT = TXT10.Text
        If Val(strTXT) < 1 Or Val(strTXT) > mlngҳ�� Then Exit Sub  '����ȷҳ������
        SetPage Val(strTXT)
    End If
End Sub

Private Sub TXT10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    gobjCommFun.ShowTipInfo TXT10.hWnd, "��ʾ��ǰҳ������������ָ��ҳ���������س���ת��ָ��ҳ"
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    pi1.ActiveFastKey
End Sub

Private Sub UserControl_Terminate()
    Call ExecuteCommand("����ؼ�")
    If Not mColRs Is Nothing Then Set mColRs = Nothing
    If Not mRsBR Is Nothing Then Set mRsBR = Nothing
    If Not mRsAll Is Nothing Then Set mRsAll = Nothing
    If Not mrsReturn Is Nothing Then Set mrsReturn = Nothing
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\" & App.ProductName & "\" & "�����б�_" & mlngModule, "������λ", mstrFilterName)
    mstrLocalID = ""
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '���ܣ���������û�������������
    Call PropBag.WriteProperty("CanCheck", m_CanCheck, m_def_CanCheck)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_CanCheck = PropBag.ReadProperty("CanCheck", m_def_CanCheck)
End Sub

Public Sub UserInit(ByVal frmMain As Object, str���� As String, Optional imgList As Object, Optional ByVal lngModule As Long = 0, Optional ByVal strIDKindstr As String = "")
    '���ܣ���ʼ��ÿ���ؼ������ݺ�������ɫ�����
    '������str�����ɡ�a|b|c|d|e|f|g|h;a|b|c|d|e|f|g|h;.....�����,һ��9�飬����9�����Ժ���ļ��飬
    '               ��9�����ݷֱ����pic2�����ϵ�1~8��label�ؼ����һ��image�ؼ���
    '      a|b|c|d|e|f|g|h����ʾ����|��������|�Ƿ���ʾ|����|�ֺ�|������ɫ|������ɫ|ͼ��;
    '      ��ʾ���ݣ��ַ�����ʾ�ڽ����ϵ�����
    '      �������ݣ��ַ�������ڿؼ���tag�е�����
    '      �Ƿ���ʾ�����֣�0��ʾ��ʾ��Ĭ�ϣ���1��ʾ����ʾ,����imgCard�ؼ���0��ʾ��ʽһȡͼ��1��ʾ��ʽ��ȡͼ
    '      ���壺�ַ������֣��ַ�������Ϊ�ջ���Ϊ0��ʾĬ��,�����е�����Ϊ"����""����"��,��imgCard��Ч
    '      �ֺţ����֣�����Ĵ�С,��imgCard��Ч
    '      ������ɫ�����֣��������ɫ��Ϊ�ջ���0��ʾĬ�Ϻ�ɫ����RGBת��������ֱ�ʾ,��imgCard��Ч
    '      ������ɫ�����֣��ؼ��ı�����ɫ��Ϊ�ջ���0��ʾĬ�Ͽؼ���ɫ����RGBת��������ֱ�ʾ,��imgCard��Ч
    '      ͼ�꣺ֻ��imgCard��Ч������ͼƬ��imgList�еı��
    '            ע����ʽһ��ʾ������Ա��Ƭ��ȡͼ���б�ŵ�ͼƬ����ʽ����ʾÿ����Ա��Ƭ����ȡ�����¼����ͼ���ֶ��б�ŵ�ͼƬ
    '               �������Ϊ��ʽһ������Ա��ͼƬ��һ������ʽ�����Ը��ݲ�ͬ����ı�ͼƬ
    '      imgList�����Ҫ���ͼ�꣬��Ҫ�ṩͼ����Դ��imglist��ͼ�궼�Ǹ���imglist��ȡ�ã�Ŀǰֻ֧��imglist
    Dim ArrS
    Dim ArrM
    Dim lngi As Long
    Dim lngj As Long
    Dim strCardName As String
    
    On Error GoTo ErrHand
    mlngModule = lngModule
    Set mImgList = imgList
    strCardName = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\" & App.ProductName & "\" & "�����б�_" & lngModule, "������λ", "����")
    '��ʼ��Patidentify�ؼ�
    Call CreateSquareCardObject(frmMain, 2200, lngModule)
    '���δ����IDKindStr��ʹ��Ĭ�ϵ�
    If strIDKindstr = "" Then strIDKindstr = "��|����|0|0|0|0|0|0;ס|סԺ��|0|0|0|0|0|0;��|�����|0|0|0|0|0|0;��|���￨|0|0|8|0|0|0;��|�������֤|0|0|0|0|0|0;IC|IC��|1|0|0|0|0|0"
    If Not gobjCardSquare Is Nothing Then
        strIDKindstr = gobjCardSquare.zlGetIDKindStr(strIDKindstr)
    End If
    '���������Nothing,���������壬������ر�ʱ�ᴥ��active�¼���Ӧ���Ƕ��ˢ��ε��ø÷��������⣩
    Call pi1.zlInit(Nothing, 2200, , gcnOracle, gstrDBUser, gobjCardSquare, strIDKindstr)
    pi1.FindPatiShowName = False
    pi1.IDKindIDX = pi1.GetKindIndex(strCardName)
    pi1.AutoSize = True
'    PI1.ShowPropertySet = True
    pi1.objIDKind.AllowAutoICCard = True
    pi1.objIDKind.AllowAutoIDCard = True
    
    ArrS = Split(str����, ";")
    ReDim mArr����(0 To 8, 0 To 7) 'mArr�����Ӧ9���ؼ���ÿ���ؼ���8�����ԣ�����ǹ̶�����
    For lngi = 0 To UBound(ArrS)
        If lngi > 8 Then Exit For
        ArrM = Split(ArrS(lngi), "|")
        For lngj = 0 To UBound(ArrM)
            If lngj > 7 Then Exit For
            mArr����(lngi, lngj) = ArrM(lngj)
        Next
    Next
    '�ڳ�ʼ����������Ⱦ�Ҫ�ı�ؼ������������
    SetLabelProper
    mblnInit = True
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function SetLabelProper()
    Dim objFont As StdFont
    Set objFont = picFont.Font
    '��ʼ��pic2�ϵĿؼ���
    lbl1(0).FontName = IIf(mArr����(0, 3) & "" <> "", mArr����(0, 3), "����") '����
    lbl1(0).FontSize = IIf(Val(mArr����(0, 4) & "") <> 0, Val(mArr����(0, 4)), 9) '�ֺ�
    lbl1(0).ForeColor = IIf(Val(mArr����(0, 5) & "") <> 0, GetRGBFromOLEColor(Val(mArr����(0, 5))), &H0&)        '������ɫ
    lbl1(0).BackColor = IIf(Val(mArr����(0, 6) & "") <> 0, GetRGBFromOLEColor(Val(mArr����(0, 6))), &HFFFFFF) '�ؼ�������ɫ
    Set picFont.Font = lbl1(0).Font
    lbl1(0).Height = picFont.TextHeight("��") '���ø߶�
    
    lbl2(0).FontName = IIf(mArr����(1, 3) & "" <> "", mArr����(1, 3), "����") '����
    lbl2(0).FontSize = IIf(Val(mArr����(1, 4) & "") <> 0, Val(mArr����(1, 4)), 16) '�ֺ�
    lbl2(0).ForeColor = IIf(Val(mArr����(1, 5) & "") <> 0, GetRGBFromOLEColor(Val(mArr����(1, 5))), &H0&)        '������ɫ
    lbl2(0).BackColor = IIf(Val(mArr����(1, 6) & "") <> 0, GetRGBFromOLEColor(Val(mArr����(1, 6))), &HFFFFFF) '�ؼ�������ɫ

    lbl3(0).FontName = IIf(mArr����(2, 3) & "" <> "", mArr����(2, 3), "����") '����
    lbl3(0).FontSize = IIf(Val(mArr����(2, 4) & "") <> 0, Val(mArr����(2, 4)), 9) '�ֺ�
    lbl3(0).ForeColor = IIf(Val(mArr����(2, 5) & "") <> 0, GetRGBFromOLEColor(Val(mArr����(2, 5))), &H0&)        '������ɫ
    lbl3(0).BackColor = IIf(Val(mArr����(2, 6) & "") <> 0, GetRGBFromOLEColor(Val(mArr����(2, 6))), &HFFFFFF) '�ؼ�������ɫ
    Set picFont.Font = lbl3(0).Font
    lbl3(0).Height = picFont.TextHeight("��") '���ø߶�

    lbl4(0).FontName = IIf(mArr����(3, 3) & "" <> "", mArr����(3, 3), "����") '����
    lbl4(0).FontSize = IIf(Val(mArr����(3, 4) & "") <> 0, Val(mArr����(3, 4)), 9) '�ֺ�
    lbl4(0).ForeColor = IIf(Val(mArr����(3, 5) & "") <> 0, GetRGBFromOLEColor(Val(mArr����(3, 5))), &H0&)        '������ɫ
    lbl4(0).BackColor = IIf(Val(mArr����(3, 6) & "") <> 0, GetRGBFromOLEColor(Val(mArr����(3, 6))), &HFFFFFF) '�ؼ�������ɫ
    Set picFont.Font = lbl4(0).Font
    lbl4(0).Height = picFont.TextHeight("��") '���ø߶�
    
    lbl5(0).FontName = IIf(mArr����(4, 3) & "" <> "", mArr����(4, 3), "����") '����
    lbl5(0).FontSize = IIf(Val(mArr����(4, 4) & "") <> 0, Val(mArr����(4, 4)), 9) '�ֺ�
    lbl5(0).ForeColor = IIf(Val(mArr����(4, 5) & "") <> 0, GetRGBFromOLEColor(Val(mArr����(4, 5))), &H0&)        '������ɫ
    lbl5(0).BackColor = IIf(Val(mArr����(4, 6) & "") <> 0, GetRGBFromOLEColor(Val(mArr����(4, 6))), &HFFFFFF) '�ؼ�������ɫ
    Set picFont.Font = lbl5(0).Font
    lbl5(0).Height = picFont.TextHeight("��") '���ø߶�
    
    lbl6(0).FontName = IIf(mArr����(5, 3) & "" <> "", mArr����(5, 3), "����") '����
    lbl6(0).FontSize = IIf(Val(mArr����(5, 4) & "") <> 0, Val(mArr����(5, 4)), 9) '�ֺ�
    lbl6(0).ForeColor = IIf(Val(mArr����(5, 5) & "") <> 0, GetRGBFromOLEColor(Val(mArr����(5, 5))), &H0&)        '������ɫ
    lbl6(0).BackColor = IIf(Val(mArr����(5, 6) & "") <> 0, GetRGBFromOLEColor(Val(mArr����(5, 6))), &HFFFFFF) '�ؼ�������ɫ
    Set picFont.Font = lbl6(0).Font
    lbl6(0).Height = picFont.TextHeight("��") '���ø߶�

    lbl7(0).FontName = IIf(mArr����(6, 3) & "" <> "", mArr����(6, 3), "����") '����
    lbl7(0).FontSize = IIf(Val(mArr����(6, 4) & "") <> 0, Val(mArr����(6, 4)), 9) '�ֺ�
    lbl7(0).ForeColor = IIf(Val(mArr����(6, 5) & "") <> 0, GetRGBFromOLEColor(Val(mArr����(6, 5))), &H0&)        '������ɫ
    lbl7(0).BackColor = IIf(Val(mArr����(6, 6) & "") <> 0, GetRGBFromOLEColor(Val(mArr����(6, 6))), &HFFFFFF) '�ؼ�������ɫ
    Set picFont.Font = lbl7(0).Font
    lbl7(0).Height = picFont.TextHeight("��") '���ø߶�

    lbl8(0).FontName = IIf(mArr����(7, 3) & "" <> "", mArr����(7, 3), "����") '����
    lbl8(0).FontSize = IIf(Val(mArr����(7, 4) & "") <> 0, Val(mArr����(7, 4)), 9) '�ֺ�
    lbl8(0).ForeColor = IIf(Val(mArr����(7, 5) & "") <> 0, GetRGBFromOLEColor(Val(mArr����(7, 5))), &H0&)        '������ɫ
    lbl8(0).BackColor = IIf(Val(mArr����(7, 6) & "") <> 0, GetRGBFromOLEColor(Val(mArr����(7, 6))), &HFFFFFF) '�ؼ�������ɫ
    Set picFont.Font = lbl8(0).Font
    lbl8(0).Height = picFont.TextHeight("��") '���ø߶�
    
    Set picFont.Font = objFont
    Set ImgCard(0).Picture = Nothing
End Function

Public Function GetCheckedData() As ADODB.Recordset
    '���ܣ����ض��ѡ�пؼ������ݣ����᷵��image�е����ݣ�
    Dim lngi As Long
    Dim strName As String 'mArr�����п���Ϊ�գ�����ղ���ͬʱ��Ϊ��Ŀ�����Ե�mArr����Ϊ��ʱ�Զ���һ����Ŀ
    strName = "�Զ���"
    If pic2(0).Visible = False Then Exit Function '���ý���û�д��ݲ���ʱ���޷���ȡ
    If pic2.Count <= 0 Then Exit Function '������������ݣ���ô�����ť����Ч��
    If m_CanCheck = False Then Exit Function '���m_CanCheck = False���޷�get����
    Set mrsReturn = New ADODB.Recordset
    With mrsReturn '��ʼ��rsReturn
        For lngi = 0 To UBound(mArr����)
            .Fields.Append IIf(mArr����(lngi, 0) = "", strName & lngi, mArr����(lngi, 0)), adLongVarChar, 100, adFldIsNullable
            .Fields.Append IIf(mArr����(lngi, 1) = "", strName & lngi, mArr����(lngi, 1)), adLongVarChar, 100, adFldIsNullable
        Next
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        For lngi = 0 To pic2.Count - 1
            If Val(pic4(lngi).Tag) = 1 Then
                .AddNew
                .Fields(0).Value = lbl1(lngi).Caption
                .Fields(1).Value = lbl1(lngi).Tag
                .Fields(2).Value = lbl2(lngi).Caption
                .Fields(3).Value = lbl2(lngi).Tag
                .Fields(4).Value = lbl3(lngi).Caption
                .Fields(5).Value = lbl3(lngi).Tag
                .Fields(6).Value = lbl4(lngi).Caption
                .Fields(7).Value = lbl4(lngi).Tag
                .Fields(8).Value = lbl5(lngi).Caption
                .Fields(9).Value = lbl5(lngi).Tag
                .Fields(10).Value = lbl6(lngi).Caption
                .Fields(11).Value = lbl6(lngi).Tag
                .Fields(12).Value = lbl7(lngi).Caption
                .Fields(13).Value = lbl7(lngi).Tag
                .Fields(14).Value = lbl8(lngi).Caption
                .Fields(15).Value = lbl8(lngi).Tag
                .Update
            End If
        Next
        If .RecordCount > 0 Then
            .MoveFirst
        End If
    End With
    Set GetCheckedData = mrsReturn
'    RaiseEvent GetChecked
End Function

'������Щȫ����Ϊ��ʵ�ֵ��ѡ��ѡ���Ŀ��
Private Sub lbl1_Click(Index As Integer)
    Call SelectPeopleCard(Index, True)
End Sub

Private Sub chk1_Click(Index As Integer)
    pic4(Index).Tag = chk1(Index).Value
End Sub
Private Sub Pic4_Click(Index As Integer)
    chk1(Index).Value = IIf(chk1(Index).Value = 1, 0, 1)
End Sub

Private Sub lbl2_Click(Index As Integer)
    Call SelectPeopleCard(Index, True)
End Sub

Private Sub lbl3_Click(Index As Integer)
    Call SelectPeopleCard(Index, True)
End Sub

Private Sub lbl4_Click(Index As Integer)
    Call SelectPeopleCard(Index, True)
End Sub

Private Sub lbl5_Click(Index As Integer)
    Call SelectPeopleCard(Index, True)
End Sub

Private Sub lbl6_Click(Index As Integer)
    Call SelectPeopleCard(Index, True)
End Sub

Private Sub lbl7_Click(Index As Integer)
    Call SelectPeopleCard(Index, True)
End Sub

Private Sub lbl8_Click(Index As Integer)
    Call SelectPeopleCard(Index, True)
End Sub

Private Sub pi1_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    pi1.ShowPropertySet = True
    If mstrFilterName <> objCard.���� Then pi1.Text = ""
    mstrFilterName = objCard.����
End Sub

Private Sub pic1_Resize()
    Dim lngi As Long
    Dim lngVSMAX As Long
    Dim lngHeight As Long
    Dim blnVshVisible As Boolean
    
    '�ؼ���������ҿؼ���ʾ�Ž���λ�õ���
    If mbln��ʼ�� = False Then Exit Sub

    '���pic2�б仯����ĸ����ؼ��Ĵ�Сλ�õ�
    If mlngCount = 0 Then Exit Sub '�������������Ҫ�Կؼ����е���
    
    On Error Resume Next
    Call LockWindowUpdate(UserControl.hWnd)
    lngHeight = (mlngCount - 1) * 950 + 900 - pic1.ScaleHeight   'pic2(mlngCount - 1).Top - pic2(0).Top + pic2(mlngCount - 1).Height + 50 - pic1.ScaleHeight
    If lngHeight > 10000 Then
        mdblVSϵ�� = lngHeight / 10000
    Else
        mdblVSϵ�� = 1
    End If
    VS1.Left = pic1.Width - VS1.Width
    VS1.Top = pic1.ScaleTop
    VS1.Height = pic1.ScaleHeight
    If lngHeight >= 0 Then
        VS1.Min = 1
        VS1.Max = IIf(lngHeight > 10000, 10000, lngHeight)
        VS1.Value = 1
        If lngHeight > 950 Then
            VS1.SmallChange = 950 / mdblVSϵ��
            VS1.LargeChange = 950 / mdblVSϵ��
        Else
            VS1.SmallChange = lngHeight / mdblVSϵ��
            VS1.LargeChange = lngHeight / mdblVSϵ��
        End If
        blnVshVisible = True
        VS1.Visible = True
    Else
        blnVshVisible = False
        VS1.Visible = False
    End If

    For lngi = 0 To mlngCount - 1
        pic2(lngi).Move 100, 950 * lngi + 50, pi1.Width + chkFilter.Width - IIf(blnVshVisible = True, VS1.Width, 0) + 100, 850
        pic2(lngi).Tag = pic2(lngi).Top  '����ÿ��pic2��ԭʼ����λ�ã�������Ҫ��
        pic2(lngi).AutoRedraw = True
        pic2(lngi).PaintPicture pic2(lngi).Picture, 0, 0, pic2(lngi).ScaleWidth, pic2(lngi).ScaleHeight '���ر���ͼƬ
        
        shpLeft(lngi).Move 0, 0, 45, pic2(lngi).Height
        shpRight(lngi).Move pic2(lngi).Width - 45, 0, 45, pic2(lngi).Height
        shpTop(lngi).Move 0, 0, pic2(lngi).Width, 45
        shpBottom(lngi).Move 0, pic2(lngi).Height - 45, pic2(lngi).Width, 45
        
        If m_CanCheck = True Then '��ʾ��ߵ�ѡ����
            pic4(lngi).Move 60, (pic2(lngi).Height - pic4(lngi).Height) \ 2
            lbl2(lngi).Move pic4(lngi).Left + pic4(lngi).Width + 30, (pic2(lngi).Height - lbl2(lngi).Height) / 2 - 60

            chk1(lngi).Visible = True
            pic4(lngi).Visible = True
        Else
            lbl2(lngi).Move 60, (pic2(lngi).Height - lbl2(lngi).Height) / 2 - 60
            chk1(lngi).Visible = False
            pic4(lngi).Visible = False
        End If
        If Val(mArr����(0, 2)) = 1 Then lbl1(lngi).Height = 50  '���⴦��,��ֻ��ʾ��ɫʱ���������
        lbl1(lngi).Move lbl2(lngi).Left + lbl2(lngi).Width + 45, (pic2(lngi).Height - lbl1(lngi).Height) / 2, pic2(lngi).Width - lbl2(lngi).Left - lbl2(lngi).Width - 45 - 120
        lbl4(lngi).Move lbl1(lngi).Left, 100, (lbl1(lngi).Width - 60) / 3
        lbl5(lngi).Move lbl4(lngi).Left + lbl4(lngi).Width + 30, lbl4(lngi).Top, lbl4(lngi).Width
        lbl6(lngi).Move lbl5(lngi).Left + lbl5(lngi).Width + 30, lbl5(lngi).Top, lbl1(lngi).Width - lbl4(lngi).Width - lbl5(lngi).Width - 60
        
        lbl7(lngi).Move lbl1(lngi).Left, pic2(lngi).Height - 300, (lbl1(lngi).Width - 30) / 2
        lbl8(lngi).Move lbl7(lngi).Left + lbl7(lngi).Width + 30, lbl7(lngi).Top, lbl1(lngi).Width - lbl7(lngi).Width - 30
        lbl3(lngi).Move lbl2(lngi).Left, lbl7(lngi).Top, lbl2(lngi).Width
        ImgCard(lngi).Move 60, 30, 200, 200
        ImgCard(lngi).ZOrder 0
    Next
    If mlngSelTab >= 0 And mlngSelTab < mlngCount Then Call SelectPeopleCard(mlngSelTab)
    Call LockWindowUpdate(0)
    If Err <> 0 Then Err.Clear
End Sub

Public Sub SetCardFocus(strTitle As String, strfind As String)
    '���ܣ������ṩ�ı�������ݲ����������ѡ�����λ
    '������strTitle-Ҫ��λ���ݵ����ͱ��粡��id����ҳid�������ȣ�strFind-��strTitle��Ӧ������123��123��������
    Dim ArrTitle
    Dim ArrFind
    Dim lngi As Integer
    Dim lngj As Integer
    Dim lngCount As Long
    Dim strCopyTitle As String
    Dim rsData As ADODB.Recordset
    Dim strFilter As String
    Dim lng����ҳ As Long
    
    If strTitle = "" Or strfind = "" Then Exit Sub
    
    On Error GoTo ErrHand
    If Not mRsAll.EOF Then
        mRsAll.MoveFirst
    Else
        Exit Sub
    End If
    
    CopyRecord mRsAll, rsData
    
    If rsData Is Nothing Then Exit Sub '���mrsbr����û������ʱ��ֱ����������
    
    lngCount = 0
    ArrTitle = Split(strTitle, "'")
    ArrFind = Split(strfind, "'")
    
    For lngi = 0 To UBound(ArrTitle)
        strFilter = strFilter & ArrTitle(lngi) & "=" & ArrFind(lngi) & " and "
    Next
    strFilter = Left(strFilter, Len(strFilter) - 4)
    
    rsData.Filter = strFilter
    
    If rsData.RecordCount > 0 Then
        lngCount = rsData.Bookmark
        If pic10.Visible = True Then '����ж�ҳ���ݣ���Ҫ����ѯ������������λ����һ���Ĵ�����Ҫ�ı䵱ǰҳ
            lng����ҳ = Fix((lngCount - 1) / 50) + IIf(mRsAll.RecordCount Mod 50 = 0, 0, 1)
            lngCount = (lngCount - 1) Mod 50
            If lng����ҳ <> Val(TXT10.Text) Then
                SetPage lng����ҳ
            End If
        Else
            lngCount = lngCount - 1
        End If
    End If
    Call SelectPeopleCard(lngCount)
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function ExecuteCommand(ParamArray varCmd() As Variant) As Boolean
    Dim intLoop As Integer
    Dim rsSAD As New ADODB.Recordset
    Dim lngi As Long
    Dim lngj As Long
    Dim lngHeight As Long
    On Error GoTo Error
    
    Call SQLRecord(rsSAD)
    
    For intLoop = 0 To UBound(varCmd)
        Select Case varCmd(intLoop)
            Case "��ʼ�ؼ�":
                '����ѡ���λ�ú͸���
                If mlngCount = 0 Then ExecuteCommand = False: Exit Function
                Call LockWindowUpdate(UserControl.hWnd)
                For lngi = 0 To mlngCount - 1
                    If lngi = 0 Then
                        pic2(lngi).Visible = True
                    Else
                        Load pic2(lngi)
                        Load pic4(lngi)
                        Load chk1(lngi)
                        Load lbl1(lngi)
                        Load lbl2(lngi)
                        Load lbl3(lngi)
                        Load lbl4(lngi)
                        Load lbl5(lngi)
                        Load lbl6(lngi)
                        Load lbl7(lngi)
                        Load lbl8(lngi)
                        Load ImgCard(lngi)
                        Load shpLeft(lngi): shpLeft(lngi).Visible = False
                        Load shpRight(lngi): shpRight(lngi).Visible = False
                        Load shpTop(lngi): shpTop(lngi).Visible = False
                        Load shpBottom(lngi): shpBottom(lngi).Visible = False
                        
                         '����ǩ����������
                        Set pic4(lngi).Container = pic2(lngi)
                        Set chk1(lngi).Container = pic4(lngi)
                        Set lbl1(lngi).Container = pic2(lngi)
                        Set lbl2(lngi).Container = pic2(lngi)
                        Set lbl3(lngi).Container = pic2(lngi)
                        Set lbl4(lngi).Container = pic2(lngi)
                        Set lbl5(lngi).Container = pic2(lngi)
                        Set lbl6(lngi).Container = pic2(lngi)
                        Set lbl7(lngi).Container = pic2(lngi)
                        Set lbl8(lngi).Container = pic2(lngi)
                        Set ImgCard(lngi).Container = pic2(lngi)
                        Set ImgCard(lngi).Picture = Nothing
                        Set shpLeft(lngi).Container = pic2(lngi)
                        Set shpRight(lngi).Container = pic2(lngi)
                        Set shpTop(lngi).Container = pic2(lngi)
                        Set shpBottom(lngi).Container = pic2(lngi)
                        
                        pic2(lngi).Visible = True
                        pic4(lngi).Visible = True
                        chk1(lngi).Visible = True
                        lbl1(lngi).Visible = True
                        lbl2(lngi).Visible = True
                        lbl3(lngi).Visible = True
                        lbl4(lngi).Visible = True
                        lbl5(lngi).Visible = True
                        lbl6(lngi).Visible = True
                        lbl7(lngi).Visible = True
                        lbl8(lngi).Visible = True
                        ImgCard(lngi).Visible = True
                    End If
                Next
                Call LockWindowUpdate(0)
                UserControl.Refresh
                For lngi = 0 To mlngCount - 1
                    Call LoadData(lngi, mRsBR, mImgList)
                    UserControl.Refresh
                    mRsBR.MoveNext
                Next
                
                If mRsBR.RecordCount <> 0 Then
                    mRsBR.MoveFirst
                End If
                mbln��ʼ�� = True
                Call pic1_Resize
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "����ؼ�":
                mbln��ʼ�� = False
                VS1.Visible = False
                For lngi = 0 To pic2.Count - 1
                    If lngi = 0 Then
                        pic2(lngi).Visible = False
                        shpLeft(lngi).Visible = False
                        shpRight(lngi).Visible = False
                        shpTop(lngi).Visible = False
                        shpBottom(lngi).Visible = False
                    Else
                        Unload chk1(lngi)
                        Unload pic4(lngi)
                        Unload lbl1(lngi)
                        Unload lbl2(lngi)
                        Unload lbl3(lngi)
                        Unload lbl4(lngi)
                        Unload lbl5(lngi)
                        Unload lbl6(lngi)
                        Unload lbl7(lngi)
                        Unload lbl8(lngi)
                        Unload ImgCard(lngi)
                        Unload shpLeft(lngi)
                        Unload shpRight(lngi)
                        Unload shpTop(lngi)
                        Unload shpBottom(lngi)
                        Unload pic2(lngi)
                    End If
                Next
                UserControl.Refresh
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End Select
    Next
    ExecuteCommand = True
    Exit Function
Error:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    ExecuteCommand = False
End Function

Private Sub pic2_Click(Index As Integer)
    Call SelectPeopleCard(Index, True)
End Sub

Private Sub LoadData(Index As Long, rsData As ADODB.Recordset, Optional imgList As Object)
    '����:�����ݼ��ص�ҳ����,
    '����:rsdata-Ҫ���ص����ݣ�imgList-imgList�ؼ������ͼ��
    '����:
    Dim blnHaveImage As Boolean
    Dim lngj As Long
    Dim blnFunc As Boolean
    Dim lngImgList As Long
    Dim str��ɫ As String, str������ɫ As String '��Ų�ѯ���ı�����ɫת��ΪRGB�������
    Dim str������ As String '���marr�����е�������
    Dim str���� As String '���marr�����еĸ�������
    Dim lng��ɫ As Long, lng������ɫ As Long '��str��ɫ����ʹ��
    If Not imgList Is Nothing Then blnHaveImage = True

    If Val(mArr����(8, 2) & "") = 0 And (mArr����(8, 0) Like "ͼ��" Or mArr����(8, 1) Like "ͼ��") _
        And blnHaveImage Then
        If Val(mArr����(8, 7)) > 0 And Val(mArr����(8, 7)) <= imgList.ListImages.Count Then
        'ͼ����ȡ��ʽһ,ͼƬ��������0��imagelist���������֮��
            ImgCard(Index).Picture = imgList.ListImages(Val(mArr����(8, 7))).Picture
        End If
    End If
    
    For lngj = 0 To UBound(mArr����, 1)
        str��ɫ = getRsValue(mArr����(lngj, 0) & "��ɫ", rsData) '��¼������ ������+��ɫ������ʽ������ʱ����������ȡ��������ӵ���Ӧ�Ŀؼ�
        lng��ɫ = GetRGBFromOLEColor(Val(str��ɫ))
        str������ɫ = getRsValue(mArr����(lngj, 0) & "������ɫ", rsData)
        lng������ɫ = GetRGBFromOLEColor(Val(str������ɫ))
        str������ = getRsValue(mArr����(lngj, 0), rsData)
        str���� = getRsValue(mArr����(lngj, 1), rsData)
        Select Case lngj
            Case 0: 'lbl1
                lbl1(Index).Caption = IIf(Val(mArr����(lngj, 2)) = 0, str������, "")
                lbl1(Index).Tag = str����
                If str��ɫ <> "" Then
                    lbl1(Index).ForeColor = lng��ɫ
                End If
                If str������ɫ <> "" Then
                    lbl1(Index).BackColor = lng������ɫ
                End If
                
            Case 1: 'lbl2
                lbl2(Index).Caption = IIf(Val(mArr����(lngj, 2)) = 0, str������, "")
                lbl2(Index).Tag = str����
                If str��ɫ <> "" Then
                    lbl2(Index).ForeColor = lng��ɫ
                End If
                If str������ɫ <> "" Then
                    lbl2(Index).BackColor = lng������ɫ
                End If
                
            Case 2: 'lbl3
                lbl3(Index).Caption = IIf(Val(mArr����(lngj, 2)) = 0, str������, "")
                lbl3(Index).Tag = str����
                If str��ɫ <> "" Then
                    lbl3(Index).ForeColor = lng��ɫ
                End If
                If str������ɫ <> "" Then
                    lbl3(Index).BackColor = lng������ɫ
                End If
                
            Case 3: 'lbl4
                lbl4(Index).Caption = IIf(Val(mArr����(lngj, 2)) = 0, str������, "")
                lbl4(Index).Tag = str����
                If str��ɫ <> "" Then
                    lbl4(Index).ForeColor = lng��ɫ
                End If
                If str������ɫ <> "" Then
                    lbl4(Index).BackColor = lng������ɫ
                End If
                
            Case 4: 'lbl5
                lbl5(Index).Caption = IIf(Val(mArr����(lngj, 2)) = 0, str������, "")
                lbl5(Index).Tag = str����
                If str��ɫ <> "" Then
                    lbl5(Index).ForeColor = lng��ɫ
                End If
                If str������ɫ <> "" Then
                    lbl5(Index).BackColor = lng������ɫ
                End If
                
            Case 5: 'lbl6
                lbl6(Index).Caption = IIf(Val(mArr����(lngj, 2)) = 0, str������, "")
                lbl6(Index).Tag = str����
                If str��ɫ <> "" Then
                    lbl6(Index).ForeColor = lng��ɫ
                End If
                If str������ɫ <> "" Then
                    lbl6(Index).BackColor = lng������ɫ
                End If
                
            Case 6: 'lbl7
                lbl7(Index).Caption = IIf(Val(mArr����(lngj, 2)) = 0, str������, "")
                lbl7(Index).Tag = str����
                If str��ɫ <> "" Then
                    lbl7(Index).ForeColor = lng��ɫ
                End If
                If str������ɫ <> "" Then
                    lbl7(Index).BackColor = lng������ɫ
                End If
                
            Case 7: 'lbl8
                lbl8(Index).Caption = IIf(Val(mArr����(lngj, 2)) = 0, str������, "")
                lbl8(Index).Tag = str����
                If str��ɫ <> "" Then
                    lbl8(Index).ForeColor = lng��ɫ
                End If
                If str������ɫ <> "" Then
                    lbl8(Index).BackColor = lng������ɫ
                End If
                
            Case 8: 'imgCard����ʽ����ͼƬ��������0��imagelist���������֮��
                'imgCard�Ĵ���ʽ��label�Ĵ���ʽ��ͬ
                If mArr����(lngj, 0) = "ͼ��" And blnHaveImage Then
                    lngImgList = Val(str������)
                    If lngImgList > 0 And lngImgList <= imgList.ListImages.Count Then
                        ImgCard(Index).Picture = imgList.ListImages(lngImgList).Picture
                    End If
                End If
        End Select
    Next
End Sub

Private Function getRsValue(name, rs As ADODB.Recordset) As String
    '��ȡ��Ӧ��Ӧ�е�����
    Dim str As String
    On Error Resume Next
    str = rs.Fields(name).Value & ""
    If Err.Description <> "" Then
        str = ""
        Err.Description = ""
    End If
    getRsValue = str
End Function

Private Sub SelectPeopleCard(ByVal Index As Integer, Optional ByVal blnClick As Boolean = False)
'���ܣ�����ѡ�����״̬����UCE�ؼ�����ʾ
'������Index--��Ƭ������blnClick���Ƿ��ǵ��ѡ��Ƭ(�����ж϶�ε��ͬһ��Ƭˢ�µĴ���)
    Dim lngi As Long
    Dim strRetrun As String
    
    mlngSelTab = Index
    If pic2.Count > 0 Then
        For lngi = 0 To pic2.Count - 1 'Ϊ����ʾ����
            shpLeft(lngi).Visible = IIf(lngi = Index, True, False)
            shpRight(lngi).Visible = IIf(lngi = Index, True, False)
            shpTop(lngi).Visible = IIf(lngi = Index, True, False)
            shpBottom(lngi).Visible = IIf(lngi = Index, True, False)
            If lngi = Index Then
               shpLeft(lngi).ZOrder 0
               shpRight(lngi).ZOrder 0
               shpTop(lngi).ZOrder 0
               shpBottom(lngi).ZOrder 0
               If pic2(lngi).Visible And pic2(lngi).Enabled Then pic2(lngi).SetFocus
            End If
        Next
    End If
    
     '���й�����ʱ����ѡ�е�������pic1δ��ʾ���֣���ôҪ�ƶ��������ѷ����û��鿴
    If VS1.Visible = True And pic2(Index).Top + pic2(Index).Height > pic1.ScaleHeight Then '���ؼ�λ����ʾ��������ʱ
        VS1.Value = (Abs(pic2(Index).Top - pic2(0).Top - pic1.ScaleHeight + pic2(Index).Height + 50)) / mdblVSϵ��
    ElseIf VS1.Visible = True And pic2(Index).Top < 0 Then  '���ؼ�λ����ʾ��������ʱ
        VS1.Value = (Abs(pic2(Index).Top - pic2(0).Top + 50)) / mdblVSϵ�� + 1
    End If
    '����һ�����˵����ݣ����᷵��image�е�����
    strRetrun = lbl1(Index).Caption & "'" & lbl1(Index).Tag & "'" & lbl2(Index).Caption & "'" & lbl2(Index).Tag & "'" & lbl3(Index).Caption & "'" & lbl3(Index).Tag & "'" & lbl4(Index).Caption & "'" & lbl4(Index).Tag & "'" & _
                 lbl5(Index).Caption & "'" & lbl5(Index).Tag & "'" & lbl6(Index).Caption & "'" & lbl6(Index).Tag & "'" & lbl7(Index).Caption & "'" & lbl7(Index).Tag & "'" & lbl8(Index).Caption & "'" & lbl8(Index).Tag

    If mlngLocalIDNum >= 0 Then '��ȡ��ѡ���ID
        mstrLocalID = GetValue(mlngLocalIDNum, Index)
    End If
    If mstrReturn = strRetrun And blnClick = True Then Exit Sub
    mstrReturn = strRetrun
    RaiseEvent CardChanged
End Sub

Private Sub UserControl_Resize()
    '�ؼ��ı��Сpic1Ҳ��ı��С
    On Error GoTo Errorhand
    pic3.Left = UserControl.ScaleLeft
    pic3.Top = UserControl.ScaleTop
    pic3.Width = UserControl.ScaleWidth
    pic3.Height = 575
    
    '����pi1��λ��
    pi1.Left = pic3.ScaleLeft + 100
    pi1.Top = pic3.ScaleTop + 75
    pi1.Width = pic3.ScaleWidth - 250 - chkFilter.Width - 50
    '�������˰�ťλ��
    chkFilter.Move pi1.Left + pi1.Width + 50, pi1.Top
    
    If mlngҳ�� = 1 Or mlngҳ�� = 0 Then
        pic1.Left = UserControl.ScaleLeft
        pic1.Top = UserControl.ScaleTop + pic3.Height
        pic1.Width = UserControl.ScaleWidth
        If UserControl.ScaleHeight > pic3.Height Then
            pic1.Height = UserControl.ScaleHeight - pic3.Height
        End If
        pic10.Visible = False
'        pic2(0).Enabled = True
    Else
        pic1.Left = UserControl.ScaleLeft
        pic1.Top = UserControl.ScaleTop + pic3.Height
        pic1.Width = UserControl.ScaleWidth
        If UserControl.ScaleHeight > pic3.Height + pic10.Height Then
            pic1.Height = UserControl.ScaleHeight - pic3.Height - pic10.Height
        End If
        
        pic10.Move UserControl.ScaleLeft, pic1.Top + pic1.Height, pic1.Width
        pic10.Visible = True
    End If
Errorhand:
End Sub

Private Sub VS1_Change()
    VS1_Scroll
End Sub

Private Sub VS1_Scroll()
    '�������֣�picturebox�ƶ�
    Dim lngi As Long
    If pic2.Count > 0 Then
        Call LockWindowUpdate(UserControl.hWnd)
        For lngi = 0 To pic2.Count - 1
            pic2(lngi).Top = Val(pic2(lngi).Tag) - VS1.Value * mdblVSϵ��
        Next
        Call LockWindowUpdate(0)
    End If
End Sub

Private Function GetRGBFromOLEColor(ByVal dwOleColour As Long) As Long
    '��VB����ɫת��ΪRGB��ʾ
    Dim clrref As Long
    Dim r As Long, g As Long, b As Long
    
    OleTranslateColor dwOleColour, 0, clrref
    
    b = (clrref \ 65536) And &HFF
    g = (clrref \ 256) And &HFF
    r = clrref And &HFF
    
    GetRGBFromOLEColor = RGB(r, g, b)
End Function


Private Sub RsTitelCopy(ByVal RsProm As ADODB.Recordset, ToRs)
    '���ܣ��½�ToRs��¼������RsProm�Ľṹ���Ƶ�ToRs��
    '������RsProm-ԭ��¼����ToRs-�½��ļ�¼������Ϊ�г�����Ҫ���붯̬�����ļ�¼������Щ��¼�����������У�����tors������һ���Ǽ�¼������
    Dim lngi As Long
    Set ToRs = New ADODB.Recordset
    With ToRs '��ʼ��rsReturn
        For lngi = 0 To RsProm.Fields.Count - 1
            .Fields.Append RsProm.Fields(lngi).name, adLongVarChar, 100, adFldIsNullable
        Next
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub CopyRecord(ByVal RsProm As ADODB.Recordset, ToRs As ADODB.Recordset)
    '���ܣ�����¼��RsProm�Ľṹ�������ݶ����Ƹ�ToRs
    '������RsProm-Ҫ��ֵ�ļ�¼����ToRs-Ŀ���¼��
    Dim lngi As Long
    Dim lngj As Long
    Call RsTitelCopy(RsProm, ToRs)
    With ToRs
        If RsProm.RecordCount > 0 Then '��ǰû�ж�rsbr���������жϻᱨ��
            For lngi = 0 To RsProm.RecordCount - 1
                .AddNew
                For lngj = 0 To RsProm.Fields.Count - 1
                    .Fields(lngj).Value = RsProm.Fields(lngj).Value
                Next
                .Update
                RsProm.MoveNext
            Next
            RsProm.MoveFirst
            If .RecordCount > 0 Then
                .MoveFirst
            End If
        End If
    End With
End Sub

