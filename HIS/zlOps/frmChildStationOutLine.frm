VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmChildStationOutLine 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11955
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   11955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3465
      Index           =   0
      Left            =   150
      ScaleHeight     =   3465
      ScaleWidth      =   11265
      TabIndex        =   0
      Top             =   -75
      Width           =   11265
      Begin VB.Frame fra 
         Height          =   3405
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   11655
         Begin VB.PictureBox picConver 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   3315
            ScaleHeight     =   240
            ScaleWidth      =   1755
            TabIndex        =   49
            Top             =   240
            Width           =   1755
         End
         Begin VB.ListBox lst 
            Height          =   900
            Left            =   1140
            Style           =   1  'Checkbox
            TabIndex        =   48
            Top             =   2385
            Width           =   3915
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   7
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   2010
            Width           =   1845
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   4
            Left            =   9315
            TabIndex        =   42
            Top             =   1665
            Width           =   405
         End
         Begin VB.TextBox txt 
            Height          =   870
            Index           =   3
            Left            =   5925
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   45
            Top             =   2385
            Width           =   3570
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   6
            Left            =   5910
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   1650
            Width           =   2310
         End
         Begin VB.CheckBox chk 
            Caption         =   "��̨����"
            Height          =   195
            Index           =   5
            Left            =   8295
            TabIndex        =   41
            Top             =   1725
            Width           =   1020
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   5
            Left            =   9270
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   2025
            Width           =   960
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   4
            Left            =   7470
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   2025
            Width           =   960
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   2
            Left            =   5910
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   2025
            Width           =   960
         End
         Begin VB.CheckBox chk 
            Caption         =   "��Ⱦ����"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   31
            Top             =   2430
            Width           =   1020
         End
         Begin VB.CheckBox chk 
            Caption         =   "��Ⱦ����"
            Height          =   195
            Index           =   3
            Left            =   3240
            TabIndex        =   46
            Top             =   2055
            Width           =   1020
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   2
            Left            =   5910
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   1290
            Width           =   2310
         End
         Begin VB.CommandButton cmd 
            Height          =   330
            Index           =   1
            Left            =   4725
            Picture         =   "frmChildStationOutLine.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "��ѡ����ݼ���F3"
            Top             =   1260
            Width           =   345
         End
         Begin VB.PictureBox picConver 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   3
            Left            =   3300
            ScaleHeight     =   240
            ScaleWidth      =   1755
            TabIndex        =   26
            Top             =   960
            Width           =   1755
         End
         Begin VB.PictureBox picConver 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   2
            Left            =   1200
            ScaleHeight     =   240
            ScaleWidth      =   1770
            TabIndex        =   25
            Top             =   960
            Width           =   1770
         End
         Begin VB.PictureBox picConver 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   5
            Left            =   3285
            ScaleHeight     =   240
            ScaleWidth      =   1755
            TabIndex        =   24
            Top             =   1680
            Width           =   1755
         End
         Begin VB.PictureBox picConver 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   4
            Left            =   1200
            ScaleHeight     =   240
            ScaleWidth      =   1770
            TabIndex        =   23
            Top             =   1680
            Width           =   1770
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   1
            Left            =   1155
            TabIndex        =   6
            Top             =   1290
            Width           =   3570
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   0
            Left            =   5910
            MaxLength       =   10
            TabIndex        =   5
            Top             =   570
            Width           =   2310
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   3
            Left            =   5910
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   915
            Width           =   2310
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   1
            Left            =   5910
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   210
            Width           =   2310
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   0
            ItemData        =   "frmChildStationOutLine.frx":058A
            Left            =   1155
            List            =   "frmChildStationOutLine.frx":058C
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   555
            Width           =   3930
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   0
            Left            =   1155
            TabIndex        =   7
            Top             =   195
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   106692611
            CurrentDate     =   39275
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   1
            Left            =   3090
            TabIndex        =   8
            Top             =   210
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   529
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   106692611
            CurrentDate     =   39275
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   2
            Left            =   1155
            TabIndex        =   9
            Top             =   930
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   106692611
            CurrentDate     =   39275
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   3
            Left            =   3255
            TabIndex        =   10
            Top             =   930
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   106692611
            CurrentDate     =   39275
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   4
            Left            =   1155
            TabIndex        =   11
            Top             =   1650
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   106692611
            CurrentDate     =   39275
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   5
            Left            =   3240
            TabIndex        =   12
            Top             =   1650
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   106692611
            CurrentDate     =   39275
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ʱ��"
            Height          =   195
            Index           =   2
            Left            =   75
            TabIndex        =   29
            Top             =   990
            Width           =   1095
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ʱ��"
            Height          =   195
            Index           =   4
            Left            =   60
            TabIndex        =   30
            Top             =   1695
            Width           =   1095
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            Height          =   180
            Index           =   9
            Left            =   330
            TabIndex        =   43
            Top             =   2055
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "˵    ��"
            Height          =   180
            Index           =   13
            Left            =   5145
            TabIndex        =   40
            Top             =   2385
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����̶�"
            Height          =   180
            Index           =   18
            Left            =   5145
            TabIndex        =   38
            Top             =   1695
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            Height          =   180
            Index           =   17
            Left            =   8535
            TabIndex        =   34
            Top             =   2070
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Ƶ���"
            Height          =   180
            Index           =   8
            Left            =   6915
            TabIndex        =   33
            Top             =   2070
            Width           =   540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�� �� ��"
            Height          =   180
            Index           =   4
            Left            =   5145
            TabIndex        =   32
            Top             =   2070
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            Height          =   180
            Index           =   7
            Left            =   5145
            TabIndex        =   22
            Top             =   990
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            Height          =   180
            Index           =   6
            Left            =   5145
            TabIndex        =   21
            Top             =   1350
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ʽ"
            Height          =   180
            Index           =   5
            Left            =   345
            TabIndex        =   20
            Top             =   1335
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Һ����"
            Height          =   180
            Index           =   3
            Left            =   5145
            TabIndex        =   19
            Top             =   630
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�� �� ��"
            Height          =   180
            Index           =   2
            Left            =   5145
            TabIndex        =   18
            Top             =   270
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������ģ"
            Height          =   180
            Index           =   1
            Left            =   345
            TabIndex        =   17
            Top             =   600
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��"
            Height          =   180
            Index           =   0
            Left            =   345
            TabIndex        =   16
            Top             =   255
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Height          =   180
            Index           =   14
            Left            =   2940
            TabIndex        =   15
            Top             =   270
            Width           =   180
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Height          =   180
            Index           =   15
            Left            =   3015
            TabIndex        =   14
            Top             =   975
            Width           =   180
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Height          =   180
            Index           =   16
            Left            =   3015
            TabIndex        =   13
            Top             =   1695
            Width           =   180
         End
      End
   End
   Begin XtremeSuiteControls.TabControl tbc 
      Height          =   2550
      Left            =   255
      TabIndex        =   47
      Top             =   4845
      Width           =   2820
      _Version        =   589884
      _ExtentX        =   4974
      _ExtentY        =   4498
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmChildStationOutLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################
'���������弶��������
Private mlngLoop As Long
Private mstrSQL As String
Private mblnStartUp As Boolean
Private mlngKey As Long
Private mlngDeptKey As Long
Private mfrmMain As Object
Private mblnDataChanged As Boolean
Private mblnAllowModify As Boolean
Private mblnReading As Boolean

'Private WithEvents mclsVsfDiagBefore As clsVsf
'Private WithEvents mclsVsfDiagAfter As clsVsf

Private WithEvents mfrmChildStationPerson As frmChildStationPerson
Attribute mfrmChildStationPerson.VB_VarHelpID = -1
Private WithEvents mfrmClildStationOps As frmClildStationOps
Attribute mfrmClildStationOps.VB_VarHelpID = -1
Private WithEvents mfrmChildStationDiagnose As frmChildStationDiagnose
Attribute mfrmChildStationDiagnose.VB_VarHelpID = -1

Private Type Items
    ����ʽ As String
End Type

Private usrSaveItem As Items
Public Event AfterDataChanged()
Public Event AfterMakeCharge()

'######################################################################################################################
'�������Զ�����̻���

Public Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
    
    If blnData = False Then
        If Not (mfrmChildStationPerson Is Nothing) Then mfrmChildStationPerson.DataChanged = False
        If Not (mfrmClildStationOps Is Nothing) Then mfrmClildStationOps.DataChanged = False
        If Not (mfrmChildStationDiagnose Is Nothing) Then mfrmChildStationDiagnose.DataChanged = False
    End If
    
    If mblnReading = False Then
        RaiseEvent AfterDataChanged
    End If
End Property

Public Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged Or mfrmChildStationPerson.DataChanged Or mfrmClildStationOps.DataChanged Or mfrmChildStationDiagnose.DataChanged
End Property

Public Function InitData(ByVal frmMain As Object, Optional ByVal blnAllowModify As Boolean = True, Optional ByVal lngDeptKey As Long = 0) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mblnAllowModify = blnAllowModify
    mlngDeptKey = lngDeptKey
    
    Set mfrmMain = frmMain
        
    If ExecuteCommand("��ʼ�ؼ�") = False Then Exit Function
    If ExecuteCommand("��ʼ����") = False Then Exit Function
    
    Call ExecuteCommand("�ؼ�״̬")
    
    DataChanged = False
    If Not (mfrmChildStationPerson Is Nothing) Then mfrmChildStationPerson.DataChanged = False
    If Not (mfrmClildStationOps Is Nothing) Then mfrmClildStationOps.DataChanged = False
    If Not (mfrmChildStationDiagnose Is Nothing) Then mfrmChildStationDiagnose.DataChanged = False
    
    
    InitData = True
    
End Function

Public Function RefreshData(ByVal lngKey As Long, Optional ByVal blnAllowModify As Boolean = True, Optional ByVal lngDeptKey As Long = 0) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mblnAllowModify = blnAllowModify
    mlngDeptKey = lngDeptKey
    mlngKey = lngKey
    
    Call ExecuteCommand("�������")
    Call ExecuteCommand("�ؼ�״̬")
        
    If mlngKey > 0 Then
        If ExecuteCommand("��ȡ����") = False Then Exit Function
    End If

    DataChanged = False
    
    RefreshData = True
    
End Function

Public Function ValidData() As Boolean
    '******************************************************************************************************************
    '���ܣ�У��༭���ݵ���Ч��
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim lngIndex As Long
    
    If dtp(0).Value > dtp(1).Value Then
        tbc.Item(0).Selected = True
        ShowSimpleMsg "������ʼʱ�䲻�ܴ�����������ʱ�䣡"
        Call LocationObj(dtp(0))
        Exit Function
    End If
    
    If IsNull(dtp(1).Value) = False Then
        If Abs(DateDiff("h", CDate(Format(dtp(0).Value, "YYYY-MM-DD HH:MM")), CDate(Format(dtp(1).Value, "YYYY-MM-DD HH:MM")))) > 12 Then
            tbc.Item(0).Selected = True
            ShowSimpleMsg "������ʼʱ�����������ʱ��֮�䲻�ܴ���12Сʱ��"
            Call LocationObj(dtp(0))
            Exit Function
        End If
    End If
    
    If dtp(2).Value > dtp(3).Value And chk(2).Value = 1 Then
        tbc.Item(0).Selected = True
        ShowSimpleMsg "����ʼʱ�䲻�ܴ����������ʱ�䣡"
        Call LocationObj(dtp(2))
        Exit Function
    End If
    
    If chk(2).Value = 1 And Trim(txt(1).Text) = "" Then
        tbc.Item(0).Selected = True
        ShowSimpleMsg "����ָ������ʽ��"
        Call LocationObj(txt(1))
        Exit Function
    End If
    
    If chk(2).Value = 1 And cbo(3).ListIndex = -1 Then
        tbc.Item(0).Selected = True
        ShowSimpleMsg "����ָ������������"
        Call LocationObj(cbo(0))
        Exit Function
    End If
    
    If dtp(4).Value > dtp(5).Value And chk(4).Value = 1 Then
        tbc.Item(0).Selected = True
        ShowSimpleMsg "������ʼʱ�䲻�ܴ�����������ʱ�䣡"
        Call LocationObj(dtp(4))
        Exit Function
    End If
        
    If CheckAllNumber(txt(0).Text) = False Then
        tbc.Item(0).Selected = True
        ShowSimpleMsg "��Һ��������Ϊȫ���֣�"
        Call LocationObj(txt(0))
        Exit Function
    End If
    
    If mfrmChildStationPerson.ValidData = False Then
        tbc.Item(0).Selected = True
        Exit Function
    End If
'
'    '�����������Ƿ��зǷ��ַ�������
'    For lngIndex = 2 To 3
'        With vsf(lngIndex)
'            For lngLoop = 1 To .Rows - 1
'                If Val(.RowData(lngLoop)) > 0 Then
'                    If StrIsValid(.TextMatrix(lngLoop, .ColIndex("�������")), 100) = False Then
'                        tbc.Item(2).Selected = True
'                        Call LocationGrid(vsf(lngIndex), lngLoop, .ColIndex("�������"))
'                        Exit Function
'                    End If
'                End If
'            Next
'        End With
'    Next
    
    ValidData = True
    
End Function

Public Function SaveData(ByRef rsSQL As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim lngOrderKey As Long
    Dim lng����id As Long
    Dim lng��ҳid As Long
    Dim lngRow As Long
    Dim str��Ⱦ���� As String
    Dim lngLoop As Long
    
    On Error GoTo errHand
    
    strSQL = "Select a.* From ����ҽ����¼ a,����������¼ b Where a.ID=b.ҽ��id And b.ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngKey)
    If rs.BOF = True Then Exit Function
    
    lng����id = zlCommFun.NVL(rs("����id").Value, 0)
    lng��ҳid = zlCommFun.NVL(rs("��ҳid").Value, 0)
    lngOrderKey = zlCommFun.NVL(rs("ID").Value, 0)
    
    For lngLoop = 0 To lst.ListCount - 1
        If lst.Selected(lngLoop) Then
            str��Ⱦ���� = str��Ⱦ���� & ";" & lst.List(lngLoop)
        End If
    Next
    If str��Ⱦ���� <> "" Then str��Ⱦ���� = Mid(str��Ⱦ����, 2)
    
    '
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "zl_����������¼_Update(" & mlngKey & "," & _
                                        "To_Date('" & Format(dtp(0).Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                        "To_Date('" & Format(dtp(0).Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                        IIf(IsNull(dtp(1).Value), "Null", "To_Date('" & Format(dtp(1).Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')") & "," & _
                                        IIf(chk(2).Value = 1, "To_Date('" & Format(dtp(2).Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", "NULL") & "," & _
                                        IIf(chk(2).Value = 1, "To_Date('" & Format(dtp(3).Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", "NULL") & "," & _
                                        IIf(chk(2).Value = 1, "'" & txt(1).Text & "'", "Null") & "," & _
                                        IIf(chk(2).Value = 1, Val(cmd(1).Tag), "Null") & "," & _
                                        IIf(chk(2).Value = 1, "'" & txt(2).Text & "'", "Null") & "," & _
                                        IIf(chk(2).Value = 1, "'" & zlCommFun.GetNeedName(cbo(3).Text) & "'", "Null") & "," & _
                                        Val(txt(0).Text) & "," & _
                                        IIf(chk(4).Value = 1, "To_Date('" & Format(dtp(4).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')", "NULL") & "," & _
                                        IIf(chk(4).Value = 1, "To_Date('" & Format(dtp(5).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')", "NULL") & ",'" & _
                                        cbo(1).Text & "'," & _
                                        mlngDeptKey & ",'" & _
                                        zlCommFun.GetNeedName(cbo(0).Text) & "','" & _
                                        zlCommFun.GetNeedName(cbo(6).Text) & "','" & _
                                        zlCommFun.GetNeedName(cbo(2).Text) & "','" & _
                                        zlCommFun.GetNeedName(cbo(4).Text) & "','" & _
                                        zlCommFun.GetNeedName(cbo(5).Text) & "'," & _
                                        Val(txt(4).Text) & ",'" & _
                                        zlCommFun.GetNeedName(cbo(7).Text) & "'," & _
                                        chk(0).Value & ",'" & _
                                        str��Ⱦ���� & "'," & _
                                        chk(3).Value & ",'" & txt(3).Text & "')"
                                        
    Call SQLRecordAdd(rsSQL, strSQL)
            

    If mfrmChildStationPerson.DataChanged Then Call mfrmChildStationPerson.SaveData(rsSQL)
    If mfrmClildStationOps.DataChanged Then Call mfrmClildStationOps.SaveData(rsSQL)
    If mfrmChildStationDiagnose.DataChanged Then Call mfrmChildStationDiagnose.SaveData(rsSQL)
    '
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "Zl_����������¼_Updateadvice(" & mlngKey & ")"
    Call SQLRecordAdd(rsSQL, strSQL)
    
    SaveData = True
    
    Exit Function
    
    '������
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitTabControl()
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    With tbc
        With .PaintManager

            .Appearance = xtpTabAppearanceVisio
            .ClientFrame = xtpTabFrameSingleLine
            .COLOR = xtpTabColorOffice2003
            .ColorSet.ButtonSelected = &HFFC0C0     '&HD2BDB6
            .ColorSet.ButtonNormal = &HFFC0C0       '&HD2BDB6
            .ShowIcons = True
        End With
        Set .Icons = frmPubIcons.imgPublic.Icons
        
        Set mfrmChildStationPerson = New frmChildStationPerson
        Set mfrmClildStationOps = New frmClildStationOps
        Set mfrmChildStationDiagnose = New frmChildStationDiagnose
        
        Call mfrmChildStationPerson.InitData(Me, mblnAllowModify)
        Call mfrmClildStationOps.InitData(Me, mblnAllowModify)
        
        Call mfrmChildStationDiagnose.InitData(Me, mblnAllowModify)
        
        
        .InsertItem 0, "�������", picPane(0).hWnd, 0
        .InsertItem 1, "������Ա", mfrmChildStationPerson.hWnd, 0
        .InsertItem 2, "�������", mfrmClildStationOps.hWnd, 0
        .InsertItem 3, "������", mfrmChildStationDiagnose.hWnd, 0
                
        
        .Item(0).Selected = True
        
    End With
    
End Sub

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim blnAllowModify As Boolean
    Dim strSQL As String
    Dim strTmp As String
    Dim intLoop As Integer
    
    On Error GoTo errHand
    
    mblnReading = True
    
    Select Case strCommand
    '--------------------------------------------------------------------------------------------------------------
    Case "��ʼ�ؼ�"
                

        txt(2).BackColor = COLOR.��ɫ
        
        Call InitTabControl
    '--------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"
                 
        '����������볤��
        '--------------------------------------------------------------------------------------------------------------
        txt(0).MaxLength = 10
        txt(3).MaxLength = 255

        '����������ģ
        '--------------------------------------------------------------------------------------------------------------
        gstrSQL = "SELECT ����||'-'||���� As ����,0 FROM ����������ģ"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If rs.BOF = False Then Call AddComboData(cbo(0), rs, False)
                
        '��������
        '--------------------------------------------------------------------------------------------------------------
        With cbo(3)
            .Clear
            .AddItem "1-��"
            .AddItem "2-��"
            .AddItem "3-��"
            .AddItem "4-Σ(��)"
        End With
    
        '�� �� ��
        '--------------------------------------------------------------------------------------------------------------
        With cbo(2)
            .Clear
            .AddItem "1-����"
            .AddItem "2-��"
            .AddItem "3-��"
            .ListIndex = 0
        End With
        
        '�Ƶ���
        '--------------------------------------------------------------------------------------------------------------
        With cbo(4)
            .Clear
            .AddItem "1-����"
            .AddItem "2-��"
            .AddItem "3-��"
            .ListIndex = 0
        End With
        
        '��������
        '--------------------------------------------------------------------------------------------------------------
        With cbo(5)
            .Clear
            .AddItem "1-����"
            .AddItem "2-��"
            .AddItem "3-��"
            .ListIndex = 0
        End With
        
        '���������̶�
        '--------------------------------------------------------------------------------------------------------------
        With cbo(6)
            .Clear
            .AddItem ""
            strSQL = "Select ����,����,����,ȱʡ��־ From ���������̶�"
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If rs.BOF = False Then
                Do While Not rs.EOF
                    .AddItem rs("����").Value & "-" & rs("����").Value
                    If rs("ȱʡ��־").Value = 1 Then .ListIndex = .NewIndex
                    rs.MoveNext
                Loop
            End If
            If .ListCount > 0 And .ListIndex = -1 Then .ListIndex = 0
        End With
            

        '�������ʷ���
        '--------------------------------------------------------------------------------------------------------------
        With cbo(7)
            .Clear
            .AddItem ""
            strSQL = "Select ����,����,����,ȱʡ��־ From �������ʷ���"
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If rs.BOF = False Then
                Do While Not rs.EOF
                    .AddItem rs("����").Value & "-" & rs("����").Value
                    If rs("ȱʡ��־").Value = 1 Then .ListIndex = .NewIndex
                    rs.MoveNext
                Loop
            End If
            If .ListCount > 0 And .ListIndex = -1 Then .ListIndex = 0
        End With
        
        
        '������Ⱦ����
        '--------------------------------------------------------------------------------------------------------------
        With lst
            .Clear
            strSQL = "Select ����,����,���� From ������Ⱦ����"
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If rs.BOF = False Then
                Do While Not rs.EOF
                    .AddItem rs("����").Value
                    rs.MoveNext
                Loop
            End If
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "�ؼ�״̬"
    
        blnAllowModify = mblnAllowModify
        If mlngKey <= 0 Then blnAllowModify = False
        
        txt(0).Locked = Not blnAllowModify
        txt(1).Locked = Not blnAllowModify
        txt(2).Locked = Not blnAllowModify
        txt(3).Locked = Not blnAllowModify
'        txt(4).Locked = Not blnAllowModify
        cbo(0).Locked = Not blnAllowModify
        cbo(1).Locked = Not blnAllowModify
        cbo(3).Locked = Not blnAllowModify
        cbo(7).Locked = Not blnAllowModify
        
        cbo(2).Locked = Not blnAllowModify
        cbo(4).Locked = Not blnAllowModify
        cbo(5).Locked = Not blnAllowModify
        cbo(6).Locked = Not blnAllowModify
        cmd(1).Enabled = blnAllowModify
        dtp(0).Enabled = blnAllowModify
        dtp(1).Enabled = blnAllowModify
        dtp(2).Enabled = blnAllowModify
        dtp(3).Enabled = blnAllowModify
        dtp(4).Enabled = blnAllowModify
        dtp(5).Enabled = blnAllowModify
        
        chk(2).Enabled = blnAllowModify
        chk(4).Enabled = blnAllowModify
        
        chk(0).Enabled = blnAllowModify
        chk(3).Enabled = blnAllowModify
        chk(5).Enabled = blnAllowModify
        
'        lst.Enabled = blnAllowModify
        
    '------------------------------------------------------------------------------------------------------------------
    Case "�������"
        
        txt(0).Text = ""
        txt(1).Text = ""
        cmd(1).Tag = ""
        txt(2).Text = ""
        txt(3).Text = ""
        txt(4).Text = ""
        chk(2).Value = 0
        chk(4).Value = 0
        chk(0).Value = 0
        chk(3).Value = 0
        chk(5).Value = 0
        
        cbo(2).ListIndex = 0
        cbo(4).ListIndex = 0
        cbo(5).ListIndex = 0
        cbo(6).ListIndex = 0
        cbo(7).ListIndex = 0
        
        lst.Selected(0) = False
        lst.Selected(1) = False
        lst.Selected(2) = False
        lst.Selected(3) = False
        
        mfrmChildStationPerson.ClearData
        mfrmClildStationOps.ClearData
        mfrmChildStationDiagnose.ClearData
        
        DataChanged = False
        If Not (mfrmChildStationPerson Is Nothing) Then mfrmChildStationPerson.DataChanged = False
        If Not (mfrmClildStationOps Is Nothing) Then mfrmClildStationOps.DataChanged = False
        If Not (mfrmChildStationDiagnose Is Nothing) Then mfrmChildStationDiagnose.DataChanged = False
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ����"
        
    
        'ҽ��ִ�з���
        '--------------------------------------------------------------------------------------------------------------
        cbo(1).Clear
        gstrSQL = "SELECT ִ�м�,RowNum As ID FROM ҽ��ִ�з��� WHERE ����id=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngDeptKey)
        If rs.BOF = False Then Call AddComboData(cbo(1), rs)
        
        '1.��ȡ������������
        '--------------------------------------------------------------------------------------------------------------
        gstrSQL = "SELECT A.*,C.�Ա�,C.��ǰ����id,C.סԺ�� FROM ����������¼ A,������Ϣ C WHERE A.����id=C.����id AND A.ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
        If rs.BOF = False Then
            
'            mlng����id = zlCommFun.NVL(rs("����id"), 0)
'            mlng��ҳid = zlCommFun.NVL(rs("��ҳid"), 0)
'            mlngDeptKey = zlCommFun.NVL(rs("��ǰ����id"), 0)
            
'            If zlCommFun.NVL(rs("�Ա�")) Like "*��*" Then mstr�Ա� = mstr�Ա� & ",1"
'            If zlCommFun.NVL(rs("�Ա�")) Like "*Ů*" Then mstr�Ա� = mstr�Ա� & ",2"
            
            If IsNull(rs("������ʼʱ��")) = False Then
                dtp(0).Value = Format(zlCommFun.NVL(rs("������ʼʱ��")), dtp(0).CustomFormat)
                
                If IsNull(rs("��������ʱ��")) = False Then
                    dtp(1).Value = Format(zlCommFun.NVL(rs("��������ʱ��")), dtp(1).CustomFormat)
                    picConver(2).Visible = False
                Else
                    dtp(1).Value = Null
                    picConver(2).Visible = True
                End If
                
                If IsNull(rs("����ʼʱ��")) = False Then
                    chk(2).Value = 1
                    picConver(2).Visible = False
                    picConver(3).Visible = False
                    dtp(2).Value = Format(zlCommFun.NVL(rs("����ʼʱ��")), dtp(2).CustomFormat)
                    dtp(3).Value = Format(zlCommFun.NVL(rs("�������ʱ��")), dtp(3).CustomFormat)
                Else
                    chk(2).Value = 0
                    picConver(2).Visible = True
                    picConver(3).Visible = True
                    dtp(2).Value = Format(zlCommFun.NVL(rs("������ʼʱ��")), dtp(2).CustomFormat)
                    dtp(3).Value = Format(zlCommFun.NVL(rs("������ʼʱ��")) + 1, dtp(3).CustomFormat)
                End If
                
                If IsNull(rs("������ʼʱ��")) = False Then
                    chk(4).Value = 1
                    picConver(4).Visible = False
                    picConver(5).Visible = False
                    dtp(4).Value = Format(zlCommFun.NVL(rs("������ʼʱ��")), dtp(4).CustomFormat)
                    dtp(5).Value = Format(zlCommFun.NVL(rs("��������ʱ��")), dtp(5).CustomFormat)
                Else
                    chk(4).Value = 0
                    picConver(4).Visible = True
                    picConver(5).Visible = True
                    dtp(4).Value = Format(zlCommFun.NVL(rs("������ʼʱ��")), dtp(4).CustomFormat)
                    dtp(5).Value = Format(zlCommFun.NVL(rs("������ʼʱ��")) + 1, dtp(5).CustomFormat)
                End If

            End If
            
            zlControl.CboLocate cbo(0), zlCommFun.NVL(rs("������ģ").Value)
            zlControl.CboLocate cbo(1), zlCommFun.NVL(rs("������").Value)
            zlControl.CboLocate cbo(3), zlCommFun.NVL(rs("��������").Value)
            
            zlControl.CboLocate cbo(6), zlCommFun.NVL(rs("�����̶�").Value)
            
            zlControl.CboLocate cbo(2), zlCommFun.NVL(rs("������").Value)
            zlControl.CboLocate cbo(4), zlCommFun.NVL(rs("�Ƶ���").Value)
            zlControl.CboLocate cbo(5), zlCommFun.NVL(rs("��������").Value)
            
            chk(0).Value = zlCommFun.NVL(rs("��Ⱦ����").Value, 0)
            If chk(0).Value = 1 Then
                strTmp = ";" & zlCommFun.NVL(rs("��Ⱦ����").Value) & ";"
                For intLoop = 0 To lst.ListCount - 1
                    If InStr(strTmp, ";" & lst.List(intLoop) & ";") > 0 Then
                        lst.Selected(intLoop) = True
                    End If
                Next
            End If
            
            Call zlControl.CboLocate(cbo(7), zlCommFun.NVL(rs("��������").Value))
            
            chk(3).Value = zlCommFun.NVL(rs("��Ⱦ����").Value, 0)
            txt(4).Text = zlCommFun.NVL(rs("��̨����").Value)
            chk(5).Value = IIf(Val(txt(4).Text) > 0, 1, 0)
            
            
            txt(0).Text = zlCommFun.NVL(rs("��Һ����").Value)
            txt(1).Text = zlCommFun.NVL(rs("����ʽ").Value)
            cmd(1).Tag = zlCommFun.NVL(rs("����ʽid").Value)
            txt(2).Text = zlCommFun.NVL(rs("��������").Value)
            txt(3).Text = zlCommFun.NVL(rs("˵��").Value)
            
        End If
        
        lst.Enabled = (chk(0).Value = 1)
        txt(4).Enabled = (chk(5).Value = 1)
        
        '������Ա
        '--------------------------------------------------------------------------------------------------------------
        Call mfrmChildStationPerson.RefreshData(mlngKey, mlngDeptKey, mblnAllowModify)
        
        Call mfrmClildStationOps.RefreshData(mlngKey, mblnAllowModify)
        
        Call mfrmChildStationDiagnose.RefreshData(mlngKey, mblnAllowModify)
               
        
    End Select
    
    mblnReading = False
    
    ExecuteCommand = True

    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    mblnReading = False
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Sub cbo_Change(Index As Integer)
    DataChanged = True
End Sub

Private Sub cbo_Click(Index As Integer)
    DataChanged = True
End Sub

Private Sub chk_Click(Index As Integer)

    Select Case Index
    Case 0
        lst.Enabled = (chk(Index).Value = 1)
        On Error Resume Next
        If lst.Enabled Then lst.SetFocus
    Case 2
        picConver(2).Visible = Not (chk(Index).Value = 1)
        picConver(3).Visible = Not (chk(Index).Value = 1)
        
        If cbo(3).Enabled = False Then
            cbo(3).ListIndex = -1
        ElseIf cbo(3).ListIndex = -1 Then
            cbo(3).ListIndex = 0
        End If
    Case 4
        picConver(4).Visible = Not (chk(Index).Value = 1)
        picConver(5).Visible = Not (chk(Index).Value = 1)
    Case 5
        txt(4).Enabled = (chk(Index).Value = 1)
        On Error Resume Next
        If txt(4).Enabled Then txt(4).SetFocus
    End Select
    
    DataChanged = True
    
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    Select Case Index
    '------------------------------------------------------------------------------------------------------------------
    Case 1      '����ʽ
        gstrSQL = GetPublicSQL(SQL.����ʽѡ��)
        
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)
 
        If ShowPubSelect(Me, txt(1), 2, "����,900,0,;����,2400,0,;��������,900,0,", Me.Name & "\����ʽѡ��", "����±���ѡ��һ������ʽ", rsData, rs, 8790, 4500) = 1 Then
            
            cmd(1).Tag = zlCommFun.NVL(rs("ID").Value)
            txt(1).Text = zlCommFun.NVL(rs("����").Value)
            txt(2).Text = zlCommFun.NVL(rs("��������").Value)

            txt(1).Tag = ""

            usrSaveItem.����ʽ = txt(1).Text
            
            DataChanged = True


        End If

    End Select
End Sub


Private Sub dtp_Change(Index As Integer)
    DataChanged = True
End Sub

Private Sub dtp_Click(Index As Integer)
    If Index = 1 Then
        picConver(1).Visible = IsNull(dtp(Index).Value)
    End If
End Sub

Private Sub Form_Load()
    Me.BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    fra(0).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    picPane(0).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
'    picPane(2).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    chk(2).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    chk(4).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    chk(0).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
'    chk(1).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    chk(3).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    chk(5).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    tbc.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload mfrmChildStationPerson
    Unload mfrmClildStationOps
    Unload mfrmChildStationDiagnose
    
End Sub


Private Sub mfrmChildStationPerson_AfterDataChanged()
    Call ExecuteCommand("�ؼ�״̬")
End Sub

Private Sub mfrmChildStationOps_AfterDataChanged()
    Call ExecuteCommand("�ؼ�״̬")
End Sub

Private Sub mfrmClildStationOps_AfterDataChanged()
    RaiseEvent AfterDataChanged
End Sub

Private Sub mfrmClildStationOps_AfterMakeCharge()
    RaiseEvent AfterMakeCharge
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next

    Select Case Index
    Case 0
        fra(0).Move 0, -75, picPane(Index).Width, picPane(Index).Height + 75

        cbo(1).Move cbo(1).Left, cbo(1).Top, fra(0).Width - cbo(1).Left - 75
        cbo(3).Move cbo(3).Left, cbo(3).Top, fra(0).Width - cbo(3).Left - 75
        txt(0).Move txt(0).Left, txt(0).Top, fra(0).Width - txt(0).Left - 75
        txt(2).Move txt(2).Left, txt(2).Top, fra(0).Width - txt(2).Left - 75

        txt(3).Move txt(3).Left, txt(3).Top, fra(0).Width - txt(3).Left - 75, fra(0).Height - txt(3).Top - 75
        
        lst.Move lst.Left, lst.Top, lst.Width, fra(0).Height - lst.Top - 75
        
    End Select
End Sub


Private Sub txt_Change(Index As Integer)
    If mblnReading Then Exit Sub

    DataChanged = True

    Select Case Index
    Case 1
        txt(Index).Tag = "Changed"
    End Select

End Sub

Private Sub txt_GotFocus(Index As Integer)

    zlControl.TxtSelAll txt(Index)

    Select Case Index
    Case 1
        zlCommFun.OpenIme True
    End Select

End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case 1
        If KeyCode = vbKeyDelete Then
            KeyCode = 0
            txt(Index).Text = ""
            txt(2).Text = ""
            cmd(1).Tag = ""
            txt(Index).Tag = ""
            usrSaveItem.����ʽ = ""
        End If
    End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strText As String
    Dim strTmp As String
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim bytMode As Byte

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Select Case Index
        Case 1
            If txt(Index).Tag <> "" Then
                txt(Index).Tag = ""

                strText = UCase(txt(Index).Text)
                bytMode = GetApplyMode(strText)

                strText = strText & "%"
                strTmp = strText & IIf(ParamInfo.��Ŀ����ƥ�䷽ʽ = 1, "", "%")

                gstrSQL = GetPublicSQL(SQL.����ʽ����, bytMode)

                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp)
                If ShowPubSelect(Me, txt(Index), 2, "����,990,0,1;����,1500,0,0;��������,900,0,0", Me.Name & "\����ʽ����", "�������ѡ��һ������ʽ", rsData, rs) = 1 Then
                    
                    cmd(1).Tag = zlCommFun.NVL(rs("ID").Value)
                    txt(Index).Text = zlCommFun.NVL(rs("����").Value)
                    txt(2).Text = zlCommFun.NVL(rs("��������").Value)

                    DataChanged = True

                    usrSaveItem.����ʽ = txt(Index).Text

                Else
                    txt(Index).Text = usrSaveItem.����ʽ
                    txt(Index).Tag = ""
                    Exit Sub
                End If

            End If
        Case Else
            zlCommFun.PressKey vbKeyTab
        End Select
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)

    Select Case Index
    Case 1
        zlCommFun.OpenIme False
    End Select

End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)

    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    If Cancel Then Exit Sub

    Select Case Index
    Case 1
        If (txt(Index).Tag = "Changed") Then
            txt(Index).Text = usrSaveItem.����ʽ
            txt(Index).Tag = ""
        End If
    End Select

End Sub

