VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCISBorrowEdit 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   7455
      Index           =   0
      Left            =   195
      ScaleHeight     =   7455
      ScaleWidth      =   9945
      TabIndex        =   2
      Top             =   75
      Width           =   9945
      Begin VB.Frame fra 
         Height          =   7005
         Left            =   60
         TabIndex        =   26
         Top             =   45
         Width           =   9225
         Begin VB.TextBox txtBorrowUser 
            Height          =   300
            Left            =   1065
            ScrollBars      =   2  'Vertical
            TabIndex        =   46
            Top             =   825
            Width           =   4500
         End
         Begin VB.CommandButton cmd 
            Height          =   300
            Index           =   3
            Left            =   5580
            Picture         =   "frmCISBorrowEdit.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   1515
            Width           =   315
         End
         Begin VB.CommandButton cmd 
            Height          =   300
            Index           =   2
            Left            =   5580
            Picture         =   "frmCISBorrowEdit.frx":6852
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   1185
            Width           =   315
         End
         Begin VB.CommandButton cmd 
            Height          =   300
            Index           =   0
            Left            =   5580
            Picture         =   "frmCISBorrowEdit.frx":D0A4
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   810
            Width           =   315
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   0
            Left            =   4740
            TabIndex        =   36
            Top             =   480
            Width           =   1170
         End
         Begin VB.TextBox txtConver 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Index           =   0
            Left            =   1095
            TabIndex        =   32
            Top             =   510
            Width           =   1260
         End
         Begin VB.TextBox txtConver 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Index           =   1
            Left            =   2625
            TabIndex        =   31
            Top             =   510
            Width           =   1260
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   5
            Left            =   510
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            Top             =   150
            Width           =   1155
         End
         Begin VB.PictureBox picPane 
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   1410
            Index           =   1
            Left            =   315
            ScaleHeight     =   1410
            ScaleWidth      =   8700
            TabIndex        =   27
            Top             =   5100
            Width           =   8700
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   9
               Left            =   2355
               ScrollBars      =   2  'Vertical
               TabIndex        =   43
               Top             =   1065
               Width           =   1530
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   1
               Left            =   615
               ScrollBars      =   2  'Vertical
               TabIndex        =   42
               Top             =   1065
               Width           =   915
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   8
               Left            =   615
               ScrollBars      =   2  'Vertical
               TabIndex        =   21
               Top             =   720
               Width           =   915
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   7
               Left            =   615
               ScrollBars      =   2  'Vertical
               TabIndex        =   13
               Top             =   375
               Width           =   915
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   6
               Left            =   630
               ScrollBars      =   2  'Vertical
               TabIndex        =   9
               Top             =   30
               Width           =   915
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   0
               Left            =   4665
               ScrollBars      =   2  'Vertical
               TabIndex        =   25
               Top             =   720
               Width           =   3075
            End
            Begin VB.TextBox txtConver 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   240
               Index           =   2
               Left            =   4695
               TabIndex        =   29
               Top             =   420
               Width           =   1245
            End
            Begin VB.TextBox txtConver 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   240
               Index           =   3
               Left            =   6255
               TabIndex        =   28
               Top             =   420
               Width           =   1455
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   2
               Left            =   2355
               ScrollBars      =   2  'Vertical
               TabIndex        =   11
               Top             =   45
               Width           =   1530
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   3
               Left            =   2355
               ScrollBars      =   2  'Vertical
               TabIndex        =   15
               Top             =   390
               Width           =   1530
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   4
               Left            =   2355
               ScrollBars      =   2  'Vertical
               TabIndex        =   23
               Top             =   720
               Width           =   1530
            End
            Begin MSComCtl2.DTPicker dtp 
               Height          =   300
               Index           =   2
               Left            =   4665
               TabIndex        =   17
               Top             =   390
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   91684867
               CurrentDate     =   39500
            End
            Begin MSComCtl2.DTPicker dtp 
               Height          =   300
               Index           =   3
               Left            =   6210
               TabIndex        =   19
               Top             =   390
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   91684867
               CurrentDate     =   39500
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�黹ʱ��"
               Height          =   180
               Index           =   19
               Left            =   1605
               TabIndex        =   45
               Top             =   1110
               Width           =   720
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�黹��"
               Height          =   180
               Index           =   18
               Left            =   45
               TabIndex        =   44
               Top             =   1125
               Width           =   540
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������"
               Height          =   180
               Index           =   0
               Left            =   30
               TabIndex        =   8
               Top             =   75
               Width           =   540
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�Ǽ�ʱ��"
               Height          =   180
               Index           =   1
               Left            =   1605
               TabIndex        =   10
               Top             =   90
               Width           =   720
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��׼��"
               Height          =   180
               Index           =   6
               Left            =   45
               TabIndex        =   12
               Top             =   450
               Width           =   540
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��׼ʱ��"
               Height          =   180
               Index           =   7
               Left            =   1605
               TabIndex        =   14
               Top             =   435
               Width           =   720
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�ܽ���"
               Height          =   180
               Index           =   8
               Left            =   45
               TabIndex        =   20
               Top             =   780
               Width           =   540
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�ܽ�ʱ��"
               Height          =   180
               Index           =   9
               Left            =   1605
               TabIndex        =   22
               Top             =   765
               Width           =   720
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�ܽ�����"
               Height          =   180
               Index           =   10
               Left            =   3930
               TabIndex        =   24
               Top             =   780
               Width           =   720
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����ʱ��"
               Height          =   180
               Index           =   11
               Left            =   3930
               TabIndex        =   16
               Top             =   450
               Width           =   720
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��"
               Height          =   180
               Index           =   12
               Left            =   6015
               TabIndex        =   18
               Top             =   435
               Width           =   330
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsf 
            Height          =   1725
            Index           =   1
            Left            =   1065
            TabIndex        =   7
            Top             =   1185
            Width           =   4500
            _cx             =   7937
            _cy             =   3043
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16772055
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   12698049
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   2
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   1
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   1
            Left            =   2595
            TabIndex        =   34
            Top             =   480
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   122945539
            CurrentDate     =   39500
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   0
            Left            =   1065
            TabIndex        =   3
            Top             =   480
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   122945539
            CurrentDate     =   39500
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Index           =   17
            Left            =   3930
            TabIndex        =   41
            Top             =   885
            Width           =   90
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Index           =   16
            Left            =   3915
            TabIndex        =   40
            Top             =   2955
            Width           =   90
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Index           =   15
            Left            =   7995
            TabIndex        =   35
            Top             =   150
            Width           =   120
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000011&
            X1              =   510
            X2              =   1815
            Y1              =   375
            Y2              =   375
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Height          =   180
            Index           =   2
            Left            =   2400
            TabIndex        =   1
            Top             =   525
            Width           =   180
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            Height          =   180
            Index           =   3
            Left            =   3990
            TabIndex        =   4
            Top             =   525
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�������:"
            Height          =   180
            Index           =   13
            Left            =   60
            TabIndex        =   0
            Top             =   540
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Index           =   14
            Left            =   90
            TabIndex        =   33
            Top             =   150
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������Ա:"
            Height          =   180
            Index           =   4
            Left            =   60
            TabIndex        =   5
            Top             =   900
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���Ĳ���:"
            Height          =   180
            Index           =   5
            Left            =   60
            TabIndex        =   6
            Top             =   1245
            Width           =   810
         End
      End
   End
End
Attribute VB_Name = "frmCISBorrowEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

Private mfrmMain As Object
Private mlngKey As Long
Private mlngReferKey As Long
Private mblnReading As Boolean
Private mstrSQL As String
Private mblnDataChanged As Boolean
Private mblnAllowModify As Boolean
Private mbytMode As Byte
Private mlngMoudal As Long
Private mstrPrivs As String

Private mblnBorrowAccount As Boolean '��������¼�����ԭ��
Private WithEvents mclsPatient As clsVsf
Attribute mclsPatient.VB_VarHelpID = -1

Public Event AfterDataChanged()
Public Event ViewDocument(ByVal strNo As String, ByVal lng����ID As Long, ByVal lng��ҳID As Long)

'######################################################################################################################
Public Property Let AllowModify(blnData As Boolean)
    mblnAllowModify = blnData
End Property

Public Property Get AllowModify() As Boolean
    AllowModify = mblnAllowModify
End Property

Public Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
    If mblnReading = False Then
        RaiseEvent AfterDataChanged
    End If
End Property

Public Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Public Function InitData(ByVal frmMain As Object, ByVal lngMoudal As Long, ByVal blnAllowModify As Boolean, ByVal strPrivs As String, ByVal blnBorrowAccount As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Set mfrmMain = frmMain
    mblnAllowModify = blnAllowModify
    mlngMoudal = lngMoudal
    mstrPrivs = strPrivs
    mblnBorrowAccount = blnBorrowAccount
    If ExecuteCommand("��ʼ�ؼ�") = False Or ExecuteCommand("��ʼ����") = False Then Exit Function
    Call ExecuteCommand("�ؼ�״̬")
        
    DataChanged = False
End Function

Public Function AddPerson() As Boolean
    
    If cmd(0).Enabled And cmd(0).Visible Then
        Call cmd_Click(0)
    End If
    
    AddPerson = True
End Function

Public Function RemovePerson() As Boolean
    
    If cmd(1).Enabled And cmd(1).Visible Then
        Call cmd_Click(1)
    End If
    
    RemovePerson = True
End Function

Public Function AddPatient() As Boolean
    
    If cmd(2).Enabled And cmd(2).Visible Then
        Call cmd_Click(2)
    End If
    
    AddPatient = True
End Function

Public Function RemovePatient() As Boolean
    
    If cmd(3).Enabled And cmd(3).Visible Then
        Call cmd_Click(3)
    End If
    
    RemovePatient = True
    
End Function

Public Function ClearData() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    ClearData = ExecuteCommand("�������")
End Function

Public Function RefreshData(ByVal lngKey As Long, ByVal blnAllowModify As Boolean, ByVal blnBorrowAccount As Boolean) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mblnAllowModify = blnAllowModify
    mlngKey = lngKey
    mbytMode = 2
    mblnBorrowAccount = blnBorrowAccount
    Call ExecuteCommand("�������")
    Call ExecuteCommand("��ʼ����")
            
    If ExecuteCommand("��ȡ����", mlngKey) = False Then Exit Function
    
    Call ExecuteCommand("�ؼ�״̬")
    DataChanged = False
    
    RefreshData = True
    
End Function

Public Function NewData(Optional ByVal lngReferKey As Long = 0) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mblnAllowModify = True
    mlngKey = 0
    mlngReferKey = lngReferKey
    
    mbytMode = 1
   
    Call ExecuteCommand("�������")
    Call ExecuteCommand("��ʼ����")
    Call ExecuteCommand("�ؼ�״̬")
    Call ExecuteCommand("ȱʡ����")

    DataChanged = True
    
    dtp(0).SetFocus
        
    NewData = True
End Function

Public Function Aduit() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    mbytMode = 3
    
    Call ExecuteCommand("�ؼ�״̬")
    Call ExecuteCommand("ȱʡ����")

    DataChanged = True
    
    dtp(2).SetFocus
        
    Aduit = True
End Function

Public Function Revert() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    mbytMode = 5
    
    Call ExecuteCommand("�ؼ�״̬")
    Call ExecuteCommand("ȱʡ����")

    DataChanged = True
    
    txt(1).SetFocus
        
    Revert = True
End Function


Public Function Refuse() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    mbytMode = 4
    
    Call ExecuteCommand("�ؼ�״̬")
    Call ExecuteCommand("ȱʡ����")

    DataChanged = True
    
    Call LocationObj(txt(0))
        
    Refuse = True
End Function

Public Function ValidData(ByVal blnBorrowReason As Boolean) As Boolean
    '******************************************************************************************************************
    '���ܣ�У��༭���ݵ���Ч��
    '������
    '���أ�
    '******************************************************************************************************************
    Dim i As Long
    Select Case mbytMode
    Case 1, 2
        
        If StrIsValid(cbo(0).Text, 255) = False Then
          cbo(0).SetFocus
          Exit Function
        End If
        
        If txtBorrowUser.Text = "" Or txtBorrowUser.Tag = "" Then
            ShowSimpleMsg "�����Ľ�����Ա����Ϊ��ֵ���������룡"
            txtBorrowUser.SetFocus
            Exit Function
        End If
        
        With vsf(1)
            If .Rows = 2 And Val(.TextMatrix(1, .ColIndex("����id"))) = 0 And Val(.TextMatrix(1, .ColIndex("��ҳid"))) = 0 Then
                ShowSimpleMsg "���ĵĲ��˲�������Ϊ��ֵ���������룡"
                mclsPatient.SetFocus
                Exit Function
            End If
        End With
        
        If Format(dtp(1).Value, dtp(1).CustomFormat) < Format(dtp(0).Value, dtp(0).CustomFormat) Then
            ShowSimpleMsg "�����Ľ�������Ľ��Ľ���ʱ�䲻��С�ڿ�ʼʱ�䣡"
            dtp(1).SetFocus
            Exit Function
        End If
        
        If DateDiff("d", dtp(0).Value, dtp(1).Value) > Val(GetPara("���������", mfrmMain.ģ���, "30")) Then
            
            ShowSimpleMsg "�������ĵ������ʱ�䲻�ܳ���" & Val(GetPara("���������", mfrmMain.ģ���, "30")) & "�죡"
            dtp(1).SetFocus
            Exit Function

        End If
        
        If blnBorrowReason Then
            If cbo(0).Text = "" Then
                ShowSimpleMsg "�����벡��������������!"
                cbo(0).SetFocus
                Exit Function
            End If
        End If
        
    Case 3
        If Format(dtp(3).Value, dtp(3).CustomFormat) < Format(dtp(2).Value, dtp(2).CustomFormat) Then
            ShowSimpleMsg "��������׼���ĵĽ��Ľ���ʱ�䲻��С�ڿ�ʼʱ�䣡"
            dtp(3).SetFocus
            Exit Function
        End If
        
        If DateDiff("d", dtp(2).Value, dtp(3).Value) > Val(GetPara("���������", mfrmMain.ģ���, "30")) Then
            
            ShowSimpleMsg "�������ĵ������ʱ�䲻�ܳ���" & Val(GetPara("���������", mfrmMain.ģ���, "30")) & "�죡"
            dtp(3).SetFocus
            Exit Function

        End If
        
        '����Ƿ����Ѿ����ĵĲ���
        With vsf(1)
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("�����洢״̬")) = "��Ժ" Then
                    ValidData = True
                Else
                    MsgBox "ѡ��Ĳ���:[" & .TextMatrix(i, .ColIndex("����")) & "]�Ѿ���[" & .TextMatrix(i, .ColIndex("���������")) & "]������,������ѡ��!", vbInformation, gstrSysName
                    ValidData = False
                    Exit Function
                End If
            Next
        End With
        
        
        
    Case 4
        If Trim(txt(0).Text) = "" Then
            ShowSimpleMsg "�ܾ�����ʱ�䣬��������ܾ����ɣ�"
            LocationObj txt(0)
            Exit Function
        End If
    Case 5

        
        
    End Select
    
    ValidData = True
    
End Function

Public Function SaveData(ByRef rsSQL As ADODB.Recordset, ByRef lngKey As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strSQL As String
    Dim lngLoop As Long
    Dim strTmp As String
    
    On Error GoTo errHand
    
    Select Case mbytMode
    Case 1, 2
        If mlngKey = 0 Then
            '����
            lngKey = zlDatabase.GetNextId("�������ļ�¼")
            txt(5).Text = zlDatabase.GetNextNo(91)
        Else
            '�޸�
            lngKey = mlngKey
            
            strSQL = "zl_����������Ա_Update(" & lngKey & ",Null)"
            Call SQLRecordAdd(rsSQL, strSQL)
        
            strSQL = "zl_������������_Update(" & lngKey & ",Null)"
            Call SQLRecordAdd(rsSQL, strSQL)
        End If
    
        strSQL = "zl_�������ļ�¼_Update(" & lngKey & ",'" & txt(5).Text & "','" & txt(6).Text & "','" & cbo(0).Text & "',To_Date('" & Format(dtp(0).Value, dtp(0).CustomFormat) & " 00:00:00','yyyy-mm-dd hh24:mi:ss'),To_Date('" & Format(dtp(1).Value, dtp(1).CustomFormat) & " 23:59:59','yyyy-mm-dd hh24:mi:ss'),To_Date('" & txt(2).Text & ":00','yyyy-mm-dd hh24:mi:ss'))"
        Call SQLRecordAdd(rsSQL, strSQL)
                    
        strTmp = ""
        strTmp = txtBorrowUser.Tag
        
        strSQL = "zl_����������Ա_Update(" & lngKey & ",'" & strTmp & "')"
        Call SQLRecordAdd(rsSQL, strSQL)
        
        strTmp = ""
        With vsf(1)
            For lngLoop = 1 To .Rows - 1
                If Val(.TextMatrix(lngLoop, .ColIndex("����id"))) > 0 And Val(.TextMatrix(lngLoop, .ColIndex("��ҳid"))) > 0 Then
                    If strTmp = "" Then
                        strTmp = Val(.TextMatrix(lngLoop, .ColIndex("����id"))) & ":" & Val(.TextMatrix(lngLoop, .ColIndex("��ҳid")))
                    Else
                        strTmp = strTmp & ";" & Val(.TextMatrix(lngLoop, .ColIndex("����id"))) & ":" & Val(.TextMatrix(lngLoop, .ColIndex("��ҳid")))
                    End If
                End If
            Next
        End With
        strSQL = "zl_������������_Update(" & lngKey & ",'" & strTmp & "')"
        Call SQLRecordAdd(rsSQL, strSQL)
    '------------------------------------------------------------------------------------------------------------------
    Case 3
        strSQL = "zl_�������ļ�¼_Authorize(" & lngKey & ",To_Date('" & Format(dtp(2).Value, dtp(2).CustomFormat) & " 00:00:00','yyyy-mm-dd hh24:mi:ss'),To_Date('" & Format(dtp(3).Value, dtp(3).CustomFormat) & " 23:59:59','yyyy-mm-dd hh24:mi:ss'),'" & txt(7).Text & "',To_Date('" & txt(3).Text & ":00','yyyy-mm-dd hh24:mi:ss'))"
        Call SQLRecordAdd(rsSQL, strSQL)
    '------------------------------------------------------------------------------------------------------------------
    Case 4
        strSQL = "zl_�������ļ�¼_Refuse(" & lngKey & ",'" & txt(8).Text & "','" & txt(0).Text & "',To_Date('" & txt(4).Text & ":00','yyyy-mm-dd hh24:mi:ss'))"
        Call SQLRecordAdd(rsSQL, strSQL)
    '------------------------------------------------------------------------------------------------------------------
    Case 5
        strSQL = "zl_�������ļ�¼_Revert(" & lngKey & ",'" & txt(1).Text & "',To_Date('" & txt(9).Text & ":00','yyyy-mm-dd hh24:mi:ss'))"
        Call SQLRecordAdd(rsSQL, strSQL)
    '------------------------------------------------------------------------------------------------------------------
    End Select
    
    SaveData = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
End Function

'######################################################################################################################
Private Function ExecuteCommand(ByVal strCmd As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim blnAllowModify As Boolean
    Dim lngNum As Long
        
    On Error GoTo errHand
    
    mblnReading = True
    Call SQLRecord(rsSQL)
    
    Select Case strCmd
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ�ؼ�"

        Set mclsPatient = New clsVsf
        With mclsPatient
            Call .Initialize(Me.Controls, vsf(1), True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            If AllowModify Then
                Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ָʾ��]", False)
            Else
                Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
            End If
            Call .AppendColumn("����", 810, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("����id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("��ҳid", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("�Ա�", 600, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("����", 600, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("����״��", 900, flexAlignLeftCenter, flexDTString, "", , True)
            
            Call .AppendColumn("סԺ��", 810, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("������", 810, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("סԺ����", 810, flexAlignLeftCenter, flexDTString, "", , True)
            
            Call .AppendColumn("��Ժʱ��", 1670, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm", , True)
            Call .AppendColumn("��Ժʱ��", 1670, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm", , True)
            Call .AppendColumn("��Ժ����", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("���״̬", 0, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("���������", 0, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("�����洢״̬", 0, flexAlignLeftCenter, flexDTString, "", , True)
            
            
            If AllowModify Then
                Call .InitializeEdit(True, True, True)
                Call .InitializeEditColumn(.ColIndex("����"), True, vbVsfEditCommand)
                .IndicatorCol = 0
                Set .IndicatorIcon = GetImageList(16).ListImages("��ǰ").Picture
            End If
            .AppendRows = True
        End With
                
        '��������
        '----------------------------------------------------------------------------------------------------------
        cbo(0).Clear
        cbo(0).AddItem ""
        Set rs = gclsPackage.GetDictTableData("��������")
        If rs.BOF = False Then
            Do While Not rs.EOF
                cbo(0).AddItem rs("����").Value
                If rs("ȱʡ��־").Value = 1 Then cbo(0).ListIndex = cbo(0).NewIndex
                rs.MoveNext
            Loop
        End If
        If cbo(0).ListCount > 0 And cbo(0).ListIndex = -1 Then cbo(0).ListIndex = 0
            
    '------------------------------------------------------------------------------------------------------------------
    Case "�ؼ�״̬"
    
        blnAllowModify = mblnAllowModify
        If (mlngKey = 0 And mbytMode = 2) Or lbl(15).Caption <> "" Then blnAllowModify = False
        
        cmd(0).Enabled = blnAllowModify
        cmd(2).Enabled = blnAllowModify
        cmd(3).Enabled = blnAllowModify
        
        Select Case mbytMode
        Case 1, 2
            txt(0).Locked = Not blnAllowModify
            cbo(0).Locked = Not blnAllowModify
            txt(2).Locked = True
            txt(3).Locked = True
            txt(4).Locked = True
            txt(5).Locked = True
            txt(6).Locked = True
            txt(7).Locked = True
            txt(8).Locked = True
            txt(1).Locked = True
            txt(9).Locked = True
            
            dtp(0).Enabled = blnAllowModify
            dtp(1).Enabled = blnAllowModify
            dtp(2).Enabled = blnAllowModify
            dtp(3).Enabled = blnAllowModify
            
            If blnAllowModify Then
                txtBorrowUser.Enabled = True
                Call mclsPatient.InitializeEdit(True, True, True)
            Else
                txtBorrowUser.Enabled = False
                Call mclsPatient.InitializeEdit(False, False, False)
            End If
        Case 3          '��׼
            txt(0).Locked = True
            cbo(0).Locked = True
            txt(2).Locked = True
            txt(3).Locked = False
            txt(4).Locked = True
            txt(5).Locked = True
            txt(6).Locked = True
            txt(7).Locked = False
            txt(8).Locked = True
            txt(1).Locked = True
            txt(9).Locked = True
            
            dtp(0).Enabled = False
            dtp(1).Enabled = False
            dtp(2).Enabled = True
            dtp(3).Enabled = True
            txtBorrowUser.Enabled = False
            Call mclsPatient.InitializeEdit(False, False, False)
        Case 4          '�ܽ�
            txt(0).Locked = False
            cbo(0).Locked = True
            txt(2).Locked = True
            txt(3).Locked = True
            txt(4).Locked = False
            txt(5).Locked = True
            txt(6).Locked = True
            txt(7).Locked = True
            txt(8).Locked = False
            txt(1).Locked = True
            txt(9).Locked = True
            
            dtp(0).Enabled = False
            dtp(1).Enabled = False
            dtp(2).Enabled = False
            dtp(3).Enabled = False
            txtBorrowUser.Enabled = False
            Call mclsPatient.InitializeEdit(False, False, False)
        Case 5          '�黹
            txt(0).Locked = True
            cbo(0).Locked = True
            txt(2).Locked = True
            txt(3).Locked = True
            txt(4).Locked = True
            txt(5).Locked = True
            txt(6).Locked = True
            txt(7).Locked = True
            txt(8).Locked = True
            txt(1).Locked = False
            txt(9).Locked = False
            
            dtp(0).Enabled = False
            dtp(1).Enabled = False
            dtp(2).Enabled = False
            dtp(3).Enabled = False
            txtBorrowUser.Enabled = False
            Call mclsPatient.InitializeEdit(False, False, False)
        End Select
            
        For lngNum = 0 To 9
            If txt(lngNum).Locked Then
                txt(lngNum).Enabled = False
            Else
                txt(lngNum).Enabled = True
            End If
        Next
    '------------------------------------------------------------------------------------------------------------------
    Case "������Ϣ"
        
        With vsf(1)
            If Val(.RowData(.Rows - 1)) > 0 Then
                lbl(16).Caption = "���ĵĲ������� " & .Rows - 1 & " ��"
            Else
                lbl(16).Caption = "���ĵĲ������� " & .Rows - 2 & " ��"
            End If
        End With
            
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"

        txt(0).MaxLength = GetMaxLength("�������ļ�¼", "��������")
                    
    '------------------------------------------------------------------------------------------------------------------
    Case "ˢ������"
        
        ExecuteCommand = ExecuteCommand("��ȡ����", Val(varParam(0)))
        GoTo endHand
        
    '------------------------------------------------------------------------------------------------------------------
    Case "�������"
                
        txt(0).Text = ""
        cbo(0).Text = ""
        txt(2).Text = ""
        txt(3).Text = ""
        txt(4).Text = ""
        txt(5).Text = ""
        txt(6).Text = ""
        txt(7).Text = ""
        txt(8).Text = ""
        txt(1).Text = ""
        txt(9).Text = ""
        lbl(15).Caption = ""
        txtConver(0).Visible = True
        txtConver(1).Visible = True
        txtConver(2).Visible = True
        txtConver(3).Visible = True
        dtp(0).Enabled = False
        dtp(1).Enabled = False
        dtp(2).Enabled = False
        dtp(3).Enabled = False
        txtBorrowUser.Text = ""
        txtBorrowUser.Tag = ""
        mclsPatient.ClearGrid
        
        Call ExecuteCommand("������Ϣ")
    '------------------------------------------------------------------------------------------------------------------
    Case "ȱʡ����"
        
        Select Case mbytMode
        Case 1, 2
            
            dtp(0).Value = Format(zlDatabase.Currentdate, dtp(0).CustomFormat)
            
            If Val(GetPara("������������", mfrmMain.ģ���, "7")) = 0 Then
                dtp(1).Value = Format(zlDatabase.Currentdate + 8, dtp(1).CustomFormat)
            Else
                dtp(1).Value = Format(zlDatabase.Currentdate + 1 + Val(GetPara("������������", mfrmMain.ģ���, "7")), dtp(1).CustomFormat)
            End If

            txtConver(0).Visible = False
            txtConver(1).Visible = False
            dtp(0).Enabled = True
            dtp(1).Enabled = True
            txt(6).Text = UserInfo.����
            txt(2).Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
        Case 3
            txt(7).Text = UserInfo.����
            txt(3).Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
            dtp(2).Value = dtp(0).Value
            dtp(3).Value = dtp(1).Value
            txtConver(2).Visible = False
            txtConver(3).Visible = False
        Case 4
            txt(8).Text = UserInfo.����
            txt(4).Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
        Case 5
            txt(1).Text = UserInfo.����
            txt(9).Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
        End Select
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ����"
        
        Call ExecuteCommand("�������")
        mblnReading = True
        
        If Val(varParam(0)) > 0 Then
            Set rs = gclsPackage.GetBorrow(1, Val(varParam(0)))
            If rs.BOF = False Then
                txt(5).Text = zlCommFun.NVL(rs("No").Value)
                cbo(0).Text = zlCommFun.NVL(rs("��������").Value)
                txt(0).Text = zlCommFun.NVL(rs("�ܽ�����").Value)
                            
                txtConver(0).Visible = IsNull(rs("����ʱ��").Value)
                txtConver(1).Visible = IsNull(rs("��������").Value)
                txtConver(2).Visible = IsNull(rs("����ʱ��").Value)
                txtConver(3).Visible = IsNull(rs("��������").Value)
                
                If IsNull(rs("����ʱ��").Value) = False Then dtp(0).Value = Format(rs("����ʱ��").Value, dtp(0).CustomFormat)
                If IsNull(rs("��������").Value) = False Then dtp(1).Value = Format(rs("��������").Value, dtp(1).CustomFormat)
                If IsNull(rs("����ʱ��").Value) = False Then dtp(2).Value = Format(rs("����ʱ��").Value, dtp(2).CustomFormat)
                If IsNull(rs("��������").Value) = False Then dtp(3).Value = Format(rs("��������").Value, dtp(3).CustomFormat)
                
                txt(6).Text = zlCommFun.NVL(rs("������").Value)
                txt(7).Text = zlCommFun.NVL(rs("��׼��").Value)
                txt(8).Text = zlCommFun.NVL(rs("�ܽ���").Value)
                txt(1).Text = zlCommFun.NVL(rs("�ջ���").Value)
                 
                If IsNull(rs("�Ǽ�ʱ��").Value) = False Then txt(2).Text = Format(rs("�Ǽ�ʱ��").Value, "yyyy-MM-dd HH:mm")
                If IsNull(rs("��׼ʱ��").Value) = False Then txt(3).Text = Format(rs("��׼ʱ��").Value, "yyyy-MM-dd HH:mm")
                If IsNull(rs("�ܽ�ʱ��").Value) = False Then txt(4).Text = Format(rs("�ܽ�ʱ��").Value, "yyyy-MM-dd HH:mm")
                If IsNull(rs("�黹ʱ��").Value) = False Then txt(9).Text = Format(rs("�黹ʱ��").Value, "yyyy-MM-dd HH:mm")
                
                dtp(0).Enabled = Not txtConver(0).Visible
                dtp(1).Enabled = Not txtConver(1).Visible
                dtp(2).Enabled = Not txtConver(2).Visible
                dtp(3).Enabled = Not txtConver(3).Visible
                
                Select Case rs("��¼״̬").Value
                Case 1
                    lbl(15).Caption = ""
                Case 2
                    lbl(15).Caption = "����׼"
                Case 3
                    lbl(15).Caption = "�Ѿܾ�"
                Case 4
                    lbl(15).Caption = "�ѹ黹"
                End Select
            End If
        End If
        
        If lbl(15).Caption = "" And mlngKey > 0 Then
            Call mclsPatient.ModifyColumn(0, "", 255, flexAlignCenterCenter, flexDTString, "", "[ָʾ��]", False)
            
        Else
            Call mclsPatient.ModifyColumn(0, "", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
        End If
    
        If Val(varParam(0)) > 0 Then
               
            Set rs = gclsPackage.GetBorrowPerson(Val(varParam(0)))
            If rs.BOF = False Then
                txtBorrowUser.Text = zlCommFun.NVL(rs!����)
                txtBorrowUser.Tag = zlCommFun.NVL(rs!ID, 0)
            End If
            
            Set rs = gclsPackage.GetBorrowPatient(Val(varParam(0)))
            If rs.BOF = False Then
                Call mclsPatient.LoadGrid(rs)
            End If
        End If
        
        Call ExecuteCommand("������Ϣ")
        
    End Select

    ExecuteCommand = True
    
    GoTo endHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
endHand:
    mblnReading = False
End Function

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then
        If mblnBorrowAccount = False Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    
    Dim lngLoop As Long
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim bytRet As Byte
    
    Select Case Index
    Case 0

        Set rsData = gclsPackage.GetOperationPerson
        bytRet = ShowPubSelect(Me, txtBorrowUser, 2, "���,1200,0,;����,1200,0,;����,900,0,;����,1200,0,", Me.Name & "\������Աѡ��", "����±���ѡ��һ������������Ա", rsData, rs, 8790, 4500, False, txtBorrowUser.Tag)
                    
        If bytRet = 1 Then
            If rs.RecordCount = 1 Then
                txtBorrowUser.Text = zlCommFun.NVL(rs("����").Value)
                txtBorrowUser.Tag = zlCommFun.NVL(rs("ID").Value, 0)
            End If
            DataChanged = True
        End If
   
    Case 1
        
'        lngLoop = vsf(0).Row
'        Call mclsPerson.DeleteRow(vsf(0).Row)
'
'        If lngLoop <= vsf(0).Rows - 1 Then
'            vsf(0).Row = lngLoop
'        Else
'            vsf(0).Row = vsf(0).Rows - 1
'        End If
        
    Case 2
    
        If frmSearchPatient.ShowEdit(Me, rs, mlngMoudal, mstrPrivs) Then
            If rs.RecordCount > 0 Then
                rs.MoveFirst
                With vsf(1)
                    Do While Not rs.EOF
                                                                                
                        If mclsPatient.CheckHave(rs("ID").Value, False) = False Then
                            If Trim(.RowData(.Rows - 1)) <> "" And Trim(.RowData(.Rows - 1)) <> "0" Then .Rows = .Rows + 1
                            .RowData(.Rows - 1) = Trim(rs("ID").Value)
                            .TextMatrix(.Rows - 1, .ColIndex("����id")) = Val(rs("����id").Value)
                            .TextMatrix(.Rows - 1, .ColIndex("��ҳid")) = Val(rs("��ҳid").Value)
                            .TextMatrix(.Rows - 1, .ColIndex("����")) = Trim(rs("����").Value)
                            .TextMatrix(.Rows - 1, .ColIndex("�Ա�")) = Trim(rs("�Ա�").Value)
                            .TextMatrix(.Rows - 1, .ColIndex("����")) = Trim(rs("����").Value)
                            .TextMatrix(.Rows - 1, .ColIndex("����״��")) = Trim(rs("����״��").Value)
                            .TextMatrix(.Rows - 1, .ColIndex("��Ժʱ��")) = Trim(rs("��Ժʱ��").Value)
                            .TextMatrix(.Rows - 1, .ColIndex("��Ժʱ��")) = Trim(rs("��Ժʱ��").Value)
                            .TextMatrix(.Rows - 1, .ColIndex("��Ժ����")) = Trim(rs("��Ժ����").Value)
                            
                            .TextMatrix(.Rows - 1, .ColIndex("סԺ��")) = Trim(rs("סԺ��").Value)
                            .TextMatrix(.Rows - 1, .ColIndex("������")) = Trim(rs("������").Value)
                            .TextMatrix(.Rows - 1, .ColIndex("סԺ����")) = Trim(rs("סԺ����").Value)
                            
                            DataChanged = True
                        End If
                        
                        rs.MoveNext
                    Loop
                End With
            End If
        End If
        
    Case 3
                
        lngLoop = vsf(1).Row
        Call mclsPatient.DeleteRow(vsf(1).Row)
        
        If lngLoop <= vsf(1).Rows - 1 Then
            vsf(1).Row = lngLoop
        Else
            vsf(1).Row = vsf(1).Rows - 1
        End If
        
    End Select
    
    Call ExecuteCommand("������Ϣ")
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    picPane(0).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    picPane(1).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    fra.BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    txt(5).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    picPane(0).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsPatient = Nothing
End Sub

Private Sub mclsPatient_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    Call ExecuteCommand("������Ϣ")
    DataChanged = True
End Sub

Private Sub mclsPatient_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = (Val(vsf(1).RowData(Row)) = 0)
End Sub

Private Sub mclsPerson_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    Call ExecuteCommand("������Ϣ")
    DataChanged = True
End Sub

Private Sub mclsPerson_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = (Val(vsf(0).RowData(Row)) = 0)
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next

    Select Case Index
    Case 0
        fra.Move 0, -75, picPane(Index).Width, picPane(Index).Height + 75
        
        cbo(0).Move cbo(0).Left, cbo(0).Top, fra.Width - cbo(0).Left - 45
'        vsf(0).Move vsf(0).Left, vsf(0).Top, fra.Width - vsf(0).Left - 45
        
        txtBorrowUser.Move txtBorrowUser.Left, txtBorrowUser.Top, fra.Width - txtBorrowUser.Left - 45 - cmd(0).Width
        cmd(0).Move txtBorrowUser.Left + txtBorrowUser.Width + 15, txtBorrowUser.Top
        
        vsf(1).Move txtBorrowUser.Left, vsf(1).Top, txtBorrowUser.Width, fra.Height - vsf(1).Top - (picPane(1).Height + 45) - 75
        cmd(2).Move vsf(1).Left + vsf(1).Width + 15, vsf(1).Top
        cmd(3).Move cmd(2).Left, cmd(2).Top + cmd(2).Height + 15
        
        
        picPane(1).Move 30, vsf(1).Top + vsf(1).Height + 45, fra.Width - 60
        
        lbl(15).Move fra.Width - 900
        
        mclsPatient.AppendRows = True
    Case 1
        txt(0).Move txt(0).Left, txt(0).Top, picPane(Index).Width - txt(0).Left - 45
    End Select
End Sub

Private Sub txt_Change(Index As Integer)
    
    If mblnReading Then Exit Sub
    
    
    DataChanged = True
End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txt(Index)
    
    Select Case Index
    Case 0, 1
        zlCommFun.OpenIme True
    End Select
        
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)

    Select Case Index
    Case 0, 1
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
    
End Sub

Private Sub txtBorrowUser_KeyPress(KeyAscii As Integer)
    Dim rsData As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim bytRet As Byte
    
    If KeyAscii = vbKeyReturn Then
        Set rsData = gclsPackage.GetOperationPerson(UCase(txtBorrowUser.Text))

         If ShowPubSelect(Me, txtBorrowUser, 2, "���,1200,0,;����,1200,0,;����,900,0,;����,1200,0,", Me.Name & "\������Ա����", "����±���ѡ��һ��������Ա", rsData, rs, 8790, 4500, , txtBorrowUser.Tag, , True) = 1 Then

             txtBorrowUser.Text = zlCommFun.NVL(rs("����").Value)
             txtBorrowUser.Tag = zlCommFun.NVL(rs("ID").Value, 0)
             DataChanged = True
         End If
    End If
End Sub

Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    '�༭����
    Select Case Index
    Case 0
        
    Case 1
        Call mclsPatient.AfterEdit(Row, Col)
    End Select
    
    DataChanged = True
    
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    '�༭����
    Select Case Index
    Case 0
 
    Case 1
        Call mclsPatient.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    End Select
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Select Case Index
    Case 0

    Case 1
        mclsPatient.AppendRows = True
    End Select
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Select Case Index
    Case 0
'        mclsPerson.AppendRows = True
    Case 1
        mclsPatient.AppendRows = True
    End Select
End Sub

Private Sub vsf_BeforeUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Index
    Case 0
'        Call mclsPerson.BeforeResizeColumn(Col, Cancel)
    Case 1
        Call mclsPatient.BeforeResizeColumn(Col, Cancel)
    End Select
End Sub

Private Sub vsf_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)

    Dim lngLoop As Long
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim bytRet As Byte
    
    With vsf(Index)
        Select Case Index
        '--------------------------------------------------------------------------------------------------------------
        Case 0
'            If Col = .ColIndex("����") Then
'
'                Set rsData = gclsPackage.GetOperationPerson
'                bytRet = ShowPubSelect(Me, vsf(Index), 2, "���,1200,0,;����,1200,0,;����,900,0,;����,1200,0,", Me.Name & "\������Աѡ��", "����±���ѡ��һ������������Ա", rsData, rs, 8790, 4500, True, Val(.RowData(Row)))
'
'                If bytRet = 1 Then
'
'                    For lngLoop = 1 To rs.RecordCount
'
'                        If mclsPerson.CheckHave(zlCommFun.NVL(rs("ID").Value), False) = False Then
'
'                            If Val(.RowData(.Rows - 1)) > 0 Then .Rows = .Rows + 1
'
'                            .EditText = zlCommFun.NVL(rs("����").Value)
'                            .TextMatrix(.Rows - 1, .ColIndex("����")) = zlCommFun.NVL(rs("����").Value)
'                            .TextMatrix(.Rows - 1, .ColIndex("���")) = zlCommFun.NVL(rs("���").Value)
'                            .TextMatrix(.Rows - 1, .ColIndex("����")) = zlCommFun.NVL(rs("����").Value)
'                            .RowData(.Rows - 1) = zlCommFun.NVL(rs("ID").Value, 0)
'
'                            DataChanged = True
'                        End If
'
'                        rs.MoveNext
'                    Next
'
'                    mclsPerson.AppendRows = True
'
'                    DataChanged = True
'
'                End If
'
'            End If
        '--------------------------------------------------------------------------------------------------------------
        Case 1
            If frmSearchPatient.ShowEdit(Me, rs, mlngMoudal, mstrPrivs) Then
                If rs.RecordCount > 0 Then
                    rs.MoveFirst
                    With vsf(1)
                        Do While Not rs.EOF
                                                                                    
                            If mclsPatient.CheckHave(rs("ID").Value, False) = False Then
                                If Trim(.RowData(.Rows - 1)) <> "" And Trim(.RowData(.Rows - 1)) <> "0" Then .Rows = .Rows + 1
                                .RowData(.Rows - 1) = Trim(rs("ID").Value)
                                .TextMatrix(.Rows - 1, .ColIndex("����id")) = Val(rs("����id").Value)
                                .TextMatrix(.Rows - 1, .ColIndex("��ҳid")) = Val(rs("��ҳid").Value)
                                .TextMatrix(.Rows - 1, .ColIndex("����")) = Trim(rs("����").Value)
                                .TextMatrix(.Rows - 1, .ColIndex("�Ա�")) = Trim(rs("�Ա�").Value)
                                .TextMatrix(.Rows - 1, .ColIndex("����")) = Trim(rs("����").Value)
                                .TextMatrix(.Rows - 1, .ColIndex("����״��")) = Trim(rs("����״��").Value)
                                .TextMatrix(.Rows - 1, .ColIndex("��Ժʱ��")) = Trim(rs("��Ժʱ��").Value)
                                .TextMatrix(.Rows - 1, .ColIndex("��Ժʱ��")) = Trim(rs("��Ժʱ��").Value)
                                .TextMatrix(.Rows - 1, .ColIndex("��Ժ����")) = Trim(rs("��Ժ����").Value)
                                
                                .TextMatrix(.Rows - 1, .ColIndex("סԺ��")) = Trim(rs("סԺ��").Value)
                                .TextMatrix(.Rows - 1, .ColIndex("������")) = Trim(rs("������").Value)
                                .TextMatrix(.Rows - 1, .ColIndex("סԺ����")) = Trim(rs("סԺ����").Value)
                                
                                DataChanged = True
                            End If
                            
                            rs.MoveNext
                        Loop
                    End With
                End If
            End If
        End Select
    End With
    Call ExecuteCommand("������Ϣ")
End Sub

Private Sub vsf_DblClick(Index As Integer)
    '�༭����
    Select Case Index
    Case 0
'        Call mclsPerson.DbClick
    Case 1
        Call mclsPatient.DbClick
        
        If lbl(15).Caption = "����׼" And DataChanged = False Then
            With vsf(1)
                RaiseEvent ViewDocument(txt(5).Text, Val(.TextMatrix(.Row, .ColIndex("����id"))), Val(.TextMatrix(.Row, .ColIndex("��ҳid"))))
            End With
        End If
        
    End Select
End Sub

Private Sub vsf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    '�༭����
    Select Case Index
    Case 0
'        Call mclsPerson.KeyDown(KeyCode, Shift)
    Case 1
        Call mclsPatient.KeyDown(KeyCode, Shift)
    End Select
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    
    'ToDo...
    If KeyAscii = vbKeyReturn Then Call vsf_DblClick(Index)
    
    '�༭����,������
    Select Case Index
    Case 0
'        Call mclsPerson.KeyPress(KeyAscii)
    Case 1
        Call mclsPatient.KeyPress(KeyAscii)
    End Select
End Sub

Private Sub vsf_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    '�༭����
    Select Case Index
    Case 0
'        Call mclsPerson.KeyPressEdit(KeyAscii)
    Case 1
        Call mclsPatient.KeyPressEdit(KeyAscii)
    End Select
End Sub

Private Sub vsf_KeyDownEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim StrText As String
    Dim bytRet As Byte
    Dim blnCard As Boolean
    Dim bytFilterMode As Byte
    
    With vsf(Index)
        
        If InStr(.EditText, "'") > 0 Then
            KeyCode = 0
            .EditText = ""
            Exit Sub
        End If
                            
        StrText = .EditText
        
        Select Case Index
        '----------------------------------------------------------------------------------------------------------
        Case 0
'            If KeyCode = vbKeyReturn Then
'                If Col = .ColIndex("����") Then
'
'                    Set rsData = gclsPackage.GetOperationPerson(UCase(StrText))
'
'                    If ShowPubSelect(Me, vsf(Index), 2, "���,1200,0,;����,1200,0,;����,900,0,;����,1200,0,", Me.Name & "\������Ա����", "����±���ѡ��һ��������Ա", rsData, rs, 8790, 4500, , Val(.RowData(Row)), , True) = 1 Then
'
'                        If mclsPerson.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
'                            ShowSimpleMsg "ѡ�����Ա��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
'                            Exit Sub
'                        End If
'
'                        .EditText = zlCommFun.NVL(rs("����").Value)
'                        .Cell(flexcpData, Row, Col) = zlCommFun.NVL(rs("����").Value)
'                        .TextMatrix(Row, .ColIndex("����")) = zlCommFun.NVL(rs("����").Value)
'                        .TextMatrix(Row, .ColIndex("���")) = zlCommFun.NVL(rs("���").Value)
'                        .TextMatrix(Row, .ColIndex("����")) = zlCommFun.NVL(rs("����").Value)
'                        .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
'
'                        DataChanged = True
'                    Else
'                        .Cell(flexcpData, Row, Col) = .Cell(flexcpData, Row, Col)
'                        .EditText = .Cell(flexcpData, Row, Col)
'                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
'                    End If
'
'                End If
'            Else
'                DataChanged = True
'            End If
        '----------------------------------------------------------------------------------------------------------
        Case 1
        
            If Col = .ColIndex("����") Then

                If KeyCode <> 8 And KeyCode <> 13 Then StrText = StrText & Chr(KeyCode)

                '���Ƿ��ַ�
                If InStr(StrText, "'") > 0 Then
                    KeyCode = 0
                    ShowSimpleMsg "�ڸ����������зǷ��ַ� ' ��"
                    .EditText = ""
                    .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                    Exit Sub
                End If

                '����Ƿ�Ϊ���￨����
                blnCard = VsfInputIsCard(vsf(Index), KeyCode, ParamInfo.ϵͳ��)
                If blnCard And Len(.EditText) = ParamInfo.���￨���볤�� - 1 And KeyCode <> 8 And KeyCode <> vbKeyReturn Then
                    .EditSelStart = Len(.EditText)
                    bytFilterMode = 1
                End If
            End If

            If KeyCode = vbKeyReturn Then
                If Col = .ColIndex("����") Then
                    If blnCard Then
                        '�Ǿ��￨
                        bytFilterMode = 1
                    Else
                        '�Ǿ��￨
                        blnCard = False
                        StrText = .EditText
                        
                        Select Case UCase(Left(StrText, 1))
                        Case "-", "A"                   '����id
                            bytFilterMode = 2
                            StrText = IIf(IsNumeric(Mid(StrText, 2)), Val(Mid(StrText, 2)), "0")
                        Case "+", "B"                   'סԺ��
                            bytFilterMode = 3
                            StrText = IIf(IsNumeric(Mid(StrText, 2)), Val(Mid(StrText, 2)), "0")
                        Case "*", "D"                   '�����
                            bytFilterMode = 4
                            StrText = IIf(IsNumeric(Mid(StrText, 2)), Val(Mid(StrText, 2)), "0")
                        Case "/", "C"                   '��ǰ����
                            bytFilterMode = 5
                            StrText = IIf(IsNumeric(Mid(StrText, 2)), Val(Mid(StrText, 2)), "0")
                        Case "F"                        '������
                            bytFilterMode = 7
                            StrText = Mid(StrText, 2)
                        Case Else                       '����
                            bytFilterMode = 6
                        End Select
                        
                    End If
                End If
            End If
            
            If Col = .ColIndex("����") Then
                
                If bytFilterMode > 0 Then
                    
                    Set rsData = gclsPackage.GetPatient(bytFilterMode, StrText)
                    
                    If rsData.RecordCount > 0 Then
                        If rsData.RecordCount = 1 Then
                            bytRet = 1
                            Set rs = rsData
                        Else
                            bytRet = ShowPubSelect(Me, vsf(Index), 2, "����,1200,0,0;�Ա�,810,0,0;��Ժʱ��,1667,0,0;��Ժʱ��,1667,0,0;���֤��,1500,0,0", mfrmMain.Name & "\���˹���ѡ��", "�������ѡ��һ������", rsData, rs, 8790, 4500)
                        End If
                        
                        If bytRet = 1 Then
                        
                            If mclsPatient.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                ShowSimpleMsg "ѡ��Ĳ��ˡ�" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                                
                                '��ԭԭ��������
                                .Cell(flexcpData, Row, Col) = .Cell(flexcpData, Row, Col)
                                .EditText = .Cell(flexcpData, Row, Col)
                                .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                                Exit Sub
                            End If
                                      
                            .EditText = zlCommFun.NVL(rs("����"))
                            .Cell(flexcpData, Row, Col) = zlCommFun.NVL(rs("����"))
                            .TextMatrix(Row, .ColIndex("����")) = zlCommFun.NVL(rs("����").Value)
                            .TextMatrix(Row, .ColIndex("�Ա�")) = zlCommFun.NVL(rs("�Ա�").Value)
                            .TextMatrix(Row, .ColIndex("����")) = zlCommFun.NVL(rs("����").Value)
                            .TextMatrix(Row, .ColIndex("����״��")) = zlCommFun.NVL(rs("����״��").Value)
                            .TextMatrix(Row, .ColIndex("��Ժʱ��")) = zlCommFun.NVL(rs("��Ժʱ��").Value)
                            .TextMatrix(Row, .ColIndex("��Ժʱ��")) = zlCommFun.NVL(rs("��Ժʱ��").Value)
                            .TextMatrix(Row, .ColIndex("��Ժ����")) = zlCommFun.NVL(rs("��Ժ����").Value)
                            .TextMatrix(Row, .ColIndex("����id")) = zlCommFun.NVL(rs("����id").Value)
                            .TextMatrix(Row, .ColIndex("��ҳid")) = zlCommFun.NVL(rs("��ҳid").Value)
                            DataChanged = True
                            .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
    
                            If blnCard Then
                                .Cell(flexcpData, Row, Col) = StrText
                                .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                                KeyCode = 13
                            End If
                        Else
                            '��ԭԭ��������
                            .Cell(flexcpData, Row, Col) = .Cell(flexcpData, Row, Col)
                            .EditText = .Cell(flexcpData, Row, Col)
                            .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                        End If
                    Else
                        '��ԭԭ��������
                        .Cell(flexcpData, Row, Col) = .Cell(flexcpData, Row, Col)
                        .EditText = .Cell(flexcpData, Row, Col)
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                    End If
                End If
            
            End If
            
        End Select
    End With
    
    Call ExecuteCommand("������Ϣ")
End Sub

Private Sub vsf_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Select Case Index
        Case 0
'            Call mclsPerson.AutoAddRow(vsf(Index).MouseRow, vsf(Index).MouseCol)
        Case 1
            Call mclsPatient.AutoAddRow(vsf(Index).MouseRow, vsf(Index).MouseCol)
        End Select
    End Select
End Sub

Private Sub vsf_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    With vsf(Index)
        
        If .MouseCol = .ColIndex("����") And Index = 1 And (mbytMode = 1 Or mbytMode = 2) Then
            If .ToolTipText = "" Then .ToolTipText = "�����������Ҳ��˵ķ�����1.'-'��'A'+����id;2.'+'��'B'+סԺ��;3.'/'��'C'+����;4.'*'��'D'+�����;5.��������������"
        Else
            If .ToolTipText <> "" Then .ToolTipText = ""
        End If
    End With
End Sub

Private Sub vsf_SetupEditWindow(Index As Integer, ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    '�༭����
    Select Case Index
    Case 0
'        Call mclsPerson.EditSelAll
    Case 1
        Call mclsPatient.EditSelAll
    End Select
End Sub

Private Sub vsf_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '�༭����
    Select Case Index
    Case 0
'        Call mclsPerson.BeforeEdit(Row, Col, Cancel)
    Case 1
        Call mclsPatient.BeforeEdit(Row, Col, Cancel)
    End Select
End Sub

