VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmSquareRefundBalance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����˿� - ���ѿ�"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8580
   Icon            =   "frmSquareRefundBalance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fra�˿���� 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   30
      TabIndex        =   5
      Top             =   3390
      Width           =   8505
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   1785
         Index           =   2
         Left            =   2610
         ScaleHeight     =   1755
         ScaleWidth      =   5865
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   60
         Width           =   5895
         Begin VB.TextBox txt 
            Enabled         =   0   'False
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
            IMEMode         =   3  'DISABLE
            Index           =   4
            Left            =   930
            TabIndex        =   20
            Tag             =   "1"
            Top             =   960
            Width           =   4860
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
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
            Index           =   5
            Left            =   930
            TabIndex        =   22
            Tag             =   "1"
            Top             =   1380
            Width           =   4860
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
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
            Index           =   3
            Left            =   930
            TabIndex        =   18
            Tag             =   "1"
            Top             =   540
            Width           =   4860
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   1950
            TabIndex        =   14
            Top             =   112
            Width           =   1410
         End
         Begin VB.TextBox txt 
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
            Height          =   315
            Index           =   2
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   120
            Width           =   1590
         End
         Begin VB.ComboBox cbo֧����ʽ 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   930
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   112
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "��  ��"
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
            Index           =   6
            Left            =   270
            TabIndex        =   19
            Top             =   1012
            Width           =   630
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�������"
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
            Index           =   7
            Left            =   60
            TabIndex        =   21
            Top             =   1440
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������"
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
            Index           =   5
            Left            =   270
            TabIndex        =   17
            Top             =   592
            Width           =   630
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�� ��"
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
            Index           =   3
            Left            =   330
            TabIndex        =   12
            Top             =   165
            Width           =   570
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�� ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   4
            Left            =   3600
            TabIndex        =   15
            Top             =   165
            Width           =   570
         End
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   1785
         Index           =   1
         Left            =   0
         ScaleHeight     =   1755
         ScaleWidth      =   2535
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   60
         Width           =   2565
         Begin XtremeSuiteControls.ShortcutCaption ShortcutCaption1 
            Height          =   315
            Left            =   15
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   870
            Width           =   2505
            _Version        =   589884
            _ExtentX        =   4419
            _ExtentY        =   556
            _StockProps     =   6
            Caption         =   "�˷Ѻϼ�"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   9
            Left            =   1710
            TabIndex        =   10
            Top             =   1335
            Width           =   660
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   2
            Left            =   1725
            TabIndex        =   8
            Top             =   450
            Width           =   660
         End
         Begin XtremeSuiteControls.ShortcutCaption lbl�˿�ϼ� 
            Height          =   315
            Left            =   15
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   15
            Width           =   2505
            _Version        =   589884
            _ExtentX        =   4419
            _ExtentY        =   556
            _StockProps     =   6
            Caption         =   "��ǰδ��"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
         End
      End
   End
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   795
      Left            =   30
      TabIndex        =   26
      Top             =   5250
      Width           =   8505
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   210
         TabIndex        =   25
         Top             =   210
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   5790
         TabIndex        =   23
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   6990
         TabIndex        =   24
         Top             =   240
         Width           =   1100
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "������0.00"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   8
         Left            =   4020
         TabIndex        =   28
         Top             =   325
         Visible         =   0   'False
         Width           =   1395
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3345
      Index           =   0
      Left            =   30
      ScaleHeight     =   3315
      ScaleWidth      =   8475
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   8505
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   795
         TabIndex        =   2
         Top             =   120
         Width           =   2055
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfBlance 
         Height          =   2715
         Left            =   60
         TabIndex        =   4
         Top             =   510
         Width           =   8325
         _cx             =   14684
         _cy             =   4789
         Appearance      =   2
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
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483644
         GridColorFixed  =   -2147483648
         TreeColor       =   -2147483643
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSquareRefundBalance.frx":000C
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
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
         Editable        =   2
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
      Begin VB.Label lblPatiInfo 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   3000
         TabIndex        =   27
         Top             =   180
         Width           =   90
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��500.00"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   1
         Left            =   6990
         TabIndex        =   3
         Tag             =   "��"
         Top             =   165
         Width           =   1395
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmSquareRefundBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'��ڲ���
Private mfrmMain As Form, mlngModule As Long, mstrPrivs As String
Private mlng����� As Long, mlng��ID As Long

'ģ�����
Private mblnFirst As Boolean, mintSucces As Integer
Private mblnNotClick As Boolean

Private Enum mLableIndex
    lbl_��� = 1
    lbl_��ǰδ�� = 2
    lbl_�˿�ϼ� = 9
    lbl_�Ҳ� = 4
    lbl_��� = 8
End Enum
Private Enum mTextIndex
    txt_���� = 0
    txt_��� = 1
    txt_�Ҳ� = 2
    txt_������ = 3
    txt_�ʺ� = 4
    txt_������� = 5
End Enum
Private Enum mPictureIndex
    pic_�����ϸ = 0
    pic_�ɿ�ϼ� = 1
    pic_�ɿ���Ϣ = 2
End Enum

Private Type Ty_CardType
    str������ As String
    str����ǰ׺ As String
    lng���ų��� As Long
    bln�������� As Boolean
    bln�ϸ���� As Boolean
    str������� As String
    bln�ض����� As Boolean
    lng�������� As Long
    lng����ID As Long
End Type
Private mCardType As Ty_CardType

'֧�����
Private mobjPayCards As Cards
Private mlngPre֧����ʽ As Long
Private Type TY_PayMoney
    dbl�˿�ϼ� As Double
    dbl��ǰδ�� As Double
    dbl������� As Double
    strԭ������� As String
    
    lng�����ID As Long
    strˢ������ As String
    strˢ������ As String
    str������ˮ�� As String
    str����˵�� As String
End Type
Private mCurCardPay As TY_PayMoney '���ο�֧��
Private mBytMoney As Byte '�ֱҴ������

Public Function ShowMe(ByVal frmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal lng����� As Long) As Boolean
    '�������
    '��Σ�
    '   frmMain - ������
    '   lngModule - ģ���
    '   strPrivs - Ȩ�޴�
    '   lng����� As Long - ���ѿ����
    '���أ������ɹ�����True,���򷵻�False
    Set mfrmMain = frmMain: mlngModule = lngModule: mstrPrivs = strPrivs:
    mlng����� = lng�����
    mlng��ID = 0
    
    mintSucces = 0
    On Error Resume Next
    Me.Show 1, frmMain
    ShowMe = mintSucces > 0
End Function

Private Function CardIsValid(ByVal bytMode As Byte, Optional ByVal lng��ID As Long, _
    Optional ByVal str���� As String, Optional ByVal blnSaveAfter As Boolean) As Boolean
    '��鿨��Ϣ
    '��Σ�
    '   bytMode 0-�����ѿ�ID���أ�1-�����ѿ����ż���
    '   blnSaveAfter �Ƿ񱣴�����ǰ���
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strWhere As String, dbl��� As Double, dblʧЧ��� As Double
    
    On Error GoTo ErrHandler
    If bytMode = 1 Then
        strWhere = " And a.���� = [2] And a.�ӿڱ�� = [3]" & vbNewLine & _
                   " And a.��� = (Select Max(���) From ���ѿ���Ϣ Where ���� = a.���� And �ӿڱ�� = a.�ӿڱ��)"
    Else
        strWhere = " And a.Id = [1]"
    End If
    
    strSQL = _
        "Select a.ID, a.�ɷ��ֵ, a.����, a.���,to_char(a.��Ч��,'yyyy-mm-dd hh24:mi:ss') as ��Ч��, " & vbNewLine & _
        "       (Select Max(���) From ���ѿ���Ϣ B Where a.���� = b.���� And a.�ӿڱ�� = b.�ӿڱ��) As ������," & vbNewLine & _
        "       To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, " & vbNewLine & _
        "       To_Char(a.ͣ������, 'yyyy-mm-dd hh24:mi:ss') As ͣ������, a.���," & vbNewLine & _
        "       b.����, b.�Ա�, b.����" & vbNewLine & _
        "From ���ѿ���Ϣ A,������Ϣ B" & vbNewLine & _
        "Where a.����ID = b.����ID(+) " & strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��ID, str����, mlng�����)
    
    If rsTemp.EOF Then
        ShowMsgbox "δ�ҵ���ص�" & mCardType.str������ & "��Ϣ�������Ѿ�������ɾ����"
        Exit Function
    End If
    
    str���� = Nvl(rsTemp!����)
    '��鿨���Ƿ�Ϸ�
    If Val(Nvl(rsTemp!���)) < Val(Nvl(rsTemp!������)) Then
        ShowMsgbox "���ܶ���ʷ���Ž�������˿�(����Ϊ:" & str���� & ")��"
        Exit Function
    End If
    
    If Nvl(rsTemp!����ʱ��, "3000-01-01") < "3000-01-01" Then
        ShowMsgbox "����Ϊ:" & str���� & " ��" & mCardType.str������ & "�ѱ����գ�����������˿"
        Exit Function
    End If
    
    'ͣ�õ�Ҳ���Ի��պ�ȡ������
    If Nvl(rsTemp!ͣ������, "3000-01-01") < "3000-01-01" Then
        ShowMsgbox "����Ϊ:" & str���� & " ��" & mCardType.str������ & "�Ѿ�ֹͣʹ�ã�����������˿"
        Exit Function
    End If
    
    dbl��� = Val(Nvl(rsTemp!���))
    dblʧЧ��� = 0
    '���Ч��
    If Nvl(rsTemp!��Ч��, "3000-01-01 00:00:00") < Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") Then
        If Val(Nvl(rsTemp!�ɷ��ֵ)) = 1 Then
            '�����ֵ�ģ����ڵģ������˿�
            dblʧЧ��� = zlGetʧЧ���(Val(Nvl(rsTemp!id)))
            dbl��� = dbl��� - dblʧЧ���
            If dbl��� <= 0 Then dbl��� = 0
        End If
    End If
    
    If dbl��� <= 0 Then
        ShowMsgbox "����Ϊ:" & str���� & " ��" & mCardType.str������ & "��ǰ�������ܽ�������˿"
        Exit Function
    End If
    If blnSaveAfter Then CardIsValid = True: Exit Function
    
    If bytMode = 1 Then
        mlng��ID = Val(Nvl(rsTemp!id))
    Else
        txt(txt_����).Text = str����
    End If
    lbl(lbl_���).Caption = lbl(lbl_���).Tag & Format(Val(Nvl(rsTemp!���)), "0.00") & _
        IIf(dblʧЧ��� > 0, "(��ʧЧ:" & Format(dblʧЧ���, "0.00") & ")", "")
    
'    If nvl(rsTemp!����) = "" Then
'        lblPatiInfo.Caption = ""
'    Else
'        lblPatiInfo.Caption = "������" & nvl(rsTemp!����) & " �Ա�" & nvl(rsTemp!�Ա�) & " ���䣺" & nvl(rsTemp!����)
'    End If
    
    CardIsValid = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cbo֧����ʽ_Click()
    On Error GoTo ErrHandler
    If mblnNotClick Then Exit Sub
    If mlngPre֧����ʽ = cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex) Then Exit Sub
    mlngPre֧����ʽ = cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex)
    
    txt(txt_���).Text = ""
    Call SetControlProperty
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetControlProperty(Optional ByVal blnLoadDefault As Boolean)
    '���ÿؼ�����
    '���:
    '   blnLoadDefault-�Ƿ����ȱʡֵ
    Dim objCard As Card
    Dim blnEnabled As Boolean
    Dim dblTemp As Double, dblMoney As Double
    
    On Error GoTo ErrHandler
    Set objCard = GetCurCard()
    
    '֧Ʊ��һ��ͨ���ϰ�һ��ͨ��������ɿλ
    '1-�ֽ���㷽ʽ,2-������ҽ������,3-ҽ�������ʻ�,4-ҽ������ͳ��,5-���տ���,6-�����ۿ�,7-һ��ͨ����,8-���㿨����
    blnEnabled = InStr(",2,7,8,", "," & objCard.�������� & ",") > 0
    txt(txt_������).Enabled = objCard.�������� <> 1
    txt(txt_�ʺ�).Enabled = objCard.�������� <> 1
    txt(txt_�������).Enabled = objCard.�������� <> 1
    If objCard.�������� = 1 Then
        txt(txt_������).Text = ""
        txt(txt_�ʺ�).Text = ""
        txt(txt_�������).Text = ""
        
        dblMoney = CentMoney(mCurCardPay.dbl��ǰδ��, mBytMoney)
    Else
        dblMoney = RoundEx(mCurCardPay.dbl��ǰδ��, 2)
    End If
    mCurCardPay.dbl������� = mCurCardPay.dbl��ǰδ�� - dblMoney
    
    Call zl_SetCtlBackColor(Array(txt(txt_������), txt(txt_�ʺ�), txt(txt_�������)), Me)
                
    'ȱʡ��������
    txt(txt_���).Locked = False
    If objCard.�ӿ���� > 0 Then '��������
        txt(txt_���).Text = Format(dblMoney, "0.00")
        txt(txt_���).Locked = True
    ElseIf objCard.�������� = 1 Then '�ֽ���
        txt(txt_���).Text = Format(dblMoney, "0.00")
    Else
        txt(txt_���).Text = Format(dblMoney, "0.00")
        txt(txt_���).Locked = True
    End If
    lbl(lbl_���).Caption = FormatEx(mCurCardPay.dbl�������, 6, , , 2)
    lbl(lbl_���).Visible = Val(lbl(lbl_���).Caption) <> 0
    lbl(lbl_���).Caption = "������" & lbl(lbl_���).Caption
    lbl(lbl_��ǰδ��).Caption = Format(mCurCardPay.dbl��ǰδ�� - mCurCardPay.dbl�������, "0.00")
    
    '�����Ҳ�
    Call SetLblCaption
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo֧����ʽ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub cmdOK_Click()
    Dim lngRow As Long
    Dim objCard As Card, lng�����ID As Long
    Dim str�˿���Ϣ As String, lng������� As Long
    Dim dblDelMoney As Double, str������� As String
    
    On Error GoTo ErrHandler
    If mlng��ID = 0 Then
        ShowMsgbox "����ȷ¼�뿨�ţ�"
        zlControl.ControlSetFocus txt(txt_����)
        Exit Sub
    End If
    If CardIsValid(0, mlng��ID, , True) = False Then Exit Sub
    If Check�ɿ���� = False Then Exit Sub
    
    lng������� = zlDatabase.GetNextId("���˿������¼")
    '�������˻ز���
    With vsfBlance
        For lngRow = 1 To .Rows - 1
            lng�����ID = Val(.Cell(flexcpData, lngRow, .ColIndex("����")))
            dblDelMoney = Val(.Cell(flexcpData, lngRow, .ColIndex("�˿���")))
            If lng�����ID = 0 Then Exit For
            If Val(.Cell(flexcpChecked, lngRow, .ColIndex("����"))) <> 1 And dblDelMoney > 0 Then
                Set objCard = GetCurCard(lng�����ID)
                str������� = .Cell(flexcpData, lngRow, .ColIndex("���㷽ʽ"))
                
                mCurCardPay.strԭ������� = ""
                mCurCardPay.lng�����ID = lng�����ID
                mCurCardPay.strˢ������ = .TextMatrix(lngRow, .ColIndex("����"))
                mCurCardPay.strˢ������ = ""
                mCurCardPay.str������ˮ�� = .TextMatrix(lngRow, .ColIndex("������ˮ��"))
                mCurCardPay.str����˵�� = .TextMatrix(lngRow, .ColIndex("����˵��"))
                If CheckThreeSwapIsValied(objCard, dblDelMoney, str�������) = False Then GoTo ErrCheckDelAll
                If SaveData(objCard, dblDelMoney, str�������, lng�������) = False Then GoTo ErrCheckDelAll
                
                str�˿���Ϣ = str�˿���Ϣ & vbCrLf & _
                    .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) & ":" & Format(dblDelMoney, "0.00")
            End If
        Next
    End With
    
    '���ֲ��֣�����֧��ת�˼�����
    str������� = ""
    With vsfBlance
        For lngRow = 1 To .Rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("�˿���"))) > 0 Then
                If Val(.Cell(flexcpData, lngRow, .ColIndex("����"))) = 0 Then
                    str������� = str������� & "," & .Cell(flexcpData, lngRow, .ColIndex("���㷽ʽ"))
                ElseIf Val(.Cell(flexcpChecked, lngRow, .ColIndex("����"))) = 1 Then
                    str������� = str������� & "," & .Cell(flexcpData, lngRow, .ColIndex("���㷽ʽ"))
                End If
            End If
        Next
        If str������� <> "" Then str������� = Mid(str�������, 2)
    End With
    Set objCard = GetCurCard()
    mCurCardPay.strԭ������� = ""
    mCurCardPay.lng�����ID = IIf(objCard.�ӿ���� > 0, objCard.�ӿ����, 0)
    mCurCardPay.strˢ������ = ""
    mCurCardPay.strˢ������ = ""
    mCurCardPay.str������ˮ�� = ""
    mCurCardPay.str����˵�� = ""
    If CheckThreeSwapIsValied(objCard, mCurCardPay.dbl��ǰδ��, str�������) = False Then GoTo ErrCheckDelAll
    If SaveData(objCard, mCurCardPay.dbl��ǰδ��, str�������, lng�������, mCurCardPay.dbl�������) = False Then GoTo ErrCheckDelAll
    
    mintSucces = mintSucces + 1
    Unload Me
    Exit Sub
ErrCheckDelAll:
    '�����;ʧ�ܣ�����Ҫˢ�½���
    If str�˿���Ϣ <> "" Then
        If MsgBox("�ѳɹ��˿�����£�" & str�˿���Ϣ & vbCrLf & "�Ƿ��ʣ��δ�ɹ����ּ����˷ѣ�", _
            vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            If CardIsValid(0, mlng��ID) Then
                If LoadCardData(mlng��ID) = False Then
                    mlng��ID = 0: Call ClearData
                    zlControl.ControlSetFocus txt(txt_����): Exit Sub
                End If
            Else
                mlng��ID = 0: Call ClearData
                zlControl.ControlSetFocus txt(txt_����): Exit Sub
            End If
        Else
            mintSucces = mintSucces + 1
            Unload Me
        End If
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SaveData(ByVal objCard As Card, ByVal dblDelMoney As Double, _
     ByVal str������� As String, ByVal lng������� As Long, _
     Optional ByVal dbl���� As Double) As Boolean
    '��������
    '��Σ�
    '   str������� - ����ö��ŷָ�
    Dim strSQL As String, blnTrain As Boolean
    
    On Error GoTo ErrHandler
    'Zl_���ѿ���Ϣ_����˿�
    strSQL = "Zl_���ѿ���Ϣ_����˿�("
    '  ���ѿ�id_In   ���ѿ���Ϣ.Id%Type,
    strSQL = strSQL & "" & mlng��ID & ","
    '  �������_In   Varchar2,
    strSQL = strSQL & "'" & str������� & "',"
    '  ���㷽ʽ_In   �ʻ��ɿ����.���㷽ʽ%Type,
    strSQL = strSQL & "'" & objCard.���㷽ʽ & "',"
    '  �˿���_In   ���˿������¼.ʵ�ս��%Type,
    strSQL = strSQL & "" & dblDelMoney & ","
    '  �����_In   ���˿������¼.ʵ�ս��%Type,
    strSQL = strSQL & "" & dbl���� & ","
    '  �˿�ʱ��_In   ���ѿ���Ϣ.����ʱ��%Type,
    strSQL = strSQL & "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
    '  ����Ա���_In ���˿������¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '  ����Ա����_In ���˿������¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  �������_In   ���˿������¼.�������%Type,
    strSQL = strSQL & "" & lng������� & ","
    '  ������_In       ���˿������¼.��λ������%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_������).Text) & "',"
    '  �ʺ�_In         ���˿������¼.��λ�ʺ�%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_�ʺ�).Text) & "',"
    '  �������_In   ���˿������¼.�������%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_�������).Text) & "',"
    '  �����id_In   ���˿������¼.�����id%Type := Null,
    strSQL = strSQL & "" & IIf(mCurCardPay.lng�����ID = 0, "NULL", mCurCardPay.lng�����ID) & ","
    '  ���㿨��_In   ���˿������¼.���㿨��%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng�����ID = 0, "NULL", "'" & mCurCardPay.strˢ������ & "'") & ","
    '  ������ˮ��_In ���˿������¼.������ˮ��%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng�����ID = 0, "NULL", "'" & mCurCardPay.str������ˮ�� & "'") & ","
    '  ����˵��_In   ���˿������¼.����˵��%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng�����ID = 0, "NULL", "'" & mCurCardPay.str����˵�� & "'") & ","
    '  �ɿ�_In       ���˿������¼.�ɿ�%Type := Null,
    strSQL = strSQL & "" & IIf(objCard.�������� = 1, -1 * Round(Val(txt(txt_���).Text), 4), "NULL") & ","
    '  �Ҳ�_In       ���˿������¼.�Ҳ�%Type := Null
    strSQL = strSQL & "" & IIf(objCard.�������� = 1, -1 * Round(Val(txt(txt_�Ҳ�).Tag), 4), "NULL") & ")"

    blnTrain = True
    gcnOracle.BeginTrans
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    '����������
    If objCard.�ӿ���� > 0 Then
        If ExecuteThreeSwapPay(objCard, lng�������, dblDelMoney) = False Then Exit Function
    Else
        gcnOracle.CommitTrans
    End If
    blnTrain = False
    SaveData = True
    Exit Function
ErrHandler:
    If blnTrain Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetLblCaption()
    '�����Ҳ�����ʾ
    Dim dbl�Ҳ� As Double
    
    On Error GoTo ErrHandler
    dbl�Ҳ� = RoundEx(Val(txt(txt_���).Text) - (mCurCardPay.dbl��ǰδ�� - mCurCardPay.dbl�������), 6)
    txt(txt_�Ҳ�).Tag = dbl�Ҳ�
    txt(txt_�Ҳ�).Text = Format(-1 * dbl�Ҳ�, "0.00")
    lbl(lbl_�Ҳ�).ForeColor = IIf(dbl�Ҳ� <= 0, vbBlack, vbRed)
    txt(txt_�Ҳ�).ForeColor = IIf(dbl�Ҳ� <= 0, vbBlack, vbRed)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function LoadCardData(ByVal lng��ID As Long) As Boolean
    '���ؽɿ���ϸ���ݵ��ؼ�
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim lngRow As Long
    Dim objCard As Card
    
    On Error GoTo ErrHandler
    If lng��ID = 0 Then Exit Function
    
    '������ǰ�����ݲ���������˿�
    strSQL = _
        "Select a.���㷽ʽ, a.�����id, a.����, a.������ˮ��, a.����˵��," & vbNewLine & _
        "       Sum(a.���) As ���, a.����, Sum(a.ʵ�ʽɿ�) As �˿���," & vbNewLine & _
        "       f_List2str(Cast(Collect(To_Char(�������)) As t_Strlist)) As �������" & vbNewLine & _
        "From �ʻ��ɿ���� A" & vbNewLine & _
        "Where a.���� = 1 And Nvl(a.��Ч��, Sysdate) >= Sysdate And a.���ѿ�id = [1] And a.������� > 0" & vbNewLine & _
        "Group By a.���㷽ʽ, a.�����id, a.����, a.������ˮ��, a.����˵��, a.����" & vbNewLine & _
        "Order By a.�����id, ���㷽ʽ"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��ID)

    If rsTemp.EOF Then
        ShowMsgbox "��ǰ���޿�����"
        Exit Function
    End If
    
    With vsfBlance
        .Redraw = flexRDNone
        .Clear 1
        .Rows = rsTemp.RecordCount + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) = Nvl(rsTemp!���㷽ʽ)
            .Cell(flexcpData, lngRow, .ColIndex("���㷽ʽ")) = Nvl(rsTemp!�������)
            .TextMatrix(lngRow, .ColIndex("���")) = Format(Nvl(rsTemp!���), "0.00")
            .TextMatrix(lngRow, .ColIndex("����")) = Format(Nvl(rsTemp!����), "0.00") & "%"
            .TextMatrix(lngRow, .ColIndex("�˿���")) = Format(Val(Nvl(rsTemp!�˿���)), "0.00")
            .Cell(flexcpData, lngRow, .ColIndex("�˿���")) = Val(Nvl(rsTemp!�˿���))
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
            .Cell(flexcpData, lngRow, .ColIndex("����")) = Val(Nvl(rsTemp!�����id))
            .TextMatrix(lngRow, .ColIndex("������ˮ��")) = Nvl(rsTemp!������ˮ��)
            .TextMatrix(lngRow, .ColIndex("����˵��")) = Nvl(rsTemp!����˵��)
            If Val(Nvl(rsTemp!�����id)) > 0 Then
                Set objCard = GetCurCard(Val(Nvl(rsTemp!�����id)))
                If objCard.�Ƿ����� And objCard.�Ƿ�ȱʡ���� Then
                    .Cell(flexcpChecked, lngRow, .ColIndex("����")) = 1
                Else
                    .Cell(flexcpChecked, lngRow, .ColIndex("����")) = 2
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbBlue
                End If
            End If
            
            rsTemp.MoveNext
            lngRow = lngRow + 1
        Loop
        .Redraw = flexRDBuffered
    End With
    
    Call Calc�˿���(True)

    LoadCardData = True
    Exit Function
ErrHandler:
    vsfBlance.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��" & Chr(22), Chr(KeyAscii)) > 0 Then 'Chr(22):Ctrl+V
        KeyAscii = 0: Exit Sub
    End If
End Sub

Private Sub Form_Load()
    mblnFirst = True
    
    Call ClearData
    If InitData() = False Then Unload Me: Exit Sub
    If Load֧����ʽ() = False Then Unload Me: Exit Sub
    
    If mlng��ID <> 0 Then
        If CardIsValid(0, mlng��ID) Then
            If LoadCardData(mlng��ID) = False Then
                mlng��ID = 0
                txt(txt_����).Text = ""
            End If
        Else
            mlng��ID = 0
        End If
    End If
    
    pic(pic_�����ϸ).AutoRedraw = True: zlControl.PicShowFlat pic(pic_�����ϸ)
    pic(pic_�ɿ�ϼ�).AutoRedraw = True: zlControl.PicShowFlat pic(pic_�ɿ�ϼ�)
    pic(pic_�ɿ���Ϣ).AutoRedraw = True: zlControl.PicShowFlat pic(pic_�ɿ���Ϣ)
    cbo.SetListWidth cbo֧����ʽ, cbo֧����ʽ.Width * 2
    
    txt(txt_�Ҳ�).BackColor = Me.BackColor
    Me.Caption = "����˿� - " & mCardType.str������
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    On Error Resume Next
    '���㶨λ
    zlControl.ControlSetFocus txt(txt_����)
End Sub

Private Function InitData() As Boolean
    '��ʼ��ģ�����
    Dim rsTemp As New ADODB.Recordset
    Dim ty_Temp As Ty_CardType
    Dim strValue As String
    
    On Error GoTo ErrHandler
    Set rsTemp = zlGet���ѿ��ӿ�()
    rsTemp.Filter = "���=" & mlng�����
    If rsTemp.EOF Then
        ShowMsgbox "δ���ֿ������Ϣ�����ܼ�����"
        Exit Function
    End If
    
    '���ѿ��ֱҴ���ʽ
    strValue = zlDatabase.GetPara(14, glngSys, , 0)
    mBytMoney = Val(IIf(Len(strValue) = 1, strValue, Mid(strValue, 4, 1)))
    
    mCardType = ty_Temp '�Զ���Type��ʼ��
    With mCardType
        .str������ = Nvl(rsTemp!����)
        .str����ǰ׺ = Nvl(rsTemp!ǰ׺�ı�)
        .lng���ų��� = Val(Nvl(rsTemp!���ų���))
        .bln�������� = Val(Nvl(rsTemp!�Ƿ�����)) = 1
        .bln�ϸ���� = Val(Nvl(rsTemp!�Ƿ��ϸ����)) = 1
        .str������� = Nvl(rsTemp!�������)
        .bln�ض����� = Val(Nvl(rsTemp!�Ƿ��ض�����)) = 1
    End With
    
    If Init֧����ʽ() = False Then Exit Function
    
    InitData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txt_Change(Index As Integer)
    If mblnNotClick Then Exit Sub
    
    On Error GoTo ErrHandler
    Select Case Index
    Case txt_���
        Call SetLblCaption
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index <> txt_���� Then
        If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
    Case txt_����
        Call zlControl.TxtCheckKeyPress(txt(Index), KeyAscii, m�ı�ʽ)
        If InStr(1, "'~��|`-'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        Call BrushCard(txt(Index), KeyAscii)
    Case txt_���
        Call zlControl.TxtCheckKeyPress(txt(Index), KeyAscii, m���ʽ)
    Case Else
        Call zlControl.TxtCheckKeyPress(txt(Index), KeyAscii, m�ı�ʽ)
    End Select
End Sub

Private Sub BrushCard(ByVal objEdit As Object, KeyAscii As Integer)
    'ˢ��
    Static sngBegin As Single
    Dim sngNow As Single
    Dim blnCard As Boolean
    Dim lng������ As Long
    
    On Error GoTo ErrHandler
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    blnCard = zlCommFun.InputIsCard(objEdit, KeyAscii, False)
    If blnCard And Len(objEdit.Text) = mCardType.lng���ų��� - 1 And KeyAscii <> 8 _
        Or KeyAscii = 13 And Trim(objEdit.Text) <> "" Then '�ﵽ���ų��Ȼ�س����ҿ���Ϣ
        
        If KeyAscii <> 13 Then
            objEdit.Text = objEdit.Text & Chr(KeyAscii)
            objEdit.SelStart = Len(objEdit.Text)
        End If
        KeyAscii = 0
        
        If CardIsValid(1, , objEdit.Text) Then
            If LoadCardData(mlng��ID) = False Then
                mlng��ID = 0: Call ClearData
                zlControl.TxtSelAll objEdit: Exit Sub
            End If
        Else
            mlng��ID = 0: Call ClearData
            zlControl.TxtSelAll objEdit: Exit Sub
        End If
        zlCommFun.PressKey vbKeyTab
    ElseIf KeyAscii = 13 And Trim(objEdit.Text) = "" Then
        zlCommFun.PressKey vbKeyTab
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        '��ȫˢ�����
        If KeyAscii <> 0 And KeyAscii > 32 Then
            sngNow = timer
            If objEdit.Text = "" Then
                sngBegin = sngNow
            ElseIf Format((sngNow - sngBegin) / (Len(objEdit) + 1), "0.000") >= 0.04 Then '>0.007>=0.01
                objEdit.Text = Chr(KeyAscii)
                objEdit.SelStart = 1
                KeyAscii = 0
                sngBegin = sngNow
            End If
        End If
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Function Init֧����ʽ() As Boolean
    '��ʼ��֧����ʽ
    '˵����
    '   ֻ�����ֽ�֧Ʊ���������Ľ��㷽ʽ
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim i As Long, objCards As Cards, objCard As Card
    Dim lngKey As Long
    
    On Error GoTo ErrHandler
    Set mobjPayCards = New Cards
    
    Set rsTemp = Get���㷽ʽ("���ѿ�")
    If Not gobjSquare Is Nothing Then
        ' zlGetCards(ByVal BytType As Byte)
            '���:bytType-  0-����ҽ�ƿ�;
        '                        1-���õ�ҽ�ƿ�,
        '                        2-���д��������˻���������
        '                        3-���õ������˻���ҽ�ƿ�
       Set objCards = gobjSquare.objSquareCard.zlGetCards(0)
    End If
    
    With rsTemp
        .Filter = 0
        If .RecordCount <> 0 Then .MoveFirst
        lngKey = 1
        Do While Not .EOF
            For i = 1 To objCards.count
                If objCards(i).���㷽ʽ = Nvl(rsTemp!����) Then blnFind = True: Exit For
            Next
            If Not blnFind Then
                If (Val(Nvl(rsTemp!����)) = 1 Or Val(Nvl(rsTemp!����)) = 2) _
                    And Val(Nvl(rsTemp!Ӧ����)) = 0 Then
                    Set objCard = New Card
                    objCard.���� = Mid(Nvl(!����), 1, 1)
                    objCard.�ӿڱ��� = Nvl(!����)
                    objCard.�ӿڳ����� = ""
                    objCard.�ӿ���� = -1 * lngKey
                    objCard.���㷽ʽ = Nvl(!����)
                    objCard.���� = Nvl(!����)
                    objCard.���� = True
                    objCard.ȱʡ��־ = Val(Nvl(rsTemp!ȱʡ)) = 1
                    objCard.���� = True
                    objCard.�������� = Val(!����)
                    
                    mobjPayCards.Add objCard, "K" & lngKey
                    lngKey = lngKey + 1
              End If
            End If
            .MoveNext
        Loop
    End With
    
    '��������
    For i = 1 To objCards.count
        rsTemp.Filter = "����='" & objCards(i).���㷽ʽ & "'"
        If Not rsTemp.EOF And objCards(i).���� And Not objCards(i).���ѿ� Then
            mobjPayCards.Add objCards(i), "K" & lngKey
            lngKey = lngKey + 1
        End If
    Next
    
    If mobjPayCards.count = 0 Then
        ShowMsgbox "���ѿ�����û�п��õĽ��㷽ʽ�����ȵ������㷽ʽ���������á�"
        Exit Function
    End If
    Init֧����ʽ = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Load֧����ʽ() As Boolean
    '����֧����ʽ
    '˵��:
    '   ȱʡ���㷽ʽ�Ĺ�������˳�����£�
    '   1.���㷽ʽӦ�������õ�ȱʡ��
    '   2.����Ϊ"1-�ֽ���㷽ʽ"�Ľ��㷽ʽ
    Dim objCard As Card, i As Long
    Dim str���㷽ʽ As String
    
    On Error GoTo ErrHandler
    mlngPre֧����ʽ = 0

    mblnNotClick = True
    With cbo֧����ʽ
        .Clear
        For i = 1 To mobjPayCards.count
            Set objCard = mobjPayCards(i)
            If objCard.���� And Not objCard.���ѿ� And InStr(str���㷽ʽ & "|", "|" & objCard.���㷽ʽ & "|") = 0 Then
                '�����˻���֧����ʽ��ʾΪҽ�ƿ����ƣ�������ʾ���㷽ʽ
                If objCard.�ӿ���� > 0 Then
                    If objCard.�Ƿ�ת�ʼ����� Then
                        .AddItem objCard.����
                        .ItemData(.NewIndex) = i
                    End If
                Else
                    .AddItem objCard.���㷽ʽ
                    .ItemData(.NewIndex) = i
                End If
                
                str���㷽ʽ = str���㷽ʽ & "|" & objCard.���㷽ʽ
            End If
            
            '����ȱʡֵ
            If objCard.ȱʡ��־ And .ListIndex < 0 Then .ListIndex = .NewIndex
            If objCard.�������� = 1 And .ListIndex < 0 Then .ListIndex = .NewIndex
        Next
            
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
    End With
    mblnNotClick = False
    Call cbo֧����ʽ_Click
    Load֧����ʽ = True
    Exit Function
ErrHandler:
    mblnNotClick = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetCurCard(Optional ByVal lngҽ�ƿ�ID As Long) As Card
    '��ȡ��ǰ֧���������󣬻�ͨ�������ID��ȡ������
    Dim intIndex As Integer
    Dim objCard As Card
    
    On Error GoTo ErrHandler
    If lngҽ�ƿ�ID = 0 Then
        If cbo֧����ʽ.ListIndex <> -1 Then
            intIndex = cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex)
            If intIndex <= 0 Then Exit Function
            Set objCard = mobjPayCards(intIndex)
        End If
    Else
        If Not mobjPayCards Is Nothing Then
            For Each objCard In mobjPayCards
                If objCard.�ӿ���� = lngҽ�ƿ�ID Then Exit For
            Next
        End If
    End If
    If objCard Is Nothing Then Set objCard = New Card
    Set GetCurCard = objCard
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check�ɿ����() As Boolean
    '����:���ɿ����
    Dim objCard As Card
    Dim strTitle As String
    
    On Error GoTo ErrHandler
    If mCurCardPay.dbl�˿�ϼ� <= 0 Then
        ShowMsgbox "��ǰ���޿��������ܽ�������˿"
        Exit Function
    End If
    
    If cbo֧����ʽ.ListIndex = -1 Then
        ShowMsgbox "��ǰ�˿ʽδѡ�����飡"
        zlControl.ControlSetFocus cbo֧����ʽ
        Exit Function
    End If
    
    If zlDblIsValid(Trim(txt(txt_���).Text), 16, True, False, txt(txt_���).hWnd, strTitle) = False Then Exit Function
    
    Set objCard = GetCurCard()
    If objCard.�������� <> 1 Then
        If RoundEx(mCurCardPay.dbl��ǰδ�� - mCurCardPay.dbl�������, 6) = 0 Then
            ShowMsgbox "��ǰδ�˿���Ϊ�㣬����ʹ�÷��ֽ���㷽ʽ��"
            zlControl.ControlSetFocus cbo֧����ʽ
            Exit Function
        End If
        
        If Val(txt(txt_���).Text) = 0 Then
            ShowMsgbox "δ�����˿�����飡"
            zlControl.ControlSetFocus txt(txt_���)
            Exit Function
        End If
    End If
    If Val(txt(txt_���).Text) <> 0 Then
        If Val(txt(txt_���).Text) < RoundEx(mCurCardPay.dbl��ǰδ�� - mCurCardPay.dbl�������, 6) Then
            ShowMsgbox "�˿���(" & Format(Val(txt(txt_���).Text), "0.00") & ")���㱾��δ�˽��(" & _
                Format(RoundEx(mCurCardPay.dbl��ǰδ�� - mCurCardPay.dbl�������, 6), "0.00") & ")�����飡"
            zlControl.ControlSetFocus txt(txt_���)
            Exit Function
        End If
        
        If objCard.�������� <> 1 And Val(txt(txt_���).Text) > RoundEx(mCurCardPay.dbl��ǰδ�� - mCurCardPay.dbl�������, 6) Then
            ShowMsgbox "�˿���(" & Format(Val(txt(txt_���).Text), "0.00") & ")�����˱���δ�˽��(" & _
                Format(RoundEx(mCurCardPay.dbl��ǰδ�� - mCurCardPay.dbl�������, 6), "0.00") & ")�����飡"
            zlControl.ControlSetFocus txt(txt_���)
            Exit Function
        End If
    End If
    
    If zlCommFun.StrIsValid(Trim(txt(txt_������).Text), 50, txt(txt_������).hWnd, "������") = False Then Exit Function
    If zlCommFun.StrIsValid(Trim(txt(txt_�ʺ�).Text), 20, txt(txt_�ʺ�).hWnd, "�ʺ�") = False Then Exit Function
    If zlCommFun.StrIsValid(Trim(txt(txt_�������).Text), 30, txt(txt_�������).hWnd, "�������") = False Then Exit Function
    Check�ɿ���� = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckThreeSwapIsValied(ByVal objCard As Card, ByVal dblDelMoney As Double, _
    ByVal str������� As String) As Boolean
    '����:������ˢ����֤
    '���:objCard-��ǰ��
    '    str������� - ����ö��ŷָ������ڻ�ȡԭ�������
    '����:ˢ���ɹ�,����true,���򷵻�False
    Dim strXMLExpend As String, strBalanceIDs As String
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If objCard.�ӿ���� <= 0 Then CheckThreeSwapIsValied = True: Exit Function
    If dblDelMoney = 0 Then CheckThreeSwapIsValied = True: Exit Function
    
    If objCard.���� = False Then
        ShowMsgbox objCard.���� & "δ���ã���˲����˻أ������ѡ�����֣�"
        Exit Function
    End If
    
    '��ȡԭ������ţ��Լ�ȫ�˼��
    mCurCardPay.strԭ������� = ""
    strSQL = _
        "Select /*+cardinality(j,10)*/Nvl(Sum(a.ʵ�ս��), 0) As �ɿ�ϼ�," & vbNewLine & _
        "       f_List2str(Cast(Collect(Distinct To_Char(b.�������)) As t_Strlist)) As �������" & vbNewLine & _
        "From ���˿������¼ A, ���˿������¼ B, Table(f_Num2list([1])) J" & vbNewLine & _
        "Where a.������� = b.������� And b.������� = j.Column_Value And a.��¼���� In (1, 2)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str�������)
    If rsTemp.EOF = False Then
        mCurCardPay.strԭ������� = Nvl(rsTemp!�������)
        If objCard.�Ƿ�ת�ʼ����� = False And objCard.�Ƿ�ȫ�� Then
            If Val(Nvl(rsTemp!�ɿ�ϼ�)) <> dblDelMoney Then
                ShowMsgbox objCard.���� & "��֧�ֲ����ˣ���˲����˻أ������ѡ�����֣�" & _
                    "(ԭ֧����" & FormatEx(Val(Nvl(rsTemp!�ɿ�ϼ�)), 2) & _
                    "�����˿��" & FormatEx(dblDelMoney, 2) & ")"
                Exit Function
            End If
        End If
    End If
    
    If objCard.�Ƿ�ת�ʼ����� Then
        '   zlBrushCard(frmMain As Object, _
        ByVal lngModule As Long, _
        ByVal rsClassMoney As ADODB.Recordset, _
        ByVal lngCardTypeID As Long, _
        ByVal bln���ѿ� As Boolean, _
        ByVal strPatiName As String, ByVal strSex As String, _
        ByVal strOld As String, ByRef dbl��� As Double, _
        Optional ByRef strCardNo As String, _
        Optional ByRef strPassWord As String, _
        Optional ByRef bln�˷� As Boolean = False, _
        Optional ByRef blnShowPatiInfor As Boolean = False, _
        Optional ByRef bln���� As Boolean = False, _
        Optional ByVal bln�����ֹ As Boolean = True, _
        Optional ByRef varSquareBalance As Variant, _
        Optional ByVal blnתԤ�� As Boolean = False, _
        Optional ByVal blnAllPay As Boolean = False, _
        Optional ByVal strXmlIn As String = "") As Boolean
        '       strXmlIn-����������XML���,Ŀǰ��ʽ����:
        '       <IN>
        '           <CZLX>0</CZLX>    //��������,0-��������ˢ��,1-ת�˵���ˢ��,2-�˿����ˢ��
        '       </IN>
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, _
            objCard.�ӿ����, False, "", "", "", -1 * dblDelMoney, mCurCardPay.strˢ������, mCurCardPay.strˢ������, _
            False, False, False, True, Nothing, False, True, "<IN><CZLX>1</CZLX></IN>") = False Then Exit Function
        
        '����ת�ʽӿ�
        'zlTransferAccountsCheck ת�ʼ��ӿ�
        '������  ��������    ��/��   ��ע
        'frmMain Object  In  ���õ�������
        'lngModule   Long    In  HIS����ģ���
        'lngCardTypeID   Long    In  �����ID
        'strCardNo   String  In  ����
        'dblDelMoney    Double  In  ת�ʽ��(����ʱΪ����)
        'strBalanceID    String  In  ԭ֧���������,���ò����¼.������Ż���Ԥ����¼.�������
        'strXMLExpend String In   XML��:
        '                            <IN>
        '                                <CZLX>��������</CZLX> //0��NULL:������ҵ��;1-�������˷�ҵ��
        '                                       2-����ҵ��;3-�����˷�ҵ��4-�����˷�ҵ��5-���ѿ������˷�ҵ��
        '                            </IN>
        '                    Out  XML��:
        '                            <OUT>
        '                               <ERRMSG>������Ϣ</ERRMSG >
        '                            </OUT>
        '    Boolean ��������    �������ݺϷ�,����True:���򷵻�False
        '˵��:
        '��. ������ת��ʱ��һЩ�Ϸ��Լ�飬������ת��ʱ�����Ի���֮��ĵȴ������������������ķ�����
        '��. �����ڼ�����Ҫ����ΪTrue�����������ת�ʹ��ܵĵ��á�
        '����XML��
        strXMLExpend = "<IN><CZLX>5</CZLX></IN>"
        If gobjSquare.objSquareCard.zlTransferAccountsCheck(Me, mlngModule, objCard.�ӿ����, _
            mCurCardPay.strˢ������, dblDelMoney, mCurCardPay.strԭ�������, strXMLExpend) = False Then
            Call ShowThreeSwapErrMsg(0, strXMLExpend)
            Exit Function
        End If
    Else
        'zlReturnCheck(frmMain As Object, ByVal lngModule As Long, _
            ByVal lngCardTypeID As Long, bln���ѿ� As Boolean, ByVal strCardNo As String, _
            ByVal strBalanceIDs As String, _
            ByVal dblDelMoney As Double, ByVal strSwapNo As String, _
            ByVal strSwapMemo As String, ByRef strXMLExpend As String) As Boolean
            '---------------------------------------------------------------------------------------------------------------------------------------------
            '����:�ʻ����˽���ǰ�ļ��
            '���:frmMain-���õ�������
            '       lngModule-���õ�ģ���
            '       lngCardTypeID-�����ID
            '       strCardNo-����
            '       strBalanceIDs   String  In  ����֧�����漰�Ľ���ID ��ʽ:�շ�����|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
            '                                   �շ�����: 1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�,6-���ղ������,7-���ѿ��տ
            '                                           ����=7ʱ��IDΪ���˿������¼.�������
            '       dblDelMoney-�˿���
            '       strSwapNo-������ˮ��(�˿�ʱ���)
            '       strSwapMemo-����˵��(�˿�ʱ����)
            '       strXMLExpend    XML IN  ��ѡ����(��չ��):
            '        <TFDATA> //�˷�����
            '          <YCTF>1</YCTF> //�Ƿ��쳣����:1-�쳣����;0-�˷� �˽ڵ����û��
            '          <TFLIST> //�˷��б�
            '            <NO></NO> // �˷ѵ���
            '            <TFITEM> //�˷���
            '              <SerialNum></SerialNum> //���
            '              ��
            '            </TFITEM>
            '          </TFLIST>
            '          ....
            '        </TFDATA >
            '����:�˿�Ϸ�,����true,���򷵻�Flase
        strBalanceIDs = "7|" & mCurCardPay.strԭ�������
        If gobjSquare.objSquareCard.zlReturncheck(Me, mlngModule, objCard.�ӿ����, _
            objCard.���ѿ�, mCurCardPay.strˢ������, strBalanceIDs, dblDelMoney, _
            mCurCardPay.str������ˮ��, mCurCardPay.str����˵��, strXMLExpend) = False Then Exit Function
    
        If objCard.�Ƿ��˿��鿨 Then
           '����ˢ������
            'zlBrushCard(frmMain As Object, _
            ByVal lngModule As Long, _
            ByVal rsClassMoney As ADODB.Recordset, _
            ByVal lngCardTypeID As Long, _
            ByVal bln���ѿ� As Boolean, _
            ByVal strPatiName As String, ByVal strSex As String, _
            ByVal strOld As String, ByVal dbl��� As Double, _
            Optional ByRef strCardNo As String, _
            Optional ByRef strPassWord As String, _
            Optional ByRef bln�˷� As Boolean = False, _
            Optional ByRef blnShowPatiInfor As Boolean = False, _
            Optional ByRef bln���� As Boolean = False, _
            Optional ByVal bln�����ֹ As Boolean = True, _
            Optional ByRef varSquareBalance As Variant, _
            Optional ByVal blnתԤ�� As Boolean = False, _
            Optional ByVal blnAllPay As Boolean = False, _
            Optional ByVal strXmlIn As String = "") As Boolean
            '       strXmlIn-����������XML���,Ŀǰ��ʽ����:
            '       <IN>
            '           <CZLX>0</CZLX>    //��������,0-��������ˢ��,1-ת�˵���ˢ��,2-�˿����ˢ��
            '       </IN>
            If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, _
                objCard.�ӿ����, False, "", "", "", dblDelMoney, mCurCardPay.strˢ������, mCurCardPay.strˢ������, _
                True, False, False, True, Nothing, False, True, "<IN><CZLX>2</CZLX></IN>") = False Then Exit Function
        End If
    End If
    CheckThreeSwapIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteThreeSwapPay(ByVal objCard As Card, ByVal lng������� As Long, _
    ByVal dblDelMoney As Double) As Boolean
    '����:һ��֧ͨ��(�����ӿ�)
    '���:
    '   objCard-��ǰ��
    '   dblDelMoney-����֧�����
    '����:
    '����:ִ�гɹ�,����true,���򷵻�False
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim strSwapExtendInfor As String, strTemp As String
    Dim strXMLExpend As String
    
    On Error GoTo ErrHandler
    If objCard.�ӿ���� <= 0 Then ExecuteThreeSwapPay = True: Exit Function
    If dblDelMoney = 0 Then ExecuteThreeSwapPay = True: Exit Function
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection

    If objCard.�Ƿ�ת�ʼ����� Then
        'zlTransferAccountsMoney
        '������  ��������    ��/��   ��ע
        'frmMain Object  In  ���õ�������
        'lngModule   Long    In  HIS����ģ���
        'lngCardTypeID   Long    In  �����ID
        'strCardNo   String  In  ����
        'strBalanceID    String  In  ����ID ����֧���������,���ò����¼.������Ż���Ԥ����¼.������Ż��˿������¼.�������
        'dblDelMoney    Double  In  ת�ʽ��
        'strSwapGlideNO  String  Out ������ˮ��
        'strSwapMemo String  Out ����˵��
        'strSwapExtendInfor  String  In �˷�ҵ��ʱ�����뱾���˷ѵĳ���ID:
        '                               ��ʽ:�շ�����1|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
        '                               �շ�����:1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�,6-���ղ������,7-���ѿ������տ�(IDΪ�������)
        '                           Out ������չ��Ϣ: ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n
        'strXMLExpend String In   XML��:
        '                            <IN>
        '                                <CZLX>��������</CZLX> //0��NULL:������ҵ��;1-�������˷�ҵ��
        '                                       2-����ҵ��;3-�����˷�ҵ��4-�����˷�ҵ��5-���ѿ������˷�ҵ��
        '                            </IN>
        '                    Out  XML��:
        '                            <OUT>
        '                               <ERRMSG>������Ϣ</ERRMSG >
        '                            </OUT>
        '    Boolean ��������    True:���óɹ�,False:����ʧ��
        '˵��:
        '��. ��ҽ���������ʱ���е�����ת��ʱ���á�
        '��. һ����˵���ɹ�ת�ʺ󣬶�Ӧ�ô�ӡ��صĽ���Ʊ�ݣ����Է��ڴ˽ӿڽ��д���.
        '��. ��ת�ʳɹ��󣬷��ؽ�����ˮ�ź���ؽ���˵���������������������Ϣ�����Է�����չ��Ϣ�з���.
        '����XML��
        strXMLExpend = "<IN><CZLX>5</CZLX></IN>"
        strSwapExtendInfor = "7|" & mCurCardPay.strԭ�������: strTemp = strSwapExtendInfor
        If gobjSquare.objSquareCard.zlTransferAccountsMoney(Me, mlngModule, objCard.�ӿ����, _
            mCurCardPay.strˢ������, lng�������, dblDelMoney, _
            mCurCardPay.str������ˮ��, mCurCardPay.str����˵��, strSwapExtendInfor, strXMLExpend) = False Then
            gcnOracle.RollbackTrans: Call ShowThreeSwapErrMsg(1, strXMLExpend)
            Exit Function
        End If
        gcnOracle.CommitTrans: ExecuteThreeSwapPay = True
        
        Call zlAddUpdateSwapSQL(False, lng�������, objCard.�ӿ����, objCard.���ѿ�, _
            mCurCardPay.strˢ������, mCurCardPay.str������ˮ��, mCurCardPay.str����˵��, cllUpdate, 1, 0, 1)
        zlExecuteProcedureArrAy cllUpdate, Me.Caption
        If strTemp <> strSwapExtendInfor Then
            Call zlAddThreeSwapSQLToCollection(False, lng�������, objCard.�ӿ����, objCard.���ѿ�, _
                mCurCardPay.strˢ������, strSwapExtendInfor, cllThreeSwap, 0, 1)
            zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
        End If
    Else
        'zlReturnMoney(frmMain As Object, ByVal lngModule As Long, _
            ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByVal strBalanceIDs As String, _
            ByVal dblDelMoney As Double, _
            ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, _
            ByRef strSwapExtendInfor As String) As Boolean
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '����:�ʻ��ۿ���˽���
        '���:frmMain-���õ�������
        '       lngModule-���õ�ģ���
        '       lngCardTypeID-�����ID:ҽ�ƿ����.ID
        '       strCardNo-����
        '       strBalanceIDs-����֧�����漰�Ľ���ID(����ԭ����ID):
        '                           ��ʽ:�շ�����(|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
        '                           �շ�����:1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�,6-���ղ������,7-���ѿ��տ�
        '       dblDelMoney-�˿���
        '       strSwapNo-������ˮ��(�ۿ�ʱ�Ľ�����ˮ��)
        '       strSwapMemo-����˵��(�ۿ�ʱ�Ľ���˵��)
        '       strSwapExtendInfor-���룬�����˷ѵĳ���ID��
        '                           ��ʽ:�շ�����1|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
        '                           �շ�����:1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�,6-���ղ������,7-���ѿ��տ�
        '       strSwapExtendInfor-���������׵���չ��Ϣ
        '           ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n ÿ����Ŀ�в��ܰ���|�ַ�
        strSwapExtendInfor = "7|" & lng�������
        If gobjSquare.objSquareCard.zlReturnMoney(Me, mlngModule, objCard.�ӿ����, objCard.���ѿ�, _
            mCurCardPay.strˢ������, "7|" & mCurCardPay.strԭ�������, dblDelMoney, _
            mCurCardPay.str������ˮ��, mCurCardPay.str����˵��, strSwapExtendInfor) = False Then
            gcnOracle.RollbackTrans: Exit Function
        End If
        gcnOracle.CommitTrans: ExecuteThreeSwapPay = True
        
        Call zlAddUpdateSwapSQL(False, lng�������, objCard.�ӿ����, objCard.���ѿ�, _
            mCurCardPay.strˢ������, mCurCardPay.str������ˮ��, mCurCardPay.str����˵��, cllUpdate, 1, 0, 1)
        zlExecuteProcedureArrAy cllUpdate, Me.Caption
        If strTemp <> strSwapExtendInfor Then
            Call zlAddThreeSwapSQLToCollection(False, lng�������, objCard.�ӿ����, objCard.���ѿ�, _
                mCurCardPay.strˢ������, strSwapExtendInfor, cllThreeSwap, 0, 1)
            zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
        End If
    End If
    ExecuteThreeSwapPay = True
    Exit Function
ErrHandler:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowThreeSwapErrMsg(ByVal bytType As Byte, ByVal strXMLErrMsg As String)
    '����:����ת�˼�������ҵ�������ʾ
    '����:
    '   bytType:0-ת�˼��,1-ת�˽���
    '   strXMLErrMsg:��ʽ����
    '            <OUT>
    '               <ERRMSG>������Ϣ</ERRMSG >
    '            </OUT>
    Dim strValue As String
    
    On Error GoTo errHandle
    '����������Ϣ
    If strXMLErrMsg <> "" Then
        If zlXML.OpenXMLDocument(strXMLErrMsg) = False Then strValue = ""
        If zlXML.GetSingleNodeValue("OUT/ERRMSG", strValue) = False Then strValue = ""
        Call zlXML.CloseXMLDocument
    End If
    '��ʾ������Ϣ
    If Trim(strValue) = "" Then
        If bytType = 0 Then
            strValue = vbCrLf & "���׼��ʧ�ܣ�"
        Else
            strValue = vbCrLf & "����ʧ�ܣ�"
        End If
    End If
    MsgBox strValue, vbExclamation, gstrSysName
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function CheckThreeBalanceToCash(ByVal objCard As Card) As Boolean
    '���������ּ��
    Dim str����Ա As String
    
    On Error GoTo errHandle
    If Not (objCard.�ӿ���� > 0 And Not objCard.���ѿ�) Then CheckThreeBalanceToCash = True: Exit Function
    If objCard.�Ƿ����� Then CheckThreeBalanceToCash = True: Exit Function
    
    If zlstr.IsHavePrivs(GetPrivFunc(glngSys, 1151), "�����˿�ǿ������") Then
        If MsgBox(objCard.���� & "��֧�����֣���ȷ��Ҫ����ǿ��������", _
            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Else
        str����Ա = zlDatabase.UserIdentifyByUser(Me, objCard.���� & "ǿ�����֣�Ȩ����֤��", _
            glngSys, mlngModule, "�����˿�ǿ������", , True)
        If str����Ա = "" Then Exit Function
    End If
    CheckThreeBalanceToCash = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub vsfBlance_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfBlance
        If Col = .ColIndex("����") Then
            If Val(.Cell(flexcpChecked, Row, .ColIndex("����"))) = 1 Then '����
                .Cell(flexcpForeColor, Row, 0, Row, .Cols - 1) = .ForeColor
                .ForeColorSel = .ForeColor
            Else
                .Cell(flexcpForeColor, Row, 0, Row, .Cols - 1) = vbBlue
                .ForeColorSel = vbBlue
            End If
        
            Call Calc�˿���
        End If
    End With
End Sub

Private Sub vsfBlance_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    vsfBlance.ForeColorSel = vsfBlance.Cell(flexcpForeColor, NewRow, 0)
End Sub

Private Sub vsfBlance_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngҽ�ƿ�ID As Long
    With vsfBlance
        If Col <> .ColIndex("����") Then Cancel = True: Exit Sub
        lngҽ�ƿ�ID = Val(.Cell(flexcpData, Row, .ColIndex("����")))
        If lngҽ�ƿ�ID <= 0 Then Cancel = True: Exit Sub
    End With
End Sub

Private Sub vsfBlance_GotFocus()
    If vsfBlance.Row < vsfBlance.FixedRows And vsfBlance.Rows > 1 Then
        vsfBlance.Row = 1
    End If
End Sub

Private Sub vsfBlance_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsfBlance.Col <> vsfBlance.ColIndex("����") Then
        If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub vsfBlance_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngҽ�ƿ�ID As Long, objCard As Card
    
    On Error GoTo errHandle
    With vsfBlance
        If Col <> .ColIndex("����") Then Exit Sub
        lngҽ�ƿ�ID = Val(.Cell(flexcpData, Row, .ColIndex("����")))
        If lngҽ�ƿ�ID <= 0 Then Exit Sub
        
        '���ּ��
        If Val(.TextMatrix(Row, .ColIndex("�˿���"))) = 0 Or Abs(Val(.EditText)) <> 1 Then Exit Sub
        Set objCard = GetCurCard(lngҽ�ƿ�ID)
        If objCard.�Ƿ����� Then Exit Sub
        If CheckThreeBalanceToCash(objCard) = False Then Cancel = True: Exit Sub
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Calc�˿���(Optional ByVal bln����ϼ� As Boolean)
    '�����˿���
    Dim lngRow As Long
    
    On Error GoTo ErrHandler
    If bln����ϼ� Then
        mCurCardPay.dbl�˿�ϼ� = 0
        mCurCardPay.dbl��ǰδ�� = 0
        With vsfBlance
            For lngRow = 1 To .Rows - 1
                mCurCardPay.dbl�˿�ϼ� = mCurCardPay.dbl�˿�ϼ� + Val(.TextMatrix(lngRow, .ColIndex("�˿���")))
            Next
            mCurCardPay.dbl�˿�ϼ� = Round(mCurCardPay.dbl�˿�ϼ�, 6)
        End With
    End If
    
    mCurCardPay.dbl��ǰδ�� = 0
    mCurCardPay.dbl������� = 0
    With vsfBlance
        For lngRow = 1 To .Rows - 1
            If Val(.Cell(flexcpData, lngRow, .ColIndex("����"))) > 0 Then
                If Val(.Cell(flexcpChecked, lngRow, .ColIndex("����"))) = 1 Then
                    mCurCardPay.dbl��ǰδ�� = mCurCardPay.dbl��ǰδ�� + Val(.Cell(flexcpData, lngRow, .ColIndex("�˿���")))
                End If
            Else
                mCurCardPay.dbl��ǰδ�� = mCurCardPay.dbl��ǰδ�� + Val(.Cell(flexcpData, lngRow, .ColIndex("�˿���")))
            End If
        Next
        mCurCardPay.dbl��ǰδ�� = Round(mCurCardPay.dbl��ǰδ��, 6)
    End With
    
    lbl(lbl_�˿�ϼ�).Caption = FormatEx(mCurCardPay.dbl�˿�ϼ�, 6, , , 2)
    
    Call SetControlProperty
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ClearData()
    '�������
    Dim tyPayMoney As TY_PayMoney
    
    On Error GoTo ErrHandler
    mCurCardPay = tyPayMoney
    
    lbl(lbl_���).Caption = lbl(lbl_���).Tag & "0.00"
    lblPatiInfo.Caption = ""
    
    vsfBlance.Clear 1
    vsfBlance.Rows = 1
    
    lbl(lbl_�˿�ϼ�).Caption = "0.00"
    lbl(lbl_��ǰδ��).Caption = "0.00"
    
    txt(txt_���).Text = ""
    Call SetControlProperty
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
