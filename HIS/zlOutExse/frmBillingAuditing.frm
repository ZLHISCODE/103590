VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBillingAuditing 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���˻��۵������"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8595
   Icon            =   "frmBillingAuditing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   90
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   8400
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   3780
      Width           =   8400
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshBill 
      Height          =   2265
      Left            =   30
      TabIndex        =   8
      ToolTipText     =   "˫�����ݲ鿴��ϸ"
      Top             =   3825
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   3995
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmBillingAuditing.frx":058A
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   960
      Left            =   0
      ScaleHeight     =   960
      ScaleWidth      =   8595
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6120
      Width           =   8595
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
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
         Left            =   300
         TabIndex        =   16
         Top             =   525
         Width           =   1200
      End
      Begin VB.CommandButton cmdFlash 
         Caption         =   "ˢ��(&R)"
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
         Left            =   300
         TabIndex        =   15
         Top             =   90
         Width           =   1200
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "ѡ��(&S)"
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
         Left            =   1695
         TabIndex        =   11
         Top             =   90
         Width           =   1200
      End
      Begin VB.CommandButton cmdCls 
         Caption         =   "���(&M)"
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
         Left            =   3000
         TabIndex        =   12
         Top             =   90
         Width           =   1200
      End
      Begin VB.CommandButton cmdClsAll 
         Caption         =   "ȫ��(&C)"
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
         Left            =   3000
         TabIndex        =   14
         Top             =   525
         Width           =   1200
      End
      Begin VB.CommandButton cmdSelAll 
         Caption         =   "ȫѡ(&A)"
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
         Left            =   1695
         TabIndex        =   13
         Top             =   525
         Width           =   1200
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "�˳�(&X)"
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
         Left            =   7140
         TabIndex        =   10
         Top             =   525
         Width           =   1200
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "���(&O)"
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
         Left            =   7140
         TabIndex        =   9
         Top             =   90
         Width           =   1200
      End
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   -15
      ScaleHeight     =   240
      ScaleWidth      =   9585
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1110
      Width           =   9585
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "δ��˻��۵�,��ǰ�ϼ�:"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Tag             =   "δ��˻��۵�,��ǰ�ϼ�:"
         Top             =   30
         Width           =   1980
      End
   End
   Begin VB.Frame fraInfo 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   45
      TabIndex        =   17
      Top             =   -45
      Width           =   8505
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   360
         Left            =   555
         TabIndex        =   30
         Top             =   180
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   635
         Appearance      =   2
         IDKindStr       =   "��|����|0;ҽ|ҽ����|0;��|���֤��|0;IC|IC����|1;��|�����|0;��|���￨|0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   12
         FontName        =   "����"
         IDKind          =   -1
         ShowPropertySet =   -1  'True
         BackColor       =   -2147483633
      End
      Begin VB.TextBox txtʣ�� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6800
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   675
         Width           =   1440
      End
      Begin VB.TextBox txt���� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   675
         Width           =   1200
      End
      Begin VB.TextBox txtԤ�� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   675
         Width           =   1200
      End
      Begin VB.TextBox txt�ѱ� 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6795
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   1440
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5205
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   180
         Width           =   960
      End
      Begin VB.TextBox txt�Ա� 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   180
         Width           =   615
      End
      Begin VB.TextBox txtPatient 
         BackColor       =   &H00EBFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1230
         MaxLength       =   100
         TabIndex        =   0
         ToolTipText     =   "�ȼ�:F6"
         Top             =   180
         Width           =   1680
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   30
         X2              =   8450
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000015&
         X1              =   30
         X2              =   8450
         Y1              =   585
         Y2              =   585
      End
      Begin VB.Label lblʣ��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʣ���"
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
         Left            =   6000
         TabIndex        =   24
         Top             =   765
         Width           =   630
      End
      Begin VB.Label lblδ����� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "δ�����"
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
         Left            =   3000
         TabIndex        =   23
         Top             =   765
         Width           =   840
      End
      Begin VB.Label lblԤ����� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ԥ�����"
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
         Left            =   120
         TabIndex        =   22
         Top             =   765
         Width           =   840
      End
      Begin VB.Label lbl�ѱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ѱ�"
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
         Left            =   6255
         TabIndex        =   21
         Top             =   255
         Width           =   420
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
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
         Left            =   4680
         TabIndex        =   20
         Top             =   255
         Width           =   420
      End
      Begin VB.Label lbl�Ա� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
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
         Left            =   3480
         TabIndex        =   19
         Top             =   255
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   420
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   2415
      Left            =   30
      TabIndex        =   7
      ToolTipText     =   "˫�����ݲ鿴��ϸ"
      Top             =   1365
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   4260
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmBillingAuditing.frx":08A4
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   29
      Top             =   7080
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBillingAuditing.frx":0BBE
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10557
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmBillingAuditing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Private mstrPrivs As String
Private mlngModule As Long
Private mrsInfo As New ADODB.Recordset
Private mrsList As ADODB.Recordset
Attribute mrsList.VB_VarHelpID = -1
Private mlngCurRow As Long, mlngTopRow As Long
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mintSucces As Integer
Private mblnNotClick As Boolean

'-----------------------------------------------------------------------------------
'���㿨���
Private mstrPassWord As String
'-----------------------------------------------------------------------------------
Private mobjDrugMachine As Object '�Զ���ҩ��(�£�
Private mblnDrugMachine As Boolean

Public Function zlShowCard(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '����:��˳ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-03 15:23:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mintSucces = 0: mlngModule = lngModule: mstrPrivs = strPrivs
    Me.Show 1, frmMain
    zlShowCard = mintSucces > 0
End Function
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCls_Click()
    Dim i As Long, intS As Integer
    intS = 1
    If mshList.Row > mshList.RowSel Then intS = -1
    For i = mshList.Row To mshList.RowSel Step intS
        If mshList.TextMatrix(i, 1) <> "" Then
            mshList.TextMatrix(i, 0) = ""
        End If
    Next
    lblTotal.Caption = lblTotal.Tag & Format(CalcTotal, gstrDec)
End Sub

Private Sub cmdClsAll_Click()
    Dim i As Long
    mshList.Redraw = False
    For i = 1 To mshList.Rows - 1
        If mshList.TextMatrix(i, 1) <> "" Then
            mshList.TextMatrix(i, 0) = ""
        End If
    Next
    lblTotal.Caption = lblTotal.Tag & Format(CalcTotal, gstrDec)
    mshList.Redraw = True
End Sub

Private Sub cmdFlash_Click()
    If mrsInfo.State = 0 Then
        MsgBox "û��ȷ������,�������벡����Ϣ��", vbInformation, gstrSysName
        txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
    End If
    Call ShowBills
End Sub

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim strDel As String, i As Long, str���ʱ�� As String, Curdate As Date, curTotal As Currency
    Dim arrSQL As Variant, strNos As String, strNo As String, blnTrans As Boolean
    
    arrSQL = Array()
    For i = 1 To mshList.Rows - 1
        If mshList.TextMatrix(i, 0) <> "" And mshList.TextMatrix(i, 1) <> "" Then
            If str���ʱ�� = "" Then
                Curdate = zlDatabase.Currentdate
                str���ʱ�� = "To_Date('" & Format(Curdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
            End If
            strNo = mshList.TextMatrix(i, 1)
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_������ʼ�¼_Verify('" & strNo & "','" & UserInfo.��� & "','" & UserInfo.���� & "',Null," & str���ʱ�� & ")"
            strDel = strDel & "," & i
            
            strNos = strNos & "," & strNo
        End If
    Next
    If UBound(arrSQL) = -1 Then
        MsgBox "û��ѡ��Ҫ��˵Ļ��۵��ݣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    curTotal = CalcTotal
    If curTotal <> 0 And gdblԤ��������鿨 <> 0 Then
        If Not zlDatabase.PatiIdentify(Me, glngSys, Val(mrsInfo!����ID), curTotal, , , , IIf(-1 * gdblԤ��������鿨 >= curTotal, False, True), , , , (gdblԤ��������鿨 = 2)) Then Exit Sub
    End If
    
    strNos = Mid(strNos, 2)
    
    '���ñ���
    If Not AuditingWarnByPatient(strNos) Then Exit Sub
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    If gbln��˴�ӡ Then
        For i = 0 To UBound(Split(strNos, ","))
            strNo = Split(strNos, ",")(i)
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1122", Me, "NO=" & strNo, "ҩƷ��λ=" & IIf(gblnҩ����λ, 1, 0), "PrintEmpty=0", 2)
        Next
    End If
    
    '110319
    If mblnDrugMachine Then
        '�����ʽ��1|����1,������1;����2,������2
        Dim strData As String, strReturn As String
        strData = "1|" & "9," & Replace(strNos, ",", ";9,")
        Call mobjDrugMachine.Operation(gstrDBUser, Val("21-��ҩ[�����סԺ������ϸ�ϴ�]"), strData, strReturn)
    End If
    
    On Error GoTo 0
    
    'ע�ⷽ��
    strDel = Mid(strDel, 2)
    For i = UBound(Split(strDel, ",")) To 0 Step -1
        If mshList.Rows > 2 Then
            mshList.RemoveItem CLng(Split(strDel, ",")(i))
        Else
            mshList.Clear
            mshList.Rows = 2
            Call SetHeader
        End If
    Next
    
    Call mshList_EnterCell
    
    lblTotal.Caption = lblTotal.Tag & Format(CalcTotal, gstrDec)
    Call RefreshMoney
    
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    mintSucces = mintSucces + 1
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSel_Click()
    Dim i As Long, intS As Integer
    intS = 1
    If mshList.Row > mshList.RowSel Then intS = -1
    For i = mshList.Row To mshList.RowSel Step intS
        If mshList.TextMatrix(i, 1) <> "" Then
            mshList.TextMatrix(i, 0) = "��"
        End If
    Next
    lblTotal.Caption = lblTotal.Tag & Format(CalcTotal, gstrDec)
End Sub

Private Sub cmdSelAll_Click()
    Dim i As Long
    mshList.Redraw = False
    For i = 1 To mshList.Rows - 1
        If mshList.TextMatrix(i, 1) <> "" Then
            mshList.TextMatrix(i, 0) = "��"
        End If
    Next
    lblTotal.Caption = lblTotal.Tag & Format(CalcTotal, gstrDec)
    mshList.Redraw = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF4
            If Shift <> vbCtrlMask Then Exit Sub
            If IDKind.Enabled Then
                Dim intIndex As Integer
                intIndex = IDKind.GetKindIndex("IC����")
                If intIndex <= 0 Then Exit Sub
                IDKind.IDKind = intIndex: Call IDKind_Click(IDKind.GetCurCard)
            End If
        Case vbKeyF6
            txtPatient.SetFocus
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    Call initCardSquareData
    Me.Height = 7815 'û��дResize
    Set mobjIDCard = New clsIDCard
    Call mobjIDCard.SetParent(Me.hWnd)
    Call CreateDrugPacker
    Call SetHeader
    Call SetBill
    txtPatient.MaxLength = zlGetPatiInforMaxLen.intPatiName
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    Set mobjICCard = Nothing
    
    Set mrsList = Nothing
    Set mrsInfo = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub

 
Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If mobjICCard Is Nothing Then Exit Sub
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text <> "" Then
            Call FindPati(objCard, True, txtPatient.Text)
        End If
        Exit Sub
    End If
    
    lng�����ID = objCard.�ӿ����
    If lng�����ID <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '����:�����ӿ�
    '    '���:frmMain-���õĸ�����
    '    '       lngModule-���õ�ģ���
    '    '       strExpand-��չ����,������
    '    '       blnOlnyCardNO-������ȡ����
    '    '����:strOutCardNO-���صĿ���
    '    '       strOutPatiInforXML-(������Ϣ����.XML��)
    '    '����:��������    True:���óɹ�,False:����ʧ��\
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModule, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then
        Call FindPati(objCard, True, txtPatient.Text)
    End If
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    If mblnNotClick Then Exit Sub
    Set gobjSquare.objCurCard = objCard
    If txtPatient.Text <> "" Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub
Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    Dim lngPreIDKind As Long
      '����:60010
    If txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.����
    If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    Dim lngPreIDKind As Long
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    txtPatient.Text = strCardNo
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("IC��", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    '���֤ʶ��
    Dim lngPreIDKind As Long
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("���֤", CardTypeName)
    txtPatient.Text = strID
    If objCard Is Nothing Then Exit Sub
    If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
End Sub

Private Sub mshList_DblClick()
    If mshList.MouseRow = 0 Then Exit Sub
    If mshList.TextMatrix(mshList.Row, 1) = "" Then Exit Sub
    
    If mshList.MouseCol = 0 Then
        If mshList.TextMatrix(mshList.Row, 0) = "" Then
            mshList.TextMatrix(mshList.Row, 0) = "��"
        Else
            mshList.TextMatrix(mshList.Row, 0) = ""
        End If
        lblTotal.Caption = lblTotal.Tag & Format(CalcTotal, gstrDec)
    Else
       Call ShowBill
    End If
End Sub

Private Sub mshList_EnterCell()
    If mshList.Row = 0 Or mshList.TextMatrix(mshList.Row, 1) = "" Then
        mshBill.Clear
        mshBill.Rows = 2
        Call SetBill
        Exit Sub
    End If
    mlngCurRow = mshList.Row: mlngTopRow = mshList.TopRow
    
    Call ShowDetail(mshList.TextMatrix(mshList.Row, 1))
End Sub

Private Sub ShowDetail(Optional strNo As String)
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim strSQL As String
    
    On Error GoTo errH
    
    '��ϸ�����е�ʣ�������ͽ��
    strSQL = _
    " Select C.���� as ���,Nvl(E.����,B.����) as ����," & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "E1.���� as ��Ʒ��,", "") & "B.���," & _
            IIf(gblnҩ����λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X." & gstrҩ����λ & ")", "A.���㵥λ") & " as ��λ," & _
    "       Avg(Nvl(A.����,1)*A.����)" & IIf(gblnҩ����λ, "/Nvl(X." & gstrҩ����װ & ",1)", "") & " as ����, " & _
    "       Ltrim(To_Char(Sum(A.��׼����)" & IIf(gblnҩ����λ, "*Nvl(X." & gstrҩ����װ & ",1)", "") & ",'99999" & gstrFeePrecisionFmt & "')) as ����," & _
    "       Ltrim(To_Char(Sum(A.Ӧ�ս��),'99999" & gstrDec & "')) as Ӧ�ս��," & _
    "       Ltrim(To_Char(Sum(A.ʵ�ս��),'99999" & gstrDec & "')) as ʵ�ս��," & _
    "       D.���� as ִ�п���" & _
    " From ������ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,�շ���Ŀ���� E,ҩƷ��� X" & _
            IIf(gTy_System_Para.bytҩƷ������ʾ = 2, ",�շ���Ŀ���� E1", "") & _
    " Where A.�շ�ϸĿID=B.ID and A.�շ����=C.���� And A.ִ�в���ID=D.ID(+)" & _
    "       And A.NO=[1] And A.��¼����=2 And A.�����־ In(1,3,4) And A.��¼״̬=0" & _
    "       And A.����ID+0=[2]" & _
    "       And A.�շ�ϸĿID=X.ҩƷID(+) And A.����Ա���� is NULL And A.������ is Not NULL" & _
    "       And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
            IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3", "") & _
    " Group by Nvl(A.�۸񸸺�,A.���),C.����," & _
    "       Nvl(E.����,B.����),B.���,A.���㵥λ,D.����," & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "E1.����,", "") & " X.ҩƷID,X." & gstrҩ����λ & ",Nvl(X." & gstrҩ����װ & ",1)" & _
    " Order by Nvl(A.�۸񸸺�,A.���)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, CLng(mrsInfo!����ID))
    
    mshBill.Redraw = False
    mshBill.ClearStructure
    mshBill.Clear
    mshBill.Rows = 2
    If Not rsTmp.EOF Then Set mshBill.DataSource = rsTmp
    Call SetBill
    mshBill.Redraw = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetBill()
    Dim strHead As String
    Dim i As Long
    
    strHead = "���,1,650|����,1,1500" & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "|��Ʒ��,1,2000", "") & "|���,1,1500|��λ,1,500|����,1,750|����,7,750|Ӧ�ս��,7,850|ʵ�ս��,7,850|ִ�п���,1,1000"
    With mshBill
        .COLS = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshBill, App.ProductName & "\" & Me.Name)
        For i = 0 To .COLS - 1
            If .TextMatrix(0, i) = "��Ʒ��" Then
                If gTy_System_Para.bytҩƷ������ʾ = 2 Then
                    If .ColWidth(i) <= 0 Then .ColWidth(i) = 2000
                Else
                    .ColWidth(i) = 0
                End If
            End If
        Next
        
        .RowHeight(0) = 320
        
        .Col = 0: .ColSel = .COLS - 1
    End With
End Sub

Private Sub mshList_KeyPress(KeyAscii As Integer)
    
    If mshList.TextMatrix(mshList.Row, 1) = "" Then Exit Sub
    
    If KeyAscii = 32 Then
        If mshList.TextMatrix(mshList.Row, 0) = "" Then
            mshList.TextMatrix(mshList.Row, 0) = "��"
        Else
            mshList.TextMatrix(mshList.Row, 0) = ""
        End If
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        Call ShowBill
    End If
End Sub

Private Sub ShowBill()
    Dim strNo As String

    On Error Resume Next
        
    strNo = mshList.TextMatrix(mshList.Row, 1)
    
    frmCharge.mlngModul = 1122
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mbytInFun = 2
    frmCharge.mbytInState = 1
    frmCharge.mstrTime = ""
    frmCharge.mblnDelete = False
    frmCharge.mstrInNO = strNo
    frmCharge.mblnNOMoved = False
    frmCharge.mbytBilling = 1
    frmCharge.Show 1, Me
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If mshList.Height + Y < 1000 Or mshBill.Height - Y < 1000 Then Exit Sub
        pic.Top = pic.Top + Y
        mshList.Height = mshList.Height + Y
        mshBill.Top = mshBill.Top + Y
        mshBill.Height = mshBill.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub txtPatient_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    IDKind.SetAutoReadCard (txtPatient.Text = "")
End Sub

Private Sub txtPatient_GotFocus()
    zlControl.TxtSelAll txtPatient
     If txtPatient.Locked Then Exit Sub
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPatient.Text = "")
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtPatient.Text = "")
    IDKind.SetAutoReadCard (txtPatient.Text = "")

End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
    If IDKind.GetCurCard Is Nothing Then Exit Sub
       If IDKind.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean, blnICCard As Boolean
    
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
    '����:51488
    If (IDKind.Cards.������� = "�ո��" Or IDKind.Cards.������� = " ") And Chr(KeyAscii) = " " Then KeyAscii = 0: Exit Sub
 
    If IDKind.GetCurCard.���� Like "����*" Then
        '103563,ֻҪ����ĵ�һ���ַ��ǡ�-+*����������ȫ���֣�����Ϊ����ˢ��
        If Not (InStr("-+*", Left(txtPatient.Text, 1)) > 0 And IsNumeric(Mid(txtPatient.Text, 2))) Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        End If
    ElseIf IDKind.GetCurCard.���� = "�����" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtPatient.IMEMode = 0
    End If
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        Call FindPati(IDKind.GetCurCard, blnCard, txtPatient.Text)
    End If
End Sub
Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ҳ���
    '����:���˺�
    '����:2012-08-31 17:54:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnICCard As Boolean, blnIDCard As Boolean
    
    '��ȡ������Ϣ
    Call ClearPati
    mshList.Clear: mshList.Rows = 2
    Call SetHeader

    If objCard.���� Like "IC��*" And objCard.ϵͳ And mstrPassWord <> "" Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If objCard.���� Like "*���֤*" And objCard.ϵͳ And mstrPassWord <> "" Then blnIDCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    
    If Not GetPatient(objCard, txtPatient.Text, blnCard) Then
        txtPatient.Text = ""
        If blnCard Then
            sta.Panels(2) = "����ȷ��������Ϣ�������Ƿ���ȷˢ����"
            txtPatient.SetFocus: Exit Sub
        End If
        sta.Panels(2) = "����ı�ʶ���ܶ�ȡ������Ϣ�����������Ƿ���ȷ��"
        txtPatient.SetFocus: Exit Sub
    End If
    '���￨������
    If (objCard.���� Like "IC��*" Or objCard.���� Like "*���֤*") And objCard.ϵͳ = True And blnCard Then blnCard = False
    If Mid(gstrCardPass, 4, 1) = "1" And (blnCard Or blnICCard Or blnIDCard) Then
        If Not zlCommFun.VerifyPassWord(Me, mstrPassWord, mrsInfo!����, mrsInfo!�Ա�, "" & mrsInfo!����) Then
            Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
        End If
    End If
    txtPatient.PasswordChar = ""
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    txtPatient.Text = "" & mrsInfo!����
    txt�Ա�.Text = "" & mrsInfo!�Ա�
    txt����.Text = "" & mrsInfo!����
    txt�ѱ�.Text = "" & mrsInfo!�ѱ�
    Call RefreshMoney
    Call ShowBills
    mshList.SetFocus
 
End Sub
 
 Private Sub RefreshMoney()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetMoneyInfo(mrsInfo!����ID, , , 1)
    If Not rsTmp Is Nothing Then
        txtԤ��.Text = Format(rsTmp!Ԥ�����, "0.00")
        txt����.Text = Format(rsTmp!�������, gstrDec)
        txtʣ��.Text = Format(rsTmp!Ԥ����� - rsTmp!�������, "0.00")
    Else
        txtԤ��.Text = ""
        txt����.Text = ""
        txtʣ��.Text = ""
    End If
End Sub

Private Sub ClearPati()
    txt�Ա�.Text = ""
    txt����.Text = ""
    txt�ѱ�.Text = ""
    
    txt����.Text = ""
    txtԤ��.Text = ""
    txtʣ��.Text = ""
End Sub

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, ByVal blnCard As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:blnCard=�Ƿ���￨ˢ��
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-03 16:49:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    Dim strPati As String, vRect As RECT, blnCancel As Boolean
    Dim strIF As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSQL = _
            "Select A.����ID,A.���￨��,A.����֤��,A.����,A.�Ա�,A.����,A.�ѱ�,A.����,A.��������" & _
            " From ������Ϣ A" & _
            " Where A.ͣ��ʱ�� is NULL "
            
    If blnCard = True And objCard.���� Like "����*" And InStr("-+*", Left(strInput, 1)) = 0 Then    '103563
        lng�����ID = IDKind.GetDefaultCardTypeID
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg, lng�����ID) = False Then GoTo NotFoundPati:
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng����ID
        blnHavePassWord = True
        strSQL = strSQL & " And A.����ID=[1] "
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '����ID
        strSQL = strSQL & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��(���˳�Ժ)
        strSQL = strSQL & " And A.סԺ��=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����
        strSQL = strSQL & " And A.�����=[1]"
        '75087,Ƚ����,2014-7-29,���ﲡ���շ�ʱ,����Ҫ���������������,ֻ��Ҫ��������ŵ����˳��ż����ҵ��������Ĳ�����Ϣ������
        strInput = "*" & zlCommFun.GetFullNO(Mid(strInput, 2), 3)
    Else
        Select Case objCard.����
            Case "����", "��������￨"  '��������
                 strPati = _
                    " Select A.����ID as ID,A.����ID,A.����,A.�Ա�,A.����,A.���￨��,A.����֤��,A.�����,A.��������,A.���֤��,A.��ͥ��ַ,A.������λ" & _
                    " From ������Ϣ A Where  A.���� Like  [1]  " & _
                                IIf(gintNameDays = 0, "", " And (A.����ʱ��>Trunc(Sysdate-" & gintNameDays & ") Or A.�Ǽ�ʱ��>Trunc(Sysdate-" & gintNameDays & "))") & _
                    " And Rownum<101" & _
                    " Order by A.����"
                vRect = zlControl.GetControlRect(txtPatient.hWnd)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "����Find", False, "", "", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%")
                If blnCancel Then Exit Function
                If Not rsTmp Is Nothing Then
                    strInput = rsTmp!����ID
                    strSQL = strSQL & " And A.����ID=[2]"
                Else
                    Exit Function
                End If
                
            Case "ҽ����"
                strInput = UCase(strInput)
                strSQL = strSQL & " And A.ҽ����=[2]"
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.�����=[2]"
                '75087,Ƚ����,2014-7-29,���ﲡ���շ�ʱ,����Ҫ���������������,ֻ��Ҫ��������ŵ����˳��ż����ҵ��������Ĳ�����Ϣ������
                strInput = zlCommFun.GetFullNO(strInput, 3)
            Case Else
                '��������,��ȡ��صĲ���ID
                If objCard.�ӿ���� > 0 Then
                    lng�����ID = objCard.�ӿ����
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng����ID = 0 Then GoTo NotFoundPati:
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng����ID <= 0 Then GoTo NotFoundPati:
                strSQL = strSQL & " And A.����ID=[1]"
                strInput = "-" & lng����ID
                blnHavePassWord = True
        End Select
    End If
        
    txtPatient.ForeColor = Me.ForeColor
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    If mrsInfo.RecordCount <> 0 Then '
        '75259:���ϴ���2014-7-10������������ʾ��ɫ����
        Call SetPatiColor(txtPatient, Nvl(mrsInfo!��������), IIf(IsNull(mrsInfo!����), Me.ForeColor, vbRed))
        mstrPassWord = strPassWord
        If Not blnHavePassWord Then mstrPassWord = Nvl(mrsInfo!����֤��)
        GetPatient = True
    Else
        Set mrsInfo = New ADODB.Recordset
    End If
    Exit Function
NotFoundPati:
    Set mrsInfo = New ADODB.Recordset
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set mrsInfo = New ADODB.Recordset
End Function

Private Sub ShowBills(Optional blnSort As Boolean)
'����:��������ȡ�����б�(���˹���)
'����:strIF=��"AND"��ʼ��������
'     blnSort=�����¶�ȡ����,��������ʾ�����������
    Dim i As Long, Curdate As Date
    
    On Error GoTo errH
    
    If Not blnSort Then
        sta.Panels(2).Text = "���ڶ�ȡ���˻��۵���,���Ժ� ..."
        Screen.MousePointer = 11
        DoEvents
        Me.Refresh
        
        gstrSQL = _
        " Select NULL as ���,A.NO as ���ݺ�," & _
        "       B.���� as ��������,A.������ as ҽ��,A.�ѱ�," & _
        "       LTrim(To_Char(Sum(A.Ӧ�ս��),'999999999" & gstrDec & "')) as Ӧ�ս��," & _
        "       LTrim(To_Char(Sum(A.ʵ�ս��),'999999999" & gstrDec & "')) as ʵ�ս��," & _
        "       A.������,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as ����ʱ��" & _
        " From ������ü�¼ A,���ű� B" & _
        " Where A.��¼����=2 And A.�����־ In(1,3,4) And A.��¼״̬=0" & _
        "       And A.������ is Not Null And A.����Ա���� is NULL And A.��������ID=B.ID" & _
        "       And A.����ID=[1]" & _
        " Group by A.NO,B.����,A.������,A.�ѱ�,A.�Ǽ�ʱ��,A.������" & _
        " Order by ����ʱ�� Desc,���ݺ� Desc"
        Set mrsList = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsInfo!����ID))
    End If
    
    mshList.Redraw = False
    mshList.ClearStructure
    mshList.Clear
    mshList.Rows = 2
    
    If mrsList.EOF Then
        sta.Panels(2).Text = "û�з��ֻ��۵���"
    Else
        Set mshList.DataSource = mrsList
        sta.Panels(2).Text = "�� " & mrsList.RecordCount & " �Ż��۵���"
    End If
    Call SetHeader
        
    lblTotal.Caption = lblTotal.Tag & Format(CalcTotal, gstrDec)
    
    mshList.Redraw = True
    Screen.MousePointer = 0
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetHeader()
    Dim strHead As String
    Dim i As Long
    
    strHead = "���,4,500|���ݺ�,1,820|��������,1,1000|ҽ��,1,750|�ѱ�,1,500|Ӧ�ս��,7,850|ʵ�ս��,7,850|������,1,700|����ʱ��,4,1850"
    With mshList
        .COLS = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshList, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 320
        
        '�ָ��ϴ���
        If mlngCurRow = 0 Then mlngCurRow = 1
        If mlngTopRow = 0 Then mlngTopRow = 1
        If mlngCurRow <= .Rows - 1 Then
            .Row = mlngCurRow
        Else
            .Row = .Rows - 1
        End If
        If mlngTopRow <= .Rows - 1 Then
            .TopRow = mlngTopRow
        Else
            .TopRow = .Row
        End If
        
        .Col = 0: .ColSel = .COLS - 1
                
        Call mshList_EnterCell
    End With
End Sub

Private Sub mshList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshList.MouseRow = 0 Then
        mshList.MousePointer = 99
    Else
        mshList.MousePointer = 0
    End If
End Sub

Private Sub mshList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = mshList.MouseCol
    
    If Button = 1 And mshList.MousePointer = 99 Then
        If mshList.TextMatrix(0, lngCol) = "" Then Exit Sub
        If mshList.TextMatrix(mshList.Row, 1) = "" Then Exit Sub
        If mrsList Is Nothing Then Exit Sub
        
        Set mshList.DataSource = Nothing

        mrsList.Sort = mshList.TextMatrix(0, lngCol) & IIf(mshList.ColData(lngCol) = 0, "", " DESC")
        mshList.ColData(lngCol) = (mshList.ColData(lngCol) + 1) Mod 2
        
        Call ShowBills(True)
    End If
End Sub

Private Function CalcTotal() As Currency
    Dim i As Long
    For i = 1 To mshList.Rows - 1
        If mshList.TextMatrix(i, 0) <> "" Then
            CalcTotal = CalcTotal + Val(mshList.TextMatrix(i, 6))
        End If
    Next
End Function

Private Function AuditingWarnByPatient(ByVal strNos As String) As Boolean
'���ܣ���˻��۵�ʱ���Է��ý��б���
'������str���=ָ��������Ҫ��˵��к�,Ϊ�ձ�ʾ������
    Dim rsWarn As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str���s As String, cur��� As Currency, cur��� As Currency
    Dim strWarn As String, intWarn As Integer
    
    On Error GoTo errH
    
    '���������Ϣ
    strSQL = _
        " Select A.�շ����,B.���� as �������,Sum(A.ʵ�ս��) as ���" & _
        " From ������ü�¼ A,�շ���Ŀ��� B" & _
        " Where A.��¼����=2 And A.�����־ In(1,3,4) And A.��¼״̬=0" & _
        "       And A.�շ����=B.���� And A.������ is Not Null And A.����Ա���� is NULL" & _
                IIf(strNos <> "", " And Instr(','||[2]||',',','||A.NO||',')>0", "") & _
        "       And A.����ID=[1]" & _
        " Group by A.�շ����,B.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mrsInfo!����ID), strNos)
    
    For i = 1 To rsTmp.RecordCount
        If InStr(str���s, rsTmp!�շ���� & rsTmp!�������) = 0 Then
            str���s = str���s & "," & rsTmp!�շ���� & rsTmp!�������
        End If
        cur��� = cur��� + rsTmp!���
        rsTmp.MoveNext
    Next
    str���s = Mid(str���s, 2)
    
    If cur��� > 0 Then
        '���������Ϣ
        strSQL = "Select A.����,C.Ԥ�����-C.������� as ���,zl_PatiDayCharge(A.����ID) as ���ն�," & _
            " Decode(A.������,null,A.������,Zl_Patientsurety(A.����ID,Null)) ������,Zl_Patiwarnscheme(A.����id, Null) As ���ò���" & _
            " From ������Ϣ A,������� C" & _
            " Where A.����ID=C.����ID(+) And C.����(+)=1 And A.����ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mrsInfo!����ID))
        Set rsWarn = GetUnitWarn(rsTmp!���ò���, "0")
        
        cur��� = Nvl(rsTmp!���, 0)
        If gbln�����������۷��� Then cur��� = Nvl(rsTmp!���, 0) - GetPriceMoneyTotal(0, mrsInfo!����ID) + cur���
        '���౨��
        For i = 0 To UBound(Split(str���s, ","))
            intWarn = BillingWarn(mstrPrivs, rsTmp!����, rsTmp!���ò���, rsWarn, _
                cur���, Nvl(rsTmp!���ն�, 0), cur���, Nvl(rsTmp!������, 0), _
                Left(Split(str���s, ",")(i), 1), Mid(Split(str���s, ",")(i), 2), strWarn)
            If intWarn = 2 Or intWarn = 3 Then Exit Function
        Next
    End If
    AuditingWarnByPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub
Private Sub initCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������رս��㿨����
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    Set gobjSquare.objDefaultCard = IDKind.GetfaultCard
    If IDKind.Cards.��ȱʡ������ And Not gobjSquare.objDefaultCard Is Nothing Then
        gobjSquare.blnȱʡ�������� = gobjSquare.objDefaultCard.�������Ĺ��� <> ""
        gobjSquare.intȱʡ���ų��� = gobjSquare.objDefaultCard.���ų���
    Else
        gobjSquare.blnȱʡ�������� = IDKind.Cards.������ʾ
        gobjSquare.intȱʡ���ų��� = 100
    End If
    gobjSquare.bln��ȱʡ������ = IDKind.Cards.��ȱʡ������
End Sub

Private Sub CreateDrugPacker()
    '����:����������ҩ��(�Զ���ҩ��)
    Dim objComLib As New zl9ComLib.clsComLib
    Dim strPrivs As String
    Dim strMessage As String
    
    mblnDrugMachine = False

    Err = 0: On Error Resume Next
    If Val(zlDatabase.GetPara("����ҩƷ�Զ����豸�ӿ�", glngSys, Val("9010-ҩƷ�Զ����豸�ӿ�"))) = 1 Then
        '�����½ӿ�
        Set mobjDrugMachine = CreateObject("zlDrugMachine.clsDrugMachine")
        If Err = 0 Then mblnDrugMachine = True
    End If
    
    Err = 0: On Error GoTo 0
    If mblnDrugMachine Then
        'Ȩ�޼��
        strPrivs = GetPrivFunc(glngSys, Val("9010-ҩƷ�Զ����豸�ӿ�"))
        If InStr(";" & strPrivs & ";", ";����;") > 0 Then
            mblnDrugMachine = mobjDrugMachine.Init(1, objComLib, strMessage)
        Else
            mblnDrugMachine = False
        End If
    End If
End Sub

