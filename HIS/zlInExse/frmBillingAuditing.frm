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
   ClientWidth     =   9930
   Icon            =   "frmBillingAuditing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   9930
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
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   3780
      Width           =   8400
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshBill 
      Height          =   2265
      Left            =   30
      TabIndex        =   10
      ToolTipText     =   "˫�����ݲ鿴��ϸ"
      Top             =   3825
      Width           =   9810
      _ExtentX        =   17304
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
      ScaleWidth      =   9930
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   6120
      Width           =   9930
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   300
         TabIndex        =   18
         Top             =   525
         Width           =   1100
      End
      Begin VB.CommandButton cmdFlash 
         Caption         =   "ˢ��(&R)"
         Height          =   350
         Left            =   300
         TabIndex        =   17
         Top             =   90
         Width           =   1100
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "ѡ��(&S)"
         Height          =   350
         Left            =   1695
         TabIndex        =   13
         Top             =   90
         Width           =   1100
      End
      Begin VB.CommandButton cmdCls 
         Caption         =   "���(&M)"
         Height          =   350
         Left            =   2880
         TabIndex        =   14
         Top             =   90
         Width           =   1100
      End
      Begin VB.CommandButton cmdClsAll 
         Caption         =   "ȫ��(&C)"
         Height          =   350
         Left            =   2880
         TabIndex        =   16
         Top             =   525
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelAll 
         Caption         =   "ȫѡ(&A)"
         Height          =   350
         Left            =   1695
         TabIndex        =   15
         Top             =   525
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "�˳�(&X)"
         Height          =   350
         Left            =   7140
         TabIndex        =   12
         Top             =   525
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "���(&O)"
         Height          =   350
         Left            =   7140
         TabIndex        =   11
         Top             =   90
         Width           =   1100
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
      ScaleWidth      =   9855
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1110
      Width           =   9855
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "δ��˻��۵�,��ǰ�ϼ�:"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   120
         TabIndex        =   31
         Tag             =   "δ��˻��۵�,��ǰ�ϼ�:"
         Top             =   30
         Width           =   1980
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   1125
      Left            =   45
      TabIndex        =   19
      Top             =   -45
      Width           =   9795
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   315
         Left            =   450
         TabIndex        =   34
         Top             =   195
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   556
         Appearance      =   2
         IDKindStr       =   $"frmBillingAuditing.frx":08A4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "����"
         IDKind          =   -1
         ShowPropertySet =   -1  'True
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         BackColor       =   -2147483633
      End
      Begin VB.TextBox txtʣ�� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   6870
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   705
         Width           =   1125
      End
      Begin VB.TextBox txt���� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3975
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   705
         Width           =   1125
      End
      Begin VB.TextBox txtԤ�� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   705
         Width           =   1125
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   8340
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   195
         Width           =   1320
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   6870
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   195
         Width           =   720
      End
      Begin VB.TextBox txtסԺ�� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   5385
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   195
         Width           =   840
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3990
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   195
         Width           =   480
      End
      Begin VB.TextBox txt�Ա� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   195
         Width           =   495
      End
      Begin VB.TextBox txtPatient 
         BackColor       =   &H00EBFFFF&
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   1110
         MaxLength       =   100
         TabIndex        =   0
         ToolTipText     =   "�ȼ�:F6"
         Top             =   195
         Width           =   1155
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
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   7875
         TabIndex        =   28
         Top             =   255
         Width           =   360
      End
      Begin VB.Label lblʣ��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʣ���"
         Height          =   180
         Left            =   6255
         TabIndex        =   27
         Top             =   765
         Width           =   540
      End
      Begin VB.Label lblδ����� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "δ�����"
         Height          =   180
         Left            =   3165
         TabIndex        =   26
         Top             =   765
         Width           =   720
      End
      Begin VB.Label lblԤ����� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ԥ�����"
         Height          =   180
         Left            =   330
         TabIndex        =   25
         Top             =   765
         Width           =   720
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   6420
         TabIndex        =   24
         Top             =   255
         Width           =   360
      End
      Begin VB.Label lblסԺ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��"
         Height          =   180
         Left            =   4785
         TabIndex        =   23
         Top             =   255
         Width           =   540
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   3540
         TabIndex        =   22
         Top             =   255
         Width           =   360
      End
      Begin VB.Label lbl�Ա� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   2385
         TabIndex        =   21
         Top             =   255
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000007&
         Height          =   180
         Left            =   90
         TabIndex        =   20
         Top             =   255
         Width           =   360
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   2415
      Left            =   30
      TabIndex        =   9
      ToolTipText     =   "˫�����ݲ鿴��ϸ"
      Top             =   1365
      Width           =   9795
      _ExtentX        =   17277
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
      MouseIcon       =   "frmBillingAuditing.frx":093A
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   33
      Top             =   7080
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBillingAuditing.frx":0C54
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12912
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
Private mlngModule As String
Private mlngUnitID As Long '��ǰ��ѡ��Ĳ���ID
Private mstrUnitIDs As String   '��ǰ����Ա�����в���ID
Private mstrPrivs As String
Private mstrPrivsOpt As String '���ʲ���1150ģ�����Ȩ����
Private mrsInfo As New ADODB.Recordset
Private mrsList As ADODB.Recordset
Attribute mrsList.VB_VarHelpID = -1
Private mlngCurRow As Long, mlngTopRow As Long
Private mobjICCard As Object
Private mintSucces As Integer
'-----------------------------------------------------------------------------------
'���㿨���
Private mstrPassWord As String
'-----------------------------------------------------------------------------------
Public Function zlCardShow(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, lngUnitID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��˳������
    '���:lngUnitID-��ǰ��ѡ��Ĳ���ID
    '����:
    '����:���һ�γɹ�����,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-03 17:30:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mintSucces = 0: mlngModule = lngModule: mstrPrivs = strPrivs: mlngUnitID = lngUnitID
    Me.Show 1, frmMain
    zlCardShow = mintSucces > 0
End Function
 

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
    Dim strDel As String, i As Long, str���ʱ�� As String, Curdate As Date
    Dim arrSQL As Variant, strNos As String, strNO As String, blnTrans As Boolean
    
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State <> 1 Then Exit Sub
    If mrsInfo.EOF Then Exit Sub
    If zlIsAllowFeeChange(Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ҳID))) = False Then
         Exit Sub
    End If
    
    arrSQL = Array()
    For i = 1 To mshList.Rows - 1
        If mshList.TextMatrix(i, 0) <> "" And mshList.TextMatrix(i, 1) <> "" Then
            If str���ʱ�� = "" Then
                Curdate = zlDatabase.Currentdate
                str���ʱ�� = "To_Date('" & Format(Curdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
            End If
            strNO = mshList.TextMatrix(i, 1)
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_סԺ���ʼ�¼_Verify('" & strNO & "','" & UserInfo.��� & "','" & UserInfo.���� & "',NULL," & mrsInfo!����ID & "," & str���ʱ�� & ")"
            strDel = strDel & "," & i
            
            strNos = strNos & "," & strNO
        End If
    Next
    If UBound(arrSQL) = -1 Then
        MsgBox "û��ѡ��Ҫ��˵Ļ��۵��ݣ�", vbInformation, gstrSysName
        Exit Sub
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
            strNO = Split(strNos, ",")(i)
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133", Me, "NO=" & strNO, "�Ǽ�ʱ��=" & Format(Curdate, "yyyy-MM-dd HH:mm:ss"), "ҩƷ��λ=" & IIf(gblnסԺ��λ, 1, 0), "PrintEmpty=0", "�ش�=0", 2)
        Next
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
    gblnOK = True
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
        Case vbKeyF6
            txtPatient.SetFocus
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    Me.Height = 7815 'û��дResize
    gblnOK = False
    mstrPrivsOpt = ";" & GetInsidePrivs(Enum_Inside_Program.p���ʲ���) & ";"
        
    Call SetHeader
    Call SetBill
    Call initCardSquareData
    mstrUnitIDs = GetUserUnits
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlngUnitID = 0
    Call SaveWinState(Me, App.ProductName)
End Sub


Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    Set gobjSquare.objCurCard = objCard
    If txtPatient.Text <> "" Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

 
Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.����
    Call FindPati(objCard, True, txtPatient.Text)
End Sub

Private Sub mshList_DblClick()
    Dim strNO As String
    
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
        Err.Clear
        On Error Resume Next
            
        strNO = mshList.TextMatrix(mshList.Row, 1)
        If BillisBatch(strNO) Then '��������
            frmBillings.mstrPrivs = mstrPrivs
            frmBillings.mbytInState = 1
            frmBillings.mstrInNO = strNO
            frmBillings.Show 1, Me
        ElseIf BillisSimple(strNO) Then '�򵥼���
            frmSimpleBilling.mstrPrivs = mstrPrivs
            frmSimpleBilling.mbytInState = 1
            frmSimpleBilling.mstrInNO = strNO
            frmSimpleBilling.Show 1, Me
        Else '���ʵ�
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 1
            frmCharge.mstrInNO = strNO
            frmCharge.Show 1, Me
        End If
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

Private Sub ShowDetail(Optional strNO As String)
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim strSQL As String
    
    On Error GoTo errH
    
    '��ϸ�����е�ʣ�������ͽ��
    strSQL = _
    " Select C.���� as ���,Nvl(E.����,B.����) as ����" & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, ",E1.���� as ��Ʒ��", "") & ",B.���," & _
            IIf(gblnסԺ��λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X.סԺ��λ)", "A.���㵥λ") & " as ��λ," & _
    "       Avg(Nvl(A.����,1)*A.����)" & IIf(gblnסԺ��λ, "/Nvl(X.סԺ��װ,1)", "") & " as ����, " & _
    "       Ltrim(To_Char(Sum(A.��׼����)" & IIf(gblnסԺ��λ, "*Nvl(X.סԺ��װ,1)", "") & ",'99999" & gstrFeePrecisionFmt & "')) as ����," & _
    "       Ltrim(To_Char(Sum(A.Ӧ�ս��),'99999" & gstrDec & "')) as Ӧ�ս��," & _
    "       Ltrim(To_Char(Sum(A.ʵ�ս��),'99999" & gstrDec & "')) as ʵ�ս��," & _
    "       D.���� as ִ�п���" & _
    " From סԺ���ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,�շ���Ŀ���� E,ҩƷ��� X" & _
        IIf(gTy_System_Para.bytҩƷ������ʾ = 2, ",�շ���Ŀ���� E1", "") & _
    " Where A.�շ�ϸĿID=B.ID and A.�շ����=C.���� And A.ִ�в���ID=D.ID(+)" & _
    "       And A.NO=[1] And A.��¼����=2 And A.�����־=2 And A.��¼״̬=0" & _
    "       And A.����ID+0=[2] And Nvl(A.��ҳID,0)=[3]" & _
    "       And A.�շ�ϸĿID=X.ҩƷID(+) And A.����Ա���� is NULL And A.������ is Not NULL" & _
    "       And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
    IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E.����(+)=1 And E1.����(+)=3", "") & _
    " Group by Nvl(A.�۸񸸺�,A.���),C.����," & _
    "       Nvl(E.����,B.����)" & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, ",E1.����", "") & ",B.���,A.���㵥λ,D.����,X.ҩƷID,X.סԺ��λ,Nvl(X.סԺ��װ,1)" & _
    " Order by Nvl(A.�۸񸸺�,A.���)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, CLng(mrsInfo!����ID), Val("" & mrsInfo!��ҳID))
    
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
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshBill, App.ProductName & "\" & Me.Name)
        
        For i = 0 To .Cols - 1
            If .TextMatrix(0, i) = "��Ʒ��" Then
                If gTy_System_Para.bytҩƷ������ʾ = 2 Then
                    If .ColWidth(i) <= 0 Then .ColWidth(i) = 2000
                Else
                    .ColWidth(i) = 0
                End If
            End If
        Next
        
        .RowHeight(0) = 320
        .Col = 0: .ColSel = .Cols - 1
    End With
End Sub

Private Sub mshList_KeyPress(KeyAscii As Integer)
    Dim strNO As String
    
    If mshList.TextMatrix(mshList.Row, 1) = "" Then Exit Sub
    
    If KeyAscii = 32 Then
        If mshList.TextMatrix(mshList.Row, 0) = "" Then
            mshList.TextMatrix(mshList.Row, 0) = "��"
        Else
            mshList.TextMatrix(mshList.Row, 0) = ""
        End If
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        Err.Clear
        On Error Resume Next
            
        strNO = mshList.TextMatrix(mshList.Row, 1)
        If BillisBatch(strNO) Then '��������
            frmBillings.mstrPrivs = mstrPrivs
            frmBillings.mbytInState = 1
            frmBillings.mstrInNO = strNO
            frmBillings.Show 1, Me
        ElseIf BillisSimple(strNO) Then '�򵥼���
            frmSimpleBilling.mstrPrivs = mstrPrivs
            frmSimpleBilling.mbytInState = 1
            frmSimpleBilling.mstrInNO = strNO
            frmSimpleBilling.Show 1, Me
        Else '���ʵ�
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 1
            frmCharge.mstrInNO = strNO
            frmCharge.Show 1, Me
        End If
    End If
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
    If txtPatient.Locked Then Exit Sub
    Call IDKIND.SetAutoReadCard(txtPatient.Text = "")
End Sub

Private Sub txtPatient_GotFocus()
    zlControl.TxtSelAll txtPatient
    If txtPatient.Locked Then Exit Sub
    Call IDKIND.SetAutoReadCard(txtPatient.Text = "")
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
'    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
'    If IDKind.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13 Then
        With frmPatiSelect
            If InStr(mstrPrivs, ";���в���;") > 0 Then
                .mlngUnitID = 0
            Else
                .mlngUnitID = mlngUnitID
            End If
            Set .mfrmParent = Me
            .mstrPrivs = mstrPrivs
            .Show 1, Me
        End With
    Else
        If IDKIND.GetCurCard.���� Like "����*" Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKIND.ShowPassText)
        ElseIf IDKIND.IDKIND = IDKIND.GetKindIndex("�����") Or IDKIND.IDKIND = IDKIND.GetKindIndex("סԺ��") Then
            If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
                If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
            End If
            txtPatient.PasswordChar = IIf(IDKIND.ShowPassText, "*", "")
        End If
    End If
    Me.Refresh
    If blnCard And Len(txtPatient.Text) = IDKIND.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        Call FindPati(IDKIND.GetCurCard, blnCard, txtPatient.Text)
    End If
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ҳ���
    '����:���˺�
    '����:2012-08-29 17:53:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnICCard As Boolean, blnIDCard As Boolean
    Dim strErrMsg As String
   '��ȡ������Ϣ
    Call ClearPati
    mshList.Clear: mshList.Rows = 2
    Call SetHeader
    If Not GetPatient(objCard, txtPatient.Text, blnCard, strErrMsg) Then
        txtPatient.Text = ""
        If blnCard Then
            If strErrMsg <> "" Then
                sta.Panels(2) = strErrMsg
            Else
                sta.Panels(2) = "����ȷ��������Ϣ�������Ƿ���ȷˢ����ѡ��Ĳ��˲���סԺ���ˣ�"
            End If
            txtPatient.SetFocus: Exit Sub
        Else
            If strErrMsg <> "" Then
                sta.Panels(2) = strErrMsg
            Else
                sta.Panels(2) = "����ı�ʶ���ܶ�ȡ������Ϣ�����������Ƿ���ȷ��ѡ��Ĳ��˲���סԺ���ˣ�"
            End If
            txtPatient.SetFocus: Exit Sub
        End If
        Exit Sub
    End If
    
    '54899
    If objCard.���� Like "IC��*" And objCard.ϵͳ = True And mstrPassWord <> "" Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If objCard.���� Like "*���֤*" And objCard.ϵͳ = True And mstrPassWord <> "" Then blnIDCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
     If (objCard.���� Like "IC��*" Or objCard.���� Like "*���֤*") And objCard.ϵͳ = True And blnCard Then blnCard = False
    '���￨������
    If Mid(gstrCardPass, 6, 1) = "1" And (blnCard Or blnICCard Or blnIDCard) Then
        If Not zlCommFun.VerifyPassWord(Me, mstrPassWord, mrsInfo!����, mrsInfo!�Ա�, "" & mrsInfo!����) Then
            Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
        End If
    End If
    
    If zlIsAllowFeeChange(Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ҳID)), Val(Nvl(mrsInfo!��˱�־))) = False Then
        Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
    End If
    
    txtPatient.PasswordChar = ""
    txtPatient.Text = IIf(IsNull(mrsInfo!����), "", mrsInfo!����)
    txt�Ա�.Text = IIf(IsNull(mrsInfo!�Ա�), "", mrsInfo!�Ա�)
    txt����.Text = IIf(IsNull(mrsInfo!����), "", mrsInfo!����)
    txt����.Text = IIf(IsNull(mrsInfo!����), "��ͥ����", mrsInfo!����)
    txtסԺ��.Text = IIf(IsNull(mrsInfo!סԺ��), "", mrsInfo!סԺ��)
    txt����.Text = GET��������(mrsInfo!����ID)
    
    txtPatient.ForeColor = zlDatabase.GetPatiColor(Nvl(mrsInfo!��������))
    
    Call RefreshMoney
    Call ShowBills
    mshList.SetFocus
End Sub


Private Sub RefreshMoney()
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = GetMoneyInfo(mrsInfo!����ID, , , 2)
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
    txt����.Text = ""
    txtסԺ��.Text = ""
    txt����.Text = ""
    txt����.Text = ""
    txtԤ��.Text = ""
    txtʣ��.Text = ""
End Sub

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, ByVal blnCard As Boolean, Optional ByRef strOut As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:blnCard=�Ƿ���￨ˢ��
    '����:
    '����:
    '����:���˺�
    '����:2011-08-03 17:34:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng�����ID As Long, strPassWord As String, strErrMsg As String, bln���в��� As Boolean
    Dim strIF As String, strWhere As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    Dim rsOutSel As ADODB.Recordset
    
    On Error GoTo errH
        
    'a.�Ƿ����ǿ�Ƽ���Ȩ��
    If InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") > 0 And InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") > 0 Then
        strIF = ""
    ElseIf InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") > 0 Then
        strIF = " And ((B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3) Or Nvl(X.�������,0)<>0)"
    ElseIf InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") > 0 Then
        strIF = " And ((B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3) Or Nvl(X.�������,0)=0)"
    Else
        strIF = " And B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3"
    End If
    
    'b.�Ƿ���Լ����в�������
    bln���в��� = True
    If InStr(mstrPrivs, ";���в���;") <= 0 Then
        bln���в��� = False
        If InStr(1, mstrUnitIDs, ",") = 0 Then
            strIF = strIF & " And B.��ǰ����ID+0=[3]"
        Else
            strIF = strIF & " And B.��ǰ����ID+0 IN(Select Column_Value From Table(Cast(f_num2list([4]) As zlTools.t_numlist)))"
        End If
    End If
       
    'c.�Ƿ����۲��˼���Ȩ��
    If (InStr(mstrPrivsOpt, ";�������ۼ���;") > 0 And gbln��������) And (InStr(mstrPrivsOpt, ";סԺ���ۼ���;") > 0 And gblnסԺ����) Then
        strIF = strIF & " And Nvl(B.��������,0) IN(0,1,2)"
    ElseIf InStr(mstrPrivsOpt, ";�������ۼ���;") > 0 And gbln�������� Then
        strIF = strIF & " And Nvl(B.��������,0) IN(0,1)"
    ElseIf InStr(mstrPrivsOpt, ";סԺ���ۼ���;") > 0 And gblnסԺ���� Then
        strIF = strIF & " And Nvl(B.��������,0) IN(0,2)"
    Else
        strIF = strIF & " And Nvl(B.��������,0)=0"
    End If
    
    strSQL = _
            "Select A.����ID,B.��ҳID,B.��ǰ����ID as ����ID,B.��Ժ����ID as ����ID,B.��Ժ����,B.��Ժ����," & _
            "   A.���￨��,A.����֤��,A.סԺ��,B.��Ժ���� as ����,X.�������,B.״̬," & _
            "   nvl(B.����,A.����) as ����,nvl(b.�Ա�,A.�Ա�) as �Ա�,A.����,B.�ѱ�,B.סԺҽʦ,B.ҽ�Ƹ��ʽ," & _
            "   A.������,Decode(A.������,null,A.������,Zl_Patientsurety(A.����ID,B.��ҳID)) ������," & _
            "   zl_PatiDayCharge(A.����ID) as ���ն�,B.����,Nvl(B.��������,0) as ��������,B.��������,b.��˱�־" & _
            " From ������Ϣ A,������ҳ B,������� X " & _
            " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID" & _
            "       And Nvl(B.��ҳID,0)<>0 And A.����ID=X.����ID(+) And X.����(+)=1 And X.����(+)=2 And A.ͣ��ʱ�� is NULL " & strIF
    If blnCard = True And objCard.���� Like "����*" Then    'ˢ��
        If IDKIND.Cards.��ȱʡ������ And Not IDKIND.GetfaultCard Is Nothing Then
            lng�����ID = IDKIND.GetfaultCard.�ӿ����
        Else
            lng�����ID = "-1"
        End If
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng����ID
        blnHavePassWord = True
        strWhere = strWhere & " And A.����ID=[1] "
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '����ID
        strWhere = strWhere & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "/" Then   '��λ��
        '41654 And IsNumeric(Mid(strInput, 2))
        strInput = Mid(strInput, 2)
        If mlngUnitID = 0 Then '������ȷ��������ͨ������ȷ������
            Set mrsInfo = New ADODB.Recordset: Exit Function
        End If
        strSQL = _
            "Select A.����ID,B.��ҳID,B.��ǰ����ID as ����ID,B.��Ժ����ID as ����ID,B.��Ժ����,B.��Ժ����," & _
            "   A.���￨��,A.����֤��,A.סԺ��,B.��Ժ���� as ����,X.�������,B.״̬," & _
            "   nvl(B.����,A.����) as ����,nvl(b.�Ա�,A.�Ա�) as �Ա�,A.����,B.�ѱ�,B.סԺҽʦ,B.ҽ�Ƹ��ʽ," & _
            "   A.������,Decode(A.������,null,A.������,Zl_Patientsurety(A.����ID,B.��ҳID)) ������," & _
            "   zl_PatiDayCharge(A.����ID) as ���ն�,B.����,Nvl(B.��������,0) as ��������,B.��������,B.��˱�־" & _
            " From ������Ϣ A,������ҳ B,��λ״����¼ C,������� X" & _
            " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID" & _
            " And Nvl(B.��ҳID,0)<>0 And A.����ID=C.����ID And A.����ID=X.����ID(+) And X.����(+)=1 And X.����(+)=2 And A.ͣ��ʱ�� is NULL " & _
            " And C.����ID=[3] And C.����=[2] " & strIF
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��(������Ժ)
        strWhere = strWhere & " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [1])"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����(ҽ������)
        strWhere = strWhere & " And A.�����=[1]"
    Else
        Select Case objCard.����
            Case "����", "��������￨"
                If mrsInfo.State = 1 Then
                    If mrsInfo.EOF = False Then
                        If mrsInfo!���� = Trim(txtPatient.Text) Then GetPatient = True: Exit Function
                    End If
                End If
                If zlSelectChargePatiFromInputName(Me, mstrPrivsOpt, strInput, bln���в���, mstrUnitIDs, gintOutDay, lng����ID, strOut, txtPatient.hWnd, txtPatient.Height) = False Then
                     Set mrsInfo = New Recordset: Exit Function
                End If
                strInput = "-" & lng����ID
                strWhere = strWhere & " And A.����ID=[1]"
            Case "ҽ����"
                strInput = UCase(strInput)
                strWhere = strWhere & " And A.ҽ����=[2]"
             Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.�����=[2]"
            Case "סԺ��"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [2])"
            Case Else
                '��������,��ȡ��صĲ���IDs
                If objCard.�ӿ���� > 0 Then
                    lng�����ID = objCard.�ӿ����
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng����ID = 0 Then GoTo NotFoundPati:
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng����ID <= 0 Then GoTo NotFoundPati:
                strWhere = strWhere & " And A.����ID=[1]"
                strInput = "-" & lng����ID
                blnHavePassWord = True
        End Select
    End If
       
    strSQL = strSQL & vbCrLf & strWhere
    txtPatient.ForeColor = Me.ForeColor
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput, mlngUnitID, mstrUnitIDs)
    
    If mrsInfo.RecordCount = 0 Then GoTo NotFoundPati:
    mstrPassWord = strPassWord
    If Not blnHavePassWord Then
        mstrPassWord = Nvl(mrsInfo!����֤��)
    End If
    GetPatient = True
    Exit Function
NotFoundPati:
    Set mrsInfo = New ADODB.Recordset
    If strWhere = "" Then Exit Function
    
    'δ�ҵ����ˣ���Ҫ�Ըò��˵ľ��������Ϣ������ʾ
    strSQL = _
    " Select A.����ID,B.��ҳID,B.��ǰ����ID as ����ID,B.��Ժ����ID as ����ID,a.��Ժ,B.��Ժ����,B.��Ժ����,X.�������,B.״̬, " & _
    "       nvl(B.����,A.����) as ����,nvl(b.�Ա�,A.�Ա�) as �Ա�,nvl(b.����,A.����) as ����,B.�ѱ�,Nvl(B.��������,0) as ��������,B.��������" & _
    " From ������Ϣ A,������ҳ B,������� X" & _
    " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID" & _
    "   And Nvl(B.��ҳID,0)<>0 And A.����ID=X.����ID(+) and X.����(+)=1 and X.����(+)=2 And A.ͣ��ʱ�� is NULL " & strWhere
    
    Set rsOutSel = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    If rsOutSel.EOF Then Exit Function
    '1.�������
    If InStr(mstrPrivs, ";���в���;") <= 0 Then
        If InStr(1, "," & mstrUnitIDs & ",", "," & Val(rsOutSel!����ID) & ",") = 0 Then
            strOut = "����:��" & Nvl(rsOutSel!����) & "�������㸺��Ĳ���,���ܶԸò��˽��м��˲���!"
            Exit Function
        End If
    End If
    
    '2.���۲��˼��(�Ƿ����۲��˼���Ȩ��)
    If (InStr(mstrPrivsOpt, ";�������ۼ���;") > 0 And gbln��������) And (InStr(mstrPrivsOpt, ";סԺ���ۼ���;") > 0 And gblnסԺ����) Then
        '0-��ͨסԺ����,1-�������۲���,2-סԺ���۲���
    ElseIf InStr(mstrPrivsOpt, ";�������ۼ���;") > 0 And gbln�������� Then
        If Val(Nvl(rsOutSel!��������)) = 2 Then
            strOut = "����:��" & Nvl(rsOutSel!����) & "��ΪסԺ���۲���,�㲻�߱���סԺ���ۼ��ʡ�Ȩ��,���ܶԸò��˽��м��˲���!"
            Exit Function
        End If
    ElseIf InStr(mstrPrivsOpt, ";סԺ���ۼ���;") > 0 And gblnסԺ���� Then
        If Val(Nvl(rsOutSel!��������)) = 1 Then
            strOut = "����:��" & Nvl(rsOutSel!����) & "��Ϊ�������۲���,�㲻�߱����������ۼ��ʡ�Ȩ��,���ܶԸò��˽��м��˲���!"
            Exit Function
        End If
    Else
        If Val(Nvl(rsOutSel!��������)) <> 0 Then
            strOut = "����:��" & Nvl(rsOutSel!����) & "��Ϊ" & IIf(Val(Nvl(rsOutSel!��������)) = 1, "����", "סԺ") & "���۲���,�㲻�߱��������סԺ ���ۼ��ʡ�Ȩ��,���ܶԸò��˽��м��˲���!"
            Exit Function
        End If
    End If
    
        '124007
    If InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") > 0 And InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") > 0 Then
        strErrMsg = ""
    ElseIf InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") > 0 Then
        If Not (Val(Nvl(rsOutSel!״̬)) <> 3 And IsNull(rsOutSel!��Ժ����) Or Val(Nvl(rsOutSel!�������)) <> 0) Then
              
                If Val(Nvl(rsOutSel!״̬)) = 3 And IsNull(rsOutSel!��Ժ����) Then
                    strErrMsg = "�����Ѿ�Ԥ��Ժ�����ܶԲ��˽��м��˲���!"
                Else
                    strErrMsg = "������" & Format(rsOutSel!��Ժ����, "yyyy��mm��DD��") & " ��Ժ�����ܶԲ��˽��м��˲���!"
                End If
        End If
    ElseIf InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") > 0 Then
        If Not (Val(Nvl(rsOutSel!״̬)) <> 3 And IsNull(rsOutSel!��Ժ����) Or Val(Nvl(rsOutSel!�������)) = 0) Then
                If Val(Nvl(rsOutSel!״̬)) = 3 And IsNull(rsOutSel!��Ժ����) Then
                strErrMsg = "�����Ѿ�Ԥ��Ժ�����ܶԲ��˽��м��˲���!"
                Else
                strErrMsg = "������" & Format(rsOutSel!��Ժ����, "yyyy��mm��DD��") & " ��Ժ�����ܶԲ��˽��м��˲���!"
                End If
        End If
    Else
        If Not (Val(Nvl(rsOutSel!״̬)) <> 3 And IsNull(rsOutSel!��Ժ����)) Then
            If Val(Nvl(rsOutSel!״̬)) = 3 And IsNull(rsOutSel!��Ժ����) Then
                strErrMsg = "�����Ѿ�Ԥ��Ժ�����ܶԲ��˽��м��˲���!"
            Else
                strErrMsg = "������" & Format(rsOutSel!��Ժ����, "yyyy��mm��DD��") & " ��Ժ�����ܶԲ��˽��м��˲���!"
            End If
        End If
    End If
    If strErrMsg <> "" Then
        strOut = strErrMsg
        Exit Function
    End If
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
        " From סԺ���ü�¼ A,���ű� B" & _
        " Where A.��¼����=2 And A.�����־=2 And A.��¼״̬=0" & _
        "       And A.������ is Not Null And A.����Ա���� is NULL" & _
        "       And A.��������ID=B.ID" & _
        "       And A.����ID=[1] And Nvl(A.��ҳID,0)=[2]" & _
        " Group by A.NO,B.����,A.������,A.�ѱ�,A.�Ǽ�ʱ��,A.������" & _
        " Order by ����ʱ�� Desc,���ݺ� Desc"
        Set mrsList = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsInfo!����ID), Val("" & mrsInfo!��ҳID))
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
        .Cols = UBound(Split(strHead, "|")) + 1
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
        
        i = MshGetColNum(mshList, "ҽ��")
        'If InStr(mstrPrivsOpt, "ҽ����ѯ") = 0 Then .ColWidth(i) = 0
        
        .Col = 0: .ColSel = .Cols - 1
                
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
        If mshList.TextMatrix(0, lngCol) = "���" Then
           mshList.Col = lngCol
            If mshList.ColData(lngCol) = 1 Then
                mshList.Sort = flexSortStringNoCaseAscending
            Else
               mshList.Sort = flexSortStringNoCaseDescending
            End If
            mshList.ColData(lngCol) = (mshList.ColData(lngCol) + 1) Mod 2
            Exit Sub
        End If
        
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
        " From סԺ���ü�¼ A,�շ���Ŀ��� B" & _
        " Where A.��¼����=2 And A.�����־=2 And A.��¼״̬=0" & _
        " And A.�շ����=B.���� And A.������ is Not Null And A.����Ա���� is NULL" & _
        IIf(strNos <> "", " And Instr(','||[3]||',',','||A.NO||',')>0", "") & _
        " And A.����ID=[1] And Nvl(A.��ҳID,0)=[2]" & _
        " Group by A.�շ����,B.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mrsInfo!����ID), Val("" & mrsInfo!��ҳID), strNos)
    
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
        strSQL = "Select B.��ǰ����ID ����ID,A.סԺ��,A.��ǰ���� As ����,nvl(B.����,A.����) as ����,C.Ԥ�����-C.������� as ���,zl_PatiDayCharge(A.����ID) as ���ն�," & _
            " Decode(A.������,null,A.������,Zl_Patientsurety(A.����ID,B.��ҳID)) ������,Zl_Patiwarnscheme(B.����id, B.��ҳid) As ���ò���" & _
            " From ������Ϣ A,������ҳ B,������� C" & _
            " Where A.����ID=B.����ID(+) And Nvl(A.��ҳID,0)=B.��ҳID(+)" & _
            " And A.����ID=C.����ID(+) And C.����(+)=1 And C.����(+)=2 " & _
            " And A.����ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mrsInfo!����ID))
        Set rsWarn = GetUnitWarn(rsTmp!���ò���, Val(Nvl(rsTmp!����ID)))    '����:43862
        
        cur��� = Nvl(rsTmp!���, 0)
        If gbln�����������۷��� Then cur��� = Nvl(rsTmp!���, 0) - GetPriceMoneyTotal(1, mrsInfo!����ID) + cur���
        '���౨��
        For i = 0 To UBound(Split(str���s, ","))
            intWarn = BillingWarn(mstrPrivsOpt, rsTmp!���� & IIf(Nvl(rsTmp!סԺ��) = "", "", "(סԺ��:" & rsTmp!סԺ�� & " ����:" & rsTmp!���� & ")"), Val("" & rsTmp!����ID), rsTmp!���ò���, rsWarn, _
                cur���, Nvl(rsTmp!���ն�, 0), cur���, Nvl(rsTmp!������, 0), _
                Left(Split(str���s, ",")(i), 1), Mid(Split(str���s, ",")(i), 2), strWarn)
            If intWarn = 2 Or intWarn = 3 Then Exit Function
        Next
    End If
    AuditingWarnByPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub initCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���㿨����������Ϣ
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    Call IDKIND.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    Set objCard = IDKIND.GetfaultCard
    Set gobjSquare.objDefaultCard = objCard
    If IDKIND.Cards.��ȱʡ������ And Not objCard Is Nothing Then
        gobjSquare.blnȱʡ�������� = objCard.�������Ĺ��� <> ""
        gobjSquare.intȱʡ���ų��� = objCard.���ų���
    Else
        gobjSquare.blnȱʡ�������� = IDKIND.Cards.������ʾ
        gobjSquare.intȱʡ���ų��� = 100
    End If
    gobjSquare.bln��ȱʡ������ = IDKIND.Cards.��ȱʡ������
End Sub
Private Sub txtPatient_LostFocus()
    Call IDKIND.SetAutoReadCard(False)
End Sub
