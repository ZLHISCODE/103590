VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmBabyReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�������Ǽ�"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9915
   Icon            =   "frmBabyReg.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   9915
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "��ӡ����(&S)"
      Height          =   350
      Left            =   2520
      TabIndex        =   45
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelivery 
      Caption         =   "������Ϣ(&F)"
      Height          =   350
      Left            =   5160
      TabIndex        =   14
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Frame fraSplit 
      Height          =   75
      Left            =   0
      TabIndex        =   43
      Top             =   4880
      Width           =   10680
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8640
      TabIndex        =   16
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7440
      TabIndex        =   15
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��ӡ���(&P)"
      Height          =   350
      Left            =   3840
      TabIndex        =   13
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "����(&A)"
      Height          =   350
      Left            =   120
      TabIndex        =   11
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "ɾ��(&D)"
      Height          =   350
      Left            =   1320
      TabIndex        =   12
      Top             =   5040
      Width           =   1100
   End
   Begin VB.Frame fraMotherInfo 
      Height          =   855
      Left            =   120
      TabIndex        =   30
      Top             =   0
      Width           =   9660
      Begin VB.Label lblOut 
         AutoSize        =   -1  'True
         Caption         =   "25��x"
         Height          =   180
         Index           =   2
         Left            =   7920
         TabIndex        =   42
         Top             =   240
         Width           =   450
      End
      Begin VB.Label lblOut 
         AutoSize        =   -1  'True
         Caption         =   "������x"
         Height          =   180
         Index           =   4
         Left            =   4440
         TabIndex        =   41
         Top             =   540
         Width           =   630
      End
      Begin VB.Label lblOut 
         AutoSize        =   -1  'True
         Caption         =   "�и�1x"
         Height          =   180
         Index           =   1
         Left            =   4440
         TabIndex        =   40
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lblOut 
         AutoSize        =   -1  'True
         Caption         =   "������xx"
         Height          =   180
         Index           =   3
         Left            =   960
         TabIndex        =   39
         Top             =   540
         Width           =   720
      End
      Begin VB.Label lblOut 
         AutoSize        =   -1  'True
         Caption         =   "20101118xx"
         Height          =   180
         Index           =   0
         Left            =   960
         TabIndex        =   38
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��  �ң�"
         Height          =   180
         Left            =   240
         TabIndex        =   35
         Top             =   540
         Width           =   720
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��  ����"
         Height          =   180
         Left            =   3720
         TabIndex        =   34
         Top             =   540
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��  �䣺"
         Height          =   180
         Left            =   7200
         TabIndex        =   33
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��  ����"
         Height          =   180
         Left            =   3720
         TabIndex        =   32
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lbl��ʶ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ�ţ�"
         Height          =   180
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   720
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsBaby 
      Height          =   1845
      Left            =   120
      TabIndex        =   24
      Top             =   960
      Width           =   9660
      _cx             =   17039
      _cy             =   3254
      Appearance      =   3
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
      BackColorSel    =   16444122
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
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
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmBabyReg.frx":058A
      ScrollTrack     =   -1  'True
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
   Begin VB.Frame fraBabyInput 
      Height          =   1935
      Left            =   120
      TabIndex        =   17
      Top             =   2880
      Width           =   9660
      Begin VB.TextBox txtBaby 
         BackColor       =   &H8000000E&
         Height          =   300
         Index           =   10
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1440
         Width           =   1575
      End
      Begin VB.ComboBox cboBaby 
         Height          =   300
         Index           =   9
         Left            =   7800
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtBaby 
         BackColor       =   &H8000000E&
         Height          =   300
         Index           =   3
         Left            =   4560
         MaxLength       =   100
         TabIndex        =   10
         Top             =   1440
         Width           =   4575
      End
      Begin VB.TextBox txtBaby 
         BackColor       =   &H8000000E&
         Height          =   300
         Index           =   0
         Left            =   1080
         MaxLength       =   60
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cboBaby 
         Height          =   300
         Index           =   8
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
      End
      Begin VB.ComboBox cboBaby 
         Height          =   300
         Index           =   7
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox cboBaby 
         Height          =   300
         Index           =   6
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtBaby 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Height          =   300
         Index           =   5
         Left            =   7800
         MaxLength       =   10
         TabIndex        =   8
         Top             =   720
         Width           =   810
      End
      Begin VB.TextBox txtBaby 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Height          =   300
         Index           =   1
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   1
         Top             =   720
         Width           =   810
      End
      Begin VB.TextBox txtBaby 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Height          =   300
         Index           =   4
         Left            =   7800
         MaxLength       =   10
         TabIndex        =   7
         Top             =   360
         Width           =   810
      End
      Begin VB.TextBox txtBaby 
         BackColor       =   &H8000000E&
         Height          =   300
         Index           =   2
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
         Height          =   180
         Index           =   13
         Left            =   240
         TabIndex        =   46
         Top             =   1500
         Width           =   720
      End
      Begin VB.Label lblERRInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   240
         TabIndex        =   44
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   9
         Left            =   8760
         TabIndex        =   37
         Top             =   420
         Width           =   360
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Index           =   10
         Left            =   8760
         TabIndex        =   36
         Top             =   780
         Width           =   180
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ע˵��"
         Height          =   180
         Index           =   12
         Left            =   3720
         TabIndex        =   29
         Top             =   1500
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
         Height          =   180
         Index           =   11
         Left            =   240
         TabIndex        =   28
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ѫ  ��"
         Height          =   180
         Index           =   7
         Left            =   7200
         TabIndex        =   27
         Top             =   1140
         Width           =   540
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "̥��״��"
         Height          =   180
         Index           =   4
         Left            =   3720
         TabIndex        =   26
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ӥ������"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   25
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Index           =   8
         Left            =   1920
         TabIndex        =   23
         Top             =   780
         Width           =   180
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��  ��"
         Height          =   180
         Index           =   6
         Left            =   7200
         TabIndex        =   22
         Top             =   780
         Width           =   540
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��  ��"
         Height          =   180
         Index           =   5
         Left            =   7200
         TabIndex        =   21
         Top             =   420
         Width           =   540
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���䷽ʽ"
         Height          =   180
         Index           =   3
         Left            =   3720
         TabIndex        =   20
         Top             =   780
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   780
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ӥ���Ա�"
         Height          =   180
         Index           =   1
         Left            =   3720
         TabIndex        =   18
         Top             =   420
         Width           =   720
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000004&
      X1              =   0
      X2              =   8280
      Y1              =   4920
      Y2              =   4920
   End
End
Attribute VB_Name = "frmBabyReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_BABY = 9
Private mlng����ID As Long
Private mlng����ID As Long
Private mbln���� As Boolean
Private mblnChange As Boolean
Private mblnOK As Boolean
Private mstrPrivs As String
Private marrDelBaby() As Variant
Private mblnWristletPrint As Boolean    '�Ƿ��ӡ�������
Private mfrmParent As Object
Private mcolBaby As Collection     '���ڸ�������ֵ��λ��ʾ�� KEY(����):VALUE��_�кţ�

Private Const M_CON_ColorUnEnabled = &H80000016
Private Const M_CON_ColorEnabled = &H8000000E

Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML

Private Enum mCol
    col��� = 0
    ColӤ������ = 1
    ColӤ���Ա� = 2
    Col������� = 3
    Col���䷽ʽ = 4
    Col̥��״�� = 5
    Col�� = 6
    Col���� = 7
    COlѪ�� = 8
    Col����ʱ�� = 9
    Col����ʱ�� = 10
    Col��ע˵�� = 11
End Enum

Private Enum M_E_SHOW
    IX_��ʶ�� = 0
    IX_���� = 1
    IX_���� = 2
    IX_���� = 3
    IX_���� = 4
End Enum

Private Enum M_E_BABY
    B_���� = 0
    B_������� = 1
    B_����ʱ�� = 2
    B_��ע˵�� = 3
    B_�� = 4
    B_���� = 5
    
    B_�Ա� = 6
    B_���䷽ʽ = 7
    B_̥��״�� = 8
    B_Ѫ�� = 9
    B_����ʱ�� = 10
End Enum

Public Function ShowMe(ByVal lng����ID As Long, ByVal lng����ID As Long, strPrivs As String, frmParent As Object, Optional bln���� As Boolean) As Boolean
'������lng����ID=סԺ����Ϊ��ҳID�����ﲡ��Ϊ�Һ�ID��
    mlng����ID = lng����ID
    mlng����ID = lng����ID
    mbln���� = bln����
    mstrPrivs = strPrivs
    Set mfrmParent = frmParent
    
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Sub cboBaby_Click(Index As Integer)
    If Me.Visible Then
        ChangeBabyInfo vsBaby.Row, mcolBaby("_" & Index), cboBaby(Index)
    End If
End Sub

Private Sub cboBaby_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call ChangeBabyInfo(vsBaby.Row, mcolBaby("_" & Index), cboBaby(Index))
        Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
End Sub

Private Sub cmdAdd_Click()
    Call AddNewBabyRow
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    
    On Error GoTo errH

    With vsBaby
        If .Rows = .FixedRows Or .Row <= .FixedRows - 1 Then Exit Sub
        If .RowData(.Row) <> 0 Then
            '��Ӥ��ҵ�����ݽ��м��
            If mbln���� Then
                'by lesfeng 2009-12-29 �����  ���˷��ü�¼ --��������ü�¼ ����ֻ������ "û����ҳid ȥ��And A.��ҳID is Null"
                strSQL = _
                    " Select Distinct 1 as ��־,A.Ӥ���� as Ӥ�� From ������ü�¼ A,���˹Һż�¼ B" & _
                    "   Where A.����ID=[1]  And B.ID=[2] And A.Ӥ����>=[3] And A.�Ǽ�ʱ��>=B.�Ǽ�ʱ�� and B.��¼����=1 and B.��¼״̬=1 And A.��¼״̬ =1 " & _
                    " Union ALL Select Distinct 2,A.Ӥ�� From ����ҽ����¼ A,���˹Һż�¼ B Where A.����ID=[1] And A.�Һŵ�=B.NO And B.ID=[2] And A.Ӥ��>=[3] And B.��¼����=1 and B.��¼״̬=1 And A.ҽ��״̬ <> 4" & _
                    " Union ALL Select Distinct 3,Ӥ�� From ���Ӳ�����¼ Where ����ID=[1] And ��ҳID=[2] And Ӥ��>=[3]" & _
                    " Union ALL Select Distinct 4,Ӥ�� From ���˻����¼ Where ����ID=[1] And ��ҳID=[2] And Ӥ��>=[3]" & _
                    " Union ALL Select Distinct 4,Ӥ�� From ���˻����ļ� Where ����ID=[1] And ��ҳID=[2] And Ӥ��>=[3]"
            Else
                'by lesfeng 2009-12-29 �����  ���˷��ü�¼ --��סԺ���ü�¼ ����ֻ��סԺ
                strSQL = _
                    " Select Distinct 1 as ��־,Ӥ���� as Ӥ�� From סԺ���ü�¼ Where ����ID=[1] And ��ҳID=[2] And Ӥ����>=[3] And ��¼״̬ = 1" & _
                    " Union ALL Select Distinct 2,Ӥ�� From ����ҽ����¼ Where ����ID=[1] And ��ҳID=[2] And Ӥ��>=[3] And ҽ��״̬ <> 4" & _
                    " Union ALL Select Distinct 3,Ӥ�� From ���Ӳ�����¼ Where ����ID=[1] And ��ҳID=[2] And Ӥ��>=[3]" & _
                    " Union ALL Select Distinct 4,Ӥ�� From ���˻����¼ Where ����ID=[1] And ��ҳID=[2] And Ӥ��>=[3]" & _
                    " Union ALL Select Distinct 4,Ӥ�� From ���˻����ļ� Where ����ID=[1] And ��ҳID=[2] And Ӥ��>=[3]"
            End If
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng����ID, Val(.RowData(.Row)))
            If Not rsTmp.EOF Then
                MsgBox "�ò��˵�Ӥ��" & rsTmp!Ӥ�� & "�Ѵ�����Ч��" & Decode(rsTmp!��־, 1, "����", 2, "ҽ��", 3, "����", 4, "����") & "���ݣ���ǰ�в���ɾ����", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        If MsgBox("ȷʵҪɾ��Ӥ��" & .TextMatrix(.Row, mCol.col���) & _
            IIf(.TextMatrix(.Row, mCol.ColӤ������) <> "", """" & .TextMatrix(.Row, mCol.ColӤ������) & """", "") & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        For i = .Row + 1 To .Rows - 1
            .TextMatrix(i, mCol.col���) = .TextMatrix(i, mCol.col���) - 1
        Next
        i = .Row
        ReDim Preserve marrDelBaby(UBound(marrDelBaby) + 1)
        marrDelBaby(UBound(marrDelBaby)) = .RowData(.Row)
        SetBabyInfo .Row, 2   '�����Ƭ��Ϣ
        .RemoveItem .Row
        .Row = IIf(i <= .Rows - 1, i, .Rows - 1)
        mblnChange = True
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdDelivery_Click()
    Dim objMedRecPage As zlMedRecPage.clsInOutMedRec
    
    Set objMedRecPage = New zlMedRecPage.clsInOutMedRec
    Call objMedRecPage.InitMedRec(gcnOracle, glngSys, glngModul)
    Call objMedRecPage.EditDelivery(Me, mlng����ID, mlng����ID)
End Sub

Private Sub cmdOK_Click()
    Dim arrSQL As Variant, arrBaby As Variant, arrItem As Variant
    Dim blnTrans As Boolean
    Dim str����ʱ�� As String
    Dim strSQL As String, strNO As String, strErr As String
    Dim intAddCount As Integer
    Dim i As Long, j As Long
    Dim rsTmp As ADODB.Recordset
    Dim blnLis As Boolean
    Dim strDieDate As String

    If Not CheckBaby() Then Exit Sub
    blnLis = Sys.IsSysSetUp(2500)
    arrSQL = Array()
    If blnLis Then arrBaby = Array()
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_������������¼_Delete(" & mlng����ID & "," & mlng����ID & ")"
    With vsBaby
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, mCol.ColӤ������) = "" And .TextMatrix(i, mCol.ColӤ���Ա�) = "" _
                And .TextMatrix(i, mCol.Col���䷽ʽ) = "" And .TextMatrix(i, mCol.Col̥��״��) = "" _
                And .TextMatrix(i, mCol.Col��) = "" And .TextMatrix(i, mCol.Col����) = "" And .TextMatrix(i, mCol.COlѪ��) = "" Then
                MsgBox "Ӥ��" & .TextMatrix(i, mCol.col���) & "����Ϣ¼�벻������", vbInformation, gstrSysName
                .Row = i: .ShowCell .Row, .Col: .SetFocus: Exit Sub
            End If
            If .TextMatrix(i, mCol.Col����ʱ��) <> "" And .TextMatrix(i, mCol.Col����ʱ��) <> "" Then
                If Format(.TextMatrix(i, mCol.Col����ʱ��), "YYYY-MM-dd HH:mm") > Format(.TextMatrix(i, mCol.Col����ʱ��), "YYYY-MM-dd HH:mm") Then
                    MsgBox "Ӥ��" & .TextMatrix(i, mCol.col���) & "�ġ�����ʱ�䡿����С�ڡ�����ʱ�䡿��", vbInformation, gstrSysName
                    .Row = i: .ShowCell .Row, .Col: .SetFocus: Exit Sub
                End If
            End If
            If .TextMatrix(i, mCol.Col����ʱ��) = "" Then
                strDieDate = ""
            Else
                strDieDate = ",to_date('" & Mid(.TextMatrix(i, mCol.Col����ʱ��), 1, 10) & " " & Val(Mid(.TextMatrix(i, mCol.Col����ʱ��), 12, 2)) & _
                ":" & Val(Mid(.TextMatrix(i, mCol.Col����ʱ��), 15, 2)) & "','yyyy-mm-dd hh24:mi:ss')"
            End If
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_������������¼_Insert(" & _
                mlng����ID & "," & mlng����ID & "," & .TextMatrix(i, mCol.col���) & "," & _
                "'" & .TextMatrix(i, mCol.ColӤ������) & "','" & zlCommFun.GetNeedName(.TextMatrix(i, mCol.ColӤ���Ա�)) & "'," & _
                ZVal(.TextMatrix(i, mCol.Col�������)) & ",'" & zlCommFun.GetNeedName(.TextMatrix(i, mCol.Col���䷽ʽ)) & "'," & _
                "'" & zlCommFun.GetNeedName(.TextMatrix(i, mCol.Col̥��״��)) & "',to_date('" & _
                Mid(.TextMatrix(i, mCol.Col����ʱ��), 1, 10) & " " & Val(Mid(.TextMatrix(i, mCol.Col����ʱ��), 12, 2)) & _
                ":" & Val(Mid(.TextMatrix(i, mCol.Col����ʱ��), 15, 2)) & "','yyyy-mm-dd hh24:mi:ss')," & ZVal(.TextMatrix(i, mCol.Col��)) & _
                "," & ZVal(.TextMatrix(i, mCol.Col����)) & ",'" & zlCommFun.GetNeedName(.TextMatrix(i, mCol.COlѪ��)) & "','" & _
                .TextMatrix(i, mCol.Col��ע˵��) & "'" & strDieDate & ")"
            If blnLis Then
                ReDim Preserve arrBaby(UBound(arrBaby) + 1)
                arrBaby(UBound(arrBaby)) = .TextMatrix(i, mCol.col���) & ";" & .TextMatrix(i, mCol.ColӤ������) & ";" & zlCommFun.GetNeedName(.TextMatrix(i, mCol.ColӤ���Ա�))
            End If
            If .RowData(i) = "" Then intAddCount = intAddCount + 1
        Next
    End With
    
    On Error GoTo errH
    
    If blnLis Then
        '���ύ����֮ǰ��ʼ�����˹�����������
        If CreatePublicPatient() Then
            If Not gobjPublicPatient.InitLis(True) Then Exit Sub
        Else
            Exit Sub
        End If
        
        If mbln���� Then
            strSQL = "Select a.NO  From ���˹Һż�¼ A Where a.Id = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
            If Not rsTmp.EOF Then strNO = rsTmp!NO & ""
        End If
        
        strSQL = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From �Ա� Order by ���� "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        
        For j = LBound(arrBaby) To UBound(arrBaby)
            arrItem = Split(arrBaby(j), ";")
            rsTmp.Filter = "����='" & arrItem(2) & "'"
            If Not rsTmp.EOF Then arrItem(2) = rsTmp!����
            arrBaby(j) = arrItem(0) & ";" & arrItem(1) & ";" & arrItem(2)    'ÿ��Ӥ����Ӧ����š��������Ա�
        Next
    End If
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    '����Ӥ��ҽ����Ҫͬ���޸�LIS��Ӥ��������
    If blnLis Then
        For i = LBound(arrBaby) To UBound(arrBaby)
             arrItem = Split(CStr(arrBaby(i)), ";")
             If Not gobjPublicPatient.ModifyBabyInfo(mlng����ID, IIf(mbln����, 0, mlng����ID), IIf(mbln����, strNO, ""), CLng(arrItem(0)), arrItem(1), arrItem(2), strErr) Then
                 gcnOracle.RollbackTrans: blnTrans = False
                 MsgBox "LIS ϵͳ����������Ϣ����ʧ�ܣ�" & vbCrLf & IIf(strErr <> "", "����ԭ��:" & strErr, ""), vbOKOnly + vbInformation, Me.Caption
                 Exit Sub
             End If
         Next
     End If
    gcnOracle.CommitTrans: blnTrans = False
    
    On Error Resume Next
    
    If mbln���� = False Then
        For i = 0 To UBound(marrDelBaby)
            If mclsMipModule.IsConnect = True Then
                mclsXML.ClearXmlText '��������е�XML
                'patient_id      ����id  1   N
                mclsXML.appendData "patient_id", mlng����ID, xsNumber
                'page_id     ��ҳid  1   N
                mclsXML.appendData "page_id", mlng����ID, xsNumber
                'baby_serial     ���    1   N
                mclsXML.appendData "baby_serial", CInt(marrDelBaby(i)), xsNumber
                mclsMipModule.CommitMessage "ZLHIS_PATIENT_013", mclsXML.XmlText
            End If
        Next i
        '�������Ǽǻ�ɾ��������Ϣ
        For i = vsBaby.FixedRows To vsBaby.Rows - 1
            If mclsMipModule.IsConnect = True And Val(vsBaby.TextMatrix(i, mCol.col���)) <> Val(vsBaby.RowData(i)) Then
                'ɾ��Ӥ���ᵼ�º���Ӥ������ŷ����仯
                If Val(vsBaby.TextMatrix(i, mCol.col���)) <> Val(vsBaby.RowData(i)) And Val(vsBaby.RowData(i)) <> 0 Then
                    mclsXML.ClearXmlText '��������е�XML
                    'patient_id      ����id  1   N
                    mclsXML.appendData "patient_id", mlng����ID, xsNumber
                    'page_id     ��ҳid  1   N
                    mclsXML.appendData "page_id", mlng����ID, xsNumber
                    'baby_serial     ���    1   N
                    mclsXML.appendData "baby_serial", Val(vsBaby.RowData(i)), xsNumber
                    mclsMipModule.CommitMessage "ZLHIS_PATIENT_013", mclsXML.XmlText
                End If
                
                mclsXML.ClearXmlText '��������е�XML
                'in_patient 1
                mclsXML.AppendNode "in_patient"
                'patient_id      ����id  1   N
                mclsXML.appendData "patient_id", mlng����ID, xsNumber
                'page_id     ��ҳid  1   N
                mclsXML.appendData "page_id", mlng����ID, xsNumber
                'patient_name        ����    1   S
                mclsXML.appendData "patient_name", lblOut(IX_����).Caption, xsString
                'patient_sex     �Ա�    0..1    S
                mclsXML.appendData "patient_sex", lblOut(IX_����).Tag, xsString
                'in_number       סԺ��  1   S
                mclsXML.appendData "in_number", lblOut(IX_��ʶ��).Caption, xsString
                mclsXML.AppendNode "in_patient", True
                'patient_baby 1
                mclsXML.AppendNode "patient_baby"
                'baby_serial     ���    1   N
                mclsXML.appendData "baby_serial", Val(vsBaby.TextMatrix(i, mCol.col���)), xsNumber
                'baby_name       ����    1   S
                mclsXML.appendData "baby_name", vsBaby.TextMatrix(i, mCol.ColӤ������), xsString
                'baby_sex        �Ա�    0..1    S
                mclsXML.appendData "baby_sex", zlCommFun.GetNeedName(vsBaby.TextMatrix(i, mCol.ColӤ���Ա�)), xsString
                'baby_birth      ��������    0..1    S
                str����ʱ�� = Format(Mid(vsBaby.TextMatrix(i, mCol.Col����ʱ��), 1, 10) & " " & Val(Mid(vsBaby.TextMatrix(i, mCol.Col����ʱ��), 12, 2)) & _
                    ":" & Val(Mid(vsBaby.TextMatrix(i, mCol.Col����ʱ��), 15, 2)), "YYYY-MM-DD HH:mm:ss")
                mclsXML.appendData "baby_birth", str����ʱ��, xsString
                mclsXML.AppendNode "patient_baby", True
                mclsMipModule.CommitMessage "ZLHIS_PATIENT_011", mclsXML.XmlText
            End If
        Next i
    End If
    If Err <> 0 Then Err.Clear
    
    On Error GoTo errH
    
    '��ӡ�������
    If InStr(mstrPrivs, "Ӥ�������ӡ") Then
        mblnWristletPrint = True
        If gbytBabyWristletPrint = 0 Then
            mblnWristletPrint = False
        Else
            If gbytBabyWristletPrint = 2 And intAddCount > 0 Then
                If MsgBox("�Ƿ��ӡ���������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    mblnWristletPrint = False
                End If
            End If
        End If
        
        If mblnWristletPrint Then

            With vsBaby
                For i = .FixedRows To .Rows - 1
                    If (.RowData(i) = "") Then
                        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_3", Me, "����ID=" & mlng����ID, "��ҳID=" & mlng����ID, "���=" & .TextMatrix(i, mCol.col���), 2)
                    End If
                Next
            End With
        End If
    End If
    
    mblnChange = False
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdPrint_Click()
    With vsBaby
        If .RowData(.Row) = "" Then MsgBox "������������Ҫ��ȷ��ʱ���ܴ�ӡ�����", vbInformation + vbOKOnly, gstrSysName: Exit Sub
        
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_3", Me, "����ID=" & mlng����ID, "��ҳID=" & mlng����ID, "���=" & .TextMatrix(.Row, mCol.col���), 2)
    End With
End Sub

Private Sub cmdPrintSet_Click()
'����:�����ӡ����
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_3", Me)
End Sub

Private Sub Form_Activate()
    '�����б䶯,ʹ��Ƭֵ�뵱ǰ�б���һ��
    If vsBaby.Rows > 1 Then
        vsBaby.Row = 0
        vsBaby.Row = vsBaby.Rows - 1
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    
    mblnChange = False
    mblnOK = False
    marrDelBaby = Array()
    
    On Error GoTo errH
    
    '������Ϣ
    If mbln���� Then
        lbl��ʶ��.Caption = "�����"
        lbl����.Caption = "����"
        lbl����.Visible = False: lblOut(IX_����).Visible = False
        
        strSQL = "Select B.����� as ��ʶ��,B.����,B.�Ա�,B.����,C.���� as ����,NULL as ����" & _
            " From ���˹Һż�¼ B,���ű� C" & _
            " Where B.ִ�в���ID=C.ID And B.ID=[1] And B.��¼����=1 and B.��¼״̬=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    Else
        strSQL = "Select B.סԺ�� as ��ʶ��,NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա�,B.����,C.���� as ����,D.���� as ����" & _
            " From ������Ϣ A,������ҳ B,���ű� C,���ű� D" & _
            " Where A.����ID=B.����ID And B.��Ժ����ID=C.ID And B.��ǰ����ID=D.ID" & _
            " And A.����ID=[1] And B.��ҳID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng����ID)
    End If
    
    lblOut(IX_��ʶ��).Caption = Nvl(rsTmp!��ʶ��)
    lblOut(IX_����).Caption = Nvl(rsTmp!����)
    lblOut(IX_����).Caption = Nvl(rsTmp!����)
    lblOut(IX_����).Tag = Nvl(rsTmp!�Ա�)
    lblOut(IX_����).Caption = Nvl(rsTmp!����)
    lblOut(IX_����) = Nvl(rsTmp!����)
    
    'Ӥ����Ϣ
    strSQL = "Select ���,Ӥ������,Ӥ���Ա�,�������,���䷽ʽ,̥��״��,��,����,Ѫ��,����ʱ��,����ʱ��,��ע˵��" & _
        " From ������������¼ Where ����ID=[1] And ��ҳID=[2] Order by ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng����ID)
    With vsBaby
        If Not rsTmp.EOF Then
            .Rows = .FixedRows + rsTmp.RecordCount
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, mCol.col���) = rsTmp!���
                .TextMatrix(i, mCol.ColӤ������) = Nvl(rsTmp!Ӥ������)
                .TextMatrix(i, mCol.ColӤ���Ա�) = Nvl(rsTmp!Ӥ���Ա�)
                .TextMatrix(i, mCol.Col�������) = Nvl(rsTmp!�������)
                .TextMatrix(i, mCol.Col���䷽ʽ) = Nvl(rsTmp!���䷽ʽ)
                .TextMatrix(i, mCol.Col̥��״��) = Nvl(rsTmp!̥��״��)
                .TextMatrix(i, mCol.Col��) = gclsBase.FormatEx(rsTmp!��, 2)  '����
                .TextMatrix(i, mCol.Col����) = gclsBase.FormatEx(rsTmp!����, 2)  '��
                .TextMatrix(i, mCol.COlѪ��) = Nvl(rsTmp!Ѫ��)
                .TextMatrix(i, mCol.Col����ʱ��) = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm")
                .TextMatrix(i, mCol.Col����ʱ��) = Format(Nvl(rsTmp!����ʱ��), "yyyy-MM-dd HH:mm")
                .TextMatrix(i, mCol.Col��ע˵��) = Nvl(rsTmp!��ע˵��)
                .RowData(i) = Val(rsTmp!���) '��������������
                
                '��������
                .Cell(flexcpData, i, mCol.ColӤ������) = Nvl(rsTmp!Ӥ������)
                .Cell(flexcpData, i, mCol.ColӤ���Ա�) = Nvl(rsTmp!Ӥ���Ա�)
                .Cell(flexcpData, i, mCol.Col�������) = Nvl(rsTmp!�������)
                .Cell(flexcpData, i, mCol.Col���䷽ʽ) = Nvl(rsTmp!���䷽ʽ)
                .Cell(flexcpData, i, mCol.Col̥��״��) = Nvl(rsTmp!̥��״��)
                .Cell(flexcpData, i, mCol.Col��) = gclsBase.FormatEx(rsTmp!��, 2)
                .Cell(flexcpData, i, mCol.Col����) = gclsBase.FormatEx(rsTmp!����, 2)
                .Cell(flexcpData, i, mCol.COlѪ��) = Nvl(rsTmp!Ѫ��)
                .Cell(flexcpData, i, mCol.Col����ʱ��) = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm")
                .Cell(flexcpData, i, mCol.Col����ʱ��) = Format(Nvl(rsTmp!����ʱ��), "yyyy-MM-dd HH:mm")
                .Cell(flexcpData, i, mCol.Col��ע˵��) = Nvl(rsTmp!��ע˵��)
                rsTmp.MoveNext
            Next
        Else
            .Rows = 1
        End If

        Call SetCardEnable(.Rows > 1)
    End With
    
    
    '�Ա�ѡ��
    Call ReadDict("�Ա�", cboBaby(B_�Ա�))
    '���䷽ʽѡ��
    Call ReadDict("���䷽ʽ", cboBaby(B_���䷽ʽ))
    '̥��״��ѡ��
    Call ReadDict("̥��״��", cboBaby(B_̥��״��))
    'Ѫ�ͼ���
    Call ReadDict("Ѫ��", cboBaby(B_Ѫ��))
    
    'ͨ����Ƭ�����ҵ��к�
    Set mcolBaby = New Collection
    
    With mcolBaby
        .Add mCol.ColӤ������, "_" & B_����
        .Add mCol.Col�������, "_" & B_�������
        .Add mCol.Col����ʱ��, "_" & B_����ʱ��
        .Add mCol.Col��ע˵��, "_" & B_��ע˵��
        .Add mCol.Col��, "_" & B_��
        .Add mCol.Col����, "_" & B_����
        .Add mCol.Col���䷽ʽ, "_" & B_���䷽ʽ
        .Add mCol.Col̥��״��, "_" & B_̥��״��
        .Add mCol.COlѪ��, "_" & B_Ѫ��
        .Add mCol.ColӤ���Ա�, "_" & B_�Ա�
        .Add mCol.Col����ʱ��, "_" & B_����ʱ��
        '--ͨ���к��ҵ���Ƭ����
        .Add B_����, "C" & mCol.ColӤ������
        .Add B_�������, "C" & mCol.Col�������
        .Add B_����ʱ��, "C" & mCol.Col����ʱ��
        .Add B_��ע˵��, "C" & mCol.Col��ע˵��
        .Add B_��, "C" & mCol.Col��
        .Add B_����, "C" & mCol.Col����
        .Add B_���䷽ʽ, "C" & mCol.Col���䷽ʽ
        .Add B_̥��״��, "C" & mCol.Col̥��״��
        .Add B_Ѫ��, "C" & mCol.COlѪ��
        .Add B_�Ա�, "C" & mCol.ColӤ���Ա�
        .Add B_����ʱ��, "C" & mCol.Col����ʱ��
    End With
    
    cmdPrint.Visible = InStr(mstrPrivs, "Ӥ�������ӡ") > 0
    
    '������Ϣ����
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, 1132, mstrPrivs, mfrmParent.hWnd)
    Call AddMipModule(mclsMipModule)
    Set mclsXML = New zl9ComLib.clsXML
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Then
        If MsgBox("�����Ѿ����޸ģ�ȷʵҪ�������˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
    
    Set mcolBaby = Nothing
    
    'ж����Ϣ����
    If Not (mclsMipModule Is Nothing) Then
        Call mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    If Not (mclsXML Is Nothing) Then
        Set mclsXML = Nothing
    End If
End Sub

Private Sub txtBaby_Change(Index As Integer)
    vsBaby.TextMatrix(vsBaby.Row, mcolBaby("_" & Index)) = txtBaby(Index).Text
    If vsBaby.Cell(flexcpData, vsBaby.Row, mcolBaby("_" & Index)) <> vsBaby.TextMatrix(vsBaby.Row, mcolBaby("_" & Index)) Then mblnChange = True
End Sub

Private Sub txtBaby_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtBaby(Index)
End Sub

Private Sub txtBaby_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = Asc("'") Then
        KeyAscii = 0: Exit Sub
    End If
    
    If KeyAscii = vbKeyReturn Then
        '�س�13
        KeyAscii = 0
        Select Case Index
        
        Case B_����
            If Trim(txtBaby(Index).Text) = "" Then
                txtBaby(Index).Text = Trim(lblOut(IX_����).Caption & "֮Ӥ" & vsBaby.TextMatrix(vsBaby.Row, mCol.col���))
            End If
        Case B_�������
            If Trim(txtBaby(Index).Text) = "" Then
                txtBaby(Index).Text = 1
            End If
        Case B_����ʱ��
            If Trim(txtBaby(Index).Text) = "" Then
                txtBaby(Index).Text = Format(zlDatabase.Currentdate, "YYYY-MM-dd HH:mm")
            ElseIf Trim(txtBaby(Index).Text) <> "" Then
                txtBaby(Index).Text = GetFullDate(txtBaby(Index).Text)
            End If
        Case B_����ʱ��
            If Trim(txtBaby(Index).Text) <> "" Then
                txtBaby(Index).Text = GetFullDate(txtBaby(Index).Text)
            End If
        Case B_��, B_����
            txtBaby(Index).Text = gclsBase.FormatEx(txtBaby(Index).Text, 2) '��ౣ����λ
        End Select
        Call ChangeBabyInfo(vsBaby.Row, mcolBaby("_" & Index), txtBaby(Index))
        If Index = B_��ע˵�� Then
            If vsBaby.Row = vsBaby.Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                vsBaby.Row = vsBaby.Row + 1
                txtBaby(B_����).SetFocus: Exit Sub
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
        Exit Sub
    End If
    
    If KeyAscii = vbKeyBack Then
        vsBaby.TextMatrix(vsBaby.Row, mcolBaby("_" & Index)) = ""
    End If
    
    Select Case Index
    Case B_����ʱ��, B_����ʱ��
        If InStr("/-0123456789:" & Chr(32) & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    Case B_�������
        If InStr("0123456789" & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    Case B_��, B_����
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End Select
End Sub

Private Sub txtBaby_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTmp As String
    
    strTmp = txtBaby(Index).Text
    If Index = B_���� Then
        If LenB(StrConv(strTmp, vbFromUnicode)) > 16 Then
            zlCommFun.ShowTipInfo txtBaby(Index).hWnd, strTmp
        Else
            zlCommFun.ShowTipInfo txtBaby(Index).hWnd, ""
        End If
        
    End If

    If Index = B_����ʱ�� Or Index = B_���� Then
        If strTmp = "" Then
            txtBaby(Index).ToolTipText = "�س�����ȱʡֵ"
        Else
            txtBaby(Index).ToolTipText = ""
        End If
    End If
End Sub

Private Sub txtBaby_Validate(Index As Integer, Cancel As Boolean)
    Dim re As New RegExp
    Dim strMsg As String
    
    Select Case Index
    Case B_����
        If zlCommFun.ActualLen(txtBaby(Index).Text) > 100 Then
            strMsg = "��Ӥ�����������ֻ��������50�����ֻ�100���ַ���"
        End If
    Case B_����ʱ��
        re.Pattern = "^([1-9][0-9]{3})-((01|03|05|07|08|10|12)-(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)-(0[1-9]|[1-2][0-9]|30)|02-(0[1-9]|[1-2][0-9]))\s([0-1][0-9]|2[0-3]):([0-5][0-9])$"
        If Not re.Test(Trim(txtBaby(Index).Text)) Then
            strMsg = "������ʱ�䡿������Ч�����ڸ�ʽ[YYYY-MM-dd hh:mm]��"
            txtBaby(Index).SetFocus
        Else
            If CDate(Format(Trim(txtBaby(Index).Text), "YYYY-MM-dd HH:mm")) > CDate(Format(zlDatabase.Currentdate, "YYYY-MM-dd HH:mm")) Then
                strMsg = "������ʱ�䡿���ڵ�ǰϵͳʱ�䣡"
            End If
        End If
    Case B_����ʱ��
        If Trim(txtBaby(Index).Text) <> "" Then
            re.Pattern = "^([1-9][0-9]{3})-((01|03|05|07|08|10|12)-(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)-(0[1-9]|[1-2][0-9]|30)|02-(0[1-9]|[1-2][0-9]))\s([0-1][0-9]|2[0-3]):([0-5][0-9])$"
            If Not re.Test(Trim(txtBaby(Index).Text)) Then
                strMsg = "������ʱ�䡿������Ч�����ڸ�ʽ[YYYY-MM-dd hh:mm]��"
                txtBaby(Index).SetFocus
            End If
        End If
    Case B_��, B_����
        If txtBaby(Index).Text = "" Then
            Exit Sub  '������ǰû��¼����\���ص����
        ElseIf Not IsNumeric(txtBaby(Index).Text) Then
            strMsg = IIf(Index = B_��, "������", "�����ء�") & "������Ч���������ͣ�"
        ElseIf Len(txtBaby(Index).Text) > 10 Then
            strMsg = IIf(Index = B_��, "������", "�����ء�") & "���ֻ����¼��10���ַ���"
        End If
    Case B_��ע˵��
        If zlCommFun.ActualLen(txtBaby(Index).Text) > 100 Then
            strMsg = "����ע˵�������ֻ��������50�����ֻ�100���ַ���"
        End If
    End Select
    
    If strMsg <> "" Then
        ShowErrInfo "��ʾ:" & strMsg
        zlControl.TxtSelAll txtBaby(Index)
        Cancel = True: Exit Sub
    Else
        ShowErrInfo strMsg
    End If
End Sub

Private Sub vsBaby_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngTmp As Long
    
    With vsBaby
        If .Rows > 1 Then
            If NewRow <> OldRow And NewRow > 0 Then
                Call SetBabyInfo(NewRow)
            End If
        Else
            Call SetBabyInfo(0, 2)
        End If
        Call SetCardEnable(.Rows > 1)
    End With
End Sub

Private Sub vsBaby_Click()
    With vsBaby
        If .Rows > 1 Then
            
        End If
    End With
End Sub

Private Sub vsBaby_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Call cmdDel_Click
    End If
End Sub

Private Function AddNewBabyRow() As Boolean
'���ܣ�����һ�����У�������ȱʡֵ
'���أ���������������У�����ʾ������False
    Dim lngRow As Long
    Dim strMsg As String
    Dim i As Long
    
    
    With vsBaby
        If .Rows - 1 >= MAX_BABY Then
            MsgBox "���˵�Ӥ����̫�࣬�������������ӡ�", vbInformation, gstrSysName
            Exit Function
        End If
        For i = 1 To .Rows - 1
            If .TextMatrix(i, mCol.ColӤ������) = "" Then
                .Row = i
                txtBaby(B_����).SetFocus
                Exit Function
            End If
        Next
        
        .AddItem "", .Rows
        .Row = .Rows - 1: Call SetBabyInfo(.Row, 1)
        .TextMatrix(.Row, mCol.col���) = .Rows - 1
        .ShowCell .Row, mCol.col���
        txtBaby(B_����).SetFocus
        mblnChange = True
    End With
    
    AddNewBabyRow = True
End Function

Private Function ReadDict(strDict As String, cbo As ComboBox) As Boolean
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo errH
    strSQL = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From " & strDict & " Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    cbo.Clear
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 Then
                cbo.ListIndex = cbo.NewIndex
                cbo.ItemData(cbo.NewIndex) = 1
            End If
            rsTmp.MoveNext
        Next
    End If
    ReadDict = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ChangeBabyInfo(ByVal lngRow As Long, ByVal lngCol As Long, ByVal objControl As Object)
    Dim strTmp As String

    If lngRow < vsBaby.FixedRows Then Exit Sub
    If TypeName(objControl) = "ComboBox" Then
        strTmp = zlCommFun.GetNeedName(Trim(objControl.Text))
    Else
        strTmp = Trim(objControl.Text)
    End If
    vsBaby.TextMatrix(lngRow, lngCol) = strTmp
End Sub

Private Sub SetBabyInfo(ByVal lngRow As Long, Optional ByVal bytFunc As Byte = 0)
'����:�����������ʾ����Ƭ
'����:bytFunc =0,ѡ����������ʾ����Ƭ��=1��������ȱʡֵ=2�����Ƭ��Ϣ
    With vsBaby
        If lngRow = 0 Then Exit Sub
        If bytFunc = 0 Then
            On Error Resume Next
            '������ǰû��Ѫ�͵ĸ�ֵ�ᱨ��cboBaby(B_Ѫ��).Text��ֵʱ,�Ҳ���ֵ
            txtBaby(B_����).Text = .TextMatrix(lngRow, mCol.ColӤ������)   'ȱʡ����
            txtBaby(B_��).Text = .TextMatrix(lngRow, mCol.Col��)
            txtBaby(B_����).Text = .TextMatrix(lngRow, mCol.Col����)
            txtBaby(B_�������).Text = .TextMatrix(lngRow, mCol.Col�������)
            txtBaby(B_����ʱ��).Text = .TextMatrix(lngRow, mCol.Col����ʱ��)
            txtBaby(B_����ʱ��).Text = .TextMatrix(lngRow, mCol.Col����ʱ��)
            cboBaby(B_̥��״��).Text = cbo.Locate(cboBaby(B_̥��״��), .TextMatrix(lngRow, mCol.Col̥��״��), False)
            cboBaby(B_Ѫ��).Text = cbo.Locate(cboBaby(B_Ѫ��), .TextMatrix(lngRow, mCol.COlѪ��), False)
            cboBaby(B_���䷽ʽ).Text = cbo.Locate(cboBaby(B_���䷽ʽ), .TextMatrix(lngRow, mCol.Col���䷽ʽ), False)
            cboBaby(B_�Ա�).Text = cbo.Locate(cboBaby(B_�Ա�), .TextMatrix(lngRow, mCol.ColӤ���Ա�), False)
            
            txtBaby(B_��ע˵��).Text = .TextMatrix(lngRow, mCol.Col��ע˵��)
            
            Err.Clear: On Error GoTo 0
        ElseIf bytFunc = 1 Then
            .TextMatrix(lngRow, mCol.ColӤ������) = txtBaby(B_����).Text
            txtBaby(B_�������).Text = 1
            .TextMatrix(lngRow, mCol.Col�������) = txtBaby(B_�������).Text
            txtBaby(B_����ʱ��).Text = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm")
            .TextMatrix(lngRow, mCol.Col����ʱ��) = txtBaby(B_����ʱ��).Text
            .TextMatrix(lngRow, mCol.Col����ʱ��) = txtBaby(B_����ʱ��).Text
            .TextMatrix(lngRow, mCol.Col��) = txtBaby(B_��).Text
            .TextMatrix(lngRow, mCol.Col����) = txtBaby(B_����).Text
             
            .TextMatrix(lngRow, mCol.ColӤ���Ա�) = zlCommFun.GetNeedName(cboBaby(B_�Ա�).Text)
            .TextMatrix(lngRow, mCol.Col���䷽ʽ) = zlCommFun.GetNeedName(cboBaby(B_���䷽ʽ).Text)
            .TextMatrix(lngRow, mCol.Col̥��״��) = zlCommFun.GetNeedName(cboBaby(B_̥��״��).Text)
            .TextMatrix(lngRow, mCol.COlѪ��) = zlCommFun.GetNeedName(cboBaby(B_Ѫ��).Text)
            
        ElseIf bytFunc = 2 Then
        '�����Ƭ��Ϣ
            txtBaby(B_����).Text = ""
            txtBaby(B_�������).Text = ""
            txtBaby(B_��).Text = ""
            txtBaby(B_����).Text = ""
            txtBaby(B_��ע˵��).Text = ""
            txtBaby(B_����ʱ��).Text = ""
            txtBaby(B_����ʱ��).Text = ""
        End If
    End With
End Sub
Private Function CheckBaby() As Boolean
'����:����ǰ���
    Dim i As Long, k As Long
    Dim strErr As String
    Dim j As Long
    Dim strName As String
    
    '���¼����Ϣ����Ϊ��
    With vsBaby
        strErr = ""
        For i = .FixedRows To .Rows - 1
            For j = mCol.ColӤ������ To mCol.COlѪ��
                '������ǰû��������ʱ,��Ҫ�����¼��
                If j = mCol.Col�� Or j = mCol.Col���� Then Exit For
                
                If .TextMatrix(i, j) = "" Then
                    strErr = "��š�" & .TextMatrix(i, mCol.col���) & "����" & .TextMatrix(0, j) & "Ϊ�գ�"
                    Exit For
                End If
            Next
            If strErr <> "" Then Exit For
            
            For k = 1 To .Rows - 1
                If .TextMatrix(i, mCol.ColӤ������) = .TextMatrix(k, mCol.ColӤ������) And k <> i Then
                    strErr = "��Ӥ������" & .TextMatrix(i, mCol.ColӤ������) & "���ظ���ӡ�": j = mCol.ColӤ������
                    Exit For
                End If
                
                If .TextMatrix(i, mCol.Col�������) <> .TextMatrix(k, mCol.Col�������) Then
                    strErr = "Ӥ����" & .TextMatrix(i, mCol.ColӤ������) & "�����������Ӥ����" & .TextMatrix(k, mCol.ColӤ������) & "�����������һ�£�": j = mCol.Col�������
                    Exit For
                End If
            Next
            If k <= .Rows - 1 Then Exit For
            
        Next
        
        If i <= .Rows - 1 Then
            If strErr <> "" Then
                MsgBox strErr, vbInformation + vbOKOnly, Me.Caption
                 .Row = i: .ShowCell .Row, j
                If j = mCol.Col����ʱ�� Or j = mCol.ColӤ������ Or j = mCol.Col������� Or j = mCol.Col�� Or j = mCol.Col���� Or j = mCol.Col��ע˵�� Then
                    txtBaby(mcolBaby("C" & j)).SetFocus
                Else
                    cboBaby(mcolBaby("C" & j)).SetFocus
                End If
             
                Exit Function
            End If
        End If
    End With
    
    CheckBaby = True
End Function

Private Function GetFullDate(ByVal strText As String, Optional blnTime As Boolean = True) As String
'���ܣ�������������ڼ�,�������������ڴ�(yyyy-MM-dd[ HH:mm])
'������blnTime=�Ƿ���ʱ�䲿��
    Dim Curdate As Date, strTmp As String
    
    If strText = "" Then Exit Function
    Curdate = zlDatabase.Currentdate
    strTmp = strText
    
    If InStr(strTmp, "-") > 0 Or InStr(strTmp, "/") Or InStr(strTmp, ":") > 0 Then
        '���봮�а������ڷָ���
        If IsDate(strTmp) Then
            strTmp = Format(strTmp, "yyyy-MM-dd HH:mm")
            If Right(strTmp, 5) = "00:00" And InStr(strText, ":") = 0 Then
                'ֻ���������ڲ���
                strTmp = Mid(strTmp, 1, 11) & Format(Curdate, "HH:mm")
            ElseIf Left(strTmp, 10) = "1899-12-30" Then
                'ֻ������ʱ�䲿��
                strTmp = Format(Curdate, "yyyy-MM-dd") & Right(strTmp, 6)
            End If
        Else
            '����Ƿ�����,����ԭ����
            strTmp = strText
        End If
    Else
        '���������ڷָ���
        If Len(strTmp) <= 2 Then
            '��������dd
            strTmp = Format(strTmp, "00")
            strTmp = Format(Curdate, "yyyy-MM") & "-" & strTmp & " " & Format(Curdate, "HH:mm")
        ElseIf Len(strTmp) <= 4 Then
            '��������MMdd
            strTmp = Format(strTmp, "0000")
            strTmp = Format(Curdate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Right(strTmp, 2) & " " & Format(Curdate, "HH:mm")
        ElseIf Len(strTmp) <= 6 Then
            '��������yyMMdd
            strTmp = Format(strTmp, "000000")
            strTmp = Format(Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & "-" & Right(strTmp, 2), "yyyy-MM-dd") & " " & Format(Curdate, "HH:mm")
        ElseIf Len(strTmp) <= 8 Then
            '��������MMddHHmm
            strTmp = Format(strTmp, "00000000")
            strTmp = Format(Curdate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & " " & Mid(strTmp, 5, 2) & ":" & Right(strTmp, 2)
            If Not IsDate(strTmp) Then
                '��������yyyyMMdd
                strTmp = Format(strText, "00000000")
                strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Right(strTmp, 2) & " " & Format(Curdate, "HH:mm")
            End If
        Else
            '��������yyyyMMddHHmm
            strTmp = Format(strTmp, "000000000000")
            strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & Mid(strTmp, 9, 2) & ":" & Right(strTmp, 2)
        End If
    End If
    
    If IsDate(strTmp) And Not blnTime Then
        strTmp = Format(strTmp, "yyyy-MM-dd")
    End If
    GetFullDate = strTmp
End Function

Private Sub SetCardEnable(ByVal blnEnable As Boolean)

    If fraBabyInput.Enabled = blnEnable Then Exit Sub
    cmdDel.Enabled = blnEnable
    cmdPrint.Enabled = blnEnable
    fraBabyInput.Enabled = blnEnable

    If blnEnable Then
        txtBaby(B_����).BackColor = M_CON_ColorEnabled
        txtBaby(B_�������).BackColor = M_CON_ColorEnabled
        txtBaby(B_����ʱ��).BackColor = M_CON_ColorEnabled
        txtBaby(B_����ʱ��).BackColor = M_CON_ColorEnabled
        txtBaby(B_��).BackColor = M_CON_ColorEnabled
        txtBaby(B_����).BackColor = M_CON_ColorEnabled
        txtBaby(B_��ע˵��).BackColor = M_CON_ColorEnabled
        
        cboBaby(B_�Ա�).BackColor = M_CON_ColorEnabled
        cboBaby(B_Ѫ��).BackColor = M_CON_ColorEnabled
        cboBaby(B_���䷽ʽ).BackColor = M_CON_ColorEnabled
        cboBaby(B_̥��״��).BackColor = M_CON_ColorEnabled
        
    Else
        txtBaby(B_����).BackColor = M_CON_ColorUnEnabled
        txtBaby(B_�������).BackColor = M_CON_ColorUnEnabled
        txtBaby(B_����ʱ��).BackColor = M_CON_ColorUnEnabled
        txtBaby(B_����ʱ��).BackColor = M_CON_ColorUnEnabled
        txtBaby(B_��).BackColor = M_CON_ColorUnEnabled
        txtBaby(B_����).BackColor = M_CON_ColorUnEnabled
        txtBaby(B_��ע˵��).BackColor = M_CON_ColorUnEnabled
        
        cboBaby(B_�Ա�).BackColor = M_CON_ColorUnEnabled
        cboBaby(B_Ѫ��).BackColor = M_CON_ColorUnEnabled
        cboBaby(B_���䷽ʽ).BackColor = M_CON_ColorUnEnabled
        cboBaby(B_̥��״��).BackColor = M_CON_ColorUnEnabled
    End If
End Sub

Private Sub ShowErrInfo(ByVal strMsg As String)
    If strMsg = "" Then
        lblERRInfo.Visible = False
    Else
        lblERRInfo.Visible = True
        lblERRInfo.Caption = strMsg
    End If
End Sub
