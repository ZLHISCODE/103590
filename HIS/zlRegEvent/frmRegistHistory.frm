VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmRegistHistory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������ιҺ���Ϣ��ѯ"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10200
   Icon            =   "frmRegistHistory.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin zlIDKind.IDKindNew IDKind 
      Height          =   375
      Left            =   600
      TabIndex        =   13
      Top             =   120
      Width           =   680
      _ExtentX        =   1191
      _ExtentY        =   661
      Appearance      =   2
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
      BackColor       =   -2147483633
   End
   Begin VB.Frame Frame1 
      Height          =   90
      Left            =   0
      TabIndex        =   4
      Top             =   975
      Width           =   10785
   End
   Begin VB.TextBox txtPatient 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1275
      TabIndex        =   3
      Top             =   120
      Width           =   3090
   End
   Begin VB.Frame Frame2 
      Height          =   90
      Left            =   0
      TabIndex        =   2
      Top             =   5760
      Width           =   10785
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   8550
      TabIndex        =   1
      Top             =   5970
      Width           =   1230
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7260
      TabIndex        =   0
      Top             =   5970
      Width           =   1230
   End
   Begin VSFlex8Ctl.VSFlexGrid vsRegist 
      Height          =   4470
      Left            =   90
      TabIndex        =   5
      Top             =   1170
      Width           =   9975
      _cx             =   17595
      _cy             =   7885
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmRegistHistory.frx":0442
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
      ExplorerBar     =   7
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   12
      Top             =   195
      Width           =   480
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "ҽ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   11
      Top             =   645
      Width           =   480
   End
   Begin VB.Label txt���� 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1275
      TabIndex        =   10
      Top             =   570
      Width           =   3135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�Ա�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5970
      TabIndex        =   9
      Top             =   645
      Width           =   480
   End
   Begin VB.Label txt�Ա� 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6510
      TabIndex        =   8
      Top             =   585
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7965
      TabIndex        =   7
      Top             =   645
      Width           =   480
   End
   Begin VB.Label txt���� 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8520
      TabIndex        =   6
      Top             =   585
      Width           =   1275
   End
End
Attribute VB_Name = "frmRegistHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mrsInfo As ADODB.Recordset
Private mblnOlnyBJYB As Boolean
Private mstr�ű� As String
Private mlng����ID As Long
Private mstrPrivs As String, mintIDKind As Integer
Private mblnOk As Boolean
Private mbln����סԺ���˹Һ� As Boolean
Private Const mlngModule = 1111
Private mblnNotClick As Boolean
'-----------------------------------------------------------------------------------
'���㿨���
Private mstrPassWord As String
'-----------------------------------------------------------------------------------

Public Function ShowRegist(ByVal frmMain As Form, ByVal strPrivs As String, _
     ByVal bln����סԺ���˹Һ� As Boolean, blnOlnyBjYb As Boolean, _
    ByRef lng����ID As Long, ByRef str�ű� As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�ѡ���˵����ιҺ���Ϣ
    '��Σ�blnOlnyBjYb- �Ƿ񱱾�ҽ��
    '���Σ�str�ű�-������ѡ��ĺű�
    '         lng����ID-���ص�ѡ��Ĳ���ID
    '���أ��ɹ�����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-07-16 14:34:24
    '˵����28604
    '------------------------------------------------------------------------------------------------------------------------
    mblnOlnyBJYB = blnOlnyBjYb: mlng����ID = lng����ID: mstrPrivs = strPrivs: mblnOk = False
    mbln����סԺ���˹Һ� = bln����סԺ���˹Һ�
    str�ű� = ""
    Me.Show 1, frmMain
    str�ű� = mstr�ű�: lng����ID = mlng����ID
    ShowRegist = mblnOk
End Function

Private Sub cmdCancel_Click()
    mblnOk = False: Unload Me
End Sub

Private Sub cmdOK_Click()
      With vsRegist
            If .Row < 0 Then Exit Sub
            If Trim(.TextMatrix(.Row, .ColIndex("�ű�"))) = "" Then Exit Sub
            If Val(.RowData(.Row)) = 0 Then Exit Sub
            
            If Val(.RowData(.Row)) <> mlng����ID And mlng����ID <> 0 Then
                If MsgBox("ע��:" & vbCrLf & " ����Ϊ�� " & txtPatient.Text & "���Ĳ��˲��ǹҺ�ȷ���Ĳ���,�Ƿ����?" & vbCrLf & _
                "ѡ���ǡ�:��ʾ�Ե�ǰѡ��Ĳ�����Ϊ�ҺŲ��ˡ�" & vbCrLf & _
                "ѡ�񡺷�:��ʾ���Դ˲���Ϊ׼�����ز�ѯ���档", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
            End If
            mlng����ID = Val(.RowData(.Row))
            mstr�ű� = Trim(.TextMatrix(.Row, .ColIndex("�ű�")))
            
            mblnOk = True
            Unload Me:
      End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnCancel As Boolean
    Select Case KeyCode
        Case vbKeyF4
            If Shift = vbCtrlMask Then
                If IDKind.Enabled Then IDKind.IDKind = IDKind.GetKindIndex("IC����"): Call IDKind_Click(IDKind.GetCurCard)
            ElseIf Me.ActiveControl Is txtPatient Then
                If IDKind.Enabled Then
                    If Shift = vbShiftMask Then
                        IDKind.IDKind = IIf(IDKind.IDKind = 0, UBound(Split(IDKind.IDkindStr, ";")), IDKind.IDKind - 1)
                    Else
                        IDKind.IDKind = IIf(IDKind.IDKind = UBound(Split(IDKind.IDkindStr, ";")), 0, IDKind.IDKind + 1)
                    End If
                End If
            End If
        Case vbKeyF11
            If txtPatient.Enabled And txtPatient.Visible And Not txtPatient.Locked Then
                If Me.ActiveControl Is txtPatient Then
                    IDKind.IDKind = IIf(IDKind.IDKind = UBound(Split(IDKind.IDkindStr, ";")), IDKind.GetKindIndex("����"), IDKind.IDKind + 1)
                Else
                    txtPatient.SetFocus
                End If
            End If
        Case vbKeyReturn
       
    End Select
End Sub
Private Sub Form_Load()
    Dim strTemp As String
    
    Call InitIDKind
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    
    Call mobjIDCard.SetParent(Me.Hwnd)
    Call mobjICCard.SetParent(Me.Hwnd)
    Set mobjICCard.gcnOracle = gcnOracle
    Call GetRegInFor(g˽��ģ��, Me.Name, "idkind", strTemp)
    mintIDKind = Val(strTemp)
    If mintIDKind > 0 And mintIDKind <= IDKind.ListCount Then IDKind.IDKind = mintIDKind
    If mlng����ID <> 0 Then
        txtPatient.Text = "-" & mlng����ID
        Call GetPatient(Trim(txtPatient.Text))
    End If
    Call InitVsGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
    mintIDKind = IDKind.IDKind
    Call SaveRegInFor(g˽��ģ��, Me.Name, "idkind", mintIDKind)
End Sub
 
Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    If IsCardType(IDKind, "IC����") Then
        If Not mobjICCard Is Nothing Then
            txtPatient.Text = mobjICCard.Read_Card()
            If txtPatient.Text <> "" Then
                Call GetPatient(Trim(txtPatient))
            End If
        End If
        Exit Sub
    End If
    lng�����ID = IDKind.GetCurCard.�ӿ����
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
        Call GetPatient(Trim(txtPatient))
    End If
End Sub

Private Sub IDKind_ItemClick(index As Integer, objCard As zlIDKind.Card)
    Set gobjSquare.objCurCard = objCard
    txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    
'    If IDKind.GetCardNoLen <> 0 Then
'        txtPatient.MaxLength = IDKind.GetCardNoLen
'    Else
'        txtPatient.MaxLength = zlGetPatiInforMaxLen.intPatiName
'    End If
    
    '��Ҫ�����Ϣ,����ˢ����,���л�,���������ʾʧȥ����
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    Dim lngPreIDKind As Long, intIndex As Integer
    Dim dtDate As Date
    Dim blnNew As Boolean

    If txtPatient.Locked Or txtPatient.Text <> "" Then Exit Sub 'Or Not Me.ActiveControl Is txtPatient
    mblnNotClick = True

    intIndex = IDKind.GetKindIndex(objCard.����)
    lngPreIDKind = IDKind.IDKind
    If intIndex > 0 Then IDKind.IDKind = intIndex
    txtPatient.Text = objPatiInfor.����
    Call txtPatient_KeyPress(vbKeyReturn)
    IDKind.IDKind = lngPreIDKind
    mblnNotClick = False
End Sub

Private Sub txtPatient_Change()
    txtPatient.Tag = "": txtPatient.ForeColor = Me.ForeColor
    If Me.ActiveControl Is txtPatient Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard txtPatient.Text = ""
    End If
End Sub

Private Sub txtPatient_GotFocus()
     zlControl.TxtSelAll txtPatient
      If txtPatient.Text = "" And Not txtPatient.Locked Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(True)
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(True)
    End If
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    Dim blnCard As Boolean, blnICCard As Boolean
    
    On Error GoTo errH
    If txtPatient.Locked Then Exit Sub
    
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
    If IDKind.GetCurCard.���� Like "����*" Then
        blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, glngSys, IDKind.ShowPassText)
    ElseIf IDKind.GetCurCard.���� = "�����" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not (IsNumeric(Chr(KeyAscii)) Or Chr(KeyAscii) = "-") Then KeyAscii = 0: Exit Sub
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
        ElseIf IsNumeric(txtPatient.Tag) Then
            KeyAscii = 0
            'ˢ�²�����Ϣ:"-����ID"
            Call GetPatient(txtPatient.Tag, False)
            Exit Sub
        End If
        KeyAscii = 0
        If IDKind.IDKind = IDKind.GetKindIndex("IC����") Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
        Call GetPatient(txtPatient.Text, blnCard)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub GetPatient(ByVal strInput As String, Optional blnCard As Boolean)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡ������Ϣ
    '��Σ�blnCard=�Ƿ���￨ˢ��
    '���Σ�
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-07-16 14:24:14
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strTemp As String
    Dim blnSame As Boolean, blnCancel As Boolean
    Dim cur��� As Currency, curMoney As Currency
    Dim i As Integer, strPati As String
    Dim vRect As RECT, str����Ժ As String
    Dim strSQL As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    
    On Error GoTo errH
    If mbln����סԺ���˹Һ� = False Then
        str����Ժ = " And Not Exists(Select 1 From ������ҳ Where ����ID=B.����ID And ��ҳID=B.��ҳID And Nvl(��������,0)=0 And ��Ժ���� is Null)"
    End If
    
    strSQL = ""
    If (blnCard Or IDKind.IDKind = IDKindDefaultKind) And Not (Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2))) Then
        lng�����ID = IDKind.GetDefaultCardTypeID
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
        'If lng����ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng����ID
        blnHavePassWord = True
        strSQL = strSQL & " And A.����ID=[1] " & str����Ժ
        strInput = UCase(strInput)
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then
        '�����
        strSQL = strSQL & " And B.�����=[2]" & str����Ժ
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then
        '����ID
        strSQL = strSQL & " And B.����ID=[2]" & str����Ժ
    Else
        Select Case IDKind.GetCurCard.����
            Case "����", "��������￨"
                '����
                blnSame = False
                If Not mrsInfo Is Nothing Then
                    If txtPatient.Text = mrsInfo!���� Then blnSame = True
                End If
                If Not blnSame Then
                    If Not gblnSeekName Or gblnSeekName And Len(txtPatient.Text) < 2 Then
                        Set mrsInfo = Nothing: Exit Sub
                    Else
                        strPati = _
                            " Select 1 as ����ID,B.����ID as ID,B.����ID,B.����,B.�Ա�,B.����,B.�����,B.��������,B.���֤��,B.��ͥ��ַ,B.������λ" & _
                            " From ������Ϣ B" & _
                            " Where Rownum <101 And B.ͣ��ʱ�� is NULL And B.���� Like [1]" & str����Ժ & _
                            IIf(gintNameDays = 0, "", " And Nvl(B.����ʱ��,B.�Ǽ�ʱ��)>Trunc(Sysdate-[2])")
                     
                        strPati = strPati & " Order by ����ID,����"
                            
                        vRect = zlControl.GetControlRect(txtPatient.Hwnd)
                        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%", gintNameDays)
                        If Not rsTmp Is Nothing Then
                            If rsTmp!ID = 0 Then '�����²���
                                MsgBox "δ�ҵ����������Ĳ�����Ϣ,����!", vbInformation + vbDefaultButton1, gstrSysName
                                 txtPatient.Text = ""
                                Call txtPatient_GotFocus
                                Set mrsInfo = Nothing: Exit Sub
                            Else '�Բ���ID��ȡ
                                strInput = rsTmp!����ID
                                strSQL = strSQL & " And A.����ID=[1]"
                            End If
                        Else 'ȡ��ѡ��
                           MsgBox "δ�ҵ����������Ĳ�����Ϣ,����!", vbInformation + vbDefaultButton1, gstrSysName
                            txtPatient.Text = ""
                            Call txtPatient_GotFocus
                            Set mrsInfo = Nothing: Exit Sub
                        End If
                    End If
                Else
                    '�޸����⣺39164
                    strInput = mrsInfo!����ID
                    strSQL = strSQL & " And A.����ID=[1]"
                End If
            Case "ҽ����"
                strInput = UCase(strInput)
                If mblnOlnyBJYB And zlCommFun.ActualLen(strInput) >= 9 Then
                    '������ҽ������Ч:������:����:26982
                    strSQL = strSQL & " And B.ҽ���� like [3] " & str����Ժ
                    strTemp = Left(strInput, 9) & "%"
                Else
                    strSQL = strSQL & " And B.ҽ����=[1]" & str����Ժ
                End If
            Case "���֤��", "�������֤", "���֤"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strInput = "-" & lng����ID
                strSQL = strSQL & " And A.����ID=[2] " & str����Ժ
            Case "IC����", "IC��"
                strInput = UCase(strInput)
                 If gobjSquare.objSquareCard.zlGetPatiID("IC��", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strInput = "-" & lng����ID
                strSQL = strSQL & " And A.����ID=[2] " & str����Ժ
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And B.�����=[1]" & str����Ժ
            Case "�Һŵ���"
                 strInput = GetFullNO(strInput, 12)
                 txtPatient.Text = strInput
                strSQL = strSQL & " And A.NO=[1]" & str����Ժ

         Case Else
            '��������,��ȡ��صĲ���ID
            If IDKind.GetCurCard.�ӿ���� > 0 Then
                lng�����ID = IDKind.GetCurCard.�ӿ����
                If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                If lng����ID = 0 Then lng����ID = 0
            Else
                If gobjSquare.objSquareCard.zlGetPatiID(IDKind.GetCurCard.����, strInput, False, lng����ID, _
                    strPassWord, strErrMsg) = False Then lng����ID = 0
            End If
            If lng����ID <= 0 Then lng����ID = 0
            strSQL = strSQL & " And A.����ID=[1]" & str����Ժ
            strInput = "-" & lng����ID
            blnHavePassWord = True
        End Select
    End If
    
    strSQL = "" & _
            "   Select distinct A.NO,A.�ű�,A.ִ�в���id,C.���� as  �Һſ���, B.����ID," & _
            "            to_char(A.�Ǽ�ʱ��,'yyyy-mm-dd hh24:mi:ss') as �Һ�ʱ��  " & vbNewLine & _
            "   From ���˹Һż�¼ A, ������Ϣ B,���ű� C" & vbNewLine & _
            "   Where  A.ִ�в���ID=C.ID (+) " & _
            "               And B.����id =A.����id(+) and a.��¼����=1 and�� a.��¼״̬=1  " & strSQL & _
            "    Order by �Һ�ʱ�� Desc"
                                                             
    'û�����ú�����,������ǰ�Ĵ���ʽ,����ֻ��ȡ�����ԤԼ��(���ʧ��Լ��,���Ժ�ɫ������ʾ)
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, CStr(Mid(strInput, 2)))
    If rsTmp.RecordCount = 0 Then
        vsRegist.Clear 1: vsRegist.Rows = 2: vsRegist.Row = 1
        MsgBox "δ�ҵ����������Ĳ�����Ϣ,����!", vbInformation + vbDefaultButton1, gstrSysName
        txtPatient.Text = ""
        txtPatient.PasswordChar = ""
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtPatient.IMEMode = 0
        Call txtPatient_GotFocus
        Exit Sub
    End If
    
    If Val(Nvl(rsTmp!����ID)) <> 0 Then
        strSQL = "Select A.*,B.���� �������� From ������Ϣ A,������� B Where A.���� = B.���(+) And A.ͣ��ʱ�� is NULL "
        strSQL = strSQL & " And A.����id=[1]"
        Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(rsTmp!����ID)))
        If mrsInfo.EOF = False Then
            txtPatient.Text = Nvl(mrsInfo!����)
            txt����.Caption = Nvl(mrsInfo!��������):
            txt�Ա� = Nvl(mrsInfo!�Ա�)
            txt���� = Nvl(mrsInfo!����)
            txtPatient.PasswordChar = ""
            '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
            txtPatient.IMEMode = 0
            '74428�����ϴ���2014-7-8������������ʾ��ɫ����
            Call SetPatiColor(txtPatient, Nvl(mrsInfo!��������), IIf(Trim(txt����.Caption) = "", txtPatient.ForeColor, vbRed))
        Else
            txt����.Caption = "": txt�Ա� = "": txt���� = ""
        End If
        
    Else
        Set mrsInfo = Nothing
        txt����.Caption = "": txt�Ա� = "": txt���� = ""
    End If
    
    Dim lngRow As Long
    With vsRegist
        .Clear 1: .Rows = 2
        If rsTmp.RecordCount <> 0 Then .Rows = rsTmp.RecordCount + 1
        lngRow = 1
        Do While Not rsTmp.EOF
            .TextMatrix(lngRow, .ColIndex("��־")) = lngRow
            .TextMatrix(lngRow, .ColIndex("���ݺ�")) = Nvl(rsTmp!NO)
            .TextMatrix(lngRow, .ColIndex("�ű�")) = Nvl(rsTmp!�ű�)
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTmp!�Һſ���)
            .TextMatrix(lngRow, .ColIndex("�Һ�ʱ��")) = Nvl(rsTmp!�Һ�ʱ��)
            .RowData(lngRow) = Val(Nvl(rsTmp!����ID))
            lngRow = lngRow + 1
            rsTmp.MoveNext
        Loop
        zl_vsGrid_Para_Restore mlngModule, vsRegist, Me.Caption, "�Һŵ��б�", True
        .ColWidth(.ColIndex("��־")) = 285
    End With
    Call txtPatient_GotFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
  
Private Sub InitVsGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2009-09-09 15:45:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intType As Integer
    With vsRegist
        .ColData(.ColIndex("��־")) = "1|1"
        .ColData(.ColIndex("ԤԼ���ݺ�")) = "1|0"
    End With
End Sub

Private Sub vsRegist_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsRegist, Me.Caption, "�Һŵ��б�", True, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub vsRegist_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsRegist
        If Col = .ColIndex("��־") Then Cancel = True
    End With
End Sub

Private Sub vsRegist_DblClick()
        Call cmdOK_Click
End Sub

Private Sub vsRegist_GotFocus()
    vsRegist.BackColorSel = &H8000000D
End Sub

Private Sub vsRegist_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub vsRegist_LostFocus()
    vsRegist.BackColorSel = GRD_LOSTFOCUS_COLORSEL
End Sub
Private Sub vsRegist_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsRegist, Me.Caption, "�Һŵ��б�", True, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    Dim lngPreIDKind As Long
    If Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        lngPreIDKind = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("���֤��")
        txtPatient.Text = strID
        Call GetPatient(Trim(txtPatient.Text))
        IDKind.IDKind = lngPreIDKind
        If Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then mobjICCard.SetEnabled (txtPatient.Text = "")
    End If
End Sub


Private Sub mobjICCard_ShowICCardInfo(ByVal strNO As String)
    Dim lngPreIDKind As Long
    If Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        lngPreIDKind = IDKind.IDKind
        mblnNotClick = True
        IDKind.IDKind = IDKind.GetKindIndex("IC����")
        txtPatient.Text = strNO
        If txtPatient.Text <> "" Then
            Call GetPatient(Trim(txtPatient.Text))
        Else
            Call mobjICCard.SetEnabled(False) '��������Ϸ������������ü����Զ���ȡ
        End If
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
        If Me.ActiveControl Is txtPatient And Not txtPatient.Locked Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
    End If
End Sub

Private Function InitIDKind() As Boolean
    Dim objCard As Card
    Dim lngCardID As Long
    Call IDKind.zlInit(Me, glngSys, mlngModule, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", Me.txtPatient)
    lngCardID = Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, mlngModule, 0))
    If lngCardID <> 0 Then
        IDKind.DefaultCardType = lngCardID
    End If
    Set objCard = IDKind.GetfaultCard
    IDKind.ShowPropertySet = InStr(";" & mstrPrivs & ";", "��������") > 0
    If IDKind.Cards.��ȱʡ������ And Not objCard Is Nothing Then
        gobjSquare.blnȱʡ�������� = objCard.�������Ĺ��� <> ""
        gobjSquare.intȱʡ���ų��� = objCard.���ų���
        Set gobjSquare.objDefaultCard = objCard
       
    Else
        gobjSquare.blnȱʡ�������� = IDKind.Cards.������ʾ
        gobjSquare.intȱʡ���ų��� = 100
    End If
End Function

Private Function IsCardType(ByVal IDKindCtl As IDKindNew, ByVal strCardName As String) As Boolean
    If IDKindCtl Is Nothing Then Exit Function
    If UCase(TypeName(IDKindCtl)) <> "IDKINDNEW" Then Exit Function
    Select Case strCardName
     Case "����", "��������￨"
          IsCardType = IDKindCtl.GetCurCard.���� Like "����*"
     Case "���֤", "���֤��", "�������֤"
          IsCardType = IDKindCtl.GetCurCard.���� Like "*���֤*"
     Case "IC����", "IC��"
          IsCardType = IDKindCtl.GetCurCard.���� Like "IC��*"
     Case "ҽ����"
          IsCardType = IDKindCtl.GetCurCard.���� = "ҽ����"
     Case "�����"
          IsCardType = IDKindCtl.GetCurCard.���� = "�����"
     Case Else
            If IDKindCtl.GetCurCard Is Nothing Then Exit Function
            If Not IsNumeric(strCardName) Or Val(strCardName) <= 0 Then Exit Function
            If IDKindCtl.GetCurCard.�ӿ���� <= 0 Then Exit Function
            IsCardType = IDKindCtl.GetCurCard.�ӿ���� = Val(strCardName)
     End Select
End Function
Private Function IDKindDefaultKind() As Long
    Dim lngIndex As Long
    'IDkind��Ĭ��Kind
    If IDKind.DefaultCardType = "" Then
        lngIndex = -1
    Else
        If IsNumeric(IDKind.DefaultCardType) Then
           lngIndex = IDKind.GetKindIndex(IDKind.GetfaultCard.����)
        Else
           lngIndex = IDKind.GetKindIndex(IDKind.DefaultCardType)
        End If
    End If
    IDKindDefaultKind = lngIndex
End Function
