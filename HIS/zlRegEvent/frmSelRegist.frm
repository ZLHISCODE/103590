VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmSelRegist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ԤԼ�Һŵ���ȡ"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9945
   Icon            =   "frmSelRegist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   9945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin zlIDKind.IDKindNew IDKind 
      Height          =   375
      Left            =   700
      TabIndex        =   14
      Top             =   210
      Width           =   615
      _ExtentX        =   1085
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
      Left            =   7275
      TabIndex        =   11
      Top             =   5175
      Width           =   1230
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
      TabIndex        =   10
      Top             =   5175
      Width           =   1230
   End
   Begin VB.Frame Frame2 
      Height          =   90
      Left            =   0
      TabIndex        =   9
      Top             =   4965
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
      Left            =   1320
      TabIndex        =   2
      Top             =   210
      Width           =   3090
   End
   Begin VB.Frame Frame1 
      Height          =   90
      Left            =   0
      TabIndex        =   0
      Top             =   1065
      Width           =   10785
   End
   Begin VSFlex8Ctl.VSFlexGrid vsRegist 
      Height          =   3660
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   9720
      _cx             =   17145
      _cy             =   6456
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
      GridLines       =   2
      GridLinesFixed  =   9
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSelRegist.frx":0442
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
      Begin VB.PictureBox picImgPlan 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   45
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   13
         Top             =   75
         Width           =   210
         Begin VB.Image imgColPlan 
            Height          =   195
            Left            =   0
            Picture         =   "frmSelRegist.frx":0541
            ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
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
      TabIndex        =   8
      Top             =   675
      Width           =   1275
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
      Top             =   735
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
      TabIndex        =   6
      Top             =   675
      Width           =   1200
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
      TabIndex        =   5
      Top             =   735
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
      Left            =   705
      TabIndex        =   4
      Top             =   660
      Width           =   3705
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
      Left            =   210
      TabIndex        =   3
      Top             =   735
      Width           =   480
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
      Left            =   210
      TabIndex        =   1
      Top             =   285
      Width           =   480
   End
End
Attribute VB_Name = "frmSelRegist"
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
Private mintԤԼʧЧ���� As Integer
Private mstrPrivs As String, mintIDKind As Integer
Private mstrNo As String
Private mblnOk As Boolean
Private mbln����סԺ���˹Һ� As Boolean
Private mblnNotClick As Boolean
Private Const mlngModule = 1111
Private mlng����ID As Long
Private mbyt���� As Byte
Private mstr����IDs As String
'-----------------------------------------------------------------------------------
'���㿨���
Private mstrPassWord As String
'-----------------------------------------------------------------------------------
Public Function ShowRegist(ByVal frmMain As Form, ByVal strPrivs As String, _
     blnOlnyBjYb As Boolean, intԤԼʧЧ���� As Integer, _
    ByRef strOutNo As String, _
    ByRef rsOutPatiInfor As ADODB.Recordset, Optional ByVal lng����ID As Long, _
    Optional byt���� As Byte = 0, Optional str����IDs As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�ѡ�����ȡԤԼ��
    '��Σ�blnOlnyBjYb- �Ƿ񱱾�ҽ��
    '      lng����ID-����Ĳ���ID(����ʱ��ȱʡ�ò�����Ϣ,������Ҫ����ˢ�²���
    '      byt����:0-�ҺŴ���;1-����̨����
    '      str����IDs-���ƵĿ���ID
    '���Σ�strOutNo-���ص�ԤԼ���ݺ�
    '         rsOutPatiInfor-���صĲ�����Ϣ
    '         lng����ID
    '���أ��ɹ�����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-07-16 14:34:24
    '˵����31182
    '------------------------------------------------------------------------------------------------------------------------
    mblnOlnyBJYB = blnOlnyBjYb: mintԤԼʧЧ���� = intԤԼʧЧ����: mstrPrivs = strPrivs: mblnOk = False
    mbyt���� = byt����: mstr����IDs = str����IDs
    Set mrsInfo = rsOutPatiInfor: strOutNo = "": mlng����ID = lng����ID
    Me.Show 1, frmMain
    strOutNo = mstrNo: Set rsOutPatiInfor = mrsInfo
    ShowRegist = mblnOk
End Function

Private Sub cmdCancel_Click()

    mblnOk = False: Unload Me
End Sub

Private Sub cmdOK_Click()
      With vsRegist
            If .Row < 0 Then Exit Sub
            If Trim(.TextMatrix(.Row, .ColIndex("ԤԼ���ݺ�"))) = "" Then Exit Sub
            mstrNo = Trim(.TextMatrix(.Row, .ColIndex("ԤԼ���ݺ�")))
            mblnOk = True
            Unload Me
      End With
End Sub

Private Sub Form_Activate()
    
    '�ڴ��弤��ʱ,��������в�����Ϣ,���Դ˳�ʼ������,��ʼ����Ӧ��Ϣ
    If mlng����ID > 0 Then
        txtPatient.Text = "-" & mlng����ID
        Call txtPatient_KeyPress(13)    '�س���ȡ��Ϣ
    End If
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
    mblnOk = False
    
    Call NewCardObject '47007
    InitIDKind
    Set mobjICCard.gcnOracle = gcnOracle
    Call GetRegInFor(g˽��ģ��, Me.Name, "idkind", strTemp)
    mintIDKind = Val(strTemp)
    If mintIDKind > 0 And mintIDKind <= IDKind.ListCount Then IDKind.IDKind = mintIDKind
    Call InitVsGrid
    mbln����סԺ���˹Һ� = zlDatabase.GetPara("����סԺ���˹Һ�", glngSys, mlngModule, 0) = "1"
End Sub
Private Function InitIDKind() As Boolean
    Dim objCard As Card
    Dim lngCardID As Long
 
    Call IDKind.zlInit(Me, glngSys, mlngModule, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
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
Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    mintIDKind = IDKind.IDKind
    Call SaveRegInFor(g˽��ģ��, Me.Name, "idkind", mintIDKind)
    txtPatient.Enabled = False
    '47007
    CloseIDCard
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
    
    If mrsInfo Is Nothing Then
        blnNew = True
    ElseIf mrsInfo.State <> 1 Then
        blnNew = True
    End If
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
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtPatient_GotFocus()
    zlControl.TxtSelAll txtPatient
    If txtPatient.Text = "" And Not txtPatient.Locked Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(True)
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(True)
        IDKind.SetAutoReadCard True
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
    
    If IsCardType(IDKind, "����") Then
        blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, glngSys, IDKind.ShowPassText)
    ElseIf IsCardType(IDKind, "�����") Then
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
        If IsCardType(IDKind, "IC����") Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
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
    Dim strTmp As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    
    On Error GoTo errH
    If Not mbln����סԺ���˹Һ� And mbyt���� = 0 Then   '����̨������
        str����Ժ = " And Not Exists(Select 1 From ������ҳ Where ����ID(+)=B.����ID And ��ҳID(+)=B.��ҳID And Nvl(��������,0)=0 And ��Ժ���� is Null)"
    End If
    
    strSQL = ""
    
    If (blnCard Or IDKind.IDKind = IDKindDefaultKind) And Not (Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2))) Then
        lng�����ID = IDKind.GetDefaultCardTypeID
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
        If lng����ID <= 0 Then lng����ID = 0
        strInput = "-" & lng����ID
        blnHavePassWord = True
        strSQL = strSQL & " And A.����ID=[2] " & str����Ժ
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
                    If (Not gblnSeekName) Or (gblnSeekName And Len(txtPatient.Text) < 2) Then
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
                                Set mrsInfo = Nothing: Exit Sub
                            Else '�Բ���ID��ȡ
                                strInput = rsTmp!����ID
                                strSQL = strSQL & " And A.����ID=[1]"
                            End If
                        Else 'ȡ��ѡ��
                            txtPatient.Text = ""
                            txtPatient.PasswordChar = ""
                            '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
                            txtPatient.IMEMode = 0
                            Set mrsInfo = Nothing: Exit Sub
                        End If
                    End If
                Else
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
            Case IDKind.GetKindIndex("ԤԼ����")
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
                strSQL = strSQL & " And B.����ID=[2]" & str����Ժ
                strInput = "-" & lng����ID
                blnHavePassWord = True
        End Select
    End If
    
    strTmp = strSQL: strSQL = ""
    strSQL = strSQL & vbNewLine & " Select distinct A.NO,A.���㵥λ as �ű�,A.ִ�в���id,C.���� as  �Һſ���, A.����ID,D.���� as �Һ���Ŀ,   "
    strSQL = strSQL & vbNewLine & "       to_char(A.����ʱ��,'yyyy-mm-dd hh24:mi:ss') as ԤԼʱ��,B.���֤��  "
    strSQL = strSQL & vbNewLine & " From ������ü�¼ A, ������Ϣ B,���ű� C,�շ���ĿĿ¼ D"
    strSQL = strSQL & vbNewLine & " Where A.��¼���� = 4 And A.��¼״̬ = 0 and ���=1  And A.ִ�в���ID=C.ID "
    strSQL = strSQL & vbNewLine & "       And A.����id = B.����id(+) And  a.�շ�ϸĿId=d.ID(+) "
    If mstr����IDs <> "" Then strSQL = strSQL & " And A.ִ�в���ID In (Select  /*+cardinality(J,10) */ Column_Value From Table(f_num2list([5]))) "

    strSQL = strSQL & vbNewLine & IIf(mintԤԼʧЧ���� > 0, "  And A.����ʱ�� between trunc(sysdate) and  trunc(sysdate)+1-1/24/60/60 ", _
                                  "  And ((nvl(A.�Ӱ��־,0) =0 And A.����ʱ�� > Trunc(Sysdate) - [3]) or  (nvl(A.�Ӱ��־,0) =1 And A.����ʱ�� > Trunc(Sysdate) - [4]))   ")
    strSQL = strSQL & vbNewLine & strTmp
    
    
    'û�����ú�����,������ǰ�Ĵ���ʽ,����ֻ��ȡ�����ԤԼ��(���ʧ��Լ��,���Ժ�ɫ������ʾ)
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, CStr(Mid(strInput, 2)), gSysPara.Sy_Reg.bytNODaysGeneral, gSysPara.Sy_Reg.bytNoDayseMergency, mstr����IDs)
    If rsTmp.RecordCount = 0 Then
        vsRegist.Clear 1: vsRegist.Rows = 2: vsRegist.Row = 1
        Exit Sub
    End If
    
    If Val(Nvl(rsTmp!����ID)) <> 0 Then
        Call zlAutoCalcBackLists(Val(Nvl(rsTmp!����ID))) '�Զ����������
        
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
    If rsTmp.RecordCount = 1 Then
        'ֻ��һ��,ֱ�ӷ���:
        mstrNo = Nvl(rsTmp!NO): mblnOk = True: Unload Me
        Exit Sub
    End If
    With vsRegist
        .Clear 1: .Rows = 2
        If rsTmp.RecordCount <> 0 Then .Rows = rsTmp.RecordCount + 1
        lngRow = 1
        Do While Not rsTmp.EOF
            .TextMatrix(lngRow, .ColIndex("ԤԼ���ݺ�")) = Nvl(rsTmp!NO)
            .TextMatrix(lngRow, .ColIndex("�ű�")) = Nvl(rsTmp!�ű�)
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTmp!�Һſ���)
            .TextMatrix(lngRow, .ColIndex("ԤԼʱ��")) = Nvl(rsTmp!ԤԼʱ��)
            .TextMatrix(lngRow, .ColIndex("���֤��")) = Nvl(rsTmp!���֤��)
            .TextMatrix(lngRow, .ColIndex("�Һ���Ŀ")) = Nvl(rsTmp!�Һ���Ŀ)
            lngRow = lngRow + 1
            rsTmp.MoveNext
        Loop
        
        zl_vsGrid_Para_Restore mlngModule, vsRegist, Me.Caption, "ԤԼ���б�", True
        .ColWidth(.ColIndex("��־")) = 285
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub imgColPlan_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgPlan.Hwnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlan.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsRegist, lngLeft, lngTop, imgColPlan.Height)
    zl_vsGrid_Para_Save mlngModule, vsRegist, Me.Caption, "ԤԼ���б�", True, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub picImgPlan_Click()
    Call imgColPlan_Click
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
    zl_vsGrid_Para_Save mlngModule, vsRegist, Me.Caption, "ԤԼ���б�", True, , InStr(1, mstrPrivs, ";��������;") > 0
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
    zl_vsGrid_Para_Save mlngModule, vsRegist, Me.Caption, "ԤԼ���б�", True, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    Dim lngPreIDKind As Long
    If Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        mblnNotClick = True
        lngPreIDKind = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("���֤��")
        txtPatient.Text = strID
        Call GetPatient(Trim(txtPatient.Text))
        If mblnOk Then Exit Sub
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
    End If
End Sub


Private Sub mobjICCard_ShowICCardInfo(ByVal strNO As String)
    Dim lngPreIDKind As Long
    If Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        lngPreIDKind = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("IC����")
        txtPatient.Text = strNO
        If txtPatient.Text <> "" Then
            Call GetPatient(Trim(txtPatient.Text))
        Else
            Call mobjICCard.SetEnabled(False) '��������Ϸ������������ü����Զ���ȡ
        End If
        IDKind.IDKind = lngPreIDKind
        If Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
    End If
End Sub

Private Sub CloseIDCard()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ر�������������
    '����:���˺�
    '����:2012-03-09 16:26:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mobjIDCard Is Nothing Then
        mobjIDCard.SetEnabled (False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        mobjICCard.SetEnabled (False)
        Set mobjICCard = Nothing
    End If
End Sub
Private Sub NewCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���µĿ�����
    '����:���˺�
    '����:2012-03-09 16:28:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.Hwnd)
    End If
    If mobjICCard Is Nothing Then
        Set mobjICCard = New clsICCard
        Call mobjICCard.SetParent(Me.Hwnd)
    End If
End Sub

'�ؼ������Ƿ�ƥ��
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
'��ȡidkind��Ĭ��kindֵ
Private Function IDKindDefaultKind() As Long
    Dim lngIndex As Long
    'IDkind��Ĭ��Kind
    If IDKind.DefaultCardType = "" Then
        lngIndex = -1
    Else
        If Not IsNumeric(IDKind.DefaultCardType) Then
           lngIndex = IDKind.GetKindIndex(IDKind.GetfaultCard.����)
        Else
           lngIndex = IDKind.GetKindIndex(IDKind.DefaultCardType)
        End If
    End If
    IDKindDefaultKind = lngIndex
End Function
