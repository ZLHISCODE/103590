VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#3.4#0"; "zlIDKind.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPatiInfoChangeLog 
   Caption         =   "���˻�����Ϣ�䶯��־"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13440
   Icon            =   "frmPatiInfoChangeLog.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7980
   ScaleWidth      =   13440
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraPati 
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   345
      Width           =   13440
      Begin VB.TextBox txtPatient 
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
         Left            =   1500
         TabIndex        =   2
         Top             =   270
         Width           =   2340
      End
      Begin VB.CommandButton cmdPati 
         Height          =   360
         Left            =   3840
         Picture         =   "frmPatiInfoChangeLog.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "ѡ����(F2)"
         Top             =   270
         Width           =   360
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   360
         Left            =   855
         TabIndex        =   1
         ToolTipText     =   "��ݼ�F4"
         Top             =   270
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   635
         Appearance      =   2
         IDKindStr       =   $"frmPatiInfoChangeLog.frx":6DDC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   10.5
         FontName        =   "����"
         IDKind          =   -1
         BackColor       =   -2147483633
      End
      Begin VB.Label lblID 
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
         Left            =   330
         TabIndex        =   7
         Top             =   345
         Width           =   420
      End
      Begin VB.Label lblSex 
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
         Left            =   4320
         TabIndex        =   4
         Top             =   345
         Width           =   630
      End
      Begin VB.Label lblAge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���䣺"
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
         Left            =   6090
         TabIndex        =   5
         Top             =   345
         Width           =   630
      End
      Begin VB.Label lblBirthday 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������ڣ�"
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
         Left            =   8355
         TabIndex        =   6
         Top             =   345
         Width           =   1050
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VsgData 
      Height          =   5865
      Left            =   0
      TabIndex        =   8
      Top             =   1185
      Width           =   13425
      _cx             =   23680
      _cy             =   10345
      Appearance      =   0
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
      BackColorSel    =   16764057
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   255
      RowHeightMax    =   5000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPatiInfoChangeLog.frx":6E63
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
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
      AutoSizeMouse   =   0   'False
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
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   7605
      Width           =   13440
      _ExtentX        =   23707
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatiInfoChangeLog.frx":6EC5
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20796
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPatiInfoChangeLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------
'65802:������,2013-11-14
'------------------------------------------

Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1

Private mlng����ID As Long
Private mblnNotClick As Boolean
Private mstrPrivs As String

Private Enum VFGDATACOL
    ���� = 0
    ����ID = 1
    �䶯��Ŀ = 2
    ԭ��Ϣ = 3
    ����Ϣ = 4
    �䶯ʱ�� = 5
    �䶯�� = 6
    �䶯ģ�� = 7
    �䶯˵�� = 8
End Enum


Public Sub ShowMe(frmParent As Object, ByVal strPrivs As String, Optional ByVal lng����ID As Long = 0)
'--------------------------------------------------------------------------------------------
'����:�鿴���˻�����Ϣ�䶯��־
'����:
'   frmParent:���ô������
'   strPrivs:Ȩ�޹����ַ���
'   lng����ID:����ID<>0��ֱ����ȡ����
'--------------------------------------------------------------------------------------------
    mstrPrivs = strPrivs
    mlng����ID = lng����ID
    mblnNotClick = False
    
    Me.Show 1, frmParent
End Sub
    
Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Long
    Dim objControl As CommandBarControl
    
    Select Case Control.ID
        Case conMenu_File_PrintSet '��ӡ����
            Call zlPrintSet
        Case conMenu_File_Preview  'Ԥ��
            Call OutputList(2)
        Case conMenu_File_Print   '��ӡ
            Call OutputList(1)
        Case conMenu_File_Excel   '�����Excel
            Call OutputList(3)
        Case conMenu_View_Refresh 'ˢ��
            Call LoadPatiChangeInfo(Val(txtPatient.Tag))
        Case conMenu_View_ToolBar_Button '������
            Control.Checked = Not Control.Checked
            For i = 2 To cbsMain.Count
                cbsMain(i).Visible = Not cbsMain(i).Visible
            Next
            cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Text '��ť����
            Control.Checked = Not Control.Checked
            For i = 2 To cbsMain.Count
                For Each objControl In cbsMain(i).Controls
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                Next
            Next
            cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Size '��ͼ��
            cbsMain.Options.LargeIcons = Not cbsMain.Options.LargeIcons
            Control.Checked = Not Control.Checked
            cbsMain.RecalcLayout
        Case conMenu_View_StatusBar '״̬��
            staThis.Visible = Not staThis.Visible
            Control.Checked = Not Control.Checked
            cbsMain.RecalcLayout
        Case conMenu_View_Refresh
            
        Case conMenu_Help_Web_Home 'Web�ϵ�����
            Call zlHomePage(hwnd)
        Case conMenu_Help_Web_Forum '������̳
            Call zlWebForum(hwnd)
        Case conMenu_Help_Web_Mail '���ͷ���
            Call zlMailTo(hwnd)
        Case conMenu_Help_About '����
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_Help_Help '����
            Call ShowHelp(App.ProductName, hwnd, Name, Int((glngSys) / 100))
        Case conMenu_File_Exit '�˳�
            Unload Me
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If staThis.Visible Then Bottom = staThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    With Me.fraPati
        .Left = lngLeft: .Top = lngTop
        .Width = lngRight - lngLeft
    End With
    
    With VsgData
        .Left = lngLeft: .Top = fraPati.Top + fraPati.Height
        .Width = lngRight - lngLeft
        .Height = lngBottom - .Top
    End With
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_View_ToolBar_Button '������
            If cbsMain.Count >= 2 Then
                Control.Checked = Me.cbsMain(2).Visible
            End If
        Case conMenu_View_ToolBar_Text 'ͼ������
            If cbsMain.Count >= 2 Then
                Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
            End If
        Case conMenu_View_ToolBar_Size '��ͼ��
            Control.Checked = Me.cbsMain.Options.LargeIcons
        Case conMenu_View_StatusBar '״̬��
            Control.Checked = Me.staThis.Visible
    End Select
End Sub

Private Sub cmdPati_Click()
    frmPatiSel.mstrPrivs = mstrPrivs
    frmPatiSel.Show 1, Me
    If frmPatiSel.mlng����ID <> 0 Then
        txtPatient.Text = "-" & frmPatiSel.mlng����ID
        mblnNotClick = True
        IDKind.IDKind = IDKind.GetKindIndex("����")
        mblnNotClick = False
        Call txtPatient_KeyPress(vbKeyReturn)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intIndex As Integer
    If KeyCode = vbKeyF4 Then
        If Shift = vbCtrlMask And IDKind.Enabled Then
            intIndex = IDKind.GetKindIndex("IC����")
            If intIndex < 0 Then Exit Sub
            IDKind.IDKind = intIndex: Call IDKind_Click(IDKind.GetCurCard)
        End If
    ElseIf KeyCode = vbKeyF2 Then
        Call cmdPati_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call CreateMobjCard
    Call CreateSquareCardObject(Me, 1101)
     '��ʼ��
    Call IDKind.zlInit(Me, 100, 1101, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    
    If Not gobjSquare.objSquareCard Is Nothing Then
        IDKind.IDKindStr = gobjSquare.objSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
    End If
    
    Call InitMainMunus
    
    RestoreWinState Me, App.ProductName
    
    If mlng����ID <> 0 Then
        txtPatient.Text = "-" & mlng����ID
        mblnNotClick = True
        IDKind.IDKind = IDKind.GetKindIndex("����")
        mblnNotClick = False
        Call txtPatient_KeyPress(vbKeyReturn)
    Else
        txtPatient.Text = ""
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
    SaveFlexState VsgData, App.ProductName & "\" & Me.Name
    SaveWinState Me, App.ProductName
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand As String
    Dim strOutPatiInforXml As String
    
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = New clsICCard
            Call mobjICCard.SetParent(Me.hwnd)
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txtPatient.Text = mobjICCard.Read_Card()
            If txtPatient.Text <> "" Then
                Call txtPatient_KeyPress(vbKeyReturn)
            End If
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
    If gobjSquare.objSquareCard.zlReadCard(Me, 1101, lng�����ID, False, strExpand, strOutCardNO, strOutPatiInforXml) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    
    Set gobjSquare.objCurCard = objCard
    txtPatient.PasswordChar = "": txtPatient.IMEMode = 0
    '��Ҫ�����Ϣ,����ˢ����,���л�,���������ʾʧȥ����
    If txtPatient.Text <> "" And mblnNotClick = False Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Text <> "" Or txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.����
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("IC��", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strCardNo
    If txtPatient.Text <> "" Then Call FindPati(objCard, False)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("���֤", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strID
    If txtPatient.Text <> "" Then Call FindPati(objCard, False)
End Sub

Private Sub txtPatient_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")
End Sub

Private Sub txtPatient_GotFocus()
    zlControl.TxtSelAll txtPatient
    If Not mobjIDCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then mobjIDCard.SetEnabled (True)
    If Not mobjICCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then mobjICCard.SetEnabled (True)
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")
End Sub

Private Function FindPati(ByVal objCard As Card, Optional blnCard As Boolean = False) As Boolean
    If Not GetPatient(objCard, txtPatient.Text, blnCard) Then
        If IsNumeric(txtPatient.Text) Then
            txtPatient.PasswordChar = "": txtPatient.IMEMode = 0: txtPatient.Text = ""
        End If
        Call zlControl.TxtSelAll(txtPatient)
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Call InitVsfDate(False)
    Else
        txtPatient.PasswordChar = ""
        txtPatient.IMEMode = 0
        Call LoadPatiChangeInfo(Val(txtPatient.Tag))
    End If
    FindPati = True
End Function

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, Optional blnCard As Boolean = False) As Boolean
'���ܣ���ȡ������Ϣ
    Dim lng�����ID As Long, lng����ID As Long
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnDo As Boolean
    Dim blnHavePassWord As Boolean
    Dim strPassWord As String, strErrMsg As String
    Dim strCard As String, strPati As String
    Dim vRect As RECT
    Dim blnCancel As Boolean
    
    On Error GoTo errH
    
    strSQL = "Select A.����ID,A.����,A.�Ա�,A.����,A.��������,A.��������,A.����" & _
        " From ������Ϣ A" & _
        " Where A.ͣ��ʱ�� is NULL"
        
    If blnCard = True And objCard.���� Like "����*" Then    'ˢ��
        If IDKind.Cards.��ȱʡ������ And Not IDKind.GetfaultCard Is Nothing Then
            lng�����ID = IDKind.GetfaultCard.�ӿ����
        Else
            lng�����ID = "-1"
        End If
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng����ID
        strSQL = strSQL & " And A.����ID=[1]"
        blnHavePassWord = True
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '����ID
        strSQL = strSQL & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��
        strSQL = strSQL & " And A.סԺ��=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����
        strSQL = strSQL & " And A.�����=[1]"
    Else
        Select Case objCard.����
            Case "����", "��������￨"
                If gblnShowCard = True Then
                    strCard = "A.���￨�� as ���￨,A.���￨�� as ���￨��,"
                Else
                    strCard = "LPAD('*',Length(A.���￨��),'*') as ���￨,A.���￨�� as ���￨��,"
                End If
                'ͨ������ģ�����Ҳ���(�������벡�˱�ʶʱ)
                strPati = _
                    " Select A.����ID ID,A.����ID,A.�����,A.סԺ��," & strCard & "A.����,A.�Ա�,A.����,A.�ѱ� as ����ѱ�," & _
                    "   B.���� as ����,C.���� as ����,A.��ǰ���� as ����,To_Char(A.��Ժʱ��,'YYYY-MM-DD') as ��Ժʱ��," & _
                    "   To_Char(A.��Ժʱ��,'YYYY-MM-DD') as ��Ժʱ��,A.סԺ����,To_Char(A.��������,'YYYY-MM-DD') as ��������," & _
                    "   A.����,A.����,A.����,A.ѧ��,A.ְҵ,A.���,A.���֤��,A.��ͥ��ַ,A.������λ,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD') as �Ǽ�ʱ��," & _
                    "   Nvl(P.��������,Decode(P.����,Null,'��ͨ����','ҽ������')) ��������" & _
                    " From ������ҳ P,������Ϣ A,���ű� B,���ű� C" & _
                    " Where A.��ǰ����ID=B.ID(+) And A.��ǰ����ID=C.ID(+) And A.����ID=P.����ID(+) And A.��ҳID=P.��ҳID(+)" & _
                    "   And Nvl(P.��ҳID(+),0)<>0 And A.ͣ��ʱ�� is NULL And A.���� Like [1]" & _
                    " Order by A.����,A.�Ǽ�ʱ�� Desc"
                
                vRect = zlControl.GetControlRect(txtPatient.hwnd)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%")
                            
                'ֻ��һ������ʱ,blncancel����false,��ȡ������Ҳ��һ��
                If Not rsTmp Is Nothing Then
                    strSQL = strSQL & " And A.����ID=[1]"
                    lng����ID = Val(Nvl(rsTmp!����ID))
                    If lng����ID <= 0 Then GoTo NotFoundPati:
                    strInput = "-" & lng����ID
                ElseIf blnCancel = True Then
                    strSQL = strSQL & " And A.����ID=[1]"
                    lng����ID = Val(txtPatient.Tag)
                    If lng����ID <= 0 Then GoTo NotFoundPati:
                    strInput = "-" & lng����ID
                Else
                    GoTo NotFoundPati
                End If
            Case "ҽ����"
                strInput = UCase(strInput)
                strSQL = strSQL & " And A.ҽ����=[2]"
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.�����=[2]"
            Case Else
                '��������,��ȡ��صĲ���ID
                If Val(objCard.�ӿ����) > 0 Then
                    lng�����ID = Val(objCard.�ӿ����)
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
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    
    blnDo = Not rsTmp.EOF
    
    If blnDo Then
        txtPatient.Tag = rsTmp!����ID
        txtPatient.Text = rsTmp!����
        '74426:���ϴ�,2014-7-9,����������ʾ��ɫ����
        Call SetPatiColor(txtPatient, Nvl(rsTmp!��������), IIf(IsNull(rsTmp!����), Me.ForeColor, vbRed))
        lblSex.Caption = "�Ա�" & Nvl(rsTmp!�Ա�)
        lblAge.Caption = "���䣺" & Nvl(rsTmp!����)
        lblBirthday.Caption = "�������ڣ�" & Format(Nvl(rsTmp!��������), "YYYY-MM-DD HH:mm")
        
        GetPatient = True
    Else
NotFoundPati:
        txtPatient.Tag = ""
        txtPatient.Text = ""
        txtPatient.ForeColor = Me.ForeColor
        lblSex.Caption = "�Ա�"
        lblAge.Caption = "���䣺"
        lblBirthday.Caption = "�������ڣ�"
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Then Exit Sub
    Call IDKind.ActiveFastKey
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    Dim blnCard As Boolean
    
    If IDKind.GetCurCard.���� = "����" Then
        blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
    ElseIf IDKind.GetCurCard.���� = "�����" Or IDKind.GetCurCard.���� = "סԺ��" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        txtPatient.IMEMode = 0
    End If
    
    'ˢ����ϻ���������س�
    If blnCard And Len(Me.txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Me.txtPatient.Text <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        
        '��ȡ������Ϣ
        Call FindPati(IDKind.GetCurCard, blnCard)
    End If
End Sub

Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (False)
    txtPatient.Text = Trim(txtPatient.Text)
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtPatient.hwnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub CreateMobjCard()
    '����������
    Set mobjIDCard = New clsIDCard
    Call mobjIDCard.SetParent(Me.hwnd)
    Set mobjICCard = New clsICCard
    Call mobjICCard.SetParent(Me.hwnd)
    Set mobjICCard.gcnOracle = gcnOracle
End Sub

Private Sub InitVsfDate(Optional blnSet As Boolean)
'����:��ʼ��������Ϣ�䶯��־���

    Dim strHead As String
    Dim i As Integer

    strHead = "����,4,500|����ID,1,0|�䶯��Ŀ,4,1000|ԭ��Ϣ,4,1500|����Ϣ,4,1500|�䶯ʱ��,1,2000|�䶯��,1,1000|�䶯ģ��,1,1200|�䶯˵��,1,4000"
    
    With VsgData
        .Redraw = False
        .Clear
        .Rows = 2
        .Cols = UBound(Split(strHead, "|")) + 1
        
        .MergeCells = flexMergeRestrictColumns
        .MergeCellsFixed = flexMergeFree
        
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Or blnSet Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
        Next
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        
        If Not Visible Or blnSet Then Call RestoreFlexState(VsgData, App.ProductName & "\" & Me.Name)
        .FixedCols = 1
        .FixedRows = 1
        .ColHidden(����ID) = True

        .RowHeight(0) = 320
        .RowHeight(1) = 300
        '�ָ��ϴ���
        .Row = 1
        .Col = 1:
        .Redraw = True
    End With
    staThis.Panels(2).Text = "����ȷ������"
End Sub

Private Sub LoadPatiChangeInfo(ByVal lng����ID As Long)
'����:��ȡ�����ز�����Ϣ�䶯����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim lngRow As Long
    Dim strChangeDate As String
    Dim strTmp As String, lngNum As Long '����
    
    On Error GoTo Errhand
    
    If lng����ID = 0 Then
        Call InitVsfDate(True)
        Exit Sub
    End If
    
    strSQL = " Select ����id, �䶯��Ŀ, ԭ��Ϣ, ����Ϣ, �䶯ʱ��, �䶯��, �䶯ģ��, ˵�� �䶯˵��" & vbNewLine & _
            " From ������Ϣ�䶯" & vbNewLine & _
            " Where ����id = [1]" & vbNewLine & _
            " Order By �䶯ʱ��, �䶯��Ŀ Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "������Ϣ�䶯", lng����ID)
    With VsgData
        Call InitVsfDate(True)
        lngNum = 0
        strChangeDate = ""
        .Redraw = flexRDNone
        Do While Not rsTmp.EOF
            lngRow = rsTmp.AbsolutePosition
            If Format(strChangeDate, "YYYY-MM-DD HH:mm:ss") <> Format(Nvl(rsTmp!�䶯ʱ��), "YYYY-MM-DD HH:mm:ss") Then
                lngNum = lngNum + 1
                strChangeDate = Format(Nvl(rsTmp!�䶯ʱ��), "YYYY-MM-DD HH:mm:ss")
                If lngNum Mod 2 = 1 Then
                    strTmp = ""
                Else
                    strTmp = " "
                End If
            End If
            
            If lngRow > .Rows - 1 Then .Rows = .Rows + 1
            .TextMatrix(lngRow, ����) = lngNum
            .TextMatrix(lngRow, ����ID) = Nvl(rsTmp!����ID)
            .TextMatrix(lngRow, �䶯��Ŀ) = Nvl(rsTmp!�䶯��Ŀ)
            If Nvl(rsTmp!�䶯��Ŀ) = "��������" Then
                .TextMatrix(lngRow, ԭ��Ϣ) = Format(Nvl(rsTmp!ԭ��Ϣ), "YYYY-MM-DD HH:mm")
                .TextMatrix(lngRow, ����Ϣ) = Format(Nvl(rsTmp!����Ϣ), "YYYY-MM-DD HH:mm")
            Else
                .TextMatrix(lngRow, ԭ��Ϣ) = Nvl(rsTmp!ԭ��Ϣ)
                .TextMatrix(lngRow, ����Ϣ) = Nvl(rsTmp!����Ϣ)
            End If
            .TextMatrix(lngRow, �䶯ʱ��) = Format(Nvl(rsTmp!�䶯ʱ��), "YYYY-MM-DD HH:mm:ss") & strTmp
            .TextMatrix(lngRow, �䶯��) = Nvl(rsTmp!�䶯��) & strTmp
            .TextMatrix(lngRow, �䶯ģ��) = Nvl(rsTmp!�䶯ģ��) & strTmp
            .TextMatrix(lngRow, �䶯˵��) = Nvl(rsTmp!�䶯˵��) & strTmp
        rsTmp.MoveNext
        Loop
        
        .WordWrap = False
        .MergeCol(����) = False
        .MergeCol(�䶯ʱ��) = False
        .MergeCol(�䶯��) = False
        .MergeCol(�䶯ģ��) = False
        .MergeCol(�䶯˵��) = False
        
        .AutoResize = True
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 0, .Cols - 1
        .AutoResize = False
        .RowHeight(0) = 320
        
        .MergeCol(����) = True
        .MergeCol(�䶯ʱ��) = True
        .MergeCol(�䶯��) = True
        .MergeCol(�䶯ģ��) = True
        .MergeCol(�䶯˵��) = True
        
        For lngRow = .FixedRows To .Rows - 1
            If .RowHeight(lngRow) < 300 Then .RowHeight(lngRow) = 300
        Next lngRow
        .Row = 1
        .Redraw = flexRDDirect
    End With
    
    If rsTmp.RecordCount > 0 Then
        staThis.Panels(2).Text = "���˹������ˡ�" & lngNum & "���λ�����Ϣ�䶯"
    Else
        staThis.Panels(2).Text = "�޻�����Ϣ�䶯��¼"
    End If
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub InitMainMunus()
    Dim objBar As CommandBar
    Dim objMenu As CommandBarPopup
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
        
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    Set Me.cbsMain.Icons = zlCommFun.GetPubIcons
    With Me.cbsMain.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 24, 24
        .LargeIcons = True
    End With
    
    '�˵�����
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup '��xtpControlPopup���͵�����ID�����¸�ֵ
    With objMenu.CommandBar.Controls
        .Add xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��"
        .Add(xtpControlButton, conMenu_File_Preview, "��ӡԤ��(&V)").BeginGroup = True
        .Add xtpControlButton, conMenu_File_Print, "��ӡ(&P)"
        .Add xtpControlButton, conMenu_File_Excel, "�����Excel(&L)"
        .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)").BeginGroup = True
    End With


    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
       Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)") '����
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False '����
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False '����
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False '����
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)") '����
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)")
    End With
    
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set objPopup = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�����")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, "������ҳ(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, "������̳(&F)", -1, False    '����
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): objControl.BeginGroup = True
    End With

    
    '���������⴦��
    '-----------------------------------------------------
'    '���˵��Ҳ�Ĳ���
'    With cbsMain.ActiveMenuBar.Controls
'        Set objCustom = .Add(xtpControlCustom, conMenu_View_Find, "")
'        objCustom.Handle = picFind.hwnd
'        objCustom.flags = xtpFlagRightAlign
'        IDKind.BackColor = picFind.BackColor
'    End With

    '����������
    '-----------------------------------------------------
    Set objBar = Me.cbsMain.Add("������", xtpBarTop)
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ") '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��") '����

        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): objControl.BeginGroup = True '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�") '����
    End With
    
     For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    '�����
    With cbsMain.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    cbsMain.RecalcLayout
End Sub


Private Sub OutputList(bytStyle As Byte)
'���ܣ���������˱䶯��־
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, lngRow As Long
    
    lngRow = VsgData.Row
    
    '��ͷ
    objOut.Title.Text = "���˻�����Ϣ�䶯��־"
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    objRow.Add "������" & txtPatient.Text
    objRow.Add lblSex.Caption
    objRow.Add lblAge.Caption
    objRow.Add lblBirthday.Caption
    objOut.UnderAppRows.Add objRow
    
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    '����
    VsgData.Redraw = False
    Set objOut.Body = VsgData
    
    '���
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    VsgData.Row = lngRow
    VsgData.Redraw = True
End Sub

Private Sub VsgData_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngRow As Long
    With VsgData
        .Redraw = flexRDNone
        .WordWrap = False
        .MergeCol(����) = False
        .MergeCol(�䶯ʱ��) = False
        .MergeCol(�䶯��) = False
        .MergeCol(�䶯ģ��) = False
        .MergeCol(�䶯˵��) = False
        .AutoResize = True
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 0, .Cols - 1
        .AutoResize = False

        .MergeCol(����) = True
        .MergeCol(�䶯ʱ��) = True
        .MergeCol(�䶯��) = True
        .MergeCol(�䶯ģ��) = True
        .MergeCol(�䶯˵��) = True

        For lngRow = .FixedRows To .Rows - 1
            If .RowHeight(lngRow) < 300 Then .RowHeight(lngRow) = 300
        Next lngRow
        .Redraw = flexRDDirect
    End With
End Sub

