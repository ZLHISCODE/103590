VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmAdviceRollSend 
   AutoRedraw      =   -1  'True
   Caption         =   "���ڷ����ջ�"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9540
   Icon            =   "frmAdviceRollSend.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   9540
   Begin VB.Frame fraSetup 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   9315
      Begin VB.Frame fraBaby 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6120
         TabIndex        =   9
         Top             =   50
         Visible         =   0   'False
         Width           =   3195
         Begin VB.OptionButton optBaby 
            Caption         =   "Ӥ��ҽ��"
            Height          =   180
            Index           =   2
            Left            =   2175
            TabIndex        =   12
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton optBaby 
            Caption         =   "����ҽ��"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton optBaby 
            Caption         =   "����ҽ��"
            Height          =   180
            Index           =   1
            Left            =   1080
            TabIndex        =   10
            Top             =   0
            Width           =   1020
         End
      End
   End
   Begin VB.Frame fraInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   60
      TabIndex        =   5
      Top             =   525
      Width           =   9435
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C0FFFF&
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   60
         Width           =   90
      End
   End
   Begin VB.TextBox txtPer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   7290
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "100%"
      Top             =   6255
      Visible         =   0   'False
      Width           =   405
   End
   Begin MSComctlLib.ProgressBar psb 
      Height          =   270
      Left            =   2115
      TabIndex        =   1
      Top             =   6210
      Visible         =   0   'False
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   6150
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAdviceRollSend.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13917
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   1376
            MinWidth        =   2
            Picture         =   "frmAdviceRollSend.frx":0E1E
            Text            =   "ͨ��"
            TextSave        =   "ͨ��"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   1376
            MinWidth        =   2
            Picture         =   "frmAdviceRollSend.frx":1408
            Text            =   "����"
            TextSave        =   "����"
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
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   714
      BandCount       =   1
      _CBWidth        =   9540
      _CBHeight       =   405
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   345
      Width1          =   3525
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   345
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   9420
         _ExtentX        =   16616
         _ExtentY        =   609
         ButtonWidth     =   1349
         ButtonHeight    =   609
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȫѡ"
               Key             =   "ȫѡ"
               Description     =   "ȫѡ"
               Object.ToolTipText     =   "ȫѡ(Ctrl+A)"
               Object.Tag             =   "ȫѡ"
               ImageKey        =   "ȫѡ"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȫ��"
               Key             =   "ȫ��"
               Description     =   "ȫ��"
               Object.ToolTipText     =   "ȫ��(Ctrl+R)"
               Object.Tag             =   "ȫ��"
               ImageKey        =   "ȫ��"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�ջ�"
               Key             =   "�ջ�"
               Description     =   "�ջ�"
               Object.ToolTipText     =   "�����ջ�ѡ���ҽ��(Ctrl+E)"
               Object.Tag             =   "�ջ�"
               ImageKey        =   "�ջ�"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "�����������������������嵥(F12)"
               Object.Tag             =   "����"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "����(F1)"
               Object.Tag             =   "����"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�(ALT+X)"
               Object.Tag             =   "�˳�"
               ImageKey        =   "�˳�"
            EndProperty
         EndProperty
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   4935
      Left            =   0
      TabIndex        =   7
      Top             =   1185
      Width           =   9540
      _cx             =   16828
      _cy             =   8705
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
      BackColorSel    =   16771802
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAdviceRollSend.frx":19F2
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
      OwnerDraw       =   1
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmAdviceRollSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mMainPrivs As String 'IN
Private mlng����ID As Long 'IN:���ڼ�¼������Ĳ������ϴη��Ͳ���
Private mlng����ID As Long 'IN
Private mlng��ҳID As Long 'IN
Private mblnAutoRoll As Boolean 'In,ֹͣ���Զ��ջ�
Private mblnOnePati As Boolean  '�»�ʿվ����(������ģʽ)����ֹͣȷ�ϲ���ʱ���ó����ջ�

Private mblnRoll As Boolean 'OUT:�Ƿ�ɹ��ջع���
Private mblnAdjustNum As Boolean '�Ƿ��г����ջص�����Ȩ�ޣ�����������
Private mblnֻ��ʾ��ǰ����ҽ��  As Boolean

Private mrsBill As ADODB.Recordset
Private mbln���ڸ��� As Boolean
Private mblnFirst As Boolean
Private mblnReturn As Boolean
Private mstr����IDs As String   '��ǰ������Ӧ�Ŀ���IDs+��ǰ����ID
Private mintҽ������Χ As Integer    'ҽ������Χ   0-����ҽ��,1-����ҽ��,2-Ӥ��ҽ��
Private mblnFirstLoad As Boolean
Private mlngҽ������ID As Long
Private mlngӤ������ID As Long
Private mblnLimit As Boolean '���η��͸�ҩ;�������Ƿ��Խ���ʱ������
Private mbln�������� As Boolean '�Ǵ������Һ������Һ֮���Ƿ���Խ�����������

Private Const COL_ѡ�� = 0
Private Const COL_���� = 1
Private Const COL_���� = 2
Private Const COL_סԺ�� = 3
Private Const COL_���� = 4
Private Const COL_Ӥ�� = 5
Private Const col_ҽ������ = 6
Private Const COL_��� = 7
Private Const COL_���� = 8
Private Const COL_��λ = 9
Private Const COL_Ƶ�� = 10
Private Const COL_�÷� = 11
Private Const COL_ִ��ʱ�� = 12
Private Const COL_�ϴ�ִ�� = 13
Private Const COL_��ֹʱ�� = 14
Private Const COL_ִ�п��� = 15
Private Const COL_����ID = 16
Private Const COL_��ҳID = 17
Private Const COL_���� = 18
Private Const COL_ID = 19
Private Const COL_���ID = 20
Private Const COL_������� = 21
Private Const COL_ҩƷID = 22
Private Const COL_���˿���ID = 23
Private Const COL_��������ID = 24
Private Const COL_����ҽ�� = 25
Private Const COL_ִ�п���ID = 26
Private Const COL_���� = 27
Private Const COL_������ = 28
Private Const COL_���� = 29
Private Const COL_����ϵ�� = 30
Private Const COL_סԺ��װ = 31
Private Const COL_�ɷ���� = 32
Private Const COL_ִ������ = 33
Private Const COL_�ϴ� = 34 '�ջغ�Ӧ�õ��ϴ�ִ��ʱ��
Private Const COL_�������� = 35 '��ҺҩƷҽ�����ж�
Private Const COL_ִ�з��� = 36
Private Const COL_�������� = 37

Private Property Let Progress(ByVal vNewValue As Single)
'vNewValue=0-100
    If vNewValue = 0 Then
        psb.value = 0: txtPer.Text = ""
        psb.Visible = False: txtPer.Visible = False
    Else
        psb.value = vNewValue
        txtPer.Text = CInt(psb.value) & "%"
        psb.Visible = True: txtPer.Visible = True
        txtPer.Refresh
    End If
End Property


Public Function ShowMe(frmParent As Object, ByVal MainPrivs As String, _
    ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
     ByVal blnOnePati As Boolean, ByVal blnAutoRoll As Boolean, Optional ByVal lngҽ������ID As Long, Optional ByVal lngӤ������ID As Long) As Boolean
'������
'       blnOnePati=������ģʽ
    mMainPrivs = MainPrivs
    
    mlng����ID = lng����ID
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mlngҽ������ID = lngҽ������ID
    mlngӤ������ID = lngӤ������ID
    mblnOnePati = blnOnePati
    mblnAutoRoll = blnAutoRoll
        
    Me.Show 1, frmParent
    ShowMe = mblnRoll
End Function

Private Sub Form_Activate()
    Dim blnAutoRoll As Boolean
    
    If mblnFirst Then
        mblnFirst = False
        '������ģʽ
        If mblnOnePati Then
            Call LoadAdviceRoll(mlng����ID, mlng��ҳID)
            tbr.Buttons("����").Visible = False
                    
            If mblnAutoRoll Then
                Call tbr_ButtonClick(tbr.Buttons("�ջ�"))
            End If
        Else
            If Not ResetSend Then Unload Me: Exit Sub
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        Call tbr_ButtonClick(tbr.Buttons("ȫѡ"))
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call tbr_ButtonClick(tbr.Buttons("ȫ��"))
    ElseIf KeyCode = vbKeyE And Shift = vbCtrlMask Then
        Call tbr_ButtonClick(tbr.Buttons("�ջ�"))
    ElseIf KeyCode = vbKeyF12 And Shift = 0 Then
        Call tbr_ButtonClick(tbr.Buttons("����"))
    ElseIf KeyCode = vbKeyF1 And Shift = 0 Then
        Call tbr_ButtonClick(tbr.Buttons("����"))
    ElseIf KeyCode = vbKeyX And Shift = vbAltMask Then
        Call tbr_ButtonClick(tbr.Buttons("�˳�"))
    End If
End Sub

Private Sub Form_Load()

    '���ù�����ťͼ��
    Set tbr.HotImageList = frmIcons.imgColor
    Set tbr.ImageList = frmIcons.imgGray
    tbr.Buttons("ȫѡ").Image = "ȫѡ"
    tbr.Buttons("ȫ��").Image = "ȫ��"
    tbr.Buttons("�ջ�").Image = "ִ��"
    tbr.Buttons("����").Image = "����"
    tbr.Buttons("����").Image = "����"
    tbr.Buttons("�˳�").Image = "�˳�"
    tbr.ButtonHeight = 500
    mblnFirstLoad = True
        
    Call InitAdviceTable
    Call RestoreWinState(Me, App.ProductName)
    
    mblnRoll = False
    mblnFirst = True
    mbln���ڸ��� = Val(zlDatabase.GetPara("�����ջز�����������", glngSys, pסԺҽ������)) = 1
    mblnAdjustNum = InStr(GetInsidePrivs(pסԺҽ������), "�����ջص���") > 0
    mblnֻ��ʾ��ǰ����ҽ�� = Val(zlDatabase.GetPara("ֻ��ʾ��ǰ������ҽ��", glngSys, pסԺҽ������, "0")) = 1
    mblnLimit = Val(zlDatabase.GetPara("ҩ���������ƽ���ʱ��", glngSys, pסԺҽ������, 0)) = 1
    mstr����IDs = Get����IDs(mlng����ID)
    mbln�������� = Val(zlDatabase.GetPara("��Һ��Һ����ҩ��������������", glngSys, 1345, 0)) = 1
    
End Sub

Private Sub InitBillSet()
'���ܣ���ʼ��ҽ�����ʵ������ɼ�¼��
    Set mrsBill = New ADODB.Recordset
    
    mrsBill.Fields.Append "Key", adVarChar, 100
    mrsBill.Fields.Append "NO", adVarChar, 8
    mrsBill.CursorLocation = adUseClient
    mrsBill.LockType = adLockOptimistic
    mrsBill.CursorType = adOpenStatic
    mrsBill.Open
End Sub

Private Sub Form_Resize()
    Dim lngW As Long
    Dim i As Long
    
    On Error Resume Next
    
    fraInfo.Top = cbr.Height
    fraInfo.Left = 0
    fraInfo.Width = Me.ScaleWidth
    
    fraSetup.Top = fraInfo.Top + fraInfo.Height
    fraSetup.Left = 0
    fraSetup.Width = Me.ScaleWidth
    
    fraBaby.Left = fraSetup.Width - fraBaby.Width
    
    vsAdvice.Left = 0
    vsAdvice.Top = IIF(fraSetup.Visible, fraSetup.Top + fraSetup.Height, fraInfo.Top + fraInfo.Height)
    vsAdvice.Width = Me.ScaleWidth
    vsAdvice.Height = Me.ScaleHeight - fraInfo.Height - cbr.Height - stbThis.Height
    
    psb.Top = Me.ScaleHeight - stbThis.Height + 60
    psb.Left = stbThis.Panels(1).Width + 90
    
    For i = 1 To stbThis.Panels.Count
        If i <> 2 And stbThis.Panels(i).Visible Then
            lngW = lngW + (stbThis.Panels(i).Width + 60)
        End If
    Next
    psb.Width = Me.ScaleWidth - lngW - txtPer.Width - 500
    
    txtPer.Left = psb.Left + psb.Width
    txtPer.Top = psb.Top + (psb.Height - txtPer.Height) / 2
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    mMainPrivs = ""
    mlng����ID = 0
    mlng����ID = 0
    mstr����IDs = ""
    mlng��ҳID = 0
    mblnLimit = False
    Set mrsBill = Nothing
End Sub

Private Sub InitAdviceTable()
'���ܣ���ʼ���嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = ",300,4;����,850,1;����,750,1;סԺ��,750,1;����,500,4;Ӥ��,550,1;" & _
        "ҽ������,2000,1;���,2000,1;�ջ���,700,7;��λ,450,1;Ƶ��,1000,1;�÷�,1000,1;" & _
        "ִ��ʱ��,1000,1;�ϴ�ִ��,1530,1;��ֹʱ��,1530,1;ִ�п���,850,1;" & _
        "����ID;��ҳID;����;ID;���ID;�������;ҩƷID;���˿���ID;��������ID;����ҽ��;ִ�п���ID;" & _
        "����;������;����;����ϵ��;סԺ��װ;�ɷ����;ִ������;�ϴ�;��������;ִ�з���;��������"
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .FrozenCols = COL_ѡ�� + 1 - .FixedCols
        .ColDataType(COL_ѡ��) = flexDTBoolean
        .RowHeight(0) = 320
    End With
End Sub

Private Function ResetSend() As Boolean
'���ܣ����÷�������
    With frmAdviceRollSendCond
        .mMainPrivs = mMainPrivs
        .mlng����ID = mlng����ID
        If mlngӤ������ID <> 0 Then
            If mlngӤ������ID = mlngҽ������ID Then
                .mlng����ID = mlngӤ������ID
            End If
        End If
        .mlng����ID = mlng����ID
        .Show 1, Me
        If .mblnOK Then
            mlng����ID = .mlng����ID
            mstr����IDs = Get����IDs(mlng����ID)
            mlngҽ������ID = mlng����ID
            Call LoadAdviceRoll(.mstr����IDs, .mstr��ҳIDs)
        End If
        ResetSend = .mblnOK
    End With
End Function

Private Sub optBaby_Click(Index As Integer)
    mintҽ������Χ = Index
    '������ģʽ
    If Not mblnFirstLoad Then
        If mblnOnePati Then
            Call LoadAdviceRoll(mlng����ID, mlng��ҳID)
        Else
            Call LoadAdviceRoll(frmAdviceRollSendCond.mstr����IDs, frmAdviceRollSendCond.mstr��ҳIDs)
        End If
    End If
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim strҽ��IDs As String, i As Long, strMsg As String, strMsgAll As String
    
    Select Case Button.Key
        Case "ȫѡ"
            With vsAdvice
                For i = .FixedRows To .Rows - 1
                    If .RowHidden(i) = False And Val(.TextMatrix(i, COL_ѡ��)) = 0 Then
                        If RowCanRoll(i, strMsg) Then
                            .TextMatrix(i, COL_ѡ��) = 1
                            Call RowSelectSame(i)
                        Else
                            strMsgAll = strMsgAll & vbCrLf & .TextMatrix(i, col_ҽ������) & ":" & strMsg
                        End If
                    End If
                Next
                If strMsgAll <> "" Then
                    MsgBox "����ҽ�����ܳ����ջأ�" & strMsgAll, vbInformation, gstrSysName
                End If
            End With
        Case "ȫ��"
            vsAdvice.Cell(flexcpText, vsAdvice.FixedRows, COL_ѡ��, vsAdvice.Rows - 1, COL_ѡ��) = 0
        Case "�ջ�"
            With vsAdvice
                For i = .FixedRows To .Rows - 1
                    If Val(.TextMatrix(i, COL_ѡ��)) <> 0 And Val(.TextMatrix(i, COL_ID)) <> 0 Then
                        strҽ��IDs = strҽ��IDs & "," & Val(.TextMatrix(i, COL_ID))
                    End If
                Next
                If strҽ��IDs = "" Then
                    MsgBox "������ѡ��һ��Ҫ�ջص�ҽ����", vbInformation, gstrSysName
                    Exit Sub
                Else
                    strҽ��IDs = Mid(strҽ��IDs, 2)
                    
                    '��Ҫ�ջ�ҽ���ķ��ͷ��ý��н��ʼ��
                    If Not CheckRollMoneyBalance(strҽ��IDs) Then Exit Sub
                    
                    '��鲢��ʾ���շѶ���Ϊһ��ֻ��һ�Σ���һ�η���ֻ��һ�εȵ��շ���Ŀ
                    Call CheckRollPriceItem(strҽ��IDs)
                End If
            End With
            If mblnAutoRoll Then
                If RollAdvice(UBound(Split(strҽ��IDs, ",")) + 1) Then mblnRoll = True: Unload Me
            Else
                If MsgBox("ȷʵҪ�Ե�ǰѡ���ҽ��ִ���ջز�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    If RollAdvice(UBound(Split(strҽ��IDs, ",")) + 1) Then mblnRoll = True: Unload Me
                End If
            End If
        Case "����"
            Call ResetSend
        Case "����"
            ShowHelp App.ProductName, Me.hwnd, Me.Name
        Case "�˳�"
            Unload Me
    End Select
End Sub

Private Sub CheckRollPriceItem(ByVal strҽ��IDs As String)
'���ܣ���鲢��ʾ���շѶ���Ϊһ��ֻ��һ�Σ���һ�η���ֻ��һ�εȵ��շ���Ŀ�������  ҽ��ִ�мƼ����ݣ��������ջأ������ֹ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strTmp As String, i As Long
       
    strSQL = "Select  /*+ rule*/Distinct c.���� as �շ�����,e.���� as ��������" & vbNewLine & _
        "From ����ҽ���Ƽ� A,Table(f_Num2list([1])) B,�շ���ĿĿ¼ C,����ҽ����¼ D,������ĿĿ¼ E" & vbNewLine & _
        "Where a.ҽ��id = b.Column_Value And Nvl(a.�շѷ�ʽ, 0) <> 0 and a.�շ�ϸĿid=c.id And a.ҽ��id = d.id And d.������Ŀid = e.id"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�����ջ�", strҽ��IDs)
    For i = 1 To rsTmp.RecordCount
        strTmp = strTmp & vbCrLf & rsTmp!�������� & "��" & rsTmp!�շ�����
        If i > 9 Then
            strTmp = strTmp & "......"
            Exit For
        End If
        rsTmp.MoveNext
    Next
    If strTmp <> "" Then
        strSQL = "select Column_Value from Table(f_Num2list([1]))" & vbNewLine & _
            "minus" & vbNewLine & _
            "Select ҽ��id From ҽ��ִ�мƼ� Where ҽ��id In (select Column_Value from Table(f_Num2list([1]))) Group By ҽ��id"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�����ջ�", strҽ��IDs)
        If rsTmp.RecordCount <> 0 Then
            MsgBox "��鷢��Ҫ�ջ�ҽ���ķ��ô�������һ��ֻ��һ�λ�һ�η���ֻ��һ�ε���Ŀ��" & vbCrLf & _
                strTmp & vbCrLf & "�����޷���ȷ�ջ����������ǽ����ᱻ�Զ��ջأ������ʹ�������������ջء�", vbInformation, gstrSysName
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CheckRollMoneyBalance(ByVal strҽ��IDs As String) As Boolean
'���ܣ���Ҫ�ջ�ҽ���ķ��ͷ��ý��н��ʼ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    CheckRollMoneyBalance = True
    If gbytBillOpt = 0 Then Exit Function
    
    'ȡҽ��������͵ļ���NO
    strSQL = "Select Column_Value From Table(f_Num2list([1]))"
    strSQL = _
        " Select B.����,A.ҽ��ID,Decode(Instr(',4,5,6,7,',B.�������),0,C.����,B.ҽ������) as ҽ������,Max(A.NO) as NO" & _
        " From ����ҽ������ A,����ҽ����¼ B,������ĿĿ¼ C,(" & strSQL & ") X" & _
        " Where A.ҽ��ID=B.ID And B.������ĿID=C.ID And A.��¼����=2 And A.ҽ��ID=X.Column_Value" & _
        " Group by B.����,A.ҽ��ID,Decode(Instr(',4,5,6,7,',B.�������),0,C.����,B.ҽ������)"
    
    'ȡ��ЩNO�Ľ������(�ǻ���δ����)
    strSQL = "Select B.����,B.ҽ��ID,B.ҽ������,A.NO,Nvl(A.�۸񸸺�,A.���) as ���,Sum(Nvl(A.���ʽ��,0)) as ���ʽ��" & _
        " From סԺ���ü�¼ A,(" & strSQL & ") B" & _
        " Where A.NO=B.NO And A.ҽ�����=B.ҽ��ID And A.��¼���� IN(2,12) And A.��¼״̬=1" & _
        " Group by B.����,B.ҽ��ID,B.ҽ������,A.NO,Nvl(A.�۸񸸺�,A.���) Having Sum(Nvl(A.���ʽ��,0))<>0"
    strSQL = "Select /*+ Rule*/ ����,ҽ��ID,ҽ������ From (" & strSQL & ") Group by ����,ҽ��ID,ҽ������"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�����ջ�", strҽ��IDs)
    
    strSQL = ""
    Do While Not rsTmp.EOF
        If UBound(Split(strSQL, vbCrLf)) > 10 Then
            strSQL = strSQL & vbCrLf & "�� ��"
            Exit Do
        Else
            strSQL = strSQL & vbCrLf & "��" & rsTmp!���� & "��" & rsTmp!ҽ������
        End If
        rsTmp.MoveNext
    Loop
    
    If strSQL <> "" Then
        If gbytBillOpt = 1 Then
            If MsgBox("Ҫ�ջص�����ҽ����������ͷ��ô����ѽ��ʵ������" & vbCrLf & strSQL & vbCrLf & vbCrLf & "ȷʵҪִ���ջز�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                CheckRollMoneyBalance = False
            End If
        ElseIf gbytBillOpt = 2 Then
            MsgBox "Ҫ�ջص�����ҽ����������ͷ��ô����ѽ��ʵ������" & vbCrLf & strSQL & vbCrLf & vbCrLf & "����ִ���ջز�����", vbInformation, gstrSysName
            CheckRollMoneyBalance = False
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub RowSelectSame(ByVal lngRow As Long, Optional lngBegin As Long, Optional lngEnd As Long)
'���ܣ����ݿɼ��е�ѡ��״̬,�����ҽ��һ��ѡ��
    Dim lngS��ID As Long, lngO��ID As Long, i As Long
    
    With vsAdvice
        lngBegin = lngRow: lngEnd = lngRow
        lngS��ID = IIF(Val(.TextMatrix(lngRow, COL_���ID)) <> 0, Val(.TextMatrix(lngRow, COL_���ID)), Val(.TextMatrix(lngRow, COL_ID)))
        For i = lngRow + 1 To .Rows - 1
            lngO��ID = IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID)))
            If lngO��ID = lngS��ID Then
                .TextMatrix(i, COL_ѡ��) = .TextMatrix(lngRow, COL_ѡ��)
                lngEnd = i
            Else
                Exit For
            End If
        Next
        For i = lngRow - 1 To .FixedRows Step -1
            lngO��ID = IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID)))
            If lngO��ID = lngS��ID Then
                .TextMatrix(i, COL_ѡ��) = .TextMatrix(lngRow, COL_ѡ��)
                lngBegin = i
            Else
                Exit For
            End If
        Next
    End With
End Sub

Private Sub GetGroupRow(ByVal lngRow As Long, Optional lngBegin As Long, Optional lngEnd As Long)
'���ܣ����ݵ�ǰҽ���з���һ��ҽ�����з�Χ
    Dim lngS��ID As Long, lngO��ID As Long, i As Long
    
    With vsAdvice
        lngBegin = lngRow: lngEnd = lngRow
        lngS��ID = IIF(Val(.TextMatrix(lngRow, COL_���ID)) <> 0, Val(.TextMatrix(lngRow, COL_���ID)), Val(.TextMatrix(lngRow, COL_ID)))
        For i = lngRow + 1 To .Rows - 1
            lngO��ID = IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID)))
            If lngO��ID = lngS��ID Then
                lngEnd = i
            Else
                Exit For
            End If
        Next
        For i = lngRow - 1 To .FixedRows Step -1
            lngO��ID = IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID)))
            If lngO��ID = lngS��ID Then
                lngBegin = i
            Else
                Exit For
            End If
        Next
    End With
End Sub

Private Sub vsAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = COL_ѡ�� Then Call RowSelectSame(Row)
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewCol = COL_���� Then
        If Not CellEditable(NewRow, NewCol) Then
            vsAdvice.FocusRect = flexFocusLight
        Else
            vsAdvice.FocusRect = flexFocusHeavy
        End If
    Else
        vsAdvice.FocusRect = flexFocusLight
    End If
End Sub

Private Sub vsAdvice_AfterUserFreeze()
    With vsAdvice
        If .FrozenCols < COL_ѡ�� + 1 - .FixedCols Then
            .FrozenCols = COL_ѡ�� + 1 - .FixedCols
        End If
    End With
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    With vsAdvice
        If Col = col_ҽ������ Or Col = COL_��� Then
            If Not .ColHidden(COL_���) Then
                .AutoSize col_ҽ������, COL_���
            Else
                .AutoSize col_ҽ������
            End If
            .RowHeight(0) = 320
        ElseIf Row = -1 Then
            lngW = Me.TextWidth(.TextMatrix(.FixedRows - 1, Col) & "A")
            If .ColWidth(Col) < lngW Then
                .ColWidth(Col) = lngW
            ElseIf .ColWidth(Col) > .Width * 0.5 Then
                .ColWidth(Col) = .Width * 0.5
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COL_ѡ�� Then Cancel = True
End Sub

Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
'˵�������ص��кŷ�Χ��������ҩ;�����к�
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, COL_�������) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInһ����ҩ = blnTmp
    End With
End Function

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        lngLeft = COL_Ƶ��: lngRight = COL_�÷�
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = COL_����: lngRight = COL_Ӥ��
        End If
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        
        If Not RowInһ����ҩ(Row, lngBegin, lngEnd) Then Exit Sub
        
        vRect.Left = Left '������߱����
        vRect.Right = Right - 1 '�����ұ߱����
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '���б�����������
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 2 '���б����±���(���������õ��±��ߴ�Ϊ2)
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, OS.SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    Dim i As Long
    With vsAdvice
        If KeyAscii = 13 Then
            KeyAscii = 0
            For i = .Row + 1 To .Rows - 1
                If Not .RowHidden(i) Then
                    .Row = i: Exit For
                End If
            Next
            If i > .Rows - 1 And Not .RowHidden(.FixedRows) Then .Row = .FixedRows
            Call .ShowCell(.Row, .Col)
        End If
    End With
End Sub

Private Function AcceptInput(ByVal Row As Long, ByVal Col As Long) As Boolean
'���ܣ���鲢������������
    Dim lngBegin As Long, lngEnd As Long, i As Long
    Dim dblOnce As Double, dblModify As Double
    Dim lng���� As Long, lngMin���� As Long
    
    AcceptInput = False
    With vsAdvice
        If Val(.EditText) = Val(.TextMatrix(Row, Col)) Then AcceptInput = True: Exit Function
        
        '���������Ч��
        If Val(.TextMatrix(Row, COL_���ID)) <> 0 And InStr(",5,6,", "," & .TextMatrix(Row, COL_�������) & ",") > 0 Then
            If CheckAdvcieComPound(Val(.TextMatrix(Row, COL_���ID))) Then
                MsgBox "��Һ��ҩ�ļ�¼�������޸��ջ�����", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
            End If
        End If
        If Not IsNumeric(.EditText) Or Val(.EditText) < 0 Or Val(.EditText) > LONG_MAX Then
            MsgBox "������󣬲��Ǵ��ڵ���������ֻ�������ֵ����", vbInformation, gstrSysName
            .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
        End If
        If Val(.EditText) > Val(.TextMatrix(Row, COL_������)) Then
            MsgBox "�ջ������ܴ��� " & .TextMatrix(Row, COL_������) & .TextMatrix(Row, COL_��λ) & "��", vbInformation, gstrSysName
            .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
        End If
        If .TextMatrix(Row, COL_�������) = "E" And Val(.TextMatrix(Row, COL_ID)) = Val(.TextMatrix(Row - 1, COL_���ID)) _
            And InStr(",E,7,", .TextMatrix(Row - 1, COL_�������)) > 0 Then
            If Val(.EditText) <> Int(.EditText) Then
                MsgBox "��ҩ�䷽�ջظ���ӦΪ������", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
            End If
        End If
        
        '���յ�ǰ����ֵ
        .EditText = FormatEx(.EditText, 5)
        If InStr(",5,6,", .TextMatrix(Row, COL_�������)) > 0 Then
            'ҩƷҪ�ܷ������Լ���ҩ����
            If Val(.TextMatrix(Row, COL_�ɷ����)) = 0 Then
                '�ɷ���
            ElseIf Val(.TextMatrix(Row, COL_�ɷ����)) = 1 Or Val(.TextMatrix(Row, COL_�ɷ����)) < 0 Then
                '������:����סԺ��װ,����������ջ�ֵ�ٴ���
                .EditText = Int(Val(.EditText))
            ElseIf Val(.TextMatrix(Row, COL_�ɷ����)) = 2 Then
                'һ����:�㵥������������סԺ��װ,����������ջ�ֵ�ٴ���
                dblOnce = IntEx(Val(.TextMatrix(Row, COL_����)) / Val(.TextMatrix(Row, COL_����ϵ��)) / Val(.TextMatrix(Row, COL_סԺ��װ)))
                .EditText = Int(Val(.EditText) / dblOnce) * dblOnce
            End If
        End If
        .TextMatrix(Row, Col) = .EditText
        .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
        .Cell(flexcpFontBold, Row, Col) = Val(.TextMatrix(Row, Col)) <> Val(.TextMatrix(Row, COL_������)) '���Ϊ�޸Ĺ�
        If Val(.TextMatrix(Row, Col)) = 0 Then
            .TextMatrix(Row, COL_ѡ��) = 0
        Else
            If RowCanRoll(Row) Then
                .TextMatrix(Row, COL_ѡ��) = 1
            Else
                .TextMatrix(Row, COL_ѡ��) = 0
            End If
        End If
        Call RowSelectSame(Row, lngBegin, lngEnd)
        
        '�����������ֵ
        If InStr(",5,6,", .TextMatrix(Row, COL_�������)) > 0 Then
            '��ҩ;��
            lngMin���� = LONG_MAX
            For i = lngBegin To lngEnd
                If InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0 Then
                    If Val(.TextMatrix(i, COL_����)) = Val(.TextMatrix(i, COL_������)) Then
                        lng���� = 0 'δ�䶯��,�ָ�ԭ����
                    Else
                        '�󱾴��޸����ջص�����ִ�еĴ���,һ����ҩ����С��Ϊ׼
                        dblModify = Val(.TextMatrix(i, COL_������)) - Val(.TextMatrix(i, COL_����))
                        If Val(.TextMatrix(i, COL_�ɷ����)) = 0 Then
                            '�ɷ���,����������ջ�ֵ�ٴ���
                            lng���� = Int(dblModify * Val(.TextMatrix(i, COL_סԺ��װ)) * Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_����)))
                        ElseIf Val(.TextMatrix(i, COL_�ɷ����)) = 1 Or Val(.TextMatrix(i, COL_�ɷ����)) < 0 Then
                            '������:����סԺ��װ,����������ʵ���ջ���������ô���
                            lng���� = IntEx(Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_סԺ��װ)) * Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_����)))
                            lng���� = Val(.TextMatrix(i, COL_����)) - lng����
                        ElseIf Val(.TextMatrix(i, COL_�ɷ����)) = 2 Then
                            'һ����:�㵥������������סԺ��װ,����������ջ�ֵ�ٴ���
                            lng���� = Int(dblModify / IntEx(Val(.TextMatrix(i, COL_����)) / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_סԺ��װ))))
                        End If
                    End If
                    If lng���� < 0 Then lng���� = 0
                    If lng���� < lngMin���� Then lngMin���� = lng����
                ElseIf .TextMatrix(i, COL_�������) = "E" Then
                    If lngMin���� <> LONG_MAX Then
                        If Val(.TextMatrix(i, COL_����)) - lngMin���� >= 0 Then
                            .TextMatrix(i, COL_����) = Val(.TextMatrix(i, COL_����)) - lngMin����
                            .Cell(flexcpData, i, COL_����) = .TextMatrix(i, COL_����)
                        End If
                    End If
                End If
            Next
        Else
            '��ҩ�䷽���Լ���ҩƷ���:ͬ���뵱ǰ��������ͬ
            For i = lngBegin To lngEnd
                If i <> Row Then
                    .TextMatrix(i, COL_����) = .TextMatrix(Row, COL_����)
                    .Cell(flexcpData, i, COL_����) = .TextMatrix(i, COL_����)
                End If
            Next
        End If
    End With
    AcceptInput = True
End Function

Private Sub vsAdvice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsAdvice
        If mblnReturn Then mblnReturn = False
        If KeyAscii = 13 Then
            If Col = COL_���� Then
                KeyAscii = 0
                mblnReturn = True
                If Not AcceptInput(Row, Col) Then Exit Sub
                '��λ��һ����
                Call vsAdvice.FinishEditing(False)
                Call vsAdvice_KeyPress(13)
            End If
        Else
            If Col = COL_���� Then
                If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0: Exit Sub
                End If
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsAdvice.EditSelStart = 0
    vsAdvice.EditSelLength = zlCommFun.ActualLen(vsAdvice.EditText)
End Sub

Private Function CellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    CellEditable = True
    
    If lngCol = COL_���� Then
        If Not mblnAdjustNum Then
            CellEditable = False
        End If
    ElseIf lngCol <> COL_ѡ�� Then
        CellEditable = False
    End If
End Function

Private Sub vsAdvice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strMsg As String
    
    If Not CellEditable(Row, Col) Then
        Cancel = True
    Else
        If Col = COL_���� Then
            vsAdvice.EditMaxLength = 10
        Else
            vsAdvice.EditMaxLength = 0
        End If
        
        If Col = COL_ѡ�� And Val(vsAdvice.TextMatrix(Row, Col)) = 0 Then
            If Not RowCanRoll(Row, strMsg) Then
                Cancel = True
                MsgBox strMsg, vbInformation, gstrSysName
            End If
        End If
    End If
End Sub

Private Function LoadAdviceRoll(ByVal str����IDs As String, ByVal str��ҳIDs As String) As Boolean
'���ܣ���ȡָ�����˵ĳ��ڷ���ҽ���嵥,����ҩƷ����ҩҽ��
'������str����IDs=��������ID���ַ���
    Dim rsAdvice As New ADODB.Recordset
    Dim rsDrug As New ADODB.Recordset
    Dim rsSend As New ADODB.Recordset
    Dim strSQL As String, str������ As String, lngҩƷID As Long
    Dim str���� As String, lng������ As Long, lng����ID As Long
    Dim strPause As String, lng���� As Long, dbl���� As Double, dbl����All As Double
    Dim arr�ֽ�ʱ�� As Variant, str�ֽ�ʱ�� As String, str�ϴ�ʱ�� As String
    Dim datBegin As Date, lngRow As Long, i As Long, j As Long, k As Long
    Dim lngDel��ID As Long
    Dim int�ɷ���� As Integer, strUnRoll As String
    
    Screen.MousePointer = 11
    lblInfo.Caption = "���ڶ�ȡ����...."
    
    vsAdvice.Rows = vsAdvice.FixedRows
    vsAdvice.ColHidden(COL_���) = True
    vsAdvice.ColHidden(COL_����) = True
    vsAdvice.ColHidden(COL_Ӥ��) = True
    Me.Refresh
    
    If DeptIsWoman(0, Get����IDs(mlng����ID)) Then
        If mblnFirstLoad Then
            fraSetup.Visible = True
            fraBaby.Visible = True
            'ҽ������Χ
            mintҽ������Χ = Val(zlDatabase.GetPara("ҽ������Χ", glngSys, pסԺҽ������, "0"))
            optBaby(mintҽ������Χ).value = True
            mblnFirstLoad = False
        End If
    Else
        mblnFirstLoad = True
        fraSetup.Visible = False
        optBaby(0).value = True
    End If
    Call Form_Resize
    
    strUnRoll = zlDatabase.GetPara("��ҩ���ջ�", glngSys, pסԺҽ������)
    
    '����������ȼ�,��ǰ����ҽ���Ͷ������ֲ����͵�ҽ��(��ҩ;��,�䷽�巨,�÷�Ҳ����Ϊ����)
    'Ӧ�ò������������(����)
    'ע��"������"������ֹ���첻����
    '��ҩ�÷���ʹ����Ҳ�̶�Ҫ������(�������ջظ���)
    str������ = "(A.ִ��ʱ�䷽�� is NULL And (Nvl(A.Ƶ�ʴ���,0)=0 Or Nvl(A.Ƶ�ʼ��,0)=0 Or A.Ƶ�ʼ�� is NULL))"
    
    For k = 0 To UBound(Split(str����IDs, ","))
        strSQL = "Select A.ID,A.���ID,Nvl(A.���ID,A.ID) as ��ID,Nvl(X.���,A.���) as ���," & _
            " D.���� as ����,A.����ID,A.��ҳID,B.����,A.�շ�ϸĿID,A.����,B.סԺ��,B.��Ժ���� as ����,B.��Ժ����," & _
            " A.Ӥ��,A.ҽ������,A.�������,A.������ĿID,A.���˿���ID,A.��������ID,A.����ҽ��,A.�ܸ�����,A.��������," & _
            " A.ִ��Ƶ�� as Ƶ��,E.���㵥λ,E.���� as ������Ŀ,Nvl(F.����,Decode(Nvl(A.ִ������,0),5,'-')) as ִ�п���,A.ִ�п���ID,A.ִ������," & _
            " A.��ʼִ��ʱ��,A.ִ��ʱ�䷽��,A.�ϴ�ִ��ʱ��,A.ִ����ֹʱ��,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ," & _
            " A.�ɷ����,Decode(Instr(',5,6,',A.�������),0,NULL,G.����) as ��ҩ;��,A.�״�����,e.��������,e.ִ�з���,b.��������" & _
            " From ����ҽ����¼ A,����ҽ����¼ X,������ҳ B,������Ϣ C,���ű� D,������ĿĿ¼ E,���ű� F,������ĿĿ¼ G" & _
            " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID" & _
            " And A.����ID=C.����ID And B.��Ժ����ID=D.ID And A.������ĿID=E.ID" & _
            " And A.ִ�п���ID=F.ID(+) And A.����ID=[1] And A.��ҳID=[2]" & _
            " And (Nvl(A.ִ������,0)<>0 Or A.�������='E' And E.��������='4')" & _
            " And Not(A.�������='H' And E.��������='1' And E.ִ��Ƶ��=2) And Not(A.�������='Z' And E.�������� In('4','14'))" & _
            " And Nvl(A.ҽ����Ч,0)=0 And A.���ID=X.ID(+) And X.������ĿID = G.ID(+)" & _
            " And ((Not " & str������ & " And A.ִ����ֹʱ��<A.�ϴ�ִ��ʱ��)" & _
            " Or (" & str������ & " And Trunc(A.ִ����ֹʱ��)<Trunc(A.�ϴ�ִ��ʱ��)+1))" & _
            " And A.��ʼִ��ʱ�� is Not NULL And Nvl(A.ҽ��״̬,0)<>-1" & _
            " And Nvl(A.ִ�б��,0)<>-1 And A.������Դ<>3 And NVL(a.ִ��Ƶ��,'��')<>'��Ҫʱ' And NVL(a.ִ��Ƶ��,'��')<>'��Ҫʱ'" & _
            IIF(mblnֻ��ʾ��ǰ����ҽ��, " And instr(',' || [3] || ',',',' || Decode(NVL(A.Ӥ��,0),0,a.���˿���ID,NVL(b.Ӥ������ID,a.���˿���ID)) || ',')>0 ", "") & _
            Decode(mintҽ������Χ, 1, " And nvl(a.Ӥ��,0) = 0 ", 2, " And nvl(a.Ӥ��,0) <> 0 ", "") & _
            " And (B.Ӥ������ID is null or B.Ӥ������ID is not null and B.Ӥ������ID=[4] and NVL(A.Ӥ��,0)<>0 or B.Ӥ������ID is not null and B.Ӥ������ID<>[4] and NVL(A.Ӥ��,0)=0)" & _
            " Order by D.����,LPAD(B.��Ժ����,10,' '),A.Ӥ��,���,��ID,A.���"
        On Error GoTo errH
        Set rsAdvice = zlDatabase.OpenSQLRecord(strSQL, "�����ջ�", Val(Split(str����IDs, ",")(k)), Val(Split(str��ҳIDs, ",")(k)), mstr����IDs, mlngҽ������ID)
        
        '���㲢��ʾ�ջ��嵥
        '----------------------------------------------------------------------------------------------------------
        If Not rsAdvice.EOF Then
            With vsAdvice
                .Redraw = flexRDNone
                For i = 1 To rsAdvice.RecordCount
                    If NVL(rsAdvice!���ID, rsAdvice!ID) = lngDel��ID Then
                        GoTo NextLoop 'һ��ҽ�����������ٴ���
                    Else
                        lngDel��ID = 0
                    End If
                    
                    '���뵱ǰ��
                    .Rows = .Rows + 1: lngRow = .Rows - 1
                    
                    '���������(��Ȼ�����������,��Ҳ������)
                    If rsAdvice!������� = "7" Then
                        .RowHidden(lngRow) = True '��ζ��ҩ��
                    ElseIf rsAdvice!������� = "E" And NVL(rsAdvice!���ID, 0) = Val(.TextMatrix(lngRow - 1, COL_���ID)) And NVL(rsAdvice!���ID, 0) <> 0 Then
                        .RowHidden(lngRow) = True '�䷽�巨��
                    ElseIf rsAdvice!������� = "E" And rsAdvice!ID = Val(.TextMatrix(lngRow - 1, COL_���ID)) _
                        And InStr(",5,6,", .TextMatrix(lngRow - 1, COL_�������)) > 0 Then
                        .RowHidden(lngRow) = True '��ҩ;��
                    ElseIf InStr(",D,F,G,E,", rsAdvice!�������) > 0 And Not IsNull(rsAdvice!���ID) Then
                        .RowHidden(lngRow) = True '��鲿λ,��������,��������,��Ѫ;��
                    End If
                    
                    'һ���и�ֵ
                    '---------------------------------------------------------------
                    If NVL(rsAdvice!Ӥ��, 0) = 0 Then
                        .TextMatrix(lngRow, COL_Ӥ��) = "����"
                    Else
                        .TextMatrix(lngRow, COL_Ӥ��) = "Ӥ��" & rsAdvice!Ӥ��
                        .ColHidden(COL_Ӥ��) = False '��Ӥ��ҽ��ʱ����ʾ
                    End If
                    
                    .TextMatrix(lngRow, COL_����) = rsAdvice!����
                    If InStr(str���� & ",", "," & rsAdvice!���� & ",") = 0 Then
                        If str���� <> "" Then .ColHidden(COL_����) = False
                        str���� = str���� & "," & rsAdvice!����
                    End If
                    
                    .TextMatrix(lngRow, COL_����ID) = rsAdvice!����ID
                    .TextMatrix(lngRow, COL_��ҳID) = rsAdvice!��ҳID
                    .TextMatrix(lngRow, COL_����) = NVL(rsAdvice!����)
                    .TextMatrix(lngRow, COL_����) = rsAdvice!����
                    .TextMatrix(lngRow, COL_סԺ��) = NVL(rsAdvice!סԺ��)
                    .TextMatrix(lngRow, COL_����) = NVL(rsAdvice!����)
                    .TextMatrix(lngRow, COL_ID) = rsAdvice!ID
                    .TextMatrix(lngRow, COL_���ID) = NVL(rsAdvice!���ID)
                    .TextMatrix(lngRow, COL_�������) = rsAdvice!�������
                    .TextMatrix(lngRow, col_ҽ������) = NVL(rsAdvice!ҽ������)
                    .TextMatrix(lngRow, COL_��λ) = NVL(rsAdvice!���㵥λ)
                    .TextMatrix(lngRow, COL_Ƶ��) = NVL(rsAdvice!Ƶ��)
                    .TextMatrix(lngRow, COL_ִ��ʱ��) = NVL(rsAdvice!ִ��ʱ�䷽��)
                    .TextMatrix(lngRow, COL_�ϴ�ִ��) = Format(NVL(rsAdvice!�ϴ�ִ��ʱ��), "yyyy-MM-dd HH:mm")
                    .TextMatrix(lngRow, COL_��ֹʱ��) = Format(NVL(rsAdvice!ִ����ֹʱ��), "yyyy-MM-dd HH:mm")
                    .TextMatrix(lngRow, COL_ִ�п���) = NVL(rsAdvice!ִ�п���)
                    .TextMatrix(lngRow, COL_ִ�п���ID) = NVL(rsAdvice!ִ�п���ID, 0)
                    .TextMatrix(lngRow, COL_ִ������) = NVL(rsAdvice!ִ������, 0)
                    
                    .TextMatrix(lngRow, COL_���˿���ID) = NVL(rsAdvice!���˿���id, 0)
                    .TextMatrix(lngRow, COL_��������ID) = NVL(rsAdvice!��������id, 0)
                    .TextMatrix(lngRow, COL_����ҽ��) = NVL(rsAdvice!����ҽ��)
                    .TextMatrix(lngRow, COL_��������) = NVL(rsAdvice!��������)
                    .TextMatrix(lngRow, COL_ִ�з���) = NVL(rsAdvice!ִ�з���)
                    .TextMatrix(lngRow, COL_��������) = NVL(rsAdvice!��������)
                    
                    '�����ջش���(Ҫ����ͣʱ��,��Ȼ���ܶ��ջ�)
                    '---------------------------------------------------------------
                    lng���� = 0: str�ֽ�ʱ�� = "": str�ϴ�ʱ�� = ""
                    strPause = GetAdvicePause(rsAdvice!ID)
                    If IsNull(rsAdvice!ִ��ʱ�䷽��) And (NVL(rsAdvice!Ƶ�ʴ���, 0) = 0 Or NVL(rsAdvice!Ƶ�ʼ��, 0) = 0 Or IsNull(rsAdvice!�����λ)) Then
                        '"������"�ĳ���
                        Call Calc�����Գ�������(rsAdvice!��ʼִ��ʱ��, rsAdvice!�ϴ�ִ��ʱ��, "", "", strPause, "", "", str�ֽ�ʱ��)
                        arr�ֽ�ʱ�� = Split(str�ֽ�ʱ��, ",")
                        For j = 0 To UBound(arr�ֽ�ʱ��)
                            If Format(arr�ֽ�ʱ��(j), "yyyy-MM-dd") <= Format(rsAdvice!ִ����ֹʱ��, "yyyy-MM-dd") Then
                                str�ϴ�ʱ�� = Format(arr�ֽ�ʱ��(j), "yyyy-MM-dd HH:mm:ss")
                            Else
                                lng���� = lng���� + 1
                            End If
                        Next
                    Else
                        '"��ѡƵ��"����
                        str�ֽ�ʱ�� = Calc���ڷֽ�ʱ��(rsAdvice!��ʼִ��ʱ��, rsAdvice!�ϴ�ִ��ʱ��, strPause, NVL(rsAdvice!ִ��ʱ�䷽��), rsAdvice!Ƶ�ʴ���, rsAdvice!Ƶ�ʼ��, rsAdvice!�����λ, rsAdvice!��ʼִ��ʱ��)
                        arr�ֽ�ʱ�� = Split(str�ֽ�ʱ��, ",")
                        For j = 0 To UBound(arr�ֽ�ʱ��)
                            If arr�ֽ�ʱ��(j) <= Format(rsAdvice!ִ����ֹʱ��, "yyyy-MM-dd HH:mm:ss") Then
                                str�ϴ�ʱ�� = Format(arr�ֽ�ʱ��(j), "yyyy-MM-dd HH:mm:ss")
                            Else
                                lng���� = lng���� + 1
                            End If
                        Next
                    End If
                    If lng���� = 0 Then '�����ջص����
                        lngDel��ID = NVL(rsAdvice!���ID, rsAdvice!ID)
                        .RemoveItem lngRow: GoTo NextLoop
                    End If
                    
                    '�����ջ�����
                    '---------------------------------------------------------------
                    If rsAdvice!������� = "7" Then
                        '�����ǰ���䷽���������������൱��ÿ�εĵ���
                        .TextMatrix(lngRow, COL_����) = lng���� * NVL(rsAdvice!�ܸ�����, 1)
                    ElseIf InStr(",5,6,", rsAdvice!�������) > 0 Then
                        '�����г�ҩ
                        '------------------
                        '��ȡԭҩƷ���(�Ա�ҩ�޶�Ӧ����,��ҩƷĿ¼ȡһ�����)
                        lngҩƷID = 0
                        If Not IsNull(rsAdvice!�շ�ϸĿID) Then
                            lngҩƷID = rsAdvice!�շ�ϸĿID
                        Else
                            '������͵�ҩƷ�����е�ҩƷID:ҩƷ�϶���д�˷��ͼ�¼
                            'ҩƷֻ��һ��������Ŀ(�۸񸸺�=NULL),���ų��÷����ѱ���Ϊ����(���磺���۵���ɾ��)
                            lngҩƷID = GetLastSendMediCineID(Val(rsAdvice!ID), CDate(rsAdvice!�ϴ�ִ��ʱ��), Val(rsAdvice!�������� & ""))
                        End If
                        '�޷��ͻ��޶�Ӧ���õ�ҩƷ,Ҳ���ջ�(���Ա�ҩ���򻮼۵���ɾ��)
                        If lngҩƷID = 0 Then
                            lngDel��ID = NVL(rsAdvice!���ID, rsAdvice!ID)
                            .RemoveItem lngRow: GoTo NextLoop
                        End If
                                                
                        '�Ѿ����͹�,һ���й����Ϣ
                        strSQL = "Select A.ҩƷID,A.����ϵ��,A.סԺ��װ,A.סԺ��λ," & _
                            " A.ҩ������,A.סԺ�ɷ���� As �ɷ����,Nvl(C.����,B.����) as ����,B.���,B.����,A.��ҩ����" & _
                            " From ҩƷ��� A,�շ���ĿĿ¼ B,�շ���Ŀ���� C" & _
                            " Where A.ҩƷID=B.ID And A.ҩ��ID=[1] And A.ҩƷID=[2]" & _
                            " And B.ID=C.�շ�ϸĿID(+) And C.����(+)=1 And C.����(+)=[3] And Rownum=1"
                        Set rsDrug = zlDatabase.OpenSQLRecord(strSQL, "�����ջ�", Val(rsAdvice!������ĿID), lngҩƷID, IIF(gbytҩƷ������ʾ = 0, 1, 3))
                        
                        'һ����ҩ��Ͳ����ջأ�����ж�η��͵�,ֻ������һ���Ƿ�ҩ(��Ϊ�ֱ��ж��������ջش�����̫���ӣ�һ�����ְ����˷�ҩ����һ��)
                        If Not IsNull(rsDrug!��ҩ����) Then
                            If InStr("," & strUnRoll & ",", "," & rsDrug!��ҩ���� & ",") > 0 Then
                                If CheckMedicineSended(Val(rsAdvice!ID), CDate(rsAdvice!�ϴ�ִ��ʱ��)) Then
                                    lngDel��ID = NVL(rsAdvice!���ID, rsAdvice!ID)
                                    .RemoveItem lngRow: GoTo NextLoop
                                End If
                            End If
                        End If
                        
                        int�ɷ���� = NVL(rsAdvice!�ɷ����, NVL(rsDrug!�ɷ����, 0))
                        
                        .TextMatrix(lngRow, COL_ҩƷID) = rsDrug!ҩƷID '��¼���ڱ���ʱ����
                        .Cell(flexcpData, lngRow, COL_ҩƷID) = Val(NVL(rsDrug!ҩ������, 0))
                        
                        .TextMatrix(lngRow, COL_��λ) = rsDrug!סԺ��λ
                        .TextMatrix(lngRow, COL_���) = rsDrug!���� & IIF(Not IsNull(rsDrug!����), "(" & rsDrug!���� & ")", "") & IIF(Not IsNull(rsDrug!���), " " & rsDrug!���, "")
                        
                        '�����������Լ����ջ�����(סԺ��λ)
                        dbl���� = 0
                        If int�ɷ���� = 0 Then
                            '�ɷ���
                            dbl���� = NVL(rsAdvice!��������, 0) * lng���� / rsDrug!����ϵ�� / rsDrug!סԺ��װ
                            If str�ϴ�ʱ�� = "" And NVL(rsAdvice!�״�����, 0) <> 0 Then
                                '����ϴ�ʱ��Ϊ��������״�
                                dbl���� = dbl���� + (NVL(rsAdvice!�״�����, 0) - NVL(rsAdvice!��������, 0)) / rsDrug!����ϵ�� / rsDrug!סԺ��װ
                            End If
                        ElseIf int�ɷ���� = 1 Then
                            '������:������,����һ��סԺ��λ����,��İ�С�ڵ�������
                            dbl���� = Int(NVL(rsAdvice!��������, 0) * lng���� / rsDrug!����ϵ�� / rsDrug!סԺ��װ)
                            If str�ϴ�ʱ�� = "" And NVL(rsAdvice!�״�����, 0) <> 0 Then
                                '����ϴ�ʱ��Ϊ��������״�
                                dbl���� = dbl���� + (NVL(rsAdvice!�״�����, 0) - NVL(rsAdvice!��������, 0)) / rsDrug!����ϵ�� / rsDrug!סԺ��װ
                            End If
                        ElseIf int�ɷ���� = 2 Then
                            'һ����(��ʱʧЧ)
                            dbl���� = lng���� * IntEx(NVL(rsAdvice!��������, 0) / rsDrug!����ϵ�� / rsDrug!סԺ��װ)
                            If str�ϴ�ʱ�� = "" And NVL(rsAdvice!�״�����, 0) <> 0 Then
                                '����ϴ�ʱ��Ϊ��������״�
                                dbl���� = dbl���� + (NVL(rsAdvice!�״�����, 0) - NVL(rsAdvice!��������, 0)) / rsDrug!����ϵ�� / rsDrug!סԺ��װ
                            End If
                        ElseIf int�ɷ���� < 0 Then
                            'N���ڷ�����Ч:�ջ���=�ϴη�����-�ջغ�Ӧ����
                            
                            'Ӧ�ϴη���ĩ��ʱ��=��ǰ�ϴ�ִ��ʱ��
                            If str�ϴ�ʱ�� <> "" Then
                                'ֹͣ��ֹʱ���ڶ�η���֮��
                                strSQL = "Select Min(�״�ʱ��) as �״�ʱ��,Max(ĩ��ʱ��) as ĩ��ʱ��" & _
                                    " From ����ҽ������ Where ҽ��ID=[1] And [2]<=ĩ��ʱ��"
                                Set rsSend = zlDatabase.OpenSQLRecord(strSQL, "�����ջ�", Val(rsAdvice!ID), CDate(str�ϴ�ʱ��))
                            Else
                                strSQL = "Select �״�ʱ��,ĩ��ʱ�� From ����ҽ������ Where ҽ��ID=[1]" & _
                                    " And ���ͺ�=(Select Max(���ͺ�) From ����ҽ������ Where ҽ��ID=[1])"
                                Set rsSend = zlDatabase.OpenSQLRecord(strSQL, "�����ջ�", Val(rsAdvice!ID))
                            End If
                            
                            '�����ϴη��͵�����:���صķֽ�ʱ�����Ӧ���ҩ;����"��������"��ͬ
                            datBegin = Calc�����ڿ�ʼʱ��(rsAdvice!��ʼִ��ʱ��, rsSend!�״�ʱ��, rsAdvice!Ƶ�ʼ��, rsAdvice!�����λ)
                            str�ֽ�ʱ�� = Calc���ڷֽ�ʱ��(datBegin, rsSend!ĩ��ʱ��, strPause, NVL(rsAdvice!ִ��ʱ�䷽��), rsAdvice!Ƶ�ʴ���, rsAdvice!Ƶ�ʼ��, rsAdvice!�����λ, rsAdvice!��ʼִ��ʱ��)
                            dbl����All = Calc����ҩƷ����(rsAdvice!��ʼִ��ʱ��, 0, str�ֽ�ʱ��, _
                                    NVL(rsAdvice!��������, 0), rsDrug!����ϵ��, rsDrug!סԺ��װ, int�ɷ����, _
                                    CDate("3000-01-01"), strPause, NVL(rsAdvice!ִ��ʱ�䷽��), _
                                    rsAdvice!Ƶ�ʴ���, rsAdvice!Ƶ�ʼ��, rsAdvice!�����λ, mblnLimit, NVL(rsAdvice!�״�����, 0))
                            If str�ϴ�ʱ�� <> "" Then
                                str�ֽ�ʱ�� = Calc���ڷֽ�ʱ��(datBegin, CDate(str�ϴ�ʱ��), strPause, NVL(rsAdvice!ִ��ʱ�䷽��), rsAdvice!Ƶ�ʴ���, rsAdvice!Ƶ�ʼ��, rsAdvice!�����λ, rsAdvice!��ʼִ��ʱ��)
                                dbl���� = Calc����ҩƷ����(rsAdvice!��ʼִ��ʱ��, 0, str�ֽ�ʱ��, _
                                        NVL(rsAdvice!��������, 0), rsDrug!����ϵ��, rsDrug!סԺ��װ, int�ɷ����, _
                                        CDate("3000-01-01"), strPause, NVL(rsAdvice!ִ��ʱ�䷽��), _
                                        rsAdvice!Ƶ�ʴ���, rsAdvice!Ƶ�ʼ��, rsAdvice!�����λ, mblnLimit, NVL(rsAdvice!�״�����, 0))
                                dbl���� = dbl����All - dbl����
                            Else
                                'Ϊ�ձ�ʾȫ���ջص����
                                dbl���� = dbl����All
                            End If
                        End If
                        .TextMatrix(lngRow, COL_����) = FormatEx(dbl����, 5)
                                                
                        'ҩƷ������Ϣ
                        .TextMatrix(lngRow, COL_����ϵ��) = rsDrug!����ϵ��
                        .TextMatrix(lngRow, COL_סԺ��װ) = rsDrug!סԺ��װ
                        .TextMatrix(lngRow, COL_�ɷ����) = int�ɷ����
                        
                        '��ʾҩƷ��ҩ;��
                        .TextMatrix(lngRow, COL_�÷�) = "" & rsAdvice!��ҩ;��
                        
                        .ColHidden(COL_���) = gblnҩƷ�������ҽ��
                    Else
                        '��ҩҽ��
                        '------------------
                        .TextMatrix(lngRow, COL_����) = lng���� * NVL(rsAdvice!��������, 1)
                        If str�ϴ�ʱ�� = "" And NVL(rsAdvice!�״�����, 0) <> 0 Then
                            '����ϴ�ʱ��Ϊ��������״�
                            dbl���� = dbl���� + (NVL(rsAdvice!�״�����, 0) - NVL(rsAdvice!��������, 0))
                        End If
                        
                        '��ҩ�䷽��λ
                        If rsAdvice!������� = "E" And rsAdvice!ID = Val(.TextMatrix(lngRow - 1, COL_���ID)) _
                            And InStr(",E,7,", .TextMatrix(lngRow - 1, COL_�������)) > 0 Then
                            .TextMatrix(lngRow, COL_��λ) = "��"
                        End If
                    End If
                    
                    .TextMatrix(lngRow, COL_����) = NVL(rsAdvice!��������, 0)   '��ҩ�洢����Ƶ��
                    .TextMatrix(lngRow, COL_������) = .TextMatrix(lngRow, COL_����)
                    .TextMatrix(lngRow, COL_����) = lng����
                    .TextMatrix(lngRow, COL_�ϴ�) = str�ϴ�ʱ�� '����Ϊ��,��ȫ���ջص����
                    .Cell(flexcpData, lngRow, COL_����) = .TextMatrix(lngRow, COL_����) '��������ָ�
                    
                    '��������
                    '---------------------------------------------------------------
                    '���˼������ָ�
                    If rsAdvice!����ID <> lng����ID Then
                        lng������ = lng������ + 1
                        If lng����ID <> 0 Then
                            For j = lngRow - 1 To .FixedRows Step -1
                                If Not .RowHidden(j) Then
                                    .CellBorderRange j, .FixedCols, j, .Cols - 1, vbBlack, 0, 0, 0, 2, 0, 0
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                    lng����ID = rsAdvice!����ID

NextLoop:           '---------------------------------------------------------------
                    Progress = i / rsAdvice.RecordCount * 100
                    rsAdvice.MoveNext
                Next
            End With
        End If
    Next
    
    lblInfo.Caption = "����" & IIF(str���� = "", " ", "(" & Mid(str����, 2) & ") ") & lng������ & " �����˵�ҽ��"
    With vsAdvice
        .RowHeight(0) = 320
        If Not .ColHidden(COL_���) Then
            .AutoSize col_ҽ������, COL_���
        Else
            .AutoSize col_ҽ������
        End If
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        
        .Col = .FixedCols
        For i = .FixedRows To .Rows - 1
            If Not .RowHidden(i) Then
                .Row = i: Exit For
            End If
        Next
        Call .ShowCell(.Row, .Col)
        .Redraw = flexRDDirect
        
        'ֻ��һ��ʱ��ѡ��
        k = 0
        For i = .FixedRows To .Rows - 1
            If Not .RowHidden(i) Then k = k + 1
            If k > 1 Then Exit For
        Next
        If k = 1 Or mblnOnePati Then    '�����˵���ģʽ��ȫѡ
            If Val(.TextMatrix(.Rows - 1, COL_ID)) <> 0 Then Call tbr_ButtonClick(tbr.Buttons("ȫѡ"))
        End If
        If mblnAdjustNum Then
            .Cell(flexcpBackColor, .FixedRows, COL_����, .Rows - 1, COL_����) = COLEditBackColor       'ǳ��
        End If
    End With
    Progress = 0: Screen.MousePointer = 0
    LoadAdviceRoll = True
    Exit Function
errH:
    vsAdvice.Redraw = flexRDDirect
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        vsAdvice.Redraw = flexRDNone: Resume
    End If
    Call SaveErrLog
    lblInfo.Caption = "": Progress = 0
End Function

Private Function RowCanRoll(ByVal lngRow As Long, Optional strMsg As String) As Boolean
'���ܣ��ж�ָ�����Ƿ������ջ�(һ��ҽ��һ���ж�)
'������strMsg=���ز������ջص�ԭ����ʾ
    Dim lngBegin As Long, lngEnd As Long, i As Long
    
    strMsg = "": RowCanRoll = True
    
    With vsAdvice
        If mbln���ڸ��� Then
            If InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) > 0 Then
                If Val(.Cell(flexcpData, lngRow, COL_ҩƷID)) = 1 Then
                    strMsg = "���������ҩƷ����������ʽ���ʣ���ҽ�������ջء�"
                    RowCanRoll = False: Exit Function
                End If
            End If
        End If
        
        Call GetGroupRow(lngRow, lngBegin, lngEnd)
        For i = lngBegin To lngEnd
            If Val(.TextMatrix(i, COL_����)) <> 0 And Val(.TextMatrix(i, COL_����)) > 0 Then
                If mbln���ڸ��� Then
                    If Not gclsInsure.GetCapability(support��������, Val(.TextMatrix(i, COL_����ID)), Val(.TextMatrix(i, COL_����))) Then
                        strMsg = "��ҽ�������������಻��������ʽ���ʣ���ҽ�������ջء�"
                        RowCanRoll = False: Exit Function
                    End If
                Else
                    If Not gclsInsure.GetCapability(support�����ݳ�������, Val(.TextMatrix(i, COL_����ID)), Val(.TextMatrix(i, COL_����))) Then
                        strMsg = "��ҽ�������������಻�����ݳ������ã���ҽ�������ջء�"
                        RowCanRoll = False: Exit Function
                    End If
                End If
            End If
        Next
    End With
End Function

Private Sub GetCurBillSet(ByVal strKey As String, strNO As String)
'���ܣ���ȡ��ǰ���ʵ��ݵ�NO
    mrsBill.Filter = "Key='" & strKey & "'"
    If mrsBill.EOF Then
        mrsBill.AddNew
        mrsBill!Key = strKey
        mrsBill!NO = zlDatabase.GetNextNo(14)
        mrsBill.Update
    End If
    strNO = mrsBill!NO
End Sub

Public Function RollAdvice(ByVal lngCount As Long) As Boolean
'���ܣ�����ҽ������(��������м��ʱ���)
'������lngCount=��ѡ�������
'˵����������˷����ύ
    Dim arrSQL() As Variant
    Dim strSQL As String, strNOKey As String
    Dim curDate As Date, blnTran As Boolean
    Dim strNO As String, strTmp As String
    Dim i As Long, j As Long, k As Long
    Dim int�䷽�� As Integer, dbl���� As Double, dbl����ϵ�� As Double, dblסԺ��װ As Double
    Dim strҽ��IDs As String, str��ҩҽ��IDs As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    Screen.MousePointer = 11
    
    If mbln���ڸ��� Then Call InitBillSet
    curDate = zlDatabase.Currentdate
    arrSQL = Array()
    int�䷽�� = 1
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_ѡ��)) <> 0 And Val(.TextMatrix(i, COL_ִ������)) <> 0 Then '�ſ�����
                If mbln���ڸ��� Then
                    dbl����ϵ�� = Val(.TextMatrix(i, COL_����ϵ��))
                    If dbl����ϵ�� = 0 Then dbl����ϵ�� = 1
                    If InStr(",7,", .TextMatrix(i, COL_�������)) > 0 Then
                        '������ʾ�ĸ�������Ҫ���Ե�������ÿ��n����
                        dbl���� = Format(Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_����)) / dbl����ϵ��, "0.00000")
                    Else
                        dblסԺ��װ = Val(.TextMatrix(i, COL_סԺ��װ))
                        If dblסԺ��װ = 0 Then dblסԺ��װ = 1
                        '��ԭΪ�ۼ۵�λ������(���ͼ�¼���ۼ۵�λ)
                        dbl���� = Format(Val(.TextMatrix(i, COL_����)) * dblסԺ��װ * dbl����ϵ��, "0.00000")
                    End If
                    
                    If CheckAllPrice(Val(.TextMatrix(i, COL_ID)), dbl����, Val(.TextMatrix(i, COL_��������))) Then
                        strNO = "�������۵�"
                    Else
                        '�������ݺŷ���ؼ���:�뷢��ʱ�ķֺŹ�����ͬ
                        '-----------------------------------------------------------------------------------------
                        If InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0 Then
                            '������ҩ��"����(����ID,��ҳID)_���˿���ID_��������ID_����ҽ��_ִ�п���ID"�ֺš�
                            strNOKey = "������ҩ_" & Val(.TextMatrix(i, COL_����ID)) & "_" & Val(.TextMatrix(i, COL_��ҳID)) & "_" & _
                                Val(.TextMatrix(i, COL_���˿���ID)) & "_" & Val(.TextMatrix(i, COL_��������ID)) & "_" & _
                                .TextMatrix(i, COL_����ҽ��) & "_" & Val(.TextMatrix(i, COL_ִ�п���ID))
                        ElseIf InStr(",4,M,", .TextMatrix(i, COL_�������)) > 0 Then
                            '���ϰ�"����(����ID,��ҳID)_���˿���ID_��������ID_����ҽ��_ִ�п���ID"�ֺš�
                            strNOKey = "����ҽ��_" & Val(.TextMatrix(i, COL_����ID)) & "_" & Val(.TextMatrix(i, COL_��ҳID)) & "_" & _
                                Val(.TextMatrix(i, COL_���˿���ID)) & "_" & Val(.TextMatrix(i, COL_��������ID)) & "_" & _
                                .TextMatrix(i, COL_����ҽ��) & "_" & Val(.TextMatrix(i, COL_ִ�п���ID))
                        ElseIf .TextMatrix(i, COL_�������) = "7" Then
                            'һ���䷽�е����в�ҩ����һ���������ݺ�
                            strNOKey = "��ҩ�䷽_" & Val(.TextMatrix(i, COL_����ID)) & "_" & Val(.TextMatrix(i, COL_��ҳID)) & "_" & int�䷽��
                        ElseIf Val(.TextMatrix(i, COL_���ID)) <> 0 And .TextMatrix(i, COL_�������) = "C" Then
                            'һ���ɼ��ļ�����Ϸ�����ͬ�ĵ��ݺţ��걾�ɼ��������䵥���ĵ��ݺ�
                            strNOKey = "һ���ɼ�_" & Val(.TextMatrix(i, COL_���ID))
                        ElseIf Val(.TextMatrix(i, COL_���ID)) <> 0 And InStr(",F,D,", .TextMatrix(i, COL_�������)) > 0 Then
                            '��鲿λ�͸�����������Ҫҽ��������ͬ���ݺţ�����������䵥���ĵ��ݺš�
                            strNOKey = "��ҩҽ��_" & Val(.TextMatrix(i, COL_���ID))
                        Else
                            '������ҩҽ��ÿ��ҽ��һ���������ݺ�(������ҩ;�����䷽�巨���÷����ɼ���ʽ������ʽ����Ѫҽ��/��Ѫ;��)
                            strNOKey = "��ҩҽ��_" & Val(.TextMatrix(i, COL_ID))
                        End If
                            
                        Call GetCurBillSet(strNOKey, strNO)
                    End If
                End If
                '�ܲ�����Һ��¼����ҽ������ҩ��ʽΪ��Һ��ִ��ҩ��Ϊ��������
                If gstr��Һ�������� <> "" And Not mbln�������� Then
                    If .TextMatrix(i, COL_�������) = "E" And .TextMatrix(i, COL_��������) = "2" And .TextMatrix(i, COL_ִ�з���) = "1" Then
                        For j = i - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(i, COL_ID)) = Val(.TextMatrix(j, COL_���ID)) Then
                                If InStr("," & gstr��Һ�������� & ",", "," & Val(.TextMatrix(j, COL_ִ�п���ID)) & ",") > 0 Then
                                    strҽ��IDs = strҽ��IDs & "," & Val(.TextMatrix(i, COL_ID)): Exit For
                                End If
                            Else
                                Exit For
                            End If
                        Next
                    End If
                End If
                
                If InStr(",7,", .TextMatrix(i, COL_�������)) > 0 Then
                    str��ҩҽ��IDs = str��ҩҽ��IDs & "," & Val(.TextMatrix(i, COL_ID))
                End If
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                
                If .TextMatrix(i, COL_�ϴ�) = "" Then
                    strTmp = "NULL"
                Else
                    strTmp = "To_Date('" & .TextMatrix(i, COL_�ϴ�) & "','YYYY-MM-DD HH24:MI:SS')"
                End If
                
                arrSQL(UBound(arrSQL)) = _
                    IIF(Val(.TextMatrix(i, COL_ҩƷID)) = 0, "999999999", Val(.TextMatrix(i, COL_ҩƷID))) & ":" & _
                    "ZL_����ҽ����¼_�ջ�(" & Val(.TextMatrix(i, COL_����)) & "," & Val(.TextMatrix(i, COL_ID)) & "," & strTmp & "," & _
                    "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                    IIF(mbln���ڸ���, "'" & strNO & "'", "NULL") & ")"
                
                '������ҩ�䷽��
                If .TextMatrix(i, COL_�������) = "E" And Val(.TextMatrix(i, COL_ID)) = Val(.TextMatrix(i - 1, COL_���ID)) _
                    And InStr(",E,7,", .TextMatrix(i - 1, COL_�������)) > 0 Then '��ҩ�÷�
                    int�䷽�� = int�䷽�� + 1
                End If
                
                '---------------------------------
                k = k + 1
                Progress = k / (lngCount * 2) * 100
            End If
        Next
        
        If strҽ��IDs <> "" Then
            strҽ��IDs = Mid(strҽ��IDs, 2)
            If Drug��Һ(strҽ��IDs) Then
                If MsgBox("�����ջص���Һҽ����ҩƷ�а����Ѿ���Һ�ļ�¼���������ջأ��Ƿ�����ջ�����δ��Һ�ļ�¼��", vbQuestion + vbYesNo + vbDefaultButton2, "�����ջ�") = vbNo Then
                    Progress = 0: Screen.MousePointer = 0
                    RollAdvice = False
                    Exit Function
                End If
            End If
        End If
        
        If str��ҩҽ��IDs <> "" Then
            str��ҩҽ��IDs = Mid(str��ҩҽ��IDs, 2)
            strSQL = "select 1 from סԺ���ü�¼ a where a.��¼״̬ In (0, 1, 3) And Nvl(a.ִ��״̬,0)<>0" & _
                " and a.ҽ����� in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) and rownum<2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str��ҩҽ��IDs)
            If Not rsTmp.EOF Then
                If MsgBox("�����ջص��в�ҩ�д����Ѿ���ҩ�ģ�ֻ���ջ�δ��ҩ�Ĳ��֣��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, "�����ջ�") = vbNo Then
                    Progress = 0: Screen.MousePointer = 0
                    RollAdvice = False
                    Exit Function
                End If
            End If
        End If
        
        '��ҩƷID����(�����),��ҩ;������ҩ���ں���
        For i = 0 To UBound(arrSQL) - 1
            For j = i + 1 To UBound(arrSQL)
                If Val(Left(arrSQL(j), InStr(arrSQL(j), ":") - 1)) < Val(Left(arrSQL(i), InStr(arrSQL(i), ":") - 1)) Then
                    strTmp = arrSQL(j)
                    arrSQL(j) = arrSQL(i)
                    arrSQL(i) = strTmp
                End If
            Next
        Next
                
        '�ύ����
        gcnOracle.BeginTrans: blnTran = True
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(Mid(arrSQL(i), InStr(arrSQL(i), ":") + 1), "�����ջ�")
            '---------------------------------
            k = k + 1
            Progress = k / (lngCount * 2) * 100
        Next
        gcnOracle.CommitTrans: blnTran = False
        
        '�ύ�ɹ�,ɾ�����ջ���
        .Redraw = flexRDNone
        For i = .Rows - 1 To .FixedRows Step -1
            If Val(.TextMatrix(i, COL_ѡ��)) <> 0 Then
                .RemoveItem i
            End If
        Next
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        For i = .FixedRows To .Rows + 1
            If Not .RowHidden(i) Then
                .Row = i: Exit For
            End If
        Next
        .ShowCell .Row, .Col
        .Redraw = flexRDDirect
    End With
    Progress = 0: Screen.MousePointer = 0
    RollAdvice = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTran Then gcnOracle.RollbackTrans
    If err.Number <> 0 Then '��ҽ���ϴ�ʧ���˳�û�д���
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
    Progress = 0
End Function

Private Function Drug��Һ(ByVal strҽ��IDs As String) As Boolean
'���ܣ�ҽ�����Ƿ�����Ѿ�����Һ��
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    strSQL = "Select 1 from ����ҽ����¼ A,��Һ��ҩ��¼ B Where a.Id=b.ҽ��id And (b.����״̬ In (4,5,6,7,8) AND NVL(B.�Ƿ���,0) = 0) And b.ִ��ʱ��>a.ִ����ֹʱ��" & _
        " and a.Id In (Select Column_Value From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist))) And Rownum<2"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strҽ��IDs)
    Drug��Һ = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckAllPrice(ByVal lngҽ��ID As Long, ByVal dbl�ջ����� As Double, ByVal lng�������� As Long) As Boolean
'���ܣ�����ջش�����ҽ����Ӧ�����Ƿ�ȫ��δ��˵Ļ��۵����Ա�ȷ��ֱ���޸Ļ��۵�������ȡ�µĵ��ݺ�
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim lng�ջ��� As Long, lng������ As Long
    
    CheckAllPrice = False
    
    '���������ǰ��ۼ۵�λ�洢��
    strSQL = "Select Sum(a.��������) ��������" & vbNewLine & _
            "From ����ҽ������ A" & vbNewLine & _
            "Where a.ҽ��id = [1] And a.��¼���� = 2 And Not Exists" & vbNewLine & _
            " (Select 1 From " & IIF(lng�������� = 1, "����", "סԺ") & "���ü�¼ B Where a.ҽ��id = b.ҽ����� And a.No = b.No And b.��¼���� = 2 And ��¼״̬ <> 0)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
    If rsTmp.RecordCount > 0 Then
        If dbl�ջ����� <= Val("" & rsTmp!��������) Then CheckAllPrice = True
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsAdvice_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COL_���� And Not mblnReturn Then
        vsAdvice.Refresh    '����е�����ʾ����ˢ�µĻ���һ����ҩͨ��Drawcell�������ĵ�Ԫ����ٴ���ʾ
        If Not AcceptInput(Row, Col) Then
            Cancel = True
        End If
    End If
End Sub

Private Function CheckAdvcieComPound(ByVal lngҽ��ID As Long) As Boolean
'���ܣ�����ҽ��ID���ж��Ƿ�����Һ��ҩ��ҩƷ
    Dim strSQL As String, rsTmp As Recordset
    
    strSQL = "Select 1 from ��Һ��ҩ��¼ Where ҽ��ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
    CheckAdvcieComPound = rsTmp.RecordCount > 0
End Function
