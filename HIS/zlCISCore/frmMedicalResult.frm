VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmMedicalResult 
   Caption         =   "���������"
   ClientHeight    =   5700
   ClientLeft      =   2775
   ClientTop       =   4050
   ClientWidth     =   9180
   Icon            =   "frmMedicalResult.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   9180
   Begin VB.PictureBox picTitle 
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   9165
      TabIndex        =   1
      Top             =   735
      Width           =   9165
      Begin VB.CommandButton cmdSearch 
         Caption         =   "��ʼ����(&F)"
         Height          =   350
         Left            =   7785
         TabIndex        =   2
         Top             =   90
         Width           =   1320
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "����:���� �Ա�:�� ����:60 ����״��:�ѻ�"
         Height          =   180
         Left            =   60
         TabIndex        =   3
         Top             =   195
         Width           =   3510
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   5340
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMedicalResult.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11113
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
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
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   2205
      Left            =   345
      TabIndex        =   4
      Top             =   1605
      Width           =   4920
      _cx             =   8678
      _cy             =   3889
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
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
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
      Begin VB.Line lnX 
         Index           =   1
         Visible         =   0   'False
         X1              =   -555
         X2              =   1230
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Line lnY 
         Index           =   1
         Visible         =   0   'False
         X1              =   270
         X2              =   270
         Y1              =   420
         Y2              =   1635
      End
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9180
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   9060
         _ExtentX        =   15981
         _ExtentY        =   1270
         ButtonWidth     =   1402
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsHotMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&A.ȫѡ"
               Key             =   "ȫѡ"
               Object.ToolTipText     =   "ȫѡ(Alt+A)"
               Object.Tag             =   "&A.ȫѡ"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&C.ȫ��"
               Key             =   "ȫ��"
               Object.ToolTipText     =   "ȫ��(Alt+C)"
               Object.Tag             =   "&C.ȫ��"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&S.����"
               Key             =   "����"
               Object.ToolTipText     =   "����(Alt+S)"
               Object.Tag             =   "&S.����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&H.����"
               Key             =   "����"
               Object.ToolTipText     =   "����(Alt+H)"
               Object.Tag             =   "&H.����"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&X.�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�(Alt+X)"
               Object.Tag             =   "&X.�˳�"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   7950
      Top             =   1050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalResult.frx":0E1E
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalResult.frx":1598
            Key             =   "ClearAll"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalResult.frx":1D12
            Key             =   "save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalResult.frx":1F2C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalResult.frx":214C
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   8625
      Top             =   1050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalResult.frx":236C
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalResult.frx":2AE6
            Key             =   "ClearAll"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalResult.frx":3260
            Key             =   "save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalResult.frx":347A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalResult.frx":369A
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileSelAll 
         Caption         =   "ȫѡ(&A)"
      End
      Begin VB.Menu mnuFileClsAll 
         Caption         =   "ȫ��(&C)"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "����(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "��������(&T)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&Web�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmMedicalResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean                          '����������־
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngKey As Long
Private mblnChanged As Boolean

Private mlng����id As Long
Private mlngҽ��id As Long
Private mstr�Һŵ� As String

'�������Զ�����̻���************************************************************************************************

Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '����:
    'ֵ��:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngSvrKey As Long
    
    mnuFileSave.Enabled = True
       
    
    If vData = False Then
        mnuFileSave.Enabled = False
    
    End If
    
    tbrThis.Buttons("����").Enabled = mnuFileSave.Enabled
            
End Property

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    
    vsf.Rows = 2
    vsf.RowData(1) = 0
    vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
    Call AppendRows(vsf, lnX, lnY)
        
End Function

Public Function ShowEdit(ByVal frmMain As Object, ByVal strParam As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʾ�༭���壬������ô���Ľӿں���
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
                
    '����id,ҽ��id,�Һŵ�
    
    mlng����id = Val(Split(strParam, "'")(0))
    mlngҽ��id = Val(Split(strParam, "'")(1))
    mstr�Һŵ� = Split(strParam, "'")(2)
    
    Set mfrmMain = frmMain
    
    If InitData = False Then Exit Function
    
    If mlng����id > 0 Then Call ReadData(mlng����id)
    
    EditChanged = False
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function ReadData(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ȡ����
    '����:  lngKey      ����������
    '����:  True        ��ȡ�ɹ�
    '       False       ��ȡʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    
    '��ȡ������Ϣ
    strSQL = "SELECT * FROM ������Ϣ WHERE ����id=" & lngKey
    Call OpenRecord(rs, strSQL, Me.Caption)
    If rs.BOF = False Then
        lblInfo.Caption = "����:" & zlCommFun.NVL(rs("����")) & " �Ա�:" & zlCommFun.NVL(rs("�Ա�")) & " ����:" & zlCommFun.NVL(rs("����")) & " ����״��:" & zlCommFun.NVL(rs("����״��"))
    End If
                                
    Call AppendRows(vsf, lnX, lnY)
    
    ReadData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʼ������
    '����:  True        ��ʼ���ɹ�
    '       False       ��ʼ��ʧ��
    '------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHand
    
    vsf.Cols = 3
    vsf.ColWidth(0) = 255
    vsf.ExtendLastCol = True
    
    vsf.TextMatrix(0, 0) = ""
    vsf.TextMatrix(0, 1) = "�������"
    vsf.TextMatrix(0, 2) = "�쳣���"
    
    vsf.ColWidth(1) = 2400
    vsf.Editable = flexEDKbdMouse
    vsf.ColDataType(0) = flexDTBoolean
    
    Call AppendRows(vsf, lnX, lnY)
    
    InitData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ValidEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  У�����ݵ���Ч��
    '����:  True        ������Ч
    '       False       ������Ч
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
        
    ValidEdit = True
    
End Function

Private Function SaveEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��������
    '����:  True        ����ɹ�
    '       False       ����ʧ��
    '------------------------------------------------------------------------------------------------------------------
    
    Dim lngLoop As Long
    
    
    On Error GoTo errHand
    
    
    
    SaveEdit = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then Resume
    
End Function

Private Function AppendRows(ByVal objVsf As Object, ByRef objLineX As Variant, ByRef objLineY As Variant, Optional ByVal lngHideRows As Long = 0) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '����:������ؼ��Ŀ���
    '����:objVsf Ҫ�����еı��ؼ�����
    '����:���ɹ�����True,���򷵻� False
    '--------------------------------------------------------------------------------------------------------
    Dim lngTop As Long
    Dim lngLoop As Long
    Dim lngIndex As Long
    Dim lngLastRow As Long
    
    On Error GoTo errHand
    
    If objVsf.Rows = 0 Then Exit Function
    
    For lngLoop = objVsf.Rows - 1 To 1 Step -1
        If objVsf.RowHidden(lngLoop) = False Then
            lngLastRow = lngLoop
            Exit For
        End If
    Next
    
    lngTop = objVsf.Cell(flexcpTop, lngLastRow, 0) + objVsf.RowHeight(lngLastRow)
    
    '1.�������е���
    For lngLoop = 1 To objLineX.UBound
        objLineX(lngLoop).Visible = False
    Next
    
    For lngLoop = 1 To objLineY.UBound
        objLineY(lngLoop).Visible = False
    Next
    
    '2.���¼�����Ҫ������
    For lngLoop = 1 To objVsf.Cols - 1

        If objLineY.UBound < lngLoop Then Load objLineY(lngLoop)

        With objLineY(lngLoop)

            .ZOrder

            .X1 = objVsf.Cell(flexcpLeft, 0, lngLoop) - 15
            .X2 = .X1
            .Y1 = lngTop
            .Y2 = objVsf.Height

            .BorderColor = objVsf.GridColor

            .Visible = True
        End With

    Next

    '3.���¼�����Ҫ�ĺ���
    lngIndex = 0
    Do While (lngTop + objVsf.RowHeight(0)) < objVsf.Height

        lngIndex = lngIndex + 1
        If objLineX.UBound < lngIndex Then Load objLineX(lngIndex)

        With objLineX(lngIndex)

            .ZOrder

            .X1 = 0
            .X2 = objVsf.Width
            .Y1 = lngTop + objVsf.RowHeight(0) + 15
            .Y2 = .Y1

            .BorderColor = objVsf.GridColor

            .Visible = True

            lngTop = .Y1
        End With

    Loop
        
    AppendRows = True
    
    Exit Function
    
errHand:
    
End Function

Private Sub cmdSearch_Click()
    Dim lngLoop As Long
    Dim strTmp As String
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    Dim blnFound As Boolean
    Dim bytSave As Byte
    Dim strResult As String
    Dim str�Ա� As String
    Dim str���� As String
    Dim strValue As String
    Dim strWorn As String
    Dim strRefence As String
    
    '��ʼ����������
    
    cmdSearch.Enabled = False
    
    bytSave = vsf.Redraw
    vsf.Redraw = flexRDNone
    
    vsf.Rows = 2
    vsf.RowData(1) = 0
    vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
        
    stbThis.Panels(2).Text = "����ȷ��ɸ�鷶Χ..."
    DoEvents
    
    '��ȡ�������������Ŀ
    strSQL = _
            "SELECT A.����,A.���,B.������,A.�Ƿ񼲲� " & _
            "FROM   �����Ͻ��� A, " & _
                    "���������� B " & _
            "Where A.��� = B.������ " & _
            "GROUP BY A.����,A.���,B.������,A.�Ƿ񼲲�"
            
    Call OpenRecord(rs, strSQL, Me.Caption)
    If rs.BOF = False Then
    
        strSQL = "Select �Ա�,���� From ������Ϣ Where ����id=[1]"
        Set rs2 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����id)
        If rs2.BOF Then Exit Sub
        
        str�Ա� = zlCommFun.NVL(rs2("�Ա�").Value)
        str���� = zlCommFun.NVL(rs2("����").Value)
'        If str���� = "" Then str���� = zlCommFun.NVL(rs2("ʵ������").Value)
        
        
        Do While Not rs.EOF
            strResult = ""
            
            stbThis.Panels(2).Text = "���ڽ��С�" & zlCommFun.NVL(rs("����").Value) & "������..."
            DoEvents
            
            '��ȡ�ж�����
            strSQL = _
                    "SELECT B.ID,B.����,B.�滻��,B.������,A.��ϵʽ,A.����ֵ,A.�Ա�,A.��ʼ����,A.�������� " & _
                    "FROM ���������� A, " & _
                         "����������Ŀ B " & _
                    "Where A.��ĿID = B.ID " & _
                          "AND A.������=" & rs("���").Value & " " & _
                          IIf(zlCommFun.NVL(rs("������")) = "", "AND A.������ IS NULL", "AND A.������='" & zlCommFun.NVL(rs("������")) & "'")
                          
            Call OpenRecord(rs2, strSQL, Me.Caption)
            If rs2.BOF = False Then
                
                '
                blnFound = True
                                
                Do While Not rs2.EOF
                    
                    strTmp = ""
                    
                    If zlCommFun.NVL(rs2("�滻��").Value, 0) = 0 Then
                        '�����滻��
                        
                        '��ȡ����
                        strSQL = "SELECT S.����,S.��������,Y.���� " & _
                                " FROM ���˲��������� S,���˲������� X,����Ԫ��Ŀ¼ Y " & _
                                " WHERE X.ID=S.����id And Y.����(+)=X.Ԫ�ر��� And S.������ID+0=" & rs2("ID").Value & _
                                "       AND S.����ID=(" & _
                                "           SELECT MAX(S.����ID)" & _
                                "           from (SELECT S.����ID FROM ���˲��������� S WHERE S.������ID+0=" & rs2("ID").Value & ") S," & _
                                "                (SELECT C.ID,C.������¼ID " & _
                                "                 FROM ���˲�����¼ L,���˲������� C" & _
                                "                 WHERE L.ID=C.������¼ID" & _
                                "                       AND L.����ID=" & mlng����id
                                
                        
                        If mlngҽ��id > 0 Then
                            strSQL = strSQL & "                       AND L.ID IN (SELECT DISTINCT ����id FROM ����ҽ������ WHERE ҽ��id IN (SELECT ID FROM ����ҽ����¼ WHERE ������Դ=4 AND ID=" & mlngҽ��id & " UNION ALL SELECT ID FROM ����ҽ����¼ WHERE ������Դ=4 AND ���id=" & mlngҽ��id & "))"
                        Else
                            strSQL = strSQL & "                       AND L.ID IN (SELECT DISTINCT ����id FROM ����ҽ������ WHERE ҽ��id IN (SELECT ID FROM ����ҽ����¼ WHERE ������Դ=4 AND �Һŵ�='" & mstr�Һŵ� & "' and ����id=" & mlng����id & "))"
                        End If
                        
                        'strSQL = strSQL & "                       AND L.ID IN (SELECT DISTINCT ����id FROM ����ҽ������ WHERE ҽ��id IN (SELECT ID FROM ����ҽ����¼ WHERE ������Դ=4 AND �Һŵ�='" & mstr�Һŵ� & "' and ����id=" & mlng����id & "))"
                        
                        strSQL = strSQL & "       ) C" & _
                            "           WHERE C.ID=S.����ID)"
                            
                        Call OpenRecord(rs3, strSQL, Me.Caption)
                        If rs3.BOF = False Then
                            strTmp = zlCommFun.NVL(rs3("��������").Value, "")
                        Else
                            blnFound = False
                            Exit Do
                        End If
                        
                    Else
                        '���滻��
                        
                        strTmp = GetSpecValue(rs2("������").Value, CStr(mlng����id), "0", 0)
                        
                    End If
                    
                    
                    '���ò��������ж�
                    If zlCommFun.NVL(rs2("�Ա�").Value) <> "" Then
                                                
                        If InStr(str�Ա�, zlCommFun.NVL(rs2("�Ա�").Value)) = 0 Then
                            '������
                            blnFound = False
                            Exit Do
                        End If
                        
                    End If
                    
                    If zlCommFun.NVL(rs2("��ʼ����").Value) <> "" Or zlCommFun.NVL(rs2("��������").Value) <> "" Then
                        '
                        If zlVerifyAge(str����, zlCommFun.NVL(rs2("��ʼ����").Value), zlCommFun.NVL(rs2("��ʼ����").Value)) = False Then
                            blnFound = False
                            Exit Do
                        End If
                        
                    End If
                        
                    strValue = strTmp
                    strWorn = ""
                    strRefence = ""
                    
                    If UCase(zlCommFun.NVL(rs3("����").Value, "")) = "ZL9CISCORE.USRVERIFYREPORT" Then
                        If strTmp <> "" Then
                            
                            strValue = Split(strTmp, "'")(0)
                            strWorn = Split(strTmp, "'")(1)
                            strRefence = Split(strTmp, "'")(2)
                            
                            strTmp = Split(strTmp, "'")(0) & "(" & Split(strTmp, "'")(2) & ")"
                        End If
                    End If
                    
                    '���������ж�
                    If Not zlVerifyValue(strValue, zlCommFun.NVL(rs2("����"), 0), zlCommFun.NVL(rs2("��ϵʽ")), zlCommFun.NVL(rs2("����ֵ")), strWorn, strRefence) Then
                        blnFound = False
                        Exit Do
                    End If
                    
                    If strTmp <> "" Then
                        If strResult <> zlCommFun.NVL(rs3("����").Value, "") & ":" & strTmp Then
                            strResult = strResult & zlCommFun.NVL(rs3("����").Value, "") & ":" & strTmp
                        End If
                    End If
                    
                    
                    rs2.MoveNext
                Loop
                
                If blnFound Then
                    
                    '������������
                            
                    '�ж��Ƿ��Ѿ�����
                    For lngLoop = 1 To vsf.Rows - 1
                        If Val(vsf.RowData(lngLoop)) = rs("���").Value Then
                            Exit For
                        End If
                    Next
                                                
                    If lngLoop >= vsf.Rows Then
                        
                        'û�о�������
                        
                        If Val(vsf.RowData(vsf.Rows - 1)) > 0 Then
                            vsf.Rows = vsf.Rows + 1
                        End If
                        
                        vsf.RowData(vsf.Rows - 1) = rs("���").Value
                        vsf.TextMatrix(vsf.Rows - 1, 0) = "1"
                        vsf.TextMatrix(vsf.Rows - 1, 1) = zlCommFun.NVL(rs("����").Value)
                        vsf.TextMatrix(vsf.Rows - 1, 2) = strResult
                        vsf.Cell(flexcpData, vsf.Rows - 1, 1) = zlCommFun.NVL(rs("�Ƿ񼲲�").Value, 0)
                                                
                    End If
                    
                End If
                
            End If
            
            rs.MoveNext
        Loop
    End If
    
    If Val(vsf.RowData(1)) = 0 Then
        stbThis.Panels(2).Text = "û���ҵ����������"
        
        EditChanged = False
    Else
        stbThis.Panels(2).Text = "���ҵ� " & vsf.Rows - 1 & " �����������"
        
        EditChanged = True
    End If
    
    vsf.Redraw = bytSave
    
    cmdSearch.Enabled = True
    AppendRows vsf, lnX, lnY
    
End Sub

Private Function zlVerifyAge(ByVal str���� As String, ByVal str��ʼ���� As String, ByVal str�������� As String) As Boolean
    
    Dim strAgeNumber As String
    Dim strAgeNumberBegin As String
    Dim strAgeNumberEnd As String
    Dim strAgeUnit As String
    
    On Error GoTo errHand
    
    If str��ʼ���� = "" And str�������� = "" Then
        zlVerifyAge = True
        Exit Function
    End If
    
    If str��ʼ���� = "" And str�������� <> "" Then str��ʼ���� = str��������
    If str��ʼ���� <> "" And str�������� = "" Then str�������� = str��ʼ����
        
    Call AnalyseAge(str��ʼ����, strAgeNumberBegin, strAgeUnit)
    Select Case strAgeUnit
    Case "��"
        strAgeNumberBegin = Val(strAgeNumberBegin) * 30
    Case "��"
        strAgeNumberBegin = Val(strAgeNumberBegin) * 365
    End Select
    
    Call AnalyseAge(str��������, strAgeNumberEnd, strAgeUnit)
    Select Case strAgeUnit
    Case "��"
        strAgeNumberEnd = Val(strAgeNumberEnd) * 30
    Case "��"
        strAgeNumberEnd = Val(strAgeNumberEnd) * 365
    End Select
    
    Call AnalyseAge(str����, strAgeNumber, strAgeUnit)
    Select Case strAgeUnit
    Case "��"
        strAgeNumber = Val(strAgeNumber) * 30
    Case "��"
        strAgeNumber = Val(strAgeNumber) * 365
    End Select
        
    If Val(strAgeNumber) >= Val(strAgeNumberBegin) And Val(strAgeNumber) <= Val(strAgeNumberEnd) Then
        zlVerifyAge = True
        Exit Function
    End If
    
    zlVerifyAge = False
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function AnalyseAge(strOld As String, ByRef strAgeNumber As String, ByRef strAgeUnit As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    '����:�����ݿ��б�������䰴���Ƶĸ�ʽ���ص�����
    
    Dim strTmp As Long
    
    If strOld = "��" Then Exit Function
    
    If InStr(strOld, "��") > 0 Then
        If InStr(strOld, "��") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "��") - 1)
            strAgeNumber = strTmp
            strAgeUnit = "��"
        Else
            strAgeNumber = strOld
            strAgeUnit = ""
        End If
    ElseIf InStr(strOld, "��") > 0 Then
        If InStr(strOld, "��") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "��") - 1)
            strAgeNumber = strTmp
            strAgeUnit = "��"
        Else
            strAgeNumber = strOld
            strAgeUnit = ""
        End If
    ElseIf InStr(strOld, "��") > 0 Then
        If InStr(strOld, "��") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "��") - 1)
            strAgeNumber = strTmp
            strAgeUnit = "��"
        Else
            strAgeNumber = strOld
            strAgeUnit = ""
        End If
    ElseIf IsNumeric(strOld) Then
        strAgeNumber = strOld
        strAgeUnit = "��"
    Else
        strAgeNumber = strOld
        strAgeUnit = ""
    End If
    
    AnalyseAge = True
    
End Function

Private Function zlVerifyValue(strVerify As String, bytType As Byte, ByVal strFormula As String, ByVal strAskValue As String, ByVal strWorn As String, ByVal str������� As String) As Boolean
    '-------------------------------------------------
    '���ܣ��жϵ�ǰ�����Ƿ������������ʽ
    '��Σ� strVerify-���жϵ���ֵ
    '       bytType-��ֵ����
    '       strFormula-��ϵʽ������˵����
    '       strAskValue-Ҫ�����ֵ��Χ��
    '���Σ���ȷ����true�����򷵻�false
    '-------------------------------------------------
    Dim aryTemp() As String
    Dim varTmp As Variant
    
    zlVerifyValue = False
    
    
    Select Case strAskValue
    Case "[���ֵ]"
    
        If InStr(str�������, "��") > 0 Then
            varTmp = Split(str�������, "��")
            strAskValue = Val(varTmp(0))
        End If
        
    Case "[���ֵ]"
        
        If InStr(str�������, "��") > 0 Then
            varTmp = Split(str�������, "��")
            strAskValue = Val(varTmp(1))
        End If
        
    Case "[ƫ��]"
        
        If Trim(strWorn) = "ƫ��" Then zlVerifyValue = True
        Exit Function
        
    Case "[ƫ��]"
        
        If Trim(strWorn) = "ƫ��" Then zlVerifyValue = True
        Exit Function
        
    Case "[�쳣]"
        
        If Trim(strWorn) = "�쳣" Then zlVerifyValue = True
        Exit Function
        
    End Select
    
    Select Case Val(bytType)
    Case 0  '��ֵ
        Select Case Trim(strFormula)
        Case "����"
            If Val(strVerify) = Val(strAskValue) Then zlVerifyValue = True
        Case "������"
            If Val(strVerify) <> Val(strAskValue) Then zlVerifyValue = True
        Case "����"
            If Val(strVerify) > Val(strAskValue) Then zlVerifyValue = True
        Case "С��"
            If Val(strVerify) < Val(strAskValue) Then zlVerifyValue = True
        Case "С�ڵ���"
            If Val(strVerify) <= Val(strAskValue) Then zlVerifyValue = True
        Case "���ڵ���"
            If Val(strVerify) >= Val(strAskValue) Then zlVerifyValue = True
        Case "����", "�ڷ�Χ��"
            aryTemp = Split(strAskValue, "��")
            If UBound(aryTemp) = 1 Then
                aryTemp(0) = Trim(aryTemp(0))
                aryTemp(1) = Trim(aryTemp(1))
                
                If Val(strVerify) >= Val(aryTemp(0)) And Val(strVerify) <= Val(aryTemp(1)) Then zlVerifyValue = True
                If Val(strVerify) >= Val(aryTemp(1)) And Val(strVerify) <= Val(aryTemp(0)) Then zlVerifyValue = True
            End If
'        Case "����"
'            strAskValue = Replace(strAskValue, Space(1), "")
'            If InStr(1, "," & strAskValue & ",", "," & Val(strVerify) & ",") > 0 Then zlVerifyValue = True
'        Case "������"
'            strAskValue = Replace(strAskValue, Space(1), "")
'            If InStr(1, "," & strAskValue & ",", "," & Val(strVerify) & ",") = 0 Then zlVerifyValue = True
        End Select
    Case 1  '����
        Select Case Trim(strFormula)
        Case "����"
            If Trim(strVerify) = Trim(strAskValue) Then zlVerifyValue = True
        
        Case "����"
            
            If Trim(strVerify) > Trim(strAskValue) Then zlVerifyValue = True
            
        Case "С��"
            
            If Trim(strVerify) < Trim(strAskValue) Then zlVerifyValue = True
            
        Case "���ڵ���"
            
            If Trim(strVerify) >= Trim(strAskValue) Then zlVerifyValue = True
            
        Case "С�ڵ���"
            
            If Trim(strVerify) <= Trim(strAskValue) Then zlVerifyValue = True
            
        Case "������"
            If Trim(strVerify) <> Trim(strAskValue) Then zlVerifyValue = True
            
        Case "����"
            If InStr(1, Trim(strVerify), Trim(strAskValue)) > 0 Then zlVerifyValue = True
            
        Case "������"
            If InStr(1, Trim(strVerify), Trim(strAskValue)) = 0 Then zlVerifyValue = True
'        Case "����"
'            strAskValue = Replace(strAskValue, Space(1), "")
'            If InStr(1, "," & strAskValue & ",", "," & Trim(strVerify) & ",") > 0 Then zlVerifyValue = True
'        Case "������"
'            strAskValue = Replace(strAskValue, Space(1), "")
'            If InStr(1, "," & strAskValue & ",", "," & Trim(strVerify) & ",") = 0 Then zlVerifyValue = True
        End Select
'    Case 2  '����
'        strVerify = Format(strVerify, "YYYY-MM-DD")
'        Select Case Trim(strFormula)
'        Case "����"
'            strAskValue = Format(strAskValue, "YYYY-MM-DD")
'            If Trim(strVerify) = Trim(strAskValue) Then zlVerifyValue = True
'        Case "������"
'            strAskValue = Format(strAskValue, "YYYY-MM-DD")
'            If Trim(strVerify) <> Trim(strAskValue) Then zlVerifyValue = True
'        Case "����"
'            strAskValue = Format(strAskValue, "YYYY-MM-DD")
'            If Trim(strVerify) > Trim(strAskValue) Then zlVerifyValue = True
'        Case "����"
'            strAskValue = Format(strAskValue, "YYYY-MM-DD")
'            If Trim(strVerify) < Trim(strAskValue) Then zlVerifyValue = True
'        Case "������"
'            strAskValue = Format(strAskValue, "YYYY-MM-DD")
'            If Trim(strVerify) <= Trim(strAskValue) Then zlVerifyValue = True
'        Case "������"
'            strAskValue = Format(strAskValue, "YYYY-MM-DD")
'            If Trim(strVerify) >= Trim(strAskValue) Then zlVerifyValue = True
'        Case "����", "�ڷ�Χ��"
'            aryTemp = Split(strAskValue, "��")
'            If UBound(aryTemp) = 1 Then
'                aryTemp(0) = Format(Trim(aryTemp(0)), "YYYY-MM-DD")
'                aryTemp(1) = Format(Trim(aryTemp(1)), "YYYY-MM-DD")
'                If Trim(strVerify) >= Trim(aryTemp(0)) And Trim(strVerify) <= Trim(aryTemp(1)) Then zlVerifyValue = True
'                If Trim(strVerify) >= Trim(aryTemp(1)) And Trim(strVerify) <= Trim(aryTemp(0)) Then zlVerifyValue = True
'            End If
'        End Select
    Case 2  '�߼�
        Select Case Trim(strFormula)
        Case "����"
            If Val(strVerify) = Trim(strAskValue) Then zlVerifyValue = True
        Case "������"
            If Val(strVerify) <> Trim(strAskValue) Then zlVerifyValue = True
        End Select
    Case Else
    End Select
End Function


'���������弰��ؼ����¼�����******************************************************************************************

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        Select Case KeyCode
        Case vbKeyA
            If tbrThis.Buttons("ȫѡ").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("ȫѡ"))
        Case vbKeyC
            If tbrThis.Buttons("ȫ��").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("ȫ��"))
        Case vbKeyS
            If tbrThis.Buttons("����").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("����"))
        Case vbKeyH
            If tbrThis.Buttons("����").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("����"))
        Case vbKeyX
            If tbrThis.Buttons("�˳�").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("�˳�"))
        End Select
    ElseIf Shift = 0 Then
        If KeyCode = vbKeyEscape Then
            If tbrThis.Buttons("�˳�").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("�˳�"))
        End If
    End If
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
        
    With picTitle
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = Me.ScaleWidth
    End With
    
    With vsf
        .Left = 0
        .Top = picTitle.Top + picTitle.Height + 30
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
        
    cmdSearch.Left = picTitle.Width - cmdSearch.Width - 60
    
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If mnuFileSave.Enabled Then
        Cancel = (MsgBox("�������޸ĵ����ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
        If Cancel Then Exit Sub
    End If
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub mnuFileClear_Click()
    If MsgBox("ȷʵҪ��������õ���Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    vsf.Rows = 2
    vsf.RowData(1) = 0
    vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
    
    EditChanged = True
    
End Sub

Private Sub mnuFileClsAll_Click()
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            If vsf.TextMatrix(lngLoop, 0) <> "0" Then
                vsf.TextMatrix(lngLoop, 0) = "0"
                EditChanged = True
            End If
        End If
    Next
    
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileRestore_Click()
    
    If MsgBox("ȷʵҪ�ָ���ǰ��ѡ��Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    Call ClearData
    Call ReadData(mlngKey)
    EditChanged = False
    
End Sub

Private Sub mnuFileSave_Click()
    
    If ValidEdit = False Then Exit Sub
    
    If SaveEdit Then
        
        On Error Resume Next
        
        Call mfrmMain.EditRefresh(vsf)
        
        On Error GoTo 0
        
        EditChanged = False
        
        Unload Me
        
    End If
    
End Sub

Private Sub mnuFileSelAll_Click()
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            If vsf.TextMatrix(lngLoop, 0) = "0" Then
                vsf.TextMatrix(lngLoop, 0) = "1"
                EditChanged = True
            End If
        End If
    Next
    
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpTopic_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intLoop As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For intLoop = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(intLoop).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(intLoop).Tag, "")
    Next
    cbrThis.Bands(1).MINHEIGHT = tbrThis.Height
    Call Form_Resize
    
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "����"
        Call mnuFileSave_Click
    Case "ȫѡ"
        Call mnuFileSelAll_Click
    Case "ȫ��"
        Call mnuFileClsAll_Click
    Case "����"
        Call mnuHelpTopic_Click
    Case "�˳�"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    vsf.TextMatrix(Row, Col) = Abs(vsf.Value)
    EditChanged = True
    
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col = 0)
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    If Col <> 0 Then Cancel = True
    If Val(vsf.RowData(Row)) = 0 Then Cancel = True
    
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

