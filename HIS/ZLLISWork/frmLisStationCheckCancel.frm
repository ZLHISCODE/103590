VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmLisStationCheckCancel 
   Caption         =   "ȡ������"
   ClientHeight    =   7155
   ClientLeft      =   2775
   ClientTop       =   4050
   ClientWidth     =   10920
   Icon            =   "frmLisStationCheckCancel.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   4125
      Left            =   150
      TabIndex        =   0
      Top             =   765
      Width           =   6825
      _cx             =   12039
      _cy             =   7276
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
      Begin VB.Line lnY 
         Index           =   0
         Visible         =   0   'False
         X1              =   270
         X2              =   270
         Y1              =   420
         Y2              =   1635
      End
      Begin VB.Line lnX 
         Index           =   0
         Visible         =   0   'False
         X1              =   -555
         X2              =   1230
         Y1              =   555
         Y2              =   555
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   6795
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmLisStationCheckCancel.frx":038A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14182
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
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   10920
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
         TabIndex        =   3
         Top             =   30
         Width           =   10800
         _ExtentX        =   19050
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
               Caption         =   "&S.����"
               Key             =   "ȡ��"
               Object.ToolTipText     =   "����(Alt+S)"
               Object.Tag             =   "&S.����"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&A.ȫѡ"
               Key             =   "ȫѡ"
               Object.ToolTipText     =   "ȫѡ(Alt+A)"
               Object.Tag             =   "&A.ȫѡ"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&C.ȫ��"
               Key             =   "ȫ��"
               Object.ToolTipText     =   "ȫ��(Alt+C)"
               Object.Tag             =   "&C.ȫ��"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&H.����"
               Key             =   "����"
               Object.ToolTipText     =   "����(Alt+H)"
               Object.Tag             =   "&H.����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&X.�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�(Alt+X)"
               Object.Tag             =   "&X.�˳�"
               ImageIndex      =   4
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
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationCheckCancel.frx":0C1E
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationCheckCancel.frx":1398
            Key             =   "ClearAll"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationCheckCancel.frx":1B12
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationCheckCancel.frx":1D32
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationCheckCancel.frx":1F52
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationCheckCancel.frx":226C
            Key             =   ""
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
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationCheckCancel.frx":2486
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationCheckCancel.frx":2C00
            Key             =   "ClearAll"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationCheckCancel.frx":337A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationCheckCancel.frx":359A
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationCheckCancel.frx":37BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationCheckCancel.frx":3AD4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileCancel 
         Caption         =   "ȡ��(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSelAll 
         Caption         =   "ȫѡ(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFileClsAll 
         Caption         =   "ȫ��(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
         Shortcut        =   ^X
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
Attribute VB_Name = "frmLisStationCheckCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean                          '����������־
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngKey As Long   '�걾ID
Private mblnChanged As Boolean
Private blnComm As Boolean '�Ƿ�����˫��ͨ��
Private miType As Integer '�걾���:0=��ͨ��1=����
Private objLISComm As Object
Private mblnReserveSample As Boolean, mblnShow As Boolean
Private mWinsockC As Winsock

Private Enum mCol
    ѡ�� = 0
    ҽ������
    ������
    ����ʱ��
    �������
End Enum

'�������Զ�����̻���************************************************************************************************
Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '����:
    'ֵ��:
    '------------------------------------------------------------------------------------------------------------------

    mnuFileCancel.Enabled = vData

    tbrThis.Buttons("ȡ��").Enabled = mnuFileCancel.Enabled


End Property

Public Function ShowEdit(ByVal frmMain As Object, ByRef lngKey As Long, WinsockC As Winsock, _
    Optional ByVal blnReserveSample As Boolean = False, Optional ByVal blnShow As Boolean = True) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʾ�༭���壬������ô���Ľӿں���
    '����:  frmMain         ���ô������
    '       lngKey          �걾id
    '       blnReserveSample�Ƿ����걾��תΪ������
    '       blnShow         �Ƿ���ʾѡ�����
    '����:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
    mblnReserveSample = blnReserveSample
    mblnShow = blnShow

    mlngKey = lngKey
    blnComm = Val(zldatabase.GetPara("��������˫��", 100, 1208, 0))

    Set mfrmMain = frmMain

    If InitData = False Then Exit Function
    If ReadData(mlngKey) = False Then Exit Function
    
    If Val(vsf.RowData(1)) < 0 Then Exit Function
    
    Set mWinsockC = WinsockC
    
    If vsf.Rows = 2 Or Not blnShow Then
        'ֻ����һ������,ֱ��ȡ��,����ѡ��
        Call mnuFileCancel_Click
        ShowEdit = mblnOK
        Exit Function
    End If
    
    EditChanged = (Val(vsf.RowData(1)) > 0)

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
    Dim strBill As String
    Dim strSQL As String
    
    On Error GoTo ErrHand
    
    strSQL = "Select Distinct 1 AS ѡ��,ҽ������,����ҽ�� As ������,����ʱ�� As ����ʱ��,B.����  As �������,A.ID,Nvl(C.�걾���,0) As �걾��� " & _
                "From ����ҽ����¼ A,���ű� B," & _
                "(Select ҽ��ID,MAX(�걾���) As �걾��� From (Select ҽ��id,Nvl(�걾���,0) As �걾��� From ����걾��¼ Where ID=[1] " & _
                "Union All Select Distinct ҽ��id,0 From ������Ŀ�ֲ� Where �걾ID = [1]) GROUP BY ҽ��ID) C " & _
                "Where C.ҽ��id=A.ID " & _
                    "And A.��������id=B.ID"
    Set rs = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)

    If rs.BOF = False Then
        miType = rs("�걾���")
        
        Call FillGrid(vsf, rs)
        Call AppendRows(vsf, lnX, lnY)
    End If
    
    ReadData = True

    Exit Function

ErrHand:
    If ErrCenter = 1 Then Resume

End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʼ������
    '����:  True        ��ʼ���ɹ�
    '       False       ��ʼ��ʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String

    On Error GoTo ErrHand

    strVsf = "ѡ��,450,1,1,1,;ҽ������,3000,1,1,1,;������,900,1,1,1,;����ʱ��,1800,1,1,1,;�������,1500,1,1,1,"

    Call CreateVsf(vsf, strVsf)
    vsf.Cols = vsf.Cols + 1
    vsf.ColWidth(vsf.Cols - 1) = 15
    vsf.ColDataType(0) = flexDTBoolean
    vsf.Editable = True

    Call AppendRows(vsf, lnX, lnY)

    InitData = True

    Exit Function

ErrHand:
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

Private Function SaveData() As Boolean

    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim strSQL() As String
    Dim rs As New ADODB.Recordset, strQrySQL As String
    Dim dtSendTime As Date '����ʱ��
    Dim strDevices As String, aDevice() As String, strAdviceIDs As String, i As Integer
    Dim intEmerge As Integer                '�Ƿ����ҽ��
    Dim lngBeginDate As Long

    intEmerge = Val(zldatabase.GetPara("����걾", 100, 1208, 0))
    

    On Error GoTo ErrHand

    ReDim strSQL(1 To 1)
    
    Me.MousePointer = vbHourglass
    strAdviceIDs = "": strDevices = ""
    For lngLoop = 1 To vsf.Rows - 1
        
        If Abs(Val(vsf.TextMatrix(lngLoop, mCol.ѡ��))) = 1 And Val(vsf.RowData(lngLoop)) > 0 Then
            If Not mblnReserveSample Then
                strSQL(ReDimArray(strSQL)) = "ZL_����걾��¼_ȡ������(" & Val(vsf.RowData(lngLoop)) & ")"
            
                '����˫��ͨ��
                If blnComm Then
                    strAdviceIDs = strAdviceIDs & "," & vsf.RowData(lngLoop)
                    
                    strQrySQL = "Select Distinct ����ID From ����걾��¼ A,������Ŀ�ֲ� B" & _
                        " Where B.ҽ��ID=[1] And B.�걾ID+0=A.ID"
                    Set rs = zldatabase.OpenSQLRecord(strQrySQL, Me.Caption, Val(vsf.RowData(lngLoop)))
                    Do While Not rs.EOF
                        If InStr(strDevices, "," & zlCommFun.Nvl(rs(0), 0)) = 0 Then
                            strDevices = strDevices & "," & zlCommFun.Nvl(rs(0), 0)
                        End If
                        
                        rs.MoveNext
                    Loop
                End If
            Else
                strSQL(ReDimArray(strSQL)) = "ZL_����걾��¼_תΪ����(" & Val(vsf.RowData(lngLoop)) & ")"
            End If
        End If
    Next
    '����˫��ͨ��
    If blnComm And Not mblnReserveSample Then
        If Len(strDevices) > 0 Then strDevices = Mid(strDevices, 2)
        If Len(strAdviceIDs) > 0 Then strAdviceIDs = Mid(strAdviceIDs, 2)
        
        aDevice = Split(strDevices, ",")
        For i = 0 To UBound(aDevice)
            SendSample mWinsockC, mWinsockC.LocalIP, CLng(Val(aDevice(i))), "", 0, strAdviceIDs, True, IIf(intEmerge = 1 And miType = 1, 1, 0)
        Next
        frmLabMain.mblnSendComplete = False
        lngBeginDate = Timer
        Do
            DoEvents
        Loop Until frmLabMain.mblnSendComplete = True Or (CLng(Timer) - lngBeginDate > 2)
    End If
    Me.MousePointer = vbDefault
    
    blnTran = True
    
    gcnOracle.BeginTrans
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zldatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    gcnOracle.CommitTrans
    blnTran = False
    
    SaveData = True
    
    Exit Function
    
ErrHand:
    
    Me.MousePointer = vbDefault
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
        
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        Select Case KeyCode
        Case vbKeyA
            If tbrThis.Buttons("ȫѡ").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("ȫѡ"))
        Case vbKeyC
            If tbrThis.Buttons("ȫ��").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("ȫ��"))
        Case vbKeyS
            If tbrThis.Buttons("ȡ��").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("ȡ��"))
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

'���������弰��ؼ����¼�����******************************************************************************************
Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    With vsf
        .Left = 30
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0) + 30
        .Width = Me.ScaleWidth - .Left - 30
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0) - 30
    End With

    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub


Private Sub mnuFileClearAll_Click()
    Dim lngLoop As Long

    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            vsf.TextMatrix(lngLoop, mCol.ѡ��) = 0
        End If
    Next

    EditChanged = False
End Sub

Private Sub mnuFileCancel_Click()
    '
    If mblnShow Then
        If MsgBox("���Ҫȡ����ǰ����ĺ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
        
    If SaveData = False Then Exit Sub
    
    mblnOK = True
    
    Unload Me
        
        
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileSelectAll_Click()
    Dim lngLoop As Long

    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            vsf.TextMatrix(lngLoop, mCol.ѡ��) = 1
            EditChanged = True
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
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call Form_Resize

End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "ȫѡ"
        Call mnuFileSelectAll_Click
    Case "ȫ��"
        Call mnuFileClearAll_Click
    Case "ȡ��"
        Call mnuFileCancel_Click
    Case "����"
        Call mnuHelpTopic_Click
    Case "�˳�"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    Call tbrThis_ButtonClick(Button)
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngLoop As Long

    If Abs(Val(vsf.TextMatrix(Row, mCol.ѡ��))) = 1 Then
        EditChanged = True
        Exit Sub
    End If

    For lngLoop = 1 To vsf.Rows - 1
        If Abs(Val(vsf.TextMatrix(lngLoop, mCol.ѡ��))) = 1 Then
            EditChanged = True
            Exit Sub
        End If
    Next

    If lngLoop = vsf.Rows Then EditChanged = False

End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> mCol.ѡ�� Or Val(vsf.RowData(Row)) <= 0 Then
        Cancel = True
    End If
End Sub


