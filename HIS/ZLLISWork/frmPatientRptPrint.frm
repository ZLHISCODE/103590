VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmPatientRptPrint 
   Caption         =   "�������鱨���ӡ"
   ClientHeight    =   5730
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPatientRptPrint.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5730
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CheckBox chkPrinted 
      Caption         =   "�����Ѵ�ӡ��¼"
      Height          =   240
      Left            =   75
      TabIndex        =   8
      Top             =   5325
      Width           =   8565
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgData 
      Height          =   3855
      Left            =   510
      TabIndex        =   7
      Top             =   1950
      Width           =   8430
      _cx             =   14870
      _cy             =   6800
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   0
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
      BackColorSel    =   16635590
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
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
      AllowUserFreezing=   1
      BackColorFrozen =   -2147483633
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Frame FraWhere 
      Height          =   1080
      Left            =   90
      TabIndex        =   0
      Top             =   555
      Width           =   8490
      Begin VB.ComboBox cboPatient 
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   225
         Width           =   1560
      End
      Begin VB.ComboBox cboDir 
         Height          =   315
         Left            =   990
         TabIndex        =   13
         Text            =   "cboDir"
         Top             =   615
         Width           =   1560
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&P"
         Height          =   315
         Left            =   5850
         TabIndex        =   12
         Top             =   615
         Width           =   285
      End
      Begin VB.TextBox txtItem 
         Height          =   315
         Left            =   3585
         TabIndex        =   11
         Top             =   615
         Width           =   2265
      End
      Begin VB.TextBox txtPatiNo 
         Height          =   300
         Left            =   7065
         TabIndex        =   4
         Top             =   225
         Width           =   1170
      End
      Begin MSComCtl2.DTPicker dtpS 
         Height          =   300
         Left            =   3585
         TabIndex        =   1
         Top             =   225
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   112001027
         CurrentDate     =   40954
      End
      Begin MSComCtl2.DTPicker dtpE 
         Height          =   300
         Left            =   4920
         TabIndex        =   2
         Top             =   225
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   112001027
         CurrentDate     =   40954
      End
      Begin VB.Label lblDir 
         AutoSize        =   -1  'True
         Caption         =   "סԺҽ��"
         Height          =   195
         Left            =   105
         TabIndex        =   10
         Top             =   660
         Width           =   720
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "������Ŀ"
         Height          =   195
         Left            =   2700
         TabIndex        =   9
         Top             =   660
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "������ҡ�"
         Height          =   195
         Left            =   105
         TabIndex        =   6
         Top             =   270
         Width           =   900
      End
      Begin VB.Label lblNo 
         AutoSize        =   -1  'True
         Caption         =   "סԺ�š�"
         Height          =   195
         Left            =   6315
         TabIndex        =   5
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lbldate 
         AutoSize        =   -1  'True
         Caption         =   "�������ڡ�"
         Height          =   195
         Left            =   2700
         TabIndex        =   3
         Top             =   270
         Width           =   900
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   180
      Top             =   105
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Menu mnuPop 
      Caption         =   "ѡ��"
      Visible         =   0   'False
      Begin VB.Menu mnuPatiNo 
         Caption         =   "סԺ��"
      End
      Begin VB.Menu mnuBedNo 
         Caption         =   "����"
      End
   End
   Begin VB.Menu mnudate 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu mnuAppdate 
         Caption         =   "��������"
      End
      Begin VB.Menu mnuAdjdate 
         Caption         =   "�������"
      End
   End
End
Attribute VB_Name = "frmPatientRptPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsDept As ADODB.Recordset  '����Ա���ڲ�����¼��
Private mlngOptDeptID As Long           '����ID
Private mlngSys As Long                 'ϵͳ��
Private mcnOracle As ADODB.Connection
Private mstrPrivs As String         'ģ��Ȩ��

Private mintPrint As Integer    '�Ƿ��ڴ�ӡ
Private Enum Col
    ѡ�� = 0: ����: �Ա�: ����: סԺ��: ����: ������Ŀ: �������: ������: �����: ���ʱ��: ҽ��id: ���ͺ�: ����ID: ID
End Enum
Private Function CbsSetting(ByRef cbsMain As CommandBars)
'���ܣ������ڲ˵����岿��
'˵����
'1.���й��еĲ˵��Ͱ�ť�����У���Ϊ�Ӵ��崦��˵��Ļ�׼
'2.�����������������ҵ��Ĳ�ͬ�����ܲ�ͬ
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        
    End With
    cbsMain.EnableCustomization False
    
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    cbsMain.ActiveMenuBar.Closeable = False
      
End Function

Private Sub CbsButtonInit(ByRef cbsMain As CommandBars, Buttons As Collection, _
                         Optional blnLargeIcons As Boolean = False, _
                         Optional Position As XTPBarPosition)
    '�����������˵�
    'cbsMain :����������
    'Buttons :�˵�����,ÿ��Ԫ�صĸ�ʽΪ �˵�id,����,�Ƿ����
    'blnLargeIcons :�Ƿ��ͼ��
    'Position      :�˵�λ��
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    
    Dim strButton As Variant
    Dim varButton As Variant

    Call CbsSetting(cbsMain)
    '����������:������������
    '-----------------------------------------------------
    Set objBar = cbsMain.ActiveMenuBar
    cbsMain.Options.LargeIcons = blnLargeIcons  'Сͼ��
    objBar.Position = Position   '�������ڶ���

    For Each strButton In Buttons
        varButton = Split(strButton, ",")
        With objBar.Controls
            Set objControl = .Add(xtpControlButton, Val(varButton(0)), varButton(1))     '����
            objControl.Style = xtpButtonIconAndCaption
            If UCase(varButton(2)) = "TRUE" Then objControl.BeginGroup = True '����
        End With
    Next
    cbsMain.RecalcLayout
End Sub

Public Sub ShowME(cnOracle As ADODB.Connection, ByVal lngSys As Long, Objfrm As Object, lngOpterDeptID As Long, MainPrivs As String)
    
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo hErr
    
    mlngOptDeptID = lngOpterDeptID
    mlngSys = lngSys
    mstrPrivs = MainPrivs
    Set mcnOracle = cnOracle
    
    strSQL = "Select a.Id, a.����, a.���� From ���ű� A, �������Ҷ�Ӧ D Where a.Id = d.����id And d.����id = [1]"

    Set mrsDept = zlDatabase.OpenSQLRecord(strSQL, "ȡ��������", mlngOptDeptID)
    If mrsDept.EOF Then
        MsgBox "����Աû�пɲ����Ŀ��ң�", vbQuestion, Me.Caption
        Exit Sub
    End If

    Me.Show , Objfrm
    Exit Sub
hErr:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub LoadBaseData()
    'װ�벡������
    Dim lngloop As Long
    dtpS.Value = zlDatabase.Currentdate
    dtpE.Value = dtpS.Value
    txtPatiNo = ""
    cboPatient.Clear
    cboPatient.AddItem "<���п���>"
    
    Do Until mrsDept.EOF
        cboPatient.AddItem Trim("" & mrsDept!���� & "-" & mrsDept!����)
        If mlngOptDeptID = Val("" & mrsDept!ID) Then
            cboPatient.ListIndex = cboPatient.NewIndex
        End If
        cboPatient.ItemData(cboPatient.NewIndex) = Val("" & mrsDept!ID)
        mrsDept.MoveNext
    Loop
    If cboPatient.ListIndex = -1 Then cboPatient.ListIndex = 0
    Call InitDoctors(cboPatient.ItemData(cboPatient.ListIndex))
    Call mnuPatiNo_Click
    Call mnuAppdate_Click
    Call vfgSetting(0, vfgData, "ѡ��,800,1; ����,1200,1;�Ա�,600,1;����,600,1;סԺ��,900,1;����,900,1;������Ŀ,2000,1;�������,1200,1;������,900,1;�����,900,1;���ʱ��,1200,1;ҽ��id,0,1;���ͺ�,0,1;����ID,0,1;ID,0,1")
    vfgData.ColDataType(Col.ѡ��) = flexDTBoolean
End Sub

Private Sub LoadDataToVfg()
    '���ݽ����ϵ�������������������ؼ�
    '
    Dim strSQL As String, rsTmp As ADODB.Recordset, intIndex As Integer
    Dim dateS As Date, dateE As Date, lngPatiDeptID As Long, strNO As String, strDepts As String, strPatients As String
    
    On Error GoTo hErr
    
    txtPatiNo.SetFocus
    
    dateS = CDate(Format(dtpS.Value, "yyyy-MM-dd 00:00:00"))
    dateE = CDate(Format(dtpE.Value, "yyyy-MM-dd 23:59:59"))
    If dateE < dateS Then
        MsgBox "�������ڲ���С�ڿ�ʼ���ڣ�", vbInformation, Me.Caption
        Exit Sub
    End If
    If DateDiff("d", dateS, dateE) > 31 Then
        If MsgBox("��ѯ�����ڷ�Χ������31�죬��Ӱ��ϵͳ��Ӧ�ٶȣ��Ƿ������", vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Sub
    End If
    
    Call vfgSetting(0, vfgData, "ѡ��,800,1; ����,1200,1;�Ա�,600,1;����,600,1;סԺ��,900,1;����,900,1;������Ŀ,2000,1;�������,1200,1;������,900,1;�����,900,1;���ʱ��,1200,1;ҽ��id,0,1;���ͺ�,0,1;����ID,0,1;ID,0,1")
    vfgData.ColDataType(Col.ѡ��) = flexDTBoolean
    
    
    lngPatiDeptID = Val(cboPatient.ItemData(cboPatient.ListIndex))
    
    strNO = Trim(DelInvalidChar(txtPatiNo))
    strDepts = ""
    If lngPatiDeptID <= 0 Then
        If cboPatient.ListCount > 0 Then
            For intIndex = 1 To cboPatient.ListCount - 1
                If Val(cboPatient.ItemData(intIndex)) <> 0 Then strDepts = strDepts & "," & Val(cboPatient.ItemData(intIndex))
            Next
        End If
    End If
    If strDepts <> "" Then
        strDepts = Mid(strDepts, 2)
        strSQL = "select /*+ RULE */ 0 as ѡ��,a.����, a.�Ա�, a.����, a.סԺ��, a.����, a.������Ŀ,f.���� as �������,a.������,a.�����,a.���ʱ��,a.ҽ��id,b.���ͺ�,a.����id,a.id " & vbNewLine & _
            "From ����걾��¼ A, ����ҽ������ B, ���ű� F,(Select * From Table(Cast(f_str2list([5]) As zltools.t_strlist))) G " & vbNewLine & _
            "where a.������Դ = 2 and a.ҽ��id=b.ҽ��id and a.�������id=f.id and  a.���ʱ�� Is Not Null "
        strSQL = strSQL & " And A.�������id = G.Column_Value "
    Else
        strSQL = "select 0 as ѡ��,a.����, a.�Ա�, a.����, a.סԺ��, a.����, a.������Ŀ,f.���� as �������,a.������,a.�����,a.���ʱ��,a.ҽ��id,b.���ͺ�,a.����id,a.id " & vbNewLine & _
            "From ����걾��¼ A, ����ҽ������ B, ���ű� F" & vbNewLine & _
            "where a.������Դ = 2 and a.ҽ��id=b.ҽ��id and a.�������id=f.id  And a.���ʱ�� Is Not Null "
    End If
    
    
    If chkPrinted.Value = 0 Then
        strSQL = strSQL & " And (a.��ӡ���� Is Null Or a.��ӡ���� <= 0)"
    End If
    
    
    If lngPatiDeptID <> 0 Then
        strSQL = strSQL & " and a.�������id = [3] "
    End If
    If strNO <> "" Then
        If lblNo.Caption = "סԺ�š�" Then
            strSQL = strSQL & " and a.סԺ�� = [4] "
        Else
            strSQL = strSQL & " and a.���� = [4] "
        End If
    End If
    
    If lbldate.Caption = "�������ڡ�" Then
        strSQL = strSQL & " and a.����ʱ�� Between [1] And [2] "
    Else
        strSQL = strSQL & " and a.���ʱ�� Between [1] And [2] "
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dateS, dateE, lngPatiDeptID, strNO, strDepts)
    With vfgData
        Do Until rsTmp.EOF
            If strPatients = "" Then
                strPatients = rsTmp("����id")
            Else
                If InStr(strPatients, rsTmp("����id")) = 0 Then
                    strPatients = strPatients & "," & rsTmp("����id")
                End If
            End If
            rsTmp.MoveNext
        Loop
    End With
    
    strSQL = "select distinct 0 as ѡ��,a.����, a.�Ա�, a.����, a.סԺ��, a.����, a.������Ŀ,f.���� as �������,a.������,a.�����,a.���ʱ��,a.ҽ��id,b.���ͺ�,a.����id,a.id " & vbNewLine & _
            "From ����걾��¼ A, ����ҽ������ B, ���ű� F,(Select * From Table(Cast(f_str2list([5]) As zltools.t_strlist))) G��" & vbNewLine & _
            "where a.������Դ = 2 and a.ҽ��id=b.ҽ��id and a.�������id=f.id and  a.���ʱ�� Is Not Null and A.����id = G.Column_Value "
    
    If cboDir.Text <> "" Then
        strSQL = strSQL & " and  a.������=[6]"
    End If
    
    If txtItem.Text <> "" Then
        strSQL = strSQL & " and a.������Ŀ=[7]"
    End If
    
    If chkPrinted.Value = 0 Then
        strSQL = strSQL & " And (a.��ӡ���� Is Null Or a.��ӡ���� <= 0)"
    End If
    
    If lbldate.Caption = "�������ڡ�" Then
        strSQL = strSQL & " and a.����ʱ�� Between [1] And [2] "
    Else
        strSQL = strSQL & " and a.���ʱ�� Between [1] And [2] "
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dateS, dateE, lngPatiDeptID, strNO, strPatients, cboDir.Text, txtItem.Text)
    
    With vfgData
        Do Until rsTmp.EOF
            
            If Val(.TextMatrix(.Rows - 1, Col.ID)) <> 0 Then .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, Col.ID) = Val("" & rsTmp!ID)
            .TextMatrix(.Rows - 1, Col.ѡ��) = Val("" & rsTmp!ѡ��)
            .TextMatrix(.Rows - 1, Col.����) = Trim("" & rsTmp!����)
            .TextMatrix(.Rows - 1, Col.�Ա�) = Trim("" & rsTmp!�Ա�)
            .TextMatrix(.Rows - 1, Col.����) = Trim("" & rsTmp!����)

            .TextMatrix(.Rows - 1, Col.סԺ��) = Trim("" & rsTmp!סԺ��)
            .TextMatrix(.Rows - 1, Col.����) = Trim("" & rsTmp!����)
            .TextMatrix(.Rows - 1, Col.������Ŀ) = Trim("" & rsTmp!������Ŀ)
            .TextMatrix(.Rows - 1, Col.�������) = Trim("" & rsTmp!�������)
            .TextMatrix(.Rows - 1, Col.������) = Trim("" & rsTmp!������)
            
            .TextMatrix(.Rows - 1, Col.�����) = Trim("" & rsTmp!�����)
            .TextMatrix(.Rows - 1, Col.���ʱ��) = Trim("" & rsTmp!���ʱ��)
            .TextMatrix(.Rows - 1, Col.ҽ��id) = Trim("" & rsTmp!ҽ��id)
            .TextMatrix(.Rows - 1, Col.���ͺ�) = Trim("" & rsTmp!���ͺ�)
            .TextMatrix(.Rows - 1, Col.����ID) = Trim("" & rsTmp!����ID)
            
            rsTmp.MoveNext
        Loop
        If rsTmp.RecordCount > 0 Then
            chkPrinted.Caption = "�����Ѵ�ӡ��¼    " & "����" & .Rows - 1 & "����¼"
        Else
            chkPrinted.Caption = "�����Ѵ�ӡ��¼"
        End If
    End With
    Exit Sub
hErr:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitDoctors(ByVal lngdptID As Long)
'���ܣ���ȡ��ǰ���������а�����������Ա
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strOldDoctor As String
    
    strOldDoctor = Me.cboDir.Text
    
    Me.cboDir.Clear
    
    '����ҽ����ʿ
    If lngdptID = 0 Then
        strSQL = _
            "Select Distinct A.ID,A.���,A.����,Upper(A.����) as ����," & _
            " C.��Ա����,Nvl(A.Ƹ�μ���ְ��,0) as ְ��" & _
            " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
            " Where A.ID=B.��ԱID And A.ID=C.��ԱID" & _
            " And C.��Ա���� IN('ҽ��')  " & _
            " And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) "

    Else
        strSQL = _
            "Select Distinct A.ID,A.���,A.����,Upper(A.����) as ����," & _
            " C.��Ա����,Nvl(A.Ƹ�μ���ְ��,0) as ְ��" & _
            " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
            " Where A.ID=B.��ԱID And A.ID=C.��ԱID" & _
            " And C.��Ա���� IN('ҽ��') And B.����ID=[1] " & _
            " And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) "
    End If
    strSQL = strSQL & " Order by ����,��Ա���� Desc"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngdptID)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboDir.AddItem rsTmp!����
            cboDir.ItemData(cboDir.ListCount - 1) = rsTmp!ID
            If rsTmp!���� = strOldDoctor Then
                cboDir.ListIndex = cboDir.NewIndex
            End If
            
            If rsTmp!ID = UserInfo.ID And cboDir.ListIndex = -1 Then cboDir.ListIndex = cboDir.NewIndex
            rsTmp.MoveNext
        Next
        
        If cboDir.ListCount = 1 And cboDir.ListIndex = -1 Then cboDir.ListIndex = 0
    End If
End Sub

Private Sub cboDir_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
'        If mbln΢������Ŀ = True Then
'            Call cbodir_Validate(False)
'            dtp(0).SetFocus
'        Else
'            zlCommFun.PressKey vbKeyTab
            Call cboDir_Validate(False)
'            SetFocusNextIndex Me.cboDir.TabIndex + 2
'            gintSelectFocus = 2
'        End If
    End If
End Sub

Private Sub cboDir_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
        
'    If cboDir.ListIndex <> -1 Then mstrReqDoctor = Me.cboDir.Text: Exit Sub '��ѡ��
    If cboDir.Text = "" Then '������
        Exit Sub
    End If

    strInput = UCase(NeedName(cboDir.Text))
    'ȫԺҽ��
    strSQL = "Select Distinct ����ID From ��������˵�� Where ������� IN(1,2,3)"
    strSQL = "Select Distinct A.ID,A.���,A.����,A.����" & _
        " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
        " Where A.ID=B.��ԱID And A.ID=C.��ԱID And C.��Ա����='ҽ��'" & _
        " And B.����ID IN(" & strSQL & ")" & _
        " And (Upper(A.���) Like [1] Or Upper(A.����) Like [2] Or Upper(A.����) Like [2])" & _
        " And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & _
        " Order by A.����"

    On Error GoTo errH
    vRect = GetControlRect(cboDir.hWnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����ҽ��", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cboDir.Height, blnCancel, False, True, strInput & "%", strInput & "%")
    If Not rsTmp Is Nothing Then
        cboDir.Text = rsTmp!����
'        Me.dtp(0).SetFocus
'        SetFocusNextIndex Me.cbodir.TabIndex
    Else
        If Not blnCancel Then
            MsgBox "δ�ҵ���Ӧ��ҽ����", vbInformation, gstrSysName
        End If
        Cancel = True:  Exit Sub
    End If
'    If Len(Trim(Me.cboDir.Text)) > 0 Then mstrReqDoctor = Me.cboDir.Text
'    gintSelectFocus = 2
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vfgSelect(ByVal intSelect As Integer)
    Dim iRow As Integer
    With vfgData
        For iRow = .FixedRows To .Rows - 1
            If intSelect = 1 Then
                .TextMatrix(iRow, 0) = 1
            Else
                .TextMatrix(iRow, 0) = 0
            End If
        Next
    End With
End Sub
Private Sub PrintSelect()
    '��ӡѡ�еļ�¼
    Dim iRow As Integer, intCount As Integer, intCurr As Integer
    Dim lngRedeID As Long, lngSendID As Long, lngPatiID As Long, lngSampleID As Long
    mintPrint = 1
    intCount = 0: intCurr = 0
    With vfgData
        For iRow = .FixedRows To .Rows - 1
            If Val(.TextMatrix(iRow, Col.ѡ��)) <> 0 Then
                intCount = intCount + 1
            End If
            
        Next
        If intCount <= 0 Then
            MsgBox "��ѡ��Ҫ��ӡ�ı������ִ�д˲�����", vbInformation, Me.Caption
            mintPrint = 0
            Exit Sub
        End If
        If intCount > 300 Then
            If MsgBox("Ҫ��ӡ�ı��泬��300�ݣ���ӡʱ���Ƚϳ����Ƿ������" & vbNewLine & "��[��]��ʼ��ӡ����[��]����ӡ��", vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                mintPrint = 0
                Exit Sub
            End If
        End If
        
        For iRow = .FixedRows To .Rows - 1
            DoEvents
            If Val(.TextMatrix(iRow, Col.ѡ��)) <> 0 Then
                intCurr = intCurr + 1
                lngRedeID = Val(.TextMatrix(iRow, Col.ҽ��id))
                lngSendID = Val(.TextMatrix(iRow, Col.���ͺ�))
                lngPatiID = Val(.TextMatrix(iRow, Col.����ID))
                lngSampleID = Val(.TextMatrix(iRow, Col.ID))
                Call ReportPrint(lngRedeID, lngSendID, lngPatiID, lngSampleID, intCurr, True)
                .TextMatrix(iRow, Col.ѡ��) = 0
            End If
        Next
    End With
    mintPrint = 0
    
End Sub
Private Sub ReportPrint(ByVal lngRedeID As Long, ByVal lngSendID As Long, ByVal lngPatiID As Long, ByVal lngSampleID As Long, _
                        ByVal intCount As Integer, ByVal blnPrint As Boolean)
    '���������ӡ
    'lngRedeID :ҽ��ID
    'lngSendID :���ͺ�
    'lngPatiID :����ID
    'lngSampleID :�걾ID
    
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim blnCurrMoved As Boolean
    
    Dim strSQL As String
    Dim strChart(1 To 9) As String
    Dim intLoop As Integer
    Dim lngKey As Long
    
    Me.MousePointer = 11
    zlCommFun.ShowFlash "���ڴ�ӡ��" & intCount & "�ݱ���...", Me
    
    '����ͼ�ι��Զ��屨�����
    strSQL = "select id from ����ͼ���� where �걾id = [1] order by ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "LISWork.LIS_Report", lngSampleID)
    intLoop = 1
    Do Until rsTmp.EOF
        strChart(intLoop) = App.path & "\" & rsTmp("ID") & ".cht"
        Call LoadImageData(App.path, rsTmp("ID"))
        intLoop = intLoop + 1
        rsTmp.MoveNext
    Loop
    
    If GetReportCode(lngRedeID, lngSendID, strReportCode, strReportParaNo, bytReportParaMode, blnCurrMoved) Then
'        Call ReportOpen(gcnOracle, mlngSys, strReportCode, Me, "NO=" & strReportParaNo, "����=" & bytReportParaMode, "ҽ��ID=" & lngRedeID, _
'                        "����ID=" & lngPatiID, "�걾ID=" & lngSampleID, "���ҽ��=" & lngRedeID, "����걾=" & lngSampleID, _
'                        "ͼ��1=" & strChart(1), "ͼ��2=" & strChart(2), "ͼ��3=" & strChart(3), "ͼ��4=" & strChart(4), _
'                        "ͼ��5=" & strChart(5), "ͼ��6=" & strChart(6), "ͼ��7=" & strChart(7), "ͼ��8=" & strChart(8), _
'                        "ͼ��9=" & strChart(9), IIf(blnPrint, 2, 1))

            Call ReportOpen(mcnOracle, mlngSys, strReportCode, Me, "NO=" & strReportParaNo, "����=" & bytReportParaMode, "ҽ��ID=" & lngRedeID, _
                            "����ID=" & lngPatiID, "�걾ID=" & lngSampleID, _
                            "ͼ��1=" & strChart(1), "ͼ��2=" & strChart(2), "ͼ��3=" & strChart(3), "ͼ��4=" & strChart(4), _
                            "ͼ��5=" & strChart(5), "ͼ��6=" & strChart(6), "ͼ��7=" & strChart(7), "ͼ��8=" & strChart(8), _
                            "ͼ��9=" & strChart(9), IIf(blnPrint, 2, 1))
    End If
    
    
    On Error GoTo errH
    If blnPrint = True Then
        strSQL = "ZL_����걾��¼_�걾�ʿ�(" & lngSampleID & ",'',1)"   '��ӡ������1
        zlDatabase.ExecuteProcedure strSQL, gstrSysName
    End If
    Me.MousePointer = 0
    zlCommFun.StopFlash
    
    On Error Resume Next
    'ɾ��ͼ���ļ�
    For intLoop = 1 To 9
        Kill strChart(intLoop)
    Next
    Exit Sub
errH:
    Me.MousePointer = 0
    zlCommFun.StopFlash
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub vfgSetting(ByVal LngStyle As Long, ByRef objVfg As VSFlexGrid, Optional ByVal strTtile As String, Optional VsfImg As ImageList)
    'lngStyle��0 Ĭ�����ã�ͳһVfg�������
    'strHead��  �����ʽ��
    '           ����1,���,���뷽ʽ;����2,���,���뷽ʽ;.......
    '           ���뷽ʽȡֵ, * ��ʾ����ȡֵ
    '           FlexAlignLeftTop       0   ����
    '           flexAlignLeftCenter    1   ����  *
    '           flexAlignLeftBottom    2   ����
    '           flexAlignCenterTop     3   ����
    '           flexAlignCenterCenter  4   ����  *
    '           flexAlignCenterBottom  5   ����
    '           flexAlignRightTop      6   ����
    '           flexAlignRightCenter   7   ����  *
    '           flexAlignRightBottom   8   ����
    '           flexAlignGeneral       9   ����
    'objVfg:    Ҫ��ʼ���Ŀؼ�
    'VsfImg:    ImageListͼ�꼯�ؼ�����

    Dim arrHead As Variant, i As Long, strHead As String
    If strTtile = "" Then
        strHead = "��1��,900,1;��2��,900,1;��3��,900,1"
    Else
        strHead = strTtile
    End If
    arrHead = Split(strHead, ";")
    
    
    With objVfg
        '1.�߿�
        .Appearance = flex3DLight
        .BorderStyle = flexBorderFlat
        .GridLines = flexGridFlat
        .GridColorFixed = flexGridFlat
        
        '2.��ɫ
        .BackColor = vbWindowBackground '���ڱ���
        .BackColorAlternate = vbWindowBackground
        .BackColorBkg = vbWindowBackground
        .BackColorFixed = vbButtonFace '��ť����
        .BackColorFrozen = &H0&         '��
        .FloodColor = &HC0&             '��
        .BackColorSel = &HFFEBD7        'ǳ��
        .ForeColor = vbWindowText       '�����ı�
        .ForeColorFixed = vbButtonText  '��ť�ı�
        .ForeColorFrozen = &H0&         '��
        .ForeColorSel = vbWindowText
        
        .GridColor = vbApplicationWorkspace 'Ӧ�ó�������
        .GridColorFixed = vbApplicationWorkspace
        .SheetBorder = vbWindowBackground
        .TreeColor = vbButtonShadow         '��ť��Ӱ
        
        '3.��ʼ������

        .Redraw = False
        .Clear
        .Cols = 2
        .FixedRows = 1: .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .ColKey(i) = Split(arrHead(i), ",")(0) '��������ΪcolKeyֵ

            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                
                'Ϊ��֧��zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0 'Ϊ��֧��zl9PrintMode
            End If
        Next
        
        '�̶������־���
        If .FixedRows > 0 And .Cols > 0 Then
            .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        End If
        .RowHeight(0) = 300
        .RowHeightMin = 300
        .WordWrap = True '�Զ�����
        .AutoSizeMode = flexAutoSizeRowHeight '�Զ��и�
        .AutoResize = True '�Զ�
        .Redraw = True
        
        
        '4.��������
        .SelectionMode = flexSelectionByRow     '����ѡ��
        .ExplorerBar = flexExNone               '�����������Ӧ�������ƶ��У�����
        .AllowUserResizing = flexResizeColumns  '�ɵ����п�
        .Editable = flexEDNone                  'ֻ��
        
    End With
    
End Sub
'-------------------------------------------------------------------------------------------------

Private Sub cboPatient_Click()
    Call InitDoctors(cboPatient.ItemData(cboPatient.ListIndex))
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_Manage_SelectAllImages       'ȫѡ
           If mintPrint = 0 Then Call vfgSelect(1)
        Case conMenu_Manage_UnSelectAllImages      'ȫ��
            If mintPrint = 0 Then Call vfgSelect(0)
        Case conMenu_File_PrintSet        '��ӡ����
            If mintPrint = 0 Then Call zlPrintSet
        Case conMenu_File_Print           '��ӡ
            If mintPrint = 0 Then Call PrintSelect
        Case conMenu_View_Find            '��ѯ
            If mintPrint = 0 Then Call LoadDataToVfg
        Case ConMenu_pop_Dept
            Label3.Caption = "������ҡ�"
            InitDepts 0
        Case ConMenu_pop_DeptDistrict
            Label3.Caption = "���벡����"
            InitDepts 1
        Case conMenu_File_Exit            '�˳�
            Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    
    With FraWhere
        .Left = lngLeft + 15
        .Top = lngTop + 15
        .Width = lngRight - lngLeft - 30
    End With
    With vfgData
        .Left = lngLeft + 15
        .Top = FraWhere.Top + FraWhere.Height + 15
        .Width = lngRight - lngLeft - 30
        .Height = chkPrinted.Top - .Top - 30
    End With
    chkPrinted.Left = lngLeft + 15
    chkPrinted.Width = FraWhere.Width

End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control.ID <> conMenu_File_Exit Then Control.Enabled = mintPrint = 0   '��ӡʱ������ִ������������
End Sub


Private Function ShowOpenTree()
    '-----------------------------------------------------------------------------------------
    '����:������+�б�ṹ��������Ŀ����
    '����:������2;�ɹ�����1;ȡ������0
    '-----------------------------------------------------------------------------------------
    Dim strLvw As String
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim objPoint As POINTAPI
    
    On Error GoTo ErrHand
    
    strLvw = "����,1200,0,1;����,2700,0,0;�걾��λ,900,0,0"

    ShowOpenTree = 2
    
    strSQL = "Select * " & vbNewLine & _
             "   From (Select Distinct ID, �ϼ�id, 0 As ĩ��, ����, ����, Null + 0 As �걾��λ," & vbNewLine & _
             "                          Decode(�ϼ�id, Null, ID * Power(10, 20), �ϼ�id * Power(10, 20) + ID) As ����" & vbNewLine & _
             "          From ���Ʒ���Ŀ¼" & vbNewLine & _
             "          Where ���� = 5" & vbNewLine & _
             "          Start With ID In (Select Distinct ����id From ������ĿĿ¼ Where ��� = 'C')" & vbNewLine & _
             "          Connect By Prior �ϼ�id = ID" & vbNewLine & _
             "          Union All" & vbNewLine & _
             "          Select Distinct a.Id, a.����id As �ϼ�id, 1 As ĩ��, a.����, a.����, a.�걾��λ, 1 As ����" & vbNewLine & _
             "          From ������ĿĿ¼ A, ���鱨����Ŀ B" & vbNewLine & _
             "          Where a.��� = 'C' And (a.�����Ŀ = 1 Or a.����Ӧ�� = 1) And a.Id = b.������Ŀid(+) And" & vbNewLine & _
             "                (a.����ʱ�� = To_Date('30000101', 'YYYYMMDD') Or a.����ʱ�� Is Null)) A" & vbNewLine & _
             "   Order By a.ĩ��, a.����, a.����"
            
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    If rs.BOF Then Exit Function
    
    Call ClientToScreen(cmdFind.hWnd, objPoint)
    
    If frmSelectExplorer.ShowSelect(Me, _
                            rs, _
                            objPoint.X * 15 - 30, objPoint.Y * 15 + txtItem.Height - 30, 5400, 2400, _
                            txtItem.Height, _
                            "������Ŀ����ѡ��", _
                            strLvw, _
                            "��ѡ��һ��������Ŀ") Then
        
        txtItem.Text = zlCommFun.Nvl(rs("����").Value) & IIf(rs("�걾��λ") = "", "", "(" & zlCommFun.Nvl(rs("�걾��λ").Value) & ")")
        txtItem.Tag = ""
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub cmdFind_Click()
    ShowOpenTree
End Sub

Private Sub Form_Load()
    Dim Menus As New Collection
    'ȫѡ  ȫ�� ��ӡ �˳�
    Menus.Add conMenu_Manage_SelectAllImages & ",ȫѡ(&A),False"
    Menus.Add conMenu_Manage_UnSelectAllImages & ",ȫ��(&U),False"
    Menus.Add conMenu_File_PrintSet & ",����(&P),True"
    Menus.Add conMenu_File_Print & ",��ӡ(&P),False"
    Menus.Add conMenu_View_Find & ",��ѯ(&F),True"
    Menus.Add conMenu_File_Exit & ",�˳�(&Q),True"
    
    Call CbsButtonInit(cbsMain, Menus, True, xtpBarTop)
    Set Menus = Nothing

    Call LoadBaseData
    mintPrint = 0
    
End Sub

Private Sub Form_Resize()
    Me.chkPrinted.Top = Me.ScaleHeight - Me.chkPrinted.Height - 15
    Call cbsMain_Resize
End Sub



Private Sub Label3_Click()
    Dim objPopup As CommandBar
    Dim cbrControl As CommandBarControl
    Dim vPoint As POINTAPI
    On Error Resume Next

    Set objPopup = Me.cbsMain.Add("Popup", xtpBarPopup)
    With objPopup.Controls
        Set cbrControl = .Add(xtpControlButton, ConMenu_pop_Dept, "�������")
        Set cbrControl = .Add(xtpControlButton, ConMenu_pop_DeptDistrict, "���벡��")
    End With
    vPoint.X = Label3.Left / Screen.TwipsPerPixelX
    vPoint.Y = (Label3.Top + Label3.Height + 30) / Screen.TwipsPerPixelY
    ClientToScreen FraWhere.hWnd, vPoint

    objPopup.ShowPopup , vPoint.X * Screen.TwipsPerPixelX, vPoint.Y * Screen.TwipsPerPixelY
End Sub

Private Function InitDepts(intDeptView As Integer, Optional strErr As String) As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strDeptIDs As String, lngPreDept As Long

    If cboPatient.ListIndex <> -1 Then
        lngPreDept = cboPatient.ItemData(cboPatient.ListIndex)
    End If

    On Error GoTo errH

    If intDeptView = 0 Then
        '�����Ҷ�ȡ��ʾ
        '�����ż���۲��ҵĲ��˻�û���ϴ�������ֻ�Դ����в��˵Ŀ��ҵ�����
        If InStr(mstrPrivs, "ȫԺ����") > 0 Then

            strSQL = _
                " Select Distinct A.ID,A.����,A.����" & _
                " From ���ű� A,��������˵�� B" & _
                " Where B.����ID=A.ID And B.��������='�ٴ�'" & _
                " And (B.������� IN(2,3) Or (B.�������=1 And Exists(Select 1 From ��λ״����¼ C Where B.����ID = C.����ID)))" & _
                " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                " Order by A.����"
        Else
            '����Ȩ�޵Ŀ��ң��������ڿ���+�������������Ŀ���
            strSQL = _
                " Select A.ID,A.����,A.����,Nvl(C.ȱʡ,0) as ȱʡ" & _
                " From ���ű� A,��������˵�� B,������Ա C" & _
                " Where B.����ID=A.ID And A.ID=C.����ID And C.��ԱID=[1]" & _
                " And (B.������� IN(2,3) Or (B.�������=1 And Exists(Select 1 From ��λ״����¼ C Where B.����ID = C.����ID)))" & _
                " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                " And B.��������='�ٴ�'"
            strSQL = strSQL & " Union " & _
                " Select C.ID,C.����,C.����,Nvl(A.ȱʡ,0) As ȱʡ" & _
                " From ������Ա A,�������Ҷ�Ӧ B,���ű� C" & _
                " Where A.����ID=B.����ID And B.����ID=C.ID And A.��ԱID=[1]" & _
                " And Exists(Select 1 From ��������˵�� Where ��������='����' And ����ID=B.����ID)" & _
                " And Not Exists(Select 1 From ��������˵�� Where ��������='�ٴ�' And ����ID=B.����ID)" & _
                " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)"
            If InStr(mstrPrivs, "ICU����") > 0 Then
                strSQL = strSQL & " Union " & _
                    " Select A.ID,A.����,A.����,0 As ȱʡ" & _
                    " From ���ű� A" & _
                    " Where Exists(Select 1 From ��������˵�� B Where A.ID=B.����ID And B.��������='ICU')" & _
                    " And Exists(Select 1 From ��������˵�� B Where A.ID=B.����ID And B.��������='�ٴ�')" & _
                    " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)"
            End If
            strSQL = "Select ID,����,����,Max(ȱʡ) As ȱʡ From (" & strSQL & ") Group By ID,����,���� Order by ����"
        End If
    Else
        '��������ȡ��ʾ
        If InStr(mstrPrivs, "ȫԺ����") > 0 Then

            strSQL = _
                " Select Distinct A.ID,A.����,A.����" & _
                " From ���ű� A,��������˵�� B " & _
                " Where A.ID=B.����ID And B.������� in(1,2,3) And B.��������='����'" & _
                " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                " Order by A.����"
        Else
            '����Ȩ������ֱ�����ڲ���+���ڿ�����������
            strSQL = _
                " Select A.ID,A.����,A.����,Nvl(C.ȱʡ,0) as ȱʡ" & _
                " From ���ű� A,��������˵�� B,������Ա C" & _
                " Where A.ID=B.����ID And A.ID=C.����ID And C.��ԱID=[1]" & _
                " And B.������� in(1,2,3) And B.��������='����'" & _
                " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)"
            strSQL = strSQL & " Union " & _
                " Select C.ID,C.����,C.����,Nvl(A.ȱʡ,0) as ȱʡ" & _
                " From ������Ա A,�������Ҷ�Ӧ B,���ű� C" & _
                " Where A.����ID=B.����ID And B.����ID=C.ID And A.��ԱID=[1]" & _
                " And Exists(Select 1 From ��������˵�� Where ��������='�ٴ�' And ����ID=B.����ID)" & _
                " And Not Exists(Select 1 From ��������˵�� Where ��������='����' And ����ID=B.����ID)" & _
                " And (C.����ʱ�� is NULL or Trunc(C.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)"
            If InStr(mstrPrivs, "ICU����") > 0 Then
                strSQL = strSQL & " Union " & _
                    " Select A.ID,A.����,A.����,0 As ȱʡ" & _
                    " From ���ű� A" & _
                    " Where Exists(Select 1 From ��������˵�� B Where A.ID=B.����ID And B.��������='ICU')" & _
                    " And Exists(Select 1 From ��������˵�� B Where A.ID=B.����ID And B.��������='����')" & _
                    " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)"
            End If
            strSQL = "Select ID,����,����,Max(ȱʡ) as ȱʡ From (" & strSQL & ") Group by ID,����,���� Order by ����"
        End If
    End If

    cboPatient.Clear
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    
    If intDeptView = 0 Then
        cboPatient.AddItem "<���п���>"
    Else
        cboPatient.AddItem "<���в���>"
    End If
    
    For i = 1 To rsTmp.RecordCount
        cboPatient.AddItem rsTmp!���� & "-" & rsTmp!����
        cboPatient.ItemData(cboPatient.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Next
    If rsTmp.RecordCount > 0 Then
        cboPatient.ListIndex = 0
    End If
    InitDepts = True
    Exit Function
errH:
    strErr = "������(GetSampleValCount),������Ϣ:" & Err.Number & " " & Err.Description
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
End Function


Private Sub lbldate_Click()
    PopupMenu mnudate
End Sub

Private Sub lblNo_Click()
    PopupMenu mnuPop
End Sub

Private Sub mnuAdjdate_Click()
    mnuAdjdate.Checked = Not mnuAdjdate.Checked
    If mnuAdjdate.Checked = True Then
        lbldate.Caption = "������ڡ�"
        mnuAppdate.Checked = False
    Else
        mnuAppdate.Checked = True
        lbldate.Caption = "�������ڡ�"
    End If
End Sub

Private Sub mnuAppdate_Click()
    mnuAppdate.Checked = Not mnuAppdate.Checked
    If mnuAppdate.Checked = True Then
        lbldate.Caption = "�������ڡ�"
        mnuAdjdate.Checked = False
    Else
        mnuAdjdate.Checked = True
        lbldate.Caption = "������ڡ�"
    End If
End Sub

Private Sub mnuBedNo_Click()
    mnuBedNo.Checked = Not mnuBedNo.Checked
    If mnuBedNo.Checked = True Then
        lblNo.Caption = "���š�"
        mnuPatiNo.Checked = False
    Else
        mnuPatiNo.Checked = True
        lblNo.Caption = "סԺ�š�"
    End If
End Sub

Private Sub mnuPatiNo_Click()
    mnuPatiNo.Checked = Not mnuPatiNo.Checked
    If mnuPatiNo.Checked = True Then
        lblNo.Caption = "סԺ�š�"
        mnuBedNo.Checked = False
    Else
        mnuBedNo.Checked = True
        lblNo.Caption = "���š�"
    End If
End Sub

Private Sub txtItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OpenSelect (txtItem)
    End If
End Sub

Private Function OpenSelect(ByVal strText As String)
    '-----------------------------------------------------------------------------------------
    '����:���б�ṹ��������Ŀ����
    '����:������2;�ɹ�����1;ȡ������0
    '-----------------------------------------------------------------------------------------
    Dim strInput As String
    Dim rs As New ADODB.Recordset
    Dim strLvw As String
    Dim objPoint As POINTAPI
    Dim strSQL As String
    
    On Error GoTo ErrHand
    
    strLvw = "����,900,0,1;������Ŀ,3600,0,0;�걾��λ,900,0,0"
    
    strInput = "%" & UCase(strText) & "%"
    strSQL = " Select Distinct a.Id, a.����, a.���� as ������Ŀ, a.�걾��λ" & vbNewLine & _
             "    From ������ĿĿ¼ A, ���鱨����Ŀ B" & vbNewLine & _
             "    Where a.��� = 'C' And (a.�����Ŀ = 1 Or a.����Ӧ�� = 1) And a.Id = b.������Ŀid(+) And " & vbNewLine & _
             "          (a.���� Like [1] Or a.���� Like [1] Or" & vbNewLine & _
             "          a.Id In (Select ������Ŀid From ������Ŀ���� Where (���� Like [1] Or Upper(����) Like Upper([1]))))" & vbNewLine & _
             "    Order By a.����"
             
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput)
    If rs.BOF Then
        Exit Function
    End If
    
    If rs.RecordCount = 1 Then GoTo Over
            
    Call ClientToScreen(txtItem.hWnd, objPoint)
    If frmSelectList.ShowSelect(Me, rs, strLvw, objPoint.X * 15 - 30, objPoint.Y * 15 + txtItem.Height - 30, 6000, 4200, Me.Name & "\������Ŀѡ��", "����±���ѡ��һ����Ŀ") Then
        GoTo Over
    End If
    Exit Function
Over:
    txtItem.Text = zlCommFun.Nvl(rs("������Ŀ").Value) & IIf(rs("�걾��λ") = "", "", "(" & zlCommFun.Nvl(rs("�걾��λ").Value) & ")")
    txtItem.Tag = ""
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub vfgData_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iRow As Integer, iCol As Integer
    If Button = 1 Then
        With vfgData
            iCol = .MouseCol: iRow = .MouseRow
            
            If iCol = Col.ѡ�� And iRow >= .FixedRows And iRow <= .Rows - 1 Then
                If Val(.TextMatrix(iRow, Col.ѡ��)) = 0 Then
                    .TextMatrix(iRow, Col.ѡ��) = 1
                Else
                    .TextMatrix(iRow, Col.ѡ��) = 0
                End If
            End If
        End With
    End If
End Sub

