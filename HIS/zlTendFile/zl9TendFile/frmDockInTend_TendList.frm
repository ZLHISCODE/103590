VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDockInTend_TendList 
   BorderStyle     =   0  'None
   Caption         =   "�����ļ��б�"
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer TimFresh 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   0
      Top             =   3630
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   3300
      Left            =   30
      ScaleHeight     =   3300
      ScaleWidth      =   6690
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   6690
      Begin MSComctlLib.ImageList imgData 
         Left            =   1005
         Top             =   1695
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDockInTend_TendList.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDockInTend_TendList.frx":6862
               Key             =   "����"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDockInTend_TendList.frx":6DFC
               Key             =   "��ͨ"
            EndProperty
         EndProperty
      End
      Begin VB.Frame fra 
         Height          =   525
         Left            =   0
         TabIndex        =   1
         Top             =   -90
         Width           =   6015
         Begin VB.ComboBox cboBaby 
            Height          =   300
            Left            =   630
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   150
            Width           =   1350
         End
         Begin VB.Label lbl���� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�鿴"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   180
            TabIndex        =   4
            Top             =   210
            Width           =   360
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgFile 
         Height          =   1095
         Left            =   -15
         TabIndex        =   3
         Top             =   435
         Width           =   6060
         _cx             =   10689
         _cy             =   1931
         Appearance      =   2
         BorderStyle     =   0
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
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
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
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeCommandBars.ImageManager imgPublic 
      Bindings        =   "frmDockInTend_TendList.frx":7396
      Left            =   135
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmDockInTend_TendList.frx":73AA
   End
End
Attribute VB_Name = "frmDockInTend_TendList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'�󶨿�ݼ�ʱ,IDֵ������޷������͵�ȡֵ��Χ���޷���,Ҳ����0-65535
Private Const conMenu_Add As Long = 32761 '����
Private Const conMenu_Modify As Long = 32762 '�޸�
Private Const conMenu_Delete As Long = 32763 'ɾ��

Private Enum mCol
    f��־ = 0: fID: f��ʽID: f�ļ�: f��ʼ����: f����ID: f����: f����: f��������
End Enum

Private mblnInit As Boolean
Private mblnNoRefresh As Boolean
Private mstrPrivs As String                             '��ǰʹ���߶Ա�����(1255)��Ȩ�޴�
Private mlngPatiID As Long                              '����id
Private mlngPageId As Long                              '��ҳid
Private mlngDeptId As Long                              '��ǰ��������id���粡�˿��Һ͵�ǰ���Ҳ�һ�£����ܲ����鵵��Ĺ���
Private mlngFileID As Long                              '��Ҫ��λ�����ļ�ID
Private mlngFormatID As Long                            '�ļ���ʽID
Private mlng��� As Integer                             'ѡ���˱��˻�Ӥ��
Private mblnEdit As Boolean                             '�Ƿ����������ͨ�����ϼ�������ݵ�ǰ���������Ƿ�ǰ���˲���������
Private mblnDoctorStation As Boolean
Private mintCurveReSize As Integer                      '���µ��鿴�Ƿ�����Сģʽ 0��С 1ԭʼ��С
Private rsTemp As New ADODB.Recordset
Private mintBaby As Integer
Private mfrmMain As Object
Private mbytFontSize As Byte
Private mblnChange As Boolean                           '�޸ı�־
Private mblnSign As Boolean                             'ǩ����־
Private mblnArchive As Boolean                          '�鵵��־
Private mblnRefreshFontSize As Boolean                  '�Ƿ���ˢ�����ݺ��Զ������������幦��(�����ڲ����õ��Զ�ˢ��)
Private mblnTemparatureChat As Boolean                  '�Ƿ��Ǳ�׼���µ�

'���ѿɷ���鿴���µ��뻤���¼���������,����ʽ�鿴��ʧȥ����,��д������
Public Event Activate()         '���°�ť��˵�
Public Event ViewFile(ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, _
    ByVal intBaby As Integer, ByVal blnChildForm As Boolean, ByVal strPrivs As String, ByVal blnEdit As Boolean, ByVal bytSize As Byte)
Public Event ViewAnimalHeat(strPara As String, bytMode As Byte, strPrivs As String, ByVal bytSize As Byte)
Public Event ShowData(intBaby As Integer, lngFile As Long, lngDept As Long, bytSel As Byte, ByVal intCurveReSize As Integer)                 '֪ͨ����ҳ��ˢ��
Public Event PrintTendFile(ByVal bytKind As Byte, ByVal bytMode As Byte)
Public Event SaveDocument(blnSave As Boolean)                                                               '����ָ�����
Public Event SignDocument(blnOK As Boolean, blnVerify As Boolean, blnExchange As Boolean)                                           '����ȡ��ǩ��
Public Event ArchiveDocument(blnOK As Boolean)                                                              '����ȡ���鵵
Public Event SignMarker()
Public Event ViewCaveData(ByVal intDataEditor As Integer)
Public Event Viewpartogram(strPara As String, bytMode As Byte, strPrivs As String, ByVal bytSize As Byte)
Public Event ViewpartogramEditor(ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, _
    ByVal intBaby As Integer, ByVal strPrivs As String, ByVal bytSize As Byte)
Public Event ViewReSetFontSize(ByVal intSEL As Integer, ByVal bytSize As Byte)
Public Event BulkPrintDocument(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, ByVal intBaby As Integer)

Public Sub SetChange(ByVal blnChange As Boolean)
    mblnChange = blnChange
End Sub

Public Sub SetState(ByVal blnSign As Boolean, ByVal blnArchive As Boolean)
    mblnArchive = blnArchive
    mblnSign = blnSign
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strInfo As String
    Dim byt����ȼ� As Byte
    Dim rs As New ADODB.Recordset
    
    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_FileMan
        Call frmNurseFileMan.ShowEditor(mlngPatiID, mlngPageId, mintBaby, mstrPrivs, False, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)))
    Case conMenu_File_Open
        Call vfgFile_DblClick
'        With vfgFile
'            strInfo = Val(.TextMatrix(.ROW, mCol.f����ID))
'            If Val(.TextMatrix(.ROW, mCol.f����)) = -1 Then
'                '���µ��鿴������ID;��ҳID;����ID;��Ժ;�༭;Ӥ��
'                If Not CreateBodyEditor Then Exit Sub
'                RaiseEvent ViewAnimalHeat(mlngPatiId & ";" & mlngPageId & ";" & mlngDeptId & ";" & Val(.TextMatrix(.ROW, mCol.fID)) & ";0;0;" & mintBaby & ";1", 0, mstrPrivs)
'            ElseIf Val(.TextMatrix(.ROW, mCol.f����)) = 1 Then
'                '����ͼ�鿴:�ļ�ID;����ID;��ҳID;����ID
'                If Not CreatePartogram Then Exit Sub
'                RaiseEvent Viewpartogram(Val(.TextMatrix(.ROW, mCol.fID)) & ";" & mlngPatiId & ";" & mlngPageId & ";" & mlngDeptId, 1, mstrPrivs)
'            Else
'                RaiseEvent ViewFile(Val(.TextMatrix(.ROW, mCol.fID)), mlngPatiId, mlngPageId, mlngDeptId, mintBaby, False, mstrPrivs, True)
'            End If
'        End With
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_PrintSet
        With vfgFile
            If Val(.TextMatrix(.ROW, mCol.f����)) = -1 Then
                If Not CreateBodyEditor Then Exit Sub
                Call gobjBodyEditor.zlPrintSet(Me)
            ElseIf Val(.TextMatrix(.ROW, mCol.f����)) = 1 Then
                If Not CreatePartogram Then Exit Sub
                Call gobjPartogram.zlPrintSet(Me, 1)
            Else
                frmPrintSet.Show 1
            End If
        End With
    Case conMenu_File_Preview
        ''1-Ԥ��,2-��ӡ
        With vfgFile
            If Val(.TextMatrix(.ROW, mCol.f����)) = -1 Then
                RaiseEvent PrintTendFile(1, 1)
            ElseIf Val(.TextMatrix(.ROW, mCol.f����)) = 1 Then
                RaiseEvent PrintTendFile(3, 1)
            Else
                RaiseEvent PrintTendFile(2, 1)
            End If
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Print
        With vfgFile
            If Val(.TextMatrix(.ROW, mCol.f����)) = -1 Then
                RaiseEvent PrintTendFile(1, 2)
            ElseIf Val(.TextMatrix(.ROW, mCol.f����)) = 1 Then
                RaiseEvent PrintTendFile(3, 2)
            Else
                RaiseEvent PrintTendFile(2, 2)
            End If
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Excel
        With vfgFile
            If Val(.TextMatrix(.ROW, mCol.f����)) = -1 Then
                MsgBox "�Բ������µ���֧�������Excel��", vbInformation, gstrSysName
            ElseIf Val(.TextMatrix(.ROW, mCol.f����)) = 1 Then
                MsgBox "�Բ��𣬲���ͼ��֧�������Excel��", vbInformation, gstrSysName
            Else
                RaiseEvent PrintTendFile(2, 3)
            End If
        End With
    '51588:������,2012-12-12,�����ļ����������ӡ
    Case conMenu_File_Print * 100# + 1
        gstrSQL = "SELECT A.ID FROM ���˻����ļ� A,�����ļ��б� B WHERE A.��ʽID=B.ID  And A.����ID=[1] And A.��ҳID=[2]"
        Set rs = zldatabase.OpenSQLRecord(gstrSQL, "����Ƿ����Ҫ��ӡ���ļ�", mlngPatiID, mlngPageId)
        If rs.RecordCount = 0 Then
            MsgBox "�ò���û���κλ����ļ�,������ӻ����ļ���", vbInformation, gstrSysName
            Exit Sub
        End If
        RaiseEvent BulkPrintDocument(mlngPatiID, mlngPageId, mlngDeptId, mintBaby)
    Case conMenu_Tool_Sign
        RaiseEvent SignDocument(True, False, False)
    '51589:������,2013-03-01,��ӽ���ǩ��
    Case conMenu_Tool_SignShiftExchange  '����ǩ��
        RaiseEvent SignDocument(True, False, True)
    Case conMenu_Tool_SignEarse
        RaiseEvent SignDocument(False, False, False)
    Case conMenu_Tool_SignAuditAffirm
        RaiseEvent SignDocument(True, True, False)
    Case conMenu_Tool_SignAuditCancel
        RaiseEvent SignDocument(False, True, False)
    Case conMenu_Edit_Archive * 10
        RaiseEvent ArchiveDocument(True)
    Case conMenu_Edit_UnArchive
        If Val(Split(EprIsCommit(mlngPatiID, mlngPageId), "|")(2)) = 0 Then
            MsgBox "�ò��˵Ĳ������ύ���[״̬��" & gstrMecState & "]�����ܳ����鵵����ȡ���������ԣ�", vbInformation, gstrSysName
            Exit Sub
        End If
        RaiseEvent ArchiveDocument(False)
    Case conMenu_Edit_Save
        RaiseEvent SaveDocument(True)
    Case conMenu_Tool_SignVerify
        RaiseEvent SignMarker
    Case conMenu_Edit_Transf_Cancle
        RaiseEvent SaveDocument(False)
    Case conMenu_File_PrintDayDetail, conMenu_Edit_Curve, conMenu_Edit_CurveTable, conMenu_Edit_Curve_Show, conMenu_Edit_Surgery_Edit '����¼��,���ü�¼,��ʾ,����/��������
        If Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) = -1 Then
            On Error Resume Next
            Dim strDLL As String
            Dim strSQL As String
            Dim objChart As Object
            Dim rsTemp As New ADODB.Recordset
            
            strSQL = " Select �²��� From ���²��� Where Nvl(����,0)=1"
            Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "��ȡ���²���")
            If Err <> 0 Then
                strDLL = "zl9TemperatureChart"
            Else
                If rsTemp.RecordCount = 0 Then
                    strDLL = "zl9TemperatureChart"
                Else
                    strDLL = NVL(rsTemp!�²���, "zl9TemperatureChart")
                End If
            End If
            
            Err = 0
            strDLL = strDLL & ".clsBodyEditor"
            Set objChart = CreateObject(strDLL)
            If Err <> 0 Then
                MsgBox "    �������²���ʧ�ܣ�" & vbCrLf & "    ���򽫴�����׼�����²�����������չ�֣�����ָ�������²����Ƿ���ڻ����𻵣�" & vbCrLf & "    ��ϸ����" & Err.Description, vbInformation, gstrSysName
                
                '�������ָ�������²��������򴴽���׼�����²�������Ϊ���ﲻ����Ļ���������ܴ���ֱ��ʹ�����²����еĶ��󣬴Ӷ����³������
                strDLL = "zl9TemperatureChart.clsBodyEditor"
                Set objChart = CreateObject(strDLL)
            End If
            
            On Error GoTo ErrHand
            Call objChart.InitBodyEditor(glngSys, gcnOracle)
            Select Case Control.ID
                Case conMenu_File_PrintDayDetail
                    Call objChart.BodyMutilEditor(Me, mlngDeptId, mstrPrivs, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)))
                Case conMenu_Edit_Curve
                    RaiseEvent ViewCaveData(0)
                Case conMenu_Edit_CurveTable
                    RaiseEvent ViewCaveData(-1)
                Case Else
                    RaiseEvent ViewCaveData(1)
            End Select
        Else
            If Control.ID <> conMenu_File_PrintDayDetail Then Exit Sub
            Dim frmTendFileMutil As New frmTendFileMutilEditor
            Call frmTendFileMutil.ShowMe(Me, mlngDeptId, mstrPrivs, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)))
        End If
    Case conMenu_Edit_Billing '�������ݱ༭
        If Not CreatePartogram Then Exit Sub
        RaiseEvent ViewpartogramEditor(Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)), mlngPatiID, mlngPageId, mlngDeptId, 0, mstrPrivs, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)))
    Case conMenu_Tool_Option '����ѡ��
        '���������������ͼ�ͼ�¼���������ǹ����ģ�ȡ��ԭ�в������ý���
        If Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) = 1 Then '����ͼ
'            If Not CreatePartogram Then Exit Sub
'            If gobjPartogram.zlPartogramPara(Me, mstrPrivs) Then
'                Call RefreshData(mlngPatiID, mlngPageId, mlngDeptId, mblnDoctorStation, mblnEdit, Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)))
'            End If
        ElseIf Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) = -1 Then '���µ�
            If Not CreateBodyEditor Then Exit Sub
            If gobjBodyEditor.GetCaseTendBodyPara.ShowPara(Me, mstrPrivs) Then
                Call RefreshData(mlngPatiID, mlngPageId, mlngDeptId, mblnDoctorStation, mblnEdit, Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)))
            End If
        Else '��¼��
'            If frmTendPara.ShowPara(Me, mstrPrivs) Then
'                Call RefreshData(mlngPatiID, mlngPageId, mlngDeptId, mblnDoctorStation, mblnEdit, Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)))
'            End If
        End If
    End Select
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If

    Call SaveErrLog
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    If Not mblnInit Then Exit Sub
    
    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_EditPopup
         Control.Visible = (mblnDoctorStation = False)
         Control.Enabled = Control.Visible
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_FileMan
        Control.Visible = (InStr(1, mstrPrivs, "�����ļ�����") > 0 And mblnDoctorStation = False And Not gblnMoved)
        Control.Enabled = (mlngPatiID > 0) And Not mblnArchive And Control.Visible And mblnEdit
    Case conMenu_File_Open
        Control.Visible = True
        Control.Enabled = (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_PrintSet
        'Control.Enabled = Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) = -1
    Case conMenu_File_Preview, conMenu_File_Print
        Control.Enabled = (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Excel
        Control.Enabled = (vfgFile.Rows > 1 And (InStr(1, ",-1,1,", "," & Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) & ",") = 0))
    Case conMenu_File_ExportToXML, conMenu_File_RowPrint, conMenu_Edit_Audit, conMenu_Edit_Sort, _
        conMenu_Tool_Monitor, conMenu_Edit_Archive * 10 + 1
        Control.Visible = False
        Control.Enabled = False
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_PrintDayDetail
        Control.Enabled = (mblnEdit And mlngPatiID > 0 And (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.f����)) <> 1))

        Control.Visible = (InStr(1, mstrPrivs, "�����¼�Ǽ�") > 0 And mblnDoctorStation = False And Not gblnMoved And (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.f����)) <> 1))
        If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "�����¼�Ǽ�") > 0)
    Case conMenu_Edit_Curve, conMenu_Edit_CurveTable '���ü�¼
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "���µ���ͼ") > 0) And Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) = -1
        Control.Enabled = (mlngPatiID > 0) And (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0) And Control.Visible And mblnEdit
    Case conMenu_Edit_Curve_Show '��ʾ
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "���µ���ͼ") > 0) And Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) = -1 And mblnTemparatureChat = False
        Control.Enabled = (mlngPatiID > 0) And (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0) And Control.Visible And mblnEdit
    Case conMenu_Edit_Surgery_Edit '����/��������
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "���µ���ͼ") > 0) And Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) = -1 And mblnTemparatureChat = True
        Control.Enabled = (mlngPatiID > 0) And (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0) And Control.Visible And mblnEdit
    Case conMenu_Edit_Billing  '�������ݱ༭
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "����ͼ��ͼ") > 0) And Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) = 1
        Control.Enabled = (mlngPatiID > 0) And (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0) And Not mblnArchive And Control.Visible And mblnEdit
    Case conMenu_Tool_Sign  'ǩ��
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "�����¼ǩ��") > 0) And (InStr(1, ",-1,1,", "," & Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) & ",") = 0)
        Control.Enabled = (mlngPatiID > 0) And (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0) And Not mblnArchive And Not mblnChange And Control.Visible And mblnEdit
    '51589:������,2013-03-01,��ӽ���ǩ��
    Case conMenu_Tool_SignShiftExchange  '����ǩ��
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "�����¼ǩ��") > 0) And (InStr(1, ",-1,1,", "," & Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) & ",") = 0)
        Control.Enabled = (mlngPatiID > 0) And (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0) And Not mblnArchive And Not mblnChange And Control.Visible And mblnEdit
    Case conMenu_Tool_SignEarse  'ȡ��ǩ��
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "ȡ����¼ǩ��") > 0) And (InStr(1, ",-1,1,", "," & Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) & ",") = 0)
        Control.Enabled = (mlngPatiID > 0) And (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0) And Not mblnArchive And Not mblnChange And Control.Visible And mblnEdit
    Case conMenu_Tool_SignAuditAffirm, conMenu_Tool_SignAuditCancel  '��ǩ,ȡ����ǩ
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "�����¼��ǩ") > 0) And (InStr(1, ",-1,1,", "," & Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) & ",") = 0)
        Control.Enabled = (mlngPatiID > 0) And (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0) And Not mblnArchive And Not mblnChange And Control.Visible And mblnEdit
        If Control.ID = conMenu_Tool_SignAuditCancel And Control.Enabled Then
            Control.Enabled = (InStr(1, mstrPrivs, "ȡ����¼ǩ��") > 0)
        End If
    Case conMenu_Edit_Archive * 10 '�鵵
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "�����¼�鵵") > 0)
        Control.Enabled = (mlngPatiID > 0) And (Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)) > 0) And Not mblnArchive And Not mblnChange And Control.Visible And mblnEdit
    Case conMenu_Edit_UnArchive  'ȡ���鵵
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "ȡ����¼�鵵") > 0)
        Control.Enabled = (mlngPatiID > 0) And (Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)) > 0) And mblnArchive And Not mblnChange And Control.Visible And mblnEdit
    Case conMenu_Edit_Save  '����
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "�����¼�Ǽ�") > 0)
        Control.Enabled = mblnChange And Not mblnArchive And Control.Visible And (InStr(1, ",-1,1,", "," & Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) & ",") = 0)
    Case conMenu_Edit_Transf_Cancle  'ȡ��
        Control.Visible = Not mblnDoctorStation And Not gblnMoved
        Control.Enabled = mblnChange And Not mblnArchive And Control.Visible And (InStr(1, ",-1,1,", "," & Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) & ",") = 0)
    Case conMenu_Tool_SignVerify
        Control.Visible = (Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) = 0) And Not mblnDoctorStation And Not gblnMoved
        Control.Enabled = Not mblnChange And Not mblnArchive And Control.Visible And Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) = 0 And mblnEdit
    Case conMenu_Tool_Option '����ѡ��
        If Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) = 1 Then
            Control.Caption = "����ѡ��"
        Else
            Control.Caption = "����ѡ��"
        End If
        '���������������ͼ�ͼ�¼���������ǹ����ģ�ȡ��ԭ�в������ý���
        Control.Visible = Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) = -1
        Control.Enabled = (mlngPatiID > 0) And (Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)) > 0) And Control.Visible
    End Select
    
End Sub

Private Function zlRefData() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Dim intRow As Integer
    Dim lngID As Long
    
    Err = 0
    On Error GoTo ErrHand
    '------------------------------------------------------------------------------------------------------------------
    '�����ļ�ˢ��
    
    With vfgFile
        .Clear
        .Rows = 2
        .Cols = 9
        .FixedCols = 1
        
        .TextMatrix(0, mCol.f��־) = ""
        .TextMatrix(0, mCol.fID) = "ID"
        .TextMatrix(0, mCol.f��ʽID) = "��ʽID"
        .TextMatrix(0, mCol.f�ļ�) = "�ļ�"
        .TextMatrix(0, mCol.f��ʼ����) = "��ʼ����"
        .TextMatrix(0, mCol.f����ID) = "����id"
        .TextMatrix(0, mCol.f����) = "����"
        .TextMatrix(0, mCol.f����) = "����"
        .TextMatrix(0, mCol.f��������) = "��������"
        
        Set .Cell(flexcpPicture, 1, mCol.f��־) = Nothing
        .TextMatrix(1, mCol.fID) = ""
        .TextMatrix(1, mCol.f��ʽID) = ""
        .TextMatrix(1, mCol.f�ļ�) = ""
        .TextMatrix(1, mCol.f��ʼ����) = ""
        .TextMatrix(1, mCol.f����ID) = ""
        .TextMatrix(1, mCol.f����) = ""
        .TextMatrix(1, mCol.f����) = ""
        .TextMatrix(1, mCol.f��������) = ""
        
        .ColWidth(mCol.f��־) = 270
        .ColWidth(mCol.fID) = 0: .ColWidth(mCol.f��ʽID) = 0: .ColWidth(mCol.f�ļ�) = 2000: .ColWidth(mCol.f��ʼ����) = 1200
        .ColWidth(mCol.f����ID) = 0: .ColWidth(mCol.f����) = 1200: .ColWidth(mCol.f����) = 0: .ColWidth(mCol.f��������) = 0
    End With
    
    intRow = vfgFile.FixedRows
    '--------------------------------------------------------------------------------------------------------------
    gstrSQL = "" & _
        " SELECT A.ID,A.��ʽID,A.����ID,C.���� AS ����,A.�ļ�����,A.��ʼʱ��,A.����ʱ��,B.����,b.���" & vbNewLine & _
        " FROM ���˻����ļ� A,�����ļ��б� B,���ű� C" & vbNewLine & _
        " WHERE A.��ʽID=B.ID AND A.����ID=C.ID And A.����ID=[1] And A.��ҳID=[2] And A.Ӥ��=[3]" & _
        " ORDER BY B.����,A.��ʼʱ�� "
    Call SQLDIY(gstrSQL)
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiID, mlngPageId, mintBaby)
    
    With Me.vfgFile
        Do While Not rsTemp.EOF
            If Val(.TextMatrix(.Rows - 1, mCol.fID)) > 0 Then .AddItem ""
            If rsTemp!���� = -1 Then
                Set .Cell(flexcpPicture, .Rows - 1, mCol.f��־) = imgData.ListImages("����").Picture
            Else
                Set .Cell(flexcpPicture, .Rows - 1, mCol.f��־) = imgData.ListImages("��ͨ").Picture
            End If
            
            lngID = Val(NVL(rsTemp!ID, 0))
            If mlngFormatID > 0 And mlngFormatID = Val(NVL(rsTemp!��ʽID)) Then mlngFileID = lngID
            
            If mlngFileID <> 0 And lngID = mlngFileID Then
                intRow = .Rows - 1
            End If
            .TextMatrix(.Rows - 1, mCol.fID) = lngID
            .TextMatrix(.Rows - 1, mCol.f��ʽID) = NVL(rsTemp!��ʽID, 0)
            .TextMatrix(.Rows - 1, mCol.f�ļ�) = NVL(rsTemp!�ļ�����)
            .TextMatrix(.Rows - 1, mCol.f��ʼ����) = Format(NVL(rsTemp!��ʼʱ��), "yyyy-MM-dd HH:mm:ss")
            .TextMatrix(.Rows - 1, mCol.f����ID) = NVL(rsTemp!����ID)
            .TextMatrix(.Rows - 1, mCol.f����) = NVL(rsTemp!����)
            .TextMatrix(.Rows - 1, mCol.f����) = NVL(rsTemp!����)
            .TextMatrix(.Rows - 1, mCol.f��������) = Format(NVL(rsTemp!����ʱ��), "yyyy-MM-dd HH:mm:ss")
            
            rsTemp.MoveNext
        Loop
    End With
    
    'ѡ����
    Call vfgFile.Select(intRow, mCol.fID)
    
    If mblnEdit = True Then
        '41778,������,2012-09-06
        '��������ϰ���°����ݶ��Ѿ����ڣ������κ����ơ����ֻ���ϰ����ݣ�û���°档��������ļ���
        'Ӥ��Ӧ�ú�ĸ��ʹ��ͬһ��ϵͳ��
        gstrSQL = " Select 1 ��� From ���˻����¼ Where ����id = [1] And ��ҳid = [2] And Rownum < 2" & vbNewLine & _
                "   Union All" & vbNewLine & _
                "   Select 2 ��� From ���˻����ļ� Where ����id = [1] And ��ҳid = [2] And Rownum < 2"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiID, mlngPageId)
        If rsTemp.RecordCount > 0 Then
            rsTemp.MoveLast
            If Val(rsTemp!���) = 1 Then mblnEdit = False
        End If
    End If
    
    zlRefData = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitData(ByVal frmMain As Object, ByVal strPrivs As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mstrPrivs = strPrivs
    Set mfrmMain = frmMain
    
    If ExecuteCommand("��ʼ�ؼ�") = False Or ExecuteCommand("��ʼ����") = False Then Exit Function
    Call ExecuteCommand("��ע���")
    Call ExecuteCommand("�ؼ�״̬")
    
End Function

Public Function RefreshData(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lngDeptID As Long, ByVal blnDoctorStation As Boolean, _
    ByVal blnEdit As Boolean, Optional ByVal lngFileID As Long = 0, Optional ByVal lng��� As Integer = 0, Optional ByVal intCurveReSize As Integer = 0, _
    Optional blnRefreshFontSize As Boolean = True) As Boolean
    '******************************************************************************************************************
    '���ܣ�ˢ������
    '������ lngFileID Ϊ0Ĭ��ѡ���һ���ļ�����Ϊ0��ѡ����ļ���int��� 0Ϊ���˱��� ����ΪӤ�����
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    
    mblnInit = False
    mlngPatiID = lng����ID
    mlngPageId = lng��ҳID
    mlngDeptId = lngDeptID
    mblnEdit = blnEdit And Not gblnMoved
    mintCurveReSize = intCurveReSize
    mblnDoctorStation = blnDoctorStation
    mblnRefreshFontSize = blnRefreshFontSize
    '�ļ�ID<>0˵�����޸Ļ�������ļ� =0�������л��˲��ˡ�
    '�޸�����ļ����Զ���λ��������ļ���ȴ�����˽���λ����Ŀ��ʽ���ļ���(û����ͬ��ʽ���ļ���λ����һ���ļ�)
    If lngFileID <> 0 Then
        mlngFileID = lngFileID
        mlngFormatID = 0
    End If
    mlng��� = IIf(lng��� < 0, 0, lng���)
    
    Call ExecuteCommand("ˢ������")
End Function

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim intRow As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    Dim strNow As String
    Dim strNote As String
    Dim byt����ȼ� As Byte
    Static strPatient As String     '����ID|��ҳID|Ӥ��
    
    On Error GoTo ErrHand

    Call SQLRecord(rsSQL)

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ�ؼ�"
        Call InitCommandBar
'        Set mclsDockAduits = New zlRichEPR.clsDockAduits
'        Call FormSetCaption(mclsDockAduits.zlGetFormTendBody, False, False)
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"

        
    '------------------------------------------------------------------------------------------------------------------
    Case "�ؼ�״̬"

        
    '------------------------------------------------------------------------------------------------------------------
    Case "ˢ��״̬"
        

        
    '------------------------------------------------------------------------------------------------------------------
    Case "ˢ������"
    
        '�жϲ����Ƿ���ת��
        '��Ϊ�ú������ⶼ�ڵ���,�������ñ�,ֱ�Ӷ�ȡ
        '------------------------------------------------------------------------------------------------------------------
        gblnMoved = False
        
        If mlngPatiID <> 0 Then
            '��鲡�������ļ��Ƿ��Ѿ��ύ�������ҡ������Ƿ�ת��
            gstrSQL = "Select ����ת�� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
            Set rs = zldatabase.OpenSQLRecord(gstrSQL, "�ж������Ƿ�ת��", mlngPatiID, mlngPageId)
            gblnMoved = NVL(rs!����ת��, 0) <> 0
        End If
        
        mblnNoRefresh = True
        cboBaby.Clear
        cboBaby.AddItem "���˱���"
        gstrSQL = "Select a.���,Decode(a.Ӥ������,Null,NVL(C.����,b.����) ||'֮��'||Trim(To_Char(a.���,'9')),a.Ӥ������) As Ӥ������" & _
            " From ������Ϣ b,������ҳ C,������������¼ a Where b.����id=C.����id And A.����ID=C.����ID And A.��ҳID=C.��ҳID And C.����id=[1] And C.��ҳid=[2]  Order By a.���"
        Set rs = zldatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngPatiID, mlngPageId)
        If rs.BOF = False Then
            Do While Not rs.EOF
                cboBaby.AddItem rs("Ӥ������").Value: cboBaby.ItemData(cboBaby.NewIndex) = Val(NVL(rs("���").Value, 0))
                If cboBaby.ListIndex = -1 And Val(NVL(rs("���").Value, 0)) = mlng��� Then cboBaby.ListIndex = cboBaby.NewIndex
                rs.MoveNext
            Loop
        End If
        If cboBaby.ListIndex = -1 And cboBaby.ListCount > 0 Then cboBaby.ListIndex = 0
        cboBaby.Enabled = (cboBaby.ListCount > 1)
        
        Call zlRefData
        mblnNoRefresh = False
        Call ExecuteCommand("��ʾ�ļ�����", vfgFile.ROW)
        
        mblnInit = True
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʾ�ļ�����"
        'todo:Ӧ�ô��ļ�ID,���ϳ���ֻ���ܸ�ʽID,��Ҫ�޸ĳ���
        
        If Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)) <> 0 Then mlngFileID = Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID))
        If Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f��ʽID)) <> 0 Then mlngFormatID = Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f��ʽID))
        RaiseEvent ShowData(mintBaby, Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)), mlngDeptId, Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) + 1, mintCurveReSize)
        If mblnRefreshFontSize = True And Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)) <> 0 Then
            RaiseEvent ViewReSetFontSize(Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) + 1, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)))
        End If
        mblnRefreshFontSize = True
        If Not mblnDoctorStation And mblnEdit = True Then
            If Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)) = 0 And InStr(1, mstrPrivs, "�����ļ�����") > 0 And strPatient <> mlngPatiID & "|" & mlngPageId & "|" & mintBaby Then
                '���˲������ύ���,��������ļ�
                If Val(Split(EprIsCommit(mlngPatiID, mlngPageId), "|")(0)) = 0 Then Exit Function
                If frmNurseFileMan.ShowEditor(mlngPatiID, mlngPageId, mintBaby, mstrPrivs, True, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))) Then
                    Call ExecuteCommand("ˢ������")
                End If
            End If
            strPatient = mlngPatiID & "|" & mlngPageId & "|" & mintBaby
        End If
    
    End Select

    ExecuteCommand = True

    GoTo EndHand
    
    '------------------------------------------------------------------------------------------------------------------
ErrHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
EndHand:

End Function

Private Sub InitCommandBar()
    Dim strDLL As String
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    
    gstrSQL = " Select �²��� From ���²��� Where Nvl(����,0)=1"
    Call zldatabase.OpenRecordset(rsTemp, gstrSQL, "��ȡ���²���")
    If Err <> 0 Then
        mblnTemparatureChat = True
    Else
        If rsTemp.RecordCount = 0 Then
            mblnTemparatureChat = True
        Else
            If rsTemp!�²��� = "zl9TemperatureChart" Then
                mblnTemparatureChat = True
            Else
                mblnTemparatureChat = False
            End If
        End If
    End If
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.VisualTheme = xtpThemeOffice2003
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = imgPublic.Icons
    cbsMain.ActiveMenuBar.Visible = False
    
End Sub

Private Sub cboBaby_Click()
    mlng��� = cboBaby.ItemData(cboBaby.ListIndex)
    If mintBaby = mlng��� Then Exit Sub
    mintBaby = mlng���
'    mblnRefresh = True
    If mblnNoRefresh Then Exit Sub
    mblnNoRefresh = True
    Call zlRefData
    Call ExecuteCommand("��ʾ�ļ�����", vfgFile.ROW)
    mblnNoRefresh = False
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngLoop As Long, intNORule As Integer
    Dim DBeginTime As Date
    Dim lngFileID As Long
    Dim blnTrans As Boolean
    Dim ArrSQL()
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    Select Case Control.ID
        Case conMenu_Add
            If Val(Split(EprIsCommit(mlngPatiID, mlngPageId), "|")(0)) = 0 Then
                MsgBox "�ò��˵Ĳ������ύ���[״̬��" & gstrMecState & "]����������ļ�����ȡ���������ԣ�", vbInformation, gstrSysName
                Exit Sub
            End If
            If frmNurseFileEdit.ShowEditor(mlngPatiID, mlngPageId, mintBaby, mlngDeptId, "", 0, lngFileID) Then
                mintBaby = -1: mblnNoRefresh = False
                mlngFileID = lngFileID: mlngFormatID = 0
                cboBaby_Click
            End If
        Case conMenu_Modify
            lngFileID = Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID))
            If lngFileID = 0 Then Exit Sub
            
            If Val(Split(EprIsCommit(mlngPatiID, mlngPageId), "|")(2)) = 0 Then
                MsgBox "�ò��˵Ĳ������ύ���[״̬��" & gstrMecState & "]�������޸��ļ�����ȡ���������ԣ�", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If frmNurseFileEdit.ShowEditor(mlngPatiID, mlngPageId, mintBaby, mlngDeptId, "", lngFileID) Then
                mintBaby = -1: mblnNoRefresh = False
                mlngFileID = lngFileID: mlngFormatID = 0
                cboBaby_Click
            End If
        Case conMenu_Delete
            lngFileID = Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID))
            If lngFileID = 0 Then Exit Sub
            
            If Val(Split(EprIsCommit(mlngPatiID, mlngPageId), "|")(1)) = 0 Then
                MsgBox "�ò��˵Ĳ������ύ���[״̬��" & gstrMecState & "]������ɾ���ļ�����ȡ���������ԣ�", vbInformation, gstrSysName
                Exit Sub
            End If
            
            '91844,���ӻ�����ϸ��
            If Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.f����)) = -1 Then
                gstrSQL = "SELECT A.ID,B.��ʼʱ��" & _
                    " FROM ���˻������� A, ���˻����ļ� B,���˻�����ϸ C" & _
                    " Where  a.�ļ�id = b.Id And b.Id = [1]  And c.��¼id = a.Id And Rownum < 2"
            Else
                gstrSQL = "SELECT A.ID,B.��ʼʱ��" & _
                    " FROM ���˻������� A,���˻����ӡ C,���˻����ļ� B,���˻�����ϸ D" & _
                    " WHERE B.ID=[1] And A.�ļ�ID=B.ID and A.�ļ�ID=C.�ļ�ID And A.ID=C.��¼ID And d.��¼Id=a.Id And RowNum<2"
            End If
            Call SQLDIY(gstrSQL)
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "����Ƿ��������", lngFileID)
            If rsTemp.RecordCount > 0 Then
                MsgBox "���ļ��Ѿ������������ݲ�����ɾ��,���飡", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.f����)) = -1 Then
                DBeginTime = CDate(Format(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.f��ʼ����), "YYYY-MM-DD HH:mm:ss"))
                gstrSQL = " Select A.ID,A.��ʼʱ��" & _
                    " From ���˻����ļ� A,�����ļ��б� B" & _
                    " Where A.��ʽID=B.ID And A.����ID=[1] And A.��ҳID=[2] And A.Ӥ��=[3] And B.����=-1 order by A.��ʼʱ�� DESC"
                Call SQLDIY(gstrSQL)
                Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ��Ѷ������µ�", mlngPatiID, mlngPageId, mintBaby)
                rsTemp.Filter = "��ʼʱ��> '" & CStr(DBeginTime) & "'"
                If rsTemp.RecordCount > 0 Then
                    MsgBox "���ļ�֮�󻹴������������µ��ļ�,�ļ�ֻ�ܴӺ���ǰɾ��,���飡", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            If MsgBox("��ȷ��Ҫɾ��" & vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.f�ļ�) & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            'If MsgBox("���ļ����еĻ�������Ҳ��һ��ɾ�������ٴ�ȷ���Ƿ�ɾ����", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            ArrSQL = Array()
            gstrSQL = "ZL_���˻����ļ�_DELETE(" & lngFileID & ")"
            ReDim Preserve ArrSQL(UBound(ArrSQL) + 1)
            ArrSQL(UBound(ArrSQL)) = gstrSQL
            
            If Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.f����)) = -1 Then
                rsTemp.Filter = "��ʼʱ��< '" & CStr(DBeginTime) & "'"
                rsTemp.Sort = "��ʼʱ�� DESC"
                If rsTemp.RecordCount > 0 Then
                    'ȡ����һ���µ��ļ��Ľ���ʱ��
                    gstrSQL = "ZL_���˻����ļ�_STATE(" & Val(rsTemp!ID) & ",1,NULL)"
                    ReDim Preserve ArrSQL(UBound(ArrSQL) + 1)
                    ArrSQL(UBound(ArrSQL)) = gstrSQL
                End If
            End If
            
            'ɾ�������¼��ʱ��������ļ�ҳ��˳������Ҫ������ļ�֮����ļ�ҳ��
            '�˴��������ļ����ںϲ��������(��Ϊɾ���Ѿ����ƣ������ļ�������ںϲ���Ϣ����ɾ��)
            intNORule = zldatabase.GetPara("�����ļ�ҳ�����", glngSys, 1255, 0)
            If InStr(1, ",-1,1,", "," & Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.f����)) & ",") = 0 And intNORule <> 0 Then
                
                gstrSQL = " Select id " & vbNewLine & _
                    " From (" & vbNewLine & _
                    "   With ���˻����ļ�_F1 As" & vbNewLine & _
                    "   (Select a.Id, a.����id, ��ʼʱ��, ����ʱ��" & vbNewLine & _
                    "   From ���˻����ļ� a, �����ļ��б� b" & vbNewLine & _
                    "   Where a.��ʽid = b.Id And b.���� = 3 And b.���� <> 1 And b.���� <> -1 And a.����id = [1] And a.��ҳid = [2] And Nvl(a.Ӥ��, 0) = [3])" & vbNewLine & _
                    "   Select Id" & vbNewLine & _
                    "   From (Select Id, ��ʼʱ��, ����ʱ��" & vbNewLine & _
                    "       From ���˻����ļ�_F1 a" & vbNewLine & _
                    "       Where Not Exists (Select 1 From ���˻����ļ�_F1 Where a.Id = ����id))" & vbNewLine & _
                    "   Where id<>[4] And (��ʼʱ��>[5] OR (��ʼʱ��=[5] And ����ʱ��>[6])) " & vbNewLine & _
                    "   Order by ��ʼʱ��)"
                Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ���ļ�֮��Ļ����ļ�", mlngPatiID, mlngPageId, mintBaby, lngFileID, _
                    CDate(Format(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.f��ʼ����), "YYYY-MM-DD HH:mm:ss")), CDate(Format(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.f��������), "YYYY-MM-DD HH:mm:ss")))
                If rsTemp.RecordCount > 0 Then
                    gstrSQL = "Zl_���˻����ӡ_Batchretrypage(" & rsTemp!ID & ",'1;0')"
                    ReDim Preserve ArrSQL(UBound(ArrSQL) + 1)
                    ArrSQL(UBound(ArrSQL)) = gstrSQL
                End If
            End If
            
            If UBound(ArrSQL) > 0 Then gcnOracle.BeginTrans: blnTrans = True
            For lngLoop = 0 To UBound(ArrSQL)
                If CStr(ArrSQL(lngLoop)) <> "" Then Call zldatabase.ExecuteProcedure(CStr(ArrSQL(lngLoop)), "�ļ�ɾ��")
            Next
            If UBound(ArrSQL) > 0 Then gcnOracle.CommitTrans: blnTrans = False
            
            mintBaby = -1: mblnNoRefresh = False
            mlngFileID = lngFileID: mlngFormatID = 0
            cboBaby_Click
    End Select
    
    Exit Sub
ErrHand:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_Add, conMenu_Modify, conMenu_Delete
            Control.Visible = (InStr(1, mstrPrivs, "�����ļ�����") > 0 And mblnDoctorStation = False And Not gblnMoved)
            Control.Enabled = (mlngPatiID > 0) And Not mblnArchive And Control.Visible And mblnEdit
            If Control.ID = conMenu_Modify And Control.Enabled = True Then
                Control.Enabled = (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0)
            ElseIf Control.ID = conMenu_Delete And Control.Enabled = True Then
                Control.Enabled = (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0) And (InStr(1, mstrPrivs, "�����ļ�ɾ��") <> 0)
            End If
    End Select
End Sub

Private Sub Form_Activate()
    RaiseEvent Activate
End Sub

Private Sub Form_Load()
    mblnInit = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picPane.Move 0, 0, Me.Width, Me.Height
    fra.Move 10, 10, Me.Width - 30, fra.Height
    vfgFile.Move 10, fra.Height + 10, Me.Width - 20, Me.Height - vfgFile.Top - 20
End Sub

Private Sub TimFresh_Timer()
    Dim blnFileChange As Boolean
    Dim lngFileID As Long
    Dim lngBaby As Long
    Dim i As Long
    
    If Not mblnInit Then Exit Sub
    If gobjBodyEditor Is Nothing Then Exit Sub
    On Error Resume Next
    Call gobjBodyEditor.zlFileChange(blnFileChange, lngFileID, lngBaby)
    If Err <> 0 Then Err.Clear
    If blnFileChange = False Then Exit Sub
    '�������µ�ѡ����ļ����¶�λ�ļ��������ļ��б�����µ�ѡ��һ��
    If cboBaby.ItemData(cboBaby.ListIndex) = lngBaby Then
        For i = vfgFile.FixedRows To vfgFile.Rows - 1
            If Val(vfgFile.TextMatrix(i, mCol.fID)) = lngFileID And Val(vfgFile.TextMatrix(i, mCol.f����)) = -1 Then
                Call vfgFile.Select(i, mCol.fID)
                Exit For
            End If
        Next i
    Else
        For i = 0 To cboBaby.ListCount - 1
           If lngBaby = cboBaby.ItemData(i) Then
               mlngFileID = lngFileID: mlngFormatID = 0
               cboBaby.ListIndex = i
               Exit For
           End If
        Next
    End If
End Sub

Private Sub vfgFile_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnNoRefresh Then Exit Sub
    If OldRow <> NewRow Then
        
        Call ExecuteCommand("��ʾ�ļ�����", NewRow)
'        DoEvents
'        On Error Resume Next
'        vfgFile.SetFocus
    End If
End Sub

Private Sub vfgFile_DblClick()
    Dim lng����ID As Long
    Dim intEdit As Integer
    
    If Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)) = 0 Then Exit Sub
    
    lng����ID = Val(Me.vfgFile.TextMatrix(vfgFile.ROW, mCol.f����ID))

    If Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) = -1 Then
        '���µ��鿴������ID;��ҳID;����ID;��Ժ;�༭;Ӥ��
        intEdit = 0
        If (InStr(1, ";" & mstrPrivs & ";", ";���µ���ͼ;") > 0 And mblnDoctorStation = False) Then
            If (mblnEdit And mlngPatiID > 0 And mblnArchive = False) Then
                intEdit = 1
            End If
        End If
        If Not CreateBodyEditor Then Exit Sub
        RaiseEvent ViewAnimalHeat(mlngPatiID & ";" & mlngPageId & ";" & mlngDeptId & ";" & Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)) & ";0;" & intEdit & ";" & mintBaby & ";1", 0, mstrPrivs, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)))
    ElseIf Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) = 1 Then
        '����ͼ�鿴
        intEdit = 0
        If (InStr(1, ";" & mstrPrivs & ";", ";����ͼ��ͼ;") > 0 And mblnDoctorStation = False) Then
            If (mblnEdit And mlngPatiID > 0 And mblnArchive = False) Then
                intEdit = 1
            End If
        End If
        If Not CreatePartogram Then Exit Sub
        RaiseEvent Viewpartogram(Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)) & ";" & mlngPatiID & ";" & mlngPageId & ";" & mlngDeptId & ";" & intEdit, 1, mstrPrivs, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)))
        
    Else
        With vfgFile
            RaiseEvent ViewFile(Val(.TextMatrix(.ROW, mCol.fID)), mlngPatiID, mlngPageId, mlngDeptId, mintBaby, False, mstrPrivs, mblnEdit, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)))
        End With
    End If

End Sub

Private Sub vfgFile_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call vfgFile_DblClick
End Sub

Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-19 15:16
    '����:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize
    '���о���ģ������Ŵ���С����
    If Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)) = 0 Then Exit Sub
    RaiseEvent ViewReSetFontSize(Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) + 1, bytSize)
End Sub

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-19 15:16
    '����:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim CtlFont  As StdFont
    Dim objCtrl As Control
    Dim bytSize As Byte
    Dim intCol As Integer
    bytSize = IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))
    
    Me.FontSize = mbytFontSize
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("Label")
            objCtrl.FontSize = mbytFontSize
            objCtrl.Height = TextHeight("��") + 20
        Case UCase("ComboBox")
           objCtrl.FontSize = mbytFontSize
        Case UCase("Frame")
           objCtrl.FontSize = mbytFontSize
        Case UCase("VSFlexGrid")
            objCtrl.FontSize = mbytFontSize
            For intCol = 0 To objCtrl.Cols
                Select Case intCol
                    Case mCol.f�ļ�, mCol.f��ʼ����, mCol.f����
                        objCtrl.ColWidth(intCol) = BlowUp(CDbl(objCtrl.ColWidth(intCol)))
                End Select
            Next intCol
        End Select
    Next
    fra.Height = cboBaby.Height + 200
    Call Form_Resize
End Sub

Private Function BlowUp(ByRef dblChange As Double) As Double
    '�Ŵ����壬��Ԫ����
    BlowUp = dblChange
    If mbytFontSize = 9 Or mbytFontSize = 0 Then Exit Function
    BlowUp = dblChange + (dblChange * 1 / 3)
End Function

Public Sub zlRefreshViewFile()
    If Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) <> 1 Then
        Call ExecuteCommand("��ʾ�ļ�����", vfgFile.ROW)
    End If
End Sub

Public Sub StartTimer(ByVal blnStart As Boolean)
    TimFresh.Enabled = blnStart
End Sub

Private Sub vfgFile_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    If Button = 2 Then
        Set cbrPopupBar = cbsMain.Add("�Ҽ��˵�", xtpBarPopup)
        cbrPopupBar.Title = "�Ҽ��˵�"
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Add, "����(&A)"): cbrPopupItem.IconId = 1
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Modify, "�޸�(&M)"):  cbrPopupItem.IconId = 2
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Delete, "ɾ��(&D)"): cbrPopupItem.IconId = 3
        
        cbrPopupBar.ShowPopup
    End If
End Sub
