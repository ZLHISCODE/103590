VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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
End
Attribute VB_Name = "frmDockInTend_TendList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

Private Enum mCol
    f��־ = 0: fID: f��ʽID: f�ļ�: f��ʼ����: f����ID: f����: f����
End Enum

Private mblnInit As Boolean
Private mstrPrivs As String                             '��ǰʹ���߶Ա�����(1255)��Ȩ�޴�
Private mlngPatiId As Long                              '����id
Private mlngPageId As Long                              '��ҳid
Private mlngDeptId As Long                              '��ǰ��������id���粡�˿��Һ͵�ǰ���Ҳ�һ�£����ܲ����鵵��Ĺ���
Private mblnEdit As Boolean                             '�Ƿ����������ͨ�����ϼ�������ݵ�ǰ���������Ƿ�ǰ���˲���������
Private mblnDoctorStation As Boolean

Private rsTemp As New ADODB.Recordset
Private mintBaby As Integer
Private mfrmMain As Object

Private mblnChange As Boolean                           '�޸ı�־
Private mblnSign As Boolean                             'ǩ����־
Private mblnArchive As Boolean                          '�鵵��־

'���ѿɷ���鿴���µ��뻤���¼���������,����ʽ�鿴��ʧȥ����,��д������
Public Event Activate()         '���°�ť��˵�
Public Event ViewFile(ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, _
    ByVal intBaby As Integer, ByVal blnChildForm As Boolean, ByVal strPrivs As String, ByVal blnEdit As Boolean)
Public Event ViewAnimalHeat(strPara As String, bytMode As Byte, strPrivs As String)
Public Event ShowData(intBaby As Integer, lngFile As Long, lngDept As Long, bytSel As Byte)                 '֪ͨ����ҳ��ˢ��
Public Event PrintDocument(ByVal bytKind As Byte, ByVal bytMode As Byte)
Public Event SaveDocument(blnSave As Boolean)                                                               '����ָ�����
Public Event SignDocument(blnOK As Boolean, blnVerify As Boolean)                                           '����ȡ��ǩ��
Public Event ArchiveDocument(blnOK As Boolean)                                                              '����ȡ���鵵

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
    Dim Rs As New ADODB.Recordset
    
    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_FileMan
        Call frmNurseFileMan.ShowEditor(mlngPatiId, mlngPageId, mintBaby)
    Case conMenu_File_Open
        With vfgFile
            strInfo = Val(.TextMatrix(.ROW, mCol.f����ID))
            If Val(.TextMatrix(.ROW, mCol.f����)) = -1 Then
                '���µ��鿴������ID;��ҳID;����ID;��Ժ;�༭;Ӥ��
                If Not CreateBodyEditor Then Exit Sub
                RaiseEvent ViewAnimalHeat(mlngPatiId & ";" & mlngPageId & ";" & mlngDeptId & ";" & Val(.TextMatrix(.ROW, mCol.fID)) & ";0;0;" & mintBaby & ";1", 0, mstrPrivs)
            Else
                RaiseEvent ViewFile(Val(.TextMatrix(.ROW, mCol.fID)), mlngPatiId, mlngPageId, mlngDeptId, mintBaby, False, mstrPrivs, True)
            End If
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview
        ''1-Ԥ��,2-��ӡ
        With vfgFile
            If Val(.TextMatrix(.ROW, mCol.f����)) = -1 Then
                RaiseEvent PrintDocument(1, 1)
            Else
                RaiseEvent PrintDocument(2, 1)
            End If
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Print
        With vfgFile
            If Val(.TextMatrix(.ROW, mCol.f����)) = -1 Then
                RaiseEvent PrintDocument(1, 2)
            Else
                RaiseEvent PrintDocument(2, 2)
            End If
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Excel
        With vfgFile
            If Val(.TextMatrix(.ROW, mCol.f����)) = -1 Then
                MsgBox "�Բ������µ���֧�������Excel��", vbInformation, gstrSysName
            Else
                RaiseEvent PrintDocument(2, 3)
            End If
        End With
    Case conMenu_Tool_Sign
        RaiseEvent SignDocument(True, False)
    Case conMenu_Tool_SignEarse
        RaiseEvent SignDocument(False, False)
    Case conMenu_Tool_SignAuditAffirm
        RaiseEvent SignDocument(True, True)
    Case conMenu_Tool_SignAuditCancel
        RaiseEvent SignDocument(False, True)
    Case conMenu_Edit_Archive * 10
        RaiseEvent ArchiveDocument(True)
    Case conMenu_Edit_UnArchive
        RaiseEvent ArchiveDocument(False)
    Case conMenu_Edit_Save
        RaiseEvent SaveDocument(True)
    Case conMenu_Edit_Transf_Cancle
        RaiseEvent SaveDocument(False)
    Case conMenu_File_PrintDayDetail    '����¼��
        Call frmTendFileMutilEditor.ShowMe(Me, mlngDeptId, mstrPrivs)
    End Select
    Exit Sub

errHand:
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
    Case conMenu_Edit_FileMan
        Control.Visible = (InStr(1, mstrPrivs, "�����ļ�����") > 0 And mblnDoctorStation = False And Not gblnMoved)
        Control.Enabled = (mlngPatiId > 0) And Not mblnArchive And Control.Visible
    Case conMenu_File_Open
        Control.Visible = True
        Control.Enabled = (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_PrintSet
        Control.Enabled = Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) = -1
    Case conMenu_File_Preview, conMenu_File_Print
        Control.Enabled = (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Excel
        Control.Enabled = (vfgFile.Rows > 1 And Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) <> -1)
    Case conMenu_File_ExportToXML, conMenu_File_RowPrint, conMenu_Edit_Audit, conMenu_Edit_Sort, _
        conMenu_Tool_Monitor, conMenu_Edit_Archive * 10 + 1
        Control.Visible = False
        Control.Enabled = False
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_PrintDayDetail
        Control.Enabled = (mblnEdit And mlngPatiId > 0)

        Control.Visible = (InStr(1, mstrPrivs, "�����¼�Ǽ�") > 0 And mblnDoctorStation = False And Not gblnMoved)
        If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "�����¼�Ǽ�") > 0)
    Case conMenu_Tool_Sign  'ǩ��
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "�����¼ǩ��") > 0) And Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) <> -1
        Control.Enabled = (mlngPatiId > 0) And (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0) And Not mblnArchive And Not mblnChange And Control.Visible
    Case conMenu_Tool_SignEarse  'ȡ��ǩ��
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "ȡ����¼ǩ��") > 0) And Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) <> -1
        Control.Enabled = (mlngPatiId > 0) And (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0) And Not mblnArchive And Not mblnChange And Control.Visible
    Case conMenu_Tool_SignAuditAffirm, conMenu_Tool_SignAuditCancel  '��ǩ,ȡ����ǩ
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "�����¼��ǩ") > 0) And Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) <> -1
        Control.Enabled = (mlngPatiId > 0) And (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0) And Not mblnArchive And Not mblnChange And Control.Visible
        If Control.ID = conMenu_Tool_SignAuditCancel And Control.Enabled Then
            Control.Enabled = (InStr(1, mstrPrivs, "ȡ����¼ǩ��") > 0)
        End If
    Case conMenu_Edit_Archive * 10 '�鵵
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "�����¼�鵵") > 0)
        Control.Enabled = (mlngPatiId > 0) And Not mblnArchive And Not mblnChange And Control.Visible
    Case conMenu_Edit_UnArchive  'ȡ���鵵
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "ȡ����¼�鵵") > 0)
        Control.Enabled = (mlngPatiId > 0) And mblnArchive And Not mblnChange And Control.Visible
    Case conMenu_Edit_Save  '����
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "�����¼�Ǽ�") > 0)
        Control.Enabled = mblnChange And Not mblnArchive And Control.Visible
    Case conMenu_Edit_Transf_Cancle  'ȡ��
        Control.Visible = Not mblnDoctorStation And Not gblnMoved
        Control.Enabled = mblnChange And Not mblnArchive And Control.Visible
    End Select
    
End Sub

Private Function zlRefData() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Err = 0
    On Error GoTo errHand
    '------------------------------------------------------------------------------------------------------------------
    '�����ļ�ˢ��
    
    With vfgFile
        .Rows = 2
        .Cols = 8
        .FixedCols = 1
        
        .TextMatrix(0, mCol.f��־) = ""
        .TextMatrix(0, mCol.fID) = "ID"
        .TextMatrix(0, mCol.f��ʽID) = "��ʽID"
        .TextMatrix(0, mCol.f�ļ�) = "�ļ�"
        .TextMatrix(0, mCol.f��ʼ����) = "��ʼ����"
        .TextMatrix(0, mCol.f����ID) = "����id"
        .TextMatrix(0, mCol.f����) = "����"
        .TextMatrix(0, mCol.f����) = "����"
        
        Set .Cell(flexcpPicture, 1, mCol.f��־) = Nothing
        .TextMatrix(1, mCol.fID) = ""
        .TextMatrix(1, mCol.f��ʽID) = ""
        .TextMatrix(1, mCol.f�ļ�) = ""
        .TextMatrix(1, mCol.f��ʼ����) = ""
        .TextMatrix(1, mCol.f����ID) = ""
        .TextMatrix(1, mCol.f����) = ""
        .TextMatrix(1, mCol.f����) = ""
        
        .ColWidth(mCol.f��־) = 270
        .ColWidth(mCol.fID) = 0: .ColWidth(mCol.f��ʽID) = 0: .ColWidth(mCol.f�ļ�) = 2000: .ColWidth(mCol.f��ʼ����) = 1200
        .ColWidth(mCol.f����ID) = 0: .ColWidth(mCol.f����) = 1200: .ColWidth(mCol.f����) = 0
    End With
    
    '--------------------------------------------------------------------------------------------------------------
    gstrSQL = "" & _
        " SELECT A.ID,A.��ʽID,A.����ID,C.���� AS ����,A.�ļ�����,A.��ʼʱ��,B.����" & vbNewLine & _
        " FROM ���˻����ļ� A,�����ļ��б� B,���ű� C" & vbNewLine & _
        " WHERE A.��ʽID=B.ID AND A.����ID=C.ID And A.����ID=[1] And A.��ҳID=[2] And A.Ӥ��=[3]" & _
        " ORDER BY B.����,A.��ʼʱ�� "
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId, mintBaby)
    
    With Me.vfgFile
        Do While Not rsTemp.EOF
            If Val(.TextMatrix(.Rows - 1, mCol.fID)) > 0 Then .AddItem ""
            If rsTemp!���� = -1 Then
                Set .Cell(flexcpPicture, .Rows - 1, mCol.f��־) = imgData.ListImages("����").Picture
            Else
                Set .Cell(flexcpPicture, .Rows - 1, mCol.f��־) = imgData.ListImages("��ͨ").Picture
            End If

            .TextMatrix(.Rows - 1, mCol.fID) = rsTemp!ID
            .TextMatrix(.Rows - 1, mCol.f��ʽID) = rsTemp!��ʽID
            .TextMatrix(.Rows - 1, mCol.f�ļ�) = rsTemp!�ļ�����
            .TextMatrix(.Rows - 1, mCol.f��ʼ����) = Format(rsTemp!��ʼʱ��, "yyyy-MM-dd")
            .TextMatrix(.Rows - 1, mCol.f����ID) = rsTemp!����ID
            .TextMatrix(.Rows - 1, mCol.f����) = rsTemp!����
            .TextMatrix(.Rows - 1, mCol.f����) = rsTemp!����
            
            rsTemp.MoveNext
        Loop
    End With

    zlRefData = True
    Exit Function
errHand:
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

Public Function RefreshData(ByVal lng����id As Long, ByVal lng��ҳid As Long, ByVal lngDeptID As Long, ByVal blnDoctorStation As Boolean, ByVal blnEdit As Boolean) As Boolean
    '******************************************************************************************************************
    '���ܣ�ˢ������
    '������
    '���أ�
    '******************************************************************************************************************
    Dim Rs As New ADODB.Recordset
    
    mblnInit = False
    mlngPatiId = lng����id
    mlngPageId = lng��ҳid
    mlngDeptId = lngDeptID
    mblnEdit = blnEdit And Not gblnMoved
    
    mblnDoctorStation = blnDoctorStation
    
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
    Dim Rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strtmp As String
    Dim strSQL As String
    Dim strNow As String
    Dim strNote As String
    Dim byt����ȼ� As Byte
    
    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ�ؼ�"
        
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
        
        If mlngPatiId <> 0 Then
            gstrSQL = "Select ����ת�� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
            Set Rs = zlDatabase.OpenSQLRecord(gstrSQL, "�ж������Ƿ�ת��", mlngPatiId, mlngPageId)
            gblnMoved = NVL(Rs!����ת��, 0) <> 0
        End If
        
        cboBaby.Clear
        cboBaby.AddItem "���˱���"
        gstrSQL = "Select a.���,Decode(a.Ӥ������,Null,b.����||'֮��'||Trim(To_Char(a.���,'9')),a.Ӥ������) As Ӥ������ From ������������¼ a,������Ϣ b Where a.����id=[1] And a.��ҳid=[2] And a.����id=b.����id Order By a.���"
        Set Rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngPatiId, mlngPageId)
        If Rs.BOF = False Then
            Do While Not Rs.EOF
                cboBaby.AddItem Rs("Ӥ������").Value
                Rs.MoveNext
            Loop
        End If
        cboBaby.ListIndex = 0
        cboBaby.Enabled = (cboBaby.ListCount > 1)
        
        Call zlRefData
        Call ExecuteCommand("��ʾ�ļ�����", vfgFile.ROW)
        
        mblnInit = True
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʾ�ļ�����"
        'todo:Ӧ�ô��ļ�ID,���ϳ���ֻ���ܸ�ʽID,��Ҫ�޸ĳ���
        RaiseEvent ShowData(mintBaby, Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)), mlngDeptId, Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) + 1)
    End Select

    ExecuteCommand = True

    GoTo EndHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
EndHand:

End Function

Private Sub cboBaby_Click()
    If mintBaby = cboBaby.ListIndex Then Exit Sub
    
    mintBaby = cboBaby.ListIndex
'    mblnRefresh = True
    
    Call zlRefData
    Call ExecuteCommand("��ʾ�ļ�����", vfgFile.ROW)
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

Private Sub vfgFile_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow Then
        
        Call ExecuteCommand("��ʾ�ļ�����", NewRow)
        DoEvents
        
        On Error Resume Next
        vfgFile.SetFocus
    End If
End Sub

Private Sub vfgFile_DblClick()
    Dim lng����ID As Long
    Dim intEdit As Integer

    lng����ID = Val(Me.vfgFile.TextMatrix(vfgFile.ROW, mCol.f����ID))

    If Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f����)) = -1 Then
        '���µ��鿴������ID;��ҳID;����ID;��Ժ;�༭;Ӥ��

        intEdit = 0
        If (InStr(1, mstrPrivs, "���µ���ͼ") > 0 And mblnDoctorStation = False) Then
            If (mblnEdit And mlngPatiId > 0 And mblnArchive = False) Then
                intEdit = 1
            End If
        End If

        RaiseEvent ViewAnimalHeat(mlngPatiId & ";" & mlngPageId & ";" & mlngDeptId & ";" & Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)) & ";0;" & intEdit & ";" & mintBaby & ";1", 0, mstrPrivs)
    Else
        With vfgFile
            RaiseEvent ViewFile(Val(.TextMatrix(.ROW, mCol.fID)), mlngPatiId, mlngPageId, mlngDeptId, mintBaby, False, mstrPrivs, True)
        End With
    End If

End Sub

Private Sub vfgFile_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call vfgFile_DblClick
End Sub
