VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAppCheck 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�������޸�"
   ClientHeight    =   8835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13980
   ControlBox      =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmAppCheck.frx":0000
   ScaleHeight     =   8835
   ScaleWidth      =   13980
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkProcedure 
      BackColor       =   &H8000000E&
      Caption         =   "������/��������Ч��"
      Height          =   495
      Left            =   9120
      TabIndex        =   16
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CheckBox chkParameters 
      BackColor       =   &H8000000E&
      Caption         =   "���������Ƶ�һ����"
      Height          =   375
      Left            =   11520
      TabIndex        =   15
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "�������޸�"
      Height          =   465
      Index           =   0
      Left            =   1080
      TabIndex        =   13
      Top             =   3600
      Width           =   1890
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "����ͬ�������"
      Height          =   465
      Index           =   1
      Left            =   1080
      TabIndex        =   12
      Top             =   4440
      Width           =   1890
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "������Ȩ������"
      Height          =   465
      Index           =   2
      Left            =   1080
      TabIndex        =   11
      Top             =   5400
      Width           =   1890
   End
   Begin VB.CheckBox chkIndex 
      BackColor       =   &H8000000E&
      Caption         =   "���������ռ��һ����"
      Height          =   465
      Left            =   3240
      TabIndex        =   10
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CheckBox chkReport 
      BackColor       =   &H8000000E&
      Caption         =   "��鵱ǰ�汾�б����Ƿ����"
      Height          =   465
      Left            =   6360
      TabIndex        =   9
      Top             =   3600
      Width           =   2775
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   13920
      TabIndex        =   1
      Top             =   8295
      Visible         =   0   'False
      Width           =   13980
      Begin MSComctlLib.ProgressBar pgbState 
         Height          =   180
         Left            =   135
         TabIndex        =   2
         Top             =   255
         Width           =   6180
         _ExtentX        =   10901
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar pgbProgress 
         Height          =   180
         Left            =   7080
         TabIndex        =   3
         Top             =   255
         Width           =   6180
         _ExtentX        =   10901
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ڼ��"
         Height          =   180
         Left            =   135
         TabIndex        =   5
         Top             =   60
         Width           =   810
      End
      Begin VB.Label lblProgress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ɣ�"
         Height          =   180
         Left            =   6840
         TabIndex        =   4
         Top             =   0
         Width           =   720
      End
      Begin VB.Line Linepgb 
         BorderColor     =   &H80000006&
         X1              =   6600
         X2              =   6600
         Y1              =   0
         Y2              =   720
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfSelSys 
      Height          =   1695
      Left            =   1080
      TabIndex        =   6
      Top             =   900
      Width           =   11175
      _cx             =   19711
      _cy             =   2990
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
      Editable        =   2
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
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   3240
      TabIndex        =   14
      Top             =   4200
      Width           =   7695
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   360
      Picture         =   "frmAppCheck.frx":803A
      Top             =   600
      Width           =   480
   End
   Begin VB.Label lblMainPath 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "ϵͳ��װĿ¼��C:\Appsoft"
      Height          =   180
      Left            =   1080
      TabIndex        =   8
      Tag             =   "C:\Appsoft"
      Top             =   660
      Width           =   2160
   End
   Begin VB.Label lblSel 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "���ġ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   3420
      TabIndex        =   7
      Top             =   660
      Width           =   540
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������޸�"
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
      Left            =   315
      TabIndex        =   0
      Top             =   120
      Width           =   1440
   End
End
Attribute VB_Name = "frmAppCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MSTR_COL = ",300,4;���,500,1;����,2000,1;��ǰ�汾,1400,1;�ű��汾,1400,1;�����ļ�,4050,1;������,0,1;�����,0,1;,400,4"
Private Enum SysSelCol
    Col_ѡ�� = 0
    Col_ϵͳ��� = 1
    Col_ϵͳ���� = 2
    Col_��ǰ�汾 = 3
    Col_�ű��汾 = 4
    Col_�����ļ� = 5
    Col_������ = 6
    Col_����� = 7
    Col_�հ� = 8
End Enum
Private mclsRunScript As New clsRunScript
Private mrsLocalFile As New ADODB.Recordset

Private mrsSequenceFromFile As ADODB.Recordset
Private mrsViewFromFile As ADODB.Recordset
Private mrsPackageFromFile As ADODB.Recordset
Private mrsFildFromFile As ADODB.Recordset
Private mrsConstraintFromFile As ADODB.Recordset
Private mrsIndexFromFile As ADODB.Recordset
Private mrsProcedureFromFile As ADODB.Recordset

Private mrsSequenceFromDB As ADODB.Recordset
Private mrsViewFromDB As ADODB.Recordset
Private mrsPackageFromDB As ADODB.Recordset
Private mrsFildFromDB As ADODB.Recordset
Private mrsConstraintFromDB As ADODB.Recordset
Private mrsIndexFromDB As ADODB.Recordset
Private mrsProcedureFromDB As ADODB.Recordset

Private mrsDataFromFile As ADODB.Recordset
Private mrsDataFromDB As ADODB.Recordset

Private mlngSysNum As Long
Private mlngShare As Long
Private mlngProgress As Long
Private mblnzlTables As Boolean

Private Sub cmdFunction_Click(Index As Integer)
    Dim rsProData As ADODB.Recordset
    Dim rsChooseSysInfo As ADODB.Recordset
    Dim lngConsuming As Long
    Dim strSQL As String
    Dim strTemp As String
    Dim cnTools As ADODB.Connection
    Dim lngProgress As Long
    Dim strOwner As String
    
    If MsgBox("""" & Split(cmdFunction(Index).Caption, "(")(0) & """�������������Ľ϶����Դ�ͻ��ѽϳ���ʱ�䣬Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    Select Case Index
        Case 0
            Set rsChooseSysInfo = CopyNewRec(Nothing, True, , _
                                Array("ϵͳ���", adDouble, 10, 0, "ϵͳ����", adVarChar, 50, Empty, _
                                      "��ǰ�汾", adVarChar, 20, Empty, "�ű��汾", adVarChar, 20, Empty, _
                                      "�����ļ�", adVarChar, 500, Empty, "������", adVarChar, 50, Empty, _
                                      "�����", adDouble, 10, 0))
            
            If CheckChoose(rsChooseSysInfo) = False Then Exit Sub
            '���谲װ�������ű�·����¼����ʼ��
            Set mrsLocalFile = IniFilePathRecordset

            Call InirsFile

            '���ݼ�¼����ʼ��
            Set mrsDataFromFile = InitDataRecordset
            
            mlngProgress = 3 + rsChooseSysInfo.RecordCount * 8
            picStatus.Visible = True
            Enabled = False
            Call ShowFinalPro(1)
            
            If CollectObj(rsChooseSysInfo) = False Then
                picStatus.Visible = False
                Enabled = True
                Exit Sub
            End If
            
            Set rsProData = InitProDataRecordset
            lngProgress = 3
            rsChooseSysInfo.MoveFirst
            Do While Not rsChooseSysInfo.EOF
                If strOwner <> rsChooseSysInfo!������ Then
                    strOwner = rsChooseSysInfo!������
                    Call CollectObjFromDB(strOwner, rsChooseSysInfo!ϵͳ���)
                    Call GainData(mrsSequenceFromFile, mrsViewFromFile, mrsPackageFromFile, mrsFildFromFile, mrsConstraintFromFile, mrsIndexFromFile, mrsProcedureFromFile, mrsDataFromFile, _
                        mrsSequenceFromDB, mrsViewFromDB, mrsPackageFromDB, mrsFildFromDB, mrsConstraintFromDB, mrsIndexFromDB, mrsProcedureFromDB, mrsDataFromDB, _
                        IIf(chkIndex.value = 1, True, False), IIf(chkReport.value = 1, True, False), mblnzlTables, IIf(chkProcedure.value = 1, True, False), _
                        IIf(chkParameters.value = 1, True, False))
                End If
                    Call CompareCheck(rsChooseSysInfo!ϵͳ���, rsChooseSysInfo!ϵͳ����, rsProData, lngProgress)
                DoEvents
                rsChooseSysInfo.MoveNext
            Loop
            
            picStatus.Visible = False
            Enabled = True
            rsProData.Filter = ""
            If rsProData.RecordCount > 0 Then
                Call frmAppChkRpt.ShowMe(lblMainPath.Tag, rsProData, mrsDataFromFile)
            Else
                MsgBox "δ�������޸��Ķ���"
            End If
            Call Release
        Case 1
            '������ǰ�����ߵ�ȫ������Ĺ���ͬ���('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION')
            gcnOracle.Execute "Zl_Createpubsynonyms", , adCmdStoredProc
            
            MsgBox "��������ͬ�����ɣ�", vbInformation, gstrSysName
        Case 2
            '����Ȩ������
            Set cnTools = GetConnection("ZLTOOLS")
            If cnTools Is Nothing Then Exit Sub
            Call ReGrantForTools(cnTools, , True)
            MsgBox "������Ȩ��������ɣ�", vbInformation, gstrSysName
    End Select
End Sub

Private Sub InirsFile()
'���ܣ���ʼ���ű����ݼ�¼��

    Set mrsSequenceFromFile = CopyNewRec(Nothing, True, , Array("ϵͳ���", adDouble, 10, 0, "SQL", adVarChar, 2000, Empty, "����", adVarChar, 100, Empty))
    Set mrsViewFromFile = CopyNewRec(Nothing, True, , Array("ϵͳ���", adDouble, 10, 0, "SQL", adVarChar, 2000, Empty, "����", adVarChar, 100, Empty))
    Set mrsPackageFromFile = CopyNewRec(Nothing, True, , Array("ϵͳ���", adDouble, 10, 0, "SQL", adVarChar, 2000, Empty, "����", adVarChar, 100, Empty, "STATUS", adVarChar, 20, Empty))
    Set mrsFildFromFile = CopyNewRec(Nothing, True, , Array("ϵͳ���", adDouble, 10, 0, "SQL", adVarChar, 2000, Empty, "����", adVarChar, 100, Empty, _
                        "�ֶ�", adVarChar, 200, Empty, "�ֶ�����", adVarChar, 20, Empty, "�ֶγ���", adVarChar, 10, Empty))
    Set mrsConstraintFromFile = CopyNewRec(Nothing, True, , Array("ϵͳ���", adDouble, 10, 0, "SQL", adVarChar, 2000, Empty, "����", adVarChar, 100, Empty, _
                        "����", adVarChar, 100, Empty, "�ֶ�", adVarChar, 200, Empty, "��ռ�", adVarChar, 20, Empty))
    Set mrsIndexFromFile = CopyNewRec(Nothing, True, , Array("ϵͳ���", adDouble, 10, 0, "SQL", adVarChar, 2000, Empty, "����", adVarChar, 100, Empty, _
                        "����", adVarChar, 100, Empty, "�ֶ�", adVarChar, 200, Empty, "��ռ�", adVarChar, 20, Empty))
    Set mrsProcedureFromFile = CopyNewRec(Nothing, True, , Array("ϵͳ���", adDouble, 10, 0, "SQL", adVarChar, 2000, Empty, "����", adVarChar, 100, Empty, "�ֶ�", adVarChar, 1000, Empty))
    
End Sub

Private Function CheckChoose(ByRef rsFileInfor As ADODB.Recordset) As Boolean
'���ܣ����������޸�ǰ�Ƿ�ѡϵͳ;��ѡ��ϵͳ�Ƿ���ڱ��������ļ�;��ǰ�û��Ƿ��ܹ��������ѡ��ϵͳ
'������rsFileInfor��������ѡϵͳ���б�����
    Dim i As Long
    Dim strFile As String
    Dim blnFile As Boolean
    Dim strOraVer As String
    Dim strLocalVer As String
    Dim varTemp As Variant
    Dim cnTools As ADODB.Connection
    
    blnFile = False
    With vsfSelSys
        For i = .FixedRows To .Rows - .FixedRows
            If .Cell(flexcpChecked, i, Col_ѡ��) = flexChecked Then
                If .TextMatrix(i, Col_�����ļ�) = "" Then
                    strFile = IIf(strFile = "", .TextMatrix(i, Col_ϵͳ����), strFile & "��" & .TextMatrix(i, Col_ϵͳ����))
                Else
                    strOraVer = VerFull(.TextMatrix(i, Col_��ǰ�汾))
                    strLocalVer = VerFull(.TextMatrix(i, Col_�ű��汾))
                    If Split(strOraVer, ".")(1) = Split(strLocalVer, ".")(1) Then
                        varTemp = Split(.TextMatrix(i, Col_��ǰ�汾), ".")
                        If strOraVer > strLocalVer Then
                            MsgBox .TextMatrix(i, Col_ϵͳ����) & "��ǰ�汾���ڽű��汾���޷����ж������޸������飡"
                            Exit Function
                        End If
                        If UBound(varTemp) > 2 Then
                            If strOraVer <> strLocalVer Then
                                MsgBox .TextMatrix(i, Col_ϵͳ����) & "Ϊ����sp�汾���ű��汾�����뵱ǰ�汾һ�£�"
                                Exit Function
                            End If
                        End If
                        If .TextMatrix(i, Col_ϵͳ����) = "������������" Then .TextMatrix(i, Col_������) = "ZLTOOLS"
                        rsFileInfor.AddNew Array("ϵͳ���", "ϵͳ����", "��ǰ�汾", "�ű��汾", "�����ļ�", "������", "�����"), Array(IIf(.TextMatrix(i, Col_ϵͳ���) = "", 0, .TextMatrix(i, Col_ϵͳ���)), _
                                .TextMatrix(i, Col_ϵͳ����), .TextMatrix(i, Col_��ǰ�汾), .TextMatrix(i, Col_�ű��汾), _
                                .TextMatrix(i, Col_�����ļ�), .TextMatrix(i, Col_������), IIf(.TextMatrix(i, Col_�����) = "", 0, .TextMatrix(i, Col_�����)))
                        blnFile = True
                    Else
                        MsgBox "�ű��汾�뵱ǰ�汾�Ĵ�汾��һ�£��޷����м�飡"
                        Exit Function
                    End If
                End If
            End If
            If i = .Rows - .FixedRows Then
                If strFile <> "" Then
                    MsgBox strFile & "�ı��������ļ������ڣ��޷����ж������޸������飡"
                    Exit Function
                End If
                If blnFile = False Then
                    MsgBox "û�й�ѡϵͳ���޷����ж������޸�����ѡ��"
                    Exit Function
                End If
            End If
        Next
    End With
    
    If rsFileInfor.RecordCount > 0 Then rsFileInfor.MoveFirst
    Do While Not rsFileInfor.EOF
        If gstrUserName = "ZLTOOLS" Then
            If rsFileInfor!ϵͳ���� <> "������������" Then
                MsgBox "ZLTOOLS�û�ֻ�ܼ������������ߣ������¹�ѡϵͳ���л��û���"
                Exit Function
            End If
        ElseIf gblnDBA Then
            
        Else
            If rsFileInfor!������ <> gstrUserName Then
                If rsFileInfor!ϵͳ���� = "������������" Then
                    If gcnTools Is Nothing Then
                        MsgBox gstrUserName & "����DBA�û���Ҳ����ZLTOOLS�û��������ӹ������û����ܶԹ����߽��м�飡"
                        Set gcnTools = GetConnection("ZLTOOLS")
                        If gcnTools Is Nothing Then
                            MsgBox "�������û�����ʧ�ܣ��޷����м�飡"
                            Exit Function
                        End If
                    End If
                Else
                    MsgBox gstrUserName & "����DBA�û���Ҳ����" & rsFileInfor!ϵͳ���� & "�������ߣ����ܽ��и�ϵͳ�ļ�飬�����¹�ѡϵͳ���л��û���"
                    Exit Function
                End If
            End If
        End If
        rsFileInfor.MoveNext
    Loop
    CheckChoose = True
End Function

Private Function CollectObj(ByVal rsChoose As ADODB.Recordset) As Boolean
'���ܣ�����ű����������ݿ��ȡ�Ķ�����Ϣ
'������rsChoose������ѡ��ϵͳ��Ϣ
    Dim varTemp As Variant
    Dim strBigVer As String
    Dim rsTemp As ADODB.Recordset
    Dim strOwner As String
    Dim strSQL As String
    Dim rsUpgrade As New ADODB.Recordset
    Dim strFilePath As String
    Dim strSPInfo As String

    strSQL = "Select Nvl(a.ϵͳ, 0) ϵͳ, Nvl(b.����, '������������') ϵͳ����, ����汾" & vbNewLine & _
            "From Zlupgrade a, Zlsystems b" & vbNewLine & _
            "Where Length(a.����汾) > 10 And a.��Ǩ��� = 0 And a.ϵͳ = b.���(+)" & vbNewLine & _
            "Order By ϵͳ"
    Set rsUpgrade = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "���ݿ���Ǩ���")
    Do While Not rsUpgrade.EOF
        strSPInfo = strSPInfo & rsUpgrade!ϵͳ���� & vbTab & rsUpgrade!����汾 & vbCrLf
        rsUpgrade.MoveNext
    Loop
    If rsUpgrade.RecordCount > 0 Then
        If MsgBox("��ǰ�汾�����������к�������sp����������ȷ������ϵͳ������sp�ļ������ٽ��м�飡" & vbCrLf & strSPInfo, vbDefaultButton2 + vbYesNo, "��ʾ") = vbNo Then
            Exit Function
        End If
    End If
    
    With rsChoose
        .MoveFirst
        Do While Not .EOF
            If CheckSetFile(!�����ļ�, !ϵͳ���) Then
                varTemp = Split(!��ǰ�汾, ".")
                If varTemp(2) <> 0 Then
                    strBigVer = varTemp(0) & "." & varTemp(1) & ".0"
                    '���ڹ�������(GetUpgradeFiles)��ȡ���ǵ�ǰ��ǰ�汾����ǰ�ű����ļ����������øú�����ȡ��ǰ���ݿ�Ĵ�汾�����ؽű����汾��Ȼ��ɾ���汾���ڵ�ǰ�汾�Ľű��ļ�·��
                    Set rsTemp = Nothing
                    Set rsTemp = GetUpgradeFiles(rsTemp, !ϵͳ���, strBigVer, !�����ļ�, , , , , , , False)
                    Call AddUpFile(rsTemp, !��ǰ�汾, !�����)
                End If
            Else
                picStatus.Visible = False
                Enabled = True
                CollectObj = False
                Exit Function
            End If
            pgbState.value = .AbsolutePosition / .RecordCount * 100
            DoEvents
            .MoveNext
        Loop
    End With
        
    Call ShowFinalPro(2)
    
    With mrsLocalFile
        .MoveFirst
        Do While Not .EOF
            lblStatus.Caption = "�����ռ��ű�������Ϣ��" & !FilePath
            mlngSysNum = !SystemNum
            mlngShare = IIf(IsNull(!�����) = True, 0, !�����)
            Call CollectObjFromFile
            Call DealSpeObj
            pgbState.value = .AbsolutePosition / .RecordCount * 100
            DoEvents
            .MoveNext
        Loop
    End With
    Call CollectDataFromDB
    Call ShowFinalPro(3)
    CollectObj = True
End Function

Private Sub DealSpeObj()
'���ܣ�ɾ����������Ķ�����optional�ű���
    Dim strFilter As String
    Dim strSQL As String
    
    If mlngSysNum = 100 Then
        If mrsLocalFile!Filename = "ZL1_10.35.30.SQL" Then
            strFilter = "����='Ӱ���ղ����' and �ֶ�='������' and ϵͳ���=" & mlngSysNum
            Call RecDelete(mrsFildFromFile, strFilter)
        ElseIf mrsLocalFile!Filename = "ZL1_10.35.60.SQL" Then
            strFilter = "����='Ӱ����ʱ��¼_IX_����' and ϵͳ���=" & mlngSysNum
            Call RecDelete(mrsIndexFromFile, strFilter)
        ElseIf mrsLocalFile!Filename = "ZL1_10.35.80.SQL" Then
            mrsFildFromFile.Filter = "����='���ѿ�Ŀ¼'"
            Do While Not mrsFildFromFile.EOF
                mrsFildFromFile!���� = "���ѿ���Ϣ"
                mrsFildFromFile!SQL = Replace(mrsFildFromFile!SQL, "���ѿ�Ŀ¼", "���ѿ���Ϣ")
                mrsFildFromFile.Update
                mrsFildFromFile.MoveNext
            Loop
            mrsConstraintFromFile.Filter = "���� like '���ѿ�Ŀ¼*'"
            Do While Not mrsConstraintFromFile.EOF
                mrsConstraintFromFile!���� = "���ѿ���Ϣ"
                mrsConstraintFromFile!���� = Replace(mrsConstraintFromFile!����, "���ѿ�Ŀ¼", "���ѿ���Ϣ")
                mrsConstraintFromFile!SQL = Replace(mrsConstraintFromFile!SQL, "���ѿ�Ŀ¼", "���ѿ���Ϣ")
                mrsConstraintFromFile.Update
                mrsConstraintFromFile.MoveNext
            Loop
            mrsIndexFromFile.Filter = "���� like '���ѿ�Ŀ¼*'"
            Do While Not mrsIndexFromFile.EOF
                mrsIndexFromFile!���� = "���ѿ���Ϣ"
                mrsIndexFromFile!���� = Replace(mrsIndexFromFile!����, "���ѿ�Ŀ¼", "���ѿ���Ϣ")
                mrsIndexFromFile!SQL = Replace(mrsIndexFromFile!SQL, "���ѿ�Ŀ¼", "���ѿ���Ϣ")
                mrsIndexFromFile.Update
                mrsIndexFromFile.MoveNext
            Loop
            mrsSequenceFromFile.Filter = "���� like '���ѿ�Ŀ¼*'"
            Do While Not mrsSequenceFromFile.EOF
                mrsSequenceFromFile!���� = Replace(mrsSequenceFromFile!����, "���ѿ�Ŀ¼", "���ѿ���Ϣ")
                mrsSequenceFromFile!SQL = Replace(mrsSequenceFromFile!SQL, "���ѿ�Ŀ¼", "���ѿ���Ϣ")
                mrsSequenceFromFile.Update
                mrsSequenceFromFile.MoveNext
            Loop
            
            mrsFildFromFile.Filter = "����='�����ѽӿ�Ŀ¼'"
            Do While Not mrsFildFromFile.EOF
                mrsFildFromFile!���� = "���ѿ����Ŀ¼"
                mrsFildFromFile!SQL = Replace(mrsFildFromFile!SQL, "�����ѽӿ�Ŀ¼", "���ѿ����Ŀ¼")
                mrsFildFromFile.Update
                mrsFildFromFile.MoveNext
            Loop
            mrsConstraintFromFile.Filter = "���� like '�����ѽӿ�Ŀ¼*'"
            Do While Not mrsConstraintFromFile.EOF
                mrsConstraintFromFile!���� = "���ѿ����Ŀ¼"
                mrsConstraintFromFile!���� = Replace(mrsConstraintFromFile!����, "�����ѽӿ�Ŀ¼", "���ѿ����Ŀ¼")
                mrsConstraintFromFile!SQL = Replace(mrsConstraintFromFile!SQL, "�����ѽӿ�Ŀ¼", "���ѿ����Ŀ¼")
                mrsConstraintFromFile.Update
                mrsConstraintFromFile.MoveNext
            Loop
            mrsIndexFromFile.Filter = "���� like '�����ѽӿ�Ŀ¼*'"
            Do While Not mrsIndexFromFile.EOF
                mrsIndexFromFile!���� = "���ѿ����Ŀ¼"
                mrsIndexFromFile!���� = Replace(mrsIndexFromFile!����, "�����ѽӿ�Ŀ¼", "���ѿ����Ŀ¼")
                mrsIndexFromFile!SQL = Replace(mrsIndexFromFile!SQL, "�����ѽӿ�Ŀ¼", "���ѿ����Ŀ¼")
                mrsIndexFromFile.Update
                mrsIndexFromFile.MoveNext
            Loop
            
            strFilter = "����='���ѿ���ֵ��¼'"
            Call RecDelete(mrsFildFromFile, strFilter)
            strFilter = "����='���ѿ���ֵ��¼' or ���� like '���ѿ���ֵ��¼*'"
            Call RecDelete(mrsConstraintFromFile, strFilter)
            strFilter = "����='���ѿ���ֵ��¼' or ���� like '���ѿ���ֵ��¼*'"
            Call RecDelete(mrsIndexFromFile, strFilter)
            strFilter = "���� like '���ѿ���ֵ��¼*'"
            Call RecDelete(mrsSequenceFromFile, strFilter)
            
            strFilter = "����='���˿��������'"
            Call RecDelete(mrsFildFromFile, strFilter)
            strFilter = "����='���˿��������' or ���� like '���˿��������*'"
            Call RecDelete(mrsConstraintFromFile, strFilter)
            strFilter = "����='���˿��������' or ���� like '���˿��������*'"
            Call RecDelete(mrsIndexFromFile, strFilter)
            
            'ɾ�����ѿ���Ϣ�е��ֶΣ����㷽ʽ���ɿ���ID����λ�����С���λ�ʺš��������
            strFilter = "(����='���ѿ���Ϣ' and �ֶ�='���㷽ʽ') or (����='���ѿ���Ϣ' and �ֶ�='�ɿ���ID') or (����='���ѿ���Ϣ' and �ֶ�='��λ������') or (����='���ѿ���Ϣ' and �ֶ�='��λ�ʺ�') or (����='���ѿ���Ϣ' and �ֶ�='�������')"
            Call RecDelete(mrsFildFromFile, strFilter)
            
            strFilter = "����='���ѿ���Ϣ' and ����='���ѿ���Ϣ_FK_�ɿ���ID'"
            Call RecDelete(mrsConstraintFromFile, strFilter)
            
            If mblnzlTables Then
                strFilter = "(����='�����ѽӿ�Ŀ¼' and ���='��Ŀ¼') or (����='���˿��������' and ���='��Ŀ¼') or (����='���ѿ���ֵ��¼' and ���='��Ŀ¼')"
                Call RecDelete(mrsDataFromFile, strFilter)
            End If
        End If
    ElseIf mlngSysNum = 0 Then
        If mrsLocalFile!Filename = "ZLUPGRADE10.35.30.SQL" Then
            strFilter = "����='ZLRPTRUNHISTORY' and �ֶ�='ִ����ԱID' and ϵͳ���=" & mlngSysNum
            Call RecDelete(mrsFildFromFile, strFilter)
            strFilter = "����='ZLREPORTS' and �ֶ�='ִ����ԱID' and ϵͳ���=" & mlngSysNum
            Call RecDelete(mrsFildFromFile, strFilter)
        ElseIf mrsLocalFile!Filename = "ZLUPGRADE10.35.90.SQL" Then
            strFilter = "����='ZLPERIODS' and ϵͳ���=" & mlngSysNum
            Call RecDelete(mrsFildFromFile, strFilter)
            strFilter = "����='ZLPERIODS' and ϵͳ���=" & mlngSysNum
            Call RecDelete(mrsDataFromFile, strFilter)
            strFilter = "���� like 'ZLPERIODS*' and ϵͳ���=" & mlngSysNum
            Call RecDelete(mrsConstraintFromFile, strFilter)
            strFilter = "���� like 'ZLPERIODS*' and ϵͳ���=" & mlngSysNum
            Call RecDelete(mrsIndexFromFile, strFilter)
        End If
    ElseIf mlngSysNum = 2100 Then
        If mrsLocalFile!Filename = "ZL21_10.35.10.SQL" Then
            strFilter = "����='���������Ա' and �ֶ�='ָ������ӡ'"
            Call RecDelete(mrsFildFromFile, strFilter)
        End If
    ElseIf mlngSysNum = 2200 Then
        If mrsLocalFile!Filename = "ZL22_10.35.80.SQL" Then
            strFilter = "����='ѪҺ��Ѫ����'"
            Call RecDelete(mrsFildFromFile, strFilter)
            strFilter = "����='ѪҺ��Ѫ����' or ���� like 'ѪҺ��Ѫ����*'"
            Call RecDelete(mrsConstraintFromFile, strFilter)
            strFilter = "����='ѪҺ��Ѫ����' or ���� like 'ѪҺ��Ѫ����*'"
            Call RecDelete(mrsIndexFromFile, strFilter)
            
            strFilter = "����='ѪҺ��Ѫ����'"
            Call RecDelete(mrsDataFromFile, strFilter)
        End If
    ElseIf mlngSysNum = 2400 Then
        If mrsLocalFile!Filename = "ZL24_10.35.60.SQL" Then
            '������SQL����������
'            strFilter = "(����='�������ʷ���_PK' and ϵͳ���=" & mlngSysNum & ") or (����='�������ʷ���_UQ_����' and ϵͳ���=" & mlngSysNum & ")"
'            Call RecDelete(mrsConstraintFromFile, strFilter)
            strSQL = "Alter Table �������ʷ��� Add Constraint �������ʷ���_PK Primary Key (����) Using Index Pctfree 5 Tablespace zl9indexhis"
            mrsConstraintFromFile.AddNew Array("ϵͳ���", "SQL", "����", "����", "�ֶ�", "��ռ�"), _
                            Array(mlngSysNum, strSQL, "�������ʷ���", "�������ʷ���_PK", "����", "ZL9INDEXHIS")
            mrsIndexFromFile.AddNew Array("ϵͳ���", "SQL", "����", "����", "�ֶ�", "��ռ�"), _
                            Array(mlngSysNum, strSQL, "�������ʷ���", "�������ʷ���_PK", "����", "ZL9INDEXHIS")
            strSQL = "Alter Table �������ʷ��� Add Constraint �������ʷ���_UQ_���� Unique (����) Using Index Pctfree 5 Tablespace zl9indexhis"
            mrsConstraintFromFile.AddNew Array("ϵͳ���", "SQL", "����", "����", "�ֶ�", "��ռ�"), _
                            Array(mlngSysNum, strSQL, "�������ʷ���", "�������ʷ���_UQ_����", "����", "ZL9INDEXHIS")
            mrsIndexFromFile.AddNew Array("ϵͳ���", "SQL", "����", "����", "�ֶ�", "��ռ�"), _
                            Array(mlngSysNum, strSQL, "�������ʷ���", "�������ʷ���_UQ_����", "����", "ZL9INDEXHIS")
        End If
    ElseIf mlngSysNum = 2600 Then
        If mrsLocalFile!Filename = "ZL26_10.35.60.SQL" Then
            '������SQL����������
'            strFilter = "(����='��������ѡ��_PK' and ϵͳ���=" & mlngSysNum & ") or (����='���ﲥ������_PK' and ϵͳ���=" & mlngSysNum & ")"
'            Call RecDelete(mrsConstraintFromFile, strFilter)
            strSQL = "Alter Table ��������ѡ�� Add Constraint ��������ѡ��_PK Primary Key (����Ŀ¼id,����) Using Index  Tablespace zl9IndexPss"
            mrsConstraintFromFile.AddNew Array("ϵͳ���", "SQL", "����", "����", "�ֶ�", "��ռ�"), _
                            Array(mlngSysNum, strSQL, "��������ѡ��", "��������ѡ��_PK", "����Ŀ¼ID,����", "ZL9INDEXPSS")
            mrsIndexFromFile.AddNew Array("ϵͳ���", "SQL", "����", "����", "�ֶ�", "��ռ�"), _
                            Array(mlngSysNum, strSQL, "��������ѡ��", "��������ѡ��_PK", "����Ŀ¼ID,����", "ZL9INDEXPSS")
            strSQL = "Alter Table ���ﲥ������ Add Constraint ���ﲥ������_PK Primary Key (����Ŀ¼ID,�������) Using Index  Tablespace zl9IndexPss"
            mrsConstraintFromFile.AddNew Array("ϵͳ���", "SQL", "����", "����", "�ֶ�", "��ռ�"), _
                            Array(mlngSysNum, strSQL, "���ﲥ������", "���ﲥ������_PK", "����Ŀ¼ID,�������", "ZL9INDEXPSS")
            mrsIndexFromFile.AddNew Array("ϵͳ���", "SQL", "����", "����", "�ֶ�", "��ռ�"), _
                            Array(mlngSysNum, strSQL, "���ﲥ������", "���ﲥ������_PK", "����Ŀ¼ID,�������", "ZL9INDEXPSS")
        End If
    ElseIf mlngSysNum = 300 Then
        If mlngShare <> 0 Then
            strFilter = "����='�Һ���Ŀ'"
            Call RecDelete(mrsFildFromFile, strFilter)
            strFilter = "���� like '�Һ���Ŀ*'"
            Call RecDelete(mrsConstraintFromFile, strFilter)
            strFilter = "���� like '�Һ���Ŀ*'"
            Call RecDelete(mrsIndexFromFile, strFilter)
            'ɾ������ʱ�����ڵ�ģ��͹�������
            strFilter = "(���='ģ��' and ���=1001 and ����='���Ź���' and ϵͳ���=300) or (���='ģ��' and ���=1002 and ����='��Ա����' and ϵͳ���=300) or (���='ģ��' and ���=1013 and ����='�����������' and ϵͳ���=300)" & _
                " or (���='����' and ϵͳ���=300 and ���=1001) or (���='����' and ϵͳ���=300 and ���=1002) or (���='����' and ϵͳ���=300 and ���=1013)"
            Call RecDelete(mrsDataFromFile, strFilter)
        End If
    ElseIf mlngSysNum = 400 Then
        strFilter = "(���='ģ��' and ���=1001 and ����='���Ź���' and ϵͳ���=400) or (���='ģ��' and ���=1002 and ����='��Ա����' and ϵͳ���=400) or (���='ģ��' and ���=1010 and ����='�ڼ仮�ֵ���' and ϵͳ���=400) or (���='ģ��' and ���=1025 and ����='�����̹���' and ϵͳ���=400)" & _
            " or (���='����' and ϵͳ���=400 and ���=1001) or (���='����' and ϵͳ���=400 and ���=1002) or (���='����' and ϵͳ���=400 and ���=1010) or (���='����' and ϵͳ���=400 and ���=1025)"
        Call RecDelete(mrsDataFromFile, strFilter)
    ElseIf mlngSysNum = 600 Then
        strFilter = "(���='ģ��' and ���=1001 and ����='���Ź���' and ϵͳ���=600) or (���='ģ��' and ���=1002 and ����='��Ա����' and ϵͳ���=600) or (���='ģ��' and ���=1010 and ����='�ڼ仮�ֵ���' and ϵͳ���=600) or (���='ģ��' and ���=1025 and ����='�����̹���' and ϵͳ���=600)" & _
            " or (���='����' and ϵͳ���=600 and ���=1001) or (���='����' and ϵͳ���=600 and ���=1002) or (���='����' and ϵͳ���=600 and ���=1010) or (���='����' and ϵͳ���=600 and ���=1025)"
        Call RecDelete(mrsDataFromFile, strFilter)
    End If
End Sub

Public Sub ShowProgress(ByRef strSysName As String, ByRef lngNum As Long, ByRef lngCurNum As Long, ByRef strObjType As String, Optional ByRef strName As String)
'���ܣ�ÿ���һ�������ʾ������
        
    lblStatus.Caption = IIf(strName = "", "���ڼ��" & strSysName & "��" & strObjType & "...", "���ڼ��" & strSysName & "��" & strObjType & "��" & strName)
    pgbState.value = lngCurNum / lngNum * 100
    If pgbState.value = 100 Then pgbState.value = 0
End Sub

Public Sub ShowFinalPro(ByRef lngNum As Long)
'���ܣ���ʾ�ܵĽ�����
    lblProgress.Caption = "����ɣ�" & Round(lngNum / mlngProgress * 100) & "%"
    pgbProgress.value = lngNum / mlngProgress * 100
    If pgbProgress.value = 100 Then
        pgbProgress.value = 0
        lblProgress.Caption = ""
    End If
End Sub

Private Sub AddUpFile(ByVal rsTemp As ADODB.Recordset, ByVal strCurrver As String, ByVal lngShare As Long)
'���ܣ���ȡ���������ű�·����ģ��ű�·����¼����
'������rsTemp�����ݹ���������ȡ�������ű�;strCurrver����ǰ��ǰ�汾
    Dim i As Long
    Dim strVer As String
    Dim strFilter As String

    strVer = VerFull(strCurrver)
    strFilter = "FullSPVer>" & strVer & " or FileName like '*HISTORY*' or FileName like '*OPTIONAL*'"
    Call RecDelete(rsTemp, strFilter)
    
    rsTemp.Filter = "": rsTemp.Sort = "FullSPVer Asc"
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
        mrsLocalFile.AddNew Array("FilePath", "SystemNum", "FileName", "FileType", "FullVer", "�����"), Array(rsTemp!FilePath, rsTemp!ϵͳ���, UCase(rsTemp!Filename), "�����ű�", rsTemp!FullSPVer, lngShare)
        rsTemp.MoveNext
    Next
End Sub

Private Sub CollectObjFromDB(ByRef strOwner As String, ByRef lngNum As Long)
'���ܣ���ȡ���ݿ�Ķ�����Ϣ
    Dim strSQL As String
    Dim cnChoose As New ADODB.Connection
    
    If gblnDBA Then
        strSQL = "select '����' ���,a.SEQUENCE_NAME ���� from Dba_SEQUENCES a where a.SEQUENCE_OWNER='" & strOwner & "'"
        Set mrsSequenceFromDB = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "����")
        strSQL = "select '��ͼ' ���,a.Object_Name ���� from Dba_Objects a where  a.OBJECT_TYPE Like 'VIEW' and a.owner ='" & strOwner & "'"
        Set mrsViewFromDB = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ͼ")
        strSQL = "select '��' ���,a.Object_Name ����,a.STATUS from Dba_Objects a where a.OBJECT_TYPE in('PACKAGE') and OWNER='" & strOwner & "'"
        Set mrsPackageFromDB = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��")
        strSQL = "select '�ֶ�' ���, a.TABLE_NAME ����,a.COLUMN_NAME ����,a.COLUMN_NAME �ֶ�,a.DATA_TYPE �ֶ�����,a.DATA_LENGTH �ֶγ���,a.DATA_PRECISION �ֶ�ʵ�ʳ���,a.DATA_SCALE �ֶ�С������ From DBA_TAB_COLUMNS a WHERE a.OWNER='" & strOwner & "'"
        Set mrsFildFromDB = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "�ֶ�")
        strSQL = "Select ���, ����, ����, f_List2str(Cast(Collect(�ֶ� Order By Position) As t_Strlist)) �ֶ�, Status" & vbNewLine & _
                "From (Select 'Լ��' ���, a.Table_Name ����, a.Constraint_Name ����, b.Column_Name �ֶ�, b.Position Position, a.Status" & vbNewLine & _
                "       From Dba_Constraints a, Dba_Cons_Columns b" & vbNewLine & _
                "       Where a.Owner = b.Owner And a.Constraint_Name = b.Constraint_Name And a.Constraint_Type In ('R', 'P', 'U') And" & vbNewLine & _
                "             a.Owner = '" & strOwner & "'" & vbNewLine & _
                "       Order By a.Constraint_Name, b.Position)" & vbNewLine & _
                "Group By ���, ����, ����, Status"
        Set mrsConstraintFromDB = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "Լ��")
        strSQL = "Select ���, ����, ����, f_List2str(Cast(Collect(�ֶ� Order By Position) As t_Strlist)) �ֶ�, ��ռ�, Uniqueness, Status" & vbNewLine & _
                "From (Select '����' ���, a.Table_Name ����, a.Index_Name ����, a.Column_Name �ֶ�, b.Tablespace_Name ��ռ�, b.Uniqueness Uniqueness," & vbNewLine & _
                "              b.Status, a.Column_Position Position" & vbNewLine & _
                "       From All_Ind_Columns a, Dba_Indexes b" & vbNewLine & _
                "       Where a.Index_Name = b.Index_Name And a.Index_Owner = b.Owner And a.Index_Name Not Like '%$%' And" & vbNewLine & _
                "             b.Owner ='" & strOwner & "'" & vbNewLine & _
                "       Order By a.Index_Name, a.Column_Position)" & vbNewLine & _
                "Group By ���, ����, ����, ��ռ�, Uniqueness, Status"
        Set mrsIndexFromDB = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "����")
        strSQL = "Select ���, ����,f_List2str(Cast(Collect(�ֶ� Order By Position) As t_Strlist)) �ֶ�,Status" & vbNewLine & _
                "From(Select '����/����' ���, b.Object_Name ����, a.Argument_Name �ֶ�,b.Status,a.Position Position" & vbNewLine & _
                "From Dba_Arguments a, Dba_Objects b" & vbNewLine & _
                "Where a.Package_Name Is Null And a.Object_Id(+) = b.Object_Id And" & vbNewLine & _
                "b.Object_Type In ('FUNCTION', 'PROCEDURE') And Not (a.Argument_Name Is Null And a.Data_Type Is Not Null) And" & vbNewLine & _
                "b.Owner ='" & strOwner & "'" & vbNewLine & _
                "Order By b.Object_Name,a.Position)" & vbNewLine & _
                "Group By  ����,���,Status"
        Set mrsProcedureFromDB = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "����/����")
    Else
        If lngNum = 0 And gstrUserName <> "ZLTOOLS" Then
            Set cnChoose = gcnTools
        Else
            Set cnChoose = gcnOracle
        End If
        strSQL = "select '����' ���,a.SEQUENCE_NAME ���� from User_SEQUENCES a"
        Set mrsSequenceFromDB = gclsBase.OpenSQLRecord(cnChoose, strSQL, "����")
        strSQL = "select '��ͼ' ���,a.Object_Name ���� from User_Objects a where  a.OBJECT_TYPE Like 'VIEW'"
        Set mrsViewFromDB = gclsBase.OpenSQLRecord(cnChoose, strSQL, "��ͼ")
        strSQL = "select '��' ���,a.Object_Name ����,a.STATUS from User_Objects a where a.OBJECT_TYPE in('PACKAGE')"
        Set mrsPackageFromDB = gclsBase.OpenSQLRecord(cnChoose, strSQL, "��")
        strSQL = "select '�ֶ�' ���, a.TABLE_NAME ����,a.COLUMN_NAME ����,a.COLUMN_NAME �ֶ�,a.DATA_TYPE �ֶ�����,a.DATA_LENGTH �ֶγ���,a.DATA_PRECISION �ֶ�ʵ�ʳ���,a.DATA_SCALE �ֶ�С������ From User_TAB_COLUMNS a"
        Set mrsFildFromDB = gclsBase.OpenSQLRecord(cnChoose, strSQL, "�ֶ�")
        strSQL = "Select ���, ����, ����, f_List2str(Cast(Collect(�ֶ� Order By Position) As t_Strlist)) �ֶ�, Status" & vbNewLine & _
                "From (Select 'Լ��' ���, a.Table_Name ����, a.Constraint_Name ����, b.Column_Name �ֶ�, b.Position Position, a.Status" & vbNewLine & _
                "       From User_Constraints a, User_Cons_Columns b" & vbNewLine & _
                "       Where a.Constraint_Name = b.Constraint_Name And a.Constraint_Type In ('R', 'P', 'U')" & vbNewLine & _
                "       Order By a.Constraint_Name, b.Position)" & vbNewLine & _
                "Group By ���, ����, ����, Status"
        Set mrsConstraintFromDB = gclsBase.OpenSQLRecord(cnChoose, strSQL, "Լ��")
        strSQL = "Select ���, ����, ����, f_List2str(Cast(Collect(�ֶ� Order By Position) As t_Strlist)) �ֶ�, ��ռ�, Uniqueness, Status" & vbNewLine & _
                "From (Select '����' ���, a.Table_Name ����, a.Index_Name ����, a.Column_Name �ֶ�, b.Tablespace_Name ��ռ�, b.Uniqueness Uniqueness," & vbNewLine & _
                "              b.Status, a.Column_Position Position" & vbNewLine & _
                "       From User_Ind_Columns a, User_Indexes b" & vbNewLine & _
                "       Where a.Index_Name = b.Index_Name And a.Index_Name Not Like '%$%' " & vbNewLine & _
                "       Order By a.Index_Name, a.Column_Position)" & vbNewLine & _
                "Group By ���, ����, ����, ��ռ�, Uniqueness, Status"
        Set mrsIndexFromDB = gclsBase.OpenSQLRecord(cnChoose, strSQL, "����")
        strSQL = "Select ���, ����,f_List2str(Cast(Collect(�ֶ� Order By Position) As t_Strlist)) �ֶ�,Status" & vbNewLine & _
                "From(Select '����/����' ���, b.Object_Name ����, a.Argument_Name �ֶ�,b.Status,a.Position Position" & vbNewLine & _
                "From User_Arguments a, User_Objects b" & vbNewLine & _
                "Where a.Package_Name Is Null And a.Object_Id(+) = b.Object_Id And" & vbNewLine & _
                "b.Object_Type In ('FUNCTION', 'PROCEDURE') And Not (a.Argument_Name Is Null And a.Data_Type Is Not Null) " & vbNewLine & _
                "Order By b.Object_Name,a.Position)" & vbNewLine & _
                "Group By  ����,���,Status"
        Set mrsProcedureFromDB = gclsBase.OpenSQLRecord(cnChoose, strSQL, "����/����")
    End If
End Sub

Private Sub CollectDataFromDB()
'���ܣ����ݿ�������ݱ���
    Dim strSQL As String
    
    If mblnzlTables Then
        strSQL = "Select 'ģ��' ���, Nvl(ϵͳ, 0) ϵͳ���, ���, ���� ����, Null ������, Null ������ From Zlprograms Union All" & vbNewLine & _
                "Select '����' ���, Nvl(ϵͳ, 0) ϵͳ���, ���, ���� ����, Null ������, Null ������ From Zlprogfuncs Union All" & vbNewLine & _
                "Select '����' ���, Nvl(ϵͳ, 0) ϵͳ���, Null ���, Nvl(ģ�� || '', 'NULL') ����, ������, Upper(������) ������ From Zlparameters Union All" & vbNewLine & _
                "Select '����' ���, Nvl(ϵͳ, 0) ϵͳ���, Null ���, ��� ����, Null ������, Null ������ From Zlreports Union All" & vbNewLine & _
                "Select '��Ŀ¼' ���, ϵͳ ϵͳ���, Null ���, ���� ����, Null ������, Null ������ From Zltables"
    Else
        strSQL = "Select 'ģ��' ���, Nvl(ϵͳ, 0) ϵͳ���, ���, ���� ����, Null ������, Null ������ From Zlprograms Union All" & vbNewLine & _
            "Select '����' ���, Nvl(ϵͳ, 0) ϵͳ���, ���, ���� ����, Null ������, Null ������ From Zlprogfuncs Union All" & vbNewLine & _
            "Select '����' ���, Nvl(ϵͳ, 0) ϵͳ���, Null ���, Nvl(ģ�� || '', 'NULL') ����, ������, Upper(������) ������ From Zlparameters Union All" & vbNewLine & _
            "Select '����' ���, Nvl(ϵͳ, 0) ϵͳ���, Null ���, ��� ����, Null ������, Null ������ From Zlreports"
    End If
    lblStatus.Caption = "�����ռ����ݿ�Ļ���������Ϣ..."
    Set mrsDataFromDB = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��������")
    
End Sub

Private Sub CollectObjFromFile()
    '�洢���ؽű������ݶ���
    Dim i As Long
    Dim lngSys As Long
    Dim strSQL As String
    Dim strTemp As String
    Dim varTemp As Variant
    Dim varFild As Variant
    Dim strName As String
    Dim objText As TextStream
    Dim strTableName As String
    Dim strFild As String
    Dim strFildType As String
    Dim strReFild As String
    Dim strTableSpace As String
    Dim strFildLength As String
    Dim rsTemp As ADODB.Recordset
    
    If mrsLocalFile!Filename = "ZLSEQUENCE.SQL" Then
        Set objText = gobjFile.OpenTextFile(mrsLocalFile!FilePath, ForReading)
        Do While Not objText.AtEndOfStream
            strSQL = objText.ReadLine
            If strSQL <> "" And Mid(strSQL, 1, 2) <> "--" Then
                strSQL = iniSQL(strSQL)
                If strSQL Like "CREATE SEQUENCE*" Then
                    strName = Trim(Replace(strSQL, "CREATE SEQUENCE", ""))
                    strName = Mid(strName, 1, InStr(strName, " ") - 1)
                    mrsSequenceFromFile.AddNew Array("ϵͳ���", "SQL", "����"), Array(mlngSysNum, strSQL, strName)
                End If
            End If
        Loop
        objText.Close
    ElseIf mrsLocalFile!Filename = "ZLTABLE.SQL" Then
        If mclsRunScript.OpenFile(mrsLocalFile!FilePath) Then
            Do While Not mclsRunScript.EOF
                strSQL = iniSQL(mclsRunScript.SQLInfo.SQL)
                If strSQL Like "CREATE TABLE*" Then
                    strSQL = Replace(strSQL, "NUMERIC", "NUMBER")
                    strTemp = Trim(Replace(strSQL, "CREATE TABLE", ""))
                    strTableName = Trim(Replace(Mid(strTemp, 1, InStr(strTemp, "(") - 1), vbCrLf, ""))
                    strTemp = Replace(strSQL, vbCrLf, "||")
                    If InStr(strTemp, "))") > 0 Then
                        strFild = Mid(strTemp, InStr(strTemp, "(") + 1, InStr(strTemp, "))") - InStr(strTemp, "("))
                    ElseIf InStr(strTemp, "||)") > 0 Then
                        strFild = Mid(strTemp, InStr(strTemp, "(") + 1, InStr(strTemp, "||)") - InStr(strTemp, "("))
                    ElseIf InStr(strTemp, ")||") > 0 Then
                        strFild = Mid(strTemp, InStr(strTemp, "(") + 1, InStr(strTemp, ")||") - InStr(strTemp, "("))
                    Else
                        strFild = Mid(strTemp, InStr(strTemp, "(") + 1)
                    End If
                    varTemp = Split(strFild, "||")
                    For i = LBound(varTemp) To UBound(varTemp)
                        varTemp(i) = Trim(varTemp(i))
                        If InStr(varTemp(i), "TABLESPACE") > 0 Then Exit For
                        If varTemp(i) <> "" And varTemp(i) <> ")" And InStr(varTemp(i), "TABLESPACE") = 0 Then
                            strFildType = ""
                            strFildLength = ""
                            strFild = TrimEx(Mid(varTemp(i), 1, InStr(varTemp(i), " ")))
                            strTemp = Trim(Mid(varTemp(i), InStr(varTemp(i), " ") + 1))
                            If InStr(strTemp, "DATE") > 0 Then
                                strFildType = "DATE"
                            ElseIf InStr(strTemp, "LONG RAW") > 0 Then
                                strFildType = "LONG RAW"
                            Else
                                If InStr(strTemp, ")") > 0 Then
                                    strTemp = Trim(Mid(strTemp, 1, InStr(strTemp, ")") - 1))
                                ElseIf InStr(strTemp, " ") > 0 Then
                                    strTemp = Trim(Mid(strTemp, 1, InStr(strTemp, " ") - 1))
                                End If
                                If InStr(strTemp, "(") > 0 Then
                                    strFildType = Mid(strTemp, 1, InStr(strTemp, "(") - 1)
                                    strFildLength = Mid(strTemp, InStr(strTemp, "(") + 1)
                                ElseIf InStr(strTemp, ",") > 0 Then
                                    strFildType = Mid(strTemp, 1, Len(strTemp) - 1)
                                ElseIf InStr(strTemp, ")") > 0 Then
                                    strFildType = Mid(strTemp, 1, InStr(strTemp, ")") - 1)
                                Else
                                    strFildType = strTemp
                                End If
                                strFildType = Trim(Replace(strFildType, "|", ""))
                                '�ֶε����ƺ��ֶ��и�����ͬ��ֵ
                            End If
                            If strFild <> "" Then
                                mrsFildFromFile.AddNew Array("ϵͳ���", "SQL", "����", "�ֶ�", "�ֶ�����", "�ֶγ���"), _
                                                    Array(mlngSysNum, strSQL, strTableName, strFild, strFildType, strFildLength)
                            End If
                        End If
                    Next
                End If
                DoEvents
                Call mclsRunScript.ReadNextSQL
            Loop
        End If
    ElseIf mrsLocalFile!Filename = "ZLCONSTRAINT.SQL" Then
        Set objText = gobjFile.OpenTextFile(mrsLocalFile!FilePath, ForReading)
        Do While Not objText.AtEndOfStream
            strSQL = objText.ReadLine
            If strSQL <> "" And Mid(strSQL, 1, 2) <> "--" Then
            strSQL = iniSQL(strSQL)
                If strSQL Like "ALTER TABLE * ADD CONSTRAINT*" Or strSQL Like "ALTER TABLE * MODIFY * CONSTRAINT*" Then
                    varTemp = Split(strSQL, "CONSTRAINT")
                    strTemp = Trim(Replace(varTemp(1), "CONSTRAINT", ""))
                    '��ȡԼ������
                    strName = TrimEx(Mid(strTemp, 1, InStr(strTemp, " ") - 1))
                    strTemp = Trim(Replace(varTemp(0), "ALTER TABLE", ""))
                    '��ȡ����
                    strTableName = Trim(Mid(strTemp, 1, InStr(strTemp, " ")))
                    If InStr(strSQL, "ADD") > 0 Then
                        '��ȡԼ���ֶ�
                        strFild = Replace(Trim(Mid(strSQL, InStr(strSQL, "(") + 1, InStr(strSQL, ")") - InStr(strSQL, "(") - 1)), " ", "")
                        strTableSpace = GetTableSpace(strSQL)
                        If InStr(strSQL, "NOVALIDATE") = 0 Then strSQL = strSQL & " NOVALIDATE"
                        mrsConstraintFromFile.AddNew Array("ϵͳ���", "SQL", "����", "����", "�ֶ�", "��ռ�"), _
                                            Array(mlngSysNum, strSQL, strTableName, strName, strFild, strTableSpace)
                        If InStr(strSQL, "PRIMARY") > 0 Or InStr(strSQL, "UNIQUE") > 0 Then
                            If strTableSpace <> "" Then
                                strTemp = "Create Unique Index " & strName & " On " & strTableName & "(" & strFild & ") Tablespace " & strTableSpace & " Nologging"
                                strTemp = strTemp & "||" & strSQL
                            Else
                                strTemp = "Create Unique Index " & strName & " On " & strTableName & "(" & strFild & ") Nologging"
                                strTemp = strTemp & "||" & strSQL
                            End If
                            mrsIndexFromFile.AddNew Array("ϵͳ���", "SQL", "����", "����", "�ֶ�", "��ռ�"), _
                                Array(mlngSysNum, strTemp, strTableName, strName, strFild, strTableSpace)
                        End If
                    End If
                End If
            End If
        Loop
        objText.Close
    ElseIf mrsLocalFile!Filename = "ZLINDEX.SQL" Then
        Set objText = gobjFile.OpenTextFile(mrsLocalFile!FilePath, ForReading)
        Do While Not objText.AtEndOfStream
            strSQL = objText.ReadLine
            If strSQL <> "" And Mid(strSQL, 1, 2) <> "--" Then
                strSQL = iniSQL(strSQL)
                If strSQL Like "CREATE INDEX*" Then
                    varTemp = Split(strSQL, "ON")
                    strName = Trim(Replace(varTemp(0), "CREATE INDEX", ""))
                    strTableName = Trim(Mid(varTemp(1), 1, InStr(varTemp(1), "(") - 1))
                    strFild = Replace(Mid(varTemp(1), InStr(varTemp(1), "(") + 1, InStrRev(varTemp(1), ")") - InStr(varTemp(1), "(") - 1), " ", "")
                    strTableSpace = GetTableSpace(strSQL)
                    If InStr(strSQL, "NOLOGGING") = 0 Then strSQL = strSQL & " NOLOGGING"
                    mrsIndexFromFile.AddNew Array("ϵͳ���", "SQL", "����", "����", "�ֶ�", "��ռ�"), _
                        Array(mlngSysNum, strSQL, strTableName, strName, strFild, strTableSpace)
                End If
            End If
        Loop
        objText.Close
    ElseIf mrsLocalFile!Filename = "ZLVIEW.SQL" Then
        If mclsRunScript.OpenFile(mrsLocalFile!FilePath) Then
            Do While Not mclsRunScript.EOF
                strSQL = iniSQL(mclsRunScript.SQLInfo.SQL)
                If strSQL <> "" And Mid(strSQL, 1, 2) <> "--" Then
                    If strSQL Like "CREATE OR REPLACE VIEW*" Then
                        strName = Trim(Replace(strSQL, "CREATE OR REPLACE VIEW", ""))
                        strName = Mid(strName, 1, InStr(strName, " ") - 1)
                        mrsViewFromFile.AddNew Array("ϵͳ���", "SQL", "����"), Array(mlngSysNum, strSQL, strName)
                    End If
                End If
                DoEvents
                Call mclsRunScript.ReadNextSQL
            Loop
        End If
    ElseIf mrsLocalFile!Filename = "ZLPACKAGE.SQL" Then
        If mclsRunScript.OpenFile(mrsLocalFile!FilePath) Then
            Do While Not mclsRunScript.EOF
                strSQL = iniSQL(mclsRunScript.SQLInfo.SQL)
                If strSQL <> "" Then
                    strSQL = iniSQL(strSQL)
                        If strSQL Like "CREATE OR REPLACE PACKAGE*" And Not strSQL Like "CREATE OR REPLACE PACKAGE BODY*" Then
                            strName = Trim(Replace(strSQL, "CREATE OR REPLACE PACKAGE", ""))
                            strName = Mid(strName, 1, InStr(strName, " ") - 1)
                            mrsPackageFromFile.AddNew Array("ϵͳ���", "SQL", "����"), Array(mlngSysNum, strSQL, strName)
                        End If
                End If
                DoEvents
                Call mclsRunScript.ReadNextSQL
            Loop
        End If
    ElseIf mrsLocalFile!Filename = "ZLPROGRAM.SQL" Then
        If mclsRunScript.OpenFile(mrsLocalFile!FilePath) Then
            Do While Not mclsRunScript.EOF
                strSQL = iniSQL(mclsRunScript.SQLInfo.SQL)
                If strSQL <> "" Then
                    If mlngSysNum = 2700 Then
                        If strSQL Like "CREATE OR REPLACE PROCEDURE ZLHIS.ZL_�����Ա��Ŀ_REJECT*" Then
                            strSQL = Replace(strSQL, "ZLHIS.", "")
                        End If
                    End If
                    If strSQL Like "CREATE OR REPLACE PROCEDURE*" Then
                        strName = Trim(Replace(strSQL, "CREATE OR REPLACE PROCEDURE", ""))
                    Else
                        strName = Trim(Replace(strSQL, "CREATE OR REPLACE FUNCTION", ""))
                    End If
                    If InStr(strName, vbCrLf) > 0 Then strName = Mid(strName, 1, InStr(strName, vbCrLf) - 1)
                    If InStr(strName, "(") > 0 Then strName = Trim(Mid(strName, 1, InStr(strName, "(") - 1))
                    If InStr(strName, " ") > 0 Then strName = Trim(Mid(strName, 1, InStr(strName, " ") - 1))
                    strName = Trim(strName)
                    If InStr(strSQL, "(") - InStr(strSQL, strName) - Len(strName) < 5 And Mid(InStr(strSQL, "(") - InStr(strSQL, strName) - Len(strName), 1, 1) <> "-" Then
                        strTemp = Mid(strSQL, InStr(strSQL, "(") + 1, InStr(strSQL, ")") - InStr(strSQL, "(") - 1)
                        If InStr(strTemp, "= ','") > 0 Then
                            varTemp = Split(strTemp, vbCrLf)
                        Else
                            varTemp = Split(strTemp, ",")
                        End If
                        strFild = ""
                        For i = 0 To UBound(varTemp)
                            varTemp(i) = Trim(Replace(varTemp(i), vbCrLf, ""))
                            If varTemp(i) <> "" Then
                                strFild = IIf(strFild = "", Trim(Mid(varTemp(i), 1, InStr(varTemp(i), " ") - 1)), strFild & "," & Trim(Mid(varTemp(i), 1, InStr(varTemp(i), " ") - 1)))
                            End If
                        Next
                        mrsProcedureFromFile.AddNew Array("ϵͳ���", "SQL", "����", "�ֶ�"), Array(mlngSysNum, strSQL, strName, strFild)
                    ElseIf strSQL Like "CREATE OR REPLACE VIEW *" Then
                        strName = Trim(Replace(strSQL, "CREATE OR REPLACE VIEW", ""))
                        strName = Mid(strName, 1, InStr(strName, " ") - 1)
                        mrsViewFromFile.AddNew Array("ϵͳ���", "SQL", "����"), Array(mlngSysNum, mclsRunScript.SQLInfo.SQL, strName)
                    End If
                End If
                DoEvents
                Call mclsRunScript.ReadNextSQL
            Loop
        End If
    ElseIf mrsLocalFile!Filename = "ZLMANDATA.SQL" Then
        If mclsRunScript.OpenFile(mrsLocalFile!FilePath) Then
            Do While Not mclsRunScript.EOF
                strSQL = iniSQL(mclsRunScript.SQLInfo.SQL)
                If strSQL <> "" Then
                    If strSQL Like "INSERT INTO*" Then
                        If strSQL Like "INSERT INTO ZLTABLES*" Then
                            strTemp = Replace(Mid(strSQL, InStr(strSQL, "(") + 1, InStr(strSQL, ")") - InStr(strSQL, "(") - 1), " ", "")
                            varFild = Split(strTemp, ",")
                            Set rsTemp = SetSelectRecordset(strSQL, strTemp, varFild, "ZLTABLES")
                            If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                            Do While Not rsTemp.EOF
                                mrsDataFromFile.AddNew Array("���", "SQL", "����", "ϵͳ���"), Array("��Ŀ¼", rsTemp!����SQL, rsTemp!����, mlngSysNum)
                                rsTemp.MoveNext
                            Loop
                        ElseIf strSQL Like "INSERT INTO ZLPROGRAMS*" Then
                            strTemp = Replace(Mid(strSQL, InStr(strSQL, "(") + 1, InStr(strSQL, ")") - InStr(strSQL, "(") - 1), " ", "")
                            varFild = Split(strTemp, ",")
                            Set rsTemp = SetSelectRecordset(strSQL, strTemp, varFild, "ZLPROGRAMS")
                            If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                            Do While Not rsTemp.EOF
                                mrsDataFromFile.AddNew Array("���", "SQL", "ϵͳ���", "���", "����"), Array("ģ��", rsTemp!����SQL, mlngSysNum, rsTemp!���, rsTemp!����)
                                rsTemp.MoveNext
                            Loop
                        ElseIf strSQL Like "INSERT INTO ZLPROGFUNCS*" Then
                            strTemp = Replace(Mid(strSQL, InStr(strSQL, "(") + 1, InStr(strSQL, ")") - InStr(strSQL, "(") - 1), " ", "")
                            varFild = Split(strTemp, ",")
                            Set rsTemp = SetSelectRecordset(strSQL, strTemp, varFild, "ZLPROGFUNCS")
                            If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                            Do While Not rsTemp.EOF
                                mrsDataFromFile.AddNew Array("���", "SQL", "ϵͳ���", "���", "����"), Array("����", rsTemp!����SQL, mlngSysNum, rsTemp!���, rsTemp!����)
                                rsTemp.MoveNext
                            Loop
                        ElseIf strSQL Like "INSERT INTO ZLPARAMETERS*" Then
                            If strSQL = "INSERT INTO ZLPARAMETERS(ID,ϵͳ,ģ��,˽��,����,��Ȩ,�̶�,����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵��)" & vbNewLine & _
                                        "SELECT ZLPARAMETERS_ID.NEXTVAL,2500,2500,-NULL,-NULL,-NULL,-NULL,-NULL,A.* FROM (" & vbNewLine & _
                                        "SELECT ˽��,����,��Ȩ,�̶�,����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵�� FROM ZLPARAMETERS WHERE 1 = 0 UNION ALL" & vbNewLine & _
                                        "SELECT 0,0,1,1,0,0,27,'����ǩ����֤����','0','0','�ڶԱ걾���к��ա����ʱ����ǩ����','�û�ѡ��ǩ����ʽ��0Ϊ��ʹ�õ���ǩ���⣬0���ϵ�ֵ��ʾʹ�ò�ͬ����֤���Ľ���ǩ����'||CHR(13)||'0=��ʹ�õ���ǩ��'||CHR(13)||'1=����ʡ����֤����֤����'||CHR(13)||'2=����ʡ����֤��֤����'||CHR(13)||'3=����������֤����֤����'||CHR(13)||'4=ɽ��ʡ����֤����֤����'||CHR(13)||'5=��������ҽԺ��֤����'||CHR(13)||'6=����ʡҽԺ��֤����',NULL,NULL,NULL FROM DUAL UNION ALL" & vbNewLine & _
                                        "SELECT ˽��,����,��Ȩ,�̶�,����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵�� FROM ZLPARAMETERS WHERE 1 = 0) A" Then
                                strSQL = "INSERT INTO ZLPARAMETERS(ID,ϵͳ,ģ��,˽��,����,��Ȩ,�̶�,����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵��)" & vbNewLine & _
                                        "SELECT ZLPARAMETERS_ID.NEXTVAL,2500,2500,A.* FROM (" & vbNewLine & _
                                        "SELECT ˽��,����,��Ȩ,�̶�,����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵�� FROM ZLPARAMETERS WHERE 1 = 0 UNION ALL" & vbNewLine & _
                                        "SELECT 0,0,1,1,0,0,27,'����ǩ����֤����','0','0','�ڶԱ걾���к��ա����ʱ����ǩ����','�û�ѡ��ǩ����ʽ��0Ϊ��ʹ�õ���ǩ���⣬0���ϵ�ֵ��ʾʹ�ò�ͬ����֤���Ľ���ǩ����'||CHR(13)||'0=��ʹ�õ���ǩ��'||CHR(13)||'1=����ʡ����֤����֤����'||CHR(13)||'2=����ʡ����֤��֤����'||CHR(13)||'3=����������֤����֤����'||CHR(13)||'4=ɽ��ʡ����֤����֤����'||CHR(13)||'5=��������ҽԺ��֤����'||CHR(13)||'6=����ʡҽԺ��֤����',NULL,NULL,NULL FROM DUAL UNION ALL" & vbNewLine & _
                                        "SELECT ˽��,����,��Ȩ,�̶�,����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵�� FROM ZLPARAMETERS WHERE 1 = 0) A"
                            End If
                            strTemp = Replace(Mid(strSQL, InStr(strSQL, "(") + 1, InStr(strSQL, ")") - InStr(strSQL, "(") - 1), " ", "")
                            varFild = Split(strTemp, ",")
                            Set rsTemp = SetSelectRecordset(strSQL, strTemp, varFild, "ZLPARAMETERS")
                            If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                            Do While Not rsTemp.EOF
                                If IsNull(rsTemp!ϵͳ) Or InStr(rsTemp!ϵͳ, "NULL") > 0 Or rsTemp!ģ�� = """" Then
                                    lngSys = 0
                                Else
                                    lngSys = mlngSysNum
                                End If
                                If InStr(strTemp, "ģ��") > 0 Then
                                    If InStr(UCase(rsTemp!ģ��), "NULL") > 0 Or rsTemp!ģ�� = """" Then
                                        mrsDataFromFile.AddNew Array("���", "SQL", "ϵͳ���", "����", "������", "������"), Array("����", rsTemp!����SQL, lngSys, "NULL", rsTemp!������, rsTemp!������)
                                    Else
                                        mrsDataFromFile.AddNew Array("���", "SQL", "ϵͳ���", "����", "������", "������"), Array("����", rsTemp!����SQL, lngSys, rsTemp!ģ��, rsTemp!������, rsTemp!������)
                                    End If
                                Else
                                    mrsDataFromFile.AddNew Array("���", "SQL", "ϵͳ���", "����", "������", "������"), Array("����", rsTemp!����SQL, lngSys, "NULL", rsTemp!������, rsTemp!������)
                                End If
                                rsTemp.MoveNext
                            Loop
                        End If
                    End If
                End If
                DoEvents
                Call mclsRunScript.ReadNextSQL
            Loop
        End If
    ElseIf mrsLocalFile!Filename = "ZLREPORT.SQL" Then
        If mclsRunScript.OpenFile(mrsLocalFile!FilePath) Then
            Do While Not mclsRunScript.EOF
                strSQL = UCase(mclsRunScript.SQLInfo.SQL)
                If strSQL Like "INSERT INTO ZLREPORTS*" Then
                    strSQL = iniSQL(strSQL)
                    varFild = Split(Replace(Mid(strSQL, InStr(strSQL, "(") + 1, InStr(strSQL, ")") - InStr(strSQL, "(") - 1), " ", ""), ",")
                    Set rsTemp = SetSelectRecordset(strSQL, strTemp, varFild, "ZLREPORTS")
                    If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                    Do While Not rsTemp.EOF
                        mrsDataFromFile.AddNew Array("���", "SQL", "ϵͳ���", "����", "����"), Array("����", rsTemp!����SQL, mlngSysNum, rsTemp!���, rsTemp!����)
                        rsTemp.MoveNext
                    Loop
                End If
                DoEvents
                Call mclsRunScript.ReadNextSQL
            Loop
        End If
    Else
        Call GetAnyObject
    End If
End Sub

Private Sub GetAnyObject()
'���������ű������п��ܵ�SQL����
    Dim strSQL As String
    Dim strIniSQL As String
    Dim strName As String
    Dim strTableName As String
    Dim strTableSpace As String
    Dim strFild As String
    Dim strFildType As String
    Dim strReFild As String
    Dim strFildLength As String
    Dim varTemp As Variant
    Dim i As Long
    Dim lngSys As Long
    Dim strFilter As String
    Dim strTemp As String
    Dim rsTemp As ADODB.Recordset
    
    If mclsRunScript.OpenFile(mrsLocalFile!FilePath) Then
        Do While Not mclsRunScript.EOF
            strSQL = mclsRunScript.SQLInfo.SQL
            strIniSQL = iniSQL(strSQL)
            If strIniSQL <> "" Then
                If strIniSQL Like "CREATE*" Then
                    If strIniSQL Like "CREATE OR REPLACE*" Then
                        If strIniSQL Like "CREATE OR REPLACE PROCEDURE*" Or strIniSQL Like "CREATE OR REPLACE FUNCTION*" Then
                            If strIniSQL Like "CREATE OR REPLACE PROCEDURE*" Then
                                strName = Trim(Replace(strIniSQL, "CREATE OR REPLACE PROCEDURE", ""))
                            Else
                                strName = Trim(Replace(strIniSQL, "CREATE OR REPLACE FUNCTION", ""))
                            End If
                            If InStr(strName, vbCrLf) > 0 Then strName = Mid(strName, 1, InStr(strName, vbCrLf) - 1)
                            If InStr(strName, " ") > 0 Then strName = Trim(Mid(strName, 1, InStr(strName, " ") - 1))
                            If InStr(strName, "(") > 0 Then strName = Trim(Mid(strName, 1, InStr(strName, "(") - 1))
                            strFilter = "ϵͳ���='" & mlngSysNum & "' and ����='" & strName & "'"
                            Call RecDelete(mrsProcedureFromFile, strFilter)
                            If InStr(strIniSQL, "(") - InStr(strIniSQL, strName) - Len(strName) < 5 And Mid(InStr(strIniSQL, "(") - InStr(strIniSQL, strName) - Len(strName), 1, 1) <> "-" Then
                                strFild = Mid(strIniSQL, InStr(strIniSQL, "(") + 1, InStr(strIniSQL, ")") - InStr(strIniSQL, "(") - 1)
                                If InStr(strFild, "','") > 0 Then
                                    varTemp = Split(strFild, vbCrLf)
                                Else
                                    varTemp = Split(strFild, ",")
                                End If
                                strFild = ""
                                For i = 0 To UBound(varTemp)
                                    varTemp(i) = Trim(Replace(varTemp(i), vbCrLf, ""))
                                    If varTemp(i) <> "" Then
                                        strFild = IIf(strFild = "", Trim(Mid(varTemp(i), 1, InStr(varTemp(i), " ") - 1)), strFild & "," & Trim(Mid(varTemp(i), 1, InStr(varTemp(i), " ") - 1)))
                                    End If
                                Next
                                mrsProcedureFromFile.AddNew Array("ϵͳ���", "SQL", "����", "�ֶ�"), Array(mlngSysNum, strSQL, strName, strFild)
                            End If
                        ElseIf strIniSQL Like "CREATE OR REPLACE VIEW *" Then
                            strName = Trim(Replace(strIniSQL, "CREATE OR REPLACE VIEW", ""))
                            strName = Mid(strName, 1, InStr(strName, " ") - 1)
                            strFilter = "ϵͳ���='" & mlngSysNum & "' and ����='" & strName & "'"
                            Call RecDelete(mrsViewFromFile, strFilter)
                            mrsViewFromFile.AddNew Array("ϵͳ���", "SQL", "����"), Array(mlngSysNum, strSQL, strName)
                        ElseIf strIniSQL Like "CREATE OR REPLACE PACKAGE*" And Not strIniSQL Like "CREATE OR REPLACE PACKAGE BODY*" Then
                            If InStr(strIniSQL, vbCrLf) > 0 Then strName = Mid(strIniSQL, 1, InStr(strIniSQL, vbCrLf) - 1)
                             strName = Trim(Replace(strName, "CREATE OR REPLACE PACKAGE", ""))
                            strName = Mid(strName, 1, InStr(strName, " ") - 1)
                            strFilter = "ϵͳ���='" & mlngSysNum & "' and ����='" & strName & "'"
                            Call RecDelete(mrsPackageFromFile, strFilter)
                            mrsPackageFromFile.AddNew Array("ϵͳ���", "SQL", "����"), Array(mlngSysNum, strSQL, strName)
                        End If
                    ElseIf strIniSQL Like "CREATE INDEX *" Or strIniSQL Like "CREATE UNIQUE INDEX*" Then
                        varTemp = Split(strIniSQL, " ON ")
                        If strIniSQL Like "CREATE INDEX *" Then
                            strName = Trim(Replace(varTemp(0), "CREATE INDEX", ""))
                        Else
                            strName = Trim(Replace(varTemp(0), "CREATE UNIQUE INDEX", ""))
                        End If
                        strTableName = Trim(Mid(varTemp(1), 1, InStr(varTemp(1), "(") - 1))
                        strFild = Replace(Mid(varTemp(1), InStr(varTemp(1), "(") + 1, InStrRev(varTemp(1), ")") - InStr(varTemp(1), "(") - 1), " ", "")
                        strTableSpace = GetTableSpace(strIniSQL)
                        strFilter = "ϵͳ���='" & mlngSysNum & "' and ����='" & strName & "' and ����='" & strTableName & "'"
                        Call RecDelete(mrsIndexFromFile, strFilter)
                        mrsIndexFromFile.AddNew Array("ϵͳ���", "SQL", "����", "����", "�ֶ�", "��ռ�"), _
                                            Array(mlngSysNum, strIniSQL, strTableName, strName, strFild, strTableSpace)
                    ElseIf strIniSQL Like "CREATE SEQUENCE*" Then
                        strName = Trim(Replace(strIniSQL, "CREATE SEQUENCE", ""))
                        strName = Mid(strName, 1, InStr(strName, " ") - 1)
                        strFilter = "ϵͳ���='" & mlngSysNum & "' and ����='" & strName & "'"
                        Call RecDelete(mrsSequenceFromFile, strFilter)
                        mrsSequenceFromFile.AddNew Array("ϵͳ���", "SQL", "����"), Array(mlngSysNum, strSQL, strName)
                    ElseIf strIniSQL Like "CREATE TABLE *" Then
                        strIniSQL = Trim(Replace(strIniSQL, "CREATE TABLE", ""))
                        strTableName = Trim(Replace(Mid(strIniSQL, 1, InStr(strIniSQL, "(") - 1), vbCrLf, ""))
                        strIniSQL = Replace(strIniSQL, vbCrLf, "||")
                        If InStr(strIniSQL, "))") > 0 Then
                            strFild = Mid(strIniSQL, InStr(strIniSQL, "(") + 1, InStr(strIniSQL, "))") - InStr(strIniSQL, "("))
                        ElseIf InStr(strIniSQL, "||)") > 0 Then
                            strFild = Mid(strIniSQL, InStr(strIniSQL, "(") + 1, InStr(strIniSQL, "||)") - InStr(strIniSQL, "("))
                        ElseIf InStr(strIniSQL, ")||") > 0 Then
                            strFild = Mid(strIniSQL, InStr(strIniSQL, "(") + 1, InStr(strIniSQL, ")||") - InStr(strIniSQL, "("))
                        Else
                            strFild = Mid(strIniSQL, InStr(strIniSQL, "(") + 1)
                        End If
                        varTemp = Split(strFild, "||")
                        For i = LBound(varTemp) To UBound(varTemp)
                            varTemp(i) = Trim(varTemp(i))
                            If varTemp(i) <> "" And varTemp(i) <> ")" And InStr(varTemp(i), "TABLESPACE") = 0 Then
                                If InStr(varTemp(i), "TABLESPACE") > 0 Then Exit For
                                strFildType = ""
                                strFildLength = ""
                                strFild = TrimEx(Mid(varTemp(i), 1, InStr(varTemp(i), " ")))
                                strIniSQL = Trim(Mid(varTemp(i), InStr(varTemp(i), " ") + 1))
                                If InStr(strIniSQL, "DATE") > 0 Then
                                    strFildType = "DATE"
                                ElseIf InStr(strIniSQL, "LONG RAW") > 0 Then
                                    strFildType = "LONG RAW"
                                Else
                                    If InStr(strIniSQL, ")") > 0 Then
                                        strIniSQL = Trim(Mid(strIniSQL, 1, InStr(strIniSQL, ")") - 1))
                                    ElseIf InStr(strIniSQL, " ") > 0 Then
                                        strIniSQL = Trim(Mid(strIniSQL, 1, InStr(strIniSQL, " ") - 1))
                                    End If
                                    If InStr(strIniSQL, "(") > 0 Then
                                        strFildType = Mid(strIniSQL, 1, InStr(strIniSQL, "(") - 1)
                                        strFildLength = Mid(strIniSQL, InStr(strIniSQL, "(") + 1)
                                    ElseIf InStr(strIniSQL, ",") > 0 Then
                                        strFildType = Mid(strIniSQL, 1, Len(strIniSQL) - 1)
                                    ElseIf InStr(strIniSQL, ")") > 0 Then
                                        strFildType = Mid(strIniSQL, 1, InStr(strIniSQL, ")") - 1)
                                    Else
                                        strFildType = strIniSQL
                                    End If
                                    strFildType = Trim(Replace(strFildType, "|", ""))
                                End If
                                If strFild <> "" Then
                                    mrsFildFromFile.AddNew Array("ϵͳ���", "SQL", "����", "�ֶ�", "�ֶ�����", "�ֶγ���"), _
                                                        Array(mlngSysNum, strSQL, strTableName, strFild, strFildType, strFildLength)
                                End If
                            End If
                        Next
                    End If
                ElseIf strIniSQL Like "ALTER*" Then
                    If strIniSQL Like "ALTER TABLE*" Then
                        If InStr(strIniSQL, "CONSTRAINT") > 0 Then
                            If InStr(strIniSQL, "_CK_") = 0 Then
                                If strIniSQL Like "ALTER TABLE * ADD CONSTRAINT *" Then
                                    varTemp = Split(strIniSQL, "ADD CONSTRAINT")
                                    strName = Trim(Replace(varTemp(1), "CONSTRAINT", ""))
                                    '��ȡԼ������
                                    strName = TrimEx(Mid(strName, 1, InStr(strName, " ") - 1))
                                    '��ȡ����
                                    strTableName = Trim(Replace(varTemp(0), "ALTER TABLE", ""))
                                    strFild = Replace(Trim(Mid(strIniSQL, InStr(strIniSQL, "(") + 1, InStr(strIniSQL, ")") - InStr(strIniSQL, "(") - 1)), " ", "")
                                    strTableSpace = GetTableSpace(strIniSQL)
                                    strFilter = "����='" & strTableName & "' and ����='" & strName & "' and ϵͳ���=" & mlngSysNum
                                    Call RecDelete(mrsConstraintFromFile, strFilter)
                                    If InStr(strSQL, "NOVALIDATE") = 0 Then strSQL = strSQL & " NOVALIDATE"
                                    mrsConstraintFromFile.AddNew Array("ϵͳ���", "SQL", "����", "����", "�ֶ�", "��ռ�"), _
                                                        Array(mlngSysNum, strIniSQL, strTableName, strName, strFild, strTableSpace)
                                    If InStr(strIniSQL, "PRIMARY") > 0 Or InStr(strIniSQL, "UNIQUE") > 0 Then
                                        strFilter = "����='" & strTableName & "' and ����='" & strName & "' and ϵͳ���=" & mlngSysNum
                                        Call RecDelete(mrsIndexFromFile, strFilter)
                                        If strTableSpace <> "" Then
                                            strSQL = "Create Unique Index " & strName & " On " & strTableName & "(" & strFild & ") Tablespace " & strTableSpace & " Nologging"
                                            strSQL = strSQL & "||" & strIniSQL
                                        Else
                                            strSQL = "Create Unique Index " & strName & " On " & strTableName & "(" & strFild & ") Nologging"
                                            strSQL = strSQL & "||" & strIniSQL
                                        End If
                                        mrsIndexFromFile.AddNew Array("ϵͳ���", "SQL", "����", "����", "�ֶ�", "��ռ�"), _
                                                        Array(mlngSysNum, strSQL, strTableName, strName, strFild, strTableSpace)
                                    End If
                                ElseIf strIniSQL Like "ALTER TABLE*DROP CONSTRAINT*" Then
                                    varTemp = Split(strIniSQL, "DROP CONSTRAINT")
                                    strName = Trim(Replace(varTemp(1), "CONSTRAINT", ""))
                                    If InStr(strName, " ") > 0 Then strName = TrimEx(Mid(strName, 1, InStr(strName, " ") - 1))
                                    strTableName = Trim(Replace(varTemp(0), "ALTER TABLE", ""))
                                    strFilter = "����='" & strTableName & "' and ����='" & strName & "' and ϵͳ���=" & mlngSysNum
                                    Call RecDelete(mrsConstraintFromFile, strFilter)
                                    Call RecDelete(mrsIndexFromFile, strFilter)
                                'Alter Table ���˿����ؼ�¼ rename Constraint ���˿����ؼ�¼_��ҳID to ���˿����ؼ�¼_FK_��ҳID
                                ElseIf strIniSQL Like "ALTER TABLE*RENAME CONSTRAINT*" Then
                                    strTemp = Mid(strIniSQL, InStr(strIniSQL, "CONSTRAINT") + 11)
                                    varTemp = Split(strIniSQL, "RENAME CONSTRAINT")
                                    strTableName = Trim(Replace(varTemp(0), "ALTER TABLE", ""))
                                    If strTableName = "�����ѽӿ�Ŀ¼" Then
                                        strTableName = "���ѿ����Ŀ¼"
                                    ElseIf strTableName = "���ѿ�Ŀ¼" Then
                                        strTableName = "���ѿ���Ϣ"
                                    End If
                                    varTemp = Split(varTemp(1), "TO")
                                    varTemp(0) = Trim(varTemp(0))
                                    varTemp(1) = Trim(varTemp(1))
                                    mrsConstraintFromFile.Filter = "����='" & varTemp(0) & "' and ϵͳ���=" & mlngSysNum
                                    If mrsConstraintFromFile.RecordCount > 0 Then
                                        mrsConstraintFromFile!���� = varTemp(1)
                                        mrsConstraintFromFile!���� = strTableName
                                        mrsConstraintFromFile!SQL = Replace(mrsConstraintFromFile!SQL, varTemp(0), varTemp(1))
                                        mrsConstraintFromFile!SQL = Replace(mrsConstraintFromFile!SQL, mrsConstraintFromFile!����, strTableName)
                                        mrsConstraintFromFile.Update
                                    End If
                                    mrsIndexFromFile.Filter = "����='" & varTemp(0) & "' and ϵͳ���=" & mlngSysNum
                                    If mrsIndexFromFile.RecordCount > 0 Then
                                        mrsIndexFromFile!���� = varTemp(1)
                                        mrsIndexFromFile!���� = strTableName
                                        mrsIndexFromFile!SQL = "Alter Index " & mrsIndexFromFile!���� & " rebulid nologging"
                                        mrsIndexFromFile.Update
                                    End If
                                End If
                            End If
                        Else
                            If strIniSQL Like "ALTER TABLE*ADD*" Then
                                If strIniSQL = "ALTER TABLE ʱ��� ADD (" & vbNewLine & _
                                            " վ�� VARCHAR2(1)," & vbNewLine & _
                                            " ���� VARCHAR2(10)," & vbNewLine & _
                                            " ����Ԥ��ʱ�� NUMBER(18)," & vbNewLine & _
                                            " ��Ϣʱ�� VARCHAR2(200))" Then
                                    strSQL = "alter table ʱ��� add վ�� VARCHAR2(1)"
                                    mrsFildFromFile.AddNew Array("ϵͳ���", "SQL", "����", "�ֶ�", "�ֶ�����", "�ֶγ���"), _
                                        Array(mlngSysNum, strSQL, "ʱ���", "վ��", "VARCHAR2", 1)
                                    strSQL = "alter table ʱ��� add ���� VARCHAR2(10)"
                                    mrsFildFromFile.AddNew Array("ϵͳ���", "SQL", "����", "�ֶ�", "�ֶ�����", "�ֶγ���"), _
                                        Array(mlngSysNum, strSQL, "ʱ���", "����", "VARCHAR2", 10)
                                    strSQL = "alter table ʱ��� add ����Ԥ��ʱ�� NUMBER(18)"
                                    mrsFildFromFile.AddNew Array("ϵͳ���", "SQL", "����", "�ֶ�", "�ֶ�����", "�ֶγ���"), _
                                        Array(mlngSysNum, strSQL, "ʱ���", "����Ԥ��ʱ��", "NUMBER", 18)
                                    strSQL = "alter table ʱ��� add ��Ϣʱ�� VARCHAR2(200)"
                                    mrsFildFromFile.AddNew Array("ϵͳ���", "SQL", "����", "�ֶ�", "�ֶ�����", "�ֶγ���"), _
                                        Array(mlngSysNum, strSQL, "ʱ���", "��Ϣʱ��", "VARCHAR2", strFildLength)
                                ElseIf strIniSQL = "ALTER TABLE ��Ա�սɼ�¼ ADD(" & vbNewLine & _
                                                    " �Ƿ�Һ� NUMBER(1)," & vbNewLine & _
                                                    " �Ƿ���￨ NUMBER(1)," & vbNewLine & _
                                                    " �Ƿ����ѿ� NUMBER(1)," & vbNewLine & _
                                                    " �Ƿ��շ� NUMBER(1)," & vbNewLine & _
                                                    " Ԥ����� NUMBER(2)," & vbNewLine & _
                                                    " �Ƿ���� NUMBER(1))" Then
                                    strSQL = "alter table ��Ա�սɼ�¼ add �Ƿ�Һ� NUMBER(1)"
                                    mrsFildFromFile.AddNew Array("ϵͳ���", "SQL", "����", "�ֶ�", "�ֶ�����", "�ֶγ���"), _
                                        Array(mlngSysNum, strSQL, "��Ա�սɼ�¼", "�Ƿ�Һ�", "NUMBER", 1)
                                    strSQL = "alter table ��Ա�սɼ�¼ add �Ƿ���￨ NUMBER(1)"
                                    mrsFildFromFile.AddNew Array("ϵͳ���", "SQL", "����", "�ֶ�", "�ֶ�����", "�ֶγ���"), _
                                        Array(mlngSysNum, strSQL, "��Ա�սɼ�¼", "�Ƿ���￨", "NUMBER", 1)
                                    strSQL = "alter table ��Ա�սɼ�¼ add �Ƿ����ѿ� NUMBER(1)"
                                    mrsFildFromFile.AddNew Array("ϵͳ���", "SQL", "����", "�ֶ�", "�ֶ�����", "�ֶγ���"), _
                                        Array(mlngSysNum, strSQL, "��Ա�սɼ�¼", "�Ƿ����ѿ�", "NUMBER", 1)
                                    strSQL = "alter table ��Ա�սɼ�¼ add �Ƿ��շ� NUMBER(1)"
                                    mrsFildFromFile.AddNew Array("ϵͳ���", "SQL", "����", "�ֶ�", "�ֶ�����", "�ֶγ���"), _
                                        Array(mlngSysNum, strSQL, "��Ա�սɼ�¼", "�Ƿ��շ�", "NUMBER", 1)
                                    strSQL = "alter table ��Ա�սɼ�¼ add Ԥ����� NUMBER(1)"
                                    mrsFildFromFile.AddNew Array("ϵͳ���", "SQL", "����", "�ֶ�", "�ֶ�����", "�ֶγ���"), _
                                        Array(mlngSysNum, strSQL, "��Ա�սɼ�¼", "Ԥ�����", "NUMBER", 2)
                                    strSQL = "alter table ��Ա�սɼ�¼ add �Ƿ���� NUMBER(1)"
                                    mrsFildFromFile.AddNew Array("ϵͳ���", "SQL", "����", "�ֶ�", "�ֶ�����", "�ֶγ���"), _
                                        Array(mlngSysNum, strSQL, "��Ա�սɼ�¼", "�Ƿ����", "NUMBER", 1)
                                ElseIf strIniSQL = "ALTER TABLE ZLREPORTS ADD (ִ����Ա VARCHAR2(20), ���ִ��ʱ�� DATE)" Then
                                    strSQL = "alter table ZLREPORTS add ִ����Ա varchar2(20)"
                                    mrsFildFromFile.AddNew Array("ϵͳ���", "SQL", "����", "�ֶ�", "�ֶ�����", "�ֶγ���"), _
                                    Array(mlngSysNum, strSQL, "ZLREPORTS", "ִ����Ա", "VARCHAR2", 20)
                                    strSQL = "alter table ZLREPORTS add ���ִ��ʱ�� varchar2(20)"
                                    mrsFildFromFile.AddNew Array("ϵͳ���", "SQL", "����", "�ֶ�", "�ֶ�����", "�ֶγ���"), _
                                    Array(mlngSysNum, strSQL, "ZLREPORTS", "���ִ��ʱ��", "DATE", "")
'                                ElseIf strIniSQL = "ALTER TABLE ѪҺ��Ѫ���� ADD (�Ƿ�Ӥ�� NUMBER (1)" Then
'                                    strIniSQL = "ALTER TABLE ѪҺ��Ѫ���� ADD �Ƿ�Ӥ�� NUMBER (1))"
                                    
                                Else
                                    strIniSQL = Replace(strIniSQL, vbCrLf, " ")
                                    varTemp = Split(strIniSQL, "ADD")
                                    strTableName = Trim(Replace(varTemp(0), "ALTER TABLE", ""))
                                    strName = Trim(varTemp(1))
                                    If Mid(strName, 1, 1) = "(" Then
                                        strName = Mid(strName, 2, InStrRev(strName, ")") - 2)
                                    End If
                                    strFild = Mid(strName, 1, InStr(strName, " ") - 1)
                                    If InStr(strName, "(") > 0 Then
                                        strFildType = Trim(Replace(strName, strFild, ""))
                                        strFildType = Trim(Mid(strFildType, 1, InStr(strFildType, "(") - 1))
                                        strFildLength = Trim(Mid(strName, InStr(strName, "(") + 1, InStr(strName, ")") - InStr(strName, "(") - 1))
                                    Else
                                        strFildType = Trim(Replace(strName, strFild, ""))
                                        If InStr(strFildType, " ") > 0 Then strFildType = Mid(strFildType, 1, InStr(strFildType, " ") - 1)
                                        If InStr(strFildType, ")") > 0 Then strFildType = Trim(Mid(strFildType, 1, InStr(strFildType, ")") - 1))
                                    End If
                                    strFilter = "����='" & strTableName & "' and �ֶ�='" & strFild & "' and ϵͳ���=" & mlngSysNum
                                    Call RecDelete(mrsFildFromFile, strFilter)
                                    mrsFildFromFile.AddNew Array("ϵͳ���", "SQL", "����", "�ֶ�", "�ֶ�����", "�ֶγ���"), _
                                        Array(mlngSysNum, strSQL, strTableName, strFild, strFildType, strFildLength)
                                End If
                            ElseIf strIniSQL Like "ALTER TABLE*MODIFY*" Then
                                varTemp = Split(strIniSQL, "MODIFY")
                                varTemp(1) = Trim(varTemp(1))
                                If InStr(varTemp(1), "NULL") = 0 And InStr(varTemp(1), "DEFAULT") = 0 Then
                                    strTableName = Trim(Replace(varTemp(0), "ALTER TABLE", ""))
                                    varTemp(1) = Trim(varTemp(1))
                                    If Mid(varTemp(1), 1, 1) = "(" Then varTemp(1) = Mid(varTemp(1), 2, Len(varTemp(1)) - 2)
                                    strFild = Mid(varTemp(1), 1, InStr(varTemp(1), " ") - 1)
                                    strTemp = Trim(Replace(varTemp(1), strFild, ""))
                                    If InStr(strTemp, "(") > 0 Then
                                        strFildType = Mid(strTemp, 1, InStr(strTemp, "(") - 1)
                                        strFildLength = Mid(strTemp, InStr(strTemp, "(") + 1, InStr(strTemp, ")") - InStr(strTemp, "(") - 1)
                                    Else
                                        strFildType = strTemp
                                        strFildLength = ""
                                    End If
                                    mrsFildFromFile.Filter = "����='" & strTableName & "' and �ֶ�='" & strFild & "'"
                                    If mrsFildFromFile.RecordCount > 0 Then
                                        mrsFildFromFile!�ֶ� = strFild
                                        mrsFildFromFile!�ֶ����� = strFildType
                                        mrsFildFromFile!�ֶγ��� = strFildLength
                                        mrsFildFromFile!SQL = strIniSQL
                                        mrsFildFromFile.Update
                                    End If
                                End If
                            ElseIf strIniSQL Like "ALTER TABLE*DROP COLUMN*" Then
                                varTemp = Split(strIniSQL, "DROP COLUMN")
                                strTableName = Trim(Replace(varTemp(0), "ALTER TABLE", ""))
                                strFild = Trim(varTemp(1))
                                strFilter = "����='" & strTableName & "' and �ֶ�='" & strFild & "' and ϵͳ���=" & mlngSysNum
                                Call RecDelete(mrsFildFromFile, strFilter)
                            ElseIf strIniSQL Like "ALTER TABLE*RENAME COLUMN*" Then
                                varTemp = Split(strIniSQL, "RENAME COLUMN")
                                strTableName = Trim(Replace(varTemp(0), "ALTER TABLE", ""))
                                varTemp = Split(varTemp(1), "TO")
                                strFild = Trim(varTemp(0))
                                strTemp = Trim(varTemp(1))
                                If strTemp Like "*BAK" Then
                                    strFilter = "����='" & strTableName & "' and �ֶ�='" & strFild & "'"
                                    Call RecDelete(mrsFildFromFile, strFilter)
                                Else
                                    mrsFildFromFile.Filter = "����='" & strTableName & "' and �ֶ�='" & strFild & "'"
                                    If mrsFildFromFile.RecordCount > 0 Then
                                        mrsFildFromFile!�ֶ� = strTemp
                                        mrsFildFromFile.Update
                                    End If
                                End If
                            ElseIf strIniSQL Like "ALTER TABLE*RENAME TO*" Then
                                varTemp = Split(strIniSQL, "RENAME TO")
                                strTableName = Trim(Replace(varTemp(0), "ALTER TABLE", ""))
                                strTemp = Trim(varTemp(1))
                                If strTemp Like "*BAK" Then
                                    strName = Trim(Replace(strIniSQL, "DROP TABLE", ""))
                                    strFilter = "����='" & strTableName & "' and ϵͳ���=" & mlngSysNum
                                    Call RecDelete(mrsFildFromFile, strFilter)
                                    strFilter = "���� like '" & strTableName & "*' and ϵͳ���=" & mlngSysNum
                                    Call RecDelete(mrsConstraintFromFile, strFilter)
                                    strFilter = "���� like '" & strTableName & "*' and ϵͳ���=" & mlngSysNum
                                    Call RecDelete(mrsIndexFromFile, strFilter)
                                    strFilter = "���� like '" & strTableName & "*' and ϵͳ���=" & mlngSysNum
                                    Call RecDelete(mrsSequenceFromFile, strFilter)
                                    strFilter = "���='��Ŀ¼' and ���� = '" & strTableName & "' and ϵͳ���=" & mlngSysNum
                                    Call RecDelete(mrsDataFromFile, strFilter)
                                Else
                                    mrsFildFromFile.Filter = "����='" & strTableName & "'"
                                    Do While Not mrsFildFromFile.EOF
                                        mrsFildFromFile!���� = strTemp
                                        mrsFildFromFile.Update
                                        mrsFildFromFile.MoveNext
                                    Loop
                                    mrsConstraintFromFile.Filter = "���� like '" & strTableName & "*'"
                                    Do While Not mrsConstraintFromFile.EOF
                                        mrsConstraintFromFile!���� = strTemp
                                        mrsConstraintFromFile!���� = Replace(mrsConstraintFromFile!����, strTableName, strTemp)
                                        mrsConstraintFromFile.Update
                                        mrsConstraintFromFile.MoveNext
                                    Loop
                                    mrsIndexFromFile.Filter = "���� like '" & strTableName & "*'"
                                    Do While Not mrsIndexFromFile.EOF
                                        mrsIndexFromFile!���� = Replace(mrsIndexFromFile!����, strTableName, strTemp)
                                        mrsIndexFromFile.Update
                                        mrsIndexFromFile.MoveNext
                                    Loop
                                    mrsSequenceFromFile.Filter = "���� like '" & strTableName & "*'"
                                    Do While Not mrsSequenceFromFile.EOF
                                        mrsSequenceFromFile!���� = Replace(mrsSequenceFromFile!����, strTableName, strTemp)
                                        mrsSequenceFromFile.Update
                                        mrsSequenceFromFile.MoveNext
                                    Loop
                                End If
                            End If
                        End If
                    ElseIf strIniSQL Like "ALTER INDEX*" Then
                        If strIniSQL Like "ALTER INDEX*RENAME TO*" Then
                            strTemp = Replace(strIniSQL, "ALTER INDEX", "")
                            varTemp = Split(strTemp, "RENAME TO")
                            varTemp(0) = Trim(varTemp(0))
                            varTemp(1) = Trim(varTemp(1))
                            strTableName = Mid(varTemp(1), 1, InStr(varTemp(1), "_") - 1)
                            mrsIndexFromFile.Filter = "����='" & varTemp(0) & "'"
                            If mrsIndexFromFile.RecordCount > 0 Then
                                mrsIndexFromFile!���� = varTemp(1)
                                mrsIndexFromFile!���� = strTableName
                                mrsIndexFromFile!SQL = "Alter Index " & mrsIndexFromFile!���� & " rebulid nologging"
                                mrsIndexFromFile.Update
                            End If
                        ElseIf strIniSQL Like "ALTER INDEX*REBUILD TABLESPACE*" Then
                            varTemp = Split(strIniSQL, "REBUILD TABLESPACE")
                            strTemp = Trim(Replace(varTemp(0), "ALTER INDEX", ""))
                            varTemp(1) = Trim(varTemp(1))
                            mrsIndexFromFile.Filter = "����='" & strTemp & "'"
                            If mrsIndexFromFile.RecordCount > 0 Then
                                mrsIndexFromFile!��ռ� = varTemp(1)
                                mrsIndexFromFile.Update
                            End If
                        End If
                    End If
                ElseIf strIniSQL Like "INSERT INTO*" Then
                    If strIniSQL Like "INSERT INTO ZLTABLES*" Then
                        strTemp = Replace(Mid(strIniSQL, InStr(strIniSQL, "(") + 1, InStr(strIniSQL, ")") - InStr(strIniSQL, "(") - 1), " ", "")
                        varTemp = Split(strTemp, ",")
                        Set rsTemp = SetSelectRecordset(strIniSQL, strTemp, varTemp, "ZLTABLES")
                        If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                        Do While Not rsTemp.EOF
                            mrsDataFromFile.Filter = "���='��Ŀ¼' and ����='" & rsTemp!���� & "' and ϵͳ���=" & mlngSysNum
                            If mrsDataFromFile.RecordCount = 0 Then
                                mrsDataFromFile.AddNew Array("���", "SQL", "����", "ϵͳ���"), Array("��Ŀ¼", rsTemp!����SQL, rsTemp!����, mlngSysNum)
                            End If
                            rsTemp.MoveNext
                        Loop
                    ElseIf strIniSQL Like "INSERT INTO ZLPROGRAMS*" Then
                        If InStr(strIniSQL, "����� IS NULL") > 0 And mlngShare = 0 Then
                            If strIniSQL = "INSERT INTO ZLPROGRAMS" & vbNewLine & _
                                " (���, ����, ˵��, ϵͳ, ����)" & vbNewLine & _
                                " SELECT 1082 ���, 'ҽ����Ȩ����' ����, 'ӵ�б�ģ��Ȩ�޵���Ա�ɶԱ����һ�ȫԺ�ٴ�ҽʦ������Ȩ�޽��м��й���' AS ˵��, 0 ϵͳ, 'ZL9CISBASE' ����" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPROGRAMS WHERE ϵͳ = 0 AND ��� = 1082 AND ���� = 'ҽ����Ȩ����')" Then
                                strIniSQL = "INSERT INTO ZLPROGRAMS" & vbNewLine & _
                                " (���, ����, ˵��, ϵͳ, ����)" & vbNewLine & _
                                " SELECT 1082, 'ҽ����Ȩ����', 'ӵ�б�ģ��Ȩ�޵���Ա�ɶԱ����һ�ȫԺ�ٴ�ҽʦ������Ȩ�޽��м��й���', 0, 'ZL9CISBASE'" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPROGRAMS WHERE ϵͳ = 0 AND ��� = 1082 AND ���� = 'ҽ����Ȩ����')"
                            ElseIf strIniSQL = "INSERT INTO ZLPROGRAMS" & vbNewLine & _
                                " (���, ����, ˵��, ϵͳ, ����)" & vbNewLine & _
                                " SELECT 2228 ���, '�������' ����, '���ڶԷ��Ľ�����˲���' AS ˵��, 0 ϵͳ, 'ZL9EMRINTERFACE' ����" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPROGRAMS WHERE ϵͳ = 0 AND ��� = 2228 AND ����='�������')" Then
                                strIniSQL = "INSERT INTO ZLPROGRAMS" & vbNewLine & _
                                " (���, ����, ˵��, ϵͳ, ����)" & vbNewLine & _
                                " SELECT 2228,'�������','���ڶԷ��Ľ�����˲���',0,'ZL9EMRINTERFACE'" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPROGRAMS WHERE ϵͳ = 0 AND ��� = 2228 AND ����='�������')"
                            ElseIf strIniSQL = "INSERT INTO ZLPROGRAMS" & vbNewLine & _
                                " (���, ����, ˵��, ϵͳ, ����)" & vbNewLine & _
                                " SELECT 2227 ���, 'ȡ���������' ����, '�����ڲ�����ɺ���Ҫ�ٴ��޸�ʱ������������' AS ˵��, 0 ϵͳ, 'ZL9EMRINTERFACE' ����" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPROGRAMS WHERE ϵͳ = 0 AND ��� = 2227 AND ����='ȡ���������')" Then
                                strIniSQL = "INSERT INTO ZLPROGRAMS" & vbNewLine & _
                                " (���, ����, ˵��, ϵͳ, ����)" & vbNewLine & _
                                " SELECT 2227,'ȡ���������','�����ڲ�����ɺ���Ҫ�ٴ��޸�ʱ������������',0,'ZL9EMRINTERFACE'" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPROGRAMS WHERE ϵͳ = 0 AND ��� = 2227 AND ����='ȡ���������')"
                            ElseIf strIniSQL = "INSERT INTO ZLPROGRAMS" & vbNewLine & _
                                " (���, ����, ˵��, ϵͳ, ����)" & vbNewLine & _
                                " SELECT 2226 ���, '��ĩ�ʿؽ���' ����, '������ĩ�ʿ�ǰ���н��չ�����������ͳ��' AS ˵��, 0 ϵͳ, 'ZL9EMRINTERFACE' ����" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPROGRAMS WHERE ��� = 2226 AND ����='��ĩ�ʿؽ���')" Then
                                strIniSQL = "INSERT INTO ZLPROGRAMS" & vbNewLine & _
                                " (���, ����, ˵��, ϵͳ, ����)" & vbNewLine & _
                                " SELECT 2226, '��ĩ�ʿؽ���', '������ĩ�ʿ�ǰ���н��չ�����������ͳ��', 0, 'ZL9EMRINTERFACE'" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPROGRAMS WHERE ��� = 2226 AND ����='��ĩ�ʿؽ���')"
                            ElseIf strIniSQL = "INSERT INTO ZLPROGRAMS" & vbNewLine & _
                                " (���, ����, ˵��, ϵͳ, ����)" & vbNewLine & _
                                " SELECT 2228 ���, '�������' ����, '���ڶԷ��Ľ�����˲���' AS ˵��, 0 ϵͳ, 'ZL9EMRINTERFACE' ����" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPROGRAMS WHERE ��� = 2228 AND ����='�������')" Then
                                strIniSQL = "INSERT INTO ZLPROGRAMS" & vbNewLine & _
                                " (���, ����, ˵��, ϵͳ, ����)" & vbNewLine & _
                                " SELECT 2228, '�������', '���ڶԷ��Ľ�����˲���', 0, 'ZL9EMRINTERFACE'" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPROGRAMS WHERE ��� = 2228 AND ����='�������')"
                            ElseIf strIniSQL = "INSERT INTO ZLPROGRAMS" & vbNewLine & _
                                " (���, ����, ˵��, ϵͳ, ����)" & vbNewLine & _
                                " SELECT 2227 ���, 'ȡ���������' ����, '�����ڲ�����ɺ���Ҫ�ٴ��޸�ʱ������������' AS ˵��, 0 ϵͳ, 'ZL9EMRINTERFACE' ����" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPROGRAMS WHERE ��� = 2227 AND ����='ȡ���������')" Then
                                strIniSQL = "INSERT INTO ZLPROGRAMS" & vbNewLine & _
                                " (���, ����, ˵��, ϵͳ, ����)" & vbNewLine & _
                                " SELECT 2227, 'ȡ���������', '�����ڲ�����ɺ���Ҫ�ٴ��޸�ʱ������������', 0, 'ZL9EMRINTERFACE'" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPROGRAMS WHERE ��� = 2227 AND ����='ȡ���������')"
                                
                            End If
                            strTemp = Replace(Mid(strIniSQL, InStr(strIniSQL, "(") + 1, InStr(strIniSQL, ")") - InStr(strIniSQL, "(") - 1), " ", "")
                            varTemp = Split(strTemp, ",")
                            Set rsTemp = SetSelectRecordset(strIniSQL, strTemp, varTemp, "ZLPROGRAMS")
                            If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                            Do While Not rsTemp.EOF
                                mrsDataFromFile.AddNew Array("���", "SQL", "ϵͳ���", "���", "����"), Array("ģ��", rsTemp!����SQL, mlngSysNum, Trim(rsTemp!���), Trim(rsTemp!����))
                                rsTemp.MoveNext
                            Loop
                        ElseIf strIniSQL Like "INSERT INTO ZLPROGFUNCS*" Then
                            If InStr(strIniSQL, "����� IS NULL") > 0 And mlngShare = 0 Then
                                strTemp = Replace(Mid(strIniSQL, InStr(strIniSQL, "(") + 1, InStr(strIniSQL, ")") - InStr(strIniSQL, "(") - 1), " ", "")
                                varTemp = Split(strTemp, ",")
                                Set rsTemp = SetSelectRecordset(strIniSQL, strTemp, varTemp, "ZLPROGFUNCS")
                                If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                                Do While Not rsTemp.EOF
                                    mrsDataFromFile.AddNew Array("���", "SQL", "ϵͳ���", "���", "����"), Array("����", rsTemp!����SQL, mlngSysNum, rsTemp!���, rsTemp!����)
                                    rsTemp.MoveNext
                                Loop
                            End If
                        ElseIf strIniSQL Like "INSERT INTO ZLPARAMETERS*" Then
                            If strIniSQL = "INSERT INTO ZLPARAMETERS" & vbNewLine & _
                                " (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)" & vbNewLine & _
                                " SELECT ZLPARAMETERS_ID.NEXTVAL, 0, 1124, 0, 0, 0, 0, 16, '��������Ч����'," & vbNewLine & _
                                " (SELECT DECODE(SUBSTR(NVL(����ֵ, ȱʡֵ), 1, 1), '0', '3', SUBSTR(NVL(����ֵ, ȱʡֵ), 1, 1)) AS VALIDDAY" & vbNewLine & _
                                " FROM ZLPARAMETERS" & vbNewLine & _
                                " WHERE ϵͳ = 0 AND ģ�� IS NULL AND ������ = '�Һ���Ч����'), '3', '�ɽ���ҽ���������ķ�����Ч������'" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPARAMETERS WHERE ϵͳ = 0 AND ģ�� = 1124 AND ������ = '��������Ч����')" Then
                                strIniSQL = "INSERT INTO ZLPARAMETERS" & vbNewLine & _
                                " (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)" & vbNewLine & _
                                " SELECT ZLPARAMETERS_ID.NEXTVAL, 0, 1124, 0, 0, 0, 0, 16, '��������Ч����','0', '3', '�ɽ���ҽ���������ķ�����Ч������'" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPARAMETERS WHERE ϵͳ = 0 AND ģ�� = 1124 AND ������ = '��������Ч����')"
                            End If
                            strTemp = Replace(Mid(strIniSQL, InStr(strIniSQL, "(") + 1, InStr(strIniSQL, ")") - InStr(strIniSQL, "(") - 1), " ", "")
                            varTemp = Split(strTemp, ",")
                            Set rsTemp = SetSelectRecordset(strIniSQL, strTemp, varTemp, "ZLPARAMETERS")
                            If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                            Do While Not rsTemp.EOF
                                If IsNull(rsTemp!ϵͳ) Or InStr(rsTemp!ϵͳ, "NULL") > 0 Or rsTemp!ģ�� = """" Then
                                    lngSys = 0
                                Else
                                    lngSys = mlngSysNum
                                End If
                                If InStr(strTemp, "ģ��") > 0 Then
                                    If InStr(UCase(rsTemp!ģ��), "NULL") > 0 Or rsTemp!ģ�� = """" Then
                                        mrsDataFromFile.AddNew Array("���", "SQL", "ϵͳ���", "����", "������", "������"), Array("����", rsTemp!����SQL, lngSys, "NULL", rsTemp!������, rsTemp!������)
                                    Else
                                        mrsDataFromFile.AddNew Array("���", "SQL", "ϵͳ���", "����", "������", "������"), Array("����", rsTemp!����SQL, lngSys, rsTemp!ģ��, rsTemp!������, rsTemp!������)
                                    End If
                                Else
                                    mrsDataFromFile.AddNew Array("���", "SQL", "ϵͳ���", "����", "������", "������"), Array("����", rsTemp!����SQL, lngSys, "NULL", rsTemp!������, rsTemp!������)
                                End If
                                rsTemp.MoveNext
                            Loop
                        ElseIf strIniSQL Like "INSERT INTO ZLREPORTS*" Then
                            strTemp = Replace(Mid(strIniSQL, InStr(strIniSQL, "(") + 1, InStr(strIniSQL, ")") - InStr(strIniSQL, "(") - 1), " ", "")
                            varTemp = Split(Replace(Mid(strIniSQL, InStr(strIniSQL, "(") + 1, InStr(strIniSQL, ")") - InStr(strIniSQL, "(") - 1), " ", ""), ",")
                            Set rsTemp = SetSelectRecordset(strIniSQL, strTemp, varTemp, "ZLREPORTS")
                            If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                            Do While Not rsTemp.EOF
                                mrsDataFromFile.AddNew Array("���", "SQL", "ϵͳ���", "����", "����"), Array("����", rsTemp!����SQL, mlngSysNum, rsTemp!���, rsTemp!����)
                                rsTemp.MoveNext
                            Loop
                        End If
                    End If
                ElseIf strIniSQL Like "UPDATE ZLPARAMETERS*" Then
                    If strIniSQL = "UPDATE ZLPARAMETERS" & vbNewLine & _
                            "SET ������ = 'һ��ͨ����ˢ������', ����ֵ = DECODE(����ֵ, NULL, NULL, '1|' || ����ֵ), ȱʡֵ = '1|0'," & vbNewLine & _
                            " Ӱ�����˵�� = '�����²���ǰ���Ƿ���Ҫ����ˢ������������֤��' || CHR(10) ||" & vbNewLine & _
                            " ' 1)������ʣ����ʣ����ʻ������' || CHR(10) ||" & vbNewLine & _
                            " ' 2)�������ʹ��Ԥ�����������������Ԥ���' || CHR(10) ||" & vbNewLine & _
                            " ' 3)�����շ�ʹ��Ԥ���������˷��˻�Ԥ���' || CHR(10) ||" & vbNewLine & _
                            " ' 4)����Һ�ʹ��Ԥ�����˺��˻�Ԥ���' || CHR(10) ||" & vbNewLine & _
                            " ' 5)����ҽ������Ϊ���ʵ���סԺҽ������Ϊ������ʵ������������ʷ��ã�סԺ��ʿվִ����ɣ�ҽ������վִ����ɻ�����ִ�����'," & vbNewLine & _
                            " ����ֵ���� = '������ʽ:����ˢ������|�˷�ˢ������' || CHR(10) ||" & vbNewLine & _
                            " ' 1.����ˢ������:0-������ˢ�����ƣ�1-��������ʱ��Ҫˢ����֤��2-��������ʱ���������(ֻҪ����һ�ſ�������ģ��ʹ��������������)�������ˢ����֤��' || CHR(10) ||" & vbNewLine & _
                            " ' 2.�˷�ˢ������:0-������ˢ�����ƣ�1-�����˷�ʱ��Ҫˢ����֤��2-�����˷�ʱ���������(ֻҪ����һ�ſ�������ģ��ʹ��������������)�������ˢ����֤��'," & vbNewLine & _
                            " ����˵�� = '����""���﷢��Ϊ���۵��������""������Ƿ���Ϊ���۵������Ҳ���ִ�к󱾿��Զ���˵�������򲻻ᵯ��������֤����Ϊ��û��ʵ�ʿۼ����˵ķ���'," & vbNewLine & _
                            " ����˵�� = '�������ˢ����������������֤������ܴ��ڿ�����ˢ�İ�ȫ���ա��е�ҽԺΪ�˷��㲡�˾���������ַ��գ��ڷ���ʱ��ҽԺ�벡��һ����Ҫǩ��Э�� '," & vbNewLine & _
                            " ����˵�� = '�˲������鲻����Ϊ""������ˢ������""���������ܻ���ڲ����ʽ�ȫ������Ϊ�˱�����������Ҫ��ÿ�����˶�����ˢ���������룬�Ա�֤�ʽ�ȫ'" & vbNewLine & _
                            "WHERE ϵͳ = 0 AND ģ�� IS NULL AND ������ = 28" Then
                        strIniSQL = "UPDATE ZLPARAMETERS" & vbNewLine & _
                            "SET ������ = 'һ��ͨ����ˢ������', ����ֵ = """", ȱʡֵ = '1|0'," & vbNewLine & _
                            " Ӱ�����˵�� = '�����²���ǰ���Ƿ���Ҫ����ˢ������������֤��' || CHR(10) ||" & vbNewLine & _
                            " ' 1)������ʣ����ʣ����ʻ������' || CHR(10) ||" & vbNewLine & _
                            " ' 2)�������ʹ��Ԥ�����������������Ԥ���' || CHR(10) ||" & vbNewLine & _
                            " ' 3)�����շ�ʹ��Ԥ���������˷��˻�Ԥ���' || CHR(10) ||" & vbNewLine & _
                            " ' 4)����Һ�ʹ��Ԥ�����˺��˻�Ԥ���' || CHR(10) ||" & vbNewLine & _
                            " ' 5)����ҽ������Ϊ���ʵ���סԺҽ������Ϊ������ʵ������������ʷ��ã�סԺ��ʿվִ����ɣ�ҽ������վִ����ɻ�����ִ�����'," & vbNewLine & _
                            " ����ֵ���� = '������ʽ:����ˢ������|�˷�ˢ������' || CHR(10) ||" & vbNewLine & _
                            " ' 1.����ˢ������:0-������ˢ�����ƣ�1-��������ʱ��Ҫˢ����֤��2-��������ʱ���������(ֻҪ����һ�ſ�������ģ��ʹ��������������)�������ˢ����֤��' || CHR(10) ||" & vbNewLine & _
                            " ' 2.�˷�ˢ������:0-������ˢ�����ƣ�1-�����˷�ʱ��Ҫˢ����֤��2-�����˷�ʱ���������(ֻҪ����һ�ſ�������ģ��ʹ��������������)�������ˢ����֤��'," & vbNewLine & _
                            " ����˵�� = '����""���﷢��Ϊ���۵��������""������Ƿ���Ϊ���۵������Ҳ���ִ�к󱾿��Զ���˵�������򲻻ᵯ��������֤����Ϊ��û��ʵ�ʿۼ����˵ķ���'," & vbNewLine & _
                            " ����˵�� = '�������ˢ����������������֤������ܴ��ڿ�����ˢ�İ�ȫ���ա��е�ҽԺΪ�˷��㲡�˾���������ַ��գ��ڷ���ʱ��ҽԺ�벡��һ����Ҫǩ��Э�� '," & vbNewLine & _
                            " ����˵�� = '�˲������鲻����Ϊ""������ˢ������""���������ܻ���ڲ����ʽ�ȫ������Ϊ�˱�����������Ҫ��ÿ�����˶�����ˢ���������룬�Ա�֤�ʽ�ȫ'" & vbNewLine & _
                            "WHERE ϵͳ = 0 AND ģ�� IS NULL AND ������ = 28"
                    End If
                    strTemp = Mid(strIniSQL, InStr(strIniSQL, "SET") + 3, InStr(strIniSQL, "WHERE") - InStr(strIniSQL, "SET") - 3)
                    If (InStr(strTemp, "������") > 0 Or InStr(strTemp, "������") > 0 Or InStr(strTemp, "ģ��") > 0) Then
                        If InStr(strIniSQL, "NOT EXISTS") > 0 Then
                            strTemp = Mid(strIniSQL, 1, InStr(strIniSQL, "NOT EXISTS") - 1)
                        Else
                            strTemp = strIniSQL
                        End If
                        strTemp = Mid(strTemp, InStr(strTemp, "WHERE") + 6)
                        varTemp = Split(strTemp, "AND")
                        strFild = ""
                        strReFild = ""
                        strName = ""
                        For i = 0 To UBound(varTemp)
                            varTemp(i) = Trim(varTemp(i))
                            If varTemp(i) <> "" Then
                                If InStr(varTemp(i), "IS NULL") = 0 Then
                                    strTemp = Trim(Mid(varTemp(i), 1, InStr(varTemp(i), "=") - 1))
                                    If InStr(strTemp, "������") > 0 Then
                                        strFild = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                    ElseIf InStr(strTemp, "������") > 0 Then
                                        strReFild = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                    ElseIf InStr(strTemp, "ģ��") > 0 Then
                                        strName = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                    End If
                                Else
                                    strTemp = Mid(varTemp(i), 1, InStr(varTemp(i), " ") - 1)
                                    If InStr(strTemp, "������") > 0 Then
                                        strFild = "NULL"
                                    ElseIf InStr(strTemp, "������") > 0 Then
                                        strReFild = "NULL"
                                    ElseIf InStr(strTemp, "ģ��") > 0 Then
                                        strName = "NULL"
                                    End If
                                End If
                            End If
                        Next
                        If strFild <> "" And strReFild = "" Then
                            mrsDataFromFile.Filter = "ϵͳ���=" & mlngSysNum & " and ����=" & strName & " and ������='" & strFild & "'"
                        ElseIf strFild = "" And strReFild <> "" Then
                            mrsDataFromFile.Filter = "ϵͳ���=" & mlngSysNum & " and ����=" & strName & " and ������='" & strReFild & "'"
                        ElseIf strFild <> "" And strReFild <> "" Then
                            mrsDataFromFile.Filter = "ϵͳ���=" & mlngSysNum & " and ����=" & strName & " and ������='" & strReFild & "' and ������='" & strFild & "'"
                        End If
                        If mrsDataFromFile.RecordCount > 0 Then
                            strFild = ""
                            strReFild = ""
                            strName = ""
                            strTemp = Replace(Mid(strIniSQL, InStr(strIniSQL, "SET") + 3, InStr(strIniSQL, "WHERE") - InStr(strIniSQL, "SET") - 3), vbCrLf, " ")
                            varTemp = Split(strTemp, ",")
                            For i = 0 To UBound(varTemp)
                                strTemp = Trim(Mid(varTemp(i), 1, InStr(varTemp(i), "=") - 1))
                                If strTemp = "������" Then
                                    strFild = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                ElseIf strTemp = "������" Then
                                    strReFild = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                ElseIf strTemp = "ģ��" Then
                                    strName = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                End If
                            Next
                            If strFild <> "" Then mrsDataFromFile!������ = strFild
                            If strReFild <> "" Then mrsDataFromFile!������ = strReFild
                            If strName <> "" Then mrsDataFromFile!���� = strName
                            mrsDataFromFile.Update
                        End If
                    End If
                ElseIf strIniSQL Like "DELETE*" Then
                    If strIniSQL Like "DELETE ZLPARAMETERS*" Or strIniSQL Like "DELETE FROM ZLPARAMETERS*" Then
                        strTemp = Mid(strIniSQL, InStr(strIniSQL, "WHERE") + 6)
                        If InStr(strTemp, "IN") = 0 And InStr(strTemp, "OR") = 0 Then
                            strName = ""
                            strFild = ""
                            strReFild = ""
                            varTemp = Split(strTemp, "AND")
                            For i = 0 To UBound(varTemp)
                                If InStr(varTemp(i), "NULL") > 0 Then
                                    If InStr(varTemp(i), "ģ��") > 0 Then
                                        strName = "NULL"
                                    End If
                                Else
                                    strTemp = Trim(Mid(varTemp(i), 1, InStr(varTemp(i), "=") - 1))
                                    If InStr(strTemp, "NVL") > 0 Then
                                        strTemp = Trim(Mid(strTemp, InStr(strTemp, "(") + 1, InStr(strTemp, ",") - InStr(strTemp, "(") - 1))
                                    End If
                                    If strTemp = "ģ��" Then
                                        strName = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                    ElseIf strTemp = "������" Then
                                        strFild = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                    ElseIf strTemp = "������" Then
                                        strReFild = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                    End If
                                End If
                            Next
                            If strFild <> "" And strReFild <> "" Then
                                strFilter = "���='����' and ϵͳ���=" & mlngSysNum & " and ����=" & strName & " and ������='" & strFild & "' and ������='" & strReFild & "'"
                            ElseIf strFild = "" Then
                                strFilter = "���='����' and ϵͳ���=" & mlngSysNum & " and ����=" & strName & " and ������='" & strReFild & "'"
                            ElseIf strReFild = "" Then
                                strFilter = "���='����' and ϵͳ���=" & mlngSysNum & " and ����=" & strName & " and ������='" & strFild & "'"
                            End If
                            If strName = "NULL" Then
                                strFilter = Replace(strFilter, "=NULL", "='NULL'")
                            End If
                            Call RecDelete(mrsDataFromFile, strFilter)
                        ElseIf InStr(strTemp, "IN") = 0 And InStr(strTemp, "OR") > 0 Then
                            varTemp = Split(strTemp, "AND")
                            For i = 0 To UBound(varTemp)
                                If InStr(varTemp(i), "(") = 0 Then
                                    strTemp = Trim(Mid(varTemp(i), 1, InStr(varTemp(i), "=") - 1))
                                    If strTemp = "������" Then
                                        strFild = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                    End If
                                Else
                                    strName = Mid(varTemp(i), InStr(varTemp(i), "(") + 1, InStr(varTemp(i), ")") - InStr(varTemp(i), "(") - 1)
                                End If
                            Next
                            If strName = "ģ�� = 1291 OR ģ�� = 1294" Then
                                strFilter = "���='����' and ϵͳ���=" & mlngSysNum & " and ����=1291 and ������='" & strFild & "'"
                                Call RecDelete(mrsDataFromFile, strFilter)
                                strFilter = "���='����' and ϵͳ���=" & mlngSysNum & " and ����=1294 and ������='" & strFild & "'"
                                Call RecDelete(mrsDataFromFile, strFilter)
                            End If
                        ElseIf strIniSQL = "DELETE FROM ZLPARAMETERS WHERE ϵͳ=&N_SYSTEM AND NVL(ģ��,0) IN (1252,1253) AND ������='�Զ�����Ƥ��'" Then
                            strFilter = "���='����' and ϵͳ���=" & mlngSysNum & " and ����=1252 and ������='�Զ�����Ƥ��'"
                            Call RecDelete(mrsDataFromFile, strFilter)
                            strFilter = "���='����' and ϵͳ���=" & mlngSysNum & " and ����=1253 and ������='�Զ�����Ƥ��'"
                            Call RecDelete(mrsDataFromFile, strFilter)
                        ElseIf strIniSQL = "DELETE ZLPARAMETERS WHERE ϵͳ = &N_SYSTEM AND (ģ�� = 1252 AND ������ IN (22, 24) OR ģ�� = 1253 AND ������ IN (17, 19, 45))" Then
                            strFilter = "���='����' and ϵͳ���=" & mlngSysNum & " and ����=1252 and ������=22"
                            Call RecDelete(mrsDataFromFile, strFilter)
                            strFilter = "���='����' and ϵͳ���=" & mlngSysNum & " and ����=1252 and ������=24"
                            Call RecDelete(mrsDataFromFile, strFilter)
                            strFilter = "���='����' and ϵͳ���=" & mlngSysNum & " and ����=1253 and ������=17"
                            Call RecDelete(mrsDataFromFile, strFilter)
                            strFilter = "���='����' and ϵͳ���=" & mlngSysNum & " and ����=1253 and ������=19"
                            Call RecDelete(mrsDataFromFile, strFilter)
                            strFilter = "���='����' and ϵͳ���=" & mlngSysNum & " and ����=1253 and ������=45"
                            Call RecDelete(mrsDataFromFile, strFilter)
                        End If
                    ElseIf strIniSQL Like "DELETE ZLPROGFUNCS*" Or strIniSQL Like "DELETE FROM ZLPROGFUNCS*" Then
                        strTemp = Mid(strIniSQL, InStr(strIniSQL, "WHERE") + 6)
                        If InStr(strTemp, "OR") = 0 Then
                            varTemp = Split(strTemp, "AND")
                            For i = 0 To UBound(varTemp)
                                strTemp = Trim(Mid(varTemp(i), 1, InStr(varTemp(i), "=") - 1))
                                If strTemp = "���" Then
                                    strFild = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                ElseIf strTemp = "����" Then
                                    strReFild = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                End If
                            Next
                            strFilter = "���='����' and ���=" & strFild & " and ����='" & strReFild & "' and ϵͳ���=" & mlngSysNum
                            Call RecDelete(mrsDataFromFile, strFilter)
                        Else
                            varTemp = Split(strTemp, "OR")
                            For i = 0 To UBound(varTemp)
                                strFild = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                strFilter = "���='����' and ����='" & strFild & "' and ϵͳ���=" & mlngSysNum
                                Call RecDelete(mrsDataFromFile, strFilter)
                            Next
                        End If
                    End If
                ElseIf strIniSQL Like "DROP*" Then
                    If strIniSQL Like "DROP SEQUENCE*" Then
                        strName = Trim(Replace(strIniSQL, "DROP SEQUENCE", ""))
                        strFilter = "����='" & strName & "' and ϵͳ���=" & mlngSysNum
                        Call RecDelete(mrsSequenceFromFile, strFilter)
                    ElseIf strIniSQL Like "DROP INDEX*" Then
                        strName = Trim(Replace(strIniSQL, "DROP INDEX", ""))
                        strFilter = "����='" & strName & "' and ϵͳ���=" & mlngSysNum
                        Call RecDelete(mrsIndexFromFile, strFilter)
                    ElseIf strIniSQL Like "DROP TABLE*" Then
                        strName = Trim(Replace(strIniSQL, "DROP TABLE", ""))
                        strFilter = "����='" & strName & "' and ϵͳ���=" & mlngSysNum
                        Call RecDelete(mrsFildFromFile, strFilter)
                        strFilter = "���� like '" & strName & "*' and ϵͳ���=" & mlngSysNum
                        Call RecDelete(mrsConstraintFromFile, strFilter)
                        strFilter = "���� like '" & strName & "*' and ϵͳ���=" & mlngSysNum
                        Call RecDelete(mrsIndexFromFile, strFilter)
                        strFilter = "���� like '" & strName & "*' and ϵͳ���=" & mlngSysNum
                        Call RecDelete(mrsSequenceFromFile, strFilter)
                        strFilter = "���='��Ŀ¼' and ���� = '" & strName & "' and ϵͳ���=" & mlngSysNum
                        Call RecDelete(mrsDataFromFile, strFilter)
                    ElseIf strIniSQL Like "DROP PROCEDURE*" Then
                        strName = Trim(Replace(strIniSQL, "DROP PROCEDURE", ""))
                        strFilter = "����='" & strName & "' and ϵͳ���=" & mlngSysNum
                        Call RecDelete(mrsProcedureFromFile, strFilter)
                    ElseIf strIniSQL Like "DROP FUNCTION*" Then
                        strName = Trim(Replace(strIniSQL, "DROP FUNCTION", ""))
                        strFilter = "����='" & strName & "' and ϵͳ���=" & mlngSysNum
                        Call RecDelete(mrsProcedureFromFile, strFilter)
                    ElseIf strIniSQL Like "DROP VIEW*" Then
                        strName = Trim(Replace(strIniSQL, "DROP VIEW", ""))
                        strFilter = "����='" & strName & "' and ϵͳ���=" & mlngSysNum
                        Call RecDelete(mrsViewFromFile, strFilter)
                    End If
                ElseIf strIniSQL Like "UPDATE ZLTABLES*" Then
                    strTemp = Mid(strIniSQL, InStr(strIniSQL, "SET") + 3, InStr(strIniSQL, "WHERE") - InStr(strIniSQL, "SET") - 3)
                    If InStr(strTemp, "����") > 0 Then
                        strTemp = Mid(strIniSQL, InStr(strIniSQL, "WHERE") + 6)
                        varTemp = Split(strTemp, "AND")
                        strTemp = ""
                        strFild = ""
                        strReFild = ""
                        For i = 0 To UBound(varTemp)
                            strTemp = Trim(Mid(varTemp(i), 1, InStr(varTemp(i), "=") - 1))
                            If strTemp <> "" And strTemp = "����" Then
                                strTableName = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                            End If
                        Next
                        strFilter = "ϵͳ���=" & mlngSysNum
                        If strTableName <> "" Then strFilter = strFilter & " and ����='" & strTableName & "'"
                        mrsDataFromFile.Filter = strFilter
                        If mrsDataFromFile.RecordCount > 0 Then
                            strTemp = Mid(strIniSQL, InStr(strIniSQL, "SET") + 3, InStr(strIniSQL, "WHERE") - InStr(strIniSQL, "SET") - 3)
                            varTemp = Split(strTemp, ",")
                            strTemp = ""
                            strFild = ""
                            strReFild = ""
                            For i = 0 To UBound(varTemp)
                                strTemp = Trim(Mid(varTemp(i), 1, InStr(varTemp(i), "=") - 1))
                                If strTemp = "����" Then
                                    strTableName = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                End If
                            Next
                            If strTableName <> "" Then
                                mrsDataFromFile!���� = strTableName
                                mrsDataFromFile.Update
                            End If
                        End If
                    End If
                End If
            End If
            DoEvents
            Call mclsRunScript.ReadNextSQL
        Loop
    End If
End Sub

Private Function iniSQL(ByVal strSQL As String) As String
    
    strSQL = Trim(UCase(strSQL))
    strSQL = Replace(strSQL, "ZLTOOLS.", "")
    strSQL = Replace(strSQL, Chr(0), " ")
    strSQL = Replace(strSQL, vbTab, " ")
    Do While InStr(strSQL, "  ") > 0
        strSQL = Replace(strSQL, "  ", " ")
    Loop
    If Mid(strSQL, 1, 11) = "INSERT INTO" Or Mid(strSQL, 1, 6) = "UPDATE" Then
        Call ReplaceMark(strSQL, strSQL)
    Else
        If strSQL Like "CREATE TABLE*" Or strSQL Like "CREATE OR REPLACE PROCEDURE*" Or strSQL Like "CREATE OR REPLACE FUNCTION*" Then
            strSQL = TrimAllComment(strSQL)
        End If
    End If
    
    If Right(strSQL, 1) = ";" Then
        strSQL = Mid(strSQL, 1, Len(strSQL) - 1)
    End If
    iniSQL = strSQL
End Function

Private Function ReplaceMark(ByRef strSQL As String, ByVal strCut As String) As String
    Dim strTemp As String
    Dim strCutSQL As String
    Dim strReplaceCutSQL As String
    Dim strIniSQL As String
    Dim lngBegin As Long
    
    If InStr(strCut, "'") > 0 Then
        lngBegin = InStr(strCut, "'") + 1
        strTemp = Mid(strCut, lngBegin)
        strCutSQL = "'" & Mid(strTemp, 1, InStr(strTemp, "'"))
        If InStr(strCutSQL, ",") > 0 Then
            If strCutSQL = "','" Then
                If InStr(Mid(strTemp, 1, 6), "||") > 0 Then
                    strCutSQL = "'" & Mid(strTemp, 1, 6)
                    strReplaceCutSQL = Replace(strCutSQL, ",", "��")
                    strSQL = Replace(strSQL, strCutSQL, strReplaceCutSQL)
                End If
            Else
                strReplaceCutSQL = Replace(strCutSQL, ",", "��")
                strSQL = Replace(strSQL, strCutSQL, strReplaceCutSQL)
            End If
        End If
        lngBegin = InStr(strTemp, "'") + 1
        strTemp = Mid(strTemp, lngBegin)
        Call ReplaceMark(strSQL, strTemp)
    End If
End Function

Public Function TrimAllComment(ByVal strSQL As String) As String
'���ܣ�ȥ��д�ڵ���strSQL�������"--"����"/"ע��(ֻ��Ա�ģ�飬��Ҫ������ȥ������߹���/�����ֶλ���������ע��)
'˵������Ҫ��RunSQLFile���Ӻ���
    Dim strTemp As String
    Dim strModifySQL As String
    Dim varTemp As Variant
    Dim blnStr As Boolean
    Dim i As Long
    
    If Mid(strSQL, 1, 2) = "--" Or strSQL = "" Or Mid(strSQL, 1, 1) = "/" Then Exit Function
    varTemp = Split(strSQL, vbCrLf)
    For i = 0 To UBound(varTemp)
        If InStr(varTemp(i), "--") > 0 Then
            strTemp = Mid(varTemp(i), 1, InStr(varTemp(i), "--") - 1)
            strModifySQL = IIf(strModifySQL = "", strTemp, strModifySQL & vbCrLf & strTemp)
        ElseIf InStr(varTemp(i), "/") > 0 Then
            strTemp = Mid(varTemp(i), 1, InStr(varTemp(i), "/") - 1)
            strModifySQL = IIf(strModifySQL = "", strTemp, strModifySQL & vbCrLf & strTemp)
        Else
            strModifySQL = IIf(strModifySQL = "", varTemp(i), strModifySQL & vbCrLf & varTemp(i))
        End If
    Next
    TrimAllComment = strModifySQL
End Function

Public Function GetTableSpace(ByVal strSQL As String) As String
'���ܣ����һ��������б�ռ䣬�򷵻ر�ռ��������ޣ��򷵻ؿ�
    Dim strTemp As String
    
    If InStr(strSQL, "TABLESPACE") > 0 Then
        strSQL = Replace(strSQL, vbCrLf, " ")
        strTemp = Trim(Right(strSQL, Len(strSQL) - InStrRev(strSQL, "TABLESPACE") - 10))
        If InStr(strTemp, " ") > 0 Then
            GetTableSpace = Mid(strTemp, 1, InStr(strTemp, " ") - 1)
        ElseIf InStr(strTemp, ";") > 0 Then
            GetTableSpace = Mid(strTemp, 1, InStr(strTemp, ";") - 1)
        Else
            GetTableSpace = strTemp
        End If
    Else
        GetTableSpace = ""
    End If
End Function

Private Function CheckSetFile(ByVal strPath As String, ByVal strSysNum As Long) As Boolean
'���ܣ���鱾�ذ�װ�ű�
'������strPath-��·����strSysNum-ϵͳ���
    Dim strMainPath As String
    Dim strProblem As String
    
    If strSysNum = 0 Then
        Call AddFilePath(UCase(strPath), "ZLSERVER.SQL", "�������ļ�", strSysNum, strProblem)
        CheckSetFile = True
        Exit Function
    End If
    
    strMainPath = UCase(Mid(strPath, 1, InStrRev(strPath, "\")))
    
    '��鰲װ�ű��Ƿ����
    Call AddFilePath(strMainPath & "ZLSEQUENCE.SQL", "ZLSEQUENCE.SQL", "�����ļ�", strSysNum, strProblem)
    Call AddFilePath(strMainPath & "ZLTABLE.SQL", "ZLTABLE.SQL", "���ݱ��ļ�", strSysNum, strProblem)
    Call AddFilePath(strMainPath & "ZLCONSTRAINT.SQL", "ZLCONSTRAINT.SQL", "Լ���ļ�", strSysNum, strProblem)
    Call AddFilePath(strMainPath & "ZLINDEX.SQL", "ZLINDEX.SQL", "�����ļ�", strSysNum, strProblem)
    Call AddFilePath(strMainPath & "ZLVIEW.SQL", "ZLVIEW.SQL", "��ͼ�ļ�", strSysNum, strProblem)
    Call AddFilePath(strMainPath & "ZLPROGRAM.SQL", "ZLPROGRAM.SQL", "���������ļ�", strSysNum, strProblem)
    Call AddFilePath(strMainPath & "ZLMANDATA.SQL", "ZLMANDATA.SQL", "���������ļ�", strSysNum, strProblem)
    Call AddFilePath(strMainPath & "ZLREPORT.SQL", "ZLREPORT.SQL", "���������ļ�", strSysNum, strProblem)
    
    If strProblem <> "" Then
        MsgBox "���·�������װ������ļ���ʧ�����ܼ�����������" & strProblem, vbExclamation, gstrSysName
        Exit Function
    End If
    
    '��Ѫ�⣬�豸�����������鰲װ�ű�û�а��ű��ļ���������������
    If Dir(strMainPath & "ZLPACKAGE.SQL") <> "" Then
        mrsLocalFile.AddNew Array("FilePath", "SystemNum", "FileName", "FileType"), _
                            Array(strMainPath & "ZLPACKAGE.SQL", strSysNum, "ZLPACKAGE.SQL", "��װ�ű�")
    End If
    CheckSetFile = True
End Function

Private Sub AddFilePath(ByVal strPath As String, ByVal strFileName As String, ByVal strFileType As String, ByVal strSysNum As Long, ByRef strProblem As String)

    If Dir(strPath) = "" Then
        strProblem = strProblem & vbCr & strFileType & strPath
    Else
        mrsLocalFile.AddNew Array("FilePath", "SystemNum", "FileName", "FileType"), Array(strPath, strSysNum, strFileName, "��װ�ű�")
    End If
End Sub

Private Sub cmdFunction_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngTemp As Long
    Dim i As Long
    
    For i = 0 To cmdFunction.UBound
        If i = Index Then
            If cmdFunction(Index) Is ActiveControl And cmdFunction(Index).FontBold = True Then Exit Sub
            
            For lngTemp = 0 To cmdFunction.UBound
                cmdFunction(lngTemp).FontBold = False
            Next
            cmdFunction(i).FontBold = True
            cmdFunction(i).SetFocus
            Select Case i
            Case 0
                lblNote.Caption = "ͨ���ռ����������ذ�װ�������ű�����ֹ����ǰ�汾���е����ݽṹ�ͻ����������ݣ������ݿ�������ʹ�õĽ��жԱȼ�顣" & vbNewLine & _
                            vbCrLf & "�������ݽṹ���������ֶΡ�Լ�������������С���ͼ�������洢���̡�" & vbNewLine & _
                                    "�����������ݰ�����ģ�顢���ܡ�������������Ŀ¼��" & vbNewLine & _
                            vbCrLf & "�����̱ȽϺ�ʱ�������ĵȴ������������������ĵ�������ѡ��ȫ���򲿷���������޸���" & vbNewLine & _
                                    "�޸����������漰��ض���Ķ�ռ����������Ӱ����ز�Ʒ���ܵ��������У�������ҵ������ڼ�ִ�С�" & vbNewLine & _
                                    "һ�㽨����������ɺ�ִ�б��������Լ��������������ű�ִ�г������Ժ���ܵ��µĽṹ�����ݲ�������"
            Case 1
                lblNote.Caption = "����ͬ���ָ��Ӧ��ϵͳ�����ߵ�ʵ�ʶ��󣨱��洢���̵ȣ���������ͨ����Աִ��SQLʱ���ʣ��Ա�����SQL�Ķ�������ǰ���������ǰ׺��" & vbNewLine & _
                            vbCrLf & "���ȱʧ����ͬ��ʣ���ͨ������Աִ�����SQLʱ�Ϳ��ܳ���" & vbNewLine & _
                                    "������ɺ���Զ�����������������һ������δͨ�������������ʱִ�нű����������"
            Case 2
                lblNote.Caption = "�Թ������û�ZLTOOLSִ��Ȩ�޼�����������������ȱ�ٵĹ���ͬ��ʣ�ZLTOOLS���ж���ģ���" & vbNewLine & _
                                    "ZLTOOLS�����ж�������Public�Ĺ���Ȩ�ޡ�����Ӧ��ϵͳ�����ߺ���ʷ�ռ������ߵ�ȫ��Ȩ�ޡ�" & vbNewLine & _
                            vbCrLf & "������ɺ���Զ�����������������һ������δͨ�������������ʱִ�нű����������"
            End Select
        End If
    Next
End Sub

Private Sub Form_Load()

    lblNote.Caption = "ͨ���ռ����������ذ�װ�������ű�����ֹ����ǰ�汾���е����ݽṹ�ͻ����������ݣ������ݿ�������ʹ�õĽ��жԱȼ�顣" & vbNewLine & _
                vbCrLf & "�������ݽṹ���������ֶΡ�Լ�������������С���ͼ�������洢���̡�" & vbNewLine & _
                        "�����������ݰ�����ģ�顢���ܡ�������������Ŀ¼��" & vbNewLine & _
                vbCrLf & "�����̱ȽϺ�ʱ�������ĵȴ������������������ĵ�������ѡ��ȫ���򲿷���������޸���" & vbNewLine & _
                        "�޸����������漰��ض���Ķ�ռ����������Ӱ����ز�Ʒ���ܵ��������У�������ҵ������ڼ�ִ�С�" & vbNewLine & _
                        "һ�㽨����������ɺ�ִ�б��������Լ��������������ű�ִ�г������Ժ���ܵ��µĽṹ�����ݲ�������"
    Call IniVSF
    Call GetVersion
    On Error Resume Next
    gcnOracle.Execute "select ���� from zltables"
    If err.Number = 0 Then mblnzlTables = True
    err.Clear: On Error GoTo 0
End Sub

Private Sub IniVSF()
'���ܣ���ʼ��VSF
    Dim rsSys As ADODB.Recordset
    Dim strSQL As String
    Dim i As Long
    
    Set rsSys = GetSystemList
    Call InitTable(vsfSelSys, MSTR_COL)

    With vsfSelSys
        .TextMatrix(.Rows - 1, Col_ϵͳ����) = "������������"
        .TextMatrix(.Rows - 1, Col_��ǰ�汾) = GetToolsVersion
        .TextMatrix(.Rows - 1, Col_�����) = 0
        Do While Not rsSys.EOF
            If Val(Split(rsSys!ϵͳ�汾��, ".")(0)) > 9 And rsSys!ϵͳ��� <> "2300" Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, Col_ϵͳ���) = rsSys!ϵͳ��� & ""
                .TextMatrix(.Rows - 1, Col_ϵͳ����) = rsSys!ϵͳ���� & ""
                .TextMatrix(.Rows - 1, Col_��ǰ�汾) = rsSys!ϵͳ�汾�� & ""
                .TextMatrix(.Rows - 1, Col_������) = rsSys!ϵͳ������ & ""
                .TextMatrix(.Rows - 1, Col_�����) = rsSys!����� & ""
            End If
            rsSys.MoveNext
        Loop
        .Cell(flexcpChecked, 0, 0, .Rows - .FixedRows) = flexChecked
        For i = 0 To .Rows - 1
            .rowHeight(i) = 300
        Next
    End With
End Sub

Private Sub GetVersion(Optional ByVal strMainPath As String)
'��ȡ�����ļ���ϵͳ�汾��
    Dim i As Long
    Dim strTemp As String
    Dim strMaxVer As String
    Dim strPath As String
    Dim varTemp As Variant
    Dim rsTemp As ADODB.Recordset
    Dim rsSetFile As ADODB.Recordset
    Dim blnExist As Boolean
    
    Set rsSetFile = GetSystemSetupIni
    '�����ߵĵ�ǰ�ű��汾��ȡ
    If strMainPath <> "" Then
        strTemp = strMainPath
        strPath = strTemp & "\TOOLS\zlServer.sql"
        lblMainPath.Caption = "ϵͳ��װĿ¼��" & strMainPath
    Else
        rsSetFile.Filter = "ϵͳ���=100"
        If rsSetFile.RecordCount <> 0 Then
            strTemp = rsSetFile!�ļ���
            strPath = Mid(strTemp, 1, 1) & ":\APPSOFT\TOOLS\zlServer.sql"
            lblMainPath.Caption = "ϵͳ��װĿ¼��" & Mid(strTemp, 1, 1) & ":\Appsoft"
        Else
            strPath = "C:\APPSOFT\TOOLS\ZLSERVER.SQL"
            lblMainPath.Caption = "ϵͳ��װĿ¼��C:\Appsoft"
        End If
    End If
    blnExist = True
    With vsfSelSys
        'Ӧ��ϵͳ�ĵ�ǰ�ű��汾��ȡ
        For i = .FixedRows To .Rows - .FixedRows
            If i = .FixedRows Then
                strPath = Replace(strPath, "\\", "\")
                If Dir(strPath) <> "" Then
                    .TextMatrix(1, Col_�����ļ�) = strPath
                Else
                    blnExist = False
                    .TextMatrix(1, Col_�����ļ�) = ""
                End If
            Else
                rsSetFile.Filter = "ϵͳ���=" & .TextMatrix(i, Col_ϵͳ���) & ""
                If rsSetFile.RecordCount > 0 Then
                    If strMainPath <> "" Then
                        varTemp = Split(rsSetFile!�ļ���, "APPSOFT")
                        strTemp = strMainPath & varTemp(1)
                    Else
                        strTemp = rsSetFile!�ļ���
                    End If
                    strTemp = Replace(strTemp, "\\", "\")
                    If Dir(strTemp) <> "" Then
                        .TextMatrix(i, Col_�����ļ�) = strTemp
                    Else
                        blnExist = False
                        .TextMatrix(i, Col_�����ļ�) = ""
                    End If
                End If
            End If
            If blnExist Then
                strMaxVer = ""
                varTemp = Split(.TextMatrix(i, Col_��ǰ�汾), ".")
                Set rsTemp = GetUpgradeFiles(rsTemp, Val(.TextMatrix(i, Col_ϵͳ���)), "10.34.0", .TextMatrix(i, Col_�����ļ�), "", "", strMaxVer)
                .Cell(flexcpText, i, Col_�ű��汾) = strMaxVer
            Else
                .Cell(flexcpText, i, Col_�ű��汾) = ""
            End If
        Next
        If blnExist = False Then MsgBox "��װĿ¼û�нű��ļ���������ѡ��"
        .Row = 1
        Call .ShowCell(1, 1)
    End With
End Sub

Private Sub Form_Resize()
    Dim i As Long
    
    On Error Resume Next
    With imgMain
        .Top = 700
        .Left = ScaleLeft + 200
    End With
    
    With vsfSelSys
        .Top = lblMainPath.Top + lblMainPath.Height + 50
        .Width = ScaleWidth - .Left - imgMain.Left
        .ColWidth(Col_�հ�) = ScaleWidth - .Left - 5000 - 280
        .Height = .Rows * 300 + 50
    End With
    
    cmdFunction(0).Top = vsfSelSys.Top + vsfSelSys.Height + 400
    cmdFunction(0).Left = vsfSelSys.Left

    chkIndex.Top = cmdFunction(0).Top
    chkIndex.Left = cmdFunction(0).Left + cmdFunction(0).Width + 300
    chkReport.Top = chkIndex.Top
    chkReport.Left = chkIndex.Left + chkIndex.Width + 500
    chkProcedure.Top = chkReport.Top
    chkProcedure.Left = chkReport.Left + chkReport.Width + 500
    chkParameters.Top = chkProcedure.Top
    chkParameters.Left = chkProcedure.Left + chkProcedure.Width + 500
    
    lblNote.Left = chkIndex.Left
    lblNote.Top = cmdFunction(0).Top + cmdFunction(0).Height + 50
    lblNote.Width = ScaleWidth - lblNote.Left
    lblNote.Height = 1600
    
    cmdFunction(2).Top = lblNote.Top + lblNote.Height - cmdFunction(2).Height
    cmdFunction(2).Left = cmdFunction(0).Left
    
    cmdFunction(1).Top = cmdFunction(2).Top - cmdFunction(1).Height - 50
    cmdFunction(1).Left = cmdFunction(1).Left
    
End Sub

Private Sub picStatus_Resize()
    If picStatus.ScaleWidth < 1000 Then Exit Sub
    
    With pgbProgress
        .Left = 150
        .Width = (picStatus.ScaleWidth - 150 * 4) / 2
        .Top = 240
        .Height = 180
        lblProgress.Left = .Left
        lblProgress.Top = .Top - lblProgress.Height
    End With
    
    
    With pgbState
        .Left = pgbProgress.Left + pgbProgress.Width + 300
        .Width = pgbProgress.Width
        .Top = 240
        .Height = 180
        lblStatus.Left = .Left
        lblStatus.Top = lblProgress.Top
    End With
    
    With Linepgb
        .x1 = pgbState.Left - 150
        .X2 = pgbState.Left - 150
        .y1 = 0
        .Y2 = picStatus.Height
    End With
End Sub

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�

End Sub

Private Sub lblSel_Click()
    Dim strFolderName As String
    Dim strOldPath As String
    strFolderName = lblMainPath.Tag
    
    strFolderName = OpenFolder(Me, "ѡ��ϵͳ��װĿ¼")
    If strFolderName = "" Then Exit Sub
    lblMainPath.Tag = strFolderName
    Call GetVersion(strFolderName)
End Sub

Private Sub vsfSelSys_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    Dim strNum As String
    
    With vsfSelSys
        If Col = Col_ѡ�� Then
            If Row = 0 Then
                If .Cell(flexcpChecked, 0, Col_ѡ��) = flexChecked Then
                    .Cell(flexcpChecked, 0, Col_ѡ��) = flexChecked
                    For i = .FixedRows To .Rows - .FixedRows
                        .Cell(flexcpChecked, i, Col_ѡ��) = flexChecked
                    Next
                Else
                    .Cell(flexcpChecked, 0, Col_ѡ��) = flexUnchecked
                    For i = .FixedRows To .Rows - .FixedRows
                        .Cell(flexcpChecked, i, Col_ѡ��) = flexUnchecked
                    Next
                End If
            ElseIf Row <> 0 Then
                If .Cell(flexcpChecked, 0, Col_ѡ��) = flexChecked Then
                    .Cell(flexcpChecked, 0, Col_ѡ��) = flexUnchecked
                End If
                For i = .FixedRows To .Rows - .FixedRows
                    If .Cell(flexcpChecked, i, Col_ѡ��) = flexUnchecked Then
                        Exit For
                    Else
                        If i = .Rows - .FixedRows Then
                            .Cell(flexcpChecked, 0, Col_ѡ��) = flexChecked
                        End If
                    End If
                Next
            End If
        End If
    End With
End Sub

Private Function CheckHistorySpaceEx(ByVal lngSys As Long) As Boolean
    '����:��鵱ǰϵͳ���Ƿ�����ʷ��ռ�ı���Ϣ

    Dim rsTmp As New ADODB.Recordset
    
    On Error Resume Next '���ܵ�ǰ�û�û�б�Ȩ��
    gstrSQL = "Select ����,������ From Zltools.Zlbakspaces Where ϵͳ = " & lngSys & "  And ��ǰ = 1 And ֻ�� = 0"
    Call OpenRecordset(rsTmp, gstrSQL, "��ȡ��ʷ��ռ�������")
    CheckHistorySpaceEx = Not rsTmp.EOF
    On Error GoTo 0
End Function

Private Sub Release()
'������ɺ��ͷ�ģ�鴰��

    Set mrsSequenceFromFile = Nothing
    Set mrsViewFromFile = Nothing
    Set mrsPackageFromFile = Nothing
    Set mrsFildFromFile = Nothing
    Set mrsConstraintFromFile = Nothing
    Set mrsIndexFromFile = Nothing
    Set mrsProcedureFromFile = Nothing
    Set mrsDataFromFile = Nothing
    
    Set mrsSequenceFromDB = Nothing
    Set mrsViewFromDB = Nothing
    Set mrsPackageFromDB = Nothing
    Set mrsFildFromDB = Nothing
    Set mrsConstraintFromDB = Nothing
    Set mrsIndexFromDB = Nothing
    Set mrsProcedureFromDB = Nothing
    Set mrsDataFromDB = Nothing
End Sub

Private Sub vsfSelSys_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    If Col <> 0 Then Cancel = True
End Sub


