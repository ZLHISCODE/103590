VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCISAduitPDF 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���������PDF"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12525
   Icon            =   "frmCISAduitPDF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   12525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdUnSelectAll 
      Cancel          =   -1  'True
      Caption         =   "ȫ��(&U)"
      Height          =   350
      Left            =   7845
      TabIndex        =   14
      Top             =   5685
      Width           =   1200
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "ȫѡ(&S)"
      Height          =   350
      Left            =   6660
      TabIndex        =   15
      Top             =   5685
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      Caption         =   "����������嵥(����������������)"
      Height          =   5415
      Left            =   90
      TabIndex        =   7
      Top             =   150
      Width           =   8955
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   5010
         Left            =   45
         TabIndex        =   8
         ToolTipText     =   "˫��ѡ��"
         Top             =   315
         Width           =   8820
         _cx             =   15557
         _cy             =   8837
         Appearance      =   2
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
         Cols            =   6
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
         MergeCells      =   1
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
         Begin VB.PictureBox picInfo 
            BackColor       =   &H00FFEBD7&
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   7470
            Picture         =   "frmCISAduitPDF.frx":000C
            ScaleHeight     =   225
            ScaleMode       =   0  'User
            ScaleWidth      =   283.333
            TabIndex        =   13
            Top             =   285
            Width           =   250
         End
         Begin MSComctlLib.ImageList img16 
            Left            =   0
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCISAduitPDF.frx":685E
                  Key             =   "Selected"
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.Frame fraPageScope 
      Caption         =   "���ѡ��(&R)"
      Height          =   5415
      Left            =   9120
      TabIndex        =   2
      Top             =   150
      Width           =   3330
      Begin VB.CommandButton cmdPath 
         Caption         =   "��"
         Height          =   315
         Left            =   2910
         TabIndex        =   3
         Top             =   4898
         Width           =   210
      End
      Begin VB.ListBox lst 
         Height          =   4470
         Left            =   180
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   240
         Width           =   2985
      End
      Begin VB.Frame Frame1 
         Height          =   120
         Left            =   30
         TabIndex        =   4
         Top             =   4710
         Width           =   3270
      End
      Begin VB.TextBox txtPath 
         Height          =   300
         Left            =   960
         TabIndex        =   11
         Top             =   4905
         Width           =   1965
      End
      Begin VB.ComboBox cboPrinterName 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   4930
         Width           =   2205
      End
      Begin VB.Label Label1 
         Caption         =   "���λ��"
         Height          =   255
         Left            =   150
         TabIndex        =   6
         Top             =   4995
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "����豸"
         Height          =   255
         Left            =   150
         TabIndex        =   10
         Top             =   4995
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   11250
      TabIndex        =   1
      Top             =   5685
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   9120
      TabIndex        =   0
      Top             =   5685
      Width           =   1200
   End
   Begin VB.Label Label2 
      Height          =   270
      Left            =   90
      TabIndex        =   9
      Top             =   5730
      Width           =   6495
   End
End
Attribute VB_Name = "frmCISAduitPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnDoctorAdvice As Boolean
Private mblnPDF As Boolean
Private mstrPrintDocIDs As String '�����������ĵ�ֻ��ӡһ��
Private WithEvents mclsDockAduits As zlRichEPR.clsDockAduits
Attribute mclsDockAduits.VB_VarHelpID = -1
Private mfrmTipInfo As New frmTipInfo
Private Enum mCol
    ѡ��
    ����ID
    ��ҳID
    ����
    ����
    סԺ��
    �Ա�
    ����
    ��Ժ����
    ��Ժ����ID
    ��Ժ����
    ��Ժ����
    ��ӡ��¼
End Enum

Public Sub ShowMe(ByVal frmObj As Object, ByVal parVsf As VSFlexGrid, ByVal intType As Integer, ByVal blnDoctorAdvice As Boolean, ByVal blnPDF As Boolean)
Dim strPath As String, strSelect As String, strTmp As String, intCount As Integer, arySerial As Variant, strPrinterName As String
'blnDoctorAdvice False=ҽ���� True=ҽ����
'----------------------------------------------------------------------------------------
'1-סԺҽ��;2-סԺ����;3-������;4-�����¼;5-��ҳ��¼;6-ҽ������;7-����֤��;8-֪���ļ�
    On Error GoTo errHand
    mblnPDF = blnPDF
    If blnPDF = False Then '�豸������� ���PDF
        strPrinterName = GetRegister(˽��ģ��, "��ӡ����", "��ӡ��", Printer.DeviceName)
        strSelect = "," & GetRegister(˽��ģ��, "��ӡ����", "��ӡ����", "1,2,3,4,5,6,7,8,9") & ","
        With cboPrinterName
            .Clear
            For intCount = 0 To Printers.count - 1
                .AddItem Printers(intCount).DeviceName
                If Printers(intCount).DeviceName = strPrinterName Then .ListIndex = intCount
            Next
        End With
        Call zlControl.CboSetWidth(cboPrinterName.hWnd, 3000)
        Label1.Visible = False
        txtPath.Visible = False
        cmdPath.Visible = False
        picInfo.Visible = False
        Me.Caption = "�������"
    Else '���PDF
        Me.Caption = "���������PDF"
        strSelect = "," & GetRegister(˽��ģ��, "��ӡ����", "���PDF", "1,2,3,4,5,6,7,8,9") & ","
        Label3.Visible = False
        cboPrinterName.Visible = False
        picInfo.Visible = True
        strPath = GetRegister(˽��ģ��, "��ӡ����", "PDFλ��", App.Path)
        txtPath.Text = strPath: txtPath.ToolTipText = strPath
    End If
    
    mstrPrintDocIDs = ""
    mblnDoctorAdvice = blnDoctorAdvice
    Call FillVfg(parVsf, intType)
    
    strTmp = Trim(zlDatabase.GetPara("��������˳��", ParamInfo.ϵͳ��, 1560, "5;1;6;2;3;4;8;7;9"))
    If strTmp = "" Then strTmp = "5;1;6;2;3;4;8;7;9"
    arySerial = Split(strTmp, ";")
    
    With lst
        For intCount = 0 To UBound(arySerial)
            Select Case Val(arySerial(intCount))
            Case 1
                .AddItem "סԺҽ��": .ItemData(.NewIndex) = 1
                If InStr(strSelect, ",1,") > 0 Then .Selected(.NewIndex) = True
            Case 2
                .AddItem "סԺ����": .ItemData(.NewIndex) = 2
                If InStr(strSelect, ",2,") > 0 Then .Selected(.NewIndex) = True
            Case 3
                .AddItem "������": .ItemData(.NewIndex) = 3
                If InStr(strSelect, ",3,") > 0 Then .Selected(.NewIndex) = True
            Case 4
                .AddItem "�����¼": .ItemData(.NewIndex) = 4
                If InStr(strSelect, ",4,") > 0 Then .Selected(.NewIndex) = True
            Case 5
                .AddItem "��ҳ����": .ItemData(.NewIndex) = 5
                If InStr(strSelect, ",5,") > 0 Then .Selected(.NewIndex) = True
                .AddItem "��ҳ����": .ItemData(.NewIndex) = 52
                If InStr(strSelect, ",52,") > 0 Then .Selected(.NewIndex) = True
                .AddItem "��ҳ��ҳһ": .ItemData(.NewIndex) = 53
                If InStr(strSelect, ",53,") > 0 Then .Selected(.NewIndex) = True
                .AddItem "��ҳ��ҳ��": .ItemData(.NewIndex) = 54
                If InStr(strSelect, ",54,") > 0 Then .Selected(.NewIndex) = True
            Case 6
                .AddItem "ҽ������": .ItemData(.NewIndex) = 6
                If InStr(strSelect, ",6,") > 0 Then .Selected(.NewIndex) = True
            Case 7
                .AddItem "����֤��": .ItemData(.NewIndex) = 7
                If InStr(strSelect, ",7,") > 0 Then .Selected(.NewIndex) = True
            Case 8
                .AddItem "֪���ļ�": .ItemData(.NewIndex) = 8
                If InStr(strSelect, ",8,") > 0 Then .Selected(.NewIndex) = True
            Case 9
                .AddItem "�ٴ�·��": .ItemData(.NewIndex) = 9
                If InStr(strSelect, ",9,") > 0 Then .Selected(.NewIndex) = True
            End Select
        Next

        .ListIndex = 0
    End With
    
    Me.Show 1, frmObj
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub FillVfg(ByVal parVsf As VSFlexGrid, ByVal intType As Integer)
Dim i As Integer
    On Error GoTo errHand
    With vsf
        .Clear
        .Rows = parVsf.Rows
        .Cols = 13
        .RowHeight(0) = 350
        .TextMatrix(0, mCol.ѡ��) = "ѡ��"
        .TextMatrix(0, mCol.����ID) = "����ID"
        .TextMatrix(0, mCol.��ҳID) = "��ҳID"
        .TextMatrix(0, mCol.����) = "����"
        .TextMatrix(0, mCol.����) = "����"
        .TextMatrix(0, mCol.סԺ��) = "סԺ��"
        .TextMatrix(0, mCol.�Ա�) = "�Ա�"
        .TextMatrix(0, mCol.����) = "����"
        .TextMatrix(0, mCol.��Ժ����) = "��Ժ����"
        .TextMatrix(0, mCol.��Ժ����ID) = "��Ժ����ID"
        .TextMatrix(0, mCol.��Ժ����) = "��Ժ����"
        .TextMatrix(0, mCol.��Ժ����) = "��Ժ����"
        .TextMatrix(0, mCol.��ӡ��¼) = "��ӡ��¼"
        
        .ColWidth(mCol.ѡ��) = 400
        .ColWidth(mCol.����ID) = 0
        .ColWidth(mCol.��ҳID) = 0
        .ColWidth(mCol.����) = 1200
        .ColWidth(mCol.����) = 600
        .ColWidth(mCol.סԺ��) = 1200
        .ColWidth(mCol.��Ժ����) = 1800
        .ColWidth(mCol.�Ա�) = 500
        .ColWidth(mCol.��Ժ����ID) = 0
        If intType = 0 Then
            .ColWidth(mCol.��Ժ����) = 1800
            .ColWidth(mCol.��Ժ����) = 800
        Else
            .ColWidth(mCol.��Ժ����) = 0
            .ColWidth(mCol.��Ժ����) = 0
        End If
        .ColWidth(mCol.����) = 500
        .ColWidth(mCol.��ӡ��¼) = 0
                
        For i = 1 To parVsf.Rows - 1
            .RowHeight(i) = 350
            .Cell(flexcpData, i, mCol.ѡ��) = 0
            .TextMatrix(i, mCol.����ID) = parVsf.TextMatrix(i, parVsf.ColIndex("����ID"))
            .TextMatrix(i, mCol.��ҳID) = parVsf.TextMatrix(i, parVsf.ColIndex("��ҳID"))
            If intType = 1 Then '��Ժ
                .TextMatrix(i, mCol.����) = parVsf.TextMatrix(i, parVsf.ColIndex("����"))
                .TextMatrix(i, mCol.����) = parVsf.TextMatrix(i, parVsf.ColIndex("����"))
                .TextMatrix(i, mCol.סԺ��) = parVsf.TextMatrix(i, parVsf.ColIndex("סԺ��"))
                .TextMatrix(i, mCol.�Ա�) = parVsf.TextMatrix(i, parVsf.ColIndex("�Ա�"))
                .TextMatrix(i, mCol.����) = parVsf.TextMatrix(i, parVsf.ColIndex("����"))
                .TextMatrix(i, mCol.��Ժ����) = parVsf.TextMatrix(i, parVsf.ColIndex("��Ժ����"))
                .TextMatrix(i, mCol.��Ժ����ID) = parVsf.TextMatrix(i, parVsf.ColIndex("��Ժ����ID"))
            Else
                .TextMatrix(i, mCol.����) = parVsf.TextMatrix(i, parVsf.ColIndex("����"))
                If parVsf.ColIndex("����") <> -1 Then
                    .TextMatrix(i, mCol.����) = parVsf.TextMatrix(i, parVsf.ColIndex("����"))
                End If
                If parVsf.ColIndex("סԺ��") <> -1 Then
                    .TextMatrix(i, mCol.סԺ��) = parVsf.TextMatrix(i, parVsf.ColIndex("סԺ��"))
                End If
                If parVsf.ColIndex("����") <> -1 Then
                    .TextMatrix(i, mCol.����) = parVsf.TextMatrix(i, parVsf.ColIndex("����"))
                End If
                If parVsf.ColIndex("��Ժ����") <> -1 Then
                    .TextMatrix(i, mCol.��Ժ����) = parVsf.TextMatrix(i, parVsf.ColIndex("��Ժ����"))
                End If
                
                If parVsf.ColIndex("��Ժ����") <> -1 Then
                    .TextMatrix(i, mCol.��Ժ����) = parVsf.TextMatrix(i, parVsf.ColIndex("��Ժ����"))
                End If
                
                If parVsf.ColIndex("��Ժ����") <> -1 Then
                    .TextMatrix(i, mCol.��Ժ����) = parVsf.TextMatrix(i, parVsf.ColIndex("��Ժ����"))
                End If
                .TextMatrix(i, mCol.��Ժ����ID) = parVsf.TextMatrix(i, parVsf.ColIndex("��Ժ����ID"))
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
        .Row = parVsf.Row
        .TopRow = .Row
        If .Rows = 2 Then vsf_DblClick
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub PrintWithActiveEXE(ByVal strRegRange As String, ByVal strRange As String, ByVal strParPath As String)
Dim i As Long, strPrintContent As String, rs As New ADODB.Recordset, l As Long, objPrint As Object
Dim objRichEMR As Object, arrPar As Variant, arrParOne As Variant, j As Long, strEmrId As String
Dim varParam As Variant, strReportNO As String, lng����ID As Long, blnNewTends As Boolean, intSel As Integer, strEprName As String
Dim lngPatient As Long, lngPageID As Long, lngDept As Long, lngInNo As Long, strPath As String, strFileName As String, blnDataMove As Boolean, strName As String
    On Error GoTo errHand
    For i = 1 To vsf.Rows - 1
        With vsf
            If .Cell(flexcpData, i, mCol.ѡ�� = 1) Then 'ѡ�в����
                l = l + 1
                lngPatient = Val(.TextMatrix(i, mCol.����ID))
                lngInNo = Val(.TextMatrix(i, mCol.סԺ��))
                lngPageID = Val(.TextMatrix(i, mCol.��ҳID))
                lngDept = Val(.TextMatrix(i, mCol.��Ժ����ID))
                strName = NVL(.TextMatrix(i, mCol.����))
                strPath = strParPath & "\" & strName & "_" & lngInNo
                If Not gobjFSO.FolderExists(strPath) Then
                    Call gobjFSO.CreateFolder(strPath)
                End If

                '��ȡ��¼
                Set rs = gclsPackage.GetCISStruct(lngPatient, lngPageID, lngDept, blnDataMove)
                Do Until rs.EOF
                    If NVL(rs("�ϼ�id").Value) = "" Then
                        If InStr(strRange, "," & rs("ID").Value & ",") > 0 Then
                            Select Case rs("ID").Value
                            Case "R5" '��ҳ
                                'ϵͳ��,������,����id,��ҳid,1��/2��/3��һ/4����,PDFFileName
                                lng����ID = GetlngID(lngPatient, lngPageID)
                                Select Case Val(zlDatabase.GetPara("������ҳ��׼", glngSys, 1261, "0"))
                                Case 0 '��������׼
                                    If Have��������(lng����ID, "��ҽ��") Then
                                        strReportNO = "ZL1_INSIDE_1261_4"
                                    Else
                                        strReportNO = "ZL1_INSIDE_1261_1"
                                    End If
                                Case 1    '�Ĵ�ʡ��׼
                                    If Have��������(lng����ID, "��ҽ��") Then
                                        strReportNO = "ZL1_INSIDE_1261_6"
                                    Else
                                        strReportNO = "ZL1_INSIDE_1261_5"
                                    End If
                                Case 2    '����ʡ��׼
                                    If Have��������(lng����ID, "��ҽ��") Then
                                        strReportNO = "ZL1_INSIDE_1261_8"
                                    Else
                                        strReportNO = "ZL1_INSIDE_1261_7"
                                    End If
                                Case 3     '����ʡ��׼
                                    If Have��������(lng����ID, "��ҽ��") Then
                                        strReportNO = "ZL1_INSIDE_1261_10"
                                    Else
                                        strReportNO = "ZL1_INSIDE_1261_9"
                                    End If
                                Case Else '�����޸�ʱδ����
                                    If Have��������(lng����ID, "��ҽ��") Then
                                        strReportNO = "ZL1_INSIDE_1261_4"
                                    Else
                                        strReportNO = "ZL1_INSIDE_1261_1"
                                    End If
                                End Select
                                If InStr("," & strRegRange & ",", ",5,") > 0 Then '����
                                    strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_��ҳ����.PDF"
                                    strPrintContent = strPrintContent & "|" & strName & ",R5," & glngSys & "," & strReportNO & "," & lngPatient & "," & lngPageID & ",1," & strFileName
                                End If
                                
                                If InStr("," & strRegRange & ",", ",52,") > 0 Then '����
                                    strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_��ҳ����.PDF"
                                    strPrintContent = strPrintContent & "|" & strName & ",R5," & glngSys & "," & strReportNO & "," & lngPatient & "," & lngPageID & ",2," & strFileName
                                End If
                                
                                If InStr("," & strRegRange & ",", ",53,") > 0 Then '��һ
                                    strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_��ҳ��ҳһ.PDF"
                                    strPrintContent = strPrintContent & "|" & strName & ",R5," & glngSys & "," & strReportNO & "," & lngPatient & "," & lngPageID & ",3," & strFileName
                                End If
                                
                                If InStr("," & strRegRange & ",", ",54,") > 0 Then '����
                                    strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_��ҳ��ҳ��.PDF"
                                    strPrintContent = strPrintContent & "|" & strName & ",R5," & glngSys & "," & strReportNO & "," & lngPatient & "," & lngPageID & ",4," & strFileName
                                End If
                            Case "R1"               'ҽ��
                                'ϵͳ��,������,����id,��ҳid,ҽ����A0,A1/ҽ����B,PDFFileName
                                If mblnDoctorAdvice Then
                                    strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_ҽ��.PDF"
                                    strPrintContent = strPrintContent & "|" & strName & ",R1," & glngSys & ",zl1_INSIDE_1254_1," & lngPatient & "," & lngPageID & ",A0," & strFileName
                                    strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_����.PDF"
                                    strPrintContent = strPrintContent & "|" & strName & ",R1," & glngSys & ",zl1_INSIDE_1254_2," & lngPatient & "," & lngPageID & ",A1," & strFileName
                                Else
                                    strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_ҽ��.PDF"
                                    strPrintContent = strPrintContent & "|" & strName & ",R1," & glngSys & ",ZL1_INSIDE_1560," & lngPatient & "," & lngPageID & ",B," & strFileName
                                End If
                            Case "R9"               '�ٴ�·��
                                'FileName,����ID,��ҳID
                                strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_�ٴ�·��.PDF"
                                strPrintContent = strPrintContent & "|" & strName & ",R9," & glngSys & "," & strFileName & "," & lngPatient & "," & lngPageID
                            End Select
                        End If
                    Else
                        If InStr(strRange, "," & rs("�ϼ�id").Value & ",") > 0 Then
                            varParam = Split(rs("����").Value, ";")
                            Select Case rs("�ϼ�id").Value
                            Case "R2"               'סԺ����
                                'ϵͳ��,FileName,ID
                                strEprName = Split(rs("����").Value, "��")(0)
                                strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_" & strEprName & "_" & Val(varParam(0)) & ".PDF"
                                strPrintContent = strPrintContent & "|" & strName & ",R2," & glngSys & "," & strFileName & "," & Val(varParam(0))
                            Case "R3"               '������
                                'ϵͳ��,FileName,ID
                                strEprName = Split(rs("����").Value, "��")(0)
                                strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_" & strEprName & "_" & Val(varParam(0)) & ".PDF"
                                strPrintContent = strPrintContent & "|" & strName & ",R3," & glngSys & "," & strFileName & "," & Val(varParam(0))
                            Case "R4"               '�����¼
                                'ϵͳ��,�°�N/�ɰ�O,���µ�1/�����¼��2/����ͼ3,FileName,����ID,��ҳID,����ID,Ӥ�����,lngKey/lngFileID,Period
                                blnNewTends = Get�°滤��(lngPatient, lngPageID)
                                If blnNewTends = False Then
                                    If UBound(varParam) >= 1 Then
                                        If Val(varParam(1)) = -1 Then '���µ�
                                            strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_���µ�_" & Val(varParam(0)) & ".PDF"
                                            strPrintContent = strPrintContent & "|" & strName & ",R4," & glngSys & ",O,1," & strFileName & "," & lngPatient & "," & lngPageID & "," & Val(Split(varParam(0), "_")(0)) & "," & Val(varParam(4))
                                        Else '�����¼
                                            strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_�����¼_" & Val(varParam(3)) & ".PDF"
                                            strPrintContent = strPrintContent & "|" & strName & ",R4," & glngSys & ",O,2," & strFileName & "," & lngPatient & "," & lngPageID & "," & Val(Split(varParam(0), "_")(0)) & "," & Val(varParam(4)) & "," & Val(varParam(3)) & "," & CStr(varParam(2))
                                        End If
                                    End If
                                Else
                                    '�˲������� ����
                                    varParam = Split(rs("����").Value, ";")
                                    If UBound(varParam) >= 1 Then
                                        Select Case Val(varParam(1))
                                            Case -1 '���µ�
                                                intSel = 1
                                            Case 1  '����ͼ
                                                intSel = 3
                                            Case Else '��¼��
                                                intSel = 2
                                        End Select
                                        strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_" & Decode(intSel, 1, "���µ�", 2, "�����¼", "����ͼ") & "_" & Val(varParam(3)) & ".PDF"
                                        strPrintContent = strPrintContent & "|" & strName & ",R4," & glngSys & ",N," & intSel & "," & strFileName & "," & lngPatient & "," & lngPageID & "," & Val(varParam(0)) & "," & Val(varParam(4)) & "," & Val(varParam(3))
                                    End If
                                End If
                            Case "R6"               'ҽ������
                                'ϵͳ��,FileName,ID
                                strEprName = Split(rs("����").Value, "��")(0)
                                If UBound(Split(strEprName, ">")) > 0 Then
                                    strEprName = Split(strEprName, ">")(1)
                                End If
                                strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_" & strEprName & "_" & Val(varParam(0)) & ".PDF"
                                strPrintContent = strPrintContent & "|" & strName & ",R6," & glngSys & "," & strFileName & "," & Val(varParam(0))
                            Case "R7"               '����֤��
                                'ϵͳ��,FileName,ID
                                strEprName = Split(rs("����").Value, "��")(0)
                                strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_" & strEprName & "_" & Val(varParam(0)) & ".PDF"
                                strPrintContent = strPrintContent & "|" & strName & ",R7," & glngSys & "," & strFileName & "," & Val(varParam(0))
                            Case "R8"               '֪���ļ�
                                'ϵͳ��,FileName,ID
                                strEprName = Split(rs("����").Value, "��")(0)
                                strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_" & strEprName & "_" & Val(varParam(0)) & ".PDF"
                                strPrintContent = strPrintContent & "|" & strName & ",R8," & glngSys & "," & strFileName & "," & Val(varParam(0))
                            End Select
                        End If
                    End If
                    rs.MoveNext
                Loop
                
                If Not gobjEmr Is Nothing Then
                    If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then
                        Set gobjEmr = Nothing
                    End If
                    If Not gobjEmr Is Nothing Then
                        Set rs = gclsPackage.GetEmrCISStruct(lngPatient, lngPageID)
                        Do Until rs.EOF
                            strEmrId = Split(rs!����, "|")(0)
                            If InStr(strPrintContent, strEmrId) = 0 Then
                                If UBound(Split(rs!����, "|")) = 0 Then
                                    strEprName = Split(rs("����").Value, "��")(0)
                                Else
                                    strEprName = rs!Title
                                End If
                                strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_" & strEprName & "_" & strEmrId & ".PDF"
                                strPrintContent = strPrintContent & "|" & strName & ",EMR," & glngSys & "," & strFileName & "," & strEmrId
                            End If
                            rs.MoveNext
                        Loop
                    End If
                End If
                
                ''����ѭ��,ÿ10���������һ�Σ��Լ���zlCisAuditPrint��ʼ������ʱ��
                If l Mod 10 = 0 Then
                    strPrintContent = Mid(strPrintContent, 2)
                    Set objPrint = Nothing
                    Set objPrint = CreateObject("zlCisAuditPrint.clsPrint")
                    Call objPrint.PrintDocument(Me, gstrInputSeverName, gstrInputUser, gstrInputPwd, strPrintContent, "TinyPDF")
                    
                    '�²������
                    arrPar = Split(strPrintContent, "|")
                    For j = 0 To UBound(arrPar)
                        arrParOne = Split(arrPar(j), ",")
                        If arrParOne(1) = "EMR" Then             '�²���
                            Label2.Caption = "��ʼ���" & arrParOne(0) & "����"
                                                            
                            If objRichEMR Is Nothing Then
                                Set objRichEMR = DynamicCreate("zlRichEMR.clsDockContent", "�°没��", False)
                                If Not objRichEMR Is Nothing Then Call objRichEMR.Init(gobjEmr, gcnOracle, glngSys, 0)
                            End If
                            Call objRichEMR.zlShowDoc(arrParOne(4), "")
                            Call zlCommFun.PDFFile(arrParOne(3))
                            Call objRichEMR.zlPrintDoc(False, "TinyPDF")
                        End If
                    Next
                    
                    l = 0: strPrintContent = ""
                End If
            End If
        End With
    Next
    
    If l <> 0 Then
        strPrintContent = Mid(strPrintContent, 2)
        Set objPrint = Nothing
        Set objPrint = CreateObject("zlCisAuditPrint.clsPrint")
        Call objPrint.PrintDocument(Me, gstrInputSeverName, gstrInputUser, gstrInputPwd, strPrintContent, "TinyPDF")
        
        '�²������
        arrPar = Split(strPrintContent, "|")
        For j = 0 To UBound(arrPar)
            arrParOne = Split(arrPar(j), ",")
            If arrParOne(1) = "EMR" Then             '�²���
                Label2.Caption = "��ʼ���" & arrParOne(0) & "����"
                                                
                If objRichEMR Is Nothing Then
                    Set objRichEMR = DynamicCreate("zlRichEMR.clsDockContent", "�°没��", False)
                    If Not objRichEMR Is Nothing Then Call objRichEMR.Init(gobjEmr, gcnOracle, glngSys, 0)
                End If
                Call objRichEMR.zlShowDoc(arrParOne(4), "")
                Call zlCommFun.PDFFile(arrParOne(3))
                Call objRichEMR.zlPrintDoc(False, "TinyPDF")
            End If
        Next
        
        l = 0: strPrintContent = ""
    End If
    Exit Sub
errHand:
    zlCommFun.StopFlash
    If ErrCenter = 1 Then
        Resume
    End If
    Label2.Caption = ""
    mstrPrintDocIDs = ""
End Sub

Private Sub cmdOk_Click()
Dim strRange As String, strRegRange As String, i As Integer, strErr As String
Dim strParPath As String, strPrinterName As String
    
    'ͳ�Ʋ���¼��ӡ���
    strRange = ""
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) = True Then
            strRegRange = strRegRange & "," & lst.ItemData(i)
            If InStr(",5,52,53,54,", "," & lst.ItemData(i) & ",") > 0 Then '��ҳ�������棬���Ͷ���5
                If InStr(strRange, "R5") = 0 Then 'û��
                    strRange = strRange & ",R5"
                End If
            Else
                strRange = strRange & ",R" & lst.ItemData(i)
            End If
        End If
    Next
    If strRange <> "" Then
        strRange = strRange & ","
        strRegRange = Mid(strRegRange, 2)
    Else
        MsgBox "��ѡ����Ҫ����ĵ�����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mblnPDF Then
        Call SetRegister(˽��ģ��, "��ӡ����", "���PDF", strRegRange)
        '���λ��
        If txtPath.Text = "" Then
            MsgBox "��ѡ�����������λ�ã�", vbInformation, gstrSysName
            Exit Sub
        Else
            strParPath = txtPath.Text
            If gobjFSO.FolderExists(strParPath) = False Then
                MsgBox "ָ��Ŀ¼�����ڣ����飡", vbInformation, gstrSysName
                Exit Sub
            End If
            Call SetRegister(˽��ģ��, "��ӡ����", "PDFλ��", txtPath.Text)
        End If
        On Error Resume Next
        Err.Clear
        Call zlCommFun.PDFInitialize(strErr)
        If Err.Number <> 0 Then
            Err.Raise vbObjectError, , "PDF�豸��ʼ��ʧ��"
        End If
    Else
        strPrinterName = cboPrinterName.Text
        Call SetRegister(˽��ģ��, "��ӡ����", "��ӡ����", strRegRange)
        Call SetRegister(˽��ģ��, "��ӡ����", "��ӡ��", strPrinterName)
    End If
    On Error GoTo 0

    cmdCancel.Enabled = False: cmdOK.Enabled = False: fraPageScope.Enabled = False
    cmdSelectAll.Enabled = False: cmdUnSelectAll.Enabled = False: Frame2.Enabled = False
    
    If mblnPDF Then
        Call PrintWithActiveEXE(strRegRange, strRange, strParPath)
    Else
        Call PrintDocument(strRegRange, strRange, strParPath, strPrinterName)
    End If
    
    Label2.Caption = "��������"
    mstrPrintDocIDs = ""
    cmdCancel.Enabled = True: cmdOK.Enabled = True: fraPageScope.Enabled = True: cmdSelectAll.Enabled = True: cmdUnSelectAll.Enabled = True: Frame2.Enabled = True
End Sub


Private Sub PrintDocument(ByVal strRegRange As String, ByVal strRange As String, ByVal strParPath As String, ByVal strPrinterName As String)
Dim i As Integer, rs As New ADODB.Recordset, blnTrans As Boolean, lngNo As Long
Dim clsPath As zlCISPath.clsDockPath, clsTendsNew As zl9TendFile.clsTendFile, objPacsDoc As Object
Dim varParam As Variant, strReportNO As String, lng����ID As Long, blnNewTends As Boolean, intSel As Integer, strEprName As String
Dim lngPatient As Long, lngPageID As Long, lngDept As Long, lngInNo As Long, strPath As String, strFileName As String, blnDataMove As Boolean, strName As String
    
    On Error GoTo errHand

    '�������
    If mclsDockAduits Is Nothing Then
        Set mclsDockAduits = New zlRichEPR.clsDockAduits
    End If
    Set clsPath = New zlCISPath.clsDockPath
    Set clsTendsNew = New zl9TendFile.clsTendFile: Call clsTendsNew.InitTendFile(gcnOracle, glngSys)
    
    '���ô�ӡ
    For i = 1 To vsf.Rows - 1
        With vsf
            If .Cell(flexcpData, i, mCol.ѡ�� = 1) Then 'ѡ�в����
                .Row = i
                lngPatient = Val(.TextMatrix(i, mCol.����ID))
                lngInNo = Val(.TextMatrix(i, mCol.סԺ��))
                lngPageID = Val(.TextMatrix(i, mCol.��ҳID))
                lngDept = Val(.TextMatrix(i, mCol.��Ժ����ID))
                strName = NVL(.TextMatrix(i, mCol.����))
                '������ӡ
                Call gclsPackage.FuncPrintBatch(lngPatient, lngPageID, lngDept, strRange, strRegRange, mclsDockAduits, clsPath, clsTendsNew, _
                    mblnPDF, strParPath, strName, lngInNo, Me, Label2.Caption, blnDataMove, strPrinterName, mblnDoctorAdvice, mstrPrintDocIDs)
                .TopRow = i
                .Cell(flexcpData, i, mCol.ѡ��) = 0
                Set .Cell(flexcpPicture, i, mCol.ѡ��) = Nothing
            End If
        End With
    Next
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Label2.Caption = ""
    mstrPrintDocIDs = ""
End Sub

Private Sub cmdPath_Click()
Dim strPath As String
    strPath = zl9Comlib.OS.OpenDir(Me.hWnd, "��ѡ�񵼳��ļ�λ��")
    If strPath = "" Then Exit Sub
    txtPath.Text = strPath: txtPath.ToolTipText = strPath
End Sub

Private Sub cmdSelectAll_Click()
    vsf.Cell(flexcpData, 1, mCol.ѡ��, vsf.Rows - 1, mCol.ѡ��) = 1
    Set vsf.Cell(flexcpPicture, 1, mCol.ѡ��, vsf.Rows - 1, mCol.ѡ��) = img16.ListImages("Selected").Picture
End Sub

Private Sub cmdUnSelectAll_Click()
    vsf.Cell(flexcpData, 1, mCol.ѡ��, vsf.Rows - 1, mCol.ѡ��) = 0
    Set vsf.Cell(flexcpPicture, 1, mCol.ѡ��, vsf.Rows - 1, mCol.ѡ��) = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
If cmdCancel.Enabled = False Then
    Cancel = 1
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If cmdCancel.Enabled = False Then
    Cancel = 1
End If

Set mclsDockAduits = Nothing
Unload mfrmTipInfo
Set mfrmTipInfo = Nothing
End Sub

Private Sub mclsDockAduits_AfterEprPrint(ByVal lngRecordId As Long)
    mstrPrintDocIDs = mstrPrintDocIDs & lngRecordId & ","
End Sub
Private Sub picInfo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'��ʾָ�������б��е���ʷǩ����¼
Dim strTipInfo As String, lngRow As Long
    If picInfo.Visible = False Then Exit Sub
    
    lngRow = vsf.MouseRow
    If lngRow <= 0 Then Exit Sub
    
    strTipInfo = vsf.Cell(flexcpData, lngRow, mCol.��ӡ��¼)

    If strTipInfo = "" Then '���û�л�ȡ������������ȡ����¼���б���
        strTipInfo = GetPrintLog(vsf.TextMatrix(lngRow, mCol.����ID), vsf.TextMatrix(lngRow, mCol.��ҳID)) '��ȡ��ӡ��¼
        vsf.Cell(flexcpData, lngRow, mCol.��ӡ��¼) = strTipInfo
    End If
    
    mfrmTipInfo.ShowTipInfo picInfo.hWnd, strTipInfo, True
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    If picInfo.Visible Then
        picInfo.Move vsf.Cell(flexcpLeft, NewTopRow, mCol.����) + vsf.Cell(flexcpWidth, NewTopRow, mCol.����) - picInfo.Width - 30
    End If
End Sub

Private Sub vsf_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Then
        Call vsf_DblClick
    End If
End Sub

Private Sub vsf_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lngCol As Long, lngRow As Long
    lngCol = vsf.MouseCol: lngRow = vsf.MouseRow
    If lngRow <= 0 Then picInfo.Visible = False: Exit Sub
    If Val(vsf.TextMatrix(lngRow, mCol.����ID)) <> 0 Then
        If Val(picInfo.Tag) = lngRow And picInfo.Visible Then Exit Sub
        picInfo.Tag = lngRow
        picInfo.Move vsf.Cell(flexcpLeft, lngRow, mCol.����) + vsf.Cell(flexcpWidth, lngRow, mCol.����) - picInfo.Width - 30, vsf.Cell(flexcpTop, lngRow, mCol.����) + 15
        If vsf.RowSel = lngRow Then
            picInfo.BackColor = vsf.BackColorSel
        Else
            picInfo.BackColor = &H80000005
        End If
        picInfo.Visible = True
    Else
        picInfo.Visible = False
    End If
End Sub
Private Sub vsf_SelChange()
    If picInfo.Visible Then
        picInfo.BackColor = vsf.BackColorSel
    End If
End Sub
Private Sub vsf_DblClick()
Dim lngRow As Long, l As Long, lCheck As Long
    With vsf
        lngRow = .Row
        If lngRow < 1 Then Exit Sub
        If .Cell(flexcpData, lngRow, mCol.ѡ��) = 0 Then
            .Cell(flexcpData, lngRow, mCol.ѡ��) = 1
            Set .Cell(flexcpPicture, lngRow, mCol.ѡ��) = img16.ListImages("Selected").Picture
        Else
            .Cell(flexcpData, lngRow, mCol.ѡ��) = 0
            Set .Cell(flexcpPicture, lngRow, mCol.ѡ��) = Nothing
        End If
        
        For l = 1 To .Rows - 1
            If .Cell(flexcpData, l, mCol.ѡ��) = 1 Then
                lCheck = lCheck + 1
            End If
        Next
        Frame2.Caption = "����������嵥(����������������)" & " ��" & .Rows - 1 & "�У���ѡ��" & lCheck & "��"
    End With
End Sub
Private Function GetPrintLog(ByVal lngPatient As Long, ByVal lngPageID As Long) As String
Dim rs As New ADODB.Recordset
    gstrSQL = "Select ��ӡ���� As ��ӡ��, ��ӡ����, ��ӡ��, ��ӡʱ�� From ������ӡ��¼ Where ����id = [1] And ��ҳid = [2] Order By ��ӡʱ��, ��ӡ���"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatient, lngPageID)
    Do Until rs.EOF
        GetPrintLog = GetPrintLog & vbCrLf & Rpad(rs!��ӡ��, 10) & Rpad(Format(rs!��ӡʱ��, "yyyy-mm-dd hh:MM"), 20) & Rpad(rs!��ӡ����, 40)
        rs.MoveNext
    Loop
    GetPrintLog = Rpad("��ӡ��", 10) & Rpad("��ӡʱ��", 20) & Rpad("��ӡ����", 40) & GetPrintLog
End Function
