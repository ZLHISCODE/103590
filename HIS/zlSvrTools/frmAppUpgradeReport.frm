VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAppUpgradeReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ϵͳ������Ǩ"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9345
   Icon            =   "frmAppUpgradeReport.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdSelect 
      Caption         =   "��ѡ(&R)"
      Height          =   350
      Index           =   2
      Left            =   2610
      TabIndex        =   4
      Top             =   5760
      Width           =   1100
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "ȫ��(&C)"
      Height          =   350
      Index           =   1
      Left            =   1380
      TabIndex        =   3
      Top             =   5760
      Width           =   1100
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "ȫѡ(&S)"
      Height          =   350
      Index           =   0
      Left            =   150
      TabIndex        =   2
      Top             =   5760
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshReport 
      Height          =   4635
      Left            =   120
      TabIndex        =   5
      Top             =   900
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   8176
      _Version        =   393216
      Rows            =   4
      Cols            =   5
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      GridColor       =   8421504
      GridColorFixed  =   8421504
      GridColorUnpopulated=   8421504
      FocusRect       =   2
      GridLinesFixed  =   1
      ScrollBars      =   2
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8070
      TabIndex        =   1
      Top             =   5760
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6750
      TabIndex        =   0
      Top             =   5760
      Width           =   1100
   End
   Begin MSComctlLib.ProgressBar prgImport 
      Height          =   345
      Left            =   150
      TabIndex        =   7
      Top             =   5760
      Visible         =   0   'False
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   270
      Picture         =   "frmAppUpgradeReport.frx":014A
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lbl˵�� 
      Caption         =   "�±��г���ѡ����Ҫ"
      Height          =   615
      Left            =   1020
      TabIndex        =   6
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "frmAppUpgradeReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mstrPath As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strError As String, objReport As Object
    Dim lngRow As Long
    
    prgImport.Max = mshReport.Rows - 1
    prgImport.Visible = True
    prgImport.ZOrder
    MousePointer = vbHourglass
    
    If gobjReport Is Nothing Then
        Set gobjReport = CreateObject("zl9Report.clsReport")
    End If
    With mshReport
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, 0) = "��" Then
                If gobjReport.ReportImport(mstrPath & .TextMatrix(lngRow, 3), gcnOracle, .TextMatrix(lngRow, 1), .TextMatrix(lngRow, 4) = "��") = False Then
                    strError = strError & vbCrLf & .TextMatrix(lngRow, 1) & Space(30 - Len(.TextMatrix(lngRow, 1))) & .TextMatrix(lngRow, 2)
                End If
            End If
            
            prgImport.Value = lngRow
            DoEvents
        Next
    End With
    
    MousePointer = vbDefault
    
    If strError <> "" Then
        MsgBox "���ֱ����ڵ���ʱ���ִ������飺" & vbCrLf & vbCrLf & strError, vbInformation, gstrSysName
    End If
    
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdSelect_Click(Index As Integer)
    Dim lngRow As Long
    
    With mshReport
        For lngRow = 1 To .Rows - 1
            Select Case Index
                Case 0
                    .TextMatrix(lngRow, 0) = "��"
                Case 1
                    .TextMatrix(lngRow, 0) = ""
                Case 2
                    .TextMatrix(lngRow, 0) = IIf(.TextMatrix(lngRow, 0) = "��", "", "��")
            End Select
        Next
    End With
End Sub

Private Sub Form_Load()
    lbl˵��.Caption = "    �±��г���������ʱ��Ҫ����ı���������ʹ��Ĭ�ϵ�ѡ�񡣵�Ȼ����Ҳ���ܸ��ݾ������ȡ�����֣��Ժ�����ڱ������������ֹ����롣���ڸ�ʽ�ڱ���λ�����޸ĵı�������ѡ��ֻ��������Դ�����������Ļ�����ڱ������ƽ����������Ч��Ԫ�ء�"
End Sub

Public Function UpdateReport(ByVal strInstallFile As String, ByVal strVerSource As String, ByVal strVerDest As String) As Boolean
'���ܣ���ɱ��������
'������strSetupFile   ��װ�����ļ�
'      strVerSource   ����ǰ�İ汾
'      strVerDest     ������İ汾
    Dim objSys As New Scripting.FileSystemObject, objFolder As Scripting.Folder, objText As Scripting.TextStream
    Dim rsReports As New ADODB.Recordset, strLine As String, str��� As String
    Dim lngRow As Long, varLine As Variant, varArray As Variant
    Dim dblVerSource As Double, dblVerDest As Double, dblVer As Double
    
    On Error Resume Next
    '��õ��������ļ����Ŀ¼
    mstrPath = Left(strInstallFile, Len(strInstallFile) - Len("zlSetup.ini")) & "..\��������"
    Set objFolder = objSys.GetFolder(mstrPath)
    If Err <> 0 Then
        MsgBox "�򿪵��������ļ����Ŀ¼����", vbInformation, gstrSysName
        Exit Function
    End If
    mstrPath = objFolder.Path & "\" '�õ�ַ�Ѿ������ˣ�ȥ��..
    
    Set objText = objSys.OpenTextFile(mstrPath & "zlReport.ini")
    If Err <> 0 Then
        MsgBox "�򿪵�������˵���ļ�zlReport.ini����", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�õ�����Щ������Ҫ����������
    dblVerSource = GetVerDouble(strVerSource)
    dblVerDest = GetVerDouble(strVerDest)
    
    rsReports.Fields.Append "���", adVarChar, 50
    rsReports.Fields.Append "����", adVarChar, 500
    rsReports.Open
    
    Do Until objText.AtEndOfStream
        strLine = Trim(objText.ReadLine)
        
        If Left(strLine, 1) = "[" And Right(strLine, 1) = "]" Then
            'ȡ�ð汾��
            dblVer = GetVerDouble(Mid(strLine, 2, Len(strLine) - 2))
        ElseIf InStr(strLine, "|") > 0 Then  '�����п�������ע����
            If dblVer > dblVerSource And dblVer <= dblVerDest Then
                '�ô����ϰ汾Ҫ�󣺱ȵ�ǰ�汾�£�����С�ڵ������°汾
                str��� = Split(strLine, "|")(0)
                
                rsReports.Filter = "���='" & str��� & "'"
                If rsReports.EOF = False Then
                    '�ñ�������Ѿ����ڣ������
                    rsReports("����") = strLine
                Else
                    rsReports.AddNew Array("���", "����"), Array(str���, strLine)
                End If
            End If
        End If
    Loop
    
    rsReports.Filter = 0
    If rsReports.RecordCount > 0 Then
        '�б�������
        mshReport.Rows = rsReports.RecordCount + 1
        Call InitTable
        
        '��д����
        lngRow = 1
        rsReports.Sort = "���"
        rsReports.MoveFirst
        Do Until rsReports.EOF
            varArray = Split(rsReports("����"), "|")
            mshReport.TextMatrix(lngRow, 0) = "��"
            mshReport.TextMatrix(lngRow, 1) = varArray(0)
            mshReport.TextMatrix(lngRow, 2) = varArray(1)
            mshReport.TextMatrix(lngRow, 3) = varArray(2)
            mshReport.TextMatrix(lngRow, 4) = IIf(varArray(3) = "1", "��", "")
            
            lngRow = lngRow + 1
            rsReports.MoveNext
        Loop
        
        frmAppUpgradeReport.Show vbModal
        UpdateReport = mblnOK
    Else
        '���봦��
        UpdateReport = True
    End If
End Function

Private Sub InitTable()
'���ܣ��Ա��ĸ�ʽ���г�ʼ
    With mshReport
        '����
        .TextMatrix(0, 0) = "�Ƿ�����"
        .TextMatrix(0, 1) = "������"
        .TextMatrix(0, 2) = "��������"
        .TextMatrix(0, 3) = "�ļ���"
        .TextMatrix(0, 4) = "ֻ��������Դ"
        
        .ColWidth(0) = 1000
        .ColWidth(1) = 2400
        .ColWidth(2) = 4000
        .ColWidth(3) = 0
        .ColWidth(4) = 1200
        
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignCenterCenter
        
        '��ͷ���ж���
        .Col = 0
        .Row = 0
        .ColSel = .Cols - 1
        .RowSel = 0
        .FillStyle = flexFillRepeat
        .CellAlignment = flexAlignCenterCenter
        
        '��һ��
        .Col = 0
        .Row = 1
        .ColSel = 0
        .RowSel = .Rows - 1
        .FillStyle = flexFillRepeat
        .CellAlignment = flexAlignCenterCenter
        .CellForeColor = RGB(255, 0, 0)
        .CellFontBold = True
'        .CellTextStyle
        
        
        .FillStyle = flexFillSingle
        .AllowBigSelection = False
        .Row = .FixedRows: .Col = 0
    End With
End Sub

Private Sub mshReport_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeySpace Then Exit Sub
    
    With mshReport
        If .Col = 0 Or .Col = .Cols - 1 Then
            .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = "��", "", "��")
        End If
    End With
End Sub

Private Sub mshReport_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    With mshReport
        If (.MouseCol = 0 Or .MouseCol = .Cols - 1) And .MouseRow > 0 Then
            .TextMatrix(.MouseRow, .MouseCol) = IIf(.TextMatrix(.MouseRow, .MouseCol) = "��", "", "��")
        End If
    End With
End Sub
