VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{B26F6243-4C7D-11D1-910E-00600807163F}#2.78#0"; "Xcdzip35.ocx"
Begin VB.Form frmPriceImp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ�����ݵ���"
   ClientHeight    =   3075
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6210
   FillColor       =   &H80000012&
   Icon            =   "frmPriceImp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ر�(&X)"
      Height          =   350
      Left            =   4845
      TabIndex        =   9
      Top             =   2550
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   135
      Picture         =   "frmPriceImp.frx":058A
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2550
      Width           =   1100
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "����(&I)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   3735
      TabIndex        =   7
      Top             =   2550
      Width           =   1100
   End
   Begin VB.Frame frmLine 
      Height          =   45
      Left            =   -30
      TabIndex        =   5
      Top             =   2370
      Width           =   6360
   End
   Begin VB.TextBox txtFile 
      Height          =   300
      Left            =   1110
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1245
      Width           =   4785
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "ѡ��(&S)��"
      Height          =   350
      Left            =   4815
      TabIndex        =   2
      Top             =   885
      Width           =   1100
   End
   Begin MSComctlLib.ProgressBar pgbImp 
      Height          =   240
      Left            =   1110
      TabIndex        =   4
      Top             =   2115
      Visible         =   0   'False
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComDlg.CommonDialog cdgThis 
      Left            =   120
      Top             =   1620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblImp 
      AutoSize        =   -1  'True
      Caption         =   "���ڵ���ҽ������"
      Height          =   180
      Left            =   1095
      TabIndex        =   6
      Top             =   1875
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      Caption         =   "��׼ҽ���ļ�"
      Height          =   180
      Left            =   1095
      TabIndex        =   1
      Top             =   975
      Width           =   1080
   End
   Begin XCEEDZIPLib.XceedZip zip 
      Left            =   135
      Top             =   825
      _Version        =   131150
      _ExtentX        =   794
      _ExtentY        =   794
      _StockProps     =   0
      Compression     =   6
      ClearDisks      =   0   'False
      ExtractDirectory=   ""
      FilesToProcess  =   ""
      IncludeDirectoryEntries=   0   'False
      IncludeHiddenFiles=   0   'False
      IncludeVolumeLabel=   0   'False
      ModifiedDate    =   "01011980"
      MoveFiles       =   0   'False
      MultidiskMode   =   0   'False
      Overwrite       =   0
      Password        =   ""
      Recurse         =   0   'False
      SelfExtracting  =   0   'False
      SfxBinary       =   ""
      SfxConfigFile   =   ""
      StoredExtensions=   ".ZIP;.LZH;.ARC;.ARJ;.ZOO"
      TempPath        =   ""
      UsePaths        =   -1  'True
      UseTempFile     =   -1  'True
      ZipFileName     =   ""
      InternalState   =   "7f6ba9d4"
      SfxExtractDirectory=   ""
      SfxRunExePath   =   ""
      SfxReadmePath   =   ""
      SfxDefaultPassword=   ""
      SfxOverwrite    =   0
      SfxPromptForDirectory=   -1  'True
      SfxShowProgress =   -1  'True
      SfxPromptForPassword=   -1  'True
      SfxPromptCreateDirectory=   -1  'True
      SfxProgramGroup =   ""
      SfxProgramGroupItems=   ""
      SfxRegisterExtensions=   ""
      SfxInstallMode  =   0   'False
   End
   Begin VB.Image imgICO 
      Height          =   480
      Left            =   135
      Picture         =   "frmPriceImp.frx":0E54
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblNote 
      Caption         =   "    ѡ����ȷ��ҽ�������ļ�����ҽ�������������뱾ϵͳ���Ա㱣֤ϵͳ���շ���Ŀ�ͼ۸���ϼ۸����ߵĹ涨��"
      Height          =   390
      Left            =   705
      TabIndex        =   0
      Top             =   165
      Width           =   5220
   End
End
Attribute VB_Name = "frmPriceImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private strTmpPath As String

Dim objFile As New FileSystemObject
Dim objText As TextStream

Private Function GetDateString(ByVal strDateString As String) As String
    'һ����ʱ�İ����봮ת��Ϊ��׼���ڸ�ʽ���ִ��ĺ�������Ҫ������ȥ������
    'ֻ���һ������ڸ�ʽ������"2005-8-15 13:51:32:953"��ȥ���������"2005-8-15 13:51:32"
    Dim strInput As String      '���������ִ�
    Dim strOutput As String     '����ִ�
    Dim strDatePart As String   '���ڲ���
    Dim strTimePart As String   'ʱ�䲿��
    Dim intSpace As Integer     '�ո�����λ��
    Const cstDateTimeFormat = "yyyy-mm-dd hh:mm:ss"     '��׼������ʱ���ʽ
    Const cstDateFormat = "yyyy-mm-dd"                  '��׼�����ڸ�ʽ
    Const cstTimeFormat = "hh:mm:ss"                    '��׼��ʱ���ʽ
    Dim strTime() As String     '��ʱ����ʱ��ָ�������
    
    '��ȥ����β�ո�
    strInput = Trim(strDateString)
    
    '���Ϊ�վ��˳�
    If strInput = "" Then
        strOutput = ""
        GetDateString = strOutput
        Exit Function
    End If
    
    '������봮����ת��Ϊ���ڣ���ֱ��ת��Ϊ��׼�����ڸ�ʽ�ִ����
    If IsDate(strInput) Then
        strOutput = Format(CDate(strInput), cstDateTimeFormat)
        GetDateString = strOutput
        Exit Function
    End If
    
    '�ж����봮�м��Ƿ���ڿո�
    intSpace = InStr(strInput, " ")
    If intSpace > 0 Then    '���ڿո�ͷָ�Ϊ���ں�ʱ�䲿��
        strDatePart = Mid(strInput, 1, intSpace - 1)
        strTimePart = Mid(strInput, intSpace + 1)
    Else    'û�пո���Ȼ������ȷ�����ڴ���ֻ���˳�
        GetDateString = ""
        Exit Function
    End If
    
    '�ж������ִ������Ƿ����ת��Ϊ����
    If IsDate(strDatePart) Then
        strDatePart = Format(CDate(strDatePart), cstDateFormat)
    Else    '����ת����ֻ���˳�
        GetDateString = ""
        Exit Function
    End If
    
    '����:�ָ���ʱ�䲿�ַֽ������
    strTime = Split(strTimePart, ":")
    
    If UBound(strTime) > 2 Then     '����������޴���2�����ֽ����4�����֣�Ҳ��˵�����к���
        'ֻ����ʱ���벿��
        strTimePart = strTime(0) & ":" & strTime(1) & ":" & strTime(2)
    End If
    
    '�ж�ʱ���ִ������Ƿ����ת��Ϊ����
    If IsDate(strTimePart) Then
        strTimePart = Format(CDate(strTimePart), cstTimeFormat)
    Else    '����ת����ֻ���˳�
        GetDateString = ""
        Exit Function
    End If
    
    '��������ִ�
    strOutput = strDatePart & " " & strTimePart
    
    GetDateString = strOutput
    
End Function
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdExecute_Click()
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim lngCount As Long
    Dim strLine As String
    Dim aryField() As String
    
    If MsgBox("������ڵ����׼ҽ���ļ���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    Err = 0: On Error Resume Next
    Set objText = objFile.OpenTextFile(strTmpPath & "\item.txt")
    If Err <> 0 Then
        MsgBox "�޷���ҽ�۱�׼�ļ���", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    Do Until objText.AtEndOfStream
        objText.ReadLine
    Loop
    lngCount = objText.Line
    objText.Close
    
    '��ʼ����
    Err = 0: On Error GoTo ErrHand
    Set objText = objFile.OpenTextFile(strTmpPath & "\item.txt")
    
    Me.lblImp.Visible = True: Me.pgbImp.Visible = True
    DoEvents
    
    gcnOracle.BeginTrans
    gstrSQL = "Delete From ��׼ҽ�۹淶"
    gcnOracle.Execute gstrSQL
    
    Do While Not objText.AtEndOfStream
        strLine = Trim(objText.ReadLine)
        aryField = Split(strLine, vbTab)
        gstrSQL = "Insert Into ��׼ҽ�۹淶(��Ŀ����, ��Ŀ����, ƴ����, ��Ŀ����, �Ƽ۵�λ, ��Ŀ�ں�, ��������, ��Ŀ˵��, ��Ŀ�۸�, �ظ���־, ҽԺ�ȼ�, ע����־, �������, ����޼�, ����޼�, ��������)"
        gstrSQL = gstrSQL & " Values('" & Trim(Replace(aryField(0), "'", "''")) & "'"
        gstrSQL = gstrSQL & ",'" & Trim(Replace(aryField(1), "'", "''")) & "'"
        gstrSQL = gstrSQL & ",'" & Trim(Replace(aryField(2), "'", "''")) & "'"
        gstrSQL = gstrSQL & ",'" & Trim(Replace(aryField(3), "'", "''")) & "'"
        gstrSQL = gstrSQL & ",'" & Trim(Replace(aryField(4), "'", "''")) & "'"
        gstrSQL = gstrSQL & ",'" & Trim(Replace(aryField(5), "'", "''")) & "'"
        gstrSQL = gstrSQL & ",'" & Trim(Replace(aryField(6), "'", "''")) & "'"
        gstrSQL = gstrSQL & ",'" & Trim(Replace(aryField(7), "'", "''")) & "'"
        gstrSQL = gstrSQL & "," & Format(IIf(Not IsNumeric(Replace(aryField(8), "'", "''")), 0, Replace(aryField(8), "'", "''")), "0.00")
        gstrSQL = gstrSQL & ",'" & Trim(Replace(aryField(9), "'", "''")) & "'"
        gstrSQL = gstrSQL & ",'" & Trim(Replace(aryField(10), "'", "''")) & "'"
        gstrSQL = gstrSQL & ",'" & Trim(Replace(aryField(11), "'", "''")) & "'"
        gstrSQL = gstrSQL & ",'" & Trim(Replace(aryField(12), "'", "''")) & "'"
        gstrSQL = gstrSQL & "," & Format(IIf(Not IsNumeric(Replace(aryField(13), "'", "''")), 0, Replace(aryField(13), "'", "''")), "0.00")
        gstrSQL = gstrSQL & "," & Format(IIf(Not IsNumeric(Replace(aryField(14), "'", "''")), 0, Replace(aryField(14), "'", "''")), "0.00")
        If InStr(1, aryField(15), ".") > 0 Then
            aryField(15) = Mid(aryField(15), 1, InStr(1, aryField(15), ".") - 1)
        End If
'        gstrSQL = gstrSQL & ",to_date('" & Format(aryField(15), "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS'))"
        If GetDateString(aryField(15)) = "" Then
            gstrSQL = gstrSQL & ",NULL)"
        Else
            gstrSQL = gstrSQL & ",to_date('" & GetDateString(aryField(15)) & "','YYYY-MM-DD HH24:MI:SS'))"
        End If
        gcnOracle.Execute gstrSQL
        Me.pgbImp.Value = Int(objText.Line / lngCount * 100)
    Loop
    gcnOracle.CommitTrans
    objText.Close
    
    MsgBox "��׼ҽ�۵���ɹ���ɣ�", vbExclamation, gstrSysName
    Me.lblImp.Visible = False: Me.pgbImp.Visible = False
    Exit Sub

ErrHand:
    gcnOracle.RollbackTrans
    objText.Close
    MsgBox "��׼ҽ�۵���ʧ�ܣ���ϵͳ����Ա���ҽ���ļ���", vbExclamation, gstrSysName
    Me.lblImp.Visible = False: Me.pgbImp.Visible = False
End Sub

Private Sub cmdFile_Click()
    With Me.cdgThis
        .FileName = Me.txtFile.Text
        .DialogTitle = "ѡ���׼ҽ���ļ�"
        .Filter = "(��׼ҽ���ļ�)|*.zl"
        .ShowOpen
        If .FileName = "" Then
            Exit Sub
        Else
            Me.txtFile.Text = .FileName
        End If
        If Dir(Me.txtFile.Text) = "" Then
            MsgBox "ҽ���ļ������ڣ�", vbExclamation, gstrSysName
            Me.txtFile.Text = ""
            Me.cmdExecute.Enabled = False
            Exit Sub
        End If
    End With
    
    Err = 0: On Error Resume Next
    Kill strTmpPath & "\item.txt"
    Err = 0: On Error GoTo 0
    With Me.zip
        .FilesToProcess = "*"
        .Password = "zlhis"
        .UsePaths = False
        .ZipFileName = Me.txtFile.Text
        .ExtractDirectory = strTmpPath
        .Extract (0)
    End With
    
    If Dir(strTmpPath & "\item.txt") = "" Then
        MsgBox "���ļ�������ȷ��ҽ���ļ���", vbExclamation, gstrSysName
        Me.txtFile.Text = ""
        Me.cmdExecute.Enabled = False
    Else
        Me.cmdExecute.Enabled = True
    End If
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Form_Load()
    Dim strInput As String * 255
    Call GetTempPath(255, strInput)
    strTmpPath = Left(strInput, InStr(strInput, Chr(0)) - 1)
End Sub
