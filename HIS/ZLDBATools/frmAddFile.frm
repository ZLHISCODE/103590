VERSION 5.00
Begin VB.Form frmAddFile 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�����ļ����"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   6525
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2460
      Width           =   6525
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Default         =   -1  'True
         Height          =   345
         Left            =   4080
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   345
         Left            =   5280
         TabIndex        =   7
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblPgs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   195
         Width           =   45
      End
      Begin VB.Line lineBottom 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   10000
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line lineBottom 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   10000
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.TextBox txtDataFile 
      Height          =   300
      Left            =   1710
      TabIndex        =   2
      Top             =   1560
      Width           =   3945
   End
   Begin VB.CheckBox chkSpaceExtd 
      Caption         =   "�Զ���չ�ռ�"
      Height          =   270
      Left            =   2880
      TabIndex        =   5
      ToolTipText     =   "AUTOEXTEND ON Next (��ռ��С/10)M"
      Top             =   1965
      Value           =   1  'Checked
      Width           =   1425
   End
   Begin VB.TextBox txtSpaceSize 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1710
      MaxLength       =   6
      TabIndex        =   3
      Text            =   "500"
      Top             =   1950
      Width           =   735
   End
   Begin VB.TextBox txtTableSpace 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1710
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   2160
   End
   Begin VB.TextBox txtFileAmount 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   1710
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "1"
      Top             =   1230
      Width           =   300
   End
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ļ������ļ�ĩβ���ֵ�����"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2640
      TabIndex        =   13
      Top             =   1290
      Width           =   2520
   End
   Begin VB.Label lblPrompt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ϊ��ǰ��ռ���������ļ�"
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   240
      Width           =   3975
   End
   Begin VB.Image Img 
      Height          =   480
      Left            =   240
      Picture         =   "frmAddFile.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblDataFile 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��һ���ļ�"
      Height          =   180
      Left            =   720
      TabIndex        =   10
      Top             =   1620
      Width           =   900
   End
   Begin VB.Label lblBakSpace 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ݱ�ռ���"
      Height          =   225
      Left            =   480
      TabIndex        =   8
      Top             =   900
      Width           =   1125
   End
   Begin VB.Label lblFileAmount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����          ���ļ�"
      Height          =   195
      Index           =   0
      Left            =   1065
      TabIndex        =   4
      Top             =   1290
      Width           =   1530
   End
   Begin VB.Label lblFileSize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ʼ��С                     M"
      Height          =   195
      Left            =   855
      TabIndex        =   9
      Top             =   2010
      Width           =   1785
   End
End
Attribute VB_Name = "frmAddFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnCreate As Boolean

Public Function ShowAddFile(ByVal strTableSpace As String) As Boolean
    
    txtTableSpace.Text = strTableSpace
    txtDataFile.Text = GetFileName(, strTableSpace)
    
    Me.Show 1
    ShowAddFile = mblnCreate
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function GetFileName(Optional ByVal strFile As String, Optional ByVal strTableSpace As String) As String
    '���ݵ�ǰ�������ļ�����,��ȡ��һ�������ļ�
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strTmp As String, i As Integer
    
    If strFile = "" Then
        strSQL = "Select Max(File_Name) Max_File From Dba_Data_Files Where Tablespace_Name =[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "��ȡ�����ļ���", strTableSpace)
        strFile = rsTmp!Max_file
    End If
    
    If InStr(1, strFile, ".DBF") > 0 Then
        strFile = Left(strFile, InStr(1, strFile, ".DBF") - 1)
    End If
    
    If IsNumeric(Right(strFile, 4)) Then
        '����λΪ����,���������� ZLHD2017\2018 ���ְ����Ϊ����ı��������ļ�
        strFile = strFile & "_01.DBF"
    Else
        '����,ȡĩ������+1
        i = 1
        Do While IsNumeric(Right(strFile, i))
            i = i + 1
        Loop
        
        If i = 1 Then
            'û������
            strFile = strFile & "01.DBF"
        Else
            strTmp = Format(Val(Right(strFile, i - 1)) + 1, Lpad("", i - 1, "0"))
            strFile = Left(strFile, Len(strFile) - i + 1) & strTmp & ".DBF"
        End If
    End If
    
    GetFileName = strFile
End Function


Private Sub cmdOK_Click()
    
    '���ݼ��
    If Len(Trim(txtDataFile.Text)) = 0 Then
        MsgBox "�붨��" & txtTableSpace.Text & "��ռ�������ļ���", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If Val(txtSpaceSize.Text) > 32000 Then
        MsgBox "��ռ�" & txtTableSpace.Text & "����32G�ˡ�", vbExclamation, gstrSysName
        Exit Sub
    End If

    
    If AddDatafile(txtTableSpace.Text, txtDataFile.Text, txtFileAmount.Text, txtSpaceSize.Text, chkSpaceExtd.Value) Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub txtDataFile_GotFocus()
    txtDataFile.SelStart = Len(txtDataFile.Text)
End Sub

Private Sub txtDataFile_KeyPress(KeyAscii As Integer)
    OnlyStrCK KeyAscii, "\", "_", "/"
End Sub

Private Sub txtFileAmount_GotFocus()
    txtFileAmount.SelStart = Len(txtFileAmount.Text)
End Sub

Private Sub txtFileAmount_KeyPress(KeyAscii As Integer)
    OnlyIntCK KeyAscii
End Sub

Private Sub txtSpaceSize_GotFocus()
    txtSpaceSize.SelStart = Len(txtSpaceSize.Text)
End Sub

Private Function AddDatafile(ByVal strTableSpace As String, ByVal strFile As String, ByVal intNum As Integer, ByVal lngSize As Long, ByVal blnAutoExtend As Boolean) As Boolean
    'Ϊ��ռ���������ļ�
    '����:strTableSpace - ��ռ�����,strFile - �׸��ļ��� , intNum - ����ļ����� ,lngSize  - ��ʼ�ļ���С, blnAutoExtend - �Ƿ��Զ���չ
    Dim strErrMsg As String, strSQL As String
    Dim strNextFile As String, i As Integer, strTmp As String
    
    On Error Resume Next
    
    lblPgs.Caption = "���ڴ��������ļ�������"
    
    For i = 1 To intNum
        If strNextFile = "" Then
            strNextFile = strFile
        Else
            strNextFile = GetFileName(strNextFile)
        End If
        
        strTmp = IIf(InStr(1, strNextFile, "\") > 0, "\", "/")
        strTmp = Mid(strNextFile, InStrRev(strNextFile, strTmp) + 1, InStr(1, strNextFile, ".") - 1)
        lblPgs.Caption = "���ڴ��������ļ�" & strTmp & "������"
        lblPgs.Refresh
        
        strSQL = "Alter TableSpace " & strTableSpace & " Add DataFile '" & strNextFile & "' Size " & lngSize & "M  AutoExtend  " & IIf(blnAutoExtend, "On", "Flase")
        gcnOracle.Execute strSQL
        
        If Err.Number <> 0 Then
            strErrMsg = "��������ļ� " & strTmp & "�������� ����ԭ�� ��" & vbNewLine & Err.Description
            
             If MsgBox(strErrMsg & vbNewLine & "�Ƿ�����������������ļ�������ǽ����������ȡ�����˳���ǰ������", vbYesNo, "����") = vbYes Then
                strErrMsg = ""
                Err.Clear
            Else
                lblPgs.Caption = "������ȡ��"
                Exit Function
            End If
        End If
    Next
    
    mblnCreate = True
    AddDatafile = mblnCreate
End Function

Private Sub txtSpaceSize_KeyPress(KeyAscii As Integer)
    OnlyIntCK KeyAscii
End Sub
