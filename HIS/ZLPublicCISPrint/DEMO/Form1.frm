VERSION 5.00
Begin VB.Form frmPDF 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PDF���"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8340
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtXML 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   2400
      Width           =   8055
   End
   Begin VB.TextBox txtList 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   5400
      Width           =   8055
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "�ļ��嵥"
      Height          =   360
      Left            =   5760
      TabIndex        =   16
      Top             =   4800
      Width           =   1100
   End
   Begin VB.TextBox txtPath 
      Height          =   300
      Left            =   960
      TabIndex        =   10
      Text            =   "C:\Users\Administrator\Desktop\PDF\PDF���"
      Top             =   1620
      Width           =   7215
   End
   Begin VB.CheckBox chkMerge 
      Caption         =   "�����ϲ�"
      Height          =   255
      Left            =   4320
      TabIndex        =   9
      Top             =   923
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.TextBox Text 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   5
      Left            =   5280
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "HIS"
      Top             =   113
      Width           =   855
   End
   Begin VB.TextBox Text 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "aqa"
      Top             =   113
      Width           =   855
   End
   Begin VB.TextBox Text 
      Height          =   300
      Index           =   3
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":6852
      Top             =   113
      Width           =   855
   End
   Begin VB.TextBox Text 
      Height          =   300
      Index           =   2
      Left            =   7320
      TabIndex        =   3
      Text            =   "testbase"
      Top             =   113
      Width           =   855
   End
   Begin VB.TextBox Text 
      Height          =   300
      Index           =   1
      Left            =   2520
      TabIndex        =   6
      Text            =   "1"
      Top             =   900
      Width           =   855
   End
   Begin VB.TextBox Text 
      Height          =   300
      Index           =   0
      Left            =   720
      TabIndex        =   4
      Text            =   "4211"
      Top             =   900
      Width           =   855
   End
   Begin VB.CommandButton cmdPDF 
      Caption         =   "���PDF"
      Height          =   360
      Left            =   7080
      TabIndex        =   15
      Top             =   4800
      Width           =   1100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "���PDFʱ����XML:˫�������ı�����XMLʾ��"
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   120
      TabIndex        =   20
      Top             =   2160
      Width           =   3780
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "ִ���ļ��嵥�õ���XML�ַ���"
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   120
      TabIndex        =   18
      Top             =   5040
      Width           =   2430
   End
   Begin VB.Label lblPati 
      AutoSize        =   -1  'True
      Caption         =   "���λ��"
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   720
   End
   Begin VB.Label lblPati 
      AutoSize        =   -1  'True
      Caption         =   "���ݿ�����"
      Height          =   180
      Index           =   5
      Left            =   4320
      TabIndex        =   13
      Top             =   180
      Width           =   900
   End
   Begin VB.Label lblPati 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   4
      Left            =   2040
      TabIndex        =   12
      Top             =   180
      Width           =   360
   End
   Begin VB.Label lblPati 
      AutoSize        =   -1  'True
      Caption         =   "�û�"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   173
      Width           =   360
   End
   Begin VB.Label lblPati 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   2
      Left            =   6840
      TabIndex        =   8
      Top             =   180
      Width           =   360
   End
   Begin VB.Label lblPati 
      AutoSize        =   -1  'True
      Caption         =   "��ҳID"
      Height          =   180
      Index           =   1
      Left            =   1920
      TabIndex        =   7
      Top             =   960
      Width           =   540
   End
   Begin VB.Label lblPati 
      AutoSize        =   -1  'True
      Caption         =   "����ID"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   540
   End
End
Attribute VB_Name = "frmPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPreInfo As String
Private mobjPrint As Object

Private Sub cmdPDF_Click()
          
          '2559,1,C:\Users\Administrator\Desktop\PDF\�¶���_16100006_1_��ҳ����.PDF
          Dim objPrint As Object
          Dim strNoPDF As String

   On Error GoTo errH

10        If Text(3).Text = "" Then
20            MsgBox "�û�����Ϊ�գ�"
30            Exit Sub
40        End If
50        If Text(4).Text = "" Then
60            MsgBox "���벻��Ϊ�գ�"
70            Exit Sub
80        End If
90        If Text(2).Text = "" Then
100           MsgBox "������Ϊ�գ�"
110           Exit Sub
120       End If
          
130       If Text(0).Text = "" Then
140           MsgBox "����ID����Ϊ�գ�"
150           Exit Sub
160       End If
          
170       If Text(1).Text = "" Then
180           MsgBox "��ҳID����Ϊ�գ�"
190           Exit Sub
200       End If
          
210       If txtPath.Text = "" Then
220           MsgBox "PDF���λ�ò���Ϊ�գ�"
230           Exit Sub
240       End If
250       On Error GoTo errH
260       Set mobjPrint = Nothing
270       cmdPDF.Enabled = False
280       Set mobjPrint = CreateObject("ZLPublicCISPrint.clsPrint")
290       If Not mobjPrint.InitPrint(Trim(Text(2).Text), Trim(Text(3).Text), Trim(Text(4).Text), Trim(Text(5).Text)) Then Exit Sub
       
300       Call mobjPrint.PrintDocument(Val(Text(0).Text), Val(Text(1).Text), Trim(txtPath.Text), Trim(txtXML.Text), chkMerge.Value = vbChecked, strNoPDF)
       
310       cmdPDF.Enabled = True
320       Exit Sub
errH:
         
330       cmdPDF.Enabled = True

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdPDF_Click of Form frmPDF"
End Sub

Public Function GetRegister(ByVal enmRegister As REGISTER, ByVal strSection As String, ByVal strKey As String, ByVal strDefKeyValue As String, Optional ByVal strUser As String) As String
    '******************************************************************************************************************
    '���ܣ� ��ָ����ע����Ϣ��ȡ����
    '������ enmRegister-ע������
    '       strSection-ע���Ŀ¼
    '       strKey-����
    '       strDefKeyValue-ȱʡ��ֵ
    '���أ� strKeyValue-��ֵ
    '******************************************************************************************************************

    Dim strValue As String
    
    On Error GoTo errHand
    
    Select Case enmRegister
    Case ע����Ϣ
        
        strValue = GetSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKey, strDefKeyValue)
        
    Case ˽��ģ��

        strValue = GetSetting("ZLSOFT", "˽��ģ��\" & strUser & "\" & App.ProductName & "\" & strSection, strKey, strDefKeyValue)
        
    Case ˽��ȫ��

        strValue = GetSetting("ZLSOFT", "˽��ȫ��\" & strUser & "\" & strSection, strKey, strDefKeyValue)
        
    Case ����ģ��

        strValue = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & strSection, strKey, strDefKeyValue)
        
    Case ����ȫ��
        
        strValue = GetSetting("ZLSOFT", "����ȫ��\" & strSection, strKey, strDefKeyValue)
        
    End Select
    
    GetRegister = strValue
    
errHand:
End Function

Public Function SetRegister(ByVal enmRegister As REGISTER, ByVal strSection As String, ByVal strKey As String, ByVal strKeyValue As String, Optional ByVal strUser As String) As Boolean
    '******************************************************************************************************************
    '���ܣ� ��ָ������Ϣ������ע�����
    '������ enmRegister-ע������
    '       strSection-ע���Ŀ¼
    '       strKey-����
    '       strKeyValue-��ֵ
    '���أ�
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Select Case enmRegister
    Case ע����Ϣ
        
        Call SaveSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue)
        
    Case ˽��ģ��

        Call SaveSetting("ZLSOFT", "˽��ģ��\" & strUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue)
        
    Case ˽��ȫ��

        Call SaveSetting("ZLSOFT", "˽��ȫ��\" & strUser & "\" & strSection, strKey, strKeyValue)
        
    Case ����ģ��

        Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & strSection, strKey, strKeyValue)
        
    Case ����ȫ��
        
        Call SaveSetting("ZLSOFT", "����ȫ��\" & strSection, strKey, strKeyValue)
        
    End Select
    
    SetRegister = True
    
errHand:
    
End Function

Private Sub cmdList_Click()
 
    Dim strNoPDF As String
    Dim rsTmp As ADODB.Recordset
    
    If Text(3).Text = "" Then
        MsgBox "�û�����Ϊ�գ�"
        Exit Sub
    End If
    If Text(4).Text = "" Then
        MsgBox "���벻��Ϊ�գ�"
        Exit Sub
    End If
    If Text(2).Text = "" Then
        MsgBox "������Ϊ�գ�"
        Exit Sub
    End If
    
    If Text(0).Text = "" Then
        MsgBox "����ID����Ϊ�գ�"
        Exit Sub
    End If
    
    If Text(1).Text = "" Then
        MsgBox "��ҳID����Ϊ�գ�"
        Exit Sub
    End If
    
    If txtPath.Text = "" Then
        MsgBox "PDF���λ�ò���Ϊ�գ�"
        Exit Sub
    End If
    On Error GoTo errH
    Set mobjPrint = Nothing
    Set mobjPrint = CreateObject("ZLPublicCISPrint.clsPrint")
    If Not mobjPrint.InitPrint(Trim(Text(2).Text), Trim(Text(3).Text), Trim(Text(4).Text), Trim(Text(5).Text)) Then Exit Sub
    txtList.Text = mobjPrint.GetPrintList(Val(Text(0).Text), Val(Text(1).Text))
    
    Exit Sub
errH:
    
End Sub

Private Sub Form_Load()
          Dim strPath As String
          Dim objFSO As New FileSystemObject
   On Error GoTo Form_Load_Error

10        strPath = GetRegister(˽��ģ��, "��ӡ����", "PDFλ��", App.Path, Trim(Text(3).Text))
20        If Not objFSO.FolderExists(strPath) And strPath <> "" Then
30            Call objFSO.CreateFolder(strPath)
40        End If
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form frmPDF"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SetRegister(˽��ģ��, "��ӡ����", "PDFλ��", Trim(txtPath.Text), Trim(Text(3).Text))
End Sub

Private Sub txtXML_DblClick()
    txtXML.Text = "<items><item><id></id><file_path></file_path></item></items>"
End Sub
