VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPatholRequisition_AntibodyInf 
   Caption         =   "��ϸ��Ϣ"
   ClientHeight    =   6825
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   7155
   Icon            =   "frmPatholRequisition_AntibodyInf.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   7155
   StartUpPosition =   3  '����ȱʡ
   Begin RichTextLib.RichTextBox txtContext 
      Height          =   6015
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   10610
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmPatholRequisition_AntibodyInf.frx":179A
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�� ��(&E)"
      Height          =   400
      Left            =   5760
      TabIndex        =   0
      Top             =   6240
      Width           =   1215
   End
End
Attribute VB_Name = "frmPatholRequisition_AntibodyInf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Public Sub ShowAntibodyInf(ByVal lngAntibodyId As Long, owner As Form)
'��ʾ������Ϣ
    Call LoadAntibodyInf(lngAntibodyId)
    
    Call Me.Show(1, owner)
End Sub


Private Sub LoadAntibodyInf(ByVal lngAntibodyId As Long)
'��ȡ������Ϣ
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim strFormats As String
    Dim strTemp As String
    
    
    strFormats = "{\rtf1\ansi\ansicpg936\deff0\deflang1033\deflangfe2052{\fonttbl{\f0\fnil\fcharset134 \'cb\'ce\'cc\'e5;}}" & _
                "{\colortbl ;\red0\green0\blue160;\red255\green0\blue0;\red0\green77\blue187;} " & _
                "{\*\generator Msftedit 5.41.21.2509;}\viewkind4\uc1\pard\sa200\sl276\slmult1\lang2052\b\f0\fs22\par " & _
                " \'bf\'b9\'cc\'e5\'c3\'fb\'b3\'c6\'a3\'ba\cf1\b0  [Value1]\par" & _
                "\cf0 --------------------------------------------------\par" & _
                "\b  \'ca\'b9\'d3\'c3\'c8\'cb\'b7\'dd\'a3\'ba\cf2\b0  [Value2]\cf0\par" & _
                "\b  \'d2\'d1\'d3\'c3\'c8\'cb\'b7\'dd\'a3\'ba\cf2\b0  [Value3]\par" & _
                "\cf0 --------------------------------------------------\par" & _
                "\b  \'c9\'fa\'b2\'fa\'c8\'d5\'c6\'da\'a3\'ba\cf1\b0  [Value4]\par" & _
                "\cf0\b  \'b9\'fd\'c6\'da\'c8\'d5\'c6\'da\'a3\'ba\cf1\b0  [Value5]\par" & _
                "\cf0\b  \'d3\'d0 \'d0\'a7 \'c6\'da\'a3\'ba\cf1\b0  [Value6]\par" & _
                "\cf0 --------------------------------------------------\par" & _
                "\b  \'bf\'cb \'c2\'a1 \'d0\'d4\'a3\'ba\cf1\b0  [Value7]\par" & _
                "\cf0\b  \'d7\'f7\'d3\'c3\'b6\'d4\'cf\'f3\'a3\'ba\cf1\b0  [Value8]\par" & _
                "\cf0\b  \'c0\'ed\'bb\'af\'d0\'d4\'d6\'ca\'a3\'ba\cf1\b0  [Value9]\par" & _
                "\cf0\b  \'d3\'a6\'d3\'c3\'c7\'e9\'bf\'f6\'a3\'ba\cf1\b0  [Value10]\par" & _
                "\cf0--------------------------------------------------\par" & _
                "\b  \cf0\'b5\'c7 \'bc\'c7 \'c8\'cb\'a3\'ba\cf1\b0  [Value11]\par" & _
                "\cf0\b  \'b5\'c7\'bc\'c7\'ca\'b1\'bc\'e4\'a3\'ba\cf1\b0  [Value12]\par" & _
                "\cf0--------------------------------------------------\par" & _
                "\b  \'b1\'b8    \'d7\'a2\'a3\'ba\b0\par" & _
                "\cf1  [Value13]\par" & _
                "\cf0 --------------------------------------------------\par" & _
                "\b  \'b7\'b4\'c0\'a1\'bc\'c7\'c2\'bc\'a3\'ba\b0\par\cf1 [Value14]\par" & _
                "\cf0 --------------------------------------------------\cf3\par}"

    
    strSql = "select ��������,ʹ���˷�,�����˷�,��������,��������,��Ч��,��¡��,���ö���,������,Ӧ�����,�Ǽ���,�Ǽ�ʱ��,��ע " & _
            " from ��������Ϣ where ����ID=[1]"
            
            
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAntibodyId)
    
    txtContext.Text = "����ϸ��Ϣ......"
    If rsData.RecordCount <= 0 Then Exit Sub
    
    txtContext.Text = ""
    
    'txtContext.Text = " �������ƣ�" & Nvl(rsData!��������) & vbCrLf & vbCrLf
    strFormats = Replace(strFormats, "[Value1]", Nvl(rsData!��������))
    
    'txtContext.Text = txtContext.Text & " --------------------------------------------------" & vbCrLf
    'txtContext.Text = txtContext.Text & " ʹ���˷ݣ�" & Nvl(rsData!ʹ���˷�) & vbCrLf
    'txtContext.Text = txtContext.Text & " �����˷ݣ�" & Nvl(rsData!�����˷�) & vbCrLf & vbCrLf
    
    strFormats = Replace(strFormats, "[Value2]", Nvl(rsData!ʹ���˷�))
    strFormats = Replace(strFormats, "[Value3]", Nvl(rsData!�����˷�))
    
    'txtContext.Text = txtContext.Text & " --------------------------------------------------" & vbCrLf
    'txtContext.Text = txtContext.Text & " �������ڣ�" & Format((rsData!��������), gstrDateFormat) & vbCrLf
    'txtContext.Text = txtContext.Text & " �������ڣ�" & Format((rsData!��������), gstrDateFormat) & vbCrLf
    'txtContext.Text = txtContext.Text & " �� Ч �ڣ�" & Nvl(rsData!��Ч��) & "��" & vbCrLf & vbCrLf
    strFormats = Replace(strFormats, "[Value4]", Nvl(rsData!��������))
    strFormats = Replace(strFormats, "[Value5]", Nvl(rsData!��������))
    strFormats = Replace(strFormats, "[Value6]", Nvl(rsData!��Ч��) & "��")
    
    'txtContext.Text = txtContext.Text & " --------------------------------------------------" & vbCrLf
    'txtContext.Text = txtContext.Text & " �� ¡ �ԣ�" & IIf(Val(Nvl(rsData!��¡��)) = 0, "����¡", "���¡") & vbCrLf
    'txtContext.Text = txtContext.Text & " ���ö���" & Nvl(rsData!���ö���) & vbCrLf
    'txtContext.Text = txtContext.Text & " �����ʣ�" & Nvl(rsData!������) & vbCrLf
    'txtContext.Text = txtContext.Text & " Ӧ�������" & vbCrLf & "    " & Nvl(rsData!Ӧ�����) & vbCrLf & vbCrLf
    strFormats = Replace(strFormats, "[Value7]", Decode(Val(Nvl(rsData!��¡��)), 0, "����¡(Ũ����)", 1, "����¡(������)", 2, "���¡(Ũ����)", 3, "���¡(������)"))
    strFormats = Replace(strFormats, "[Value8]", Nvl(rsData!���ö���))
    strFormats = Replace(strFormats, "[Value9]", Nvl(rsData!������))
    strFormats = Replace(strFormats, "[Value10]", Nvl(rsData!Ӧ�����))
    
    'txtContext.Text = txtContext.Text & " --------------------------------------------------" & vbCrLf
    'txtContext.Text = txtContext.Text & " �� �� �ˣ�" & Nvl(rsData!�Ǽ���) & vbCrLf
    'txtContext.Text = txtContext.Text & " �Ǽ�ʱ�䣺" & Format((rsData!�Ǽ�ʱ��), gstrFullDateTimeFormat) & vbCrLf & vbCrLf
    strFormats = Replace(strFormats, "[Value11]", Nvl(rsData!�Ǽ���))
    strFormats = Replace(strFormats, "[Value12]", Format((rsData!�Ǽ�ʱ��), gstrFullDateTimeFormat))
    
    'txtContext.Text = txtContext.Text & " --------------------------------------------------" & vbCrLf
    'txtContext.Text = txtContext.Text & " ��    ע��" & vbCrLf & "    " & Nvl(rsData!��ע) & vbCrLf & vbCrLf
    'txtContext.Text = txtContext.Text & " --------------------------------------------------" & vbCrLf & vbCrLf & vbCrLf
    strFormats = Replace(strFormats, "[Value13]", "    " & Nvl(rsData!��ע))
    
    

    '��ȡ���巴����¼
    strSql = "select decode(ʵ������,0,'�����黯',1,'����Ⱦɫ',3,'���Ӳ���','����') as ʵ������,�������,��������,����ҽ��,����ʱ�� from �����巴�� where ����ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAntibodyId)
    
    If rsData.RecordCount <= 0 Then
        strFormats = Replace(strFormats, "[Value14]", "")
        txtContext.SelRTF = strFormats
        Exit Sub
    End If
    
    
    'txtContext.Text = txtContext.Text & " ������¼��" & vbCrLf
    
    While Not rsData.EOF
        If strTemp <> "" Then strTemp = strTemp & "\par"
        
        strTemp = strTemp & "    [ʵ�����ͣ�" & rsData!ʵ������ & "] [�������ۣ�" & rsData!�������� & "] [����ҽ����" & rsData!����ҽ�� & "] [����ʱ�䣺" & rsData!����ʱ�� & "]" & "\par"
        strTemp = strTemp & "    ���������" & rsData!������� & "\par"
        rsData.MoveNext
    Wend
    
    'txtContext.Text = txtContext.Text & " --------------------------------------------------" & vbCrLf & vbCrLf
    strFormats = Replace(strFormats, "[Value14]", strTemp)
    txtContext.SelRTF = strFormats
    

End Sub


Private Sub AdjustFace()
    txtContext.Left = 120
    txtContext.Top = 120
    txtContext.Width = Me.Width - 360
    txtContext.Height = Me.Height - cmdExit.Height - 840
    
    cmdExit.Left = Me.Width - cmdExit.Width - 240
    cmdExit.Top = Me.Height - cmdExit.Height - 600
End Sub



Private Sub cmdExit_Click()
    Call Me.Hide
End Sub

Private Sub Form_Load()
On Error Resume Next
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Call AdjustFace
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
End Sub
