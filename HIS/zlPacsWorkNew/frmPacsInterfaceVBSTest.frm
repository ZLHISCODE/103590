VERSION 5.00
Begin VB.Form frmPacsInterfaceVBSTest 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "������֤"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txt 
      Height          =   270
      Index           =   0
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&S)"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "����ֵ"
      Height          =   180
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   540
   End
   Begin VB.Label lab 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   540
   End
End
Attribute VB_Name = "frmPacsInterfaceVBSTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintParsCount As Integer '���ظ��Ĳ�������
Private mstrVBSOld As String
Private mstrVBSTest As String
Private mobjOwner As frmPacsInterfaceCfg
Private mintVBSTest As Integer
Public Function zlShowMe(ByVal strVBS As String, objOwner As frmPacsInterfaceCfg) As Integer
    Set mobjOwner = objOwner
    mstrVBSOld = strVBS
    
    Call LoadControlAndLayout
    
    Call Me.Show(1, objOwner)
    zlShowMe = mintVBSTest
End Function
Private Sub LoadControlAndLayout()
'����mstrVBSOld�е�Ԥ�����������̬���ɿؼ����Ҳ���
    Dim strTmp As String
    Dim intTMP As Integer
    Dim lngL As Long, lngT As Long, lngW As Long, lngH As Long
    Dim strParName As String
    Dim blHaveSamePar As Boolean '�Ƿ��Ѿ�����һ���Ĳ���
    Dim i As Integer
    
    mintParsCount = 0
    strTmp = mstrVBSOld
    While InStr(strTmp, "[[")
        
        '���ܣ���ȡ��������
        strTmp = Mid(strTmp, InStr(strTmp, "[[") + 2)
        strParName = Mid(strTmp, 1, InStr(strTmp, "]]") - 1)
        blHaveSamePar = False
        
        For i = 0 To mintParsCount - 1
            If lab(i).Caption = strParName Then
                blHaveSamePar = True
                Exit For
            End If
        Next
        
        If mintParsCount > 0 Then
            '�����ų��ظ����
            If Not blHaveSamePar Then
                mintParsCount = mintParsCount + 1
                
                Load lab(mintParsCount - 1)
                lab(mintParsCount - 1).Caption = strParName
                lab(mintParsCount - 1).AutoSize = True
                Call lab(mintParsCount - 1).Move(240, 240 + 360 * mintParsCount)
                lab(mintParsCount - 1).Visible = True
        
                Load txt(mintParsCount - 1)
                txt(mintParsCount - 1).Text = ""
                txt(mintParsCount - 1).tag = strParName
                Call txt(mintParsCount - 1).Move(1560, 240 + 360 * mintParsCount, 1215, 270)
                txt(mintParsCount - 1).Visible = True
                
            Call SetDefaltValue(strParName, False)
            End If
        Else
            lab(0).Caption = strParName
            txt(0).tag = strParName

            Call SetDefaltValue(strParName, True)
            mintParsCount = mintParsCount + 1
        End If
        
        strTmp = Mid(strTmp, InStr(strTmp, "]]") + 2)
    Wend
    
    Call cmdOK.Move(1080, lab(mintParsCount - 1).Top + lab(mintParsCount - 1).Height + 360)
    Call cmdCancel.Move(2160, lab(mintParsCount - 1).Top + lab(mintParsCount - 1).Height + 360)
    
    lngW = 3315
    lngL = mobjOwner.Left + (mobjOwner.Width - lngW) / 2
    lngH = cmdOK.Top + cmdOK.Height + 600
    lngT = mobjOwner.Top + (mobjOwner.Height - lngH) / 2
    
    Call Me.Move(lngL, lngT, lngW, lngH)
    
End Sub

Private Sub ReplacePars(ByVal strParName As String)
'�滻��������
On Error GoTo ErrorHnad
    Dim strValue As String
    Dim i As Integer

    If strParName = "��ǰ���ھ��" Then
        mstrVBSTest = Replace(mstrVBSTest, "[[" & strParName & "]]", mobjOwner.hWnd)
    Else
        For i = 0 To mintParsCount - 1
            If txt(i).tag = strParName Then
                strValue = txt(i).Text
                Exit For
            End If
        Next
        mstrVBSTest = Replace(mstrVBSTest, "[[" & strParName & "]]", strValue)
    End If
    
    Exit Sub
ErrorHnad:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub cmdCancel_Click()
    mintVBSTest = δ����
    Unload Me
End Sub

Private Sub CmdOK_Click()
'�����滻�붨�������Ȼ�����
On Error GoTo ErrorHnad
    mstrVBSTest = mstrVBSOld
    Call ReplacePars("�û���")
    Call ReplacePars("�˺���")
    Call ReplacePars("ϵͳ��")
    Call ReplacePars("ģ���")
    Call ReplacePars("����ID")
    Call ReplacePars("����ID")
    Call ReplacePars("ҽ��ID")
    Call ReplacePars("����")
    Call ReplacePars("�����")
    Call ReplacePars("סԺ��")
    Call ReplacePars("���֤��")
    Call ReplacePars("Ӱ�����")
    Call ReplacePars("��ǰ���ھ��")
    
    Call TestExecuteSub(mstrVBSTest)
    Exit Sub
ErrorHnad:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub


Private Function TestExecuteSub(ByVal strVBS As String) As Boolean
'����vbs�ű�ʵ�ֹ���
On Error GoTo ErrorHnad
    Dim objCall As Object
    Dim strTempVBS As String
    
    '�����ű�ִ�ж���
    Set objCall = CreateObject("ScriptControl")
    objCall.Timeout = 60000
    objCall.Language = "vbscript"
    
    Call objCall.AddCode(strVBS)
    
    Call objCall.Run(Trim("ExcuteSub"))
    
    TestExecuteSub = True
    mintVBSTest = ͨ��
    Unload Me
    Exit Function
ErrorHnad:
    If err.Description <> "" Then MsgBox err.Description, vbExclamation, gstrSysName
End Function

Private Sub Form_Terminate()
    Set mobjOwner = Nothing
End Sub

Private Sub SetDefaltValue(ByVal strName As String, ByVal blnFirst As Boolean)
'��������ʼֵ�Ĵ���
    Dim lngIndex As Long '(�ؼ�����)
    
    If blnFirst Then
        lngIndex = 0
    Else
        lngIndex = mintParsCount - 1
    End If
    
    Select Case strName
        Case "�û���"
            txt(lngIndex).Text = UserInfo.����
             
        Case "�˺���"
            txt(lngIndex).Text = UserInfo.�û���
                                
        Case "ϵͳ��"
            txt(lngIndex).Text = "100"
            
        Case "ģ���"
            txt(lngIndex).Text = "1291"
            
        Case "����ID"
            txt(lngIndex).Text = UserInfo.����ID
        
        Case "����ID"
            txt(lngIndex).Text = "1"
            
        Case "ҽ��ID"
            txt(lngIndex).Text = "101"
            
        Case "����"
            txt(lngIndex).Text = "110"
            
        Case "�����"
            txt(lngIndex).Text = "1"
        
        Case "סԺ��"
            txt(lngIndex).Text = "110"
            
        Case "���֤��"
            txt(lngIndex).Text = "500105190001010000"
            
        Case "Ӱ�����"
            txt(lngIndex).Text = "CT"
                                
        Case "��ǰ���ھ��"
            txt(lngIndex).Text = Me.hWnd
            txt(lngIndex).Enabled = False
    End Select

End Sub
