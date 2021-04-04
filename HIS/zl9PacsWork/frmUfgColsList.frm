VERSION 5.00
Begin VB.Form frmUfgColsList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����б�����"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUfgColsList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&S)"
      Height          =   405
      Left            =   1440
      TabIndex        =   4
      Top             =   3960
      Width           =   825
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�(&E)"
      Height          =   405
      Left            =   2400
      TabIndex        =   3
      Top             =   3960
      Width           =   820
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "�ָ�Ĭ��(&D)"
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   1185
   End
   Begin VB.ListBox lstUfgColsName 
      Height          =   3435
      ItemData        =   "frmUfgColsList.frx":6852
      Left            =   120
      List            =   "frmUfgColsList.frx":6854
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ѡ��Ҫ��ʾ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   1680
   End
End
Attribute VB_Name = "frmUfgColsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mUcData As ucFlexGrid
Private mStrDefaultColNames As String



Public Sub ShowUfgColsListWindow(ByRef UcData As ucFlexGrid, ByVal strDefaultColNames As String)
'�������б��岢����Ĭ����ʾ����
    
    '��UcData�������ģ�鼶����
    Set mUcData = UcData
    mStrDefaultColNames = strDefaultColNames
    
    '���ü���Ĭ����ʾ����
    Call LoadColsList(UcData)
    
    '���ش���
    Call Show(1)
End Sub



Private Sub cmdOK_Click()
'ȷ���������Բ���
On Error GoTo ErrHandle

    Dim i As Integer
    Dim strColsName As String
    Dim strColName  As String
    Dim strProperty As String
    Dim objProperty As Scripting.Dictionary
    
    '��ť�Ƿ񱻵��
    cmdOK.Tag = True
    
    If cmdDefault.Tag = "True" Then mUcData.ColNames = mStrDefaultColNames
    
    For i = 0 To lstUfgColsName.ListCount - 1
        '�ж��Ƿ�ѡ��ĳ��
        If lstUfgColsName.Selected(i) Then
            strColsName = strColsName + lstUfgColsName.list(i) & ","
        End If
    Next

    '��hide����д��flexcpdata
    For i = 1 To mUcData.DataGrid.Cols - 1
        strColName = mUcData.DataGrid.Cell(flexcpText, 0, i)
        
        Set objProperty = mUcData.DataGrid.Cell(flexcpData, 0, i)
        
        If Not objProperty Is Nothing Then
            strProperty = Mid(objProperty(TColPro.cpProperty), InStrRev(objProperty(TColPro.cpProperty), "@") + 1)
            
            '�ж�ƥ��ѡ���� ƥ����ɾ��hide����  δƥ������hide����
            If InStr(strColsName, strColName) = 0 Then
            
                '��������ַ����д��� uncfg ���Ծ�����
                If InStr(strProperty, "uncfg") = 0 Then
                    '��������ַ��������� hide ���Ծ�����  û����׷��
                    If InStr(strProperty, "hide") Then
                        
                        objProperty(TColPro.cpProperty) = mUcData.GetFieldName(i) & "@" & strProperty
                        
                    Else
                        '����������Ժ� �����е������ַ���
                        objProperty(TColPro.cpProperty) = mUcData.GetFieldName(i) & "@" & strProperty & ",hide"
                    End If
                    
                    '������
                     mUcData.DataGrid.ColHidden(i) = True
                     mUcData.DataGrid.Cell(flexcpData, 0, i)(TColPro.cpIsHide) = True
                     '���ü���CheckBoxλ�÷���
                     mUcData.RefreshCbxPostion
                End If
                
            Else
            
                 '��������ַ����д��� uncfg ���Ծ�����
                If InStr(strProperty, "uncfg") = 0 Then
                    '��������ַ��������� hide ���Ծ�ȥ��  û�������
                    If InStr(strProperty, "hide") Then
                        
                        '��hide����ɾ�� �������е������ַ���
                        strProperty = Replace(strProperty, ",hide", "")
                        objProperty(TColPro.cpProperty) = mUcData.GetFieldName(i) & "@" & strProperty
                    Else
                    
                        objProperty(TColPro.cpProperty) = mUcData.GetFieldName(i) & "@" & strProperty
                        
                    End If
                    
                     '��ʾ��
                     mUcData.DataGrid.ColHidden(i) = False
                     mUcData.DataGrid.Cell(flexcpData, 0, i)(TColPro.cpIsHide) = False
                     '���ü���CheckBoxλ�÷���
                     mUcData.RefreshCbxPostion
                End If
        
            End If
        End If
    Next
    
    Call Me.Hide
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub LoadColsList(UcData As ucFlexGrid)
'����Ĭ����ʾ����
    Dim i As Integer
    Dim j As Integer
    Dim strProperty As String
    Dim objProperty As Scripting.Dictionary
    
    '���ؼ��±긳��ʼֵ
    j = 0
    '���List�ؼ�
    lstUfgColsName.Clear
    
    For i = 1 To UcData.DataGrid.Cols - 1
        Set objProperty = mUcData.DataGrid.Cell(flexcpData, 0, i)
        
        If Not objProperty Is Nothing Then
            strProperty = Mid(objProperty(TColPro.cpProperty), InStrRev(objProperty(TColPro.cpProperty), "@") + 1)
            
            If InStr(strProperty, "uncfg") = 0 Then
                '�����ַ�
                frmUfgColsList.lstUfgColsName.list(j) = UcData.DataGrid.Cell(flexcpText, 0, i)
                
                '�ж��Ƿ�Ĭ��Ϊ������
                If InStr(strProperty, "hide") = 0 Then
                    '����Ĭ����ʾ��
                    frmUfgColsList.lstUfgColsName.Selected(j) = True
                End If
                
                j = j + 1
            End If
        End If
    Next

End Sub


Private Sub cmdDefault_Click()
'�ָ�Ĭ�Ϲ�ѡ
On Error GoTo ErrHandle
    Dim i As Integer
    Dim j As Integer
    Dim strProperty As String
    Dim strTemp As String
    Dim strColNames() As String
    
    '��ť�Ƿ񱻵��
    cmdDefault.Tag = True
    
    '���������ô����벢��������
    strColNames() = Split(mStrDefaultColNames, "|")
    
     '���List�ؼ�
    lstUfgColsName.Clear
    
    For i = 1 To UBound(strColNames()) - 1
        strProperty = strColNames(i)
        
        If InStr(strProperty, "uncfg") = 0 Then
            
            '�ж��ַ����Ƿ���� ��>�� ��,�����ţ���Ҫ���н�ȡ����
             If InStr(strProperty, ">") > 0 Then
                strTemp = Mid(strProperty, 1, InStr(strProperty, ">") - 1)

                If InStr(strTemp, ",") > 0 Then
                    strTemp = Mid(strTemp, 1, InStr(strTemp, ",") - 1)
                Else
                    strTemp = strTemp
                End If
             Else
                If InStr(strProperty, ",") > 0 Then
                    strTemp = Mid(strProperty, 1, InStr(strProperty, ",") - 1)
                Else
                    strTemp = strProperty
                End If
             End If

            '�����ַ�
            frmUfgColsList.lstUfgColsName.list(j) = strTemp
            
            '�ж��Ƿ�Ĭ��Ϊ������
            If InStr(strProperty, "hide") = 0 Then
                '����Ĭ����ʾ��
                frmUfgColsList.lstUfgColsName.Selected(j) = True

            End If
            
            j = j + 1
        End If
    Next

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdExit_Click()
'ж�ش���
On Error GoTo ErrHandle

    Unload Me
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
'�������ö�
    '���Ĭ�����������ڿգ�����ûָ�Ĭ�ϰ�ť
    If Trim(mStrDefaultColNames) = "" Then cmdDefault.Enabled = False
    '�������ö�
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3

End Sub
