VERSION 5.00
Begin VB.Form frmResult 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10125
   ControlBox      =   0   'False
   Icon            =   "frmResult.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraCriticalValues 
      Caption         =   "Σ��ֵ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   1890
      Begin VB.OptionButton optCriticalValues 
         Caption         =   "��ͨ"
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
         Index           =   1
         Left            =   480
         TabIndex        =   20
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton optCriticalValues 
         Caption         =   "Σ��"
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
         Index           =   2
         Left            =   480
         TabIndex        =   19
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.Frame fraReportLevel 
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   6120
      TabIndex        =   11
      Top             =   120
      Width           =   1890
      Begin VB.OptionButton optReportLevel 
         Caption         =   "��"
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
         Index           =   4
         Left            =   360
         TabIndex        =   15
         Top             =   1440
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optReportLevel 
         Caption         =   "��"
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
         Index           =   3
         Left            =   360
         TabIndex        =   14
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optReportLevel 
         Caption         =   "��"
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
         Index           =   1
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optReportLevel 
         Caption         =   "��"
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
         Index           =   2
         Left            =   360
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame fraFuHeLevel 
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   8115
      TabIndex        =   7
      Top             =   120
      Width           =   1890
      Begin VB.OptionButton optFuHeLevel 
         Caption         =   "�� �� ��"
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
         Index           =   3
         Left            =   360
         TabIndex        =   10
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton optFuHeLevel 
         Caption         =   "��    ��"
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
         Index           =   1
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optFuHeLevel 
         Caption         =   "��������"
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
         Index           =   2
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.Frame fraImageLevel 
      Caption         =   "Ӱ������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   4140
      TabIndex        =   2
      Top             =   120
      Width           =   1890
      Begin VB.OptionButton optImageLevel 
         Caption         =   "��"
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
         Index           =   4
         Left            =   360
         TabIndex        =   17
         Top             =   1500
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optImageLevel 
         Caption         =   "��"
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
         Index           =   3
         Left            =   360
         TabIndex        =   16
         Top             =   1160
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optImageLevel 
         Caption         =   "��"
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
         Index           =   2
         Left            =   360
         TabIndex        =   6
         Top             =   780
         Width           =   1455
      End
      Begin VB.OptionButton optImageLevel 
         Caption         =   "��"
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
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   420
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame fraResult 
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   2145
      TabIndex        =   1
      Top             =   120
      Width           =   1890
      Begin VB.OptionButton optResult 
         Caption         =   "����"
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
         Index           =   2
         Left            =   480
         TabIndex        =   4
         Top             =   1320
         Width           =   855
      End
      Begin VB.OptionButton optResult 
         Caption         =   "����"
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
         Index           =   1
         Left            =   480
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   350
      Left            =   6720
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmResult.frx":000C
      TabIndex        =   0
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   1100
   End
End
Attribute VB_Name = "frmResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintResult As Integer    '�����
Private mintImageLevel As Integer '��Ƭ����
Private mintFuHeLevel As Integer  '�������
Private mintReportLevel As Integer '��������
Private mintCriticalValues As Integer 'Σ�����
Private mstrResult As String
Public mlngModul As Long      'ģ��ŵ���

Public Function zlGetResult(frmParent As Form, ByVal lngModul As Long, ByVal strQueryId As String, lngCur����ID As Long, strResultInput As String) As String
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim blnShowResult As Boolean
    Dim blnShowCriticalValues As Boolean
    Dim blnShowImageLevel As Boolean
    Dim blnShowReportLevel As Boolean
    Dim blnShowFuHeLevel As Boolean
    Dim strImageLevel As String
    Dim strReportLevel As String
    Dim intTxtLen As Integer
    Dim i As Integer
    Dim lngFramCount As Long
    
    zlGetResult = ""
    mlngModul = lngModul
    
    If strQueryId Like String(Len(strQueryId), "#") Then
        strSql = "Select a.Ӱ������,a.�������,a.��������,b.�������,a.Σ��״̬ From Ӱ�����¼ a,����ҽ������ b  " _
                & " Where a.ҽ��ID= b.ҽ��ID And a.���ͺ� =b.���ͺ� And  a.ҽ��ID=[1] "
    Else
        strSql = "Select B.Σ��״̬, A.�������, B.Ӱ������, A.��������, B.�������,B.ҽ��ID " & _
                 "From Ӱ�񱨸��¼ A, Ӱ�����¼ B " & _
                 "Where A.ID=[1] and A.ҽ��id = B.ҽ��id"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ������", strQueryId)
    
    '�����Ѿ�ѡ������Խ��
    If Nvl(rsTemp!�������) = 1 Then
        optResult(1).value = True   '���԰�ť
        mintResult = 1
    Else
        optResult(2).value = True   '���԰�ť
        mintResult = 2
    End If
    
    
    '�����¼Ϊ����ʹ��Ĭ��ֵ
    If Nvl(rsTemp!�������) = "" Then
         '�����ѡ���� ��Ͻ��Ĭ������ �Զ�ѡ�����԰�ť ��֮ѡ�����԰�ť
        If Val(GetDeptPara(lngCur����ID, "��Ͻ��Ĭ������", 0)) = 1 Then
            optResult(1).value = True
            mintResult = 1
        Else
            optResult(2).value = True
            mintResult = 2
        End If
    End If
    
    lngFramCount = 1
    
    blnShowResult = True
    blnShowCriticalValues = True
    blnShowImageLevel = True
    blnShowReportLevel = True
    blnShowFuHeLevel = True
    
    If InStr(strResultInput, "Σ��״̬") > 0 Then
        blnShowCriticalValues = True
        
        If Nvl(rsTemp!Σ��״̬) = 1 Then
            optCriticalValues(2).value = True   'Σ����ť
            mintCriticalValues = 2
        Else
            optCriticalValues(1).value = True   '������ť
            mintCriticalValues = 1
        End If
        
        If mintCriticalValues = 2 Then
            '���ΪΣ��״̬Ϊ'Σ��',����Ϊ'����'���Ҳ��ɸ���
            optResult(1).value = True
            optResult(1).Enabled = False
            optResult(2).Enabled = False
        End If
        
        lngFramCount = lngFramCount + 1
    Else
        blnShowCriticalValues = False
        fraCriticalValues.Visible = False
    End If
    
    
    
    If InStr(strResultInput, "�������") > 0 Then
        blnShowResult = True
        lngFramCount = lngFramCount + 1
    Else
        blnShowResult = False
        fraResult.Visible = False
    End If
    
    
    
    'Ӱ����������
    strImageLevel = Nvl(GetDeptPara(lngCur����ID, "Ӱ�������ȼ�", "��,��"))
    intTxtLen = Len(strImageLevel) - Len(Replace(strImageLevel, ",", "")) + 1
    
    If mlngModul = 1290 Then
        If InStr(strResultInput, "Ӱ������") <= 0 Then        '����ʾ
            fraImageLevel.Visible = False
            blnShowImageLevel = False
            
            If IsNull(rsTemp!Ӱ������) Then mintImageLevel = 0
        Else
            lngFramCount = lngFramCount + 1
            mintImageLevel = 1
        End If
    Else
        fraImageLevel.Visible = False
        blnShowImageLevel = False
        If IsNull(rsTemp!Ӱ������) Then mintImageLevel = 0
    End If
    
    '�̶������4��Ӱ��ȼ�  ����ѭ��4��
    For i = 1 To 4
        If i <= intTxtLen Then
            optImageLevel(i).Visible = True
            
            If Trim(Split(strImageLevel, ",")(i - 1)) <> "" Then
                optImageLevel(i).Caption = Trim(Split(strImageLevel, ",")(i - 1))
            Else
                optImageLevel(i).Caption = "δ����"
            End If
            
            If Nvl(rsTemp!Ӱ������) = i Then
                optImageLevel(i).value = True
                mintImageLevel = i
            End If

        Else
            optImageLevel(i).Visible = False
        End If
    Next i
    
    'ͨ�����õĵȼ��������ж�top ��ֵ
    Select Case intTxtLen
        Case 2
            optImageLevel(1).Top = 600
            optImageLevel(2).Top = 1320
        Case 3
            optImageLevel(1).Top = 480
            optImageLevel(2).Top = 960
            optImageLevel(3).Top = 1440
        Case 4
            optImageLevel(1).Top = 420
            optImageLevel(2).Top = 780
            optImageLevel(3).Top = 1160
            optImageLevel(4).Top = 1500
    End Select
    
    
    
     '������������
    strReportLevel = Nvl(GetDeptPara(lngCur����ID, "���������ȼ�", "��,��"))
    intTxtLen = Len(strReportLevel) - Len(Replace(strReportLevel, ",", "")) + 1
    
    If InStr(strResultInput, "��������") <= 0 Then       '����ʾ
        fraReportLevel.Visible = False
        blnShowReportLevel = False
        
        If IsNull(rsTemp!��������) Then mintReportLevel = 0
    Else
        lngFramCount = lngFramCount + 1
        mintReportLevel = 1
    End If
    
    '�̶������4���ȼ�  ����ѭ��4��
    For i = 1 To 4
        If i <= intTxtLen Then
            optReportLevel(i).Visible = True
            
            If Trim(Split(strReportLevel, ",")(i - 1)) <> "" Then
                optReportLevel(i).Caption = Trim(Split(strReportLevel, ",")(i - 1))
            Else
                optReportLevel(i).Caption = "δ����"
            End If
            
            If Nvl(rsTemp!��������) = i Then
                optReportLevel(i).value = True
                mintReportLevel = i
            End If
        Else
            optReportLevel(i).Visible = False
        End If
    Next i

    'ͨ�����õĵȼ��������ж�top ��ֵ
    Select Case intTxtLen
        Case 2
            optReportLevel(1).Top = 600
            optReportLevel(2).Top = 1320
        Case 3
            optReportLevel(1).Top = 480
            optReportLevel(2).Top = 960
            optReportLevel(3).Top = 1440
        Case 4
            optReportLevel(1).Top = 420
            optReportLevel(2).Top = 780
            optReportLevel(3).Top = 1160
            optReportLevel(4).Top = 1500
    End Select
    
    '��ʾ�������
    If Nvl(rsTemp!�������) = "" Or Nvl(rsTemp!�������) = "����" Then
        optFuHeLevel(1).value = True
        mintFuHeLevel = 1
    ElseIf Nvl(rsTemp!�������) = "��������" Then
        optFuHeLevel(2).value = True
        mintFuHeLevel = 2
    Else
        optImageLevel(3).value = True
        mintFuHeLevel = 3
    End If

    If mlngModul = 1294 Or InStr(strResultInput, "�������") <= 0 Then        '����ʾ
        fraFuHeLevel.Visible = False
        blnShowFuHeLevel = False
        mintFuHeLevel = 1
    Else
        lngFramCount = lngFramCount + 1
    End If
    
    Me.Width = IIf(blnShowResult, fraResult.Width, 0) + IIf(blnShowCriticalValues, fraCriticalValues.Width, 0) + IIf(blnShowImageLevel, fraImageLevel.Width, 0) + IIf(blnShowReportLevel, fraReportLevel.Width, 0) + IIf(blnShowFuHeLevel, fraFuHeLevel.Width, 0) + 120 + lngFramCount * 120
    
    cmdOK.Left = Me.Width - cmdOK.Width - 240
    
    If Not blnShowCriticalValues Then
        fraResult.Left = fraCriticalValues.Left
    Else
        fraResult.Left = fraCriticalValues.Left + fraCriticalValues.Width + 120
    End If
    
    If Not blnShowResult Then
        fraImageLevel.Left = fraResult.Left
    Else
        fraImageLevel.Left = fraResult.Left + fraResult.Width + 120
    End If

    If Not blnShowImageLevel Then
        fraReportLevel.Left = fraImageLevel.Left
    Else
        fraReportLevel.Left = fraImageLevel.Left + fraImageLevel.Width + 120
    End If

    If Not blnShowReportLevel Then
        fraFuHeLevel.Left = fraReportLevel.Left
    Else
        fraFuHeLevel.Left = fraReportLevel.Left + fraReportLevel.Width + 120
    End If
    
    If blnShowResult = False And blnShowCriticalValues = False And blnShowImageLevel = False And blnShowReportLevel = False And blnShowFuHeLevel = False Then
        Unload Me
        Exit Function
    End If

    Me.Show 1, frmParent
    zlGetResult = mstrResult
End Function

Private Sub cmdCancel_Click()
    mstrResult = ""
    Unload Me
End Sub

Private Sub CmdOK_Click()
    mstrResult = mintCriticalValues & "-" & mintResult & "-" & mintImageLevel & "-" & mintReportLevel & "-" & mintFuHeLevel
    Unload Me
End Sub

Private Sub Form_Load()
    '�����ö�
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3 '�������ö�
End Sub

Private Sub optCriticalValues_Click(Index As Integer)
    mintCriticalValues = Index
    If mintCriticalValues = 1 Then
        optResult(1).Enabled = True
        optResult(2).Enabled = True
    Else
        '���ΪΣ��״̬Ϊ'Σ��',����Ϊ'����'���Ҳ��ɸ���
        optResult(1).value = True
        optResult(1).Enabled = False
        optResult(2).Enabled = False
    End If
End Sub

Private Sub optFuHeLevel_Click(Index As Integer)
    mintFuHeLevel = Index
End Sub

Private Sub optImageLevel_Click(Index As Integer)
    mintImageLevel = Index
End Sub

Private Sub optReportLevel_Click(Index As Integer)
    mintReportLevel = Index
End Sub

Private Sub optResult_Click(Index As Integer)
    mintResult = Index
End Sub


