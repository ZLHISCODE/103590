VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSchPlanTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ԤԼ--ʱ��ƻ�����"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4740
   Icon            =   "frmSchPlanTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   4740
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2760
      TabIndex        =   5
      Top             =   2760
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   840
      TabIndex        =   4
      Top             =   2760
      Width           =   1100
   End
   Begin VB.TextBox txtCapacity 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "30"
      Top             =   2100
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker dpTimeStart 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "HH:mm:ss"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   885
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   46006274
      CurrentDate     =   .333333333333333
   End
   Begin VB.ComboBox cboSchExamTimeCalcType 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   322
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker dpTimeEnd 
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   1485
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   46006274
      CurrentDate     =   .5
   End
   Begin VB.Label Label4 
      Caption         =   "ԤԼ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "����ʱ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "��ʼʱ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "���ʱ�����㷽��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frmSchPlanTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mlngTimeProjectID As Long
Dim mlngPlanID As Long

Public Sub zlShowMe(frmParent As Form, lngTimeProjectID As Long, lngPlanID As Long)
'------------------------------------------------
'���ܣ�װ��ʱ���ı���ʽ�ͻ�������
'������ frmParent -- ������
'       lngTimeProjectID -- ʱ��ƻ�ID���������������=0
'       lngPlanID -- ԤԼ����ID
'���أ���
'------------------------------------------------
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    mlngTimeProjectID = lngTimeProjectID
    mlngPlanID = lngPlanID
    
    If mlngTimeProjectID <> 0 Then
        '�����ݿ��ȡ��ǰ��ʱ��ƻ�
        strSql = "select ��ʼʱ��,����ʱ��,ԤԼ����,���㷽��,ԤԼ����ID from Ӱ��ԤԼʱ��ƻ� where id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "Ӱ��ԤԼʱ������", mlngTimeProjectID)
        If rsTemp.EOF = False Then
            cboSchExamTimeCalcType.ListIndex = IIF(NVL(rsTemp!���㷽��, 1) = 1, 0, 1)
            dpTimeStart = Format(rsTemp!��ʼʱ��, "hh:mm:ss")
            dpTimeEnd = Format(rsTemp!����ʱ��, "hh:mm:ss")
            txtCapacity = NVL(rsTemp!ԤԼ����)
        End If
    End If
    
    Me.Show 1, frmParent
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    '����ʱ��ƻ�����
    If saveTimeProject = True Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    cboSchExamTimeCalcType.AddItem "���˴�ƽ��"
    cboSchExamTimeCalcType.AddItem "��Ŀʱ��"
    cboSchExamTimeCalcType.ListIndex = 0
End Sub

Private Function saveTimeProject() As Boolean
'------------------------------------------------
'���ܣ�����ʱ��ƻ�����
'������
'���أ�True -- �ɹ��� False -- ʧ��
'------------------------------------------------
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strS1 As String
    Dim strS2 As String
    Dim strE1 As String
    Dim strE2 As String
    
    On Error GoTo err
    '�ȼ���������ݵ���Ч��
    If dpTimeEnd.value <= dpTimeStart.value Then
        MsgBox "���������뿪ʼʱ��ͽ���ʱ�䣬��ʼʱ��Ӧ��С�ڽ���ʱ�䡣", vbOKOnly, "���ԤԼ��ʾ"
        dpTimeEnd.SetFocus
        Exit Function
    End If
    
    If Val(txtCapacity.Text) = 0 Then
        MsgBox "����������ԤԼ������", vbOKOnly, "���ԤԼ��ʾ"
        txtCapacity.SetFocus
        Exit Function
    End If
    
    '���ƽ�����ʱ��С��2���ӣ�������ʾ
    If DateDiff("n", dpTimeStart.value, dpTimeEnd.value) / Val(txtCapacity.Text) < 2 Then
        MsgBox "����ԤԼ�����Ƿ�����������ʱ�����ƽ�����ʱ��С��2���ӡ�", vbOKOnly, "���ԤԼ��ʾ"
        txtCapacity.SetFocus
        Exit Function
    End If
    
    '�ж�ʱ��ƻ��Ƿ�����ظ���ʱ��
    strSql = "select ID,��ʼʱ��,����ʱ�� from Ӱ��ԤԼʱ��ƻ� where ԤԼ����ID=[1] order by ��ʼʱ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ѯ�ظ���ʱ��ƻ�", mlngPlanID)
    strS1 = Format(dpTimeStart.value, "HH:MM")
    strE1 = Format(dpTimeEnd.value, "HH:MM")
    
    While rsTemp.EOF = False
        If rsTemp!ID <> mlngTimeProjectID Then
            strS2 = Format(NVL(rsTemp!��ʼʱ��), "HH:MM")
            strE2 = Format(NVL(rsTemp!����ʱ��), "HH:MM")
            If (strS1 <= strS2 And strS2 < strE1) _
                Or (strS1 < strE2 And strE2 <= strE1) _
                Or (strS2 <= strS1 And strS1 < strE2) Then
            
                MsgBox "���������뿪ʼʱ��ͽ���ʱ�䣬����ƻ��������ƻ�����ʱ���ظ���", vbOKOnly, "���ԤԼ��ʾ"
                dpTimeStart.SetFocus
                Exit Function
            End If
        End If
        rsTemp.MoveNext
    Wend
    
    strSql = "Zl_Ӱ��ԤԼʱ��ƻ�_����(" & mlngTimeProjectID & "," & mlngPlanID & "," _
            & zlStr.To_Date(CDate(dpTimeStart.value)) & "," & zlStr.To_Date(CDate(dpTimeEnd.value)) _
            & "," & Val(txtCapacity.Text) & "," & IIF(cboSchExamTimeCalcType.ListIndex = 0, 1, 2) & ")"
    zlDatabase.ExecuteProcedure strSql, "������ԤԼʱ��ƻ�"
    
    saveTimeProject = True
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
End Function

Private Sub txtCapacity_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
