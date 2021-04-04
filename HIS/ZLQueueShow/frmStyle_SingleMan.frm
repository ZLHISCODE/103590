VERSION 5.00
Begin VB.Form frmStyle_SingleMan 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7290
   ClientLeft      =   -60
   ClientTop       =   -45
   ClientWidth     =   11955
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   11955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer tmrRefreshInterval 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7560
      Top             =   240
   End
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   6600
      Top             =   240
   End
   Begin VB.Label lblClinicName1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   8280
      TabIndex        =   6
      Top             =   1680
      Width           =   240
   End
   Begin VB.Label lblPatientInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0002F6FC&
      Height          =   435
      Left            =   1920
      TabIndex        =   5
      Top             =   3120
      Width           =   240
   End
   Begin VB.Image imgDoctor 
      Height          =   1215
      Left            =   7320
      Picture         =   "frmStyle_SingleMan.frx":0000
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1035
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2014��01��19��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   9600
      TabIndex        =   0
      Top             =   600
      Width           =   1470
   End
   Begin VB.Label lblWeek 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   10080
      TabIndex        =   1
      Top             =   120
      Width           =   990
   End
   Begin VB.Label lblClinicName0 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   7920
      TabIndex        =   2
      Top             =   1320
      Width           =   240
   End
   Begin VB.Label lblDoctorJob 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   8640
      TabIndex        =   4
      Top             =   3480
      Width           =   240
   End
   Begin VB.Label lblDoctorName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   8760
      TabIndex        =   3
      Top             =   2880
      Width           =   240
   End
   Begin VB.Image imgBack 
      Height          =   7215
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmStyle_SingleMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements ISty

'��Ҫʵ�ֵĽӿڷ������£�
'
'
'��lcd��ʾ����
'public sub ISty_Show(byval lngWindowNo as long)
'lngWindowNo:���ڱ�ţ����ݴ��ڱ�Ŷ�ȡ������Ϣ����������ʾ
'
'end sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private mlngWindowNo As Long            '���ڱ��
Private mlngRefreshInterval As Long     '��ѯʱ����
Private mlngInterval As Long            '�ۼ�ʱ����
Private mstrStyleTylePath As String     '������ʽͼƬ·��
Private mstrClinicNames As String       '�ٴ��Ŷ�ҵ���µ���������
Private mLcdCommonParameter As TLcdCommonParameter

Private Type TPageObj
    tpBackImage     As TRect        '����(Ƥ��)

    tpClinicName    As TRect     '��������
    tpDoctorPhoto   As TRect     'ҽ����Ƭ
    tpDoctorName    As TRect     'ҽ������
    tpDoctorJob     As TRect     'ҽ��ְ��
    tpPatientInfo   As TRect     '������Ϣ
    tpWeek          As TRect     '����
    tpDate          As TRect     '����
End Type

Private mtpPageObj As TPageObj

Private Sub GetSkinObj(ByVal strSkinName As String)
'��ȡ��ʽ�����ļ����Խ���ؼ�λ�ý��г�ʼ��
    
    Call SetIniFile(strSkinName)
    
    With mtpPageObj
        '����ͼ��С
        .tpBackImage.lngWidth = Val(ReadValue("Ƥ���ֱ���", "��"))
        .tpBackImage.lngHeight = Val(ReadValue("Ƥ���ֱ���", "��"))
        
        '��������
        .tpClinicName.lngLeft = Val(ReadValue("��������", "��"))
        .tpClinicName.lngTop = Val(ReadValue("��������", "��"))
        .tpClinicName.lngWidth = Val(ReadValue("��������", "��"))
        .tpClinicName.lngHeight = Val(ReadValue("��������", "��"))
        
        'ҽ����Ƭ
        .tpDoctorPhoto.lngLeft = Val(ReadValue("ҽ����Ƭ", "��"))
        .tpDoctorPhoto.lngTop = Val(ReadValue("ҽ����Ƭ", "��"))
        .tpDoctorPhoto.lngWidth = Val(ReadValue("ҽ����Ƭ", "��"))
        .tpDoctorPhoto.lngHeight = Val(ReadValue("ҽ����Ƭ", "��"))
        
        'ҽ������
        .tpDoctorName.lngLeft = Val(ReadValue("ҽ������", "��"))
        .tpDoctorName.lngTop = Val(ReadValue("ҽ������", "��"))
        .tpDoctorName.lngWidth = Val(ReadValue("ҽ������", "��"))
        .tpDoctorName.lngHeight = Val(ReadValue("ҽ������", "��"))
        
        'ҽ��ְ��
        .tpDoctorJob.lngLeft = Val(ReadValue("ҽ��ְ��", "��"))
        .tpDoctorJob.lngTop = Val(ReadValue("ҽ��ְ��", "��"))
        .tpDoctorJob.lngWidth = Val(ReadValue("ҽ��ְ��", "��"))
        .tpDoctorJob.lngHeight = Val(ReadValue("ҽ��ְ��", "��"))
        
        '������Ϣ
        .tpPatientInfo.lngLeft = Val(ReadValue("������Ϣ", "��"))
        .tpPatientInfo.lngTop = Val(ReadValue("������Ϣ", "��"))
        .tpPatientInfo.lngWidth = Val(ReadValue("������Ϣ", "��"))
        .tpPatientInfo.lngHeight = Val(ReadValue("������Ϣ", "��"))
        
        '��������
        .tpWeek.lngLeft = Val(ReadValue("��������", "��"))
        .tpWeek.lngTop = Val(ReadValue("��������", "��"))
        .tpWeek.lngWidth = Val(ReadValue("��������", "��"))
        .tpWeek.lngHeight = Val(ReadValue("��������", "��"))
        
        '��������
        .tpDate.lngLeft = Val(ReadValue("��������", "��"))
        .tpDate.lngTop = Val(ReadValue("��������", "��"))
        .tpDate.lngWidth = Val(ReadValue("��������", "��"))
        .tpDate.lngHeight = Val(ReadValue("��������", "��"))
    End With
End Sub

Public Sub ISty_RefreshQueueData(Optional ByVal lngQueueId As Long)
'ˢ�½�����ʾ����
    Call LoadCallingData
    'Call SetStyleFont
    
    '����ˢ�º󽫼�ʱ����0
    mlngInterval = 0
End Sub

'��lcd��ʾ����
Public Sub ISty_Show(ByVal lngWindowNo As Long)
'lngWindowNo:���ڱ�ţ����ݴ��ڱ�Ŷ�ȡ������Ϣ����������ʾ
    mlngWindowNo = lngWindowNo
    
    Call InitMonitor    '��ʼ������������
    
    Call InitLocalPars
    
    Call LoadCallingData
    
    Call SetStyleFont

    Call Show
End Sub

Public Function ISty_ShowCfg(ByVal lngWindowNo As Long, objOwner As Object) As Boolean
'�򿪶�Ӧ����ʽ���ô���
    Dim objConfig As frmStyle_CommonCfg
    
    Set objConfig = New frmStyle_CommonCfg
            
    ISty_ShowCfg = objConfig.OpenShowConfig(lngWindowNo, TShowStyle.ssSingleMan, Me)
End Function


Public Function ISty_MsgProcess(ByVal lngWindowNo As Long, _
    ByVal strMsgKey As String, ByVal strXmlContext As String, rsData As ADODB.Recordset) As Boolean
'��Ϣ���մ���

    Dim strValue As String

On Error GoTo ErrorHand
    
    '�ж���Ϣ�еĶ��������Ƿ���Ҫ���д���Ķ�������
    rsData.Filter = "node_name='queue_name'"
    If rsData.RecordCount <= 0 Then
        Debug.Print "��Ϣ��Ч����⵽δ������Ч�Ķ������ƣ���ֹ��Ϣ����"
        Exit Function
    End If

    strValue = Nvl(rsData!node_value)

    If InStr(mLcdCommonParameter.strQueryQueueNames, strValue) <= 0 Then
        Debug.Print "����Ϣ�������в����ڵ�ǰҵ����Χ��������Ϣ����"
        Exit Function
    End If
    
    '���ݽ��յ�����Ϣ���д���......
    Select Case strMsgKey
        Case G_STR_MSG_QUEUE_001, G_STR_MSG_QUEUE_002, G_STR_MSG_QUEUE_003
            Call ISty_RefreshQueueData
    End Select

    Exit Function
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Function


Public Function ISty_WindNo() As Long
'��ȡ��ǰ��ʽ���ڵı��
    ISty_WindNo = mlngWindowNo
End Function


Private Function InitLocalPars() As Boolean
'��ʼ�����ز�������
    Dim i As Integer
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
    Dim lngCurLCDNo As Long
    Dim strBusinessType As String
    Dim strLCDLocation As String
    Dim strQueryQueueNames As String

On Error GoTo ErrorHand
    If gobjFile.FolderExists(App.Path & "\Skin\��������ʽ") Then
        mstrStyleTylePath = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "Ƥ����ʽ", App.Path & "\Skin\��������ʽ\�����˿�������") & ".jpg"
    Else
        mstrStyleTylePath = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "Ƥ����ʽ", App.Path & "\zlQueueShow\Skin\��������ʽ\�����˿�������") & ".jpg"
    End If
    
    imgBack.Picture = LoadPicture(mstrStyleTylePath)
    
    Call GetSkinObj(Replace(mstrStyleTylePath, ".jpg", ".ini"))
    
    '��ʾ�����
    lngCurLCDNo = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "��ʾ�����", 1)) - 1
    If lngCurLCDNo < 0 Then lngCurLCDNo = 0
        
    '��ʾģʽ,0-ȫ����1-�Զ���
    If Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "��ʾģʽ", 0)) = 0 Then
        Call SetFullScreenWindow(Me, lngCurLCDNo)
    Else
        strLCDLocation = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "�Զ���λ��")
        
        If strLCDLocation <> "" Then
            mLcdCommonParameter.recPos.lngLeft = Mid(Split(strLCDLocation, "|")(0), 3) * Screen.TwipsPerPixelX
            mLcdCommonParameter.recPos.lngTop = Mid(Split(strLCDLocation, "|")(1), 3) * Screen.TwipsPerPixelY
            mLcdCommonParameter.recPos.lngWidth = Mid(Split(strLCDLocation, "|")(2), 3) * Screen.TwipsPerPixelX
            mLcdCommonParameter.recPos.lngHeight = Mid(Split(strLCDLocation, "|")(3), 3) * Screen.TwipsPerPixelY
        End If
        
        Call SetCustomWindow(Me, lngCurLCDNo, mLcdCommonParameter.recPos)
    End If

    '�Ŷ��б�����ʾ�Ķ�����
    strQueryQueueNames = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "��ʾ����", "")
    
    mLcdCommonParameter.blnConvertQueueName = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "ת����������", 0)) = 1
    
    If strQueryQueueNames <> "" Then
        If mLcdCommonParameter.blnConvertQueueName Then    'ת�����ϰ汾��ʽ�Ķ�������
            Select Case glngBusinessType
                Case TBusinessType.btClinical
                    mLcdCommonParameter.strQueryQueueNames = Split(Split(strQueryQueueNames, "|")(1), "_")(0)
                    mstrClinicNames = Split(strQueryQueueNames, ":")(1)
                    
                Case TBusinessType.btPacs
                    For i = 0 To UBound(Split(strQueryQueueNames, ","))
                        If InStr(strQueryQueueNames, "���Ҷ���") > 0 Then    '�������Ŷ�
                            mLcdCommonParameter.strQueryQueueNames = mLcdCommonParameter.strQueryQueueNames & "," & Split(Split(Split(Split(strQueryQueueNames, ",")(i), "|")(1), ":")(0), "_")(1) & "-" & Split(Split(strQueryQueueNames, ",")(i), "|")(0)
                        Else
                            mLcdCommonParameter.strQueryQueueNames = mLcdCommonParameter.strQueryQueueNames & "," & Split(Split(Split(strQueryQueueNames, ",")(i), "|")(1), "_")(0) & ":" & Split(Split(strQueryQueueNames, ",")(i), ":")(1)
                        End If
                    Next
                    
                    mLcdCommonParameter.strQueryQueueNames = Mid(mLcdCommonParameter.strQueryQueueNames, 2)
                    
                Case TBusinessType.btPeis
                    strSql = "select վ������ from ���վ��ֲ� where ִ�п���id=[1]"
                    Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡվ������", Val(Split(Split(strQueryQueueNames, "|")(1), "_")(0)))
                    
                    If rsRecord.RecordCount > 0 Then
                        mLcdCommonParameter.strQueryQueueNames = Nvl(rsRecord!վ������) & ":" & Split(strQueryQueueNames, ":")(1)
                    End If
                'case
                '
                '
            End Select
        Else                                                                '''''''''�°�������Ƹ�ʽ
            Select Case glngBusinessType
                Case TBusinessType.btClinical
                    mLcdCommonParameter.strQueryQueueNames = Split(Split(strQueryQueueNames, "|")(1), "_")(0)
                    mstrClinicNames = Split(strQueryQueueNames, ":")(1)
                    
                Case TBusinessType.btPacs
                    For i = 0 To UBound(Split(strQueryQueueNames, ","))
                        mLcdCommonParameter.strQueryQueueNames = mLcdCommonParameter.strQueryQueueNames & "," & Split(Split(strQueryQueueNames, ",")(i), "|")(0) & "-" & Split(Split(strQueryQueueNames, ",")(i), ":")(1)
                    Next
                    
                    mLcdCommonParameter.strQueryQueueNames = Mid(mLcdCommonParameter.strQueryQueueNames, 2)
                    
                Case TBusinessType.btPeis
                    strSql = "select վ������ from ���վ��ֲ� where ִ�п���id=[1]"
                    Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡվ������", Val(Split(Split(strQueryQueueNames, "|")(1), "_")(0)))
                    
                    If rsRecord.RecordCount > 0 Then
                        mLcdCommonParameter.strQueryQueueNames = Nvl(rsRecord!վ������) & ":" & Split(strQueryQueueNames, ":")(1)
                    End If
                'case
                '
                '
            End Select
        End If
    End If
    
    '��ǰ��������
    mLcdCommonParameter.strCurDiagnoseRoom = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "����ִ�м�", "")
    If mLcdCommonParameter.strCurDiagnoseRoom = "" Then
        strSql = "select d.���� from �ϻ���Ա�� A,��Ա�� B,������Ա C,���ű� D " & _
                 "where A.��ԱID=B.ID And b.id=c.��Աid and c.����id=d.id and c.ȱʡ=1 and A.�û���=[1]"
        
        Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "", UCase(gstrUserName))
        
        If rsRecord.RecordCount > 0 Then lblClinicName0.Caption = Nvl(rsRecord!����)
    Else
        If InStr(strQueryQueueNames, "���Ҷ���") > 0 Then
            lblClinicName0.Caption = Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(0)
        Else
            If Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "���ұ����Ƿ���ʾ������", 0)) = 1 Then
                If Trim(Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(1)) <> "" Then
                    lblClinicName0.Caption = Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(0) & Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(1)
                Else
                    lblClinicName0.Caption = Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(0)
                End If
            Else
                If Trim(Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(1)) <> "" Then
                    lblClinicName0.Caption = Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(1)
                Else
                    lblClinicName0.Caption = Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(0)
                End If
            End If
        End If
    End If
    
    If Len(lblClinicName0.Caption) <= 5 Then lblClinicName0 = FormatStr(lblClinicName0.Caption)
    
    lblClinicName1.Caption = lblClinicName0.Caption
    
    mLcdCommonParameter.blnFontAutoSizeToList = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "�б���������Ӧ", True)
    
    '�Ŷ��б���ѯ���
    mlngRefreshInterval = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "��ѯ���", 30))

    Call LoadDoctorInfo
    
    tmrRefreshInterval.Enabled = True
Exit Function
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Function

Private Function FormatStr(ByVal strSources As String) As String
'���ܣ����ַ����еĺ���֮����Ͽո�
    Dim i As Integer
    Dim strResult As String
    Dim strCurS As String
    Dim strNextS As String
    
    If Len(strSources) <= 1 Then Exit Function
    
    For i = 1 To Len(strSources) - 1
        strCurS = Mid(strSources, i, 1)
        strNextS = Mid(strSources, i + 1, 1)
        strResult = strResult & strCurS
        
        If Not (Asc(strNextS) < 255 And Asc(strNextS) > 0) Then
            strResult = strResult & " "
        End If
    Next
    
    FormatStr = strResult & strNextS
End Function

Private Sub LoadDoctorInfo()
'���ض�Ӧִ�м��ҽ���Ϳ��������Ϣ
    Dim i As Integer
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
    Dim blnNotWorkingTime As Boolean

    Dim strDoctorInfo As String     '�����ʽ��"ҽ��1��������ְλ|ҽ��2��������ְλ|��������"
    Dim strDoctorPhoto As String    '�����ʽ��"ҽ��1����Ƭ|ҽ��2����Ƭ|��������"
    Dim strIntroduction As String   '�����ʽ��"ҽ��1�ļ��|ҽ��2�ļ��|��������"
    Dim strWorkingTime As String    '�����ʽ��"ҽ��1��ֵ��ʱ��|ҽ��2��ֵ��ʱ��|��������"
    
    blnNotWorkingTime = True
    
    strDoctorInfo = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "ҽ����Ϣ")    '
    strDoctorPhoto = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "ҽ����Ƭ")    '
    strWorkingTime = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "ֵ��ʱ��")   '
    strIntroduction = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "ҽ�����")   '
    
    '���ݵ�ǰ���ڶ�ȡ��Ӧ��ҽ��������Ϣ
    For i = 0 To UBound(Split(Mid(strWorkingTime, 2), "|"))
        If Split(Mid(strWorkingTime, 2), "|")(i) = lblWeek.Caption Then
            blnNotWorkingTime = False
            
            lblDoctorName.Caption = Split(Split(Mid(strDoctorInfo, 2), "|")(i), "-")(1)
            
            If lblDoctorName.Caption <> "" Then
                strSql = "select רҵ����ְ�� from ��Ա�� where id=[1]"
                Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "", Val(Split(Split(Mid(strDoctorInfo, 2), "|")(i), "-")(0)))
                
                If rsRecord.RecordCount > 0 Then
                    lblDoctorJob.Caption = Nvl(rsRecord!רҵ����ְ��)
                End If
            End If
            
            Call LoadPictureInfo(imgDoctor, Split(Mid(strDoctorPhoto, 2), "|")(i))
            
            Exit Sub
        End If
    Next
    
    '���ֻ����ҽ����û��ָ��ҽ����ֵ����Ϣ�����ȡ��½��Ա��Ϣ
    If blnNotWorkingTime Then
        strSql = "select B.����,B.ID from �ϻ���Ա�� A,��Ա�� B where A.��ԱID=B.ID And A.�û���=[1]"
        Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "", UCase(gstrUserName))

        If rsRecord.RecordCount > 0 Then
            For i = 0 To UBound(Split(Mid(strDoctorInfo, 2), "|"))
                If Split(Split(Mid(strDoctorInfo, 2), "|")(i), "-")(0) = Nvl(rsRecord!����) Then
                    lblDoctorName.Caption = Split(Split(Mid(strDoctorInfo, 2), "|")(i), "-")(1)
                    
                    If lblDoctorName.Caption <> "" Then
                        strSql = "select רҵ����ְ�� from ��Ա�� where id=[1]"
                        Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "", Split(Split(Mid(strDoctorInfo, 2), "|")(i), "-")(0))
                        
                        If rsRecord.RecordCount > 0 Then
                            lblDoctorJob.Caption = Nvl(rsRecord!רҵ����ְ��)
                        End If
                    End If
                    
                    Call LoadPictureInfo(imgDoctor, Split(Mid(strDoctorPhoto, 2), "|")(i))

                    Exit Sub
                End If
            Next
        End If
    End If
End Sub

Private Sub SetStyleFont()
'���ý�����ؼ���������
    '���ý�����ؼ���������
    Dim i As Integer
    Dim strFontPropertys As String           '��ʽ:"����:����|�ֺ�:20|����:FALSE|б��:FALSE|�»���:FALSE|ǰ��ɫ:16777215"
    Dim strFontProperty() As String
On Error GoTo ErrorHand

    '��������
    strFontPropertys = Trim(ReadValue("��������", "������������", "����:΢���ź�|�ֺ�:50|����:TRUE|б��:FALSE|�»���:FALSE|ǰ��ɫ:194300"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblClinicName0, strFontProperty)
        Call SetControlFont(lblClinicName1, strFontProperty)
        lblClinicName0.ForeColor = vbBlack
    End If
    
    'ҽ��������ְ��
    strFontPropertys = Trim(ReadValue("��������", "ҽ����Ϣ����", "����:΢���ź�|�ֺ�:22|����:TRUE|б��:FALSE|�»���:FALSE|ǰ��ɫ:0"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblDoctorName, strFontProperty)
        Call SetControlFont(lblDoctorJob, strFontProperty)
    End If
    
    '������Ϣ����
    strFontPropertys = Trim(ReadValue("��������", "������Ϣ����", "����:΢���ź�|�ֺ�:70|����:TRUE|б��:FALSE|�»���:FALSE|ǰ��ɫ:16777215"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblPatientInfo, strFontProperty)
    End If
    
    '����
    strFontPropertys = Trim(ReadValue("��������", "��������", "����:΢���ź�|�ֺ�:15|����:FALSE|б��:FALSE|�»���:FALSE|ǰ��ɫ:16777215"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblWeek, strFontProperty)
    End If

    '����
    strFontPropertys = Trim(ReadValue("��������", "��������", "����:΢���ź�|�ֺ�:15|����:FALSE|б��:FALSE|�»���:FALSE|ǰ��ɫ:16777215"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblDate, strFontProperty)
    End If
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub LoadCallingData()
'���ش��ں����е�����
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
    Dim dblHeightScale As Double, dblWidhtScale As Double
    
On Error GoTo ErrorHand:
    lblPatientInfo.Caption = ""
    
    If mLcdCommonParameter.strQueryQueueNames = "" Or glngBusinessType < 0 Then Exit Sub
    
    Select Case glngBusinessType
        Case TBusinessType.btClinical
            strSql = "select �ŶӺ���,��������,�Ŷ�״̬,����ʱ��,��ע,�Ŷ����,���� from �ŶӽкŶ��� where ��������=[1] and ����=[2] and ҵ������=[3] and " & _
                     "�Ŷ�״̬ in (1,7,8,9) And �Ŷ�ʱ�� > trunc(sysdate) "
            
            Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡ�Ŷ���Ϣ", mLcdCommonParameter.strQueryQueueNames, mstrClinicNames, glngBusinessType)
            
        Case TBusinessType.btPacs
            strSql = "select �ŶӺ���,��������,�Ŷ�״̬,����ʱ��,��ע,�Ŷ����,���� from �ŶӽкŶ��� where ��������=[1] and ����=[2] and ҵ������=[3] and " & _
                     "�Ŷ�״̬ in (1,7,8,9) And �Ŷ�ʱ�� > trunc(sysdate) "
            
            If mLcdCommonParameter.strCurDiagnoseRoom = "" Then Exit Sub
            Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡ�Ŷ���Ϣ", mLcdCommonParameter.strQueryQueueNames, CStr(Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(1)), glngBusinessType)
            
        Case TBusinessType.btPeis
            strSql = "select �ŶӺ���,��������,�Ŷ�״̬,����ʱ��,��ע,�Ŷ����,���� from �ŶӽкŶ��� where ��������=[1] and ҵ������=[2] and " & _
                     "�Ŷ�״̬ in (1,7,8,9) And �Ŷ�ʱ�� > trunc(sysdate) "
            
            Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡ�Ŷ���Ϣ", mLcdCommonParameter.strQueryQueueNames, glngBusinessType)
         
        'case
        '
        '
    End Select

    If rsRecord.RecordCount >= 1 Then rsRecord.Filter = "�Ŷ�״̬=9"
    If rsRecord.RecordCount <= 0 Then rsRecord.Filter = "�Ŷ�״̬=1"
    If rsRecord.RecordCount <= 0 Then rsRecord.Filter = "�Ŷ�״̬=7"
    If rsRecord.RecordCount <= 0 Then rsRecord.Filter = "�Ŷ�״̬=8"
    If rsRecord.RecordCount <= 0 Then Exit Sub
    
    rsRecord.Sort = "����ʱ�� desc"
    lblPatientInfo.Caption = Format(Nvl(rsRecord!�ŶӺ���), "000") & "��   " & Nvl(rsRecord!��������) & IIf(Len(Trim(Nvl(rsRecord!��������))) <= 3, "   ", "")
    
    '�������õ�ǰ������Ϣ��ʾλ�ú������С
    dblHeightScale = imgBack.Height / mtpPageObj.tpBackImage.lngHeight
    dblWidhtScale = imgBack.Width / mtpPageObj.tpBackImage.lngWidth
    
    If mLcdCommonParameter.blnFontAutoSizeToList Then
        lblPatientInfo.Caption = Trim(lblPatientInfo.Caption)
        
        lblPatientInfo.Visible = False
        lblPatientInfo.FontSize = dblHeightScale * mtpPageObj.tpPatientInfo.lngHeight / 20

        While lblPatientInfo.Width > dblHeightScale * mtpPageObj.tpPatientInfo.lngWidth Or lblPatientInfo.Height > dblHeightScale * mtpPageObj.tpPatientInfo.lngHeight
            lblPatientInfo.FontSize = lblPatientInfo.FontSize - 1
        Wend
        lblPatientInfo.Visible = True
    End If
    
    lblPatientInfo.Left = dblWidhtScale * mtpPageObj.tpPatientInfo.lngLeft + dblWidhtScale * mtpPageObj.tpPatientInfo.lngWidth / 2 - lblPatientInfo.Width / 2
    lblPatientInfo.Top = dblHeightScale * mtpPageObj.tpPatientInfo.lngTop + dblHeightScale * mtpPageObj.tpPatientInfo.lngHeight / 2 - lblPatientInfo.Height / 2
Exit Sub
ErrorHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHand
    mlngInterval = 0
    tmrRefreshInterval.Interval = 1000
    
    Call refreshWeekLab
    
    lblDate.Caption = Format(Now, "yyyy��mm��dd�� hh:mm:ss")
Exit Sub
ErrorHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    Dim dblHeightScale As Double, dblWidhtScale As Double
    
    '���屳��
    imgBack.Left = 0
    imgBack.Top = 0
    imgBack.Height = Me.ScaleHeight
    imgBack.Width = Me.ScaleWidth
    
    dblHeightScale = imgBack.Height / mtpPageObj.tpBackImage.lngHeight
    dblWidhtScale = imgBack.Width / mtpPageObj.tpBackImage.lngWidth

    '��������
    lblClinicName0.Left = dblWidhtScale * mtpPageObj.tpClinicName.lngLeft + dblWidhtScale * mtpPageObj.tpClinicName.lngWidth / 2 - lblClinicName0.Width / 2
    lblClinicName0.Top = dblHeightScale * mtpPageObj.tpClinicName.lngTop + dblHeightScale * mtpPageObj.tpClinicName.lngHeight / 2 - lblClinicName0.Height / 2
    
    lblClinicName1.Left = lblClinicName0.Left - 50
    lblClinicName1.Top = lblClinicName0.Top - 50
    'ҽ����Ƭ
    Call ResizeImg(imgDoctor, dblWidhtScale * mtpPageObj.tpDoctorPhoto.lngLeft, dblHeightScale * mtpPageObj.tpDoctorPhoto.lngTop, dblWidhtScale * mtpPageObj.tpDoctorPhoto.lngWidth, dblHeightScale * mtpPageObj.tpDoctorPhoto.lngHeight)

    'ҽ������
    lblDoctorName.Left = dblWidhtScale * mtpPageObj.tpDoctorName.lngLeft + dblWidhtScale * mtpPageObj.tpDoctorName.lngWidth / 2 - lblDoctorName.Width / 2
    lblDoctorName.Top = dblHeightScale * mtpPageObj.tpDoctorName.lngTop + dblHeightScale * mtpPageObj.tpDoctorName.lngHeight / 2 - lblDoctorName.Height / 2
    
    'ҽ��ְλ
    lblDoctorJob.Left = dblWidhtScale * mtpPageObj.tpDoctorJob.lngLeft + dblWidhtScale * mtpPageObj.tpDoctorJob.lngWidth / 2 - lblDoctorJob.Width / 2
    lblDoctorJob.Top = dblHeightScale * mtpPageObj.tpDoctorJob.lngTop + dblHeightScale * mtpPageObj.tpDoctorJob.lngHeight / 2 - lblDoctorJob.Height / 2
    
    '������Ϣ
    If mLcdCommonParameter.blnFontAutoSizeToList Then
        lblPatientInfo.Caption = Trim(lblPatientInfo.Caption)
        
        lblPatientInfo.Visible = False
        lblPatientInfo.FontSize = dblHeightScale * mtpPageObj.tpPatientInfo.lngHeight / 20

        While lblPatientInfo.Width > dblHeightScale * mtpPageObj.tpPatientInfo.lngWidth Or lblPatientInfo.Height > dblHeightScale * mtpPageObj.tpPatientInfo.lngHeight
            lblPatientInfo.FontSize = lblPatientInfo.FontSize - 1
        Wend
        lblPatientInfo.Visible = True
    End If
    
    lblPatientInfo.Left = dblWidhtScale * mtpPageObj.tpPatientInfo.lngLeft + dblWidhtScale * mtpPageObj.tpPatientInfo.lngWidth / 2 - lblPatientInfo.Width / 2
    lblPatientInfo.Top = dblHeightScale * mtpPageObj.tpPatientInfo.lngTop + dblHeightScale * mtpPageObj.tpPatientInfo.lngHeight / 2 - lblPatientInfo.Height / 2
    
    '����
    lblDate.Left = dblWidhtScale * mtpPageObj.tpDate.lngLeft + dblWidhtScale * mtpPageObj.tpDate.lngWidth / 2 - lblDate.Width / 2
    lblDate.Top = dblHeightScale * mtpPageObj.tpDate.lngTop + dblHeightScale * mtpPageObj.tpDate.lngHeight / 2 - lblDate.Height / 2
    
    '����
    lblWeek.Left = dblWidhtScale * mtpPageObj.tpWeek.lngLeft + dblWidhtScale * mtpPageObj.tpWeek.lngWidth / 2 - lblWeek.Width / 2
    lblWeek.Top = dblHeightScale * mtpPageObj.tpWeek.lngTop + dblHeightScale * mtpPageObj.tpWeek.lngHeight / 2 - lblWeek.Height / 2
End Sub

Private Sub tmrRefreshInterval_Timer()
On Error GoTo ErrorHand
    mlngInterval = mlngInterval + 1

    '��Timer�ۼƵ�ʱ��С����ѯʱ��ʱ������ˢ���Ŷ�����
    If mlngInterval < mlngRefreshInterval Then Exit Sub
    '���ۼƵ�ʱ����0
    mlngInterval = 0
    
    Call LoadCallingData
'    Call SetStyleFont
Exit Sub
ErrorHand:
    Debug.Print Err.Description
    Err.Clear
End Sub

Private Sub tmrTime_Timer()
On Error GoTo ErrorHand
    Call refreshWeekLab
    lblDate.Caption = Format(Now, "yyyy��mm��dd�� hh:mm:ss")
Exit Sub
ErrorHand:
    Debug.Print Err.Description
    Err.Clear
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHand
    If KeyAscii = vbKeyEscape Then Call CloseStyleWindow
Exit Sub
ErrorHand:
    Debug.Print Err.Description
    Err.Clear
End Sub

Private Sub refreshWeekLab()
    Select Case Weekday(Date)
        Case 1
            lblWeek.Caption = "������"
        Case 2
            lblWeek.Caption = "����һ"
        Case 3
            lblWeek.Caption = "���ڶ�"
        Case 4
            lblWeek.Caption = "������"
        Case 5
            lblWeek.Caption = "������"
        Case 6
            lblWeek.Caption = "������"
        Case 7
            lblWeek.Caption = "������"
    End Select
End Sub


