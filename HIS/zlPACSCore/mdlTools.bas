Attribute VB_Name = "mdlTools"
Option Explicit
'--------------------------------------------------------
'��  �ܣ���ģ��Ϊ�˵�����ť���Ų���
'�����ˣ�����
'�������ڣ�2004.6.12
'���̺����嵥��
'    BarterIco():        ����ͼ�񼯵�Tagֵ��
'    CreateMenu()��      �����˵�
'    StatusBarTip():     ��״̬����ʾ�˵��ļ򵥰���
'    funcGetShiftStr():  ͨ������shift״̬��ֵ�����������ַ�����ʾ��shift״̬��
'    ArrayToolBar():     ���°�һ����˳��ڷŹ�����λ��
'    ReplaceToolBarIcon():   �滻��ǰͼ��Ϊ16,24,32
'    PutToolbar():       �ڷŹ�������Top,Left,Right,Bottom
'�޸ļ�¼��
'    2005.7.02    �ƽݣ�����
'-------------------------------------------------------
Public blfrmRefresh As Boolean          ''''�����Ƿ�ˢ�£����ڹ�����ˢ��ʱ���岻һ��ˢ��)
Public gstrSysName As String

Public Sub BarterIco(ImgBox As ImageList)
'------------------------------------------------
'���ܣ�����ͼ�񼯵�Tagֵ����Tagֵ��������ʵ�ָ���������ͼ�꣩
'������ImgBox--��Ҫ����ͼ��Tagֵ��ͼ�񼯡�
'���أ���
'�ϼ���������̣�frmViewer.Form_Load��
'�¼���������̣���
'���õ��ⲿ������cMouseUsage
'�����ˣ�����
'------------------------------------------------
    If cMouseUsage("101").lngMouseKey = 1 Then
        ImgBox.ListImages("����L").Tag = IIf(ImgBox.ListImages("����L").Tag = "", ImgBox.ListImages("����R").Tag, ImgBox.ListImages("����L").Tag)
        ImgBox.ListImages("����R").Tag = ""
    Else
        ImgBox.ListImages("����R").Tag = IIf(ImgBox.ListImages("����R").Tag = "", ImgBox.ListImages("����L").Tag, ImgBox.ListImages("����R").Tag)
        ImgBox.ListImages("����L").Tag = ""
    End If
    
    If cMouseUsage("103").lngMouseKey = 1 Then
        ImgBox.ListImages("����L").Tag = IIf(ImgBox.ListImages("����L").Tag = "", ImgBox.ListImages("����R").Tag, ImgBox.ListImages("����L").Tag)
        ImgBox.ListImages("����R").Tag = ""
    Else
        ImgBox.ListImages("����R").Tag = IIf(ImgBox.ListImages("����R").Tag = "", ImgBox.ListImages("����L").Tag, ImgBox.ListImages("����R").Tag)
        ImgBox.ListImages("����L").Tag = ""
    End If

    If cMouseUsage("201").lngMouseKey = 1 Then
        ImgBox.ListImages("�ü�L").Tag = IIf(ImgBox.ListImages("�ü�L").Tag = "", ImgBox.ListImages("�ü�R").Tag, ImgBox.ListImages("�ü�L").Tag)
        ImgBox.ListImages("�ü�R").Tag = ""
        
        ImgBox.ListImages("��ѡL").Tag = IIf(ImgBox.ListImages("��ѡL").Tag = "", ImgBox.ListImages("��ѡR").Tag, ImgBox.ListImages("��ѡL").Tag)
        ImgBox.ListImages("��ѡR").Tag = ""
    Else
        ImgBox.ListImages("�ü�R").Tag = IIf(ImgBox.ListImages("�ü�R").Tag = "", ImgBox.ListImages("�ü�L").Tag, ImgBox.ListImages("�ü�R").Tag)
        ImgBox.ListImages("�ü�L").Tag = ""
        
        ImgBox.ListImages("��ѡR").Tag = IIf(ImgBox.ListImages("��ѡR").Tag = "", ImgBox.ListImages("��ѡL").Tag, ImgBox.ListImages("��ѡR").Tag)
        ImgBox.ListImages("��ѡL").Tag = ""
    End If


    If cMouseUsage("102").lngMouseKey = 1 Then
        ImgBox.ListImages("�ֶ�����L").Tag = IIf(ImgBox.ListImages("�ֶ�����L").Tag = "", ImgBox.ListImages("�ֶ�����R").Tag, ImgBox.ListImages("�ֶ�����L").Tag)
        ImgBox.ListImages("�ֶ�����R").Tag = ""
    Else
        ImgBox.ListImages("�ֶ�����R").Tag = IIf(ImgBox.ListImages("�ֶ�����R").Tag = "", ImgBox.ListImages("�ֶ�����L").Tag, ImgBox.ListImages("�ֶ�����R").Tag)
        ImgBox.ListImages("�ֶ�����L").Tag = ""
    End If

    If cMouseUsage("105").lngMouseKey = 1 Then
        ImgBox.ListImages("����Ӧ����L").Tag = IIf(ImgBox.ListImages("����Ӧ����L").Tag = "", ImgBox.ListImages("����Ӧ����R").Tag, ImgBox.ListImages("����Ӧ����L").Tag)
        ImgBox.ListImages("����Ӧ����R").Tag = ""
    Else
        ImgBox.ListImages("����Ӧ����R").Tag = IIf(ImgBox.ListImages("����Ӧ����R").Tag = "", ImgBox.ListImages("����Ӧ����L").Tag, ImgBox.ListImages("����Ӧ����R").Tag)
        ImgBox.ListImages("����Ӧ����L").Tag = ""
    End If
    
    If cMouseUsage("106").lngMouseKey = 1 Then
        ImgBox.ListImages("��ά���L").Tag = IIf(ImgBox.ListImages("��ά���L").Tag = "", ImgBox.ListImages("��ά���R").Tag, ImgBox.ListImages("��ά���L").Tag)
        ImgBox.ListImages("��ά���R").Tag = ""
    Else
        ImgBox.ListImages("��ά���R").Tag = IIf(ImgBox.ListImages("��ά���R").Tag = "", ImgBox.ListImages("��ά���L").Tag, ImgBox.ListImages("��ά���R").Tag)
        ImgBox.ListImages("��ά���L").Tag = ""
    End If

    If cMouseUsage("8").lngMouseKey = 1 Then
        ImgBox.ListImages("����L").Tag = IIf(ImgBox.ListImages("����L").Tag = "", ImgBox.ListImages("����R").Tag, ImgBox.ListImages("����L").Tag)
        ImgBox.ListImages("����R").Tag = ""
    Else
        ImgBox.ListImages("����R").Tag = IIf(ImgBox.ListImages("����R").Tag = "", ImgBox.ListImages("����L").Tag, ImgBox.ListImages("����R").Tag)
        ImgBox.ListImages("����L").Tag = ""
    End If

    If cMouseUsage("4").lngMouseKey = 1 Then
        ImgBox.ListImages("��ͷL").Tag = IIf(ImgBox.ListImages("��ͷL").Tag = "", ImgBox.ListImages("��ͷR").Tag, ImgBox.ListImages("��ͷL").Tag)
        ImgBox.ListImages("��ͷR").Tag = ""
    Else
        ImgBox.ListImages("��ͷR").Tag = IIf(ImgBox.ListImages("��ͷR").Tag = "", ImgBox.ListImages("��ͷL").Tag, ImgBox.ListImages("��ͷR").Tag)
        ImgBox.ListImages("��ͷL").Tag = ""
    End If

    If cMouseUsage("3").lngMouseKey = 1 Then
        ImgBox.ListImages("��ԲL").Tag = IIf(ImgBox.ListImages("��ԲL").Tag = "", ImgBox.ListImages("��ԲR").Tag, ImgBox.ListImages("��ԲL").Tag)
        ImgBox.ListImages("��ԲR").Tag = ""
    Else
        ImgBox.ListImages("��ԲR").Tag = IIf(ImgBox.ListImages("��ԲR").Tag = "", ImgBox.ListImages("��ԲL").Tag, ImgBox.ListImages("��ԲR").Tag)
        ImgBox.ListImages("��ԲL").Tag = ""
    End If

    If cMouseUsage("7").lngMouseKey = 1 Then
        ImgBox.ListImages("�Ƕ�L").Tag = IIf(ImgBox.ListImages("�Ƕ�L").Tag = "", ImgBox.ListImages("�Ƕ�R").Tag, ImgBox.ListImages("�Ƕ�L").Tag)
        ImgBox.ListImages("�Ƕ�R").Tag = ""
    Else
        ImgBox.ListImages("�Ƕ�R").Tag = IIf(ImgBox.ListImages("�Ƕ�R").Tag = "", ImgBox.ListImages("�Ƕ�L").Tag, ImgBox.ListImages("�Ƕ�R").Tag)
        ImgBox.ListImages("�Ƕ�L").Tag = ""
    End If
    
    If cMouseUsage("6").lngMouseKey = 1 Then
        ImgBox.ListImages("����L").Tag = IIf(ImgBox.ListImages("����L").Tag = "", ImgBox.ListImages("����R").Tag, ImgBox.ListImages("����L").Tag)
        ImgBox.ListImages("����R").Tag = ""
    Else
        ImgBox.ListImages("����R").Tag = IIf(ImgBox.ListImages("����R").Tag = "", ImgBox.ListImages("����L").Tag, ImgBox.ListImages("����R").Tag)
        ImgBox.ListImages("����L").Tag = ""
    End If
    
    If cMouseUsage("5").lngMouseKey = 1 Then
        ImgBox.ListImages("����L").Tag = IIf(ImgBox.ListImages("����L").Tag = "", ImgBox.ListImages("����R").Tag, ImgBox.ListImages("����L").Tag)
        ImgBox.ListImages("����R").Tag = ""
    Else
        ImgBox.ListImages("����R").Tag = IIf(ImgBox.ListImages("����R").Tag = "", ImgBox.ListImages("����L").Tag, ImgBox.ListImages("����R").Tag)
        ImgBox.ListImages("����L").Tag = ""
    End If
    
    If cMouseUsage("1").lngMouseKey = 1 Then
        ImgBox.ListImages("ֱ��L").Tag = IIf(ImgBox.ListImages("ֱ��L").Tag = "", ImgBox.ListImages("ֱ��R").Tag, ImgBox.ListImages("ֱ��L").Tag)
        ImgBox.ListImages("ֱ��R").Tag = ""
        ImgBox.ListImages("Ѫ�ܲ���L").Tag = IIf(ImgBox.ListImages("Ѫ�ܲ���L").Tag = "", ImgBox.ListImages("Ѫ�ܲ���R").Tag, ImgBox.ListImages("Ѫ�ܲ���L").Tag)
        ImgBox.ListImages("Ѫ�ܲ���R").Tag = ""
        ImgBox.ListImages("���ر�L").Tag = IIf(ImgBox.ListImages("���ر�L").Tag = "", ImgBox.ListImages("���ر�R").Tag, ImgBox.ListImages("���ر�L").Tag)
        ImgBox.ListImages("���ر�R").Tag = ""
    Else
        ImgBox.ListImages("ֱ��R").Tag = IIf(ImgBox.ListImages("ֱ��R").Tag = "", ImgBox.ListImages("ֱ��L").Tag, ImgBox.ListImages("ֱ��R").Tag)
        ImgBox.ListImages("ֱ��L").Tag = ""
        ImgBox.ListImages("Ѫ�ܲ���R").Tag = IIf(ImgBox.ListImages("Ѫ�ܲ���R").Tag = "", ImgBox.ListImages("Ѫ�ܲ���L").Tag, ImgBox.ListImages("Ѫ�ܲ���R").Tag)
        ImgBox.ListImages("Ѫ�ܲ���L").Tag = ""
        ImgBox.ListImages("���ر�R").Tag = IIf(ImgBox.ListImages("���ر�R").Tag = "", ImgBox.ListImages("���ر�L").Tag, ImgBox.ListImages("���ر�R").Tag)
        ImgBox.ListImages("���ر�L").Tag = ""
    End If

    If cMouseUsage("2").lngMouseKey = 1 Then
        ImgBox.ListImages("����L").Tag = IIf(ImgBox.ListImages("����L").Tag = "", ImgBox.ListImages("����R").Tag, ImgBox.ListImages("����L").Tag)
        ImgBox.ListImages("����R").Tag = ""
    Else
        ImgBox.ListImages("����R").Tag = IIf(ImgBox.ListImages("����R").Tag = "", ImgBox.ListImages("����L").Tag, ImgBox.ListImages("����R").Tag)
        ImgBox.ListImages("����L").Tag = ""
    End If

    If cMouseUsage("104").lngMouseKey = 1 Then
        ImgBox.ListImages("����L").Tag = IIf(ImgBox.ListImages("����L").Tag = "", ImgBox.ListImages("����R").Tag, ImgBox.ListImages("����L").Tag)
        ImgBox.ListImages("����R").Tag = ""
    Else
        ImgBox.ListImages("����R").Tag = IIf(ImgBox.ListImages("����R").Tag = "", ImgBox.ListImages("����L").Tag, ImgBox.ListImages("����R").Tag)
        ImgBox.ListImages("����L").Tag = ""
    End If
    
End Sub

Public Sub CreateMenu(ToolBars As Object, IconX As Integer, IconY As Integer)
    '------------------------------------------------
    '���ܣ�                                  �����˵�
    '������
    '           IconX                        ����ͼ��X��С
    '           IconY                        ����ͼ��Y��С
    '���أ�                                  ��
    '�ϼ���������̣�                        frViewer_load
    '�¼���������̣�                        ��
    '���õ��ⲿ������                        ��
    '�����ˣ�                                ���� 2005-6-27
    '------------------------------------------------
    Dim control As CommandBarControl
    Dim ControlFile As CommandBarPopup
    Dim ControlSelect As CommandBarPopup
    
    ToolBars.Options.UseDisabledIcons = True
    ToolBars.ActiveMenuBar.EnableDocking xtpFlagHideWrap
    '�����˵�
    '''''''''''''''''''''''''''''''''''''''�ļ��˵�''''''''''''''''''''''''''''''''''''''''''''''
    Set ControlFile = ToolBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "�ļ�(&F)", -1, False)
    With ControlFile.CommandBar.Controls
        .Add xtpControlButton, ID_File_Open, "��(&O)"
        .Add xtpControlButton, ID_File_Close, "�ر�����(&C)"
        .Add xtpControlButton, ID_File_DelAllPhoto, "ɾ������ͼ��(&K)"
        .Add xtpControlButton, ID_File_DelReport, "ɾ������ͼ��(&D)"
        
        
        Set control = .Add(xtpControlButton, ID_File_SaveFile, "�����ļ�(&S)")
        control.BeginGroup = True
        
        .Add xtpControlButton, ID_File_SaveASFile, "����ļ�(&A)", -1, False
        .Add xtpControlButton, ID_File_SaveToCD, "����CD", -1, False
        .Add xtpControlButton, ID_File_SAveASReport, "���汨��ͼ(&R)", -1, False
        
        Set ControlSelect = .Add(xtpControlPopup, ID_File_Send, "����")
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_File_Send_GetHost, "��������(&H)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_File_Send_OutPowerPoint, "�����PowerPoint"
        ControlSelect.BeginGroup = True
        .Add xtpControlButton, ID_File_OpenDicomDir, "��DICOMDIR"
        .Add xtpControlButton, ID_File_PhotoProperty, "ͼ������(&I)"
        
        Set control = .Add(xtpControlButton, ID_File_Exit, "�˳�(&X)")
        control.BeginGroup = True
    End With
    ''''''''''''''''''''''''''''''''''''''''''��ͼ�˵�''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set ControlFile = ToolBars.ActiveMenuBar.Controls.Add(xtpControlPopup, ID_View, "��ͼ(&V)", -1, False)
    With ControlFile.CommandBar.Controls
        .Add xtpControlButton, ID_View_UpSeries, "��һ����"
        .Add xtpControlButton, ID_View_DownSeries, "��һ����"
        .Add xtpControlButton, ID_View_Typeset, "���氲��(T)"
        Set control = .Add(xtpControlButton, ID_View_OneBrowse, "�����й۲�")
        control.BeginGroup = True
        Set control = .Add(xtpControlButton, ID_View_PropertyShow, "������ʾ(&P)")
        control.Checked = True
        Set control = .Add(xtpControlButton, ID_View_LableShow, "��ע��ʾ(&L)")
        control.Checked = True
        Set control = .Add(xtpControlButton, ID_View_ShowOverlay, "��ʾOverlay")
        control.Checked = True
        .Add xtpControlButton, ID_View_ShowMiniSeries, "��ʾ��������ͼ(&M)"
        .Add xtpControlButton, ID_View_ViewAllSeries, "ȫ���й�Ƭ"
        
        Set ControlSelect = .Add(xtpControlPopup, ID_View_PhotoSerial, "ͼ��˳��(&S)")
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_PhotoNumber, "ͼ���(&1)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_BedASC, "��λ����(&2)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_BedDESC, "��λ����(&3)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_CollectionTime, "�ɼ�ʱ��(&4)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_PhotoTime, "ͼ��ʱ��(&5)"
        
        Set ControlSelect = .Add(xtpControlPopup, ID_View_ShowScale, "��ʾ����(&Z)")
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_AutoShow, "����Ӧ(&O)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_50%, "50%"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_100%, "100%"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_showScale_150%, "150%"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_200%, "200%"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_showScale_250%, "250%"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_showScale_300%, "300%"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_showScale_400%, "400%"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_Custom, "�Զ���(&A)"
        ControlSelect.BeginGroup = True
        
        .Add xtpControlButton, ID_View_FullScreen, "ȫ����ʾ(&U)", -1, False
    End With
    ''''''''''''''''''''''''''''''''''''''''''''''�����˵�'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set ControlFile = ToolBars.ActiveMenuBar.Controls.Add(xtpControlPopup, ID_Active, "����(&A)", -1, False)
    With ControlFile.CommandBar.Controls
        
        Set ControlSelect = .Add(xtpControlPopup, ID_Active_Select, "ѡ��(&S)")
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Select_OneSelect, "����ѡ��(&O)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Select_SelectAllSerial, "ѡ����������(&S)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Acitve_Select_SelectAllPhoto, "ͼ��ȫѡ(&A)"
        
        Set ControlSelect = .Add(xtpControlPopup, ID_Active_Also, "ͬ��")
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Also_Serial, "����ͬ��(&S)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Also_Photo, "ͼ��ͬ��(&I)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Also_ManualSerial, "�ֹ�����ͬ��"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Also_LockSerial, "����/��������"
                    
        Set control = .Add(xtpControlButton, ID_Active_Shuttle, "����(&T)")
        control.BeginGroup = True
        
        .Add xtpControlButton, ID_Active_Cruise, "����(&M)"
        .Add xtpControlButton, ID_Active_Cut, "�ü�"
        .Add xtpControlButton, ID_ACtive_FrameSelectImage, "��ѡͼ��"
        .Add xtpControlButton, ID_Active_Zoom, "����(&Z)"
        .Add xtpControlButton, ID_Active_ReSetAll, "�ָ�����(&A)"
        .Add xtpControlButton, ID_ACtive_Mouse_Value, "���������ʾCTֵ(&S)"
        .Add xtpControlButton, ID_Tool_NothinMouseState, "����������״̬(ESC)"
        
        Set ControlSelect = .Add(xtpControlPopup, ID_Active_AdjustWindow, "����(&W)")
            ControlSelect.CommandBar.Controls.Add xtpControlSplitButtonPopup, ID_Active_AdjustWindow_HandAdjustWindow, "�ֶ�����"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_AdjustWindow_AutoAdjustWindow, "����Ӧ����"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_AdjustWindow_HandAdjustWindow_Custom, "�Զ���(&A)"
            
        Set ControlSelect = .Add(xtpControlPopup, ID_Active_PointingLine, "��λ��(&P)")
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_PointingLine_ALL, "���ж�λ��(&O)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_PointingLine_FirstLast, "��β��λ��(&1)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_PointingLine_Now, "��ǰ��λ��(&2)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_PointingLine_3DLine, "3D��궨λ(&M)"
        ControlSelect.BeginGroup = True
        
        Set ControlSelect = .Add(xtpControlPopup, ID_Active_Eddy, "��ת(&R)")
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Eddy_LeftRight, "���ҷ�ת(&X)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Eddy_TopButton, "��ֱ��ת(&Y)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Eddy_Left90, "����90��"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Eddy_Right90, "����90��"
        ControlSelect.BeginGroup = True
        
        .Add xtpControlButton, ID_Active_ReverseVideo, "����"
        
        Set ControlSelect = .Add(xtpControlPopup, ID_Active_SieveLens, "�˾�(&A)")
            ControlSelect.CommandBar.Controls.Add xtpControlButtonPopup, ID_Active_SieveLens_Model, "�˾�ģ��"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_SieveLens_LancetMinus, "��ǿǿ�ȼ���(&D)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_SieveLens_LancetAdd, "��ǿǿ������(&U)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_SieveLens_FlatnessMinus, "ƽ������"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_SieveLens_FlatnessAdd, "ƽ������"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Sievelens_LeftMoveMinus, "��ǿ���ȼ���(&M)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Sievelens_LeftMoveAdd, "��ǿ������ǿ(&T)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Sievelens_PhotoReset, "ͼ��ԭ(&R)"
        
        
        Set ControlSelect = .Add(xtpControlPopup, ID_Active_Lable, "��ע(&M)")
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_Text, "����(&T)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_Arrowhead, "��ͷ(&P)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_Ellipse, "��Բ(&E)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_Angle, "�Ƕ�(&G)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_Curve, "����(&C)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_Area, "����(&A)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_BeeLine, "ֱ��(&B)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_Rect, "����(&R)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_VasMeasure, "Ѫ����խ����"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_CadioThoracicRatio, "���رȲ���"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_AreaBeeLinePhoto, "����ֱ��ͼ(&H)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_AdjustLine, "У׼(&V)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_ClearLbale, "�����ע(&E)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Active_Lable_DelSelectLable, "ɾ����ע(&D)"
        
    End With
    '''''''''''''''''''''''''''''''''''''''''''����''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set ControlFile = ToolBars.ActiveMenuBar.Controls.Add(xtpControlPopup, ID_Tool, "����(&T)", -1, False)
    
    With ControlFile.CommandBar.Controls
        .Add xtpControlButton, ID_Tool_Movie, "��Ӱ"
        .Add xtpControlButton, ID_Tool_Magnifier, "�Ŵ�(&G)"
        
        Set control = .Add(xtpControlButton, ID_Tool_ArrowyCoronaryReset, "ʸ��״�ؽ�(&V)")
        control.BeginGroup = True
        Set control = .Add(xtpControlButton, ID_Tool_SlopeReconstruction, "б���ؽ�(&S)")
        
        .Add xtpControlButton, ID_Tool_NumberMinusShadow, "���ּ�Ӱ(&D)"
        .Add xtpControlButton, ID_Tool_BogusColour, "α�ʹ۲�(&C)"
        
        Set control = .Add(xtpControlButton, ID_Tool_FilmPrint, "��Ƭ��ӡ(&P)")
        Set ControlSelect = .Add(xtpControlPopup, ID_Tool_Film_AddSeries, "��ӡ����")
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Tool_Film_AddSeries, "��ӡ����"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Tool_Film_AddImage, "��ӡͼ��"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Tool_Film_AddSelected, "��ӡ��ѡͼ"
            Set ControlSelect = ControlSelect.CommandBar.Controls.Add(xtpControlButtonPopup, ID_Tool_Film_AddInterval, "�����ӡ")
            ControlSelect.CommandBar.SetPopupToolBar True
            ControlSelect.CommandBar.Title = "�����ӡ"
            ControlSelect.ToolTipText = "�����ӡ��ǰ����"
        control.BeginGroup = True
        
        .Add xtpControlButton, ID_Tool_PhotoUnite, "ͼ��ƴ��(&I)"
        'ͨ��ʹ��app.logmode���жϵ�ǰ��������Դ����ĵ���״̬������exe�ļ���ִ��״̬��
        'App.LogMode = 0Ϊ����״̬�������Դ����ĵ���״̬��������Ƭѡ��Ĳ˵�����exe�ļ���ִ��״̬������˵�����ʾ��
        If App.LogMode = 0 Then .Add xtpControlButton, ID_Tool_LableTool, "��ע����"
        
        Set control = .Add(xtpControlButton, ID_Tool_LookPhotoOption, "��Ƭѡ��(&O)")
        control.BeginGroup = True
        
        Set ControlSelect = .Add(xtpControlPopup, ID_ToolBar, "������(&B)")
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_ToolBar_Left, "����(&L)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_ToolBar_Right, "����(&R)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_ToolBar_Top, "����(&T)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_ToolBar_Button, "����(&B)"
            
            Set control = ControlSelect.CommandBar.Controls.Add(xtpControlButton, ID_toolBar_16Icon, "16*16ͼ��")
            control.BeginGroup = True
            
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_ToolBar_24Icon, "24*24ͼ��"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_ToolBar_32Icon, "32*32ͼ��"
            
            Set control = ControlSelect.CommandBar.Controls.Add(xtpControlButton, ID_ToolBar_Hide, "����(&H)")
            control.BeginGroup = True
    End With
    ''''''''''''''''''''''''''''''''''''''''''����''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set ControlFile = ToolBars.ActiveMenuBar.Controls.Add(xtpControlPopup, ID_Help, "����(&H)", -1, False)
    With ControlFile.CommandBar.Controls
        .Add xtpControlButton, ID_Help_Help, "����(&H)"
        
        Set ControlSelect = .Add(xtpControlPopup, ID_Help_WebZLSOFT, "WEB�ϵ�����")
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Help_WebZLSOFT_WEB, "������ҳ(&H)"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Help_WebZLSOFT_Mail, "���ͷ���(&K)"
        
        Set control = .Add(xtpControlButton, ID_Help_UpdateDB, "�������ݿ�(&U)")
        control.BeginGroup = True
        
        Set control = .Add(xtpControlButton, ID_Help_About, "����(&A)")
        control.BeginGroup = True
    End With
    
    
    
    
    '���ӿ�ݼ�
    'Ctrl
    ToolBars.KeyBindings.Add FCONTROL, Asc("O"), ID_File_Open
    ToolBars.KeyBindings.Add FCONTROL, Asc("Q"), ID_File_Close
    ToolBars.KeyBindings.Add FCONTROL, Asc("R"), ID_File_SAveASReport
    ToolBars.KeyBindings.Add FCONTROL, Asc("X"), ID_File_Exit
    ToolBars.KeyBindings.Add FCONTROL, Asc("J"), ID_View_Typeset
    ToolBars.KeyBindings.Add FCONTROL, Asc("L"), ID_Active_Also_LockSerial
    ToolBars.KeyBindings.Add FCONTROL, Asc("1"), ID_Active_Select_OneSelect
    ToolBars.KeyBindings.Add FCONTROL, Asc("M"), ID_View_ShowMiniSeries
    ToolBars.KeyBindings.Add FCONTROL, Asc("G"), ID_Tool_Magnifier
'    ToolBars.KeyBindings.Add FCONTROL, Asc("A"), ID_Active_Select_SelectAllSerial
    ToolBars.KeyBindings.Add FCONTROL, Asc("A"), ID_Active_Also_ManualSerial
    
    'Alt
    ToolBars.KeyBindings.Add FALT, Asc("T"), ID_Active_Shuttle
    ToolBars.KeyBindings.Add FALT, Asc("M"), ID_Active_Cruise
    ToolBars.KeyBindings.Add FALT, Asc("J"), ID_ACtive_FrameSelectImage
    ToolBars.KeyBindings.Add FALT, Asc("R"), ID_Active_ReSetAll
    ToolBars.KeyBindings.Add FALT, Asc("B"), ID_Active_ReverseVideo
    ToolBars.KeyBindings.Add FALT, Asc("H"), ID_ToolBar_Hide
    
    
    
    '����������
    Dim ToolBar As CommandBar
    Dim ControlPopup As CommandBarPopup
    
    Set ToolBar = ToolBars.Add("��������", xtpBarBottom)
    ToolBar.SetIconSize IconX, IconY
    
    With ToolBar.Controls
        .Add xtpControlButton, ID_File_SAveASReport, "���汨��ͼ"
        .Add xtpControlButton, ID_File_Open, "��"
        .Add xtpControlButton, ID_Tool_FilmPrint, "��Ƭ���"
        Set ControlSelect = .Add(xtpControlSplitButtonPopup, ID_Tool_Film_AddSeries, "��ӡ����")
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Tool_Film_AddSeries, "��ӡ����"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Tool_Film_AddImage, "��ӡͼ��"
            ControlSelect.CommandBar.Controls.Add xtpControlButton, ID_Tool_Film_AddSelected, "��ӡ��ѡͼ"
            Set ControlSelect = ControlSelect.CommandBar.Controls.Add(xtpControlButtonPopup, ID_Tool_Film_AddInterval, "�����ӡ")
            ControlSelect.CommandBar.SetPopupToolBar True
            ControlSelect.CommandBar.Title = "�����ӡ"
            ControlSelect.ToolTipText = "�����ӡ��ǰ����"
    End With
    
    
    Set ToolBar = ToolBars.Add("ͼ�����", xtpBarBottom)
    ToolBar.SetIconSize IconX, IconY
    With ToolBar.Controls
        .Add xtpControlButton, ID_Active_Eddy_LeftRight, "ˮƽ����"
        .Add xtpControlButton, ID_Active_Eddy_TopButton, "��ֱ����"
        .Add xtpControlButton, ID_Active_Eddy_Left90, "��ת90��"
        .Add xtpControlButton, ID_Active_Eddy_Right90, "��ת90��"
        .Add xtpControlButton, ID_Active_ReverseVideo, "����"
        .Add xtpControlButton, ID_Tool_NumberMinusShadow, "DSA���ּ�Ӱ"
    End With
    
    Set ToolBar = ToolBars.Add("����������", xtpBarBottom)
    ToolBar.SetIconSize IconX, IconY
    With ToolBar.Controls
        .Add xtpControlButton, ID_Tool_NothinMouseState, "���"
        .Add xtpControlButton, ID_ACtive_Mouse_Value, "���������ʾCTֵ"
        .Add xtpControlButton, ID_Active_Lable_Text, "����"
        .Add xtpControlButton, ID_Active_Lable_Arrowhead, "��ͷ"
        .Add xtpControlButton, ID_Active_Lable_Ellipse, "��Բ"
        .Add xtpControlButton, ID_Active_Lable_Angle, "�Ƕ�"
        .Add xtpControlButton, ID_Active_Lable_Curve, "����"
        .Add xtpControlButton, ID_Active_Lable_Area, "����"
        .Add xtpControlButton, ID_Active_Lable_BeeLine, "ֱ��"
        .Add xtpControlButton, ID_Active_Lable_Rect, "����"
        .Add xtpControlButton, ID_Active_Lable_VasMeasure, "Ѫ����խ����"
        .Add xtpControlButton, ID_Active_Lable_CadioThoracicRatio, "���رȲ���"
        .Add xtpControlButton, ID_Active_Lable_ClearLbale, "�����ע"
        .Add xtpControlButton, ID_Active_Lable_AdjustLine, "У׼"
    End With

    Set ToolBar = ToolBars.Add("��ƽ�湤����", xtpBarBottom)
    ToolBar.SetIconSize IconX, IconY
    With ToolBar.Controls
        .Add xtpControlButton, ID_Active_PointingLine_ALL, "��ʾ���ж�λ��"
        .Add xtpControlButton, ID_Active_PointingLine_FirstLast, "��ʾ��β��λ��"
        .Add xtpControlButton, ID_Active_PointingLine_Now, "��ʾ��ǰ��λ��"
        .Add xtpControlButton, ID_Active_PointingLine_3DLine, "��ά���"
        .Add xtpControlButton, ID_Tool_ArrowyCoronaryReset, "ʸ/��״λ�ؽ�"
        .Add xtpControlButton, ID_Tool_SlopeReconstruction, "б���ؽ�"
    End With

    Set ToolBar = ToolBars.Add("�������", xtpBarBottom)
    ToolBar.SetIconSize IconX, IconY
    With ToolBar.Controls
        .Add xtpControlButton, ID_ACtive_FrameSelectImage, "��ѡͼ��"
        .Add xtpControlButton, ID_Active_Also_Photo, "ͼ���ʽͬ��"
        .Add xtpControlButton, ID_Active_Also_Serial, "���м�ͼ��λ��ͬ��"
        .Add xtpControlButton, ID_Active_Also_ManualSerial, "�ֹ�����ͬ��"
        .Add xtpControlButton, ID_Active_Also_LockSerial, "����/��������"
        .Add xtpControlButton, ID_View_ShowMiniSeries, "��ʾ��������ͼ"
        .Add xtpControlButton, ID_View_ViewAllSeries, "ȫ���й�Ƭ"
    End With

    Set ToolBar = ToolBars.Add("ͨ�ù�����", xtpBarBottom)
    ToolBar.Closeable = False
    ToolBar.SetIconSize IconX, IconY
    With ToolBar.Controls
        .Add xtpControlButton, ID_Tool_Magnifier, "�Ŵ�"
        Set ControlPopup = .Add(xtpControlSplitButtonPopup, ID_Active_AdjustWindow_HandAdjustWindow, "�ֿص���")

        .Add xtpControlButton, ID_Active_Cruise, "����"

        Set ControlPopup = .Add(xtpControlSplitButtonPopup, ID_Active_Zoom, "����")
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_AutoShow, "����Ӧ"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_50%, "50%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_100%, "100%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_showScale_150%, "150%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_200%, "200%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_showScale_250%, "250%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_showScale_300%, "300%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_showScale_400%, "400%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_Custom, "�Զ���(&A)"

        Set ControlPopup = .Add(xtpControlSplitButtonPopup, ID_Active_Shuttle, "����")
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_PhotoNumber, "ͼ���"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_BedASC, "��λ����"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_BedDESC, "��λ����"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_CollectionTime, "�ɼ�ʱ��"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_PhotoTime, "ͼ��ʱ��"

        .Add xtpControlButton, ID_Tool_Movie, "��Ӱ����"
        .Add xtpControlButton, ID_Active_Select_SelectAllSerial, "ѡ����������"
        .Add xtpControlButton, ID_Acitve_Select_SelectAllPhoto, "ѡ�����������е�ͼ��"
        .Add xtpControlButton, ID_View_UpSeries, "��һ����"
        .Add xtpControlButton, ID_View_DownSeries, "��һ����"
        .Add xtpControlButton, ID_View_Typeset, "�������"
        .Add xtpControlButton, ID_View_FullScreen, "ȫ����ʾ"
        Set control = .Add(xtpControlButton, ID_View_PropertyShow, "��/��������Ϣ")
        control.Checked = True
        .Add xtpControlButton, ID_Active_ReSetAll, "ȫ���ָ�"
        .Add xtpControlButton, ID_View_OneBrowse, "���/�۲�ģʽ"
    End With
    
    Set ToolBar = ToolBars.Add("ͼ����ǿ", xtpBarBottom)
    ToolBar.SetIconSize IconX, IconY
    With ToolBar.Controls
        Set ControlPopup = .Add(xtpControlSplitButtonPopup, ID_Active_SieveLens_Model, "�˾�ģ��")
        .Add xtpControlButton, ID_Active_SieveLens_LancetMinus, "��Ե��ǿǿ�ȼ���"
        .Add xtpControlButton, ID_Active_SieveLens_LancetAdd, "��Ե��ǿǿ������"
        .Add xtpControlButton, ID_Active_Sievelens_LeftMoveMinus, "��Ե��ǿ���ȼ���"
        .Add xtpControlButton, ID_Active_Sievelens_LeftMoveAdd, "��Ե��ǿ��������"
        .Add xtpControlButton, ID_Active_SieveLens_FlatnessMinus, "ƽ������"
        .Add xtpControlButton, ID_Active_SieveLens_FlatnessAdd, "ƽ������"
        .Add xtpControlButton, ID_Active_Sievelens_PhotoReset, "ͼ��ԭ"
        .Add xtpControlButton, ID_Tool_BogusColour, "α��"
    End With
    
    ToolBars.EnableCustomization True
    
End Sub

Public Function StatusBarTip(control As CommandBarControl) As String
'------------------------------------------------
'���ܣ���״̬����ʾ�˵��ļ򵥰���
'������Control--��ʾ�����Ĳ˵��ؼ�
'���أ�������Ϣ
'�ϼ���������̣�frmViewer.ComToolBar_ControlSelected
'�¼���������̣���
'���õ��ⲿ��������
'�����ˣ�����
'------------------------------------------------
    If control Is Nothing Then
        StatusBarTip = ""
        Exit Function
    End If
    Select Case control.Id
        ''''''''''''''''''''''''''�ļ��˵�'''''''''''''''''''''''''''''''''''
        Case ID_File_Open                                                               '���ļ�
            StatusBarTip = "���µ�ͼ���ļ����й۲�"
            
        Case ID_File_Close                                                              '�ر�����
            StatusBarTip = "�رյ�ǰ����ͼ�񣬹رպ����ͨ���Ű��ٴε�������ͼ��"
            
        Case ID_File_DelAllPhoto                                                        'ɾ������ͼ��
            StatusBarTip = "ɾ������������ͼ��"
            
        Case ID_File_DelReport                                                          'ɾ������ͼ��
            StatusBarTip = "ɾ������ͼ��"
            
        Case ID_File_SaveFile                                                           '�����ļ�
            StatusBarTip = "�����ļ�"
            
        Case ID_File_SaveASFile                                                         '����ļ�
            StatusBarTip = "����ǰѡ�е�ͼ�����Ϊ�ļ�"
            
        Case ID_File_SaveToCD                                                           '����CD
            StatusBarTip = "����ǰѡ�е�ͼ�󱣴浽CD������"
            
        Case ID_File_SAveASReport                                                       '���汨��ͼ��
            StatusBarTip = "����ǰѡ�е�ͼ�󱣴�Ϊ����ͼ��"
            
        Case ID_File_Send_GetHost                                                       '��������
            StatusBarTip = "���͵���������"
            
        Case ID_File_Send_OutPowerPoint                                                 '�����PowerPoint
            StatusBarTip = "�����PowerPoint"
            
        Case ID_File_OpenDicomDir                                                       '��DICOMDIR
            StatusBarTip = "��DICOMDIR�е�ͼ��"
            
        Case ID_File_PhotoProperty                                                      'ͼ������
            StatusBarTip = "�鿴������ͼ���еĲ��ˡ���顢���к�ͼ��������Ϣ"
            
        Case ID_File_Exit                                                               '�˳�
            StatusBarTip = "�˳�"
            
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''��ͼ''''''''''''''''''''''''''''''''''''''''
        Case ID_View_UpSeries                                                           '��һ����
            StatusBarTip = "�л�����һ������"
            
        Case ID_View_DownSeries                                                         '��һ����
            StatusBarTip = "�л�����һ������"
            
        Case ID_View_Typeset                                                            '���氲��
            StatusBarTip = "���������м�������ͼ�����Ļ��ʾ���в���"
            
        Case ID_View_OneBrowse                                                          '����۲�ģʽ
            StatusBarTip = "ʹ�����ģʽ���ǹ۲�ģʽ�쿴ͼ��"
            
        Case ID_View_PropertyShow                                                       'ͼ���ϲ�����Ϣ��ʾ
            StatusBarTip = "��ʾ�����ز��˻�����Ϣ"
            
        Case ID_View_LableShow                                                          '��ע��ʾ
            StatusBarTip = "��ʾ/���ر�ע��Ϣ"
            
        Case ID_View_ShowMiniSeries                                                     '��ʾ��������ͼ
            StatusBarTip = "��ʾ/������������ͼ"
            
        Case ID_View_PhotoSerial_PhotoNumber                                            'ͼ��˳��_ͼ���
            StatusBarTip = "��ͼ���˳�����"
            
        Case ID_View_PhotoSerial_BedASC                                                 '��λ����
            StatusBarTip = "�Դ�λ����˳�����"
            
        Case ID_View_PhotoSerial_BedDESC                                                '��λ����
            StatusBarTip = "�Դ�λ����˳�����"
            
        Case ID_View_PhotoSerial_CollectionTime                                         '�ɼ�ʱ��
            StatusBarTip = "�Բɼ�ʱ��˳�����"
            
        Case ID_View_PhotoSerial_PhotoTime                                              'ͼ��ʱ��
            StatusBarTip = "��ͼ��ʱ��˳�����"
            
        Case ID_View_ShowScale_AutoShow                                                 '����Ӧ
            StatusBarTip = "����Ļ���ʵĴ�С��ʾͼ��"
            
        Case ID_View_ShowScale_50%                                                      '50%
            StatusBarTip = "��50%��Сͼ����ʾ"
            
        Case ID_View_ShowScale_100%                                                     '100%
            StatusBarTip = "��ͼ��������С��ʾ"
            
        Case ID_View_showScale_150%                                                     '150%
            StatusBarTip = "��150%ͼ���С��ʾ"
            
        Case ID_View_ShowScale_200%                                                     '200%
            StatusBarTip = "��200%��Сͼ����ʾ"
            
        Case ID_View_showScale_250%                                                     '250%
            StatusBarTip = "��250%��Сͼ����ʾ"
            
        Case ID_View_showScale_300%                                                     '300%
            StatusBarTip = "��300%��Сͼ����ʾ"
            
        Case ID_View_showScale_400%                                                     '400%
            StatusBarTip = "��400%��Сͼ����ʾ"
            
        Case ID_View_ShowScale_Custom                                                   '�Զ���
            StatusBarTip = "�Զ���ͼ����ʾ��С"
            
        Case ID_View_FullScreen                                                         'ȫ����ʾ
            StatusBarTip = "ȫ��Ļ��ʾͼ����й۲�"
            
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''����''''''''''''''''''''''''''''''''''''
        Case ID_Active_Select_OneSelect                                                 '����ѡ��
            StatusBarTip = "ѡ��ǰͼ��"
        
        Case ID_Active_Select_SelectAllSerial                                           'ѡ����������
            StatusBarTip = "ѡ�д򿪵��������У��Ա�Ϊ����������׼��"
            
        Case ID_Acitve_Select_SelectAllPhoto                                            'ѡ������ͼ��
            StatusBarTip = "ѡ���ȡ���Ե�ǰ���е�����ͼ���ѡ���־"
            
        Case ID_Active_Also_Serial                                                      '����ͬ��
            StatusBarTip = "�Ե�ǰ�򿪵�������ͼ��λ�õ�ͬ������"
            
        Case ID_Active_Also_ManualSerial                                                '�ֹ�����ͬ��
            StatusBarTip = "�ֹ��Ե�ǰ�򿪵�������ͼ��λ�õ�ͬ������"
        
        Case ID_Active_Also_LockSerial                                                  '��������
            StatusBarTip = "������������У����ø������Ƿ�μ��ֶ�����ͬ�����롰����Ctrl+������������������ͬ"
            
        Case ID_Active_Also_Photo                                                       'ͼ��ͬ��
            StatusBarTip = "�Ե�ǰ�����ڵ�ͼ����������ͬ��"
            
        Case ID_Active_Shuttle                                                          '����
            StatusBarTip = "�ڹ۲�����ͨ�������ƶ�ֱ���л�ͼ��" & funcGetShiftStr(cMouseUsage("101").lngShift)
            
        Case ID_Active_Cruise                                                           '����
            StatusBarTip = "�ڹ۲������ƶ�ͼ���λ�ã��Ա��ڸ��õع۲�" & funcGetShiftStr(cMouseUsage("103").lngShift)

        Case ID_Active_Cut                                                              '�ü�
            StatusBarTip = "��ͼ����вü�" & funcGetShiftStr(cMouseUsage("201").lngShift)
            
        Case ID_ACtive_FrameSelectImage                                                 '��ѡ
            StatusBarTip = "�϶����ο�ѡ��ֲ�ͼ��" & funcGetShiftStr(cMouseUsage("201").lngShift)
                        
        Case ID_Active_Zoom                                                             '����
            StatusBarTip = "�ڹ۲�������С��Ŵ�ͼ��" & funcGetShiftStr(cMouseUsage("104").lngShift)
            
        Case ID_Active_ReSetAll                                                         '�ָ�����
            StatusBarTip = "ȡ�����������εȲ������ָ�ͼ��ԭʼ״̬"
            
        Case ID_Active_AdjustWindow_HandAdjustWindow                                    '�ֶ�����
            StatusBarTip = "�����ֶ����Ƶ�ͼ�󴰿�λ����ģʽ" & funcGetShiftStr(cMouseUsage("102").lngShift)
            
        Case ID_Active_AdjustWindow_AutoAdjustWindow                                    '����Ӧ����
            StatusBarTip = "�������ʵ���ģʽ��ͨ��ѡ��һ�����򣬽�������Ӧ����" & funcGetShiftStr(cMouseUsage("105").lngShift)
            
        Case ID_Active_AdjustWidnow_CustomAdjustWindow                                  '�Զ������
            StatusBarTip = "������ʵĴ���λ���е���"
            
        Case ID_Active_PointingLine_ALL                                                 '���ж�λ��
            StatusBarTip = "��ʾ����ͼ������ж�λ��"
            
        Case ID_Active_PointingLine_FirstLast                                           '��λ��λ��
            StatusBarTip = "��ʾ����ͼ�����β��λ��"
            
        Case ID_Active_PointingLine_Now                                                 '��ǰ��λ��
            StatusBarTip = "��ʾ��ǰͼ���Ӧ�Ķ�λ��"
            
        Case ID_Active_PointingLine_3DLine                                              '3D���
            StatusBarTip = "��ʾ��ǰͼ�����ָ�����ά��Ӧλ�õ�" & funcGetShiftStr(cMouseUsage("106").lngShift)
            
        Case ID_Active_Eddy_LeftRight                                                   '������ת
            StatusBarTip = "��ͼ��������ҷ�ת����й۲�"
            
        Case ID_Active_Eddy_TopButton                                                   '��ֱ��ת
            StatusBarTip = "��ͼ����д�ֱ��ת����й۲�"
            
        Case ID_Active_Eddy_Left90                                                      '����90
            StatusBarTip = "��ͼ���������90�����й۲�"
            
        Case ID_Active_Eddy_Right90                                                     '����90
            StatusBarTip = "��ͼ���������90�����й۲�"
            
        Case ID_Active_ReverseVideo                                                     '����
            StatusBarTip = "�Ե�ǰͼ����ͬ��������ͼ����кڰ׷�ת�۲�"
        
        Case ID_Active_SieveLens_Model                                                  '�˾�ģ��
            StatusBarTip = "Ӧ��Ԥ�����úõ��˾�ģ��"
                 
        Case ID_Active_SieveLens_LancetMinus                                            '�񻯼���
            StatusBarTip = "����ͼ����ǿǿ��"
            
        Case ID_Active_SieveLens_LancetAdd                                              '������
            StatusBarTip = "����ͼ����ǿǿ��"
            
        Case ID_Active_SieveLens_FlatnessMinus                                          'ƽ������
            StatusBarTip = "����ͼ��ƽ��Ч��"
            
        Case ID_Active_SieveLens_FlatnessAdd                                            'ƽ������
            StatusBarTip = "����ͼ��ƽ��Ч��"
            
        Case ID_Active_Sievelens_LeftMoveMinus                                          '��������
            StatusBarTip = "����ͼ����ǿ����"
            
        Case ID_Active_Sievelens_LeftMoveAdd                                            '��������
            StatusBarTip = "����ͼ����ǿ����"
            
        Case ID_Active_Sievelens_PhotoReset                                             'ͼ��ԭ
            StatusBarTip = "ȡ���˾���ǿ�۲�Ч�����ָ�ͼ��ԭʼ״̬"
            
        Case ID_Active_Lable_Text                                                       '����
            StatusBarTip = "��ӡ����֡����͵ı�ע" & funcGetShiftStr(cMouseUsage("8").lngShift)
            
        Case ID_Active_Lable_Arrowhead                                                  '��ͷ
            StatusBarTip = "��ӡ���ͷ�����͵ı�ע" & funcGetShiftStr(cMouseUsage("4").lngShift)
         
        Case ID_Active_Lable_Ellipse                                                    '��Բ
            StatusBarTip = "��ӡ���Բ�����͵ı�ע" & funcGetShiftStr(cMouseUsage("3").lngShift)
        
        Case ID_Active_Lable_Angle                                                      '�Ƕ�
            StatusBarTip = "��ӡ��Ƕȡ����͵ı�ע" & funcGetShiftStr(cMouseUsage("7").lngShift)
        
        Case ID_Active_Lable_Curve                                                      '����
            StatusBarTip = "��ӡ����ߡ����͵ı�ע" & funcGetShiftStr(cMouseUsage("6").lngShift)
        
        Case ID_Active_Lable_Area                                                       '����
            StatusBarTip = "�������ķ��������ʽ�ı�ע" & funcGetShiftStr(cMouseUsage("5").lngShift)
        
        Case ID_Active_Lable_BeeLine                                                    'ֱ��
            StatusBarTip = "��ӡ�ֱ�ߡ����͵ı�ע" & funcGetShiftStr(cMouseUsage("1").lngShift)
        
        Case ID_Active_Lable_Rect                                                       '����
            StatusBarTip = "��ӡ����Ρ����͵ı�ע" & funcGetShiftStr(cMouseUsage("2").lngShift)
        
        Case ID_Active_Lable_AreaBeeLinePhoto                                           '����ֱ��ͼ
            StatusBarTip = "��ѡ�еľ�����Բ������ĻҶ����ֱ��ͼ�Ա�"
            
        Case ID_Active_Lable_VasMeasure                                                 'Ѫ����խ����
            StatusBarTip = "��ͼ�����Ѫ����խ����" & funcGetShiftStr(cMouseUsage("1").lngShift)
            
        Case ID_Active_Lable_CadioThoracicRatio                                         '���رȲ���
            StatusBarTip = "������������������в���" & funcGetShiftStr(cMouseUsage("1").lngShift)
            
        Case ID_Active_Lable_AdjustLine                                                 'У׼
            StatusBarTip = "��ѡ�е�ֱ�߽��г��ȵ��ֹ�У׼���޸ı�ע"
            
        Case ID_Active_Lable_ClearLbale                                                 '������б�ע
            StatusBarTip = "������еı�ע"
            
        Case ID_Active_Lable_DelSelectLable                                             'ɾ����ǰ��ע
            StatusBarTip = "ɾ����ǰѡ�еı�ע�����"
            
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''���߲˵�''''''''''''''''''''''''''''''''''''
        Case ID_Tool_Movie                                                              '��Ӱ
            StatusBarTip = "��ѭ�����Ӱڷ�ʽ�������Ŷ�֡����ͼ��"
            
        Case ID_Tool_Magnifier                                                          '�Ŵ�
            StatusBarTip = "ʹ�÷Ŵ󾵣���ͼ����оֲ��ķŴ���С�۲�"
            
        Case ID_Tool_ArrowyCoronaryReset                                                'ʸ��״�ؽ�
            StatusBarTip = "��ͼ�����ʸ��״λ��ά�ؽ��۲�"
            
        Case ID_Tool_SlopeReconstruction                                                'б���ؽ�
            StatusBarTip = "��ͼ�����б���ά�ؽ��۲�"
            
        Case ID_Tool_NumberMinusShadow                                                  '���ּ�Ӱ
            StatusBarTip = "��ͼ��������ּ�Ӱ�Ĺ۲�"
            
        Case ID_Tool_BogusColour                                                        'α��
            StatusBarTip = "���ò���α��ɫ��ʽ�۲�ͼ��"
            
        Case ID_Tool_FilmPrint                                                          '��Ƭ��ӡ
            StatusBarTip = "�����ѡ�����к�ͼ��Ľ�Ƭ��ӡ����"
            
        Case ID_Tool_PhotoUnite                                                         'ͼ��ƴ��
            StatusBarTip = "����ͬ���ͼ��ͼ��ƴ��"
            
        Case ID_Tool_LableTool                                                          '��ע����
            StatusBarTip = "��ע����"
            
        Case ID_Tool_LookPhotoOption                                                    '��Ƭѡ��
            StatusBarTip = "��Ƭ����վ�Ļ�������"
            
        Case ID_ToolBar_Left                                                            '����������
            StatusBarTip = "����������ڷ�"
            
        Case ID_ToolBar_Right                                                           '����������
            StatusBarTip = "���������Ұڷ�"
        
        Case ID_ToolBar_Top                                                             '����������
            StatusBarTip = "���������ϰڷ�"
            
        Case ID_ToolBar_Button                                                          '����������
            StatusBarTip = "���������°ڷ�"
            
        Case ID_toolBar_16Icon                                                          '������ͼ��16*16��ʾ
            StatusBarTip = "��������16*16ͼ����ʾ"
            
        Case ID_ToolBar_24Icon                                                          '������ͼ��24*24��ʾ
            StatusBarTip = "��������24*24ͼ����ʾ"
        
        Case ID_ToolBar_32Icon                                                          '������ͼ��32*32��ʾ
            StatusBarTip = "��������32*32ͼ����ʾ"
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''����'''''''''''''''''''''''''''''''''''''
        Case ID_Help_Help                                                               '����
            StatusBarTip = "��Ƭվ����"
        
        Case ID_Help_WebZLSOFT_WEB                                                      '������ҳ
            StatusBarTip = "��������ҳ"
        
        Case ID_Help_WebZLSOFT_Mail                                                     '���ͷ���
            StatusBarTip = "���ͷ���"
            
        Case ID_Help_About                                                              '����
            StatusBarTip = "����"
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End Select
End Function

Private Function funcGetShiftStr(lngShift As Long) As String
'------------------------------------------------
'���ܣ�ͨ������shift״̬��ֵ�����������ַ�����ʾ��shift״̬��
'������lngShift--��ʾshifit״̬����ֵ
'���أ����ַ�����ʾ��shift״̬��
'�ϼ���������̣�mdlTools.StatusBarTip
'�¼���������̣���
'���õ��ⲿ��������
'�����ˣ��ƽ�
'------------------------------------------------
    'shift ���÷���shift,ctrl,alt �ֱ���1��2��4��ʾ��ͨ���ۼ�ʵ��
    funcGetShiftStr = ""
    If lngShift - 4 >= 0 Then
        funcGetShiftStr = " Alt "
        lngShift = lngShift - 4
    End If
    If lngShift - 2 >= 0 Then
        funcGetShiftStr = funcGetShiftStr & " Ctrl "
        lngShift = lngShift - 2
    End If
    If lngShift = 1 Then
        funcGetShiftStr = funcGetShiftStr & " Shift "
    End If
End Function

Public Sub ArrayToolBar(ToolBars As Object, frmTop As Long, frmLeft As Long, frmHeight As Long, frmWidth As Long)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����               ���°�һ����˳��ڷŹ�����λ��
    '����
    '    ToolBars       �������ؼ�
    '    frmTop         ��ǰ����Top
    '    frmLeft        ��ǰ����Left
    '    frmWidth       ��ǰ����Width
    '    frmHieht       ��ǰ����Height
    '����               ��
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim MenuToolBar As CommandBar                           '�˵�
    Dim MainToolBar As CommandBar                           '��������
    Dim PhotoToolBar As CommandBar                          'ͼ��������
    Dim ScaleToolBar As CommandBar                          '����������
    Dim PlaneToolBar As CommandBar                          'ƽ�湤����
    Dim ObjectToolBar As CommandBar                         '���󹤾���
    Dim CommToolBar As CommandBar                           'ͨ�ù�����
    Dim PhotoStrongToolBar As CommandBar                    'ͼ����ǿ������
    Dim NowPosiTion As Integer                              '��ǰ������λ��
    Const intState As Integer = 360                         '״̬����
    Dim ToolBarLeft As Long                                 '����������ߵ�λ��
    Dim ToolBarTop As Long                                  '���������ϱߵ�λ��
    
    Dim OldTop As Long, OldLeft As Long, OldRight As Long, OldBottom As Long                    '��һ������λ��
    Dim NowTop As Long, NowLeft As Long, NowRight As Long, NowBottom As Long                    '��ǰ������λ��
    
    
    
    blfrmRefresh = False
    
    '��������
    Set MainToolBar = ToolBars.Item(2)
    NowPosiTion = MainToolBar.Position
    Set MenuToolBar = ToolBars.Item(1)
    MenuToolBar.GetWindowRect NowLeft, NowTop, NowRight, NowBottom
    ToolBarLeft = NowLeft
    ToolBarTop = NowTop
    Select Case MainToolBar.Position
        Case 0
            ToolBars.DockToolBar MainToolBar, ToolBarLeft, NowBottom, NowPosiTion
        Case 1
            ToolBars.DockToolBar MainToolBar, ToolBarLeft, frmHeight + frmTop, NowPosiTion
        Case 2
            ToolBars.DockToolBar MainToolBar, ToolBarLeft, ToolBarTop, NowPosiTion
        Case 3
            ToolBars.DockToolBar MainToolBar, frmLeft + frmWidth, ToolBarTop, NowPosiTion
    End Select
    
    ToolBars.RecalcLayout
    
    '�������������
    MainToolBar.GetWindowRect OldLeft, OldTop, OldRight, OldBottom
    
    
    Set ObjectToolBar = ToolBars.Item(6)
    ObjectToolBar.GetWindowRect NowLeft, NowTop, NowRight, NowBottom
    Select Case ObjectToolBar.Position
        Case 0
            If frmWidth - (OldRight - frmLeft) > NowRight - NowLeft Then
                ToolBars.DockToolBar ObjectToolBar, OldRight, (OldBottom + OldTop) / 2, NowPosiTion
            Else
                ToolBars.DockToolBar ObjectToolBar, ToolBarLeft, OldBottom, NowPosiTion
            End If
        Case 1
            If frmWidth - (OldRight - frmLeft) > NowRight - NowLeft Then
                ToolBars.DockToolBar ObjectToolBar, OldRight, (OldBottom + OldTop) / 2, NowPosiTion
            Else
                ToolBars.DockToolBar ObjectToolBar, ToolBarLeft, OldTop, NowPosiTion
            End If
        Case 2
            If frmHeight - (OldBottom - frmTop) - intState > NowBottom - NowTop Then
                ToolBars.DockToolBar ObjectToolBar, (OldLeft + OldRight) / 2, OldBottom, NowPosiTion
            Else
                ToolBars.DockToolBar ObjectToolBar, OldRight, ToolBarTop, NowPosiTion
            End If
        Case 3
            If frmHeight - (OldBottom - frmTop) - intState > NowBottom - NowTop Then
                ToolBars.DockToolBar ObjectToolBar, (OldLeft + OldRight) / 2, OldBottom, NowPosiTion
            Else
                ToolBars.DockToolBar ObjectToolBar, OldLeft - (OldRight - OldLeft), ToolBarTop, NowPosiTion
            End If
    End Select
    ToolBars.RecalcLayout
    
    'ͼ��������
    ObjectToolBar.GetWindowRect OldLeft, OldTop, OldRight, OldBottom
    
    
    
    Set PhotoToolBar = ToolBars.Item(3)
    PhotoToolBar.GetWindowRect NowLeft, NowTop, NowRight, NowBottom
    Select Case PhotoToolBar.Position
        Case 0
            If frmWidth - (OldRight - frmLeft) > NowRight - NowLeft Then
                ToolBars.DockToolBar PhotoToolBar, OldRight, (OldBottom + OldTop) / 2, NowPosiTion
            Else
                ToolBars.DockToolBar PhotoToolBar, ToolBarLeft, OldBottom, NowPosiTion
            End If
        Case 1
            If frmWidth - (OldRight - frmLeft) > NowRight - NowLeft Then
                ToolBars.DockToolBar PhotoToolBar, OldRight, (OldBottom + OldTop) / 2, NowPosiTion
            Else
                ToolBars.DockToolBar PhotoToolBar, ToolBarLeft, OldTop, NowPosiTion
            End If
        Case 2
            If frmHeight - (OldBottom - frmTop) - intState > NowBottom - NowTop Then
                ToolBars.DockToolBar PhotoToolBar, (OldLeft + OldRight) / 2, OldBottom, NowPosiTion
            Else
                ToolBars.DockToolBar PhotoToolBar, OldRight, ToolBarTop, NowPosiTion
            End If
        Case 3
            If frmHeight - (OldBottom - frmTop) - intState > NowBottom - NowTop Then
                ToolBars.DockToolBar PhotoToolBar, (OldLeft + OldRight) / 2, OldBottom, NowPosiTion
            Else
                ToolBars.DockToolBar PhotoToolBar, OldLeft - (OldRight - OldLeft), ToolBarTop, NowPosiTion
            End If
    End Select
    ToolBars.RecalcLayout

    '����������
    PhotoToolBar.GetWindowRect OldLeft, OldTop, OldRight, OldBottom
    
    
    
    Set ScaleToolBar = ToolBars.Item(4)
    ScaleToolBar.GetWindowRect NowLeft, NowTop, NowRight, NowBottom
    Select Case ScaleToolBar.Position
        Case 0
            If frmWidth - (OldRight - frmLeft) > NowRight - NowLeft Then
                ToolBars.DockToolBar ScaleToolBar, OldRight, (OldBottom + OldTop) / 2, NowPosiTion
            Else
                ToolBars.DockToolBar ScaleToolBar, ToolBarLeft, OldBottom, NowPosiTion
            End If
        Case 1
            If frmWidth - (OldRight - frmLeft) > NowRight - NowLeft Then
                ToolBars.DockToolBar ScaleToolBar, OldRight, (OldBottom + OldTop) / 2, NowPosiTion
            Else
                ToolBars.DockToolBar ScaleToolBar, ToolBarLeft, OldTop, NowPosiTion
            End If
        Case 2
            If frmHeight - (OldBottom - frmTop) - intState > NowBottom - NowTop Then
                ToolBars.DockToolBar ScaleToolBar, (OldLeft + OldRight) / 2, OldBottom, NowPosiTion
            Else
                ToolBars.DockToolBar ScaleToolBar, OldRight, ToolBarTop, NowPosiTion
            End If
        Case 3
            If frmHeight - (OldBottom - frmTop) - intState > NowBottom - NowTop Then
                ToolBars.DockToolBar ScaleToolBar, (OldLeft + OldRight) / 2, OldBottom, NowPosiTion
            Else
                ToolBars.DockToolBar ScaleToolBar, OldLeft - (OldRight - OldLeft), ToolBarTop, NowPosiTion
            End If
    End Select
    ToolBars.RecalcLayout

    'ƽ�湤����
    ScaleToolBar.GetWindowRect OldLeft, OldTop, OldRight, OldBottom
    
    
    Set PlaneToolBar = ToolBars.Item(5)
    PlaneToolBar.GetWindowRect NowLeft, NowTop, NowRight, NowBottom
    Select Case ScaleToolBar.Position
        Case 0
            If frmWidth - (OldRight - frmLeft) > NowRight - NowLeft Then
                ToolBars.DockToolBar PlaneToolBar, OldRight, (OldBottom + OldTop) / 2, NowPosiTion
            Else
                ToolBars.DockToolBar PlaneToolBar, ToolBarLeft, OldBottom, NowPosiTion
            End If
        Case 1
            If frmWidth - (OldRight - frmLeft) > NowRight - NowLeft Then
                ToolBars.DockToolBar PlaneToolBar, OldRight, (OldBottom + OldTop) / 2, NowPosiTion
            Else
                ToolBars.DockToolBar PlaneToolBar, ToolBarLeft, OldTop, NowPosiTion
            End If
        Case 2
            If frmHeight - (OldBottom - frmTop) - intState > NowBottom - NowTop Then
                ToolBars.DockToolBar PlaneToolBar, (OldLeft + OldRight) / 2, OldBottom, NowPosiTion
            Else
                ToolBars.DockToolBar PlaneToolBar, OldRight, ToolBarTop, NowPosiTion
            End If
        Case 3
            If frmHeight - (OldBottom - frmTop) - intState > NowBottom - NowTop Then
                ToolBars.DockToolBar PlaneToolBar, (OldLeft + OldRight) / 2, OldBottom, NowPosiTion
            Else
                ToolBars.DockToolBar PlaneToolBar, OldLeft - (OldRight - OldLeft), ToolBarTop, NowPosiTion
            End If
    End Select
    ToolBars.RecalcLayout

    'ͨ�ù�����
    PlaneToolBar.GetWindowRect OldLeft, OldTop, OldRight, OldBottom
    
    
    
    Set CommToolBar = ToolBars.Item(7)
    CommToolBar.GetWindowRect NowLeft, NowTop, NowRight, NowBottom
    Select Case ScaleToolBar.Position
        Case 0
            If frmWidth - (OldRight - frmLeft) > NowRight - NowLeft Then
                ToolBars.DockToolBar CommToolBar, OldRight, (OldBottom + OldTop) / 2, NowPosiTion
            Else
                ToolBars.DockToolBar CommToolBar, ToolBarLeft, OldBottom, NowPosiTion
            End If
        Case 1
            If frmWidth - (OldRight - frmLeft) > NowRight - NowLeft Then
                ToolBars.DockToolBar CommToolBar, OldRight, (OldBottom + OldTop) / 2, NowPosiTion
            Else
                ToolBars.DockToolBar CommToolBar, ToolBarLeft, OldTop, NowPosiTion
            End If
        Case 2
            If frmHeight - (OldBottom - frmTop) - intState > NowBottom - NowTop Then
                ToolBars.DockToolBar CommToolBar, (OldLeft + OldRight) / 2, OldBottom, NowPosiTion
            Else
                ToolBars.DockToolBar CommToolBar, OldRight, ToolBarTop, NowPosiTion
            End If
        Case 3
            If frmHeight - (OldBottom - frmTop) - intState > NowBottom - NowTop Then
                ToolBars.DockToolBar CommToolBar, (OldLeft + OldRight) / 2, OldBottom, NowPosiTion
            Else
                ToolBars.DockToolBar CommToolBar, OldLeft - (OldRight - OldLeft), ToolBarTop, NowPosiTion
            End If
    End Select
    
    ToolBars.RecalcLayout
    
    'ͨ�ù�����
    CommToolBar.GetWindowRect OldLeft, OldTop, OldRight, OldBottom
    
    
    
    Set PhotoStrongToolBar = ToolBars.Item(8)
    PhotoStrongToolBar.GetWindowRect NowLeft, NowTop, NowRight, NowBottom
    Select Case ScaleToolBar.Position
        Case 0
            If frmWidth - (OldRight - frmLeft) > NowRight - NowLeft Then
                ToolBars.DockToolBar PhotoStrongToolBar, OldRight, (OldBottom + OldTop) / 2, NowPosiTion
            Else
                ToolBars.DockToolBar PhotoStrongToolBar, ToolBarLeft, OldBottom, NowPosiTion
            End If
        Case 1
            If frmWidth - (OldRight - frmLeft) > NowRight - NowLeft Then
                ToolBars.DockToolBar PhotoStrongToolBar, OldRight, (OldBottom + OldTop) / 2, NowPosiTion
            Else
                ToolBars.DockToolBar PhotoStrongToolBar, ToolBarLeft, OldTop, NowPosiTion
            End If
        Case 2
            If frmHeight - (OldBottom - frmTop) - intState > NowBottom - NowTop Then
                ToolBars.DockToolBar PhotoStrongToolBar, (OldLeft + OldRight) / 2, OldBottom, NowPosiTion
            Else
                ToolBars.DockToolBar PhotoStrongToolBar, OldRight, ToolBarTop, NowPosiTion
            End If
        Case 3
            If frmHeight - (OldBottom - frmTop) - intState > NowBottom - NowTop Then
                ToolBars.DockToolBar PhotoStrongToolBar, (OldLeft + OldRight) / 2, OldBottom, NowPosiTion
            Else
                ToolBars.DockToolBar PhotoStrongToolBar, OldLeft - (OldRight - OldLeft), ToolBarTop, NowPosiTion
            End If
    End Select
    
    blfrmRefresh = True
    ToolBars.RecalcLayout
End Sub


Sub ReplaceToolBarIcon(ObjToolBar As Object, imgList As ImageList, IconX As Integer, IconY As Integer)
    '------------------------------------------------
    '���ܣ�                                  �滻��ǰͼ��Ϊ16,24,32
    '������
    '           objToolbar                   ����������
    '           imglist                      ͼ��������
    '           IconX                        ����ͼ��X��С
    '           IconY                        ����ͼ��Y��С
    '���أ�                                  ��
    '�ϼ���������̣�                        ComToolBar_Execute
    '�¼���������̣�                        ��
    '���õ��ⲿ������                        ��
    '�����ˣ�                                ���� 2005-6-29
    '------------------------------------------------
    Dim i As Integer
    For i = 2 To ObjToolBar.Count
        ObjToolBar.Item(i).SetIconSize IconX, IconY
    Next
    ObjToolBar.AddImageList imgList
End Sub

Sub PutToolbar(ObjToolBar As Object, Position As Integer)
    '------------------------------------------------
    '���ܣ�                                  �ڷŹ�������Top,Left,Right,Bottom
    '������
    '           objToolbar                   ����������
    '           position                     �ڷ�λ�� 0=Top,1=Bottom 2=Left 3=Right
    '���أ�                                  ��
    '�ϼ���������̣�                        ComToolBar_Execute
    '�¼���������̣�                        ��
    '���õ��ⲿ������                        ��
    '�����ˣ�                                ���� 2005-6-29
    '------------------------------------------------
    Dim i As Integer
    For i = 2 To ObjToolBar.Count
        ObjToolBar.Item(i).Position = Position
    Next
End Sub

Public Sub WriteLog(ByVal ErrorType As Integer, ErrorNum As Long, ErrorDesc As String)
    Dim strSQL As String
    
    On Error GoTo errh
    
    If blLocalRun = False Then Exit Sub
    If cnAccess.State = adStateClosed Then Exit Sub
    
    strSQL = "Insert Into ������־(����ʱ��,��������,�����,������Ϣ) " & _
        "Values(cDate('" & Date & " " & Time() & "')," & ErrorType & "," & ErrorNum & ",'" & Replace(ErrorDesc, "'", "''") & "')"
    cnAccess.Execute strSQL
    
    Exit Sub
errh:
    MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
End Sub


