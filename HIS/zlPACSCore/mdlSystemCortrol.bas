Attribute VB_Name = "mdlSystemCortrol"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''[ͼ�����ġ�ϵͳ��ע��˵��]''''''''''''''''''''''''''''''''''
''''ͼ���ڲ�������50��ϵͳ��ע��
''''ͬʱ���м����������͵ı�ע��LabelType = 7����λ��עdoLabelSpecial����LabelType = 11�����doLabelRuler��
''''
''''1#      �ü��þ��ο�
''''2#      �ü��ú�ɫ�ڸǾ���
''''3#      �ü��ú�ɫ�ڸǾ���
''''4#      �ü��ú�ɫ�ڸǾ���
''''5#      �ü��ú�ɫ�ڸǾ���
''''6#      1#���׵�TEXT��
''''7#      ��λ��ע��,�ĸ���λ��ע��˳���ǹ̶��ģ�һ������������
''''8#      ��λ��ע��
''''9#      ��λ��ע��
''''10#     ��λ��ע��
''''11#     ��ע��������ã�
''''12#     ��ע��������ã�
''''13#     ��ע��������ã�
''''14#     ��ע��������ã�
''''15#     ��ע��������ã�
''''16#     ��ע��������ã�
''''17#     ��ע��������ã�
''''18#     ��ע��������ã�
''''19#     ��ע��������ã�
''''20#     ��ע��������ã�
''''21#     ʸ��״�ؽ�������
''''22#     ʸ��״�ؽ��ĺ���
''''23#     ʸ��״�ؽ��ıߵ�
''''24#     ʸ��״�ؽ��ıߵ�
''''25#     ʸ��״�ؽ��ıߵ�
''''26#     ʸ��״�ؽ��ıߵ�
''''27#     ʸ��״�ؽ������ĵ�
''''28#     ʸ��״�ؽ��Ľ��ͼ��λ��
''''29#     ʸ��״�ؽ��Ľ��ͼʸ״λ���߹�״λ��
''''30#     ��ǰͼ�񴰿�/��λ��ʾ
''''31#     �����Ľ���Ϣ����
''''32#     �����Ľ���Ϣ����
''''33#     �����Ľ���Ϣ����
''''34#     �����Ľ���Ϣ����
''''35#     ���˱����
''''36#     ���˱����
''''37#     ���˱����
''''38#     ���˱����
''''39#     ���˱����ĵ�λ
''''40#     ���˱���ϵĵ�λ
''''41#     ���˱���ҵĵ�λ
''''42#     ���˱���µĵ�λ
''''43#     ��ӡ���

''''''''''''''''''''''''''''''''''''''''[���ݿ�����]''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'�������ݿ�����

Public Const intSliceOffset = 10


Public cnAccess As ADODB.Connection
Public gcnOracle As ADODB.Connection         ''''ORACLE���ݿ����Ӵ�
Public strADOcn  As String
Public rsTemp As Recordset
Public blLocalRun As Boolean                '��Ƭվ�Ƿ��ڱ��ػ򵥻�����
Public glngUserID  As Long                  '����Ӱ��������ñ�Ĭ�ϴ���λ�ȱ���ص��û�ID
Public PstrCheckUID As String               '���µļ��UID
Public PstrFtpHost As String                'FTP������ַ
Public PstrFtpUser As String                'FTP�û���
Public PstrFtpPwd As String                 'FTP����
Public PstrFtpPath As String                'FTPĿ¼����ͼ���ļ����ڵľ���Ŀ¼
Public PstrBufferImagePath As String        '��������Ŀ¼
''''''''''''''''''''''''''''''''''''''''Ȩ�ޱ�������'''''''''''''''''''''''''''''''''''''''
Public mstrPrivs As String
Public glngSys As Long                      'ϵͳ��
''''''''''''''''''''''''''''''''''''''''[���沼��]'''''''''''''''''''''''''''''''''''''''''
Public intSpaceSize As Integer                          ''����֮��ļ����ȡ��߶�
Public intMaxAreaX As Long                              ''�������ɻ��ֵ����и���
Public intMaxAreaY As Long                              ''�������ɻ��ֵ����и���
Public Const G_INT_MAX_IMG_COL = 8                      ''���������Ի��ֵ�ͼ���������ͼ������
Public Const G_INT_MAX_IMG_ROW = 8                      ''���������Ի��ֵ�ͼ���������ͼ������

Public lngDefaultImageBorderColor As Long               ''Ĭ�ϣ�δѡ�У��ǵ�ǰ��ͼ��߿���ɫ
Public lngDefaultImageBorderLineStyle As Long           ''Ĭ�ϣ�δѡ�У��ǵ�ǰ��ͼ��߿�����
Public lngDefaultImageBorderLineWidth As Long           ''Ĭ�ϣ�δѡ�У��ǵ�ǰ��ͼ��߿��߿�

Public lngCurrentImageBorderColor As Long               ''��ǰͼ��߿���ɫ
Public lngCurrentSeriesBorderColor As Long              ''��ǰ��δѡ�У����б߿���ɫ
Public lngCurrentImageBorderLineStyle As Long           ''��ǰͼ��߿�����
Public lngCurrentImageBorderLineWidth As Long           ''��ǰͼ��߿��߿�

Public lngSelectedImageBorderColor As Long              ''ѡ��ͼ��߿���ɫ
Public lngSelectedImageBorderLineStyle As Long          ''ѡ��ͼ��߿�����
Public lngSelectedImageBorderLineWidth As Long          ''ѡ��ͼ��߿��߿�

Public lngCellSpacing   As Long                         ''ͼ����
Public lngImageIdentifierSize   As Long                 ''ͼ��ѡ���Ǵ�С
Public blnDsipSpilthBorder As Boolean                   ''����߿��Ƿ���ʾ
Public lngSelectImageForeColour As Long                 ''ѡ��ͼ���ʶ���ɫ
Public lngViewerBackColor   As Long                     ''Viewer������ɫ
Public lngProgramBackColor As Long                      ''��Ƭ����վ��ɫ,���򱳾���ɫ
Public blnDockMiniImage As Boolean                      ''����ͼͣ���ڲ˵���
Public blnShowMiniImageInfo As Boolean                  ''����ͼ����ʾͼ����Ϣ
Public blnShowMPRLine As Boolean                        ''MPR����ʱ����ʾλ�ø�����
Public blnSquareFrame As Boolean                        ''�Ƿ������ο�ѡ����ͼ
Public blnShowPrintTag As Boolean                       ''�Ƿ���ʾ��Ƭ�Ѵ�ӡ�ı��
Public blnPrintFilmBeep As Boolean                      ''��Ƭ��ӡʱ�Ƿ���ʾ������������ӽ�Ƭ����ӡ
Public gblnCompareSize As Boolean                       ''����FTP�ļ���С�Ա�

Public Const G_INT_MPR_RADIUS = 16                      ''ʸ��״�ؽ��ľ��ֱ��

'''''''''''''''''''''''''''''''''''''''''''[LABEL����]''''''''''''''''''''''''''''''''''''
Public Const G_INT_SYS_LABEL_COUNT = 50                 ''ϵͳ�ڲ�ʹ�õ�ͼ���ע����
Public Const G_INT_SYS_LABEL_HIDE_LEFT = -12000         ''����ϵͳ��עʱ������ע�ƶ��������λ��
Public Const G_INT_SYS_LABEL_HIDE_TOP = 12000           ''����ϵͳ��עʱ������ע�ƶ������ϱ�λ��
Public Const G_INT_SYS_LABEL_TIWEI = 7                  ''��λ��ע��,�ϣ��ң��¡��ĸ���ע�̶�˳������
Public Const G_INT_SYS_LABEL_WWWL = 30                  ''����λ��ע
Public Const G_INT_SYS_LABEL_PAT_INFO = 31              ''�����Ľ���Ϣ���ϣ����£����£����ϡ��ĸ���ע�̶�˳��������31-34
Public Const G_INT_SYS_LABEL_RULLER = 35                ''���˱����Ϣ���ϣ��ң��¡��ĸ���ע�̶�˳���������������ĸ������ĵ�λ��ע��35-42
Public Const G_INT_SYS_LABEL_MPRV = 21                  ''MPRʸ��״λ�ؽ�������
Public Const G_INT_SYS_LABEL_MPRH = 22                  ''MPRʸ��״λ�ؽ��ĺ���
Public Const G_INT_SYS_LABEL_MPR_POINT_V1 = 23          ''MPRʸ��״λ�ؽ������ߵ�һ���˵�
Public Const G_INT_SYS_LABEL_MPR_POINT_V2 = 25          ''MPRʸ��״λ�ؽ������ߵڶ����˵�
Public Const G_INT_SYS_LABEL_MPR_POINT_H1 = 24          ''MPRʸ��״λ�ؽ��ĺ��ߵ�һ���˵�
Public Const G_INT_SYS_LABEL_MPR_POINT_H2 = 26          ''MPRʸ��״λ�ؽ��ĺ��ߵڶ����˵�
Public Const G_INT_SYS_LABEL_MPR_POINT_O = 27           ''MPRʸ��״λ�ؽ������ߺͺ��ߵ����ĵ�
Public Const G_INT_SYS_LABEL_MPR_RESULT_H = 28          ''ʸ��״�ؽ��Ľ��ͼ��λͶӰ��
Public Const G_INT_SYS_LABEL_MPR_RESULT_V = 29          ''ʸ��״�ؽ��Ľ��ͼʸ״λ���߹�״λͶӰ��
Public Const G_INT_SYS_LABEL_PRINT_TAG = 43             '' ��ӡ���

Public intTextoOffX As Long, intTextoOffY As Long       ''��ע���ֵ�ƫ����
Public lngLabelColor As Long                            ''��ע��ʾɫ����ɫ
Public lngLabelSelectedColor As Long                    ''��עѡ��ɫ����ɫ
Public lngLabelLineStyleNorm As Long                    ''��ע��������
Public lngLabelLineWidthNorm As Long                    ''��ע�����߿�
Public lngLabelFontSize As Long                         ''��ע���ִ�С
Public lngLabelLineStyleSel As Long                     ''��עѡ������
Public lngLabelLineWidthSel As Long                     ''��עѡ���߿�
Public intPeriodSize As Long                            ''ѡ������С
Public lngPeriodColor As Long                           ''ѡ������ɫ
Public blnLabelTextScaleFontSize As Boolean             ''��ע���ִ�С�Ƿ�����ͼ��һ������
Public intSelectLabelStyle  As Integer                  ''��Ҫ���ı�עDicomOBJECT���ͱ���,�ڰ��±�ע��ť��ʱ����д
Public bROIArea As Boolean                              ''��ʾ���
Public bROIMean As Boolean                              ''��ʾƽ��ֵ
Public bROIStandardDeviation As Boolean                 ''��ʾ������
Public bROILength As Boolean                            ''��ʾ�ܳ�
Public bROIMax As Boolean                               ''��ʾ���ֵ
Public bROIMin As Boolean                               ''��ʾ��Сֵ
Public bROITextChinese As Boolean                       ''������ʾ�Ĺ���������Ϣʹ������
Public lngWinWidthLevelLocation As Long                 ''����λ��λ�� 1-�ϱߣ�2-�±ߣ�3-��ߣ�4-�ұ�
Public intNarrowThreshold As Integer                    ''Ѫ����խ������Ԥ����ֵ
Public intStandardThreshold   As Integer                ''Ѫ����խ������Ԥ����ֵ
Public intVasEdgeWidth As Integer                       ''Ѫ����խ��������ʾѪ�ܱڶ�ֱ�ߵĿ��

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public blnRulerDsipLeft As Boolean                      ''�Ƿ���ʾ��߱��
Public blnRulerDsipBottom   As Boolean                  ''�Ƿ���ʾ�ײ����
Public blnRulerDsipRight   As Boolean                   ''�Ƿ���ʾ�ұ߱��
Public blnRulerDsipTop   As Boolean                     ''�Ƿ���ʾ�������
Public intRulerWidth As Long                            ''��߿��
Public intRulerHeight   As Long                         ''��߸߶�
Public intRulerTop   As Long
Public intRulerLeft   As Long
Public lngRulerLeftColor   As Long                      ''�����ɫ
Public intRulerLineWidth   As Long                      ''����߿�

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public lngReferenceLineColor   As Long                  ''��λ����ɫ
Public lngReferenceLineStyle   As Long                  ''��λ������
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public lngStackStep As Long                             ''���󲽳�
Public lngZoomStep As Long                              ''���Ų���
Public lngCruiseStep As Long                            ''���β���
Public lngWidthLevelStep As Long                        ''��������
Public intMouseWheelRoll As Integer                     ''�����ֹ������÷�
Public intMouseWheelDrag As Integer                     ''�����ֹ������÷�
''''''''''''''''''''''''''''''''''''''''[������Ϣ]'''''''''''''''''''''''''''''''''''''''''
Public blnAnatomicMarkersLeft As Boolean                ''�Ƿ���ʾ�����λ���
Public blnAnatomicMarkersTop   As Boolean               ''�Ƿ���ʾ������λ���
Public blnAnatomicMarkersBottom   As Boolean            ''�Ƿ���ʾ�ײ���λ���
Public blnAnatomicMarkersRight   As Boolean             ''�Ƿ���ʾ�ұ���λ���
Public blnChinaMark   As Boolean                        ''�Ƿ���ú�����ʾ��λ���
Public lngPatientInfoInvisibleSize As Long              ''ͼ��С��Xʱ������ʾ������Ϣ
Public lngpatientInfoColor As Long                      ''������Ϣ��ɫ
Public blnpatientInfoScaleFontSize As Boolean           ''������Ϣ���ִ�С�Ƿ�����ͼ��һ������
Public blnHidePatientInfo As Boolean                    ''�Ƿ���ʾ������Ϣ

''''''''''''''''''''''''''''''''''''''''[ͼ���ֵ]'''''''''''''''''''''''''''''''''''''''''
Public Const intMagnificationMode = 3                   ''ͼ���ֵģʽ

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public lngPatientInfoFontSize As Long                   ''������Ϣ��ʾ�����С
Public blnPatientInfoFontBold As Boolean                ''������Ϣ��ʾ�������
Public blnPatientInfoFontItalic As Boolean              ''������Ϣ��ʾ����б��
Public strPatientInfoFontName As String                 ''������Ϣ��ʾ��������
Public lngPatientInfoTitle As Long                      ''������Ϣʹ�õ���ͷ��0--��ʹ����ͷ��1--������ͷ��2--Ӣ����ͷ
Public bShowFilmConfig As Boolean                       ''�ڵ�����ఴťʱ���Ƿ񵯳���Ƭ���ô���

Public blnInterfaceParaModified As Boolean              ''��¼Ӱ�����ϵͳ������ֵ�Ƿ����ı䣿

''''''''''''''''''''''''''''''''''''''''''''''''''''

Public lngReferenceLineSpacing As Long                  ''��λ�ߵ���ʾ���
Public cstrPrintAE As String                            ''��ӡʱʹ�õı�����AE����
Public intFilmFontSize As Integer                       ''��¼��ӡ��Ƭʹ�õı�ע���ִ�С
Public blnPrintOkEcho As Boolean                        ''��ӡ��ɺ󣬵�����ʾ�Ի���

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''[�˵����Ƶ���ʱ����]''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Button_miSerialPlaceInPhase  As Boolean           ''����ͬ��
Public Button_miSerialManualSyn As Boolean               ''�ֹ�����ͬ��
Public Button_miImageInPhase As Boolean                  ''ͼ��״̬ͬ��
Public Button_miLookOrBrowse As Boolean                  ''����۲�ģʽ
Public Button_miCutOut As Boolean                        ''�ü�
Public Button_miFrameSelectImage As Boolean              ''��ѡͼ��
Public Button_miStack As Boolean                         ''����
Public Button_miWidthLevel As Boolean                    ''�ֶ�����
Public Button_miZoom As Boolean                          ''����
Public Button_miCruise As Boolean                        ''����
Public Button_mi3dCursor As Boolean                      ''3D���
Public Button_miAutoWidthLevel As Boolean                ''����Ӧ����
Public Button_miDispPatientInfo As Boolean               ''������ʾ
Public Button_miLabeltext As Boolean                     ''����
Public Button_miDispLabelInfo As Boolean                 ''��ע��ʾ
Public Button_miLabelAngle As Boolean                    ''�Ƕ�
Public Button_miLabelPolygon As Boolean                  ''����
Public Button_miAllReferLine As Boolean                  ''���ж�λ��
Public Button_miFLReferLine As Boolean                   ''��β��λ��
Public Button_miCurrentReferLine As Boolean              ''��ǰ��λ��
Public Button_miLabelRectangle As Boolean                ''����
Public Button_miLabelLine As Boolean                     ''ֱ��
Public Button_miLabelEllipse As Boolean                  ''��Բ
Public Button_miLabelArrowhead As Boolean                ''��ͷ
Public Button_miLabelPolyLine As Boolean                 ''����
Public Button_miLabelVasMeasure As Boolean               ''Ѫ����խ����
Public Button_miLabelCadiothoracicRatio As Boolean       ''���رȲ���
Public Button_miFullScreen As Boolean                    ''ȫ����ʾ
Public Button_miMouseShowValue As Boolean                ''���������ʾCTֵ
Public Button_miShowMiniSeries As Boolean                ''��ʾ��������ͼ
Public Button_miViewAllSeries As Boolean                 ''ȫ���й�Ƭ
Public Button_miShowOverlay As Boolean                   ''��ʾOverlay


'''''''''''''''''''''''''�˵��Ͱ�ť��������'''''''''''''''''''''''''''''''''''''''
'---------------------------------------------
'-----------------�ļ��˵�--------------------
'---------------------------------------------
'Ԥ����100������ʹ�õ�15
Public Const ID_File = 101                                               ''�ļ�
Public Const ID_File_Open = 102                                          ''���ļ�
Public Const ID_File_Close = 103                                         ''�ر�����
Public Const ID_File_DelAllPhoto = 104                                   ''ɾ������ͼ��
Public Const ID_File_DelReport = 105                                     ''ɾ������ͼ��
Public Const ID_File_SaveFile = 106                                      ''�����ļ�
Public Const ID_File_SaveASFile = 107                                    ''����ļ�
Public Const ID_File_SaveToCD = 115                                      ''����CD
Public Const ID_File_SAveASReport = 108                                  ''���汨��ͼ
'***************************************
Public Const ID_File_Send = 109                                          ''����
Public Const ID_File_Send_GetHost = 110                                  ''��������
Public Const ID_File_Send_OutPowerPoint = 111                            ''�����PowerPoint
'***************************************
Public Const ID_File_OpenDicomDir = 114                                  ''��DICOMDIR
Public Const ID_File_PhotoProperty = 112                                 ''ͼ������
Public Const ID_File_Exit = 113                                          ''�˳�
'-----------------------------------------------
'------------------��ͼ�˵�---------------------
'-----------------------------------------------
'Ԥ����200-300������ʹ��251
Public Const ID_View = 200                                              ''��ͼ
Public Const ID_View_Typeset = 201                                      ''���氲��
Public Const ID_View_OneBrowse = 202                                    ''�����й۲�
Public Const ID_View_PropertyShow = 203                                 ''������ʾ
Public Const ID_View_LableShow = 204                                    ''��ע��ʾ
Public Const ID_View_UpSeries = 247                                     ''��һ����
Public Const ID_View_DownSeries = 248                                   ''��һ����
Public Const ID_View_ShowMiniSeries = 249                               ''��ʾ��������ͼ
Public Const ID_View_ViewAllSeries = 250                                ''ȫ���й�Ƭ
Public Const ID_View_ShowOverlay = 251                                  ''��ʾOverlay
'***********************************************
Public Const ID_View_PhotoSerial = 205                                  ''ͼ��˳��
Public Const ID_View_PhotoSerial_PhotoNumber = 206                      ''ͼ���
Public Const ID_View_PhotoSerial_BedASC = 207                           ''��λ����
Public Const ID_View_PhotoSerial_BedDESC = 208                          ''��λ����
Public Const ID_View_PhotoSerial_CollectionTime = 209                   ''�ɼ�ʱ��
Public Const ID_View_PhotoSerial_PhotoTime = 210                        ''ͼ��ʱ��
'************************************************
Public Const ID_View_ShowScale = 230                                    ''��ʾ����
Public Const ID_View_ShowScale_AutoShow = 240                           ''����Ӧ
Public Const ID_View_ShowScale_50% = 241                                ''50%
Public Const ID_View_ShowScale_100% = 242                               ''100%
Public Const ID_View_ShowScale_200% = 243                               ''200%
Public Const ID_View_showScale_400% = 244                               ''400%
Public Const ID_View_showScale_150% = 2411                              ''150%
Public Const ID_View_showScale_250% = 2412                              ''250%
Public Const ID_View_showScale_300% = 2413                              ''300%
Public Const ID_View_ShowScale_Custom = 245                             ''�Զ���
'************************************************
Public Const ID_View_FullScreen = 246                                   ''ȫ����ʾ
''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''�����˵�''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''
'Ԥ����300-400������ʹ��367,����349-360�ǿ�ݼ�������ID
Public Const ID_Active = 300                                            ''����
'***********************************************
Public Const ID_Active_Select = 301                                     ''ѡ��
Public Const ID_Active_Select_OneSelect = 302                           ''����ѡ��
Public Const ID_Active_Select_SelectAllSerial = 303                     ''ѡ����������
Public Const ID_Acitve_Select_SelectAllPhoto = 304                      ''ѡ������ͼ��
'************************************************
Public Const ID_Active_Also = 305                                       ''ͬ��
Public Const ID_Active_Also_Serial = 306                                ''����ͬ��
Public Const ID_Active_Also_Photo = 307                                 ''ͼ��ͬ��
Public Const ID_Active_Also_ManualSerial = 363                          ''�ֹ�����ͬ��
Public Const ID_Active_Also_LockSerial = 364                            ''��������
'************************************************
Public Const ID_Active_Shuttle = 308                                    ''����
Public Const ID_Active_Cruise = 309                                     ''����
Public Const ID_Active_Cut = 310                                        ''�ü�
Public Const ID_Active_Zoom = 311                                       ''����
Public Const ID_Active_ReSetAll = 312                                   ''�ָ�����
'************************************************
Public Const ID_Active_AdjustWindow = 313                               ''����
Public Const ID_Active_AdjustWindow_HandAdjustWindow = 314              ''�ֿص���
Public Const ID_Active_AdjustWindow_HandAdjustWindow_ReSet = 349        ''�ֿص���_�ָ�
Public Const ID_Active_AdjustWindow_HandAdjustWindow_Custom = 350       ''�ֿص���_�Զ���
Public Const ID_Active_AdjustWindow_HandAdjustWindow_F3 = 351           ''�ֿص���_F3
Public Const ID_Active_AdjustWindow_HandAdjustWindow_F4 = 352           ''�ֿص���_F4
Public Const ID_Active_AdjustWindow_HandAdjustWindow_F5 = 353           ''�ֿص���_F5
Public Const ID_Active_AdjustWindow_HandAdjustWindow_F6 = 354           ''�ֿص���_F6
Public Const ID_Active_AdjustWindow_HandAdjustWindow_F7 = 355           ''�ֿص���_F7
Public Const ID_Active_AdjustWindow_HandAdjustWindow_F8 = 356           ''�ֿص���_F8
Public Const ID_Active_AdjustWindow_HandAdjustWindow_F9 = 357           ''�ֿص���_F9
Public Const ID_Active_AdjustWindow_HandAdjustWindow_F10 = 358          ''�ֿص���_F10
Public Const ID_Active_AdjustWindow_HandAdjustWindow_F11 = 359          ''�ֿص���_F11
Public Const ID_Active_AdjustWindow_HandAdjustWindow_F12 = 360          ''�ֿص���_F12
Public Const ID_Active_AdjustWindow_AutoAdjustWindow = 315              ''����Ӧ����
Public Const ID_Active_AdjustWidnow_CustomAdjustWindow = 316            ''�Զ������
'************************************************
Public Const ID_Active_PointingLine = 317                               ''��λ��
Public Const ID_Active_PointingLine_ALL = 318                           ''���ж�λ��
Public Const ID_Active_PointingLine_FirstLast = 319                     ''��λ��λ��
Public Const ID_Active_PointingLine_Now = 320                           ''��ǰ��λ��
Public Const ID_Active_PointingLine_3DLine = 321                        ''3D��궨λ
'************************************************
Public Const ID_Active_Eddy = 322                                       ''��ת
Public Const ID_Active_Eddy_LeftRight = 323                             ''���ҷ�ת
Public Const ID_Active_Eddy_TopButton = 324                             ''��ֱ��ת
Public Const ID_Active_Eddy_Left90 = 325                                ''����90
Public Const ID_Active_Eddy_Right90 = 326                               ''����90
'************************************************
Public Const ID_Active_ReverseVideo = 327                               ''����
'************************************************
Public Const ID_Active_SieveLens = 328                                  ''�˾�
Public Const ID_Active_SieveLens_Model = 32810                          ''�����˾�ģ�壬��32810��ʼ��32850�����֧��40��
Public Const ID_Active_SieveLens_LancetMinus = 329                      ''��Ե��ǿǿ�ȼ���
Public Const ID_Active_SieveLens_LancetAdd = 330                        ''��Ե��ǿǿ������
Public Const ID_Active_SieveLens_FlatnessMinus = 331                    ''ƽ������
Public Const ID_Active_SieveLens_FlatnessAdd = 332                      ''ƽ������
Public Const ID_Active_Sievelens_LeftMoveMinus = 333                    ''��Ե��ǿ���ȼ���
Public Const ID_Active_Sievelens_LeftMoveAdd = 334                      ''��Ե��ǿ��������
Public Const ID_Active_Sievelens_PhotoReset = 335                       ''ͼ��ԭ
'************************************************
Public Const ID_Active_Lable = 336                                      ''��ע
Public Const ID_Active_Lable_Text = 337                                 ''����
Public Const ID_Active_Lable_Arrowhead = 338                            ''��ͷ
Public Const ID_Active_Lable_Ellipse = 339                              ''��Բ
Public Const ID_Active_Lable_Angle = 340                                ''�Ƕ�
Public Const ID_Active_Lable_Curve = 341                                ''����
Public Const ID_Active_Lable_Area = 342                                 ''����
Public Const ID_Active_Lable_BeeLine = 343                              ''ֱ��
Public Const ID_Active_Lable_Rect = 344                                 ''����
Public Const ID_Active_Lable_AreaBeeLinePhoto = 345                     ''����ֱ��ͼ
Public Const ID_Active_Lable_AdjustLine = 346                           ''У׼
Public Const ID_Active_Lable_ClearLbale = 347                           ''�����ע
Public Const ID_Active_Lable_DelSelectLable = 348                       ''ɾ����ע
Public Const ID_Active_Lable_VasMeasure = 361                           ''��խѪ�ܲ���
Public Const ID_ACtive_Mouse_Value = 362                                ''���������ʾCTֵ
Public Const ID_ACtive_FrameSelectImage = 365                           ''��ѡͼ��
Public Const ID_ACtive_SaveInReport = 366                               ''��ǰͼ���ɱ���ͼ
Public Const ID_Active_Lable_CadioThoracicRatio = 367                   ''���رȲ���
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''���߲˵�''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'Ԥ����400-500������ʹ��420
Public Const ID_Tool = 400                                              ''����
Public Const ID_Tool_Movie = 401                                        ''��Ӱ
Public Const ID_Tool_Magnifier = 402                                    ''�Ŵ�
Public Const ID_Tool_ArrowyCoronaryReset = 403                          ''ʸ��״�ؽ�
Public Const ID_Tool_NumberMinusShadow = 404                            ''���ּ�Ӱ
Public Const ID_Tool_BogusColour = 405                                  ''α�ʹ۲�
Public Const ID_Tool_FilmPrint = 406                                    ''��Ƭ��ӡ
Public Const ID_Tool_Film_AddSeries = 40601                             ''��Ƭ��ӡ--��ӡ��ǰ����
Public Const ID_Tool_Film_AddImage = 40602                              ''��Ƭ��ӡ -- ��ӡ��ǰͼ��
Public Const ID_Tool_Film_AddSelected = 40603                           ''��Ƭ��ӡ -- ��ӡ��ǰѡ��
Public Const ID_Tool_Film_AddInterval = 40604                           ''��Ƭ��ӡ -- �����ӡ��ǰ����
Public Const ID_Tool_PhotoUnite = 407                                   ''ͼ��ƴ��
Public Const ID_Tool_LableTool = 408                                    ''��ע����
Public Const ID_Tool_LookPhotoOption = 409                              ''��Ƭѡ��
'*************************************
Public Const ID_ToolBar = 410                                           ''������
Public Const ID_ToolBar_Left = 411                                      ''����
Public Const ID_ToolBar_Right = 412                                     ''����
Public Const ID_ToolBar_Top = 413                                       ''����
Public Const ID_ToolBar_Button = 414                                    ''����
Public Const ID_toolBar_16Icon = 415                                    ''16*16ͼ��
Public Const ID_ToolBar_24Icon = 416                                    ''24*24ͼ��
Public Const ID_ToolBar_32Icon = 417                                    ''32*32ͼ��
Public Const ID_ToolBar_Hide = 418                                      ''���ع�����
Public Const ID_Tool_NothinMouseState = 419                             ''����������ʹ��״̬
Public Const ID_Tool_SlopeReconstruction = 420                          ''б���ؽ�
'*************************************
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''�����˵�''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''�����˵�'''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''Ԥ����500---800'''''''''''''''''''''
'''''''''����frmImageSpelling���幤�߰�ť����801-807'''''''''
Public Const ID_frmImageSpelling_CompleteSpelling = 801                 ''���ƴ��
Public Const ID_frmImageSpelling_SavePhoto = 802                        ''����ͼ��
Public Const ID_frmImageSpelling_DelPhoto = 803                         ''ɾ��ͼ��
Public Const ID_frmImageSpelling_ZoomOut = 804                          ''����ͼ��
Public Const ID_frmImageSpelling_Quit = 806                             ''�˳�
Public Const ID_frmImageSpelling_CutOut = 807                           ''�ü�ͼ��
Public Const ID_frmImageSpelling_Move = 808                             ''�ƶ�ͼ��
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''����frmFilm���幤�߰�ť����830-872'''''''''''''''
Public Const ID_frmFilm_TakePictures = 830                              ''����
Public Const ID_frmFilm_FilmCol = 831                                   ''����
Public Const ID_frmFilm_FilmRow = 832                                   ''����
Public Const ID_frmFilm_RectPhotCase = 833                              ''������ͼ���
Public Const ID_frmFilm_FormatCustom = 834                              ''��ʽ����
Public Const ID_frmFilm_FilmSize = 835                                  ''��Ƭ��С
Public Const ID_frmFilm_Format = 836                                    ''��ʽ
Public Const ID_frmFilm_Camera = 837                                    ''���
Public Const ID_frmFilm_Quit = 838                                      ''�˳�
Public Const ID_frmFilm_DeleteImg = 839                                 ''ɾ��ͼ��
Public Const ID_frmFilm_WinLevel = 840                                  ''����
Public Const ID_frmFilm_Pan = 841                                       ''����
Public Const ID_frmFilm_Zoom = 842                                      ''����
Public Const ID_frmFilm_RotateLeft = 843                                ''������ת
Public Const ID_frmFilm_RotateRight = 844                               ''������ת
Public Const ID_frmFilm_FlipHorizontal = 845                            ''���Ҿ���
Public Const ID_frmFilm_FlipVertical = 846                              ''���¾���
Public Const ID_frmFilm_Resume = 847                                    ''�ָ�
Public Const ID_frmFilm_ImgSynchronal = 848                             ''ͼ��ͬ��
Public Const ID_frmFilm_Divide = 851                                    ''ͼ��ָ�
Public Const ID_frmFilm_UnDivide = 852                                  ''ȡ���ָ�
Public Const ID_frmFilm_Invert = 853                                    ''����
Public Const ID_frmFilm_SelAll = 854                                    ''ȫѡ
Public Const ID_frmFilm_RectZoom = 855                                  ''��ѡ����
Public Const ID_frmFilm_CutOut = 856                                    ''�ü�
Public Const ID_frmFilm_CutOut_14X17 = 85601                            ''�ü����̶�������14*17
Public Const ID_frmFilm_CutOut_11X14 = 85602                            ''�ü����̶�������11*14
Public Const ID_frmFilm_CutOut_10X14 = 85603                            ''�ü����̶�������10*14
Public Const ID_frmFilm_CutOut_8X10 = 85604                             ''�ü����̶�������8*10
Public Const ID_frmFilm_CutOut_14X14 = 85605                            ''�ü����̶�������14*14
Public Const ID_frmFilm_CutOut_17X14 = 85606                            ''�ü����̶�������17*14
Public Const ID_frmFilm_CutOut_14X11 = 85607                            ''�ü����̶�������14*11
Public Const ID_frmFilm_CutOut_14X10 = 85608                            ''�ü����̶�������14*10
Public Const ID_frmFilm_CutOut_10X8 = 85609                             ''�ü����̶�������10*8
Public Const ID_frmFilm_CutOut_Custom = 85610                           ''�ü������ɱ���

Public Const ID_frmFilm_FilterLengthUp = 857                            ''ƽ������
Public Const ID_frmFilm_FilterLengthDown = 858                          ''ƽ������
Public Const ID_frmFilm_OpenImages = 859                                ''��ͼ��
Public Const ID_frmFilm_Label = 860                                     ''��ע���������˵�
Public Const ID_frmFilm_Label_A = 861                                   ''��ע����-Anterior-ǰ
Public Const ID_frmFilm_Label_P = 862                                   ''��ע����-Posterior-��
Public Const ID_frmFilm_Label_L = 863                                   ''��ע����-Left-��
Public Const ID_frmFilm_Label_R = 864                                   ''��ע����-Right-��
Public Const ID_frmFilm_Label_S = 865                                   ''��ע����-Superior-��
Public Const ID_frmFilm_Label_I = 866                                   ''��ע����-Inferior-��
Public Const ID_frmFilm_Label_Delete = 867                              ''ɾ����־����
Public Const ID_frmFilm_SelNone = 868                                   ''ѡ��--ȫ��ͼ��
Public Const ID_frmFilm_SelSeries = 869                                 ''ѡ�� -- ѡ������
Public Const ID_frmFilm_SelInverse = 870                                ''ѡ�� - ��ѡ
Public Const ID_frmFilm_ImgIncrease = 871                               ''���� - ����
Public Const ID_frmFilm_ImgDecrease = 872                               ''���� - ����

''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''����frmPacsImg���幤�߰�ť����880-884'''''''''''''''
Public Const ID_PacsImg_SelectAllSeries = 880                           ''ȫѡ����
Public Const ID_PacsImg_UnSelectAllSeries = 881                         ''ȫ������
Public Const ID_PacsImg_SelectAllImages = 882                           ''ȫѡͼ��
Public Const ID_PacsImg_UnSelectAllImages = 883                         ''ȫ��ͼ��
Public Const ID_PacsImg_ReverseSelectImages = 884                       ''��ѡͼ��


Public Const ID_Help = 600                                              ''����
Public Const ID_Help_Help = 601                                         ''����
Public Const ID_Help_WebZLSOFT = 602                                    ''WEB�ϵ�����
Public Const ID_Help_WebZLSOFT_WEB = 603                                ''������ҳ
Public Const ID_Help_WebZLSOFT_Mail = 604                               ''���ͷ���
Public Const ID_Help_About = 605                                        ''����
Public Const ID_Help_UpdateDB = 606                                     ''�������ݿ�
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''��ݰ�ť''''''''''''''''''''''''''''''''''''
Public Const FSHIFT = 4                                            'Shift
Public Const FCONTROL = 8                                          'Ctrl
Public Const FALT = 16                                             'ALT
Public Const VK_F1 = &H70                                          'F1
Public Const VK_F2 = &H71                                          'F2
Public Const VK_F3 = &H72                                          'F3
Public Const VK_F4 = &H73                                          'F4
Public Const VK_F5 = &H74                                          'F5
Public Const VK_F6 = &H75                                          'F6
Public Const VK_F7 = &H76                                          'F7
Public Const VK_F8 = &H77                                          'F8
Public Const VK_F9 = &H78                                          'F9
Public Const VK_F10 = &H79                                         'F10
Public Const VK_F11 = &H7A                                         'F11
Public Const VK_F12 = &H7B                                         'F12
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''����������''''''''''''''''''''''''''''''''''''
Public Const ToolBar_Menu  As Integer = 1                               ''�˵�
Public Const ToolBar_Main  As Integer = 2                               ''��������
Public Const ToolBar_Photo  As Integer = 3                              ''ͼ��������
Public Const ToolBar_Scale As Integer = 4                               ''����������
Public Const ToolBar_Plane  As Integer = 5                              ''ƽ�湤����
Public Const ToolBar_Object  As Integer = 6                             ''���񹤾���
Public Const ToolBar_Comm  As Integer = 7                               ''����������
Public Const toolBar_PhotoStrong As Integer = 8
'''''''''''''''''''''''��ǰ����������''''''''''''''''''''''''''''''''
Public intToolBarIconSize As Integer                                    ''ͼ���С
Public intToolBarPosition As Integer                                    ''�ڷ�λ��
Public blToolBarHide As Boolean                                         ''���ع�����

''''''''''''''''''''''���������'''''''''''''''''''''''''''''''''''''
Public IntComBarTheme As Integer                                        ''ͳһ����������ʾ���
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''[ϵͳ��������]''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public cLabelStore As New Collection            '��¼�����ע��ʹ�õ�ͼ��ͷ��Ϣ
Public Const cProducer = "ZLPACS"
Public intStatusBarFontSize As Integer

'''''''''''''''''''''''''''[Ԥ�贰��λ F3---F12]''''''''''''''''''''''''''''''''''''''''''''''''''
Public Type TPresetWinWL
    bInUse As Boolean                               ''��ʶ����ݼ��Ƿ�����
    strModality As String                           ''Ӱ�����
    strWinWLCName As String                         ''��ݼ������õĴ���λ��������
    strWinWLEName As String                         ''��ݼ������õĴ���λӢ������
    lngWinWidth As Long                             ''��ݼ������õĴ���ֵ
    lngWinLevel As Long                             ''��ݼ������õĴ�λֵ
    intDefault As Integer                           ''�Ƿ�Ĭ�ϴ���λ
    lngID As Long                                   ''Ԥ�贰��λ��ID
End Type
Public aPresetWinWL() As TPresetWinWL         ''����Ԥ�贰��λ�����飬
                                                    ''����Ŀ�ݼ�ֵΪF3--F12����Ӧ��������±�

'''''''''''''''''''''''''''[Ԥ����Ļ����]''''''''''''''''''''''''''''''''''''''''''''''''''
Public Type TModifiedPresetLayout
    bModified As Boolean
    strModality As String                               ''��¼��Ӧ��Ӱ�����
    bSeriesAutoFormat As Boolean                        ''�򿪴���ʱ�Ƿ��Զ��������и�ʽ
    lngSeriesRows As Long                               ''Ԥ��򿪴���ʱʹ�õ���������
    lngSeriesColumns As Long                            ''Ԥ��򿪴���ʱʹ�õ���������
    bImageAutoFormat As Boolean                         ''��ͼ��ʱ�Ƿ��Զ�����ͼ���ʽ
    lngImageRows As Long                                ''Ԥ���ͼ��ʱʹ�õ�ͼ������
    lngImageColumns As Long                             ''Ԥ���ͼ��ʱʹ�õ�ͼ������
    bInvert As Boolean                                  ''��ͼ��ʱ�Ƿ��Զ����ף������á�
    bShowPatientInfo As Boolean                         ''��ͼ��ʱ�Ƿ���ʾ������Ϣ,�����á�
    bAutoSelectReferenceLine As Boolean                 ''��ͼ��ʱ�Ƿ��Զ�ѡ����ʾ��λ�ߣ�ֻ���CT,MRͼ���д����ã�,�����á�
    bAutoSelectSeriesSyn As Boolean                     ''��ͼ��ʱ�Ƿ��Զ�ѡ�����м�ͼ��λ��ͬ����ֻ���CT,MRͼ���д����ã�,�����á�
    lngInterpolationMode As Long                        ''ͼ��Ŵ�ʱ�Ĳ�ֵģʽ
    lngImageSort As Long                                ''ͼ������ʽ��0-Ĭ�ϣ�1-ͼ��ţ�2-��λ����3-��λ����4-�ɼ�ʱ�䣻5-ͼ��ʱ��
End Type

Public aPresetLayout() As TModifiedPresetLayout         ''����Ԥ����Ļ���ֵ�����
Public aModifiedPresetLayout() As TModifiedPresetLayout ''���汻�޸ĵ���Ļ���ֵ�����

'''''''''''''''''''''''''''''''''''''[Ԥ��ͼ������]'''''''''''''''''''''''''''''''''''''''''''''''''
Public Type TImageShutter
    bModified As Boolean            ''�Ƿ��޸�
    strModality As String           ''��¼��Ӧ��Ӱ�����
    intShutterType As Integer       ''���������ͣ�0����������1��Բ��������2������������4�������������
                '��Щ�������Ϳ����໥���ӣ�����ÿ������ֻ�ܹ�������һ�Ρ����磬ͬʱʹ��Բ�κͶ����������
                '����������Ϊ1+4��5���������ʹ���7����Ϊ��Ч���ͣ��Զ�����Ϊ0��
    intCenterX As Integer           ''Բ��������Բ��X����
    intCenterY As Integer           ''Բ��������Բ��Y����
    intRadius As Integer            ''Բ�������İ뾶
    intRectLeft As Integer          ''������������߽�
    intRectRight As Integer         ''�����������ұ߽�
    intRectUpper As Integer         ''�����������ϱ߽�
    intRectLower As Integer         ''�����������±߽�
    strVertices As String           ''����������Ķ��㼯��ʹ��Ӣ���ַ��ġ������������
    lngColor As Long                ''�����ĻҶ���ɫ
End Type

Public aImageShutter() As TImageShutter             ''����Ԥ���ͼ����������
Public aModifiedImageShutter() As TImageShutter     ''���汻�޸ĵ�ͼ����������

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''[����÷�����]'''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public iMouseFuncCount As Integer                       ''��깦�ܼ�������
Public cMouseUsage As New Collection                  ''��¼����÷��ļ���
Public cModifiedMouseUsage As New Collection          ''��ʱ��¼����÷����޸�״̬�ļ���
Public bMouseUsageModified As Boolean
Public Const lngDrawLabelFuncNo = 20                    ''����ע���ۺϹ������
Public Const lngDrawLabelCurrent = 1                    ''����ע���ܱ�ѡ��Ϊ��ǰ��갴ť��������ѡ�İ�ť���ܺ�
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''[������Ϣ��עλ�ú���ʾ����]''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Type TInfoLabelLocate                            ''��ʶͼ����Ϣ��ͼ�����ĸ��ǵ���ʾλ��
    lngID As Long                                       ''�����ݿ��е�ID��
    strGroup As String                                  ''ͼ����dicom��ʶ��Group��
    strElement As String                                ''ͼ����dicom��ʶ��Element��
    strEName As String                                  ''ͼ����Ϣ��Ӣ����
    strCName As String                                  ''ͼ����Ϣ��������
    bUsed As Boolean                                    ''����Ϣ�Ƿ�ѡ��
    lngLocation As Long                                 ''��ʶ��Ϣ���ڵ�λ��
    lngOrder As Long                                    ''��ʾ��Ϣ�ڱ�ѡ�н��ڵ�λ�����
    blnIsExport As Boolean                              ''��ʾ�Ƿ�����������Ϣ
End Type
Public aInfoLabelLocate() As TInfoLabelLocate           ''����ͼ����Ϣ��ʾ��ʽ������
Public lngInfoLabelCount As Long                        ''��¼����ʹ�õ�ͼ����Ϣ����
Public bInfoLabelModified As Boolean                    ''��¼�����Ľ���Ϣ�������Ƿ񱻸ı���

Public cDICOMPrinter As New Collection                  ''[DICOM��ӡ����������]
Public blnSelectedImageIfColor As Boolean               ''��ǰѡͼ���Ƿ��ǲ�ɫͼ��

'''''''''''''''''''''''''''''''''''[��Ƭ��ӡ��ͼ�񲼾�]''''''''''''''''''''''''

Public Const CUT_LABEL = "CUT"
Public Const POSTURE_LABEL = "��λ"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''ǧͼ��Ƭ'''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public ZLSeriesInfos As Collection          ''��¼��ǰ�򿪵�����ͼ���ļ���Ϣ
Public ZLShowSeriesInfos As Collection    '��¼�Ѿ���ʾ��������ͼ���״̬

Public Const ATTR_Ӱ����� As String = "8:60"
Public Const ATTR_���к� As String = "20:11"
Public Const ATTR_ͼ��� As String = "20:13"

Public Const ATTR_�ɼ����� As String = "8:22"
Public Const ATTR_�ɼ�ʱ�� As String = "8:32"
Public Const ATTR_ͼ������ As String = "8:23"
Public Const ATTR_ͼ��ʱ�� As String = "8:33"
Public Const ATTR_��� As String = "18:50"
Public Const ATTR_ͼ��λ�ò��� As String = "20:32"
Public Const ATTR_ͼ������ As String = "20:37"
Public Const ATTR_�ο�֡UID As String = "20:52"
Public Const ATTR_��Ƭλ�� As String = "20:1041"
Public Const ATTR_���� As String = "28:10"
Public Const ATTR_���� As String = "28:11"
Public Const ATTR_���ؾ��� As String = "28:30"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''MPR'''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Type MPRCube
    ZLShowSeriesInfos As clsSeriesInfo  '--- ԭ�е�ZLShowSeriesInfos�ṹ
    Images As New DicomImages           '--- ԭ��Viewer���Ѿ����ص�ͼ�񣬷���ָ�ͼ���еı�ע�����������ŵ���Ϣ
    blnIsMPR As Boolean                 '--- �Ƿ�ǰ��MPR�����У�����ǣ��ָ���ʱ�򣬲���Ҫ�滻�����е����ݡ�
    intViewerIndex As Integer           '--- �ڷ��ؽ������Viewer��Index
End Type
Public ZLMPRCube(1 To 3) As MPRCube
Public ZLMPRSeriesUID As String         '��ǰ��ά�ؽ�������UID
Public ZLMPRSlopeSeriesUID As String    '��ǰMPRб���ؽ�������UID


'---------------------------��Ƭվ�ͽ�Ƭ��ӡ���������ƣ�ע��-------------------------------
Public Const LOGIN_TYPE_ҽ����Ƭվ As String = "Ӱ���Ƭվ����"
Public Const LOGIN_TYPE_��Ƭ��ӡ�� As String = "Ӱ��Ƭ��ӡ������"
Public gintҽ����Ƭվ���� As Integer
Public gint��Ƭ��ӡ�� As Integer



'''''''''''''''''''''''''''[Ԥ���˾�����]''''''''''''''''''''''''''''''''''''''''''''''''''
Public Type TPresetFilter
    lngID As Long                                   ''Ԥ���˾���ID
    strname As String                               ''Ԥ���˾�������
    strModality As String                           ''Ӱ�����
    intUnSharpEnhancementUp As Integer              ''ͼ����ǿǿ������
    intUnSharpEnhancementDown   As Integer          ''ͼ����ǿǿ�ȼ���
    intUnSharpLengthUp  As Integer                  ''ͼ����ǿ��������
    intUnSharpLengthDown    As Integer              ''ͼ����ǿ���ȼ���
    intFilterLengthUp As Integer                    ''ͼ��ƽ������
    intFilterLengthDown As Integer                  ''ͼ��ƽ������
End Type
Public aPresetFilter() As TPresetFilter             ''����Ԥ���˾������飬
                                        

'''''''''''''''''''''''''''''[ͼ��ͬ��]'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const IMG_SYN_All = 0                            ''ȫ��ͬ��
Public Const IMG_SYN_WINDOW = 1                         ''����ͬ��
Public Const IMG_SYN_ZOOMPAN = 2                        ''���š�����ͬ��
Public Const IMG_SYN_ROTATE = 3                         ''��תͬ��
Public Const IMG_SYN_FLIP = 4                           ''����ͬ��
Public Const IMG_SYN_FILTER = 5                         ''�˾�ͬ��

