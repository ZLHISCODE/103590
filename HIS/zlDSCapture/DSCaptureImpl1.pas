{*******************************************************************************
��Ƶ�ɼ�COM����ʵ�ֵ�Ԫ
�����ˣ�TJH
������ǰ��2009-11-3

������...

DirectShow��ʽ˵����

MEDIATYPE_Video;    ��������Ƶ(����) 
MEDIATYPE_Audio;    ��������Ƶ(����)
MEDIATYPE_AnalogVideo;    ����ģ����Ƶ��һ������Ƶ�ɼ���������������� 
MEDIATYPE_AnalogAudio;    ����ģ����Ƶ��һ���������ɼ�������������� 
MEDIATYPE_Text��    �������� 
MEDIATYPE_Midi;    ����MIDI���� 
MEDIATYPE_STREAM;  //�ֽ���,��(Pullģʽ)�ļ�Դ������������� 
MEDIATYPE_Interleaved;    ������������������DV�������� 
MEDIATYPE_MPEG1SystemStream;    ����MPEG1��ϵͳ�� 
MEDIATYPE_MPEG2_PACK;    ����MPEG2�����ݰ� 
MEDIATYPE_MPEG2_PES;    ����MPEG2�������� 
MEDIATYPE_DVD_ENCRYPTED_PACK;    ����DVD�����õ���ý������ 
MEDIATYPE_DVD_NAVIGATION;


ý��������Ҫ��3������������majortype(������)��subtype(����˵������)��formattype(��ʽϸ������)��
��3���ָ�����һ��GUID����ʶ�����ǵ����÷ֱ��ǣ�majortype���Ե�����ý�����ͣ�
��ָ������һ����Ƶ (MEDIATYPE_Video)����Ƶ(MEDIATYPE_Audio)�����ֽ���(MEDIATYPE_Stream)�ȣ�

subtype����˵��majortype��ָ�����������ָ�ʽ�����磬��majortype����Ƶ��
subtype���Խ�һ��ָ������UYVY(MEDIASUBTYPE_UYVY)��YUY2(MEDIASUBTYPE_YUY2)��
RGB24(MEDIASUBTYPE_RGB24)����RGB32(MEDIASUBTYPE_RGB32)�ȣ���majortype����Ƶ��
subtype���Խ�һ��ָ������PCM��ʽ(MEDIASUBTYPE_PCM)����AC3��ʽ(MEDIASUBTYPE_DOLBY_AC3)��;

formattypeָ����һ�ֽ�һ��������ʽϸ�ڵ����ݽṹ���ͣ���ʽϸ��������������Ҫ������Ƶͼ��Ĵ�С��
֡�ʣ�����Ƶ�Ĳ���Ƶ�ʡ��������ȵȲ��������������ʽϸ�ڵ����ݿ�ָ�뱣����pbFormat��Ա�С�


*******************************************************************************}
unit DSCaptureImpl1;

{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ActiveX, AxCtrls, ZLDSVideoProcess_TLB, StdVcl, DSPack, VideoProcessDefine, DirectShow9,
  ExtCtrls, jpeg, ComCtrls, DSUtil, CapParameterCfg;

const
  WM_BEEPMSG = wm_user + $1089;

type
  TDSCapture = class(TActiveForm, IDSCapture)
    Timer1: TTimer;
    _VideoWindow: TVideoWindow;
    _FilterGraphic: TFilterGraph;
    
    timerSys: TTimer;
    stabStates: TStatusBar;
    _ImgCaptureFilter: TSampleGrabber;

    imgLogo: TImage;
    procedure Timer1Timer(Sender: TObject);
    procedure _VideoWindowKeyPress(Sender: TObject; var Key: Char);
    procedure timerSysTimer(Sender: TObject);
    procedure _VideoWindowClick(Sender: TObject);
    procedure _VideoWindowDblClick(Sender: TObject);
    procedure _VideoWindowKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure _VideoWindowMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure _VideoWindowMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure _VideoWindowKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure _VideoWindowEnter(Sender: TObject);
    procedure _VideoWindowExit(Sender: TObject);
    procedure _VideoWindowMouseWheel(Sender: TObject; Shift: TShiftState;
      WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
    procedure _VideoWindowMouseWheelDown(Sender: TObject;
      Shift: TShiftState; MousePos: TPoint; var Handled: Boolean);
    procedure _VideoWindowMouseWheelUp(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);
    procedure _VideoWindowPaint(Sender: TObject);
    procedure _VideoWindowMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
  private
    { Private declarations }
    _CaptureParameter: TCaptureParameter;  //�ɼ�����
    _frmCaptureParameterCfg: TfrmCapParameterCfg;

    _CaptureParameterCfgFileName: WideString; //�ɼ������������ļ�����

    _IsPreviewState: Boolean;       //�Ƿ�ΪԤ��״̬
    _IsCaptureVideo: Boolean;     //�Ƿ���Ƶ�ɼ�
    _TempCaptureVideoFile: WideString; //��ʱ�ɼ��ļ�����

    _IsStretch: Boolean;             //�Զ���䴰�ڴ�С
    _IsAdjustWindowSize: Boolean;    //�Զ��������ڴ�С
    _IsFit: Boolean;                 //�Զ���Ӧ���ڴ�С

    _GraphManager: ICaptureGraphBuilder2; //����GraphBuilder�е�����FILTER
    _CapSourceFilter: IBaseFilter;
    _EncoderFilter: IBaseFilter;      //��Ƶ������
    _AviMultiplexer: IBaseFilter;     //��·���ýӿ�
    _AviWriter: IBaseFilter;          //�ļ�д��ӿ�
    _SmartTee: IBaseFilter;           //���ݷ����ӿ�
    _SmartTee1: IBaseFilter;          //������ʲɼ�����PX1000E�����ֲɼ�����Ҫ��������SmartTee1��preview��������ʾ����Ƶ����
    _ColorSpace: IBaseFilter;         //��ɫת��Filter
    //_MjpegDescompress: IBaseFilter;   //mjpegѹ���ӿ�

    _HideCfgItem: Integer;            //��Ҫ���ص�������

    _IsEscKeyQuitFullScreen: Boolean;
    _IsDblClickQuitFullScreen: Boolean;
    _IsClickQuitFullScreen: Boolean;

    _CapParameterWindowPos: TCapParameterPostion;

    _RecordTimeLen: Integer;

    FEvents: IDSCaptureEvents;
    
    procedure ActivateEvent(Sender: TObject);
    procedure ClickEvent(Sender: TObject);
    procedure CreateEvent(Sender: TObject);
    procedure DblClickEvent(Sender: TObject);
    procedure DeactivateEvent(Sender: TObject);
    procedure DestroyEvent(Sender: TObject);
    procedure KeyPressEvent(Sender: TObject; var Key: Char);
    procedure PaintEvent(Sender: TObject);
    procedure EnterEvent(Sender: TObject);
    procedure ExitEvent(Sender: TObject);
    procedure MouseDownEvent(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure MouseUpEvent(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure MouseMoveEvent(Sender: TObject; Shift: TShiftState; X, Y: Integer);
    procedure KeyDownEvent(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure KeyUpEvent(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure ResizeEvent(Sender: TObject);
    procedure VideoSizeEvent(const videoWidth, videoHeight, windowWidth, windowHeight: Integer);
    procedure MouseWheelEvent(Sender: TObject; Shift: TShiftState;
      WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
    procedure MouseWheelDownEvent(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);
    procedure MouseWheelUpEvent(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);  


    //��ʼ��DSHOW���������
    procedure InitDShow();
    //����ʼ��DSHOW���������
    procedure ReInitDShow();


    //�����޸��¼�
    procedure ParameterChange(const capParameter: TCaptureParameter; const needCaptureSample: Boolean);
    //ѡ���е�vfw���õ����¼�
    procedure vfwConfigCallEvent(const operVfwConfigType: TVfwConfigType;
      const parentHandle: Integer; out errMsg: WideString);

    //������Ƶ����
    procedure ConfigVideoQuality(filter: IBaseFilter; captureParameter: TCaptureParameter);
    //������Ƶ��ʽ
    procedure ConfigVideoAnalog(filter: IBaseFilter; captureParameter: TCaptureParameter);
    //������Ƶ��ʽ
    procedure ConfigVideoFormat(filter: IBaseFilter; captureParameter: TCaptureParameter);


    //�������õ������ڴ�С
    procedure AdjustWindowSize();
    //�ɼ�ͼ��BMP����
    function CaptureImageToBmpObj(): TBitmap;
    //��ʾVFW���öԻ���
    procedure ShowVfwConfigDialog(const dialogType: Integer; const parentHandle: Integer);
    //��ʾVFWѹ����������
    procedure ShowVfwCompressCfgDialog(const dialogType: Integer; const parentHandle: Integer);

    //Ԥ��
    procedure Preview(const isCaptureVideo: Boolean; out errMsg: WideString);
  protected
    { Protected declarations }
    procedure DefinePropertyPages(DefinePropertyPage: TDefinePropertyPage); override;
    procedure EventSinkChanged(const EventSink: IUnknown); override;
    function Get_Active: WordBool; safecall;
    function Get_AlignDisabled: WordBool; safecall;
    function Get_AutoScroll: WordBool; safecall;
    function Get_AutoSize: WordBool; safecall;
    function Get_AxBorderStyle: TxActiveFormBorderStyle; safecall;
    function Get_Caption: WideString; safecall;
    function Get_Color: OLE_COLOR; safecall;
    function Get_DoubleBuffered: WordBool; safecall;
    function Get_DropTarget: WordBool; safecall;
    function Get_Enabled: WordBool; safecall;
    function Get_Font: IFontDisp; safecall;
    function Get_HelpFile: WideString; safecall;
    function Get_KeyPreview: WordBool; safecall;
    function Get_PixelsPerInch: Integer; safecall;
    function Get_PrintScale: TxPrintScale; safecall;
    function Get_Scaled: WordBool; safecall;
    function Get_ScreenSnap: WordBool; safecall;
    function Get_SnapBuffer: Integer; safecall;
    function Get_Visible: WordBool; safecall;
    function Get_VisibleDockClientCount: Integer; safecall;
    procedure _Set_Font(var Value: IFontDisp); safecall;
    procedure Set_AutoScroll(Value: WordBool); safecall;
    procedure Set_AutoSize(Value: WordBool); safecall;
    procedure Set_AxBorderStyle(Value: TxActiveFormBorderStyle); safecall;
    procedure Set_Caption(const Value: WideString); safecall;
    procedure Set_Color(Value: OLE_COLOR); safecall;
    procedure Set_DoubleBuffered(Value: WordBool); safecall;
    procedure Set_DropTarget(Value: WordBool); safecall;
    procedure Set_Enabled(Value: WordBool); safecall;
    procedure Set_Font(const Value: IFontDisp); safecall;
    procedure Set_HelpFile(const Value: WideString); safecall;
    procedure Set_KeyPreview(Value: WordBool); safecall;
    procedure Set_PixelsPerInch(Value: Integer); safecall;
    procedure Set_PrintScale(Value: TxPrintScale); safecall;
    procedure Set_Scaled(Value: WordBool); safecall;
    procedure Set_ScreenSnap(Value: WordBool); safecall;
    procedure Set_SnapBuffer(Value: Integer); safecall;
    procedure Set_Visible(Value: WordBool); safecall;

    //property   ����IsStretch��IsFit������Ҫ��Ϊ�˺���ǰ�����Ľӿڼ���
    function Get_IsStretch: WordBool; safecall;
    procedure Set_IsStretch(Value: WordBool); safecall;
    function Get_IsFit: WordBool; safecall;
    procedure Set_IsFit(Value: WordBool); safecall;
    function Get_IsAdjustWindowSize: WordBool; safecall;
    procedure Set_IsAdjustWindowSize(Value: WordBool); safecall;    

    function Get_IsShowState: WordBool; safecall;
    procedure Set_IsShowState(Value: WordBool); safecall;
    function Get_IsFullScreen: WordBool; safecall;
    procedure Set_IsFullScreen(Value: WordBool); safecall;
    function Get_CaptureState: WordBool; safecall;
    function Get_PreviewState: WordBool; safecall;
    function Get_IsClickQuitFullScreen: WordBool; safecall;
    function Get_IsDblClickQuitFullScreen: WordBool; safecall;
    function Get_IsEscKeyQuitFullScreen: WordBool; safecall;
    procedure Set_IsClickQuitFullScreen(Value: WordBool); safecall;
    procedure Set_IsDblClickQuitFullScreen(Value: WordBool); safecall;
    procedure Set_IsEscKeyQuitFullScreen(Value: WordBool); safecall;
    function Get_CurHeight: Integer; safecall;
    function Get_CurVideoHeight: Integer; safecall;
    function Get_CurVideoWidth: Integer; safecall;
    function Get_CurWidth: Integer; safecall;
    procedure Set_CurHeight(Value: Integer); safecall;
    procedure Set_CurVideoHeight(Value: Integer); safecall;
    procedure Set_CurVideoWidth(Value: Integer); safecall;
    procedure Set_CurWidth(Value: Integer); safecall;
    function Get_ShowModel: TShowModel; safecall;
    procedure Set_ShowModel(Value: TShowModel); safecall;
    function Get_CapParameterWindPos: TCapParameterPostion; safecall;
    procedure Set_CapParameterWindPos(Value: TCapParameterPostion); safecall;
    function Get_SnatchWay: TSnatchWay; safecall;
    procedure Set_SnatchWay(Value: TSnatchWay); safecall;
    function Get_ParameterCfgFileName: WideString; safecall;
    procedure Set_ParameterCfgFileName(const Value: WideString); safecall;
    function Get_HideCfgItem: Integer; safecall;
    procedure Set_HideCfgItem(Value: Integer); safecall;
    function Get_AppHandle: Integer; safecall;
    procedure Set_AppHandle(Value: Integer); safecall;
    function Get_RecordTimeLen: Integer; safecall;
    
    //�ͷ���Դ
    procedure FreeRes; safecall;
    //��ʼԤ��
    function StartPreview: WideString; safecall;
    //ֹͣԤ��
    function StopPreview: WideString; safecall;
    //�ɼ�BMPͼ���ļ�
    function CaptureBmpImageToFile(const fileName: WideString): WideString; safecall;
    //�ɼ�JPGͼ���ļ�
    function CaptureJpgImageToFile(const fileName: WideString; compressRate: SYSINT): WideString; safecall;
    //��ʼ��Ƶ�ɼ�
    function StartCaptureVideo(const fileName: WideString): WideString; safecall;
    //ֹͣ��Ƶ�ɼ�
    function StopCaptureVideo(out videoFile: WideString): WideString; safecall;
    //�ɼ���������  -- ����ҪParentHandle��ֵ
    function ShowCaptureParameterCfgDialog(parentHandle: Integer): WideString; safecall;
    
    //��ʾ�ɼ�Դfilter����  -- ��Ҫ����֧��
    function ShowVideoCaptureFilterCfg(parentHandle: Integer): WideString; safecall;
    //��ʾ��Ƶ����������  -- ��Ҫ����֧��
    function ShowVideoEncoderFilterCfg(parentHandle: Integer): WideString; safecall;
    //��ʾ�ɼ��˿�����  -- ��Ҫ����֧��
    function ShowVideoCapturePinCfg(parentHandle: Integer): WideString; safecall;

    //��ʾVFW��ʾ��ʽ����  -- ��Ҫ����֧��
    function ShowVfwVideoDisplayCfg(parentHandle: Integer): WideString; safecall;
    //��ʾVFW��Ƶ��ʽ����  -- ��Ҫ����֧��
    function ShowVfwVideoFormatCfg(parentHandle: Integer): WideString; safecall;
    //��ʾ��ƵԴ����  -- ��Ҫ����֧��
    function ShowVfwVideoSourceCfg(parentHandle: Integer): WideString; safecall;
    
    //�������ļ���ȡ�ɼ�����
    function ReadParameterFromFile: WideString; safecall;
    //ˢ�´���
    function RefreshWindow: WideString; safecall;
    //�˳�ȫ��
    function QuitFullScreen: WideString; safecall;
    //ȫ����ʾ
    function ShowFullScreen(parentHandle, monitorIndex: Integer): WideString; safecall;
    //������Ƶ����
    function UpdateVideoQuailty: WideString; safecall;
    //����ɼ������������ļ�
    function SaveParameterToFile: WideString; safecall;
    //ȡ����Ƶ�ɼ�����
    function GetCaptureParameter(var parameter: TCaptureParameter): WideString; safecall;
    //���òɼ�����
    function SetCaptureParameter(var parameter: TCaptureParameter): WideString; safecall;
    //����Ԥ��
    function RePreview: WideString; safecall;
    function CaptureImgToClipBoard: WideString; safecall;
    //��ʾvfwѹ������
    function ShowVfwCompressCfg(parentHandle: Integer): WideString; safecall;
    //��ʾvideoCrossbar����
    function ShowVideoCrossbarCfg(parentHandle: Integer): WideString; safecall;
    //�ɼ�bmpͼ��
    function CaptureBmpImage: IPictureDisp; safecall;
    //�ɼ�jpgͼ��(��ת����IPictureDisp�����ݽ���ɴ�λͼ��ʽ�����CaptureBmpImage�����չ�����ͬ)
    function CaptureJpgImage(compressRate: Integer): IPictureDisp; safecall;
    //ȡ��ʵ�ʵ���Ƶ�ֱ��ʴ�С
    function GetRealVideoSize: TVideoSize; safecall;

    procedure WM_BEEP(var msg: TMessage); message WM_BEEPMSG;
  public
    { Public declarations }
    procedure Initialize; override;
  end;


implementation

uses
  ComObj, ComServ, GraphicProcess, Types, DirectShow9Ex,
  CaptureDebug, FullScreenWindow, Clipbrd, DSCapParameterConfigObj, Math;


const
  FILTER_NAME_SMART_TEE: WideString = 'SMARTTEE>5E70D3C68F884604A5218A31DABB32A0';
  FILTER_NAME_SMART_TEE1: WideString = 'SMARTTEE1>5E70D3C68F884604A5218A31DABB32A0';
  FILTER_NAME_COLOR_CONVERT: WideString = 'COLORCONVERT>5678AB3216AAC694E533DF21AD3BB02A';
  FILTER_NAME_INFINITE_PIN: WideString = 'INFINITEPIN>AAAB236B4779AEFF25E3DF312DD3EFF33';
  FILTER_NAME_ENCODER: WideString = 'ENCODER>DFD34BC4FDF14B3CB52F6A0388B6AADF';
  FILTER_NAME_SAMPLE_GRABBER: WideString = 'GRABBER>F28421F5652C42CAA8AC4BF3522413F8';
  FILTER_NAME_AVI_WRITER: WideString = 'WRITER>A9BD3C9DFD5D4E1F-8245-516D5E655E68';
  FILTER_NAME_AVI_MULTIPLEXER: WideString = 'MULTIPLEXER>BAAD81147F054E1686A310D0CC1F18BD';
  FILTER_NAME_CAPSOURCE_FILTER: WideString = 'CAPSOURCE>AA53685F0B4748DD8FB22EE73172A3C3';
  FILTER_NAME_MJPEGDECOMPRESS: WideString = 'MJPEGDECOMPRESS>FBC51DB1F1AB4441BC165EE0E1816D54';


{$R *.DFM}

{ TDSCapture }

procedure TDSCapture.DefinePropertyPages(DefinePropertyPage: TDefinePropertyPage);
begin
  { Define property pages here.  Property pages are defined by calling
    DefinePropertyPage with the class id of the page.  For example,
      DefinePropertyPage(Class_DSCapturePage); }
end;



procedure TDSCapture.EventSinkChanged(const EventSink: IUnknown);
begin
  FEvents := EventSink as IDSCaptureEvents;
  inherited EventSinkChanged(EventSink);
end;

procedure TDSCapture.Initialize;
begin
  inherited Initialize;

  OnActivate := ActivateEvent;
  OnClick := ClickEvent;
  OnCreate := CreateEvent;
  OnDblClick := DblClickEvent;
  OnDeactivate := DeactivateEvent;
  OnDestroy := DestroyEvent;
  OnKeyPress := KeyPressEvent;
  OnPaint := PaintEvent;
  OnMouseMove := MouseMoveEvent;
  OnMouseDown := MouseDownEvent;
  OnMouseUp := MouseUpEvent;
  OnKeyDown := KeyDownEvent;
  OnKeyUp := KeyUpEvent;
  OnResize := ResizeEvent;
  OnEnter := EnterEvent;
  OnExit := ExitEvent;
  OnMouseWheel := MouseWheelEvent;
  OnMouseWheelDown := MouseWheelDownEvent;
  OnMouseWheelUp := MouseWheelUpEvent;
  
  InitDShow();
end;

function TDSCapture.Get_Active: WordBool;
begin
  Result := Active;
end;

function TDSCapture.Get_AlignDisabled: WordBool;
begin
  Result := AlignDisabled;
end;

function TDSCapture.Get_AutoScroll: WordBool;
begin
  Result := AutoScroll;
end;

function TDSCapture.Get_AutoSize: WordBool;
begin
  Result := AutoSize;
end;

function TDSCapture.Get_AxBorderStyle: TxActiveFormBorderStyle;
begin
  Result := Ord(AxBorderStyle);
end;

function TDSCapture.Get_Caption: WideString;
begin
  Result := WideString(Caption);
end;

function TDSCapture.Get_Color: OLE_COLOR;
begin
  Result := OLE_COLOR(Color);
end;

function TDSCapture.Get_DoubleBuffered: WordBool;
begin
  Result := DoubleBuffered;
end;

function TDSCapture.Get_DropTarget: WordBool;
begin
  Result := DropTarget;
end;

function TDSCapture.Get_Enabled: WordBool;
begin
  Result := Enabled;
end;

function TDSCapture.Get_Font: IFontDisp;
begin
  GetOleFont(Font, Result);
end;

function TDSCapture.Get_HelpFile: WideString;
begin
  Result := WideString(HelpFile);
end;

function TDSCapture.Get_KeyPreview: WordBool;
begin
  Result := KeyPreview;
end;

function TDSCapture.Get_PixelsPerInch: Integer;
begin
  Result := PixelsPerInch;
end;

function TDSCapture.Get_PrintScale: TxPrintScale;
begin
  Result := Ord(PrintScale);
end;

function TDSCapture.Get_Scaled: WordBool;
begin
  Result := Scaled;
end;

function TDSCapture.Get_ScreenSnap: WordBool;
begin
  Result := ScreenSnap;
end;

function TDSCapture.Get_SnapBuffer: Integer;
begin
  Result := SnapBuffer;
end;

function TDSCapture.Get_Visible: WordBool;
begin
  Result := Visible;
end;

function TDSCapture.Get_VisibleDockClientCount: Integer;
begin
  Result := VisibleDockClientCount;
end;

procedure TDSCapture._Set_Font(var Value: IFontDisp);
begin
  SetOleFont(Font, Value);
end;

procedure TDSCapture.ActivateEvent(Sender: TObject);
begin
  try
    if FEvents <> nil then FEvents.OnActivate;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSCapture', 'ActivateEvent', e.Message);
    end;
  end;
end;

procedure TDSCapture.ClickEvent(Sender: TObject);
begin
  try
    if FEvents <> nil then FEvents.OnClick;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSCapture', 'ClickEvent', e.Message);
    end;
  end;  
end;

procedure TDSCapture.CreateEvent(Sender: TObject);
begin
  try
    if FEvents <> nil then FEvents.OnCreate;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSCapture', 'ClickEvent', e.Message);
    end;
  end;    
end;

procedure TDSCapture.DblClickEvent(Sender: TObject);
begin
  try
    if FEvents <> nil then FEvents.OnDblClick;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSCapture', 'DblClickEvent', e.Message);
    end;
  end;    
end;

procedure TDSCapture.DeactivateEvent(Sender: TObject);
begin
  try
    if FEvents <> nil then FEvents.OnDeactivate;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSCapture', 'DeactivateEvent', e.Message);
    end;
  end;    
end;

procedure TDSCapture.DestroyEvent(Sender: TObject);
begin
  try
    if FEvents <> nil then FEvents.OnDestroy;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSCapture', 'DestroyEvent', e.Message);
    end;
  end;    
end;

procedure TDSCapture.KeyPressEvent(Sender: TObject; var Key: Char);
var
  TempKey: Smallint;
begin
  try
    TempKey := Smallint(Key);
    if FEvents <> nil then FEvents.OnKeyPress(TempKey);
    Key := Char(TempKey);
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSCapture', 'KeyPressEvent', e.Message);
    end;
  end;    
end;

procedure TDSCapture.PaintEvent(Sender: TObject);
begin
  try
    if FEvents <> nil then FEvents.OnPaint;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSCapture', 'PaintEvent', e.Message);
    end;
  end;    
end;

procedure TDSCapture.Set_AutoScroll(Value: WordBool);
begin
  AutoScroll := Value;
end;

procedure TDSCapture.Set_AutoSize(Value: WordBool);
begin
  AutoSize := Value;
end;

procedure TDSCapture.Set_AxBorderStyle(Value: TxActiveFormBorderStyle);
begin
  AxBorderStyle := TActiveFormBorderStyle(Value);
end;

procedure TDSCapture.Set_Caption(const Value: WideString);
begin
  Caption := TCaption(Value);
end;

procedure TDSCapture.Set_Color(Value: OLE_COLOR);
begin
  Self.Color := TColor(Value);
  _VideoWindow.Color := TColor(Value);  
end;

procedure TDSCapture.Set_DoubleBuffered(Value: WordBool);
begin
  DoubleBuffered := Value;
end;

procedure TDSCapture.Set_DropTarget(Value: WordBool);
begin
  DropTarget := Value;
end;

procedure TDSCapture.Set_Enabled(Value: WordBool);
begin
  Enabled := Value;
end;

procedure TDSCapture.Set_Font(const Value: IFontDisp);
begin
  SetOleFont(Font, Value);
end;

procedure TDSCapture.Set_HelpFile(const Value: WideString);
begin
  HelpFile := String(Value);
end;

procedure TDSCapture.Set_KeyPreview(Value: WordBool);
begin
  KeyPreview := Value;
end;

procedure TDSCapture.Set_PixelsPerInch(Value: Integer);
begin
  PixelsPerInch := Value;
end;

procedure TDSCapture.Set_PrintScale(Value: TxPrintScale);
begin
  PrintScale := TPrintScale(Value);
end;

procedure TDSCapture.Set_Scaled(Value: WordBool);
begin
  Scaled := Value;
end;

procedure TDSCapture.Set_ScreenSnap(Value: WordBool);
begin
  ScreenSnap := Value;
end;

procedure TDSCapture.Set_SnapBuffer(Value: Integer);
begin
  SnapBuffer := Value;
end;

procedure TDSCapture.Set_Visible(Value: WordBool);
begin
  Visible := Value;
end;

procedure TDSCapture.InitDShow;
begin
  _IsPreviewState := False;
  _IsCaptureVideo := false;

  _IsStretch := True;
  _IsAdjustWindowSize := false;
  _IsFit := False;

  _TempCaptureVideoFile := '';
  _IsEscKeyQuitFullScreen := True;
  _IsDblClickQuitFullScreen := False;
  _IsClickQuitFullScreen := False;

  _HideCfgItem := 0;


  _CapParameterWindowPos := cppScreenCenter;
  
  with _FilterGraphic do begin
    Mode := gmCapture;
    GraphEdit    := False;
    LinearVolume := False;
    Active := False;
  end;

  //������ƵԤ������
  _VideoWindow.FilterGraph := _FilterGraphic;
  with _VideoWindow do begin
    Parent := Self;
    Color  := Self.Color;

    Align := alClient;
  end;

  _GraphManager := nil;
  _CapSourceFilter := nil;
  _EncoderFilter := nil;
  _AviMultiplexer := nil;
  _AviWriter := nil;
  _SmartTee := nil;
  _SmartTee1 := nil;
  _ColorSpace := nil;
  
  //_SampleGrabber := nil;
  //_CrabberPin := nil;


  //��ȡ�ɼ�����
  TCaptureParameterConfig.InitCaptureParameter(_CaptureParameter);

  stabStates.Visible := _CaptureParameter.IsShowState;

  _frmCaptureParameterCfg := nil;
end;

procedure TDSCapture.ReInitDShow;
begin
  //ֹͣԤ��
  StopPreview();
  
  if Assigned(_frmCaptureParameterCfg) then FreeAndNil(_frmCaptureParameterCfg);
end;

function TDSCapture.ShowCaptureParameterCfgDialog(
  parentHandle: Integer): WideString;
begin
  try
    Result := '';

    if _IsCaptureVideo then begin
      Result := '���ڽ�����Ƶ�ɼ������ܶ���������á�';
      Exit;
    end;

    if not _FilterGraphic.Active then begin
      Result := 'FilterGraph��δ��ʼ�������ܶ���������á�';
      Exit;
    end;

    //2014-07-22 modify by tjh
    if not Assigned(_frmCaptureParameterCfg) then begin
      //��Ҫÿ�����´����������ô��ڣ���Ϊ��vb��zl9pacscapture�£���������´�����
      //��ڶ��ν�����Ƶ�ɼ�����ʱ�����ڽ�û���κ���Ӧ
      FreeAndNil(_frmCaptureParameterCfg);
      _frmCaptureParameterCfg := nil;
    end;

    //������Ƶ���ô���
    _frmCaptureParameterCfg := TfrmCapParameterCfg.Create(Application);
    _frmCaptureParameterCfg.CapGraphBuilder2 := _GraphManager;
    _frmCaptureParameterCfg.CapSourceFilter := _CapSourceFilter;

    try
      _frmCaptureParameterCfg.InitParameterCfg(_CaptureParameterCfgFileName, _CaptureParameter);
    except
      on e: Exception do begin
        Application.MessageBox(PChar('��ʼ���ɼ�����ʱ�����쳣��������Ϣ��' + e.Message), '��ʾ', MB_OK + MB_ICONINFORMATION);
      end;
    end;

    _frmCaptureParameterCfg.HideParameterCfgItem(_HideCfgItem);
    _frmCaptureParameterCfg.PositionType := _CapParameterWindowPos;
    _frmCaptureParameterCfg.OnParameterChange := ParameterChange;
    _frmCaptureParameterCfg.OnVfwConfigCall := vfwConfigCallEvent;

    try
      _frmCaptureParameterCfg.ShowModal();
    finally
      //2014-07-22 modify by tjh
      //�ͷ����ô��ڶ���
      FreeAndNil(_frmCaptureParameterCfg);
      _frmCaptureParameterCfg := nil;
    end;
    
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;  
end;

procedure TDSCapture.Preview(const isCaptureVideo: Boolean; out errMsg: WideString);
var
  fs: IFileSinkFilter;
  hr: HRESULT;
  iCrossbar: IAMCrossbar;

  //ȡ���׸��豸����
  function GetFirstDeviceName(): WideString;
  var
    capDeviceNames: TStringList;
    hr: HRESULT;
  begin
    Result := '';

    capDeviceNames := TStringList.Create;
    try
      hr := TDS9Ex.GetDeviceNames(CLSID_VideoInputDeviceCategory, capDeviceNames);
      
      if Failed(hr) then Exit;
      if capDeviceNames.Count <= 0 then Exit;

      Result := capDeviceNames[0];
    finally
      FreeAndNil(capDeviceNames);
    end;
  end;

  //ȡ���׸��豸�����ֱ���
  function GetMaxVideoSize(BaseFilter: IBaseFilter): WideString;
  var
    VideoMediaTypes: TEnumMediaType;
    pinList: TPinList;
    formats: String;
    i: Integer;
  begin
    Result := '320X240';

    pinList := TPinList.Create(BaseFilter);
    VideoMediaTypes := TEnumMediaType.Create;
    try
      VideoMediaTypes.Assign(pinList.First);
      for i := 0 to VideoMediaTypes.Count - 1 do begin
        formats := formats + VideoMediaTypes.MediaDescription[i];
      end;

      for i := Length(SysVideoSize) - 1 downto 0 do begin
        if Pos(SysVideoSize[i], formats) > 0 then begin
          Result := SysVideoSize[i];
          Exit;
        end;
      end;
    finally
      FreeAndNil(VideoMediaTypes);
      FreeAndNil(pinList);
    end;
  end;
  
begin
  {Filter ����ͼ���£�

                                           /Capture Pin  --->����¼�����Filter
  CaptureSource>>Signal Pin ---> Smart Tee                             /Capture Pin --->����ͼ��ɼ����Filter
                                           \Preview Pin  --->Smart Tee1
                                                                       \Preview Pin --->������ƵԤ�����Filter

  }
  try
    errMsg := '';

    TDebug.OutputDebug('CAP>>>Preview Step 1');
    
    _FilterGraphic.GraphEdit := _CaptureParameter.DebugFilter;

    if Trim(_CaptureParameter.CaptureDeviceName) = '' then begin

      TDebug.OutputDebug('CAP>>>Preview Step 1.1');
      //ȡ�õ�һ���豸����
      _CaptureParameter.CaptureDeviceName := GetFirstDeviceName();   
      if Trim(_CaptureParameter.CaptureDeviceName) = '' then begin
        errMsg := 'û���ҵ���زɼ��豸����������Ӳ���������Ƿ���ȷ��';
        Exit;
      end;

      TDebug.OutputDebug('CAP>>>Preview Step 1.2');
      //ȡ�õ�һ���豸�����ֱ���...
      hr := TDS9Ex.CreateFilterByDeviceName(CLSID_VideoInputDeviceCategory, _CaptureParameter.CaptureDeviceName, _CapSourceFilter);
      if Failed(hr) then begin
        errMsg := '����CapSourceFilter��ƵԴ�ӿ�ʧ�ܡ� [�豸����:' + _CaptureParameter.CaptureDeviceName + ']  [�������:' + IntToStr(hr) + ']';
        Exit;
      end;

      TDebug.OutputDebug('CAP>>>Preview Step 1.3');
      _CaptureParameter.videoSize := GetMaxVideoSize(_CapSourceFilter);
    end;
                        
    TDebug.OutputDebug('CAP>>>Preview Step 2');

    //���òɼ�����
    if _FilterGraphic.Active then begin
      TDebug.OutputDebug('CAP>>>Preview Step 2.1');
      _IsPreviewState := False;

      _FilterGraphic.Stop;
      _FilterGraphic.ClearGraph; // �ù��̻��Զ��Ͽ���filter������

      _FilterGraphic.Active := False;

      _VideoWindow.FilterGraph := nil;
      _ImgCaptureFilter.FilterGraph := nil;

      TDebug.OutputDebug('CAP>>>Preview Step 2.2');
    end;

    TDebug.OutputDebug('CAP>>>Preview Step 3');

    _FilterGraphic.Active := True;


    //����GraphManager�ӿڶ���,�ö������_GraphBuilder�е�����FILTER
    _GraphManager := nil;
    hr := CoCreateInstance(CLSID_CaptureGraphBuilder2, nil, CLSCTX_INPROC_SERVER, IID_ICaptureGraphBuilder2, _GraphManager);
    if Failed(hr) then begin
      errMsg := '����GraphManager�ӿڹ������ʧ�ܡ� [�������:' + IntToStr(hr) + ']';
      Exit;
    end;

    TDebug.OutputDebug('CAP>>>Preview Step 4');

    //��ʼ��IGraphBuilder�ӿڶ���_GraphBuilder
    hr := _GraphManager.SetFiltergraph(_FilterGraphic as IGraphBuilder);
    if Failed(hr) then begin
      errMsg := '��ʼ��FilterGraphic����ʧ�ܡ�[�������:' + IntToStr(hr) + ']';
      Exit;
    end;

    TDebug.OutputDebug('CAP>>>Preview Step 5');

    //������ƵԴ����CapSourceFilter
    //if not Assigned(_CapSourceFilter) then begin
      _CapSourceFilter := nil;
      hr := TDS9Ex.CreateFilterByDeviceName(CLSID_VideoInputDeviceCategory, _CaptureParameter.CaptureDeviceName, _CapSourceFilter);
      if Failed(hr) then begin
        errMsg := '����CapSourceFilter��ƵԴ�ӿ�ʧ�ܣ�������ƵԴ���á�[�豸����:' + _CaptureParameter.CaptureDeviceName + '] [�������:' + IntToStr(hr) + ']';
        Exit;
      end;
    //end;

    TDebug.OutputDebug('CAP>>>Preview Step 6');

    //���_CapSourceFilter
    hr := (_FilterGraphic as IGraphBuilder).AddFilter(_CapSourceFilter, PWideChar(FILTER_NAME_CAPSOURCE_FILTER));
    if Failed(hr) then begin
      errMsg := '���CapSourceFilter��ƵԴ�ӿڵ�FilterGraphic��ʧ�ܣ�������ƵԴ���á�[�豸����:' + _CaptureParameter.CaptureDeviceName + '] [�������:' + IntToStr(hr) + ']';
      Exit;
    end;

    TDebug.OutputDebug('CAP>>>Preview Step 7');

    //ѡ��ɼ�����(����΢���� v600 �Ĳɼ����ϲ���ͨ��)
    //sdk3000��ʹ�øöδ���ʱ����������������Ƶ  2011-3-17
    //hr := _GraphManager.QueryInterface(IID_IAMCrossbar, iCrossbar);
    iCrossbar := nil;
    hr := _GraphManager.FindInterface(@LOOK_UPSTREAM_ONLY, nil, _CapSourceFilter, IID_IAMCrossbar, iCrossbar);
    if not Failed(hr) then begin
      iCrossbar.Route(_CaptureParameter.OutputCrossbar, _CaptureParameter.InputCrossbar);
      iCrossbar := nil;
    end;//}

    TDebug.OutputDebug('CAP>>>Preview Step 8');

    if not TDS9Ex.IsVfwDevice(_CaptureParameter.CaptureDeviceName) then begin
      TDebug.OutputDebug('CAP>>>Preview Step 8.1');
      //����VFW���豸���򲻽�������

      //����Ƶ���ص�ʱ����Ƶ���������֮ǰ������ֵ���м���

      //������Ƶ��ʽ ����������Ҫ����Ԥ����װ�ز���Ч��
      ConfigVideoAnalog(_CapSourceFilter, _CaptureParameter);

      TDebug.OutputDebug('CAP>>>Preview Step 8.2');
      //������Ƶ��ʽ����������Ҫ����Ԥ����װ�ز���Ч��
      ConfigVideoFormat(_CapSourceFilter, _CaptureParameter);

      TDebug.OutputDebug('CAP>>>Preview Step 8.3');
      //������Ƶ����
      ConfigVideoQuality(_CapSourceFilter, _CaptureParameter);

      TDebug.OutputDebug('CAP>>>Preview Step 8.4');
    end;
    //}

    TDebug.OutputDebug('CAP>>>Preview Step 9');

    //����SampleGrabberͼ��ɼ�����------(ʹ�øýӿڵ�ʱ�򣬲��ܶ�ͼ����в���, ���ʹ��DSPACK�ṩ��TSampleGrabber����)
    {hr := TDS9Ex.AddFilterToGraphBuilder(_FilterGraphic as IGraphBuilder, CLSID_SampleGrabber, _SampleGrabber, FILTER_NAME_SAMPLE_GRABBER);
    if Failed(hr) then begin
      errMsg := '����SampleGrabberͼ��ɼ��ӿ�ʧ�ܡ�';
      Exit;
    end;}

    //����SmartTee�����ӿ�,�����뵽FilterGraphic��------
    //(ע����ʹ��RenderStream����FILTER������ʱ,������ӵ�PREVIEW�˿ڲ����ڣ����Զ�����SmartTee��
    //����Щ�ɼ��豸��Ȼ��preview�Ķ˿ڣ�ȴ����������ݣ�����ִ��filter֮������ӣ���ʹ����SMARTTEE filter������Ϊrenderstream������Դ������
    //��������ΪSmartTee�����Pin���ǲ���PinCategory.Capture��PinCategory.Previewģʽ)
    _SmartTee := nil;
    _SmartTee1 := nil;
    _ColorSpace := nil;

    TDebug.OutputDebug('CAP>>>Preview Step 10');

    hr := TDS9Ex.AddFilterToGraphBuilder(_FilterGraphic as IGraphBuilder, CLSID_SmartTee, _SmartTee, FILTER_NAME_SMART_TEE);
    if Failed(hr) then begin
      errMsg := '����SmartTeeFilter(0)�����ӿ�,�����뵽FilterGraphic��ʱʧ�ܡ�[�������:' + IntToStr(hr) + ']';
      Exit;
    end;

    TDebug.OutputDebug('CAP>>>Preview Step 11');
    //�˶δ���������ʲɼ���PX1000E�����ֲɼ�����Ҫ��������SmartTee1��preview��������ʾ����Ƶ����
    //PX1000E Begin
    hr := TDS9Ex.AddFilterToGraphBuilder(_FilterGraphic as IGraphBuilder, CLSID_SmartTee, _SmartTee1, FILTER_NAME_SMART_TEE1);
    if Failed(hr) then begin
      errMsg := '����SmartTeeFilter(1)�����ӿ�,�����뵽FilterGraphic��ʱʧ�ܡ�[�������:' + IntToStr(hr) + ']';
      Exit;
    end;

    TDebug.OutputDebug('CAP>>>Preview Step 12');
    hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _CapSourceFilter, _SmartTee, False, 0);
    if Failed(hr) then begin
      errMsg := '����CapSourceFilter��SmartTeeFilter(0)֮�������ʱʧ�ܡ�[�������:' + IntToStr(hr) + ']';
      Exit;
    end;

    TDebug.OutputDebug('CAP>>>Preview Step 13');
    //PX1000E End

    //����CapSourceFilter��SmartTee, ���ʹ��RenderStream���������������һ��smartTee�ӿ�
    //�÷��������Ӳɼ�pin
    //hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _CapSourceFilter, _SmartTee, False);
    hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _SmartTee, _SmartTee1, True, 1);
    if Failed(hr) then begin
      errMsg := '����SmartTeeFilter(0)��SmartTeeFilter(1)֮�������ʱʧ�ܡ�[�������:' + IntToStr(hr) + ']';
      Exit;
    end;

    TDebug.OutputDebug('CAP>>>Preview Step 14');

    {hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _SmartTee, _SmartTee1, True, 1);
    if Failed(hr) then begin
      errMsg := '����ColorSpaceConvert��SmartTeeFilter֮�������ʱʧ�ܡ�[�������:' + IntToStr(hr) + ']';
      Exit;
    end;
    //}

    //ok c20a 2010-07-02   ����amcap��RenderStream���ӷ�������filter֮������ӣ�������Щ�ɼ���
    //smarttee Capture Pin��������������Ϊ0��Preview Pin��������������Ϊ1
    //��ok c20a �� micro view v500 ��preview pin�Ͳ������Ԥ������
    {hr := (_FilterGraphic as ICaptureGraphBuilder2).RenderStream(@PIN_CATEGORY_CAPTURE, @MEDIATYPE_Interleaved, _CapSourceFilter, nil, _SmartTee);
    if hr <> NOERROR then begin
      //RenderStream��������------(��ʹ��RenderStream����FILTER������ʱ�����Զ�����������SmartTee)
      hr := (_FilterGraphic as ICaptureGraphBuilder2).RenderStream(@PIN_CATEGORY_CAPTURE, @MEDIATYPE_Video, _CapSourceFilter, nil, _SmartTee);
    end;

    if hr <> NOERROR then begin
      errMsg := '����CapSourceFilter��SmartTee֮�������ʱʧ�ܡ�[�������:' + IntToStr(hr) + ']';
      Exit;
    end;
    //}
            


    //�ж��Ƿ�̬�ɼ���̬��Ƶ
    if isCaptureVideo then begin

      TDebug.OutputDebug('CAP>>>Preview Step 14.1');
      //������Ƶ�������ӿ�
      _EncoderFilter := nil;
      hr := TDS9Ex.CreateFilterByDeviceName(CLSID_VideoCompressorCategory, _CaptureParameter.EncoderName, _EncoderFilter);
      if Failed(hr) then begin
        errMsg := '����EncoderFilter��Ƶ������ʧ�ܣ�������������á�[�������:' + IntToStr(hr) + ']';
        Exit;
      end;

      TDebug.OutputDebug('CAP>>>Preview Step 14.2');
      //���EncoderFilter��Ƶ�������ӿ�
      hr := (_FilterGraphic as IGraphBuilder).AddFilter(_EncoderFilter, PWideChar(FILTER_NAME_ENCODER));
      if Failed(hr) then begin
        errMsg := '���EncoderFilter��Ƶ�������ӿڵ�FilterGraphic��ʧ�ܣ�������������á�[�������:' + IntToStr(hr) + ']';
        Exit;
      end;

      TDebug.OutputDebug('CAP>>>Preview Step 14.3');
      //����MULTIPLEXER��·���ýӿ�
      _AviMultiplexer := nil;
      hr := TDS9Ex.AddFilterToGraphBuilder(_FilterGraphic as IGraphBuilder, CLSID_AviDest, _AviMultiplexer, FILTER_NAME_AVI_MULTIPLEXER);
      if Failed(hr) then begin
        errMsg := '����AviMultiplexerFilter��·���ýӿ�,�����뵽FilterGraphic��ʱʧ�ܡ�[�������:' + IntToStr(hr) + ']';
        Exit;
      end;

      TDebug.OutputDebug('CAP>>>Preview Step 14.4');
      //�����ļ�д��ӿ�
      _AviWriter := nil;
      hr := TDS9Ex.AddFilterToGraphBuilder(_FilterGraphic as IGraphBuilder, CLSID_FileWriter, _AviWriter, FILTER_NAME_AVI_WRITER);
      if Failed(hr) then begin
        errMsg := '����AviWriterFilter�ļ�д��ӿ�,�����뵽FilterGraphic��ʱʧ�ܡ�[�������:' + IntToStr(hr) + ']';
        Exit;
      end;

      TDebug.OutputDebug('CAP>>>Preview Step 14.5');
      //��ѯ��Ƶ�ļ����ýӿ�
      fs := nil;
      hr := _AviWriter.QueryInterface(IID_IFileSinkFilter, fs);
      if Failed(hr)then begin
        errMsg := '��ѯAviWriterFilter��IFileSinkFilter�ӿ�ʱʧ�ܡ�[�������:' + IntToStr(hr) + ']';
        Exit;
      end;

      TDebug.OutputDebug('CAP>>>Preview Step 14.6');
      //������Ƶ�ļ�·��
      hr := fs.SetFileName(PWideChar(_TempCaptureVideoFile), nil);
      if FAILED(hr) then begin
        errMsg := '������Ƶ�ļ�·��ʱʧ�ܡ�[�������:' + IntToStr(hr) + ']';
        Exit;
      end;

      TDebug.OutputDebug('CAP>>>Preview Step 14.7');
      //����SmartTee��EncoderFilter֮�������------��ʹ��RenderStream��ʽ��������_SmartTee������Ϊsource filter������
      //ʹ�õ�һ��smarttee�����capture�˿ڽ������ӣ���Ϊcapture�˿�û��ȡ��ʱ���
      hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _SmartTee, _EncoderFilter, false, 0);
      if Failed(hr) then begin
        errMsg := '����SmartTeeFilter(1)��EncoderFilter֮�����ʱʧ�ܡ�[�������:' + IntToStr(hr) + ']';
        Exit;
      end;

      TDebug.OutputDebug('CAP>>>Preview Step 14.8');
      hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _EncoderFilter, _AviMultiplexer, false, 0);
      if Failed(hr) then begin
        errMsg := '����EncoderFilter��AviMultiplexerFilter֮�����ʱʧ�ܣ�������������á�[�������:' + IntToStr(hr) + ']';
        Exit;
      end;

      TDebug.OutputDebug('CAP>>>Preview Step 14.9');
      //ʹ��RenderStream����FILTER������ʱ,�������Ҫ�Զ����벢����SmartTee���з�������,�������PREVIEW�˿�ʱ��SmartTee�������Զ�����
      //�������ConnectFilters����ʽ�������ӣ�����Ҫ�ֱ���⼸��FILTER��������
      {hr := (_FilterGraphic as ICaptureGraphBuilder2).RenderStream(@PIN_CATEGORY_CAPTURE, @MEDIATYPE_VIDEO, _CapSourceFilter, _EncoderFilter, _AviMultiplexer);
      if Failed(hr) then begin
        errMsg := '����SmartTee��AviMultiplexer֮�����,������EncoderFilter��Ƶ����ӿ�ʱʧ��,�볢������������Ƶ��������';
        Exit;
      end;}

      //����AviMultiplexer��AviWriter֮�������------����Ҫ�Ը����ӽ��д�����ΪAviMultiplexer��AviWriter�����ⲿ������FILTER��
      hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _AviMultiplexer, _AviWriter, false, 0);
      if Failed(hr) then begin
        errMsg := '����AviMultiplexerFilter��AviWriterFilter֮�������ʱʧ�ܡ�[�������:' + IntToStr(hr) + ']';
        Exit;
      end;

      TDebug.OutputDebug('CAP>>>Preview Step 14.10');
    end;  //video capture filter link end...


    TDebug.OutputDebug('CAP>>>Preview Step 15');

    //���ݲ�ͬ��ץͼ��ʽ������ʾģʽ
    if _CaptureParameter.SnatchWay = swVMR then begin
      TDebug.OutputDebug('CAP>>>Preview Step 15.1.1');
      _VideoWindow.Mode := vmVMR;
      _VideoWindow.VMROptions.Mode := vmrWindowless;

      _VideoWindow.FilterGraph := _FilterGraphic;

      {��ֱ��ʹ��vmr9�ӿڻ�ȡͼ��ʱ����Ҫ�ϳ���ʱ�䣬�������_ImgCaptureFilter�ӿڽ���ͼ��ɼ�,
      //�������Ƶ�ط�ĳЩ��ʽ����֧��ʹ��ISampleGrabber�ӿڶ�����вɼ� }
      {_ImgCaptureFilter.FilterGraph := nil;

      //����SmartTee��VideoWindow֮�������(ʹ��vmrģʽ������Ҫ_ImgCaptureFilter����ͼ��Ĳɼ�)
      hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _SmartTee, _VideoWindow as IBaseFilter, true);
      if Failed(hr) then begin
        errMsg := '����SmartTee��VideoWindow֮�����ʱʧ�ܡ�[�������:' + IntToStr(hr) + ']';
        Exit;
      end;//}
      TDebug.OutputDebug('CAP>>>Preview Step 15.1.2');
    end else begin
      TDebug.OutputDebug('CAP>>>Preview Step 15.2.1');
      _VideoWindow.Mode := vmNormal;
      _VideoWindow.VMROptions.Mode := vmrWindowed;

      _VideoWindow.FilterGraph := _FilterGraphic;
      TDebug.OutputDebug('CAP>>>Preview Step 15.2.2');
    end;

    TDebug.OutputDebug('CAP>>>Preview Step 16');

    _ImgCaptureFilter.FilterGraph := _FilterGraphic;


    //����MjpegDescompressѹ���ӿ�
    {_MjpegDescompress := nil;
    hr := TDS9Ex.AddFilterToGraphBuilder(_FilterGraphic as IGraphBuilder, CLSID_MjpegDec, _MjpegDescompress, FILTER_NAME_MJPEGDECOMPRESS);
    if Failed(hr) then begin
      errMsg := '����MJPEGѹ���ӿ�,�����뵽FilterGraphic��ʱʧ�ܡ�[�������:' + IntToStr(hr) + ']';
      Exit;
    end; }



    //����SmartTee��ImgCaptureFilter֮�������  (��ʹ�ô���ģʽʱ����Ҫ����_ImgCaptureFilter)
    //���¾���룬�޸�Ϊʹ�òɼ��źŽţ�ֱ������ͼ�񲶻�_ImgCaptureFilter��filter��
    //hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _SmartTee1, _ImgCaptureFilter  as IBaseFilter, true, 1);

    //��Բ��ֲɼ������ڰ�װһЩ��������smarttee1��imacapturefilter֮����Զ����밲װ��ı���������ɲɼ�ͼ��ʱƫɫ��
    //�����Ҫ�ڴ�֮���ֶ�����color space converter����ת��

    TDebug.OutputDebug('CAP>>>Preview Step 17');
    //����ColorSpaceConverter Filter����Ԥ�������Ͳɼ���ͼ����ɫ����ƫ��
    hr := TDS9Ex.AddFilterToGraphBuilder(_FilterGraphic as IGraphBuilder, CLSID_Colour, _ColorSpace, FILTER_NAME_COLOR_CONVERT);
    if Failed(hr) then begin
      errMsg := '����ColorSpaceConvert��ɫ�ռ�ת���ӿ�,�����뵽FilterGraphic��ʱʧ�ܡ�[�������:' + IntToStr(hr) + ']';
      Exit;
    end;

    TDebug.OutputDebug('CAP>>>Preview Step 18');

    hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _SmartTee1, _ColorSpace, false, 0);
    if Failed(hr) then begin
      errMsg := '����SmartTeeFilter(1)��ColorSpaceConverter֮�������ʱʧ�ܡ�[�������:' + IntToStr(hr) + ']';
      Exit;
    end;

    TDebug.OutputDebug('CAP>>>Preview Step 19');

    hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _ColorSpace, _ImgCaptureFilter  as IBaseFilter, false, 0);
    if Failed(hr) then begin
      errMsg := '����ColorSpaceConverter��ImgCaptureFilter֮�������ʱʧ�ܡ�[�������:' + IntToStr(hr) + ']';
      Exit;
    end;

    TDebug.OutputDebug('CAP>>>Preview Step 20');
    {//���ܳɹ����ӵ�MjpegDescompressFilter,GraphiEdit���Զ����AVICompressorFilter
    hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _MjpegDescompress, _ImgCaptureFilter as IBaseFilter, false);
    if Failed(hr) then begin
      errMsg := '����ImgCaptureFilter��MjpegDescompressFilter֮�������ʱʧ�ܡ�[�������:' + IntToStr(hr) + ']';
      Exit;
    end; //}

    //����ImgCaptureFilter��VideoWindow֮�������
    //����videowindow������ʾʱ��ֱ��ʹ��Ԥ��������ź���ʾ��
    //hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _ImgCaptureFilter as IBaseFilter, _VideoWindow as IBaseFilter, false, 0);
    hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _SmartTee1, _VideoWindow as IBaseFilter, true, 1);
    if Failed(hr) then begin
      errMsg := '����ImgCaptureFilter��VideoWindow֮�������ʱʧ�ܡ�[�������:' + IntToStr(hr) + ']';
      Exit;
    end; //}

    TDebug.OutputDebug('CAP>>>Preview Step 21');
    //RenderStreamʹ����������(������Щ�ɼ��豸��˵��Ȼ�߱�PREVIEW�˿ڣ���ȴ����ʹ�øö˿��������)
    {hr := (_FilterGraphic as ICaptureGraphBuilder2).RenderStream(@PIN_CATEGORY_PREVIEW, @MEDIATYPE_VIDEO, _CapSourceFilter, _ImgCaptureFilter as IBaseFilter, _VideoWindow as IBaseFilter);
    if Failed(hr) then begin
      hr := (_FilterGraphic as ICaptureGraphBuilder2).RenderStream(@PIN_CATEGORY_CAPTURE, @MEDIATYPE_VIDEO, _CapSourceFilter, _ImgCaptureFilter as IBaseFilter, _VideoWindow as IBaseFilter);
      if Failed(hr) then begin
        errMsg := '����CapSourceFilter��ImgCaptureFilter֮�������ʱʧ�ܣ�����Ԥ����Ƶͼ�������豸������˿����͡�';
        Exit;
      end;
    end;}

    _FilterGraphic.Play;

    TDebug.OutputDebug('CAP>>>Preview Step 22');
    imgLogo.Visible := False;

    //��������λ��
    AdjustWindowSize();

    TDebug.OutputDebug('CAP>>>Preview Step 23');
    try
      if Assigned(_frmCaptureParameterCfg) then begin
        //��Ҫ�Բ������ô����еĲɼ����filter���и���
        _frmCaptureParameterCfg.CapGraphBuilder2 := _GraphManager;
        _frmCaptureParameterCfg.CapSourceFilter := _CapSourceFilter;
      end;
    except
      On Ex: Exception do
        TDebug.OutputDebug('CAP>>>Preview Err:' + ex.Message);
    end;

    TDebug.OutputDebug('CAP>>>Preview Step End.');
    _IsPreviewState := True;
  except
    on e: Exception do begin
      errMsg := e.Message;
      TDebug.OutputDebug('CAP>>>Preview Err:' + e.Message);
    end;
  end;
end;

procedure TDSCapture.ParameterChange(
  const capParameter: TCaptureParameter; const needCaptureSample: Boolean);
var
  tmpBitMap: TBitmap;
  previewResult: WideString;
  sErrMsg: WideString;
begin
  try
    //����ɼ��豸Ϊ�գ��򲻽�������
    if Trim(capParameter.CaptureDeviceName) = '' then Exit;

    //�ж��Ƿ����Ԥ��ģʽ
    if not _IsPreviewState then begin
      TCaptureParameterConfig.CopyParameter(capParameter, _CaptureParameter);

      //�����øı��ʱ�����û��Ԥ������ִ�п�ʼԤ��
      if Trim(_CaptureParameter.CaptureDeviceName) <> '' then
        StartPreview();

      Exit;
    end;

    if needCaptureSample then begin
      //�ɼ���Ʒͼ�����ڲü�����
      tmpBitMap := CaptureImageToBmpObj;
      try
        tmpBitMap.SaveToFile(TfrmCapParameterCfg.GetCaptureSampleFile);
      finally
        FreeAndNil(tmpBitMap);
      end;

      Exit;
    end;

    //�жϲ����Ƿ��޸ģ�����޸��������Ƶ��ʾ**************************************

    //���²ɼ��豸
    if capParameter.CaptureDeviceName <> _CaptureParameter.CaptureDeviceName then begin
      _CaptureParameter.CaptureDeviceName := capParameter.CaptureDeviceName;
      Preview(False, previewResult);
    end;

    //ˢ����Ƶ����
    if (capParameter.Brightness <> _CaptureParameter.Brightness)
      or (capParameter.Contrast <> _CaptureParameter.Contrast)
      or (capParameter.Hue <> _CaptureParameter.Hue)
      or (capParameter.Saturation <> _CaptureParameter.Saturation)
      or (capParameter.Gamma <> _CaptureParameter.Gamma)
      or (capParameter.WhiteBlance <> _CaptureParameter.WhiteBlance)
      or (capParameter.IsAutoBrightness <> _CaptureParameter.IsAutoBrightness)
      or (capParameter.IsAutoContrast <> _CaptureParameter.IsAutoContrast)
      or (capParameter.IsAutoHue <> _CaptureParameter.IsAutoHue)
      or (capParameter.IsAutoGamma <> _CaptureParameter.IsAutoGamma)
      or (capParameter.IsAutoSaturation <> _CaptureParameter.IsAutoSaturation)
      or (capParameter.IsAutoWhiteBlance <> _CaptureParameter.IsAutoWhiteBlance) then begin
      
      _CaptureParameter.Brightness  := capParameter.Brightness;
      _CaptureParameter.Contrast    := capParameter.Contrast;
      _CaptureParameter.Hue         := capParameter.Hue;
      _CaptureParameter.Saturation  := capParameter.Saturation;
      _CaptureParameter.Gamma       := capParameter.Gamma;
      _CaptureParameter.WhiteBlance := capParameter.WhiteBlance;

      _CaptureParameter.IsAutoBrightness  := capParameter.IsAutoBrightness;
      _CaptureParameter.IsAutoContrast    := capParameter.IsAutoContrast;
      _CaptureParameter.IsAutoHue         := capParameter.IsAutoHue;
      _CaptureParameter.IsAutoGamma       := capParameter.IsAutoGamma;
      _CaptureParameter.IsAutoSaturation  := capParameter.IsAutoSaturation;
      _CaptureParameter.IsAutoWhiteBlance := capParameter.IsAutoWhiteBlance;

      //�����VFW�豸���򲻽��е���
      if not TDS9Ex.IsVfwDevice(capParameter.CaptureDeviceName) then
        ConfigVideoQuality(_CapSourceFilter, _CaptureParameter);
    end;

    //ˢ����Ƶ��ʽ
    if (capParameter.VideoSize <> _CaptureParameter.VideoSize) then begin
      _CaptureParameter.VideoSize := capParameter.VideoSize;

      Preview(False, previewResult);
      //if FEvents <> nil then FEvents.OnResize;//��VB�У����е�ʱ�򣬸��¼��ſ��Ա���������ִ��
      VideoSizeEvent(_VideoWindow.Width, _VideoWindow.Height, Self.Width, Self.Height);
    end;

    //ˢ����ɫ���
    if (capParameter.ColorDepth <> _CaptureParameter.ColorDepth) then begin
      _CaptureParameter.ColorDepth := capParameter.ColorDepth;
      Preview(false, previewResult);
    end;

    //ˢ����Ƶ��ʽ
    if capParameter.VideoAnalog <> _CaptureParameter.VideoAnalog then begin
      _CaptureParameter.VideoAnalog := capParameter.VideoAnalog;
      Preview(false, previewResult);
    end;

    //ˢ����ʾģʽ
    if capParameter.VideoShowModel <> _CaptureParameter.VideoShowModel then begin
      _CaptureParameter.VideoShowModel := capParameter.VideoShowModel;
      AdjustWindowSize();
    end;

    //ˢ��ͼ��ץȡģʽ
    if capParameter.SnatchWay <> _CaptureParameter.SnatchWay then begin
      _CaptureParameter.SnatchWay := capParameter.SnatchWay;

      if _IsPreviewState and not _IsCaptureVideo then begin
        //���¿�ʼԤ��
        Preview(False, sErrMsg);
      end;
    end;

    //ˢ������˿�
    if capParameter.InputCrossbar <> _CaptureParameter.InputCrossbar then begin
      _CaptureParameter.InputCrossbar := capParameter.InputCrossbar;
      if _IsPreviewState and not _IsCaptureVideo then begin
        //���¿�ʼԤ��
        Preview(False, sErrMsg);      
      end;
    end;

    //ˢ������˿�
    if capParameter.OutputCrossbar <> _CaptureParameter.OutputCrossbar then begin
      _CaptureParameter.OutputCrossbar := capParameter.OutputCrossbar;
      if _IsPreviewState and not _IsCaptureVideo then begin
        //���¿�ʼԤ��
        Preview(False, sErrMsg);      
      end;
    end;


    //��Ƶ״̬��ʾ����
    if capParameter.IsShowState <> _CaptureParameter.IsShowState then begin
      _CaptureParameter.IsShowState := capParameter.IsShowState;

      stabStates.Visible := capParameter.IsShowState;
      AdjustWindowSize();
    end;

    //������������Ҫˢ�µ�ǰ��Ƶ��ʾ�Ĳ���
    TCaptureParameterConfig.CopyParameter(capParameter, _CaptureParameter);
  except
  end;
end;

procedure TDSCapture.ConfigVideoAnalog(filter: IBaseFilter;
  captureParameter: TCaptureParameter);
var
  curBaseFilter: IBaseFilter;
  hr: HRESULT;
  amAnalogVideoDecoder: IAMAnalogVideoDecoder;
begin
  curBaseFilter := filter;//(filter as IBaseFilter);
  if not Assigned(curBaseFilter) then Exit;

  hr := curBaseFilter.QueryInterface(IID_IAMAnalogVideoDecoder, amAnalogVideoDecoder);
  if Succeeded(hr) then begin
    amAnalogVideoDecoder.put_TVFormat(TCaptureParameterConfig.ConvertAnalogVideoStandard(captureParameter.VideoAnalog));
    amAnalogVideoDecoder := nil;
  end;

  curBaseFilter := nil;
end;

procedure TDSCapture.ConfigVideoFormat(filter: IBaseFilter;
  captureParameter: TCaptureParameter);
var
  curBaseFilter: IBaseFilter;
  curVideoSize: TVideoSize;
begin
  curBaseFilter := filter; //filter�����޸�ΪTFilter��
  if not Assigned(curBaseFilter) then Exit;

  //
  curVideoSize := TCaptureParameterConfig.ConvertVideoSizeInf(captureParameter.VideoSize);

  TDS9Ex.ConfigCaptureScale(curBaseFilter, curVideoSize.Width, curVideoSize.Height, captureParameter.ColorDepth);
end;

procedure TDSCapture.FreeRes;
begin
  try           
    ReInitDShow();
  except
  end;  
end;

procedure TDSCapture.ConfigVideoQuality(filter: IBaseFilter;
  captureParameter: TCaptureParameter);
var
  curBaseFilter: IBaseFilter;
  amVideoProcAmp: IAMVideoProcAmp;
  amCameraControl: IAMCameraControl;
  hr: HRESULT;
  isOk: Boolean;

  //ȡ���������������Ϣ
  function GetDefaultQualityInf(curAmVideoProcAmp: IAMVideoProcAmp;
    PropertyTag : TVideoProcAmpProperty; var isSucceed: Boolean): Integer;
  var
    curHr: HRESULT;
    iMinValue, iMaxValue, iStep, iCurValue, iDefault: Integer;
    iFlags : TVideoProcAmpFlags;
  begin
    try
      isSucceed := True;

      TDebug.OutputDebug('CAP>>>GetDefaultQualityInf 1');
      //ȡ����Ƶ�������õķ�Χ
      curHr := curAmVideoProcAmp.GetRange(PropertyTag, iMinValue, iMaxValue, iStep, iDefault, iFlags);
      if not Succeeded(curHr) then begin
        isSucceed := False;
        Result := 0;
        TDebug.OutputDebug('CAP>>>GetDefaultQualityInf 1.1 Return False');
        Exit;
      end;

      TDebug.OutputDebug('CAP>>>GetDefaultQualityInf 2');
      //ȡ�õ�ǰֵ
      curHr := curAmVideoProcAmp.Get(PropertyTag, iCurValue, iFlags);
      if not Succeeded(curHr) then begin
        isSucceed := False;
        Result := 0;
        TDebug.OutputDebug('CAP>>>GetDefaultQualityInf 2.1 Return False');
        Exit;
      end;


      TDebug.OutputDebug('CAP>>>GetDefaultQualityInf End Return True.');
      Result := iDefault;
    except
      On Ex: Exception do begin
        isSucceed := False;
        Result := 0;
        TDebug.OutputDebug('CAP>>>GetDefaultQualityInf Err:' + Ex.Message);
      end;
    end;
  end;
begin
  curBaseFilter := filter;

  TDebug.OutputDebug('CAP>>>ConfigVideoQuality 1');
  
  if not Assigned(curBaseFilter) then Exit;
  if Trim(captureParameter.CaptureDeviceName) = '' then Exit;


  if captureParameter.ExposureWay <> 0 then begin
    TDebug.OutputDebug('CAP>>>ConfigVideoQuality 1.5');
    hr := curBaseFilter.QueryInterface(IID_IAMCameraControl, amCameraControl);
    if Succeeded(hr) then begin

      //typedef enum  {
      //  CameraControl_Flags_Auto    = 0x0001,
      //  CameraControl_Flags_Manual  = 0x0002
      //} CameraControlFlags;

      //tagCameraControlFlags�ж����ֵ��΢�������෴���ο�����ע�Ͳ���
      amCameraControl.Set_(CameraControl_Exposure, captureParameter.ExposureValue, TCameraControlFlags(captureParameter.ExposureWay));
      TDebug.OutputDebug('CAP>>>ConfigVideoQuality 1.5 auto Exposure');
    end;
  end;

  TDebug.OutputDebug('CAP>>>ConfigVideoQuality 2');
  hr := curBaseFilter.QueryInterface(IID_IAMVideoProcAmp, amVideoProcAmp);
  if not Succeeded(hr) then Exit;

  //˵������������directshow��VideoProcAmp_Flags_Auto��ʾ�ֶ�����  VideoProcAmp_Flags_Manual��ʾ�Զ�����������������Ϊֵ�������

  isOk := False;
  
  TDebug.OutputDebug('CAP>>>ConfigVideoQuality 3');
  //����
  try
    if captureParameter.Brightness < 0 then
      captureParameter.Brightness := GetDefaultQualityInf(amVideoProcAmp, VideoProcAmp_Brightness, isOk);

    if isOk and (captureParameter.Brightness >= 0) then begin
      TDebug.OutputDebug('CAP>>>ConfigVideoQuality 3.1');
      amVideoProcAmp.Set_(
        VideoProcAmp_Brightness, captureParameter.Brightness,
        tagVideoProcAmpFlags(IfThen(captureParameter.IsAutoBrightness, Integer(VideoProcAmp_Flags_Manual), Integer(VideoProcAmp_Flags_Auto)))
        );
    end;
  except
    On Ex: Exception do
      TDebug.OutputDebug('CAP>>>ConfigVideoQuality Err:' + Ex.Message);
  end;

  TDebug.OutputDebug('CAP>>>ConfigVideoQuality 4');
  //�Աȶ�
  try
    if captureParameter.Contrast < 0 then
      captureParameter.Contrast := GetDefaultQualityInf(amVideoProcAmp, VideoProcAmp_Contrast, isOk);

    if isOk and (captureParameter.Contrast >= 0) then begin
      TDebug.OutputDebug('CAP>>>ConfigVideoQuality 4.1');
      amVideoProcAmp.Set_(
        VideoProcAmp_Contrast, captureParameter.Contrast,
        tagVideoProcAmpFlags(IfThen(captureParameter.IsAutoContrast, Integer(VideoProcAmp_Flags_Manual), Integer(VideoProcAmp_Flags_Auto)))
        );
    end;
  except
    On Ex: Exception do
      TDebug.OutputDebug('CAP>>>ConfigVideoQuality Err:' + Ex.Message);
  end;

  TDebug.OutputDebug('CAP>>>ConfigVideoQuality 5');
  //ɫ��
  try
    if captureParameter.Hue < 0 then
      captureParameter.Hue := GetDefaultQualityInf(amVideoProcAmp, VideoProcAmp_Hue, isOk);

    if isOk and (captureParameter.Hue >= 0) then begin
      TDebug.OutputDebug('CAP>>>ConfigVideoQuality 5.1');
      amVideoProcAmp.Set_(
        VideoProcAmp_Hue, captureParameter.Hue,
        tagVideoProcAmpFlags(IfThen(captureParameter.IsAutoHue, Integer(VideoProcAmp_Flags_Manual), Integer(VideoProcAmp_Flags_Auto)))
        );
    end;
  except
    On Ex: Exception do
      TDebug.OutputDebug('CAP>>>ConfigVideoQuality Err:' + Ex.Message);
  end;    

  TDebug.OutputDebug('CAP>>>ConfigVideoQuality 6');
  //���Ͷ�
  try
    if captureParameter.Saturation < 0 then
      captureParameter.Saturation := GetDefaultQualityInf(amVideoProcAmp, VideoProcAmp_Saturation, isOk);

    if isOk and (captureParameter.Saturation >= 0) then begin
      TDebug.OutputDebug('CAP>>>ConfigVideoQuality 6.1');
      amVideoProcAmp.Set_(
        VideoProcAmp_Saturation, captureParameter.Saturation,
        tagVideoProcAmpFlags(IfThen(captureParameter.IsAutoSaturation, Integer(VideoProcAmp_Flags_Manual), Integer(VideoProcAmp_Flags_Auto)))
        );
    end;
  except
    On Ex: Exception do
      TDebug.OutputDebug('CAP>>>ConfigVideoQuality Err:' + Ex.Message);
  end;

  TDebug.OutputDebug('CAP>>>ConfigVideoQuality 7');
  //٤��
  try
    if captureParameter.Gamma < 0 then
      captureParameter.Gamma := GetDefaultQualityInf(amVideoProcAmp, VideoProcAmp_Gamma, isOk);

    if isOk and (captureParameter.Gamma >= 0) then begin
      TDebug.OutputDebug('CAP>>>ConfigVideoQuality 7.1');
      amVideoProcAmp.Set_(
        VideoProcAmp_Gamma, captureParameter.Gamma,
        tagVideoProcAmpFlags(IfThen(captureParameter.IsAutoGamma, Integer(VideoProcAmp_Flags_Manual), Integer(VideoProcAmp_Flags_Auto)))
        );
    end;
  except
    On Ex: Exception do
      TDebug.OutputDebug('CAP>>>ConfigVideoQuality Err:' + Ex.Message);
  end;    

  TDebug.OutputDebug('CAP>>>ConfigVideoQuality 8');
  //��ƽ��
  try
    if captureParameter.WhiteBlance < 0 then
      captureParameter.WhiteBlance := GetDefaultQualityInf(amVideoProcAmp, VideoProcAmp_WhiteBalance, isOk);

    if isOk and (captureParameter.WhiteBlance >= 0) then begin
      TDebug.OutputDebug('CAP>>>ConfigVideoQuality 8.1');
      amVideoProcAmp.Set_(
        VideoProcAmp_WhiteBalance, captureParameter.WhiteBlance,
        tagVideoProcAmpFlags(IfThen(captureParameter.IsAutoWhiteBlance, Integer(VideoProcAmp_Flags_Manual), Integer(VideoProcAmp_Flags_Auto)))
        );
    end;
  except
    On Ex: Exception do
      TDebug.OutputDebug('CAP>>>ConfigVideoQuality Err:' + Ex.Message);
  end;


  TDebug.OutputDebug('CAP>>>ConfigVideoQuality End.');
  curBaseFilter := nil;
end;

function TDSCapture.CaptureBmpImageToFile(
  const fileName: WideString): WideString;
var
  bitMap, cutBmp: TBitmap;
  cutArea: TRect;
begin
  try
    Result := '';

    if not _IsPreviewState then begin
      Result := '��Ƶ�ɼ���δ����Ԥ��ģʽ��';
      Exit;
    end;

    bitMap := CaptureImageToBmpObj;
    try
      if not Assigned(bitMap) then begin
        Result := 'û�вɼ�����Ƶͼ��';
        Exit;
      end;

      //ת��Ϊ�Ҷ�ͼ
      //�ɼ���ͼ��һ�㶼��1024*768���ڣ����Լ����ǲ������ü����лҶ�ת����
      //��Ч����Ҳû�ж���Ӱ��
      if _CaptureParameter.IsConvertGrayImg then begin
        TGraphicProcess.ConvertBitmapToGrayscale(bitMap);
      end;


      //ͼ��ü�����
      if _CaptureParameter.IsApplyImageCut then begin
        cutBmp := TBitmap.Create;
        try
          //ȡ�òü���Χ
          cutArea.Left := Round(_CaptureParameter.LeftRate * bitMap.Width);
          cutArea.Right := cutArea.Left + Round(_CaptureParameter.WidthRate * bitMap.Width);

          cutArea.Top := Round(_CaptureParameter.TopRate * bitMap.Height);
          cutArea.Bottom := cutArea.Top + Round(_CaptureParameter.HeightRate * bitMap.Height);

          TGraphicProcess.CutImg(cutArea, bitMap, cutBmp);

          cutBmp.SaveToFile(fileName);
        finally
          FreeAndNil(cutBmp);
        end;
      end else begin
        //ֱ�ӱ���ɼ�ͼ��
        bitMap.SaveToFile(fileName);
      end;  
    finally
      FreeAndNil(bitMap);
    end;
  except
    on e: Exception do begin
      Result := e.Message;
    end;  
  end;
end;

function TDSCapture.StartCaptureVideo(
  const fileName: WideString): WideString;
var
  fileGuid: TGUID;
  curFile: String;
  sVideoFile: WideString;
begin
  Result := '';

  if not _IsPreviewState then begin
    Result := '��Ƶ�ɼ���δ����Ԥ��ģʽ��';
    Exit;
  end;

  if _IsCaptureVideo then begin
    Result := '���ڽ�����Ƶ�ɼ�������ִ�иò�����';
    Exit;
  end;

  try
    //������Ƶ�ļ�����λ��
    _TempCaptureVideoFile := fileName;
    if Trim(fileName) = '' then begin
      CreateGUID(fileGuid);
      curFile := GUIDToString(fileGuid) + '.avi';
      curFile := StringReplace(curFile, '-', '', [rfReplaceAll, rfIgnoreCase]);

      _TempCaptureVideoFile := ExtractFilePath(Application.ExeName) + CONST_TEMP_DIR + curFile;

      //���Ŀ¼�����ڣ��򴴽���Ŀ¼
      if not DirectoryExists(ExtractFilePath(Application.ExeName) + CONST_TEMP_DIR) then begin
        ForceDirectories(ExtractFilePath(Application.ExeName) + CONST_TEMP_DIR);
      end;
    end;

    //��ʼ��̬��Ƶ�ɼ�
    _IsCaptureVideo := True;
    _RecordTimeLen := 0;

    Preview(True, Result);

    if Trim(Result) <> '' then begin
      StopCaptureVideo(sVideoFile);

      Exit;
    end;

    Timer1.Enabled := True;
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

function TDSCapture.StopCaptureVideo(
  out videoFile: WideString): WideString;
var
  ms: IMediaSeeking;
  stopPos: Int64;
  hr: HRESULT;
begin
  try
    Result := '';
    if not _IsCaptureVideo then Exit; //û�п�ʼ��Ƶ�ɼ�����ִ�иò����� 

    Timer1.Enabled := False;
    videoFile := _TempCaptureVideoFile;

    stabStates.Panels.Items[4].Text := '';
    _TempCaptureVideoFile := '';

    _IsCaptureVideo := False;

    hr := (_FilterGraphic as IGraphBuilder).QueryInterface(IID_IMediaSeeking, ms);
    if not Failed(hr) then begin

      ms.GetCurrentPosition(stopPos);
      _RecordTimeLen := Round(stopPos / ONE_SECOND);

      ms := nil;
    end;

    Preview(false, Result);
  except
    on e: Exception do begin
      videoFile := '';
      Result := e.Message;
    end;
  end;
end;

procedure TDSCapture.Timer1Timer(Sender: TObject);
var
  position: int64;
  Hour, Min, Sec, MSec: Word;
  sVideoFile: WideString;
begin
  try
    if _FilterGraphic.Active then begin
      with _FilterGraphic as IMediaSeeking do
        GetCurrentPosition(position);

      DecodeTime(position div 10000 / MiliSecInOneDay, Hour, Min, Sec, MSec);
      stabStates.Panels.Items[4].Text := Format('%d:%d:%d:%d',[Hour, Min, Sec, MSec]);

      if (_CaptureParameter.IsTimeLimit and ((Hour * 3600 + Min * 60 + Sec) >= _CaptureParameter.LimitLength))
        or ((Min * 60 + Sec) >= 3600 * 12) then begin
        //�ɼ�ʱ������(���ܳ���ָ��ʱ��),���δ�������ã���������8Сʱ��3600 * 8��
        StopCaptureVideo(sVideoFile);
      end;
    end;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSCapture', 'Timer1Timer', e.Message);
    end;
  end;
end;

function TDSCapture.Get_IsStretch: WordBool;
begin
  Result := _CaptureParameter.VideoShowModel = smStretch;
end;

procedure TDSCapture.Set_IsStretch(Value: WordBool);
begin
  _IsStretch := Value;

  if (_IsStretch = False) and (_IsFit = False) and (_IsAdjustWindowSize = False) then _CaptureParameter.VideoShowModel := smNormal;
  if (_IsStretch = False) and (_IsFit = False) and (_IsAdjustWindowSize = True) then _CaptureParameter.VideoShowModel := smWindAutoFit;
  if (_IsStretch = False) and (_IsFit = True) and (_IsAdjustWindowSize = True) then _CaptureParameter.VideoShowModel := smFit;
  if (_IsStretch = False) and (_IsFit = True) and (_IsAdjustWindowSize = False) then _CaptureParameter.VideoShowModel := smFit;

  if _IsStretch then begin
    _CaptureParameter.VideoShowModel := smStretch;
    _IsFit := False;
    _IsAdjustWindowSize := False;
  end;

  //AdjustWindowSize();
end;

function TDSCapture.Get_IsShowState: WordBool;
begin
  Result := stabStates.Visible;
end;

procedure TDSCapture.Set_IsShowState(Value: WordBool);
begin
  _CaptureParameter.IsShowState := Value;
  stabStates.Visible := Value;

  //��Ҫ������Ƶ��ʾλ��   
  AdjustWindowSize();
end;

procedure TDSCapture._VideoWindowKeyPress(Sender: TObject; var Key: Char);
begin
  if Assigned(Self.OnKeyPress) then Self.OnKeyPress(Sender, Key);
end;

procedure TDSCapture.timerSysTimer(Sender: TObject);
begin
  try
    if not stabStates.Visible then Exit;
    
    stabStates.Panels.Items[1].Text := TimeToStr(Now);

    if not _IsCaptureVideo then begin
      stabStates.Panels.Items[3].Text := 'Ԥ��ģʽ';
    end else begin
      stabStates.Panels.Items[3].Text := '¼��ģʽ';
    end;

    if not _IsPreviewState then begin
      stabStates.Panels.Items[3].Text := '����ģʽ';
    end;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSCapture', 'timerSysTimer', e.Message);
    end;
  end;
end;

function TDSCapture.Get_IsFullScreen: WordBool;
begin
  Result := TfrmFullScreen.GetFullScreenState();
end;

procedure TDSCapture.Set_IsFullScreen(Value: WordBool);
begin
  if Value then begin
    ShowFullScreen(Handle, 0)
  end else begin
    QuitFullScreen();
  end;  
end;

procedure TDSCapture.AdjustWindowSize;
const
  CONTROL_BORDER_SPACE: Integer = 2;
var
  curVideoSizeInf: TVideoSize;
  zoomRate, wCutRate, hCutRate: Double;
  stateBarHeight: Integer;
begin
  //������ʾģʽ���ͣ��������λ�ü���С
  case _CaptureParameter.VideoShowModel of
    smNormal: begin //---------------------------------------------------------
      _VideoWindow.Align := alNone;
      curVideoSizeInf := TCaptureParameterConfig.ConvertVideoSizeInf(_CaptureParameter.VideoSize);      

      _VideoWindow.Width := curVideoSizeInf.Width;
      _VideoWindow.Height := curVideoSizeInf.Height;

      _VideoWindow.Left := (Self.Width - _VideoWindow.Width) div 2 - 1;

      //�ж��Ƿ���Ҫ��ȥ״̬���߶�
      stateBarHeight := 0;
      if stabStates.Visible then stateBarHeight := stabStates.Height;

      if curVideoSizeInf.Height > Self.Height - stateBarHeight then begin
        //_VideoWindow.Top := Self.Height - curVideoSizeInf.Height - stateBarHeight - 2;
        _VideoWindow.Top := (Self.Height - curVideoSizeInf.Height - stateBarHeight) div 2 - 1;
      end else begin
        _VideoWindow.Top := (Self.Height - stateBarHeight - _VideoWindow.Height) div 2 - 1;
      end;
    end;
    smFit: begin //-------------------------------------------------------------
      _VideoWindow.Align := alNone;    
      curVideoSizeInf := TCaptureParameterConfig.ConvertVideoSizeInf(_CaptureParameter.VideoSize);


      //�ж��Ƿ���Ҫ��ȥ״̬���߶�
      stateBarHeight := 0;
      if stabStates.Visible then stateBarHeight := stabStates.Height;

      //ȡ�����ű���
      if (curVideoSizeInf.Height) / curVideoSizeInf.Width > (Self.Height - stateBarHeight) / (Self.Width) then begin
        zoomRate := (Self.Height - stateBarHeight) / curVideoSizeInf.Height;
      end else begin
        zoomRate := Self.Width / curVideoSizeInf.Width;
      end;

      //�����С��ȣ��򲻽�������
      if (curVideoSizeInf.Height = Self.Height - stateBarHeight - CONTROL_BORDER_SPACE)
        and (curVideoSizeInf.Width = Self.Width - CONTROL_BORDER_SPACE) then begin
        zoomRate := 1;
      end;

      _VideoWindow.Width := Round(curVideoSizeInf.Width * zoomRate);
      _VideoWindow.Height := Round(curVideoSizeInf.Height * zoomRate) ;


      _VideoWindow.Left := (Self.Width - _VideoWindow.Width) div 2 - 1;
      _VideoWindow.Top := (Self.Height - stateBarHeight - _VideoWindow.Height) div 2 - 1;
    end;
    smStretch: begin //---------------------------------------------------------
      _VideoWindow.Align := alClient;
    end;
    smAutoFitCut: begin  //-----------------------------------------------------
      _VideoWindow.Align := alNone;
      curVideoSizeInf := TCaptureParameterConfig.ConvertVideoSizeInf(_CaptureParameter.VideoSize);
      
      if (_CaptureParameter.widthRate <= 0) then begin
        wCutRate := 1;
      end else begin
        wCutRate := _CaptureParameter.widthRate;
      end;

      if (_CaptureParameter.heightRate <= 0) then begin
        hCutRate := 1;
      end else begin
        hCutRate := _CaptureParameter.heightRate;
      end;

      Self.Width := round(curVideoSizeInf.Width * wCutRate) + CONTROL_BORDER_SPACE;

      stateBarHeight := 0;
      if stabStates.Visible then stateBarHeight := stabStates.Height;

      Self.Height := round(curVideoSizeInf.Height * hCutRate) + stateBarHeight + CONTROL_BORDER_SPACE;


      _VideoWindow.Left := 0 - Round(curVideoSizeInf.Width * _CaptureParameter.leftRate);
      _VideoWindow.Top := 0 - Round(curVideoSizeInf.Height * _CaptureParameter.topRate);

      _VideoWindow.Width := curVideoSizeInf.Width;
      _VideoWindow.Height := curVideoSizeInf.Height;
    end;
    smWindAutoFit: begin //-----------------------------------------------------
      //�жϵ�ǰ���ڴ�С�Ƿ���Ӧ�ɼ����ڴ�С
      curVideoSizeInf := TCaptureParameterConfig.ConvertVideoSizeInf(_CaptureParameter.VideoSize);

      //�ж��Ƿ���Ҫ��ȥ״̬���߶�
      stateBarHeight := 0;
      if stabStates.Visible then stateBarHeight := stabStates.Height;

      Self.Height := curVideoSizeInf.Height + stateBarHeight + CONTROL_BORDER_SPACE;
      Self.Width := curVideoSizeInf.Width + CONTROL_BORDER_SPACE;

      _VideoWindow.Width := curVideoSizeInf.Width;
      _VideoWindow.Height := curVideoSizeInf.Height;

      _VideoWindow.Left := (Self.Width - _VideoWindow.Width) div 2 - 1;

      if _VideoWindow.Height > Self.Height - stateBarHeight then begin
        _VideoWindow.Top := Self.Height - _VideoWindow.Height - stateBarHeight;
      end else begin
        _VideoWindow.Top := (Self.Height - stateBarHeight - _VideoWindow.Height) div 2 - 1;
      end;
    end;
  end;

  if imgLogo.Visible then begin
    imgLogo.Left := (Self.Width - imgLogo.Width) div 2;
    imgLogo.Top := (Self.Height - imgLogo.Height) div 2;
  end;

end;

function TDSCapture.Get_IsAdjustWindowSize: WordBool;
begin
  Result :=  _CaptureParameter.VideoShowModel = smWindAutoFit;
end;

procedure TDSCapture.Set_IsAdjustWindowSize(Value: WordBool);
begin
  //if _IsAdjustWindowSize = Value then Exit; 
  _IsAdjustWindowSize := Value;

  if (_IsAdjustWindowSize = False) and (_IsFit = False) and (_IsStretch = False) then _CaptureParameter.VideoShowModel := smNormal;
  if (_IsAdjustWindowSize = False) and (_IsFit = False) and (_IsStretch = True) then _CaptureParameter.VideoShowModel := smStretch;
  if (_IsAdjustWindowSize = False) and (_IsFit = True) and (_IsStretch = False) then _CaptureParameter.VideoShowModel := smFit;
  if (_IsAdjustWindowSize = False) and (_IsFit = True) and (_IsStretch = True) then _CaptureParameter.VideoShowModel := smFit;

  if Value then begin
    _CaptureParameter.VideoShowModel := smWindAutoFit;
    _IsFit := False;
    _IsStretch := False;
  end;  

  //AdjustWindowSize();
end;

function TDSCapture.StopPreview: WideString;
begin
  Result := '';
  
  if _IsCaptureVideo then begin
    Result := '���ڲɼ���Ƶ������ֹͣԤ����';
    Exit;
  end;

  try
    //ֹͣ��ƵԤ�������밴��������˳��ִ�У�
    if _FilterGraphic.Active then begin
                        
      _FilterGraphic.Stop;

      _FilterGraphic.ClearGraph; // �ù��̻��Զ��Ͽ���filter������

      _FilterGraphic.Active := False;
    end;

    _GraphManager := nil;

    _IsPreviewState := False;
    imgLogo.Visible := True;

    //�÷����ᴥ��videowindow ��paint�¼��������Ҫ����_IsPreviewState := False���֮��
    _VideoWindow.Refresh;
  except
    on e: Exception do begin
      Result := e.Message;
      
      TDebug.DebugMsg('TDSCapture', 'StopPreview', e.Message )
    end;    
  end;  
end;

procedure TDSCapture._VideoWindowClick(Sender: TObject);
begin
  //�˳�ȫ��
  if _IsClickQuitFullScreen {and _VideoWindow.FullScreen} then begin
    QuitFullScreen();
  end;
    
  if Assigned(Self.OnClick) then Self.OnClick(Sender);
end;

procedure TDSCapture._VideoWindowDblClick(Sender: TObject);
begin
  //�˳�ȫ��
  if _IsDblClickQuitFullScreen {and _VideoWindow.FullScreen} then begin
    QuitFullScreen();
  end;

  if Assigned(Self.OnDblClick) then Self.OnDblClick(Sender);
end;

function TDSCapture.CaptureImageToBmpObj: TBitmap;
var
  //bmpSize: Longint;
  curBitMap: TBitmap;
  bmpStream: TMemoryStream;
  //stick, etick: Cardinal;

  procedure UseVmrCapture();
  begin
    if (_CaptureParameter.SnatchWay = swVMR) and _VideoWindow.VMRGetBitmap(bmpStream) then begin
      //etick := GetTickCount;
      //ShowMessage(IntToStr(etick - stick ));
      curBitMap.LoadFromStream(bmpStream);

      Result := curBitMap;

      case _CaptureParameter.colorDepth of
        4: Result.PixelFormat := pf4bit;
        8: Result.PixelFormat := pf8bit;
        12: Result.PixelFormat := pf24bit;//pfDevice;
        16: Result.PixelFormat := pf24bit;//pf16bit;
        24: Result.PixelFormat := pf24bit;
        32: Result.PixelFormat := pf32bit;
      end;
    end;
  end;
begin
  if not _IsPreviewState then begin
    Result := nil;
    Exit;
  end;
  //stick := GetTickCount;

  Result := nil;
  curBitMap := TBitmap.Create;
  try
    bmpStream := TMemoryStream.Create;
    try
      //�ɼ���֡ͼ��     ȡ��ָ��ͼ�����ת��Ϊָ��λ����ͼ��dshow�ɼ�ʱ��ͼ��λ�����õ�Ϊ32λ
      {if (_CaptureParameter.SnatchWay = swVMR) and _VideoWindow.VMRGetBitmap(bmpStream) then begin
      //etick := GetTickCount;
      //ShowMessage(IntToStr(etick - stick ));
        curBitMap.LoadFromStream(bmpStream);

        Result := curBitMap;

        case _CaptureParameter.colorDepth of
          4: Result.PixelFormat := pf4bit;
          8: Result.PixelFormat := pf8bit;
          12: Result.PixelFormat := pf24bit;//pfDevice;
          16: Result.PixelFormat := pf24bit;//pf16bit;
          24: Result.PixelFormat := pf24bit;
          32: Result.PixelFormat := pf32bit;
        end;

        //Ϊ���ⷢ��ʱ��ռ�ù���ʱ�䣬���ʹ��postmessage��������������ʾ
        if _CaptureParameter.IsSoundHint then PostMessage(Self.Handle, WM_BEEPMSG, 0, 0);

        Exit;
      end;//}


      if {(_CaptureParameter.SnatchWay = swDEVICE) and} _ImgCaptureFilter.GetBitmap(curBitMap) then begin
        Result := curBitMap;

        case _CaptureParameter.colorDepth of
          4: Result.PixelFormat := pf4bit;
          8: Result.PixelFormat := pf8bit;
          12: Result.PixelFormat := pf24bit;//pfDevice;
          16: Result.PixelFormat := pf24bit;//pf16bit;
          24: Result.PixelFormat := pf24bit;
          32: Result.PixelFormat := pf32bit;
        end;
      end;

      if Assigned(Result) then begin
        //Ϊ���ⷢ��ʱ��ռ�ù���ʱ�䣬���ʹ��postmessage��������������ʾ
        if _CaptureParameter.IsSoundHint then PostMessage(Self.Handle, WM_BEEPMSG, 0, 0);

        Exit;
      end else begin
        //��ʹ��SampleGrabber���ܲɼ���ͼ��ʱ����ֱ��ʹ��VideoWindow��VMRGetBitmap�����ɼ�ͼ��
        if _VideoWindow.VMRGetBitmap(bmpStream) then begin

          curBitMap.LoadFromStream(bmpStream);

          Result := curBitMap;

          case _CaptureParameter.colorDepth of
            4: Result.PixelFormat := pf4bit;
            8: Result.PixelFormat := pf8bit;
            12: Result.PixelFormat := pf24bit;//pfDevice;
            16: Result.PixelFormat := pf24bit;//pf16bit;
            24: Result.PixelFormat := pf24bit;
            32: Result.PixelFormat := pf32bit;
          end;

          //Ϊ���ⷢ��ʱ��ռ�ù���ʱ�䣬���ʹ��postmessage��������������ʾ
          if _CaptureParameter.IsSoundHint then PostMessage(Self.Handle, WM_BEEPMSG, 0, 0);

          Exit;
        end;
      end;

      Result := nil;
    finally
      FreeAndNil(bmpStream);

      //etick := GetTickCount;
      //ShowMessage(IntToStr(etick - stick ));
    end;
  except
    on e: Exception do begin
      Result := nil;
      if Assigned(curBitMap) then FreeAndNil(curBitMap);

      TDebug.DebugMsg('TDSCapture', 'StopPreview', e.Message);
    end;
  end;
end;

function TDSCapture.ShowVideoCaptureFilterCfg(
  parentHandle: Integer): WideString;
begin
  try
    Result := '';

    if _IsCaptureVideo then begin
      Result := '���ڽ�����Ƶ�ɼ�������ִ�иò�����';
      Exit;
    end;

    if Trim(_CaptureParameter.CaptureDeviceName) = '' then begin
      Result := '��δ���òɼ��豸���ơ�';
      Exit;
    end;

    TDS9Ex.ShowCaptureFilterProperty(_CaptureParameter.CaptureDeviceName, parentHandle);
  except
    on e: Exception do begin
      Result := e.Message;
    end;  
  end;
end;

function TDSCapture.Get_IsFit: WordBool;
begin
  Result := _CaptureParameter.VideoShowModel = smFit;
end;

procedure TDSCapture.Set_IsFit(Value: WordBool);
begin
  _IsFit := Value;

  if (_IsFit = False) and (_IsStretch = False) and (_IsAdjustWindowSize = False) then _CaptureParameter.VideoShowModel := smNormal;
  if (_IsFit = False) and (_IsStretch = False) and (_IsAdjustWindowSize = True) then _CaptureParameter.VideoShowModel := smWindAutoFit;
  if (_IsFit = False) and (_IsStretch = True) and (_IsAdjustWindowSize = True) then _CaptureParameter.VideoShowModel := smStretch;
  if (_IsFit = False) and (_IsStretch = True) and (_IsAdjustWindowSize = False) then _CaptureParameter.VideoShowModel := smStretch;

  if _IsFit then begin
    _CaptureParameter.VideoShowModel := smFit;
    _IsStretch := False;
    _IsAdjustWindowSize := False;
  end;

  //AdjustWindowSize();
end;

//��ʾ�ɼ�pin����
function TDSCapture.ShowVideoCapturePinCfg(
  parentHandle: Integer): WideString;
var
  //pinlist: TPinList;
  //i: Integer;
  //pPinOut: IPin;
  curVideoSize: TVideoSize;
begin
  try
    Result := '';

    if not _IsPreviewState then begin
      Result := '��δ����Ԥ��ģʽ������ִ�иò�����';
      Exit;
    end;  

    if _IsCaptureVideo then begin
      Result := '���ڽ�����Ƶ�ɼ������ܶ���������á�';
      Exit;
    end;
      
    //��������Ҫ��ִ��ֹͣ����
    _FilterGraphic.Stop();
    //pinlist := TPinList.Create(_CapSourceFilter);
    try
      //����CAPTURE PIN
      //(_FilterGraphic as ICaptureGraphBuilder2).FindPin(_SourceFilter as IbaseFilter, PINDIR_OUTPUT, @PIN_CATEGORY_PREVIEW, @MEDIATYPE_Video, false, 0, pPinOut);

      //����˵pinlist.Items[0]��ʾԤ���ķֱ��ʣ�pinlist.Items[1]��ʾ����ķֱ���
      {for i := 0 to pinlist.Count - 1 do begin
        if (pinlist.PinInfo[i].dir = PINDIR_OUTPUT) and (pinlist.Connected[i]) then begin
          TDS9Ex.ShowPinPropertyPage('��Ƶ�˿�', parentHandle, pinlist.Items[i]);
          exit;
        end;
      end;}

      //ʹ��amcap��ʵ�ַ�ʽ��ʾ�ɼ��˿�����
      curVideoSize := TCaptureParameterConfig.ConvertVideoSizeInf(_captureParameter.VideoSize);

      TDS9Ex.ShowPinPropertyPage1('��Ƶ�˿�', parentHandle, _GraphManager, _CapSourceFilter, curVideoSize);
    finally
      //FreeAndNil(pinlist);
      _FilterGraphic.Play;
    end;
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

function TDSCapture.ShowVfwVideoDisplayCfg(
  parentHandle: Integer): WideString;
begin
  try
    Result := '';

    if not _IsPreviewState then begin
      Result := '��δ����Ԥ��ģʽ������ִ�иò�����';
      Exit;
    end;

    if _IsCaptureVideo then begin
      Result := '���ڽ�����Ƶ�ɼ������ܶ���������á�';
      Exit;
    end;
      
    //ShowVfwConfigDialog(VfwCaptureDialog_Display, parentHandle);
    TDS9Ex.ShowFilterPropertyPage('��Ƶ��ʾ����', parentHandle, _VideoWindow as IBaseFilter);

  except
    on e: Exception do begin
      Result := e.Message;
    end;  
  end;  
end;

function TDSCapture.ShowVfwVideoFormatCfg(
  parentHandle: Integer): WideString;
//var
//  hr: HRESULT;  
begin
  try
    Result := '';

    if not _IsPreviewState then begin
      Result := '��δ����Ԥ��ģʽ������ִ�иò�����';
      Exit;
    end;  
    
    if _IsCaptureVideo then begin
      Result := '���ڽ�����Ƶ�ɼ������ܶ���������á�';
      Exit;
    end;

    //ShowVfwConfigDialog(VfwCaptureDialog_Format, parentHandle);

    _FilterGraphic.Stop();
    try
      TDS9Ex.ShowFilterPropertyPage('��Ƶ��ʽ', parentHandle, _CapSourceFilter as IBaseFilter, ppVFWCapSource);
      //if Succeeded(hr) then begin
      //  TDS9Ex.ShowFilterPropertyPage('��Ƶ��ʽ', parentHandle, _CapSourceFilter as IBaseFilter, ppDefault);
      //end;
    finally
      _FilterGraphic.Play();
    end;
  except
    on e: Exception do begin
      Result := e.Message;
    end;  
  end;  
end;

function TDSCapture.ShowVfwVideoSourceCfg(
  parentHandle: Integer): WideString;
var
  hr: HRESULT;
begin
  try
    Result := '';

    if not _IsPreviewState then begin
      Result := '��δ����Ԥ��ģʽ������ִ�иò�����';
      Exit;
    end;

    if _IsCaptureVideo then begin
      Result := '���ڽ�����Ƶ�ɼ������ܶ���������á�';
      Exit;
    end;

    _FilterGraphic.Stop();

    try
      //ShowVfwConfigDialog(VfwCaptureDialog_Source, parentHandle);
      hr := TDS9Ex.ShowFilterPropertyPage('��ƵԴ', parentHandle, _CapSourceFilter as IBaseFilter, ppVFWCapSource);
      if Succeeded(hr) then begin
        TDS9Ex.ShowFilterPropertyPage('��ƵԴ', parentHandle, _CapSourceFilter as IBaseFilter, ppDefault);
      end;
    finally
      _FilterGraphic.Play();
    end;
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

procedure TDSCapture.ShowVfwConfigDialog(
  const dialogType: Integer; const parentHandle: Integer);
var
  hr: HRESULT;
  vfwConfig: IAMVfwCaptureDialogs;
begin
  if not _IsPreviewState then Exit;

  //��������Ҫ��ִ��ֹͣ����
  _FilterGraphic.Stop;
  try
    hr := _CapSourceFilter.QueryInterface(IID_IAMVfwCaptureDialogs, vfwConfig);
    if Succeeded(hr) then begin
      if vfwConfig.HasDialog(dialogType) = S_OK then
        vfwConfig.ShowDialog(dialogType, parentHandle);
    end;
  finally
    _FilterGraphic.Play;
    vfwConfig := nil;
  end;
end;

procedure TDSCapture.ShowVfwCompressCfgDialog(const dialogType,
  parentHandle: Integer);
begin
  raise Exception.Create('��δʵ�ָù��ܡ�');
end;

procedure TDSCapture.vfwConfigCallEvent(const operVfwConfigType: TVfwConfigType;
  const parentHandle: Integer; out errMsg: WideString);
begin
  case operVfwConfigType of
    vctVideoSourceProperty: begin //��ʾVFWԴ����
      errMsg := ShowVideoCaptureFilterCfg(parentHandle);
    end;
    vctVideoCapturePinProperty: begin  //��ʾ�ɼ��˿���������
      errMsg := ShowVideoCapturePinCfg(parentHandle);
    end;
    vctVfwVideoFormat: begin   //��ʾ��Ƶ��ʽ����
      errMsg := ShowVfwVideoFormatCfg(parentHandle);
    end;
    vctVfwVideoDisplay: begin  //��ʾ��Ƶ��ʾ����
      errMsg := ShowVfwVideoDisplayCfg(parentHandle);
    end;
    vctVideoCrossbar: begin    //��ʾvideo Crossbar����
      errMsg := ShowVideoCrossbarCfg(parentHandle);
    end;
    vctVfwCompressDialog: begin  //��ʾ��Ƶѹ������
      errMsg := ShowVfwCompressCfg(parentHandle);
    end; 
  end;
end;

function TDSCapture.StartPreview: WideString;
begin
  Result := '';
  
  if _IsPreviewState then Exit;

  Preview(False, Result);
end;

function TDSCapture.Get_CaptureState: WordBool;
begin
  Result := _IsCaptureVideo;
end;

function TDSCapture.Get_PreviewState: WordBool;
begin
  Result := _IsPreviewState;
end;

function TDSCapture.ReadParameterFromFile: WideString;
begin
  try
    Result := '';
    TfrmCapParameterCfg.ReadCaptureParameterFromFile(_CaptureParameterCfgFileName, _CaptureParameter);

    stabStates.Visible := _CaptureParameter.IsShowState;
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

function TDSCapture.CaptureJpgImageToFile(const fileName: WideString;
  compressRate: SYSINT): WideString;
var
  bitMap, cutBmp: TBitmap;
  cutArea: TRect;
  jpg: TJPEGImage;
begin
  try
    Result := '';

    if not _IsPreviewState then begin
      Result := '��Ƶ�ɼ���δ����Ԥ��ģʽ��';
      Exit;
    end;

    bitMap := CaptureImageToBmpObj;
    try
      if not Assigned(bitMap) then begin
        Result := 'û�вɼ�����Ƶͼ��';
        Exit;
      end;

      //ת��Ϊ�Ҷ�ͼ
      //�ɼ���ͼ��һ�㶼��1024*768���ڣ����Լ����ǲ������ü����лҶ�ת����
      //��Ч����Ҳû�ж���Ӱ��
      if _CaptureParameter.IsConvertGrayImg then begin
        TGraphicProcess.ConvertBitmapToGrayscale(bitMap);
      end;


      //ͼ��ü�����
      if _CaptureParameter.IsApplyImageCut then begin
        cutBmp := TBitmap.Create;
        try
          //ȡ�òü���Χ
          cutArea.Left := Round(_CaptureParameter.LeftRate * bitMap.Width);
          cutArea.Right := cutArea.Left + Round(_CaptureParameter.WidthRate * bitMap.Width);

          cutArea.Top := Round(_CaptureParameter.TopRate * bitMap.Height);
          cutArea.Bottom := cutArea.Top + Round(_CaptureParameter.HeightRate * bitMap.Height);

          TGraphicProcess.CutImg(cutArea, bitMap, cutBmp);

          jpg := TGraphicProcess.BmpConvertToJpg(cutBmp, compressRate);
        finally
          FreeAndNil(cutBmp);
        end;
      end else begin
        //ֱ�ӱ���ɼ�ͼ��
        jpg := TGraphicProcess.BmpConvertToJpg(bitMap, compressRate);
      end;

      if not Assigned(jpg) then begin
        Result := 'ͼ��ת��ʧ�ܡ�';
        Exit;
      end;

      jpg.SaveToFile(fileName);
    finally
      FreeAndNil(bitMap);
      if Assigned(jpg) then FreeAndNil(jpg);
    end;
  except
    on e: Exception do begin
      Result := e.Message;
    end;  
  end;
end;

function TDSCapture.Get_IsClickQuitFullScreen: WordBool;
begin
  Result := _IsClickQuitFullScreen;
end;

function TDSCapture.Get_IsDblClickQuitFullScreen: WordBool;
begin
  Result := _IsDblClickQuitFullScreen;
end;

function TDSCapture.Get_IsEscKeyQuitFullScreen: WordBool;
begin
  Result := _IsEscKeyQuitFullScreen;
end;

procedure TDSCapture.Set_IsClickQuitFullScreen(Value: WordBool);
begin
  _IsClickQuitFullScreen := Value;
end;

procedure TDSCapture.Set_IsDblClickQuitFullScreen(Value: WordBool);
begin
  _IsDblClickQuitFullScreen := Value;
end;

procedure TDSCapture.Set_IsEscKeyQuitFullScreen(Value: WordBool);
begin
  _IsEscKeyQuitFullScreen := Value;
end;

procedure TDSCapture.KeyDownEvent(Sender: TObject; var Key: Word; Shift: TShiftState);
var
  curShift: Integer;
  curkey: Integer;
begin
  try
    if FEvents <> nil then begin
      curShift := 0;
      move(Shift, curShift, sizeof(TShiftState));

      curkey := Key;
      FEvents.OnKeyDown(curkey, curShift);
      Key := curkey;
    end;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSCapture', 'KeyDownEvent', e.Message);
    end;
  end;      
end;

procedure TDSCapture.KeyUpEvent(Sender: TObject; var Key: Word; Shift: TShiftState);
var
  curShift: Integer;
  curKey: Integer;
begin
  try
    if FEvents <> nil then begin
      curShift := 0;
      move(Shift, curShift, sizeof(TShiftState));

      curKey := Key;
      FEvents.OnKeyUp(curKey, curshift);
      Key := curKey;
    end;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSCapture', 'KeyUpEvent', e.Message);
    end;
  end;      
end;

procedure TDSCapture.MouseDownEvent(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  curShift: Integer;
begin
  try
    if FEvents <> nil then begin
      curShift := 0;
      move(Shift, curShift, sizeof(TShiftState));

      FEvents.OnMouseDown(Integer(Button), curShift, x, y);
    end;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSCapture', 'MouseDownEvent', e.Message);
    end;
  end;    
end;

procedure TDSCapture.MouseMoveEvent(Sender: TObject; Shift: TShiftState; X, Y: Integer);
var
  curShift: Integer;
begin
  try
    if FEvents <> nil then begin
      curShift := 0;
      move(Shift, curShift, sizeof(TShiftState));

      FEvents.OnMouseMove(curShift, x, y);
    end;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSCapture', 'MouseMoveEvent', e.Message);
    end;
  end;      
end;

procedure TDSCapture.MouseUpEvent(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  curShift: Integer;
begin
  try
    if FEvents <> nil then begin
      curShift := 0;
      move(Shift, curShift, sizeof(TShiftState));

      FEvents.OnMouseUp(Integer(Button), curShift, x, y);
    end;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSCapture', 'MouseUpEvent', e.Message);
    end;
  end;    
end;

procedure TDSCapture.ResizeEvent(Sender: TObject);
begin
  //AdjustWindowSize(); //����ʾ���ڴ�С�����ı�ʱ���Զ��ı���Ƶ�������ʾλ��(�����ⲿ����RefreshWindow����)

  try
    if FEvents <> nil then FEvents.OnResize;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSCapture', 'ResizeEvent', e.Message);
    end;
  end;    
end;

procedure TDSCapture.EnterEvent(Sender: TObject);
begin
  try
    if FEvents <> nil then FEvents.OnGotFocus;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSCapture', 'EnterEvent', e.Message);
    end;
  end;    
end;

procedure TDSCapture.ExitEvent(Sender: TObject);
begin
  try
    if FEvents <> nil then FEvents.OnLostFocus;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSCapture', 'ExitEvent', e.Message);
    end;
  end;    
end;

function TDSCapture.Get_CurHeight: Integer;
begin
  Result := Self.Height;
end;

function TDSCapture.Get_CurVideoHeight: Integer;
begin
  Result := _VideoWindow.Height;
end;

function TDSCapture.Get_CurVideoWidth: Integer;
begin
  Result := _VideoWindow.Width;
end;

function TDSCapture.Get_CurWidth: Integer;
begin
  Result := Self.Width;
end;

procedure TDSCapture.Set_CurHeight(Value: Integer);
begin
  Self.Height := Value;
end;

procedure TDSCapture.Set_CurVideoHeight(Value: Integer);
begin
  _VideoWindow.Height := Value;
end;

procedure TDSCapture.Set_CurVideoWidth(Value: Integer);
begin
  _VideoWindow.Width := Value;
end;

procedure TDSCapture.Set_CurWidth(Value: Integer);
begin
  Self.Width := Value;
end;

function TDSCapture.RefreshWindow: WideString;
begin
  try
    Result := '';

    AdjustWindowSize();
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

function TDSCapture.Get_ShowModel: TShowModel;
begin
  Result := _CaptureParameter.VideoShowModel;
end;

procedure TDSCapture.Set_ShowModel(Value: TShowModel);
begin
  _CaptureParameter.VideoShowModel := Value;

  //AdjustWindowSize();
end;

procedure TDSCapture.VideoSizeEvent(const videoWidth, videoHeight, windowWidth, windowHeight: Integer);
begin
  try
    if FEvents <> nil then FEvents.OnVideoSizeChange(videoWidth, videoHeight, windowWidth, windowHeight);
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSCapture', 'VideoSizeEvent', e.Message);
    end;
  end;    
end;

function TDSCapture.Get_CapParameterWindPos: TCapParameterPostion;
begin
  Result := _CapParameterWindowPos;
end;

procedure TDSCapture.Set_CapParameterWindPos(Value: TCapParameterPostion);
begin
  _CapParameterWindowPos := Value;
end;

function TDSCapture.QuitFullScreen: WideString;
begin
  if _IsPreviewState then
    Result := TfrmFullScreen.QuitFullScreen();
end;

function TDSCapture.ShowFullScreen(parentHandle,
  monitorIndex: Integer): WideString;
begin
  if _IsPreviewState then
    Result := TfrmFullScreen.ShowFullScreen(_VideoWindow, _CaptureParameter.VideoShowModel, monitorIndex);
end;


//------------------------------------------------------------------------------


procedure TDSCapture._VideoWindowKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  //�˳�ȫ��ģʽ
  if ((Key = VK_ESCAPE) and _IsEscKeyQuitFullScreen)
    or (Key = VK_LWIN)
    or (Key = VK_RWIN)
    //or ((ssCtrl in Shift) and (ssAlt in Shift))
    or (ssCtrl in Shift)
    or (ssAlt in Shift) then begin
    QuitFullScreen();
  end;

  if Assigned(Self.OnKeyDown) then Self.OnKeyDown(Sender, Key, Shift);
end;

procedure TDSCapture._VideoWindowMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if Assigned(Self.OnMouseDown) then Self.OnMouseDown(Sender, Button, Shift, X, Y);
end;

procedure TDSCapture._VideoWindowMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if Assigned(Self.OnMouseUp) then Self.OnMouseUp(Sender, Button, Shift, X, Y);
end;

procedure TDSCapture._VideoWindowKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Assigned(Self.OnKeyUp) then Self.OnKeyUp(Sender, Key, Shift);
end;

procedure TDSCapture._VideoWindowEnter(Sender: TObject);
begin
  if Assigned(Self.OnEnter) then Self.OnEnter(Sender);
end;

procedure TDSCapture._VideoWindowExit(Sender: TObject);
begin
  if Assigned(Self.OnExit) then Self.OnExit(Sender);
end;

procedure TDSCapture._VideoWindowMouseWheel(Sender: TObject;
  Shift: TShiftState; WheelDelta: Integer; MousePos: TPoint;
  var Handled: Boolean);
begin
  if Assigned(Self.OnMouseWheel) then Self.OnMouseWheel(Sender, Shift, WheelDelta, MousePos, Handled);
end;

procedure TDSCapture._VideoWindowMouseWheelDown(Sender: TObject;
  Shift: TShiftState; MousePos: TPoint; var Handled: Boolean);
begin
  if Assigned(Self.OnMouseWheelDown) then Self.OnMouseWheelDown(Sender, Shift, MousePos, Handled);
end;

procedure TDSCapture._VideoWindowMouseWheelUp(Sender: TObject;
  Shift: TShiftState; MousePos: TPoint; var Handled: Boolean);
begin
  if Assigned(Self.OnMouseWheelUp) then Self.OnMouseWheelUp(Sender, Shift, MousePos, Handled);
end;


//------------------------------------------------------------------------------


procedure TDSCapture._VideoWindowPaint(Sender: TObject);
var
  vmrWindCtrl9: IVMRWindowlessControl9;
  vw: IVideoWindow;
  hr: HRESULT;
  videoDc: HDC;
begin
  if imgLogo.Visible then begin
    imgLogo.Left := (_VideoWindow.Width - imgLogo.Width) div 2;
    imgLogo.Top := (_VideoWindow.Height - imgLogo.Height) div 2;
  end;

  //��������Ļ�������������ʱ���ڻָ���ʱ��ˢ����Ƶ��ʾ
  if not _IsPreviewState then Exit;

  if _CaptureParameter.SnatchWay = swVMR then begin
    //Vmr��ʽ��ˢ��
    hr := (_VideoWindow as IBaseFilter).QueryInterface(IID_IVMRWindowlessControl9, vmrWindCtrl9);
    if Failed(hr) then Exit;

    videoDc := GetDC(_VideoWindow.Handle);
    try
      vmrWindCtrl9.RepaintVideo(_VideoWindow.Handle, videoDc);
    finally
      vmrWindCtrl9 := nil;
      ReleaseDC(_VideoWindow.Handle, videoDc);
    end;
  end else begin
    //device��ʽ��ˢ��
    hr := (_VideoWindow as IBaseFilter).QueryInterface(IID_IVideoWindow, vw);
    if Failed(hr) then Exit;

    try
      //ˢ����Ƶ
      vw.put_Visible(True);
    finally
      vw := nil;
    end;
  end;
end;

procedure TDSCapture.MouseWheelEvent(Sender: TObject; Shift: TShiftState;
  WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
var
  curShift: Integer;
  curHandled: WordBool;
begin
  try
    if FEvents <> nil then begin
      curShift := 0;
      move(Shift, curShift, sizeof(TShiftState));

      curHandled := Handled;
      FEvents.OnMouseWheel(curShift, WheelDelta, MousePos.X, MousePos.Y, curHandled);
      Handled := curHandled;
    end;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSCapture', 'MouseWheelEvent', e.Message);
    end;
  end;    
end;

procedure TDSCapture.MouseWheelDownEvent(Sender: TObject;
  Shift: TShiftState; MousePos: TPoint; var Handled: Boolean);
var
  curShift: Integer;
  curHandled: WordBool;
begin
  try
    if FEvents <> nil then begin
      curShift := 0;
      move(Shift, curShift, sizeof(TShiftState));

      curHandled := Handled;
      FEvents.OnMouseWheelDown(curShift, MousePos.X, MousePos.Y, curHandled);
      Handled := curHandled;
    end;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSCapture', 'MouseWheelDownEvent', e.Message);
    end;
  end;    
end;

procedure TDSCapture.MouseWheelUpEvent(Sender: TObject; Shift: TShiftState;
  MousePos: TPoint; var Handled: Boolean);
var
  curShift: Integer;
  curHandled: WordBool;
begin
  try
    if FEvents <> nil then begin
      curShift := 0;
      move(Shift, curShift, sizeof(TShiftState));

      curHandled := Handled;
      FEvents.OnMouseWheelUp(curShift, MousePos.X, MousePos.Y, curHandled);
      Handled := curHandled;
    end;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSCapture', 'MouseWheelUpEvent', e.Message);
    end;
  end;    
end;

function TDSCapture.Get_SnatchWay: TSnatchWay;
begin
  Result := _CaptureParameter.SnatchWay;
end;

procedure TDSCapture.Set_SnatchWay(Value: TSnatchWay);
var
  sErrMsg: WideString;
begin
  if _CaptureParameter.SnatchWay = Value then Exit;
  
  _CaptureParameter.SnatchWay := Value;
                  
  if _IsPreviewState and not _IsCaptureVideo then begin
    //���¿�ʼԤ��
    Preview(False, sErrMsg);
  end;
end;


function TDSCapture.UpdateVideoQuailty: WideString;
begin
  if not TDS9Ex.IsVfwDevice(_CaptureParameter.CaptureDeviceName)
    and _IsPreviewState then begin
    //������Ƶ��������VFW���豸���򲻽�������
    ConfigVideoQuality(_CapSourceFilter, _CaptureParameter);
  end;
end;

function TDSCapture.SaveParameterToFile: WideString;
begin
  try
    Result := '';

    //д��ɼ�����
    TfrmCapParameterCfg.WriteCaptureParameterToFile(_CaptureParameterCfgFileName, _CaptureParameter);
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

procedure TDSCapture._VideoWindowMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  if Assigned(Self.OnMouseMove) then Self.OnMouseMove(Sender, Shift, x, y);
end;

function TDSCapture.GetCaptureParameter(
  var parameter: TCaptureParameter): WideString;
begin
  TCaptureParameterConfig.CopyParameter(_CaptureParameter, parameter);
end;

function TDSCapture.SetCaptureParameter(
  var parameter: TCaptureParameter): WideString;
begin
  TCaptureParameterConfig.CopyParameter(parameter, _CaptureParameter);
end;

function TDSCapture.RePreview: WideString;
begin
  if _IsCaptureVideo then begin
    Result := '���ڽ�����Ƶ�ɼ�������ִ�иò�����';
    Exit;
  end;

  Preview(False, Result);
end;

function TDSCapture.ShowVideoEncoderFilterCfg(
  parentHandle: Integer): WideString;
begin
  try
    Result := '';

    if _IsCaptureVideo then begin
      Result := '���ڽ�����Ƶ�ɼ�������ִ�иò�����';
      Exit;
    end;

    if Trim(_CaptureParameter.encoderName) = '' then begin
      Result := '��δ������Ƶ���������ơ�';
      Exit;
    end;

    TDS9Ex.ShowEncoderFilterProperty(_CaptureParameter.encoderName, parentHandle);
  except
    on e: Exception do begin
      Result := e.Message;
    end;  
  end;
end;

function TDSCapture.Get_ParameterCfgFileName: WideString;
begin
  Result := _CaptureParameterCfgFileName;
end;

procedure TDSCapture.Set_ParameterCfgFileName(const Value: WideString);
begin
  _CaptureParameterCfgFileName := Value;
end;


function TDSCapture.Get_HideCfgItem: Integer;
begin
  Result := _HideCfgItem;
end;

procedure TDSCapture.Set_HideCfgItem(Value: Integer);
begin
  _HideCfgItem := Value;
end;



function TDSCapture.Get_AppHandle: Integer;
begin
  Result := Application.Handle;
end;

procedure TDSCapture.Set_AppHandle(Value: Integer);
begin
  Application.Handle := Value;
end;

function TDSCapture.CaptureImgToClipBoard: WideString;
var
  bitMap: TBitmap;
  data: THandle;
  palette :HPALETTE;
  cutArea: TRect;
  cutBmp: TBitmap;
  curFormat: Word;
  //H: THandle;
  //P: Pointer;
begin
  try
    Result := '';

    bitMap := CaptureImageToBmpObj;
    try
      if not Assigned(bitMap) then begin
        Result := 'û�вɼ�����Ƶͼ��';
        Exit;
      end;

      //ת��Ϊ�Ҷ�ͼ
      //�ɼ���ͼ��һ�㶼��1024*768���ڣ����Լ����ǲ������ü����лҶ�ת����
      //��Ч����Ҳû�ж���Ӱ��
      if _CaptureParameter.IsConvertGrayImg then begin
        TGraphicProcess.ConvertBitmapToGrayscale(bitMap);
      end;


      //ͼ��ü�����
      if _CaptureParameter.IsApplyImageCut then begin
        cutBmp := TBitmap.Create;
        try
          //ȡ�òü���Χ
          cutArea.Left := Round(_CaptureParameter.LeftRate * bitMap.Width);
          cutArea.Right := cutArea.Left + Round(_CaptureParameter.WidthRate * bitMap.Width);

          cutArea.Top := Round(_CaptureParameter.TopRate * bitMap.Height);
          cutArea.Bottom := cutArea.Top + Round(_CaptureParameter.HeightRate * bitMap.Height);

          TGraphicProcess.CutImg(cutArea, bitMap, cutBmp);



          cutBmp.SaveToClipboardFormat(curFormat, data, palette);

          {curFormat := format;
          if curFormat <= 0 then begin
            curFormat := RegisterClipboardFormat('ZLDSVIDEOPROCESS10161');

            H:= GlobalAlloc(GMEM_DDESHARE, GlobalSize(data)); //����һ���ڴ�
            //P:= GlobalLock(H);//ȡ��ָ���ڴ���ָ��

            //move(Pchar('')^, p^, GlobalSize(data));

            //GlobalUnlock(H);

          end;}

          Clipboard.SetAsHandle(curFormat, data);
          //format := curFormat;
        finally
          FreeAndNil(cutBmp);
        end;
      end else begin
        //ֱ�ӽ�ͼ���Ƶ�������
        bitMap.SaveToClipboardFormat(curFormat, data, palette);

        {curFormat := format;
        if curFormat <= 0 then begin
          curFormat := RegisterClipboardFormat('ZLDSVIDEOPROCESS10161');

          H:= GlobalAlloc(GMEM_DDESHARE, GlobalSize(data)); //����һ���ڴ�
          //P:= GlobalLock(H);//ȡ��ָ���ڴ���ָ��

          //move(Pchar('')^, p^, GlobalSize(data));

          //GlobalUnlock(H);
        end;}

        Clipboard.SetAsHandle(curFormat, data);
        //format := curFormat;
      end;
    finally
      FreeAndNil(bitMap);
    end;          
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;
                                      
function TDSCapture.ShowVfwCompressCfg(parentHandle: Integer): WideString;
begin
  try
    Result := '';

    if not _IsPreviewState then begin
      Result := '��δ����Ԥ��ģʽ������ִ�иò�����';
      Exit;
    end;

    if _IsCaptureVideo then begin
      Result := '���ڽ�����Ƶ�ɼ������ܶ���������á�';
      Exit;
    end;

    //_FilterGraphic.Stop();
    //try
      ShowVfwCompressCfgDialog(VfwCompressDialog_Config, parentHandle);
    //finally
    //  _FilterGraphic.Play();
    //end;
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

function TDSCapture.ShowVideoCrossbarCfg(
  parentHandle: Integer): WideString;
begin
  try
    Result := '';

    if not _IsPreviewState then begin
      Result := '��δ����Ԥ��ģʽ������ִ�иò�����';
      Exit;
    end;  

    if _IsCaptureVideo then begin
      Result := '���ڽ�����Ƶ�ɼ������ܶ���������á�';
      Exit;
    end;

    //��������Ҫ��ִ��ֹͣ����  ��amcap��û��ֹͣfilter graph �����ھ�����Ե�ʱ��
    //�����ֹͣ���п��ܵ�һ�ν�����������ʾ������û���κη�ӳ
    _FilterGraphic.Stop();
    try

      TDS9Ex.ShowVideoCrossbarPropertyPage('��Ƶ�˿�', parentHandle, _GraphManager, _CapSourceFilter);
    finally
      _FilterGraphic.Play;
    end;
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

function TDSCapture.CaptureBmpImage: IPictureDisp;
var
  bitMap, cutBmp: TBitmap;
  cutArea: TRect;
  picImage: TPicture;
begin
  try
    Result := nil;

    if not _IsPreviewState then begin
      Result := nil;
      Exit;
    end;

    bitMap := CaptureImageToBmpObj;
    try
      if not Assigned(bitMap) then begin
        Result := nil;
        Exit;
      end;

      //ת��Ϊ�Ҷ�ͼ
      //�ɼ���ͼ��һ�㶼��1024*768���ڣ����Լ����ǲ������ü����лҶ�ת����
      //��Ч����Ҳû�ж���Ӱ��
      if _CaptureParameter.IsConvertGrayImg then begin
        TGraphicProcess.ConvertBitmapToGrayscale(bitMap);
      end;


      //ͼ��ü�����
      if _CaptureParameter.IsApplyImageCut then begin
        cutBmp := TBitmap.Create;
        try
          //ȡ�òü���Χ
          cutArea.Left := Round(_CaptureParameter.LeftRate * bitMap.Width);
          cutArea.Right := cutArea.Left + Round(_CaptureParameter.WidthRate * bitMap.Width);

          cutArea.Top := Round(_CaptureParameter.TopRate * bitMap.Height);
          cutArea.Bottom := cutArea.Top + Round(_CaptureParameter.HeightRate * bitMap.Height);

          TGraphicProcess.CutImg(cutArea, bitMap, cutBmp);

          picImage := TPicture.Create;
          try
            picImage.Assign(cutBmp);
            GetOlePicture(picImage, Result);
          finally
            FreeAndNil(picImage);
          end;
        finally
          FreeAndNil(cutBmp);
        end;

        exit;
      end;

      //ֱ�ӱ���ɼ�ͼ��
      picImage := TPicture.Create;
      try
        picImage.Assign(bitMap);
        GetOlePicture(picImage, Result);
      finally
        FreeAndNil(picImage);
      end;

    finally
      FreeAndNil(bitMap);
    end;
  except
    on e: Exception do begin
      Result := nil;
    end;  
  end;
end;

function TDSCapture.CaptureJpgImage(compressRate: Integer): IPictureDisp;
var
  bitMap, cutBmp: TBitmap;
  cutArea: TRect;
  picImage: TPicture;
  jpg: TJPEGImage;
begin
  try
    Result := nil;

    if not _IsPreviewState then begin
      Result := nil;
      Exit;
    end;

    bitMap := CaptureImageToBmpObj;
    try
      if not Assigned(bitMap) then begin
        Result := nil;
        Exit;
      end;

      //ת��Ϊ�Ҷ�ͼ
      //�ɼ���ͼ��һ�㶼��1024*768���ڣ����Լ����ǲ������ü����лҶ�ת����
      //��Ч����Ҳû�ж���Ӱ��
      if _CaptureParameter.IsConvertGrayImg then begin
        TGraphicProcess.ConvertBitmapToGrayscale(bitMap);
      end;


      //ͼ��ü�����
      if _CaptureParameter.IsApplyImageCut then begin
        cutBmp := TBitmap.Create;
        try
          //ȡ�òü���Χ
          cutArea.Left := Round(_CaptureParameter.LeftRate * bitMap.Width);
          cutArea.Right := cutArea.Left + Round(_CaptureParameter.WidthRate * bitMap.Width);

          cutArea.Top := Round(_CaptureParameter.TopRate * bitMap.Height);
          cutArea.Bottom := cutArea.Top + Round(_CaptureParameter.HeightRate * bitMap.Height);

          TGraphicProcess.CutImg(cutArea, bitMap, cutBmp);

          jpg := TGraphicProcess.BmpConvertToJpg(cutBmp, compressRate);
        finally
          FreeAndNil(cutBmp);
        end;  
      end else begin
        jpg := TGraphicProcess.BmpConvertToJpg(bitMap, compressRate);
      end;

      if not Assigned(jpg) then begin
        Result := nil;
        Exit;
      end;

      //����jpgͼ��
      picImage := TPicture.Create;
      try
        picImage.Assign(jpg);
        GetOlePicture(picImage, Result);
      finally
        FreeAndNil(picImage);
      end;

    finally
      FreeAndNil(bitMap);
    end;
  except
    on e: Exception do begin
      Result := nil;
    end;  
  end;
end;

function TDSCapture.GetRealVideoSize: TVideoSize;
var
  pin: IPin;
  amStreamConfig: IAMStreamConfig;
  pmt: PAMMediaType;
  pvih: PVideoInfoHeader;
  
  curSize: TVideoSize;
  hr: HRESULT;
begin
  curSize.Width := 0;
  curSize.Height := 0;

  Result := curSize;

  if not Assigned(_CapSourceFilter) then Exit;

  try
    //�����Ѿ����ӵ�PIN
    hr := TDS9Ex.FindConnectedPin(_CapSourceFilter, PINDIR_OUTPUT, pin);
    if FAILED(hr) then Exit;
    if not Assigned(pin) then Exit;

    try                             
      hr := pin.QueryInterface(IID_IAMStreamConfig, amStreamConfig);
      if FAILED(hr) then Exit;

      try
        hr := amStreamConfig.GetFormat(pmt);   //ȡ��Ĭ����Ƶ��ʽ
        if FAILED(hr) then Exit;

        try
          pvih := pmt.pbFormat;
          curSize.Width := pvih^.bmiHeader.biWidth;
          curSize.Height := pvih^.bmiHeader.biHeight;

          Result := curSize;
        finally
          DeleteMediaType(pmt);
        end;
      finally
        amStreamConfig := nil;
      end;
    finally
      pin := nil;
    end;
  except end;  
end;

procedure TDSCapture.WM_BEEP(var msg: TMessage);
begin
  Windows.Beep(2000, 500);
end;

function TDSCapture.Get_RecordTimeLen: Integer;
begin
  Result := _RecordTimeLen; 
end;

initialization

  TActiveFormFactory.Create(
    ComServer,
    TActiveFormControl,
    TDSCapture,
    Class_DSCapture,
    1,
    '',
    OLEMISC_SIMPLEFRAME or OLEMISC_ACTSLIKELABEL,
    tmApartment);

finalization


end.
